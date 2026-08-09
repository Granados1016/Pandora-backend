using Google.Apis.Auth.OAuth2;
using Google.Apis.Drive.v3;
using Google.Apis.Services;
using DriveFile = Google.Apis.Drive.v3.Data.File;

namespace Pandora.API.Services;

/// <summary>
/// Sube el archivo de backup a Google Drive usando una cuenta de servicio con
/// delegación de dominio (domain-wide delegation), impersonando la cuenta que
/// aloja los backups (ej. sistemas@tuempresa.com).
///
/// Credenciales: se leen de (en este orden)
///   1) GoogleDrive:ServiceAccountJson  — el JSON completo de la cuenta de servicio
///   2) GoogleDrive:ServiceAccountKeyPath — ruta a un archivo .json local (dev)
/// En Railway se recomienda la variable de entorno GoogleDrive__ServiceAccountJson.
/// Ninguna de las dos debe commitearse al repo.
/// </summary>
public class GoogleDriveBackupUploader(IConfiguration config, ILogger<GoogleDriveBackupUploader> logger)
{
    private const string FolderMimeType = "application/vnd.google-apps.folder";

    public record UploadResult(string FileId, string WebViewLink, string FolderId);

    /// <summary>Devuelve null si no hay credenciales configuradas (feature deshabilitada).</summary>
    public bool IsConfigured =>
        !string.IsNullOrWhiteSpace(config["GoogleDrive:ServiceAccountJson"]) ||
        !string.IsNullOrWhiteSpace(config["GoogleDrive:ServiceAccountKeyPath"]);

    public async Task<UploadResult> UploadAsync(
        byte[] bytes, string fileName, string contentType,
        string impersonateEmail, string? cachedFolderId,
        CancellationToken ct)
    {
        if (string.IsNullOrWhiteSpace(impersonateEmail))
            throw new InvalidOperationException("No hay cuenta de Drive configurada para impersonar (GoogleDrive:ImpersonateEmail).");

        using var driveService = BuildDriveService(impersonateEmail);

        var folderId = string.IsNullOrWhiteSpace(cachedFolderId)
            ? await FindOrCreateFolderAsync(driveService, "Pandora Backups", ct)
            : cachedFolderId;

        var fileMetadata = new DriveFile
        {
            Name    = fileName,
            Parents = [folderId]
        };

        using var stream = new MemoryStream(bytes);
        var request = driveService.Files.Create(fileMetadata, stream, contentType);
        request.Fields = "id, webViewLink";
        var progress = await request.UploadAsync(ct);

        if (progress.Status != Google.Apis.Upload.UploadStatus.Completed)
            throw new Exception($"Falló la subida a Drive: {progress.Exception?.Message}");

        var uploaded = request.ResponseBody;
        logger.LogInformation("Backup subido a Google Drive: {File} ({Id})", fileName, uploaded.Id);

        return new UploadResult(uploaded.Id, uploaded.WebViewLink ?? "", folderId);
    }

    private DriveService BuildDriveService(string impersonateEmail)
    {
        var json = config["GoogleDrive:ServiceAccountJson"];
        GoogleCredential credential;

        if (!string.IsNullOrWhiteSpace(json))
        {
            credential = GoogleCredential.FromJson(json);
        }
        else
        {
            var path = config["GoogleDrive:ServiceAccountKeyPath"]
                ?? throw new InvalidOperationException("No hay credenciales de Google Drive configuradas.");
            if (!System.IO.File.Exists(path))
                throw new FileNotFoundException("No se encontró el archivo de credenciales de Google Drive.", path);
            using var stream = System.IO.File.OpenRead(path);
            credential = GoogleCredential.FromStream(stream);
        }

        credential = credential
            .CreateScoped(DriveService.Scope.Drive)
            .CreateWithUser(impersonateEmail);

        return new DriveService(new BaseClientService.Initializer
        {
            HttpClientInitializer = credential,
            ApplicationName       = "Pandora Backups"
        });
    }

    private static async Task<string> FindOrCreateFolderAsync(
        DriveService driveService, string folderName, CancellationToken ct)
    {
        var listRequest = driveService.Files.List();
        listRequest.Q = $"name = '{folderName}' and mimeType = '{FolderMimeType}' and trashed = false";
        listRequest.Fields = "files(id, name)";
        var existing = await listRequest.ExecuteAsync(ct);

        if (existing.Files is { Count: > 0 })
            return existing.Files[0].Id;

        var folder = new DriveFile { Name = folderName, MimeType = FolderMimeType };
        var createRequest = driveService.Files.Create(folder);
        createRequest.Fields = "id";
        var created = await createRequest.ExecuteAsync(ct);
        return created.Id;
    }
}
