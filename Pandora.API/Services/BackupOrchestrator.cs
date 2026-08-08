using Microsoft.Data.SqlClient;

namespace Pandora.API.Services;

/// <summary>
/// Orquesta un ciclo completo de backup: genera el archivo (BackupService),
/// lo reparte por correo y Google Drive según la configuración en
/// dbo.SystemSettings, y deja constancia en dbo.BackupHistory.
/// Compartido por <see cref="AutomatedBackupWorker"/> (cron diario) y por el
/// endpoint manual "Ejecutar ahora" del panel Admin.
/// </summary>
public class BackupOrchestrator(
    IConfiguration config,
    BackupService backupService,
    GoogleDriveBackupUploader driveUploader,
    ILogger<BackupOrchestrator> logger)
{
    private SqlConnection Conn() => new(config.GetConnectionString("PandoraDb"));

    public record RunSummary(
        string FileName, string Method, long SizeBytes,
        List<string> EmailedTo, string? EmailError,
        bool DriveUploaded, string? DriveError, string? DriveLink);

    public async Task<RunSummary> RunAsync(CancellationToken ct)
    {
        var settings = await LoadSettingsAsync(ct);
        var backup   = await backupService.GenerateAsync(ct);
        var contentType = backup.Method == "bak" ? "application/octet-stream" : "text/plain; charset=utf-8";

        // ── Correo ─────────────────────────────────────────────────────────
        var emailedTo  = new List<string>();
        string? emailError = null;
        var recipients = (settings.GetValueOrDefault("backup_recipient_emails") ?? "")
            .Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);

        if (recipients.Length > 0)
        {
            var connStr = config.GetConnectionString("PandoraDb")!;
            var smtpCfg = await SmtpHelper.LoadAsync(connStr, config);
            var attachment = new EmailAttachment(backup.FileName, backup.Bytes, contentType);
            var htmlBody = BuildEmailBody(backup.FileName, backup.Method, backup.Bytes.Length);

            foreach (var to in recipients)
            {
                var err = await SmtpHelper.SendAsync(smtpCfg, to, to,
                    $"🗄️ Backup automático de Pandora — {DateTime.Now:dd/MM/yyyy}", htmlBody, [attachment]);
                if (err != null)
                {
                    emailError = err;
                    logger.LogWarning("Backup automático: correo a {To} falló: {Err}", to, err);
                }
                else emailedTo.Add(to);
            }
        }

        // ── Google Drive ───────────────────────────────────────────────────
        var driveUploaded   = false;
        string? driveError  = null;
        string? driveLink   = null;
        var impersonate     = settings.GetValueOrDefault("backup_drive_impersonate_email") ?? "";
        var cachedFolderId  = settings.GetValueOrDefault("backup_drive_folder_id");

        if (driveUploader.IsConfigured && !string.IsNullOrWhiteSpace(impersonate))
        {
            try
            {
                var result = await driveUploader.UploadAsync(
                    backup.Bytes, backup.FileName, contentType, impersonate, cachedFolderId, ct);
                driveUploaded = true;
                driveLink     = result.WebViewLink;
                if (!string.IsNullOrWhiteSpace(result.FolderId) && result.FolderId != cachedFolderId)
                    await SaveSettingAsync("backup_drive_folder_id", result.FolderId, ct);
            }
            catch (Exception ex)
            {
                driveError = ex.Message;
                logger.LogWarning(ex, "Backup automático: subida a Drive falló");
            }
        }

        var summary = new RunSummary(
            backup.FileName, backup.Method, backup.Bytes.Length,
            emailedTo, emailError, driveUploaded, driveError, driveLink);

        await LogHistoryAsync(summary, ct);
        return summary;
    }

    // ── Config (dbo.SystemSettings) ───────────────────────────────────────────

    public async Task<Dictionary<string, string?>> LoadSettingsAsync(CancellationToken ct)
    {
        var dict = new Dictionary<string, string?>();
        await using var conn = Conn();
        await conn.OpenAsync(ct);
        await using var cmd = conn.CreateCommand();
        cmd.CommandText = """
            SELECT SettingKey, SettingValue FROM dbo.SystemSettings
            WHERE SettingKey IN (
                'backup_auto_enabled','backup_recipient_emails',
                'backup_drive_impersonate_email','backup_drive_folder_id')
            """;
        await using var r = await cmd.ExecuteReaderAsync(ct);
        while (await r.ReadAsync(ct))
            dict[r.GetString(0)] = r.IsDBNull(1) ? null : r.GetString(1);
        return dict;
    }

    public async Task SaveSettingAsync(string key, string value, CancellationToken ct)
    {
        await using var conn = Conn();
        await conn.OpenAsync(ct);
        await using var cmd = conn.CreateCommand();
        cmd.CommandText = """
            MERGE dbo.SystemSettings AS t
            USING (SELECT @Key AS K, @Val AS V) AS s ON t.SettingKey = s.K
            WHEN MATCHED THEN UPDATE SET t.SettingValue = s.V, t.UpdatedAt = GETUTCDATE()
            WHEN NOT MATCHED THEN INSERT (SettingKey, SettingValue, UpdatedAt)
                                  VALUES (s.K, s.V, GETUTCDATE());
            """;
        cmd.Parameters.AddWithValue("@Key", key);
        cmd.Parameters.AddWithValue("@Val", value);
        await cmd.ExecuteNonQueryAsync(ct);
    }

    // ── Historial (dbo.BackupHistory) ─────────────────────────────────────────

    private async Task LogHistoryAsync(RunSummary s, CancellationToken ct)
    {
        await using var conn = Conn();
        await conn.OpenAsync(ct);
        await using var cmd = conn.CreateCommand();
        cmd.CommandText = """
            INSERT INTO dbo.BackupHistory
                (FileName, Method, SizeBytes, EmailedTo, EmailError, DriveUploaded, DriveError, DriveLink, RanAt)
            VALUES
                (@FileName, @Method, @SizeBytes, @EmailedTo, @EmailError, @DriveUploaded, @DriveError, @DriveLink, GETUTCDATE())
            """;
        cmd.Parameters.AddWithValue("@FileName",      s.FileName);
        cmd.Parameters.AddWithValue("@Method",         s.Method);
        cmd.Parameters.AddWithValue("@SizeBytes",      s.SizeBytes);
        cmd.Parameters.AddWithValue("@EmailedTo",      (object?)(s.EmailedTo.Count > 0 ? string.Join(", ", s.EmailedTo) : null) ?? DBNull.Value);
        cmd.Parameters.AddWithValue("@EmailError",     (object?)s.EmailError ?? DBNull.Value);
        cmd.Parameters.AddWithValue("@DriveUploaded",  s.DriveUploaded);
        cmd.Parameters.AddWithValue("@DriveError",     (object?)s.DriveError ?? DBNull.Value);
        cmd.Parameters.AddWithValue("@DriveLink",      (object?)s.DriveLink ?? DBNull.Value);
        await cmd.ExecuteNonQueryAsync(ct);
    }

    public async Task<List<Dictionary<string, object?>>> GetHistoryAsync(int take, CancellationToken ct)
    {
        var rows = new List<Dictionary<string, object?>>();
        await using var conn = Conn();
        await conn.OpenAsync(ct);
        await using var cmd = conn.CreateCommand();
        cmd.CommandText = """
            SELECT TOP (@Take) Id, FileName, Method, SizeBytes, EmailedTo, EmailError,
                   DriveUploaded, DriveError, DriveLink, RanAt
            FROM dbo.BackupHistory
            ORDER BY RanAt DESC
            """;
        cmd.Parameters.AddWithValue("@Take", take);
        await using var r = await cmd.ExecuteReaderAsync(ct);
        while (await r.ReadAsync(ct))
        {
            rows.Add(new Dictionary<string, object?>
            {
                ["id"]            = r.GetInt32(0),
                ["fileName"]      = r.GetString(1),
                ["method"]        = r.GetString(2),
                ["sizeBytes"]     = r.GetInt64(3),
                ["emailedTo"]     = r.IsDBNull(4) ? null : r.GetString(4),
                ["emailError"]    = r.IsDBNull(5) ? null : r.GetString(5),
                ["driveUploaded"] = r.GetBoolean(6),
                ["driveError"]    = r.IsDBNull(7) ? null : r.GetString(7),
                ["driveLink"]     = r.IsDBNull(8) ? null : r.GetString(8),
                ["ranAt"]         = r.GetDateTime(9),
            });
        }
        return rows;
    }

    private static string BuildEmailBody(string fileName, string method, long sizeBytes)
    {
        var kb = sizeBytes / 1024.0;
        var sizeStr = kb > 1024 ? $"{kb / 1024:0.0} MB" : $"{kb:0} KB";
        var methodLabel = method == "bak" ? "SQL Server nativo (.bak)" : "script SQL autocontenible (.sql)";

        return $"""
            <html><body style="font-family:Arial,sans-serif;font-size:14px;color:#333">
            <div style="max-width:600px;margin:0 auto">
              <div style="background:#1a237e;padding:20px;border-radius:8px 8px 0 0">
                <h2 style="color:white;margin:0">🗄️ Backup automático de Pandora</h2>
              </div>
              <div style="border:1px solid #ddd;padding:24px;border-radius:0 0 8px 8px">
                <p>Se generó el respaldo diario de la base de datos.</p>
                <table style="width:100%;border-collapse:collapse">
                  <tr style="border-bottom:1px solid #eee">
                    <td style="padding:8px;font-weight:bold;color:#555;width:40%">Archivo</td>
                    <td style="padding:8px">{fileName}</td>
                  </tr>
                  <tr style="border-bottom:1px solid #eee">
                    <td style="padding:8px;font-weight:bold;color:#555">Método</td>
                    <td style="padding:8px">{methodLabel}</td>
                  </tr>
                  <tr style="border-bottom:1px solid #eee">
                    <td style="padding:8px;font-weight:bold;color:#555">Tamaño</td>
                    <td style="padding:8px">{sizeStr}</td>
                  </tr>
                  <tr>
                    <td style="padding:8px;font-weight:bold;color:#555">Fecha</td>
                    <td style="padding:8px">{DateTime.Now:dd/MM/yyyy HH:mm}</td>
                  </tr>
                </table>
                <p style="margin-top:24px;font-size:12px;color:#888">
                  Este correo fue generado automáticamente por <strong>Pandora</strong>.
                </p>
              </div>
            </div>
            </body></html>
            """;
    }
}
