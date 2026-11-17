using Microsoft.Data.SqlClient;

namespace Pandora.API.Services;

/// <summary>
/// Restaura datos a partir de un backup subido por el Admin. Dos modos, muy
/// distintos en riesgo:
///   - RestoreMissingAsync: ejecuta un .sql (generado por BackupService) tal
///     cual — como cada INSERT va envuelto en "IF NOT EXISTS", solo repone
///     filas que ya no estén, sin tocar ni borrar nada existente. Seguro.
///   - RestoreFullAsync: RESTORE DATABASE completo desde un .bak — reemplaza
///     TODA la base de datos actual por la del archivo. Destructivo e
///     irreversible; tumba las conexiones activas mientras corre.
/// </summary>
public class BackupRestoreService(
    IConfiguration config,
    BackupEncryptionService encryption,
    ILogger<BackupRestoreService> logger)
{
    private SqlConnection Conn() => new(config.GetConnectionString("PandoraDb"));

    public record RestoreResult(bool Success, string Message);

    // Quita el BOM UTF-8 inicial que BackupService antepone al .sql.
    private static string StripBom(string s) =>
        s.Length > 0 && s[0] == '﻿' ? s[1..] : s;

    // ── Modo seguro: repone solo registros faltantes desde un .sql ───────────
    public async Task<RestoreResult> RestoreMissingAsync(Stream sqlFileStream, CancellationToken ct)
    {
        byte[] raw;
        using (var ms = new MemoryStream())
        {
            await sqlFileStream.CopyToAsync(ms, ct);
            raw = ms.ToArray();
        }

        // Descifra si viene cifrado (AES-GCM); passthrough si es texto plano.
        byte[] plain;
        try { plain = encryption.Decrypt(raw); }
        catch (InvalidOperationException ex) { return new RestoreResult(false, ex.Message); }

        var script = StripBom(System.Text.Encoding.UTF8.GetString(plain));

        if (string.IsNullOrWhiteSpace(script))
            return new RestoreResult(false, "El archivo está vacío.");

        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await using var cmd = conn.CreateCommand();
            cmd.CommandTimeout = 300;
            cmd.CommandText = script;
            await cmd.ExecuteNonQueryAsync(ct);

            logger.LogInformation("Restore (registros faltantes) ejecutado correctamente.");
            return new RestoreResult(true, "Registros faltantes repuestos correctamente. No se tocó ningún registro existente.");
        }
        catch (Exception ex)
        {
            logger.LogError(ex,
                "RestoreMissingAsync falló — si el .sql viene de una versión anterior sin guardas " +
                "'IF NOT EXISTS', puede fallar por llaves duplicadas.");
            return new RestoreResult(false,
                $"Error al ejecutar el script: {ex.Message} " +
                "(si el backup es de una versión anterior sin protección anti-duplicados, prueba con uno más reciente).");
        }
    }

    // ── Modo destructivo: reemplazo total de la BD desde un .bak ─────────────
    public async Task<RestoreResult> RestoreFullAsync(Stream bakFileStream, IWebHostEnvironment env, CancellationToken ct)
    {
        var tempDir = Path.Combine(env.ContentRootPath, "backups");
        Directory.CreateDirectory(tempDir);
        var tempFile = Path.Combine(tempDir, $"restore_{DateTime.Now:yyyyMMdd_HHmmss}.bak");

        byte[] bakBytes;
        using (var ms = new MemoryStream())
        {
            await bakFileStream.CopyToAsync(ms, ct);
            bakBytes = ms.ToArray();
        }

        // Descifra si viene cifrado (.bak.enc); passthrough si es .bak en claro.
        try { bakBytes = encryption.Decrypt(bakBytes); }
        catch (InvalidOperationException ex) { return new RestoreResult(false, ex.Message); }

        await System.IO.File.WriteAllBytesAsync(tempFile, bakBytes, ct);

        var sqlPath = tempFile.Replace("\\", "/");

        // Conectar a master, no a la BD que se va a reemplazar — no se puede
        // restaurar sobre una conexión abierta hacia esa misma base.
        var masterConnStr = new SqlConnectionStringBuilder(config.GetConnectionString("PandoraDb"))
        {
            InitialCatalog = "master"
        }.ConnectionString;

        var database = new SqlConnectionStringBuilder(config.GetConnectionString("PandoraDb")).InitialCatalog;

        try
        {
            await using var conn = new SqlConnection(masterConnStr);
            await conn.OpenAsync(ct);

            await using (var cmd = conn.CreateCommand())
            {
                cmd.CommandTimeout = 300;
                cmd.CommandText = $"""
                    ALTER DATABASE [{database}] SET SINGLE_USER WITH ROLLBACK IMMEDIATE;
                    RESTORE DATABASE [{database}] FROM DISK = N'{sqlPath.Replace("'", "''")}' WITH REPLACE, STATS = 10;
                    ALTER DATABASE [{database}] SET MULTI_USER;
                    """;
                await cmd.ExecuteNonQueryAsync(ct);
            }

            // Las conexiones pooled hacia la BD vieja quedan inválidas tras el REPLACE.
            SqlConnection.ClearAllPools();

            logger.LogWarning("RESTORE DATABASE completo ejecutado sobre {Db} desde archivo subido.", database);
            return new RestoreResult(true,
                "Base de datos reemplazada por completo. Es posible que tengas que volver a iniciar sesión.");
        }
        catch (Exception ex)
        {
            logger.LogError(ex, "RestoreFullAsync falló sobre {Db}", database);

            // Intentar regresar la BD a multi-user si quedó a medias.
            try
            {
                await using var fixConn = new SqlConnection(masterConnStr);
                await fixConn.OpenAsync(CancellationToken.None);
                await using var fixCmd = fixConn.CreateCommand();
                fixCmd.CommandText = $"ALTER DATABASE [{database}] SET MULTI_USER;";
                await fixCmd.ExecuteNonQueryAsync(CancellationToken.None);
            }
            catch { /* best-effort */ }

            return new RestoreResult(false, $"Error al restaurar: {ex.Message}");
        }
        finally
        {
            try { System.IO.File.Delete(tempFile); } catch { /* ignorar */ }
        }
    }
}
