using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using System.Text;

namespace Pandora.API.Controllers;

[ApiController]
[Route("api/admin")]
[Authorize(Roles = "Admin")]
public class AdminController(
    IConfiguration config,
    IWebHostEnvironment env,
    ILogger<AdminController> logger) : ControllerBase
{
    private SqlConnection Conn() => new(config.GetConnectionString("PandoraDb"));

    // ── GET /api/admin/backup/download ────────────────────────────────────────
    /// <summary>
    /// Intenta un BACKUP DATABASE nativo (.bak).
    /// Si el motor no lo soporta (LocalDB, permisos, etc.) genera un .sql ejecutable.
    /// </summary>
    [HttpGet("backup/download")]
    [RequestTimeout(300_000)] // 5 min
    public async Task<IActionResult> DownloadBackup(CancellationToken ct)
    {
        var timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");

        await using var conn = Conn();
        await conn.OpenAsync(ct);

        // ── Obtener nombre de la base de datos ────────────────────────────────
        string database;
        await using (var cmd = conn.CreateCommand())
        {
            cmd.CommandText = "SELECT DB_NAME()";
            database = (string)(await cmd.ExecuteScalarAsync(ct) ?? "PandoraDB");
        }

        // ── Intentar BACKUP DATABASE nativo (.bak) ────────────────────────────
        // El archivo se escribe en el directorio de trabajo del servidor SQL;
        // en Docker (same-container) coincide con el del proceso .NET.
        var backupDir  = Path.Combine(env.ContentRootPath, "backups");
        Directory.CreateDirectory(backupDir);
        var backupFile = Path.Combine(backupDir, $"PandoraDB_{timestamp}.bak");
        // Normalizar separadores para SQL Server en Linux/Windows
        var sqlPath = backupFile.Replace("\\", "/");

        try
        {
            await using var cmd = conn.CreateCommand();
            cmd.CommandTimeout = 300;
            cmd.CommandText = $"""
                BACKUP DATABASE [{database}]
                TO DISK = N'{sqlPath.Replace("'", "''")}'
                WITH FORMAT, INIT,
                     NAME = N'Pandora Full Backup {timestamp}',
                     SKIP, NOREWIND, NOUNLOAD, STATS = 10
                """;

            await cmd.ExecuteNonQueryAsync(ct);

            if (!System.IO.File.Exists(backupFile))
                throw new FileNotFoundException("SQL Server no generó el archivo .bak en la ruta esperada.", backupFile);

            // Leer y borrar el archivo temporal
            var bytes    = await System.IO.File.ReadAllBytesAsync(backupFile, ct);
            System.IO.File.Delete(backupFile);

            logger.LogInformation("Backup .bak generado por {User}: {File} ({Kb} KB)",
                User.Identity?.Name, $"PandoraDB_{timestamp}.bak", bytes.Length / 1024);

            return File(bytes, "application/octet-stream", $"PandoraDB_{timestamp}.bak");
        }
        catch (Exception ex)
        {
            // Limpiar archivo parcial si quedó
            if (System.IO.File.Exists(backupFile))
                try { System.IO.File.Delete(backupFile); } catch { /* ignorar */ }

            logger.LogWarning(ex,
                "BACKUP DATABASE no disponible ({Msg}). Generando respaldo SQL ejecutable como fallback.",
                ex.Message);

            // ── Fallback: script SQL ejecutable (.sql con INSERTs) ─────────────
            return await GenerateSqlBackupAsync(conn, database, timestamp, ct);
        }
    }

    // ── Fallback: genera un .sql con INSERTs ejecutables ─────────────────────
    private async Task<IActionResult> GenerateSqlBackupAsync(
        SqlConnection conn, string database, string timestamp, CancellationToken ct)
    {
        var sb = new StringBuilder();
        sb.AppendLine("-- ============================================================");
        sb.AppendLine($"-- Pandora DB Backup (SQL Script)");
        sb.AppendLine($"-- Base de datos: {database}");
        sb.AppendLine($"-- Generado: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
        sb.AppendLine("-- ============================================================");
        sb.AppendLine();
        sb.AppendLine("SET NOCOUNT ON;");
        sb.AppendLine("BEGIN TRANSACTION;");
        sb.AppendLine();

        // Obtener tablas
        var tables = new List<string>();
        await using (var tCmd = conn.CreateCommand())
        {
            tCmd.CommandText = """
                SELECT TABLE_SCHEMA + '.' + TABLE_NAME
                FROM INFORMATION_SCHEMA.TABLES
                WHERE TABLE_TYPE = 'BASE TABLE'
                ORDER BY TABLE_NAME
                """;
            await using var tReader = await tCmd.ExecuteReaderAsync(ct);
            while (await tReader.ReadAsync(ct))
                tables.Add(tReader.GetString(0));
        }

        foreach (var table in tables)
        {
            try
            {
                await using var dCmd = conn.CreateCommand();
                dCmd.CommandText = $"SELECT * FROM {table}";
                await using var r = await dCmd.ExecuteReaderAsync(ct);

                var cols = Enumerable.Range(0, r.FieldCount)
                    .Select(i => $"[{r.GetName(i)}]").ToList();

                bool hasRows = false;
                while (await r.ReadAsync(ct))
                {
                    if (!hasRows)
                    {
                        sb.AppendLine($"-- ── {table} ──────────────────────────────");
                        hasRows = true;
                    }

                    var vals = Enumerable.Range(0, r.FieldCount).Select(i =>
                    {
                        if (r.IsDBNull(i)) return "NULL";
                        return r.GetValue(i) switch
                        {
                            string s           => $"N'{s.Replace("'", "''")}'",
                            DateTime d         => $"'{d:yyyy-MM-ddTHH:mm:ss.fff}'",
                            DateTimeOffset dto => $"'{dto:yyyy-MM-ddTHH:mm:ss.fffzzz}'",
                            bool b             => b ? "1" : "0",
                            Guid g             => $"'{g}'",
                            byte[] arr         => $"0x{Convert.ToHexString(arr)}",
                            var v              => v?.ToString() ?? "NULL",
                        };
                    }).ToList();

                    sb.AppendLine(
                        $"INSERT INTO {table} ({string.Join(", ", cols)}) " +
                        $"VALUES ({string.Join(", ", vals)});");
                }

                if (hasRows) sb.AppendLine();
            }
            catch (Exception ex)
            {
                sb.AppendLine($"-- ADVERTENCIA: no se pudo exportar {table}: {ex.Message}");
                sb.AppendLine();
            }
        }

        sb.AppendLine("COMMIT TRANSACTION;");
        sb.AppendLine($"-- Fin del respaldo — {DateTime.Now:yyyy-MM-dd HH:mm:ss}");

        logger.LogInformation("Respaldo SQL generado por {User} a las {Time}",
            User.Identity?.Name, DateTime.Now);

        var bytes = Encoding.UTF8.GetPreamble()
            .Concat(Encoding.UTF8.GetBytes(sb.ToString()))
            .ToArray();

        return File(bytes, "application/octet-stream", $"PandoraDB_{timestamp}.sql");
    }
}
