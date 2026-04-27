using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using System.Text;

namespace Pandora.API.Controllers;

[ApiController]
[Route("api/admin")]
[Authorize(Roles = "Admin")]
public class AdminController(IConfiguration config, ILogger<AdminController> logger) : ControllerBase
{
    private SqlConnection Conn() => new(config.GetConnectionString("PandoraDb"));

    [HttpGet("backup/download")]
    public async Task<IActionResult> DownloadBackup(CancellationToken ct)
    {
        try
        {
            var timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            var sb = new StringBuilder();

            sb.AppendLine($"-- ============================================================");
            sb.AppendLine($"-- Pandora DB Backup");
            sb.AppendLine($"-- Generado: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
            sb.AppendLine($"-- ============================================================");
            sb.AppendLine();
            sb.AppendLine("SET NOCOUNT ON;");
            sb.AppendLine("SET IDENTITY_INSERT ON;");
            sb.AppendLine("BEGIN TRANSACTION;");
            sb.AppendLine();

            await using var conn = Conn();
            await conn.OpenAsync(ct);

            // Obtener todas las tablas de usuario en orden
            var tables = new List<string>();
            await using var tCmd = conn.CreateCommand();
            tCmd.CommandText = """
                SELECT TABLE_SCHEMA + '.' + TABLE_NAME
                FROM INFORMATION_SCHEMA.TABLES
                WHERE TABLE_TYPE = 'BASE TABLE'
                ORDER BY TABLE_NAME
                """;
            await using var tReader = await tCmd.ExecuteReaderAsync(ct);
            while (await tReader.ReadAsync(ct))
                tables.Add(tReader.GetString(0));
            await tReader.CloseAsync();

            foreach (var table in tables)
            {
                try
                {
                    await using var dCmd = conn.CreateCommand();
                    dCmd.CommandText = $"SELECT * FROM {table}";
                    await using var r = await dCmd.ExecuteReaderAsync(ct);

                    var columnNames = Enumerable.Range(0, r.FieldCount)
                        .Select(i => $"[{r.GetName(i)}]")
                        .ToList();

                    bool hasRows = false;
                    while (await r.ReadAsync(ct))
                    {
                        if (!hasRows)
                        {
                            sb.AppendLine($"-- ── {table} ──────────────────────────────────────────");
                            hasRows = true;
                        }

                        var values = Enumerable.Range(0, r.FieldCount).Select(i =>
                        {
                            if (r.IsDBNull(i)) return "NULL";
                            var v = r.GetValue(i);
                            return v switch
                            {
                                string s   => $"N'{s.Replace("'", "''")}'",
                                DateTime d => $"'{d:yyyy-MM-ddTHH:mm:ss.fff}'",
                                DateTimeOffset dto => $"'{dto:yyyy-MM-ddTHH:mm:ss.fffzzz}'",
                                bool b     => b ? "1" : "0",
                                Guid g     => $"'{g}'",
                                byte[] bytes => $"0x{Convert.ToHexString(bytes)}",
                                _          => v.ToString() ?? "NULL",
                            };
                        }).ToList();

                        sb.AppendLine(
                            $"INSERT INTO {table} ({string.Join(", ", columnNames)}) " +
                            $"VALUES ({string.Join(", ", values)});");
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

            logger.LogInformation("Respaldo generado por {User} a las {Time}", User.Identity?.Name, DateTime.Now);

            var bytes = Encoding.UTF8.GetPreamble()
                .Concat(Encoding.UTF8.GetBytes(sb.ToString()))
                .ToArray();

            return File(bytes, "application/octet-stream", $"PandoraDB_{timestamp}.sql");
        }
        catch (Exception ex)
        {
            logger.LogError(ex, "Error generando respaldo");
            return StatusCode(500, "Error al generar el respaldo.");
        }
    }
}
