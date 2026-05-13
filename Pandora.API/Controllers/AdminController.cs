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
                // Escapar con corchetes: "dbo.Tabla" → "[dbo].[Tabla]"
                var safeName = string.Join(".", table.Split('.').Select(p => $"[{p}]"));
                dCmd.CommandText = $"SELECT * FROM {safeName}";
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

    // ── POST /api/admin/fix-encoding ──────────────────────────────────────────
    /// <summary>
    /// Corrige mojibake UTF-8→Latin-1 en todas las columnas NVARCHAR de la BD.
    /// Usar cuando aparezcan nombres con "Ã©" en lugar de "é", etc.
    /// </summary>
    [HttpPost("fix-encoding")]
    public async Task<IActionResult> FixEncoding(CancellationToken ct)
    {
        await using var conn = Conn();
        await conn.OpenAsync(ct);

        var tableCols = new (string Table, string[] Cols)[]
        {
            ("Employees",         ["FullName", "Email", "Phone", "Position"]),
            ("AppUsers",          ["FullName", "Email"]),
            ("InventoryTypes",    ["Name", "Description", "Department"]),
            ("InventoryItems",    ["Name", "Brand", "Model", "SerialNumber",
                                   "Department", "AssignedTo", "Accessories",
                                   "DecommissionReason"]),
            ("EquipmentTransfers",["FromDepartment", "FromPerson",
                                   "ToDepartment", "ToPerson", "Notes", "CreatedBy"]),
            ("Licencias",         ["Plataforma", "Area", "Responsable", "Notas"]),
            ("Comunicados",       ["Title", "Content", "Author"]),
            ("Notifications",     ["Title", "Message"]),
            ("Rooms",             ["Name", "Location"]),
            ("Reservations",      ["Title", "Description", "CreatedBy"]),
            ("RoomRequests",      ["Title", "Description", "RequestedBy"]),
            ("Procedimientos",    ["Title", "Description", "Category", "UploadedBy"]),
            ("ProcedimientoCategorias", ["Name"]),
            ("Tickets",           ["Title", "Description", "RequestedBy",
                                   "AssignedTo", "Department", "Category"]),
            ("Indicadores",       ["Nombre", "Descripcion", "Unidad",
                                   "Responsable", "Area"]),
            ("Departments",       ["Name"]),
        };

        var fixes = new (string SqlBad, string SqlGood)[]
        {
            ("NCHAR(0xC3)+NCHAR(0xA9)",  "NCHAR(0xE9)"),  // é
            ("NCHAR(0xC3)+NCHAR(0xB3)",  "NCHAR(0xF3)"),  // ó
            ("NCHAR(0xC3)+NCHAR(0xA1)",  "NCHAR(0xE1)"),  // á
            ("NCHAR(0xC3)+NCHAR(0xB1)",  "NCHAR(0xF1)"),  // ñ
            ("NCHAR(0xC3)+NCHAR(0xBA)",  "NCHAR(0xFA)"),  // ú
            ("NCHAR(0xC3)+NCHAR(0xBC)",  "NCHAR(0xFC)"),  // ü
            ("NCHAR(0xC3)+NCHAR(0xAD)",  "NCHAR(0xED)"),  // í
            ("NCHAR(0xC3)+NCHAR(0x2030)","NCHAR(0xC9)"),  // É
            ("NCHAR(0xC3)+NCHAR(0x201C)","NCHAR(0xD3)"),  // Ó
            ("NCHAR(0xC3)+NCHAR(0x0161)","NCHAR(0xDA)"),  // Ú
            ("NCHAR(0xC3)+NCHAR(0x2018)","NCHAR(0xD1)"),  // Ñ
            ("NCHAR(0xC3)+NCHAR(0x0081)","NCHAR(0xC1)"),  // Á
            ("NCHAR(0xC3)+NCHAR(0x008D)","NCHAR(0xCD)"),  // Í
            ("NCHAR(0xC2)+NCHAR(0xBF)",  "NCHAR(0xBF)"),  // ¿
            ("NCHAR(0xC2)+NCHAR(0xA1)",  "NCHAR(0xA1)"),  // ¡
        };

        string FixExpr(string col)
        {
            string expr = $"[{col}]";
            foreach (var (sqlBad, sqlGood) in fixes)
                expr = $"REPLACE({expr}, {sqlBad}, {sqlGood})";
            return expr;
        }

        var results = new List<object>();

        foreach (var (table, cols) in tableCols)
        {
            try
            {
                await using var cmd = conn.CreateCommand();
                var whereClause = string.Join(" OR ", cols.Select(c =>
                    $"([{c}] LIKE N'%' + NCHAR(0xC3) + N'%' OR [{c}] LIKE N'%' + NCHAR(0xC2) + N'%')"));
                var setClause = string.Join(", ", cols.Select(c =>
                    $"[{c}] = {FixExpr(c)}"));

                cmd.CommandText = $"""
                    IF OBJECT_ID('dbo.{table}') IS NOT NULL
                    BEGIN
                        UPDATE dbo.[{table}]
                        SET {setClause}
                        WHERE {whereClause};
                        SELECT @@ROWCOUNT;
                    END
                    ELSE SELECT 0;
                    """;

                var rows = (int)(await cmd.ExecuteScalarAsync(ct) ?? 0);
                results.Add(new { table, rowsFixed = rows });
            }
            catch (Exception ex)
            {
                results.Add(new { table, error = ex.Message });
            }
        }

        logger.LogInformation("fix-encoding ejecutado por {User}", User.Identity?.Name);
        return Ok(new { message = "Corrección completada.", details = results });
    }
}
