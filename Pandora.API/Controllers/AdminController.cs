using System.Security.Cryptography;
using System.Text;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using Pandora.API.Services;

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

    // ═══════════════════════════════════════════════════════════════════════════
    // CONFIGURACIÓN SMTP (#formulario UI)
    // ═══════════════════════════════════════════════════════════════════════════

    // ── GET /api/admin/settings/smtp ─────────────────────────────────────────
    [HttpGet("settings/smtp")]
    public async Task<IActionResult> GetSmtpSettings(CancellationToken ct)
    {
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = """
                SELECT SettingKey, SettingValue
                FROM   dbo.SystemSettings
                WHERE  SettingKey IN (
                    'smtp_host','smtp_port','smtp_from_email',
                    'smtp_from_name','smtp_use_ssl','smtp_notifications_email'
                )
                """;
            var dict = new Dictionary<string, string?>();
            await using var r = await cmd.ExecuteReaderAsync(ct);
            while (await r.ReadAsync(ct))
                dict[r.GetString(0)] = r.IsDBNull(1) ? null : r.GetString(1);

            return Ok(new {
                host              = dict.GetValueOrDefault("smtp_host",               "smtp.gmail.com"),
                port              = int.TryParse(dict.GetValueOrDefault("smtp_port"),  out var p) ? p : 587,
                fromEmail         = dict.GetValueOrDefault("smtp_from_email",          ""),
                fromName          = dict.GetValueOrDefault("smtp_from_name",           "Pandora"),
                useSsl            = dict.GetValueOrDefault("smtp_use_ssl",             "true") == "true",
                notificationsEmail= dict.GetValueOrDefault("smtp_notifications_email", ""),
                // Nunca devolver la contraseña — solo indicar si está guardada
                hasPassword       = await HasSmtpPassword(conn, ct),
            });
        }
        catch (Exception ex)
        {
            logger.LogError(ex, "GetSmtpSettings");
            return StatusCode(500, "Error al leer configuración SMTP.");
        }
    }

    // ── POST /api/admin/settings/smtp ────────────────────────────────────────
    [HttpPost("settings/smtp")]
    public async Task<IActionResult> SaveSmtpSettings([FromBody] SmtpSettingsDto dto, CancellationToken ct)
    {
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);

            // Guardar campos no sensibles
            var fields = new Dictionary<string, string>
            {
                ["smtp_host"]               = dto.Host?.Trim() ?? "",
                ["smtp_port"]               = dto.Port.ToString(),
                ["smtp_from_email"]         = dto.FromEmail?.Trim() ?? "",
                ["smtp_from_name"]          = dto.FromName?.Trim() ?? "Pandora",
                ["smtp_use_ssl"]            = dto.UseSsl ? "true" : "false",
                ["smtp_notifications_email"]= dto.NotificationsEmail?.Trim() ?? "",
            };

            foreach (var (key, value) in fields)
                await UpsertSetting(conn, key, value, ct);

            // Contraseña: solo actualizar si viene en el request (no vacía)
            if (!string.IsNullOrWhiteSpace(dto.Password))
            {
                var encrypted = EncryptSetting(dto.Password);
                await UpsertSetting(conn, "smtp_password", encrypted, ct);
            }

            logger.LogInformation("Configuración SMTP actualizada por {User}", User.Identity?.Name);
            return Ok(new { message = "Configuración SMTP guardada." });
        }
        catch (Exception ex)
        {
            logger.LogError(ex, "SaveSmtpSettings");
            return StatusCode(500, "Error al guardar configuración SMTP.");
        }
    }

    // ── POST /api/admin/settings/smtp/test ───────────────────────────────────
    [HttpPost("settings/smtp/test")]
    public async Task<IActionResult> TestSmtp([FromBody] SmtpTestDto dto, CancellationToken ct)
    {
        try
        {
            var connStr = config.GetConnectionString("PandoraDb")!;
            var smtpCfg = await SmtpHelper.LoadAsync(connStr, config);

            if (string.IsNullOrWhiteSpace(smtpCfg.Host) || string.IsNullOrWhiteSpace(smtpCfg.From) || string.IsNullOrWhiteSpace(smtpCfg.Password))
                return BadRequest(new { error = "Configura y guarda primero el SMTP antes de probar." });

            var toEmail = !string.IsNullOrWhiteSpace(dto.TestEmail) ? dto.TestEmail : smtpCfg.From;
            var htmlBody = $"""
                <div style="font-family:Arial,sans-serif;padding:24px">
                  <h2 style="color:#1a237e">✅ Prueba de correo exitosa</h2>
                  <p>El servidor SMTP está configurado correctamente en <strong>Pandora</strong>.</p>
                  <p style="color:#888;font-size:12px">Enviado el {DateTime.Now:dd/MM/yyyy HH:mm}</p>
                </div>
                """;

            var err = await SmtpHelper.SendAsync(smtpCfg, toEmail, toEmail, "✅ Prueba SMTP — Pandora", htmlBody);
            if (err != null)
            {
                logger.LogWarning("TestSmtp falló: {Error}", err);
                return BadRequest(new { error = err });
            }

            logger.LogInformation("Correo de prueba SMTP enviado a {Email} por {User}", toEmail, User.Identity?.Name);
            return Ok(new { message = $"Correo de prueba enviado a {toEmail}" });
        }
        catch (Exception ex)
        {
            logger.LogWarning(ex, "TestSmtp falló");
            return BadRequest(new { error = ex.Message });
        }
    }

    // ── Helpers ──────────────────────────────────────────────────────────────

    private static readonly byte[] _encKey =
        SHA256.HashData(Encoding.UTF8.GetBytes("Pandora_SMTP_Enc_Key_2024!"));

    private static string EncryptSetting(string plain)
    {
        using var aes = Aes.Create();
        aes.Key = _encKey;
        aes.GenerateIV();
        using var enc = aes.CreateEncryptor();
        var data = Encoding.UTF8.GetBytes(plain);
        var cipher = enc.TransformFinalBlock(data, 0, data.Length);
        var result = new byte[aes.IV.Length + cipher.Length];
        Buffer.BlockCopy(aes.IV, 0, result, 0, aes.IV.Length);
        Buffer.BlockCopy(cipher, 0, result, aes.IV.Length, cipher.Length);
        return Convert.ToBase64String(result);
    }

    private static string? DecryptSetting(string? encrypted)
    {
        if (string.IsNullOrWhiteSpace(encrypted)) return null;
        try
        {
            var raw = Convert.FromBase64String(encrypted);
            using var aes = Aes.Create();
            aes.Key = _encKey;
            var iv = new byte[16];
            Buffer.BlockCopy(raw, 0, iv, 0, 16);
            aes.IV = iv;
            using var dec = aes.CreateDecryptor();
            var cipher = new byte[raw.Length - 16];
            Buffer.BlockCopy(raw, 16, cipher, 0, cipher.Length);
            return Encoding.UTF8.GetString(dec.TransformFinalBlock(cipher, 0, cipher.Length));
        }
        catch { return null; }
    }

    private async Task UpsertSetting(SqlConnection conn, string key, string value, CancellationToken ct)
    {
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

    private async Task<bool> HasSmtpPassword(SqlConnection conn, CancellationToken ct)
    {
        await using var cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT COUNT(1) FROM dbo.SystemSettings WHERE SettingKey='smtp_password' AND SettingValue IS NOT NULL AND SettingValue <> ''";
        return (int)(await cmd.ExecuteScalarAsync(ct) ?? 0) > 0;
    }

    private async Task<(string host, int port, string from, string pass, string fromName)> LoadSmtpFromDb(CancellationToken ct)
    {
        await using var conn = Conn();
        await conn.OpenAsync(ct);
        await using var cmd = conn.CreateCommand();
        cmd.CommandText = """
            SELECT SettingKey, SettingValue FROM dbo.SystemSettings
            WHERE SettingKey IN ('smtp_host','smtp_port','smtp_from_email','smtp_password','smtp_from_name')
            """;
        var d = new Dictionary<string, string?>();
        await using var r = await cmd.ExecuteReaderAsync(ct);
        while (await r.ReadAsync(ct))
            d[r.GetString(0)] = r.IsDBNull(1) ? null : r.GetString(1);

        // Fallback a appsettings si la BD no tiene configuración
        var smtp     = config.GetSection("SmtpSettings");
        var host     = d.GetValueOrDefault("smtp_host") ?? smtp["Host"] ?? "";
        var port     = int.TryParse(d.GetValueOrDefault("smtp_port") ?? smtp["Port"], out var p) ? p : 587;
        var from     = d.GetValueOrDefault("smtp_from_email") ?? smtp["FromEmail"] ?? smtp["Username"] ?? "";
        var encPass  = d.GetValueOrDefault("smtp_password");
        var pass     = (encPass != null ? DecryptSetting(encPass) : null) ?? smtp["Password"] ?? "";
        var fromName = d.GetValueOrDefault("smtp_from_name") ?? smtp["FromName"] ?? "Pandora";
        return (host, port, from, pass, fromName);
    }
}

// ── DTOs SMTP ─────────────────────────────────────────────────────────────────
public record SmtpSettingsDto(
    string? Host,
    int     Port,
    string? FromEmail,
    string? Password,
    string? FromName,
    bool    UseSsl,
    string? NotificationsEmail
);

public record SmtpTestDto(string? TestEmail);
