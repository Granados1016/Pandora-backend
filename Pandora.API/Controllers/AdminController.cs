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
    BackupService backupService,
    BackupOrchestrator backupOrchestrator,
    BackupRestoreService backupRestoreService,
    GoogleDriveBackupUploader driveUploader,
    ILogger<AdminController> logger) : ControllerBase
{
    private SqlConnection Conn() => new(config.GetConnectionString("PandoraDb"));

    // ── GET /api/admin/backup/download ────────────────────────────────────────
    /// <summary>
    /// Sin parámetro: intenta un BACKUP DATABASE nativo (.bak) y cae a .sql si el
    /// motor no lo soporta (LocalDB, permisos, etc.). Con ?format=sql o ?format=bak
    /// fuerza ese formato específico — bak falla con 400 si el motor no lo soporta.
    /// </summary>
    [HttpGet("backup/download")]
    public async Task<IActionResult> DownloadBackup([FromQuery] string? format, CancellationToken ct)
    {
        try
        {
            var result = format?.ToLowerInvariant() switch
            {
                "sql" => await backupService.GenerateSqlAsync(ct),
                "bak" => await backupService.GenerateBakAsync(ct),
                _     => await backupService.GenerateAsync(ct),
            };

            logger.LogInformation("Backup ({Method}) descargado manualmente por {User}: {File} ({Kb} KB)",
                result.Method, User.Identity?.Name, result.FileName, result.Bytes.Length / 1024);
            return File(result.Bytes, "application/octet-stream", result.FileName);
        }
        catch (Exception ex) when (format?.ToLowerInvariant() == "bak")
        {
            logger.LogWarning(ex, "Descarga forzada de .bak falló");
            return BadRequest(new { error = "El motor de base de datos no soporta BACKUP DATABASE nativo (.bak) en este entorno." });
        }
    }

    // ═══════════════════════════════════════════════════════════════════════════
    // BACKUP AUTOMÁTICO (correo + Google Drive)
    // ═══════════════════════════════════════════════════════════════════════════

    // ── GET /api/admin/backup/settings ───────────────────────────────────────
    [HttpGet("backup/settings")]
    public async Task<IActionResult> GetBackupSettings(CancellationToken ct)
    {
        var s = await backupOrchestrator.LoadSettingsAsync(ct);
        return Ok(new
        {
            enabled            = s.GetValueOrDefault("backup_auto_enabled") == "true",
            recipientEmails    = s.GetValueOrDefault("backup_recipient_emails") ?? "",
            driveImpersonateEmail = s.GetValueOrDefault("backup_drive_impersonate_email") ?? "",
            driveConfigured    = driveUploader.IsConfigured,
        });
    }

    // ── POST /api/admin/backup/settings ──────────────────────────────────────
    [HttpPost("backup/settings")]
    public async Task<IActionResult> SaveBackupSettings([FromBody] BackupSettingsDto dto, CancellationToken ct)
    {
        await backupOrchestrator.SaveSettingAsync("backup_auto_enabled", dto.Enabled ? "true" : "false", ct);
        await backupOrchestrator.SaveSettingAsync("backup_recipient_emails", dto.RecipientEmails?.Trim() ?? "", ct);
        await backupOrchestrator.SaveSettingAsync("backup_drive_impersonate_email", dto.DriveImpersonateEmail?.Trim() ?? "", ct);

        logger.LogInformation("Configuración de backup automático actualizada por {User}", User.Identity?.Name);
        return Ok(new { message = "Configuración de backup automático guardada." });
    }

    // ── POST /api/admin/backup/run-now ───────────────────────────────────────
    /// <summary>Ejecuta un ciclo completo (generar + repartir) inmediatamente, sin esperar al cron diario.</summary>
    [HttpPost("backup/run-now")]
    public async Task<IActionResult> RunBackupNow(CancellationToken ct)
    {
        try
        {
            var summary = await backupOrchestrator.RunAsync(ct);
            logger.LogInformation("Backup ejecutado manualmente ('Ejecutar ahora') por {User}: {File}",
                User.Identity?.Name, summary.FileName);
            return Ok(new
            {
                fileName      = summary.FileName,
                method        = summary.Method,
                sizeBytes     = summary.SizeBytes,
                emailedTo     = summary.EmailedTo,
                emailError    = summary.EmailError,
                driveUploaded = summary.DriveUploaded,
                driveError    = summary.DriveError,
                driveLink     = summary.DriveLink,
            });
        }
        catch (Exception ex)
        {
            logger.LogError(ex, "RunBackupNow falló");
            return StatusCode(500, new { error = ex.Message });
        }
    }

    // ── GET /api/admin/backup/history ────────────────────────────────────────
    [HttpGet("backup/history")]
    public async Task<IActionResult> GetBackupHistory(CancellationToken ct)
    {
        var rows = await backupOrchestrator.GetHistoryAsync(20, ct);
        return Ok(rows);
    }

    // ═══════════════════════════════════════════════════════════════════════════
    // RESTAURAR (recuperar registros faltantes / reemplazo total)
    // ═══════════════════════════════════════════════════════════════════════════

    // ── POST /api/admin/backup/restore-missing ───────────────────────────────
    /// <summary>
    /// Modo seguro: sube un .sql (generado por Pandora) y solo repone las filas
    /// que ya no existan. No borra ni modifica nada existente.
    /// </summary>
    [HttpPost("backup/restore-missing")]
    [RequestSizeLimit(200 * 1024 * 1024)] // 200 MB
    public async Task<IActionResult> RestoreMissing(IFormFile file, CancellationToken ct)
    {
        if (file is null || file.Length == 0) return BadRequest(new { error = "Archivo .sql requerido." });
        if (!file.FileName.EndsWith(".sql", StringComparison.OrdinalIgnoreCase))
            return BadRequest(new { error = "Se esperaba un archivo .sql (el generado por 'Descargar Backup')." });

        await using var stream = file.OpenReadStream();
        var result = await backupRestoreService.RestoreMissingAsync(stream, ct);

        logger.LogInformation("RestoreMissing ejecutado por {User}: {Ok} — {Msg}",
            User.Identity?.Name, result.Success, result.Message);

        return result.Success ? Ok(new { message = result.Message }) : BadRequest(new { error = result.Message });
    }

    // ── POST /api/admin/backup/restore-full ──────────────────────────────────
    /// <summary>
    /// Modo destructivo: sube un .bak y reemplaza TODA la base de datos actual.
    /// Irreversible — cualquier dato creado después de ese backup se pierde.
    /// </summary>
    [HttpPost("backup/restore-full")]
    [RequestSizeLimit(500 * 1024 * 1024)] // 500 MB
    public async Task<IActionResult> RestoreFull(IFormFile file, CancellationToken ct)
    {
        if (file is null || file.Length == 0) return BadRequest(new { error = "Archivo .bak requerido." });
        if (!file.FileName.EndsWith(".bak", StringComparison.OrdinalIgnoreCase))
            return BadRequest(new { error = "Se esperaba un archivo .bak (backup nativo de SQL Server)." });

        await using var stream = file.OpenReadStream();
        var result = await backupRestoreService.RestoreFullAsync(stream, env, ct);

        logger.LogWarning("RestoreFull (REEMPLAZO TOTAL) ejecutado por {User}: {Ok} — {Msg}",
            User.Identity?.Name, result.Success, result.Message);

        return result.Success ? Ok(new { message = result.Message }) : BadRequest(new { error = result.Message });
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

            // Leer configuración en bloque explícito para garantizar que cmd y reader
            // queden completamente dispuestos antes de reutilizar la conexión.
            var dict = new Dictionary<string, string?>();
            await using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = """
                    SELECT SettingKey, SettingValue
                    FROM   dbo.SystemSettings
                    WHERE  SettingKey IN (
                        'smtp_host','smtp_port','smtp_from_email',
                        'smtp_from_name','smtp_use_ssl','smtp_notifications_email'
                    )
                    """;
                await using var r = await cmd.ExecuteReaderAsync(ct);
                while (await r.ReadAsync(ct))
                    dict[r.GetString(0)] = r.IsDBNull(1) ? null : r.GetString(1);
            } // cmd y reader dispuestos aquí

            var hasPassword = await HasSmtpPassword(conn, ct);

            return Ok(new {
                host              = dict.GetValueOrDefault("smtp_host",               "smtp.gmail.com"),
                port              = int.TryParse(dict.GetValueOrDefault("smtp_port"),  out var p) ? p : 587,
                fromEmail         = dict.GetValueOrDefault("smtp_from_email",          ""),
                fromName          = dict.GetValueOrDefault("smtp_from_name",           "Pandora"),
                useSsl            = dict.GetValueOrDefault("smtp_use_ssl",             "true") == "true",
                notificationsEmail= dict.GetValueOrDefault("smtp_notifications_email", ""),
                hasPassword,
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

// ── DTO Backup automático ────────────────────────────────────────────────────
public record BackupSettingsDto(bool Enabled, string? RecipientEmails, string? DriveImpersonateEmail);
