using MailKit.Net.Smtp;
using MailKit.Security;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using MimeKit;
using System.Security.Claims;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;

namespace Pandora.API.Controllers;

[ApiController]
[Route("api/campaigns")]
[Authorize]
public class CampaignsController(
    IConfiguration config,
    ILogger<CampaignsController> logger) : ControllerBase
{
    // ── Helpers ───────────────────────────────────────────────────────────────
    private SqlConnection Conn() => new(config.GetConnectionString("PandoraDb"));

    private string? CurrentUsername =>
        User.FindFirstValue(ClaimTypes.Name) ??
        User.FindFirstValue("name") ??
        User.Claims.FirstOrDefault(c => c.Type.EndsWith("name", StringComparison.OrdinalIgnoreCase))?.Value;

    /// <summary>
    /// Si la columna Status de EmailCampaigns es INT (creada por EF Core con enum),
    /// la convierte a NVARCHAR para que funcione con los valores en español del controller.
    /// </summary>
    private static async Task EnsureEmailCampaignsStatusIsVarcharAsync(SqlConnection conn, CancellationToken ct)
    {
        const string sql = """
            IF EXISTS (
                SELECT 1 FROM sys.columns c
                JOIN sys.objects o ON c.object_id = o.object_id
                WHERE o.name = 'EmailCampaigns'
                  AND c.name = 'Status'
                  AND c.system_type_id = 56  -- INT
            )
            BEGIN
                ALTER TABLE dbo.EmailCampaigns ALTER COLUMN Status NVARCHAR(50) NULL;
                UPDATE dbo.EmailCampaigns SET Status = CASE CAST(Status AS INT)
                    WHEN 0 THEN 'Pendiente'
                    WHEN 1 THEN 'Enviando'
                    WHEN 2 THEN 'Completado'
                    WHEN 3 THEN 'Completado con errores'
                    ELSE 'Pendiente'
                END WHERE Status NOT LIKE '%[^0-9]%' OR Status IS NULL;
                ALTER TABLE dbo.EmailCampaigns ALTER COLUMN Status NVARCHAR(50) NOT NULL;
            END
            """;
        await using var cmd = conn.CreateCommand();
        cmd.CommandText = sql;
        await cmd.ExecuteNonQueryAsync(ct);
    }

    /// <summary>Crea la tabla CampaignRecipients si no existe y agrega ReadAt si falta.</summary>
    private static async Task EnsureRecipientsTableAsync(SqlConnection conn, CancellationToken ct)
    {
        const string sql = """
            IF NOT EXISTS (SELECT 1 FROM sys.objects
                WHERE object_id = OBJECT_ID(N'dbo.CampaignRecipients') AND type = N'U')
            BEGIN
                CREATE TABLE dbo.CampaignRecipients (
                    Id           UNIQUEIDENTIFIER NOT NULL DEFAULT NEWID() PRIMARY KEY,
                    CampaignId   UNIQUEIDENTIFIER NOT NULL,
                    FullName     NVARCHAR(200)    NOT NULL,
                    Email        NVARCHAR(200)    NOT NULL,
                    Username     NVARCHAR(100)    NULL,
                    Password     NVARCHAR(200)    NULL,
                    ExtraData    NVARCHAR(MAX)    NULL,
                    Status       NVARCHAR(20)     NOT NULL DEFAULT 'Pending',
                    ErrorMessage NVARCHAR(MAX)    NULL,
                    SentAt       DATETIME2        NULL,
                    ReadAt       DATETIME2        NULL,
                    CreatedAt    DATETIME2        NOT NULL DEFAULT GETUTCDATE()
                );
            END
            ELSE IF NOT EXISTS (
                SELECT 1 FROM sys.columns c
                JOIN sys.objects o ON c.object_id = o.object_id
                WHERE o.name = 'CampaignRecipients' AND c.name = 'ReadAt'
            )
            BEGIN
                ALTER TABLE dbo.CampaignRecipients ADD ReadAt DATETIME2 NULL;
            END
            """;
        await using var cmd = conn.CreateCommand();
        cmd.CommandText = sql;
        await cmd.ExecuteNonQueryAsync(ct);
    }

    private static string ProgramLabel(int t) => t switch
    {
        1 => "Licenciatura", 2 => "Posgrado", 3 => "Preparatoria", 4 => "Notificaciones", _ => "General"
    };

    private static object ReadCampaign(SqlDataReader r) => new
    {
        id              = r.GetGuid(r.GetOrdinal("Id")),
        name            = r.IsDBNull(r.GetOrdinal("Name"))    ? "" : r.GetString(r.GetOrdinal("Name")),
        subject         = r.IsDBNull(r.GetOrdinal("Subject")) ? "" : r.GetString(r.GetOrdinal("Subject")),
        body            = r.IsDBNull(r.GetOrdinal("Body"))    ? "" : r.GetString(r.GetOrdinal("Body")),
        programType     = r.GetInt32(r.GetOrdinal("ProgramType")),
        status          = r.IsDBNull(r.GetOrdinal("Status"))  ? "Pendiente" : r.GetString(r.GetOrdinal("Status")),
        totalRecipients = r.IsDBNull(r.GetOrdinal("TotalRecipients")) ? 0 : r.GetInt32(r.GetOrdinal("TotalRecipients")),
        sentCount       = r.IsDBNull(r.GetOrdinal("SentCount"))       ? 0 : r.GetInt32(r.GetOrdinal("SentCount")),
        failedCount     = r.IsDBNull(r.GetOrdinal("FailedCount"))     ? 0 : r.GetInt32(r.GetOrdinal("FailedCount")),
        isDeleted       = r.GetBoolean(r.GetOrdinal("IsDeleted")),
        createdAt       = r.GetDateTime(r.GetOrdinal("CreatedAt")),
        sentAt          = r.IsDBNull(r.GetOrdinal("SentAt"))  ? (DateTime?)null : r.GetDateTime(r.GetOrdinal("SentAt")),
        templateId      = r.IsDBNull(r.GetOrdinal("TemplateId")) ? (Guid?)null : r.GetGuid(r.GetOrdinal("TemplateId")),
    };

    // ── GET /api/campaigns ────────────────────────────────────────────────────
    [HttpGet]
    public async Task<IActionResult> GetAll(CancellationToken ct)
    {
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = """
                SELECT Id, Name, Subject, Body, ProgramType, Status,
                       TotalRecipients, SentCount, FailedCount, IsDeleted,
                       CreatedAt, SentAt, TemplateId
                FROM dbo.EmailCampaigns
                WHERE IsDeleted = 0
                ORDER BY CreatedAt DESC
                """;
            var list = new List<object>();
            await using var r = await cmd.ExecuteReaderAsync(ct);
            while (await r.ReadAsync(ct)) list.Add(ReadCampaign(r));
            return Ok(list);
        }
        catch (Exception ex) { logger.LogError(ex, "GetAll Campaigns"); return StatusCode(500, ex.Message); }
    }

    // ── GET /api/campaigns/deleted ────────────────────────────────────────────
    [HttpGet("deleted")]
    public async Task<IActionResult> GetDeleted(CancellationToken ct)
    {
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = """
                SELECT Id, Name, Subject, Body, ProgramType, Status,
                       TotalRecipients, SentCount, FailedCount, IsDeleted,
                       CreatedAt, SentAt, TemplateId
                FROM dbo.EmailCampaigns
                WHERE IsDeleted = 1
                ORDER BY CreatedAt DESC
                """;
            var list = new List<object>();
            await using var r = await cmd.ExecuteReaderAsync(ct);
            while (await r.ReadAsync(ct)) list.Add(ReadCampaign(r));
            return Ok(list);
        }
        catch (Exception ex) { logger.LogError(ex, "GetDeleted Campaigns"); return StatusCode(500, ex.Message); }
    }

    // ── GET /api/campaigns/{id} ───────────────────────────────────────────────
    [HttpGet("{id:guid}")]
    public async Task<IActionResult> GetById(Guid id, CancellationToken ct)
    {
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = """
                SELECT Id, Name, Subject, Body, ProgramType, Status,
                       TotalRecipients, SentCount, FailedCount, IsDeleted,
                       CreatedAt, SentAt, TemplateId
                FROM dbo.EmailCampaigns
                WHERE Id = @Id
                """;
            cmd.Parameters.AddWithValue("@Id", id);
            await using var r = await cmd.ExecuteReaderAsync(ct);
            if (!await r.ReadAsync(ct)) return NotFound("Campaña no encontrada.");
            return Ok(ReadCampaign(r));
        }
        catch (Exception ex) { logger.LogError(ex, "GetById Campaign {Id}", id); return StatusCode(500, ex.Message); }
    }

    // ── POST /api/campaigns ───────────────────────────────────────────────────
    [HttpPost]
    public async Task<IActionResult> Create([FromBody] CampaignCreateDto dto, CancellationToken ct)
    {
        if (string.IsNullOrWhiteSpace(dto.Name))    return BadRequest("Nombre requerido.");
        if (string.IsNullOrWhiteSpace(dto.Subject)) return BadRequest("Asunto requerido.");
        if (dto.Recipients == null || dto.Recipients.Count == 0) return BadRequest("Destinatarios requeridos.");

        var id = Guid.NewGuid();
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await EnsureEmailCampaignsStatusIsVarcharAsync(conn, ct);
            await EnsureRecipientsTableAsync(conn, ct);

            // Insertar campaña en EmailCampaigns
            await using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = """
                    INSERT INTO dbo.EmailCampaigns
                        (Id, Name, Subject, Body, ProgramType, Status,
                         TotalRecipients, SentCount, FailedCount, IsDeleted, CreatedAt)
                    VALUES
                        (@Id, @Name, @Subject, @Body, @ProgramType, 'Pendiente',
                         @Total, 0, 0, 0, GETUTCDATE())
                    """;
                cmd.Parameters.AddWithValue("@Id",          id);
                cmd.Parameters.AddWithValue("@Name",        dto.Name.Trim());
                cmd.Parameters.AddWithValue("@Subject",     dto.Subject.Trim());
                cmd.Parameters.AddWithValue("@Body",        dto.Body ?? "");
                cmd.Parameters.AddWithValue("@ProgramType", dto.ProgramType);
                cmd.Parameters.AddWithValue("@Total",       dto.Recipients.Count);
                await cmd.ExecuteNonQueryAsync(ct);
            }

            // Insertar destinatarios en CampaignRecipients
            foreach (var rec in dto.Recipients)
            {
                await using var cmd2 = conn.CreateCommand();
                cmd2.CommandText = """
                    INSERT INTO dbo.CampaignRecipients
                        (Id, CampaignId, FullName, Email, Username, Password, ExtraData, Status, CreatedAt)
                    VALUES
                        (NEWID(), @CampaignId, @FullName, @Email, @Username, @Password, @ExtraData, 'Pending', GETUTCDATE())
                    """;
                cmd2.Parameters.AddWithValue("@CampaignId", id);
                cmd2.Parameters.AddWithValue("@FullName",   rec.FullName ?? "");
                cmd2.Parameters.AddWithValue("@Email",      rec.Email    ?? "");
                cmd2.Parameters.AddWithValue("@Username",   (object?)rec.Username  ?? DBNull.Value);
                cmd2.Parameters.AddWithValue("@Password",   (object?)rec.Password  ?? DBNull.Value);
                cmd2.Parameters.AddWithValue("@ExtraData",  rec.ExtraData != null
                    ? JsonSerializer.Serialize(rec.ExtraData) : DBNull.Value);
                await cmd2.ExecuteNonQueryAsync(ct);
            }

            logger.LogInformation("Campaign created: {Id} — {Name}", id, dto.Name);
            return Ok(new { id });
        }
        catch (Exception ex) { logger.LogError(ex, "Create Campaign"); return StatusCode(500, ex.Message); }
    }

    // ── POST /api/campaigns/{id}/send ─────────────────────────────────────────
    [HttpPost("{id:guid}/send")]
    public async Task<IActionResult> Send(Guid id, CancellationToken ct)
    {
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await EnsureEmailCampaignsStatusIsVarcharAsync(conn, ct);
            await EnsureRecipientsTableAsync(conn, ct);

            // Obtener campaña
            string subject = "", body = "";
            int programType = 1;
            await using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = "SELECT Subject, Body, ProgramType FROM dbo.EmailCampaigns WHERE Id = @Id AND IsDeleted = 0";
                cmd.Parameters.AddWithValue("@Id", id);
                await using var r = await cmd.ExecuteReaderAsync(ct);
                if (!await r.ReadAsync(ct)) return NotFound("Campaña no encontrada.");
                subject     = r.GetString(0);
                body        = r.GetString(1);
                programType = r.GetInt32(2);
            }

            // Marcar como "Enviando"
            await using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = "UPDATE dbo.EmailCampaigns SET Status = 'Enviando', UpdatedAt = GETUTCDATE() WHERE Id = @Id";
                cmd.Parameters.AddWithValue("@Id", id);
                await cmd.ExecuteNonQueryAsync(ct);
            }

            // Obtener destinatarios pendientes
            var recipients = new List<(Guid Id, string FullName, string Email, string? Username, string? Password, string? ExtraData)>();
            await using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = """
                    SELECT Id, FullName, Email, Username, Password, ExtraData
                    FROM dbo.CampaignRecipients
                    WHERE CampaignId = @CampaignId AND Status = 'Pending'
                    """;
                cmd.Parameters.AddWithValue("@CampaignId", id);
                await using var r = await cmd.ExecuteReaderAsync(ct);
                while (await r.ReadAsync(ct))
                    recipients.Add((
                        r.GetGuid(0),
                        r.GetString(1),
                        r.GetString(2),
                        r.IsDBNull(3) ? null : r.GetString(3),
                        r.IsDBNull(4) ? null : r.GetString(4),
                        r.IsDBNull(5) ? null : r.GetString(5)
                    ));
            }

            // Configuración SMTP — intentar personal del usuario primero
            string? smtpEmail = null, smtpPass = null;
            var username = CurrentUsername;
            if (!string.IsNullOrWhiteSpace(username))
            {
                await using var cmd = conn.CreateCommand();
                cmd.CommandText = "SELECT SmtpEmail, SmtpPassword FROM dbo.AppUsers WHERE LOWER(Username) = LOWER(@User)";
                cmd.Parameters.AddWithValue("@User", username);
                await using var r = await cmd.ExecuteReaderAsync(ct);
                if (await r.ReadAsync(ct))
                {
                    smtpEmail = r.IsDBNull(0) ? null : r.GetString(0);
                    smtpPass  = r.IsDBNull(1) ? null : r.GetString(1);
                }
            }

            var smtpCfg     = config.GetSection("SmtpSettings");
            var smtpHost    = smtpCfg["Host"]    ?? "smtp.gmail.com";
            var smtpPort    = int.TryParse(smtpCfg["Port"], out var p) ? p : 587;
            var fromEmail   = smtpEmail ?? smtpCfg["FromEmail"] ?? "";
            var fromPass    = smtpPass  ?? smtpCfg["Password"]  ?? "";
            var fromName    = smtpCfg["FromName"] ?? "Coordinación de TI";

            // Capturar baseUrl antes de entrar al Task.Run (HttpContext no es accesible dentro)
            var baseUrl = $"{Request.Scheme}://{Request.Host}";

            // ── Prueba de conexión SMTP antes de lanzar el background task ───────
            // Si las credenciales son incorrectas, el error se muestra inmediatamente
            // en lugar de quedar silencioso en la base de datos.
            try
            {
                using var testSmtp = new SmtpClient();
                await testSmtp.ConnectAsync(smtpHost, smtpPort, SecureSocketOptions.StartTls, ct);
                await testSmtp.AuthenticateAsync(fromEmail, fromPass, ct);
                await testSmtp.DisconnectAsync(true, ct);
            }
            catch (Exception smtpEx)
            {
                // Revertir estado de campaña a Pendiente
                await using var rollback = conn.CreateCommand();
                rollback.CommandText = "UPDATE dbo.EmailCampaigns SET Status='Pendiente', UpdatedAt=GETUTCDATE() WHERE Id=@Id";
                rollback.Parameters.AddWithValue("@Id", id);
                await rollback.ExecuteNonQueryAsync(ct);
                return BadRequest($"Error de conexión SMTP: {smtpEx.Message}");
            }

            // ── Envío en background con MailKit ───────────────────────────────────
            _ = Task.Run(async () =>
            {
                int sent = 0, failed = 0;
                foreach (var rec in recipients)
                {
                    try
                    {
                        var extraVars = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                        if (!string.IsNullOrWhiteSpace(rec.ExtraData))
                        {
                            try
                            {
                                var extra = JsonSerializer.Deserialize<Dictionary<string, JsonElement>>(rec.ExtraData);
                                if (extra != null)
                                    foreach (var kv in extra)
                                        extraVars[kv.Key] = kv.Value.GetString() ?? "";
                            }
                            catch { /* ignore */ }
                        }

                        var finalSubject = ApplyVars(subject, rec.FullName, rec.Email,
                            rec.Username, rec.Password, ProgramLabel(programType), extraVars);
                        var finalBody    = ApplyVars(body, rec.FullName, rec.Email,
                            rec.Username, rec.Password, ProgramLabel(programType), extraVars);

                        // Pixel de seguimiento invisible — registra la primera apertura
                        var trackingPixel = $"<img src=\"{baseUrl}/api/campaigns/track/{rec.Id}\" width=\"1\" height=\"1\" style=\"display:none;\" alt=\"\" />";
                        var bodyWithTracking = finalBody.Contains("</body>", StringComparison.OrdinalIgnoreCase)
                            ? finalBody.Replace("</body>", $"{trackingPixel}</body>", StringComparison.OrdinalIgnoreCase)
                            : finalBody + trackingPixel;

                        var message = new MimeMessage();
                        message.From.Add(new MailboxAddress(fromName, fromEmail));
                        message.To.Add(new MailboxAddress(rec.FullName, rec.Email));
                        message.Subject = finalSubject;
                        message.Body = BuildBodyWithInlineImages(bodyWithTracking);

                        using var mailClient = new SmtpClient();
                        await mailClient.ConnectAsync(smtpHost, smtpPort, SecureSocketOptions.StartTls);
                        await mailClient.AuthenticateAsync(fromEmail, fromPass);
                        await mailClient.SendAsync(message);
                        await mailClient.DisconnectAsync(true);

                        await using var conn2 = Conn();
                        await conn2.OpenAsync();
                        await using var cmd2 = conn2.CreateCommand();
                        cmd2.CommandText = "UPDATE dbo.CampaignRecipients SET Status='Sent', SentAt=GETUTCDATE() WHERE Id=@Id";
                        cmd2.Parameters.AddWithValue("@Id", rec.Id);
                        await cmd2.ExecuteNonQueryAsync();
                        sent++;
                    }
                    catch (Exception ex)
                    {
                        failed++;
                        try
                        {
                            await using var conn2 = Conn();
                            await conn2.OpenAsync();
                            await using var cmd2 = conn2.CreateCommand();
                            cmd2.CommandText = "UPDATE dbo.CampaignRecipients SET Status='Failed', ErrorMessage=@Err WHERE Id=@Id";
                            cmd2.Parameters.AddWithValue("@Err", ex.Message);
                            cmd2.Parameters.AddWithValue("@Id",  rec.Id);
                            await cmd2.ExecuteNonQueryAsync();
                        }
                        catch { /* ignore */ }
                    }
                }

                // Actualizar contadores y estado final
                try
                {
                    await using var conn2 = Conn();
                    await conn2.OpenAsync();
                    await using var cmd2 = conn2.CreateCommand();
                    cmd2.CommandText = """
                        UPDATE dbo.EmailCampaigns
                        SET Status = CASE WHEN @Failed = 0 THEN 'Completado' ELSE 'Completado con errores' END,
                            SentCount   = @Sent,
                            FailedCount = @Failed,
                            SentAt      = GETUTCDATE(),
                            UpdatedAt   = GETUTCDATE()
                        WHERE Id = @Id
                        """;
                    cmd2.Parameters.AddWithValue("@Sent",   sent);
                    cmd2.Parameters.AddWithValue("@Failed", failed);
                    cmd2.Parameters.AddWithValue("@Id",     id);
                    await cmd2.ExecuteNonQueryAsync();
                }
                catch { /* ignore */ }
            }, CancellationToken.None);

            return Ok(new { total = recipients.Count, sent = 0, failed = 0, message = "Envío iniciado." });
        }
        catch (Exception ex) { logger.LogError(ex, "Send Campaign {Id}", id); return StatusCode(500, ex.Message); }
    }

    // ── POST /api/campaigns/{id}/retry-failed ─────────────────────────────────
    [HttpPost("{id:guid}/retry-failed")]
    public async Task<IActionResult> RetryFailed(Guid id, CancellationToken ct)
    {
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await EnsureEmailCampaignsStatusIsVarcharAsync(conn, ct);
            await EnsureRecipientsTableAsync(conn, ct);

            // Resetear los fallidos a Pending
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = """
                UPDATE dbo.CampaignRecipients
                SET Status = 'Pending', ErrorMessage = NULL, SentAt = NULL
                WHERE CampaignId = @Id AND Status = 'Failed'
                """;
            cmd.Parameters.AddWithValue("@Id", id);
            int rows = await cmd.ExecuteNonQueryAsync(ct);

            // Actualizar estado de campaña
            await using var cmd2 = conn.CreateCommand();
            cmd2.CommandText = "UPDATE dbo.EmailCampaigns SET Status='Pendiente', UpdatedAt=GETUTCDATE() WHERE Id=@Id";
            cmd2.Parameters.AddWithValue("@Id", id);
            await cmd2.ExecuteNonQueryAsync(ct);

            return Ok(new { reset = rows });
        }
        catch (Exception ex) { logger.LogError(ex, "RetryFailed Campaign {Id}", id); return StatusCode(500, ex.Message); }
    }

    // ── POST /api/campaigns/{id}/duplicate ────────────────────────────────────
    [HttpPost("{id:guid}/duplicate")]
    public async Task<IActionResult> Duplicate(Guid id, CancellationToken ct)
    {
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await EnsureEmailCampaignsStatusIsVarcharAsync(conn, ct);
            await EnsureRecipientsTableAsync(conn, ct);

            // Obtener campaña original
            string name = "", subject = "", body = "";
            int programType = 1;
            await using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = "SELECT Name, Subject, Body, ProgramType FROM dbo.EmailCampaigns WHERE Id=@Id";
                cmd.Parameters.AddWithValue("@Id", id);
                await using var r = await cmd.ExecuteReaderAsync(ct);
                if (!await r.ReadAsync(ct)) return NotFound("Campaña no encontrada.");
                name = r.GetString(0); subject = r.GetString(1);
                body = r.GetString(2); programType = r.GetInt32(3);
            }

            var newId = Guid.NewGuid();

            // Obtener destinatarios originales
            var recs = new List<(string fn, string em, string? un, string? pw, string? ex)>();
            await using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = "SELECT FullName, Email, Username, Password, ExtraData FROM dbo.CampaignRecipients WHERE CampaignId=@Id";
                cmd.Parameters.AddWithValue("@Id", id);
                await using var r = await cmd.ExecuteReaderAsync(ct);
                while (await r.ReadAsync(ct))
                    recs.Add((r.GetString(0), r.GetString(1),
                        r.IsDBNull(2)?null:r.GetString(2), r.IsDBNull(3)?null:r.GetString(3),
                        r.IsDBNull(4)?null:r.GetString(4)));
            }

            await using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = """
                    INSERT INTO dbo.EmailCampaigns
                        (Id, Name, Subject, Body, ProgramType, Status, TotalRecipients, SentCount, FailedCount, IsDeleted, CreatedAt)
                    VALUES (@Id, @Name, @Subject, @Body, @ProgramType, 'Pendiente', @Total, 0, 0, 0, GETUTCDATE())
                    """;
                cmd.Parameters.AddWithValue("@Id",  newId);
                cmd.Parameters.AddWithValue("@Name", $"(Copia) {name}");
                cmd.Parameters.AddWithValue("@Subject", subject);
                cmd.Parameters.AddWithValue("@Body", body);
                cmd.Parameters.AddWithValue("@ProgramType", programType);
                cmd.Parameters.AddWithValue("@Total", recs.Count);
                await cmd.ExecuteNonQueryAsync(ct);
            }

            foreach (var rec in recs)
            {
                await using var cmd = conn.CreateCommand();
                cmd.CommandText = """
                    INSERT INTO dbo.CampaignRecipients
                        (Id, CampaignId, FullName, Email, Username, Password, ExtraData, Status, CreatedAt)
                    VALUES (NEWID(), @CId, @FN, @Em, @Un, @Pw, @Ex, 'Pending', GETUTCDATE())
                    """;
                cmd.Parameters.AddWithValue("@CId", newId);
                cmd.Parameters.AddWithValue("@FN",  rec.fn);
                cmd.Parameters.AddWithValue("@Em",  rec.em);
                cmd.Parameters.AddWithValue("@Un",  (object?)rec.un ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@Pw",  (object?)rec.pw ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@Ex",  (object?)rec.ex ?? DBNull.Value);
                await cmd.ExecuteNonQueryAsync(ct);
            }

            return Ok(new { id = newId });
        }
        catch (Exception ex) { logger.LogError(ex, "Duplicate Campaign {Id}", id); return StatusCode(500, ex.Message); }
    }

    // ── DELETE /api/campaigns/{id}  (soft delete) ─────────────────────────────
    [HttpDelete("{id:guid}")]
    public async Task<IActionResult> Delete(Guid id, CancellationToken ct)
    {
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = """
                UPDATE dbo.EmailCampaigns
                SET IsDeleted = 1, DeletedAt = GETUTCDATE(),
                    DeletedBy = @User, UpdatedAt = GETUTCDATE()
                WHERE Id = @Id AND IsDeleted = 0
                """;
            cmd.Parameters.AddWithValue("@Id",   id);
            cmd.Parameters.AddWithValue("@User", (object?)(CurrentUsername) ?? DBNull.Value);
            int rows = await cmd.ExecuteNonQueryAsync(ct);
            if (rows == 0) return NotFound("Campaña no encontrada.");
            return NoContent();
        }
        catch (Exception ex) { logger.LogError(ex, "Delete Campaign {Id}", id); return StatusCode(500, ex.Message); }
    }

    // ── POST /api/campaigns/{id}/restore ──────────────────────────────────────
    [HttpPost("{id:guid}/restore")]
    public async Task<IActionResult> Restore(Guid id, CancellationToken ct)
    {
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = """
                UPDATE dbo.EmailCampaigns
                SET IsDeleted = 0, DeletedAt = NULL, DeletedBy = NULL, UpdatedAt = GETUTCDATE()
                WHERE Id = @Id AND IsDeleted = 1
                """;
            cmd.Parameters.AddWithValue("@Id", id);
            int rows = await cmd.ExecuteNonQueryAsync(ct);
            if (rows == 0) return NotFound("Campaña no encontrada o ya activa.");
            return Ok();
        }
        catch (Exception ex) { logger.LogError(ex, "Restore Campaign {Id}", id); return StatusCode(500, ex.Message); }
    }

    // ── GET /api/campaigns/{id}/recipients ────────────────────────────────────
    [HttpGet("{id:guid}/recipients")]
    public async Task<IActionResult> GetRecipients(Guid id, CancellationToken ct)
    {
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await EnsureEmailCampaignsStatusIsVarcharAsync(conn, ct);
            await EnsureRecipientsTableAsync(conn, ct);
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = """
                SELECT Id, FullName, Email, Username, Status, ErrorMessage, SentAt, ReadAt
                FROM dbo.CampaignRecipients
                WHERE CampaignId = @Id
                ORDER BY FullName
                """;
            cmd.Parameters.AddWithValue("@Id", id);
            var list = new List<object>();
            await using var r = await cmd.ExecuteReaderAsync(ct);
            while (await r.ReadAsync(ct))
                list.Add(new
                {
                    id           = r.GetGuid(0),
                    fullName     = r.GetString(1),
                    email        = r.GetString(2),
                    username     = r.IsDBNull(3) ? null : r.GetString(3),
                    status       = r.IsDBNull(4) ? "Pending" : r.GetString(4),
                    errorMessage = r.IsDBNull(5) ? null : r.GetString(5),
                    sentAt       = r.IsDBNull(6) ? (DateTime?)null : r.GetDateTime(6),
                    readAt       = r.IsDBNull(7) ? (DateTime?)null : r.GetDateTime(7),
                });
            return Ok(list);
        }
        catch (Exception ex) { logger.LogError(ex, "GetRecipients Campaign {Id}", id); return StatusCode(500, ex.Message); }
    }

    // ── GET /api/campaigns/track/{recipientId} ────────────────────────────────
    // Pixel de seguimiento 1x1: registra la primera apertura del correo.
    // No requiere autenticación (se carga desde el cliente de correo del destinatario).
    [HttpGet("track/{recipientId:guid}")]
    [AllowAnonymous]
    public async Task<IActionResult> Track(Guid recipientId)
    {
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync();
            await using var cmd = conn.CreateCommand();
            // Solo registra la primera apertura (no sobreescribe si ya existe)
            cmd.CommandText = """
                UPDATE dbo.CampaignRecipients
                SET ReadAt = GETUTCDATE()
                WHERE Id = @Id AND ReadAt IS NULL
                """;
            cmd.Parameters.AddWithValue("@Id", recipientId);
            await cmd.ExecuteNonQueryAsync();
        }
        catch { /* silencioso: no interrumpir la carga del correo */ }

        // GIF transparente 1x1 px
        var pixel = Convert.FromBase64String("R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7");
        return File(pixel, "image/gif");
    }

    // ── GET /api/campaigns/{id}/export ────────────────────────────────────────
    [HttpGet("{id:guid}/export")]
    [AllowAnonymous]
    public async Task<IActionResult> Export(Guid id, [FromQuery] string? access_token, CancellationToken ct)
    {
        // Validar token de acceso via query string
        if (!string.IsNullOrWhiteSpace(access_token))
        {
            var token = HttpContext.Request.Headers.Authorization.FirstOrDefault();
            if (string.IsNullOrWhiteSpace(token))
                HttpContext.Request.Headers.Append("Authorization", $"Bearer {access_token}");
        }

        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await EnsureEmailCampaignsStatusIsVarcharAsync(conn, ct);
            await EnsureRecipientsTableAsync(conn, ct);

            var sb = new StringBuilder();
            sb.AppendLine("Nombre,Email,Usuario,Estado,Error,FechaEnvio");

            await using var cmd = conn.CreateCommand();
            cmd.CommandText = """
                SELECT FullName, Email, Username, Status, ErrorMessage, SentAt
                FROM dbo.CampaignRecipients WHERE CampaignId = @Id ORDER BY FullName
                """;
            cmd.Parameters.AddWithValue("@Id", id);
            await using var r = await cmd.ExecuteReaderAsync(ct);
            while (await r.ReadAsync(ct))
            {
                sb.AppendLine(string.Join(",",
                    EscapeCsv(r.GetString(0)),
                    EscapeCsv(r.GetString(1)),
                    EscapeCsv(r.IsDBNull(2) ? "" : r.GetString(2)),
                    EscapeCsv(r.IsDBNull(3) ? "" : r.GetString(3)),
                    EscapeCsv(r.IsDBNull(4) ? "" : r.GetString(4)),
                    EscapeCsv(r.IsDBNull(5) ? "" : r.GetDateTime(5).ToString("yyyy-MM-dd HH:mm"))
                ));
            }
            var bytes = Encoding.UTF8.GetPreamble().Concat(Encoding.UTF8.GetBytes(sb.ToString())).ToArray();
            return File(bytes, "text/csv", $"campaña_{id}.csv");
        }
        catch (Exception ex) { logger.LogError(ex, "Export Campaign {Id}", id); return StatusCode(500, ex.Message); }
    }

    // ── Helpers ───────────────────────────────────────────────────────────────
    private static string EscapeCsv(string s) =>
        s.Contains(',') || s.Contains('"') || s.Contains('\n')
            ? $"\"{s.Replace("\"", "\"\"")}\""
            : s;

    private static string ApplyVars(string template, string fullName, string email,
        string? username, string? password, string programa,
        Dictionary<string, string> extra)
    {
        var result = template
            .Replace("{{nombre}}",     fullName,          StringComparison.OrdinalIgnoreCase)
            .Replace("{{email}}",      email,             StringComparison.OrdinalIgnoreCase)
            .Replace("{{usuario}}",    username  ?? "",   StringComparison.OrdinalIgnoreCase)
            .Replace("{{contrasena}}", password  ?? "",   StringComparison.OrdinalIgnoreCase)
            .Replace("{{programa}}",   programa,          StringComparison.OrdinalIgnoreCase);

        foreach (var kv in extra)
            result = Regex.Replace(result, $"{{{{\\s*{Regex.Escape(kv.Key)}\\s*}}}}", kv.Value,
                RegexOptions.IgnoreCase);

        return result;
    }

    // Convierte imágenes data:base64 embebidas en el HTML a adjuntos CID inline
    // para que los clientes de correo las muestren correctamente.
    private static MimeEntity BuildBodyWithInlineImages(string html)
    {
        var imgRegex = new Regex(
            @"src=""(data:(?<mime>image/[^;]+);base64,(?<data>[^""]+))""",
            RegexOptions.Compiled);

        var matches = imgRegex.Matches(html);
        if (matches.Count == 0)
            return new TextPart("html") { Text = html };

        var builder = new BodyBuilder();
        var processedHtml = html;

        foreach (Match m in matches)
        {
            var mimeType  = m.Groups["mime"].Value;           // e.g. "image/png"
            var base64    = m.Groups["data"].Value;
            var bytes     = Convert.FromBase64String(base64);
            var subtype   = mimeType.Split('/').Last();        // png, jpeg, gif, webp
            var cid       = Guid.NewGuid().ToString("N");

            var part = new MimePart("image", subtype)
            {
                Content                  = new MimeContent(new MemoryStream(bytes)),
                ContentDisposition       = new ContentDisposition(ContentDisposition.Inline),
                ContentTransferEncoding  = ContentEncoding.Base64,
                ContentId                = cid,
                FileName                 = $"img_{cid}.{subtype}",
            };

            builder.LinkedResources.Add(part);
            processedHtml = processedHtml.Replace(m.Value, $@"src=""cid:{cid}""");
        }

        builder.HtmlBody = processedHtml;
        return builder.ToMessageBody();
    }
}

// ── DTOs ──────────────────────────────────────────────────────────────────────
public record CampaignCreateDto(
    string             Name,
    string             Subject,
    string?            Body,
    int                ProgramType,
    List<RecipientDto> Recipients
);

public record RecipientDto(
    string              FullName,
    string              Email,
    string?             Username,
    string?             Password,
    Dictionary<string, JsonElement>? ExtraData
);
