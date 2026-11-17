using System.Text.Json;
using MailKit.Net.Smtp;
using MailKit.Security;
using Microsoft.Data.SqlClient;
using MimeKit;
using Pandora.API.Controllers;

namespace Pandora.API.Services;

/// <summary>
/// #1 Fase 2 — Worker de cola de campañas. Vacía automáticamente las campañas
/// que quedaron "En cola (límite diario)" cuando se alcanzó el tope de envío de
/// Gmail (~2,000/día). Antes, el resto solo se reanudaba reenviando a mano; este
/// worker lo hace solo, desde la MISMA cuenta que inició la campaña
/// (<c>EmailCampaigns.SenderEmail</c>) y respetando el tope diario por tenant.
///
/// Reutiliza los helpers de render estáticos de <see cref="CampaignsController"/>
/// (ApplyVars / ProgramLabel / BuildBodyWithInlineImages), así que el correo sale
/// idéntico al del envío manual.
/// </summary>
public class CampaignQueueWorker(
    IConfiguration config,
    ILogger<CampaignQueueWorker> logger) : BackgroundService
{
    private const string QueuedStatus = "En cola (límite diario)";

    private SqlConnection Conn() => new(config.GetConnectionString("PandoraDb"));

    private int IntervalMinutes =>
        int.TryParse(config["Campaigns:QueueWorkerMinutes"], out var m) && m > 0 ? m : 20;

    private int DailyLimit =>
        int.TryParse(config["Campaigns:DailySendLimit"], out var dl) && dl > 0 ? dl : 1800;

    protected override async Task ExecuteAsync(CancellationToken ct)
    {
        await Task.Delay(TimeSpan.FromSeconds(90), ct); // esperar a que la app/BD estén listas

        while (!ct.IsCancellationRequested)
        {
            try
            {
                await DrainQueueAsync(ct);
            }
            catch (OperationCanceledException) { break; }
            catch (Exception ex)
            {
                logger.LogError(ex, "CampaignQueueWorker: error inesperado en el ciclo.");
            }

            try { await Task.Delay(TimeSpan.FromMinutes(IntervalMinutes), ct); }
            catch (OperationCanceledException) { break; }
        }
    }

    private record QueuedCampaign(Guid Id, string Subject, string Body, int ProgramType, object TenantId, string? SenderEmail);

    private async Task DrainQueueAsync(CancellationToken ct)
    {
        List<QueuedCampaign> campaigns = [];
        await using (var conn = Conn())
        {
            await conn.OpenAsync(ct);

            // Si la columna SenderEmail aún no existe (BD sin el ALTER), no se puede
            // resolver la cuenta de envío: se omite el ciclo en silencio.
            bool hasSender;
            await using (var chk = conn.CreateCommand())
            {
                chk.CommandText = "SELECT COUNT(*) FROM sys.columns WHERE object_id=OBJECT_ID('dbo.EmailCampaigns') AND name='SenderEmail'";
                hasSender = Convert.ToInt32(await chk.ExecuteScalarAsync(ct) ?? 0) > 0;
            }
            if (!hasSender) return;

            await using var cmd = conn.CreateCommand();
            cmd.CommandText = $"""
                SELECT Id, Subject, Body, ProgramType, TenantId, SenderEmail
                FROM dbo.EmailCampaigns
                WHERE Status = @Queued AND IsDeleted = 0
                """;
            cmd.Parameters.AddWithValue("@Queued", QueuedStatus);
            await using var r = await cmd.ExecuteReaderAsync(ct);
            while (await r.ReadAsync(ct))
                campaigns.Add(new QueuedCampaign(
                    r.GetGuid(0),
                    r.IsDBNull(1) ? "" : r.GetString(1),
                    r.IsDBNull(2) ? "" : r.GetString(2),
                    r.IsDBNull(3) ? 1 : r.GetInt32(3),
                    r.IsDBNull(4) ? DBNull.Value : r.GetGuid(4),
                    r.IsDBNull(5) ? null : r.GetString(5)));
        }

        if (campaigns.Count == 0) return;
        logger.LogInformation("CampaignQueueWorker: {N} campaña(s) en cola por procesar.", campaigns.Count);

        foreach (var c in campaigns)
        {
            if (ct.IsCancellationRequested) break;
            try { await ProcessCampaignAsync(c, ct); }
            catch (Exception ex) { logger.LogError(ex, "CampaignQueueWorker: fallo procesando campaña {Id}", c.Id); }
        }
    }

    private async Task ProcessCampaignAsync(QueuedCampaign c, CancellationToken ct)
    {
        // ── Resolver cuenta SMTP (la misma que inició la campaña) ────────────────
        var smtpCfg  = config.GetSection("SmtpSettings");
        var smtpHost = smtpCfg["Host"] ?? "smtp.gmail.com";
        var smtpPort = int.TryParse(smtpCfg["Port"], out var p) ? p : 587;
        var fromName = smtpCfg["FromName"] ?? "Pandora";

        var fromEmail = !string.IsNullOrWhiteSpace(c.SenderEmail) ? c.SenderEmail! : (smtpCfg["FromEmail"] ?? "");
        string fromPass = "";

        await using (var conn = Conn())
        {
            await conn.OpenAsync(ct);
            if (!string.IsNullOrWhiteSpace(c.SenderEmail))
            {
                await using var cmd = conn.CreateCommand();
                cmd.CommandText = "SELECT SmtpPassword FROM dbo.AppUsers WHERE LOWER(SmtpEmail) = LOWER(@Email)";
                cmd.Parameters.AddWithValue("@Email", c.SenderEmail);
                var pass = await cmd.ExecuteScalarAsync(ct);
                fromPass = pass is string s ? s : "";
            }
        }
        if (string.IsNullOrWhiteSpace(fromPass))
            fromPass = smtpCfg["Password"] ?? "";

        if (string.IsNullOrWhiteSpace(fromEmail) || string.IsNullOrWhiteSpace(fromPass))
        {
            logger.LogWarning("CampaignQueueWorker: campaña {Id} sin credenciales SMTP para '{Email}'; se deja en cola.",
                c.Id, fromEmail);
            return;
        }

        // ── Tope diario por tenant ───────────────────────────────────────────────
        int sentToday = await CountSentTodayAsync(c.TenantId, ct);
        int remaining = DailyLimit - sentToday;
        if (remaining <= 0)
        {
            logger.LogInformation("CampaignQueueWorker: tenant aún en el tope diario ({Sent}/{Limit}); campaña {Id} sigue en cola.",
                sentToday, DailyLimit, c.Id);
            return;
        }

        // ── Destinatarios pendientes (hasta 'remaining') ─────────────────────────
        var recipients = new List<(Guid Id, string FullName, string Email, string? Username, string? Password, string? ExtraData)>();
        await using (var conn = Conn())
        {
            await conn.OpenAsync(ct);
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = """
                SELECT TOP (@Take) Id, FullName, Email, Username, Password, ExtraData
                FROM dbo.CampaignRecipients
                WHERE CampaignId = @CampaignId AND Status = 'Pending'
                ORDER BY CreatedAt
                """;
            cmd.Parameters.AddWithValue("@Take", remaining);
            cmd.Parameters.AddWithValue("@CampaignId", c.Id);
            await using var r = await cmd.ExecuteReaderAsync(ct);
            while (await r.ReadAsync(ct))
                recipients.Add((
                    r.GetGuid(0), r.GetString(1), r.GetString(2),
                    r.IsDBNull(3) ? null : r.GetString(3),
                    r.IsDBNull(4) ? null : r.GetString(4),
                    r.IsDBNull(5) ? null : r.GetString(5)));
        }

        if (recipients.Count == 0)
        {
            await FinalizeStatusAsync(c.Id, ct); // ya no hay pendientes: cerrar estado
            return;
        }

        logger.LogInformation("CampaignQueueWorker: campaña {Id} — enviando {N} de {Remaining} disponibles hoy.",
            c.Id, recipients.Count, remaining);

        await SetStatusAsync(c.Id, "Enviando", ct);

        var baseUrl = config["Campaigns:PublicApiUrl"]; // opcional; sin él, sin pixel de seguimiento

        foreach (var rec in recipients)
        {
            if (ct.IsCancellationRequested) break;
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

                var finalSubject = CampaignsController.ApplyVars(c.Subject, rec.FullName, rec.Email,
                    rec.Username, rec.Password, CampaignsController.ProgramLabel(c.ProgramType), extraVars);
                var finalBody = CampaignsController.ApplyVars(c.Body, rec.FullName, rec.Email,
                    rec.Username, rec.Password, CampaignsController.ProgramLabel(c.ProgramType), extraVars);

                if (!string.IsNullOrWhiteSpace(baseUrl))
                {
                    var trackingPixel = $"<img src=\"{baseUrl}/api/campaigns/track/{rec.Id}\" width=\"1\" height=\"1\" style=\"display:none;\" alt=\"\" />";
                    finalBody = finalBody.Contains("</body>", StringComparison.OrdinalIgnoreCase)
                        ? finalBody.Replace("</body>", $"{trackingPixel}</body>", StringComparison.OrdinalIgnoreCase)
                        : finalBody + trackingPixel;
                }

                var message = new MimeMessage();
                message.From.Add(new MailboxAddress(fromName, fromEmail));
                message.To.Add(new MailboxAddress(rec.FullName, rec.Email));
                message.Subject = finalSubject;
                message.Body = CampaignsController.BuildBodyWithInlineImages(finalBody);

                using var mailClient = new SmtpClient();
                await mailClient.ConnectAsync(smtpHost, smtpPort, SecureSocketOptions.StartTls, ct);
                await mailClient.AuthenticateAsync(fromEmail, fromPass, ct);
                await mailClient.SendAsync(message, ct);
                await mailClient.DisconnectAsync(true, ct);

                await using var conn2 = Conn();
                await conn2.OpenAsync(ct);
                await using var cmd2 = conn2.CreateCommand();
                cmd2.CommandText = "UPDATE dbo.CampaignRecipients SET Status='Sent', SentAt=GETUTCDATE(), Password=NULL WHERE Id=@Id";
                cmd2.Parameters.AddWithValue("@Id", rec.Id);
                await cmd2.ExecuteNonQueryAsync(ct);
            }
            catch (Exception ex)
            {
                try
                {
                    await using var conn2 = Conn();
                    await conn2.OpenAsync(ct);
                    await using var cmd2 = conn2.CreateCommand();
                    cmd2.CommandText = "UPDATE dbo.CampaignRecipients SET Status='Failed', ErrorMessage=@Err WHERE Id=@Id";
                    cmd2.Parameters.AddWithValue("@Err", ex.Message);
                    cmd2.Parameters.AddWithValue("@Id", rec.Id);
                    await cmd2.ExecuteNonQueryAsync(ct);
                }
                catch { /* ignore */ }
            }
        }

        await FinalizeStatusAsync(c.Id, ct);
    }

    private async Task<int> CountSentTodayAsync(object tenantId, CancellationToken ct)
    {
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = """
                SELECT COUNT(*) FROM dbo.CampaignRecipients cr
                JOIN dbo.EmailCampaigns c ON c.Id = cr.CampaignId
                WHERE cr.Status = 'Sent' AND cr.SentAt >= @Today AND c.TenantId = @TenantId
                """;
            cmd.Parameters.AddWithValue("@Today", DateTime.UtcNow.Date);
            cmd.Parameters.AddWithValue("@TenantId", tenantId);
            return Convert.ToInt32(await cmd.ExecuteScalarAsync(ct) ?? 0);
        }
        catch { return 0; } // si el conteo falla, no se bloquea el envío
    }

    private async Task SetStatusAsync(Guid id, string status, CancellationToken ct)
    {
        await using var conn = Conn();
        await conn.OpenAsync(ct);
        await using var cmd = conn.CreateCommand();
        cmd.CommandText = "UPDATE dbo.EmailCampaigns SET Status=@S, UpdatedAt=GETUTCDATE() WHERE Id=@Id";
        cmd.Parameters.AddWithValue("@S", status);
        cmd.Parameters.AddWithValue("@Id", id);
        await cmd.ExecuteNonQueryAsync(ct);
    }

    // Recalcula contadores y estado final desde la BD (idéntico al envío manual):
    // si quedan Pending, sigue "En cola (límite diario)"; si no, Completado / con errores.
    private async Task FinalizeStatusAsync(Guid id, CancellationToken ct)
    {
        await using var conn = Conn();
        await conn.OpenAsync(ct);

        int pendientes = 0, totalSent = 0, totalFailed = 0;
        await using (var cmdC = conn.CreateCommand())
        {
            cmdC.CommandText = """
                SELECT
                  SUM(CASE WHEN Status='Pending' THEN 1 ELSE 0 END),
                  SUM(CASE WHEN Status='Sent'    THEN 1 ELSE 0 END),
                  SUM(CASE WHEN Status='Failed'  THEN 1 ELSE 0 END)
                FROM dbo.CampaignRecipients WHERE CampaignId=@Id
                """;
            cmdC.Parameters.AddWithValue("@Id", id);
            await using var rC = await cmdC.ExecuteReaderAsync(ct);
            if (await rC.ReadAsync(ct))
            {
                pendientes  = rC.IsDBNull(0) ? 0 : rC.GetInt32(0);
                totalSent   = rC.IsDBNull(1) ? 0 : rC.GetInt32(1);
                totalFailed = rC.IsDBNull(2) ? 0 : rC.GetInt32(2);
            }
        }

        string estado = pendientes > 0
            ? QueuedStatus
            : (totalFailed == 0 ? "Completado" : "Completado con errores");

        await using var cmd2 = conn.CreateCommand();
        cmd2.CommandText = """
            UPDATE dbo.EmailCampaigns
            SET Status = @Estado, SentCount = @Sent, FailedCount = @Failed, UpdatedAt = GETUTCDATE()
            WHERE Id = @Id
            """;
        cmd2.Parameters.AddWithValue("@Estado", estado);
        cmd2.Parameters.AddWithValue("@Sent", totalSent);
        cmd2.Parameters.AddWithValue("@Failed", totalFailed);
        cmd2.Parameters.AddWithValue("@Id", id);
        await cmd2.ExecuteNonQueryAsync(ct);

        logger.LogInformation("CampaignQueueWorker: campaña {Id} → {Estado} (enviados={Sent}, fallidos={Failed}, pendientes={Pend}).",
            id, estado, totalSent, totalFailed, pendientes);
    }
}
