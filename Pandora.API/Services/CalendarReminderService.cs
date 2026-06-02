using MailKit.Net.Smtp;
using MailKit.Security;
using MimeKit;
using System.Globalization;

namespace Pandora.API.Services;

/// <summary>
/// Envía recordatorios de email 24 horas y 1 hora antes de cada reservación del calendario.
/// Requiere SmtpSettings en appsettings (FromEmail + Password como mínimo).
/// </summary>
public class CalendarReminderService(
    IConfiguration config,
    ILogger<CalendarReminderService> logger) : BackgroundService
{
    private static readonly TimeZoneInfo MxTz =
        TimeZoneInfo.FindSystemTimeZoneById(
            OperatingSystem.IsWindows() ? "Central Standard Time" : "America/Mexico_City");

    protected override async Task ExecuteAsync(CancellationToken ct)
    {
        // Esperar 60 s al inicio para que la BD esté lista
        await Task.Delay(TimeSpan.FromSeconds(60), ct);

        while (!ct.IsCancellationRequested)
        {
            try
            {
                await EnsureReminderTableAsync(ct);
                await SendRemindersAsync(ct);
            }
            catch (Exception ex)
            {
                logger.LogError(ex, "CalendarReminderService error");
            }

            // Ejecutar cada hora
            await Task.Delay(TimeSpan.FromHours(1), ct);
        }
    }

    // ── Crea la tabla de deduplicación si no existe ───────────────────────────
    private async Task EnsureReminderTableAsync(CancellationToken ct)
    {
        await using var conn = new Microsoft.Data.SqlClient.SqlConnection(
            config.GetConnectionString("PandoraDb"));
        await conn.OpenAsync(ct);
        await using var cmd = conn.CreateCommand();
        cmd.CommandText = """
            IF NOT EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID(N'dbo.CalendarReminders') AND type = N'U')
            CREATE TABLE dbo.CalendarReminders (
                Id             INT IDENTITY PRIMARY KEY,
                ReservationId  UNIQUEIDENTIFIER NOT NULL,
                ReminderType   NVARCHAR(10) NOT NULL,  -- '24h' | '1h'
                SentAt         DATETIME2 NOT NULL DEFAULT GETUTCDATE(),
                CONSTRAINT UQ_CalReminder UNIQUE (ReservationId, ReminderType)
            );
            """;
        await cmd.ExecuteNonQueryAsync(ct);
    }

    // ── Busca y envía recordatorios pendientes ────────────────────────────────
    private async Task SendRemindersAsync(CancellationToken ct)
    {
        var smtp = config.GetSection("SmtpSettings");
        var host = smtp["Host"] ?? "";
        var port = int.TryParse(smtp["Port"], out var p) ? p : 587;
        var from = smtp["FromEmail"] ?? "";
        var pass = smtp["Password"] ?? "";
        var fromName = smtp["FromName"] ?? "Pandora";

        if (string.IsNullOrWhiteSpace(host) || string.IsNullOrWhiteSpace(from))
        {
            logger.LogDebug("CalendarReminderService: SMTP no configurado, omitiendo.");
            return;
        }

        var utcNow = DateTime.UtcNow;

        // Reservaciones que empiezan en ~24 h (rango ±30 min)
        var win24Start = utcNow.AddHours(23).AddMinutes(30);
        var win24End   = utcNow.AddHours(24).AddMinutes(30);

        // Reservaciones que empiezan en ~1 h (rango ±15 min)
        var win1Start  = utcNow.AddMinutes(45);
        var win1End    = utcNow.AddMinutes(75);

        var pending = new List<ReminderTask>();

        await using var conn = new Microsoft.Data.SqlClient.SqlConnection(
            config.GetConnectionString("PandoraDb"));
        await conn.OpenAsync(ct);

        // 24-hour window
        await QueryReminders(conn, pending, win24Start, win24End, "24h", ct);
        // 1-hour window
        await QueryReminders(conn, pending, win1Start, win1End, "1h", ct);

        if (pending.Count == 0) return;

        logger.LogInformation("CalendarReminderService: {Count} recordatorio(s) a enviar", pending.Count);

        // Connect SMTP once
        using var smtpClient = new SmtpClient();
        try
        {
            await smtpClient.ConnectAsync(host, port, SecureSocketOptions.StartTls, ct);
            await smtpClient.AuthenticateAsync(from, pass, ct);
        }
        catch (Exception ex)
        {
            logger.LogError(ex, "SMTP connect error en CalendarReminderService");
            return;
        }

        foreach (var task in pending)
        {
            try
            {
                var mxStart = TimeZoneInfo.ConvertTimeFromUtc(task.StartTime, MxTz);
                var mxEnd   = TimeZoneInfo.ConvertTimeFromUtc(task.EndTime,   MxTz);

                var recipients = new List<(string Name, string Email)>();
                if (!string.IsNullOrWhiteSpace(task.OrganizerEmail))
                    recipients.Add((task.OrganizerName ?? "Organizador", task.OrganizerEmail!));

                foreach (var (name, email) in recipients)
                {
                    var msg = new MimeMessage();
                    msg.From.Add(new MailboxAddress(fromName, from));
                    msg.To.Add(new MailboxAddress(name, email));
                    msg.Subject = task.ReminderType == "1h"
                        ? $"[iMET] ⏰ Tu reunión '{task.Title}' comienza en 1 hora"
                        : $"[iMET] 📅 Recordatorio: mañana tienes '{task.Title}'";

                    msg.Body = new TextPart("html")
                    {
                        Text = BuildReminderBody(
                            recipientName: name,
                            title:         task.Title,
                            roomName:      task.RoomName,
                            mxStart:       mxStart,
                            mxEnd:         mxEnd,
                            meetLink:      task.MeetLink,
                            reminderType:  task.ReminderType)
                    };

                    await smtpClient.SendAsync(msg, ct);
                }

                // Mark as sent (upsert-safe — ignore if already exists)
                await using var markCmd = conn.CreateCommand();
                markCmd.CommandText = """
                    IF NOT EXISTS (SELECT 1 FROM dbo.CalendarReminders WHERE ReservationId=@Id AND ReminderType=@Type)
                    INSERT INTO dbo.CalendarReminders (ReservationId, ReminderType) VALUES (@Id, @Type)
                    """;
                markCmd.Parameters.AddWithValue("@Id",   task.ReservationId);
                markCmd.Parameters.AddWithValue("@Type", task.ReminderType);
                await markCmd.ExecuteNonQueryAsync(ct);
            }
            catch (Exception ex)
            {
                logger.LogWarning("Reminder email failed for {Id}: {Msg}", task.ReservationId, ex.Message);
            }
        }

        await smtpClient.DisconnectAsync(true, ct);
    }

    private static async Task QueryReminders(
        Microsoft.Data.SqlClient.SqlConnection conn,
        List<ReminderTask> result,
        DateTime winStart, DateTime winEnd,
        string reminderType,
        CancellationToken ct)
    {
        await using var cmd = conn.CreateCommand();
        cmd.CommandText = """
            SELECT r.Id, r.Title, r.StartTime, r.EndTime, r.OrganizerEmail, r.OrganizerName,
                   r.MeetLink, rm.Name AS RoomName
            FROM dbo.Reservations r
            INNER JOIN dbo.Rooms rm ON r.RoomId = rm.Id
            WHERE r.StartTime BETWEEN @Start AND @End
              AND NOT EXISTS (
                SELECT 1 FROM dbo.CalendarReminders cr
                WHERE cr.ReservationId = r.Id AND cr.ReminderType = @Type
              )
            """;
        cmd.Parameters.AddWithValue("@Start", winStart);
        cmd.Parameters.AddWithValue("@End",   winEnd);
        cmd.Parameters.AddWithValue("@Type",  reminderType);

        await using var r = await cmd.ExecuteReaderAsync(ct);
        while (await r.ReadAsync(ct))
        {
            result.Add(new ReminderTask(
                ReservationId:  r.GetGuid(0),
                Title:          r.GetString(1),
                StartTime:      r.GetDateTime(2),
                EndTime:        r.GetDateTime(3),
                OrganizerEmail: r.IsDBNull(4) ? null : r.GetString(4),
                OrganizerName:  r.IsDBNull(5) ? null : r.GetString(5),
                MeetLink:       r.IsDBNull(6) ? null : r.GetString(6),
                RoomName:       r.IsDBNull(7) ? null : r.GetString(7),
                ReminderType:   reminderType
            ));
        }
    }

    private static string BuildReminderBody(
        string recipientName, string title, string? roomName,
        DateTime mxStart, DateTime mxEnd,
        string? meetLink, string reminderType)
    {
        var isHour  = reminderType == "1h";
        var header  = isHour ? "Tu reunión comienza en 1 hora" : "Recordatorio: mañana tienes una reunión";
        var intro   = isHour
            ? "Tu reunión <strong>comienza en aproximadamente 1 hora</strong>. ¡Prepárate!"
            : "Te recordamos que <strong>mañana</strong> tienes la siguiente reunión agendada.";
        var color   = isHour ? "#e65100" : "#1a237e";
        var emoji   = isHour ? "⏰" : "📅";

        var dateStr  = mxStart.ToString("dddd, dd 'de' MMMM 'de' yyyy", new CultureInfo("es-MX"));
        var startStr = mxStart.ToString("HH:mm");
        var endStr   = mxEnd.ToString("HH:mm");

        var meetRow = string.IsNullOrWhiteSpace(meetLink) ? "" : $"""
              <tr style="border-bottom:1px solid #eee">
                <td style="padding:8px;font-weight:bold;color:#555;width:40%">Link</td>
                <td style="padding:8px"><a href="{meetLink}" style="color:#1a237e">{meetLink}</a></td>
              </tr>
            """;

        return $"""
            <html><body style="font-family:Arial,sans-serif;font-size:14px;color:#333">
            <div style="max-width:600px;margin:0 auto">
              <div style="background:{color};padding:20px;border-radius:8px 8px 0 0">
                <h2 style="color:white;margin:0">{emoji} {header}</h2>
              </div>
              <div style="border:1px solid #ddd;padding:24px;border-radius:0 0 8px 8px">
                <p>Hola <strong>{recipientName}</strong>,</p>
                <p>{intro}</p>
                <table style="width:100%;border-collapse:collapse">
                  <tr style="border-bottom:1px solid #eee">
                    <td style="padding:8px;font-weight:bold;color:#555;width:40%">Evento</td>
                    <td style="padding:8px"><strong>{title}</strong></td>
                  </tr>
                  <tr style="border-bottom:1px solid #eee">
                    <td style="padding:8px;font-weight:bold;color:#555">Sala</td>
                    <td style="padding:8px">{roomName ?? "—"}</td>
                  </tr>
                  <tr style="border-bottom:1px solid #eee">
                    <td style="padding:8px;font-weight:bold;color:#555">Fecha</td>
                    <td style="padding:8px">{dateStr}</td>
                  </tr>
                  <tr style="border-bottom:1px solid #eee">
                    <td style="padding:8px;font-weight:bold;color:#555">Hora</td>
                    <td style="padding:8px">{startStr} – {endStr} (hora de México)</td>
                  </tr>
                  {meetRow}
                </table>
                <p style="margin-top:24px;font-size:12px;color:#888">
                  Este recordatorio fue generado automáticamente por <strong>Pandora — Sistema de Gestión iMET</strong>.<br/>
                  Para gestionar tus reservaciones, accede a Pandora Calendar.
                </p>
              </div>
            </div>
            </body></html>
            """;
    }

    private record ReminderTask(
        Guid    ReservationId,
        string  Title,
        DateTime StartTime,
        DateTime EndTime,
        string? OrganizerEmail,
        string? OrganizerName,
        string? MeetLink,
        string? RoomName,
        string  ReminderType);
}
