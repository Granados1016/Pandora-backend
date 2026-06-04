using Microsoft.Data.SqlClient;

namespace Pandora.API.Services;

/// <summary>
/// Servicio en background que revisa diariamente las licencias de tenants
/// y envía alertas por correo al proveedor (Pandora) cuando están por vencer.
/// Alertas: 30, 15 y 5 días antes de la expiración.
/// </summary>
public class TenantLicenseAlertService(IConfiguration config, ILogger<TenantLicenseAlertService> logger)
    : BackgroundService
{
    private SqlConnection Conn() => new(config.GetConnectionString("PandoraDb"));

    protected override async Task ExecuteAsync(CancellationToken ct)
    {
        // Esperar 2 minutos al inicio para que el sistema se estabilice
        await Task.Delay(TimeSpan.FromMinutes(2), ct);

        while (!ct.IsCancellationRequested)
        {
            try { await CheckLicensesAsync(ct); }
            catch (Exception ex) { logger.LogError(ex, "TenantLicenseAlertService error"); }

            // Revisar cada 24 horas
            await Task.Delay(TimeSpan.FromHours(24), ct);
        }
    }

    private async Task CheckLicensesAsync(CancellationToken ct)
    {
        var alertDays = new[] { 30, 15, 5 };
        var providerEmail = config["SuperAdmin:AlertEmail"] ?? config["SmtpSettings:NotificationsEmail"] ?? "";

        if (string.IsNullOrWhiteSpace(providerEmail))
        {
            logger.LogWarning("TenantLicenseAlert: no hay correo de proveedor configurado (SuperAdmin:AlertEmail).");
            return;
        }

        await using var conn = Conn();
        await conn.OpenAsync(ct);
        await using var cmd = conn.CreateCommand();
        cmd.CommandText = """
            SELECT Id, Slug, Name, DisplayName, ExpiresAt, ContactEmail
            FROM dbo.Tenants
            WHERE IsActive = 1 AND ExpiresAt IS NOT NULL
              AND ExpiresAt > GETUTCDATE()
              AND DATEDIFF(DAY, GETUTCDATE(), ExpiresAt) IN (30, 15, 5)
            """;
        var toAlert = new List<(Guid Id, string Slug, string Name, string Display, DateTime Exp, string? Contact)>();
        await using (var r = await cmd.ExecuteReaderAsync(ct))
        {
            while (await r.ReadAsync(ct))
                toAlert.Add((
                    r.GetGuid(0), r.GetString(1), r.GetString(2), r.GetString(3),
                    r.GetDateTime(4),
                    r.IsDBNull(5) ? null : r.GetString(5)
                ));
        }

        if (toAlert.Count == 0) return;

        var smtp = await SmtpHelper.LoadAsync(config.GetConnectionString("PandoraDb")!, config);
        foreach (var t in toAlert)
        {
            var days = (int)Math.Ceiling((t.Exp - DateTime.UtcNow).TotalDays);
            var body = $"""
                <html><body style="font-family:Arial,sans-serif;font-size:14px;color:#333">
                <div style="max-width:600px;margin:0 auto">
                  <div style="background:#1a237e;padding:20px;border-radius:8px 8px 0 0">
                    <h2 style="color:white;margin:0">⚠️ Licencia por vencer — {t.Display}</h2>
                  </div>
                  <div style="border:1px solid #ddd;padding:24px;border-radius:0 0 8px 8px">
                    <p>La licencia del cliente <strong>{t.Name}</strong> ({t.Slug}) vence en <strong>{days} días</strong>.</p>
                    <table style="width:100%;border-collapse:collapse">
                      <tr style="border-bottom:1px solid #eee"><td style="padding:8px;font-weight:bold;color:#555">Cliente</td><td style="padding:8px">{t.Name}</td></tr>
                      <tr style="border-bottom:1px solid #eee"><td style="padding:8px;font-weight:bold;color:#555">Slug</td><td style="padding:8px">{t.Slug}</td></tr>
                      <tr style="border-bottom:1px solid #eee"><td style="padding:8px;font-weight:bold;color:#555">Vencimiento</td><td style="padding:8px">{t.Exp:dd/MM/yyyy}</td></tr>
                      <tr><td style="padding:8px;font-weight:bold;color:#555">Contacto</td><td style="padding:8px">{t.Contact ?? "—"}</td></tr>
                    </table>
                    <p style="margin-top:16px">Renueve la licencia en el <strong>Panel de Clientes</strong> de Pandora antes de la fecha de vencimiento para evitar la suspensión del servicio.</p>
                    <p style="font-size:12px;color:#999">Pandora — Sistema de Gestión · Alerta automática</p>
                  </div>
                </div></body></html>
                """;

            var err = await SmtpHelper.SendAsync(smtp, providerEmail, "Pandora Admin",
                $"[Pandora] Licencia por vencer — {t.Display} ({days} días)", body);

            if (err == null)
                logger.LogInformation("Alerta de licencia enviada: {Slug} vence en {Days} días", t.Slug, days);
            else
                logger.LogWarning("No se pudo enviar alerta para {Slug}: {Err}", t.Slug, err);
        }
    }
}
