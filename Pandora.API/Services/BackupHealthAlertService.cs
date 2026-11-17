using Microsoft.Data.SqlClient;

namespace Pandora.API.Services;

/// <summary>
/// #9 Fase 2 (monitoreo) — Servicio en background que vigila que los respaldos
/// automáticos sigan corriendo. Una vez al día revisa dbo.BackupHistory: si el
/// respaldo automático está habilitado pero NO hay un respaldo exitoso reciente
/// (dentro de Monitoring:BackupStaleHours, por defecto 48h), envía una alerta al
/// correo del proveedor (SuperAdmin:AlertEmail). Complementa el sistema de
/// respaldos (#4): de nada sirve respaldar si nadie se entera cuando deja de
/// correr. El loop de 24h limita la alerta a una vez al día.
/// </summary>
public class BackupHealthAlertService(IConfiguration config, ILogger<BackupHealthAlertService> logger)
    : BackgroundService
{
    private SqlConnection Conn() => new(config.GetConnectionString("PandoraDb"));

    private int StaleHours =>
        int.TryParse(config["Monitoring:BackupStaleHours"], out var h) && h > 0 ? h : 48;

    protected override async Task ExecuteAsync(CancellationToken ct)
    {
        // Esperar 3 minutos al inicio para que el sistema (y la BD) se estabilicen.
        await Task.Delay(TimeSpan.FromMinutes(3), ct);

        while (!ct.IsCancellationRequested)
        {
            try { await CheckBackupHealthAsync(ct); }
            catch (Exception ex) { logger.LogError(ex, "BackupHealthAlertService error"); }

            await Task.Delay(TimeSpan.FromHours(24), ct);
        }
    }

    private async Task CheckBackupHealthAsync(CancellationToken ct)
    {
        await using var conn = Conn();
        await conn.OpenAsync(ct);

        // Si la tabla de historial aún no existe, no hay nada que vigilar.
        await using (var chk = conn.CreateCommand())
        {
            chk.CommandText = "SELECT CASE WHEN OBJECT_ID('dbo.BackupHistory') IS NULL THEN 0 ELSE 1 END";
            if (Convert.ToInt32(await chk.ExecuteScalarAsync(ct) ?? 0) == 0) return;
        }

        // Solo alertar si el respaldo automático está habilitado por el cliente.
        bool autoEnabled = false;
        await using (var cmd = conn.CreateCommand())
        {
            cmd.CommandText = "SELECT SettingValue FROM dbo.SystemSettings WHERE SettingKey = 'backup_auto_enabled'";
            autoEnabled = (await cmd.ExecuteScalarAsync(ct) as string) == "true";
        }
        if (!autoEnabled)
        {
            logger.LogDebug("BackupHealthAlert: respaldo automático deshabilitado, no se vigila.");
            return;
        }

        // ¿Hay un respaldo EXITOSO (archivo generado) dentro de la ventana?
        DateTime? lastOk = null;
        await using (var cmd = conn.CreateCommand())
        {
            cmd.CommandText = "SELECT MAX(RanAt) FROM dbo.BackupHistory WHERE SizeBytes > 0";
            var scalar = await cmd.ExecuteScalarAsync(ct);
            if (scalar is DateTime dt) lastOk = dt;
        }

        var threshold = DateTime.UtcNow.AddHours(-StaleHours);
        if (lastOk.HasValue && lastOk.Value >= threshold)
        {
            logger.LogDebug("BackupHealthAlert: último respaldo OK a las {Last} (dentro de {Hours}h).", lastOk, StaleHours);
            return; // todo en orden
        }

        // Respaldo viejo o inexistente → alertar al proveedor.
        var providerEmail = config["SuperAdmin:AlertEmail"] ?? config["SmtpSettings:NotificationsEmail"] ?? "";
        if (string.IsNullOrWhiteSpace(providerEmail))
        {
            logger.LogWarning("BackupHealthAlert: respaldo atrasado pero no hay correo de proveedor (SuperAdmin:AlertEmail).");
            return;
        }

        var lastText = lastOk.HasValue ? $"{lastOk.Value:dd/MM/yyyy HH:mm} UTC" : "nunca (sin respaldos exitosos registrados)";
        var body = $"""
            <html><body style="font-family:Arial,sans-serif;font-size:14px;color:#333">
            <div style="max-width:600px;margin:0 auto">
              <div style="background:#b71c1c;padding:20px;border-radius:8px 8px 0 0">
                <h2 style="color:white;margin:0">🛑 Respaldo automático atrasado</h2>
              </div>
              <div style="border:1px solid #ddd;padding:24px;border-radius:0 0 8px 8px">
                <p>El respaldo automático de Pandora está <strong>habilitado</strong>, pero no se detecta
                   un respaldo exitoso en las últimas <strong>{StaleHours} horas</strong>.</p>
                <table style="width:100%;border-collapse:collapse">
                  <tr style="border-bottom:1px solid #eee"><td style="padding:8px;font-weight:bold;color:#555;width:45%">Último respaldo exitoso</td><td style="padding:8px">{lastText}</td></tr>
                  <tr><td style="padding:8px;font-weight:bold;color:#555">Umbral de alerta</td><td style="padding:8px">{StaleHours} horas</td></tr>
                </table>
                <p style="margin-top:16px">Revise el panel de administración → Respaldos, o ejecute uno manual
                   ("Ejecutar ahora"), y verifique la configuración de correo/Drive.</p>
                <p style="font-size:12px;color:#999">Pandora — Sistema de Gestión · Alerta automática de monitoreo</p>
              </div>
            </div></body></html>
            """;

        var smtp = await SmtpHelper.LoadAsync(config.GetConnectionString("PandoraDb")!, config);
        var err = await SmtpHelper.SendAsync(smtp, providerEmail, "Pandora Admin",
            "[Pandora] Respaldo automático atrasado", body);

        if (err == null)
            logger.LogWarning("BackupHealthAlert: alerta enviada (último respaldo: {Last}).", lastText);
        else
            logger.LogError("BackupHealthAlert: no se pudo enviar la alerta: {Err}", err);
    }
}
