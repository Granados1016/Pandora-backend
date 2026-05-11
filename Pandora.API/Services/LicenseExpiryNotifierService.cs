using Microsoft.AspNetCore.SignalR;
using Microsoft.Data.SqlClient;
using Pandora.API.Hubs;

namespace Pandora.API.Services;

/// <summary>
/// Servicio en segundo plano que revisa licencias próximas a vencer
/// y genera notificaciones automáticas para administradores.
/// Se ejecuta cada 24 horas.
/// </summary>
public class LicenseExpiryNotifierService(
    IConfiguration config,
    IHubContext<NotificationsHub> hub,
    ILogger<LicenseExpiryNotifierService> logger) : BackgroundService
{
    private readonly int[] _thresholds = [7, 15, 30]; // días de anticipación

    protected override async Task ExecuteAsync(CancellationToken stoppingToken)
    {
        // Esperar 30 segundos después del inicio para que la BD esté lista
        await Task.Delay(TimeSpan.FromSeconds(30), stoppingToken);

        while (!stoppingToken.IsCancellationRequested)
        {
            try { await CheckAndNotifyAsync(stoppingToken); }
            catch (Exception ex)
            {
                logger.LogError(ex, "Error en LicenseExpiryNotifierService");
            }
            // Ejecutar cada 24 horas
            await Task.Delay(TimeSpan.FromHours(24), stoppingToken);
        }
    }

    private async Task CheckAndNotifyAsync(CancellationToken ct)
    {
        var connStr = config.GetConnectionString("PandoraDb")!;

        foreach (int days in _thresholds)
        {
            var expiring = await GetExpiringLicensesAsync(connStr, days, ct);
            foreach (var lic in expiring)
            {
                // Evitar notificaciones duplicadas en el mismo día
                bool alreadyNotified = await WasNotifiedTodayAsync(connStr, lic.Id, days, ct);
                if (alreadyNotified) continue;

                string title   = $"⚠️ Licencia por vencer en {days} días";
                string message = $"{lic.Plataforma} ({lic.Area}) vence el {lic.ProximoPago:dd/MM/yyyy}. Responsable: {lic.Responsable ?? "No asignado"}.";

                int notifId = await CreateNotificationAsync(connStr, title, message, ct);

                // Enviar por SignalR al grupo broadcast (todos los admins conectados)
                var payload = new { id = notifId, title, message, type = "warning" };
                await hub.Clients.Group("broadcast").SendAsync("NewNotification", payload, ct);

                logger.LogInformation(
                    "Notificación generada: licencia #{Id} '{Plataforma}' vence en {Days} días",
                    lic.Id, lic.Plataforma, days);
            }
        }
    }

    // ── Licencias que vencen exactamente en `days` días ───────────────────────
    private static async Task<List<LicenciaRow>> GetExpiringLicensesAsync(
        string connStr, int days, CancellationToken ct)
    {
        var list = new List<LicenciaRow>();
        await using var conn = new SqlConnection(connStr);
        await conn.OpenAsync(ct);
        await using var cmd = conn.CreateCommand();
        cmd.CommandText = """
            SELECT Id, Plataforma, Area, Responsable, ProximoPago
            FROM dbo.Licencias
            WHERE Estado NOT IN ('Cancelada','Dado de baja')
              AND CAST(ProximoPago AS DATE) = CAST(DATEADD(DAY, @Days, GETUTCDATE()) AS DATE)
            """;
        cmd.Parameters.AddWithValue("@Days", days);
        await using var r = await cmd.ExecuteReaderAsync(ct);
        while (await r.ReadAsync(ct))
        {
            list.Add(new LicenciaRow(
                r.GetInt32(0),
                r.GetString(1),
                r.GetString(2),
                r.IsDBNull(3) ? null : r.GetString(3),
                r.GetDateTime(4)));
        }
        return list;
    }

    // ── Verificar si ya se notificó hoy para esta licencia y umbral ───────────
    private static async Task<bool> WasNotifiedTodayAsync(
        string connStr, int licId, int days, CancellationToken ct)
    {
        await using var conn = new SqlConnection(connStr);
        await conn.OpenAsync(ct);
        await using var cmd = conn.CreateCommand();
        cmd.CommandText = """
            SELECT COUNT(1) FROM dbo.Notifications
            WHERE Title LIKE @Pattern
              AND Message LIKE @LicId
              AND CAST(CreatedAt AS DATE) = CAST(GETUTCDATE() AS DATE)
              AND IsDeleted = 0
            """;
        cmd.Parameters.AddWithValue("@Pattern", $"%vencer en {days} días%");
        cmd.Parameters.AddWithValue("@LicId",   $"% {licId} %");
        // Fallback simple: buscar por título + mensaje aproximado
        cmd.Parameters["@LicId"].Value = $"%{licId}%";
        var count = (int)(await cmd.ExecuteScalarAsync(ct) ?? 0);
        return count > 0;
    }

    // ── Insertar notificación en BD ───────────────────────────────────────────
    private static async Task<int> CreateNotificationAsync(
        string connStr, string title, string message, CancellationToken ct)
    {
        await using var conn = new SqlConnection(connStr);
        await conn.OpenAsync(ct);
        await using var cmd = conn.CreateCommand();
        cmd.CommandText = """
            INSERT INTO dbo.Notifications
                (Title, Message, Type, IsRead, TargetUser, CreatedBy, CreatedAt, IsDeleted)
            OUTPUT INSERTED.Id
            VALUES (@Title, @Message, 'warning', 0, NULL, 'system', GETUTCDATE(), 0)
            """;
        cmd.Parameters.AddWithValue("@Title",   title);
        cmd.Parameters.AddWithValue("@Message", message);
        return (int)(await cmd.ExecuteScalarAsync(ct))!;
    }

    private record LicenciaRow(int Id, string Plataforma, string Area, string? Responsable, DateTime ProximoPago);
}
