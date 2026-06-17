using Microsoft.Data.SqlClient;
using WebPush;

namespace Pandora.API.Services;

/// <summary>
/// Servicio para enviar Web Push Notifications via VAPID.
/// </summary>
public class PushService(IConfiguration config, ILogger<PushService> logger)
{
    private string ConnStr => config.GetConnectionString("PandoraDb")!;
    private string VapidPublic  => config["Push:VapidPublicKey"]  ?? config["Push__VapidPublicKey"]  ?? "";
    private string VapidPrivate => config["Push:VapidPrivateKey"] ?? config["Push__VapidPrivateKey"] ?? "";
    private string VapidSubject => config["Push:VapidSubject"]    ?? config["Push__VapidSubject"]    ?? "mailto:pandora@pandora.app";

    // ── Obtener suscripciones de un usuario ───────────────────────────────────
    public async Task<List<PushSubscriptionRecord>> GetSubscriptionsAsync(string userId)
    {
        var list = new List<PushSubscriptionRecord>();
        try
        {
            await using var conn = new SqlConnection(ConnStr);
            await conn.OpenAsync();
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = "SELECT Endpoint, P256dh, Auth FROM dbo.PushSubscriptions WHERE UserId = @UserId";
            cmd.Parameters.AddWithValue("@UserId", userId);
            await using var r = await cmd.ExecuteReaderAsync();
            while (await r.ReadAsync())
                list.Add(new PushSubscriptionRecord(r.GetString(0), r.GetString(1), r.GetString(2)));
        }
        catch (Exception ex) { logger.LogWarning(ex, "PushService.GetSubscriptions"); }
        return list;
    }

    // ── Obtener todas las suscripciones activas ───────────────────────────────
    public async Task<List<(string UserId, PushSubscriptionRecord Sub)>> GetAllSubscriptionsAsync()
    {
        var list = new List<(string, PushSubscriptionRecord)>();
        try
        {
            await using var conn = new SqlConnection(ConnStr);
            await conn.OpenAsync();
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = "SELECT UserId, Endpoint, P256dh, Auth FROM dbo.PushSubscriptions";
            await using var r = await cmd.ExecuteReaderAsync();
            while (await r.ReadAsync())
                list.Add((r.GetString(0), new PushSubscriptionRecord(r.GetString(1), r.GetString(2), r.GetString(3))));
        }
        catch (Exception ex) { logger.LogWarning(ex, "PushService.GetAllSubscriptions"); }
        return list;
    }

    // ── Enviar notificación a un usuario específico ───────────────────────────
    public async Task SendToUserAsync(string userId, string title, string body, string? url = null, string? icon = null)
    {
        if (string.IsNullOrWhiteSpace(VapidPublic) || string.IsNullOrWhiteSpace(VapidPrivate)) return;
        var subs = await GetSubscriptionsAsync(userId);
        foreach (var sub in subs)
            await SendAsync(sub, title, body, url, icon);
    }

    // ── Enviar notificación por username (para módulos que guardan nombre, no ID) ─
    public async Task SendToUsernameAsync(string username, string title, string body, string? url = null)
    {
        if (string.IsNullOrWhiteSpace(VapidPublic) || string.IsNullOrWhiteSpace(VapidPrivate)) return;
        try
        {
            await using var conn = new SqlConnection(ConnStr);
            await conn.OpenAsync();
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = "SELECT Endpoint, P256dh, Auth FROM dbo.PushSubscriptions WHERE UserName = @UserName";
            cmd.Parameters.AddWithValue("@UserName", username);
            await using var r = await cmd.ExecuteReaderAsync();
            var subs = new List<PushSubscriptionRecord>();
            while (await r.ReadAsync())
                subs.Add(new PushSubscriptionRecord(r.GetString(0), r.GetString(1), r.GetString(2)));
            foreach (var sub in subs)
                await SendAsync(sub, title, body, url);
        }
        catch (Exception ex) { logger.LogWarning(ex, "PushService.SendToUsername"); }
    }

    // ── Enviar notificación a todos ───────────────────────────────────────────
    public async Task SendToAllAsync(string title, string body, string? url = null)
    {
        if (string.IsNullOrWhiteSpace(VapidPublic) || string.IsNullOrWhiteSpace(VapidPrivate)) return;
        var subs = await GetAllSubscriptionsAsync();
        foreach (var (_, sub) in subs)
            await SendAsync(sub, title, body, url);
    }

    // ── Enviar a una suscripción individual ───────────────────────────────────
    private async Task SendAsync(PushSubscriptionRecord sub, string title, string body, string? url, string? icon = null)
    {
        try
        {
            var webPushClient = new WebPushClient();
            webPushClient.SetVapidDetails(VapidSubject, VapidPublic, VapidPrivate);

            var subscription = new PushSubscription(sub.Endpoint, sub.P256dh, sub.Auth);
            var payload = System.Text.Json.JsonSerializer.Serialize(new {
                title,
                body,
                url   = url ?? "/",
                icon  = icon ?? "/vite.svg",
                badge = "/vite.svg",
            });
            await webPushClient.SendNotificationAsync(subscription, payload);
        }
        catch (WebPushException ex) when (ex.StatusCode == System.Net.HttpStatusCode.Gone ||
                                          ex.StatusCode == System.Net.HttpStatusCode.NotFound)
        {
            // Suscripción expirada — limpiar de la BD
            _ = Task.Run(() => RemoveSubscriptionAsync(sub.Endpoint));
        }
        catch (Exception ex) { logger.LogWarning(ex, "PushService.Send failed for {Endpoint}", sub.Endpoint[..Math.Min(30, sub.Endpoint.Length)]); }
    }

    private async Task RemoveSubscriptionAsync(string endpoint)
    {
        try
        {
            await using var conn = new SqlConnection(ConnStr);
            await conn.OpenAsync();
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = "DELETE FROM dbo.PushSubscriptions WHERE Endpoint = @E";
            cmd.Parameters.AddWithValue("@E", endpoint);
            await cmd.ExecuteNonQueryAsync();
        }
        catch { /* best effort */ }
    }
}

public record PushSubscriptionRecord(string Endpoint, string P256dh, string Auth);
