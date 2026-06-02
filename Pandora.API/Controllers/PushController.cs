using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using System.Security.Claims;
using Pandora.API.Services;

namespace Pandora.API.Controllers;

[ApiController]
[Route("api/push")]
[Authorize]
public class PushController(
    IConfiguration config,
    PushService pushService,
    ILogger<PushController> logger) : ControllerBase
{
    private SqlConnection Conn() => new(config.GetConnectionString("PandoraDb"));
    private string CurrentUser   => User.FindFirstValue(ClaimTypes.Name) ?? "sistema";
    private string CurrentUserId => User.FindFirstValue(ClaimTypes.NameIdentifier)
                                 ?? User.FindFirstValue("sub") ?? CurrentUser;

    // ── GET /api/push/vapid-key ───────────────────────────────────────────────
    [HttpGet("vapid-key")]
    [AllowAnonymous]
    public IActionResult GetVapidKey() =>
        Ok(new { publicKey = config["Push:VapidPublicKey"] ?? config["Push__VapidPublicKey"] ?? "" });

    // ── POST /api/push/subscribe ──────────────────────────────────────────────
    [HttpPost("subscribe")]
    public async Task<IActionResult> Subscribe([FromBody] SubscribeDto dto, CancellationToken ct)
    {
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await using var cmd = conn.CreateCommand();
            // INSERT OR UPDATE (por si ya existe)
            cmd.CommandText = """
                IF EXISTS (SELECT 1 FROM dbo.PushSubscriptions WHERE UserId=@UserId AND Endpoint=@Endpoint)
                    UPDATE dbo.PushSubscriptions SET P256dh=@P256dh, Auth=@Auth WHERE UserId=@UserId AND Endpoint=@Endpoint
                ELSE
                    INSERT INTO dbo.PushSubscriptions (UserId, UserName, Endpoint, P256dh, Auth)
                    VALUES (@UserId, @UserName, @Endpoint, @P256dh, @Auth)
                """;
            cmd.Parameters.AddWithValue("@UserId",   CurrentUserId);
            cmd.Parameters.AddWithValue("@UserName", CurrentUser);
            cmd.Parameters.AddWithValue("@Endpoint", dto.Endpoint);
            cmd.Parameters.AddWithValue("@P256dh",   dto.P256dh);
            cmd.Parameters.AddWithValue("@Auth",     dto.Auth);
            await cmd.ExecuteNonQueryAsync(ct);
            return Ok(new { message = "Suscripción registrada." });
        }
        catch (Exception ex) { logger.LogError(ex, "Push.Subscribe"); return StatusCode(500, ex.Message); }
    }

    // ── DELETE /api/push/unsubscribe ──────────────────────────────────────────
    [HttpDelete("unsubscribe")]
    public async Task<IActionResult> Unsubscribe([FromBody] UnsubscribeDto dto, CancellationToken ct)
    {
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = "DELETE FROM dbo.PushSubscriptions WHERE UserId=@UserId AND Endpoint=@Endpoint";
            cmd.Parameters.AddWithValue("@UserId",   CurrentUserId);
            cmd.Parameters.AddWithValue("@Endpoint", dto.Endpoint);
            await cmd.ExecuteNonQueryAsync(ct);
            return NoContent();
        }
        catch (Exception ex) { logger.LogError(ex, "Push.Unsubscribe"); return StatusCode(500, ex.Message); }
    }

    // ── POST /api/push/test (admin) ───────────────────────────────────────────
    [HttpPost("test")]
    [Authorize(Roles = "Admin")]
    public async Task<IActionResult> Test(CancellationToken ct)
    {
        await pushService.SendToUserAsync(CurrentUserId,
            "🔔 Pandora — Prueba de notificación",
            "Las notificaciones push están funcionando correctamente.",
            "/");
        return Ok(new { message = "Notificación enviada." });
    }
}

public record SubscribeDto(string Endpoint, string P256dh, string Auth);
public record UnsubscribeDto(string Endpoint);
