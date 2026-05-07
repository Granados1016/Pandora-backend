using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.SignalR;
using Microsoft.Data.SqlClient;
using Pandora.API.Hubs;
using System.Security.Claims;

namespace Pandora.API.Controllers;

[ApiController]
[Route("api/notifications")]
[Authorize]
public class NotificationsController(
    IConfiguration config,
    ILogger<NotificationsController> logger,
    IHubContext<NotificationsHub> hub) : ControllerBase
{
    private SqlConnection Conn() => new(config.GetConnectionString("PandoraDb"));
    private string? CurrentUser() => User.FindFirstValue(ClaimTypes.Name) ?? User.FindFirstValue("name");

    // ── GET /api/notifications ────────────────────────────────────────────────
    /// <summary>Devuelve las últimas 50 notificaciones para el usuario actual.</summary>
    [HttpGet]
    public async Task<IActionResult> GetAll(CancellationToken ct)
    {
        var user = CurrentUser();
        if (user is null) return Unauthorized();

        var items = new List<object>();

        await using var conn = Conn();
        await conn.OpenAsync(ct);
        await using var cmd = conn.CreateCommand();
        cmd.CommandText = """
            SELECT TOP 50
                Id, Title, Message, Type, IsRead, TargetUser, CreatedBy, CreatedAt
            FROM dbo.Notifications
            WHERE (TargetUser = @User OR TargetUser IS NULL)
              AND IsDeleted = 0
            ORDER BY CreatedAt DESC
            """;
        cmd.Parameters.AddWithValue("@User", user);

        await using var reader = await cmd.ExecuteReaderAsync(ct);
        while (await reader.ReadAsync(ct))
        {
            items.Add(new
            {
                id        = reader.GetInt32(0),
                title     = reader.GetString(1),
                message   = reader.GetString(2),
                type      = reader.GetString(3),
                isRead    = reader.GetBoolean(4),
                targetUser = reader.IsDBNull(5) ? null : reader.GetString(5),
                createdBy = reader.GetString(6),
                createdAt = reader.GetDateTime(7),
            });
        }

        return Ok(items);
    }

    // ── GET /api/notifications/unread-count ───────────────────────────────────
    [HttpGet("unread-count")]
    public async Task<IActionResult> UnreadCount(CancellationToken ct)
    {
        var user = CurrentUser();
        if (user is null) return Unauthorized();

        await using var conn = Conn();
        await conn.OpenAsync(ct);
        await using var cmd = conn.CreateCommand();
        cmd.CommandText = """
            SELECT COUNT(*)
            FROM dbo.Notifications
            WHERE (TargetUser = @User OR TargetUser IS NULL)
              AND IsRead = 0
              AND IsDeleted = 0
            """;
        cmd.Parameters.AddWithValue("@User", user);

        var count = (int)(await cmd.ExecuteScalarAsync(ct))!;
        return Ok(new { count });
    }

    // ── PATCH /api/notifications/{id}/read ────────────────────────────────────
    [HttpPatch("{id:int}/read")]
    public async Task<IActionResult> MarkRead(int id, CancellationToken ct)
    {
        var user = CurrentUser();
        if (user is null) return Unauthorized();

        await using var conn = Conn();
        await conn.OpenAsync(ct);
        await using var cmd = conn.CreateCommand();
        cmd.CommandText = """
            UPDATE dbo.Notifications
            SET IsRead = 1
            WHERE Id = @Id
              AND (TargetUser = @User OR TargetUser IS NULL)
              AND IsDeleted = 0
            """;
        cmd.Parameters.AddWithValue("@Id",   id);
        cmd.Parameters.AddWithValue("@User", user);
        await cmd.ExecuteNonQueryAsync(ct);

        return NoContent();
    }

    // ── PATCH /api/notifications/read-all ────────────────────────────────────
    [HttpPatch("read-all")]
    public async Task<IActionResult> MarkAllRead(CancellationToken ct)
    {
        var user = CurrentUser();
        if (user is null) return Unauthorized();

        await using var conn = Conn();
        await conn.OpenAsync(ct);
        await using var cmd = conn.CreateCommand();
        cmd.CommandText = """
            UPDATE dbo.Notifications
            SET IsRead = 1
            WHERE (TargetUser = @User OR TargetUser IS NULL)
              AND IsRead = 0
              AND IsDeleted = 0
            """;
        cmd.Parameters.AddWithValue("@User", user);
        await cmd.ExecuteNonQueryAsync(ct);

        return NoContent();
    }

    // ── POST /api/notifications ───────────────────────────────────────────────
    [HttpPost]
    [Authorize(Roles = "Admin")]
    public async Task<IActionResult> Create([FromBody] NotificationDto dto, CancellationToken ct)
    {
        var creator = CurrentUser() ?? "system";

        int newId;
        await using (var conn = Conn())
        {
            await conn.OpenAsync(ct);
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = """
                INSERT INTO dbo.Notifications
                    (Title, Message, Type, IsRead, TargetUser, CreatedBy, CreatedAt, IsDeleted)
                OUTPUT INSERTED.Id
                VALUES (@Title, @Message, @Type, 0, @TargetUser, @CreatedBy, GETUTCDATE(), 0)
                """;
            cmd.Parameters.AddWithValue("@Title",      dto.Title);
            cmd.Parameters.AddWithValue("@Message",    dto.Message);
            cmd.Parameters.AddWithValue("@Type",       dto.Type);
            cmd.Parameters.AddWithValue("@TargetUser", (object?)dto.TargetUser ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@CreatedBy",  creator);
            newId = (int)(await cmd.ExecuteScalarAsync(ct))!;
        }

        // SignalR
        var notif = new { id = newId, title = dto.Title, message = dto.Message, type = dto.Type };
        if (string.IsNullOrWhiteSpace(dto.TargetUser))
            await hub.Clients.Group("broadcast").SendAsync("NewNotification", notif, ct);
        else
            await hub.Clients.Group($"user-{dto.TargetUser}").SendAsync("NewNotification", notif, ct);

        return Ok(new { id = newId });
    }

    // ── DELETE /api/notifications/{id} ───────────────────────────────────────
    [HttpDelete("{id:int}")]
    [Authorize(Roles = "Admin")]
    public async Task<IActionResult> Delete(int id, CancellationToken ct)
    {
        await using var conn = Conn();
        await conn.OpenAsync(ct);
        await using var cmd = conn.CreateCommand();
        cmd.CommandText = """
            UPDATE dbo.Notifications
            SET IsDeleted = 1
            WHERE Id = @Id AND IsDeleted = 0
            """;
        cmd.Parameters.AddWithValue("@Id", id);
        var rows = await cmd.ExecuteNonQueryAsync(ct);

        return rows > 0 ? NoContent() : NotFound();
    }
}

public record NotificationDto(
    string Title,
    string Message,
    string Type = "info",
    string? TargetUser = null);
