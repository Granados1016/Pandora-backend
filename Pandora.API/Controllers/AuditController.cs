using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;

namespace Pandora.API.Controllers;

[ApiController]
[Route("api/audit")]
[Authorize(Roles = "Admin")]
public class AuditController(IConfiguration config) : ControllerBase
{
    private SqlConnection Conn() => new(config.GetConnectionString("PandoraDb"));

    // ── GET /api/audit ────────────────────────────────────────────────────────
    [HttpGet]
    public async Task<IActionResult> GetAll(
        [FromQuery] string? search   = null,
        [FromQuery] string? module   = null,
        [FromQuery] int     page     = 1,
        [FromQuery] int     pageSize = 50,
        CancellationToken   ct       = default)
    {
        if (page < 1) page = 1;
        if (pageSize < 1) pageSize = 50;

        var hasSearch = !string.IsNullOrWhiteSpace(search);
        var hasModule = !string.IsNullOrWhiteSpace(module);

        var where = new System.Text.StringBuilder("WHERE 1=1");
        if (hasSearch) where.Append(" AND (UserName LIKE @Search OR Action LIKE @Search OR Module LIKE @Search OR Details LIKE @Search)");
        if (hasModule) where.Append(" AND Module = @Module");

        int total;
        var items = new List<object>();

        await using var conn = Conn();
        await conn.OpenAsync(ct);

        // Total count
        await using (var countCmd = conn.CreateCommand())
        {
            countCmd.CommandText = $"SELECT COUNT(*) FROM dbo.AuditLogs {where}";
            if (hasSearch) countCmd.Parameters.AddWithValue("@Search", $"%{search}%");
            if (hasModule) countCmd.Parameters.AddWithValue("@Module", module!);
            total = (int)(await countCmd.ExecuteScalarAsync(ct))!;
        }

        // Paged items
        await using var cmd = conn.CreateCommand();
        cmd.CommandText = $"""
            SELECT Id, UserName, Action, Module, EntityId, Details, IpAddress, CreatedAt
            FROM dbo.AuditLogs
            {where}
            ORDER BY CreatedAt DESC
            OFFSET @Offset ROWS FETCH NEXT @PageSize ROWS ONLY
            """;
        if (hasSearch) cmd.Parameters.AddWithValue("@Search", $"%{search}%");
        if (hasModule) cmd.Parameters.AddWithValue("@Module", module!);
        cmd.Parameters.AddWithValue("@Offset",   (page - 1) * pageSize);
        cmd.Parameters.AddWithValue("@PageSize", pageSize);

        await using var reader = await cmd.ExecuteReaderAsync(ct);
        while (await reader.ReadAsync(ct))
        {
            items.Add(new
            {
                id        = reader.GetInt32(0),
                userName  = reader.GetString(1),
                action    = reader.GetString(2),
                module    = reader.GetString(3),
                entityId  = reader.IsDBNull(4) ? null : reader.GetString(4),
                details   = reader.IsDBNull(5) ? null : reader.GetString(5),
                ipAddress = reader.IsDBNull(6) ? null : reader.GetString(6),
                createdAt = reader.GetDateTime(7),
            });
        }

        var totalPages = (int)Math.Ceiling(total / (double)pageSize);
        return Ok(new { items, total, page, pageSize, totalPages });
    }

    // ── GET /api/audit/modules ────────────────────────────────────────────────
    [HttpGet("modules")]
    public async Task<IActionResult> GetModules(CancellationToken ct)
    {
        var modules = new List<string>();

        await using var conn = Conn();
        await conn.OpenAsync(ct);
        await using var cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT DISTINCT Module FROM dbo.AuditLogs ORDER BY Module";

        await using var reader = await cmd.ExecuteReaderAsync(ct);
        while (await reader.ReadAsync(ct))
            modules.Add(reader.GetString(0));

        return Ok(modules);
    }
}

// ── AuditHelper ──────────────────────────────────────────────────────────────
public static class AuditHelper
{
    public static async Task Log(
        string  connStr,
        string  userName,
        string  action,
        string  module,
        string? entityId = null,
        string? details  = null,
        string? ip       = null)
    {
        try
        {
            await using var conn = new Microsoft.Data.SqlClient.SqlConnection(connStr);
            await conn.OpenAsync();
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = """
                INSERT INTO dbo.AuditLogs (UserName, Action, Module, EntityId, Details, IpAddress)
                VALUES (@User, @Action, @Module, @EntityId, @Details, @Ip)
                """;
            cmd.Parameters.AddWithValue("@User",     userName);
            cmd.Parameters.AddWithValue("@Action",   action);
            cmd.Parameters.AddWithValue("@Module",   module);
            cmd.Parameters.AddWithValue("@EntityId", (object?)entityId ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@Details",  (object?)details  ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@Ip",       (object?)ip       ?? DBNull.Value);
            await cmd.ExecuteNonQueryAsync();
        }
        catch { /* Audit no debe romper el flujo principal */ }
    }
}
