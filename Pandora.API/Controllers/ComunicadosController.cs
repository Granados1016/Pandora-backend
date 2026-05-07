using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using System.Security.Claims;

namespace Pandora.API.Controllers;

[ApiController]
[Route("api/comunicados")]
[Authorize]
public class ComunicadosController(
    IConfiguration config,
    ILogger<ComunicadosController> logger) : ControllerBase
{
    private SqlConnection Conn() => new(config.GetConnectionString("PandoraDb"));
    private string? CurrentUser() => User.FindFirstValue(ClaimTypes.Name) ?? User.FindFirstValue("name");

    // ── GET /api/comunicados ──────────────────────────────────────────────────
    [HttpGet]
    [AllowAnonymous]
    public async Task<IActionResult> GetAll(
        [FromQuery] int page     = 1,
        [FromQuery] int pageSize = 10,
        CancellationToken ct     = default)
    {
        if (page < 1) page = 1;
        if (pageSize < 1) pageSize = 10;

        int total;
        var items = new List<object>();

        await using var conn = Conn();
        await conn.OpenAsync(ct);

        // Total count
        await using (var countCmd = conn.CreateCommand())
        {
            countCmd.CommandText = """
                SELECT COUNT(*)
                FROM dbo.Comunicados
                WHERE IsDeleted = 0
                  AND IsPublished = 1
                  AND (ExpiresAt IS NULL OR ExpiresAt > GETUTCDATE())
                """;
            total = (int)(await countCmd.ExecuteScalarAsync(ct))!;
        }

        // Paged items
        await using var cmd = conn.CreateCommand();
        cmd.CommandText = """
            SELECT Id, Title, Content, Priority, Author, CreatedAt, ExpiresAt
            FROM dbo.Comunicados
            WHERE IsDeleted = 0
              AND IsPublished = 1
              AND (ExpiresAt IS NULL OR ExpiresAt > GETUTCDATE())
            ORDER BY CreatedAt DESC
            OFFSET @Offset ROWS FETCH NEXT @PageSize ROWS ONLY
            """;
        cmd.Parameters.AddWithValue("@Offset",   (page - 1) * pageSize);
        cmd.Parameters.AddWithValue("@PageSize", pageSize);

        await using var reader = await cmd.ExecuteReaderAsync(ct);
        while (await reader.ReadAsync(ct))
        {
            items.Add(new
            {
                id        = reader.GetInt32(0),
                title     = reader.GetString(1),
                content   = reader.GetString(2),
                priority  = reader.GetString(3),
                author    = reader.GetString(4),
                createdAt = reader.GetDateTime(5),
                expiresAt = reader.IsDBNull(6) ? (DateTime?)null : reader.GetDateTime(6),
            });
        }

        var totalPages = (int)Math.Ceiling(total / (double)pageSize);
        return Ok(new { items, total, page, pageSize, totalPages });
    }

    // ── GET /api/comunicados/{id} ─────────────────────────────────────────────
    [HttpGet("{id:int}")]
    public async Task<IActionResult> GetOne(int id, CancellationToken ct)
    {
        await using var conn = Conn();
        await conn.OpenAsync(ct);
        await using var cmd = conn.CreateCommand();
        cmd.CommandText = """
            SELECT Id, Title, Content, Priority, Author, IsPublished, ExpiresAt, CreatedAt
            FROM dbo.Comunicados
            WHERE Id = @Id AND IsDeleted = 0
            """;
        cmd.Parameters.AddWithValue("@Id", id);

        await using var reader = await cmd.ExecuteReaderAsync(ct);
        if (!await reader.ReadAsync(ct)) return NotFound();

        return Ok(new
        {
            id          = reader.GetInt32(0),
            title       = reader.GetString(1),
            content     = reader.GetString(2),
            priority    = reader.GetString(3),
            author      = reader.GetString(4),
            isPublished = reader.GetBoolean(5),
            expiresAt   = reader.IsDBNull(6) ? (DateTime?)null : reader.GetDateTime(6),
            createdAt   = reader.GetDateTime(7),
        });
    }

    // ── POST /api/comunicados ─────────────────────────────────────────────────
    [HttpPost]
    [Authorize(Roles = "Admin")]
    public async Task<IActionResult> Create([FromBody] ComunicadoDto dto, CancellationToken ct)
    {
        var author = CurrentUser() ?? "system";

        await using var conn = Conn();
        await conn.OpenAsync(ct);
        await using var cmd = conn.CreateCommand();
        cmd.CommandText = """
            INSERT INTO dbo.Comunicados
                (Title, Content, Priority, Author, IsPublished, ExpiresAt, IsDeleted)
            OUTPUT INSERTED.Id
            VALUES (@Title, @Content, @Priority, @Author, @IsPublished, @ExpiresAt, 0)
            """;
        cmd.Parameters.AddWithValue("@Title",       dto.Title);
        cmd.Parameters.AddWithValue("@Content",     dto.Content);
        cmd.Parameters.AddWithValue("@Priority",    dto.Priority);
        cmd.Parameters.AddWithValue("@Author",      author);
        cmd.Parameters.AddWithValue("@IsPublished", dto.IsPublished);
        cmd.Parameters.AddWithValue("@ExpiresAt",   (object?)dto.ExpiresAt ?? DBNull.Value);

        var newId = (int)(await cmd.ExecuteScalarAsync(ct))!;
        return Ok(new { id = newId });
    }

    // ── PUT /api/comunicados/{id} ─────────────────────────────────────────────
    [HttpPut("{id:int}")]
    [Authorize(Roles = "Admin")]
    public async Task<IActionResult> Update(int id, [FromBody] ComunicadoDto dto, CancellationToken ct)
    {
        await using var conn = Conn();
        await conn.OpenAsync(ct);
        await using var cmd = conn.CreateCommand();
        cmd.CommandText = """
            UPDATE dbo.Comunicados
            SET Title       = @Title,
                Content     = @Content,
                Priority    = @Priority,
                IsPublished = @IsPublished,
                ExpiresAt   = @ExpiresAt
            WHERE Id = @Id AND IsDeleted = 0
            """;
        cmd.Parameters.AddWithValue("@Title",       dto.Title);
        cmd.Parameters.AddWithValue("@Content",     dto.Content);
        cmd.Parameters.AddWithValue("@Priority",    dto.Priority);
        cmd.Parameters.AddWithValue("@IsPublished", dto.IsPublished);
        cmd.Parameters.AddWithValue("@ExpiresAt",   (object?)dto.ExpiresAt ?? DBNull.Value);
        cmd.Parameters.AddWithValue("@Id",          id);

        var rows = await cmd.ExecuteNonQueryAsync(ct);
        return rows > 0 ? NoContent() : NotFound();
    }

    // ── DELETE /api/comunicados/{id} ──────────────────────────────────────────
    [HttpDelete("{id:int}")]
    [Authorize(Roles = "Admin")]
    public async Task<IActionResult> Delete(int id, CancellationToken ct)
    {
        await using var conn = Conn();
        await conn.OpenAsync(ct);
        await using var cmd = conn.CreateCommand();
        cmd.CommandText = """
            UPDATE dbo.Comunicados
            SET IsDeleted = 1
            WHERE Id = @Id AND IsDeleted = 0
            """;
        cmd.Parameters.AddWithValue("@Id", id);
        var rows = await cmd.ExecuteNonQueryAsync(ct);

        return rows > 0 ? NoContent() : NotFound();
    }
}

public record ComunicadoDto(
    string Title,
    string Content,
    string Priority    = "normal",
    bool IsPublished   = true,
    DateTime? ExpiresAt = null);
