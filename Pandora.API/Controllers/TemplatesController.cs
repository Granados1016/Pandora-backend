using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using System.Text.Json;

namespace Pandora.API.Controllers;

[ApiController]
[Route("api/templates")]
[Authorize]
public class TemplatesController(
    IConfiguration config,
    ILogger<TemplatesController> logger) : ControllerBase
{
    // ── Helpers ──────────────────────────────────────────────────────────────

    private SqlConnection CreateConn()
        => new(config.GetConnectionString("PandoraDb"));

    /// <summary>Crea la tabla Templates si no existe.</summary>
    private static async Task EnsureTableAsync(SqlConnection conn, CancellationToken ct = default)
    {
        const string sql = """
            IF NOT EXISTS (
                SELECT 1 FROM sys.objects
                WHERE object_id = OBJECT_ID(N'dbo.Templates') AND type = N'U')
            BEGIN
                CREATE TABLE dbo.Templates (
                    Id          UNIQUEIDENTIFIER NOT NULL DEFAULT NEWID() PRIMARY KEY,
                    Name        NVARCHAR(200)    NOT NULL,
                    Subject     NVARCHAR(500)    NOT NULL,
                    Body        NVARCHAR(MAX)    NOT NULL,
                    ProgramType INT              NOT NULL DEFAULT 1,
                    IsPlainText BIT              NOT NULL DEFAULT 1,
                    IsActive    BIT              NOT NULL DEFAULT 1,
                    Variables   NVARCHAR(2000)   NULL,     -- JSON array: ["var1","var2"]
                    CreatedAt   DATETIME2        NOT NULL DEFAULT GETUTCDATE(),
                    UpdatedAt   DATETIME2        NULL
                );
            END
            """;
        await using var cmd = conn.CreateCommand();
        cmd.CommandText = sql;
        await cmd.ExecuteNonQueryAsync(ct);
    }

    private static List<string> DeserializeVars(object? raw)
    {
        if (raw is null or DBNull) return [];
        try { return JsonSerializer.Deserialize<List<string>>(raw.ToString()!) ?? []; }
        catch { return []; }
    }

    private static string? SerializeVars(List<string>? vars)
        => vars is { Count: > 0 }
            ? JsonSerializer.Serialize(vars)
            : null;

    // ── GET /api/templates ────────────────────────────────────────────────────

    [HttpGet]
    public async Task<IActionResult> GetAll(CancellationToken ct = default)
    {
        try
        {
            await using var conn = CreateConn();
            await conn.OpenAsync(ct);
            await EnsureTableAsync(conn, ct);

            await using var cmd = conn.CreateCommand();
            cmd.CommandText = """
                SELECT Id, Name, Subject, Body, ProgramType,
                       IsPlainText, IsActive, Variables, CreatedAt
                FROM   dbo.Templates
                ORDER  BY CreatedAt DESC
                """;

            var list = new List<object>();
            await using var r = await cmd.ExecuteReaderAsync(ct);
            while (await r.ReadAsync(ct))
            {
                list.Add(new
                {
                    id          = r.GetGuid(r.GetOrdinal("Id")),
                    name        = r.GetString(r.GetOrdinal("Name")),
                    subject     = r.GetString(r.GetOrdinal("Subject")),
                    body        = r.GetString(r.GetOrdinal("Body")),
                    programType = r.GetInt32(r.GetOrdinal("ProgramType")),
                    isPlainText = r.GetBoolean(r.GetOrdinal("IsPlainText")),
                    isActive    = r.GetBoolean(r.GetOrdinal("IsActive")),
                    variables   = DeserializeVars(r.IsDBNull(r.GetOrdinal("Variables"))
                                                    ? null : r.GetString(r.GetOrdinal("Variables"))),
                    createdAt   = r.GetDateTime(r.GetOrdinal("CreatedAt")),
                });
            }
            return Ok(list);
        }
        catch (Exception ex)
        {
            logger.LogError(ex, "GetAll Templates failed");
            return StatusCode(500, $"Error al cargar plantillas: {ex.Message}");
        }
    }

    // ── GET /api/templates/{id} ───────────────────────────────────────────────

    [HttpGet("{id:guid}")]
    public async Task<IActionResult> GetById(Guid id, CancellationToken ct = default)
    {
        try
        {
            await using var conn = CreateConn();
            await conn.OpenAsync(ct);
            await EnsureTableAsync(conn, ct);

            await using var cmd = conn.CreateCommand();
            cmd.CommandText = """
                SELECT Id, Name, Subject, Body, ProgramType,
                       IsPlainText, IsActive, Variables, CreatedAt
                FROM   dbo.Templates
                WHERE  Id = @Id
                """;
            cmd.Parameters.AddWithValue("@Id", id);

            await using var r = await cmd.ExecuteReaderAsync(ct);
            if (!await r.ReadAsync(ct)) return NotFound("Plantilla no encontrada.");

            return Ok(new
            {
                id          = r.GetGuid(r.GetOrdinal("Id")),
                name        = r.GetString(r.GetOrdinal("Name")),
                subject     = r.GetString(r.GetOrdinal("Subject")),
                body        = r.GetString(r.GetOrdinal("Body")),
                programType = r.GetInt32(r.GetOrdinal("ProgramType")),
                isPlainText = r.GetBoolean(r.GetOrdinal("IsPlainText")),
                isActive    = r.GetBoolean(r.GetOrdinal("IsActive")),
                variables   = DeserializeVars(r.IsDBNull(r.GetOrdinal("Variables"))
                                                ? null : r.GetString(r.GetOrdinal("Variables"))),
                createdAt   = r.GetDateTime(r.GetOrdinal("CreatedAt")),
            });
        }
        catch (Exception ex)
        {
            logger.LogError(ex, "GetById Template {Id} failed", id);
            return StatusCode(500, ex.Message);
        }
    }

    // ── POST /api/templates ───────────────────────────────────────────────────

    [HttpPost]
    public async Task<IActionResult> Create([FromBody] TemplateDto dto, CancellationToken ct = default)
    {
        if (string.IsNullOrWhiteSpace(dto.Name))    return BadRequest("El nombre es obligatorio.");
        if (string.IsNullOrWhiteSpace(dto.Subject)) return BadRequest("El asunto es obligatorio.");
        if (string.IsNullOrWhiteSpace(dto.Body))    return BadRequest("El cuerpo es obligatorio.");

        var id = Guid.NewGuid();
        try
        {
            await using var conn = CreateConn();
            await conn.OpenAsync(ct);
            await EnsureTableAsync(conn, ct);

            await using var cmd = conn.CreateCommand();
            cmd.CommandText = """
                INSERT INTO dbo.Templates
                    (Id, Name, Subject, Body, ProgramType, IsPlainText, IsActive, Variables, CreatedAt)
                VALUES
                    (@Id, @Name, @Subject, @Body, @ProgramType, @IsPlainText, @IsActive, @Variables, GETUTCDATE())
                """;

            cmd.Parameters.AddWithValue("@Id",          id);
            cmd.Parameters.AddWithValue("@Name",        dto.Name.Trim());
            cmd.Parameters.AddWithValue("@Subject",     dto.Subject.Trim());
            cmd.Parameters.AddWithValue("@Body",        dto.Body);
            cmd.Parameters.AddWithValue("@ProgramType", dto.ProgramType);
            cmd.Parameters.AddWithValue("@IsPlainText", dto.IsPlainText);
            cmd.Parameters.AddWithValue("@IsActive",    dto.IsActive);
            cmd.Parameters.AddWithValue("@Variables",   (object?)SerializeVars(dto.Variables) ?? DBNull.Value);

            await cmd.ExecuteNonQueryAsync(ct);
            logger.LogInformation("Template created: {Id} — {Name}", id, dto.Name);
            return Ok(new { id });
        }
        catch (Exception ex)
        {
            logger.LogError(ex, "Create Template failed");
            return StatusCode(500, $"Error al guardar la plantilla: {ex.Message}");
        }
    }

    // ── PUT /api/templates/{id} ───────────────────────────────────────────────

    [HttpPut("{id:guid}")]
    public async Task<IActionResult> Update(Guid id, [FromBody] TemplateDto dto, CancellationToken ct = default)
    {
        if (string.IsNullOrWhiteSpace(dto.Name))    return BadRequest("El nombre es obligatorio.");
        if (string.IsNullOrWhiteSpace(dto.Subject)) return BadRequest("El asunto es obligatorio.");
        if (string.IsNullOrWhiteSpace(dto.Body))    return BadRequest("El cuerpo es obligatorio.");

        try
        {
            await using var conn = CreateConn();
            await conn.OpenAsync(ct);
            await EnsureTableAsync(conn, ct);

            await using var cmd = conn.CreateCommand();
            cmd.CommandText = """
                UPDATE dbo.Templates
                SET    Name        = @Name,
                       Subject     = @Subject,
                       Body        = @Body,
                       ProgramType = @ProgramType,
                       IsPlainText = @IsPlainText,
                       IsActive    = @IsActive,
                       Variables   = @Variables,
                       UpdatedAt   = GETUTCDATE()
                WHERE  Id = @Id
                """;

            cmd.Parameters.AddWithValue("@Id",          id);
            cmd.Parameters.AddWithValue("@Name",        dto.Name.Trim());
            cmd.Parameters.AddWithValue("@Subject",     dto.Subject.Trim());
            cmd.Parameters.AddWithValue("@Body",        dto.Body);
            cmd.Parameters.AddWithValue("@ProgramType", dto.ProgramType);
            cmd.Parameters.AddWithValue("@IsPlainText", dto.IsPlainText);
            cmd.Parameters.AddWithValue("@IsActive",    dto.IsActive);
            cmd.Parameters.AddWithValue("@Variables",   (object?)SerializeVars(dto.Variables) ?? DBNull.Value);

            int rows = await cmd.ExecuteNonQueryAsync(ct);
            if (rows == 0) return NotFound("Plantilla no encontrada.");

            logger.LogInformation("Template updated: {Id} — {Name}", id, dto.Name);
            return Ok(new { id });
        }
        catch (Exception ex)
        {
            logger.LogError(ex, "Update Template {Id} failed", id);
            return StatusCode(500, $"Error al actualizar la plantilla: {ex.Message}");
        }
    }

    // ── DELETE /api/templates/{id} ────────────────────────────────────────────

    [HttpDelete("{id:guid}")]
    public async Task<IActionResult> Delete(Guid id, CancellationToken ct = default)
    {
        try
        {
            await using var conn = CreateConn();
            await conn.OpenAsync(ct);

            await using var cmd = conn.CreateCommand();
            cmd.CommandText = "DELETE FROM dbo.Templates WHERE Id = @Id";
            cmd.Parameters.AddWithValue("@Id", id);

            int rows = await cmd.ExecuteNonQueryAsync(ct);
            if (rows == 0) return NotFound("Plantilla no encontrada.");

            logger.LogInformation("Template deleted: {Id}", id);
            return NoContent();
        }
        catch (Exception ex)
        {
            logger.LogError(ex, "Delete Template {Id} failed", id);
            return StatusCode(500, ex.Message);
        }
    }
}

// ── DTO ───────────────────────────────────────────────────────────────────────

public record TemplateDto(
    string       Name,
    string       Subject,
    string       Body,
    int          ProgramType,
    bool         IsPlainText,
    bool         IsActive,
    List<string> Variables
);
