using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using System.Security.Claims;

namespace Pandora.API.Controllers;

/// <summary>
/// Módulo de Capacitaciones — CRUD de cursos/entrenamientos del personal.
/// </summary>
[ApiController]
[Route("api/capacitaciones")]
[Authorize]
public class CapacitacionesController(
    IConfiguration config,
    ILogger<CapacitacionesController> logger) : ControllerBase
{
    private SqlConnection Conn() => new(config.GetConnectionString("PandoraDb"));
    private bool IsAdmin => User.IsInRole("Admin");
    private string CurrentUser => User.FindFirstValue(ClaimTypes.Name) ?? "Desconocido";

    // ── Ensure tables exist ───────────────────────────────────────────────────
    private static async Task EnsureTablesAsync(SqlConnection conn, CancellationToken ct)
    {
        await using var cmd = conn.CreateCommand();
        cmd.CommandText = """
            IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'Capacitaciones')
            BEGIN
                CREATE TABLE dbo.Capacitaciones (
                    Id          UNIQUEIDENTIFIER NOT NULL DEFAULT NEWID() PRIMARY KEY,
                    Title       NVARCHAR(200)    NOT NULL,
                    Description NVARCHAR(MAX)    NULL,
                    Category    NVARCHAR(100)    NULL,
                    Instructor  NVARCHAR(200)    NULL,
                    StartDate   DATE             NOT NULL,
                    EndDate     DATE             NULL,
                    Duration    INT              NULL,   -- horas
                    Status      NVARCHAR(50)     NOT NULL DEFAULT 'Pendiente',
                    Modality    NVARCHAR(50)     NULL,   -- Presencial / Virtual / Híbrida
                    MaxParticipants INT          NULL,
                    CreatedBy   NVARCHAR(100)    NOT NULL,
                    CreatedAt   DATETIME2        NOT NULL DEFAULT GETUTCDATE(),
                    UpdatedAt   DATETIME2        NULL,
                    IsDeleted   BIT              NOT NULL DEFAULT 0
                )
            END;
            IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'CapacitacionParticipantes')
            BEGIN
                CREATE TABLE dbo.CapacitacionParticipantes (
                    Id              UNIQUEIDENTIFIER NOT NULL DEFAULT NEWID() PRIMARY KEY,
                    CapacitacionId  UNIQUEIDENTIFIER NOT NULL,
                    Username        NVARCHAR(100)    NOT NULL,
                    FullName        NVARCHAR(200)    NULL,
                    Status          NVARCHAR(50)     NOT NULL DEFAULT 'Inscrito',   -- Inscrito / Completado / Cancelado
                    CompletedAt     DATETIME2        NULL,
                    Score           DECIMAL(5,2)     NULL,
                    Notes           NVARCHAR(500)    NULL,
                    CreatedAt       DATETIME2        NOT NULL DEFAULT GETUTCDATE()
                )
            END;
            """;
        await cmd.ExecuteNonQueryAsync(ct);
    }

    // ── GET /api/capacitaciones ───────────────────────────────────────────────
    [HttpGet]
    public async Task<IActionResult> GetAll(
        [FromQuery] string? status   = null,
        [FromQuery] string? category = null,
        [FromQuery] string? search   = null,
        CancellationToken ct = default)
    {
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await EnsureTablesAsync(conn, ct);

            var where = new List<string> { "c.IsDeleted = 0" };
            if (!string.IsNullOrWhiteSpace(status))   where.Add("c.Status = @Status");
            if (!string.IsNullOrWhiteSpace(category)) where.Add("c.Category = @Category");
            if (!string.IsNullOrWhiteSpace(search))   where.Add("(c.Title LIKE @Q OR c.Instructor LIKE @Q OR c.Description LIKE @Q)");

            await using var cmd = conn.CreateCommand();
            cmd.CommandText = $"""
                SELECT c.Id, c.Title, c.Description, c.Category, c.Instructor,
                       c.StartDate, c.EndDate, c.Duration, c.Status, c.Modality,
                       c.MaxParticipants, c.CreatedBy, c.CreatedAt,
                       (SELECT COUNT(*) FROM dbo.CapacitacionParticipantes p
                        WHERE p.CapacitacionId = c.Id AND p.Status != 'Cancelado') AS EnrolledCount
                FROM dbo.Capacitaciones c
                WHERE {string.Join(" AND ", where)}
                ORDER BY c.StartDate DESC
                """;
            if (!string.IsNullOrWhiteSpace(status))   cmd.Parameters.AddWithValue("@Status",   status);
            if (!string.IsNullOrWhiteSpace(category)) cmd.Parameters.AddWithValue("@Category", category);
            if (!string.IsNullOrWhiteSpace(search))   cmd.Parameters.AddWithValue("@Q", $"%{search}%");

            var list = new List<object>();
            await using var r = await cmd.ExecuteReaderAsync(ct);
            while (await r.ReadAsync(ct))
                list.Add(Map(r));
            return Ok(list);
        }
        catch (Exception ex) { logger.LogError(ex, "GetAll Capacitaciones"); return StatusCode(500, ex.Message); }
    }

    // ── GET /api/capacitaciones/{id} ─────────────────────────────────────────
    [HttpGet("{id:guid}")]
    public async Task<IActionResult> GetById(Guid id, CancellationToken ct)
    {
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = """
                SELECT c.Id, c.Title, c.Description, c.Category, c.Instructor,
                       c.StartDate, c.EndDate, c.Duration, c.Status, c.Modality,
                       c.MaxParticipants, c.CreatedBy, c.CreatedAt,
                       (SELECT COUNT(*) FROM dbo.CapacitacionParticipantes p
                        WHERE p.CapacitacionId = c.Id AND p.Status != 'Cancelado') AS EnrolledCount
                FROM dbo.Capacitaciones c
                WHERE c.Id = @Id AND c.IsDeleted = 0
                """;
            cmd.Parameters.AddWithValue("@Id", id);
            await using var r = await cmd.ExecuteReaderAsync(ct);
            if (!await r.ReadAsync(ct)) return NotFound();
            var cap = Map(r);
            await r.CloseAsync();

            // Participantes
            await using var pCmd = conn.CreateCommand();
            pCmd.CommandText = """
                SELECT Id, Username, FullName, Status, CompletedAt, Score, Notes, CreatedAt
                FROM dbo.CapacitacionParticipantes
                WHERE CapacitacionId = @Id
                ORDER BY FullName
                """;
            pCmd.Parameters.AddWithValue("@Id", id);
            await using var pr = await pCmd.ExecuteReaderAsync(ct);
            var parts = new List<object>();
            while (await pr.ReadAsync(ct))
                parts.Add(new {
                    id          = pr.GetGuid(0),
                    username    = pr.GetString(1),
                    fullName    = pr.IsDBNull(2) ? null : pr.GetString(2),
                    status      = pr.GetString(3),
                    completedAt = pr.IsDBNull(4) ? (DateTime?)null : pr.GetDateTime(4),
                    score       = pr.IsDBNull(5) ? (decimal?)null : pr.GetDecimal(5),
                    notes       = pr.IsDBNull(6) ? null : pr.GetString(6),
                    createdAt   = pr.GetDateTime(7),
                });

            return Ok(new { capacitacion = cap, participantes = parts });
        }
        catch (Exception ex) { logger.LogError(ex, "GetById Capacitacion {Id}", id); return StatusCode(500, ex.Message); }
    }

    // ── POST /api/capacitaciones ─────────────────────────────────────────────
    [HttpPost]
    [Authorize(Roles = "Admin")]
    public async Task<IActionResult> Create([FromBody] CapacitacionDto dto, CancellationToken ct)
    {
        if (string.IsNullOrWhiteSpace(dto.Title)) return BadRequest("Título requerido.");
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await EnsureTablesAsync(conn, ct);
            var newId = Guid.NewGuid();
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = """
                INSERT INTO dbo.Capacitaciones
                    (Id, Title, Description, Category, Instructor, StartDate, EndDate, Duration, Status, Modality, MaxParticipants, CreatedBy, CreatedAt)
                VALUES
                    (@Id, @Title, @Desc, @Cat, @Instr, @Start, @End, @Dur, @Status, @Mod, @Max, @CreatedBy, GETUTCDATE())
                """;
            cmd.Parameters.AddWithValue("@Id",        newId);
            cmd.Parameters.AddWithValue("@Title",     dto.Title.Trim());
            cmd.Parameters.AddWithValue("@Desc",      string.IsNullOrWhiteSpace(dto.Description)  ? DBNull.Value : (object)dto.Description);
            cmd.Parameters.AddWithValue("@Cat",       string.IsNullOrWhiteSpace(dto.Category)     ? DBNull.Value : (object)dto.Category);
            cmd.Parameters.AddWithValue("@Instr",     string.IsNullOrWhiteSpace(dto.Instructor)   ? DBNull.Value : (object)dto.Instructor);
            cmd.Parameters.AddWithValue("@Start",     dto.StartDate);
            cmd.Parameters.AddWithValue("@End",       dto.EndDate.HasValue ? (object)dto.EndDate.Value : DBNull.Value);
            cmd.Parameters.AddWithValue("@Dur",       dto.Duration.HasValue ? (object)dto.Duration.Value : DBNull.Value);
            cmd.Parameters.AddWithValue("@Status",    string.IsNullOrWhiteSpace(dto.Status) ? "Pendiente" : dto.Status);
            cmd.Parameters.AddWithValue("@Mod",       string.IsNullOrWhiteSpace(dto.Modality)     ? DBNull.Value : (object)dto.Modality);
            cmd.Parameters.AddWithValue("@Max",       dto.MaxParticipants.HasValue ? (object)dto.MaxParticipants.Value : DBNull.Value);
            cmd.Parameters.AddWithValue("@CreatedBy", CurrentUser);
            await cmd.ExecuteNonQueryAsync(ct);
            return CreatedAtAction(nameof(GetById), new { id = newId }, new { id = newId });
        }
        catch (Exception ex) { logger.LogError(ex, "Create Capacitacion"); return StatusCode(500, ex.Message); }
    }

    // ── PUT /api/capacitaciones/{id} ─────────────────────────────────────────
    [HttpPut("{id:guid}")]
    [Authorize(Roles = "Admin")]
    public async Task<IActionResult> Update(Guid id, [FromBody] CapacitacionDto dto, CancellationToken ct)
    {
        if (string.IsNullOrWhiteSpace(dto.Title)) return BadRequest("Título requerido.");
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = """
                UPDATE dbo.Capacitaciones SET
                    Title = @Title, Description = @Desc, Category = @Cat, Instructor = @Instr,
                    StartDate = @Start, EndDate = @End, Duration = @Dur, Status = @Status,
                    Modality = @Mod, MaxParticipants = @Max, UpdatedAt = GETUTCDATE()
                WHERE Id = @Id AND IsDeleted = 0
                """;
            cmd.Parameters.AddWithValue("@Id",    id);
            cmd.Parameters.AddWithValue("@Title", dto.Title.Trim());
            cmd.Parameters.AddWithValue("@Desc",  string.IsNullOrWhiteSpace(dto.Description)  ? DBNull.Value : (object)dto.Description);
            cmd.Parameters.AddWithValue("@Cat",   string.IsNullOrWhiteSpace(dto.Category)     ? DBNull.Value : (object)dto.Category);
            cmd.Parameters.AddWithValue("@Instr", string.IsNullOrWhiteSpace(dto.Instructor)   ? DBNull.Value : (object)dto.Instructor);
            cmd.Parameters.AddWithValue("@Start", dto.StartDate);
            cmd.Parameters.AddWithValue("@End",   dto.EndDate.HasValue ? (object)dto.EndDate.Value : DBNull.Value);
            cmd.Parameters.AddWithValue("@Dur",   dto.Duration.HasValue ? (object)dto.Duration.Value : DBNull.Value);
            cmd.Parameters.AddWithValue("@Status",string.IsNullOrWhiteSpace(dto.Status) ? "Pendiente" : dto.Status);
            cmd.Parameters.AddWithValue("@Mod",   string.IsNullOrWhiteSpace(dto.Modality)     ? DBNull.Value : (object)dto.Modality);
            cmd.Parameters.AddWithValue("@Max",   dto.MaxParticipants.HasValue ? (object)dto.MaxParticipants.Value : DBNull.Value);
            int rows = await cmd.ExecuteNonQueryAsync(ct);
            if (rows == 0) return NotFound();
            return NoContent();
        }
        catch (Exception ex) { logger.LogError(ex, "Update Capacitacion {Id}", id); return StatusCode(500, ex.Message); }
    }

    // ── DELETE /api/capacitaciones/{id} ──────────────────────────────────────
    [HttpDelete("{id:guid}")]
    [Authorize(Roles = "Admin")]
    public async Task<IActionResult> Delete(Guid id, CancellationToken ct)
    {
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = "UPDATE dbo.Capacitaciones SET IsDeleted=1, UpdatedAt=GETUTCDATE() WHERE Id=@Id AND IsDeleted=0";
            cmd.Parameters.AddWithValue("@Id", id);
            int rows = await cmd.ExecuteNonQueryAsync(ct);
            if (rows == 0) return NotFound();
            return NoContent();
        }
        catch (Exception ex) { logger.LogError(ex, "Delete Capacitacion {Id}", id); return StatusCode(500, ex.Message); }
    }

    // ── POST /api/capacitaciones/{id}/participantes ───────────────────────────
    [HttpPost("{id:guid}/participantes")]
    [Authorize(Roles = "Admin")]
    public async Task<IActionResult> AddParticipant(Guid id, [FromBody] ParticipanteDto dto, CancellationToken ct)
    {
        if (string.IsNullOrWhiteSpace(dto.Username)) return BadRequest("Username requerido.");
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = """
                IF NOT EXISTS (SELECT 1 FROM dbo.CapacitacionParticipantes
                               WHERE CapacitacionId=@CapId AND Username=@User AND Status!='Cancelado')
                BEGIN
                  INSERT INTO dbo.CapacitacionParticipantes
                    (Id, CapacitacionId, Username, FullName, Status, CreatedAt)
                  VALUES (NEWID(), @CapId, @User, @FN, 'Inscrito', GETUTCDATE())
                END
                """;
            cmd.Parameters.AddWithValue("@CapId", id);
            cmd.Parameters.AddWithValue("@User",  dto.Username.Trim().ToLower());
            cmd.Parameters.AddWithValue("@FN",    string.IsNullOrWhiteSpace(dto.FullName) ? DBNull.Value : (object)dto.FullName);
            await cmd.ExecuteNonQueryAsync(ct);
            return Ok(new { added = true });
        }
        catch (Exception ex) { return StatusCode(500, ex.Message); }
    }

    // ── PUT /api/capacitaciones/{id}/participantes/{partId} ───────────────────
    [HttpPut("{id:guid}/participantes/{partId:guid}")]
    [Authorize(Roles = "Admin")]
    public async Task<IActionResult> UpdateParticipant(Guid id, Guid partId, [FromBody] UpdateParticipanteDto dto, CancellationToken ct)
    {
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = """
                UPDATE dbo.CapacitacionParticipantes
                SET Status=@Status, Score=@Score, Notes=@Notes,
                    CompletedAt = CASE WHEN @Status='Completado' THEN GETUTCDATE() ELSE CompletedAt END
                WHERE Id=@Id AND CapacitacionId=@CapId
                """;
            cmd.Parameters.AddWithValue("@Id",     partId);
            cmd.Parameters.AddWithValue("@CapId",  id);
            cmd.Parameters.AddWithValue("@Status", dto.Status ?? "Inscrito");
            cmd.Parameters.AddWithValue("@Score",  dto.Score.HasValue ? (object)dto.Score.Value : DBNull.Value);
            cmd.Parameters.AddWithValue("@Notes",  string.IsNullOrWhiteSpace(dto.Notes) ? DBNull.Value : (object)dto.Notes);
            int rows = await cmd.ExecuteNonQueryAsync(ct);
            if (rows == 0) return NotFound();
            return NoContent();
        }
        catch (Exception ex) { return StatusCode(500, ex.Message); }
    }

    // ── GET /api/capacitaciones/categorias ────────────────────────────────────
    [HttpGet("categorias")]
    public async Task<IActionResult> GetCategorias(CancellationToken ct)
    {
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await EnsureTablesAsync(conn, ct);
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = "SELECT DISTINCT Category FROM dbo.Capacitaciones WHERE IsDeleted=0 AND Category IS NOT NULL ORDER BY Category";
            var list = new List<string>();
            await using var r = await cmd.ExecuteReaderAsync(ct);
            while (await r.ReadAsync(ct)) list.Add(r.GetString(0));
            return Ok(list);
        }
        catch (Exception ex) { return StatusCode(500, ex.Message); }
    }

    // ── Mapper ────────────────────────────────────────────────────────────────
    private static object Map(SqlDataReader r) => new
    {
        id              = r.GetGuid(0),
        title           = r.GetString(1),
        description     = r.IsDBNull(2)  ? null : r.GetString(2),
        category        = r.IsDBNull(3)  ? null : r.GetString(3),
        instructor      = r.IsDBNull(4)  ? null : r.GetString(4),
        startDate       = r.GetDateTime(5),
        endDate         = r.IsDBNull(6)  ? (DateTime?)null : r.GetDateTime(6),
        duration        = r.IsDBNull(7)  ? (int?)null : r.GetInt32(7),
        status          = r.GetString(8),
        modality        = r.IsDBNull(9)  ? null : r.GetString(9),
        maxParticipants = r.IsDBNull(10) ? (int?)null : r.GetInt32(10),
        createdBy       = r.GetString(11),
        createdAt       = r.GetDateTime(12),
        enrolledCount   = r.GetInt32(13),
    };
}

// ── DTOs ──────────────────────────────────────────────────────────────────────
public record CapacitacionDto(
    string   Title,
    string?  Description,
    string?  Category,
    string?  Instructor,
    DateTime StartDate,
    DateTime? EndDate,
    int?     Duration,
    string?  Status,
    string?  Modality,
    int?     MaxParticipants
);

public record ParticipanteDto(string Username, string? FullName);
public record UpdateParticipanteDto(string? Status, decimal? Score, string? Notes);
