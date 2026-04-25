using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;

namespace Pandora.API.Controllers;

[ApiController]
[Route("api/calendar")]
[Authorize]
public class CalendarController(IConfiguration config, ILogger<CalendarController> logger) : ControllerBase
{
    private SqlConnection Conn() => new(config.GetConnectionString("PandoraDb"));

    // ════════════════════════════════════════════════════════════════════════
    //  ROOMS
    // ════════════════════════════════════════════════════════════════════════

    // ── GET /api/calendar/rooms ───────────────────────────────────────────────
    [HttpGet("rooms")]
    public async Task<IActionResult> GetRooms(CancellationToken ct)
    {
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = """
                SELECT Id, Name, Capacity, Location, Color, Description, IsActive, CreatedAt
                FROM dbo.Rooms ORDER BY Name
                """;
            var list = new List<object>();
            await using var r = await cmd.ExecuteReaderAsync(ct);
            while (await r.ReadAsync(ct))
                list.Add(new
                {
                    id          = r.GetGuid(r.GetOrdinal("Id")),
                    name        = r.GetString(r.GetOrdinal("Name")),
                    capacity    = r.GetInt32(r.GetOrdinal("Capacity")),
                    location    = r.IsDBNull(r.GetOrdinal("Location"))    ? null : r.GetString(r.GetOrdinal("Location")),
                    color       = r.IsDBNull(r.GetOrdinal("Color"))       ? null : r.GetString(r.GetOrdinal("Color")),
                    description = r.IsDBNull(r.GetOrdinal("Description")) ? null : r.GetString(r.GetOrdinal("Description")),
                    isActive    = r.GetBoolean(r.GetOrdinal("IsActive")),
                    createdAt   = r.GetDateTime(r.GetOrdinal("CreatedAt")),
                });
            return Ok(list);
        }
        catch (Exception ex) { logger.LogError(ex, "GetRooms"); return StatusCode(500, ex.Message); }
    }

    // ── POST /api/calendar/rooms ──────────────────────────────────────────────
    [HttpPost("rooms")]
    public async Task<IActionResult> CreateRoom([FromBody] RoomDto dto, CancellationToken ct)
    {
        if (string.IsNullOrWhiteSpace(dto.Name)) return BadRequest("Nombre requerido.");
        var id = Guid.NewGuid();
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = """
                INSERT INTO dbo.Rooms (Id, Name, Capacity, Location, Color, Description, IsActive, CreatedAt)
                VALUES (@Id, @Name, @Capacity, @Location, @Color, @Desc, @IsActive, GETUTCDATE())
                """;
            cmd.Parameters.AddWithValue("@Id",       id);
            cmd.Parameters.AddWithValue("@Name",     dto.Name.Trim());
            cmd.Parameters.AddWithValue("@Capacity", dto.Capacity);
            cmd.Parameters.AddWithValue("@Location", (object?)dto.Location    ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@Color",    (object?)dto.Color       ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@Desc",     (object?)dto.Description ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@IsActive", dto.IsActive);
            await cmd.ExecuteNonQueryAsync(ct);
            return Ok(new { id });
        }
        catch (Exception ex) { logger.LogError(ex, "CreateRoom"); return StatusCode(500, ex.Message); }
    }

    // ── PUT /api/calendar/rooms/{id} ──────────────────────────────────────────
    [HttpPut("rooms/{id:guid}")]
    public async Task<IActionResult> UpdateRoom(Guid id, [FromBody] RoomDto dto, CancellationToken ct)
    {
        if (string.IsNullOrWhiteSpace(dto.Name)) return BadRequest("Nombre requerido.");
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = """
                UPDATE dbo.Rooms
                SET Name = @Name, Capacity = @Capacity, Location = @Location,
                    Color = @Color, Description = @Desc, IsActive = @IsActive, UpdatedAt = GETUTCDATE()
                WHERE Id = @Id
                """;
            cmd.Parameters.AddWithValue("@Id",       id);
            cmd.Parameters.AddWithValue("@Name",     dto.Name.Trim());
            cmd.Parameters.AddWithValue("@Capacity", dto.Capacity);
            cmd.Parameters.AddWithValue("@Location", (object?)dto.Location    ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@Color",    (object?)dto.Color       ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@Desc",     (object?)dto.Description ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@IsActive", dto.IsActive);
            int rows = await cmd.ExecuteNonQueryAsync(ct);
            if (rows == 0) return NotFound("Sala no encontrada.");
            return Ok(new { id });
        }
        catch (Exception ex) { logger.LogError(ex, "UpdateRoom {Id}", id); return StatusCode(500, ex.Message); }
    }

    // ── DELETE /api/calendar/rooms/{id} ───────────────────────────────────────
    [HttpDelete("rooms/{id:guid}")]
    public async Task<IActionResult> DeleteRoom(Guid id, CancellationToken ct)
    {
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = "UPDATE dbo.Rooms SET IsActive = 0, UpdatedAt = GETUTCDATE() WHERE Id = @Id";
            cmd.Parameters.AddWithValue("@Id", id);
            int rows = await cmd.ExecuteNonQueryAsync(ct);
            if (rows == 0) return NotFound("Sala no encontrada.");
            return NoContent();
        }
        catch (Exception ex) { logger.LogError(ex, "DeleteRoom {Id}", id); return StatusCode(500, ex.Message); }
    }

    // ════════════════════════════════════════════════════════════════════════
    //  RESERVATIONS
    // ════════════════════════════════════════════════════════════════════════

    // ── GET /api/calendar/reservations ────────────────────────────────────────
    [HttpGet("reservations")]
    public async Task<IActionResult> GetReservations(
        [FromQuery] DateTime? rangeStart,
        [FromQuery] DateTime? rangeEnd,
        [FromQuery] Guid? roomId,
        CancellationToken ct)
    {
        try
        {
            var start = rangeStart ?? DateTime.UtcNow.AddMonths(-1);
            var end   = rangeEnd   ?? DateTime.UtcNow.AddMonths(2);

            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await using var cmd = conn.CreateCommand();

            var sql = """
                SELECT r.Id, r.Title, r.Description, r.RoomId,
                       r.StartTime, r.EndTime,
                       r.OrganizerName, r.OrganizerEmail, r.MeetLink,
                       r.IsRecurring, r.RecurrenceRule, r.ParentReservationId,
                       r.CreatedAt,
                       rm.Name AS RoomName, rm.Color AS RoomColor, rm.Capacity AS RoomCapacity
                FROM dbo.Reservations r
                INNER JOIN dbo.Rooms rm ON r.RoomId = rm.Id
                WHERE r.StartTime < @End AND r.EndTime > @Start
                """;
            if (roomId.HasValue) sql += " AND r.RoomId = @RoomId";
            sql += " ORDER BY r.StartTime";

            cmd.CommandText = sql;
            cmd.Parameters.AddWithValue("@Start", start);
            cmd.Parameters.AddWithValue("@End",   end);
            if (roomId.HasValue) cmd.Parameters.AddWithValue("@RoomId", roomId.Value);

            var list = new List<object>();
            await using var r = await cmd.ExecuteReaderAsync(ct);
            while (await r.ReadAsync(ct))
                list.Add(ReadReservation(r));
            return Ok(list);
        }
        catch (Exception ex) { logger.LogError(ex, "GetReservations"); return StatusCode(500, ex.Message); }
    }

    // ── GET /api/calendar/reservations/check-conflict ─────────────────────────
    [HttpGet("reservations/check-conflict")]
    public async Task<IActionResult> CheckConflict(
        [FromQuery] Guid roomId,
        [FromQuery] DateTime start,
        [FromQuery] DateTime end,
        [FromQuery] Guid? excludeId,
        CancellationToken ct)
    {
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await using var cmd = conn.CreateCommand();
            var sql = """
                SELECT COUNT(1) FROM dbo.Reservations
                WHERE RoomId = @RoomId AND StartTime < @End AND EndTime > @Start
                """;
            if (excludeId.HasValue) sql += " AND Id <> @ExcludeId";
            cmd.CommandText = sql;
            cmd.Parameters.AddWithValue("@RoomId", roomId);
            cmd.Parameters.AddWithValue("@Start",  start);
            cmd.Parameters.AddWithValue("@End",    end);
            if (excludeId.HasValue) cmd.Parameters.AddWithValue("@ExcludeId", excludeId.Value);
            var count = (int)await cmd.ExecuteScalarAsync(ct)!;
            return Ok(new { hasConflict = count > 0 });
        }
        catch (Exception ex) { logger.LogError(ex, "CheckConflict"); return StatusCode(500, ex.Message); }
    }

    // ── GET /api/calendar/reservations/{id} ───────────────────────────────────
    [HttpGet("reservations/{id:guid}")]
    public async Task<IActionResult> GetReservationById(Guid id, CancellationToken ct)
    {
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = """
                SELECT r.Id, r.Title, r.Description, r.RoomId,
                       r.StartTime, r.EndTime,
                       r.OrganizerName, r.OrganizerEmail, r.MeetLink,
                       r.IsRecurring, r.RecurrenceRule, r.ParentReservationId,
                       r.CreatedAt,
                       rm.Name AS RoomName, rm.Color AS RoomColor, rm.Capacity AS RoomCapacity
                FROM dbo.Reservations r
                INNER JOIN dbo.Rooms rm ON r.RoomId = rm.Id
                WHERE r.Id = @Id
                """;
            cmd.Parameters.AddWithValue("@Id", id);
            await using var r = await cmd.ExecuteReaderAsync(ct);
            if (!await r.ReadAsync(ct)) return NotFound("Reservación no encontrada.");
            return Ok(ReadReservation(r));
        }
        catch (Exception ex) { logger.LogError(ex, "GetReservationById {Id}", id); return StatusCode(500, ex.Message); }
    }

    // ── POST /api/calendar/reservations ───────────────────────────────────────
    [HttpPost("reservations")]
    public async Task<IActionResult> CreateReservation([FromBody] ReservationDto dto, CancellationToken ct)
    {
        if (string.IsNullOrWhiteSpace(dto.Title))  return BadRequest("Título requerido.");
        if (dto.RoomId == Guid.Empty)              return BadRequest("Sala requerida.");
        if (dto.StartTime >= dto.EndTime)          return BadRequest("La hora de inicio debe ser antes del fin.");

        var id = Guid.NewGuid();
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = """
                INSERT INTO dbo.Reservations
                    (Id, Title, Description, RoomId, StartTime, EndTime,
                     OrganizerName, OrganizerEmail, MeetLink,
                     IsRecurring, RecurrenceRule, ParentReservationId, CreatedAt)
                VALUES
                    (@Id, @Title, @Desc, @RoomId, @Start, @End,
                     @OrgName, @OrgEmail, @MeetLink,
                     @IsRecurring, @RRule, @ParentId, GETUTCDATE())
                """;
            cmd.Parameters.AddWithValue("@Id",          id);
            cmd.Parameters.AddWithValue("@Title",       dto.Title.Trim());
            cmd.Parameters.AddWithValue("@Desc",        (object?)dto.Description       ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@RoomId",      dto.RoomId);
            cmd.Parameters.AddWithValue("@Start",       dto.StartTime);
            cmd.Parameters.AddWithValue("@End",         dto.EndTime);
            cmd.Parameters.AddWithValue("@OrgName",     (object?)dto.OrganizerName  ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@OrgEmail",    (object?)dto.OrganizerEmail ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@MeetLink",    (object?)dto.MeetLink       ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@IsRecurring", dto.IsRecurring);
            cmd.Parameters.AddWithValue("@RRule",       (object?)dto.RecurrenceRule         ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@ParentId",    (object?)dto.ParentReservationId    ?? DBNull.Value);
            await cmd.ExecuteNonQueryAsync(ct);
            return Ok(new { id });
        }
        catch (Exception ex) { logger.LogError(ex, "CreateReservation"); return StatusCode(500, ex.Message); }
    }

    // ── PUT /api/calendar/reservations/{id} ───────────────────────────────────
    [HttpPut("reservations/{id:guid}")]
    public async Task<IActionResult> UpdateReservation(Guid id, [FromBody] ReservationDto dto, CancellationToken ct)
    {
        if (string.IsNullOrWhiteSpace(dto.Title)) return BadRequest("Título requerido.");
        if (dto.StartTime >= dto.EndTime)          return BadRequest("La hora de inicio debe ser antes del fin.");
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = """
                UPDATE dbo.Reservations
                SET Title             = @Title,
                    Description       = @Desc,
                    RoomId            = @RoomId,
                    StartTime         = @Start,
                    EndTime           = @End,
                    OrganizerName     = @OrgName,
                    OrganizerEmail    = @OrgEmail,
                    MeetLink          = @MeetLink,
                    IsRecurring       = @IsRecurring,
                    RecurrenceRule    = @RRule,
                    UpdatedAt         = GETUTCDATE()
                WHERE Id = @Id
                """;
            cmd.Parameters.AddWithValue("@Id",          id);
            cmd.Parameters.AddWithValue("@Title",       dto.Title.Trim());
            cmd.Parameters.AddWithValue("@Desc",        (object?)dto.Description       ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@RoomId",      dto.RoomId);
            cmd.Parameters.AddWithValue("@Start",       dto.StartTime);
            cmd.Parameters.AddWithValue("@End",         dto.EndTime);
            cmd.Parameters.AddWithValue("@OrgName",     (object?)dto.OrganizerName  ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@OrgEmail",    (object?)dto.OrganizerEmail ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@MeetLink",    (object?)dto.MeetLink       ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@IsRecurring", dto.IsRecurring);
            cmd.Parameters.AddWithValue("@RRule",       (object?)dto.RecurrenceRule ?? DBNull.Value);
            int rows = await cmd.ExecuteNonQueryAsync(ct);
            if (rows == 0) return NotFound("Reservación no encontrada.");
            return Ok(new { id });
        }
        catch (Exception ex) { logger.LogError(ex, "UpdateReservation {Id}", id); return StatusCode(500, ex.Message); }
    }

    // ── DELETE /api/calendar/reservations/{id} ────────────────────────────────
    [HttpDelete("reservations/{id:guid}")]
    public async Task<IActionResult> DeleteReservation(
        Guid id, [FromQuery] bool deleteAll = false, CancellationToken ct = default)
    {
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await using var cmd = conn.CreateCommand();

            if (deleteAll)
            {
                // Eliminar todas las instancias de la recurrencia
                cmd.CommandText = """
                    DELETE FROM dbo.Reservations
                    WHERE Id = @Id OR ParentReservationId = @Id
                    """;
            }
            else
            {
                cmd.CommandText = "DELETE FROM dbo.Reservations WHERE Id = @Id";
            }
            cmd.Parameters.AddWithValue("@Id", id);
            int rows = await cmd.ExecuteNonQueryAsync(ct);
            if (rows == 0) return NotFound("Reservación no encontrada.");
            return NoContent();
        }
        catch (Exception ex) { logger.LogError(ex, "DeleteReservation {Id}", id); return StatusCode(500, ex.Message); }
    }

    // ── Helper ────────────────────────────────────────────────────────────────
    private static object ReadReservation(SqlDataReader r) => new
    {
        id                   = r.GetGuid(r.GetOrdinal("Id")),
        title                = r.GetString(r.GetOrdinal("Title")),
        description          = r.IsDBNull(r.GetOrdinal("Description"))          ? null : r.GetString(r.GetOrdinal("Description")),
        roomId               = r.GetGuid(r.GetOrdinal("RoomId")),
        startTime            = r.GetDateTime(r.GetOrdinal("StartTime")),
        endTime              = r.GetDateTime(r.GetOrdinal("EndTime")),
        organizerName        = r.IsDBNull(r.GetOrdinal("OrganizerName"))         ? null : r.GetString(r.GetOrdinal("OrganizerName")),
        organizerEmail       = r.IsDBNull(r.GetOrdinal("OrganizerEmail"))        ? null : r.GetString(r.GetOrdinal("OrganizerEmail")),
        meetLink             = r.IsDBNull(r.GetOrdinal("MeetLink"))              ? null : r.GetString(r.GetOrdinal("MeetLink")),
        isRecurring          = r.GetBoolean(r.GetOrdinal("IsRecurring")),
        recurrenceRule       = r.IsDBNull(r.GetOrdinal("RecurrenceRule"))        ? null : r.GetString(r.GetOrdinal("RecurrenceRule")),
        parentReservationId  = r.IsDBNull(r.GetOrdinal("ParentReservationId"))   ? (Guid?)null : r.GetGuid(r.GetOrdinal("ParentReservationId")),
        createdAt            = r.GetDateTime(r.GetOrdinal("CreatedAt")),
        roomName             = r.IsDBNull(r.GetOrdinal("RoomName"))              ? null : r.GetString(r.GetOrdinal("RoomName")),
        roomColor            = r.IsDBNull(r.GetOrdinal("RoomColor"))             ? null : r.GetString(r.GetOrdinal("RoomColor")),
        roomCapacity         = r.IsDBNull(r.GetOrdinal("RoomCapacity"))          ? 0    : r.GetInt32(r.GetOrdinal("RoomCapacity")),
    };
}

// ── DTOs ──────────────────────────────────────────────────────────────────────
public record RoomDto(
    string  Name,
    int     Capacity,
    string? Location,
    string? Color,
    string? Description,
    bool    IsActive
);

public record ReservationDto(
    string    Title,
    string?   Description,
    Guid      RoomId,
    DateTime  StartTime,
    DateTime  EndTime,
    string?   OrganizerName,
    string?   OrganizerEmail,
    string?   MeetLink,
    bool      IsRecurring,
    string?   RecurrenceRule,
    Guid?     ParentReservationId
);
