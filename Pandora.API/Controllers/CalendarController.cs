using MailKit.Net.Smtp;
using MailKit.Security;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using MimeKit;
using System.Globalization;
using System.Text.Json;

namespace Pandora.API.Controllers;

[ApiController]
[Route("api/calendar")]
[Authorize]
public class CalendarController(IConfiguration config, ILogger<CalendarController> logger) : ControllerBase
{
    private static readonly TimeZoneInfo MxTz = TimeZoneInfo.FindSystemTimeZoneById(
        OperatingSystem.IsWindows() ? "Central Standard Time" : "America/Mexico_City");

    private SqlConnection Conn() => new(config.GetConnectionString("PandoraDb"));

    private async Task EnsureAttendeesColumnAsync(SqlConnection conn, CancellationToken ct)
    {
        await using var cmd = conn.CreateCommand();
        cmd.CommandText = """
            IF NOT EXISTS (
                SELECT 1 FROM sys.columns
                WHERE object_id = OBJECT_ID('dbo.Reservations') AND name = 'AttendeesJson')
            ALTER TABLE dbo.Reservations ADD AttendeesJson NVARCHAR(MAX) NULL
            """;
        await cmd.ExecuteNonQueryAsync(ct);
    }

    // ════════════════════════════════════════════════════════════════════════
    //  ROOMS
    // ════════════════════════════════════════════════════════════════════════

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
            await EnsureAttendeesColumnAsync(conn, ct);
            await using var cmd = conn.CreateCommand();

            var sql = """
                SELECT r.Id, r.Title, r.Description, r.RoomId,
                       r.StartTime, r.EndTime,
                       r.OrganizerName, r.OrganizerEmail, r.MeetLink,
                       r.IsRecurring, r.RecurrenceRule, r.ParentReservationId,
                       r.CreatedAt, r.AttendeesJson,
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

    [HttpGet("reservations/{id:guid}")]
    public async Task<IActionResult> GetReservationById(Guid id, CancellationToken ct)
    {
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await EnsureAttendeesColumnAsync(conn, ct);
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = """
                SELECT r.Id, r.Title, r.Description, r.RoomId,
                       r.StartTime, r.EndTime,
                       r.OrganizerName, r.OrganizerEmail, r.MeetLink,
                       r.IsRecurring, r.RecurrenceRule, r.ParentReservationId,
                       r.CreatedAt, r.AttendeesJson,
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
            await EnsureAttendeesColumnAsync(conn, ct);

            // Paso 1 — nombre de sala
            string roomName;
            using (var roomCmd = conn.CreateCommand())
            {
                roomCmd.CommandText = "SELECT Name FROM dbo.Rooms WHERE Id = @Id AND IsActive = 1";
                roomCmd.Parameters.AddWithValue("@Id", dto.RoomId);
                var nameObj = await roomCmd.ExecuteScalarAsync(ct);
                if (nameObj is null || nameObj is DBNull) return BadRequest("Sala no encontrada o inactiva.");
                roomName = (string)nameObj;
            }

            // Paso 2 — verificar conflicto de horario
            int conflicts;
            using (var conflictCmd = conn.CreateCommand())
            {
                conflictCmd.CommandText = """
                    SELECT COUNT(1) FROM dbo.Reservations
                    WHERE RoomId = @RoomId AND StartTime < @End AND EndTime > @Start
                    """;
                conflictCmd.Parameters.AddWithValue("@RoomId", dto.RoomId);
                conflictCmd.Parameters.AddWithValue("@Start",  dto.StartTime);
                conflictCmd.Parameters.AddWithValue("@End",    dto.EndTime);
                conflicts = (int)(await conflictCmd.ExecuteScalarAsync(ct))!;
            }

            if (conflicts > 0)
                return Conflict($"La sala '{roomName}' ya está reservada en ese horario. Por favor elige otro horario o sala.");

            var attendeesJson = dto.Attendees?.Count > 0
                ? JsonSerializer.Serialize(dto.Attendees)
                : null;

            await using var cmd = conn.CreateCommand();
            cmd.CommandText = """
                INSERT INTO dbo.Reservations
                    (Id, Title, Description, RoomId, StartTime, EndTime,
                     OrganizerName, OrganizerEmail, MeetLink,
                     IsRecurring, RecurrenceRule, ParentReservationId, AttendeesJson, CreatedAt)
                VALUES
                    (@Id, @Title, @Desc, @RoomId, @Start, @End,
                     @OrgName, @OrgEmail, @MeetLink,
                     @IsRecurring, @RRule, @ParentId, @AttendeesJson, GETUTCDATE())
                """;
            cmd.Parameters.AddWithValue("@Id",            id);
            cmd.Parameters.AddWithValue("@Title",         dto.Title.Trim());
            cmd.Parameters.AddWithValue("@Desc",          (object?)dto.Description         ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@RoomId",        dto.RoomId);
            cmd.Parameters.AddWithValue("@Start",         dto.StartTime);
            cmd.Parameters.AddWithValue("@End",           dto.EndTime);
            cmd.Parameters.AddWithValue("@OrgName",       (object?)dto.OrganizerName       ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@OrgEmail",      (object?)dto.OrganizerEmail      ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@MeetLink",      (object?)dto.MeetLink            ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@IsRecurring",   dto.IsRecurring);
            cmd.Parameters.AddWithValue("@RRule",         (object?)dto.RecurrenceRule      ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@ParentId",      (object?)dto.ParentReservationId ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@AttendeesJson", (object?)attendeesJson           ?? DBNull.Value);
            await cmd.ExecuteNonQueryAsync(ct);

            _ = Task.Run(() => SendReservationEmailsAsync(dto, id, roomName, isUpdate: false), CancellationToken.None);

            return Ok(new { id });
        }
        catch (Exception ex) { logger.LogError(ex, "CreateReservation"); return StatusCode(500, ex.Message); }
    }

    [HttpPut("reservations/{id:guid}")]
    public async Task<IActionResult> UpdateReservation(Guid id, [FromBody] ReservationDto dto, CancellationToken ct)
    {
        if (string.IsNullOrWhiteSpace(dto.Title)) return BadRequest("Título requerido.");
        if (dto.StartTime >= dto.EndTime)          return BadRequest("La hora de inicio debe ser antes del fin.");
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await EnsureAttendeesColumnAsync(conn, ct);

            // Paso 1 — nombre de sala
            string roomName;
            using (var roomCmd = conn.CreateCommand())
            {
                roomCmd.CommandText = "SELECT Name FROM dbo.Rooms WHERE Id = @RoomId AND IsActive = 1";
                roomCmd.Parameters.AddWithValue("@RoomId", dto.RoomId);
                var nameObj = await roomCmd.ExecuteScalarAsync(ct);
                if (nameObj is null || nameObj is DBNull) return BadRequest("Sala no encontrada o inactiva.");
                roomName = (string)nameObj;
            }

            // Paso 2 — verificar conflicto excluyendo la reservación actual
            int conflicts;
            using (var conflictCmd = conn.CreateCommand())
            {
                conflictCmd.CommandText = """
                    SELECT COUNT(1) FROM dbo.Reservations
                    WHERE RoomId = @RoomId AND StartTime < @End AND EndTime > @Start AND Id <> @Id
                    """;
                conflictCmd.Parameters.AddWithValue("@RoomId", dto.RoomId);
                conflictCmd.Parameters.AddWithValue("@Start",  dto.StartTime);
                conflictCmd.Parameters.AddWithValue("@End",    dto.EndTime);
                conflictCmd.Parameters.AddWithValue("@Id",     id);
                conflicts = (int)(await conflictCmd.ExecuteScalarAsync(ct))!;
            }

            if (conflicts > 0)
                return Conflict($"La sala '{roomName}' ya está reservada en ese horario. Por favor elige otro horario o sala.");

            var attendeesJson = dto.Attendees?.Count > 0
                ? JsonSerializer.Serialize(dto.Attendees)
                : null;

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
                    AttendeesJson     = @AttendeesJson,
                    UpdatedAt         = GETUTCDATE()
                WHERE Id = @Id
                """;
            cmd.Parameters.AddWithValue("@Id",            id);
            cmd.Parameters.AddWithValue("@Title",         dto.Title.Trim());
            cmd.Parameters.AddWithValue("@Desc",          (object?)dto.Description    ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@RoomId",        dto.RoomId);
            cmd.Parameters.AddWithValue("@Start",         dto.StartTime);
            cmd.Parameters.AddWithValue("@End",           dto.EndTime);
            cmd.Parameters.AddWithValue("@OrgName",       (object?)dto.OrganizerName  ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@OrgEmail",      (object?)dto.OrganizerEmail ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@MeetLink",      (object?)dto.MeetLink       ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@IsRecurring",   dto.IsRecurring);
            cmd.Parameters.AddWithValue("@RRule",         (object?)dto.RecurrenceRule ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@AttendeesJson", (object?)attendeesJson      ?? DBNull.Value);
            int rows = await cmd.ExecuteNonQueryAsync(ct);
            if (rows == 0) return NotFound("Reservación no encontrada.");

            _ = Task.Run(() => SendReservationEmailsAsync(dto, id, roomName, isUpdate: true), CancellationToken.None);

            return Ok(new { id });
        }
        catch (Exception ex) { logger.LogError(ex, "UpdateReservation {Id}", id); return StatusCode(500, ex.Message); }
    }

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

    // ── Helpers ───────────────────────────────────────────────────────────────

    private static object ReadReservation(SqlDataReader r)
    {
        object? attendees = null;
        try
        {
            var ord = r.GetOrdinal("AttendeesJson");
            if (!r.IsDBNull(ord))
                attendees = JsonSerializer.Deserialize<List<object>>(r.GetString(ord));
        }
        catch { /* columna aún no migrada o JSON malformado */ }

        return new
        {
            id                  = r.GetGuid(r.GetOrdinal("Id")),
            title               = r.GetString(r.GetOrdinal("Title")),
            description         = r.IsDBNull(r.GetOrdinal("Description"))        ? null : r.GetString(r.GetOrdinal("Description")),
            roomId              = r.GetGuid(r.GetOrdinal("RoomId")),
            startTime           = DateTime.SpecifyKind(r.GetDateTime(r.GetOrdinal("StartTime")), DateTimeKind.Utc),
            endTime             = DateTime.SpecifyKind(r.GetDateTime(r.GetOrdinal("EndTime")),   DateTimeKind.Utc),
            organizerName       = r.IsDBNull(r.GetOrdinal("OrganizerName"))      ? null : r.GetString(r.GetOrdinal("OrganizerName")),
            organizerEmail      = r.IsDBNull(r.GetOrdinal("OrganizerEmail"))     ? null : r.GetString(r.GetOrdinal("OrganizerEmail")),
            meetLink            = r.IsDBNull(r.GetOrdinal("MeetLink"))           ? null : r.GetString(r.GetOrdinal("MeetLink")),
            isRecurring         = r.GetBoolean(r.GetOrdinal("IsRecurring")),
            recurrenceRule      = r.IsDBNull(r.GetOrdinal("RecurrenceRule"))     ? null : r.GetString(r.GetOrdinal("RecurrenceRule")),
            parentReservationId = r.IsDBNull(r.GetOrdinal("ParentReservationId"))? (Guid?)null : r.GetGuid(r.GetOrdinal("ParentReservationId")),
            createdAt           = DateTime.SpecifyKind(r.GetDateTime(r.GetOrdinal("CreatedAt")), DateTimeKind.Utc),
            roomName            = r.IsDBNull(r.GetOrdinal("RoomName"))           ? null : r.GetString(r.GetOrdinal("RoomName")),
            roomColor           = r.IsDBNull(r.GetOrdinal("RoomColor"))          ? null : r.GetString(r.GetOrdinal("RoomColor")),
            roomCapacity        = r.IsDBNull(r.GetOrdinal("RoomCapacity"))       ? 0    : r.GetInt32(r.GetOrdinal("RoomCapacity")),
            attendees,
        };
    }

    private async Task SendReservationEmailsAsync(
        ReservationDto dto, Guid reservationId, string roomName, bool isUpdate)
    {
        try
        {
            var smtp     = config.GetSection("SmtpSettings");
            var host     = smtp["Host"]     ?? "";
            var port     = int.TryParse(smtp["Port"], out var p) ? p : 587;
            var from     = smtp["FromEmail"] ?? "";
            var pass     = smtp["Password"]  ?? "";
            var fromName = smtp["FromName"]  ?? "Pandora";

            if (string.IsNullOrWhiteSpace(host) || string.IsNullOrWhiteSpace(from) || string.IsNullOrWhiteSpace(pass))
                return;

            // Lista de destinatarios: organizador + asistentes
            var recipients = new List<(string Name, string Email)>();
            if (!string.IsNullOrWhiteSpace(dto.OrganizerEmail))
                recipients.Add((dto.OrganizerName ?? "Organizador", dto.OrganizerEmail!));
            if (dto.Attendees != null)
                foreach (var a in dto.Attendees.Where(a => !string.IsNullOrWhiteSpace(a.Email)))
                    recipients.Add((a.Name, a.Email));

            if (recipients.Count == 0) return;

            // Convertir a hora de México para el correo
            var mxStart  = TimeZoneInfo.ConvertTimeFromUtc(dto.StartTime, MxTz);
            var mxEnd    = TimeZoneInfo.ConvertTimeFromUtc(dto.EndTime,   MxTz);
            var dateStr  = mxStart.ToString("dddd, dd 'de' MMMM 'de' yyyy", new CultureInfo("es-MX"));
            var startStr = mxStart.ToString("HH:mm");
            var endStr   = mxEnd.ToString("HH:mm");

            var subject = isUpdate
                ? $"[iMET] Reservación actualizada: {dto.Title}"
                : $"[iMET] Nueva reservación de sala: {dto.Title}";

            using var smtpClient = new SmtpClient();
            await smtpClient.ConnectAsync(host, port, SecureSocketOptions.StartTls);
            await smtpClient.AuthenticateAsync(from, pass);

            foreach (var (name, email) in recipients)
            {
                var body = BuildEmailBody(
                    recipientName: name,
                    title:         dto.Title,
                    roomName:      roomName,
                    dateStr:       dateStr,
                    startStr:      startStr,
                    endStr:        endStr,
                    organizerName: dto.OrganizerName ?? "—",
                    meetLink:      dto.MeetLink,
                    description:   dto.Description,
                    isUpdate:      isUpdate,
                    reservationId: reservationId
                );

                var msg = new MimeMessage();
                msg.From.Add(new MailboxAddress(fromName, from));
                msg.To.Add(new MailboxAddress(name, email));
                msg.Subject = subject;
                msg.Body    = new TextPart("html") { Text = body };

                await smtpClient.SendAsync(msg);
            }

            await smtpClient.DisconnectAsync(true);
            logger.LogInformation("Reservation emails sent for {Id} to {Count} recipients", reservationId, recipients.Count);
        }
        catch (Exception ex)
        {
            logger.LogWarning("Reservation email failed for {Id}: {Msg}", reservationId, ex.Message);
        }
    }

    private static string BuildEmailBody(
        string recipientName, string title, string roomName,
        string dateStr, string startStr, string endStr,
        string organizerName, string? meetLink, string? description,
        bool isUpdate, Guid reservationId)
    {
        var headerText  = isUpdate ? "Reservación actualizada" : "Nueva reservación de sala";
        var headerColor = isUpdate ? "#0d47a1" : "#1a237e";

        var meetRow = string.IsNullOrWhiteSpace(meetLink) ? "" : $"""
              <tr style="border-bottom:1px solid #eee">
                <td style="padding:8px;font-weight:bold;color:#555;width:40%">Link de reunión</td>
                <td style="padding:8px"><a href="{meetLink}" style="color:#1a237e">{meetLink}</a></td>
              </tr>
            """;

        var descRow = string.IsNullOrWhiteSpace(description) ? "" : $"""
              <tr style="border-bottom:1px solid #eee">
                <td style="padding:8px;font-weight:bold;color:#555">Descripción</td>
                <td style="padding:8px">{description}</td>
              </tr>
            """;

        return $"""
            <html><body style="font-family:Arial,sans-serif;font-size:14px;color:#333">
            <div style="max-width:600px;margin:0 auto">
              <div style="background:{headerColor};padding:20px;border-radius:8px 8px 0 0">
                <h2 style="color:white;margin:0">📅 {headerText}</h2>
              </div>
              <div style="border:1px solid #ddd;padding:24px;border-radius:0 0 8px 8px">
                <p>Hola <strong>{recipientName}</strong>,</p>
                <p>Te informamos sobre la siguiente reservación de sala en el sistema Pandora de iMET:</p>
                <table style="width:100%;border-collapse:collapse">
                  <tr style="border-bottom:1px solid #eee">
                    <td style="padding:8px;font-weight:bold;color:#555;width:40%">Evento</td>
                    <td style="padding:8px"><strong>{title}</strong></td>
                  </tr>
                  <tr style="border-bottom:1px solid #eee">
                    <td style="padding:8px;font-weight:bold;color:#555">Sala</td>
                    <td style="padding:8px">{roomName}</td>
                  </tr>
                  <tr style="border-bottom:1px solid #eee">
                    <td style="padding:8px;font-weight:bold;color:#555">Fecha</td>
                    <td style="padding:8px">{dateStr}</td>
                  </tr>
                  <tr style="border-bottom:1px solid #eee">
                    <td style="padding:8px;font-weight:bold;color:#555">Horario</td>
                    <td style="padding:8px">{startStr} – {endStr} hrs (hora de México)</td>
                  </tr>
                  <tr style="border-bottom:1px solid #eee">
                    <td style="padding:8px;font-weight:bold;color:#555">Organizador</td>
                    <td style="padding:8px">{organizerName}</td>
                  </tr>
                  {meetRow}
                  {descRow}
                </table>
                <p style="margin-top:20px;font-size:12px;color:#999;text-align:center">
                  Pandora — Sistema de Gestión iMET · Reservación #{reservationId.ToString()[..8]}
                </p>
              </div>
            </div>
            </body></html>
            """;
    }
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

public record AttendeeDto(string Name, string Email, bool IsExternal);

public record ReservationDto(
    string              Title,
    string?             Description,
    Guid                RoomId,
    DateTime            StartTime,
    DateTime            EndTime,
    string?             OrganizerName,
    string?             OrganizerEmail,
    string?             MeetLink,
    bool                IsRecurring,
    string?             RecurrenceRule,
    Guid?               ParentReservationId,
    List<AttendeeDto>?  Attendees
);
