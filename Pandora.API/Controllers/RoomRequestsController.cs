using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using System.Net.Mail;
using System.Net;

namespace Pandora.API.Controllers;

[ApiController]
[Route("api/room-requests")]
[Authorize]
public class RoomRequestsController(
    IConfiguration config,
    ILogger<RoomRequestsController> logger) : ControllerBase
{
    // ── Helpers ──────────────────────────────────────────────────────────────

    private SqlConnection CreateConn()
        => new(config.GetConnectionString("PandoraDb"));

    /// <summary>
    /// Crea la tabla RoomRequests con el esquema actual.
    /// Si existe con esquema anterior (sin columna Area), la elimina y recrea.
    /// </summary>
    private async Task EnsureTableAsync(SqlConnection conn, CancellationToken ct = default)
    {
        try
        {
            // Migración: tabla vieja sin columna Area → eliminar
            const string migrate = """
                IF OBJECT_ID('dbo.RoomRequests', 'U') IS NOT NULL
                   AND NOT EXISTS (
                       SELECT 1 FROM sys.columns
                        WHERE object_id = OBJECT_ID('dbo.RoomRequests')
                          AND name = 'Area')
                BEGIN
                    DROP TABLE dbo.RoomRequests;
                END
                """;

            const string create = """
                IF NOT EXISTS (
                    SELECT 1 FROM sys.objects
                    WHERE object_id = OBJECT_ID(N'dbo.RoomRequests') AND type = N'U')
                BEGIN
                    CREATE TABLE dbo.RoomRequests (
                        Id               UNIQUEIDENTIFIER NOT NULL DEFAULT NEWID() PRIMARY KEY,
                        Area             NVARCHAR(200)    NOT NULL,
                        ResponsibleName  NVARCHAR(200)    NOT NULL,
                        ActivityName     NVARCHAR(500)    NOT NULL,
                        PreferredRoom    NVARCHAR(100)    NOT NULL,
                        Disposition      NVARCHAR(50)     NULL,
                        RequestedDate    DATE             NOT NULL,
                        StartTime        TIME             NOT NULL,
                        EndTime          TIME             NOT NULL,
                        AttendeeCount    INT              NOT NULL DEFAULT 1,
                        Resources        NVARCHAR(500)    NULL,
                        CoffeeBreak      BIT              NOT NULL DEFAULT 0,
                        CoffeeBreakItems NVARCHAR(500)    NULL,
                        Status           NVARCHAR(20)     NOT NULL DEFAULT 'Pendiente',
                        AdminNotes       NVARCHAR(500)    NULL,
                        ReservationId    UNIQUEIDENTIFIER NULL,
                        CreatedAt        DATETIME2        NOT NULL DEFAULT GETUTCDATE(),
                        UpdatedAt        DATETIME2        NULL
                    );
                END
                """;

            await using var cmd = conn.CreateCommand();
            cmd.CommandText = migrate;
            await cmd.ExecuteNonQueryAsync(ct);
            cmd.CommandText = create;
            await cmd.ExecuteNonQueryAsync(ct);
        }
        catch (Exception ex)
        {
            logger.LogError(ex, "EnsureTableAsync failed");
            throw;
        }
    }

    private static object? N(SqlDataReader r, string col)
        => r.IsDBNull(r.GetOrdinal(col)) ? null : r.GetValue(r.GetOrdinal(col));

    // ── GET /api/room-requests?status=Pendiente ───────────────────────────────

    [HttpGet]
    public async Task<IActionResult> GetAll(
        [FromQuery] string? status = null,
        CancellationToken ct = default)
    {
        try
        {
            await using var conn = CreateConn();
            await conn.OpenAsync(ct);
            await EnsureTableAsync(conn, ct);

            await using var cmd = conn.CreateCommand();
            cmd.CommandText = $"""
                SELECT Id, Area, ResponsibleName, ActivityName,
                       PreferredRoom, Disposition,
                       RequestedDate, StartTime, EndTime,
                       AttendeeCount, Resources,
                       CoffeeBreak, CoffeeBreakItems,
                       Status, AdminNotes, ReservationId, CreatedAt
                FROM   dbo.RoomRequests
                {(status is not null ? "WHERE Status = @status" : "")}
                ORDER BY CreatedAt DESC
                """;

            if (status is not null)
                cmd.Parameters.AddWithValue("@status", status);

            var list = new List<object>();
            await using var reader = await cmd.ExecuteReaderAsync(ct);
            while (await reader.ReadAsync(ct))
            {
                var start = (TimeSpan)reader["StartTime"];
                var end   = (TimeSpan)reader["EndTime"];
                list.Add(new
                {
                    id               = reader.GetGuid(reader.GetOrdinal("Id")),
                    area             = reader.GetString(reader.GetOrdinal("Area")),
                    responsibleName  = reader.GetString(reader.GetOrdinal("ResponsibleName")),
                    activityName     = reader.GetString(reader.GetOrdinal("ActivityName")),
                    preferredRoom    = reader.GetString(reader.GetOrdinal("PreferredRoom")),
                    disposition      = N(reader, "Disposition"),
                    requestedDate    = ((DateTime)reader["RequestedDate"]).ToString("yyyy-MM-dd"),
                    startTime        = $"{start.Hours:D2}:{start.Minutes:D2}",
                    endTime          = $"{end.Hours:D2}:{end.Minutes:D2}",
                    attendeeCount    = reader.GetInt32(reader.GetOrdinal("AttendeeCount")),
                    resources        = N(reader, "Resources"),
                    coffeeBreak      = reader.GetBoolean(reader.GetOrdinal("CoffeeBreak")),
                    coffeeBreakItems = N(reader, "CoffeeBreakItems"),
                    status           = reader.GetString(reader.GetOrdinal("Status")),
                    adminNotes       = N(reader, "AdminNotes"),
                    reservationId    = N(reader, "ReservationId"),
                    createdAt        = reader.GetDateTime(reader.GetOrdinal("CreatedAt")),
                });
            }
            return Ok(list);
        }
        catch (Exception ex)
        {
            logger.LogError(ex, "GetAll RoomRequests failed");
            return StatusCode(500, $"Error al cargar solicitudes: {ex.Message}");
        }
    }

    // ── POST /api/room-requests ───────────────────────────────────────────────

    [HttpPost]
    public async Task<IActionResult> Create(
        [FromBody] RoomRequestDto dto,
        CancellationToken ct = default)
    {
        if (string.IsNullOrWhiteSpace(dto.Area))            return BadRequest("El área es obligatoria.");
        if (string.IsNullOrWhiteSpace(dto.ResponsibleName)) return BadRequest("El nombre del responsable es obligatorio.");
        if (string.IsNullOrWhiteSpace(dto.ActivityName))    return BadRequest("El nombre de la actividad es obligatorio.");
        if (string.IsNullOrWhiteSpace(dto.PreferredRoom))   return BadRequest("El aula es obligatoria.");

        var id = Guid.NewGuid();
        try
        {
            await using var conn = CreateConn();
            await conn.OpenAsync(ct);
            await EnsureTableAsync(conn, ct);

            await using var cmd = conn.CreateCommand();
            cmd.CommandText = """
                INSERT INTO dbo.RoomRequests
                    (Id, Area, ResponsibleName, ActivityName,
                     PreferredRoom, Disposition,
                     RequestedDate, StartTime, EndTime,
                     AttendeeCount, Resources,
                     CoffeeBreak, CoffeeBreakItems,
                     Status, CreatedAt)
                VALUES
                    (@Id, @Area, @Responsible, @Activity,
                     @Room, @Disposition,
                     @Date, @Start, @End,
                     @Count, @Resources,
                     @Coffee, @CoffeeItems,
                     'Pendiente', GETUTCDATE())
                """;

            cmd.Parameters.AddWithValue("@Id",          id);
            cmd.Parameters.AddWithValue("@Area",        dto.Area.Trim());
            cmd.Parameters.AddWithValue("@Responsible", dto.ResponsibleName.Trim());
            cmd.Parameters.AddWithValue("@Activity",    dto.ActivityName.Trim());
            cmd.Parameters.AddWithValue("@Room",        dto.PreferredRoom.Trim());
            cmd.Parameters.AddWithValue("@Disposition", string.IsNullOrWhiteSpace(dto.Disposition)
                                                            ? DBNull.Value : dto.Disposition.Trim());
            cmd.Parameters.AddWithValue("@Date",        dto.RequestedDate);
            cmd.Parameters.AddWithValue("@Start",       dto.StartTime);
            cmd.Parameters.AddWithValue("@End",         dto.EndTime);
            cmd.Parameters.AddWithValue("@Count",       dto.AttendeeCount);
            cmd.Parameters.AddWithValue("@Resources",   string.IsNullOrWhiteSpace(dto.Resources)
                                                            ? DBNull.Value : dto.Resources.Trim());
            cmd.Parameters.AddWithValue("@Coffee",      dto.CoffeeBreak);
            cmd.Parameters.AddWithValue("@CoffeeItems", string.IsNullOrWhiteSpace(dto.CoffeeBreakItems)
                                                            ? DBNull.Value : dto.CoffeeBreakItems.Trim());

            await cmd.ExecuteNonQueryAsync(ct);
            logger.LogInformation("RoomRequest created: {Id} — {Activity} por {Responsible}", id, dto.ActivityName, dto.ResponsibleName);

            // Enviar notificación por email (no bloquear si falla)
            _ = Task.Run(() => SendNotificationEmailAsync(dto, id), CancellationToken.None);

            return Ok(new { id });
        }
        catch (Exception ex)
        {
            logger.LogError(ex, "Create RoomRequest failed");
            return StatusCode(500, $"Error al guardar la solicitud: {ex.Message}");
        }
    }

    // ── PUT /api/room-requests/{id}/status ───────────────────────────────────

    [HttpPut("{id:guid}/status")]
    public async Task<IActionResult> UpdateStatus(
        Guid id,
        [FromBody] UpdateStatusDto dto,
        CancellationToken ct = default)
    {
        if (string.IsNullOrWhiteSpace(dto.Status)) return BadRequest("El estado es obligatorio.");
        try
        {
            await using var conn = CreateConn();
            await conn.OpenAsync(ct);
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = """
                UPDATE dbo.RoomRequests
                SET    Status        = @Status,
                       AdminNotes    = @AdminNotes,
                       ReservationId = @ReservationId,
                       UpdatedAt     = GETUTCDATE()
                WHERE  Id = @Id
                """;
            cmd.Parameters.AddWithValue("@Id",           id);
            cmd.Parameters.AddWithValue("@Status",        dto.Status.Trim());
            cmd.Parameters.AddWithValue("@AdminNotes",    string.IsNullOrWhiteSpace(dto.AdminNotes)
                                                              ? DBNull.Value : dto.AdminNotes.Trim());
            cmd.Parameters.AddWithValue("@ReservationId", dto.ReservationId.HasValue
                                                              ? (object)dto.ReservationId.Value : DBNull.Value);

            int rows = await cmd.ExecuteNonQueryAsync(ct);
            if (rows == 0) return NotFound("Solicitud no encontrada.");
            return Ok();
        }
        catch (Exception ex)
        {
            logger.LogError(ex, "UpdateStatus failed for {Id}", id);
            return StatusCode(500, ex.Message);
        }
    }

    // ── DELETE /api/room-requests/{id} ───────────────────────────────────────

    [HttpDelete("{id:guid}")]
    public async Task<IActionResult> Delete(Guid id, CancellationToken ct = default)
    {
        try
        {
            await using var conn = CreateConn();
            await conn.OpenAsync(ct);
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = "DELETE FROM dbo.RoomRequests WHERE Id = @Id";
            cmd.Parameters.AddWithValue("@Id", id);
            int rows = await cmd.ExecuteNonQueryAsync(ct);
            if (rows == 0) return NotFound("Solicitud no encontrada.");
            return NoContent();
        }
        catch (Exception ex)
        {
            logger.LogError(ex, "Delete RoomRequest failed for {Id}", id);
            return StatusCode(500, ex.Message);
        }
    }

    // ── Email de notificación ─────────────────────────────────────────────────

    private async Task SendNotificationEmailAsync(RoomRequestDto dto, Guid requestId)
    {
        try
        {
            var smtp     = config.GetSection("SmtpSettings");
            var host     = smtp["Host"]               ?? "";
            var port     = int.TryParse(smtp["Port"], out var p) ? p : 587;
            var user     = smtp["Username"]           ?? "";
            var pass     = smtp["Password"]           ?? "";
            var from     = smtp["FromEmail"]          ?? "";
            var fromName = smtp["FromName"]           ?? "Pandora";
            var to       = smtp["NotificationsEmail"] ?? from;
            var useSsl   = smtp["UseSsl"]?.ToLower() == "true";

            if (string.IsNullOrWhiteSpace(host) || string.IsNullOrWhiteSpace(to)
                || string.IsNullOrWhiteSpace(from)) return;

            var startStr = dto.StartTime.Length >= 5 ? dto.StartTime[..5] : dto.StartTime;
            var endStr   = dto.EndTime.Length   >= 5 ? dto.EndTime[..5]   : dto.EndTime;

            var body = $"""
                <html><body style="font-family:Arial,sans-serif;font-size:14px;color:#333">
                <div style="max-width:600px;margin:0 auto">
                  <div style="background:#1a237e;padding:20px;border-radius:8px 8px 0 0">
                    <h2 style="color:white;margin:0">📅 Nueva solicitud de espacio</h2>
                  </div>
                  <div style="border:1px solid #ddd;padding:24px;border-radius:0 0 8px 8px">
                    <table style="width:100%;border-collapse:collapse">
                      <tr style="border-bottom:1px solid #eee">
                        <td style="padding:8px;font-weight:bold;color:#555;width:40%">Área solicitante</td>
                        <td style="padding:8px">{dto.Area}</td>
                      </tr>
                      <tr style="border-bottom:1px solid #eee">
                        <td style="padding:8px;font-weight:bold;color:#555">Responsable</td>
                        <td style="padding:8px">{dto.ResponsibleName}</td>
                      </tr>
                      <tr style="border-bottom:1px solid #eee">
                        <td style="padding:8px;font-weight:bold;color:#555">Actividad</td>
                        <td style="padding:8px"><strong>{dto.ActivityName}</strong></td>
                      </tr>
                      <tr style="border-bottom:1px solid #eee">
                        <td style="padding:8px;font-weight:bold;color:#555">Aula preferida</td>
                        <td style="padding:8px">{dto.PreferredRoom}</td>
                      </tr>
                      <tr style="border-bottom:1px solid #eee">
                        <td style="padding:8px;font-weight:bold;color:#555">Disposición</td>
                        <td style="padding:8px">{dto.Disposition ?? "N/A"}</td>
                      </tr>
                      <tr style="border-bottom:1px solid #eee">
                        <td style="padding:8px;font-weight:bold;color:#555">Fecha</td>
                        <td style="padding:8px">{dto.RequestedDate}</td>
                      </tr>
                      <tr style="border-bottom:1px solid #eee">
                        <td style="padding:8px;font-weight:bold;color:#555">Horario</td>
                        <td style="padding:8px">{startStr} – {endStr} hrs</td>
                      </tr>
                      <tr style="border-bottom:1px solid #eee">
                        <td style="padding:8px;font-weight:bold;color:#555">Número de personas</td>
                        <td style="padding:8px">{dto.AttendeeCount}</td>
                      </tr>
                      <tr style="border-bottom:1px solid #eee">
                        <td style="padding:8px;font-weight:bold;color:#555">Recursos</td>
                        <td style="padding:8px">{dto.Resources ?? "—"}</td>
                      </tr>
                      <tr>
                        <td style="padding:8px;font-weight:bold;color:#555">Coffee Break</td>
                        <td style="padding:8px">
                          {(dto.CoffeeBreak ? "✅ Sí" : "❌ No")}
                          {(!string.IsNullOrEmpty(dto.CoffeeBreakItems) ? $" — {dto.CoffeeBreakItems}" : "")}
                        </td>
                      </tr>
                    </table>
                    <div style="margin-top:20px;text-align:center">
                      <a href="http://localhost:3000/calendar/solicitudes"
                         style="background:#1a237e;color:white;padding:10px 24px;border-radius:6px;text-decoration:none;font-weight:bold">
                        Ver en Pandora →
                      </a>
                    </div>
                    <p style="margin-top:20px;font-size:12px;color:#999;text-align:center">
                      Pandora — Sistema de Gestión iMET · Solicitud #{requestId.ToString()[..8]}
                    </p>
                  </div>
                </div>
                </body></html>
                """;

            using var client = new SmtpClient(host, port)
            {
                Credentials = new NetworkCredential(user, pass),
                EnableSsl   = useSsl,
                Timeout     = 10_000,
            };

            using var msg = new MailMessage
            {
                From       = new MailAddress(from, fromName),
                Subject    = $"[Pandora] Nueva solicitud de espacio — {dto.ActivityName}",
                Body       = body,
                IsBodyHtml = true,
            };
            msg.To.Add(to);

            await client.SendMailAsync(msg);
            logger.LogInformation("Notification email sent for request {Id} to {To}", requestId, to);
        }
        catch (Exception ex)
        {
            // No propagar — el email no debe bloquear la solicitud
            logger.LogWarning("Email notification failed for {Id}: {Msg}", requestId, ex.Message);
        }
    }
}

// ── DTOs ─────────────────────────────────────────────────────────────────────

public record RoomRequestDto(
    string  Area,
    string  ResponsibleName,
    string  ActivityName,
    string  PreferredRoom,
    string? Disposition,
    string  RequestedDate,     // "yyyy-MM-dd"
    string  StartTime,         // "HH:mm:ss"
    string  EndTime,           // "HH:mm:ss"
    int     AttendeeCount,
    string? Resources,         // comma-separated
    bool    CoffeeBreak,
    string? CoffeeBreakItems   // comma-separated
);

public record UpdateStatusDto(
    string  Status,
    string? AdminNotes,
    Guid?   ReservationId
);
