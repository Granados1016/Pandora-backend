using ClosedXML.Excel;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.SignalR;
using Microsoft.Data.SqlClient;
using Pandora.API.Hubs;
using Pandora.API.Services;
using System.Security.Claims;

namespace Pandora.API.Controllers;

[ApiController]
[Route("api/vacaciones")]
[Authorize]
public class VacacionesController(
    IConfiguration config,
    IWebHostEnvironment env,
    ILogger<VacacionesController> logger,
    IHubContext<NotificationsHub> hub) : ControllerBase
{
    private SqlConnection Conn() => new(config.GetConnectionString("PandoraDb"));

    private string CurrentUser =>
        User.FindFirstValue(ClaimTypes.Name) ??
        User.FindFirstValue("name") ?? "Desconocido";

    private string CurrentFullName =>
        User.FindFirstValue("fullName") ??
        User.FindFirstValue(ClaimTypes.GivenName) ??
        CurrentUser;

    private bool IsAdmin =>
        User.IsInRole("Admin") ||
        User.Claims.Any(c => c.Type == ClaimTypes.Role && c.Value == "Admin");

    private string DocsPath =>
        Path.Combine(env.ContentRootPath, "storage", "vacaciones-docs");

    // ── Helper: cuenta días hábiles (excluye sáb/dom + festivos) ─────────────
    private static int CountWorkdays(DateOnly start, DateOnly end, HashSet<DateOnly> holidays)
    {
        int count = 0;
        for (var d = start; d <= end; d = d.AddDays(1))
        {
            if (d.DayOfWeek is DayOfWeek.Saturday or DayOfWeek.Sunday) continue;
            if (holidays.Contains(d)) continue;
            count++;
        }
        return count;
    }

    // ── Cargar festivos del año (recurrentes + específicos) ───────────────────
    private async Task<HashSet<DateOnly>> LoadHolidaysAsync(int year, SqlConnection conn, CancellationToken ct)
    {
        var set = new HashSet<DateOnly>();
        await using var cmd = conn.CreateCommand();
        cmd.CommandText = """
            SELECT HolidayDate, IsRecurring FROM dbo.VacationHolidays
            WHERE IsDeleted = 0
              AND (YEAR(HolidayDate) = @Year OR IsRecurring = 1)
            """;
        cmd.Parameters.AddWithValue("@Year", year);
        await using var r = await cmd.ExecuteReaderAsync(ct);
        while (await r.ReadAsync(ct))
        {
            var date = DateOnly.FromDateTime(r.GetDateTime(0));
            var isRecurring = r.GetBoolean(1);
            // Para recurrentes, usar mes/día del año solicitado
            var effective = isRecurring ? new DateOnly(year, date.Month, date.Day) : date;
            set.Add(effective);
        }
        return set;
    }

    // ── GET /api/vacaciones/mi-calendario/{year} ──────────────────────────────
    [HttpGet("mi-calendario/{year:int}")]
    public async Task<IActionResult> MiCalendario(int year, CancellationToken ct)
    {
        var username = CurrentUser;
        await using var conn = Conn();
        await conn.OpenAsync(ct);

        var holidays = await LoadHolidaysAsync(year, conn, ct);

        // Solicitudes del usuario en el año
        var markedDays = new List<object>();

        // Festivos
        foreach (var h in holidays)
            markedDays.Add(new { date = h.ToString("yyyy-MM-dd"), type = "holiday", label = "Festivo" });

        // Solicitudes
        await using var cmd = conn.CreateCommand();
        cmd.CommandText = """
            SELECT Id, StartDate, EndDate, Type, Status
            FROM dbo.VacationRequests
            WHERE Username = @Username
              AND IsDeleted = 0
              AND (YEAR(StartDate) = @Year OR YEAR(EndDate) = @Year)
            """;
        cmd.Parameters.AddWithValue("@Username", username);
        cmd.Parameters.AddWithValue("@Year", year);
        await using var r = await cmd.ExecuteReaderAsync(ct);
        while (await r.ReadAsync(ct))
        {
            int    id     = r.GetInt32(0);
            var    start  = DateOnly.FromDateTime(r.GetDateTime(1));
            var    end    = DateOnly.FromDateTime(r.GetDateTime(2));
            string type   = r.GetString(3);
            string status = r.GetString(4);

            for (var d = start; d <= end; d = d.AddDays(1))
            {
                if (d.Year != year) continue;
                markedDays.Add(new { date = d.ToString("yyyy-MM-dd"), type = status.ToLower(), requestId = id, requestType = type });
            }
        }

        return Ok(markedDays);
    }

    // ── GET /api/vacaciones/mis-dias ──────────────────────────────────────────
    [HttpGet("mis-dias")]
    public async Task<IActionResult> MisDias(CancellationToken ct)
    {
        var username = CurrentUser;
        int year = DateTime.UtcNow.Year;

        await using var conn = Conn();
        await conn.OpenAsync(ct);

        // Política: días asignados
        int totalDays = 15;
        await using (var pol = conn.CreateCommand())
        {
            pol.CommandText = """
                SELECT TotalDays FROM dbo.VacationPolicies
                WHERE Username = @Username AND Year = @Year
                """;
            pol.Parameters.AddWithValue("@Username", username);
            pol.Parameters.AddWithValue("@Year", year);
            var val = await pol.ExecuteScalarAsync(ct);
            if (val != null) totalDays = (int)val;
        }

        // Días usados (aprobados en el año)
        int usedDays = 0;
        await using (var used = conn.CreateCommand())
        {
            used.CommandText = """
                SELECT ISNULL(SUM(TotalDays), 0)
                FROM dbo.VacationRequests
                WHERE Username = @Username
                  AND YEAR(StartDate) = @Year
                  AND Status = 'Aprobado'
                  AND IsDeleted = 0
                """;
            used.Parameters.AddWithValue("@Username", username);
            used.Parameters.AddWithValue("@Year", year);
            usedDays = (int)(await used.ExecuteScalarAsync(ct) ?? 0);
        }

        // Días en solicitudes pendientes
        int pendingDays = 0;
        await using (var pend = conn.CreateCommand())
        {
            pend.CommandText = """
                SELECT ISNULL(SUM(TotalDays), 0)
                FROM dbo.VacationRequests
                WHERE Username = @Username
                  AND YEAR(StartDate) = @Year
                  AND Status = 'Pendiente'
                  AND IsDeleted = 0
                """;
            pend.Parameters.AddWithValue("@Username", username);
            pend.Parameters.AddWithValue("@Year", year);
            pendingDays = (int)(await pend.ExecuteScalarAsync(ct) ?? 0);
        }

        return Ok(new
        {
            year,
            totalDays,
            usedDays,
            pendingDays,
            availableDays = totalDays - usedDays
        });
    }

    // ── POST /api/vacaciones/solicitar ────────────────────────────────────────
    [HttpPost("solicitar")]
    public async Task<IActionResult> Solicitar([FromBody] VacacionSolicitudDto dto, CancellationToken ct)
    {
        if (dto.StartDate > dto.EndDate)
            return BadRequest("La fecha de inicio no puede ser posterior a la fecha fin.");

        var username = CurrentUser;
        var fullName = CurrentFullName;

        await using var conn = Conn();
        await conn.OpenAsync(ct);

        var holidays = await LoadHolidaysAsync(dto.StartDate.Year, conn, ct);
        int totalDays = CountWorkdays(dto.StartDate, dto.EndDate, holidays);

        if (totalDays == 0)
            return BadRequest("El rango seleccionado no contiene días hábiles.");

        // Verificar solapamiento con solicitudes existentes
        await using var checkCmd = conn.CreateCommand();
        checkCmd.CommandText = """
            SELECT COUNT(*) FROM dbo.VacationRequests
            WHERE Username = @Username
              AND IsDeleted = 0
              AND Status IN ('Pendiente','Aprobado')
              AND StartDate <= @End AND EndDate >= @Start
            """;
        checkCmd.Parameters.AddWithValue("@Username", username);
        checkCmd.Parameters.AddWithValue("@Start", dto.StartDate.ToDateTime(TimeOnly.MinValue));
        checkCmd.Parameters.AddWithValue("@End",   dto.EndDate.ToDateTime(TimeOnly.MinValue));
        int overlap = (int)(await checkCmd.ExecuteScalarAsync(ct) ?? 0);
        if (overlap > 0)
            return Conflict("Ya tienes una solicitud que se traslapa con esas fechas.");

        // Insertar solicitud
        int newId;
        await using var ins = conn.CreateCommand();
        ins.CommandText = """
            INSERT INTO dbo.VacationRequests
                (Username, FullName, StartDate, EndDate, TotalDays, Type, Notes, Status, CreatedAt, IsDeleted)
            OUTPUT INSERTED.Id
            VALUES (@Username, @FullName, @Start, @End, @TotalDays, @Type, @Notes, 'Pendiente', GETUTCDATE(), 0)
            """;
        ins.Parameters.AddWithValue("@Username", username);
        ins.Parameters.AddWithValue("@FullName", fullName);
        ins.Parameters.AddWithValue("@Start",    dto.StartDate.ToDateTime(TimeOnly.MinValue));
        ins.Parameters.AddWithValue("@End",      dto.EndDate.ToDateTime(TimeOnly.MinValue));
        ins.Parameters.AddWithValue("@TotalDays",totalDays);
        ins.Parameters.AddWithValue("@Type",     dto.Type ?? "Vacaciones");
        ins.Parameters.AddWithValue("@Notes",    (object?)dto.Notes ?? DBNull.Value);
        newId = (int)(await ins.ExecuteScalarAsync(ct))!;

        // Notificar a Admin por SignalR
        await hub.Clients.Group("broadcast").SendAsync("NewNotification", new
        {
            id      = 0,
            title   = "📅 Nueva solicitud de vacaciones",
            message = $"{fullName} solicitó {totalDays} día(s) ({dto.StartDate:dd/MM} – {dto.EndDate:dd/MM})",
            type    = "vacacion",
            isRead  = false,
            path    = "/vacaciones/admin",
        }, ct);

        logger.LogInformation("Solicitud de vacaciones #{Id} creada por {User}", newId, username);
        return Ok(new { id = newId, totalDays });
    }

    // ── DELETE /api/vacaciones/{id}/cancelar ──────────────────────────────────
    [HttpDelete("{id:int}/cancelar")]
    public async Task<IActionResult> Cancelar(int id, CancellationToken ct)
    {
        var username = CurrentUser;
        await using var conn = Conn();
        await conn.OpenAsync(ct);
        await using var cmd = conn.CreateCommand();
        cmd.CommandText = """
            UPDATE dbo.VacationRequests
            SET Status = 'Cancelado'
            WHERE Id = @Id AND Username = @Username
              AND Status = 'Pendiente' AND IsDeleted = 0
            """;
        cmd.Parameters.AddWithValue("@Id",       id);
        cmd.Parameters.AddWithValue("@Username", username);
        int rows = await cmd.ExecuteNonQueryAsync(ct);
        return rows > 0 ? NoContent() : NotFound();
    }

    // ── GET /api/vacaciones/mis-solicitudes ───────────────────────────────────
    [HttpGet("mis-solicitudes")]
    public async Task<IActionResult> MisSolicitudes(CancellationToken ct)
    {
        var username = CurrentUser;
        await using var conn = Conn();
        await conn.OpenAsync(ct);
        await using var cmd = conn.CreateCommand();
        cmd.CommandText = """
            SELECT Id, StartDate, EndDate, TotalDays, Type, Status, Notes, ReviewNotes, CreatedAt,
                   CASE WHEN DocumentPath IS NOT NULL THEN 1 ELSE 0 END AS HasDocument,
                   ReviewedBy, ReviewedAt
            FROM dbo.VacationRequests
            WHERE Username = @Username AND IsDeleted = 0
            ORDER BY CreatedAt DESC
            """;
        cmd.Parameters.AddWithValue("@Username", username);
        var items = new List<object>();
        await using var r = await cmd.ExecuteReaderAsync(ct);
        while (await r.ReadAsync(ct))
            items.Add(new {
                id          = r.GetInt32(0),
                startDate   = r.GetDateTime(1).ToString("yyyy-MM-dd"),
                endDate     = r.GetDateTime(2).ToString("yyyy-MM-dd"),
                totalDays   = r.GetInt32(3),
                type        = r.GetString(4),
                status      = r.GetString(5),
                notes       = r.IsDBNull(6) ? null : r.GetString(6),
                reviewNotes = r.IsDBNull(7) ? null : r.GetString(7),
                createdAt   = r.GetDateTime(8),
                hasDocument = r.GetInt32(9) == 1,
                reviewedBy  = r.IsDBNull(10) ? null : r.GetString(10),
                reviewedAt  = r.IsDBNull(11) ? (DateTime?)null : r.GetDateTime(11),
            });
        return Ok(items);
    }

    // ════════════════════════════════════════════════════════════════════════════
    // ADMIN endpoints
    // ════════════════════════════════════════════════════════════════════════════

    // ── GET /api/vacaciones/admin/solicitudes ─────────────────────────────────
    [HttpGet("admin/solicitudes")]
    [Authorize(Roles = "Admin")]
    public async Task<IActionResult> AdminSolicitudes(
        [FromQuery] string? status = null, CancellationToken ct = default)
    {
        await using var conn = Conn();
        await conn.OpenAsync(ct);
        await using var cmd = conn.CreateCommand();
        cmd.CommandText = """
            SELECT Id, Username, FullName, StartDate, EndDate, TotalDays,
                   Type, Status, Notes, ReviewedBy, ReviewedAt, ReviewNotes, CreatedAt,
                   CASE WHEN DocumentPath IS NOT NULL THEN 1 ELSE 0 END AS HasDocument,
                   ISNULL(ApprovalStage,'PendienteJefe') AS ApprovalStage,
                   JefeReviewedBy, JefeReviewedAt, JefeNotes
            FROM dbo.VacationRequests
            WHERE IsDeleted = 0
              AND (@Status IS NULL OR Status = @Status)
            ORDER BY CreatedAt DESC
            """;
        cmd.Parameters.AddWithValue("@Status", (object?)status ?? DBNull.Value);
        var items = new List<object>();
        await using var r = await cmd.ExecuteReaderAsync(ct);
        while (await r.ReadAsync(ct))
            items.Add(new {
                id              = r.GetInt32(0),
                username        = r.GetString(1),
                fullName        = r.IsDBNull(2) ? r.GetString(1) : r.GetString(2),
                startDate       = r.GetDateTime(3).ToString("yyyy-MM-dd"),
                endDate         = r.GetDateTime(4).ToString("yyyy-MM-dd"),
                totalDays       = r.GetInt32(5),
                type            = r.GetString(6),
                status          = r.GetString(7),
                notes           = r.IsDBNull(8)  ? null : r.GetString(8),
                reviewedBy      = r.IsDBNull(9)  ? null : r.GetString(9),
                reviewedAt      = r.IsDBNull(10) ? (DateTime?)null : r.GetDateTime(10),
                reviewNotes     = r.IsDBNull(11) ? null : r.GetString(11),
                createdAt       = r.GetDateTime(12),
                hasDocument     = r.GetInt32(13) == 1,
                approvalStage   = r.GetString(14),
                jefeReviewedBy  = r.IsDBNull(15) ? null : r.GetString(15),
                jefeReviewedAt  = r.IsDBNull(16) ? (DateTime?)null : r.GetDateTime(16),
                jefeNotes       = r.IsDBNull(17) ? null : r.GetString(17),
            });
        return Ok(items);
    }

    // ── PUT /api/vacaciones/jefe/{id}/revisar  (primera aprobación) ──────────
    [HttpPut("jefe/{id:int}/revisar")]
    [Authorize(Roles = "Admin")]
    public async Task<IActionResult> RevisarJefe(int id, [FromBody] RevisionDto dto, CancellationToken ct)
    {
        if (dto.Status is not ("AprobadoJefe" or "Rechazado"))
            return BadRequest("Status debe ser AprobadoJefe o Rechazado.");

        var reviewer = CurrentUser;
        await using var conn = Conn();
        await conn.OpenAsync(ct);

        string reqUsername = "", reqFullName = "", reqEmail = "", reqStart = "", reqEnd = "";
        int reqDays = 0;
        await using (var sel = conn.CreateCommand())
        {
            sel.CommandText = """
                SELECT vr.Username, vr.FullName, ISNULL(au.Email,''),
                       vr.StartDate, vr.EndDate, vr.TotalDays
                FROM dbo.VacationRequests vr
                LEFT JOIN dbo.AppUsers au ON LOWER(au.Username)=LOWER(vr.Username)
                WHERE vr.Id=@Id AND vr.IsDeleted=0
                """;
            sel.Parameters.AddWithValue("@Id", id);
            await using var r = await sel.ExecuteReaderAsync(ct);
            if (!await r.ReadAsync(ct)) return NotFound();
            reqUsername = r.GetString(0);
            reqFullName = r.IsDBNull(1) ? r.GetString(0) : r.GetString(1);
            reqEmail    = r.IsDBNull(2) ? "" : r.GetString(2);
            reqStart    = r.GetDateTime(3).ToString("dd/MM/yyyy");
            reqEnd      = r.GetDateTime(4).ToString("dd/MM/yyyy");
            reqDays     = r.GetInt32(5);
        }

        // Si rechaza el jefe, el status final es Rechazado
        var finalStatus = dto.Status == "Rechazado" ? "Rechazado" : "Pendiente";
        await using var cmd = conn.CreateCommand();
        cmd.CommandText = """
            UPDATE dbo.VacationRequests
            SET ApprovalStage   = @Stage,
                Status          = @Status,
                JefeReviewedBy  = @JefeBy,
                JefeReviewedAt  = GETUTCDATE(),
                JefeNotes       = @JefeNotes
            WHERE Id=@Id AND ApprovalStage='PendienteJefe' AND IsDeleted=0
            """;
        cmd.Parameters.AddWithValue("@Id",       id);
        cmd.Parameters.AddWithValue("@Stage",    dto.Status == "Rechazado" ? "Rechazado" : "PendienteRRHH");
        cmd.Parameters.AddWithValue("@Status",   finalStatus);
        cmd.Parameters.AddWithValue("@JefeBy",   reviewer);
        cmd.Parameters.AddWithValue("@JefeNotes",(object?)dto.Notes ?? DBNull.Value);
        int rows = await cmd.ExecuteNonQueryAsync(ct);
        if (rows == 0) return NotFound();

        var icon = dto.Status == "AprobadoJefe" ? "✅" : "❌";
        var msg  = dto.Status == "AprobadoJefe"
            ? "Tu solicitud fue aprobada por tu jefe directo y está en revisión de RRHH."
            : $"Tu solicitud fue rechazada por tu jefe directo. {dto.Notes}";

        await hub.Clients.Group($"user-{reqUsername}").SendAsync("NewNotification", new
        {
            id=0, title=$"{icon} Revisión de jefe directo",
            message=msg, type="vacacion", isRead=false, path="/vacaciones",
        }, ct);

        // Notificar a admins si fue aprobada por jefe (para segunda revisión)
        if (dto.Status == "AprobadoJefe")
            await hub.Clients.Group("broadcast").SendAsync("NewNotification", new
            {
                id=0, title="📋 Solicitud pendiente de revisión RRHH",
                message=$"{reqFullName} — {reqStart} al {reqEnd} ({reqDays} días). Aprobada por jefe directo.",
                type="vacacion", isRead=false, path="/vacaciones/admin",
            }, ct);

        if (!string.IsNullOrWhiteSpace(reqEmail) && dto.Status == "Rechazado")
            _ = SendVacacionEmailAsync(reqEmail, reqFullName, "Rechazado", dto.Notes, reqStart, reqEnd, reqDays);

        return NoContent();
    }

    // ── PUT /api/vacaciones/admin/{id}/revisar  (segunda aprobación — RRHH) ──
    [HttpPut("admin/{id:int}/revisar")]
    [Authorize(Roles = "Admin")]
    public async Task<IActionResult> Revisar(int id, [FromBody] RevisionDto dto, CancellationToken ct)
    {
        if (dto.Status is not ("Aprobado" or "Rechazado"))
            return BadRequest("Status debe ser Aprobado o Rechazado.");

        var reviewer = CurrentUser;

        await using var conn = Conn();
        await conn.OpenAsync(ct);

        string reqUsername = "", reqFullName = "", reqEmail = "", reqStart = "", reqEnd = "";
        int reqDays = 0;
        await using (var sel = conn.CreateCommand())
        {
            sel.CommandText = """
                SELECT vr.Username, vr.FullName,
                       ISNULL(au.Email, ''), vr.StartDate, vr.EndDate, vr.TotalDays
                FROM   dbo.VacationRequests vr
                LEFT JOIN dbo.AppUsers au ON LOWER(au.Username) = LOWER(vr.Username)
                WHERE  vr.Id = @Id AND vr.IsDeleted = 0
                """;
            sel.Parameters.AddWithValue("@Id", id);
            await using var r = await sel.ExecuteReaderAsync(ct);
            if (!await r.ReadAsync(ct)) return NotFound();
            reqUsername = r.GetString(0);
            reqFullName = r.IsDBNull(1) ? r.GetString(0) : r.GetString(1);
            reqEmail    = r.IsDBNull(2) ? "" : r.GetString(2);
            reqStart    = r.GetDateTime(3).ToString("dd/MM/yyyy");
            reqEnd      = r.GetDateTime(4).ToString("dd/MM/yyyy");
            reqDays     = r.GetInt32(5);
        }

        await using var cmd = conn.CreateCommand();
        cmd.CommandText = """
            UPDATE dbo.VacationRequests
            SET Status = @Status, ApprovalStage = @Stage,
                ReviewedBy = @ReviewedBy, ReviewedAt = GETUTCDATE(), ReviewNotes = @ReviewNotes
            WHERE Id = @Id AND ApprovalStage = 'PendienteRRHH' AND IsDeleted = 0
            """;
        cmd.Parameters.AddWithValue("@Id",          id);
        cmd.Parameters.AddWithValue("@Status",      dto.Status);
        cmd.Parameters.AddWithValue("@Stage",       dto.Status);
        cmd.Parameters.AddWithValue("@ReviewedBy",  reviewer);
        cmd.Parameters.AddWithValue("@ReviewNotes", (object?)dto.Notes ?? DBNull.Value);
        int rows = await cmd.ExecuteNonQueryAsync(ct);
        if (rows == 0) return NotFound();

        var icon  = dto.Status == "Aprobado" ? "✅" : "❌";
        await hub.Clients.Group($"user-{reqUsername}").SendAsync("NewNotification", new
        {
            id      = 0,
            title   = $"{icon} Tu solicitud fue {dto.Status.ToLower()}",
            message = dto.Notes ?? $"Tu solicitud de vacaciones ha sido {dto.Status.ToLower()} por RRHH.",
            type    = "vacacion",
            isRead  = false,
            path    = "/vacaciones",
        }, ct);

        logger.LogInformation("Solicitud #{Id} {Status} por {Reviewer}", id, dto.Status, reviewer);

        if (!string.IsNullOrWhiteSpace(reqEmail))
            _ = SendVacacionEmailAsync(reqEmail, reqFullName, dto.Status, dto.Notes, reqStart, reqEnd, reqDays);

        return NoContent();
    }

    // ── GET /api/vacaciones/admin/festivos/{year} ─────────────────────────────
    [HttpGet("admin/festivos/{year:int}")]
    [Authorize(Roles = "Admin")]
    public async Task<IActionResult> GetFestivos(int year, CancellationToken ct)
    {
        await using var conn = Conn();
        await conn.OpenAsync(ct);
        await using var cmd = conn.CreateCommand();
        cmd.CommandText = """
            SELECT Id, HolidayDate, Name, IsRecurring
            FROM dbo.VacationHolidays
            WHERE IsDeleted = 0
              AND (YEAR(HolidayDate) = @Year OR IsRecurring = 1)
            ORDER BY MONTH(HolidayDate), DAY(HolidayDate)
            """;
        cmd.Parameters.AddWithValue("@Year", year);
        var items = new List<object>();
        await using var r = await cmd.ExecuteReaderAsync(ct);
        while (await r.ReadAsync(ct))
            items.Add(new {
                id          = r.GetInt32(0),
                holidayDate = r.GetDateTime(1).ToString("yyyy-MM-dd"),
                name        = r.GetString(2),
                isRecurring = r.GetBoolean(3),
            });
        return Ok(items);
    }

    // ── POST /api/vacaciones/admin/festivos ───────────────────────────────────
    [HttpPost("admin/festivos")]
    [Authorize(Roles = "Admin")]
    public async Task<IActionResult> AddFestivo([FromBody] FestivoDto dto, CancellationToken ct)
    {
        await using var conn = Conn();
        await conn.OpenAsync(ct);
        await using var cmd = conn.CreateCommand();
        cmd.CommandText = """
            INSERT INTO dbo.VacationHolidays (HolidayDate, Name, IsRecurring)
            OUTPUT INSERTED.Id
            VALUES (@Date, @Name, @IsRecurring)
            """;
        cmd.Parameters.AddWithValue("@Date",        dto.HolidayDate.ToDateTime(TimeOnly.MinValue));
        cmd.Parameters.AddWithValue("@Name",        dto.Name);
        cmd.Parameters.AddWithValue("@IsRecurring", dto.IsRecurring);
        int newId = (int)(await cmd.ExecuteScalarAsync(ct))!;
        return Ok(new { id = newId });
    }

    // ── DELETE /api/vacaciones/admin/festivos/{id} ────────────────────────────
    [HttpDelete("admin/festivos/{id:int}")]
    [Authorize(Roles = "Admin")]
    public async Task<IActionResult> DeleteFestivo(int id, CancellationToken ct)
    {
        await using var conn = Conn();
        await conn.OpenAsync(ct);
        await using var cmd = conn.CreateCommand();
        cmd.CommandText = "UPDATE dbo.VacationHolidays SET IsDeleted=1 WHERE Id=@Id AND IsDeleted=0";
        cmd.Parameters.AddWithValue("@Id", id);
        int rows = await cmd.ExecuteNonQueryAsync(ct);
        return rows > 0 ? NoContent() : NotFound();
    }

    // ── GET /api/vacaciones/admin/politicas ───────────────────────────────────
    [HttpGet("admin/politicas")]
    [Authorize(Roles = "Admin")]
    public async Task<IActionResult> GetPoliticas([FromQuery] int? year = null, CancellationToken ct = default)
    {
        int y = year ?? DateTime.UtcNow.Year;
        await using var conn = Conn();
        await conn.OpenAsync(ct);
        await using var cmd = conn.CreateCommand();
        cmd.CommandText = """
            SELECT u.Username, u.FullName,
                   ISNULL(p.TotalDays, 15)  AS TotalDays,
                   ISNULL(p.Id, 0)          AS PolicyId,
                   ISNULL(e.Position, '')   AS Position,
                   ISNULL(d.Name, '')       AS Department,
                   ISNULL(u.Email, '')      AS Email,
                   ISNULL(
                       (SELECT ISNULL(SUM(r.TotalDays),0)
                        FROM dbo.VacationRequests r
                        WHERE r.Username = u.Username
                          AND YEAR(r.StartDate) = @Year
                          AND r.Status = 'Aprobado'
                          AND r.IsDeleted = 0), 0) AS UsedDays
            FROM dbo.AppUsers u
            LEFT JOIN dbo.VacationPolicies p
                ON p.Username = u.Username AND p.Year = @Year
            LEFT JOIN dbo.Employees e
                ON LOWER(e.Email) = LOWER(u.Email) OR LOWER(e.FullName) = LOWER(u.FullName)
            LEFT JOIN dbo.Departments d
                ON d.Id = e.DepartmentId
            WHERE u.IsActive = 1
            ORDER BY u.FullName
            """;
        cmd.Parameters.AddWithValue("@Year", y);
        var items = new List<object>();
        await using var r = await cmd.ExecuteReaderAsync(ct);
        while (await r.ReadAsync(ct))
        {
            int total = r.GetInt32(2);
            int used  = r.GetInt32(7);
            items.Add(new {
                username   = r.GetString(0),
                fullName   = r.IsDBNull(1) ? r.GetString(0) : r.GetString(1),
                totalDays  = total,
                policyId   = r.GetInt32(3),
                position   = r.GetString(4),
                department = r.GetString(5),
                email      = r.GetString(6),
                usedDays   = used,
                available  = total - used,
                year       = y,
            });
        }
        return Ok(items);
    }

    // ── PUT /api/vacaciones/admin/politicas ───────────────────────────────────
    [HttpPut("admin/politicas")]
    [Authorize(Roles = "Admin")]
    public async Task<IActionResult> SetPolitica([FromBody] PoliticaDto dto, CancellationToken ct)
    {
        await using var conn = Conn();
        await conn.OpenAsync(ct);
        await using var cmd = conn.CreateCommand();
        cmd.CommandText = """
            MERGE dbo.VacationPolicies AS target
            USING (SELECT @Username AS Username, @Year AS Year) AS source
                ON target.Username = source.Username AND target.Year = source.Year
            WHEN MATCHED THEN
                UPDATE SET TotalDays = @TotalDays
            WHEN NOT MATCHED THEN
                INSERT (Username, Year, TotalDays)
                VALUES (@Username, @Year, @TotalDays);
            """;
        cmd.Parameters.AddWithValue("@Username",  dto.Username);
        cmd.Parameters.AddWithValue("@Year",      dto.Year);
        cmd.Parameters.AddWithValue("@TotalDays", dto.TotalDays);
        await cmd.ExecuteNonQueryAsync(ct);
        return NoContent();
    }
    // ── POST /api/vacaciones/{id}/documento ──────────────────────────────────
    [HttpPost("{id:int}/documento")]
    [Consumes("multipart/form-data")]
    public async Task<IActionResult> UploadDocumento(int id, IFormFile? file, CancellationToken ct)
    {
        if (file is null || file.Length == 0)
            return BadRequest("No se recibió ningún archivo.");

        var allowedExts = new[] { ".pdf", ".jpg", ".jpeg", ".png" };
        var ext = Path.GetExtension(file.FileName).ToLowerInvariant();
        if (!allowedExts.Contains(ext))
            return BadRequest("Solo se permiten archivos PDF, JPG o PNG.");
        if (file.Length > 10 * 1024 * 1024)
            return BadRequest("El archivo no puede superar 10 MB.");

        var username = CurrentUser;

        await using var conn = Conn();
        await conn.OpenAsync(ct);

        // Verificar que la solicitud existe y pertenece al usuario (o admin)
        string? owner;
        string? oldPath;
        await using (var sel = conn.CreateCommand())
        {
            sel.CommandText = """
                SELECT Username, DocumentPath
                FROM dbo.VacationRequests
                WHERE Id = @Id AND IsDeleted = 0
                """;
            sel.Parameters.AddWithValue("@Id", id);
            await using var rr = await sel.ExecuteReaderAsync(ct);
            if (!await rr.ReadAsync(ct)) return NotFound();
            owner   = rr.GetString(0);
            oldPath = rr.IsDBNull(1) ? null : rr.GetString(1);
        }

        if (!IsAdmin && owner != username) return Forbid();

        // Eliminar archivo anterior si existe
        if (!string.IsNullOrEmpty(oldPath) && System.IO.File.Exists(oldPath))
            System.IO.File.Delete(oldPath);

        // Guardar nuevo archivo
        Directory.CreateDirectory(DocsPath);
        var fileName = $"vacreq_{id}_{Guid.NewGuid():N}{ext}";
        var filePath = Path.Combine(DocsPath, fileName);
        await using (var stream = System.IO.File.Create(filePath))
            await file.CopyToAsync(stream, ct);

        // Actualizar BD
        await using var upd = conn.CreateCommand();
        upd.CommandText = """
            UPDATE dbo.VacationRequests
            SET DocumentPath = @Path
            WHERE Id = @Id AND IsDeleted = 0
            """;
        upd.Parameters.AddWithValue("@Path", filePath);
        upd.Parameters.AddWithValue("@Id",   id);
        await upd.ExecuteNonQueryAsync(ct);

        logger.LogInformation("Documento subido para solicitud #{Id} por {User}", id, username);
        return Ok(new { fileName });
    }

    // ── GET /api/vacaciones/{id}/documento ────────────────────────────────────
    [HttpGet("{id:int}/documento")]
    public async Task<IActionResult> GetDocumento(int id, CancellationToken ct)
    {
        var username = CurrentUser;

        await using var conn = Conn();
        await conn.OpenAsync(ct);

        await using var cmd = conn.CreateCommand();
        cmd.CommandText = """
            SELECT Username, DocumentPath
            FROM dbo.VacationRequests
            WHERE Id = @Id AND IsDeleted = 0
            """;
        cmd.Parameters.AddWithValue("@Id", id);
        await using var r = await cmd.ExecuteReaderAsync(ct);
        if (!await r.ReadAsync(ct)) return NotFound();

        var owner   = r.GetString(0);
        var docPath = r.IsDBNull(1) ? null : r.GetString(1);

        if (!IsAdmin && owner != username)
            return Forbid();
        if (string.IsNullOrEmpty(docPath) || !System.IO.File.Exists(docPath))
            return NotFound("No hay documento adjunto.");

        var fileExt = Path.GetExtension(docPath).ToLowerInvariant();
        var contentType = fileExt switch
        {
            ".pdf"            => "application/pdf",
            ".png"            => "image/png",
            ".jpg" or ".jpeg" => "image/jpeg",
            _                 => "application/octet-stream"
        };

        var bytes = await System.IO.File.ReadAllBytesAsync(docPath, ct);
        return File(bytes, contentType, Path.GetFileName(docPath));
    }

    // ── GET /api/vacaciones/admin/export-excel ────────────────────────────────
    [HttpGet("admin/export-excel")]
    [Authorize(Roles = "Admin")]
    public async Task<IActionResult> ExportExcel([FromQuery] int? year = null, [FromQuery] string? status = null, CancellationToken ct = default)
    {
        int y = year ?? DateTime.UtcNow.Year;
        await using var conn = Conn();
        await conn.OpenAsync(ct);
        await using var cmd = conn.CreateCommand();
        cmd.CommandText = $"""
            SELECT r.Id, u.FullName, u.Username, u.Email,
                   r.StartDate, r.EndDate, r.TotalDays, r.Status,
                   r.Type, r.Notes,
                   r.ReviewedBy, r.ReviewedAt, r.ReviewNotes,
                   r.CreatedAt
            FROM dbo.VacationRequests r
            JOIN dbo.AppUsers u ON u.Username = r.Username
            WHERE r.IsDeleted = 0
              AND (YEAR(r.StartDate) = @Year OR YEAR(r.EndDate) = @Year)
              {(string.IsNullOrWhiteSpace(status) ? "" : "AND r.Status = @Status")}
            ORDER BY r.CreatedAt DESC
            """;
        cmd.Parameters.AddWithValue("@Year", y);
        if (!string.IsNullOrWhiteSpace(status)) cmd.Parameters.AddWithValue("@Status", status);

        var rows = new List<(int id, string name, string user, string email,
            string start, string end, int days, string sts, string type, string? notes,
            string? reviewedBy, string? reviewedAt, string? reviewNotes, string created)>();

        await using var r = await cmd.ExecuteReaderAsync(ct);
        while (await r.ReadAsync(ct))
        {
            rows.Add((
                r.GetInt32(0),
                r.IsDBNull(1) ? "" : r.GetString(1),
                r.GetString(2),
                r.IsDBNull(3) ? "" : r.GetString(3),
                r.GetFieldValue<DateOnly>(4).ToString("dd/MM/yyyy"),
                r.GetFieldValue<DateOnly>(5).ToString("dd/MM/yyyy"),
                r.GetInt32(6),
                r.GetString(7),
                r.IsDBNull(8) ? "Vacaciones" : r.GetString(8),
                r.IsDBNull(9)  ? null : r.GetString(9),
                r.IsDBNull(10) ? null : r.GetString(10),
                r.IsDBNull(11) ? null : ((DateTime)r.GetValue(11)).ToString("dd/MM/yyyy HH:mm"),
                r.IsDBNull(12) ? null : r.GetString(12),
                ((DateTime)r.GetValue(13)).ToString("dd/MM/yyyy")
            ));
        }

        using var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Vacaciones");

        // Encabezados
        string[] headers = ["#", "Nombre", "Usuario", "Correo", "Inicio", "Fin", "Días", "Tipo", "Estado", "Notas", "Revisado por", "Fecha revisión", "Observaciones", "Solicitado"];
        for (int c = 0; c < headers.Length; c++)
        {
            var cell = ws.Cell(1, c + 1);
            cell.Value = headers[c];
            cell.Style.Font.Bold = true;
            cell.Style.Fill.BackgroundColor = XLColor.FromHtml("#1a237e");
            cell.Style.Font.FontColor = XLColor.White;
        }

        // Datos
        for (int i = 0; i < rows.Count; i++)
        {
            var row = rows[i];
            var rn = i + 2;
            ws.Cell(rn, 1).Value  = row.id;
            ws.Cell(rn, 2).Value  = row.name;
            ws.Cell(rn, 3).Value  = row.user;
            ws.Cell(rn, 4).Value  = row.email;
            ws.Cell(rn, 5).Value  = row.start;
            ws.Cell(rn, 6).Value  = row.end;
            ws.Cell(rn, 7).Value  = row.days;
            ws.Cell(rn, 8).Value  = row.type;
            ws.Cell(rn, 9).Value  = row.sts;
            ws.Cell(rn, 10).Value = row.notes ?? "";
            ws.Cell(rn, 11).Value = row.reviewedBy ?? "";
            ws.Cell(rn, 12).Value = row.reviewedAt ?? "";
            ws.Cell(rn, 13).Value = row.reviewNotes ?? "";
            ws.Cell(rn, 14).Value = row.created;

            var statusColor = row.sts switch
            {
                "Aprobada" => XLColor.FromHtml("#e8f5e9"),
                "Rechazada" => XLColor.FromHtml("#ffebee"),
                "Pendiente" => XLColor.FromHtml("#fff8e1"),
                _ => XLColor.White
            };
            ws.Row(rn).Style.Fill.BackgroundColor = statusColor;
        }

        ws.Columns().AdjustToContents();
        ws.SheetView.FreezeRows(1);

        using var ms = new MemoryStream();
        wb.SaveAs(ms);
        ms.Position = 0;

        var fileName = $"Vacaciones_{y}_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
        return File(ms.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
    }

    // ── GET /api/vacaciones/admin/equipo/{year}/{month} ───────────────────────
    [HttpGet("admin/equipo/{year:int}/{month:int}")]
    [Authorize(Roles = "Admin")]
    public async Task<IActionResult> GetEquipoCalendario(int year, int month, CancellationToken ct)
    {
        await using var conn = Conn();
        await conn.OpenAsync(ct);
        await using var cmd = conn.CreateCommand();
        cmd.CommandText = """
            SELECT r.Id, u.FullName, u.Username,
                   r.StartDate, r.EndDate, r.TotalDays, r.Status, r.Type
            FROM dbo.VacationRequests r
            JOIN dbo.AppUsers u ON u.Username = r.Username
            WHERE r.IsDeleted = 0
              AND r.Status = 'Aprobada'
              AND (
                  (YEAR(r.StartDate) = @Year AND MONTH(r.StartDate) = @Month)
                  OR (YEAR(r.EndDate) = @Year AND MONTH(r.EndDate) = @Month)
                  OR (r.StartDate <= DATEFROMPARTS(@Year, @Month, 1)
                      AND r.EndDate >= EOMONTH(DATEFROMPARTS(@Year, @Month, 1)))
              )
            ORDER BY r.StartDate, u.FullName
            """;
        cmd.Parameters.AddWithValue("@Year",  year);
        cmd.Parameters.AddWithValue("@Month", month);

        var items = new List<object>();
        await using var r = await cmd.ExecuteReaderAsync(ct);
        while (await r.ReadAsync(ct))
        {
            items.Add(new
            {
                id        = r.GetInt32(0),
                fullName  = r.IsDBNull(1) ? r.GetString(2) : r.GetString(1),
                username  = r.GetString(2),
                startDate = r.GetFieldValue<DateOnly>(3).ToString("yyyy-MM-dd"),
                endDate   = r.GetFieldValue<DateOnly>(4).ToString("yyyy-MM-dd"),
                totalDays = r.GetInt32(5),
                status    = r.GetString(6),
                type      = r.IsDBNull(7) ? "Vacaciones" : r.GetString(7),
            });
        }
        return Ok(items);
    }

    // ── GET /api/vacaciones/admin/calendario-area/{year}/{month} ─────────────
    [HttpGet("admin/calendario-area/{year:int}/{month:int}")]
    [Authorize(Roles = "Admin")]
    public async Task<IActionResult> GetCalendarioArea(int year, int month, CancellationToken ct)
    {
        await using var conn = Conn();
        await conn.OpenAsync(ct);

        // Festivos del mes
        var holidays = await LoadHolidaysAsync(year, conn, ct);

        // Ausencias aprobadas que tocan este mes
        await using var cmd = conn.CreateCommand();
        cmd.CommandText = """
            SELECT r.Id, u.FullName, u.Username,
                   ISNULL(d.Name,'Sin área') AS Department,
                   r.StartDate, r.EndDate, r.TotalDays, r.Type
            FROM dbo.VacationRequests r
            JOIN dbo.AppUsers u ON u.Username = r.Username
            LEFT JOIN dbo.Employees e ON LOWER(e.Email)=LOWER(u.Email)
            LEFT JOIN dbo.Departments d ON d.Id = e.DepartmentId
            WHERE r.IsDeleted=0
              AND r.Status='Aprobado'
              AND (
                (YEAR(r.StartDate)=@Year AND MONTH(r.StartDate)=@Month)
                OR (YEAR(r.EndDate)=@Year AND MONTH(r.EndDate)=@Month)
                OR (r.StartDate <= DATEFROMPARTS(@Year,@Month,1)
                    AND r.EndDate >= EOMONTH(DATEFROMPARTS(@Year,@Month,1)))
              )
            ORDER BY d.Name, u.FullName
            """;
        cmd.Parameters.AddWithValue("@Year",  year);
        cmd.Parameters.AddWithValue("@Month", month);

        var daysInMonth = DateTime.DaysInMonth(year, month);
        var items = new List<object>();
        await using var r = await cmd.ExecuteReaderAsync(ct);
        while (await r.ReadAsync(ct))
        {
            var start = DateOnly.FromDateTime(r.GetDateTime(4));
            var end   = DateOnly.FromDateTime(r.GetDateTime(5));
            // Días ausentes dentro del mes (solo hábiles)
            var days = new List<int>();
            for (int d = 1; d <= daysInMonth; d++)
            {
                var date = new DateOnly(year, month, d);
                if (date < start || date > end) continue;
                if (date.DayOfWeek is DayOfWeek.Saturday or DayOfWeek.Sunday) continue;
                if (holidays.Contains(date)) continue;
                days.Add(d);
            }
            items.Add(new {
                id         = r.GetInt32(0),
                fullName   = r.IsDBNull(1) ? r.GetString(2) : r.GetString(1),
                username   = r.GetString(2),
                department = r.GetString(3),
                startDate  = start.ToString("yyyy-MM-dd"),
                endDate    = end.ToString("yyyy-MM-dd"),
                totalDays  = r.GetInt32(6),
                type       = r.IsDBNull(7) ? "Vacaciones" : r.GetString(7),
                daysInMonth = days,
            });
        }

        return Ok(new { year, month, daysInMonth, holidays = holidays.Where(h => h.Year == year && h.Month == month).Select(h => h.Day).ToList(), items });
    }

    // ── Enviar correo de notificación de vacación (#14) ───────────────────────
    private async Task SendVacacionEmailAsync(
        string toEmail, string toName,
        string status, string? reviewNotes,
        string startDate, string endDate, int totalDays)
    {
        try
        {
            var cfg   = await SmtpHelper.LoadAsync(config.GetConnectionString("PandoraDb")!, config);
            var icon  = status == "Aprobado" ? "✅" : "❌";
            var color = status == "Aprobado" ? "#2e7d32" : "#b71c1c";
            var body  = $"""
                <!DOCTYPE html><html lang="es"><head><meta charset="UTF-8"></head>
                <body style="font-family:Arial,sans-serif;background:#f5f5f5;margin:0;padding:20px">
                  <div style="max-width:560px;margin:0 auto;background:white;border-radius:8px;overflow:hidden;box-shadow:0 2px 8px rgba(0,0,0,.1)">
                    <div style="background:#1a237e;padding:24px;text-align:center">
                      <h1 style="color:white;margin:0;font-size:22px">PANDORA</h1>
                      <p style="color:rgba(255,255,255,.7);margin:4px 0 0;font-size:13px">Sistema de Gestión</p>
                    </div>
                    <div style="padding:28px">
                      <h2 style="color:{color};font-size:18px;margin:0 0 16px">{icon} Solicitud {status.ToLower()}</h2>
                      <p style="color:#333;margin:0 0 12px">Hola <strong>{toName}</strong>,</p>
                      <p style="color:#555;margin:0 0 20px">Tu solicitud de vacaciones ha sido <strong style="color:{color}">{status.ToLower()}</strong>.</p>
                      <div style="background:#f9f9f9;border:1px solid #eee;border-radius:6px;padding:16px;margin-bottom:20px">
                        <p style="margin:0 0 8px"><strong>Período:</strong> {startDate} al {endDate}</p>
                        <p style="margin:0"><strong>Total:</strong> {totalDays} día{(totalDays != 1 ? "s" : "")}</p>
                        {(reviewNotes != null ? $"<p style=\"margin:8px 0 0\"><strong>Nota:</strong> {reviewNotes}</p>" : "")}
                      </div>
                      <p style="color:#888;font-size:12px">Puedes revisar el estado en el módulo de Vacaciones.</p>
                    </div>
                    <div style="background:#f9f9f9;border-top:1px solid #eee;padding:14px 28px;text-align:center">
                      <p style="color:#aaa;font-size:11px;margin:0">Pandora — Coordinación de TI</p>
                    </div>
                  </div>
                </body></html>
                """;

            var err = await SmtpHelper.SendAsync(cfg, toEmail, toName,
                $"{icon} Tu solicitud de vacaciones ha sido {status.ToLower()} — Pandora", body);

            if (err != null) logger.LogWarning("Correo vacación no enviado a {Email}: {Err}", toEmail, err);
            else logger.LogInformation("Correo de vacación enviado a {Email}", toEmail);
        }
        catch (Exception ex)
        {
            logger.LogWarning(ex, "No se pudo enviar correo de vacación a {Email}", toEmail);
        }
    }
}

// ── DTOs ─────────────────────────────────────────────────────────────────────
public record VacacionSolicitudDto(
    DateOnly  StartDate,
    DateOnly  EndDate,
    string?   Type  = "Vacaciones",
    string?   Notes = null);

public record RevisionDto(
    string  Status,
    string? Notes = null);

public record FestivoDto(
    DateOnly HolidayDate,
    string   Name,
    bool     IsRecurring = true);

public record PoliticaDto(
    string Username,
    int    Year,
    int    TotalDays);
