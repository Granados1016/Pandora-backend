using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using Pandora.API.Extensions;

namespace Pandora.API.Controllers;

/// <summary>
/// Endpoints exclusivos para que Azul AI actúe en Pandora.
/// Protegidos por X-Sso-Secret (sin JWT de usuario).
/// </summary>
[ApiController]
[Route("api/azul")]
[AllowAnonymous]
public class AzulIntegrationController(IConfiguration config, ILogger<AzulIntegrationController> logger) : ControllerBase
{
    private SqlConnection Conn() => new(config.GetConnectionString("PandoraDb"));

    private bool ValidateSecret()
    {
        var expected = config["Azul:SsoSecret"] ?? "";
        var received = Request.Headers["X-Sso-Secret"].FirstOrDefault() ?? "";
        return !string.IsNullOrWhiteSpace(expected) && received == expected;
    }

    // ── POST /api/azul/ticket ─────────────────────────────────────────────────
    [HttpPost("ticket")]
    public async Task<IActionResult> CreateTicket(
        [FromBody] AzulTicketRequest req,
        CancellationToken ct)
    {
        if (!ValidateSecret()) return Unauthorized(new { error = "Secreto SSO inválido." });
        if (string.IsNullOrWhiteSpace(req.Titulo)) return BadRequest(new { error = "Título requerido." });

        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);

            // Obtener template activo
            await using var cmdTpl = conn.CreateCommand();
            cmdTpl.CommandText = "SELECT TOP 1 Id FROM dbo.TicketTemplate WHERE IsActive = 1";
            var tplId = (Guid?)await cmdTpl.ExecuteScalarAsync(ct);
            if (tplId is null) return StatusCode(500, new { error = "No hay plantilla activa en Pandora." });

            // Generar número de ticket
            await using var cmdNum = conn.CreateCommand();
            cmdNum.CommandText = "SELECT ISNULL(COUNT(*), 0) FROM dbo.Tickets";
            var count        = Convert.ToInt32(await cmdNum.ExecuteScalarAsync(ct) ?? 0);
            var ticketNumber = $"TKT-{(count + 1):D4}";
            var id           = Guid.NewGuid();

            // Insertar ticket
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = """
                INSERT INTO dbo.Tickets
                    (Id, TicketNumber, Title, TemplateId, Status, Priority,
                     Area, SubmittedBy, SubmittedByEmail, CreatedAt)
                VALUES
                    (@Id, @Num, @Title, @TplId, 'Abierto', @Priority,
                     @Area, @User, @Email, GETUTCDATE())
                """;
            cmd.Parameters.AddWithValue("@Id",       id);
            cmd.Parameters.AddWithValue("@Num",      ticketNumber);
            cmd.Parameters.AddWithValue("@Title",    req.Titulo.Trim());
            cmd.Parameters.AddWithValue("@TplId",    tplId.Value);
            cmd.Parameters.AddWithValue("@Priority", string.IsNullOrWhiteSpace(req.Prioridad) ? "Media" : req.Prioridad.Trim());
            cmd.Parameters.AddWithValue("@Area",     string.IsNullOrWhiteSpace(req.Area) ? DBNull.Value : (object)req.Area.Trim());
            cmd.Parameters.AddWithValue("@User",     string.IsNullOrWhiteSpace(req.NombreUsuario) ? "Azul AI" : req.NombreUsuario.Trim());
            cmd.Parameters.AddWithValue("@Email",    string.IsNullOrWhiteSpace(req.EmailUsuario)  ? DBNull.Value : (object)req.EmailUsuario.Trim());
            await cmd.ExecuteNonQueryAsync(ct);

            // Si hay descripción, guardarla como comentario inicial
            if (!string.IsNullOrWhiteSpace(req.Descripcion))
            {
                await using var cmdC = conn.CreateCommand();
                cmdC.CommandText = """
                    INSERT INTO dbo.TicketComments (Id, TicketId, AuthorName, IsAdmin, Body)
                    VALUES (NEWID(), @TicketId, @Author, 0, @Body)
                    """;
                cmdC.Parameters.AddWithValue("@TicketId", id);
                cmdC.Parameters.AddWithValue("@Author",   req.NombreUsuario?.Trim() ?? "Azul AI");
                cmdC.Parameters.AddWithValue("@Body",     req.Descripcion.Trim());
                await cmdC.ExecuteNonQueryAsync(ct);
            }

            logger.LogInformation("Ticket creado desde Azul AI: {Num} — {Title} por {User}", ticketNumber, req.Titulo, req.NombreUsuario);
            return Ok(new { id, ticketNumber, message = $"Ticket {ticketNumber} creado correctamente en Pandora." });
        }
        catch (Exception ex)
        {
            logger.LogError(ex, "AzulIntegration CreateTicket failed");
            return StatusCode(500, new { error = ex.Message });
        }
    }

    // ── POST /api/azul/sala ───────────────────────────────────────────────────
    [HttpPost("sala")]
    public async Task<IActionResult> CreateRoomRequest(
        [FromBody] AzulSalaRequest req,
        CancellationToken ct)
    {
        if (!ValidateSecret()) return Unauthorized(new { error = "Secreto SSO inválido." });
        if (string.IsNullOrWhiteSpace(req.Actividad))  return BadRequest(new { error = "Actividad requerida." });
        if (string.IsNullOrWhiteSpace(req.Sala))       return BadRequest(new { error = "Sala requerida." });
        if (string.IsNullOrWhiteSpace(req.Fecha))      return BadRequest(new { error = "Fecha requerida (yyyy-MM-dd)." });
        if (string.IsNullOrWhiteSpace(req.HoraInicio)) return BadRequest(new { error = "Hora de inicio requerida (HH:mm)." });
        if (string.IsNullOrWhiteSpace(req.HoraFin))    return BadRequest(new { error = "Hora de fin requerida (HH:mm)." });

        // Normalizar hora a HH:mm:ss
        static string NormalizeTime(string t) =>
            t.Length == 5 ? t + ":00" : t;

        var id = Guid.NewGuid();
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);

            // Asegurar que la tabla existe
            await using var ensureCmd = conn.CreateCommand();
            ensureCmd.CommandText = """
                IF NOT EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID(N'dbo.RoomRequests') AND type = N'U')
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
            await ensureCmd.ExecuteNonQueryAsync(ct);

            await using var cmd = conn.CreateCommand();
            cmd.CommandText = """
                INSERT INTO dbo.RoomRequests
                    (Id, Area, ResponsibleName, ActivityName,
                     PreferredRoom, RequestedDate, StartTime, EndTime,
                     AttendeeCount, Resources, CoffeeBreak, Status, CreatedAt)
                VALUES
                    (@Id, @Area, @Responsible, @Activity,
                     @Room, @Date, @Start, @End,
                     @Count, @Resources, 0, 'Pendiente', GETUTCDATE())
                """;
            cmd.Parameters.AddWithValue("@Id",          id);
            cmd.Parameters.AddWithValue("@Area",        string.IsNullOrWhiteSpace(req.Area) ? "Sin área" : req.Area.Trim());
            cmd.Parameters.AddWithValue("@Responsible", string.IsNullOrWhiteSpace(req.NombreUsuario) ? "Azul AI" : req.NombreUsuario.Trim());
            cmd.Parameters.AddWithValue("@Activity",    req.Actividad.Trim());
            cmd.Parameters.AddWithValue("@Room",        req.Sala.Trim());
            cmd.Parameters.AddWithValue("@Date",        req.Fecha.Trim());
            cmd.Parameters.AddWithValue("@Start",       NormalizeTime(req.HoraInicio.Trim()));
            cmd.Parameters.AddWithValue("@End",         NormalizeTime(req.HoraFin.Trim()));
            cmd.Parameters.AddWithValue("@Count",       req.Asistentes > 0 ? req.Asistentes : 1);
            cmd.Parameters.AddWithValue("@Resources",   string.IsNullOrWhiteSpace(req.Recursos) ? DBNull.Value : (object)req.Recursos.Trim());
            await cmd.ExecuteNonQueryAsync(ct);

            logger.LogInformation("Sala agendada desde Azul AI: {Activity} en {Room} el {Date} por {User}",
                req.Actividad, req.Sala, req.Fecha, req.NombreUsuario);

            return Ok(new
            {
                id,
                message = $"Solicitud de sala registrada en Pandora para el {req.Fecha} de {req.HoraInicio} a {req.HoraFin} en {req.Sala}. Quedará pendiente de aprobación."
            });
        }
        catch (Exception ex)
        {
            logger.LogError(ex, "AzulIntegration CreateRoomRequest failed");
            return StatusCode(500, new { error = ex.Message });
        }
    }

    // ── POST /api/azul/buscar ─────────────────────────────────────────────────
    [HttpPost("buscar")]
    public async Task<IActionResult> Search(
        [FromBody] AzulSearchRequest req,
        CancellationToken ct)
    {
        if (!ValidateSecret()) return Unauthorized(new { error = "Secreto SSO inválido." });
        if (string.IsNullOrWhiteSpace(req.Query)) return Ok("No se especificó consulta de búsqueda.");

        var term    = req.Query.Trim();
        var modulo  = (req.Modulo ?? "").ToLowerInvariant();
        var results = new List<SearchResult>();

        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);

            if (modulo == "" || modulo == "tickets")
                await TryAzulSearch(conn, ct, results, """
                    SELECT TOP 5 'ticket' AS Type, CAST(Id AS NVARCHAR(36)), Title AS Label,
                           TicketNumber + ' — ' + Status AS Subtitle, CreatedAt AS Date
                    FROM dbo.Tickets
                    WHERE IsDeleted = 0
                      AND (Title LIKE '%' + @Q + '%' OR TicketNumber LIKE '%' + @Q + '%'
                        OR SubmittedBy LIKE '%' + @Q + '%')
                    ORDER BY CreatedAt DESC
                    """, term);

            if (modulo == "" || modulo == "inventario")
                await TryAzulSearch(conn, ct, results, """
                    SELECT TOP 5 'inventario' AS Type, CAST(Id AS NVARCHAR(36)), Name AS Label,
                           ISNULL(Brand,'') + ' ' + ISNULL(Model,'') + ' — ' + ISNULL(Department,'Sin área') AS Subtitle,
                           CreatedAt AS Date
                    FROM dbo.InventoryItems
                    WHERE IsActive = 1
                      AND (Name LIKE '%' + @Q + '%' OR Brand LIKE '%' + @Q + '%'
                        OR Model LIKE '%' + @Q + '%' OR InventoryNumber LIKE '%' + @Q + '%')
                    ORDER BY CreatedAt DESC
                    """, term);

            if (modulo == "" || modulo == "licencias")
                await TryAzulSearch(conn, ct, results, """
                    SELECT TOP 5 'licencia' AS Type, CAST(Id AS NVARCHAR(10)), Plataforma AS Label,
                           Area + ' — ' + Estado AS Subtitle, GETUTCDATE() AS Date
                    FROM dbo.Licencias
                    WHERE (Plataforma LIKE '%' + @Q + '%' OR Area LIKE '%' + @Q + '%'
                        OR Responsable LIKE '%' + @Q + '%')
                    """, term);

            if (modulo == "" || modulo == "calendario")
                await TryAzulSearch(conn, ct, results, """
                    SELECT TOP 5 'evento' AS Type, CAST(Id AS NVARCHAR(36)), Title AS Label,
                           ISNULL(Location,'') + ' — ' + CONVERT(VARCHAR(16), StartTime, 120) AS Subtitle,
                           StartTime AS Date
                    FROM dbo.CalendarEvents
                    WHERE IsDeleted = 0
                      AND (Title LIKE '%' + @Q + '%' OR Location LIKE '%' + @Q + '%')
                    ORDER BY StartTime DESC
                    """, term);

            if (modulo == "" || modulo == "procedimientos")
                await TryAzulSearch(conn, ct, results, """
                    SELECT TOP 5 'procedimiento' AS Type, CAST(Id AS NVARCHAR(36)), Title AS Label,
                           ISNULL(Description,'Sin descripción') AS Subtitle, UploadedAt AS Date
                    FROM dbo.Procedimientos
                    WHERE IsDeleted = 0
                      AND (Title LIKE '%' + @Q + '%' OR Description LIKE '%' + @Q + '%')
                    ORDER BY UploadedAt DESC
                    """, term);
        }
        catch (Exception ex)
        {
            logger.LogWarning(ex, "AzulIntegration Buscar failed para query '{Query}'", term);
        }

        if (results.Count == 0)
            return Ok($"No encontré resultados en Pandora para \"{term}\"" +
                      (string.IsNullOrWhiteSpace(req.Modulo) ? "." : $" en el módulo {req.Modulo}."));

        // Formatear respuesta como texto legible para Azul AI
        var sb = new System.Text.StringBuilder();
        sb.AppendLine($"Encontré {results.Count} resultado(s) en Pandora para \"{term}\":");
        sb.AppendLine();
        foreach (var r in results)
        {
            sb.AppendLine($"• [{r.Type.ToUpper()}] {r.Label}");
            if (!string.IsNullOrWhiteSpace(r.Subtitle))
                sb.AppendLine($"  {r.Subtitle}");
        }

        logger.LogInformation("Azul buscar: {Count} resultados para '{Query}' módulo='{Modulo}'",
            results.Count, term, modulo);

        return Ok(sb.ToString().Trim());
    }

    private static async Task TryAzulSearch(
        SqlConnection conn, CancellationToken ct,
        List<SearchResult> results, string sql, string term)
    {
        try
        {
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = sql;
            cmd.Parameters.AddWithValue("@Q", term);
            await using var r = await cmd.ExecuteReaderAsync(ct);
            while (await r.ReadAsync(ct))
            {
                results.Add(new SearchResult(
                    Type:     r.GetString(0),
                    Id:       r.GetValue(1)?.ToString() ?? "",
                    Label:    r.IsDBNull(2) ? "" : r.GetString(2),
                    Subtitle: r.IsDBNull(3) ? "" : r.GetString(3),
                    Date:     r.GetUtcDateTimeOrNull(4)
                ));
            }
        }
        catch { /* módulo puede no existir en este entorno */ }
    }
}

// ── Request / Response DTOs ───────────────────────────────────────────────────

public record AzulTicketRequest(
    string  Titulo,
    string? Descripcion,
    string? Prioridad,        // Alta | Media | Baja
    string? Area,
    string? NombreUsuario,
    string? EmailUsuario
);

public record AzulSearchRequest(
    string  Query,
    string? Modulo   // tickets | inventario | licencias | calendario | procedimientos | "" (todos)
);

public record AzulSalaRequest(
    string  Actividad,
    string  Sala,
    string  Fecha,            // yyyy-MM-dd
    string  HoraInicio,       // HH:mm
    string  HoraFin,          // HH:mm
    string? Area,
    int     Asistentes,
    string? Recursos,
    string? NombreUsuario,
    string? EmailUsuario
);

internal record SearchResult(
    string    Type,
    string    Id,
    string    Label,
    string    Subtitle,
    DateTime? Date
);
