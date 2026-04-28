using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.RateLimiting;
using Microsoft.Data.SqlClient;
using System.Net.Mail;
using System.Net;
using System.Security.Claims;
using System.Text.Json;

namespace Pandora.API.Controllers;

[ApiController]
[Route("api/tickets")]
[Authorize]
public class TicketsController(
    IConfiguration config,
    IWebHostEnvironment env,
    ILogger<TicketsController> logger) : ControllerBase
{
    private SqlConnection Conn() => new(config.GetConnectionString("PandoraDb"));

    private string CurrentUser =>
        User.FindFirstValue(ClaimTypes.Name) ??
        User.FindFirstValue("name") ??
        User.Claims.FirstOrDefault(c => c.Type.EndsWith("name", StringComparison.OrdinalIgnoreCase))?.Value ??
        "Desconocido";

    private string? CurrentEmail =>
        User.FindFirstValue(ClaimTypes.Email) ??
        User.FindFirstValue("email") ??
        User.Claims.FirstOrDefault(c => c.Type.EndsWith("email", StringComparison.OrdinalIgnoreCase))?.Value;

    private bool IsAdmin =>
        User.IsInRole("Admin") ||
        User.Claims.Any(c => c.Type == ClaimTypes.Role && c.Value == "Admin");

    private static string? Ns(SqlDataReader r, string col) =>
        r.IsDBNull(r.GetOrdinal(col)) ? null : r.GetString(r.GetOrdinal(col));

    private static object? N(SqlDataReader r, string col) =>
        r.IsDBNull(r.GetOrdinal(col)) ? null : r.GetValue(r.GetOrdinal(col));

    private string AttachmentsPath()
    {
        var path = Path.Combine(env.ContentRootPath, "storage", "ticket-attachments");
        Directory.CreateDirectory(path);
        return path;
    }

    // ── Table init ────────────────────────────────────────────────────────────

    private async Task EnsureTablesAsync(SqlConnection conn, CancellationToken ct = default)
    {
        await using var cmd = conn.CreateCommand();
        cmd.CommandText = """
            IF NOT EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID(N'dbo.TicketTemplate') AND type = N'U')
            BEGIN
                CREATE TABLE dbo.TicketTemplate (
                    Id          UNIQUEIDENTIFIER NOT NULL DEFAULT NEWID() PRIMARY KEY,
                    Name        NVARCHAR(200)    NOT NULL DEFAULT 'Tickets de Soporte',
                    Description NVARCHAR(500)    NULL,
                    IsActive    BIT              NOT NULL DEFAULT 1,
                    CreatedAt   DATETIME2        NOT NULL DEFAULT GETUTCDATE(),
                    UpdatedAt   DATETIME2        NULL
                );
                INSERT INTO dbo.TicketTemplate (Id, Name, Description, IsActive)
                VALUES (NEWID(), 'Tickets de Soporte', 'Formulario de atención de incidencias', 1);
            END

            IF NOT EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID(N'dbo.TemplateFields') AND type = N'U')
            BEGIN
                CREATE TABLE dbo.TemplateFields (
                    Id          UNIQUEIDENTIFIER NOT NULL DEFAULT NEWID() PRIMARY KEY,
                    TemplateId  UNIQUEIDENTIFIER NOT NULL,
                    Label       NVARCHAR(200)    NOT NULL,
                    FieldType   NVARCHAR(50)     NOT NULL DEFAULT 'Text',
                    IsRequired  BIT              NOT NULL DEFAULT 0,
                    Options     NVARCHAR(MAX)    NULL,
                    Placeholder NVARCHAR(200)    NULL,
                    HelpText    NVARCHAR(500)    NULL,
                    Width       NVARCHAR(10)     NOT NULL DEFAULT 'full',
                    SortOrder   INT              NOT NULL DEFAULT 0,
                    CreatedAt   DATETIME2        NOT NULL DEFAULT GETUTCDATE()
                );
            END

            IF NOT EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID(N'dbo.Tickets') AND type = N'U')
            BEGIN
                CREATE TABLE dbo.Tickets (
                    Id               UNIQUEIDENTIFIER NOT NULL DEFAULT NEWID() PRIMARY KEY,
                    TicketNumber     NVARCHAR(20)     NOT NULL,
                    Title            NVARCHAR(500)    NOT NULL,
                    TemplateId       UNIQUEIDENTIFIER NOT NULL,
                    Status           NVARCHAR(30)     NOT NULL DEFAULT 'Abierto',
                    Priority         NVARCHAR(20)     NOT NULL DEFAULT 'Media',
                    Department       NVARCHAR(200)    NULL,
                    AssignedTo       NVARCHAR(200)    NULL,
                    SubmittedBy      NVARCHAR(200)    NOT NULL,
                    SubmittedByEmail NVARCHAR(200)    NULL,
                    DelayNote        NVARCHAR(1000)   NULL,
                    CloseNote        NVARCHAR(1000)   NULL,
                    CreatedAt        DATETIME2        NOT NULL DEFAULT GETUTCDATE(),
                    UpdatedAt        DATETIME2        NULL
                );
            END

            IF NOT EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID(N'dbo.TicketFieldValues') AND type = N'U')
            BEGIN
                CREATE TABLE dbo.TicketFieldValues (
                    Id       UNIQUEIDENTIFIER NOT NULL DEFAULT NEWID() PRIMARY KEY,
                    TicketId UNIQUEIDENTIFIER NOT NULL,
                    FieldId  UNIQUEIDENTIFIER NOT NULL,
                    Value    NVARCHAR(MAX)    NULL
                );
            END

            IF NOT EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID(N'dbo.TicketAttachments') AND type = N'U')
            BEGIN
                CREATE TABLE dbo.TicketAttachments (
                    Id           UNIQUEIDENTIFIER NOT NULL DEFAULT NEWID() PRIMARY KEY,
                    TicketId     UNIQUEIDENTIFIER NOT NULL,
                    OriginalName NVARCHAR(500)    NOT NULL,
                    StoredName   NVARCHAR(500)    NOT NULL,
                    ContentType  NVARCHAR(100)    NOT NULL,
                    FileSize     BIGINT           NOT NULL,
                    UploadedAt   DATETIME2        NOT NULL DEFAULT GETUTCDATE()
                );
            END

            IF NOT EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID(N'dbo.TicketComments') AND type = N'U')
            BEGIN
                CREATE TABLE dbo.TicketComments (
                    Id         UNIQUEIDENTIFIER NOT NULL DEFAULT NEWID() PRIMARY KEY,
                    TicketId   UNIQUEIDENTIFIER NOT NULL,
                    AuthorName NVARCHAR(200)    NOT NULL,
                    IsAdmin    BIT              NOT NULL DEFAULT 0,
                    Body       NVARCHAR(MAX)    NOT NULL,
                    CreatedAt  DATETIME2        NOT NULL DEFAULT GETUTCDATE()
                );
            END

            -- Migración: agregar columna Area a Tickets si no existe
            IF NOT EXISTS (
                SELECT 1 FROM sys.columns
                WHERE object_id = OBJECT_ID('dbo.Tickets') AND name = 'Area')
            BEGIN
                ALTER TABLE dbo.Tickets ADD Area NVARCHAR(100) NULL;
            END

            -- Configuración de correos por área
            IF NOT EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID(N'dbo.TicketAreaConfigs') AND type = N'U')
            BEGIN
                CREATE TABLE dbo.TicketAreaConfigs (
                    Id                UNIQUEIDENTIFIER NOT NULL DEFAULT NEWID() PRIMARY KEY,
                    Area              NVARCHAR(100)    NOT NULL UNIQUE,
                    NotificationEmail NVARCHAR(200)    NULL,
                    UpdatedAt         DATETIME2        NULL
                );
                INSERT INTO dbo.TicketAreaConfigs (Area) VALUES
                    (N'Dirección General'),
                    (N'Dirección Administración'),
                    (N'Dirección de programas académicos'),
                    (N'Dirección comercial'),
                    (N'Coordinación de TI'),
                    (N'Coordinación de Mercadotecnia'),
                    (N'Coordinación de Innovación'),
                    (N'Coordinación de Control escolar'),
                    (N'Diseño'),
                    (N'Community Manager'),
                    (N'Audio Visual'),
                    (N'Asesor 1'),
                    (N'Asesor 2'),
                    (N'Asesor 3');
            END

            -- Migración: sincronizar puestos en instalaciones existentes
            IF OBJECT_ID('dbo.TicketAreaConfigs') IS NOT NULL
            BEGIN
                -- Insertar puestos nuevos si no existen
                INSERT INTO dbo.TicketAreaConfigs (Area)
                SELECT src.area
                FROM (VALUES
                    (N'Dirección General'),
                    (N'Dirección Administración'),
                    (N'Dirección de programas académicos'),
                    (N'Dirección comercial'),
                    (N'Coordinación de TI'),
                    (N'Coordinación de Mercadotecnia'),
                    (N'Coordinación de Innovación'),
                    (N'Coordinación de Control escolar'),
                    (N'Diseño'),
                    (N'Community Manager'),
                    (N'Audio Visual'),
                    (N'Asesor 1'),
                    (N'Asesor 2'),
                    (N'Asesor 3')
                ) AS src(area)
                WHERE NOT EXISTS (
                    SELECT 1 FROM dbo.TicketAreaConfigs WHERE Area = src.area);

                -- Eliminar puestos anteriores que ya no aplican y no tienen tickets asociados
                DELETE FROM dbo.TicketAreaConfigs
                WHERE Area IN (N'Administración', N'Comercial', N'TI', N'MKT', N'Programas Académicos')
                  AND NOT EXISTS (
                      SELECT 1 FROM dbo.Tickets WHERE Area = dbo.TicketAreaConfigs.Area);
            END
            """;
        await cmd.ExecuteNonQueryAsync(ct);
    }

    // ════════════════════════════════════════════════════════════════════════════
    //  TEMPLATE
    // ════════════════════════════════════════════════════════════════════════════

    [HttpGet("template")]
    public async Task<IActionResult> GetTemplate(CancellationToken ct)
    {
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await EnsureTablesAsync(conn, ct);

            await using var cmd = conn.CreateCommand();
            cmd.CommandText = "SELECT TOP 1 Id, Name, Description FROM dbo.TicketTemplate WHERE IsActive = 1 ORDER BY CreatedAt";
            await using var r = await cmd.ExecuteReaderAsync(ct);
            if (!await r.ReadAsync(ct)) return NotFound("No hay plantilla activa.");
            var tplId = r.GetGuid(r.GetOrdinal("Id"));
            var tpl = new { id = tplId, name = r.GetString(r.GetOrdinal("Name")), description = Ns(r, "Description") };
            await r.CloseAsync();

            await using var cmd2 = conn.CreateCommand();
            cmd2.CommandText = """
                SELECT Id, Label, FieldType, IsRequired, Options, Placeholder, HelpText, Width, SortOrder
                FROM dbo.TemplateFields WHERE TemplateId = @TplId ORDER BY SortOrder, CreatedAt
                """;
            cmd2.Parameters.AddWithValue("@TplId", tplId);
            var fields = new List<object>();
            await using var r2 = await cmd2.ExecuteReaderAsync(ct);
            while (await r2.ReadAsync(ct))
                fields.Add(new
                {
                    id          = r2.GetGuid(r2.GetOrdinal("Id")),
                    label       = r2.GetString(r2.GetOrdinal("Label")),
                    fieldType   = r2.GetString(r2.GetOrdinal("FieldType")),
                    isRequired  = r2.GetBoolean(r2.GetOrdinal("IsRequired")),
                    options     = Ns(r2, "Options"),
                    placeholder = Ns(r2, "Placeholder"),
                    helpText    = Ns(r2, "HelpText"),
                    width       = r2.GetString(r2.GetOrdinal("Width")),
                    sortOrder   = r2.GetInt32(r2.GetOrdinal("SortOrder")),
                });
            return Ok(new { template = tpl, fields });
        }
        catch (Exception ex) { logger.LogError(ex, "GetTemplate failed"); return StatusCode(500, ex.Message); }
    }

    [HttpPut("template")]
    public async Task<IActionResult> UpdateTemplate([FromBody] TemplateMetaDto dto, CancellationToken ct)
    {
        if (!IsAdmin) return Forbid();
        if (string.IsNullOrWhiteSpace(dto.Name)) return BadRequest("Nombre requerido.");
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = """
                UPDATE dbo.TicketTemplate
                SET Name = @Name, Description = @Desc, UpdatedAt = GETUTCDATE()
                WHERE IsActive = 1
                """;
            cmd.Parameters.AddWithValue("@Name", dto.Name.Trim());
            cmd.Parameters.AddWithValue("@Desc", string.IsNullOrWhiteSpace(dto.Description) ? DBNull.Value : dto.Description.Trim());
            await cmd.ExecuteNonQueryAsync(ct);
            return Ok();
        }
        catch (Exception ex) { logger.LogError(ex, "UpdateTemplate failed"); return StatusCode(500, ex.Message); }
    }

    [HttpPost("template/fields")]
    public async Task<IActionResult> AddField([FromBody] TemplateFieldDto dto, CancellationToken ct)
    {
        if (!IsAdmin) return Forbid();
        if (string.IsNullOrWhiteSpace(dto.Label)) return BadRequest("Etiqueta requerida.");
        var id = Guid.NewGuid();
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);

            await using var cmd0 = conn.CreateCommand();
            cmd0.CommandText = "SELECT TOP 1 Id FROM dbo.TicketTemplate WHERE IsActive = 1";
            var tplId = (Guid?)await cmd0.ExecuteScalarAsync(ct);
            if (tplId == null) return NotFound("No hay plantilla activa.");

            await using var cmd1 = conn.CreateCommand();
            cmd1.CommandText = "SELECT ISNULL(MAX(SortOrder), -1) + 1 FROM dbo.TemplateFields WHERE TemplateId = @TplId";
            cmd1.Parameters.AddWithValue("@TplId", tplId.Value);
            var sortOrder = Convert.ToInt32(await cmd1.ExecuteScalarAsync(ct) ?? 0);

            await using var cmd = conn.CreateCommand();
            cmd.CommandText = """
                INSERT INTO dbo.TemplateFields (Id, TemplateId, Label, FieldType, IsRequired, Options, Placeholder, HelpText, Width, SortOrder)
                VALUES (@Id, @TplId, @Label, @FType, @Required, @Opts, @Ph, @Help, @Width, @Sort)
                """;
            cmd.Parameters.AddWithValue("@Id",       id);
            cmd.Parameters.AddWithValue("@TplId",    tplId.Value);
            cmd.Parameters.AddWithValue("@Label",    dto.Label.Trim());
            cmd.Parameters.AddWithValue("@FType",    dto.FieldType ?? "Text");
            cmd.Parameters.AddWithValue("@Required", dto.IsRequired);
            cmd.Parameters.AddWithValue("@Opts",     string.IsNullOrWhiteSpace(dto.Options)     ? DBNull.Value : dto.Options.Trim());
            cmd.Parameters.AddWithValue("@Ph",       string.IsNullOrWhiteSpace(dto.Placeholder) ? DBNull.Value : dto.Placeholder.Trim());
            cmd.Parameters.AddWithValue("@Help",     string.IsNullOrWhiteSpace(dto.HelpText)    ? DBNull.Value : dto.HelpText.Trim());
            cmd.Parameters.AddWithValue("@Width",    dto.Width ?? "full");
            cmd.Parameters.AddWithValue("@Sort",     sortOrder);
            await cmd.ExecuteNonQueryAsync(ct);
            return Ok(new { id, sortOrder });
        }
        catch (Exception ex) { logger.LogError(ex, "AddField failed"); return StatusCode(500, ex.Message); }
    }

    [HttpPut("template/fields/{fieldId:guid}")]
    public async Task<IActionResult> UpdateField(Guid fieldId, [FromBody] TemplateFieldDto dto, CancellationToken ct)
    {
        if (!IsAdmin) return Forbid();
        if (string.IsNullOrWhiteSpace(dto.Label)) return BadRequest("Etiqueta requerida.");
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = """
                UPDATE dbo.TemplateFields
                SET Label = @Label, FieldType = @FType, IsRequired = @Required,
                    Options = @Opts, Placeholder = @Ph, HelpText = @Help, Width = @Width
                WHERE Id = @Id
                """;
            cmd.Parameters.AddWithValue("@Id",       fieldId);
            cmd.Parameters.AddWithValue("@Label",    dto.Label.Trim());
            cmd.Parameters.AddWithValue("@FType",    dto.FieldType ?? "Text");
            cmd.Parameters.AddWithValue("@Required", dto.IsRequired);
            cmd.Parameters.AddWithValue("@Opts",     string.IsNullOrWhiteSpace(dto.Options)     ? DBNull.Value : dto.Options.Trim());
            cmd.Parameters.AddWithValue("@Ph",       string.IsNullOrWhiteSpace(dto.Placeholder) ? DBNull.Value : dto.Placeholder.Trim());
            cmd.Parameters.AddWithValue("@Help",     string.IsNullOrWhiteSpace(dto.HelpText)    ? DBNull.Value : dto.HelpText.Trim());
            cmd.Parameters.AddWithValue("@Width",    dto.Width ?? "full");
            int rows = await cmd.ExecuteNonQueryAsync(ct);
            if (rows == 0) return NotFound("Campo no encontrado.");
            return Ok();
        }
        catch (Exception ex) { logger.LogError(ex, "UpdateField {Id}", fieldId); return StatusCode(500, ex.Message); }
    }

    [HttpDelete("template/fields/{fieldId:guid}")]
    public async Task<IActionResult> DeleteField(Guid fieldId, CancellationToken ct)
    {
        if (!IsAdmin) return Forbid();
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = "DELETE FROM dbo.TemplateFields WHERE Id = @Id";
            cmd.Parameters.AddWithValue("@Id", fieldId);
            int rows = await cmd.ExecuteNonQueryAsync(ct);
            if (rows == 0) return NotFound("Campo no encontrado.");
            return NoContent();
        }
        catch (Exception ex) { logger.LogError(ex, "DeleteField {Id}", fieldId); return StatusCode(500, ex.Message); }
    }

    [HttpPut("template/fields/reorder")]
    public async Task<IActionResult> ReorderFields([FromBody] List<ReorderDto> items, CancellationToken ct)
    {
        if (!IsAdmin) return Forbid();
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            foreach (var item in items)
            {
                await using var cmd = conn.CreateCommand();
                cmd.CommandText = "UPDATE dbo.TemplateFields SET SortOrder = @Order WHERE Id = @Id";
                cmd.Parameters.AddWithValue("@Order", item.SortOrder);
                cmd.Parameters.AddWithValue("@Id",    item.Id);
                await cmd.ExecuteNonQueryAsync(ct);
            }
            return Ok();
        }
        catch (Exception ex) { logger.LogError(ex, "ReorderFields failed"); return StatusCode(500, ex.Message); }
    }

    // ════════════════════════════════════════════════════════════════════════════
    //  TICKETS
    // ════════════════════════════════════════════════════════════════════════════

    [HttpGet]
    public async Task<IActionResult> GetTickets(
        [FromQuery] string? status   = null,
        [FromQuery] string? priority = null,
        [FromQuery] string? area     = null,
        [FromQuery] string? search   = null,
        CancellationToken ct = default)
    {
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await EnsureTablesAsync(conn, ct);

            var where = new List<string>();
            if (!IsAdmin)                              where.Add("SubmittedBy = @User");
            if (!string.IsNullOrWhiteSpace(status))   where.Add("Status = @Status");
            if (!string.IsNullOrWhiteSpace(priority)) where.Add("Priority = @Priority");
            if (!string.IsNullOrWhiteSpace(area))     where.Add("(Area = @Area OR Department = @Area)");
            if (!string.IsNullOrWhiteSpace(search))   where.Add("(Title LIKE @Search OR TicketNumber LIKE @Search OR Department LIKE @Search OR Area LIKE @Search)");

            var sql = $"""
                SELECT Id, TicketNumber, Title, Status, Priority, Area, Department, AssignedTo, SubmittedBy, CreatedAt, UpdatedAt
                FROM dbo.Tickets
                {(where.Count > 0 ? "WHERE " + string.Join(" AND ", where) : "")}
                ORDER BY CreatedAt DESC
                """;

            await using var cmd = conn.CreateCommand();
            cmd.CommandText = sql;
            if (!IsAdmin)                              cmd.Parameters.AddWithValue("@User",     CurrentUser);
            if (!string.IsNullOrWhiteSpace(status))   cmd.Parameters.AddWithValue("@Status",   status);
            if (!string.IsNullOrWhiteSpace(priority)) cmd.Parameters.AddWithValue("@Priority", priority);
            if (!string.IsNullOrWhiteSpace(area))     cmd.Parameters.AddWithValue("@Area",     area);
            if (!string.IsNullOrWhiteSpace(search))   cmd.Parameters.AddWithValue("@Search",   $"%{search}%");

            var list = new List<object>();
            await using var r = await cmd.ExecuteReaderAsync(ct);
            while (await r.ReadAsync(ct))
                list.Add(new
                {
                    id           = r.GetGuid(r.GetOrdinal("Id")),
                    ticketNumber = r.GetString(r.GetOrdinal("TicketNumber")),
                    title        = r.GetString(r.GetOrdinal("Title")),
                    status       = r.GetString(r.GetOrdinal("Status")),
                    priority     = r.GetString(r.GetOrdinal("Priority")),
                    area         = Ns(r, "Area"),
                    department   = Ns(r, "Department"),
                    assignedTo   = Ns(r, "AssignedTo"),
                    submittedBy  = r.GetString(r.GetOrdinal("SubmittedBy")),
                    createdAt    = r.GetDateTime(r.GetOrdinal("CreatedAt")),
                    updatedAt    = N(r, "UpdatedAt"),
                });
            return Ok(list);
        }
        catch (Exception ex) { logger.LogError(ex, "GetTickets failed"); return StatusCode(500, ex.Message); }
    }

    [HttpGet("{id:guid}")]
    public async Task<IActionResult> GetTicket(Guid id, CancellationToken ct)
    {
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);

            await using var cmd = conn.CreateCommand();
            cmd.CommandText = """
                SELECT Id, TicketNumber, Title, TemplateId, Status, Priority, Department,
                       AssignedTo, SubmittedBy, SubmittedByEmail, DelayNote, CloseNote, CreatedAt, UpdatedAt
                FROM dbo.Tickets WHERE Id = @Id
                """;
            cmd.Parameters.AddWithValue("@Id", id);
            await using var r = await cmd.ExecuteReaderAsync(ct);
            if (!await r.ReadAsync(ct)) return NotFound("Ticket no encontrado.");

            var ticket = new
            {
                id               = r.GetGuid(r.GetOrdinal("Id")),
                ticketNumber     = r.GetString(r.GetOrdinal("TicketNumber")),
                title            = r.GetString(r.GetOrdinal("Title")),
                templateId       = r.GetGuid(r.GetOrdinal("TemplateId")),
                status           = r.GetString(r.GetOrdinal("Status")),
                priority         = r.GetString(r.GetOrdinal("Priority")),
                department       = Ns(r, "Department"),
                assignedTo       = Ns(r, "AssignedTo"),
                submittedBy      = r.GetString(r.GetOrdinal("SubmittedBy")),
                submittedByEmail = Ns(r, "SubmittedByEmail"),
                delayNote        = Ns(r, "DelayNote"),
                closeNote        = Ns(r, "CloseNote"),
                createdAt        = r.GetDateTime(r.GetOrdinal("CreatedAt")),
                updatedAt        = N(r, "UpdatedAt"),
            };
            await r.CloseAsync();

            // Field values with labels
            await using var cmd2 = conn.CreateCommand();
            cmd2.CommandText = """
                SELECT fv.FieldId, fv.Value, tf.Label, tf.FieldType, tf.SortOrder
                FROM dbo.TicketFieldValues fv
                JOIN dbo.TemplateFields tf ON fv.FieldId = tf.Id
                WHERE fv.TicketId = @Id
                ORDER BY tf.SortOrder
                """;
            cmd2.Parameters.AddWithValue("@Id", id);
            var fieldValues = new List<object>();
            await using var r2 = await cmd2.ExecuteReaderAsync(ct);
            while (await r2.ReadAsync(ct))
                fieldValues.Add(new
                {
                    fieldId   = r2.GetGuid(r2.GetOrdinal("FieldId")),
                    label     = r2.GetString(r2.GetOrdinal("Label")),
                    fieldType = r2.GetString(r2.GetOrdinal("FieldType")),
                    value     = Ns(r2, "Value"),
                    sortOrder = r2.GetInt32(r2.GetOrdinal("SortOrder")),
                });
            await r2.CloseAsync();

            // Attachments
            await using var cmd3 = conn.CreateCommand();
            cmd3.CommandText = """
                SELECT Id, OriginalName, ContentType, FileSize, UploadedAt
                FROM dbo.TicketAttachments WHERE TicketId = @Id ORDER BY UploadedAt
                """;
            cmd3.Parameters.AddWithValue("@Id", id);
            var attachments = new List<object>();
            await using var r3 = await cmd3.ExecuteReaderAsync(ct);
            while (await r3.ReadAsync(ct))
                attachments.Add(new
                {
                    id           = r3.GetGuid(r3.GetOrdinal("Id")),
                    originalName = r3.GetString(r3.GetOrdinal("OriginalName")),
                    contentType  = r3.GetString(r3.GetOrdinal("ContentType")),
                    fileSize     = r3.GetInt64(r3.GetOrdinal("FileSize")),
                    uploadedAt   = r3.GetDateTime(r3.GetOrdinal("UploadedAt")),
                });
            await r3.CloseAsync();

            // Comments
            await using var cmd4 = conn.CreateCommand();
            cmd4.CommandText = """
                SELECT Id, AuthorName, IsAdmin, Body, CreatedAt
                FROM dbo.TicketComments WHERE TicketId = @Id ORDER BY CreatedAt
                """;
            cmd4.Parameters.AddWithValue("@Id", id);
            var comments = new List<object>();
            await using var r4 = await cmd4.ExecuteReaderAsync(ct);
            while (await r4.ReadAsync(ct))
                comments.Add(new
                {
                    id         = r4.GetGuid(r4.GetOrdinal("Id")),
                    authorName = r4.GetString(r4.GetOrdinal("AuthorName")),
                    isAdmin    = r4.GetBoolean(r4.GetOrdinal("IsAdmin")),
                    body       = r4.GetString(r4.GetOrdinal("Body")),
                    createdAt  = r4.GetDateTime(r4.GetOrdinal("CreatedAt")),
                });

            return Ok(new { ticket, fieldValues, attachments, comments });
        }
        catch (Exception ex) { logger.LogError(ex, "GetTicket {Id}", id); return StatusCode(500, ex.Message); }
    }

    [HttpPost]
    [EnableRateLimiting("tickets-policy")]
    [RequestSizeLimit(55_000_000)]
    [Consumes("multipart/form-data")]
    public async Task<IActionResult> CreateTicket([FromForm] CreateTicketFormDto dto, CancellationToken ct)
    {
        if (string.IsNullOrWhiteSpace(dto.Title)) return BadRequest("Título requerido.");

        var files = dto.Files ?? [];
        if (files.Count > 5) return BadRequest("Máximo 5 archivos adjuntos.");
        foreach (var f in files)
        {
            if (f.Length > 10 * 1024 * 1024) return BadRequest($"El archivo '{f.FileName}' excede el límite de 10 MB.");
            if (!f.ContentType.StartsWith("image/"))
                return BadRequest($"Solo se permiten imágenes. Archivo no válido: {f.FileName}");
        }

        var id = Guid.NewGuid();
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await EnsureTablesAsync(conn, ct);

            await using var cmdTpl = conn.CreateCommand();
            cmdTpl.CommandText = "SELECT TOP 1 Id FROM dbo.TicketTemplate WHERE IsActive = 1";
            var tplId = (Guid?)await cmdTpl.ExecuteScalarAsync(ct);
            if (tplId == null) return BadRequest("No hay plantilla de tickets activa.");

            await using var cmdNum = conn.CreateCommand();
            cmdNum.CommandText = "SELECT ISNULL(COUNT(*), 0) FROM dbo.Tickets";
            var count = Convert.ToInt32(await cmdNum.ExecuteScalarAsync(ct) ?? 0);
            var ticketNumber = $"TKT-{(count + 1):D4}";

            await using var cmd = conn.CreateCommand();
            cmd.CommandText = """
                INSERT INTO dbo.Tickets
                    (Id, TicketNumber, Title, TemplateId, Status, Priority, Area, Department,
                     AssignedTo, SubmittedBy, SubmittedByEmail, CreatedAt)
                VALUES
                    (@Id, @Num, @Title, @TplId, 'Abierto', @Priority, @Area, @Dept,
                     @Assigned, @User, @Email, GETUTCDATE())
                """;
            cmd.Parameters.AddWithValue("@Id",       id);
            cmd.Parameters.AddWithValue("@Num",      ticketNumber);
            cmd.Parameters.AddWithValue("@Title",    dto.Title.Trim());
            cmd.Parameters.AddWithValue("@TplId",    tplId.Value);
            cmd.Parameters.AddWithValue("@Priority", string.IsNullOrWhiteSpace(dto.Priority) ? "Media" : dto.Priority);
            cmd.Parameters.AddWithValue("@Area",     string.IsNullOrWhiteSpace(dto.Area)           ? DBNull.Value : (object)dto.Area.Trim());
            cmd.Parameters.AddWithValue("@Dept",     string.IsNullOrWhiteSpace(dto.Department)     ? DBNull.Value : (object)dto.Department.Trim());
            cmd.Parameters.AddWithValue("@Assigned", string.IsNullOrWhiteSpace(dto.AssignedTo)     ? DBNull.Value : (object)dto.AssignedTo.Trim());
            cmd.Parameters.AddWithValue("@User",     CurrentUser);
            cmd.Parameters.AddWithValue("@Email",    string.IsNullOrWhiteSpace(dto.SubmittedByEmail) ? DBNull.Value : (object)dto.SubmittedByEmail.Trim());
            await cmd.ExecuteNonQueryAsync(ct);

            // Field values
            if (!string.IsNullOrWhiteSpace(dto.FieldValuesJson))
            {
                var fieldValues = JsonSerializer.Deserialize<Dictionary<string, string>>(
                    dto.FieldValuesJson,
                    new JsonSerializerOptions { PropertyNameCaseInsensitive = true });

                if (fieldValues != null)
                {
                    foreach (var (fieldId, value) in fieldValues)
                    {
                        if (!Guid.TryParse(fieldId, out var fGuid) || string.IsNullOrWhiteSpace(value)) continue;
                        await using var cmdFv = conn.CreateCommand();
                        cmdFv.CommandText = """
                            INSERT INTO dbo.TicketFieldValues (Id, TicketId, FieldId, Value)
                            VALUES (NEWID(), @TicketId, @FieldId, @Value)
                            """;
                        cmdFv.Parameters.AddWithValue("@TicketId", id);
                        cmdFv.Parameters.AddWithValue("@FieldId",  fGuid);
                        cmdFv.Parameters.AddWithValue("@Value",    value.Trim());
                        await cmdFv.ExecuteNonQueryAsync(ct);
                    }
                }
            }

            // Save attachments
            var storagePath = AttachmentsPath();
            foreach (var file in files)
            {
                var ext        = Path.GetExtension(file.FileName);
                var storedName = $"{Guid.NewGuid()}{ext}";
                var filePath   = Path.Combine(storagePath, storedName);

                await using var fs = System.IO.File.Create(filePath);
                await file.CopyToAsync(fs, ct);

                await using var cmdA = conn.CreateCommand();
                cmdA.CommandText = """
                    INSERT INTO dbo.TicketAttachments (Id, TicketId, OriginalName, StoredName, ContentType, FileSize)
                    VALUES (NEWID(), @TicketId, @OrigName, @StoredName, @CType, @Size)
                    """;
                cmdA.Parameters.AddWithValue("@TicketId",   id);
                cmdA.Parameters.AddWithValue("@OrigName",   file.FileName);
                cmdA.Parameters.AddWithValue("@StoredName", storedName);
                cmdA.Parameters.AddWithValue("@CType",      file.ContentType);
                cmdA.Parameters.AddWithValue("@Size",       file.Length);
                await cmdA.ExecuteNonQueryAsync(ct);
            }

            logger.LogInformation("Ticket created: {Num} — {Title} by {User}", ticketNumber, dto.Title, CurrentUser);
            // Notificación a correo general del sistema
            _ = Task.Run(() => SendCreatedEmailAsync(ticketNumber, dto.Title, dto.Area ?? dto.Department ?? "—", CurrentUser), CancellationToken.None);
            // Notificación al correo configurado del área
            if (!string.IsNullOrWhiteSpace(dto.Area))
                _ = Task.Run(() => SendAreaNotificationAsync(ticketNumber, dto.Title, dto.Area, CurrentUser, dto.SubmittedByEmail), CancellationToken.None);

            return Ok(new { id, ticketNumber });
        }
        catch (Exception ex) { logger.LogError(ex, "CreateTicket failed"); return StatusCode(500, ex.Message); }
    }

    [HttpPut("{id:guid}/status")]
    public async Task<IActionResult> UpdateStatus(Guid id, [FromBody] UpdateTicketStatusDto dto, CancellationToken ct)
    {
        if (!IsAdmin) return Forbid();
        if (string.IsNullOrWhiteSpace(dto.Status)) return BadRequest("Estado requerido.");
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);

            await using var cmdGet = conn.CreateCommand();
            cmdGet.CommandText = "SELECT TicketNumber, Title, SubmittedByEmail FROM dbo.Tickets WHERE Id = @Id";
            cmdGet.Parameters.AddWithValue("@Id", id);
            await using var r = await cmdGet.ExecuteReaderAsync(ct);
            if (!await r.ReadAsync(ct)) return NotFound("Ticket no encontrado.");
            var ticketNumber   = r.GetString(r.GetOrdinal("TicketNumber"));
            var title          = r.GetString(r.GetOrdinal("Title"));
            var submitterEmail = Ns(r, "SubmittedByEmail");
            await r.CloseAsync();

            await using var cmd = conn.CreateCommand();
            cmd.CommandText = """
                UPDATE dbo.Tickets
                SET Status = @Status, AssignedTo = @AssignedTo, Department = @Dept,
                    DelayNote = @DelayNote, CloseNote = @CloseNote, UpdatedAt = GETUTCDATE()
                WHERE Id = @Id
                """;
            cmd.Parameters.AddWithValue("@Id",        id);
            cmd.Parameters.AddWithValue("@Status",    dto.Status.Trim());
            cmd.Parameters.AddWithValue("@AssignedTo", string.IsNullOrWhiteSpace(dto.AssignedTo) ? DBNull.Value : dto.AssignedTo.Trim());
            cmd.Parameters.AddWithValue("@Dept",       string.IsNullOrWhiteSpace(dto.Department) ? DBNull.Value : dto.Department.Trim());
            cmd.Parameters.AddWithValue("@DelayNote",  string.IsNullOrWhiteSpace(dto.DelayNote)  ? DBNull.Value : dto.DelayNote.Trim());
            cmd.Parameters.AddWithValue("@CloseNote",  string.IsNullOrWhiteSpace(dto.CloseNote)  ? DBNull.Value : dto.CloseNote.Trim());
            await cmd.ExecuteNonQueryAsync(ct);

            if (!string.IsNullOrWhiteSpace(dto.Notes))
            {
                await using var cmdC = conn.CreateCommand();
                cmdC.CommandText = """
                    INSERT INTO dbo.TicketComments (Id, TicketId, AuthorName, IsAdmin, Body)
                    VALUES (NEWID(), @TicketId, @Author, 1, @Body)
                    """;
                cmdC.Parameters.AddWithValue("@TicketId", id);
                cmdC.Parameters.AddWithValue("@Author",   CurrentUser);
                cmdC.Parameters.AddWithValue("@Body",     $"[Cambio de estatus → {dto.Status}] {dto.Notes}");
                await cmdC.ExecuteNonQueryAsync(ct);
            }

            if (!string.IsNullOrWhiteSpace(submitterEmail))
            {
                if (dto.Status == "En Espera" && !string.IsNullOrWhiteSpace(dto.DelayNote))
                    _ = Task.Run(() => SendDelayEmailAsync(submitterEmail, ticketNumber, title, dto.DelayNote!), CancellationToken.None);
                else if (dto.Status == "Resuelto")
                    _ = Task.Run(() => SendResolvedEmailAsync(submitterEmail, ticketNumber, title, dto.CloseNote ?? "Tu ticket ha sido resuelto."), CancellationToken.None);
            }

            return Ok();
        }
        catch (Exception ex) { logger.LogError(ex, "UpdateStatus {Id}", id); return StatusCode(500, ex.Message); }
    }

    [HttpPost("{id:guid}/comments")]
    public async Task<IActionResult> AddComment(Guid id, [FromBody] AddCommentDto dto, CancellationToken ct)
    {
        if (string.IsNullOrWhiteSpace(dto.Body)) return BadRequest("El comentario no puede estar vacío.");
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = """
                INSERT INTO dbo.TicketComments (Id, TicketId, AuthorName, IsAdmin, Body)
                VALUES (NEWID(), @TicketId, @Author, @IsAdmin, @Body)
                """;
            cmd.Parameters.AddWithValue("@TicketId", id);
            cmd.Parameters.AddWithValue("@Author",   CurrentUser);
            cmd.Parameters.AddWithValue("@IsAdmin",  IsAdmin);
            cmd.Parameters.AddWithValue("@Body",     dto.Body.Trim());
            await cmd.ExecuteNonQueryAsync(ct);
            return Ok();
        }
        catch (Exception ex) { logger.LogError(ex, "AddComment {Id}", id); return StatusCode(500, ex.Message); }
    }

    [HttpDelete("{id:guid}")]
    public async Task<IActionResult> DeleteTicket(Guid id, CancellationToken ct)
    {
        if (!IsAdmin) return Forbid();
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);

            await using var cmdA = conn.CreateCommand();
            cmdA.CommandText = "SELECT StoredName FROM dbo.TicketAttachments WHERE TicketId = @Id";
            cmdA.Parameters.AddWithValue("@Id", id);
            var storedNames = new List<string>();
            await using var r = await cmdA.ExecuteReaderAsync(ct);
            while (await r.ReadAsync(ct)) storedNames.Add(r.GetString(0));
            await r.CloseAsync();

            await using var cmd = conn.CreateCommand();
            cmd.CommandText = """
                DELETE FROM dbo.TicketComments    WHERE TicketId = @Id;
                DELETE FROM dbo.TicketAttachments WHERE TicketId = @Id;
                DELETE FROM dbo.TicketFieldValues WHERE TicketId = @Id;
                DELETE FROM dbo.Tickets           WHERE Id       = @Id;
                """;
            cmd.Parameters.AddWithValue("@Id", id);
            await cmd.ExecuteNonQueryAsync(ct);

            var storagePath = AttachmentsPath();
            foreach (var name in storedNames)
            {
                var fp = Path.Combine(storagePath, name);
                if (System.IO.File.Exists(fp)) System.IO.File.Delete(fp);
            }
            return NoContent();
        }
        catch (Exception ex) { logger.LogError(ex, "DeleteTicket {Id}", id); return StatusCode(500, ex.Message); }
    }

    [HttpGet("{id:guid}/attachments/{attachId:guid}")]
    public async Task<IActionResult> GetAttachment(Guid id, Guid attachId, CancellationToken ct)
    {
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = "SELECT OriginalName, StoredName, ContentType FROM dbo.TicketAttachments WHERE Id = @Id AND TicketId = @TicketId";
            cmd.Parameters.AddWithValue("@Id",       attachId);
            cmd.Parameters.AddWithValue("@TicketId", id);
            await using var r = await cmd.ExecuteReaderAsync(ct);
            if (!await r.ReadAsync(ct)) return NotFound("Archivo no encontrado.");
            var originalName = r.GetString(r.GetOrdinal("OriginalName"));
            var storedName   = r.GetString(r.GetOrdinal("StoredName"));
            var contentType  = r.GetString(r.GetOrdinal("ContentType"));
            await r.CloseAsync();

            var filePath = Path.Combine(AttachmentsPath(), storedName);
            if (!System.IO.File.Exists(filePath)) return NotFound("Archivo no encontrado en disco.");
            var bytes = await System.IO.File.ReadAllBytesAsync(filePath, ct);
            return File(bytes, contentType, originalName);
        }
        catch (Exception ex) { logger.LogError(ex, "GetAttachment {AttachId}", attachId); return StatusCode(500, ex.Message); }
    }

    // ════════════════════════════════════════════════════════════════════════════
    //  EMAIL NOTIFICATIONS
    // ════════════════════════════════════════════════════════════════════════════

    private (SmtpClient? client, MailMessage? msg) BuildEmail(string subject, string toEmail)
    {
        var smtp     = config.GetSection("SmtpSettings");
        var host     = smtp["Host"]     ?? "";
        var port     = int.TryParse(smtp["Port"], out var p) ? p : 587;
        var user     = smtp["Username"] ?? "";
        var pass     = smtp["Password"] ?? "";
        var from     = smtp["FromEmail"] ?? "";
        var fromName = smtp["FromName"]  ?? "Pandora";
        var useSsl   = smtp["UseSsl"]?.ToLower() == "true";

        if (string.IsNullOrWhiteSpace(host) || string.IsNullOrWhiteSpace(from) || string.IsNullOrWhiteSpace(toEmail))
            return (null, null);

        var client = new SmtpClient(host, port)
        {
            Credentials = new NetworkCredential(user, pass),
            EnableSsl   = useSsl,
            Timeout     = 10_000,
        };
        var msg = new MailMessage { From = new MailAddress(from, fromName), Subject = subject, IsBodyHtml = true };
        msg.To.Add(toEmail);
        return (client, msg);
    }

    private string NotifEmail => config.GetSection("SmtpSettings")["NotificationsEmail"]
                               ?? config.GetSection("SmtpSettings")["FromEmail"]
                               ?? "";

    // ════════════════════════════════════════════════════════════════════════════
    //  AREA CONFIGS — correos de notificación por área
    // ════════════════════════════════════════════════════════════════════════════

    // ── GET /api/tickets/positions — lista pública de puestos (para el formulario) ──
    [HttpGet("positions")]
    public async Task<IActionResult> GetPositions(CancellationToken ct)
    {
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await EnsureTablesAsync(conn, ct);
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = "SELECT Area FROM dbo.TicketAreaConfigs ORDER BY Area";
            var list = new List<string>();
            await using var r = await cmd.ExecuteReaderAsync(ct);
            while (await r.ReadAsync(ct))
                list.Add(r.GetString(0));
            return Ok(list);
        }
        catch (Exception ex) { logger.LogError(ex, "GetPositions"); return StatusCode(500, ex.Message); }
    }

    [HttpGet("area-configs")]
    [Authorize(Roles = "Admin")]
    public async Task<IActionResult> GetAreaConfigs(CancellationToken ct)
    {
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await EnsureTablesAsync(conn, ct);
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = "SELECT Id, Area, NotificationEmail, UpdatedAt FROM dbo.TicketAreaConfigs ORDER BY Area";
            var list = new List<object>();
            await using var r = await cmd.ExecuteReaderAsync(ct);
            while (await r.ReadAsync(ct))
                list.Add(new
                {
                    id                = r.GetGuid(r.GetOrdinal("Id")),
                    area              = r.GetString(r.GetOrdinal("Area")),
                    notificationEmail = Ns(r, "NotificationEmail"),
                    updatedAt         = N(r, "UpdatedAt"),
                });
            return Ok(list);
        }
        catch (Exception ex) { logger.LogError(ex, "GetAreaConfigs"); return StatusCode(500, ex.Message); }
    }

    [HttpPut("area-configs")]
    [Authorize(Roles = "Admin")]
    public async Task<IActionResult> UpdateAreaConfigs([FromBody] List<AreaConfigDto> configs, CancellationToken ct)
    {
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await EnsureTablesAsync(conn, ct);
            foreach (var cfg in configs)
            {
                if (string.IsNullOrWhiteSpace(cfg.Area)) continue;
                await using var cmd = conn.CreateCommand();
                cmd.CommandText = """
                    UPDATE dbo.TicketAreaConfigs
                    SET NotificationEmail = @Email, UpdatedAt = GETUTCDATE()
                    WHERE Area = @Area
                    """;
                cmd.Parameters.AddWithValue("@Email", string.IsNullOrWhiteSpace(cfg.NotificationEmail)
                    ? DBNull.Value : (object)cfg.NotificationEmail.Trim());
                cmd.Parameters.AddWithValue("@Area", cfg.Area);
                await cmd.ExecuteNonQueryAsync(ct);
            }
            return Ok();
        }
        catch (Exception ex) { logger.LogError(ex, "UpdateAreaConfigs"); return StatusCode(500, ex.Message); }
    }

    [HttpPost("area-configs")]
    [Authorize(Roles = "Admin")]
    public async Task<IActionResult> CreateAreaConfig([FromBody] CreateAreaConfigDto dto, CancellationToken ct)
    {
        if (string.IsNullOrWhiteSpace(dto.Area)) return BadRequest("El nombre del puesto es requerido.");
        var id = Guid.NewGuid();
        var area = dto.Area.Trim();
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await EnsureTablesAsync(conn, ct);

            await using var cmdCheck = conn.CreateCommand();
            cmdCheck.CommandText = "SELECT COUNT(1) FROM dbo.TicketAreaConfigs WHERE Area = @Area";
            cmdCheck.Parameters.AddWithValue("@Area", area);
            var exists = Convert.ToInt32(await cmdCheck.ExecuteScalarAsync(ct) ?? 0);
            if (exists > 0) return Conflict("Ya existe un puesto con ese nombre.");

            await using var cmd = conn.CreateCommand();
            cmd.CommandText = "INSERT INTO dbo.TicketAreaConfigs (Id, Area) VALUES (@Id, @Area)";
            cmd.Parameters.AddWithValue("@Id",   id);
            cmd.Parameters.AddWithValue("@Area", area);
            await cmd.ExecuteNonQueryAsync(ct);

            logger.LogInformation("Puesto creado: {Area} por {User}", area, CurrentUser);
            return Ok(new { id, area, notificationEmail = (string?)null });
        }
        catch (Exception ex) { logger.LogError(ex, "CreateAreaConfig"); return StatusCode(500, ex.Message); }
    }

    [HttpDelete("area-configs/{id:guid}")]
    [Authorize(Roles = "Admin")]
    public async Task<IActionResult> DeleteAreaConfig(Guid id, CancellationToken ct)
    {
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);

            // Verificar si hay tickets que referencian este puesto
            await using var cmdCheck = conn.CreateCommand();
            cmdCheck.CommandText = """
                SELECT TOP 1 1 FROM dbo.Tickets t
                JOIN dbo.TicketAreaConfigs tac ON t.Area = tac.Area
                WHERE tac.Id = @Id
                """;
            cmdCheck.Parameters.AddWithValue("@Id", id);
            var inUse = await cmdCheck.ExecuteScalarAsync(ct);
            if (inUse != null)
                return Conflict("No se puede eliminar: existen tickets asociados a este puesto.");

            await using var cmd = conn.CreateCommand();
            cmd.CommandText = "DELETE FROM dbo.TicketAreaConfigs WHERE Id = @Id";
            cmd.Parameters.AddWithValue("@Id", id);
            int rows = await cmd.ExecuteNonQueryAsync(ct);
            if (rows == 0) return NotFound("Puesto no encontrado.");

            logger.LogInformation("Puesto eliminado Id={Id} por {User}", id, CurrentUser);
            return NoContent();
        }
        catch (Exception ex) { logger.LogError(ex, "DeleteAreaConfig {Id}", id); return StatusCode(500, ex.Message); }
    }

    private async Task SendAreaNotificationAsync(string ticketNumber, string title, string area, string submittedBy, string? submittedByEmail)
    {
        try
        {
            string? areaEmail = null;
            await using var conn2 = Conn();
            await conn2.OpenAsync();
            await using var cmd = conn2.CreateCommand();
            cmd.CommandText = "SELECT NotificationEmail FROM dbo.TicketAreaConfigs WHERE Area = @Area";
            cmd.Parameters.AddWithValue("@Area", area);
            areaEmail = (string?)await cmd.ExecuteScalarAsync();
            if (string.IsNullOrWhiteSpace(areaEmail)) return;

            var (client, msg) = BuildEmail($"[HelpDesk] Nuevo ticket {ticketNumber} — {area}", areaEmail);
            if (client == null || msg == null) return;
            msg.Body = $"""
                <html><body style="font-family:Arial,sans-serif;font-size:14px;color:#333">
                <div style="max-width:600px;margin:0 auto">
                  <div style="background:#1a237e;padding:20px;border-radius:8px 8px 0 0">
                    <h2 style="color:white;margin:0">🎫 Nuevo ticket — Área: {area}</h2>
                  </div>
                  <div style="border:1px solid #ddd;padding:24px;border-radius:0 0 8px 8px">
                    <table style="width:100%;border-collapse:collapse">
                      <tr style="border-bottom:1px solid #eee"><td style="padding:8px;font-weight:bold;color:#555;width:40%">Ticket</td><td style="padding:8px"><strong>{ticketNumber}</strong></td></tr>
                      <tr style="border-bottom:1px solid #eee"><td style="padding:8px;font-weight:bold;color:#555">Título</td><td style="padding:8px">{System.Net.WebUtility.HtmlEncode(title)}</td></tr>
                      <tr style="border-bottom:1px solid #eee"><td style="padding:8px;font-weight:bold;color:#555">Área</td><td style="padding:8px">{area}</td></tr>
                      <tr><td style="padding:8px;font-weight:bold;color:#555">Solicitante</td><td style="padding:8px">{submittedBy}{(string.IsNullOrWhiteSpace(submittedByEmail) ? "" : $" &lt;{submittedByEmail}&gt;")}</td></tr>
                    </table>
                    <p style="margin-top:16px;font-size:12px;color:#999">Por favor atiende este ticket a la brevedad. — Pandora HelpDesk iMET</p>
                  </div>
                </div></body></html>
                """;
            await client.SendMailAsync(msg);
        }
        catch (Exception ex) { logger.LogWarning("SendAreaNotification failed: {Msg}", ex.Message); }
    }

    private async Task SendCreatedEmailAsync(string ticketNumber, string title, string department, string submittedBy)
    {
        try
        {
            var (client, msg) = BuildEmail($"[Pandora Tickets] Nuevo ticket — {ticketNumber}", NotifEmail);
            if (client == null || msg == null) return;
            msg.Body = $"""
                <html><body style="font-family:Arial,sans-serif;font-size:14px;color:#333">
                <div style="max-width:600px;margin:0 auto">
                  <div style="background:#1a237e;padding:20px;border-radius:8px 8px 0 0">
                    <h2 style="color:white;margin:0">🎫 Nuevo ticket recibido</h2>
                  </div>
                  <div style="border:1px solid #ddd;padding:24px;border-radius:0 0 8px 8px">
                    <table style="width:100%;border-collapse:collapse">
                      <tr style="border-bottom:1px solid #eee"><td style="padding:8px;font-weight:bold;color:#555;width:40%">Ticket</td><td style="padding:8px"><strong>{ticketNumber}</strong></td></tr>
                      <tr style="border-bottom:1px solid #eee"><td style="padding:8px;font-weight:bold;color:#555">Título</td><td style="padding:8px">{title}</td></tr>
                      <tr style="border-bottom:1px solid #eee"><td style="padding:8px;font-weight:bold;color:#555">Departamento</td><td style="padding:8px">{department}</td></tr>
                      <tr><td style="padding:8px;font-weight:bold;color:#555">Solicitante</td><td style="padding:8px">{submittedBy}</td></tr>
                    </table>
                    <div style="margin-top:20px;text-align:center">
                      <a href="http://localhost:3000/tickets" style="background:#1a237e;color:white;padding:10px 24px;border-radius:6px;text-decoration:none;font-weight:bold">Ver tickets →</a>
                    </div>
                    <p style="margin-top:20px;font-size:12px;color:#999;text-align:center">Pandora — Sistema de Gestión iMET</p>
                  </div>
                </div></body></html>
                """;
            await client.SendMailAsync(msg);
        }
        catch (Exception ex) { logger.LogWarning("SendCreatedEmail failed: {Msg}", ex.Message); }
    }

    private async Task SendDelayEmailAsync(string toEmail, string ticketNumber, string title, string delayNote)
    {
        try
        {
            var (client, msg) = BuildEmail($"[Pandora Tickets] Actualización — {ticketNumber}: en espera", toEmail);
            if (client == null || msg == null) return;
            msg.Body = $"""
                <html><body style="font-family:Arial,sans-serif;font-size:14px;color:#333">
                <div style="max-width:600px;margin:0 auto">
                  <div style="background:#e65100;padding:20px;border-radius:8px 8px 0 0">
                    <h2 style="color:white;margin:0">⏳ Tu ticket está en espera</h2>
                  </div>
                  <div style="border:1px solid #ddd;padding:24px;border-radius:0 0 8px 8px">
                    <p>Tu ticket <strong>{ticketNumber}</strong> — <em>{title}</em> ha pasado al estatus <strong>En Espera</strong>.</p>
                    <div style="background:#fff3e0;border-left:4px solid #e65100;padding:12px 16px;border-radius:4px;margin:16px 0">
                      <strong>Motivo de la espera:</strong><br/>{delayNote}
                    </div>
                    <p>Nos comunicaremos contigo en cuanto tengamos una actualización. Disculpa la espera.</p>
                    <p style="font-size:12px;color:#999;margin-top:20px">Pandora — Sistema de Gestión iMET</p>
                  </div>
                </div></body></html>
                """;
            await client.SendMailAsync(msg);
        }
        catch (Exception ex) { logger.LogWarning("SendDelayEmail failed: {Msg}", ex.Message); }
    }

    private async Task SendResolvedEmailAsync(string toEmail, string ticketNumber, string title, string closeNote)
    {
        try
        {
            var (client, msg) = BuildEmail($"[Pandora Tickets] Resuelto — {ticketNumber}", toEmail);
            if (client == null || msg == null) return;
            msg.Body = $"""
                <html><body style="font-family:Arial,sans-serif;font-size:14px;color:#333">
                <div style="max-width:600px;margin:0 auto">
                  <div style="background:#1b5e20;padding:20px;border-radius:8px 8px 0 0">
                    <h2 style="color:white;margin:0">✅ Ticket resuelto</h2>
                  </div>
                  <div style="border:1px solid #ddd;padding:24px;border-radius:0 0 8px 8px">
                    <p>Tu ticket <strong>{ticketNumber}</strong> — <em>{title}</em> ha sido marcado como <strong>Resuelto</strong>.</p>
                    <div style="background:#e8f5e9;border-left:4px solid #1b5e20;padding:12px 16px;border-radius:4px;margin:16px 0">
                      <strong>Resolución:</strong><br/>{closeNote}
                    </div>
                    <p>Gracias por usar Pandora Tickets. Si el problema persiste, puedes abrir un nuevo ticket.</p>
                    <p style="font-size:12px;color:#999;margin-top:20px">Pandora — Sistema de Gestión iMET</p>
                  </div>
                </div></body></html>
                """;
            await client.SendMailAsync(msg);
        }
        catch (Exception ex) { logger.LogWarning("SendResolvedEmail failed: {Msg}", ex.Message); }
    }
}

// ── DTOs ──────────────────────────────────────────────────────────────────────

public record TemplateMetaDto(string Name, string? Description);

public record TemplateFieldDto(
    string  Label,
    string? FieldType,
    bool    IsRequired,
    string? Options,
    string? Placeholder,
    string? HelpText,
    string? Width
);

public record ReorderDto(Guid Id, int SortOrder);

public class CreateTicketFormDto
{
    public string           Title             { get; set; } = "";
    public string?          Priority          { get; set; }
    public string?          Area              { get; set; }
    public string?          Department        { get; set; }
    public string?          AssignedTo        { get; set; }
    public string?          SubmittedByEmail  { get; set; }
    public string?          FieldValuesJson   { get; set; }
    public List<IFormFile>? Files             { get; set; }
}

public record AreaConfigDto(string Area, string? NotificationEmail);

public record CreateAreaConfigDto(string Area);

public record UpdateTicketStatusDto(
    string  Status,
    string? AssignedTo,
    string? Department,
    string? Notes,
    string? DelayNote,
    string? CloseNote
);

public record AddCommentDto(string Body);
