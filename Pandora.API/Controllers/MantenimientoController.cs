using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using System.Security.Claims;
using System.Text;

namespace Pandora.API.Controllers;

[ApiController]
[Route("api/mantenimiento")]
[Authorize]
public class MantenimientoController(
    IConfiguration config,
    ILogger<MantenimientoController> logger,
    Pandora.API.Services.PushService push) : ControllerBase
{
    private SqlConnection Conn() => new(config.GetConnectionString("PandoraDb"));
    private string CurrentUser => User.FindFirstValue(ClaimTypes.Name)
                               ?? User.FindFirstValue("name") ?? "sistema";

    // ── GET /api/mantenimiento ────────────────────────────────────────────────
    [HttpGet]
    public async Task<IActionResult> GetAll(
        [FromQuery] string? search    = null,
        [FromQuery] string? estado    = null,
        [FromQuery] string? tipo      = null,
        [FromQuery] string? prioridad = null,
        [FromQuery] string? tecnico   = null,
        [FromQuery] DateTime? desde   = null,
        [FromQuery] DateTime? hasta   = null,
        CancellationToken ct = default)
    {
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await using var cmd = conn.CreateCommand();

            var where = new List<string> { "IsDeleted = 0" };
            if (!string.IsNullOrWhiteSpace(search))
            {
                where.Add("(Folio LIKE @s OR Titulo LIKE @s OR NombreEquipo LIKE @s OR TecnicoAsignado LIKE @s)");
                cmd.Parameters.AddWithValue("@s", $"%{search.Trim()}%");
            }
            if (!string.IsNullOrWhiteSpace(estado))    { where.Add("Estado = @est");    cmd.Parameters.AddWithValue("@est", estado); }
            if (!string.IsNullOrWhiteSpace(tipo))      { where.Add("TipoMantenimiento = @tipo"); cmd.Parameters.AddWithValue("@tipo", tipo); }
            if (!string.IsNullOrWhiteSpace(prioridad)) { where.Add("Prioridad = @prio"); cmd.Parameters.AddWithValue("@prio", prioridad); }
            if (!string.IsNullOrWhiteSpace(tecnico))   { where.Add("TecnicoAsignado LIKE @tec"); cmd.Parameters.AddWithValue("@tec", $"%{tecnico.Trim()}%"); }
            if (desde.HasValue) { where.Add("FechaProgramada >= @desde"); cmd.Parameters.AddWithValue("@desde", desde.Value); }
            if (hasta.HasValue) { where.Add("FechaProgramada <= @hasta"); cmd.Parameters.AddWithValue("@hasta", hasta.Value); }

            cmd.CommandText = $"""
                SELECT Id, Folio, Titulo, TipoMantenimiento, Estado, Prioridad,
                       NombreEquipo, Ubicacion, TecnicoAsignado,
                       FechaProgramada, FechaRealizada, DuracionMinutos,
                       CostoEstimado, CostoReal, CreadoEn
                FROM dbo.Mantenimientos
                WHERE {string.Join(" AND ", where)}
                ORDER BY FechaProgramada DESC
                """;

            var list = new List<object>();
            await using var r = await cmd.ExecuteReaderAsync(ct);
            while (await r.ReadAsync(ct))
                list.Add(MapRow(r));
            return Ok(list);
        }
        catch (Exception ex) { logger.LogError(ex, "Mantenimiento.GetAll"); return StatusCode(500, ex.Message); }
    }

    // ── GET /api/mantenimiento/{id} ───────────────────────────────────────────
    [HttpGet("{id:guid}")]
    public async Task<IActionResult> GetById(Guid id, CancellationToken ct)
    {
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);

            await using var cmd = conn.CreateCommand();
            cmd.CommandText = """
                SELECT Id, Folio, Titulo, Descripcion, TipoMantenimiento, Estado, Prioridad,
                       EquipoId, NombreEquipo, Ubicacion, TecnicoAsignado, EmailTecnico,
                       FechaProgramada, FechaRealizada, DuracionMinutos,
                       ChecklistJson, Notas, CostoEstimado, CostoReal,
                       CreadoPor, CreadoEn, ActualizadoEn
                FROM dbo.Mantenimientos
                WHERE Id = @Id AND IsDeleted = 0
                """;
            cmd.Parameters.AddWithValue("@Id", id);

            object? entry = null;
            await using (var r = await cmd.ExecuteReaderAsync(ct))
                if (await r.ReadAsync(ct)) entry = MapRowFull(r);

            if (entry is null) return NotFound();

            await using var cmdSeg = conn.CreateCommand();
            cmdSeg.CommandText = """
                SELECT Id, Nota, Tipo, CreadoPor, CreadoEn
                FROM dbo.MantenimientoSeguimientos
                WHERE MantenimientoId = @Id
                ORDER BY CreadoEn ASC
                """;
            cmdSeg.Parameters.AddWithValue("@Id", id);
            var seguimientos = new List<object>();
            await using (var rs = await cmdSeg.ExecuteReaderAsync(ct))
                while (await rs.ReadAsync(ct))
                    seguimientos.Add(new {
                        id        = rs.GetGuid(0),
                        nota      = rs.GetString(1),
                        tipo      = rs.GetString(2),
                        creadoPor = rs.GetString(3),
                        creadoEn  = rs.GetDateTime(4),
                    });

            await using var cmdEv = conn.CreateCommand();
            cmdEv.CommandText = """
                SELECT Id, NombreArchivo, Mime, SubidoPor, SubidoEn
                FROM dbo.MantenimientoEvidencias
                WHERE MantenimientoId = @Id
                ORDER BY SubidoEn ASC
                """;
            cmdEv.Parameters.AddWithValue("@Id", id);
            var evidencias = new List<object>();
            await using (var re = await cmdEv.ExecuteReaderAsync(ct))
                while (await re.ReadAsync(ct))
                    evidencias.Add(new {
                        id            = re.GetGuid(0),
                        nombreArchivo = re.GetString(1),
                        mime          = re.GetString(2),
                        subidoPor     = re.GetString(3),
                        subidoEn      = re.GetDateTime(4),
                    });

            return Ok(new { entry, seguimientos, evidencias });
        }
        catch (Exception ex) { logger.LogError(ex, "Mantenimiento.GetById {Id}", id); return StatusCode(500, ex.Message); }
    }

    // ── POST /api/mantenimiento ───────────────────────────────────────────────
    [HttpPost]
    public async Task<IActionResult> Create([FromBody] MantenimientoDto dto, CancellationToken ct)
    {
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);

            // Generar folio BIT-YYYY-NNNN
            await using var cmdFolio = conn.CreateCommand();
            cmdFolio.CommandText = """
                SELECT ISNULL(MAX(CAST(SUBSTRING(Folio, 10, 4) AS INT)), 0) + 1
                FROM dbo.Mantenimientos
                WHERE Folio LIKE @Prefix AND IsDeleted = 0
                """;
            cmdFolio.Parameters.AddWithValue("@Prefix", $"MNT-{DateTime.UtcNow.Year}-%");
            var seq   = Convert.ToInt32(await cmdFolio.ExecuteScalarAsync(ct));
            var folio = $"MNT-{DateTime.UtcNow.Year}-{seq:D4}";
            var id    = Guid.NewGuid();

            await using var cmd = conn.CreateCommand();
            cmd.CommandText = """
                INSERT INTO dbo.Mantenimientos
                    (Id, Folio, Titulo, Descripcion, TipoMantenimiento, Estado, Prioridad,
                     EquipoId, NombreEquipo, Ubicacion, TecnicoAsignado, EmailTecnico,
                     FechaProgramada, FechaRealizada, DuracionMinutos,
                     ChecklistJson, Notas, CostoEstimado, CostoReal, CreadoPor)
                VALUES
                    (@Id, @Folio, @Titulo, @Desc, @Tipo, @Estado, @Prio,
                     @EquipoId, @NombreEquipo, @Ubic, @Tecnico, @EmailTec,
                     @FechaProg, @FechaReal, @Dur,
                     @Checklist, @Notas, @CostoEst, @CostoReal, @CreadoPor)
                """;
            cmd.Parameters.AddWithValue("@Id",          id);
            cmd.Parameters.AddWithValue("@Folio",       folio);
            cmd.Parameters.AddWithValue("@Titulo",      dto.Titulo.Trim());
            cmd.Parameters.AddWithValue("@Desc",        (object?)dto.Descripcion    ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@Tipo",        dto.TipoMantenimiento ?? "Preventivo");
            cmd.Parameters.AddWithValue("@Estado",      dto.Estado            ?? "Programado");
            cmd.Parameters.AddWithValue("@Prio",        dto.Prioridad         ?? "Media");
            cmd.Parameters.AddWithValue("@EquipoId",    (object?)dto.EquipoId     ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@NombreEquipo",(object?)dto.NombreEquipo  ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@Ubic",        (object?)dto.Ubicacion      ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@Tecnico",     (object?)dto.TecnicoAsignado ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@EmailTec",    (object?)dto.EmailTecnico   ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@FechaProg",   dto.FechaProgramada);
            cmd.Parameters.AddWithValue("@FechaReal",   (object?)dto.FechaRealizada ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@Dur",         (object?)dto.DuracionMinutos ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@Checklist",   (object?)dto.ChecklistJson  ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@Notas",       (object?)dto.Notas          ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@CostoEst",    (object?)dto.CostoEstimado  ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@CostoReal",   (object?)dto.CostoReal      ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@CreadoPor",   CurrentUser);
            await cmd.ExecuteNonQueryAsync(ct);

            // Enviar email al técnico si tiene email
            if (!string.IsNullOrWhiteSpace(dto.EmailTecnico))
                _ = Task.Run(() => SendTecnicoEmailAsync(dto, folio, isNew: true), CancellationToken.None);
            // Push a todos los admins
            _ = Task.Run(() => push.SendToAllAsync(
                $"🔧 Nuevo mantenimiento {folio}",
                $"{dto.Titulo} — {dto.NombreEquipo ?? "Sin equipo"}",
                $"/mantenimiento/{id}"), CancellationToken.None);

            return StatusCode(201, new { id, folio });
        }
        catch (Exception ex) { logger.LogError(ex, "Mantenimiento.Create"); return StatusCode(500, ex.Message); }
    }

    // ── PUT /api/mantenimiento/{id} ───────────────────────────────────────────
    [HttpPut("{id:guid}")]
    public async Task<IActionResult> Update(Guid id, [FromBody] MantenimientoDto dto, CancellationToken ct)
    {
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = """
                UPDATE dbo.Mantenimientos SET
                    Titulo            = @Titulo,
                    Descripcion       = @Desc,
                    TipoMantenimiento = @Tipo,
                    Estado            = @Estado,
                    Prioridad         = @Prio,
                    NombreEquipo      = @NombreEquipo,
                    Ubicacion         = @Ubic,
                    TecnicoAsignado   = @Tecnico,
                    EmailTecnico      = @EmailTec,
                    FechaProgramada   = @FechaProg,
                    FechaRealizada    = @FechaReal,
                    DuracionMinutos   = @Dur,
                    ChecklistJson     = @Checklist,
                    Notas             = @Notas,
                    CostoEstimado     = @CostoEst,
                    CostoReal         = @CostoReal,
                    ActualizadoEn     = GETUTCDATE()
                WHERE Id = @Id AND IsDeleted = 0
                """;
            cmd.Parameters.AddWithValue("@Id",          id);
            cmd.Parameters.AddWithValue("@Titulo",      dto.Titulo.Trim());
            cmd.Parameters.AddWithValue("@Desc",        (object?)dto.Descripcion     ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@Tipo",        dto.TipoMantenimiento ?? "Preventivo");
            cmd.Parameters.AddWithValue("@Estado",      dto.Estado            ?? "Programado");
            cmd.Parameters.AddWithValue("@Prio",        dto.Prioridad         ?? "Media");
            cmd.Parameters.AddWithValue("@NombreEquipo",(object?)dto.NombreEquipo    ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@Ubic",        (object?)dto.Ubicacion       ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@Tecnico",     (object?)dto.TecnicoAsignado ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@EmailTec",    (object?)dto.EmailTecnico    ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@FechaProg",   dto.FechaProgramada);
            cmd.Parameters.AddWithValue("@FechaReal",   (object?)dto.FechaRealizada  ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@Dur",         (object?)dto.DuracionMinutos ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@Checklist",   (object?)dto.ChecklistJson   ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@Notas",       (object?)dto.Notas           ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@CostoEst",    (object?)dto.CostoEstimado   ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@CostoReal",   (object?)dto.CostoReal       ?? DBNull.Value);
            await cmd.ExecuteNonQueryAsync(ct);
            return NoContent();
        }
        catch (Exception ex) { logger.LogError(ex, "Mantenimiento.Update {Id}", id); return StatusCode(500, ex.Message); }
    }

    // ── DELETE /api/mantenimiento/{id} ────────────────────────────────────────
    [HttpDelete("{id:guid}")]
    public async Task<IActionResult> Delete(Guid id, CancellationToken ct)
    {
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = "UPDATE dbo.Mantenimientos SET IsDeleted=1 WHERE Id=@Id";
            cmd.Parameters.AddWithValue("@Id", id);
            await cmd.ExecuteNonQueryAsync(ct);
            return NoContent();
        }
        catch (Exception ex) { logger.LogError(ex, "Mantenimiento.Delete {Id}", id); return StatusCode(500, ex.Message); }
    }

    // ── POST /api/mantenimiento/{id}/seguimiento ──────────────────────────────
    [HttpPost("{id:guid}/seguimiento")]
    public async Task<IActionResult> AddSeguimiento(Guid id, [FromBody] SeguimientoDto dto, CancellationToken ct)
    {
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = """
                INSERT INTO dbo.MantenimientoSeguimientos (MantenimientoId, Nota, Tipo, CreadoPor)
                VALUES (@MntId, @Nota, @Tipo, @CreadoPor);
                UPDATE dbo.Mantenimientos SET ActualizadoEn = GETUTCDATE() WHERE Id = @MntId;
                """;
            cmd.Parameters.AddWithValue("@MntId",    id);
            cmd.Parameters.AddWithValue("@Nota",     dto.Nota.Trim());
            cmd.Parameters.AddWithValue("@Tipo",     dto.Tipo ?? "Seguimiento");
            cmd.Parameters.AddWithValue("@CreadoPor", CurrentUser);
            await cmd.ExecuteNonQueryAsync(ct);
            return StatusCode(201);
        }
        catch (Exception ex) { logger.LogError(ex, "Mantenimiento.AddSeguimiento {Id}", id); return StatusCode(500, ex.Message); }
    }

    // ── POST /api/mantenimiento/{id}/evidencia ────────────────────────────────
    [HttpPost("{id:guid}/evidencia")]
    [RequestSizeLimit(20 * 1024 * 1024)] // 20 MB
    public async Task<IActionResult> UploadEvidencia(Guid id, IFormFile file, CancellationToken ct)
    {
        if (file is null || file.Length == 0) return BadRequest("Archivo requerido.");
        try
        {
            await using var ms = new MemoryStream();
            await file.CopyToAsync(ms, ct);

            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = """
                INSERT INTO dbo.MantenimientoEvidencias (MantenimientoId, NombreArchivo, Mime, Datos, SubidoPor)
                VALUES (@MntId, @Nombre, @Mime, @Datos, @SubidoPor)
                """;
            cmd.Parameters.AddWithValue("@MntId",    id);
            cmd.Parameters.AddWithValue("@Nombre",   file.FileName);
            cmd.Parameters.AddWithValue("@Mime",     file.ContentType);
            cmd.Parameters.AddWithValue("@Datos",    ms.ToArray());
            cmd.Parameters.AddWithValue("@SubidoPor", CurrentUser);
            await cmd.ExecuteNonQueryAsync(ct);
            return StatusCode(201);
        }
        catch (Exception ex) { logger.LogError(ex, "Mantenimiento.UploadEvidencia {Id}", id); return StatusCode(500, ex.Message); }
    }

    // ── GET /api/mantenimiento/{id}/evidencia/{evId} ──────────────────────────
    [HttpGet("{id:guid}/evidencia/{evId:guid}")]
    public async Task<IActionResult> GetEvidencia(Guid id, Guid evId, CancellationToken ct)
    {
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = "SELECT NombreArchivo, Mime, Datos FROM dbo.MantenimientoEvidencias WHERE Id=@EvId AND MantenimientoId=@MntId";
            cmd.Parameters.AddWithValue("@EvId",  evId);
            cmd.Parameters.AddWithValue("@MntId", id);
            await using var r = await cmd.ExecuteReaderAsync(ct);
            if (!await r.ReadAsync(ct)) return NotFound();
            var nombre = r.GetString(0);
            var mime   = r.GetString(1);
            var datos  = (byte[])r.GetValue(2);
            return File(datos, mime, nombre);
        }
        catch (Exception ex) { logger.LogError(ex, "Mantenimiento.GetEvidencia"); return StatusCode(500, ex.Message); }
    }

    // ── DELETE /api/mantenimiento/{id}/evidencia/{evId} ───────────────────────
    [HttpDelete("{id:guid}/evidencia/{evId:guid}")]
    public async Task<IActionResult> DeleteEvidencia(Guid id, Guid evId, CancellationToken ct)
    {
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = "DELETE FROM dbo.MantenimientoEvidencias WHERE Id=@EvId AND MantenimientoId=@MntId";
            cmd.Parameters.AddWithValue("@EvId",  evId);
            cmd.Parameters.AddWithValue("@MntId", id);
            await cmd.ExecuteNonQueryAsync(ct);
            return NoContent();
        }
        catch (Exception ex) { logger.LogError(ex, "Mantenimiento.DeleteEvidencia"); return StatusCode(500, ex.Message); }
    }

    // ── GET /api/mantenimiento/stats ──────────────────────────────────────────
    [HttpGet("stats")]
    public async Task<IActionResult> GetStats(CancellationToken ct)
    {
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = """
                SELECT
                    COUNT(*)                                                                           AS Total,
                    ISNULL(SUM(CASE WHEN Estado='Programado'  THEN 1 ELSE 0 END), 0)                  AS Programados,
                    ISNULL(SUM(CASE WHEN Estado='En Proceso'  THEN 1 ELSE 0 END), 0)                  AS EnProceso,
                    ISNULL(SUM(CASE WHEN Estado='Completado'  THEN 1 ELSE 0 END), 0)                  AS Completados,
                    ISNULL(SUM(CASE WHEN Estado='Cancelado'   THEN 1 ELSE 0 END), 0)                  AS Cancelados,
                    ISNULL(SUM(CASE WHEN Prioridad='Alta'     THEN 1 ELSE 0 END), 0)                  AS PrioAlta,
                    ISNULL(SUM(CASE WHEN FechaProgramada < GETUTCDATE() AND Estado NOT IN ('Completado','Cancelado') THEN 1 ELSE 0 END), 0) AS Vencidos,
                    ISNULL(SUM(CASE WHEN FechaProgramada BETWEEN GETUTCDATE() AND DATEADD(DAY,7,GETUTCDATE()) AND Estado='Programado' THEN 1 ELSE 0 END), 0) AS Proximos7Dias,
                    ISNULL(SUM(CASE WHEN MONTH(FechaProgramada)=MONTH(GETUTCDATE()) AND YEAR(FechaProgramada)=YEAR(GETUTCDATE()) THEN 1 ELSE 0 END), 0) AS EsteMes
                FROM dbo.Mantenimientos WHERE IsDeleted=0
                """;
            await using var r = await cmd.ExecuteReaderAsync(ct);
            object summary = new {};
            if (await r.ReadAsync(ct))
                summary = new {
                    total        = r.GetInt32(0),
                    programados  = r.GetInt32(1),
                    enProceso    = r.GetInt32(2),
                    completados  = r.GetInt32(3),
                    cancelados   = r.GetInt32(4),
                    prioAlta     = r.GetInt32(5),
                    vencidos     = r.GetInt32(6),
                    proximos7Dias= r.GetInt32(7),
                    esteMes      = r.GetInt32(8),
                };
            return Ok(new { summary });
        }
        catch (Exception ex) { logger.LogError(ex, "Mantenimiento.GetStats"); return StatusCode(500, ex.Message); }
    }

    // ── GET /api/mantenimiento/export ─────────────────────────────────────────
    [HttpGet("export")]
    public async Task<IActionResult> Export(CancellationToken ct)
    {
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = """
                SELECT Folio, Titulo, TipoMantenimiento, Estado, Prioridad,
                       NombreEquipo, Ubicacion, TecnicoAsignado,
                       FechaProgramada, FechaRealizada, DuracionMinutos,
                       CostoEstimado, CostoReal, CreadoPor, CreadoEn
                FROM dbo.Mantenimientos WHERE IsDeleted=0 ORDER BY FechaProgramada DESC
                """;
            var sb = new StringBuilder();
            sb.AppendLine("Folio,Titulo,Tipo,Estado,Prioridad,Equipo,Ubicacion,Tecnico,FechaProgramada,FechaRealizada,DuracionMin,CostoEst,CostoReal,CreadoPor,CreadoEn");
            await using var r = await cmd.ExecuteReaderAsync(ct);
            while (await r.ReadAsync(ct))
            {
                static string C(object? v) => v is null || v == DBNull.Value ? "" : $"\"{v}\"";
                sb.AppendLine(string.Join(",", new[] {
                    C(r[0]), C(r[1]), C(r[2]), C(r[3]), C(r[4]),
                    C(r[5]), C(r[6]), C(r[7]), C(r[8]), C(r[9]),
                    C(r[10]), C(r[11]), C(r[12]), C(r[13]), C(r[14])
                }));
            }
            var bytes = Encoding.UTF8.GetPreamble().Concat(Encoding.UTF8.GetBytes(sb.ToString())).ToArray();
            return File(bytes, "text/csv", $"mantenimiento_{DateTime.Now:yyyyMMdd}.csv");
        }
        catch (Exception ex) { logger.LogError(ex, "Mantenimiento.Export"); return StatusCode(500, ex.Message); }
    }

    // ── Helpers ───────────────────────────────────────────────────────────────
    private static object MapRow(SqlDataReader r) => new {
        id               = r.GetGuid(r.GetOrdinal("Id")),
        folio            = r.GetString(r.GetOrdinal("Folio")),
        titulo           = r.GetString(r.GetOrdinal("Titulo")),
        tipoMantenimiento= r.GetString(r.GetOrdinal("TipoMantenimiento")),
        estado           = r.GetString(r.GetOrdinal("Estado")),
        prioridad        = r.GetString(r.GetOrdinal("Prioridad")),
        nombreEquipo     = r.IsDBNull(r.GetOrdinal("NombreEquipo"))     ? null : r.GetString(r.GetOrdinal("NombreEquipo")),
        ubicacion        = r.IsDBNull(r.GetOrdinal("Ubicacion"))        ? null : r.GetString(r.GetOrdinal("Ubicacion")),
        tecnicoAsignado  = r.IsDBNull(r.GetOrdinal("TecnicoAsignado"))  ? null : r.GetString(r.GetOrdinal("TecnicoAsignado")),
        fechaProgramada  = r.GetDateTime(r.GetOrdinal("FechaProgramada")),
        fechaRealizada   = r.IsDBNull(r.GetOrdinal("FechaRealizada"))   ? (DateTime?)null : r.GetDateTime(r.GetOrdinal("FechaRealizada")),
        duracionMinutos  = r.IsDBNull(r.GetOrdinal("DuracionMinutos"))  ? (int?)null : r.GetInt32(r.GetOrdinal("DuracionMinutos")),
        costoEstimado    = r.IsDBNull(r.GetOrdinal("CostoEstimado"))    ? (decimal?)null : r.GetDecimal(r.GetOrdinal("CostoEstimado")),
        costoReal        = r.IsDBNull(r.GetOrdinal("CostoReal"))        ? (decimal?)null : r.GetDecimal(r.GetOrdinal("CostoReal")),
        creadoEn         = r.GetDateTime(r.GetOrdinal("CreadoEn")),
    };

    private static object MapRowFull(SqlDataReader r) => new {
        id               = r.GetGuid(r.GetOrdinal("Id")),
        folio            = r.GetString(r.GetOrdinal("Folio")),
        titulo           = r.GetString(r.GetOrdinal("Titulo")),
        descripcion      = r.IsDBNull(r.GetOrdinal("Descripcion"))      ? null : r.GetString(r.GetOrdinal("Descripcion")),
        tipoMantenimiento= r.GetString(r.GetOrdinal("TipoMantenimiento")),
        estado           = r.GetString(r.GetOrdinal("Estado")),
        prioridad        = r.GetString(r.GetOrdinal("Prioridad")),
        equipoId         = r.IsDBNull(r.GetOrdinal("EquipoId"))         ? (Guid?)null : r.GetGuid(r.GetOrdinal("EquipoId")),
        nombreEquipo     = r.IsDBNull(r.GetOrdinal("NombreEquipo"))     ? null : r.GetString(r.GetOrdinal("NombreEquipo")),
        ubicacion        = r.IsDBNull(r.GetOrdinal("Ubicacion"))        ? null : r.GetString(r.GetOrdinal("Ubicacion")),
        tecnicoAsignado  = r.IsDBNull(r.GetOrdinal("TecnicoAsignado"))  ? null : r.GetString(r.GetOrdinal("TecnicoAsignado")),
        emailTecnico     = r.IsDBNull(r.GetOrdinal("EmailTecnico"))     ? null : r.GetString(r.GetOrdinal("EmailTecnico")),
        fechaProgramada  = r.GetDateTime(r.GetOrdinal("FechaProgramada")),
        fechaRealizada   = r.IsDBNull(r.GetOrdinal("FechaRealizada"))   ? (DateTime?)null : r.GetDateTime(r.GetOrdinal("FechaRealizada")),
        duracionMinutos  = r.IsDBNull(r.GetOrdinal("DuracionMinutos"))  ? (int?)null : r.GetInt32(r.GetOrdinal("DuracionMinutos")),
        checklistJson    = r.IsDBNull(r.GetOrdinal("ChecklistJson"))    ? null : r.GetString(r.GetOrdinal("ChecklistJson")),
        notas            = r.IsDBNull(r.GetOrdinal("Notas"))            ? null : r.GetString(r.GetOrdinal("Notas")),
        costoEstimado    = r.IsDBNull(r.GetOrdinal("CostoEstimado"))    ? (decimal?)null : r.GetDecimal(r.GetOrdinal("CostoEstimado")),
        costoReal        = r.IsDBNull(r.GetOrdinal("CostoReal"))        ? (decimal?)null : r.GetDecimal(r.GetOrdinal("CostoReal")),
        creadoPor        = r.GetString(r.GetOrdinal("CreadoPor")),
        creadoEn         = r.GetDateTime(r.GetOrdinal("CreadoEn")),
        actualizadoEn    = r.IsDBNull(r.GetOrdinal("ActualizadoEn"))    ? (DateTime?)null : r.GetDateTime(r.GetOrdinal("ActualizadoEn")),
    };

    private async Task SendTecnicoEmailAsync(MantenimientoDto dto, string folio, bool isNew)
    {
        try
        {
            var smtp    = config.GetSection("SmtpSettings");
            var host    = smtp["Host"] ?? "";
            var port    = int.TryParse(smtp["Port"], out var p) ? p : 587;
            var from    = smtp["FromEmail"] ?? "";
            var pass    = smtp["Password"] ?? "";
            var fromName= smtp["FromName"] ?? "Pandora";
            if (string.IsNullOrWhiteSpace(host) || string.IsNullOrWhiteSpace(from)) return;

            using var client = new MailKit.Net.Smtp.SmtpClient();
            await client.ConnectAsync(host, port, MailKit.Security.SecureSocketOptions.StartTls);
            await client.AuthenticateAsync(from, pass);

            var msg = new MimeKit.MimeMessage();
            msg.From.Add(new MimeKit.MailboxAddress(fromName, from));
            msg.To.Add(new MimeKit.MailboxAddress(dto.TecnicoAsignado ?? "Técnico", dto.EmailTecnico!));
            msg.Subject = isNew
                ? $"[iMET] Nuevo mantenimiento asignado: {folio}"
                : $"[iMET] Mantenimiento actualizado: {folio}";

            var fecha = dto.FechaProgramada.ToString("dd/MM/yyyy HH:mm");
            msg.Body = new MimeKit.TextPart("html") { Text = $"""
                <html><body style="font-family:Arial,sans-serif;font-size:14px">
                <div style="max-width:600px;margin:0 auto">
                  <div style="background:#1565c0;padding:20px;border-radius:8px 8px 0 0">
                    <h2 style="color:white;margin:0">🔧 {(isNew ? "Mantenimiento asignado" : "Mantenimiento actualizado")}</h2>
                  </div>
                  <div style="border:1px solid #ddd;padding:24px;border-radius:0 0 8px 8px">
                    <p>Hola <strong>{dto.TecnicoAsignado}</strong>,</p>
                    <p>Se te ha {(isNew ? "asignado" : "actualizado")} el siguiente mantenimiento:</p>
                    <table style="width:100%;border-collapse:collapse">
                      <tr><td style="padding:8px;font-weight:bold;color:#555;width:40%">Folio</td><td style="padding:8px"><strong>{folio}</strong></td></tr>
                      <tr><td style="padding:8px;font-weight:bold;color:#555">Título</td><td style="padding:8px">{dto.Titulo}</td></tr>
                      <tr><td style="padding:8px;font-weight:bold;color:#555">Tipo</td><td style="padding:8px">{dto.TipoMantenimiento}</td></tr>
                      <tr><td style="padding:8px;font-weight:bold;color:#555">Equipo</td><td style="padding:8px">{dto.NombreEquipo ?? "—"}</td></tr>
                      <tr><td style="padding:8px;font-weight:bold;color:#555">Ubicación</td><td style="padding:8px">{dto.Ubicacion ?? "—"}</td></tr>
                      <tr><td style="padding:8px;font-weight:bold;color:#555">Fecha programada</td><td style="padding:8px">{fecha}</td></tr>
                      <tr><td style="padding:8px;font-weight:bold;color:#555">Prioridad</td><td style="padding:8px">{dto.Prioridad}</td></tr>
                    </table>
                    <p style="margin-top:20px;font-size:12px;color:#888">Pandora — Sistema de Gestión iMET</p>
                  </div>
                </div>
                </body></html>
                """ };
            await client.SendAsync(msg);
            await client.DisconnectAsync(true);
        }
        catch (Exception ex) { logger.LogWarning(ex, "Mantenimiento email failed for {Folio}", folio); }
    }
}

// ── DTOs ──────────────────────────────────────────────────────────────────────
public record MantenimientoDto(
    string    Titulo,
    string?   Descripcion,
    string?   TipoMantenimiento,
    string?   Estado,
    string?   Prioridad,
    Guid?     EquipoId,
    string?   NombreEquipo,
    string?   Ubicacion,
    string?   TecnicoAsignado,
    string?   EmailTecnico,
    DateTime  FechaProgramada,
    DateTime? FechaRealizada,
    int?      DuracionMinutos,
    string?   ChecklistJson,
    string?   Notas,
    decimal?  CostoEstimado,
    decimal?  CostoReal
);

public record SeguimientoMntDto(string Nota, string? Tipo);
