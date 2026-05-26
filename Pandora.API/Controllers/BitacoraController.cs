using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using System.Security.Claims;

namespace Pandora.API.Controllers;

[ApiController]
[Route("api/bitacora")]
[Authorize]
public class BitacoraController(
    IConfiguration config,
    ILogger<BitacoraController> logger) : ControllerBase
{
    private SqlConnection Conn() => new(config.GetConnectionString("PandoraDb"));
    private bool IsAdmin => User.IsInRole("Admin");
    private string CurrentUser => User.FindFirstValue(ClaimTypes.Name) ?? "Desconocido";

    // ── GET /api/bitacora ─────────────────────────────────────────────────────
    [HttpGet]
    public async Task<IActionResult> GetAll(
        [FromQuery] string?   search     = null,
        [FromQuery] string?   categoria  = null,
        [FromQuery] string?   estado     = null,
        [FromQuery] string?   prioridad  = null,
        [FromQuery] string?   desde      = null,
        [FromQuery] string?   hasta      = null,
        CancellationToken ct = default)
    {
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);

            var where = new List<string> { "IsDeleted = 0" };
            if (!string.IsNullOrWhiteSpace(search))
                where.Add("(Titulo LIKE @Q OR Descripcion LIKE @Q OR SistemaAfectado LIKE @Q OR UsuarioAfectado LIKE @Q OR Folio LIKE @Q)");
            if (!string.IsNullOrWhiteSpace(categoria))  where.Add("Categoria  = @Categoria");
            if (!string.IsNullOrWhiteSpace(estado))     where.Add("Estado     = @Estado");
            if (!string.IsNullOrWhiteSpace(prioridad))  where.Add("Prioridad  = @Prioridad");
            if (!string.IsNullOrWhiteSpace(desde))      where.Add("FechaIncidencia >= @Desde");
            if (!string.IsNullOrWhiteSpace(hasta))      where.Add("FechaIncidencia <= @Hasta");

            await using var cmd = conn.CreateCommand();
            cmd.CommandText = $"""
                SELECT Id, Folio, Titulo, Categoria, Prioridad, Estado, Impacto,
                       SistemaAfectado, UsuarioAfectado, ReportadoPor, AsignadoA,
                       FechaIncidencia, FechaResolucion, TiempoResolucion, CreadoEn
                FROM dbo.Bitacora
                WHERE {string.Join(" AND ", where)}
                ORDER BY CreadoEn DESC
                """;

            if (!string.IsNullOrWhiteSpace(search))    cmd.Parameters.AddWithValue("@Q",         $"%{search}%");
            if (!string.IsNullOrWhiteSpace(categoria))  cmd.Parameters.AddWithValue("@Categoria",  categoria);
            if (!string.IsNullOrWhiteSpace(estado))     cmd.Parameters.AddWithValue("@Estado",     estado);
            if (!string.IsNullOrWhiteSpace(prioridad))  cmd.Parameters.AddWithValue("@Prioridad",  prioridad);
            if (!string.IsNullOrWhiteSpace(desde))      cmd.Parameters.AddWithValue("@Desde",      desde);
            if (!string.IsNullOrWhiteSpace(hasta))      cmd.Parameters.AddWithValue("@Hasta",      hasta);

            var list = new List<object>();
            await using var r = await cmd.ExecuteReaderAsync(ct);
            while (await r.ReadAsync(ct)) list.Add(MapRow(r));
            return Ok(list);
        }
        catch (Exception ex) { logger.LogError(ex, "Bitacora.GetAll"); return StatusCode(500, ex.Message); }
    }

    // ── GET /api/bitacora/{id} ────────────────────────────────────────────────
    [HttpGet("{id:guid}")]
    public async Task<IActionResult> GetById(Guid id, CancellationToken ct)
    {
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);

            // Entrada principal
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = """
                SELECT Id, Folio, Titulo, Descripcion, Categoria, Prioridad, Estado, Impacto,
                       SistemaAfectado, UsuarioAfectado, ReportadoPor, AsignadoA,
                       FechaIncidencia, FechaResolucion, TiempoResolucion, Resolucion, CreadoEn
                FROM dbo.Bitacora WHERE Id = @Id AND IsDeleted = 0
                """;
            cmd.Parameters.AddWithValue("@Id", id);
            await using var r = await cmd.ExecuteReaderAsync(ct);
            if (!await r.ReadAsync(ct)) return NotFound("Entrada no encontrada.");
            var entry = MapRowDetail(r);
            await r.CloseAsync();

            // Seguimientos
            await using var sc = conn.CreateCommand();
            sc.CommandText = """
                SELECT Id, Nota, Tipo, CreadoPor, CreadoEn
                FROM dbo.BitacoraSeguimientos
                WHERE BitacoraId = @Id
                ORDER BY CreadoEn ASC
                """;
            sc.Parameters.AddWithValue("@Id", id);
            await using var sr = await sc.ExecuteReaderAsync(ct);
            var seguimientos = new List<object>();
            while (await sr.ReadAsync(ct))
                seguimientos.Add(new {
                    id        = sr.GetGuid(0),
                    nota      = sr.GetString(1),
                    tipo      = sr.GetString(2),
                    creadoPor = sr.GetString(3),
                    creadoEn  = sr.GetDateTime(4),
                });

            return Ok(new { entry, seguimientos });
        }
        catch (Exception ex) { logger.LogError(ex, "Bitacora.GetById {Id}", id); return StatusCode(500, ex.Message); }
    }

    // ── POST /api/bitacora ────────────────────────────────────────────────────
    [HttpPost]
    public async Task<IActionResult> Create([FromBody] BitacoraDto dto, CancellationToken ct)
    {
        if (string.IsNullOrWhiteSpace(dto.Titulo))       return BadRequest("Título requerido.");
        if (string.IsNullOrWhiteSpace(dto.Descripcion))  return BadRequest("Descripción requerida.");

        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);

            // Generar folio correlativo: BIT-YYYY-NNNN
            string folio;
            await using (var fc = conn.CreateCommand())
            {
                fc.CommandText = """
                    SELECT ISNULL(MAX(CAST(SUBSTRING(Folio, 10, 4) AS INT)), 0) + 1
                    FROM dbo.Bitacora
                    WHERE Folio LIKE @Prefix
                    """;
                fc.Parameters.AddWithValue("@Prefix", $"BIT-{DateTime.UtcNow.Year}-%");
                var seq = (int)(await fc.ExecuteScalarAsync(ct) ?? 1);
                folio = $"BIT-{DateTime.UtcNow.Year}-{seq:D4}";
            }

            var newId = Guid.NewGuid();
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = """
                INSERT INTO dbo.Bitacora
                    (Id, Folio, Titulo, Descripcion, Categoria, Prioridad, Estado, Impacto,
                     SistemaAfectado, UsuarioAfectado, ReportadoPor, AsignadoA,
                     FechaIncidencia, FechaResolucion, Resolucion, CreadoEn)
                VALUES
                    (@Id, @Folio, @Titulo, @Desc, @Cat, @Prio, @Estado, @Impacto,
                     @Sistema, @Usuario, @ReportadoPor, @Asignado,
                     @FechaInc, @FechaRes, @Resolucion, GETUTCDATE())
                """;
            cmd.Parameters.AddWithValue("@Id",           newId);
            cmd.Parameters.AddWithValue("@Folio",        folio);
            cmd.Parameters.AddWithValue("@Titulo",       dto.Titulo.Trim());
            cmd.Parameters.AddWithValue("@Desc",         dto.Descripcion.Trim());
            cmd.Parameters.AddWithValue("@Cat",          dto.Categoria  ?? "Otro");
            cmd.Parameters.AddWithValue("@Prio",         dto.Prioridad  ?? "Media");
            cmd.Parameters.AddWithValue("@Estado",       dto.Estado     ?? "Abierta");
            cmd.Parameters.AddWithValue("@Impacto",      Nv(dto.Impacto));
            cmd.Parameters.AddWithValue("@Sistema",      Nv(dto.SistemaAfectado));
            cmd.Parameters.AddWithValue("@Usuario",      Nv(dto.UsuarioAfectado));
            cmd.Parameters.AddWithValue("@ReportadoPor", dto.ReportadoPor ?? CurrentUser);
            cmd.Parameters.AddWithValue("@Asignado",     Nv(dto.AsignadoA));
            cmd.Parameters.AddWithValue("@FechaInc",     dto.FechaIncidencia ?? DateTime.UtcNow);
            cmd.Parameters.AddWithValue("@FechaRes",     dto.FechaResolucion.HasValue ? (object)dto.FechaResolucion.Value : DBNull.Value);
            cmd.Parameters.AddWithValue("@Resolucion",   Nv(dto.Resolucion));
            await cmd.ExecuteNonQueryAsync(ct);

            return CreatedAtAction(nameof(GetById), new { id = newId }, new { id = newId, folio });
        }
        catch (Exception ex) { logger.LogError(ex, "Bitacora.Create"); return StatusCode(500, ex.Message); }
    }

    // ── PUT /api/bitacora/{id} ────────────────────────────────────────────────
    [HttpPut("{id:guid}")]
    public async Task<IActionResult> Update(Guid id, [FromBody] BitacoraDto dto, CancellationToken ct)
    {
        if (string.IsNullOrWhiteSpace(dto.Titulo)) return BadRequest("Título requerido.");
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = """
                UPDATE dbo.Bitacora SET
                    Titulo          = @Titulo,
                    Descripcion     = @Desc,
                    Categoria       = @Cat,
                    Prioridad       = @Prio,
                    Estado          = @Estado,
                    Impacto         = @Impacto,
                    SistemaAfectado = @Sistema,
                    UsuarioAfectado = @Usuario,
                    AsignadoA       = @Asignado,
                    FechaIncidencia = @FechaInc,
                    FechaResolucion = @FechaRes,
                    Resolucion      = @Resolucion,
                    ActualizadoEn   = GETUTCDATE()
                WHERE Id = @Id AND IsDeleted = 0
                """;
            cmd.Parameters.AddWithValue("@Id",        id);
            cmd.Parameters.AddWithValue("@Titulo",    dto.Titulo.Trim());
            cmd.Parameters.AddWithValue("@Desc",      dto.Descripcion?.Trim() ?? "");
            cmd.Parameters.AddWithValue("@Cat",       dto.Categoria  ?? "Otro");
            cmd.Parameters.AddWithValue("@Prio",      dto.Prioridad  ?? "Media");
            cmd.Parameters.AddWithValue("@Estado",    dto.Estado     ?? "Abierta");
            cmd.Parameters.AddWithValue("@Impacto",   Nv(dto.Impacto));
            cmd.Parameters.AddWithValue("@Sistema",   Nv(dto.SistemaAfectado));
            cmd.Parameters.AddWithValue("@Usuario",   Nv(dto.UsuarioAfectado));
            cmd.Parameters.AddWithValue("@Asignado",  Nv(dto.AsignadoA));
            cmd.Parameters.AddWithValue("@FechaInc",  dto.FechaIncidencia ?? DateTime.UtcNow);
            cmd.Parameters.AddWithValue("@FechaRes",  dto.FechaResolucion.HasValue ? (object)dto.FechaResolucion.Value : DBNull.Value);
            cmd.Parameters.AddWithValue("@Resolucion",Nv(dto.Resolucion));
            int rows = await cmd.ExecuteNonQueryAsync(ct);
            if (rows == 0) return NotFound();
            return NoContent();
        }
        catch (Exception ex) { logger.LogError(ex, "Bitacora.Update {Id}", id); return StatusCode(500, ex.Message); }
    }

    // ── DELETE /api/bitacora/{id}  (Admin) ───────────────────────────────────
    [HttpDelete("{id:guid}")]
    [Authorize(Roles = "Admin")]
    public async Task<IActionResult> Delete(Guid id, CancellationToken ct)
    {
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = "UPDATE dbo.Bitacora SET IsDeleted=1, ActualizadoEn=GETUTCDATE() WHERE Id=@Id AND IsDeleted=0";
            cmd.Parameters.AddWithValue("@Id", id);
            int rows = await cmd.ExecuteNonQueryAsync(ct);
            if (rows == 0) return NotFound();
            return NoContent();
        }
        catch (Exception ex) { logger.LogError(ex, "Bitacora.Delete {Id}", id); return StatusCode(500, ex.Message); }
    }

    // ── POST /api/bitacora/{id}/seguimiento ───────────────────────────────────
    [HttpPost("{id:guid}/seguimiento")]
    public async Task<IActionResult> AddSeguimiento(Guid id, [FromBody] SeguimientoDto dto, CancellationToken ct)
    {
        if (string.IsNullOrWhiteSpace(dto.Nota)) return BadRequest("La nota no puede estar vacía.");
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);

            var newId = Guid.NewGuid();
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = """
                INSERT INTO dbo.BitacoraSeguimientos (Id, BitacoraId, Nota, Tipo, CreadoPor, CreadoEn)
                VALUES (@Id, @BitId, @Nota, @Tipo, @By, GETUTCDATE());
                UPDATE dbo.Bitacora SET ActualizadoEn=GETUTCDATE() WHERE Id=@BitId;
                """;
            cmd.Parameters.AddWithValue("@Id",    newId);
            cmd.Parameters.AddWithValue("@BitId", id);
            cmd.Parameters.AddWithValue("@Nota",  dto.Nota.Trim());
            cmd.Parameters.AddWithValue("@Tipo",  dto.Tipo ?? "Seguimiento");
            cmd.Parameters.AddWithValue("@By",    CurrentUser);
            await cmd.ExecuteNonQueryAsync(ct);

            return Ok(new { id = newId });
        }
        catch (Exception ex) { logger.LogError(ex, "Bitacora.AddSeguimiento {Id}", id); return StatusCode(500, ex.Message); }
    }

    // ── GET /api/bitacora/stats ───────────────────────────────────────────────
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
                    COUNT(*)                                                  AS Total,
                    SUM(CASE WHEN Estado='Abierta'    THEN 1 ELSE 0 END)     AS Abiertas,
                    SUM(CASE WHEN Estado='En Proceso' THEN 1 ELSE 0 END)     AS EnProceso,
                    SUM(CASE WHEN Estado='Resuelta'   THEN 1 ELSE 0 END)     AS Resueltas,
                    SUM(CASE WHEN Estado='Cerrada'    THEN 1 ELSE 0 END)     AS Cerradas,
                    SUM(CASE WHEN Prioridad='Alta'    THEN 1 ELSE 0 END)     AS PrioAlta,
                    ISNULL(AVG(CAST(TiempoResolucion AS FLOAT)), 0)          AS TiempoPromedioMin,
                    SUM(CASE WHEN MONTH(FechaIncidencia)=MONTH(GETUTCDATE())
                              AND YEAR(FechaIncidencia)=YEAR(GETUTCDATE())
                             THEN 1 ELSE 0 END)                              AS EsteMes
                FROM dbo.Bitacora WHERE IsDeleted=0;

                SELECT TOP 5 Categoria, COUNT(*) AS Total
                FROM dbo.Bitacora WHERE IsDeleted=0
                GROUP BY Categoria ORDER BY Total DESC;
                """;

            await using var r = await cmd.ExecuteReaderAsync(ct);
            object summary = new {};
            if (await r.ReadAsync(ct))
                summary = new {
                    total            = r.GetInt32(0),
                    abiertas         = r.GetInt32(1),
                    enProceso        = r.GetInt32(2),
                    resueltas        = r.GetInt32(3),
                    cerradas         = r.GetInt32(4),
                    prioAlta         = r.GetInt32(5),
                    tiempoPromedioMin= (int)r.GetDouble(6),
                    esteMes          = r.GetInt32(7),
                };

            var porCategoria = new List<object>();
            if (await r.NextResultAsync(ct))
                while (await r.ReadAsync(ct))
                    porCategoria.Add(new { categoria = r.GetString(0), total = r.GetInt32(1) });

            return Ok(new { summary, porCategoria });
        }
        catch (Exception ex) { logger.LogError(ex, "Bitacora.GetStats"); return StatusCode(500, ex.Message); }
    }

    // ── GET /api/bitacora/export-excel ────────────────────────────────────────
    [HttpGet("export-excel")]
    public async Task<IActionResult> ExportExcel(
        [FromQuery] string? estado    = null,
        [FromQuery] string? categoria = null,
        [FromQuery] string? desde     = null,
        [FromQuery] string? hasta     = null,
        CancellationToken ct = default)
    {
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);

            var where = new List<string> { "IsDeleted=0" };
            if (!string.IsNullOrWhiteSpace(estado))    where.Add("Estado=@Estado");
            if (!string.IsNullOrWhiteSpace(categoria)) where.Add("Categoria=@Categoria");
            if (!string.IsNullOrWhiteSpace(desde))     where.Add("FechaIncidencia>=@Desde");
            if (!string.IsNullOrWhiteSpace(hasta))     where.Add("FechaIncidencia<=@Hasta");

            await using var cmd = conn.CreateCommand();
            cmd.CommandText = $"""
                SELECT Folio, Titulo, Categoria, Prioridad, Estado, Impacto,
                       SistemaAfectado, UsuarioAfectado, ReportadoPor, AsignadoA,
                       FechaIncidencia, FechaResolucion, TiempoResolucion, Resolucion, CreadoEn
                FROM dbo.Bitacora WHERE {string.Join(" AND ", where)}
                ORDER BY FechaIncidencia DESC
                """;
            if (!string.IsNullOrWhiteSpace(estado))    cmd.Parameters.AddWithValue("@Estado",    estado);
            if (!string.IsNullOrWhiteSpace(categoria)) cmd.Parameters.AddWithValue("@Categoria", categoria);
            if (!string.IsNullOrWhiteSpace(desde))     cmd.Parameters.AddWithValue("@Desde",     desde);
            if (!string.IsNullOrWhiteSpace(hasta))     cmd.Parameters.AddWithValue("@Hasta",     hasta);

            // Construir CSV simple (alternativa sin dependencia de ClosedXML)
            var sb = new System.Text.StringBuilder();
            sb.AppendLine("Folio,Título,Categoría,Prioridad,Estado,Impacto,Sistema Afectado,Usuario Afectado,Reportado Por,Asignado A,Fecha Incidencia,Fecha Resolución,Tiempo Resolución (min),Resolución,Creado En");

            await using var r = await cmd.ExecuteReaderAsync(ct);
            while (await r.ReadAsync(ct))
            {
                string Esc(string? s) => $"\"{(s ?? "").Replace("\"", "\"\"")}\"";
                sb.AppendLine(string.Join(",",
                    Esc(r.GetString(0)), Esc(r.GetString(1)), Esc(r.GetString(2)),
                    Esc(r.GetString(3)), Esc(r.GetString(4)),
                    Esc(r.IsDBNull(5)  ? null : r.GetString(5)),
                    Esc(r.IsDBNull(6)  ? null : r.GetString(6)),
                    Esc(r.IsDBNull(7)  ? null : r.GetString(7)),
                    Esc(r.GetString(8)),
                    Esc(r.IsDBNull(9)  ? null : r.GetString(9)),
                    Esc(r.GetDateTime(10).ToString("yyyy-MM-dd HH:mm")),
                    Esc(r.IsDBNull(11) ? null : r.GetDateTime(11).ToString("yyyy-MM-dd HH:mm")),
                    Esc(r.IsDBNull(12) ? null : r.GetInt32(12).ToString()),
                    Esc(r.IsDBNull(13) ? null : r.GetString(13)),
                    Esc(r.GetDateTime(14).ToString("yyyy-MM-dd HH:mm"))
                ));
            }

            var bytes = System.Text.Encoding.UTF8.GetPreamble().Concat(
                System.Text.Encoding.UTF8.GetBytes(sb.ToString())).ToArray();
            return File(bytes, "text/csv; charset=utf-8",
                $"bitacora-{DateTime.UtcNow:yyyyMMdd}.csv");
        }
        catch (Exception ex) { logger.LogError(ex, "Bitacora.ExportExcel"); return StatusCode(500, ex.Message); }
    }

    // ── Helpers ───────────────────────────────────────────────────────────────
    private static object Nv(string? s) => string.IsNullOrWhiteSpace(s) ? DBNull.Value : (object)s;

    private static object MapRow(SqlDataReader r) => new
    {
        id               = r.GetGuid(r.GetOrdinal("Id")),
        folio            = r.GetString(r.GetOrdinal("Folio")),
        titulo           = r.GetString(r.GetOrdinal("Titulo")),
        categoria        = r.GetString(r.GetOrdinal("Categoria")),
        prioridad        = r.GetString(r.GetOrdinal("Prioridad")),
        estado           = r.GetString(r.GetOrdinal("Estado")),
        impacto          = r.IsDBNull(r.GetOrdinal("Impacto"))          ? null : r.GetString(r.GetOrdinal("Impacto")),
        sistemaAfectado  = r.IsDBNull(r.GetOrdinal("SistemaAfectado"))  ? null : r.GetString(r.GetOrdinal("SistemaAfectado")),
        usuarioAfectado  = r.IsDBNull(r.GetOrdinal("UsuarioAfectado"))  ? null : r.GetString(r.GetOrdinal("UsuarioAfectado")),
        reportadoPor     = r.GetString(r.GetOrdinal("ReportadoPor")),
        asignadoA        = r.IsDBNull(r.GetOrdinal("AsignadoA"))        ? null : r.GetString(r.GetOrdinal("AsignadoA")),
        fechaIncidencia  = r.GetDateTime(r.GetOrdinal("FechaIncidencia")),
        fechaResolucion  = r.IsDBNull(r.GetOrdinal("FechaResolucion"))  ? (DateTime?)null : r.GetDateTime(r.GetOrdinal("FechaResolucion")),
        tiempoResolucion = r.IsDBNull(r.GetOrdinal("TiempoResolucion")) ? (int?)null : r.GetInt32(r.GetOrdinal("TiempoResolucion")),
        creadoEn         = r.GetDateTime(r.GetOrdinal("CreadoEn")),
    };

    private static object MapRowDetail(SqlDataReader r) => new
    {
        id               = r.GetGuid(r.GetOrdinal("Id")),
        folio            = r.GetString(r.GetOrdinal("Folio")),
        titulo           = r.GetString(r.GetOrdinal("Titulo")),
        descripcion      = r.GetString(r.GetOrdinal("Descripcion")),
        categoria        = r.GetString(r.GetOrdinal("Categoria")),
        prioridad        = r.GetString(r.GetOrdinal("Prioridad")),
        estado           = r.GetString(r.GetOrdinal("Estado")),
        impacto          = r.IsDBNull(r.GetOrdinal("Impacto"))          ? null : r.GetString(r.GetOrdinal("Impacto")),
        sistemaAfectado  = r.IsDBNull(r.GetOrdinal("SistemaAfectado"))  ? null : r.GetString(r.GetOrdinal("SistemaAfectado")),
        usuarioAfectado  = r.IsDBNull(r.GetOrdinal("UsuarioAfectado"))  ? null : r.GetString(r.GetOrdinal("UsuarioAfectado")),
        reportadoPor     = r.GetString(r.GetOrdinal("ReportadoPor")),
        asignadoA        = r.IsDBNull(r.GetOrdinal("AsignadoA"))        ? null : r.GetString(r.GetOrdinal("AsignadoA")),
        fechaIncidencia  = r.GetDateTime(r.GetOrdinal("FechaIncidencia")),
        fechaResolucion  = r.IsDBNull(r.GetOrdinal("FechaResolucion"))  ? (DateTime?)null : r.GetDateTime(r.GetOrdinal("FechaResolucion")),
        tiempoResolucion = r.IsDBNull(r.GetOrdinal("TiempoResolucion")) ? (int?)null : r.GetInt32(r.GetOrdinal("TiempoResolucion")),
        resolucion       = r.IsDBNull(r.GetOrdinal("Resolucion"))       ? null : r.GetString(r.GetOrdinal("Resolucion")),
        creadoEn         = r.GetDateTime(r.GetOrdinal("CreadoEn")),
    };
}

// ── DTOs ──────────────────────────────────────────────────────────────────────
public record BitacoraDto(
    string    Titulo,
    string?   Descripcion,
    string?   Categoria,
    string?   Prioridad,
    string?   Estado,
    string?   Impacto,
    string?   SistemaAfectado,
    string?   UsuarioAfectado,
    string?   ReportadoPor,
    string?   AsignadoA,
    DateTime? FechaIncidencia,
    DateTime? FechaResolucion,
    string?   Resolucion
);

public record SeguimientoDto(string Nota, string? Tipo);
