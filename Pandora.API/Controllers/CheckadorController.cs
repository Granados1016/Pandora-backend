using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.RateLimiting;
using Microsoft.Data.SqlClient;
using System.Security.Claims;
using System.Text;

namespace Pandora.API.Controllers;

[ApiController]
[Route("api/checador")]
[Authorize]
public class CheckadorController(
    IConfiguration config,
    ILogger<CheckadorController> logger) : ControllerBase
{
    private SqlConnection Conn() => new(config.GetConnectionString("PandoraDb"));
    private string CurrentUser   => User.FindFirstValue(ClaimTypes.Name) ?? "sistema";
    private string CurrentUserId => User.FindFirstValue(ClaimTypes.NameIdentifier)
                                 ?? User.FindFirstValue("sub") ?? CurrentUser;
    private bool   IsAdmin       => User.IsInRole("Admin");

    // Coordenadas válidas para México (bounding box generoso)
    private static bool CoordsValidas(double? lat, double? lng) =>
        lat is null || lng is null ||
        (lat >= 14.5 && lat <= 32.7 && lng >= -118.5 && lng <= -86.7);

    // ── POST /api/checador/marcar ─────────────────────────────────────────────
    /// <summary>Registra entrada o salida del usuario autenticado.</summary>
    [HttpPost("marcar")]
    [EnableRateLimiting("checador-policy")]
    public async Task<IActionResult> Marcar([FromBody] MarcarDto dto, CancellationToken ct)
    {
        try
        {
            // Validar tipo
            var tipo = dto.Tipo?.Trim();
            if (tipo != "Entrada" && tipo != "Salida")
                return BadRequest("Tipo debe ser 'Entrada' o 'Salida'.");

            // Validar coordenadas — rechazo si vienen coords fuera de México
            if (!CoordsValidas(dto.Lat, dto.Lng))
                return BadRequest("Coordenadas GPS fuera del rango válido.");

            await using var conn = Conn();
            await conn.OpenAsync(ct);

            // Verificar que no haya ya un registro del mismo tipo hoy
            await using var cmdDup = conn.CreateCommand();
            cmdDup.CommandText = """
                SELECT COUNT(1) FROM dbo.CheckadorRegistros
                WHERE UserId = @UserId AND Tipo = @Tipo
                  AND CAST(CONVERT(datetime2,TODATETIMEOFFSET(Timestamp,0) AT TIME ZONE 'Central Standard Time') AS DATE)
                  = CAST(CONVERT(datetime2,TODATETIMEOFFSET(GETUTCDATE(),0) AT TIME ZONE 'Central Standard Time') AS DATE)
                """;
            cmdDup.Parameters.AddWithValue("@UserId", CurrentUserId);
            cmdDup.Parameters.AddWithValue("@Tipo",   tipo);
            var count = Convert.ToInt32(await cmdDup.ExecuteScalarAsync(ct));
            if (count > 0)
                return Conflict($"Ya existe un registro de {tipo} para hoy.");

            // Propuesta B: calcular si está dentro de zona (si viene ubicación)
            bool? esDentroDeZona = null;
            Guid? sitioId = null;
            if (dto.Lat.HasValue && dto.Lng.HasValue)
            {
                await using var cmdSitios = conn.CreateCommand();
                cmdSitios.CommandText = "SELECT Id, Lat, Lng, RadioMetros FROM dbo.CheckadorSitios WHERE Activo = 1";
                await using var rs = await cmdSitios.ExecuteReaderAsync(ct);
                esDentroDeZona = false;
                while (await rs.ReadAsync(ct))
                {
                    var sId  = rs.GetGuid(0);
                    var sLat = rs.GetDouble(1);
                    var sLng = rs.GetDouble(2);
                    var rad  = rs.GetInt32(3);
                    var dist = HaversineMetros(dto.Lat.Value, dto.Lng.Value, sLat, sLng);
                    if (dist <= rad) { esDentroDeZona = true; sitioId = sId; break; }
                }
            }

            var id = Guid.NewGuid();
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = """
                INSERT INTO dbo.CheckadorRegistros
                    (Id, UserId, UserName, Tipo, Lat, Lng, Precision, SitioId, EsDentroDeZona, Notas)
                VALUES
                    (@Id, @UserId, @UserName, @Tipo, @Lat, @Lng, @Precision, @SitioId, @EsDentro, @Notas)
                """;
            cmd.Parameters.AddWithValue("@Id",        id);
            cmd.Parameters.AddWithValue("@UserId",    CurrentUserId);
            cmd.Parameters.AddWithValue("@UserName",  CurrentUser);
            cmd.Parameters.AddWithValue("@Tipo",      tipo);
            cmd.Parameters.AddWithValue("@Lat",       (object?)dto.Lat       ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@Lng",       (object?)dto.Lng       ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@Precision", (object?)dto.Precision ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@SitioId",   (object?)sitioId       ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@EsDentro",  (object?)esDentroDeZona ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@Notas",     (object?)dto.Notas     ?? DBNull.Value);
            await cmd.ExecuteNonQueryAsync(ct);

            return Ok(new { id, tipo, timestamp = DateTime.UtcNow, esDentroDeZona });
        }
        catch (Exception ex) { logger.LogError(ex, "Checador.Marcar"); return StatusCode(500, ex.Message); }
    }

    // ── POST /api/checador/marcar-foto ────────────────────────────────────────
    /// <summary>Propuesta B — Registra con selfie (multipart).</summary>
    [HttpPost("marcar-foto")]
    [RequestSizeLimit(10 * 1024 * 1024)]
    public async Task<IActionResult> MarcarConFoto(
        [FromForm] string tipo,
        [FromForm] double? lat,
        [FromForm] double? lng,
        [FromForm] double? precision,
        [FromForm] string? notas,
        IFormFile? foto,
        CancellationToken ct)
    {
        try
        {
            tipo = tipo?.Trim() ?? "";
            if (tipo != "Entrada" && tipo != "Salida")
                return BadRequest("Tipo debe ser 'Entrada' o 'Salida'.");

            await using var conn = Conn();
            await conn.OpenAsync(ct);

            await using var cmdDup = conn.CreateCommand();
            cmdDup.CommandText = """
                SELECT COUNT(1) FROM dbo.CheckadorRegistros
                WHERE UserId=@UserId AND Tipo=@Tipo AND CAST(CONVERT(datetime2,TODATETIMEOFFSET(Timestamp,0) AT TIME ZONE 'Central Standard Time') AS DATE)=CAST(CONVERT(datetime2,TODATETIMEOFFSET(GETUTCDATE(),0) AT TIME ZONE 'Central Standard Time') AS DATE)
                """;
            cmdDup.Parameters.AddWithValue("@UserId", CurrentUserId);
            cmdDup.Parameters.AddWithValue("@Tipo",   tipo);
            if (Convert.ToInt32(await cmdDup.ExecuteScalarAsync(ct)) > 0)
                return Conflict($"Ya existe un registro de {tipo} para hoy.");

            bool? esDentroDeZona = null;
            Guid? sitioId = null;
            if (lat.HasValue && lng.HasValue)
            {
                await using var cmdSitios = conn.CreateCommand();
                cmdSitios.CommandText = "SELECT Id, Lat, Lng, RadioMetros FROM dbo.CheckadorSitios WHERE Activo=1";
                await using var rs = await cmdSitios.ExecuteReaderAsync(ct);
                esDentroDeZona = false;
                while (await rs.ReadAsync(ct))
                {
                    var dist = HaversineMetros(lat.Value, lng.Value, rs.GetDouble(1), rs.GetDouble(2));
                    if (dist <= rs.GetInt32(3)) { esDentroDeZona = true; sitioId = rs.GetGuid(0); break; }
                }
            }

            byte[]? fotoBytes = null; string? fotoMime = null;
            if (foto is { Length: > 0 })
            {
                await using var ms = new MemoryStream();
                await foto.CopyToAsync(ms, ct);
                fotoBytes = ms.ToArray(); fotoMime = foto.ContentType;
            }

            var id = Guid.NewGuid();
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = """
                INSERT INTO dbo.CheckadorRegistros
                    (Id, UserId, UserName, Tipo, Lat, Lng, Precision, SitioId, EsDentroDeZona, FotoData, FotoMime, Notas)
                VALUES (@Id,@UserId,@UserName,@Tipo,@Lat,@Lng,@Prec,@SitioId,@EsDentro,@FotoData,@FotoMime,@Notas)
                """;
            cmd.Parameters.AddWithValue("@Id",       id);
            cmd.Parameters.AddWithValue("@UserId",   CurrentUserId);
            cmd.Parameters.AddWithValue("@UserName", CurrentUser);
            cmd.Parameters.AddWithValue("@Tipo",     tipo);
            cmd.Parameters.AddWithValue("@Lat",      (object?)lat        ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@Lng",      (object?)lng        ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@Prec",     (object?)precision  ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@SitioId",  (object?)sitioId    ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@EsDentro", (object?)esDentroDeZona ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@FotoData", (object?)fotoBytes  ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@FotoMime", (object?)fotoMime   ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@Notas",    (object?)notas      ?? DBNull.Value);
            await cmd.ExecuteNonQueryAsync(ct);

            return Ok(new { id, tipo, timestamp = DateTime.UtcNow, esDentroDeZona });
        }
        catch (Exception ex) { logger.LogError(ex, "Checador.MarcarConFoto"); return StatusCode(500, ex.Message); }
    }

    // ── GET /api/checador/hoy ─────────────────────────────────────────────────
    /// <summary>Registros del usuario actual para hoy.</summary>
    [HttpGet("hoy")]
    public async Task<IActionResult> GetHoy(CancellationToken ct)
    {
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = """
                SELECT Id, Tipo, Lat, Lng, Precision, EsDentroDeZona, Notas, Timestamp
                FROM dbo.CheckadorRegistros
                WHERE UserId = @UserId AND CAST(CONVERT(datetime2,TODATETIMEOFFSET(Timestamp,0) AT TIME ZONE 'Central Standard Time') AS DATE)
                  = CAST(CONVERT(datetime2,TODATETIMEOFFSET(GETUTCDATE(),0) AT TIME ZONE 'Central Standard Time') AS DATE)
                ORDER BY Timestamp ASC
                """;
            cmd.Parameters.AddWithValue("@UserId", CurrentUserId);
            var list = new List<object>();
            await using var r = await cmd.ExecuteReaderAsync(ct);
            while (await r.ReadAsync(ct))
                list.Add(MapRegistro(r));
            return Ok(list);
        }
        catch (Exception ex) { logger.LogError(ex, "Checador.GetHoy"); return StatusCode(500, ex.Message); }
    }

    // ── GET /api/checador/registros ───────────────────────────────────────────
    /// <summary>Admin — lista de registros con filtros.</summary>
    [HttpGet("registros")]
    [Authorize(Roles = "Admin")]
    public async Task<IActionResult> GetRegistros(
        [FromQuery] string?   userId = null,
        [FromQuery] string?   tipo   = null,
        [FromQuery] DateTime? desde  = null,
        [FromQuery] DateTime? hasta  = null,
        CancellationToken ct = default)
    {
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await using var cmd = conn.CreateCommand();

            var where = new List<string> { "1=1" };
            if (!string.IsNullOrWhiteSpace(userId)) { where.Add("(UserId = @Uid OR UserName LIKE @UidL)"); cmd.Parameters.AddWithValue("@Uid", userId); cmd.Parameters.AddWithValue("@UidL", $"%{userId}%"); }
            if (!string.IsNullOrWhiteSpace(tipo))   { where.Add("Tipo = @Tipo"); cmd.Parameters.AddWithValue("@Tipo", tipo); }
            if (desde.HasValue) { where.Add("Timestamp >= @Desde"); cmd.Parameters.AddWithValue("@Desde", desde.Value); }
            if (hasta.HasValue) { where.Add("Timestamp <= @Hasta"); cmd.Parameters.AddWithValue("@Hasta", hasta.Value.AddDays(1)); }

            cmd.CommandText = $"""
                SELECT Id, UserId, UserName, Tipo, Lat, Lng, Precision, EsDentroDeZona, Notas, Timestamp
                FROM dbo.CheckadorRegistros
                WHERE {string.Join(" AND ", where)}
                ORDER BY Timestamp DESC
                """;

            var list = new List<object>();
            await using var r = await cmd.ExecuteReaderAsync(ct);
            while (await r.ReadAsync(ct))
                list.Add(new {
                    id             = r.GetGuid(0),
                    userId         = r.GetString(1),
                    userName       = r.GetString(2),
                    tipo           = r.GetString(3),
                    lat            = r.IsDBNull(4) ? (double?)null : r.GetDouble(4),
                    lng            = r.IsDBNull(5) ? (double?)null : r.GetDouble(5),
                    precision      = r.IsDBNull(6) ? (double?)null : r.GetDouble(6),
                    esDentroDeZona = r.IsDBNull(7) ? (bool?)null   : r.GetBoolean(7),
                    notas          = r.IsDBNull(8) ? null          : r.GetString(8),
                    timestamp      = r.GetDateTime(9),
                });
            return Ok(list);
        }
        catch (Exception ex) { logger.LogError(ex, "Checador.GetRegistros"); return StatusCode(500, ex.Message); }
    }

    // ── GET /api/checador/foto/{id} ───────────────────────────────────────────
    [HttpGet("foto/{id:guid}")]
    public async Task<IActionResult> GetFoto(Guid id, CancellationToken ct)
    {
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await using var cmd = conn.CreateCommand();
            // Admin ve cualquier foto; usuario normal solo la suya
            cmd.CommandText = IsAdmin
                ? "SELECT FotoData, FotoMime FROM dbo.CheckadorRegistros WHERE Id=@Id AND FotoData IS NOT NULL"
                : "SELECT FotoData, FotoMime FROM dbo.CheckadorRegistros WHERE Id=@Id AND UserId=@UserId AND FotoData IS NOT NULL";
            cmd.Parameters.AddWithValue("@Id", id);
            if (!IsAdmin) cmd.Parameters.AddWithValue("@UserId", CurrentUserId);
            await using var r = await cmd.ExecuteReaderAsync(ct);
            if (!await r.ReadAsync(ct)) return NotFound();
            return File((byte[])r.GetValue(0), r.GetString(1));
        }
        catch (Exception ex) { logger.LogError(ex, "Checador.GetFoto {Id}", id); return StatusCode(500, ex.Message); }
    }

    // ── GET /api/checador/stats ───────────────────────────────────────────────
    [HttpGet("stats")]
    [Authorize(Roles = "Admin")]
    public async Task<IActionResult> GetStats(CancellationToken ct)
    {
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = """
                SELECT
                    ISNULL(SUM(CASE WHEN CAST(CONVERT(datetime2,TODATETIMEOFFSET(Timestamp,0) AT TIME ZONE 'Central Standard Time') AS DATE)=CAST(CONVERT(datetime2,TODATETIMEOFFSET(GETUTCDATE(),0) AT TIME ZONE 'Central Standard Time') AS DATE) AND Tipo='Entrada' THEN 1 ELSE 0 END),0) AS EntradasHoy,
                    ISNULL(SUM(CASE WHEN CAST(CONVERT(datetime2,TODATETIMEOFFSET(Timestamp,0) AT TIME ZONE 'Central Standard Time') AS DATE)=CAST(CONVERT(datetime2,TODATETIMEOFFSET(GETUTCDATE(),0) AT TIME ZONE 'Central Standard Time') AS DATE) AND Tipo='Salida'  THEN 1 ELSE 0 END),0) AS SalidasHoy,
                    ISNULL(SUM(CASE WHEN MONTH(Timestamp)=MONTH(GETUTCDATE()) AND YEAR(Timestamp)=YEAR(GETUTCDATE()) THEN 1 ELSE 0 END),0) AS EsteMes,
                    ISNULL(SUM(CASE WHEN EsDentroDeZona=0 THEN 1 ELSE 0 END),0) AS FueraDeZona,
                    COUNT(DISTINCT UserId) AS UsuariosRegistrados
                FROM dbo.CheckadorRegistros
                """;
            await using var r = await cmd.ExecuteReaderAsync(ct);
            if (!await r.ReadAsync(ct)) return Ok(new {});
            return Ok(new {
                entradasHoy        = r.GetInt32(0),
                salidasHoy         = r.GetInt32(1),
                esteMes            = r.GetInt32(2),
                fueraDeZona        = r.GetInt32(3),
                usuariosRegistrados= r.GetInt32(4),
            });
        }
        catch (Exception ex) { logger.LogError(ex, "Checador.GetStats"); return StatusCode(500, ex.Message); }
    }

    // ── GET /api/checador/export ──────────────────────────────────────────────
    [HttpGet("export")]
    [Authorize(Roles = "Admin")]
    public async Task<IActionResult> Export(
        [FromQuery] DateTime? desde = null,
        [FromQuery] DateTime? hasta = null,
        CancellationToken ct = default)
    {
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await using var cmd = conn.CreateCommand();
            var where = new List<string> { "1=1" };
            if (desde.HasValue) { where.Add("Timestamp >= @Desde"); cmd.Parameters.AddWithValue("@Desde", desde.Value); }
            if (hasta.HasValue) { where.Add("Timestamp <= @Hasta"); cmd.Parameters.AddWithValue("@Hasta", hasta.Value.AddDays(1)); }
            cmd.CommandText = $"""
                SELECT UserName, Tipo, Lat, Lng, Precision, EsDentroDeZona, Notas, Timestamp
                FROM dbo.CheckadorRegistros WHERE {string.Join(" AND ", where)}
                ORDER BY Timestamp DESC
                """;

            using var wb  = new ClosedXML.Excel.XLWorkbook();
            var ws = wb.Worksheets.Add("Asistencia");

            // Header
            string[] headers = ["Usuario","Tipo","Latitud","Longitud","Precisión (m)","Dentro de Zona","Notas","Fecha y Hora"];
            for (int i = 0; i < headers.Length; i++)
            {
                var cell = ws.Cell(1, i + 1);
                cell.Value = headers[i];
                cell.Style.Font.Bold = true;
                cell.Style.Fill.BackgroundColor = ClosedXML.Excel.XLColor.FromHtml("#0d2137");
                cell.Style.Font.FontColor       = ClosedXML.Excel.XLColor.White;
            }

            await using var r = await cmd.ExecuteReaderAsync(ct);
            int row = 2;
            while (await r.ReadAsync(ct))
            {
                static string S(SqlDataReader rd, int i) => rd.IsDBNull(i) ? "" : rd.GetValue(i).ToString()!;
                ws.Cell(row, 1).Value = S(r, 0);
                ws.Cell(row, 2).Value = S(r, 1);
                ws.Cell(row, 3).Value = S(r, 2);
                ws.Cell(row, 4).Value = S(r, 3);
                ws.Cell(row, 5).Value = S(r, 4);
                ws.Cell(row, 6).Value = r.IsDBNull(5) ? "" : (r.GetValue(5).ToString() == "True" ? "Sí" : "No");
                ws.Cell(row, 7).Value = S(r, 6);
                ws.Cell(row, 8).Value = r.IsDBNull(7) ? "" : r.GetDateTime(7).ToString("yyyy-MM-dd HH:mm:ss");
                row++;
            }

            ws.Columns().AdjustToContents();
            using var ms = new MemoryStream();
            wb.SaveAs(ms);
            return File(ms.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                $"asistencia_{DateTime.Now:yyyyMMdd}.xlsx");
        }
        catch (Exception ex) { logger.LogError(ex, "Checador.Export"); return StatusCode(500, ex.Message); }
    }

    // ── CRUD Sitios (Propuesta B) ─────────────────────────────────────────────
    [HttpGet("sitios")]
    [Authorize(Roles = "Admin")]
    public async Task<IActionResult> GetSitios(CancellationToken ct)
    {
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = "SELECT Id,Nombre,Descripcion,Lat,Lng,RadioMetros,Activo,CreadoEn FROM dbo.CheckadorSitios ORDER BY Nombre";
            var list = new List<object>();
            await using var r = await cmd.ExecuteReaderAsync(ct);
            while (await r.ReadAsync(ct))
                list.Add(new {
                    id          = r.GetGuid(0), nombre = r.GetString(1),
                    descripcion = r.IsDBNull(2) ? null : r.GetString(2),
                    lat         = r.GetDouble(3), lng = r.GetDouble(4),
                    radioMetros = r.GetInt32(5),  activo = r.GetBoolean(6),
                    creadoEn    = r.GetDateTime(7),
                });
            return Ok(list);
        }
        catch (Exception ex) { logger.LogError(ex, "Checador.GetSitios"); return StatusCode(500, ex.Message); }
    }

    [HttpPost("sitios")]
    [Authorize(Roles = "Admin")]
    public async Task<IActionResult> CreateSitio([FromBody] SitioDto dto, CancellationToken ct)
    {
        try
        {
            var id = Guid.NewGuid();
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = "INSERT INTO dbo.CheckadorSitios (Id,Nombre,Descripcion,Lat,Lng,RadioMetros) VALUES (@Id,@N,@D,@Lat,@Lng,@R)";
            cmd.Parameters.AddWithValue("@Id",  id);
            cmd.Parameters.AddWithValue("@N",   dto.Nombre);
            cmd.Parameters.AddWithValue("@D",   (object?)dto.Descripcion ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@Lat", dto.Lat);
            cmd.Parameters.AddWithValue("@Lng", dto.Lng);
            cmd.Parameters.AddWithValue("@R",   dto.RadioMetros > 0 ? dto.RadioMetros : 100);
            await cmd.ExecuteNonQueryAsync(ct);
            return StatusCode(201, new { id });
        }
        catch (Exception ex) { logger.LogError(ex, "Checador.CreateSitio"); return StatusCode(500, ex.Message); }
    }

    [HttpPut("sitios/{id:guid}")]
    [Authorize(Roles = "Admin")]
    public async Task<IActionResult> UpdateSitio(Guid id, [FromBody] SitioDto dto, CancellationToken ct)
    {
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = "UPDATE dbo.CheckadorSitios SET Nombre=@N,Descripcion=@D,Lat=@Lat,Lng=@Lng,RadioMetros=@R,Activo=@A WHERE Id=@Id";
            cmd.Parameters.AddWithValue("@Id",  id);
            cmd.Parameters.AddWithValue("@N",   dto.Nombre);
            cmd.Parameters.AddWithValue("@D",   (object?)dto.Descripcion ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@Lat", dto.Lat);
            cmd.Parameters.AddWithValue("@Lng", dto.Lng);
            cmd.Parameters.AddWithValue("@R",   dto.RadioMetros > 0 ? dto.RadioMetros : 100);
            cmd.Parameters.AddWithValue("@A",   dto.Activo);
            await cmd.ExecuteNonQueryAsync(ct);
            return NoContent();
        }
        catch (Exception ex) { logger.LogError(ex, "Checador.UpdateSitio {Id}", id); return StatusCode(500, ex.Message); }
    }

    [HttpDelete("sitios/{id:guid}")]
    [Authorize(Roles = "Admin")]
    public async Task<IActionResult> DeleteSitio(Guid id, CancellationToken ct)
    {
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = "DELETE FROM dbo.CheckadorSitios WHERE Id=@Id";
            cmd.Parameters.AddWithValue("@Id", id);
            await cmd.ExecuteNonQueryAsync(ct);
            return NoContent();
        }
        catch (Exception ex) { logger.LogError(ex, "Checador.DeleteSitio {Id}", id); return StatusCode(500, ex.Message); }
    }

    // ── Helpers ───────────────────────────────────────────────────────────────
    private static object MapRegistro(SqlDataReader r) => new {
        id             = r.GetGuid(0),
        tipo           = r.GetString(1),
        lat            = r.IsDBNull(2) ? (double?)null : r.GetDouble(2),
        lng            = r.IsDBNull(3) ? (double?)null : r.GetDouble(3),
        precision      = r.IsDBNull(4) ? (double?)null : r.GetDouble(4),
        esDentroDeZona = r.IsDBNull(5) ? (bool?)null   : r.GetBoolean(5),
        notas          = r.IsDBNull(6) ? null           : r.GetString(6),
        timestamp      = r.GetDateTime(7),
    };

    /// <summary>Fórmula Haversine — distancia en metros entre dos coordenadas.</summary>
    private static double HaversineMetros(double lat1, double lng1, double lat2, double lng2)
    {
        const double R = 6371000;
        var dLat = (lat2 - lat1) * Math.PI / 180;
        var dLng = (lng2 - lng1) * Math.PI / 180;
        var a    = Math.Sin(dLat/2) * Math.Sin(dLat/2)
                 + Math.Cos(lat1 * Math.PI/180) * Math.Cos(lat2 * Math.PI/180)
                 * Math.Sin(dLng/2) * Math.Sin(dLng/2);
        return R * 2 * Math.Atan2(Math.Sqrt(a), Math.Sqrt(1-a));
    }
}

// ── DTOs ──────────────────────────────────────────────────────────────────────
public record MarcarDto(string Tipo, double? Lat, double? Lng, double? Precision, string? Notas);
public record SitioDto(string Nombre, string? Descripcion, double Lat, double Lng, int RadioMetros, bool Activo = true);
