using MailKit.Net.Smtp;
using MailKit.Security;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using MimeKit;
using System.Globalization;
using System.Security.Claims;

namespace Pandora.API.Controllers;

// ═══════════════════════════════════════════════════════════════════════════════
//  FLOTILLA — Reserva de Vehículos
//  Mismo patrón que CalendarController (salas): catálogo de recursos + reservas
//  con detección de choque de horario (StartTime < @End AND EndTime > @Start),
//  push a todos y correo al solicitante/conductor.
//  Bit de módulo: 268435456 (FLOTILLA).
// ═══════════════════════════════════════════════════════════════════════════════
[ApiController]
[Route("api/flotilla")]
[Authorize]
public class FlotillaController(
    IConfiguration config,
    ILogger<FlotillaController> logger,
    Pandora.API.Services.PushService push,
    Pandora.API.Services.TenantService tenant) : ControllerBase
{
    private static readonly TimeZoneInfo MxTz = TimeZoneInfo.FindSystemTimeZoneById(
        OperatingSystem.IsWindows() ? "Central Standard Time" : "America/Mexico_City");

    private const int FLOTILLA_BIT = 268435456;

    private SqlConnection Conn() => new(config.GetConnectionString("PandoraDb"));
    private object TenantParam => (object?)tenant.TenantId ?? DBNull.Value;

    private string? CurrentUsername =>
        User.FindFirstValue(ClaimTypes.Name) ??
        User.FindFirstValue("name") ??
        User.Claims.FirstOrDefault(c => c.Type.EndsWith("name", StringComparison.OrdinalIgnoreCase))?.Value;

    private bool IsAdmin =>
        User.IsInRole("Admin") ||
        User.Claims.Any(c => c.Type == ClaimTypes.Role && c.Value == "Admin");

    private bool HasFlotilla
    {
        get
        {
            if (IsAdmin) return true;
            var claim = User.Claims.FirstOrDefault(c => c.Type == "modules")?.Value;
            return int.TryParse(claim, out int m) && (m & FLOTILLA_BIT) != 0;
        }
    }

    // ════════════════════════════════════════════════════════════════════════
    //  VEHÍCULOS (catálogo) — gestión solo Admin, lectura cualquiera con el bit
    // ════════════════════════════════════════════════════════════════════════

    [HttpGet("vehiculos")]
    public async Task<IActionResult> GetVehiculos([FromQuery] bool includeInactive = false, CancellationToken ct = default)
    {
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await using var cmd = conn.CreateCommand();
            var sql = """
                SELECT Id, Nombre, Placa, Marca, Modelo, Anio, Color, Capacidad, ColorEtiqueta,
                       CASE WHEN FotoData IS NULL THEN 0 ELSE 1 END AS HasFoto,
                       Notas, IsActive, CreatedAt
                FROM dbo.Vehiculos WHERE TenantId = @TenantId
                """;
            if (!includeInactive) sql += " AND IsActive = 1";
            sql += " ORDER BY Nombre";
            cmd.CommandText = sql;
            cmd.Parameters.AddWithValue("@TenantId", TenantParam);

            var list = new List<object>();
            await using var r = await cmd.ExecuteReaderAsync(ct);
            while (await r.ReadAsync(ct))
                list.Add(new
                {
                    id            = r.GetGuid(0),
                    nombre        = r.GetString(1),
                    placa         = r.IsDBNull(2) ? null : r.GetString(2),
                    marca         = r.IsDBNull(3) ? null : r.GetString(3),
                    modelo        = r.IsDBNull(4) ? null : r.GetString(4),
                    anio          = r.IsDBNull(5) ? (int?)null : r.GetInt32(5),
                    color         = r.IsDBNull(6) ? null : r.GetString(6),
                    capacidad     = r.GetInt32(7),
                    colorEtiqueta = r.IsDBNull(8) ? null : r.GetString(8),
                    hasFoto       = r.GetInt32(9) == 1,
                    notas         = r.IsDBNull(10) ? null : r.GetString(10),
                    isActive      = r.GetBoolean(11),
                    createdAt     = DateTime.SpecifyKind(r.GetDateTime(12), DateTimeKind.Utc),
                });
            return Ok(list);
        }
        catch (Exception ex) { logger.LogError(ex, "GetVehiculos"); return StatusCode(500, ex.Message); }
    }

    [HttpGet("vehiculos/{id:guid}/foto")]
    public async Task<IActionResult> GetVehiculoFoto(Guid id, CancellationToken ct)
    {
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = "SELECT FotoData, FotoMime FROM dbo.Vehiculos WHERE Id = @Id AND TenantId = @TenantId";
            cmd.Parameters.AddWithValue("@Id", id);
            cmd.Parameters.AddWithValue("@TenantId", TenantParam);
            await using var r = await cmd.ExecuteReaderAsync(ct);
            if (!await r.ReadAsync(ct) || r.IsDBNull(0)) return NotFound();
            var bytes = (byte[])r[0];
            var mime  = r.IsDBNull(1) ? "image/jpeg" : r.GetString(1);
            return File(bytes, mime);
        }
        catch (Exception ex) { logger.LogError(ex, "GetVehiculoFoto {Id}", id); return StatusCode(500, ex.Message); }
    }

    [HttpPost("vehiculos")]
    [Authorize(Roles = "Admin")]
    public async Task<IActionResult> CreateVehiculo([FromBody] VehiculoDto dto, CancellationToken ct)
    {
        if (string.IsNullOrWhiteSpace(dto.Nombre)) return BadRequest("Nombre requerido.");
        var id = Guid.NewGuid();
        try
        {
            var (fotoBytes, fotoMime) = ParseDataUri(dto.FotoBase64);
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = """
                INSERT INTO dbo.Vehiculos
                    (Id, Nombre, Placa, Marca, Modelo, Anio, Color, Capacidad, ColorEtiqueta,
                     FotoData, FotoMime, Notas, IsActive, CreatedAt, TenantId)
                VALUES
                    (@Id, @Nombre, @Placa, @Marca, @Modelo, @Anio, @Color, @Capacidad, @ColorEtiqueta,
                     @FotoData, @FotoMime, @Notas, @IsActive, GETUTCDATE(), @TenantId)
                """;
            BindVehiculo(cmd, dto, fotoBytes, fotoMime);
            cmd.Parameters.AddWithValue("@Id", id);
            cmd.Parameters.AddWithValue("@TenantId", TenantParam);
            await cmd.ExecuteNonQueryAsync(ct);
            return Ok(new { id });
        }
        catch (Exception ex) { logger.LogError(ex, "CreateVehiculo"); return StatusCode(500, ex.Message); }
    }

    [HttpPut("vehiculos/{id:guid}")]
    [Authorize(Roles = "Admin")]
    public async Task<IActionResult> UpdateVehiculo(Guid id, [FromBody] VehiculoDto dto, CancellationToken ct)
    {
        if (string.IsNullOrWhiteSpace(dto.Nombre)) return BadRequest("Nombre requerido.");
        try
        {
            var (fotoBytes, fotoMime) = ParseDataUri(dto.FotoBase64);
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await using var cmd = conn.CreateCommand();
            // Si FotoBase64 viene null → no se toca la foto existente; si viene "" → se borra.
            var fotoSql = dto.FotoBase64 is null
                ? ""
                : ", FotoData = @FotoData, FotoMime = @FotoMime";
            cmd.CommandText = $"""
                UPDATE dbo.Vehiculos
                SET Nombre = @Nombre, Placa = @Placa, Marca = @Marca, Modelo = @Modelo,
                    Anio = @Anio, Color = @Color, Capacidad = @Capacidad, ColorEtiqueta = @ColorEtiqueta,
                    Notas = @Notas, IsActive = @IsActive{fotoSql}, UpdatedAt = GETUTCDATE()
                WHERE Id = @Id AND TenantId = @TenantId
                """;
            BindVehiculo(cmd, dto, fotoBytes, fotoMime);
            cmd.Parameters.AddWithValue("@Id", id);
            cmd.Parameters.AddWithValue("@TenantId", TenantParam);
            int rows = await cmd.ExecuteNonQueryAsync(ct);
            if (rows == 0) return NotFound("Vehículo no encontrado.");
            return Ok(new { id });
        }
        catch (Exception ex) { logger.LogError(ex, "UpdateVehiculo {Id}", id); return StatusCode(500, ex.Message); }
    }

    [HttpDelete("vehiculos/{id:guid}")]
    [Authorize(Roles = "Admin")]
    public async Task<IActionResult> DeleteVehiculo(Guid id, CancellationToken ct)
    {
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = "UPDATE dbo.Vehiculos SET IsActive = 0, UpdatedAt = GETUTCDATE() WHERE Id = @Id AND TenantId = @TenantId";
            cmd.Parameters.AddWithValue("@Id", id);
            cmd.Parameters.AddWithValue("@TenantId", TenantParam);
            int rows = await cmd.ExecuteNonQueryAsync(ct);
            if (rows == 0) return NotFound("Vehículo no encontrado.");
            return NoContent();
        }
        catch (Exception ex) { logger.LogError(ex, "DeleteVehiculo {Id}", id); return StatusCode(500, ex.Message); }
    }

    private static void BindVehiculo(SqlCommand cmd, VehiculoDto dto, byte[]? fotoBytes, string? fotoMime)
    {
        cmd.Parameters.AddWithValue("@Nombre",        dto.Nombre.Trim());
        cmd.Parameters.AddWithValue("@Placa",         (object?)dto.Placa         ?? DBNull.Value);
        cmd.Parameters.AddWithValue("@Marca",         (object?)dto.Marca         ?? DBNull.Value);
        cmd.Parameters.AddWithValue("@Modelo",        (object?)dto.Modelo        ?? DBNull.Value);
        cmd.Parameters.AddWithValue("@Anio",          (object?)dto.Anio          ?? DBNull.Value);
        cmd.Parameters.AddWithValue("@Color",         (object?)dto.Color         ?? DBNull.Value);
        cmd.Parameters.AddWithValue("@Capacidad",     dto.Capacidad <= 0 ? 5 : dto.Capacidad);
        cmd.Parameters.AddWithValue("@ColorEtiqueta", (object?)dto.ColorEtiqueta ?? DBNull.Value);
        cmd.Parameters.AddWithValue("@Notas",         (object?)dto.Notas         ?? DBNull.Value);
        cmd.Parameters.AddWithValue("@IsActive",      dto.IsActive);
        cmd.Parameters.AddWithValue("@FotoData",      (object?)fotoBytes ?? DBNull.Value);
        cmd.Parameters.AddWithValue("@FotoMime",      (object?)fotoMime  ?? DBNull.Value);
    }

    // data:image/png;base64,XXXX → (bytes, mime). "" → (null,null) borrado; null → no tocar.
    private static (byte[]?, string?) ParseDataUri(string? dataUri)
    {
        if (string.IsNullOrWhiteSpace(dataUri)) return (null, null);
        try
        {
            if (dataUri.StartsWith("data:", StringComparison.OrdinalIgnoreCase))
            {
                var comma = dataUri.IndexOf(',');
                var meta  = dataUri[5..comma]; // image/png;base64
                var mime  = meta.Split(';')[0];
                var bytes = Convert.FromBase64String(dataUri[(comma + 1)..]);
                return (bytes, mime);
            }
            return (Convert.FromBase64String(dataUri), "image/jpeg");
        }
        catch { return (null, null); }
    }

    // ════════════════════════════════════════════════════════════════════════
    //  RESERVAS
    // ════════════════════════════════════════════════════════════════════════

    [HttpGet("reservas")]
    public async Task<IActionResult> GetReservas(
        [FromQuery] DateTime? rangeStart,
        [FromQuery] DateTime? rangeEnd,
        [FromQuery] Guid? vehiculoId,
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
                SELECT res.Id, res.VehiculoId, res.Motivo, res.Destino, res.Conductor,
                       res.SolicitanteNombre, res.SolicitanteEmail, res.Pasajeros,
                       res.StartTime, res.EndTime, res.Estado, res.KmSalida, res.KmRegreso,
                       res.NivelGasolinaSalida, res.Notas, res.CreadoPor, res.CreatedAt,
                       v.Nombre AS VehNombre, v.Placa AS VehPlaca, v.ColorEtiqueta AS VehColor
                FROM dbo.VehiculoReservas res
                INNER JOIN dbo.Vehiculos v ON res.VehiculoId = v.Id
                WHERE res.TenantId = @TenantId AND res.StartTime < @End AND res.EndTime > @Start
                """;
            if (vehiculoId.HasValue) sql += " AND res.VehiculoId = @VehiculoId";
            sql += " ORDER BY res.StartTime";
            cmd.CommandText = sql;
            cmd.Parameters.AddWithValue("@TenantId", TenantParam);
            cmd.Parameters.AddWithValue("@Start", start);
            cmd.Parameters.AddWithValue("@End",   end);
            if (vehiculoId.HasValue) cmd.Parameters.AddWithValue("@VehiculoId", vehiculoId.Value);

            var list = new List<object>();
            await using var r = await cmd.ExecuteReaderAsync(ct);
            while (await r.ReadAsync(ct)) list.Add(ReadReserva(r));
            return Ok(list);
        }
        catch (Exception ex) { logger.LogError(ex, "GetReservas"); return StatusCode(500, ex.Message); }
    }

    [HttpGet("reservas/check-conflict")]
    public async Task<IActionResult> CheckConflict(
        [FromQuery] Guid vehiculoId,
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
                SELECT COUNT(1) FROM dbo.VehiculoReservas
                WHERE TenantId = @TenantId AND VehiculoId = @VehiculoId
                  AND Estado <> 'Cancelada' AND StartTime < @End AND EndTime > @Start
                """;
            if (excludeId.HasValue) sql += " AND Id <> @ExcludeId";
            cmd.CommandText = sql;
            cmd.Parameters.AddWithValue("@TenantId", TenantParam);
            cmd.Parameters.AddWithValue("@VehiculoId", vehiculoId);
            cmd.Parameters.AddWithValue("@Start", start);
            cmd.Parameters.AddWithValue("@End",   end);
            if (excludeId.HasValue) cmd.Parameters.AddWithValue("@ExcludeId", excludeId.Value);
            var count = (int)await cmd.ExecuteScalarAsync(ct)!;
            return Ok(new { hasConflict = count > 0 });
        }
        catch (Exception ex) { logger.LogError(ex, "CheckConflict"); return StatusCode(500, ex.Message); }
    }

    [HttpGet("reservas/{id:guid}")]
    public async Task<IActionResult> GetReservaById(Guid id, CancellationToken ct)
    {
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = """
                SELECT res.Id, res.VehiculoId, res.Motivo, res.Destino, res.Conductor,
                       res.SolicitanteNombre, res.SolicitanteEmail, res.Pasajeros,
                       res.StartTime, res.EndTime, res.Estado, res.KmSalida, res.KmRegreso,
                       res.NivelGasolinaSalida, res.Notas, res.CreadoPor, res.CreatedAt,
                       v.Nombre AS VehNombre, v.Placa AS VehPlaca, v.ColorEtiqueta AS VehColor
                FROM dbo.VehiculoReservas res
                INNER JOIN dbo.Vehiculos v ON res.VehiculoId = v.Id
                WHERE res.Id = @Id AND res.TenantId = @TenantId
                """;
            cmd.Parameters.AddWithValue("@Id", id);
            cmd.Parameters.AddWithValue("@TenantId", TenantParam);
            await using var r = await cmd.ExecuteReaderAsync(ct);
            if (!await r.ReadAsync(ct)) return NotFound("Reserva no encontrada.");
            return Ok(ReadReserva(r));
        }
        catch (Exception ex) { logger.LogError(ex, "GetReservaById {Id}", id); return StatusCode(500, ex.Message); }
    }

    [HttpPost("reservas")]
    public async Task<IActionResult> CreateReserva([FromBody] ReservaDto dto, CancellationToken ct)
    {
        if (!HasFlotilla)                           return Forbid();
        if (string.IsNullOrWhiteSpace(dto.Motivo))  return BadRequest("Motivo requerido.");
        if (dto.VehiculoId == Guid.Empty)           return BadRequest("Vehículo requerido.");
        if (dto.StartTime >= dto.EndTime)           return BadRequest("La hora de inicio debe ser antes del fin.");

        var id = Guid.NewGuid();
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);

            // Vehículo válido + activo
            string vehNombre;
            using (var vc = conn.CreateCommand())
            {
                vc.CommandText = "SELECT Nombre FROM dbo.Vehiculos WHERE Id = @Id AND IsActive = 1 AND TenantId = @TenantId";
                vc.Parameters.AddWithValue("@Id", dto.VehiculoId);
                vc.Parameters.AddWithValue("@TenantId", TenantParam);
                var nameObj = await vc.ExecuteScalarAsync(ct);
                if (nameObj is null or DBNull) return BadRequest("Vehículo no encontrado o inactivo.");
                vehNombre = (string)nameObj;
            }

            // Choque de horario
            int conflicts;
            using (var cc = conn.CreateCommand())
            {
                cc.CommandText = """
                    SELECT COUNT(1) FROM dbo.VehiculoReservas
                    WHERE TenantId = @TenantId AND VehiculoId = @VehiculoId
                      AND Estado <> 'Cancelada' AND StartTime < @End AND EndTime > @Start
                    """;
                cc.Parameters.AddWithValue("@TenantId", TenantParam);
                cc.Parameters.AddWithValue("@VehiculoId", dto.VehiculoId);
                cc.Parameters.AddWithValue("@Start", dto.StartTime);
                cc.Parameters.AddWithValue("@End",   dto.EndTime);
                conflicts = (int)(await cc.ExecuteScalarAsync(ct))!;
            }
            if (conflicts > 0)
                return Conflict($"El vehículo '{vehNombre}' ya está reservado en ese horario. Elige otro horario u otro vehículo.");

            await using var cmd = conn.CreateCommand();
            cmd.CommandText = """
                INSERT INTO dbo.VehiculoReservas
                    (Id, VehiculoId, Motivo, Destino, Conductor, SolicitanteNombre, SolicitanteEmail,
                     Pasajeros, StartTime, EndTime, Estado, KmSalida, KmRegreso, NivelGasolinaSalida,
                     Notas, CreadoPor, CreatedAt, TenantId)
                VALUES
                    (@Id, @VehiculoId, @Motivo, @Destino, @Conductor, @SolNombre, @SolEmail,
                     @Pasajeros, @Start, @End, @Estado, @KmSalida, @KmRegreso, @Gasolina,
                     @Notas, @CreadoPor, GETUTCDATE(), @TenantId)
                """;
            BindReserva(cmd, dto);
            cmd.Parameters.AddWithValue("@Id", id);
            cmd.Parameters.AddWithValue("@CreadoPor", (object?)CurrentUsername ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@TenantId", TenantParam);
            await cmd.ExecuteNonQueryAsync(ct);

            var username = CurrentUsername;
            _ = Task.Run(() => SendReservaEmailAsync(dto, id, vehNombre, isUpdate: false, username), CancellationToken.None);
            _ = Task.Run(() => push.SendToAllAsync(
                $"🚗 Nueva reserva de vehículo: {vehNombre}",
                $"{dto.Motivo} — {dto.StartTime:dd/MM HH:mm} a {dto.EndTime:HH:mm}",
                "/flotilla"), CancellationToken.None);

            return Ok(new { id });
        }
        catch (Exception ex) { logger.LogError(ex, "CreateReserva"); return StatusCode(500, ex.Message); }
    }

    [HttpPut("reservas/{id:guid}")]
    public async Task<IActionResult> UpdateReserva(Guid id, [FromBody] ReservaDto dto, CancellationToken ct)
    {
        if (!HasFlotilla)                          return Forbid();
        if (string.IsNullOrWhiteSpace(dto.Motivo)) return BadRequest("Motivo requerido.");
        if (dto.StartTime >= dto.EndTime)          return BadRequest("La hora de inicio debe ser antes del fin.");
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);

            string vehNombre;
            using (var vc = conn.CreateCommand())
            {
                vc.CommandText = "SELECT Nombre FROM dbo.Vehiculos WHERE Id = @Id AND IsActive = 1 AND TenantId = @TenantId";
                vc.Parameters.AddWithValue("@Id", dto.VehiculoId);
                vc.Parameters.AddWithValue("@TenantId", TenantParam);
                var nameObj = await vc.ExecuteScalarAsync(ct);
                if (nameObj is null or DBNull) return BadRequest("Vehículo no encontrado o inactivo.");
                vehNombre = (string)nameObj;
            }

            int conflicts;
            using (var cc = conn.CreateCommand())
            {
                cc.CommandText = """
                    SELECT COUNT(1) FROM dbo.VehiculoReservas
                    WHERE TenantId = @TenantId AND VehiculoId = @VehiculoId
                      AND Estado <> 'Cancelada' AND StartTime < @End AND EndTime > @Start AND Id <> @Id
                    """;
                cc.Parameters.AddWithValue("@TenantId", TenantParam);
                cc.Parameters.AddWithValue("@VehiculoId", dto.VehiculoId);
                cc.Parameters.AddWithValue("@Start", dto.StartTime);
                cc.Parameters.AddWithValue("@End",   dto.EndTime);
                cc.Parameters.AddWithValue("@Id",    id);
                conflicts = (int)(await cc.ExecuteScalarAsync(ct))!;
            }
            if (conflicts > 0)
                return Conflict($"El vehículo '{vehNombre}' ya está reservado en ese horario. Elige otro horario u otro vehículo.");

            await using var cmd = conn.CreateCommand();
            cmd.CommandText = """
                UPDATE dbo.VehiculoReservas
                SET VehiculoId = @VehiculoId, Motivo = @Motivo, Destino = @Destino, Conductor = @Conductor,
                    SolicitanteNombre = @SolNombre, SolicitanteEmail = @SolEmail, Pasajeros = @Pasajeros,
                    StartTime = @Start, EndTime = @End, Estado = @Estado, KmSalida = @KmSalida,
                    KmRegreso = @KmRegreso, NivelGasolinaSalida = @Gasolina, Notas = @Notas,
                    UpdatedAt = GETUTCDATE()
                WHERE Id = @Id AND TenantId = @TenantId
                """;
            BindReserva(cmd, dto);
            cmd.Parameters.AddWithValue("@Id", id);
            cmd.Parameters.AddWithValue("@TenantId", TenantParam);
            int rows = await cmd.ExecuteNonQueryAsync(ct);
            if (rows == 0) return NotFound("Reserva no encontrada.");

            var username = CurrentUsername;
            _ = Task.Run(() => SendReservaEmailAsync(dto, id, vehNombre, isUpdate: true, username), CancellationToken.None);
            return Ok(new { id });
        }
        catch (Exception ex) { logger.LogError(ex, "UpdateReserva {Id}", id); return StatusCode(500, ex.Message); }
    }

    [HttpDelete("reservas/{id:guid}")]
    public async Task<IActionResult> DeleteReserva(Guid id, CancellationToken ct)
    {
        if (!HasFlotilla) return Forbid();
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = "DELETE FROM dbo.VehiculoReservas WHERE Id = @Id AND TenantId = @TenantId";
            cmd.Parameters.AddWithValue("@Id", id);
            cmd.Parameters.AddWithValue("@TenantId", TenantParam);
            int rows = await cmd.ExecuteNonQueryAsync(ct);
            if (rows == 0) return NotFound("Reserva no encontrada.");
            return NoContent();
        }
        catch (Exception ex) { logger.LogError(ex, "DeleteReserva {Id}", id); return StatusCode(500, ex.Message); }
    }

    private static void BindReserva(SqlCommand cmd, ReservaDto dto)
    {
        cmd.Parameters.AddWithValue("@VehiculoId", dto.VehiculoId);
        cmd.Parameters.AddWithValue("@Motivo",     dto.Motivo.Trim());
        cmd.Parameters.AddWithValue("@Destino",    (object?)dto.Destino            ?? DBNull.Value);
        cmd.Parameters.AddWithValue("@Conductor",  (object?)dto.Conductor          ?? DBNull.Value);
        cmd.Parameters.AddWithValue("@SolNombre",  (object?)dto.SolicitanteNombre  ?? DBNull.Value);
        cmd.Parameters.AddWithValue("@SolEmail",   (object?)dto.SolicitanteEmail   ?? DBNull.Value);
        cmd.Parameters.AddWithValue("@Pasajeros",  (object?)dto.Pasajeros          ?? DBNull.Value);
        cmd.Parameters.AddWithValue("@Start",      dto.StartTime);
        cmd.Parameters.AddWithValue("@End",        dto.EndTime);
        cmd.Parameters.AddWithValue("@Estado",     string.IsNullOrWhiteSpace(dto.Estado) ? "Activa" : dto.Estado.Trim());
        cmd.Parameters.AddWithValue("@KmSalida",   (object?)dto.KmSalida           ?? DBNull.Value);
        cmd.Parameters.AddWithValue("@KmRegreso",  (object?)dto.KmRegreso          ?? DBNull.Value);
        cmd.Parameters.AddWithValue("@Gasolina",   (object?)dto.NivelGasolinaSalida?? DBNull.Value);
        cmd.Parameters.AddWithValue("@Notas",      (object?)dto.Notas              ?? DBNull.Value);
    }

    private static object ReadReserva(SqlDataReader r) => new
    {
        id                  = r.GetGuid(0),
        vehiculoId          = r.GetGuid(1),
        motivo              = r.GetString(2),
        destino             = r.IsDBNull(3)  ? null : r.GetString(3),
        conductor           = r.IsDBNull(4)  ? null : r.GetString(4),
        solicitanteNombre   = r.IsDBNull(5)  ? null : r.GetString(5),
        solicitanteEmail    = r.IsDBNull(6)  ? null : r.GetString(6),
        pasajeros           = r.IsDBNull(7)  ? null : r.GetString(7),
        startTime           = DateTime.SpecifyKind(r.GetDateTime(8), DateTimeKind.Utc),
        endTime             = DateTime.SpecifyKind(r.GetDateTime(9), DateTimeKind.Utc),
        estado              = r.GetString(10),
        kmSalida            = r.IsDBNull(11) ? (int?)null : r.GetInt32(11),
        kmRegreso           = r.IsDBNull(12) ? (int?)null : r.GetInt32(12),
        nivelGasolinaSalida = r.IsDBNull(13) ? null : r.GetString(13),
        notas               = r.IsDBNull(14) ? null : r.GetString(14),
        creadoPor           = r.IsDBNull(15) ? null : r.GetString(15),
        createdAt           = DateTime.SpecifyKind(r.GetDateTime(16), DateTimeKind.Utc),
        vehiculoNombre      = r.IsDBNull(17) ? null : r.GetString(17),
        vehiculoPlaca       = r.IsDBNull(18) ? null : r.GetString(18),
        vehiculoColor       = r.IsDBNull(19) ? null : r.GetString(19),
    };

    // ── Correo al solicitante / conductor ─────────────────────────────────────
    private async Task SendReservaEmailAsync(ReservaDto dto, Guid reservaId, string vehNombre, bool isUpdate, string? username)
    {
        try
        {
            var smtp     = config.GetSection("SmtpSettings");
            var host     = smtp["Host"] ?? "";
            var port     = int.TryParse(smtp["Port"], out var p) ? p : 587;
            var fromName = smtp["FromName"] ?? "Pandora";

            string? smtpEmail = null, smtpPass = null;
            if (!string.IsNullOrWhiteSpace(username))
            {
                try
                {
                    await using var userConn = Conn();
                    await userConn.OpenAsync();
                    using var userCmd = userConn.CreateCommand();
                    userCmd.CommandText = "SELECT SmtpEmail, SmtpPassword FROM dbo.AppUsers WHERE LOWER(Username) = LOWER(@User)";
                    userCmd.Parameters.AddWithValue("@User", username);
                    await using var ur = await userCmd.ExecuteReaderAsync();
                    if (await ur.ReadAsync())
                    {
                        smtpEmail = ur.IsDBNull(0) ? null : ur.GetString(0);
                        smtpPass  = ur.IsDBNull(1) ? null : ur.GetString(1);
                    }
                }
                catch (Exception ex) { logger.LogWarning("No SMTP for {User}: {Msg}", username, ex.Message); }
            }

            var from = smtpEmail ?? smtp["FromEmail"] ?? "";
            var pass = smtpPass  ?? smtp["Password"]  ?? "";
            if (string.IsNullOrWhiteSpace(host) || string.IsNullOrWhiteSpace(from) || string.IsNullOrWhiteSpace(pass))
                return;

            if (string.IsNullOrWhiteSpace(dto.SolicitanteEmail)) return;

            var mxStart  = TimeZoneInfo.ConvertTimeFromUtc(dto.StartTime, MxTz);
            var mxEnd    = TimeZoneInfo.ConvertTimeFromUtc(dto.EndTime,   MxTz);
            var dateStr  = mxStart.ToString("dddd, dd 'de' MMMM 'de' yyyy", new CultureInfo("es-MX"));
            var subject  = isUpdate
                ? $"[Pandora] Reserva de vehículo actualizada: {dto.Motivo}"
                : $"[Pandora] Nueva reserva de vehículo: {dto.Motivo}";

            var body = $"""
                <div style="font-family:Arial,sans-serif;max-width:560px;margin:auto">
                  <div style="background:#1a237e;color:#fff;padding:16px 20px;border-radius:8px 8px 0 0">
                    <h2 style="margin:0">🚗 Reserva de Vehículo</h2>
                  </div>
                  <div style="border:1px solid #e0e0e0;border-top:none;padding:20px;border-radius:0 0 8px 8px">
                    <p>Hola {dto.SolicitanteNombre ?? "—"},</p>
                    <p>{(isUpdate ? "Se actualizó" : "Se registró")} tu reserva del vehículo <b>{vehNombre}</b>.</p>
                    <table style="width:100%;border-collapse:collapse">
                      <tr><td style="padding:6px 0;color:#666">Motivo</td><td style="padding:6px 0"><b>{dto.Motivo}</b></td></tr>
                      <tr><td style="padding:6px 0;color:#666">Destino</td><td style="padding:6px 0">{dto.Destino ?? "—"}</td></tr>
                      <tr><td style="padding:6px 0;color:#666">Conductor</td><td style="padding:6px 0">{dto.Conductor ?? "—"}</td></tr>
                      <tr><td style="padding:6px 0;color:#666">Fecha</td><td style="padding:6px 0">{dateStr}</td></tr>
                      <tr><td style="padding:6px 0;color:#666">Horario</td><td style="padding:6px 0">{mxStart:HH:mm} a {mxEnd:HH:mm}</td></tr>
                    </table>
                    <p style="color:#888;font-size:12px;margin-top:16px">Pandora — Reserva de Vehículos</p>
                  </div>
                </div>
                """;

            using var client = new SmtpClient();
            await client.ConnectAsync(host, port, SecureSocketOptions.StartTls);
            await client.AuthenticateAsync(from, pass);
            var msg = new MimeMessage();
            msg.From.Add(new MailboxAddress(fromName, from));
            msg.To.Add(new MailboxAddress(dto.SolicitanteNombre ?? dto.SolicitanteEmail, dto.SolicitanteEmail));
            msg.Subject = subject;
            msg.Body    = new TextPart("html") { Text = body };
            await client.SendAsync(msg);
            await client.DisconnectAsync(true);
            logger.LogInformation("Reserva vehículo email sent for {Id}", reservaId);
        }
        catch (Exception ex) { logger.LogWarning("Reserva email failed for {Id}: {Msg}", reservaId, ex.Message); }
    }
}

// ── DTOs ───────────────────────────────────────────────────────────────────────
public record VehiculoDto(
    string Nombre, string? Placa, string? Marca, string? Modelo, int? Anio, string? Color,
    int Capacidad, string? ColorEtiqueta, string? FotoBase64, string? Notas, bool IsActive = true);

public record ReservaDto(
    Guid VehiculoId, string Motivo, string? Destino, string? Conductor,
    string? SolicitanteNombre, string? SolicitanteEmail, string? Pasajeros,
    DateTime StartTime, DateTime EndTime, string? Estado,
    int? KmSalida, int? KmRegreso, string? NivelGasolinaSalida, string? Notas);
