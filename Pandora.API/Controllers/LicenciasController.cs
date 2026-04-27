using MiniExcelLibs;
using Microsoft.AspNetCore.Authorization;
using IConfiguration = Microsoft.Extensions.Configuration.IConfiguration;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using System.Security.Claims;

namespace Pandora.API.Controllers;

[ApiController]
[Route("api/licencias")]
[Authorize]
public class LicenciasController(IConfiguration config, ILogger<LicenciasController> logger) : ControllerBase
{
    private SqlConnection Conn() => new(config.GetConnectionString("PandoraDb"));

    private bool IsAdmin =>
        User.FindFirstValue(ClaimTypes.Role) == "Admin" ||
        User.FindFirst(c => c.Type.EndsWith("role", StringComparison.OrdinalIgnoreCase))?.Value == "Admin";

    // ── GET /api/licencias ───────────────────────────────────────────────────
    [HttpGet]
    public async Task<IActionResult> GetAll(
        [FromQuery] string? area,
        [FromQuery] string? estado,
        CancellationToken ct)
    {
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = """
                SELECT Id, Numero, Plataforma, Area, Responsable, FrecuenciaPago,
                       FechaInicio, ProximoPago,
                       DATEDIFF(day, CAST(GETDATE() AS DATE), ProximoPago) AS DiasRestantes,
                       CostoMXN,
                       CASE FrecuenciaPago
                           WHEN 'Mensual'    THEN CostoMXN * 12
                           WHEN 'Trimestral' THEN CostoMXN * 4
                           WHEN 'Semestral'  THEN CostoMXN * 2
                           WHEN 'Anual'      THEN CostoMXN
                           ELSE 0
                       END AS CostoAnualMXN,
                       Estado, Notas, CreadoEn, ActualizadoEn
                FROM dbo.Licencias
                WHERE (@Area   IS NULL OR Area   = @Area)
                  AND (@Estado IS NULL OR Estado = @Estado)
                ORDER BY Numero
                """;
            cmd.Parameters.AddWithValue("@Area",   (object?)area   ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@Estado", (object?)estado ?? DBNull.Value);

            var list = new List<object>();
            await using var r = await cmd.ExecuteReaderAsync(ct);
            while (await r.ReadAsync(ct))
                list.Add(MapRow(r));
            return Ok(list);
        }
        catch (Exception ex) { logger.LogError(ex, "GetAll licencias"); return StatusCode(500, ex.Message); }
    }

    // ── GET /api/licencias/dashboard ─────────────────────────────────────────
    [HttpGet("dashboard")]
    public async Task<IActionResult> GetDashboard(CancellationToken ct)
    {
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);

            await using var cmd = conn.CreateCommand();
            cmd.CommandText = """
                SELECT
                    COUNT(*)                                                         AS Total,
                    SUM(CASE WHEN Estado = 'Activa'     THEN 1 ELSE 0 END)          AS Activas,
                    SUM(CASE WHEN Estado = 'Por vencer' THEN 1 ELSE 0 END)          AS PorVencer,
                    SUM(CASE WHEN Estado = 'Vencida'    THEN 1 ELSE 0 END)          AS Vencidas,
                    SUM(CASE WHEN Estado = 'Cancelada'  THEN 1 ELSE 0 END)          AS Canceladas,
                    ISNULL(SUM(CASE WHEN Estado != 'Cancelada' THEN
                        CASE FrecuenciaPago
                            WHEN 'Mensual'    THEN CostoMXN * 12
                            WHEN 'Trimestral' THEN CostoMXN * 4
                            WHEN 'Semestral'  THEN CostoMXN * 2
                            WHEN 'Anual'      THEN CostoMXN
                            ELSE 0 END
                    ELSE 0 END), 0)                                                  AS TotalAnualMXN,
                    ISNULL(SUM(CASE WHEN Estado != 'Cancelada' THEN
                        CASE FrecuenciaPago
                            WHEN 'Mensual'    THEN CostoMXN
                            WHEN 'Trimestral' THEN CostoMXN / 3.0
                            WHEN 'Semestral'  THEN CostoMXN / 6.0
                            WHEN 'Anual'      THEN CostoMXN / 12.0
                            ELSE 0 END
                    ELSE 0 END), 0)                                                  AS TotalMensualMXN
                FROM dbo.Licencias
                """;

            await using var r = await cmd.ExecuteReaderAsync(ct);
            await r.ReadAsync(ct);
            var stats = new
            {
                total           = r.GetInt32(r.GetOrdinal("Total")),
                activas         = r.GetInt32(r.GetOrdinal("Activas")),
                porVencer       = r.GetInt32(r.GetOrdinal("PorVencer")),
                vencidas        = r.GetInt32(r.GetOrdinal("Vencidas")),
                canceladas      = r.GetInt32(r.GetOrdinal("Canceladas")),
                totalAnualMXN   = r.GetDecimal(r.GetOrdinal("TotalAnualMXN")),
                totalMensualMXN = r.GetDecimal(r.GetOrdinal("TotalMensualMXN")),
            };
            r.Close();

            await using var cmd2 = conn.CreateCommand();
            cmd2.CommandText = """
                SELECT Area,
                       COUNT(*) AS Total,
                       ISNULL(SUM(CASE FrecuenciaPago
                           WHEN 'Mensual'    THEN CostoMXN * 12
                           WHEN 'Trimestral' THEN CostoMXN * 4
                           WHEN 'Semestral'  THEN CostoMXN * 2
                           WHEN 'Anual'      THEN CostoMXN
                           ELSE 0 END), 0) AS CostoAnual
                FROM dbo.Licencias
                WHERE Estado != 'Cancelada'
                GROUP BY Area
                ORDER BY CostoAnual DESC
                """;
            var areas = new List<object>();
            await using var r2 = await cmd2.ExecuteReaderAsync(ct);
            while (await r2.ReadAsync(ct))
                areas.Add(new
                {
                    area       = r2.GetString(r2.GetOrdinal("Area")),
                    total      = r2.GetInt32(r2.GetOrdinal("Total")),
                    costoAnual = r2.GetDecimal(r2.GetOrdinal("CostoAnual")),
                });

            return Ok(new { stats, areas });
        }
        catch (Exception ex) { logger.LogError(ex, "GetDashboard licencias"); return StatusCode(500, ex.Message); }
    }

    // ── GET /api/licencias/alertas ────────────────────────────────────────────
    [HttpGet("alertas")]
    public async Task<IActionResult> GetAlertas(CancellationToken ct)
    {
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = """
                SELECT Id, Numero, Plataforma, Area, Responsable, FrecuenciaPago,
                       FechaInicio, ProximoPago,
                       DATEDIFF(day, CAST(GETDATE() AS DATE), ProximoPago) AS DiasRestantes,
                       CostoMXN,
                       CASE FrecuenciaPago
                           WHEN 'Mensual'    THEN CostoMXN * 12
                           WHEN 'Trimestral' THEN CostoMXN * 4
                           WHEN 'Semestral'  THEN CostoMXN * 2
                           WHEN 'Anual'      THEN CostoMXN
                           ELSE 0
                       END AS CostoAnualMXN,
                       Estado, Notas, CreadoEn, ActualizadoEn
                FROM dbo.Licencias
                WHERE Estado NOT IN ('Cancelada')
                  AND DATEDIFF(day, CAST(GETDATE() AS DATE), ProximoPago) <= 10
                ORDER BY DATEDIFF(day, CAST(GETDATE() AS DATE), ProximoPago)
                """;

            var list = new List<object>();
            await using var r = await cmd.ExecuteReaderAsync(ct);
            while (await r.ReadAsync(ct))
                list.Add(MapRow(r));
            return Ok(list);
        }
        catch (Exception ex) { logger.LogError(ex, "GetAlertas licencias"); return StatusCode(500, ex.Message); }
    }

    // ── GET /api/licencias/{id} ───────────────────────────────────────────────
    [HttpGet("{id:int}")]
    public async Task<IActionResult> GetById(int id, CancellationToken ct)
    {
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = """
                SELECT Id, Numero, Plataforma, Area, Responsable, FrecuenciaPago,
                       FechaInicio, ProximoPago,
                       DATEDIFF(day, CAST(GETDATE() AS DATE), ProximoPago) AS DiasRestantes,
                       CostoMXN,
                       CASE FrecuenciaPago
                           WHEN 'Mensual'    THEN CostoMXN * 12
                           WHEN 'Trimestral' THEN CostoMXN * 4
                           WHEN 'Semestral'  THEN CostoMXN * 2
                           WHEN 'Anual'      THEN CostoMXN
                           ELSE 0
                       END AS CostoAnualMXN,
                       Estado, Notas, CreadoEn, ActualizadoEn
                FROM dbo.Licencias WHERE Id = @Id
                """;
            cmd.Parameters.AddWithValue("@Id", id);
            await using var r = await cmd.ExecuteReaderAsync(ct);
            if (!await r.ReadAsync(ct)) return NotFound();
            return Ok(MapRow(r));
        }
        catch (Exception ex) { logger.LogError(ex, "GetById licencia {Id}", id); return StatusCode(500, ex.Message); }
    }

    // ── POST /api/licencias ───────────────────────────────────────────────────
    [HttpPost]
    [Authorize(Roles = "Admin")]
    public async Task<IActionResult> Create([FromBody] LicenciaRequest req, CancellationToken ct)
    {
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = """
                INSERT INTO dbo.Licencias
                    (Numero,Plataforma,Area,Responsable,FrecuenciaPago,FechaInicio,ProximoPago,CostoMXN,Estado,Notas)
                OUTPUT INSERTED.Id
                VALUES (@Numero,@Plataforma,@Area,@Responsable,@FrecuenciaPago,@FechaInicio,@ProximoPago,@CostoMXN,@Estado,@Notas)
                """;
            AddParams(cmd, req);
            var newId = (int)(await cmd.ExecuteScalarAsync(ct))!;
            return Ok(new { id = newId });
        }
        catch (Exception ex) { logger.LogError(ex, "Create licencia"); return StatusCode(500, ex.Message); }
    }

    // ── PUT /api/licencias/{id} ───────────────────────────────────────────────
    [HttpPut("{id:int}")]
    [Authorize(Roles = "Admin")]
    public async Task<IActionResult> Update(int id, [FromBody] LicenciaRequest req, CancellationToken ct)
    {
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = """
                UPDATE dbo.Licencias SET
                    Numero=@Numero, Plataforma=@Plataforma, Area=@Area, Responsable=@Responsable,
                    FrecuenciaPago=@FrecuenciaPago, FechaInicio=@FechaInicio, ProximoPago=@ProximoPago,
                    CostoMXN=@CostoMXN, Estado=@Estado, Notas=@Notas,
                    ActualizadoEn=GETUTCDATE()
                WHERE Id=@Id
                """;
            AddParams(cmd, req);
            cmd.Parameters.AddWithValue("@Id", id);
            int rows = await cmd.ExecuteNonQueryAsync(ct);
            if (rows == 0) return NotFound();
            return Ok();
        }
        catch (Exception ex) { logger.LogError(ex, "Update licencia {Id}", id); return StatusCode(500, ex.Message); }
    }

    // ── DELETE /api/licencias/{id} ────────────────────────────────────────────
    [HttpDelete("{id:int}")]
    [Authorize(Roles = "Admin")]
    public async Task<IActionResult> Delete(int id, CancellationToken ct)
    {
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = "DELETE FROM dbo.Licencias WHERE Id = @Id";
            cmd.Parameters.AddWithValue("@Id", id);
            int rows = await cmd.ExecuteNonQueryAsync(ct);
            if (rows == 0) return NotFound();
            return Ok();
        }
        catch (Exception ex) { logger.LogError(ex, "Delete licencia {Id}", id); return StatusCode(500, ex.Message); }
    }

    // ── PUT /api/licencias/actualizar-estados ─────────────────────────────────
    [HttpPut("actualizar-estados")]
    [Authorize(Roles = "Admin")]
    public async Task<IActionResult> ActualizarEstados(CancellationToken ct)
    {
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = """
                UPDATE dbo.Licencias SET
                    Estado = CASE
                        WHEN Estado = 'Cancelada' THEN 'Cancelada'
                        WHEN DATEDIFF(day, CAST(GETDATE() AS DATE), ProximoPago) < 0  THEN 'Vencida'
                        WHEN DATEDIFF(day, CAST(GETDATE() AS DATE), ProximoPago) <= 10 THEN 'Por vencer'
                        ELSE 'Activa'
                    END,
                    ActualizadoEn = GETUTCDATE()
                WHERE Estado != 'Cancelada'
                """;
            int rows = await cmd.ExecuteNonQueryAsync(ct);
            return Ok(new { actualizadas = rows });
        }
        catch (Exception ex) { logger.LogError(ex, "ActualizarEstados"); return StatusCode(500, ex.Message); }
    }

    // ── GET /api/licencias/exportar ───────────────────────────────────────────
    [HttpGet("exportar")]
    public async Task<IActionResult> ExportarExcel(CancellationToken ct)
    {
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = """
                SELECT Numero, Plataforma, Area, FrecuenciaPago,
                       FechaInicio, ProximoPago,
                       DATEDIFF(day, CAST(GETDATE() AS DATE), ProximoPago) AS DiasRestantes,
                       CostoMXN,
                       CASE FrecuenciaPago
                           WHEN 'Mensual'    THEN CostoMXN * 12
                           WHEN 'Trimestral' THEN CostoMXN * 4
                           WHEN 'Semestral'  THEN CostoMXN * 2
                           WHEN 'Anual'      THEN CostoMXN
                           ELSE 0
                       END AS CostoAnualMXN,
                       Estado, Notas
                FROM dbo.Licencias
                ORDER BY Numero
                """;

            var rows = new List<LicenciaRow>();
            await using var r = await cmd.ExecuteReaderAsync(ct);
            while (await r.ReadAsync(ct))
                rows.Add(new LicenciaRow(
                    r.GetInt32(r.GetOrdinal("Numero")),
                    r.GetString(r.GetOrdinal("Plataforma")),
                    r.GetString(r.GetOrdinal("Area")),
                    r.GetString(r.GetOrdinal("FrecuenciaPago")),
                    r.GetDateTime(r.GetOrdinal("FechaInicio")),
                    r.GetDateTime(r.GetOrdinal("ProximoPago")),
                    r.GetInt32(r.GetOrdinal("DiasRestantes")),
                    r.GetDecimal(r.GetOrdinal("CostoMXN")),
                    r.GetDecimal(r.GetOrdinal("CostoAnualMXN")),
                    r.GetString(r.GetOrdinal("Estado")),
                    r.IsDBNull(r.GetOrdinal("Notas")) ? null : r.GetString(r.GetOrdinal("Notas"))
                ));

            var bytes = BuildExcel(rows);
            string fname = $"iMET_Control_Licencias_{DateTime.Now:yyyy-MM-dd}.xlsx";
            return File(bytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fname);
        }
        catch (Exception ex) { logger.LogError(ex, "ExportarExcel licencias"); return StatusCode(500, ex.Message); }
    }

    // ── Helpers ───────────────────────────────────────────────────────────────

    private static object MapRow(SqlDataReader r) => new
    {
        id             = r.GetInt32(r.GetOrdinal("Id")),
        numero         = r.GetInt32(r.GetOrdinal("Numero")),
        plataforma     = r.GetString(r.GetOrdinal("Plataforma")),
        area           = r.GetString(r.GetOrdinal("Area")),
        responsable    = r.IsDBNull(r.GetOrdinal("Responsable")) ? null : r.GetString(r.GetOrdinal("Responsable")),
        frecuenciaPago = r.GetString(r.GetOrdinal("FrecuenciaPago")),
        fechaInicio    = r.GetDateTime(r.GetOrdinal("FechaInicio")).ToString("yyyy-MM-dd"),
        proximoPago    = r.GetDateTime(r.GetOrdinal("ProximoPago")).ToString("yyyy-MM-dd"),
        diasRestantes  = r.GetInt32(r.GetOrdinal("DiasRestantes")),
        costoMXN       = r.GetDecimal(r.GetOrdinal("CostoMXN")),
        costoAnualMXN  = r.GetDecimal(r.GetOrdinal("CostoAnualMXN")),
        estado         = r.GetString(r.GetOrdinal("Estado")),
        notas          = r.IsDBNull(r.GetOrdinal("Notas")) ? null : r.GetString(r.GetOrdinal("Notas")),
        creadoEn       = r.GetDateTime(r.GetOrdinal("CreadoEn")),
        actualizadoEn  = r.GetDateTime(r.GetOrdinal("ActualizadoEn")),
    };

    private static void AddParams(SqlCommand cmd, LicenciaRequest req)
    {
        cmd.Parameters.AddWithValue("@Numero",         req.Numero);
        cmd.Parameters.AddWithValue("@Plataforma",     req.Plataforma);
        cmd.Parameters.AddWithValue("@Area",           req.Area);
        cmd.Parameters.AddWithValue("@Responsable",    (object?)req.Responsable    ?? DBNull.Value);
        cmd.Parameters.AddWithValue("@FrecuenciaPago", req.FrecuenciaPago);
        cmd.Parameters.AddWithValue("@FechaInicio",    req.FechaInicio.ToString("yyyy-MM-dd"));
        cmd.Parameters.AddWithValue("@ProximoPago",    req.ProximoPago.ToString("yyyy-MM-dd"));
        cmd.Parameters.AddWithValue("@CostoMXN",       req.CostoMXN);
        cmd.Parameters.AddWithValue("@Estado",         req.Estado);
        cmd.Parameters.AddWithValue("@Notas",          (object?)req.Notas ?? DBNull.Value);
    }

    private static byte[] BuildExcel(List<LicenciaRow> rows)
    {
        // ── Hoja 1: Control de Licencias ─────────────────────────────────────
        var sheet1 = rows.Select(r => new Dictionary<string, object?>
        {
            ["#"]                      = r.Numero,
            ["Plataforma / Servicio"]  = r.Plataforma,
            ["Área"]                   = r.Area,
            ["Frecuencia de Pago"]     = r.FrecuenciaPago,
            ["Fecha de Inicio"]        = r.FechaInicio.ToString("dd/MM/yyyy"),
            ["Próximo Pago"]           = r.ProximoPago.ToString("dd/MM/yyyy"),
            ["Días Restantes"]         = r.DiasRestantes,
            ["Costo (MXN)"]            = r.CostoMXN,
            ["Costo Anual (MXN)"]      = r.CostoAnualMXN,
            ["Estado"]                 = r.Estado,
            ["Notas"]                  = r.Notas ?? "",
        }).ToList();

        // ── Hoja 2: Resumen por Área ──────────────────────────────────────────
        var sheet2 = rows
            .Where(r => r.Estado != "Cancelada")
            .GroupBy(r => r.Area)
            .OrderByDescending(g => g.Sum(r => r.CostoAnualMXN))
            .Select(g =>
            {
                decimal mensual = g.Sum(r => r.FrecuenciaPago switch
                {
                    "Mensual"    => r.CostoMXN,
                    "Trimestral" => r.CostoMXN / 3m,
                    "Semestral"  => r.CostoMXN / 6m,
                    "Anual"      => r.CostoMXN / 12m,
                    _            => 0m,
                });
                return new Dictionary<string, object?>
                {
                    ["Área"]                = g.Key,
                    ["Total Licencias"]     = g.Count(),
                    ["Activas"]             = g.Count(r => r.Estado == "Activa"),
                    ["Costo Mensual (MXN)"] = mensual,
                    ["Costo Anual (MXN)"]   = g.Sum(r => r.CostoAnualMXN),
                };
            }).ToList();

        // ── Hoja 3: Calendario de Pagos ───────────────────────────────────────
        string[] meses = ["ENE","FEB","MAR","ABR","MAY","JUN","JUL","AGO","SEP","OCT","NOV","DIC"];
        var sheet3 = rows
            .Where(r => r.Estado != "Cancelada")
            .OrderBy(r => r.Numero)
            .Select(r =>
            {
                var dict = new Dictionary<string, object?>
                {
                    ["Plataforma"] = r.Plataforma,
                    ["Frecuencia"] = r.FrecuenciaPago,
                };
                for (int m = 1; m <= 12; m++)
                {
                    bool paga = r.FrecuenciaPago switch
                    {
                        "Mensual"    => true,
                        "Trimestral" => m % 3 == (r.ProximoPago.Month % 3 == 0 ? 3 : r.ProximoPago.Month % 3),
                        "Semestral"  => m == r.ProximoPago.Month || m == ((r.ProximoPago.Month + 6 - 1) % 12) + 1,
                        "Anual"      => m == r.ProximoPago.Month,
                        _            => false,
                    };
                    dict[meses[m - 1]] = paga ? (object?)r.CostoMXN : null;
                }
                return dict;
            }).ToList();

        using var ms = new MemoryStream();
        MiniExcel.SaveAs(ms, new Dictionary<string, object>
        {
            ["Control de Licencias"] = sheet1,
            ["Resumen por Área"]     = sheet2,
            ["Calendario de Pagos"]  = sheet3,
        });
        return ms.ToArray();
    }

    private record LicenciaRow(
        int Numero, string Plataforma, string Area, string FrecuenciaPago,
        DateTime FechaInicio, DateTime ProximoPago, int DiasRestantes,
        decimal CostoMXN, decimal CostoAnualMXN, string Estado, string? Notas);
}

public record LicenciaRequest(
    int Numero,
    string Plataforma,
    string Area,
    string? Responsable,
    string FrecuenciaPago,
    DateTime FechaInicio,
    DateTime ProximoPago,
    decimal CostoMXN,
    string Estado,
    string? Notas
);
