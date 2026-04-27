using ClosedXML.Excel;
using Microsoft.AspNetCore.Authorization;
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
        using var wb = new XLWorkbook();

        // ── Hoja 1: Control de Licencias ─────────────────────────────────────
        var ws = wb.Worksheets.Add("Control de Licencias");

        // Colores del documento original
        var azulOscuro  = XLColor.FromHtml("#1A237E");
        var azulMedio   = XLColor.FromHtml("#3949AB");
        var verdeActiva = XLColor.FromHtml("#E8F5E9");
        var verdeTxt    = XLColor.FromHtml("#2E7D32");
        var amarilloVencer  = XLColor.FromHtml("#FFF9C4");
        var amarilloTxt     = XLColor.FromHtml("#F57F17");
        var rojoVencida = XLColor.FromHtml("#FFEBEE");
        var rojoTxt     = XLColor.FromHtml("#C62828");
        var grisCancel  = XLColor.FromHtml("#F5F5F5");
        var grisTxt     = XLColor.FromHtml("#757575");

        // Fila 1: título
        ws.Range("A1:L1").Merge();
        ws.Cell("A1").Value = "iMET — CONTROL DE LICENCIAS Y PAGOS DE PLATAFORMAS";
        ws.Cell("A1").Style
            .Font.SetBold(true).Font.SetFontSize(14).Font.SetFontColor(XLColor.White)
            .Fill.SetBackgroundColor(azulOscuro)
            .Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center)
            .Alignment.SetVertical(XLAlignmentVerticalValues.Center);
        ws.Row(1).Height = 30;

        // Fila 2: subtítulo
        ws.Range("A2:L2").Merge();
        ws.Cell("A2").Value = $"Coordinación de Tecnologías  |  Generado: {DateTime.Now:dd/MM/yyyy}";
        ws.Cell("A2").Style
            .Font.SetItalic(true).Font.SetFontColor(XLColor.White)
            .Fill.SetBackgroundColor(azulMedio)
            .Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
        ws.Row(2).Height = 20;

        // Fila 3: encabezados
        string[] headers = ["#", "Plataforma / Servicio", "Área", "Frecuencia de Pago",
                             "Fecha de Inicio", "Próximo Pago / Renovación", "Días Restantes",
                             "Costo (MXN)", "Costo Anual (MXN)", "Estado", "Notas / Observaciones"];
        for (int i = 0; i < headers.Length; i++)
        {
            var cell = ws.Cell(3, i + 1);
            cell.Value = headers[i];
            cell.Style
                .Font.SetBold(true).Font.SetFontColor(XLColor.White)
                .Fill.SetBackgroundColor(azulOscuro)
                .Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center)
                .Alignment.SetVertical(XLAlignmentVerticalValues.Center)
                .Border.SetOutsideBorder(XLBorderStyleValues.Thin)
                .Border.SetOutsideBorderColor(XLColor.White);
        }
        ws.Row(3).Height = 22;

        // Datos
        for (int i = 0; i < rows.Count; i++)
        {
            var row = rows[i];
            int excelRow = i + 4;

            ws.Cell(excelRow, 1).Value  = row.Numero;
            ws.Cell(excelRow, 2).Value  = row.Plataforma;
            ws.Cell(excelRow, 3).Value  = row.Area;
            ws.Cell(excelRow, 4).Value  = row.FrecuenciaPago;
            ws.Cell(excelRow, 5).Value  = row.FechaInicio;
            ws.Cell(excelRow, 6).Value  = row.ProximoPago;
            ws.Cell(excelRow, 7).Value  = row.DiasRestantes;
            ws.Cell(excelRow, 8).Value  = (double)row.CostoMXN;
            ws.Cell(excelRow, 9).Value  = (double)row.CostoAnualMXN;
            ws.Cell(excelRow, 10).Value = row.Estado;
            ws.Cell(excelRow, 11).Value = row.Notas ?? "";

            // Formato de fechas
            ws.Cell(excelRow, 5).Style.DateFormat.Format = "dd/MM/yyyy";
            ws.Cell(excelRow, 6).Style.DateFormat.Format = "dd/MM/yyyy";

            // Formato de moneda
            ws.Cell(excelRow, 8).Style.NumberFormat.Format = "$#,##0.00";
            ws.Cell(excelRow, 9).Style.NumberFormat.Format = "$#,##0.00";

            // Color de fila según estado
            var (bgColor, txtColor) = row.Estado switch
            {
                "Activa"     => (verdeActiva,    verdeTxt),
                "Por vencer" => (amarilloVencer, amarilloTxt),
                "Vencida"    => (rojoVencida,    rojoTxt),
                _            => (grisCancel,     grisTxt),
            };

            var range = ws.Range(excelRow, 1, excelRow, 11);
            range.Style.Fill.SetBackgroundColor(bgColor);
            range.Style.Font.SetFontColor(txtColor);
            range.Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
            range.Style.Border.SetOutsideBorderColor(XLColor.FromHtml("#BDBDBD"));

            // Estado en negrita
            ws.Cell(excelRow, 10).Style.Font.SetBold(true);
        }

        // Fila total
        int totalRow = rows.Count + 4;
        ws.Range(totalRow, 1, totalRow, 7).Merge();
        ws.Cell(totalRow, 1).Value = "TOTAL ANUAL ESTIMADO";
        ws.Cell(totalRow, 1).Style.Font.SetBold(true).Alignment.SetHorizontal(XLAlignmentHorizontalValues.Right);

        var activasRows = rows.Where(r => r.Estado != "Cancelada").ToList();
        ws.Cell(totalRow, 8).Value = (double)activasRows.Sum(r => r.CostoMXN);
        ws.Cell(totalRow, 9).Value = (double)activasRows.Sum(r => r.CostoAnualMXN);
        ws.Cell(totalRow, 8).Style.NumberFormat.Format = "$#,##0.00";
        ws.Cell(totalRow, 9).Style.NumberFormat.Format = "$#,##0.00";
        ws.Range(totalRow, 1, totalRow, 11).Style
            .Font.SetBold(true).Fill.SetBackgroundColor(azulOscuro).Font.SetFontColor(XLColor.White)
            .Border.SetOutsideBorder(XLBorderStyleValues.Medium);

        // Anchos de columna
        ws.Column(1).Width  = 6;
        ws.Column(2).Width  = 28;
        ws.Column(3).Width  = 16;
        ws.Column(4).Width  = 18;
        ws.Column(5).Width  = 16;
        ws.Column(6).Width  = 22;
        ws.Column(7).Width  = 16;
        ws.Column(8).Width  = 16;
        ws.Column(9).Width  = 18;
        ws.Column(10).Width = 14;
        ws.Column(11).Width = 32;

        // Inmovilizar encabezados
        ws.SheetView.FreezeRows(3);

        // ── Hoja 2: Resumen por Área ──────────────────────────────────────────
        var ws2 = wb.Worksheets.Add("Resumen por Área");
        ws2.Range("A1:E1").Merge();
        ws2.Cell("A1").Value = "iMET — RESUMEN EJECUTIVO DE LICENCIAS";
        ws2.Cell("A1").Style.Font.SetBold(true).Font.SetFontSize(13).Font.SetFontColor(XLColor.White)
            .Fill.SetBackgroundColor(azulOscuro)
            .Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
        ws2.Row(1).Height = 28;

        string[] h2 = ["Área", "Total Licencias", "Activas", "Costo Mensual (MXN)", "Costo Anual (MXN)"];
        for (int i = 0; i < h2.Length; i++)
        {
            var c = ws2.Cell(2, i + 1);
            c.Value = h2[i];
            c.Style.Font.SetBold(true).Font.SetFontColor(XLColor.White)
                .Fill.SetBackgroundColor(azulMedio)
                .Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
        }

        var byArea = rows.Where(r => r.Estado != "Cancelada")
            .GroupBy(r => r.Area)
            .OrderByDescending(g => g.Sum(r => r.CostoAnualMXN));
        int ar = 3;
        foreach (var g in byArea)
        {
            decimal mensual = g.Sum(r => r.FrecuenciaPago switch
            {
                "Mensual"    => r.CostoMXN,
                "Trimestral" => r.CostoMXN / 3m,
                "Semestral"  => r.CostoMXN / 6m,
                "Anual"      => r.CostoMXN / 12m,
                _            => 0m,
            });
            ws2.Cell(ar, 1).Value = g.Key;
            ws2.Cell(ar, 2).Value = g.Count();
            ws2.Cell(ar, 3).Value = g.Count(r => r.Estado == "Activa");
            ws2.Cell(ar, 4).Value = (double)mensual;
            ws2.Cell(ar, 5).Value = (double)g.Sum(r => r.CostoAnualMXN);
            ws2.Cell(ar, 4).Style.NumberFormat.Format = "$#,##0.00";
            ws2.Cell(ar, 5).Style.NumberFormat.Format = "$#,##0.00";
            ar++;
        }
        ws2.Column(1).Width = 24; ws2.Column(2).Width = 16;
        ws2.Column(3).Width = 12; ws2.Column(4).Width = 22; ws2.Column(5).Width = 22;

        // ── Hoja 3: Calendario de Pagos ───────────────────────────────────────
        var ws3 = wb.Worksheets.Add("Calendario de Pagos");
        ws3.Range("A1:N1").Merge();
        ws3.Cell("A1").Value = "iMET — CALENDARIO ANUAL DE PAGOS DE LICENCIAS";
        ws3.Cell("A1").Style.Font.SetBold(true).Font.SetFontSize(13).Font.SetFontColor(XLColor.White)
            .Fill.SetBackgroundColor(azulOscuro)
            .Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);

        string[] meses = ["ENE","FEB","MAR","ABR","MAY","JUN","JUL","AGO","SEP","OCT","NOV","DIC"];
        ws3.Cell(2, 1).Value = "Plataforma";
        ws3.Cell(2, 2).Value = "Frecuencia";
        for (int m = 0; m < 12; m++) ws3.Cell(2, m + 3).Value = meses[m];
        ws3.Row(2).Style.Font.SetBold(true).Font.SetFontColor(XLColor.White).Fill.SetBackgroundColor(azulMedio);

        var activasCalendario = rows.Where(r => r.Estado != "Cancelada").OrderBy(r => r.Numero).ToList();
        for (int i = 0; i < activasCalendario.Count; i++)
        {
            var row = activasCalendario[i];
            int er = i + 3;
            ws3.Cell(er, 1).Value = row.Plataforma;
            ws3.Cell(er, 2).Value = row.FrecuenciaPago;

            for (int m = 1; m <= 12; m++)
            {
                bool paga = row.FrecuenciaPago switch
                {
                    "Mensual"    => true,
                    "Trimestral" => m % 3 == (row.ProximoPago.Month % 3 == 0 ? 3 : row.ProximoPago.Month % 3),
                    "Semestral"  => m == row.ProximoPago.Month || m == ((row.ProximoPago.Month + 6 - 1) % 12) + 1,
                    "Anual"      => m == row.ProximoPago.Month,
                    _            => false,
                };
                if (paga)
                {
                    ws3.Cell(er, m + 2).Value = "💳";
                    ws3.Cell(er, m + 2).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center)
                        .Fill.SetBackgroundColor(verdeActiva);
                }
            }
        }
        ws3.Column(1).Width = 26; ws3.Column(2).Width = 14;
        for (int m = 3; m <= 14; m++) ws3.Column(m).Width = 7;

        using var ms = new MemoryStream();
        wb.SaveAs(ms);
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
