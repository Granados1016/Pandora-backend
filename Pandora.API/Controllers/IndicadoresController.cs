using ClosedXML.Excel;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;

namespace Pandora.API.Controllers;

[ApiController]
[Route("api/indicadores")]
[Authorize]
public class IndicadoresController(
    IConfiguration config,
    ILogger<IndicadoresController> logger) : ControllerBase
{
    private SqlConnection Conn() => new(config.GetConnectionString("PandoraDb"));

    // ── GET /api/indicadores ──────────────────────────────────────────────────
    /// <summary>
    /// Devuelve un snapshot consolidado de KPIs del sistema.
    /// </summary>
    [HttpGet]
    public async Task<IActionResult> GetAll(CancellationToken ct)
    {
        await using var conn = Conn();
        await conn.OpenAsync(ct);

        // ── Usuarios ──────────────────────────────────────────────────────────
        int usersTotal = 0, usersActive = 0;
        await using (var cmd = conn.CreateCommand())
        {
            cmd.CommandText = "SELECT COUNT(*), SUM(CASE WHEN IsActive = 1 THEN 1 ELSE 0 END) FROM dbo.AppUsers";
            await using var r = await cmd.ExecuteReaderAsync(ct);
            if (await r.ReadAsync(ct)) { usersTotal = r.IsDBNull(0) ? 0 : r.GetInt32(0); usersActive = r.IsDBNull(1) ? 0 : r.GetInt32(1); }
        }

        // ── Mail+ ─────────────────────────────────────────────────────────────
        int mailSent = 0, mailThisMonth = 0;
        var mailMonthly = new List<object>();
        try
        {
            await using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = """
                    SELECT
                        COUNT(*) AS Total,
                        SUM(CASE WHEN MONTH(SentAt) = MONTH(GETUTCDATE()) AND YEAR(SentAt) = YEAR(GETUTCDATE()) THEN 1 ELSE 0 END) AS ThisMonth
                    FROM dbo.EmailCampaigns
                    WHERE IsDeleted = 0 AND Status = 'Sent'
                    """;
                await using var r = await cmd.ExecuteReaderAsync(ct);
                if (await r.ReadAsync(ct))
                {
                    mailSent      = r.IsDBNull(0) ? 0 : r.GetInt32(0);
                    mailThisMonth = r.IsDBNull(1) ? 0 : r.GetInt32(1);
                }
            }

            // Campañas por mes (últimos 6 meses)
            await using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = """
                    SELECT
                        FORMAT(SentAt, 'MMM yy', 'es-MX') AS Month,
                        COUNT(*) AS Cnt
                    FROM dbo.EmailCampaigns
                    WHERE IsDeleted = 0 AND Status = 'Sent'
                      AND SentAt >= DATEADD(MONTH, -5, DATEFROMPARTS(YEAR(GETUTCDATE()), MONTH(GETUTCDATE()), 1))
                    GROUP BY FORMAT(SentAt, 'MMM yy', 'es-MX'),
                             YEAR(SentAt) * 100 + MONTH(SentAt)
                    ORDER BY YEAR(SentAt) * 100 + MONTH(SentAt)
                    """;
                await using var r = await cmd.ExecuteReaderAsync(ct);
                while (await r.ReadAsync(ct))
                    mailMonthly.Add(new { month = r.GetString(0), count = r.GetInt32(1) });
            }
        }
        catch { /* tabla puede no existir en todos los entornos */ }

        // ── Inventario ────────────────────────────────────────────────────────
        int invItems = 0, invTypes = 0;
        var invByType = new List<object>();
        try
        {
            await using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = """
                    SELECT COUNT(*) FROM dbo.InventoryItems WHERE IsDeleted = 0
                    """;
                invItems = (int)(await cmd.ExecuteScalarAsync(ct) ?? 0);
            }
            await using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = "SELECT COUNT(*) FROM dbo.InventoryTypes WHERE IsDeleted = 0";
                invTypes = (int)(await cmd.ExecuteScalarAsync(ct) ?? 0);
            }
            await using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = """
                    SELECT TOP 8 t.Name, COUNT(i.Id) AS Cnt
                    FROM dbo.InventoryTypes t
                    LEFT JOIN dbo.InventoryItems i ON i.TypeId = t.Id AND i.IsDeleted = 0
                    WHERE t.IsDeleted = 0
                    GROUP BY t.Name
                    ORDER BY Cnt DESC
                    """;
                await using var r = await cmd.ExecuteReaderAsync(ct);
                while (await r.ReadAsync(ct))
                    invByType.Add(new { typeName = r.GetString(0), count = r.GetInt32(1) });
            }
        }
        catch { }

        // ── Tickets ───────────────────────────────────────────────────────────
        int ticketsTotal = 0, ticketsOpen = 0, ticketsResolved = 0;
        var ticketsByStatus = new List<object>();
        try
        {
            await using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = """
                    SELECT Status, COUNT(*) AS Cnt
                    FROM dbo.Tickets
                    WHERE IsDeleted = 0
                    GROUP BY Status
                    """;
                await using var r = await cmd.ExecuteReaderAsync(ct);
                while (await r.ReadAsync(ct))
                {
                    var status = r.GetString(0);
                    var cnt    = r.GetInt32(1);
                    ticketsTotal += cnt;
                    if (status is "Abierto" or "Open" or "Nuevo")   ticketsOpen     += cnt;
                    if (status is "Resuelto" or "Closed" or "Cerrado") ticketsResolved += cnt;
                    ticketsByStatus.Add(new { status, count = cnt });
                }
            }
        }
        catch { }

        // ── Licencias ─────────────────────────────────────────────────────────
        int licActive = 0, licExpiringSoon = 0;
        var licByEstado = new List<object>();
        try
        {
            await using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = """
                    SELECT Estado, COUNT(*) AS Cnt
                    FROM dbo.Licencias
                    WHERE IsDeleted = 0
                    GROUP BY Estado
                    """;
                await using var r = await cmd.ExecuteReaderAsync(ct);
                while (await r.ReadAsync(ct))
                {
                    var estado = r.GetString(0);
                    var cnt    = r.GetInt32(1);
                    if (estado is "Activa" or "Vigente") licActive += cnt;
                    licByEstado.Add(new { estado, count = cnt });
                }
            }
            await using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = """
                    SELECT COUNT(*) FROM dbo.Licencias
                    WHERE IsDeleted = 0 AND Estado = 'Activa'
                      AND FechaVencimiento IS NOT NULL
                      AND FechaVencimiento <= DATEADD(DAY, 30, GETUTCDATE())
                      AND FechaVencimiento >= GETUTCDATE()
                    """;
                licExpiringSoon = (int)(await cmd.ExecuteScalarAsync(ct) ?? 0);
            }
        }
        catch { }

        // ── Calendario ────────────────────────────────────────────────────────
        int calThisMonth = 0, calUpcoming = 0;
        try
        {
            await using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = """
                    SELECT
                        SUM(CASE WHEN MONTH(StartTime) = MONTH(GETUTCDATE()) AND YEAR(StartTime) = YEAR(GETUTCDATE()) THEN 1 ELSE 0 END),
                        SUM(CASE WHEN StartTime >= GETUTCDATE() THEN 1 ELSE 0 END)
                    FROM dbo.Reservations
                    WHERE IsDeleted = 0 AND Status <> 'Cancelada'
                    """;
                await using var r = await cmd.ExecuteReaderAsync(ct);
                if (await r.ReadAsync(ct))
                {
                    calThisMonth = r.IsDBNull(0) ? 0 : r.GetInt32(0);
                    calUpcoming  = r.IsDBNull(1) ? 0 : r.GetInt32(1);
                }
            }
        }
        catch { }

        logger.LogInformation("Indicadores solicitados por {User}", User.Identity?.Name);

        return Ok(new
        {
            users = new
            {
                total  = usersTotal,
                active = usersActive,
            },
            mail = new
            {
                sent       = mailSent,
                thisMonth  = mailThisMonth,
                trend      = mailThisMonth > 0 ? 1 : 0,
                trendLabel = mailThisMonth > 0 ? $"{mailThisMonth} este mes" : null,
                monthlySent = mailMonthly,
            },
            inventory = new
            {
                totalItems = invItems,
                types      = invTypes,
                byType     = invByType,
            },
            tickets = new
            {
                total    = ticketsTotal,
                open     = ticketsOpen,
                resolved = ticketsResolved,
                byStatus = ticketsByStatus,
            },
            licencias = new
            {
                active       = licActive,
                expiringSoon = licExpiringSoon,
                byEstado     = licByEstado,
            },
            calendar = new
            {
                thisMonth = calThisMonth,
                upcoming  = calUpcoming,
            },
        });
    }

    // ── GET /api/indicadores/export ───────────────────────────────────────────
    /// <summary>Exporta el snapshot de indicadores a un archivo Excel.</summary>
    [HttpGet("export")]
    public async Task<IActionResult> Export(CancellationToken ct)
    {
        // Reutilizamos la misma lógica de consulta llamando al propio método
        // (evitamos duplicar SQL — deserializamos el OkObjectResult)
        var kpiResult = await GetAll(ct) as OkObjectResult;
        if (kpiResult?.Value is null) return StatusCode(500, "Error generando indicadores.");

        // Serializar/deserializar para obtener un objeto dinámico
        var json    = System.Text.Json.JsonSerializer.Serialize(kpiResult.Value);
        using var doc = System.Text.Json.JsonDocument.Parse(json);
        var root    = doc.RootElement;

        var now     = DateTime.Now;
        using var wb = new XLWorkbook();

        // ── Hoja 1: Resumen general ──────────────────────────────────────────
        var ws = wb.AddWorksheet("Resumen");
        ws.Cell(1, 1).Value = "Indicadores Pandora";
        ws.Cell(1, 1).Style.Font.Bold = true;
        ws.Cell(1, 1).Style.Font.FontSize = 16;
        ws.Cell(2, 1).Value = $"Generado el {now:dd/MM/yyyy HH:mm}";
        ws.Cell(2, 1).Style.Font.Italic = true;

        int row = 4;
        ws.Cell(row, 1).Value = "Módulo";          StyleHeader(ws.Cell(row, 1));
        ws.Cell(row, 2).Value = "Indicador";        StyleHeader(ws.Cell(row, 2));
        ws.Cell(row, 3).Value = "Valor";            StyleHeader(ws.Cell(row, 3));

        row++;
        void AddRow(string mod, string ind, object? val) { ws.Cell(row, 1).Value = mod; ws.Cell(row, 2).Value = ind; ws.Cell(row, 3).Value = val?.ToString() ?? "-"; row++; }

        var users = root.GetPropertyOrDefault("users");
        AddRow("Usuarios", "Total",   users?.GetInt("total"));
        AddRow("Usuarios", "Activos", users?.GetInt("active"));

        var mail = root.GetPropertyOrDefault("mail");
        AddRow("Mail+", "Campañas enviadas (total)", mail?.GetInt("sent"));
        AddRow("Mail+", "Este mes",                  mail?.GetInt("thisMonth"));

        var inv = root.GetPropertyOrDefault("inventory");
        AddRow("Inventario", "Artículos totales", inv?.GetInt("totalItems"));
        AddRow("Inventario", "Categorías",        inv?.GetInt("types"));

        var tickets = root.GetPropertyOrDefault("tickets");
        AddRow("Tickets", "Total",    tickets?.GetInt("total"));
        AddRow("Tickets", "Abiertos", tickets?.GetInt("open"));
        AddRow("Tickets", "Resueltos",tickets?.GetInt("resolved"));

        var lic = root.GetPropertyOrDefault("licencias");
        AddRow("Licencias", "Activas",        lic?.GetInt("active"));
        AddRow("Licencias", "Por vencer (30d)",lic?.GetInt("expiringSoon"));

        var cal = root.GetPropertyOrDefault("calendar");
        AddRow("Calendario", "Reservas este mes", cal?.GetInt("thisMonth"));
        AddRow("Calendario", "Próximas",           cal?.GetInt("upcoming"));

        ws.Columns(1, 3).AdjustToContents();

        // ── Hoja 2: Campañas por mes ─────────────────────────────────────────
        if (mail.HasValue && mail.Value.TryGetProperty("monthlySent", out var ms2)
            && ms2.ValueKind == System.Text.Json.JsonValueKind.Array)
        {
            var wsM = wb.AddWorksheet("Campañas por mes");
            wsM.Cell(1, 1).Value = "Mes";      StyleHeader(wsM.Cell(1, 1));
            wsM.Cell(1, 2).Value = "Enviadas"; StyleHeader(wsM.Cell(1, 2));
            int r = 2;
            foreach (var item in ms2.EnumerateArray())
            {
                wsM.Cell(r, 1).Value = item.TryGetProperty("month", out var m)  ? m.GetString() : "";
                wsM.Cell(r, 2).Value = item.TryGetProperty("count", out var c)  ? c.GetInt32()  : 0;
                r++;
            }
            wsM.Columns(1, 2).AdjustToContents();
        }

        // ── Hoja 3: Inventario por categoría ────────────────────────────────
        if (inv.HasValue && inv.Value.TryGetProperty("byType", out var byType)
            && byType.ValueKind == System.Text.Json.JsonValueKind.Array)
        {
            var wsI = wb.AddWorksheet("Inventario por categoría");
            wsI.Cell(1, 1).Value = "Categoría"; StyleHeader(wsI.Cell(1, 1));
            wsI.Cell(1, 2).Value = "Artículos"; StyleHeader(wsI.Cell(1, 2));
            int r = 2;
            foreach (var item in byType.EnumerateArray())
            {
                wsI.Cell(r, 1).Value = item.TryGetProperty("typeName", out var tn) ? tn.GetString() : "";
                wsI.Cell(r, 2).Value = item.TryGetProperty("count",    out var c2) ? c2.GetInt32()  : 0;
                r++;
            }
            wsI.Columns(1, 2).AdjustToContents();
        }

        // ── Hoja 4: Tickets por estado ───────────────────────────────────────
        if (tickets.HasValue && tickets.Value.TryGetProperty("byStatus", out var byStatus)
            && byStatus.ValueKind == System.Text.Json.JsonValueKind.Array)
        {
            var wsT = wb.AddWorksheet("Tickets por estado");
            wsT.Cell(1, 1).Value = "Estado";   StyleHeader(wsT.Cell(1, 1));
            wsT.Cell(1, 2).Value = "Cantidad"; StyleHeader(wsT.Cell(1, 2));
            int r = 2;
            foreach (var item in byStatus.EnumerateArray())
            {
                wsT.Cell(r, 1).Value = item.TryGetProperty("status", out var s)  ? s.GetString() : "";
                wsT.Cell(r, 2).Value = item.TryGetProperty("count",  out var c3) ? c3.GetInt32() : 0;
                r++;
            }
            wsT.Columns(1, 2).AdjustToContents();
        }

        // ── Hoja 5: Licencias por estado ─────────────────────────────────────
        if (lic.HasValue && lic.Value.TryGetProperty("byEstado", out var byEstado)
            && byEstado.ValueKind == System.Text.Json.JsonValueKind.Array)
        {
            var wsL = wb.AddWorksheet("Licencias por estado");
            wsL.Cell(1, 1).Value = "Estado";   StyleHeader(wsL.Cell(1, 1));
            wsL.Cell(1, 2).Value = "Cantidad"; StyleHeader(wsL.Cell(1, 2));
            int r = 2;
            foreach (var item in byEstado.EnumerateArray())
            {
                wsL.Cell(r, 1).Value = item.TryGetProperty("estado", out var e)  ? e.GetString() : "";
                wsL.Cell(r, 2).Value = item.TryGetProperty("count",  out var c4) ? c4.GetInt32() : 0;
                r++;
            }
            wsL.Columns(1, 2).AdjustToContents();
        }

        // ── Serializar y devolver ────────────────────────────────────────────
        using var stream = new MemoryStream();
        wb.SaveAs(stream);
        var bytes    = stream.ToArray();
        var filename = $"Indicadores_Pandora_{now:yyyyMMdd_HHmm}.xlsx";
        logger.LogInformation("Export indicadores por {User}", User.Identity?.Name);
        return File(bytes,
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            filename);
    }

    // ── Helpers privados ─────────────────────────────────────────────────────
    private static void StyleHeader(IXLCell cell)
    {
        cell.Style.Font.Bold = true;
        cell.Style.Fill.BackgroundColor = XLColor.FromHtml("#E8EAED");
    }
}

// ── Extensiones para JsonElement nullable ─────────────────────────────────────
internal static class JsonElementExtensions
{
    public static System.Text.Json.JsonElement? GetPropertyOrDefault(
        this System.Text.Json.JsonElement el, string name) =>
        el.TryGetProperty(name, out var v) ? v : null;

    public static System.Text.Json.JsonElement? GetPropertyOrDefault(
        this System.Text.Json.JsonElement? el, string name) =>
        el.HasValue && el.Value.TryGetProperty(name, out var v) ? v : null;

    public static int? GetInt(this System.Text.Json.JsonElement? el, string name) =>
        el.HasValue && el.Value.TryGetProperty(name, out var v) && v.ValueKind == System.Text.Json.JsonValueKind.Number
            ? v.GetInt32() : null;

    public static int? GetInt(this System.Text.Json.JsonElement el, string name) =>
        el.TryGetProperty(name, out var v) && v.ValueKind == System.Text.Json.JsonValueKind.Number
            ? v.GetInt32() : null;
}
