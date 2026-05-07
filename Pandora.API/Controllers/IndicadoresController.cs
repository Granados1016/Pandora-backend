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
}
