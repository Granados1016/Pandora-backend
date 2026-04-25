using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;

namespace Pandora.API.Controllers;

[ApiController]
[Route("api/reports")]
[Authorize]
public class ReportsController(IConfiguration config, ILogger<ReportsController> logger) : ControllerBase
{
    private SqlConnection Conn() => new(config.GetConnectionString("PandoraDb"));

    // ════════════════════════════════════════════════════════════════════════
    //  MAIL+
    // ════════════════════════════════════════════════════════════════════════

    [HttpGet("mail")]
    public async Task<IActionResult> GetMailReport(CancellationToken ct)
    {
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);

            // ── KPIs ─────────────────────────────────────────────────────────
            await using var cmdKpi = conn.CreateCommand();
            cmdKpi.CommandText = """
                SELECT
                    COUNT(*)                                          AS TotalCampaigns,
                    ISNULL(SUM(TotalRecipients), 0)                   AS TotalRecipients,
                    ISNULL(SUM(SentCount), 0)                         AS TotalSent,
                    ISNULL(SUM(FailedCount), 0)                       AS TotalFailed
                FROM dbo.EmailCampaigns
                WHERE IsDeleted = 0
                """;
            int totalCampaigns = 0, totalRecipients = 0, totalSent = 0, totalFailed = 0;
            await using (var r = await cmdKpi.ExecuteReaderAsync(ct))
            {
                if (await r.ReadAsync(ct))
                {
                    totalCampaigns  = r.GetInt32(0);
                    totalRecipients = r.GetInt32(1);
                    totalSent       = r.GetInt32(2);
                    totalFailed     = r.GetInt32(3);
                }
            }
            double successRate = totalSent + totalFailed > 0
                ? Math.Round((double)totalSent / (totalSent + totalFailed) * 100, 1)
                : 0;

            // ── Tendencia mensual — últimos 12 meses ──────────────────────────
            await using var cmdTrend = conn.CreateCommand();
            cmdTrend.CommandText = """
                SELECT
                    FORMAT(SentAt, 'MMM yy', 'es-MX')  AS Label,
                    YEAR(SentAt)                         AS Yr,
                    MONTH(SentAt)                        AS Mo,
                    ISNULL(SUM(SentCount),   0)          AS Sent,
                    ISNULL(SUM(FailedCount), 0)          AS Failed
                FROM dbo.EmailCampaigns
                WHERE IsDeleted = 0
                  AND SentAt IS NOT NULL
                  AND SentAt >= DATEADD(MONTH, -11, DATEFROMPARTS(YEAR(GETUTCDATE()), MONTH(GETUTCDATE()), 1))
                GROUP BY FORMAT(SentAt, 'MMM yy', 'es-MX'), YEAR(SentAt), MONTH(SentAt)
                ORDER BY Yr, Mo
                """;
            var monthlyTrend = new List<object>();
            await using (var r = await cmdTrend.ExecuteReaderAsync(ct))
                while (await r.ReadAsync(ct))
                    monthlyTrend.Add(new { label = r.GetString(0), sent = r.GetInt32(3), failed = r.GetInt32(4) });

            // ── Por tipo de programa ──────────────────────────────────────────
            await using var cmdType = conn.CreateCommand();
            cmdType.CommandText = """
                SELECT
                    ISNULL(NULLIF(ProgramType,''), 'Sin especificar') AS Name,
                    COUNT(*) AS Value
                FROM dbo.EmailCampaigns
                WHERE IsDeleted = 0
                GROUP BY ISNULL(NULLIF(ProgramType,''), 'Sin especificar')
                ORDER BY Value DESC
                """;
            var byProgramType = new List<object>();
            await using (var r = await cmdType.ExecuteReaderAsync(ct))
                while (await r.ReadAsync(ct))
                    byProgramType.Add(new { name = r.GetString(0), value = r.GetInt32(1) });

            // ── Top campañas ──────────────────────────────────────────────────
            await using var cmdTop = conn.CreateCommand();
            cmdTop.CommandText = """
                SELECT TOP 8
                    Name,
                    ISNULL(SentCount,   0) AS Sent,
                    ISNULL(FailedCount, 0) AS Failed
                FROM dbo.EmailCampaigns
                WHERE IsDeleted = 0
                ORDER BY TotalRecipients DESC
                """;
            var topCampaigns = new List<object>();
            await using (var r = await cmdTop.ExecuteReaderAsync(ct))
                while (await r.ReadAsync(ct))
                    topCampaigns.Add(new { name = r.GetString(0), sent = r.GetInt32(1), failed = r.GetInt32(2) });

            return Ok(new
            {
                totalCampaigns,
                totalRecipients,
                totalSent,
                totalFailed,
                successRate,
                monthlyTrend,
                byProgramType,
                topCampaigns,
            });
        }
        catch (Exception ex) { logger.LogError(ex, "GetMailReport"); return StatusCode(500, ex.Message); }
    }

    // ════════════════════════════════════════════════════════════════════════
    //  INVENTARIO
    // ════════════════════════════════════════════════════════════════════════

    [HttpGet("inventory")]
    public async Task<IActionResult> GetInventoryReport(CancellationToken ct)
    {
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);

            // ── KPIs ─────────────────────────────────────────────────────────
            await using var cmdKpi = conn.CreateCommand();
            cmdKpi.CommandText = """
                SELECT
                    COUNT(*)                                                             AS Total,
                    SUM(CASE WHEN Status = 'Activo'          THEN 1 ELSE 0 END)         AS Active,
                    SUM(CASE WHEN Status = 'Mantenimiento'   THEN 1 ELSE 0 END)         AS Maintenance,
                    SUM(CASE WHEN Status = 'Dado de baja'    THEN 1 ELSE 0 END)         AS Decommissioned,
                    SUM(CASE WHEN Status = 'Almacén'         THEN 1 ELSE 0 END)         AS InStorage
                FROM dbo.InventoryItems
                WHERE IsActive = 1
                """;
            int total = 0, active = 0, maintenance = 0, decommissioned = 0, inStorage = 0;
            await using (var r = await cmdKpi.ExecuteReaderAsync(ct))
            {
                if (await r.ReadAsync(ct))
                {
                    total          = r.GetInt32(0);
                    active         = r.GetInt32(1);
                    maintenance    = r.GetInt32(2);
                    decommissioned = r.GetInt32(3);
                    inStorage      = r.GetInt32(4);
                }
            }

            // ── Por categoría (tipo) ──────────────────────────────────────────
            await using var cmdType = conn.CreateCommand();
            cmdType.CommandText = """
                SELECT
                    ISNULL(t.Name, 'Sin categoría') AS Name,
                    COUNT(*) AS Value
                FROM dbo.InventoryItems i
                LEFT JOIN dbo.InventoryTypes t ON t.Id = i.TypeId
                WHERE i.IsActive = 1
                GROUP BY ISNULL(t.Name, 'Sin categoría')
                ORDER BY Value DESC
                """;
            var byType = new List<object>();
            await using (var r = await cmdType.ExecuteReaderAsync(ct))
                while (await r.ReadAsync(ct))
                    byType.Add(new { name = r.GetString(0), value = r.GetInt32(1) });

            // ── Por estado ────────────────────────────────────────────────────
            await using var cmdStatus = conn.CreateCommand();
            cmdStatus.CommandText = """
                SELECT
                    ISNULL(NULLIF(Status,''), 'Sin estado') AS Name,
                    COUNT(*) AS Value
                FROM dbo.InventoryItems
                WHERE IsActive = 1
                GROUP BY ISNULL(NULLIF(Status,''), 'Sin estado')
                ORDER BY Value DESC
                """;
            var byStatus = new List<object>();
            await using (var r = await cmdStatus.ExecuteReaderAsync(ct))
                while (await r.ReadAsync(ct))
                    byStatus.Add(new { name = r.GetString(0), value = r.GetInt32(1) });

            // ── Transferencias por mes — últimos 6 meses ──────────────────────
            await using var cmdTransfer = conn.CreateCommand();
            cmdTransfer.CommandText = """
                SELECT
                    FORMAT(CreatedAt, 'MMM yy', 'es-MX') AS Label,
                    YEAR(CreatedAt)                        AS Yr,
                    MONTH(CreatedAt)                       AS Mo,
                    COUNT(*)                               AS Cnt
                FROM dbo.EquipmentTransfers
                WHERE CreatedAt >= DATEADD(MONTH, -5, DATEFROMPARTS(YEAR(GETUTCDATE()), MONTH(GETUTCDATE()), 1))
                GROUP BY FORMAT(CreatedAt, 'MMM yy', 'es-MX'), YEAR(CreatedAt), MONTH(CreatedAt)
                ORDER BY Yr, Mo
                """;
            var transfersByMonth = new List<object>();
            await using (var r = await cmdTransfer.ExecuteReaderAsync(ct))
                while (await r.ReadAsync(ct))
                    transfersByMonth.Add(new { label = r.GetString(0), count = r.GetInt32(3) });

            // ── Por departamento ──────────────────────────────────────────────
            await using var cmdDept = conn.CreateCommand();
            cmdDept.CommandText = """
                SELECT
                    ISNULL(d.Name, 'Sin departamento') AS Name,
                    COUNT(*) AS Value
                FROM dbo.InventoryItems i
                LEFT JOIN dbo.Departments d ON d.Id = i.DepartmentId
                WHERE i.IsActive = 1
                GROUP BY ISNULL(d.Name, 'Sin departamento')
                ORDER BY Value DESC
                """;
            var byDepartment = new List<object>();
            await using (var r = await cmdDept.ExecuteReaderAsync(ct))
                while (await r.ReadAsync(ct))
                    byDepartment.Add(new { name = r.GetString(0), value = r.GetInt32(1) });

            return Ok(new
            {
                total,
                active,
                maintenance,
                decommissioned,
                inStorage,
                byType,
                byStatus,
                transfersByMonth,
                byDepartment,
            });
        }
        catch (Exception ex) { logger.LogError(ex, "GetInventoryReport"); return StatusCode(500, ex.Message); }
    }

    // ════════════════════════════════════════════════════════════════════════
    //  CALENDARIO
    // ════════════════════════════════════════════════════════════════════════

    [HttpGet("calendar")]
    public async Task<IActionResult> GetCalendarReport(CancellationToken ct)
    {
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);

            // ── KPIs ─────────────────────────────────────────────────────────
            await using var cmdKpi = conn.CreateCommand();
            cmdKpi.CommandText = """
                SELECT
                    COUNT(*) AS Total,
                    SUM(CASE WHEN YEAR(StartTime) = YEAR(GETUTCDATE())
                              AND MONTH(StartTime) = MONTH(GETUTCDATE()) THEN 1 ELSE 0 END) AS ThisMonth,
                    SUM(CASE WHEN StartTime > GETUTCDATE() THEN 1 ELSE 0 END)               AS Upcoming,
                    (SELECT COUNT(*) FROM dbo.Rooms WHERE IsActive = 1)                     AS ActiveRooms
                FROM dbo.Reservations
                """;
            int total = 0, thisMonth = 0, upcoming = 0, activeRooms = 0;
            await using (var r = await cmdKpi.ExecuteReaderAsync(ct))
            {
                if (await r.ReadAsync(ct))
                {
                    total       = r.GetInt32(0);
                    thisMonth   = r.GetInt32(1);
                    upcoming    = r.GetInt32(2);
                    activeRooms = r.GetInt32(3);
                }
            }

            // ── Por sala ──────────────────────────────────────────────────────
            await using var cmdRoom = conn.CreateCommand();
            cmdRoom.CommandText = """
                SELECT
                    r.Name AS Name,
                    COUNT(res.Id) AS Value
                FROM dbo.Rooms r
                LEFT JOIN dbo.Reservations res ON res.RoomId = r.Id
                WHERE r.IsActive = 1
                GROUP BY r.Name
                ORDER BY Value DESC
                """;
            var byRoom = new List<object>();
            await using (var r = await cmdRoom.ExecuteReaderAsync(ct))
                while (await r.ReadAsync(ct))
                    byRoom.Add(new { name = r.GetString(0), value = r.GetInt32(1) });

            // ── Por día de la semana ──────────────────────────────────────────
            await using var cmdDow = conn.CreateCommand();
            cmdDow.CommandText = """
                SELECT
                    DATEPART(WEEKDAY, StartTime) AS DayNum,
                    COUNT(*) AS Value
                FROM dbo.Reservations
                GROUP BY DATEPART(WEEKDAY, StartTime)
                ORDER BY DayNum
                """;
            // DATEPART WEEKDAY: 1=Sunday ... 7=Saturday (default DATEFIRST=7)
            var dowLabels = new[] { "Dom", "Lun", "Mar", "Mié", "Jue", "Vie", "Sáb" };
            var dowMap    = new Dictionary<int, int>();
            await using (var r = await cmdDow.ExecuteReaderAsync(ct))
                while (await r.ReadAsync(ct))
                    dowMap[r.GetInt32(0)] = r.GetInt32(1);
            var byDayOfWeek = Enumerable.Range(1, 7)
                .Select(d => (object)new { name = dowLabels[d - 1], value = dowMap.GetValueOrDefault(d, 0) })
                .ToList();

            // ── Tendencia mensual — últimos 6 meses ───────────────────────────
            await using var cmdMonth = conn.CreateCommand();
            cmdMonth.CommandText = """
                SELECT
                    FORMAT(StartTime, 'MMM yy', 'es-MX') AS Label,
                    YEAR(StartTime)                        AS Yr,
                    MONTH(StartTime)                       AS Mo,
                    COUNT(*) AS Cnt
                FROM dbo.Reservations
                WHERE StartTime >= DATEADD(MONTH, -5, DATEFROMPARTS(YEAR(GETUTCDATE()), MONTH(GETUTCDATE()), 1))
                GROUP BY FORMAT(StartTime, 'MMM yy', 'es-MX'), YEAR(StartTime), MONTH(StartTime)
                ORDER BY Yr, Mo
                """;
            var byMonth = new List<object>();
            await using (var r = await cmdMonth.ExecuteReaderAsync(ct))
                while (await r.ReadAsync(ct))
                    byMonth.Add(new { label = r.GetString(0), count = r.GetInt32(3) });

            return Ok(new
            {
                total,
                thisMonth,
                upcoming,
                activeRooms,
                byRoom,
                byDayOfWeek,
                byMonth,
            });
        }
        catch (Exception ex) { logger.LogError(ex, "GetCalendarReport"); return StatusCode(500, ex.Message); }
    }
}
