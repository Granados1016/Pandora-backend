using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;

namespace Pandora.API.Controllers;

/// <summary>
/// Dashboard ejecutivo — consolida KPIs de todos los módulos en una sola llamada.
/// Solo accesible por Admin.
/// </summary>
[ApiController]
[Route("api/ejecutivo")]
[Authorize(Roles = "Admin")]
public class EjecutivoController(
    IConfiguration config,
    ILogger<EjecutivoController> logger) : ControllerBase
{
    private SqlConnection Conn() => new(config.GetConnectionString("PandoraDb"));

    [HttpGet("stats")]
    public async Task<IActionResult> GetStats(CancellationToken ct)
    {
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await using var cmd = conn.CreateCommand();

            cmd.CommandText = """
                -- ── Tickets ──────────────────────────────────────────────────────────
                SELECT
                    COUNT(*)                                                                           AS TktTotal,
                    ISNULL(SUM(CASE WHEN Status='Abierto'     THEN 1 ELSE 0 END), 0)                  AS TktAbiertos,
                    ISNULL(SUM(CASE WHEN Status='En Progreso' THEN 1 ELSE 0 END), 0)                  AS TktEnProgreso,
                    ISNULL(SUM(CASE WHEN Status='Resuelto' OR Status='Cerrado' THEN 1 ELSE 0 END), 0) AS TktResueltos,
                    ISNULL(SUM(CASE WHEN Priority='Crítica'   THEN 1 ELSE 0 END), 0)                  AS TktCriticos,
                    ISNULL(SUM(CASE WHEN MONTH(CreatedAt)=MONTH(GETUTCDATE()) AND YEAR(CreatedAt)=YEAR(GETUTCDATE()) THEN 1 ELSE 0 END), 0) AS TktEsteMes
                FROM dbo.Tickets;

                -- ── Inventario ────────────────────────────────────────────────────────
                IF OBJECT_ID('dbo.InventoryItems') IS NOT NULL
                    SELECT
                        COUNT(*)                                                                       AS InvTotal,
                        ISNULL(SUM(CASE WHEN Status='Activo'        THEN 1 ELSE 0 END), 0)            AS InvActivos,
                        ISNULL(SUM(CASE WHEN Status='En Reparación' THEN 1 ELSE 0 END), 0)            AS InvReparacion,
                        ISNULL(SUM(CASE WHEN Status='Dado de Baja'  THEN 1 ELSE 0 END), 0)            AS InvBaja
                    FROM dbo.InventoryItems;
                ELSE
                    SELECT 0 AS InvTotal, 0 AS InvActivos, 0 AS InvReparacion, 0 AS InvBaja;

                -- ── Bitácora ──────────────────────────────────────────────────────────
                IF OBJECT_ID('dbo.BitacoraRegistros') IS NOT NULL
                    SELECT
                        COUNT(*)                                                                       AS BitTotal,
                        ISNULL(SUM(CASE WHEN Estado='Abierta'    THEN 1 ELSE 0 END), 0)               AS BitAbiertas,
                        ISNULL(SUM(CASE WHEN Estado='En Proceso' THEN 1 ELSE 0 END), 0)               AS BitEnProceso,
                        ISNULL(SUM(CASE WHEN Prioridad='Alta'    THEN 1 ELSE 0 END), 0)               AS BitPrioAlta
                    FROM dbo.BitacoraRegistros;
                ELSE IF OBJECT_ID('dbo.Bitacora') IS NOT NULL
                    SELECT
                        COUNT(*)                                                                       AS BitTotal,
                        ISNULL(SUM(CASE WHEN Estado='Abierta'    THEN 1 ELSE 0 END), 0)               AS BitAbiertas,
                        ISNULL(SUM(CASE WHEN Estado='En Proceso' THEN 1 ELSE 0 END), 0)               AS BitEnProceso,
                        ISNULL(SUM(CASE WHEN Prioridad='Alta'    THEN 1 ELSE 0 END), 0)               AS BitPrioAlta
                    FROM dbo.Bitacora;
                ELSE
                    SELECT 0 AS BitTotal, 0 AS BitAbiertas, 0 AS BitEnProceso, 0 AS BitPrioAlta;

                -- ── Mantenimiento ─────────────────────────────────────────────────────
                IF OBJECT_ID('dbo.Mantenimientos') IS NOT NULL
                    SELECT
                        COUNT(*)                                                                       AS MntTotal,
                        ISNULL(SUM(CASE WHEN Estado='Programado' THEN 1 ELSE 0 END), 0)               AS MntProgramados,
                        ISNULL(SUM(CASE WHEN Estado='En Proceso' THEN 1 ELSE 0 END), 0)               AS MntEnProceso,
                        ISNULL(SUM(CASE WHEN FechaProgramada < GETUTCDATE() AND Estado NOT IN ('Completado','Cancelado') THEN 1 ELSE 0 END), 0) AS MntVencidos
                    FROM dbo.Mantenimientos;
                ELSE
                    SELECT 0 AS MntTotal, 0 AS MntProgramados, 0 AS MntEnProceso, 0 AS MntVencidos;

                -- ── Checador (hoy — sin AT TIME ZONE para compatibilidad con LocalDB) ─
                IF OBJECT_ID('dbo.CheckadorRegistros') IS NOT NULL
                    SELECT
                        ISNULL(SUM(CASE WHEN CAST(Timestamp AS DATE)=CAST(GETUTCDATE() AS DATE) AND Tipo='Entrada' THEN 1 ELSE 0 END),0) AS ChkEntradasHoy,
                        ISNULL(SUM(CASE WHEN CAST(Timestamp AS DATE)=CAST(GETUTCDATE() AS DATE) AND Tipo='Salida'  THEN 1 ELSE 0 END),0) AS ChkSalidasHoy,
                        COUNT(DISTINCT UserId)                                                                                           AS ChkUsuarios
                    FROM dbo.CheckadorRegistros;
                ELSE
                    SELECT 0 AS ChkEntradasHoy, 0 AS ChkSalidasHoy, 0 AS ChkUsuarios;

                -- ── Usuarios activos ──────────────────────────────────────────────────
                SELECT
                    COUNT(*)                                                                           AS UsrTotal,
                    ISNULL(SUM(CASE WHEN IsActive=1 THEN 1 ELSE 0 END), 0)                            AS UsrActivos
                FROM dbo.AppUsers;

                -- ── Licencias próximas a vencer (30 días) ─────────────────────────────
                IF OBJECT_ID('dbo.Licencias') IS NOT NULL
                    SELECT
                        COUNT(*)                                                                       AS LicTotal,
                        ISNULL(SUM(CASE WHEN ProximoPago BETWEEN CAST(GETUTCDATE() AS DATE) AND CAST(DATEADD(DAY,30,GETUTCDATE()) AS DATE) THEN 1 ELSE 0 END),0) AS LicPorVencer,
                        ISNULL(SUM(CASE WHEN ProximoPago < CAST(GETUTCDATE() AS DATE) THEN 1 ELSE 0 END),0) AS LicVencidas
                    FROM dbo.Licencias WHERE Estado='Activa';
                ELSE
                    SELECT 0 AS LicTotal, 0 AS LicPorVencer, 0 AS LicVencidas;

                -- ── Vacaciones pendientes de aprobar ──────────────────────────────────
                IF OBJECT_ID('dbo.VacationRequests') IS NOT NULL
                    SELECT
                        ISNULL(SUM(CASE WHEN Status='Pendiente' THEN 1 ELSE 0 END),0) AS VacPendientes,
                        ISNULL(SUM(CASE WHEN Status='Aprobado'  THEN 1 ELSE 0 END),0) AS VacAprobadas
                    FROM dbo.VacationRequests;
                ELSE
                    SELECT 0 AS VacPendientes, 0 AS VacAprobadas;
                """;

            await using var r = await cmd.ExecuteReaderAsync(ct);

            // ── Tickets ─────────────────────────────────────────────────────────
            await r.ReadAsync(ct);
            var tickets = new {
                total     = r.GetInt32(0), abiertos  = r.GetInt32(1),
                enProgreso= r.GetInt32(2), resueltos = r.GetInt32(3),
                criticos  = r.GetInt32(4), esteMes   = r.GetInt32(5),
            };

            // ── Inventario ───────────────────────────────────────────────────────
            await r.NextResultAsync(ct); await r.ReadAsync(ct);
            var inventario = new {
                total    = r.GetInt32(0), activos    = r.GetInt32(1),
                reparacion= r.GetInt32(2), baja      = r.GetInt32(3),
            };

            // ── Bitácora ─────────────────────────────────────────────────────────
            await r.NextResultAsync(ct); await r.ReadAsync(ct);
            var bitacora = new {
                total    = r.GetInt32(0), abiertas  = r.GetInt32(1),
                enProceso= r.GetInt32(2), prioAlta  = r.GetInt32(3),
            };

            // ── Mantenimiento ────────────────────────────────────────────────────
            await r.NextResultAsync(ct); await r.ReadAsync(ct);
            var mantenimiento = new {
                total      = r.GetInt32(0), programados = r.GetInt32(1),
                enProceso  = r.GetInt32(2), vencidos    = r.GetInt32(3),
            };

            // ── Checador ─────────────────────────────────────────────────────────
            await r.NextResultAsync(ct); await r.ReadAsync(ct);
            var checador = new {
                entradasHoy = r.GetInt32(0), salidasHoy = r.GetInt32(1),
                usuarios    = r.GetInt32(2),
            };

            // ── Usuarios ─────────────────────────────────────────────────────────
            await r.NextResultAsync(ct); await r.ReadAsync(ct);
            var usuarios = new { total = r.GetInt32(0), activos = r.GetInt32(1) };

            // ── Licencias ────────────────────────────────────────────────────────
            await r.NextResultAsync(ct); await r.ReadAsync(ct);
            var licencias = new {
                total = r.GetInt32(0), porVencer = r.GetInt32(1), vencidas = r.GetInt32(2),
            };

            // ── Vacaciones ───────────────────────────────────────────────────────
            await r.NextResultAsync(ct); await r.ReadAsync(ct);
            var vacaciones = new { pendientes = r.GetInt32(0), aprobadas = r.GetInt32(1) };

            return Ok(new { tickets, inventario, bitacora, mantenimiento, checador, usuarios, licencias, vacaciones });
        }
        catch (Exception ex)
        {
            logger.LogError(ex, "Ejecutivo.GetStats");
            return StatusCode(500, ex.Message);
        }
    }
}
