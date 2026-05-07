using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;

namespace Pandora.API.Controllers;

/// <summary>
/// Búsqueda global que agrega resultados de múltiples módulos.
/// </summary>
[ApiController]
[Route("api/search")]
[Authorize]
public class SearchController(IConfiguration config) : ControllerBase
{
    private SqlConnection Conn() => new(config.GetConnectionString("PandoraDb"));

    // ── GET /api/search?q=texto ───────────────────────────────────────────────
    [HttpGet]
    public async Task<IActionResult> Search([FromQuery] string? q, CancellationToken ct)
    {
        if (string.IsNullOrWhiteSpace(q) || q.Trim().Length < 2)
            return Ok(new { results = Array.Empty<object>(), total = 0 });

        string term = q.Trim();
        var results = new List<object>();

        await using var conn = Conn();
        await conn.OpenAsync(ct);

        // ── Procedimientos ────────────────────────────────────────────────────
        await TrySearch(conn, ct, results, """
            SELECT TOP 5 'procedimiento' AS Type, Id, Title AS Label,
                   ISNULL(Description,'') AS Subtitle, Category AS Tag
            FROM dbo.Procedimientos
            WHERE IsDeleted = 0
              AND (Title LIKE '%' + @Q + '%' OR Description LIKE '%' + @Q + '%')
            ORDER BY UploadedAt DESC
            """, term, "/procedimientos");

        // ── Comunicados ───────────────────────────────────────────────────────
        await TrySearch(conn, ct, results, """
            SELECT TOP 5 'comunicado' AS Type, Id, Title AS Label,
                   LEFT(Content,120) AS Subtitle, Priority AS Tag
            FROM dbo.Comunicados
            WHERE IsDeleted = 0 AND IsPublished = 1
              AND (Title LIKE '%' + @Q + '%' OR Content LIKE '%' + @Q + '%')
            ORDER BY CreatedAt DESC
            """, term, "/comunicados");

        // ── Inventario ────────────────────────────────────────────────────────
        await TrySearch(conn, ct, results, """
            SELECT TOP 5 'inventario' AS Type, CAST(Id AS NVARCHAR(36)), Name AS Label,
                   ISNULL(Brand,'') + ' ' + ISNULL(Model,'') AS Subtitle,
                   ISNULL(Department,'') AS Tag
            FROM dbo.InventoryItems
            WHERE IsActive = 1
              AND (Name LIKE '%' + @Q + '%' OR Brand LIKE '%' + @Q + '%'
                OR Model LIKE '%' + @Q + '%' OR InventoryNumber LIKE '%' + @Q + '%')
            ORDER BY CreatedAt DESC
            """, term, "/inventory/items");

        // ── Licencias ─────────────────────────────────────────────────────────
        await TrySearch(conn, ct, results, """
            SELECT TOP 5 'licencia' AS Type, CAST(Id AS NVARCHAR(10)), Plataforma AS Label,
                   Area AS Subtitle, Estado AS Tag
            FROM dbo.Licencias
            WHERE (Plataforma LIKE '%' + @Q + '%' OR Area LIKE '%' + @Q + '%'
                OR Responsable LIKE '%' + @Q + '%')
            ORDER BY Id
            """, term, "/licencias");

        // ── Usuarios (solo admin) ─────────────────────────────────────────────
        if (User.IsInRole("Admin"))
        {
            await TrySearch(conn, ct, results, """
                SELECT TOP 5 'usuario' AS Type, Id, FullName AS Label,
                       Username AS Subtitle, Role AS Tag
                FROM dbo.AppUsers
                WHERE IsActive = 1
                  AND (FullName LIKE '%' + @Q + '%' OR Username LIKE '%' + @Q + '%'
                    OR Email LIKE '%' + @Q + '%')
                ORDER BY FullName
                """, term, "/admin");
        }

        return Ok(new { results, total = results.Count });
    }

    private static async Task TrySearch(
        SqlConnection conn, CancellationToken ct,
        List<object> results, string sql, string term, string path)
    {
        try
        {
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = sql;
            cmd.Parameters.AddWithValue("@Q", term);
            await using var r = await cmd.ExecuteReaderAsync(ct);
            while (await r.ReadAsync(ct))
            {
                results.Add(new
                {
                    type     = r.GetString(0),
                    id       = r.GetValue(1)?.ToString(),
                    label    = r.IsDBNull(2) ? "" : r.GetString(2),
                    subtitle = r.IsDBNull(3) ? "" : r.GetString(3),
                    tag      = r.IsDBNull(4) ? "" : r.GetString(4),
                    path,
                });
            }
        }
        catch { /* módulo puede no tener la tabla en todos los entornos */ }
    }
}
