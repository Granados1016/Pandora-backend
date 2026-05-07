using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using System.Security.Claims;

namespace Pandora.API.Controllers;

/// <summary>
/// Gestión de procedimientos/documentos.
/// Los archivos se almacenan en la BD (VARBINARY MAX) para garantizar
/// persistencia en Railway sin necesidad de volúmenes externos.
/// </summary>
[ApiController]
[Route("api/procedimientos")]
[Authorize]
public class ProcedimientosController(
    IConfiguration config,
    ILogger<ProcedimientosController> logger) : ControllerBase
{
    private SqlConnection Conn() => new(config.GetConnectionString("PandoraDb"));

    private string? CurrentUser() =>
        User.FindFirstValue(ClaimTypes.Name) ?? User.FindFirstValue("name");

    // ── POST /api/procedimientos ──────────────────────────────────────────────
    /// <summary>Sube un nuevo procedimiento (binario en BD).</summary>
    [HttpPost]
    [RequestSizeLimit(52_428_800)] // 50 MB
    public async Task<IActionResult> Upload(
        IFormFile file,
        [FromForm] string title,
        [FromForm] string? description,
        [FromForm] string? category,
        CancellationToken ct)
    {
        if (file is null || file.Length == 0)
            return BadRequest("Archivo requerido.");
        if (string.IsNullOrWhiteSpace(title))
            return BadRequest("El título es requerido.");

        // Leer bytes en memoria (límite 50 MB ya forzado por RequestSizeLimit)
        using var ms = new MemoryStream();
        await file.CopyToAsync(ms, ct);
        var bytes = ms.ToArray();

        await using var conn = Conn();
        await conn.OpenAsync(ct);
        await using var cmd = conn.CreateCommand();
        cmd.CommandTimeout = 120;
        cmd.CommandText = """
            INSERT INTO dbo.Procedimientos
                (Title, Description, Category, FileName, FileContentType, FileSize, FileData, UploadedBy, UploadedAt)
            OUTPUT INSERTED.Id
            VALUES
                (@Title, @Desc, @Cat, @FileName, @ContentType, @FileSize, @FileData, @UploadedBy, GETUTCDATE())
            """;
        cmd.Parameters.AddWithValue("@Title",       title.Trim());
        cmd.Parameters.AddWithValue("@Desc",        (object?)description?.Trim() ?? DBNull.Value);
        cmd.Parameters.AddWithValue("@Cat",         (object?)category?.Trim()    ?? DBNull.Value);
        cmd.Parameters.AddWithValue("@FileName",    file.FileName);
        cmd.Parameters.AddWithValue("@ContentType", file.ContentType ?? "application/octet-stream");
        cmd.Parameters.AddWithValue("@FileSize",    file.Length);
        cmd.Parameters.AddWithValue("@FileData",    bytes);
        cmd.Parameters.AddWithValue("@UploadedBy",  CurrentUser() ?? "desconocido");

        var id = (int)(await cmd.ExecuteScalarAsync(ct) ?? 0);
        logger.LogInformation("Procedimiento #{Id} '{Title}' subido por {User} ({Size} KB)",
            id, title, CurrentUser(), bytes.Length / 1024);
        return Ok(new { id });
    }

    // ── GET /api/procedimientos ───────────────────────────────────────────────
    /// <summary>Lista procedimientos con filtros y paginación. No devuelve binarios.</summary>
    [HttpGet]
    public async Task<IActionResult> GetAll(
        [FromQuery] string? search,
        [FromQuery] string? category,
        [FromQuery] int page     = 1,
        [FromQuery] int pageSize = 20,
        CancellationToken ct = default)
    {
        page     = Math.Max(1, page);
        pageSize = Math.Clamp(pageSize, 1, 100);
        int offset = (page - 1) * pageSize;

        await using var conn = Conn();
        await conn.OpenAsync(ct);

        // Total de registros (para calcular páginas)
        int total = 0;
        await using (var cntCmd = conn.CreateCommand())
        {
            cntCmd.CommandText = """
                SELECT COUNT(*)
                FROM dbo.Procedimientos
                WHERE IsDeleted = 0
                  AND (@Search IS NULL OR Title       LIKE '%' + @Search + '%'
                                       OR Description LIKE '%' + @Search + '%')
                  AND (@Cat    IS NULL OR Category = @Cat)
                """;
            cntCmd.Parameters.AddWithValue("@Search",
                string.IsNullOrWhiteSpace(search)   ? (object)DBNull.Value : search.Trim());
            cntCmd.Parameters.AddWithValue("@Cat",
                string.IsNullOrWhiteSpace(category) ? (object)DBNull.Value : category.Trim());
            total = (int)(await cntCmd.ExecuteScalarAsync(ct) ?? 0);
        }

        await using var cmd = conn.CreateCommand();
        cmd.CommandText = """
            SELECT Id, Title, Description, Category,
                   FileName, FileContentType, FileSize, UploadedBy, UploadedAt
            FROM dbo.Procedimientos
            WHERE IsDeleted = 0
              AND (@Search IS NULL OR Title       LIKE '%' + @Search + '%'
                                   OR Description LIKE '%' + @Search + '%')
              AND (@Cat    IS NULL OR Category = @Cat)
            ORDER BY UploadedAt DESC
            OFFSET @Offset ROWS FETCH NEXT @PageSize ROWS ONLY
            """;
        cmd.Parameters.AddWithValue("@Search",
            string.IsNullOrWhiteSpace(search)   ? (object)DBNull.Value : search.Trim());
        cmd.Parameters.AddWithValue("@Cat",
            string.IsNullOrWhiteSpace(category) ? (object)DBNull.Value : category.Trim());
        cmd.Parameters.AddWithValue("@Offset",   offset);
        cmd.Parameters.AddWithValue("@PageSize", pageSize);

        var items = new List<object>();
        await using var r = await cmd.ExecuteReaderAsync(ct);
        while (await r.ReadAsync(ct))
            items.Add(new
            {
                id              = r.GetInt32(0),
                title           = r.GetString(1),
                description     = r.IsDBNull(2) ? null : r.GetString(2),
                category        = r.IsDBNull(3) ? null : r.GetString(3),
                fileName        = r.GetString(4),
                fileContentType = r.GetString(5),
                fileSize        = r.GetInt64(6),
                uploadedBy      = r.GetString(7),
                uploadedAt      = r.GetDateTime(8),
            });

        return Ok(new
        {
            items,
            total,
            page,
            pageSize,
            totalPages = (int)Math.Ceiling(total / (double)pageSize),
        });
    }

    // ── GET /api/procedimientos/{id}/view ─────────────────────────────────────
    [HttpGet("{id:int}/view")]
    [AllowAnonymous]
    public async Task<IActionResult> View(int id, [FromQuery] string? access_token, CancellationToken ct)
    {
        if (!await ValidateToken(access_token)) return Unauthorized();
        return await ServeFile(id, inline: true, ct);
    }

    // ── GET /api/procedimientos/{id}/download ─────────────────────────────────
    [HttpGet("{id:int}/download")]
    [AllowAnonymous]
    public async Task<IActionResult> Download(int id, [FromQuery] string? access_token, CancellationToken ct)
    {
        if (!await ValidateToken(access_token)) return Unauthorized();
        return await ServeFile(id, inline: false, ct);
    }

    // ── DELETE /api/procedimientos/{id} ──────────────────────────────────────
    [HttpDelete("{id:int}")]
    [Authorize(Roles = "Admin")]
    public async Task<IActionResult> Delete(int id, CancellationToken ct)
    {
        await using var conn = Conn();
        await conn.OpenAsync(ct);
        await using var cmd = conn.CreateCommand();
        // Soft-delete: pone IsDeleted=1 y borra el binario para liberar espacio en BD
        cmd.CommandText = """
            UPDATE dbo.Procedimientos
            SET IsDeleted = 1, FileData = NULL
            WHERE Id = @Id AND IsDeleted = 0
            """;
        cmd.Parameters.AddWithValue("@Id", id);
        int rows = await cmd.ExecuteNonQueryAsync(ct);
        if (rows == 0) return NotFound();
        logger.LogInformation("Procedimiento #{Id} eliminado por {User}", id, CurrentUser());
        return NoContent();
    }

    // ══════════════════════════════════════════════════════════════════════════
    // CATEGORÍAS
    // ══════════════════════════════════════════════════════════════════════════

    [HttpGet("categorias")]
    public async Task<IActionResult> GetCategorias(CancellationToken ct)
    {
        await using var conn = Conn();
        await conn.OpenAsync(ct);
        await using var cmd = conn.CreateCommand();
        cmd.CommandText = """
            SELECT Id, Name, Color, SortOrder,
                   (SELECT COUNT(*) FROM dbo.Procedimientos p
                    WHERE p.Category = c.Name AND p.IsDeleted = 0) AS UsageCount
            FROM dbo.ProcedimientoCategorias c
            WHERE IsActive = 1
            ORDER BY SortOrder, Name
            """;
        var list = new List<object>();
        await using var r = await cmd.ExecuteReaderAsync(ct);
        while (await r.ReadAsync(ct))
            list.Add(new
            {
                id         = r.GetInt32(0),
                name       = r.GetString(1),
                color      = r.IsDBNull(2) ? "default" : r.GetString(2),
                sortOrder  = r.GetInt32(3),
                usageCount = r.GetInt32(4),
            });
        return Ok(list);
    }

    [HttpPost("categorias")]
    [Authorize(Roles = "Admin")]
    public async Task<IActionResult> CreateCategoria([FromBody] CategoriaDto dto, CancellationToken ct)
    {
        if (string.IsNullOrWhiteSpace(dto.Name)) return BadRequest("El nombre es requerido.");

        await using var conn = Conn();
        await conn.OpenAsync(ct);
        await using var cmd = conn.CreateCommand();
        cmd.CommandText = """
            IF EXISTS (SELECT 1 FROM dbo.ProcedimientoCategorias WHERE Name = @Name AND IsActive = 1)
                SELECT -1
            ELSE
                INSERT INTO dbo.ProcedimientoCategorias (Name, Color, SortOrder)
                OUTPUT INSERTED.Id
                VALUES (@Name, @Color, @SortOrder)
            """;
        cmd.Parameters.AddWithValue("@Name",      dto.Name.Trim());
        cmd.Parameters.AddWithValue("@Color",     (object?)dto.Color?.Trim() ?? DBNull.Value);
        cmd.Parameters.AddWithValue("@SortOrder", dto.SortOrder);
        var result = await cmd.ExecuteScalarAsync(ct);
        int newId = Convert.ToInt32(result);
        if (newId == -1) return Conflict("Ya existe una categoría con ese nombre.");
        return Ok(new { id = newId });
    }

    [HttpPut("categorias/{id:int}")]
    [Authorize(Roles = "Admin")]
    public async Task<IActionResult> UpdateCategoria(int id, [FromBody] CategoriaDto dto, CancellationToken ct)
    {
        if (string.IsNullOrWhiteSpace(dto.Name)) return BadRequest("El nombre es requerido.");

        await using var conn = Conn();
        await conn.OpenAsync(ct);

        string? oldName = null;
        await using (var q = conn.CreateCommand())
        {
            q.CommandText = "SELECT Name FROM dbo.ProcedimientoCategorias WHERE Id = @Id AND IsActive = 1";
            q.Parameters.AddWithValue("@Id", id);
            oldName = await q.ExecuteScalarAsync(ct) as string;
        }
        if (oldName is null) return NotFound();

        await using var cmd = conn.CreateCommand();
        cmd.CommandText = """
            UPDATE dbo.ProcedimientoCategorias
            SET Name = @Name, Color = @Color, SortOrder = @SortOrder
            WHERE Id = @Id AND IsActive = 1;

            UPDATE dbo.Procedimientos
            SET Category = @Name
            WHERE Category = @OldName AND IsDeleted = 0;
            """;
        cmd.Parameters.AddWithValue("@Id",        id);
        cmd.Parameters.AddWithValue("@Name",      dto.Name.Trim());
        cmd.Parameters.AddWithValue("@Color",     (object?)dto.Color?.Trim() ?? DBNull.Value);
        cmd.Parameters.AddWithValue("@SortOrder", dto.SortOrder);
        cmd.Parameters.AddWithValue("@OldName",   oldName);
        await cmd.ExecuteNonQueryAsync(ct);
        return NoContent();
    }

    [HttpDelete("categorias/{id:int}")]
    [Authorize(Roles = "Admin")]
    public async Task<IActionResult> DeleteCategoria(int id, CancellationToken ct)
    {
        await using var conn = Conn();
        await conn.OpenAsync(ct);

        await using (var q = conn.CreateCommand())
        {
            q.CommandText = """
                SELECT COUNT(*) FROM dbo.Procedimientos p
                JOIN dbo.ProcedimientoCategorias c ON c.Name = p.Category
                WHERE c.Id = @Id AND p.IsDeleted = 0
                """;
            q.Parameters.AddWithValue("@Id", id);
            int count = (int)(await q.ExecuteScalarAsync(ct) ?? 0);
            if (count > 0)
                return Conflict($"Esta categoría tiene {count} procedimiento(s) activos. Reasígnalos antes de eliminarla.");
        }

        await using var cmd = conn.CreateCommand();
        cmd.CommandText = "UPDATE dbo.ProcedimientoCategorias SET IsActive = 0 WHERE Id = @Id";
        cmd.Parameters.AddWithValue("@Id", id);
        int rows = await cmd.ExecuteNonQueryAsync(ct);
        return rows == 0 ? NotFound() : NoContent();
    }

    // ── Helpers ───────────────────────────────────────────────────────────────
    private async Task<IActionResult> ServeFile(int id, bool inline, CancellationToken ct)
    {
        string? contentType = null, fileName = null;
        byte[]? data = null;

        await using var conn = Conn();
        await conn.OpenAsync(ct);
        await using var cmd = conn.CreateCommand();
        cmd.CommandTimeout = 120;
        cmd.CommandText = """
            SELECT FileContentType, FileName, FileData
            FROM dbo.Procedimientos
            WHERE Id = @Id AND IsDeleted = 0
            """;
        cmd.Parameters.AddWithValue("@Id", id);
        await using var r = await cmd.ExecuteReaderAsync(
            System.Data.CommandBehavior.SequentialAccess, ct);

        if (await r.ReadAsync(ct))
        {
            contentType = r.GetString(0);
            fileName    = r.GetString(1);
            data        = r.IsDBNull(2) ? null : (byte[])r.GetValue(2);
        }

        if (data is null)
            return NotFound("Archivo no disponible.");

        Response.Headers["Content-Disposition"] = inline
            ? $"inline; filename=\"{fileName}\""
            : $"attachment; filename=\"{fileName}\"";
        return File(data, contentType ?? "application/octet-stream");
    }

    private Task<bool> ValidateToken(string? token)
    {
        if (User.Identity?.IsAuthenticated == true) return Task.FromResult(true);
        if (string.IsNullOrWhiteSpace(token))       return Task.FromResult(false);
        try
        {
            var key     = config["JwtSettings:Key"]!;
            var issuer  = config["JwtSettings:Issuer"]!;
            var aud     = config["JwtSettings:Audience"]!;
            var handler = new System.IdentityModel.Tokens.Jwt.JwtSecurityTokenHandler();
            var sigKey  = new Microsoft.IdentityModel.Tokens.SymmetricSecurityKey(
                              System.Text.Encoding.UTF8.GetBytes(key));
            handler.ValidateToken(token, new Microsoft.IdentityModel.Tokens.TokenValidationParameters
            {
                ValidateIssuerSigningKey = true,  IssuerSigningKey = sigKey,
                ValidateIssuer          = true,   ValidIssuer      = issuer,
                ValidateAudience        = true,   ValidAudience    = aud,
                ValidateLifetime        = true,
            }, out _);
            return Task.FromResult(true);
        }
        catch { return Task.FromResult(false); }
    }
}

// ── DTOs ──────────────────────────────────────────────────────────────────────
public record CategoriaDto(string Name, string? Color, int SortOrder = 0);
