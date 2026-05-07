using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using System.Security.Claims;

namespace Pandora.API.Controllers;

[ApiController]
[Route("api/procedimientos")]
[Authorize]
public class ProcedimientosController(
    IConfiguration config,
    IWebHostEnvironment env,
    ILogger<ProcedimientosController> logger) : ControllerBase
{
    // ── Carpeta de almacenamiento ─────────────────────────────────────────────
    private string StorageDir =>
        Path.Combine(env.ContentRootPath, "uploads", "procedimientos");

    private SqlConnection Conn() => new(config.GetConnectionString("PandoraDb"));

    // Resolución de access_token en query string (descarga directa)
    private string? CurrentUser() =>
        User.FindFirstValue(ClaimTypes.Name)
        ?? User.FindFirstValue("name");

    // ── POST /api/procedimientos ──────────────────────────────────────────────
    /// <summary>Sube un nuevo procedimiento (archivo + metadatos).</summary>
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

        Directory.CreateDirectory(StorageDir);

        // Nombre único en disco
        var ext      = Path.GetExtension(file.FileName);
        var diskName = $"{Guid.NewGuid()}{ext}";
        var diskPath = Path.Combine(StorageDir, diskName);

        await using (var stream = System.IO.File.Create(diskPath))
            await file.CopyToAsync(stream, ct);

        await using var conn = Conn();
        await conn.OpenAsync(ct);
        await using var cmd = conn.CreateCommand();
        cmd.CommandText = """
            INSERT INTO dbo.Procedimientos
                (Title, Description, Category, FileName, FileContentType, FileSize, FilePath, UploadedBy, UploadedAt)
            OUTPUT INSERTED.Id
            VALUES
                (@Title, @Desc, @Cat, @FileName, @ContentType, @FileSize, @FilePath, @UploadedBy, GETUTCDATE())
            """;
        cmd.Parameters.AddWithValue("@Title",       title.Trim());
        cmd.Parameters.AddWithValue("@Desc",        (object?)description?.Trim() ?? DBNull.Value);
        cmd.Parameters.AddWithValue("@Cat",         (object?)category?.Trim()    ?? DBNull.Value);
        cmd.Parameters.AddWithValue("@FileName",    file.FileName);
        cmd.Parameters.AddWithValue("@ContentType", file.ContentType ?? "application/octet-stream");
        cmd.Parameters.AddWithValue("@FileSize",    file.Length);
        cmd.Parameters.AddWithValue("@FilePath",    diskName);
        cmd.Parameters.AddWithValue("@UploadedBy",  CurrentUser() ?? "desconocido");

        var id = (int)(await cmd.ExecuteScalarAsync(ct) ?? 0);
        logger.LogInformation("Procedimiento #{Id} subido por {User}", id, CurrentUser());
        return Ok(new { id });
    }

    // ── GET /api/procedimientos ───────────────────────────────────────────────
    /// <summary>Lista procedimientos con filtros opcionales.</summary>
    [HttpGet]
    public async Task<IActionResult> GetAll(
        [FromQuery] string? search,
        [FromQuery] string? category,
        CancellationToken ct)
    {
        await using var conn = Conn();
        await conn.OpenAsync(ct);
        await using var cmd = conn.CreateCommand();

        cmd.CommandText = """
            SELECT Id, Title, Description, Category,
                   FileName, FileContentType, FileSize, UploadedBy, UploadedAt
            FROM dbo.Procedimientos
            WHERE IsDeleted = 0
              AND (@Search IS NULL OR Title LIKE '%' + @Search + '%'
                                   OR Description LIKE '%' + @Search + '%')
              AND (@Cat IS NULL OR Category = @Cat)
            ORDER BY UploadedAt DESC
            """;
        cmd.Parameters.AddWithValue("@Search", string.IsNullOrWhiteSpace(search) ? (object)DBNull.Value : search.Trim());
        cmd.Parameters.AddWithValue("@Cat",    string.IsNullOrWhiteSpace(category) ? (object)DBNull.Value : category.Trim());

        var results = new List<object>();
        await using var r = await cmd.ExecuteReaderAsync(ct);
        while (await r.ReadAsync(ct))
        {
            results.Add(new
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
        }
        return Ok(results);
    }

    // ── GET /api/procedimientos/{id}/view ─────────────────────────────────────
    /// <summary>Devuelve el archivo para visualización inline (PDF/imagen).</summary>
    [HttpGet("{id:int}/view")]
    [AllowAnonymous] // acceso por token en query string
    public async Task<IActionResult> View(int id, [FromQuery] string? access_token, CancellationToken ct)
    {
        if (!await ValidateToken(access_token)) return Unauthorized();
        return await ServeFile(id, inline: true, ct);
    }

    // ── GET /api/procedimientos/{id}/download ─────────────────────────────────
    /// <summary>Descarga el archivo.</summary>
    [HttpGet("{id:int}/download")]
    [AllowAnonymous] // acceso por token en query string
    public async Task<IActionResult> Download(int id, [FromQuery] string? access_token, CancellationToken ct)
    {
        if (!await ValidateToken(access_token)) return Unauthorized();
        return await ServeFile(id, inline: false, ct);
    }

    // ── DELETE /api/procedimientos/{id} ──────────────────────────────────────
    /// <summary>Elimina un procedimiento (soft-delete). Solo Admin.</summary>
    [HttpDelete("{id:int}")]
    [Authorize(Roles = "Admin")]
    public async Task<IActionResult> Delete(int id, CancellationToken ct)
    {
        await using var conn = Conn();
        await conn.OpenAsync(ct);
        await using var cmd = conn.CreateCommand();
        cmd.CommandText = "UPDATE dbo.Procedimientos SET IsDeleted = 1 WHERE Id = @Id";
        cmd.Parameters.AddWithValue("@Id", id);
        int rows = await cmd.ExecuteNonQueryAsync(ct);
        if (rows == 0) return NotFound();
        logger.LogInformation("Procedimiento #{Id} eliminado por {User}", id, CurrentUser());
        return NoContent();
    }

    // ── Helpers ───────────────────────────────────────────────────────────────
    private async Task<IActionResult> ServeFile(int id, bool inline, CancellationToken ct)
    {
        string? diskName = null, contentType = null, fileName = null;

        await using var conn = Conn();
        await conn.OpenAsync(ct);
        await using var cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT FilePath, FileContentType, FileName FROM dbo.Procedimientos WHERE Id = @Id AND IsDeleted = 0";
        cmd.Parameters.AddWithValue("@Id", id);
        await using var r = await cmd.ExecuteReaderAsync(ct);
        if (await r.ReadAsync(ct))
        {
            diskName    = r.GetString(0);
            contentType = r.GetString(1);
            fileName    = r.GetString(2);
        }

        if (diskName is null) return NotFound();

        var path = Path.Combine(StorageDir, diskName);
        if (!System.IO.File.Exists(path)) return NotFound("Archivo no encontrado en disco.");

        var bytes = await System.IO.File.ReadAllBytesAsync(path, ct);
        Response.Headers["Content-Disposition"] = inline
            ? $"inline; filename=\"{fileName}\""
            : $"attachment; filename=\"{fileName}\"";
        return File(bytes, contentType ?? "application/octet-stream");
    }

    /// <summary>
    /// Valida un access_token pasado como query string.
    /// Si el usuario ya viene autenticado por header (bearer), retorna true directo.
    /// </summary>
    private Task<bool> ValidateToken(string? token)
    {
        if (User.Identity?.IsAuthenticated == true) return Task.FromResult(true);
        if (string.IsNullOrWhiteSpace(token))       return Task.FromResult(false);

        // Verificar que el token existe en nuestro almacén (tokens válidos están en localStorage)
        // Aquí hacemos una validación mínima: el JWT se valida implícitamente al parsearlo.
        // Para query-string usamos la misma clave JWT que el middleware.
        try
        {
            var jwtKey   = config["JwtSettings:Key"]!;
            var issuer   = config["JwtSettings:Issuer"]!;
            var audience = config["JwtSettings:Audience"]!;

            var handler  = new System.IdentityModel.Tokens.Jwt.JwtSecurityTokenHandler();
            var key      = new Microsoft.IdentityModel.Tokens.SymmetricSecurityKey(
                               System.Text.Encoding.UTF8.GetBytes(jwtKey));
            handler.ValidateToken(token, new Microsoft.IdentityModel.Tokens.TokenValidationParameters
            {
                ValidateIssuerSigningKey = true,
                IssuerSigningKey        = key,
                ValidateIssuer          = true,
                ValidIssuer             = issuer,
                ValidateAudience        = true,
                ValidAudience           = audience,
                ValidateLifetime        = true,
            }, out _);
            return Task.FromResult(true);
        }
        catch { return Task.FromResult(false); }
    }
}
