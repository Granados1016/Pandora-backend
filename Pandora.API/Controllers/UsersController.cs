using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using Pandora.Application.Features.Users;
using System.Security.Claims;

namespace Pandora.API.Controllers;

[ApiController]
[Route("api/users")]
[Authorize]
public class UsersController(IConfiguration config, ILogger<UsersController> logger) : ControllerBase
{
    // ── Helpers ───────────────────────────────────────────────────────────────
    private SqlConnection Conn() => new(config.GetConnectionString("PandoraDb"));

    private string? CurrentUsername =>
        User.FindFirstValue(ClaimTypes.Name) ??
        User.FindFirstValue("name") ??
        User.Claims.FirstOrDefault(c => c.Type.EndsWith("name", StringComparison.OrdinalIgnoreCase))?.Value;

    private bool IsAdmin =>
        User.IsInRole("Admin") ||
        User.Claims.Any(c => (c.Type.EndsWith("role", StringComparison.OrdinalIgnoreCase) ||
                               c.Type.EndsWith("roles", StringComparison.OrdinalIgnoreCase)) &&
                              c.Value.Equals("Admin", StringComparison.OrdinalIgnoreCase));

    private static string HashPassword(string password) =>
        UserService.HashPassword(password);

    // Serialize user row from reader
    private static object ReadUser(SqlDataReader r) => new
    {
        id             = r.GetGuid(r.GetOrdinal("Id")),
        username       = r.GetString(r.GetOrdinal("Username")),
        fullName       = r.IsDBNull(r.GetOrdinal("FullName"))   ? null : r.GetString(r.GetOrdinal("FullName")),
        email          = r.IsDBNull(r.GetOrdinal("Email"))      ? null : r.GetString(r.GetOrdinal("Email")),
        role           = r.IsDBNull(r.GetOrdinal("Role"))       ? "User" : r.GetString(r.GetOrdinal("Role")),
        modules        = r.GetInt32(r.GetOrdinal("Modules")),
        isActive       = r.GetBoolean(r.GetOrdinal("IsActive")),
        position       = r.IsDBNull(r.GetOrdinal("Position"))        ? null : r.GetString(r.GetOrdinal("Position")),
        smtpEmail      = r.IsDBNull(r.GetOrdinal("SmtpEmail"))        ? null : r.GetString(r.GetOrdinal("SmtpEmail")),
        profilePhotoUrl= r.IsDBNull(r.GetOrdinal("ProfilePhotoUrl"))  ? null : r.GetString(r.GetOrdinal("ProfilePhotoUrl")),
        bannerPhotoUrl = r.IsDBNull(r.GetOrdinal("BannerPhotoUrl"))   ? null : r.GetString(r.GetOrdinal("BannerPhotoUrl")),
        createdAt      = r.GetDateTime(r.GetOrdinal("CreatedAt")),
    };

    // ── GET /api/users  (Admin) ───────────────────────────────────────────────
    [HttpGet]
    public async Task<IActionResult> GetAll(CancellationToken ct)
    {
        if (!IsAdmin) return Forbid();
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = """
                SELECT Id, Username, FullName, Email, Role, Modules, IsActive,
                       Position, SmtpEmail, ProfilePhotoUrl, BannerPhotoUrl, CreatedAt
                FROM dbo.AppUsers
                ORDER BY CreatedAt DESC
                """;
            var list = new List<object>();
            await using var r = await cmd.ExecuteReaderAsync(ct);
            while (await r.ReadAsync(ct)) list.Add(ReadUser(r));
            return Ok(list);
        }
        catch (Exception ex) { logger.LogError(ex, "GetAll Users"); return StatusCode(500, ex.Message); }
    }

    // ── GET /api/users/me ─────────────────────────────────────────────────────
    [HttpGet("me")]
    public async Task<IActionResult> Me(CancellationToken ct)
    {
        var username = CurrentUsername;
        if (string.IsNullOrWhiteSpace(username)) return Unauthorized();
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = """
                SELECT Id, Username, FullName, Email, Role, Modules, IsActive,
                       Position, SmtpEmail, ProfilePhotoUrl, BannerPhotoUrl, CreatedAt
                FROM dbo.AppUsers WHERE LOWER(Username) = LOWER(@User)
                """;
            cmd.Parameters.AddWithValue("@User", username);
            await using var r = await cmd.ExecuteReaderAsync(ct);
            if (!await r.ReadAsync(ct)) return NotFound("Usuario no encontrado.");
            return Ok(ReadUser(r));
        }
        catch (Exception ex) { logger.LogError(ex, "Me"); return StatusCode(500, ex.Message); }
    }

    // ── POST /api/users  (Admin) ──────────────────────────────────────────────
    [HttpPost]
    public async Task<IActionResult> Create([FromBody] UserCreateDto dto, CancellationToken ct)
    {
        if (!IsAdmin) return Forbid();
        if (string.IsNullOrWhiteSpace(dto.Username)) return BadRequest("Usuario requerido.");
        if (string.IsNullOrWhiteSpace(dto.Password)) return BadRequest("Contraseña requerida.");

        var id   = Guid.NewGuid();
        var hash = HashPassword(dto.Password);
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = """
                INSERT INTO dbo.AppUsers
                    (Id, Username, FullName, Email, PasswordHash, Role, Modules,
                     IsActive, Position, CreatedAt)
                VALUES
                    (@Id, @Username, @FullName, @Email, @Hash, @Role, @Modules,
                     @IsActive, @Position, GETUTCDATE())
                """;
            cmd.Parameters.AddWithValue("@Id",       id);
            cmd.Parameters.AddWithValue("@Username", dto.Username.Trim().ToLower());
            cmd.Parameters.AddWithValue("@FullName", (object?)dto.FullName   ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@Email",    (object?)dto.Email       ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@Hash",     hash);
            cmd.Parameters.AddWithValue("@Role",     dto.Role ?? "User");
            cmd.Parameters.AddWithValue("@Modules",  dto.Modules);
            cmd.Parameters.AddWithValue("@IsActive", dto.IsActive);
            cmd.Parameters.AddWithValue("@Position", (object?)dto.Position ?? DBNull.Value);
            await cmd.ExecuteNonQueryAsync(ct);
            return Ok(new { id });
        }
        catch (SqlException ex) when (ex.Number == 2627)
        {
            return Conflict("El nombre de usuario ya existe.");
        }
        catch (Exception ex) { logger.LogError(ex, "Create User"); return StatusCode(500, ex.Message); }
    }

    // ── PUT /api/users/{id}  (Admin) ──────────────────────────────────────────
    [HttpPut("{id:guid}")]
    public async Task<IActionResult> Update(Guid id, [FromBody] UserUpdateDto dto, CancellationToken ct)
    {
        if (!IsAdmin) return Forbid();
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await using var cmd = conn.CreateCommand();

            // Construir SET dinámico para contraseña opcional
            var setParts = new List<string>
            {
                "FullName   = @FullName",
                "Email      = @Email",
                "Role       = @Role",
                "Modules    = @Modules",
                "IsActive   = @IsActive",
                "Position   = @Position",
                "UpdatedAt  = GETUTCDATE()"
            };
            cmd.Parameters.AddWithValue("@FullName",  (object?)dto.FullName  ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@Email",     (object?)dto.Email      ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@Role",      dto.Role ?? "User");
            cmd.Parameters.AddWithValue("@Modules",   dto.Modules);
            cmd.Parameters.AddWithValue("@IsActive",  dto.IsActive);
            cmd.Parameters.AddWithValue("@Position",  (object?)dto.Position ?? DBNull.Value);

            if (!string.IsNullOrWhiteSpace(dto.Password))
            {
                setParts.Insert(0, "PasswordHash = @Hash");
                cmd.Parameters.AddWithValue("@Hash", HashPassword(dto.Password));
            }

            cmd.CommandText = $"UPDATE dbo.AppUsers SET {string.Join(", ", setParts)} WHERE Id = @Id";
            cmd.Parameters.AddWithValue("@Id", id);

            int rows = await cmd.ExecuteNonQueryAsync(ct);
            if (rows == 0) return NotFound("Usuario no encontrado.");
            return Ok(new { id });
        }
        catch (Exception ex) { logger.LogError(ex, "Update User {Id}", id); return StatusCode(500, ex.Message); }
    }

    // ── DELETE /api/users/{id}  (Admin) ──────────────────────────────────────
    [HttpDelete("{id:guid}")]
    public async Task<IActionResult> Delete(Guid id, CancellationToken ct)
    {
        if (!IsAdmin) return Forbid();
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = "DELETE FROM dbo.AppUsers WHERE Id = @Id";
            cmd.Parameters.AddWithValue("@Id", id);
            int rows = await cmd.ExecuteNonQueryAsync(ct);
            if (rows == 0) return NotFound("Usuario no encontrado.");
            return NoContent();
        }
        catch (Exception ex) { logger.LogError(ex, "Delete User {Id}", id); return StatusCode(500, ex.Message); }
    }

    // ── PUT /api/users/me/smtp ────────────────────────────────────────────────
    [HttpPut("me/smtp")]
    public async Task<IActionResult> UpdateSmtp([FromBody] SmtpDto dto, CancellationToken ct)
    {
        var username = CurrentUsername;
        if (string.IsNullOrWhiteSpace(username)) return Unauthorized();
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await using var cmd = conn.CreateCommand();

            if (string.IsNullOrWhiteSpace(dto.SmtpPassword))
            {
                // No cambiar contraseña si viene vacío y ya hay una guardada
                cmd.CommandText = """
                    UPDATE dbo.AppUsers
                    SET SmtpEmail = @Email, UpdatedAt = GETUTCDATE()
                    WHERE LOWER(Username) = LOWER(@User)
                    """;
            }
            else
            {
                cmd.CommandText = """
                    UPDATE dbo.AppUsers
                    SET SmtpEmail = @Email, SmtpPassword = @Pass, UpdatedAt = GETUTCDATE()
                    WHERE LOWER(Username) = LOWER(@User)
                    """;
                cmd.Parameters.AddWithValue("@Pass", (object?)dto.SmtpPassword ?? DBNull.Value);
            }

            cmd.Parameters.AddWithValue("@Email", (object?)dto.SmtpEmail ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@User",  username);
            await cmd.ExecuteNonQueryAsync(ct);
            return Ok();
        }
        catch (Exception ex) { logger.LogError(ex, "UpdateSmtp"); return StatusCode(500, ex.Message); }
    }

    // ── POST /api/users/me/photo ──────────────────────────────────────────────
    [HttpPost("me/photo")]
    public async Task<IActionResult> UploadPhoto(IFormFile file, CancellationToken ct)
    {
        var username = CurrentUsername;
        if (string.IsNullOrWhiteSpace(username)) return Unauthorized();
        if (file == null || file.Length == 0) return BadRequest("Archivo requerido.");

        try
        {
            var ext  = Path.GetExtension(file.FileName).ToLower();
            var safe = new[] { ".jpg", ".jpeg", ".png", ".webp" };
            if (!safe.Contains(ext)) return BadRequest("Formato no soportado.");

            var dir  = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "uploads", "profiles");
            Directory.CreateDirectory(dir);
            var name = $"{Guid.NewGuid()}{ext}";
            var path = Path.Combine(dir, name);

            await using (var stream = System.IO.File.Create(path))
                await file.CopyToAsync(stream, ct);

            var url  = $"/uploads/profiles/{name}";

            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = """
                UPDATE dbo.AppUsers SET ProfilePhotoUrl = @Url, UpdatedAt = GETUTCDATE()
                WHERE LOWER(Username) = LOWER(@User)
                """;
            cmd.Parameters.AddWithValue("@Url",  url);
            cmd.Parameters.AddWithValue("@User", username);
            await cmd.ExecuteNonQueryAsync(ct);

            return Ok(new { url });
        }
        catch (Exception ex) { logger.LogError(ex, "UploadPhoto"); return StatusCode(500, ex.Message); }
    }

    // ── DELETE /api/users/me/photo ────────────────────────────────────────────
    [HttpDelete("me/photo")]
    public async Task<IActionResult> DeletePhoto(CancellationToken ct)
    {
        var username = CurrentUsername;
        if (string.IsNullOrWhiteSpace(username)) return Unauthorized();
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = """
                UPDATE dbo.AppUsers SET ProfilePhotoUrl = NULL, UpdatedAt = GETUTCDATE()
                WHERE LOWER(Username) = LOWER(@User)
                """;
            cmd.Parameters.AddWithValue("@User", username);
            await cmd.ExecuteNonQueryAsync(ct);
            return NoContent();
        }
        catch (Exception ex) { logger.LogError(ex, "DeletePhoto"); return StatusCode(500, ex.Message); }
    }

    // ── POST /api/users/me/banner ─────────────────────────────────────────────
    [HttpPost("me/banner")]
    public async Task<IActionResult> UploadBanner(IFormFile file, CancellationToken ct)
    {
        var username = CurrentUsername;
        if (string.IsNullOrWhiteSpace(username)) return Unauthorized();
        if (file == null || file.Length == 0) return BadRequest("Archivo requerido.");
        try
        {
            var ext  = Path.GetExtension(file.FileName).ToLower();
            var safe = new[] { ".jpg", ".jpeg", ".png", ".webp" };
            if (!safe.Contains(ext)) return BadRequest("Formato no soportado.");

            var dir  = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "uploads", "banners");
            Directory.CreateDirectory(dir);
            var name = $"{Guid.NewGuid()}{ext}";
            var path = Path.Combine(dir, name);

            await using (var stream = System.IO.File.Create(path))
                await file.CopyToAsync(stream, ct);

            var url  = $"/uploads/banners/{name}";

            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = """
                UPDATE dbo.AppUsers SET BannerPhotoUrl = @Url, UpdatedAt = GETUTCDATE()
                WHERE LOWER(Username) = LOWER(@User)
                """;
            cmd.Parameters.AddWithValue("@Url",  url);
            cmd.Parameters.AddWithValue("@User", username);
            await cmd.ExecuteNonQueryAsync(ct);
            return Ok(new { url });
        }
        catch (Exception ex) { logger.LogError(ex, "UploadBanner"); return StatusCode(500, ex.Message); }
    }

    // ── DELETE /api/users/me/banner ───────────────────────────────────────────
    [HttpDelete("me/banner")]
    public async Task<IActionResult> DeleteBanner(CancellationToken ct)
    {
        var username = CurrentUsername;
        if (string.IsNullOrWhiteSpace(username)) return Unauthorized();
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = """
                UPDATE dbo.AppUsers SET BannerPhotoUrl = NULL, UpdatedAt = GETUTCDATE()
                WHERE LOWER(Username) = LOWER(@User)
                """;
            cmd.Parameters.AddWithValue("@User", username);
            await cmd.ExecuteNonQueryAsync(ct);
            return NoContent();
        }
        catch (Exception ex) { logger.LogError(ex, "DeleteBanner"); return StatusCode(500, ex.Message); }
    }
}

// ── DTOs ──────────────────────────────────────────────────────────────────────
public record UserCreateDto(
    string  Username,
    string? FullName,
    string? Email,
    string  Password,
    string? Role,
    int     Modules,
    bool    IsActive,
    string? Position
);

public record UserUpdateDto(
    string? FullName,
    string? Email,
    string? Password,
    string? Role,
    int     Modules,
    bool    IsActive,
    string? Position
);

public record SmtpDto(string? SmtpEmail, string? SmtpPassword);
