using BCrypt.Net;
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

    private Guid? TenantId =>
        Guid.TryParse(User.FindFirstValue("tenantId"), out var tid) ? tid : null;

    private static string HashPassword(string password) =>
        UserService.HashPassword(password);

    // List-view reader — does NOT include binary photo data (avoids loading MB per user in admin list)
    private static object ReadUserSummary(SqlDataReader r) => new
    {
        id              = r.GetGuid(r.GetOrdinal("Id")),
        username        = r.GetString(r.GetOrdinal("Username")),
        fullName        = r.IsDBNull(r.GetOrdinal("FullName"))        ? null : r.GetString(r.GetOrdinal("FullName")),
        email           = r.IsDBNull(r.GetOrdinal("Email"))           ? null : r.GetString(r.GetOrdinal("Email")),
        role            = r.IsDBNull(r.GetOrdinal("Role"))            ? "User" : r.GetString(r.GetOrdinal("Role")),
        modules         = r.GetInt32(r.GetOrdinal("Modules")),
        modulesViewOnly = r.IsDBNull(r.GetOrdinal("ModulesViewOnly")) ? 0 : r.GetInt32(r.GetOrdinal("ModulesViewOnly")),
        isActive        = r.GetBoolean(r.GetOrdinal("IsActive")),
        position        = r.IsDBNull(r.GetOrdinal("Position"))        ? null : r.GetString(r.GetOrdinal("Position")),
        smtpEmail       = r.IsDBNull(r.GetOrdinal("SmtpEmail"))       ? null : r.GetString(r.GetOrdinal("SmtpEmail")),
        profilePhotoUrl = r.IsDBNull(r.GetOrdinal("ProfilePhotoUrl")) ? null : r.GetString(r.GetOrdinal("ProfilePhotoUrl")),
        bannerPhotoUrl  = r.IsDBNull(r.GetOrdinal("BannerPhotoUrl"))  ? null : r.GetString(r.GetOrdinal("BannerPhotoUrl")),
        createdAt       = r.GetDateTime(r.GetOrdinal("CreatedAt")),
    };

    // Full reader — resolves binary photo data to data URL (used only for single-user endpoints)
    private static object ReadUser(SqlDataReader r)
    {
        // Helper: resolve photo URL — prefers binary (data URL) over legacy path URL
        static string? ResolvePhoto(SqlDataReader rd, string dataCol, string mimeCol, string urlCol)
        {
            int dataOrd = rd.GetOrdinal(dataCol);
            int mimeOrd = rd.GetOrdinal(mimeCol);
            int urlOrd  = rd.GetOrdinal(urlCol);
            if (!rd.IsDBNull(dataOrd))
            {
                var bytes = (byte[])rd.GetValue(dataOrd);
                var mime  = rd.IsDBNull(mimeOrd) ? "image/jpeg" : rd.GetString(mimeOrd);
                return $"data:{mime};base64,{Convert.ToBase64String(bytes)}";
            }
            return rd.IsDBNull(urlOrd) ? null : rd.GetString(urlOrd);
        }

        return new
        {
            id               = r.GetGuid(r.GetOrdinal("Id")),
            username         = r.GetString(r.GetOrdinal("Username")),
            fullName         = r.IsDBNull(r.GetOrdinal("FullName"))        ? null : r.GetString(r.GetOrdinal("FullName")),
            email            = r.IsDBNull(r.GetOrdinal("Email"))           ? null : r.GetString(r.GetOrdinal("Email")),
            role             = r.IsDBNull(r.GetOrdinal("Role"))            ? "User" : r.GetString(r.GetOrdinal("Role")),
            modules          = r.GetInt32(r.GetOrdinal("Modules")),
            modulesViewOnly  = r.IsDBNull(r.GetOrdinal("ModulesViewOnly")) ? 0 : r.GetInt32(r.GetOrdinal("ModulesViewOnly")),
            isActive         = r.GetBoolean(r.GetOrdinal("IsActive")),
            position         = r.IsDBNull(r.GetOrdinal("Position"))        ? null : r.GetString(r.GetOrdinal("Position")),
            smtpEmail        = r.IsDBNull(r.GetOrdinal("SmtpEmail"))       ? null : r.GetString(r.GetOrdinal("SmtpEmail")),
            profilePhotoUrl  = ResolvePhoto(r, "ProfilePhotoData", "ProfilePhotoMime", "ProfilePhotoUrl"),
            bannerPhotoUrl   = ResolvePhoto(r, "BannerPhotoData",  "BannerPhotoMime",  "BannerPhotoUrl"),
            createdAt        = r.GetDateTime(r.GetOrdinal("CreatedAt")),
        };
    }

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
            // Filtrar por tenant: cada admin solo ve usuarios de su propio cliente
            if (TenantId.HasValue)
            {
                cmd.CommandText = """
                    SELECT Id, Username, FullName, Email, Role, Modules, ModulesViewOnly, IsActive,
                           Position, SmtpEmail, ProfilePhotoUrl, BannerPhotoUrl, CreatedAt
                    FROM dbo.AppUsers
                    WHERE TenantId = @TenantId
                    ORDER BY CreatedAt DESC
                    """;
                cmd.Parameters.AddWithValue("@TenantId", TenantId.Value);
            }
            else
            {
                cmd.CommandText = """
                    SELECT Id, Username, FullName, Email, Role, Modules, ModulesViewOnly, IsActive,
                           Position, SmtpEmail, ProfilePhotoUrl, BannerPhotoUrl, CreatedAt
                    FROM dbo.AppUsers
                    ORDER BY CreatedAt DESC
                    """;
            }
            var list = new List<object>();
            await using var r = await cmd.ExecuteReaderAsync(ct);
            while (await r.ReadAsync(ct)) list.Add(ReadUserSummary(r));
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
                SELECT Id, Username, FullName, Email, Role, Modules, ModulesViewOnly, IsActive,
                       Position, SmtpEmail, ProfilePhotoUrl, BannerPhotoUrl, CreatedAt,
                       ProfilePhotoData, ProfilePhotoMime, BannerPhotoData, BannerPhotoMime
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
                    (Id, Username, FullName, Email, PasswordHash, Role, Modules, ModulesViewOnly,
                     IsActive, Position, TenantId, CreatedAt)
                VALUES
                    (@Id, @Username, @FullName, @Email, @Hash, @Role, @Modules, @ModulesViewOnly,
                     @IsActive, @Position, @TenantId, GETUTCDATE())
                """;
            cmd.Parameters.AddWithValue("@Id",              id);
            cmd.Parameters.AddWithValue("@Username",        dto.Username.Trim().ToLower());
            cmd.Parameters.AddWithValue("@FullName",        (object?)dto.FullName   ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@Email",           (object?)dto.Email       ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@Hash",            hash);
            cmd.Parameters.AddWithValue("@Role",            dto.Role ?? "User");
            cmd.Parameters.AddWithValue("@Modules",         dto.Modules);
            cmd.Parameters.AddWithValue("@ModulesViewOnly", dto.ModulesViewOnly);
            cmd.Parameters.AddWithValue("@IsActive",        dto.IsActive);
            cmd.Parameters.AddWithValue("@TenantId",        TenantId.HasValue ? (object)TenantId.Value : DBNull.Value);
            cmd.Parameters.AddWithValue("@Position",        (object?)dto.Position ?? DBNull.Value);
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
                "FullName        = @FullName",
                "Email           = @Email",
                "Role            = @Role",
                "Modules         = @Modules",
                "ModulesViewOnly = @ModulesViewOnly",
                "IsActive        = @IsActive",
                "Position        = @Position",
                "UpdatedAt       = GETUTCDATE()"
            };
            cmd.Parameters.AddWithValue("@FullName",        (object?)dto.FullName  ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@Email",           (object?)dto.Email      ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@Role",            dto.Role ?? "User");
            cmd.Parameters.AddWithValue("@Modules",         dto.Modules);
            cmd.Parameters.AddWithValue("@ModulesViewOnly", dto.ModulesViewOnly);
            cmd.Parameters.AddWithValue("@IsActive",        dto.IsActive);
            cmd.Parameters.AddWithValue("@Position",        (object?)dto.Position ?? DBNull.Value);

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
    /// <summary>
    /// Guarda la foto de perfil como VARBINARY en BD y la devuelve como data URL.
    /// Evita dependencia del filesystem efímero de Railway.
    /// </summary>
    [HttpPost("me/photo")]
    public async Task<IActionResult> UploadPhoto(IFormFile file, CancellationToken ct)
    {
        var username = CurrentUsername;
        if (string.IsNullOrWhiteSpace(username)) return Unauthorized();
        if (file == null || file.Length == 0) return BadRequest("Archivo requerido.");
        if (file.Length > 5 * 1024 * 1024) return BadRequest("La imagen no puede superar 5 MB.");

        var ext  = Path.GetExtension(file.FileName).ToLower();
        var mimeMap = new Dictionary<string, string>
        {
            { ".jpg",  "image/jpeg" }, { ".jpeg", "image/jpeg" },
            { ".png",  "image/png"  }, { ".webp", "image/webp" },
            { ".gif",  "image/gif"  },
        };
        if (!mimeMap.TryGetValue(ext, out var mime)) return BadRequest("Formato no soportado (.jpg, .png, .webp, .gif).");

        try
        {
            // Leer bytes
            using var ms = new MemoryStream();
            await file.CopyToAsync(ms, ct);
            var bytes = ms.ToArray();

            // Asegura que existe la columna ProfilePhotoData en AppUsers
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await using (var alter = conn.CreateCommand())
            {
                alter.CommandText = """
                    IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.AppUsers') AND name = 'ProfilePhotoData')
                        ALTER TABLE dbo.AppUsers ADD ProfilePhotoData VARBINARY(MAX) NULL;
                    IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.AppUsers') AND name = 'ProfilePhotoMime')
                        ALTER TABLE dbo.AppUsers ADD ProfilePhotoMime NVARCHAR(50) NULL;
                    """;
                await alter.ExecuteNonQueryAsync(ct);
            }

            await using var cmd = conn.CreateCommand();
            cmd.CommandText = """
                UPDATE dbo.AppUsers
                SET ProfilePhotoData = @Data, ProfilePhotoMime = @Mime,
                    ProfilePhotoUrl  = NULL,   UpdatedAt = GETUTCDATE()
                WHERE LOWER(Username) = LOWER(@User)
                """;
            cmd.Parameters.AddWithValue("@Data", bytes);
            cmd.Parameters.AddWithValue("@Mime", mime);
            cmd.Parameters.AddWithValue("@User", username);
            await cmd.ExecuteNonQueryAsync(ct);

            // Devolver data URL para que el frontend la use directamente
            var dataUrl = $"data:{mime};base64,{Convert.ToBase64String(bytes)}";
            return Ok(new { url = dataUrl });
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
                UPDATE dbo.AppUsers
                SET ProfilePhotoUrl = NULL, ProfilePhotoData = NULL, ProfilePhotoMime = NULL,
                    UpdatedAt = GETUTCDATE()
                WHERE LOWER(Username) = LOWER(@User)
                """;
            cmd.Parameters.AddWithValue("@User", username);
            await cmd.ExecuteNonQueryAsync(ct);
            return NoContent();
        }
        catch (Exception ex) { logger.LogError(ex, "DeletePhoto"); return StatusCode(500, ex.Message); }
    }

    // ── POST /api/users/me/banner ─────────────────────────────────────────────
    /// <summary>
    /// Guarda el banner como VARBINARY en BD y lo devuelve como data URL.
    /// </summary>
    [HttpPost("me/banner")]
    public async Task<IActionResult> UploadBanner(IFormFile file, CancellationToken ct)
    {
        var username = CurrentUsername;
        if (string.IsNullOrWhiteSpace(username)) return Unauthorized();
        if (file == null || file.Length == 0) return BadRequest("Archivo requerido.");
        if (file.Length > 8 * 1024 * 1024) return BadRequest("El banner no puede superar 8 MB.");

        var ext  = Path.GetExtension(file.FileName).ToLower();
        var mimeMap = new Dictionary<string, string>
        {
            { ".jpg",  "image/jpeg" }, { ".jpeg", "image/jpeg" },
            { ".png",  "image/png"  }, { ".webp", "image/webp" },
        };
        if (!mimeMap.TryGetValue(ext, out var mime)) return BadRequest("Formato no soportado (.jpg, .png, .webp).");

        try
        {
            using var ms = new MemoryStream();
            await file.CopyToAsync(ms, ct);
            var bytes = ms.ToArray();

            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await using (var alter = conn.CreateCommand())
            {
                alter.CommandText = """
                    IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.AppUsers') AND name = 'BannerPhotoData')
                        ALTER TABLE dbo.AppUsers ADD BannerPhotoData VARBINARY(MAX) NULL;
                    IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.AppUsers') AND name = 'BannerPhotoMime')
                        ALTER TABLE dbo.AppUsers ADD BannerPhotoMime NVARCHAR(50) NULL;
                    """;
                await alter.ExecuteNonQueryAsync(ct);
            }

            await using var cmd = conn.CreateCommand();
            cmd.CommandText = """
                UPDATE dbo.AppUsers
                SET BannerPhotoData = @Data, BannerPhotoMime = @Mime,
                    BannerPhotoUrl  = NULL,   UpdatedAt = GETUTCDATE()
                WHERE LOWER(Username) = LOWER(@User)
                """;
            cmd.Parameters.AddWithValue("@Data", bytes);
            cmd.Parameters.AddWithValue("@Mime", mime);
            cmd.Parameters.AddWithValue("@User", username);
            await cmd.ExecuteNonQueryAsync(ct);

            var dataUrl = $"data:{mime};base64,{Convert.ToBase64String(bytes)}";
            return Ok(new { url = dataUrl });
        }
        catch (Exception ex) { logger.LogError(ex, "UploadBanner"); return StatusCode(500, ex.Message); }
    }

    // ── PUT /api/users/me/change-password ────────────────────────────────────
    [HttpPut("me/change-password")]
    public async Task<IActionResult> ChangePassword([FromBody] ChangePasswordDto dto, CancellationToken ct)
    {
        var username = CurrentUsername;
        if (string.IsNullOrWhiteSpace(username)) return Unauthorized();
        if (string.IsNullOrWhiteSpace(dto.CurrentPassword)) return BadRequest("Contraseña actual requerida.");
        if (string.IsNullOrWhiteSpace(dto.NewPassword))     return BadRequest("Nueva contraseña requerida.");
        if (dto.NewPassword.Length < 8) return BadRequest("La nueva contraseña debe tener al menos 8 caracteres.");

        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);

            // Verificar contraseña actual
            string? hash = null;
            await using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = "SELECT PasswordHash FROM dbo.AppUsers WHERE LOWER(Username) = LOWER(@User) AND IsActive = 1";
                cmd.Parameters.AddWithValue("@User", username);
                hash = (string?)await cmd.ExecuteScalarAsync(ct);
            }

            if (hash == null) return Unauthorized();
            if (!BCrypt.Net.BCrypt.Verify(dto.CurrentPassword, hash))
                return BadRequest("La contraseña actual es incorrecta.");

            // Actualizar contraseña
            var newHash = BCrypt.Net.BCrypt.HashPassword(dto.NewPassword);
            await using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = "UPDATE dbo.AppUsers SET PasswordHash = @Hash, UpdatedAt = GETUTCDATE() WHERE LOWER(Username) = LOWER(@User)";
                cmd.Parameters.AddWithValue("@Hash", newHash);
                cmd.Parameters.AddWithValue("@User", username);
                await cmd.ExecuteNonQueryAsync(ct);
            }

            logger.LogInformation("Contraseña cambiada para {Username}", username);
            return NoContent();
        }
        catch (Exception ex) { logger.LogError(ex, "ChangePassword"); return StatusCode(500, ex.Message); }
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
                UPDATE dbo.AppUsers
                SET BannerPhotoUrl = NULL, BannerPhotoData = NULL, BannerPhotoMime = NULL,
                    UpdatedAt = GETUTCDATE()
                WHERE LOWER(Username) = LOWER(@User)
                """;
            cmd.Parameters.AddWithValue("@User", username);
            await cmd.ExecuteNonQueryAsync(ct);
            return NoContent();
        }
        catch (Exception ex) { logger.LogError(ex, "DeleteBanner"); return StatusCode(500, ex.Message); }
    }

    // ── POST /api/users/import-csv ───────────────────────────────────────────
    [HttpPost("import-csv")]
    [Authorize(Roles = "Admin")]
    public async Task<IActionResult> ImportCsv([FromBody] List<UserImportDto> users, CancellationToken ct)
    {
        if (users == null || users.Count == 0) return BadRequest("Lista vacía.");
        var created = 0; var skipped = 0; var errors = new List<string>();
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            foreach (var u in users)
            {
                if (string.IsNullOrWhiteSpace(u.Username) || string.IsNullOrWhiteSpace(u.Password))
                { skipped++; continue; }
                try
                {
                    await using var cmd = conn.CreateCommand();
                    cmd.CommandText = """
                        IF NOT EXISTS (SELECT 1 FROM dbo.AppUsers WHERE LOWER(Username) = LOWER(@User))
                        BEGIN
                          INSERT INTO dbo.AppUsers (Id, Username, FullName, Email, Position, PasswordHash, Role, Modules, IsActive, CreatedAt)
                          VALUES (NEWID(), @User, @FullName, @Email, @Position, @Hash, @Role, @Modules, 1, GETUTCDATE())
                        END
                        """;
                    cmd.Parameters.AddWithValue("@User",     u.Username.Trim().ToLower());
                    cmd.Parameters.AddWithValue("@FullName", string.IsNullOrWhiteSpace(u.FullName)   ? DBNull.Value : (object)u.FullName.Trim());
                    cmd.Parameters.AddWithValue("@Email",    string.IsNullOrWhiteSpace(u.Email)      ? DBNull.Value : (object)u.Email.Trim());
                    cmd.Parameters.AddWithValue("@Position", string.IsNullOrWhiteSpace(u.Position)   ? DBNull.Value : (object)u.Position.Trim());
                    cmd.Parameters.AddWithValue("@Hash",     HashPassword(u.Password));
                    cmd.Parameters.AddWithValue("@Role",     string.IsNullOrWhiteSpace(u.Role) ? "User" : u.Role.Trim());
                    cmd.Parameters.AddWithValue("@Modules",  u.Modules ?? 0);
                    int rows = await cmd.ExecuteNonQueryAsync(ct);
                    if (rows > 0) created++; else skipped++;
                }
                catch (Exception ex) { errors.Add($"{u.Username}: {ex.Message}"); }
            }
            return Ok(new { created, skipped, errors });
        }
        catch (Exception ex) { logger.LogError(ex, "ImportCsv"); return StatusCode(500, ex.Message); }
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
    int     ModulesViewOnly,
    bool    IsActive,
    string? Position
);

public record UserUpdateDto(
    string? FullName,
    string? Email,
    string? Password,
    string? Role,
    int     Modules,
    int     ModulesViewOnly,
    bool    IsActive,
    string? Position
);

public record SmtpDto(string? SmtpEmail, string? SmtpPassword);
public record ChangePasswordDto(string CurrentPassword, string NewPassword);
public record UserImportDto(string Username, string? FullName, string? Email, string? Position, string Password, string? Role, int? Modules);
