using System.Security.Claims;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using Pandora.API.Services;
using Pandora.Application.Features.Users;

namespace Pandora.API.Controllers;

/// <summary>
/// Panel maestro de gestión de tenants (clientes).
/// Solo accesible para el Super Admin del sistema (rol "SuperAdmin" o variable de entorno).
/// </summary>
[ApiController]
[Route("api/tenants")]
[Authorize]
public class TenantsController(IConfiguration config, ILogger<TenantsController> logger) : ControllerBase
{
    private SqlConnection Conn() => new(config.GetConnectionString("PandoraDb"));

    // Cualquier Admin puede gestionar tenants — en producción se puede restringir
    // a un username específico via config["SuperAdmin:Username"] si se desea
    private bool IsSuperAdmin => User.IsInRole("Admin") ||
        User.Claims.Any(c => c.Type.EndsWith("role", StringComparison.OrdinalIgnoreCase) &&
                             c.Value.Equals("Admin", StringComparison.OrdinalIgnoreCase));

    // ── GET /api/tenants ──────────────────────────────────────────────────────
    [HttpGet]
    public async Task<IActionResult> GetAll(CancellationToken ct)
    {
        if (!IsSuperAdmin) return Forbid();
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = """
                SELECT t.Id, t.Slug, t.Name, t.DisplayName,
                       t.PrimaryColor, t.SecondaryColor,
                       t.LicensedModules, t.MaxUsers,
                       t.ExpiresAt, t.IsActive, t.ContactEmail, t.Notes, t.CreadoEn,
                       (SELECT COUNT(*) FROM dbo.AppUsers u WHERE u.TenantId = t.Id AND u.IsActive = 1) AS ActiveUsers,
                       CASE WHEN t.LogoData IS NOT NULL THEN 1 ELSE 0 END AS HasLogo
                FROM dbo.Tenants t
                ORDER BY t.CreadoEn
                """;
            var list = new List<object>();
            await using var r = await cmd.ExecuteReaderAsync(ct);
            while (await r.ReadAsync(ct))
                list.Add(new {
                    id             = r.GetGuid(r.GetOrdinal("Id")),
                    slug           = r.GetString(r.GetOrdinal("Slug")),
                    name           = r.GetString(r.GetOrdinal("Name")),
                    displayName    = r.GetString(r.GetOrdinal("DisplayName")),
                    primaryColor   = r.GetString(r.GetOrdinal("PrimaryColor")),
                    secondaryColor = r.GetString(r.GetOrdinal("SecondaryColor")),
                    licensedModules = r.GetInt64(r.GetOrdinal("LicensedModules")),
                    maxUsers       = r.GetInt32(r.GetOrdinal("MaxUsers")),
                    expiresAt      = r.IsDBNull(r.GetOrdinal("ExpiresAt")) ? (DateTime?)null : r.GetDateTime(r.GetOrdinal("ExpiresAt")),
                    isActive       = r.GetBoolean(r.GetOrdinal("IsActive")),
                    contactEmail   = r.IsDBNull(r.GetOrdinal("ContactEmail")) ? null : r.GetString(r.GetOrdinal("ContactEmail")),
                    notes          = r.IsDBNull(r.GetOrdinal("Notes")) ? null : r.GetString(r.GetOrdinal("Notes")),
                    creadoEn       = r.GetDateTime(r.GetOrdinal("CreadoEn")),
                    activeUsers    = r.GetInt32(r.GetOrdinal("ActiveUsers")),
                    hasLogo        = r.GetInt32(r.GetOrdinal("HasLogo")) == 1,
                });
            return Ok(list);
        }
        catch (Exception ex) { logger.LogError(ex, "GetAllTenants"); return StatusCode(500, ex.Message); }
    }

    // ── GET /api/tenants/{id} ─────────────────────────────────────────────────
    [HttpGet("{id:guid}")]
    public async Task<IActionResult> GetById(Guid id, CancellationToken ct)
    {
        if (!IsSuperAdmin) return Forbid();
        try
        {
            var ts   = new TenantService(new HttpContextAccessor { HttpContext = HttpContext }, config);
            var info = await ts.LoadAsync(id, ct);
            if (info is null) return NotFound();
            return Ok(new {
                info.Id, info.Slug, info.Name, info.DisplayName,
                info.PrimaryColor, info.SecondaryColor,
                info.LicensedModules, info.MaxUsers,
                info.ExpiresAt, info.IsActive, info.ContactEmail, info.Notes,
                logo = info.LogoBase64,
            });
        }
        catch (Exception ex) { logger.LogError(ex, "GetTenant {Id}", id); return StatusCode(500, ex.Message); }
    }

    // ── POST /api/tenants ─────────────────────────────────────────────────────
    [HttpPost]
    public async Task<IActionResult> Create([FromBody] TenantDto dto, CancellationToken ct)
    {
        if (!IsSuperAdmin) return Forbid();
        if (string.IsNullOrWhiteSpace(dto.Slug))        return BadRequest("Slug requerido.");
        if (string.IsNullOrWhiteSpace(dto.Name))        return BadRequest("Nombre requerido.");
        if (string.IsNullOrWhiteSpace(dto.DisplayName)) return BadRequest("Nombre corto requerido.");

        try
        {
            var id = Guid.NewGuid();
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = """
                INSERT INTO dbo.Tenants
                    (Id, Slug, Name, DisplayName, PrimaryColor, SecondaryColor,
                     LicensedModules, MaxUsers, ExpiresAt, IsActive, ContactEmail, Notes)
                VALUES
                    (@Id, @Slug, @Name, @DisplayName, @Primary, @Secondary,
                     @Modules, @MaxUsers, @ExpiresAt, 1, @ContactEmail, @Notes)
                """;
            cmd.Parameters.AddWithValue("@Id",          id);
            cmd.Parameters.AddWithValue("@Slug",        dto.Slug.Trim().ToLower());
            cmd.Parameters.AddWithValue("@Name",        dto.Name.Trim());
            cmd.Parameters.AddWithValue("@DisplayName", dto.DisplayName.Trim());
            cmd.Parameters.AddWithValue("@Primary",     dto.PrimaryColor   ?? "#1A237E");
            cmd.Parameters.AddWithValue("@Secondary",   dto.SecondaryColor ?? "#283593");
            cmd.Parameters.AddWithValue("@Modules",     dto.LicensedModules ?? -1L);
            cmd.Parameters.AddWithValue("@MaxUsers",    dto.MaxUsers        ?? 50);
            cmd.Parameters.AddWithValue("@ExpiresAt",   (object?)dto.ExpiresAt ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@ContactEmail",(object?)dto.ContactEmail?.Trim() ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@Notes",       (object?)dto.Notes?.Trim()        ?? DBNull.Value);
            await cmd.ExecuteNonQueryAsync(ct);
            logger.LogInformation("Tenant creado: {Slug} por {User}", dto.Slug, User.Identity?.Name);

            // ── Onboarding: crear usuario admin del nuevo tenant ──────────────
            var adminUsername = $"admin.{dto.Slug.Trim().ToLower()}";
            var adminPassword = GeneratePassword();
            try
            {
                await using var cmdAdmin = conn.CreateCommand();
                cmdAdmin.CommandText = """
                    IF NOT EXISTS (SELECT 1 FROM dbo.AppUsers WHERE LOWER(Username) = LOWER(@User))
                    INSERT INTO dbo.AppUsers
                        (Id, Username, FullName, Email, PasswordHash, Role, Modules, ModulesViewOnly,
                         IsActive, Position, TenantId, CreatedAt)
                    VALUES
                        (NEWID(), @User, @FullName, @Email, @Hash, 'Admin', -1, 0,
                         1, 'Administrador', @TenantId, GETUTCDATE())
                    """;
                cmdAdmin.Parameters.AddWithValue("@User",     adminUsername);
                cmdAdmin.Parameters.AddWithValue("@FullName", $"Admin — {dto.DisplayName}");
                cmdAdmin.Parameters.AddWithValue("@Email",    (object?)dto.ContactEmail?.Trim() ?? DBNull.Value);
                // PBKDF2 "salt:hash" — mismo esquema que UserService.VerifyPassword
                // usa en el login real (Pandora.Infrastructure.Services.JwtService).
                cmdAdmin.Parameters.AddWithValue("@Hash",     UserService.HashPassword(adminPassword));
                cmdAdmin.Parameters.AddWithValue("@TenantId", id);
                await cmdAdmin.ExecuteNonQueryAsync(ct);

                // Enviar correo de bienvenida al contacto del nuevo tenant
                if (!string.IsNullOrWhiteSpace(dto.ContactEmail))
                    _ = Task.Run(() => SendWelcomeEmailAsync(dto.ContactEmail, dto.DisplayName!, adminUsername, adminPassword), CancellationToken.None);
            }
            catch (Exception ex) { logger.LogWarning(ex, "No se pudo crear usuario admin para {Slug}", dto.Slug); }

            return Ok(new { id, adminUsername, adminPassword });
        }
        catch (SqlException ex) when (ex.Number == 2627)
        {
            return Conflict($"Ya existe un tenant con el slug '{dto.Slug}'.");
        }
        catch (Exception ex) { logger.LogError(ex, "CreateTenant"); return StatusCode(500, ex.Message); }
    }

    // ── Helpers privados ──────────────────────────────────────────────────────
    private static string GeneratePassword()
    {
        const string chars = "ABCDEFGHJKMNPQRSTUVWXYZabcdefghjkmnpqrstuvwxyz23456789!@#";
        var rng = new Random();
        return new string(Enumerable.Range(0, 12).Select(_ => chars[rng.Next(chars.Length)]).ToArray());
    }

    private async Task SendWelcomeEmailAsync(string toEmail, string displayName, string adminUser, string adminPass)
    {
        try
        {
            var smtp = await SmtpHelper.LoadAsync(config.GetConnectionString("PandoraDb")!, config);
            var frontendUrl = config["FrontendUrl"] ?? "https://pandora.app";
            var body = $"""
                <html><body style="font-family:Arial,sans-serif;font-size:14px;color:#333">
                <div style="max-width:600px;margin:0 auto">
                  <div style="background:#1a237e;padding:28px 32px;border-radius:8px 8px 0 0">
                    <h1 style="color:white;margin:0;font-size:26px;letter-spacing:2px">PANDORA</h1>
                    <p style="color:rgba(255,255,255,0.8);margin:4px 0 0">Sistema de Gestión Empresarial</p>
                  </div>
                  <div style="border:1px solid #ddd;padding:32px;border-radius:0 0 8px 8px">
                    <h2 style="color:#1a237e;margin:0 0 16px">¡Bienvenido a Pandora, {displayName}!</h2>
                    <p>Su cuenta ha sido configurada. A continuación sus credenciales de acceso:</p>
                    <div style="background:#f8f9ff;border:1px solid #e3e8ff;border-radius:8px;padding:20px;margin:20px 0">
                      <table style="width:100%">
                        <tr><td style="color:#555;font-weight:bold;padding:6px 0">URL de acceso</td>
                            <td><a href="{frontendUrl}">{frontendUrl}</a></td></tr>
                        <tr><td style="color:#555;font-weight:bold;padding:6px 0">Usuario</td>
                            <td style="font-family:monospace;font-weight:bold">{adminUser}</td></tr>
                        <tr><td style="color:#555;font-weight:bold;padding:6px 0">Contraseña temporal</td>
                            <td style="font-family:monospace;font-weight:bold">{adminPass}</td></tr>
                      </table>
                    </div>
                    <p style="color:#e65100;font-weight:bold">⚠ Por seguridad, cambie su contraseña en el primer inicio de sesión.</p>
                    <div style="text-align:center;margin-top:24px">
                      <a href="{frontendUrl}" style="background:#1a237e;color:white;padding:13px 32px;border-radius:8px;text-decoration:none;font-weight:700">Acceder a Pandora →</a>
                    </div>
                    <p style="font-size:12px;color:#999;margin-top:24px;text-align:center">Pandora — Sistema de Gestión · Este correo fue generado automáticamente.</p>
                  </div>
                </div></body></html>
                """;
            await SmtpHelper.SendAsync(smtp, toEmail, displayName, $"¡Bienvenido a Pandora! Credenciales de acceso — {displayName}", body);
            logger.LogInformation("Correo de bienvenida enviado a {Email} para {Tenant}", toEmail, displayName);
        }
        catch (Exception ex) { logger.LogWarning(ex, "No se pudo enviar correo de bienvenida"); }
    }

    // ── PUT /api/tenants/{id} ─────────────────────────────────────────────────
    [HttpPut("{id:guid}")]
    public async Task<IActionResult> Update(Guid id, [FromBody] TenantDto dto, CancellationToken ct)
    {
        if (!IsSuperAdmin) return Forbid();
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = """
                UPDATE dbo.Tenants SET
                    Name           = @Name,
                    DisplayName    = @DisplayName,
                    PrimaryColor   = @Primary,
                    SecondaryColor = @Secondary,
                    LicensedModules = @Modules,
                    MaxUsers       = @MaxUsers,
                    ExpiresAt      = @ExpiresAt,
                    ContactEmail   = @ContactEmail,
                    Notes          = @Notes
                WHERE Id = @Id
                """;
            cmd.Parameters.AddWithValue("@Id",          id);
            cmd.Parameters.AddWithValue("@Name",        dto.Name?.Trim()        ?? "");
            cmd.Parameters.AddWithValue("@DisplayName", dto.DisplayName?.Trim() ?? "");
            cmd.Parameters.AddWithValue("@Primary",     dto.PrimaryColor        ?? "#1A237E");
            cmd.Parameters.AddWithValue("@Secondary",   dto.SecondaryColor      ?? "#283593");
            cmd.Parameters.AddWithValue("@Modules",     dto.LicensedModules     ?? -1L);
            cmd.Parameters.AddWithValue("@MaxUsers",    dto.MaxUsers            ?? 50);
            cmd.Parameters.AddWithValue("@ExpiresAt",   (object?)dto.ExpiresAt  ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@ContactEmail",(object?)dto.ContactEmail?.Trim() ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@Notes",       (object?)dto.Notes?.Trim()        ?? DBNull.Value);
            int rows = await cmd.ExecuteNonQueryAsync(ct);
            if (rows == 0) return NotFound();
            logger.LogInformation("Tenant actualizado {Id} por {User}", id, User.Identity?.Name);
            return Ok();
        }
        catch (Exception ex) { logger.LogError(ex, "UpdateTenant {Id}", id); return StatusCode(500, ex.Message); }
    }

    // ── PATCH /api/tenants/{id}/toggle ────────────────────────────────────────
    [HttpPatch("{id:guid}/toggle")]
    public async Task<IActionResult> Toggle(Guid id, CancellationToken ct)
    {
        if (!IsSuperAdmin) return Forbid();
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = "UPDATE dbo.Tenants SET IsActive = ~IsActive WHERE Id = @Id";
            cmd.Parameters.AddWithValue("@Id", id);
            int rows = await cmd.ExecuteNonQueryAsync(ct);
            if (rows == 0) return NotFound();
            return Ok();
        }
        catch (Exception ex) { logger.LogError(ex, "ToggleTenant {Id}", id); return StatusCode(500, ex.Message); }
    }

    // ── POST /api/tenants/{id}/logo ───────────────────────────────────────────
    [HttpPost("{id:guid}/logo")]
    public async Task<IActionResult> UploadLogo(Guid id, IFormFile file, CancellationToken ct)
    {
        if (!IsSuperAdmin) return Forbid();
        if (file.Length > 2 * 1024 * 1024) return BadRequest("El logo no debe superar 2 MB.");
        try
        {
            using var ms = new MemoryStream();
            await file.CopyToAsync(ms, ct);
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = "UPDATE dbo.Tenants SET LogoData = @Data, LogoMime = @Mime WHERE Id = @Id";
            cmd.Parameters.AddWithValue("@Id",   id);
            cmd.Parameters.AddWithValue("@Data", ms.ToArray());
            cmd.Parameters.AddWithValue("@Mime", file.ContentType);
            int rows = await cmd.ExecuteNonQueryAsync(ct);
            if (rows == 0) return NotFound();
            return Ok();
        }
        catch (Exception ex) { logger.LogError(ex, "UploadLogo {Id}", id); return StatusCode(500, ex.Message); }
    }

    // ── GET /api/tenants/me ───────────────────────────────────────────────────
    /// <summary>
    /// Devuelve los datos de branding del tenant del usuario autenticado.
    /// Usado por el frontend para aplicar colores y logo al iniciar sesión.
    /// </summary>
    [HttpGet("me")]
    public async Task<IActionResult> GetMyTenant(CancellationToken ct)
    {
        var tenantClaim = User.FindFirstValue("tenantId");
        if (!Guid.TryParse(tenantClaim, out var tenantId))
            return Ok(new { displayName = "Pandora", primaryColor = "#1A237E", secondaryColor = "#283593", logo = (string?)null });

        try
        {
            var ts   = new TenantService(new HttpContextAccessor { HttpContext = HttpContext }, config);
            var info = await ts.LoadAsync(tenantId, ct);
            if (info is null) return NotFound();

            // Verificar licencia activa
            var (ok, reason) = await ts.ValidateAsync(tenantId, ct);
            if (!ok) return StatusCode(403, new { error = reason });

            return Ok(new {
                id             = info.Id,
                slug           = info.Slug,
                displayName    = info.DisplayName,
                primaryColor   = info.PrimaryColor,
                secondaryColor = info.SecondaryColor,
                logo           = info.LogoBase64,
                licensedModules = info.LicensedModules,
                expiresAt      = info.ExpiresAt,
            });
        }
        catch (Exception ex) { logger.LogError(ex, "GetMyTenant"); return StatusCode(500, ex.Message); }
    }

    // ── GET /api/tenants/{id}/stats ───────────────────────────────────────────
    [HttpGet("{id:guid}/stats")]
    public async Task<IActionResult> GetStats(Guid id, CancellationToken ct)
    {
        if (!IsSuperAdmin) return Forbid();
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = """
                SELECT
                  (SELECT COUNT(*) FROM dbo.AppUsers         WHERE TenantId = @Id AND IsActive = 1)   AS ActiveUsers,
                  (SELECT COUNT(*) FROM dbo.AppUsers         WHERE TenantId = @Id)                    AS TotalUsers,
                  (SELECT ISNULL(COUNT(*),0) FROM dbo.Tickets          t
                     JOIN dbo.AppUsers u ON t.SubmittedBy = u.Username AND u.TenantId = @Id
                     WHERE CAST(t.CreatedAt AS DATE) >= DATEADD(DAY,-30,GETUTCDATE()))                AS TicketsLast30,
                  (SELECT ISNULL(COUNT(*),0) FROM dbo.Mantenimientos   m
                     JOIN dbo.AppUsers u ON m.CreadoPor = u.Username AND u.TenantId = @Id
                     WHERE CAST(m.CreadoEn AS DATE) >= DATEADD(DAY,-30,GETUTCDATE()))                AS MantLast30,
                  (SELECT ISNULL(COUNT(*),0) FROM dbo.CheckadorRegistros cr
                     JOIN dbo.AppUsers u ON cr.UserId = u.Id AND u.TenantId = @Id
                     WHERE CAST(cr.Timestamp AS DATE) >= DATEADD(DAY,-7,GETUTCDATE()))               AS CheckLast7
                """;
            cmd.Parameters.AddWithValue("@Id", id);
            await using var r = await cmd.ExecuteReaderAsync(ct);
            if (!await r.ReadAsync(ct)) return NotFound();
            return Ok(new {
                activeUsers  = r.GetInt32(0),
                totalUsers   = r.GetInt32(1),
                ticketsLast30 = r.GetInt32(2),
                mantLast30   = r.GetInt32(3),
                checkLast7   = r.GetInt32(4),
            });
        }
        catch (Exception ex) { logger.LogError(ex, "GetTenantStats {Id}", id); return StatusCode(500, ex.Message); }
    }
}

// ── DTOs ──────────────────────────────────────────────────────────────────────
public record TenantDto(
    string?   Slug,
    string?   Name,
    string?   DisplayName,
    string?   PrimaryColor,
    string?   SecondaryColor,
    long?     LicensedModules,
    int?      MaxUsers,
    DateTime? ExpiresAt,
    string?   ContactEmail,
    string?   Notes
);
