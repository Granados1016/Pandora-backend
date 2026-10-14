using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using Pandora.API.Extensions;
using Pandora.Application.Features.Users;

namespace Pandora.API.Controllers;

/// <summary>
/// Consola de proveedor — gestión GLOBAL de usuarios y roles a través de todos
/// los clientes. Solo accesible para el SuperAdmin del proveedor.
/// </summary>
[ApiController]
[Route("api/superadmin")]
[Authorize(Roles = "SuperAdmin")]
public class SuperAdminController(IConfiguration config, ILogger<SuperAdminController> logger) : ControllerBase
{
    private SqlConnection Conn() => new(config.GetConnectionString("PandoraDb"));

    // Roles válidos que el SuperAdmin puede asignar.
    private static readonly string[] ValidRoles = ["SuperAdmin", "Admin", "Inventarista", "User"];

    // ── GET /api/superadmin/users ─────────────────────────────────────────────
    /// <summary>Lista TODOS los usuarios de todos los clientes, con su cliente y rol.</summary>
    [HttpGet("users")]
    public async Task<IActionResult> GetAllUsers(CancellationToken ct)
    {
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = """
                SELECT u.Id, u.Username, u.FullName, u.Email, u.Role, u.IsActive,
                       u.TenantId, ISNULL(t.DisplayName, t.Name) AS TenantName
                FROM dbo.AppUsers u
                LEFT JOIN dbo.Tenants t ON u.TenantId = t.Id
                ORDER BY TenantName, u.Username
                """;
            var list = new List<object>();
            await using var r = await cmd.ExecuteReaderAsync(ct);
            while (await r.ReadAsync(ct))
                list.Add(new
                {
                    id         = r.GetGuid(0),
                    username   = r.GetString(1),
                    fullName   = r.IsDBNull(2) ? null : r.GetString(2),
                    email      = r.IsDBNull(3) ? null : r.GetString(3),
                    role       = r.IsDBNull(4) ? "User" : r.GetString(4),
                    isActive   = !r.IsDBNull(5) && r.GetBoolean(5),
                    tenantId   = r.IsDBNull(6) ? (Guid?)null : r.GetGuid(6),
                    tenantName = r.IsDBNull(7) ? "— (sin cliente)" : r.GetString(7),
                });
            return Ok(list);
        }
        catch (Exception ex) { logger.LogError(ex, "SuperAdmin.GetAllUsers"); return StatusCode(500, ex.Message); }
    }

    // ── POST /api/superadmin/users ────────────────────────────────────────────
    /// <summary>Crea un usuario en cualquier cliente, con el rol indicado.</summary>
    [HttpPost("users")]
    public async Task<IActionResult> CreateUser([FromBody] SaCreateUserDto dto, CancellationToken ct)
    {
        if (string.IsNullOrWhiteSpace(dto.Username)) return BadRequest("Usuario requerido.");

        // Cuenta institucional/Google: sin contraseña (entra por Google) y debe
        // completar su perfil en el primer ingreso (NeedsOnboarding).
        bool googleOnly = dto.GoogleOnly;
        if (googleOnly)
        {
            if (string.IsNullOrWhiteSpace(dto.Email))
                return BadRequest("Para una cuenta institucional (Google) el correo es obligatorio.");
        }
        else if (string.IsNullOrWhiteSpace(dto.Password) || dto.Password.Length < 6)
        {
            return BadRequest("La contraseña debe tener al menos 6 caracteres.");
        }

        var role = (dto.Role ?? "User").Trim();
        var canon = ValidRoles.FirstOrDefault(v => v.Equals(role, StringComparison.OrdinalIgnoreCase)) ?? "User";

        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = """
                INSERT INTO dbo.AppUsers
                    (Id, Username, FullName, Email, PasswordHash, Role, Modules, ModulesViewOnly,
                     IsActive, Position, TenantId, NeedsOnboarding, CreatedAt)
                VALUES
                    (NEWID(), @Username, @FullName, @Email, @Hash, @Role, @Modules, 0,
                     @IsActive, @Position, @TenantId, @NeedsOnboarding, GETUTCDATE())
                """;
            var uname = dto.Username.Trim().ToLower();
            // FullName es NOT NULL en la BD → si no se captura nombre, usamos el usuario.
            var fullName = string.IsNullOrWhiteSpace(dto.FullName) ? uname : dto.FullName.Trim();
            // Cuenta Google: contraseña aleatoria no divulgada (el acceso es por Google).
            var hash = googleOnly
                ? UserService.HashPassword(Guid.NewGuid().ToString("N") + Guid.NewGuid().ToString("N"))
                : UserService.HashPassword(dto.Password!);
            cmd.Parameters.AddWithValue("@Username", uname);
            cmd.Parameters.AddWithValue("@FullName", fullName);
            // Email es NOT NULL en la BD (sin índice único) → cadena vacía si no se captura.
            cmd.Parameters.AddWithValue("@Email",    (object?)dto.Email?.Trim()    ?? string.Empty);
            cmd.Parameters.AddWithValue("@Hash",     hash);
            cmd.Parameters.AddWithValue("@Role",     canon);
            cmd.Parameters.AddWithValue("@Modules",  dto.Modules ?? 0);
            cmd.Parameters.AddWithValue("@IsActive", dto.IsActive);
            cmd.Parameters.AddWithValue("@Position", (object?)dto.Position?.Trim() ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@TenantId", dto.TenantId.HasValue ? (object)dto.TenantId.Value : DBNull.Value);
            // Solo las cuentas Google precargadas deben completar perfil al entrar.
            cmd.Parameters.AddWithValue("@NeedsOnboarding", googleOnly);
            await cmd.ExecuteNonQueryAsync(ct);
            logger.LogInformation("SuperAdmin {By} creó usuario {User} ({Role}) en tenant {Tenant}",
                User.Identity?.Name, uname, canon, dto.TenantId);
            return Ok(new { username = uname, role = canon });
        }
        catch (SqlException ex) when (ex.Number is 2627 or 2601)
        {
            return Conflict("El nombre de usuario ya existe.");
        }
        catch (Exception ex) { logger.LogError(ex, "SuperAdmin.CreateUser"); return StatusCode(500, ex.Message); }
    }

    // ── PUT /api/superadmin/users/{id}/role ───────────────────────────────────
    /// <summary>Cambia el rol de cualquier usuario (SuperAdmin / Admin / Inventarista / User).</summary>
    [HttpPut("users/{id:guid}/role")]
    public async Task<IActionResult> SetRole(Guid id, [FromBody] SetRoleDto dto, CancellationToken ct)
    {
        var role = (dto.Role ?? "").Trim();
        var canon = ValidRoles.FirstOrDefault(v => v.Equals(role, StringComparison.OrdinalIgnoreCase));
        if (canon is null)
            return BadRequest($"Rol inválido. Válidos: {string.Join(", ", ValidRoles)}.");

        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = "UPDATE dbo.AppUsers SET Role = @Role, UpdatedAt = GETUTCDATE() WHERE Id = @Id";
            cmd.Parameters.AddWithValue("@Role", canon);
            cmd.Parameters.AddWithValue("@Id",   id);
            int rows = await cmd.ExecuteNonQueryAsync(ct);
            if (rows == 0) return NotFound("Usuario no encontrado.");
            logger.LogInformation("SuperAdmin {By} cambió el rol de {Id} a {Role}", User.Identity?.Name, id, canon);
            return Ok(new { id, role = canon });
        }
        catch (Exception ex) { logger.LogError(ex, "SuperAdmin.SetRole {Id}", id); return StatusCode(500, ex.Message); }
    }

    // ── PATCH /api/superadmin/users/{id}/active ───────────────────────────────
    /// <summary>Activa o desactiva cualquier usuario.</summary>
    [HttpPatch("users/{id:guid}/active")]
    public async Task<IActionResult> ToggleActive(Guid id, CancellationToken ct)
    {
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = "UPDATE dbo.AppUsers SET IsActive = ~IsActive, UpdatedAt = GETUTCDATE() WHERE Id = @Id";
            cmd.Parameters.AddWithValue("@Id", id);
            int rows = await cmd.ExecuteNonQueryAsync(ct);
            if (rows == 0) return NotFound("Usuario no encontrado.");
            return Ok();
        }
        catch (Exception ex) { logger.LogError(ex, "SuperAdmin.ToggleActive {Id}", id); return StatusCode(500, ex.Message); }
    }
}

public record SetRoleDto(string? Role);
public record SaCreateUserDto(
    string  Username,
    string? FullName,
    string? Email,
    string? Password,
    string? Role,
    Guid?   TenantId,
    int?    Modules,
    string? Position,
    bool    IsActive = true,
    bool    GoogleOnly = false);
