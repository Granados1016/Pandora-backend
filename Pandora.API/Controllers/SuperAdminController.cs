using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using Pandora.API.Extensions;

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
