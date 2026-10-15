using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using Pandora.Application.Features.Users;
using System.Security.Claims;
using System.Text.RegularExpressions;

namespace Pandora.API.Controllers;

/// <summary>
/// Solicitudes de acceso con cuenta institucional (onboarding Google).
/// Una persona con correo del dominio permitido intenta entrar con Google, no
/// tiene cuenta y se genera una solicitud PENDIENTE. El administrador la ve aquí
/// y la aprueba (eligiendo el rol) o la rechaza. Al aprobar se crea su usuario y
/// en su siguiente inicio de sesión con Google ya entra al sistema.
/// </summary>
[ApiController]
[Route("api/access-requests")]
[Authorize(Roles = "Admin,SuperAdmin")]
public partial class AccessRequestsController(IConfiguration config, ILogger<AccessRequestsController> logger) : ControllerBase
{
    private SqlConnection Conn() => new(config.GetConnectionString("PandoraDb"));

    private bool IsSuperAdmin => User.IsInRole("SuperAdmin");
    private string? CallerName =>
        User.FindFirstValue(ClaimTypes.Name) ?? User.FindFirstValue("name");
    private Guid? CallerTenantId =>
        Guid.TryParse(User.FindFirstValue("tenantId"), out var t) ? t : null;

    private static readonly string[] ValidRoles = ["SuperAdmin", "Admin", "Inventarista", "User"];

    [GeneratedRegex("[^a-z0-9._-]")]
    private static partial Regex UsernameSanitizer();

    // ── GET /api/access-requests?status=pending ──────────────────────────────
    /// <summary>Lista las solicitudes (por defecto las pendientes).</summary>
    [HttpGet]
    public async Task<IActionResult> List([FromQuery] string status = "pending", CancellationToken ct = default)
    {
        var st = (status ?? "pending").Trim().ToLowerInvariant();
        if (st != "pending" && st != "approved" && st != "rejected" && st != "all")
            st = "pending";
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await using var cmd = conn.CreateCommand();
            // SuperAdmin ve todas; un Admin de cliente ve las de su tenant o sin asignar.
            var scope = IsSuperAdmin
                ? ""
                : " AND (ar.TenantId = @Tenant OR ar.TenantId IS NULL) ";
            cmd.CommandText = $"""
                SELECT ar.Id, ar.Email, ar.FullName, ar.Picture, ar.Position, ar.TenantId,
                       ar.Status, ar.CreatedAt, ar.DecidedAt, ar.DecidedBy,
                       ISNULL(t.DisplayName, t.Name) AS TenantName
                FROM dbo.AccessRequests ar
                LEFT JOIN dbo.Tenants t ON ar.TenantId = t.Id
                WHERE ({(st == "all" ? "1=1" : "ar.Status = @Status")}) {scope}
                ORDER BY ar.CreatedAt DESC
                """;
            if (st != "all") cmd.Parameters.AddWithValue("@Status", st);
            if (!IsSuperAdmin) cmd.Parameters.AddWithValue("@Tenant", (object?)CallerTenantId ?? DBNull.Value);

            var list = new List<object>();
            await using var r = await cmd.ExecuteReaderAsync(ct);
            while (await r.ReadAsync(ct))
                list.Add(new
                {
                    id         = r.GetGuid(0),
                    email      = r.GetString(1),
                    fullName   = r.IsDBNull(2) ? null : r.GetString(2),
                    picture    = r.IsDBNull(3) ? null : r.GetString(3),
                    position   = r.IsDBNull(4) ? null : r.GetString(4),
                    tenantId   = r.IsDBNull(5) ? (Guid?)null : r.GetGuid(5),
                    status     = r.GetString(6),
                    createdAt  = r.GetDateTime(7),
                    decidedAt  = r.IsDBNull(8) ? (DateTime?)null : r.GetDateTime(8),
                    decidedBy  = r.IsDBNull(9) ? null : r.GetString(9),
                    tenantName = r.IsDBNull(10) ? null : r.GetString(10),
                });
            return Ok(list);
        }
        catch (Exception ex) { logger.LogError(ex, "AccessRequests.List"); return StatusCode(500, ex.Message); }
    }

    // ── GET /api/access-requests/count ───────────────────────────────────────
    /// <summary>Número de solicitudes pendientes (para la insignia del menú).</summary>
    [HttpGet("count")]
    public async Task<IActionResult> Count(CancellationToken ct)
    {
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await using var cmd = conn.CreateCommand();
            var scope = IsSuperAdmin ? "" : " AND (TenantId = @Tenant OR TenantId IS NULL) ";
            cmd.CommandText = $"SELECT COUNT(*) FROM dbo.AccessRequests WHERE Status = 'pending' {scope}";
            if (!IsSuperAdmin) cmd.Parameters.AddWithValue("@Tenant", (object?)CallerTenantId ?? DBNull.Value);
            var n = (int)(await cmd.ExecuteScalarAsync(ct))!;
            return Ok(new { count = n });
        }
        catch (Exception ex) { logger.LogError(ex, "AccessRequests.Count"); return StatusCode(500, ex.Message); }
    }

    // ── POST /api/access-requests/{id}/approve ───────────────────────────────
    /// <summary>Aprueba una solicitud: crea el usuario con el rol elegido.</summary>
    [HttpPost("{id:guid}/approve")]
    public async Task<IActionResult> Approve(Guid id, [FromBody] ApproveDto dto, CancellationToken ct)
    {
        // Rol: por defecto Usuario. Solo el SuperAdmin puede otorgar SuperAdmin.
        var role = (dto?.Role ?? "User").Trim();
        var canon = ValidRoles.FirstOrDefault(v => v.Equals(role, StringComparison.OrdinalIgnoreCase)) ?? "User";
        if (canon == "SuperAdmin" && !IsSuperAdmin)
            return Forbid();

        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);

            // 1) Leer la solicitud
            string email, reqName; string? reqPos; Guid? reqTenant; string status;
            await using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = "SELECT Email, FullName, Position, TenantId, Status FROM dbo.AccessRequests WHERE Id = @Id";
                cmd.Parameters.AddWithValue("@Id", id);
                await using var r = await cmd.ExecuteReaderAsync(ct);
                if (!await r.ReadAsync(ct)) return NotFound("Solicitud no encontrada.");
                email     = r.GetString(0);
                reqName   = r.IsDBNull(1) ? email : r.GetString(1);
                reqPos    = r.IsDBNull(2) ? null : r.GetString(2);
                reqTenant = r.IsDBNull(3) ? null : r.GetGuid(3);
                status    = r.GetString(4);
            }
            if (status == "approved") return Conflict("La solicitud ya fue aprobada.");

            // 2) ¿Ya existe un usuario con ese correo? (evita duplicados)
            await using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = "SELECT COUNT(*) FROM dbo.AppUsers WHERE LOWER(Email) = LOWER(@E)";
                cmd.Parameters.AddWithValue("@E", email);
                if ((int)(await cmd.ExecuteScalarAsync(ct))! > 0)
                {
                    await MarkDecided(conn, id, "approved", ct);
                    return Ok(new { alreadyExists = true, email });
                }
            }

            // 3) Resolver tenant: Admin usa el suyo; SuperAdmin puede elegir.
            Guid? tenant = IsSuperAdmin
                ? (dto?.TenantId ?? reqTenant)
                : (CallerTenantId ?? reqTenant);
            if (tenant is null)
            {
                await using var cmdT = conn.CreateCommand();
                cmdT.CommandText = "SELECT TOP 1 Id FROM dbo.Tenants WHERE Slug = 'default' ORDER BY CreadoEn";
                var o = await cmdT.ExecuteScalarAsync(ct);
                if (o is Guid g) tenant = g;
            }

            // 4) Username único a partir del correo
            var baseName = UsernameSanitizer().Replace(email.Split('@')[0].ToLowerInvariant(), "");
            if (string.IsNullOrWhiteSpace(baseName)) baseName = "usuario";
            var username = baseName;
            for (int i = 1; i <= 50; i++)
            {
                await using var cmdU = conn.CreateCommand();
                cmdU.CommandText = "SELECT COUNT(*) FROM dbo.AppUsers WHERE LOWER(Username) = LOWER(@U)";
                cmdU.Parameters.AddWithValue("@U", username);
                if ((int)(await cmdU.ExecuteScalarAsync(ct))! == 0) break;
                username = $"{baseName}{i}";
            }

            // 5) Crear el usuario. Sin contraseña usable: entra por Google.
            var fullName = !string.IsNullOrWhiteSpace(dto?.FullName) ? dto!.FullName!.Trim() : reqName;
            var position = !string.IsNullOrWhiteSpace(dto?.Position) ? dto!.Position!.Trim() : reqPos;
            await using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = """
                    INSERT INTO dbo.AppUsers
                        (Id, Username, FullName, Email, PasswordHash, Role, Modules, ModulesViewOnly,
                         IsActive, Position, TenantId, CreatedAt)
                    VALUES
                        (NEWID(), @Username, @FullName, @Email, @Hash, @Role, @Modules, 0,
                         1, @Position, @TenantId, GETUTCDATE())
                    """;
                cmd.Parameters.AddWithValue("@Username", username);
                cmd.Parameters.AddWithValue("@FullName", fullName);
                cmd.Parameters.AddWithValue("@Email",    email);
                // Contraseña aleatoria no divulgada → el acceso es por Google (o "olvidé contraseña").
                cmd.Parameters.AddWithValue("@Hash",     UserService.HashPassword(Guid.NewGuid().ToString("N") + Guid.NewGuid().ToString("N")));
                cmd.Parameters.AddWithValue("@Role",     canon);
                cmd.Parameters.AddWithValue("@Modules",  dto?.Modules ?? 0);
                cmd.Parameters.AddWithValue("@Position", (object?)position ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@TenantId", (object?)tenant ?? DBNull.Value);
                await cmd.ExecuteNonQueryAsync(ct);
            }

            await MarkDecided(conn, id, "approved", ct);
            logger.LogInformation("AccessRequest {Id} APROBADA por {By} → usuario {User} ({Role})", id, CallerName, username, canon);
            return Ok(new { username, email, role = canon });
        }
        catch (SqlException ex) when (ex.Number is 2627 or 2601)
        {
            return Conflict("Ya existe un usuario con ese correo o nombre de usuario.");
        }
        catch (Exception ex) { logger.LogError(ex, "AccessRequests.Approve {Id}", id); return StatusCode(500, ex.Message); }
    }

    // ── POST /api/access-requests/{id}/reject ────────────────────────────────
    /// <summary>Rechaza una solicitud.</summary>
    [HttpPost("{id:guid}/reject")]
    public async Task<IActionResult> Reject(Guid id, CancellationToken ct)
    {
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            var rows = await MarkDecided(conn, id, "rejected", ct);
            if (rows == 0) return NotFound("Solicitud no encontrada.");
            logger.LogInformation("AccessRequest {Id} RECHAZADA por {By}", id, CallerName);
            return Ok(new { id, status = "rejected" });
        }
        catch (Exception ex) { logger.LogError(ex, "AccessRequests.Reject {Id}", id); return StatusCode(500, ex.Message); }
    }

    private async Task<int> MarkDecided(SqlConnection conn, Guid id, string status, CancellationToken ct)
    {
        await using var cmd = conn.CreateCommand();
        cmd.CommandText = "UPDATE dbo.AccessRequests SET Status = @S, DecidedAt = GETUTCDATE(), DecidedBy = @By WHERE Id = @Id";
        cmd.Parameters.AddWithValue("@S", status);
        cmd.Parameters.AddWithValue("@By", (object?)CallerName ?? DBNull.Value);
        cmd.Parameters.AddWithValue("@Id", id);
        return await cmd.ExecuteNonQueryAsync(ct);
    }
}

public record ApproveDto(string? Role, Guid? TenantId, int? Modules, string? FullName, string? Position);
