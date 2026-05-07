using System.IdentityModel.Tokens.Jwt;
using System.Security.Claims;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.RateLimiting;
using Microsoft.Data.SqlClient;
using Microsoft.IdentityModel.Tokens;
using Pandora.Application.DTOs;
using Pandora.Application.Interfaces;

namespace Pandora.API.Controllers;

[ApiController]
[Route("api/[controller]")]
public class AuthController(
    IJwtService jwtService,
    IConfiguration config,
    ILogger<AuthController> logger) : ControllerBase
{
    // ── Helpers ───────────────────────────────────────────────────────────────
    private SqlConnection Conn() => new(config.GetConnectionString("PandoraDb"));

    private static string NewRefreshToken() =>
        Convert.ToBase64String(RandomNumberGenerator.GetBytes(64));

    // ── POST /api/auth/login ──────────────────────────────────────────────────
    /// <summary>
    /// Login — devuelve JWT + refresh token de 7 días.
    /// Limitado a 5 intentos por IP por minuto (rate limiting).
    /// </summary>
    [EnableRateLimiting("login-policy")]
    [HttpPost("login")]
    [AllowAnonymous]
    public async Task<IActionResult> Login([FromBody] LoginRequest req, CancellationToken ct)
    {
        var response = await jwtService.LoginAsync(req, ct);
        if (response is null)
            return Unauthorized("Credenciales incorrectas.");

        // Generar refresh token y persistirlo
        string refreshToken = NewRefreshToken();
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await using var cmd = conn.CreateCommand();
            // Un refresh token activo por usuario (revoca el anterior)
            cmd.CommandText = """
                DELETE FROM dbo.RefreshTokens WHERE Username = @Username;
                INSERT INTO dbo.RefreshTokens (Token, Username, ExpiresAt, CreatedAt)
                VALUES (@Token, @Username, @ExpiresAt, GETUTCDATE());
                """;
            cmd.Parameters.AddWithValue("@Token",     refreshToken);
            cmd.Parameters.AddWithValue("@Username",  req.Username.ToLower().Trim());
            cmd.Parameters.AddWithValue("@ExpiresAt", DateTime.UtcNow.AddDays(7));
            await cmd.ExecuteNonQueryAsync(ct);
        }
        catch (Exception ex)
        {
            // No-fatal: el login funciona aunque no se pueda persistir el refresh token
            logger.LogWarning(ex, "No se pudo guardar refresh token para {User}", req.Username);
        }

        // Añadir refreshToken a la respuesta del servicio
        var json = JsonSerializer.Serialize(response);
        var dict = JsonSerializer.Deserialize<Dictionary<string, JsonElement>>(json)!;
        dict["refreshToken"] = JsonSerializer.SerializeToElement(refreshToken);
        return Ok(dict);
    }

    // ── POST /api/auth/refresh ────────────────────────────────────────────────
    /// <summary>
    /// Renueva el JWT usando un refresh token válido.
    /// Implementa rotación: el token antiguo se invalida y se emite uno nuevo.
    /// </summary>
    [HttpPost("refresh")]
    [AllowAnonymous]
    public async Task<IActionResult> Refresh([FromBody] RefreshRequest req, CancellationToken ct)
    {
        if (string.IsNullOrWhiteSpace(req.RefreshToken))
            return BadRequest("Refresh token requerido.");

        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);

            // ── Validar refresh token ─────────────────────────────────────────
            string? username = null;
            await using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = """
                    SELECT Username FROM dbo.RefreshTokens
                    WHERE Token = @Token AND ExpiresAt > GETUTCDATE()
                    """;
                cmd.Parameters.AddWithValue("@Token", req.RefreshToken);
                var result = await cmd.ExecuteScalarAsync(ct);
                username = result as string;
            }

            if (username is null)
                return Unauthorized("Refresh token inválido o expirado.");

            // ── Cargar usuario ────────────────────────────────────────────────
            Guid   userId   = Guid.Empty;
            string fullName = username;
            string role     = "User";
            int    modules  = 0;

            await using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = """
                    SELECT Id, FullName, Role, Modules
                    FROM dbo.AppUsers
                    WHERE LOWER(Username) = LOWER(@Username) AND IsActive = 1
                    """;
                cmd.Parameters.AddWithValue("@Username", username);
                await using var r = await cmd.ExecuteReaderAsync(ct);
                if (!await r.ReadAsync(ct))
                    return Unauthorized("Usuario desactivado o no encontrado.");

                userId   = r.GetGuid(0);
                fullName = r.IsDBNull(1) ? username : r.GetString(1);
                role     = r.IsDBNull(2) ? "User"   : r.GetString(2);
                modules  = r.GetInt32(3);
            }

            // ── Generar nuevo JWT ─────────────────────────────────────────────
            string jwtKey   = config["JwtSettings:Key"]!;
            string issuer   = config["JwtSettings:Issuer"]!;
            string audience = config["JwtSettings:Audience"]!;
            int    expMin   = int.TryParse(config["JwtSettings:ExpiresInMinutes"], out var m) ? m : 30;

            var claims = new[]
            {
                new Claim(ClaimTypes.Name,                   username),
                new Claim(ClaimTypes.Role,                   role),
                new Claim("fullName",                        fullName),
                new Claim("modules",                         modules.ToString()),
                new Claim(JwtRegisteredClaimNames.Sub,       userId.ToString()),
                new Claim(JwtRegisteredClaimNames.Jti,       Guid.NewGuid().ToString()),
            };

            var key    = new SymmetricSecurityKey(Encoding.UTF8.GetBytes(jwtKey));
            var creds  = new SigningCredentials(key, SecurityAlgorithms.HmacSha256);
            var jwtObj = new JwtSecurityToken(
                issuer:             issuer,
                audience:           audience,
                claims:             claims,
                expires:            DateTime.UtcNow.AddMinutes(expMin),
                signingCredentials: creds);

            string newJwt = new JwtSecurityTokenHandler().WriteToken(jwtObj);

            // ── Rotar refresh token (invalida el viejo, emite uno nuevo) ──────
            string newRefresh = NewRefreshToken();
            await using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = """
                    DELETE FROM dbo.RefreshTokens WHERE Token = @Old;
                    INSERT INTO dbo.RefreshTokens (Token, Username, ExpiresAt, CreatedAt)
                    VALUES (@New, @Username, @ExpiresAt, GETUTCDATE());
                    """;
                cmd.Parameters.AddWithValue("@Old",      req.RefreshToken);
                cmd.Parameters.AddWithValue("@New",      newRefresh);
                cmd.Parameters.AddWithValue("@Username", username);
                cmd.Parameters.AddWithValue("@ExpiresAt", DateTime.UtcNow.AddDays(7));
                await cmd.ExecuteNonQueryAsync(ct);
            }

            return Ok(new { token = newJwt, refreshToken = newRefresh });
        }
        catch (Exception ex)
        {
            logger.LogError(ex, "Error al refrescar token");
            return StatusCode(500, "Error interno al renovar la sesión.");
        }
    }

    // ── POST /api/auth/logout | /api/auth/revoke ──────────────────────────────
    /// <summary>Revoca el refresh token del usuario (logout limpio).</summary>
    [HttpPost("logout")]
    [HttpPost("revoke")]
    [AllowAnonymous]
    public async Task<IActionResult> Logout([FromBody] RefreshRequest req, CancellationToken ct)
    {
        if (string.IsNullOrWhiteSpace(req.RefreshToken))
            return NoContent();
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = "DELETE FROM dbo.RefreshTokens WHERE Token = @Token";
            cmd.Parameters.AddWithValue("@Token", req.RefreshToken);
            await cmd.ExecuteNonQueryAsync(ct);
        }
        catch (Exception ex)
        {
            logger.LogWarning(ex, "Error al revocar refresh token");
        }
        return NoContent();
    }
}

// ── DTOs locales ──────────────────────────────────────────────────────────────
public record RefreshRequest(string RefreshToken);
