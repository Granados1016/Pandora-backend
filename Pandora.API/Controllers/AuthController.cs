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
using Pandora.API.Services;
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

    // ── POST /api/auth/forgot-password (#11) ─────────────────────────────────
    /// <summary>Genera token de recuperación y envía correo al usuario.</summary>
    [HttpPost("forgot-password")]
    [AllowAnonymous]
    public async Task<IActionResult> ForgotPassword([FromBody] ForgotPasswordRequest req, CancellationToken ct)
    {
        if (string.IsNullOrWhiteSpace(req.Email))
            return BadRequest("Email requerido.");

        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);

            // Buscar usuario por email
            string? username = null, fullName = null;
            await using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = """
                    SELECT Username, FullName FROM dbo.AppUsers
                    WHERE LOWER(Email) = LOWER(@Email) AND IsActive = 1
                    """;
                cmd.Parameters.AddWithValue("@Email", req.Email.Trim());
                await using var r = await cmd.ExecuteReaderAsync(ct);
                if (await r.ReadAsync(ct))
                {
                    username = r.GetString(0);
                    fullName = r.IsDBNull(1) ? r.GetString(0) : r.GetString(1);
                }
            }

            // Siempre retornar 204 para no revelar si el email existe
            if (username is null)
            {
                logger.LogDebug("ForgotPassword: email no encontrado {Email}", req.Email);
                return NoContent();
            }

            // Generar token de reset (expira en 30 minutos)
            string token = Convert.ToBase64String(RandomNumberGenerator.GetBytes(32))
                .Replace("+", "-").Replace("/", "_").TrimEnd('=');

            await using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = """
                    DELETE FROM dbo.PasswordResetTokens WHERE Username = @Username;
                    INSERT INTO dbo.PasswordResetTokens (Token, Username, ExpiresAt, CreatedAt)
                    VALUES (@Token, @Username, @ExpiresAt, GETUTCDATE());
                    """;
                cmd.Parameters.AddWithValue("@Token",     token);
                cmd.Parameters.AddWithValue("@Username",  username);
                cmd.Parameters.AddWithValue("@ExpiresAt", DateTime.UtcNow.AddMinutes(30));
                await cmd.ExecuteNonQueryAsync(ct);
            }

            // Enviar correo
            await SendResetEmailAsync(req.Email, fullName ?? username, token);

            return NoContent();
        }
        catch (Exception ex)
        {
            logger.LogError(ex, "ForgotPassword error for {Email}", req.Email);
            return StatusCode(500, "Error al procesar la solicitud.");
        }
    }

    // ── POST /api/auth/reset-password (#11) ──────────────────────────────────
    /// <summary>Valida token y actualiza contraseña.</summary>
    [HttpPost("reset-password")]
    [AllowAnonymous]
    public async Task<IActionResult> ResetPassword([FromBody] ResetPasswordRequest req, CancellationToken ct)
    {
        if (string.IsNullOrWhiteSpace(req.Token) || string.IsNullOrWhiteSpace(req.NewPassword))
            return BadRequest("Token y nueva contraseña requeridos.");

        if (req.NewPassword.Length < 8)
            return BadRequest("La contraseña debe tener al menos 8 caracteres.");

        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);

            // Validar token
            string? username = null;
            await using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = """
                    SELECT Username FROM dbo.PasswordResetTokens
                    WHERE Token = @Token AND ExpiresAt > GETUTCDATE() AND UsedAt IS NULL
                    """;
                cmd.Parameters.AddWithValue("@Token", req.Token);
                var result = await cmd.ExecuteScalarAsync(ct);
                username = result as string;
            }

            if (username is null)
                return BadRequest("Token inválido o expirado.");

            // Actualizar contraseña
            string hash = BCrypt.Net.BCrypt.HashPassword(req.NewPassword);
            await using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = """
                    UPDATE dbo.AppUsers
                    SET PasswordHash = @Hash, UpdatedAt = GETUTCDATE()
                    WHERE LOWER(Username) = LOWER(@Username);
                    UPDATE dbo.PasswordResetTokens
                    SET UsedAt = GETUTCDATE()
                    WHERE Token = @Token;
                    """;
                cmd.Parameters.AddWithValue("@Hash",     hash);
                cmd.Parameters.AddWithValue("@Username", username);
                cmd.Parameters.AddWithValue("@Token",    req.Token);
                await cmd.ExecuteNonQueryAsync(ct);
            }

            // Invalidar sesiones activas
            await using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = "DELETE FROM dbo.RefreshTokens WHERE Username = @Username";
                cmd.Parameters.AddWithValue("@Username", username);
                await cmd.ExecuteNonQueryAsync(ct);
            }

            logger.LogInformation("Contraseña restablecida para {Username}", username);
            return NoContent();
        }
        catch (Exception ex)
        {
            logger.LogError(ex, "ResetPassword error");
            return StatusCode(500, "Error al restablecer la contraseña.");
        }
    }

    // ── Helpers: envío de correo ──────────────────────────────────────────────
    private async Task SendResetEmailAsync(string toEmail, string toName, string token)
    {
        try
        {
            var connStr     = config.GetConnectionString("PandoraDb")!;
            var smtpCfg     = await SmtpHelper.LoadAsync(connStr, config);
            var frontendUrl = config["FrontendUrl"] ?? "http://localhost:5173";
            var resetUrl    = $"{frontendUrl}/reset-password?token={Uri.EscapeDataString(token)}";

            var body = $"""
                <!DOCTYPE html><html lang="es"><head><meta charset="UTF-8"></head>
                <body style="font-family:Arial,sans-serif;background:#f5f5f5;margin:0;padding:20px">
                  <div style="max-width:560px;margin:0 auto;background:white;border-radius:8px;overflow:hidden;box-shadow:0 2px 8px rgba(0,0,0,.1)">
                    <div style="background:#1a237e;padding:24px;text-align:center">
                      <h1 style="color:white;margin:0;font-size:22px">PANDORA</h1>
                      <p style="color:rgba(255,255,255,.7);margin:4px 0 0;font-size:13px">Sistema de Gestión — iMET</p>
                    </div>
                    <div style="padding:28px">
                      <h2 style="color:#1a237e;font-size:18px;margin:0 0 16px">🔑 Recuperación de contraseña</h2>
                      <p style="color:#333;margin:0 0 12px">Hola <strong>{toName}</strong>,</p>
                      <p style="color:#555;margin:0 0 20px">Recibimos una solicitud para restablecer la contraseña de tu cuenta. Haz clic en el botón siguiente:</p>
                      <div style="text-align:center;margin:24px 0">
                        <a href="{resetUrl}"
                           style="background:#1a237e;color:white;padding:14px 28px;border-radius:6px;text-decoration:none;font-weight:bold;font-size:15px;display:inline-block">
                          Restablecer contraseña
                        </a>
                      </div>
                      <p style="color:#888;font-size:12px">Este enlace expirará en <strong>30 minutos</strong>. Si no solicitaste este cambio, ignora este correo.</p>
                    </div>
                    <div style="background:#f9f9f9;border-top:1px solid #eee;padding:14px 28px;text-align:center">
                      <p style="color:#aaa;font-size:11px;margin:0">Pandora — Coordinación de TI | iMET</p>
                    </div>
                  </div>
                </body></html>
                """;

            var err = await SmtpHelper.SendAsync(smtpCfg, toEmail, toName, "🔑 Recupera tu contraseña en Pandora", body);
            if (err != null)
                logger.LogWarning("No se pudo enviar correo de reset: {Error}", err);
            else
                logger.LogInformation("Correo de recuperación enviado a {Email}", toEmail);
        }
        catch (Exception ex)
        {
            logger.LogWarning(ex, "No se pudo enviar correo de recuperación a {Email}", toEmail);
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
public record ForgotPasswordRequest(string Email);
public record ResetPasswordRequest(string Token, string NewPassword);
