using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Caching.Memory;
using System.Security.Claims;

namespace Pandora.API.Controllers;

/// <summary>
/// SSO entre Pandora y Azul AI.
/// Flujo: Pandora genera token → Azul lo valida → usuario queda logueado en Azul.
/// </summary>
[ApiController]
[Route("api/azul-sso")]
public class AzulSsoController(
    IMemoryCache cache,
    IConfiguration config) : ControllerBase
{
    private string SsoSecret => config["Azul:SsoSecret"] ?? "";

    // ── POST /api/azul-sso/token ──────────────────────────────────────────────
    /// <summary>
    /// Genera un token SSO de un solo uso (válido 2 min).
    /// Requiere que el usuario esté autenticado en Pandora.
    /// </summary>
    [HttpPost("token")]
    [Authorize]
    public IActionResult GenerateToken()
    {
        var userId   = User.FindFirstValue(ClaimTypes.NameIdentifier)
                    ?? User.FindFirstValue(System.IdentityModel.Tokens.Jwt.JwtRegisteredClaimNames.Sub)
                    ?? "";
        var username = User.FindFirstValue(ClaimTypes.Name)     ?? "";
        var fullName = User.FindFirstValue("fullName")          ?? username;
        var role     = User.FindFirstValue(ClaimTypes.Role)     ?? "User";

        // Construir email a partir del username (Pandora usa username, Azul usa email)
        // Si el username tiene formato email lo usamos; si no, generamos uno del dominio
        var email = username.Contains('@')
            ? username
            : $"{username}@pandora.app";

        var token = Guid.NewGuid().ToString("N");
        var entry = new SsoTokenEntry(userId, email, fullName, role);

        // TTL: 2 minutos — tiempo suficiente para abrir la pestaña de Azul
        cache.Set($"azul_sso_{token}", entry, TimeSpan.FromMinutes(2));

        var azulUrl = config["Azul:FrontendUrl"] ?? "http://localhost:5174";
        return Ok(new
        {
            token,
            url = $"{azulUrl}/login?pandora_token={token}",
        });
    }

    // ── POST /api/azul-sso/validate/{token} ──────────────────────────────────
    /// <summary>
    /// Valida un token SSO y lo elimina (uso único).
    /// Solo puede ser llamado por el backend de Azul con el header correcto.
    /// </summary>
    [HttpPost("validate/{token}")]
    [AllowAnonymous]
    public IActionResult ValidateToken(
        string token,
        [FromHeader(Name = "X-Sso-Secret")] string? secret)
    {
        if (string.IsNullOrWhiteSpace(secret) || secret != SsoSecret)
            return Unauthorized(new { error = "Secreto SSO inválido." });

        var key = $"azul_sso_{token}";
        if (!cache.TryGetValue<SsoTokenEntry>(key, out var entry) || entry is null)
            return NotFound(new { error = "Token SSO inválido o expirado." });

        cache.Remove(key); // one-time use: eliminar tras validación exitosa
        return Ok(entry);
    }

    private record SsoTokenEntry(string UserId, string Email, string FullName, string Role);
}
