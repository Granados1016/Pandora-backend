using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.RateLimiting;
using Pandora.Application.DTOs;
using Pandora.Application.Interfaces;

namespace Pandora.API.Controllers;

[ApiController]
[Route("api/[controller]")]
public class AuthController(IJwtService jwtService) : ControllerBase
{
    /// <summary>
    /// Login — devuelve JWT.
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

        return Ok(response);
    }
}
