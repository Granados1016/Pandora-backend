using Microsoft.AspNetCore.Http;
using System.Threading.Tasks;

namespace Pandora.API.Middleware;

/// <summary>
/// Agrega cabeceras de seguridad HTTP a todas las respuestas.
/// Mitiga: XSS, clickjacking, sniffing de MIME, info-leakage del servidor.
/// </summary>
public class SecurityHeadersMiddleware
{
    private readonly RequestDelegate _next;

    public SecurityHeadersMiddleware(RequestDelegate next)
    {
        _next = next;
    }

    public async Task InvokeAsync(HttpContext context)
    {
        var headers = context.Response.Headers;

        // ── Evita que el navegador ejecute scripts inline no autorizados ──────
        // Ajustar si se usan CDNs externos o inline scripts legítimos.
        headers["Content-Security-Policy"] =
            "default-src 'self'; " +
            "script-src 'self' 'unsafe-inline' 'unsafe-eval'; " +
            "style-src 'self' 'unsafe-inline' https://fonts.googleapis.com; " +
            "font-src 'self' https://fonts.gstatic.com data:; " +
            "img-src 'self' data: blob:; " +
            "connect-src 'self'; " +
            "frame-ancestors 'none'; " +
            "object-src 'none';";

        // ── Clickjacking — impide embeber la app en iframes ───────────────────
        headers["X-Frame-Options"] = "DENY";

        // ── Evita que el navegador adivine el tipo MIME ───────────────────────
        headers["X-Content-Type-Options"] = "nosniff";

        // ── Controla cuánta info de referrer se envía ─────────────────────────
        headers["Referrer-Policy"] = "strict-origin-when-cross-origin";

        // ── Deshabilita funciones del navegador no necesarias ─────────────────
        headers["Permissions-Policy"] =
            "geolocation=(), microphone=(), camera=(), " +
            "payment=(), usb=(), magnetometer=()";

        // ── HSTS: fuerza HTTPS por 1 año (activar solo cuando haya TLS) ───────
        // Descomentar cuando Nginx/TLS esté configurado (Sprint 3):
        // headers["Strict-Transport-Security"] = "max-age=31536000; includeSubDomains";

        // ── Oculta que el servidor es ASP.NET Core ────────────────────────────
        headers.Remove("Server");
        headers.Remove("X-Powered-By");
        headers.Remove("X-AspNet-Version");
        headers.Remove("X-AspNetMvc-Version");

        await _next(context);
    }
}

/// <summary>Extension para registrar el middleware con una línea en Program.cs</summary>
public static class SecurityHeadersExtensions
{
    public static IApplicationBuilder UseSecurityHeaders(this IApplicationBuilder app)
        => app.UseMiddleware<SecurityHeadersMiddleware>();
}
