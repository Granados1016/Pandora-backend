using System.Text.Json;
using Microsoft.AspNetCore.Http;

namespace Pandora.API.Middleware;

/// <summary>
/// Captura excepciones no manejadas y devuelve 500 con JSON.
/// </summary>
public class ErrorHandlingMiddleware(RequestDelegate next, ILogger<ErrorHandlingMiddleware> logger)
{
    public async Task InvokeAsync(HttpContext context)
    {
        try
        {
            await next(context);
        }
        catch (Exception ex)
        {
            logger.LogError(ex, "Unhandled exception on {Method} {Path}",
                context.Request.Method, context.Request.Path);

            // Solo escribir si la respuesta no ha comenzado todavía;
            // si ya empezó no podemos cambiar los headers/status.
            if (!context.Response.HasStarted)
            {
                context.Response.StatusCode  = 500;
                context.Response.ContentType = "application/json";
                // Exponer mensaje y tipo de excepción para facilitar diagnóstico.
                // TODO: en producción estable, reemplazar por mensaje genérico.
                await context.Response.WriteAsync(JsonSerializer.Serialize(new
                {
                    error  = ex.Message,
                    type   = ex.GetType().Name,
                    detail = ex.InnerException?.Message,
                }));
            }
        }
    }
}
