# Sprint 2 — Rate Limiting + Security Headers
## Archivos a modificar en Visual Studio / VS Code

---

## 1. `Middleware/SecurityHeadersMiddleware.cs`
**YA CREADO** — no hace falta tocarlo.

---

## 2. `Pandora.API/Program.cs`

### 2a. Agrega los usings al inicio del archivo (si no existen):
```csharp
using Microsoft.AspNetCore.RateLimiting;
using System.Threading.RateLimiting;
using Pandora.API.Middleware;
```

### 2b. ANTES de `var app = builder.Build();`, agrega el Rate Limiter:
```csharp
// ── Rate Limiting — protege /auth/login contra fuerza bruta ──────────────
builder.Services.AddRateLimiter(options =>
{
    // Ventana fija: máx 5 intentos por IP por minuto
    options.AddFixedWindowLimiter("login-policy", config =>
    {
        config.PermitLimit         = 5;
        config.Window              = TimeSpan.FromMinutes(1);
        config.QueueProcessingOrder = QueueProcessingOrder.OldestFirst;
        config.QueueLimit          = 0;
    });

    // Respuesta cuando se supera el límite
    options.RejectionStatusCode = StatusCodes.Status429TooManyRequests;

    // Callback para logging (opcional)
    options.OnRejected = async (context, token) =>
    {
        context.HttpContext.Response.ContentType = "application/json";
        await context.HttpContext.Response.WriteAsync(
            "{\"error\":\"Demasiados intentos. Espera 1 minuto e intenta de nuevo.\"}",
            token);
    };
});
```

### 2c. DESPUÉS de `var app = builder.Build();`, agrega los middlewares:
```csharp
// ── Security Headers ──────────────────────────────────────────────────────
app.UseSecurityHeaders();

// ── Rate Limiter ──────────────────────────────────────────────────────────
app.UseRateLimiter();
```

> ⚠️  `UseSecurityHeaders()` y `UseRateLimiter()` deben ir ANTES de `app.UseAuthentication()`.

---

## 3. `Controllers/AuthController.cs`

### 3a. Agrega el using al inicio:
```csharp
using Microsoft.AspNetCore.RateLimiting;
```

### 3b. Aplica la política al endpoint de login:
Busca el método `Login` (o `[HttpPost("login")]`) y agrega el atributo encima:
```csharp
[EnableRateLimiting("login-policy")]
[HttpPost("login")]
[AllowAnonymous]
public async Task<IActionResult> Login([FromBody] LoginRequest request)
{
    // ... código existente sin cambios ...
}
```

---

## 4. Verificación rápida tras compilar y reiniciar

```bash
# Debe responder 200 con token:
curl -X POST http://localhost:5000/api/auth/login \
  -H "Content-Type: application/json" \
  -d '{"username":"admin","password":"PandoraAdmin2024!"}'

# Después de 5 intentos en < 1 min debe responder 429:
for i in {1..6}; do
  curl -s -o /dev/null -w "%{http_code}\n" \
    -X POST http://localhost:5000/api/auth/login \
    -H "Content-Type: application/json" \
    -d '{"username":"x","password":"x"}'
done

# Debe mostrar: 401 401 401 401 401 429

# Verificar headers de seguridad:
curl -I http://localhost:5000/api/auth/login
# Debe mostrar: X-Frame-Options, X-Content-Type-Options, Content-Security-Policy, etc.
```

---

## Notas
- El rate limiting es **por IP**. En Docker, todas las peticiones llegan desde la red interna.
  Si eso causa problemas, usa `context.HttpContext.Connection.RemoteIpAddress` como key.
- HSTS está comentado intencionalmente — se activa en Sprint 3 con Nginx + TLS.
- CSP usa `unsafe-inline` porque React/MUI inyectan estilos inline.
  En Sprint 3 se puede refinar con nonces.
