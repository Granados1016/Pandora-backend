using System.IdentityModel.Tokens.Jwt;
using System.Security.Claims;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using Google.Apis.Auth;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.RateLimiting;
using Microsoft.Data.SqlClient;
using Microsoft.IdentityModel.Tokens;
using Pandora.API.Services;
using Pandora.Application.DTOs;
using Pandora.Application.Features.Users;
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

    /// <summary>
    /// Construye un JWT firmado con la clave y configuración del proyecto.
    /// Centraliza la lógica compartida entre /refresh y /verify-otp.
    /// </summary>
    private string BuildJwt(string username, string role, string fullName, int modules, Guid userId,
                            Guid? tenantId = null, string? impersonatedBy = null, bool needsOnboarding = false)
    {
        string jwtKey   = config["JwtSettings:Key"]!;
        string issuer   = config["JwtSettings:Issuer"]!;
        string audience = config["JwtSettings:Audience"]!;
        int    expMin   = int.TryParse(config["JwtSettings:ExpiresInMinutes"], out var m) ? m : 30;

        var claims = new List<Claim>
        {
            new(ClaimTypes.Name,             username),
            new(ClaimTypes.Role,             role),
            new("fullName",                  fullName),
            new("modules",                   modules.ToString()),
            new(JwtRegisteredClaimNames.Sub, userId.ToString()),
            new(JwtRegisteredClaimNames.Jti, Guid.NewGuid().ToString()),
        };

        if (tenantId.HasValue)
            claims.Add(new Claim("tenantId", tenantId.Value.ToString()));

        // Marca de impersonación: el SuperAdmin está actuando como Admin de este tenant.
        if (!string.IsNullOrWhiteSpace(impersonatedBy))
            claims.Add(new Claim("impersonatedBy", impersonatedBy));

        // Perfil pendiente: cuenta precargada que aún debe completar nombre/puesto.
        if (needsOnboarding)
            claims.Add(new Claim("needsOnboarding", "1"));

        var key    = new SymmetricSecurityKey(Encoding.UTF8.GetBytes(jwtKey));
        var creds  = new SigningCredentials(key, SecurityAlgorithms.HmacSha256);
        var jwtObj = new JwtSecurityToken(issuer, audience, claims,
            expires:            DateTime.UtcNow.AddMinutes(expMin),
            signingCredentials: creds);
        return new JwtSecurityTokenHandler().WriteToken(jwtObj);
    }

    /// <summary>
    /// Lee TenantId del usuario desde la BD para incluirlo en el JWT.
    /// </summary>
    private async Task<Guid?> GetUserTenantIdAsync(string username, CancellationToken ct)
    {
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = "SELECT TenantId FROM dbo.AppUsers WHERE LOWER(Username) = LOWER(@U)";
            cmd.Parameters.AddWithValue("@U", username.Trim());
            var scalar = await cmd.ExecuteScalarAsync(ct);
            return scalar is Guid g ? g : null;
        }
        catch { return null; }
    }

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

        // ── Verificar licencia del tenant ─────────────────────────────────────
        try
        {
            await using var connLic = Conn();
            await connLic.OpenAsync(ct);
            await using var cmdLic = connLic.CreateCommand();
            cmdLic.CommandText = """
                SELECT t.IsActive, t.ExpiresAt, t.Name
                FROM dbo.AppUsers u
                JOIN dbo.Tenants t ON u.TenantId = t.Id
                WHERE LOWER(u.Username) = LOWER(@U)
                """;
            cmdLic.Parameters.AddWithValue("@U", req.Username.Trim());
            await using var rLic = await cmdLic.ExecuteReaderAsync(ct);
            if (await rLic.ReadAsync(ct))
            {
                var isActive  = rLic.GetBoolean(0);
                var expiresAt = rLic.IsDBNull(1) ? (DateTime?)null : rLic.GetDateTime(1);
                var tenantName = rLic.GetString(2);
                if (!isActive)
                    return StatusCode(403, new { error = "license_suspended", message = "El acceso de su organización ha sido suspendido. Contacte al proveedor." });
                if (expiresAt.HasValue && expiresAt.Value < DateTime.UtcNow)
                    return StatusCode(403, new { error = "license_expired", message = $"La licencia de {tenantName} venció el {expiresAt.Value:dd/MM/yyyy}. Contacte al proveedor para renovar." });
            }
        }
        catch (Exception ex) { logger.LogWarning(ex, "No se pudo verificar licencia para {User}", req.Username); }

        // ── Verificar 2FA (columna garantizada por startup) ───────────────────
        try
        {
            await using var conn2fa = Conn();
            await conn2fa.OpenAsync(ct);
            await using var cmd2fa = conn2fa.CreateCommand();
            cmd2fa.CommandText = "SELECT ISNULL(TwoFactorEnabled, 0) FROM dbo.AppUsers WHERE LOWER(Username) = LOWER(@U)";
            cmd2fa.Parameters.AddWithValue("@U", req.Username.Trim());
            var twoFaResult = await cmd2fa.ExecuteScalarAsync(ct);
            if (twoFaResult != null && twoFaResult != DBNull.Value && Convert.ToInt32(twoFaResult) == 1)
                return StatusCode(202, new { requires2FA = true, message = "Credenciales válidas. Se requiere verificación de dos factores." });
        }
        catch (Exception ex) { logger.LogWarning(ex, "No se pudo verificar 2FA para {User}", req.Username); }

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

        // Regenerar JWT con tenantId incluido
        var tenantId = await GetUserTenantIdAsync(req.Username, ct);

        // Extraer el token original de la respuesta de jwtService. La clave puede
        // venir como "token" o "Token" según la serialización — se busca sin
        // importar mayúsculas para no perderlo.
        var json0  = JsonSerializer.Serialize(response);
        var dict0  = JsonSerializer.Deserialize<Dictionary<string, JsonElement>>(json0)!;
        var tokKey = dict0.Keys.FirstOrDefault(k => k.Equals("token", StringComparison.OrdinalIgnoreCase));
        string? origToken = tokKey != null ? dict0[tokKey].GetString() : null;

        string? newToken = null;
        try
        {
            if (!string.IsNullOrWhiteSpace(origToken))
            {
                var handler   = new JwtSecurityTokenHandler();
                var parsed    = handler.ReadJwtToken(origToken);
                var username2 = parsed.Claims.FirstOrDefault(c => c.Type == ClaimTypes.Name || c.Type == "unique_name")?.Value ?? req.Username;
                var role2     = parsed.Claims.FirstOrDefault(c => c.Type == ClaimTypes.Role || c.Type == "role")?.Value ?? "User";
                var fullName2 = parsed.Claims.FirstOrDefault(c => c.Type == "fullName")?.Value ?? username2;
                var modules2  = int.TryParse(parsed.Claims.FirstOrDefault(c => c.Type == "modules")?.Value, out var mod) ? mod : 0;
                var userId2   = Guid.TryParse(parsed.Claims.FirstOrDefault(c => c.Type == JwtRegisteredClaimNames.Sub)?.Value, out var uid) ? uid : Guid.Empty;
                newToken = BuildJwt(username2, role2, fullName2, modules2, userId2, tenantId);
            }
        }
        catch (Exception ex) { logger.LogWarning(ex, "No se pudo regenerar JWT con tenantId"); }

        // Construir la respuesta final. Se GARANTIZA una clave "token" en minúscula
        // (el frontend lee data.token); se usa el token enriquecido con tenantId si
        // se pudo regenerar, si no el original. Se elimina cualquier clave "Token"
        // en otra capitalización para no dejar duplicados confusos.
        var dict = JsonSerializer.Deserialize<Dictionary<string, JsonElement>>(json0)!;
        foreach (var dup in dict.Keys.Where(k => k.Equals("token", StringComparison.OrdinalIgnoreCase) && k != "token").ToList())
            dict.Remove(dup);
        dict["refreshToken"] = JsonSerializer.SerializeToElement(refreshToken);
        dict["token"]        = JsonSerializer.SerializeToElement(newToken ?? origToken ?? "");
        if (tenantId.HasValue)
            dict["tenantId"] = JsonSerializer.SerializeToElement(tenantId.Value.ToString());
        return Ok(dict);
    }

    // ── POST /api/auth/google ─────────────────────────────────────────────────
    /// <summary>
    /// Login con cuenta institucional de Google (OpenID Connect).
    /// El frontend obtiene un id_token con Google Identity Services y lo envía aquí.
    /// El acceso lo controla el administrador: el usuario DEBE existir ya en Pandora
    /// (vinculado por correo). No se auto-crea ninguna cuenta.
    /// El proveedor de identidad (Google Workspace) es la autoridad de autenticación,
    /// por lo que este flujo NO exige el OTP propio de Pandora (2FA de la app).
    /// </summary>
    [EnableRateLimiting("login-policy")]
    [HttpPost("google")]
    [AllowAnonymous]
    public async Task<IActionResult> GoogleLogin([FromBody] GoogleLoginRequest req, CancellationToken ct)
    {
        if (string.IsNullOrWhiteSpace(req.IdToken))
            return BadRequest("Token de Google requerido.");

        // ── 1) Validar el id_token con Google (firma, expiración, emisor, audiencia) ──
        string[] clientIds = config.GetSection("GoogleAuth:ClientIds").Get<string[]>()
                             ?? (string.IsNullOrWhiteSpace(config["GoogleAuth:ClientId"])
                                    ? Array.Empty<string>()
                                    : new[] { config["GoogleAuth:ClientId"]! });
        if (clientIds.Length == 0)
        {
            logger.LogError("GoogleLogin: GoogleAuth:ClientId no está configurado.");
            return StatusCode(500, new { error = "google_not_configured", message = "El inicio de sesión con Google no está configurado." });
        }

        GoogleJsonWebSignature.Payload payload;
        try
        {
            payload = await GoogleJsonWebSignature.ValidateAsync(req.IdToken,
                new GoogleJsonWebSignature.ValidationSettings { Audience = clientIds });
        }
        catch (Exception ex)
        {
            logger.LogWarning(ex, "GoogleLogin: id_token inválido");
            return Unauthorized("No se pudo validar la cuenta de Google.");
        }

        if (payload.EmailVerified != true || string.IsNullOrWhiteSpace(payload.Email))
            return Unauthorized("La cuenta de Google no tiene un correo verificado.");

        var email = payload.Email.Trim();

        // ── Candado de dominio institucional ──────────────────────────────────
        // Si GoogleAuth:AllowedDomains está configurado, SOLO se aceptan correos
        // (y hosted domain de Workspace) de esos dominios. Vacío = sin restricción.
        var allowedDomains = config.GetSection("GoogleAuth:AllowedDomains").Get<string[]>()
                             ?? Array.Empty<string>();
        if (allowedDomains.Length > 0)
        {
            var atIdx       = email.IndexOf('@');
            var emailDomain = atIdx >= 0 ? email[(atIdx + 1)..].ToLowerInvariant() : "";
            var hostedDom   = (payload.HostedDomain ?? "").ToLowerInvariant();
            bool emailOk  = allowedDomains.Any(d => string.Equals(d.Trim(), emailDomain, StringComparison.OrdinalIgnoreCase));
            // Si Google reporta hosted domain (cuenta Workspace), también debe coincidir.
            bool hostedOk = string.IsNullOrEmpty(hostedDom)
                          || allowedDomains.Any(d => string.Equals(d.Trim(), hostedDom, StringComparison.OrdinalIgnoreCase));
            if (!emailOk || !hostedOk)
            {
                logger.LogWarning("GoogleLogin: dominio no permitido {Email} (hd={Hd})", email, hostedDom);
                return StatusCode(403, new { error = "domain_not_allowed", message = "Solo se permite el acceso con tu correo institucional." });
            }
        }

        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);

            // ── 2) Buscar al usuario por correo (lo aprueba/da de alta el admin) ──
            Guid userId = Guid.Empty; string username = ""; string fullName = ""; string role = "User";
            int modules = 0; Guid? tenantId = null; bool isActive = false; bool found = false;
            bool needsOnboarding = false;
            await using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = """
                    SELECT Id, Username, FullName, Role, Modules, TenantId, ISNULL(IsActive, 0),
                           ISNULL(NeedsOnboarding, 0)
                    FROM dbo.AppUsers WHERE LOWER(Email) = LOWER(@Email)
                    """;
                cmd.Parameters.AddWithValue("@Email", email);
                await using var r = await cmd.ExecuteReaderAsync(ct);
                if (await r.ReadAsync(ct))
                {
                    found    = true;
                    userId   = r.GetGuid(0);
                    username = r.IsDBNull(1) ? email : r.GetString(1);
                    fullName = r.IsDBNull(2) ? username : r.GetString(2);
                    role     = r.IsDBNull(3) ? "User" : r.GetString(3);
                    modules  = r.GetInt32(4);
                    tenantId = r.IsDBNull(5) ? null : r.GetGuid(5);
                    isActive = !r.IsDBNull(6) && r.GetBoolean(6);
                    needsOnboarding = !r.IsDBNull(7) && r.GetBoolean(7);
                }
            }

            // ── 2b) Sin cuenta → generar/actualizar SOLICITUD DE ACCESO pendiente ──
            if (!found)
            {
                var reqName = !string.IsNullOrWhiteSpace(req.FullName) ? req.FullName!.Trim()
                            : (!string.IsNullOrWhiteSpace(payload.Name) ? payload.Name!.Trim() : email);
                object reqPos = string.IsNullOrWhiteSpace(req.Position) ? DBNull.Value : req.Position!.Trim();
                object picture = string.IsNullOrWhiteSpace(payload.Picture) ? DBNull.Value : payload.Picture!;

                string? status;
                await using (var cs = conn.CreateCommand())
                {
                    cs.CommandText = "SELECT Status FROM dbo.AccessRequests WHERE LOWER(Email) = LOWER(@E)";
                    cs.Parameters.AddWithValue("@E", email);
                    status = (await cs.ExecuteScalarAsync(ct)) as string;
                }

                if (status == "rejected")
                    return StatusCode(403, new { error = "request_rejected", message = "Tu solicitud de acceso fue rechazada. Contacta al administrador." });

                await using (var up = conn.CreateCommand())
                {
                    if (status is null)
                    {
                        up.CommandText = """
                            INSERT INTO dbo.AccessRequests (Id, Email, FullName, Picture, Position, Status, CreatedAt)
                            VALUES (NEWID(), @Email, @FullName, @Picture, @Position, 'pending', GETUTCDATE())
                            """;
                    }
                    else // 'pending' (o estado residual) → refrescar datos capturados
                    {
                        up.CommandText = """
                            UPDATE dbo.AccessRequests
                            SET FullName = @FullName,
                                Position = COALESCE(@Position, Position),
                                Picture  = COALESCE(@Picture, Picture),
                                Status   = 'pending'
                            WHERE LOWER(Email) = LOWER(@Email)
                            """;
                    }
                    up.Parameters.AddWithValue("@Email", email);
                    up.Parameters.AddWithValue("@FullName", reqName);
                    up.Parameters.AddWithValue("@Picture", picture);
                    up.Parameters.AddWithValue("@Position", reqPos);
                    await up.ExecuteNonQueryAsync(ct);
                }

                logger.LogInformation("AccessRequest {State} para {Email}", status is null ? "NUEVA" : "actualizada", email);

                // Aviso por correo a los administradores (solo cuando es una solicitud NUEVA).
                if (status is null)
                {
                    var notifName = reqName;
                    _ = Task.Run(() => NotifyAdminsOfAccessRequestAsync(email, notifName), CancellationToken.None);
                }

                return Ok(new
                {
                    pending = true,
                    needsDetails = status is null,   // primera vez → el frontend pide completar datos
                    name = reqName,
                    email,
                    message = "Tu solicitud de acceso fue enviada. Un administrador debe aprobarla.",
                });
            }

            if (!isActive)
                return StatusCode(403, new { error = "user_disabled", message = "Tu cuenta está desactivada. Contacta al administrador." });

            // ── 3) Verificar licencia del tenant (igual que el login por contraseña) ──
            if (tenantId.HasValue)
            {
                await using var cmdLic = conn.CreateCommand();
                cmdLic.CommandText = "SELECT ISNULL(IsActive,1), ExpiresAt, Name FROM dbo.Tenants WHERE Id = @T";
                cmdLic.Parameters.AddWithValue("@T", tenantId.Value);
                await using var rLic = await cmdLic.ExecuteReaderAsync(ct);
                if (await rLic.ReadAsync(ct))
                {
                    var tActive = rLic.GetBoolean(0);
                    var tExp    = rLic.IsDBNull(1) ? (DateTime?)null : rLic.GetDateTime(1);
                    var tName   = rLic.IsDBNull(2) ? "tu organización" : rLic.GetString(2);
                    if (!tActive)
                        return StatusCode(403, new { error = "license_suspended", message = "El acceso de su organización ha sido suspendido. Contacte al proveedor." });
                    if (tExp.HasValue && tExp.Value < DateTime.UtcNow)
                        return StatusCode(403, new { error = "license_expired", message = $"La licencia de {tName} venció el {tExp.Value:dd/MM/yyyy}. Contacte al proveedor para renovar." });
                }
            }

            // ── 4) Emitir JWT + refresh token (misma mecánica que el login normal) ──
            string token       = BuildJwt(username, role, fullName, modules, userId, tenantId, needsOnboarding: needsOnboarding);
            string refreshToken = NewRefreshToken();
            await using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = """
                    DELETE FROM dbo.RefreshTokens WHERE Username = @Username;
                    INSERT INTO dbo.RefreshTokens (Token, Username, ExpiresAt, CreatedAt)
                    VALUES (@Token, @Username, @ExpiresAt, GETUTCDATE());
                    """;
                cmd.Parameters.AddWithValue("@Token",     refreshToken);
                cmd.Parameters.AddWithValue("@Username",  username.ToLower().Trim());
                cmd.Parameters.AddWithValue("@ExpiresAt", DateTime.UtcNow.AddDays(7));
                await cmd.ExecuteNonQueryAsync(ct);
            }

            logger.LogInformation("GoogleLogin OK para {Email} (usuario {Username})", email, username);
            return Ok(new
            {
                token,
                refreshToken,
                tenantId = tenantId?.ToString(),
                needsOnboarding,
            });
        }
        catch (Exception ex)
        {
            logger.LogError(ex, "GoogleLogin error for {Email}", email);
            return StatusCode(500, "Error al iniciar sesión con Google.");
        }
    }

    // ── POST /api/auth/complete-profile ───────────────────────────────────────
    /// <summary>
    /// Completa el perfil (nombre/puesto) de una cuenta precargada en su primer
    /// ingreso. Actualiza AppUsers, limpia NeedsOnboarding y emite un token fresco
    /// (ya sin la bandera) para que los datos queden homologados en todo Pandora.
    /// </summary>
    [Authorize]
    [HttpPost("complete-profile")]
    public async Task<IActionResult> CompleteProfile([FromBody] CompleteProfileDto dto, CancellationToken ct)
    {
        if (string.IsNullOrWhiteSpace(dto.FullName))
            return BadRequest("El nombre es obligatorio.");

        var uname = User.FindFirstValue(ClaimTypes.Name);
        if (string.IsNullOrWhiteSpace(uname)) return Unauthorized();

        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);

            await using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = """
                    UPDATE dbo.AppUsers
                    SET FullName = @FullName, Position = @Position,
                        NeedsOnboarding = 0, UpdatedAt = GETUTCDATE()
                    WHERE LOWER(Username) = LOWER(@U)
                    """;
                cmd.Parameters.AddWithValue("@FullName", dto.FullName.Trim());
                cmd.Parameters.AddWithValue("@Position", (object?)dto.Position?.Trim() ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@U", uname);
                if (await cmd.ExecuteNonQueryAsync(ct) == 0) return NotFound("Usuario no encontrado.");
            }

            Guid userId; string role; int modules; Guid? tenantId; string fullName;
            await using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = "SELECT Id, ISNULL(Role,'User'), ISNULL(Modules,0), TenantId, FullName FROM dbo.AppUsers WHERE LOWER(Username) = LOWER(@U)";
                cmd.Parameters.AddWithValue("@U", uname);
                await using var r = await cmd.ExecuteReaderAsync(ct);
                if (!await r.ReadAsync(ct)) return NotFound();
                userId   = r.GetGuid(0);
                role     = r.GetString(1);
                modules  = r.GetInt32(2);
                tenantId = r.IsDBNull(3) ? null : r.GetGuid(3);
                fullName = r.IsDBNull(4) ? uname : r.GetString(4);
            }

            string token = BuildJwt(uname, role, fullName, modules, userId, tenantId);
            logger.LogInformation("CompleteProfile OK para {User}", uname);
            return Ok(new { token });
        }
        catch (Exception ex) { logger.LogError(ex, "CompleteProfile {User}", uname); return StatusCode(500, ex.Message); }
    }

    // Aviso por correo a los administradores cuando llega una solicitud de acceso nueva.
    // Best-effort: se dispara en segundo plano y nunca rompe el login.
    private async Task NotifyAdminsOfAccessRequestAsync(string email, string name)
    {
        try
        {
            var connStr = config.GetConnectionString("PandoraDb")!;
            var recipients = new List<string>();
            await using (var conn = new SqlConnection(connStr))
            {
                await conn.OpenAsync();
                await using var cmd = conn.CreateCommand();
                cmd.CommandText = """
                    SELECT DISTINCT Email FROM dbo.AppUsers
                    WHERE Role IN ('Admin','SuperAdmin') AND ISNULL(IsActive,0) = 1
                      AND Email IS NOT NULL AND LTRIM(RTRIM(Email)) <> ''
                    """;
                await using var r = await cmd.ExecuteReaderAsync();
                while (await r.ReadAsync()) recipients.Add(r.GetString(0));
            }
            if (recipients.Count == 0) return;

            var smtp = await SmtpHelper.LoadAsync(connStr, config);
            var safeName  = System.Net.WebUtility.HtmlEncode(name);
            var safeEmail = System.Net.WebUtility.HtmlEncode(email);
            var subject = "Pandora — Nueva solicitud de acceso";
            var body = $"""
                <div style="font-family:Arial,sans-serif;color:#333">
                  <h2 style="color:#1a237e;margin:0 0 12px">Nueva solicitud de acceso</h2>
                  <p><b>{safeName}</b> solicitó acceso a Pandora con su cuenta institucional.</p>
                  <table style="border-collapse:collapse">
                    <tr><td style="padding:6px;color:#555">Correo</td><td style="padding:6px">{safeEmail}</td></tr>
                  </table>
                  <p style="margin-top:12px">Entra a <b>Solicitudes de acceso</b> en Pandora para aprobarla o rechazarla.</p>
                </div>
                """;
            foreach (var to in recipients)
                _ = await SmtpHelper.SendAsync(smtp, to, to, subject, body);
            logger.LogInformation("Aviso de solicitud de acceso enviado a {N} admin(s) por {Email}", recipients.Count, email);
        }
        catch (Exception ex) { logger.LogWarning(ex, "No se pudo notificar la solicitud de acceso de {Email}", email); }
    }

    // ── POST /api/auth/impersonate/{tenantId} ─────────────────────────────────
    /// <summary>
    /// "Entrar como" un cliente: el SuperAdmin obtiene un token de Admin scoped
    /// al tenant indicado. Con ese token actúa como administrador de ESE cliente
    /// y todo el aislamiento por tenant + los [Authorize(Roles="Admin")] aplican
    /// igual. El token lleva el claim "impersonatedBy" para auditoría. No se
    /// persiste refresh token (no debe sobrescribir la sesión real del SuperAdmin);
    /// al expirar, el SuperAdmin vuelve a "entrar como".
    /// </summary>
    [HttpPost("impersonate/{tenantId:guid}")]
    [Authorize(Roles = "SuperAdmin")]
    public async Task<IActionResult> Impersonate(Guid tenantId, CancellationToken ct)
    {
        var saUser     = User.FindFirstValue(ClaimTypes.Name) ?? "";
        var saUserId   = Guid.TryParse(User.FindFirstValue(JwtRegisteredClaimNames.Sub), out var g) ? g : Guid.Empty;
        var saFullName = User.FindFirstValue("fullName") ?? saUser;

        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            string tenantName;
            await using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = "SELECT Name FROM dbo.Tenants WHERE Id = @T";
                cmd.Parameters.AddWithValue("@T", tenantId);
                var nameObj = await cmd.ExecuteScalarAsync(ct);
                if (nameObj is null or DBNull) return NotFound(new { error = "tenant_not_found", message = "Cliente no encontrado." });
                tenantName = (string)nameObj;
            }

            // Token de Admin scoped al tenant, actuando como el SuperAdmin (auditoría)
            var token = BuildJwt(saUser, "Admin", saFullName, modules: -1, userId: saUserId,
                                 tenantId: tenantId, impersonatedBy: saUser);

            logger.LogInformation("SuperAdmin {User} impersona al tenant {Tenant} ({Id})", saUser, tenantName, tenantId);
            return Ok(new { token, tenantId = tenantId.ToString(), tenantName });
        }
        catch (Exception ex)
        {
            logger.LogError(ex, "Impersonate tenant {Id}", tenantId);
            return StatusCode(500, "Error al entrar como el cliente.");
        }
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

            Guid? tenantIdRefresh = null;
            await using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = """
                    SELECT Id, FullName, Role, Modules, TenantId
                    FROM dbo.AppUsers
                    WHERE LOWER(Username) = LOWER(@Username) AND IsActive = 1
                    """;
                cmd.Parameters.AddWithValue("@Username", username);
                await using var r = await cmd.ExecuteReaderAsync(ct);
                if (!await r.ReadAsync(ct))
                    return Unauthorized("Usuario desactivado o no encontrado.");

                userId           = r.GetGuid(0);
                fullName         = r.IsDBNull(1) ? username : r.GetString(1);
                role             = r.IsDBNull(2) ? "User"   : r.GetString(2);
                modules          = r.GetInt32(3);
                tenantIdRefresh  = r.IsDBNull(4) ? null : r.GetGuid(4);
            }

            // ── Generar nuevo JWT con tenantId ────────────────────────────────
            string newJwt = BuildJwt(username, role, fullName, modules, userId, tenantIdRefresh);

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
            // IMPORTANTE: usar el mismo esquema (PBKDF2 "salt:hash") que
            // UserService.HashPassword/VerifyPassword — no BCrypt. El login
            // real (Pandora.Infrastructure.Services.JwtService) verifica vía
            // UserService.VerifyPassword, que espera ese formato exacto;
            // un hash BCrypt no tiene ':' y siempre falla la verificación,
            // dejando al usuario bloqueado tras usar "Olvidé mi contraseña".
            string hash = UserService.HashPassword(req.NewPassword);
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

    // ── POST /api/auth/send-otp ──────────────────────────────────────────────
    // Rate-limited: evita spam de correos OTP a la cuenta del usuario.
    [EnableRateLimiting("login-policy")]
    [HttpPost("send-otp")]
    [AllowAnonymous]
    public async Task<IActionResult> SendOtp([FromBody] OtpRequest req, CancellationToken ct)
    {
        if (string.IsNullOrWhiteSpace(req.Username)) return BadRequest("Usuario requerido.");
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            string? email = null; string? fullName = null;
            await using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = """
                    SELECT Email, FullName FROM dbo.AppUsers
                    WHERE LOWER(Username) = LOWER(@Username) AND IsActive = 1
                      AND ISNULL(TwoFactorEnabled, 0) = 1
                    """;
                cmd.Parameters.AddWithValue("@Username", req.Username.Trim());
                await using var r = await cmd.ExecuteReaderAsync(ct);
                if (await r.ReadAsync(ct)) { email = r.IsDBNull(0) ? null : r.GetString(0); fullName = r.IsDBNull(1) ? null : r.GetString(1); }
            }
            if (email is null) return BadRequest("Usuario no encontrado o 2FA no habilitado.");

            string code = Random.Shared.Next(100000, 999999).ToString();
            await using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = """
                    DELETE FROM dbo.OtpCodes WHERE LOWER(Username) = LOWER(@Username);
                    INSERT INTO dbo.OtpCodes (Username, Code, ExpiresAt, Used) VALUES (@Username, @Code, @Exp, 0);
                    """;
                cmd.Parameters.AddWithValue("@Username", req.Username.Trim().ToLower());
                cmd.Parameters.AddWithValue("@Code",     code);
                cmd.Parameters.AddWithValue("@Exp",      DateTime.UtcNow.AddMinutes(10));
                await cmd.ExecuteNonQueryAsync(ct);
            }
            var smtp = await SmtpHelper.LoadAsync(config.GetConnectionString("PandoraDb")!, config);
            await SmtpHelper.SendAsync(smtp, email, fullName ?? req.Username, "Código de verificación — Pandora",
                $"<p>Hola <strong>{fullName ?? req.Username}</strong>,</p>" +
                $"<p>Tu código de verificación es: <strong style='font-size:24px;letter-spacing:6px'>{code}</strong></p>" +
                $"<p>Válido por <strong>10 minutos</strong>. Si no solicitaste este código, ignora este mensaje.</p>");
            return Ok(new { message = "Código enviado al correo registrado." });
        }
        catch (Exception ex) { logger.LogError(ex, "SendOtp {User}", req.Username); return StatusCode(500, "Error al enviar el código."); }
    }

    // ── POST /api/auth/verify-otp ────────────────────────────────────────────
    // Rate-limited: evita fuerza bruta del código OTP de 6 dígitos.
    [EnableRateLimiting("login-policy")]
    [HttpPost("verify-otp")]
    [AllowAnonymous]
    public async Task<IActionResult> VerifyOtp([FromBody] VerifyOtpRequest req, CancellationToken ct)
    {
        if (string.IsNullOrWhiteSpace(req.Username) || string.IsNullOrWhiteSpace(req.Code))
            return BadRequest("Usuario y código requeridos.");
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);

            // Atomic: validate + consume in one UPDATE — prevents OTP replay on concurrent requests
            await using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = """
                    UPDATE dbo.OtpCodes SET Used=1
                    WHERE LOWER(Username)=LOWER(@Username) AND Code=@Code AND Used=0 AND ExpiresAt>GETUTCDATE()
                    """;
                cmd.Parameters.AddWithValue("@Username", req.Username.Trim());
                cmd.Parameters.AddWithValue("@Code",     req.Code.Trim());
                int affected = await cmd.ExecuteNonQueryAsync(ct);
                if (affected == 0) return Unauthorized("Código inválido o expirado.");
            }

            // Emitir JWT
            string? role = null; int modules = 0; Guid userId = Guid.Empty; string? fn = null;
            await using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = "SELECT Id, Role, Modules, FullName FROM dbo.AppUsers WHERE LOWER(Username)=LOWER(@Username) AND IsActive=1";
                cmd.Parameters.AddWithValue("@Username", req.Username.Trim());
                await using var r = await cmd.ExecuteReaderAsync(ct);
                if (await r.ReadAsync(ct)) { userId = r.GetGuid(0); role = r.GetString(1); modules = r.GetInt32(2); fn = r.IsDBNull(3) ? null : r.GetString(3); }
            }
            if (role is null) return Unauthorized("Usuario no encontrado.");
            string uname2 = req.Username.Trim().ToLower();
            var token     = BuildJwt(uname2, role, fn ?? uname2, modules, userId);
            string refreshToken = NewRefreshToken();
            await using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = """
                    DELETE FROM dbo.RefreshTokens WHERE Username=@Username;
                    INSERT INTO dbo.RefreshTokens (Token,Username,ExpiresAt,CreatedAt) VALUES (@Token,@Username,@Exp,GETUTCDATE());
                    """;
                cmd.Parameters.AddWithValue("@Token",    refreshToken);
                cmd.Parameters.AddWithValue("@Username", req.Username.Trim().ToLower());
                cmd.Parameters.AddWithValue("@Exp",      DateTime.UtcNow.AddDays(7));
                await cmd.ExecuteNonQueryAsync(ct);
            }
            return Ok(new { token, refreshToken });
        }
        catch (Exception ex) { logger.LogError(ex, "VerifyOtp {User}", req.Username); return StatusCode(500, "Error al verificar el código."); }
    }

    // ── POST /api/auth/toggle-2fa/{username} ─────────────────────────────────
    [HttpPost("toggle-2fa/{username}")]
    [Authorize(Roles = "Admin")]
    public async Task<IActionResult> Toggle2FA(string username, [FromBody] Toggle2FADto dto, CancellationToken ct)
    {
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = "UPDATE dbo.AppUsers SET TwoFactorEnabled=@Enable WHERE LOWER(Username)=LOWER(@Username)";
            cmd.Parameters.AddWithValue("@Enable",   dto.Enable ? 1 : 0);
            cmd.Parameters.AddWithValue("@Username", username);
            await cmd.ExecuteNonQueryAsync(ct);
            return Ok(new { enabled = dto.Enable });
        }
        catch (Exception ex) { return StatusCode(500, ex.Message); }
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
                      <p style="color:rgba(255,255,255,.7);margin:4px 0 0;font-size:13px">Sistema de Gestión</p>
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
                      <p style="color:#aaa;font-size:11px;margin:0">Pandora — Coordinación de TI</p>
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
public record GoogleLoginRequest(string IdToken, string? FullName = null, string? Position = null);
public record CompleteProfileDto(string FullName, string? Position);
public record ForgotPasswordRequest(string Email);
public record ResetPasswordRequest(string Token, string NewPassword);
public record OtpRequest(string Username);
public record VerifyOtpRequest(string Username, string Code, string? TempPassword);
public record Toggle2FADto(bool Enable);
