using System.IdentityModel.Tokens.Jwt;
using System.Security.Claims;
using System.Text;
using System.Text.Json;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using Microsoft.IdentityModel.Tokens;

namespace Pandora.API.Controllers;

/// <summary>
/// Integración con el microservicio Node "google-workspace-provisioning"
/// (alta masiva de cuentas institucionales en Google Workspace, a partir
/// de un Excel de alumnos). Sigue el mismo patrón que
/// <see cref="AzulSsoController"/>: Pandora emite un JWT de corta duración
/// para que el servicio externo lo valide — nada de API keys estáticas.
///
/// Autenticación en AMBOS sentidos con el mismo secreto compartido
/// (config "GoogleWorkspaceProvisioning:ServiceKey", HS256):
///   - Pandora -> Node : POST /token (solo Admin), audience "google-workspace-provisioning-svc".
///   - Node -> Pandora : POST/PATCH acá abajo, audience "google-workspace-provisioning-callback",
///     firmado por el propio microservicio Node (ver src/auth/serviceJwt.js).
/// </summary>
[ApiController]
public class GoogleWorkspaceProvisioningController(
    IConfiguration config,
    IHttpClientFactory httpClientFactory,
    ILogger<GoogleWorkspaceProvisioningController> logger) : ControllerBase
{
    private const string IncomingAudienceForNode = "google-workspace-provisioning-svc";     // Pandora -> Node
    private const string IncomingAudienceFromNode = "google-workspace-provisioning-callback"; // Node -> Pandora
    private const int ServiceTokenMinutes = 10;

    private string ServiceKey => config["GoogleWorkspaceProvisioning:ServiceKey"] ?? "";
    private string? ServiceUrl => config["GoogleWorkspaceProvisioning:ServiceUrl"];

    private SqlConnection Conn() => new(config.GetConnectionString("PandoraDb"));

    /// <summary>
    /// Firma el JWT de servicio (audience "google-workspace-provisioning-svc")
    /// que valida el microservicio Node. Compartido entre el endpoint público
    /// <see cref="GenerateServiceToken"/> y el uso interno en
    /// <see cref="UploadBatch"/> (que llama a Node por dentro, sin exponer el
    /// token al frontend).
    /// </summary>
    private string BuildServiceToken(string requestedBy)
    {
        var key = new SymmetricSecurityKey(Encoding.UTF8.GetBytes(ServiceKey));
        var creds = new SigningCredentials(key, SecurityAlgorithms.HmacSha256);
        var jwtObj = new JwtSecurityToken(
            audience: IncomingAudienceForNode,
            claims: [new Claim("requestedBy", requestedBy)],
            expires: DateTime.UtcNow.AddMinutes(ServiceTokenMinutes),
            signingCredentials: creds);
        return new JwtSecurityTokenHandler().WriteToken(jwtObj);
    }

    // ── POST /api/google-workspace-provisioning/token ───────────────────────
    /// <summary>
    /// Emite un JWT corto (10 min) para que Pandora llame al microservicio
    /// Node. Solo Admin — dar de alta cuentas de alumnos es una operación
    /// sensible (datos de menores).
    /// </summary>
    [HttpPost("api/google-workspace-provisioning/token")]
    [Authorize(Roles = "Admin")]
    public IActionResult GenerateServiceToken()
    {
        if (string.IsNullOrWhiteSpace(ServiceKey))
        {
            logger.LogError("GoogleWorkspaceProvisioning:ServiceKey no configurado.");
            return StatusCode(500, new { error = "Integración no configurada." });
        }

        var adminUsername = User.FindFirstValue(ClaimTypes.Name) ?? "admin";
        var token = BuildServiceToken(adminUsername);
        return Ok(new { token, expiresInSeconds = ServiceTokenMinutes * 60 });
    }

    // ── POST /api/google-workspace-provisioning/upload ──────────────────────
    /// <summary>
    /// Recibe el Excel subido desde el Panel Admin y lo reenvía tal cual al
    /// microservicio Node (evita duplicar el parser de Excel en C#; ver
    /// README del microservicio "Por qué JSON y no multipart"). Solo Admin.
    /// </summary>
    [HttpPost("api/google-workspace-provisioning/upload")]
    [Authorize(Roles = "Admin")]
    [RequestSizeLimit(15 * 1024 * 1024)]
    public async Task<IActionResult> UploadBatch(IFormFile file, CancellationToken ct)
    {
        if (file is null || file.Length == 0)
            return BadRequest(new { error = "Archivo .xlsx requerido." });

        if (string.IsNullOrWhiteSpace(ServiceKey) || string.IsNullOrWhiteSpace(ServiceUrl))
        {
            logger.LogError("GoogleWorkspaceProvisioning:ServiceKey/ServiceUrl no configurados.");
            return StatusCode(500, new { error = "Integración no configurada." });
        }

        var adminUsername = User.FindFirstValue(ClaimTypes.Name) ?? "admin";
        var token = BuildServiceToken(adminUsername);

        try
        {
            using var client = httpClientFactory.CreateClient();
            client.BaseAddress = new Uri(ServiceUrl);
            client.DefaultRequestHeaders.Authorization =
                new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", token);

            using var content = new MultipartFormDataContent();
            await using var stream = file.OpenReadStream();
            using var streamContent = new StreamContent(stream);
            content.Add(streamContent, "file", file.FileName);

            using var response = await client.PostAsync("/api/v1/provisioning/batch/upload", content, ct);
            var responseBody = await response.Content.ReadAsStringAsync(ct);

            if (!response.IsSuccessStatusCode)
            {
                logger.LogWarning(
                    "google-workspace-provisioning respondió {Status} al subir el Excel", response.StatusCode);
                return StatusCode((int)response.StatusCode, System.Text.Json.JsonSerializer.Deserialize<JsonElement>(responseBody));
            }

            return Accepted(System.Text.Json.JsonSerializer.Deserialize<JsonElement>(responseBody));
        }
        catch (Exception ex)
        {
            // Sanitizado: no se expone la URL interna ni detalles de red al frontend.
            logger.LogError(ex, "No se pudo contactar al microservicio google-workspace-provisioning");
            return StatusCode(502, new { error = "El servicio de aprovisionamiento no está disponible. Intenta de nuevo en unos minutos." });
        }
    }

    // ── GET /api/google-workspace-provisioning/jobs/{id} ────────────────────
    /// <summary>
    /// Estado del job — se lee directo de la tabla local (espejada por el
    /// microservicio Node en cada cambio de estado), sin llamar a Node de
    /// nuevo. Más rápido y sigue funcionando aunque Node esté momentáneamente
    /// inalcanzable.
    /// </summary>
    [HttpGet("api/google-workspace-provisioning/jobs/{id}")]
    [Authorize(Roles = "Admin")]
    public async Task<IActionResult> GetJob(string id, CancellationToken ct)
    {
        await using var conn = Conn();
        await conn.OpenAsync(ct);
        await using var cmd = conn.CreateCommand();
        cmd.CommandText = """
            SELECT JobId, Status, Total, Completed, Failed, CreatedBy, CreatedAt, UpdatedAt
            FROM dbo.GoogleWorkspaceProvisioningJobs WHERE JobId = @JobId
            """;
        cmd.Parameters.AddWithValue("@JobId", id);

        await using var reader = await cmd.ExecuteReaderAsync(ct);
        if (!await reader.ReadAsync(ct))
            return NotFound(new { error = "Job no encontrado." });

        return Ok(new
        {
            jobId = reader.GetGuid(0),
            status = reader.GetString(1),
            total = reader.GetInt32(2),
            completed = reader.GetInt32(3),
            failed = reader.GetInt32(4),
            createdBy = reader.GetString(5),
            createdAt = reader.GetDateTime(6),
            updatedAt = reader.GetDateTime(7),
        });
    }

    // ── GET /api/google-workspace-provisioning/jobs/{id}/audit ──────────────
    /// <summary>Resultado por alumno de un job — nunca incluye la contraseña.</summary>
    [HttpGet("api/google-workspace-provisioning/jobs/{id}/audit")]
    [Authorize(Roles = "Admin")]
    public async Task<IActionResult> GetJobAudit(string id, CancellationToken ct)
    {
        await using var conn = Conn();
        await conn.OpenAsync(ct);
        await using var cmd = conn.CreateCommand();
        cmd.CommandText = """
            SELECT Matricula, PrimaryEmail, Resultado, Detalle, CreatedAt
            FROM dbo.GoogleWorkspaceProvisioningAuditLog
            WHERE JobId = @JobId
            ORDER BY CreatedAt
            """;
        cmd.Parameters.AddWithValue("@JobId", id);

        var results = new List<object>();
        await using var reader = await cmd.ExecuteReaderAsync(ct);
        while (await reader.ReadAsync(ct))
        {
            results.Add(new
            {
                matricula = reader.GetString(0),
                primaryEmail = reader.GetString(1),
                resultado = reader.GetString(2),
                detalle = reader.IsDBNull(3) ? null : reader.GetString(3),
                createdAt = reader.GetDateTime(4),
            });
        }
        return Ok(results);
    }

    // ── POST /api/internal/google-workspace-provisioning/jobs ───────────────
    /// <summary>Mirror de creación/estado de un job, reportado por el microservicio Node.</summary>
    [HttpPost("api/internal/google-workspace-provisioning/jobs")]
    [AllowAnonymous]
    public async Task<IActionResult> UpsertJob([FromBody] JsonElement body, CancellationToken ct)
    {
        if (!TryValidateServiceCallback(out var unauthorized)) return unauthorized!;
        await UpsertJobInternalAsync(body, ct);
        return Ok();
    }

    // ── PATCH /api/internal/google-workspace-provisioning/jobs/{id} ─────────
    [HttpPatch("api/internal/google-workspace-provisioning/jobs/{id}")]
    [AllowAnonymous]
    public async Task<IActionResult> UpdateJob(string id, [FromBody] JsonElement body, CancellationToken ct)
    {
        if (!TryValidateServiceCallback(out var unauthorized)) return unauthorized!;
        await UpsertJobInternalAsync(body, ct);
        return Ok();
    }

    private async Task UpsertJobInternalAsync(JsonElement body, CancellationToken ct)
    {
        string S(string prop) => body.TryGetProperty(prop, out var v) ? v.GetString() ?? "" : "";
        int I(string prop) => body.TryGetProperty(prop, out var v) && v.TryGetInt32(out var n) ? n : 0;

        await using var conn = Conn();
        await conn.OpenAsync(ct);
        await using var cmd = conn.CreateCommand();
        cmd.CommandText = """
            MERGE dbo.GoogleWorkspaceProvisioningJobs AS target
            USING (SELECT @JobId AS JobId) AS src
            ON target.JobId = src.JobId
            WHEN MATCHED THEN UPDATE SET
                Status = @Status, Total = @Total, Completed = @Completed,
                Failed = @Failed, UpdatedAt = SYSUTCDATETIME()
            WHEN NOT MATCHED THEN
                INSERT (JobId, Status, Total, Completed, Failed, CreatedBy, CreatedAt, UpdatedAt)
                VALUES (@JobId, @Status, @Total, @Completed, @Failed, @CreatedBy, SYSUTCDATETIME(), SYSUTCDATETIME());
            """;
        cmd.Parameters.AddWithValue("@JobId", S("jobId"));
        cmd.Parameters.AddWithValue("@Status", S("status"));
        cmd.Parameters.AddWithValue("@Total", I("total"));
        cmd.Parameters.AddWithValue("@Completed", I("completed"));
        cmd.Parameters.AddWithValue("@Failed", I("failed"));
        cmd.Parameters.AddWithValue("@CreatedBy", string.IsNullOrWhiteSpace(S("createdBy")) ? "unknown" : S("createdBy"));
        await cmd.ExecuteNonQueryAsync(ct);
    }

    // ── POST /api/internal/google-workspace-provisioning/audit ──────────────
    /// <summary>
    /// Log de auditoría inmutable por alumno (checklist B.6) — nunca recibe
    /// la contraseña, solo el resultado. Append-only: no hay endpoint DELETE.
    /// </summary>
    [HttpPost("api/internal/google-workspace-provisioning/audit")]
    [AllowAnonymous]
    public async Task<IActionResult> ReportAudit([FromBody] JsonElement body, CancellationToken ct)
    {
        if (!TryValidateServiceCallback(out var unauthorized)) return unauthorized!;

        string S(string prop) => body.TryGetProperty(prop, out var v) ? v.GetString() ?? "" : "";

        await using var conn = Conn();
        await conn.OpenAsync(ct);
        await using var cmd = conn.CreateCommand();
        cmd.CommandText = """
            INSERT INTO dbo.GoogleWorkspaceProvisioningAuditLog (JobId, Matricula, PrimaryEmail, Resultado, Detalle, CreatedAt)
            VALUES (@JobId, @Matricula, @PrimaryEmail, @Resultado, @Detalle, SYSUTCDATETIME());
            """;
        cmd.Parameters.AddWithValue("@JobId", S("jobId"));
        cmd.Parameters.AddWithValue("@Matricula", S("matricula"));
        cmd.Parameters.AddWithValue("@PrimaryEmail", S("primaryEmail"));
        cmd.Parameters.AddWithValue("@Resultado", S("resultado"));
        cmd.Parameters.AddWithValue("@Detalle", (object?)S("detalle") ?? DBNull.Value);
        await cmd.ExecuteNonQueryAsync(ct);
        return Ok();
    }

    /// <summary>
    /// Valida el JWT que firma el propio microservicio Node al llamar de
    /// vuelta (audience "google-workspace-provisioning-callback"), con el
    /// mismo secreto compartido que Pandora usa para emitir el suyo hacia
    /// Node.
    /// </summary>
    private bool TryValidateServiceCallback(out IActionResult? unauthorized)
    {
        unauthorized = null;
        var header = Request.Headers.Authorization.ToString();
        if (string.IsNullOrWhiteSpace(header) || !header.StartsWith("Bearer "))
        {
            unauthorized = Unauthorized(new { error = "Falta el header Authorization: Bearer <token>." });
            return false;
        }

        if (string.IsNullOrWhiteSpace(ServiceKey))
        {
            logger.LogError("GoogleWorkspaceProvisioning:ServiceKey no configurado.");
            unauthorized = StatusCode(500, new { error = "Integración no configurada." });
            return false;
        }

        var token = header["Bearer ".Length..].Trim();
        var handler = new JwtSecurityTokenHandler();
        try
        {
            handler.ValidateToken(token, new TokenValidationParameters
            {
                ValidateIssuerSigningKey = true,
                IssuerSigningKey = new SymmetricSecurityKey(Encoding.UTF8.GetBytes(ServiceKey)),
                ValidateAudience = true,
                ValidAudience = IncomingAudienceFromNode,
                ValidateIssuer = false,
                ValidateLifetime = true,
                ClockSkew = TimeSpan.FromSeconds(30),
            }, out _);
            return true;
        }
        catch (Exception ex)
        {
            logger.LogWarning(ex, "Token de callback de google-workspace-provisioning inválido");
            unauthorized = Unauthorized(new { error = "Token de servicio inválido o expirado." });
            return false;
        }
    }
}
