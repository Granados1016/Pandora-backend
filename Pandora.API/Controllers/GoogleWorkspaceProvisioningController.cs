using System.IdentityModel.Tokens.Jwt;
using System.Linq;
using System.Security.Claims;
using System.Text;
using System.Text.Json;
using ClosedXML.Excel;
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

    // ── GET /api/google-workspace-provisioning/directory/buscar ─────────────
    /// <summary>
    /// Busca en el directorio REAL de Google Workspace (proxy al
    /// microservicio Node) — a diferencia de <see cref="BuscarAudit"/>, esto
    /// incluye cuentas creadas antes de que existiera este módulo, porque
    /// consulta a Google directamente en vez del historial propio de Pandora.
    /// </summary>
    [HttpGet("api/google-workspace-provisioning/directory/buscar")]
    [Authorize(Roles = "Admin")]
    public async Task<IActionResult> BuscarDirectorio([FromQuery] string q, CancellationToken ct)
    {
        if (string.IsNullOrWhiteSpace(q) || q.Trim().Length < 2)
            return BadRequest(new { error = "Escribe al menos 2 caracteres para buscar." });

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

            var query = Uri.EscapeDataString(q.Trim());
            using var response = await client.GetAsync($"/api/v1/provisioning/users?q={query}", ct);
            var responseBody = await response.Content.ReadAsStringAsync(ct);

            if (!response.IsSuccessStatusCode)
            {
                logger.LogWarning(
                    "google-workspace-provisioning respondió {Status} al buscar en el directorio", response.StatusCode);
                return StatusCode((int)response.StatusCode, System.Text.Json.JsonSerializer.Deserialize<JsonElement>(responseBody));
            }

            var users = System.Text.Json.JsonSerializer.Deserialize<List<JsonElement>>(responseBody) ?? [];
            var mapped = users.Select(u =>
            {
                string email = u.GetProperty("primaryEmail").GetString() ?? "";
                string matricula = email.Contains('@') ? email[..email.IndexOf('@')] : email;
                bool suspended = u.TryGetProperty("suspended", out var s) && s.GetBoolean();
                return new
                {
                    matricula,
                    nombre = u.TryGetProperty("givenName", out var gn) ? gn.GetString() : null,
                    apellidos = u.TryGetProperty("familyName", out var fn) ? fn.GetString() : null,
                    primaryEmail = email,
                    resultado = suspended ? "suspendida" : "ya_existia",
                    detalle = suspended ? "Cuenta suspendida en Google Workspace" : (string?)null,
                    createdAt = u.TryGetProperty("creationTime", out var ct2) ? ct2.GetString() : null,
                };
            }).ToList();

            return Ok(mapped);
        }
        catch (Exception ex)
        {
            logger.LogError(ex, "No se pudo contactar al microservicio google-workspace-provisioning (búsqueda de directorio)");
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
            SELECT Matricula, Nombre, Apellidos, PrimaryEmail, Resultado, Detalle, CreatedAt
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
                nombre = reader.IsDBNull(1) ? null : reader.GetString(1),
                apellidos = reader.IsDBNull(2) ? null : reader.GetString(2),
                primaryEmail = reader.GetString(3),
                resultado = reader.GetString(4),
                detalle = reader.IsDBNull(5) ? null : reader.GetString(5),
                createdAt = reader.GetDateTime(6),
            });
        }
        return Ok(results);
    }

    // ── GET /api/google-workspace-provisioning/audit/buscar ─────────────────
    /// <summary>
    /// Búsqueda global por matrícula/nombre/apellidos/correo, sin importar
    /// en qué lote (job) se creó — para que el admin pueda confirmar si un
    /// alumno ya tiene cuenta sin acordarse cuándo se subió.
    /// </summary>
    [HttpGet("api/google-workspace-provisioning/audit/buscar")]
    [Authorize(Roles = "Admin")]
    public async Task<IActionResult> BuscarAudit([FromQuery] string q, CancellationToken ct)
    {
        if (string.IsNullOrWhiteSpace(q) || q.Trim().Length < 2)
            return BadRequest(new { error = "Escribe al menos 2 caracteres para buscar." });

        await using var conn = Conn();
        await conn.OpenAsync(ct);
        await using var cmd = conn.CreateCommand();
        cmd.CommandText = """
            SELECT TOP 100 JobId, Matricula, Nombre, Apellidos, PrimaryEmail, Resultado, Detalle, CreatedAt
            FROM dbo.GoogleWorkspaceProvisioningAuditLog
            WHERE Matricula LIKE @Q OR Nombre LIKE @Q OR Apellidos LIKE @Q OR PrimaryEmail LIKE @Q
            ORDER BY CreatedAt DESC
            """;
        cmd.Parameters.AddWithValue("@Q", $"%{q.Trim()}%");

        var results = new List<object>();
        await using var reader = await cmd.ExecuteReaderAsync(ct);
        while (await reader.ReadAsync(ct))
        {
            results.Add(new
            {
                jobId = reader.GetString(0),
                matricula = reader.GetString(1),
                nombre = reader.IsDBNull(2) ? null : reader.GetString(2),
                apellidos = reader.IsDBNull(3) ? null : reader.GetString(3),
                primaryEmail = reader.GetString(4),
                resultado = reader.GetString(5),
                detalle = reader.IsDBNull(6) ? null : reader.GetString(6),
                createdAt = reader.GetDateTime(7),
            });
        }
        return Ok(results);
    }

    // ── GET /api/google-workspace-provisioning/jobs/{id}/audit/exportar ─────
    /// <summary>Exporta el resultado del job a Excel — mismo diseño que Licencias.</summary>
    [HttpGet("api/google-workspace-provisioning/jobs/{id}/audit/exportar")]
    [Authorize(Roles = "Admin")]
    public async Task<IActionResult> ExportJobAudit(string id, CancellationToken ct)
    {
        await using var conn = Conn();
        await conn.OpenAsync(ct);
        await using var cmd = conn.CreateCommand();
        cmd.CommandText = """
            SELECT Matricula, Nombre, Apellidos, PrimaryEmail, Resultado, Detalle, CreatedAt
            FROM dbo.GoogleWorkspaceProvisioningAuditLog
            WHERE JobId = @JobId
            ORDER BY CreatedAt
            """;
        cmd.Parameters.AddWithValue("@JobId", id);

        var rows = new List<AuditRow>();
        await using var reader = await cmd.ExecuteReaderAsync(ct);
        while (await reader.ReadAsync(ct))
        {
            rows.Add(new AuditRow(
                reader.GetString(0),
                reader.IsDBNull(1) ? "" : reader.GetString(1),
                reader.IsDBNull(2) ? "" : reader.GetString(2),
                reader.GetString(3),
                reader.GetString(4),
                reader.IsDBNull(5) ? "" : reader.GetString(5),
                reader.GetDateTime(6)
            ));
        }

        var bytes = BuildAuditExcel(rows);
        string fname = $"Pandora_Alta_Masiva_GoogleWorkspace_{DateTime.Now:yyyy-MM-dd}.xlsx";
        return File(bytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fname);
    }

    private record AuditRow(string Matricula, string Nombre, string Apellidos, string PrimaryEmail, string Resultado, string Detalle, DateTime CreatedAt);

    /// <summary>Mismo estilo visual que LicenciasController.BuildExcel — encabezado azul, filas coloreadas por resultado.</summary>
    private static byte[] BuildAuditExcel(List<AuditRow> rows)
    {
        using var wb = new XLWorkbook();

        var cAzul     = XLColor.FromHtml("#1A237E");
        var cMedio    = XLColor.FromHtml("#3949AB");
        var cVerde    = XLColor.FromHtml("#E8F5E9");
        var cVerdeTx  = XLColor.FromHtml("#2E7D32");
        var cGris     = XLColor.FromHtml("#F5F5F5");
        var cGrisTx   = XLColor.FromHtml("#757575");
        var cRojo     = XLColor.FromHtml("#FFEBEE");
        var cRojoTx   = XLColor.FromHtml("#C62828");

        var ws = wb.Worksheets.Add("Alta masiva");

        ws.Range("A1:F1").Merge();
        ws.Cell("A1").Value = "ALTA MASIVA DE CUENTAS — GOOGLE WORKSPACE";
        ws.Cell("A1").Style.Font.SetBold(true).Font.SetFontSize(14).Font.SetFontColor(XLColor.White)
            .Fill.SetBackgroundColor(cAzul)
            .Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center)
            .Alignment.SetVertical(XLAlignmentVerticalValues.Center);
        ws.Row(1).Height = 30;

        ws.Range("A2:F2").Merge();
        ws.Cell("A2").Value = $"Pandora  |  Generado: {DateTime.Now:dd/MM/yyyy HH:mm}";
        ws.Cell("A2").Style.Font.SetItalic(true).Font.SetFontColor(XLColor.White)
            .Fill.SetBackgroundColor(cMedio)
            .Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
        ws.Row(2).Height = 20;

        string[] hdrs = ["Matrícula", "Nombre", "Apellidos", "Correo", "Resultado", "Detalle"];
        for (int i = 0; i < hdrs.Length; i++)
        {
            var c = ws.Cell(3, i + 1);
            c.Value = hdrs[i];
            c.Style.Font.SetBold(true).Font.SetFontColor(XLColor.White)
                .Fill.SetBackgroundColor(cAzul)
                .Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center)
                .Border.SetOutsideBorder(XLBorderStyleValues.Thin)
                .Border.SetOutsideBorderColor(XLColor.White);
        }
        ws.Row(3).Height = 22;

        for (int i = 0; i < rows.Count; i++)
        {
            var row = rows[i];
            int er = i + 4;
            ws.Cell(er, 1).Value = row.Matricula;
            ws.Cell(er, 2).Value = row.Nombre;
            ws.Cell(er, 3).Value = row.Apellidos;
            ws.Cell(er, 4).Value = row.PrimaryEmail;
            ws.Cell(er, 5).Value = row.Resultado switch
            {
                "creado" => "Creado",
                "ya_existia" => "Ya existía",
                "error" => "Error",
                _ => row.Resultado,
            };
            ws.Cell(er, 6).Value = row.Detalle;

            var (bg, tx) = row.Resultado switch
            {
                "creado" => (cVerde, cVerdeTx),
                "ya_existia" => (cGris, cGrisTx),
                "error" => (cRojo, cRojoTx),
                _ => (XLColor.White, XLColor.Black),
            };
            ws.Range(er, 1, er, 6).Style.Fill.SetBackgroundColor(bg).Font.SetFontColor(tx);
        }

        ws.Columns().AdjustToContents();
        ws.SheetView.FreezeRows(3);

        using var stream = new MemoryStream();
        wb.SaveAs(stream);
        return stream.ToArray();
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
            INSERT INTO dbo.GoogleWorkspaceProvisioningAuditLog (JobId, Matricula, Nombre, Apellidos, PrimaryEmail, Resultado, Detalle, CreatedAt)
            VALUES (@JobId, @Matricula, @Nombre, @Apellidos, @PrimaryEmail, @Resultado, @Detalle, SYSUTCDATETIME());
            """;
        cmd.Parameters.AddWithValue("@JobId", S("jobId"));
        cmd.Parameters.AddWithValue("@Matricula", S("matricula"));
        cmd.Parameters.AddWithValue("@Nombre", string.IsNullOrWhiteSpace(S("nombre")) ? DBNull.Value : S("nombre"));
        cmd.Parameters.AddWithValue("@Apellidos", string.IsNullOrWhiteSpace(S("apellidos")) ? DBNull.Value : S("apellidos"));
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
