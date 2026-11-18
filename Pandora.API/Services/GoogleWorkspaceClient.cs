using System.IdentityModel.Tokens.Jwt;
using System.Net.Http.Json;
using System.Security.Claims;
using System.Text;
using System.Text.Json;
using Microsoft.IdentityModel.Tokens;

namespace Pandora.API.Services;

/// <summary>
/// Cliente server-to-server hacia el microservicio Node
/// "google-workspace-provisioning". Emite el JWT de servicio (mismo esquema que
/// <c>GoogleWorkspaceProvisioningController</c>: audience
/// "google-workspace-provisioning-svc", HS256 con la ServiceKey compartida) y
/// llama a sus endpoints. Hoy se usa para la baja de colaborador (#8 Fase 3):
/// suspender la cuenta de Google de quien deja la institución.
/// </summary>
public class GoogleWorkspaceClient(
    IConfiguration config,
    IHttpClientFactory httpClientFactory,
    ILogger<GoogleWorkspaceClient> logger)
{
    private const string AudienceForNode = "google-workspace-provisioning-svc";
    private const int ServiceTokenMinutes = 10;

    private string ServiceKey => config["GoogleWorkspaceProvisioning:ServiceKey"] ?? "";
    private string? ServiceUrl => config["GoogleWorkspaceProvisioning:ServiceUrl"];

    public bool IsConfigured => !string.IsNullOrWhiteSpace(ServiceKey) && !string.IsNullOrWhiteSpace(ServiceUrl);

    /// <summary>Resultado de suspender una cuenta de Google.</summary>
    /// <param name="Status">"suspendida" | "no_existia" | "no_configurado" | "error".</param>
    public record SuspendResult(string Status, string? Detail = null)
    {
        public bool Ok => Status is "suspendida" or "no_existia";
    }

    private string BuildServiceToken(string requestedBy)
    {
        var key = new SymmetricSecurityKey(Encoding.UTF8.GetBytes(ServiceKey));
        var creds = new SigningCredentials(key, SecurityAlgorithms.HmacSha256);
        var jwt = new JwtSecurityToken(
            audience: AudienceForNode,
            claims: [new Claim("requestedBy", requestedBy)],
            expires: DateTime.UtcNow.AddMinutes(ServiceTokenMinutes),
            signingCredentials: creds);
        return new JwtSecurityTokenHandler().WriteToken(jwt);
    }

    /// <summary>
    /// Suspende la cuenta de Google del correo dado. Nunca lanza: devuelve un
    /// <see cref="SuspendResult"/> con el estado, para que la baja del usuario en
    /// Pandora no dependa de que el microservicio esté disponible.
    /// </summary>
    public async Task<SuspendResult> SuspendGoogleAccountAsync(string email, string requestedBy, CancellationToken ct)
    {
        if (!IsConfigured)
        {
            logger.LogWarning("GoogleWorkspaceClient: ServiceKey/ServiceUrl no configurados; no se suspende {Email}.", email);
            return new SuspendResult("no_configurado", "Integración con Google no configurada.");
        }

        try
        {
            using var client = httpClientFactory.CreateClient();
            client.BaseAddress = new Uri(ServiceUrl!);
            client.Timeout = TimeSpan.FromSeconds(40); // el micro reintenta ante 403/5xx
            client.DefaultRequestHeaders.Authorization =
                new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", BuildServiceToken(requestedBy));

            using var response = await client.PostAsJsonAsync(
                "/api/v1/provisioning/users/suspend", new { email }, ct);
            var bodyText = await response.Content.ReadAsStringAsync(ct);

            if (!response.IsSuccessStatusCode)
            {
                logger.LogWarning("GoogleWorkspaceClient: suspender {Email} respondió {Status}.", email, response.StatusCode);
                return new SuspendResult("error", $"El servicio respondió {(int)response.StatusCode}.");
            }

            var json = JsonSerializer.Deserialize<JsonElement>(bodyText);
            bool notFound = json.TryGetProperty("notFound", out var nf) && nf.ValueKind == JsonValueKind.True;
            return new SuspendResult(notFound ? "no_existia" : "suspendida");
        }
        catch (Exception ex)
        {
            // Sanitizado: no se expone la URL interna del micro.
            logger.LogError(ex, "GoogleWorkspaceClient: no se pudo contactar al microservicio para suspender {Email}.", email);
            return new SuspendResult("error", "El servicio de Google no está disponible. Intenta de nuevo en unos minutos.");
        }
    }
}
