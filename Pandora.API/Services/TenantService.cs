using System.Security.Claims;
using Microsoft.Data.SqlClient;

namespace Pandora.API.Services;

/// <summary>
/// Resuelve el tenant activo del request desde el claim "tenantId" del JWT.
/// Se registra como Scoped en DI.
/// </summary>
public class TenantService(IHttpContextAccessor httpContextAccessor, IConfiguration config)
{
    private SqlConnection Conn() => new(config.GetConnectionString("PandoraDb"));

    // ── Tenant del usuario autenticado ────────────────────────────────────────
    public Guid? TenantId
    {
        get
        {
            var claim = httpContextAccessor.HttpContext?.User.FindFirstValue("tenantId");
            return Guid.TryParse(claim, out var id) ? id : null;
        }
    }

    public bool HasTenant => TenantId.HasValue;

    // ── Cargar datos completos del tenant ─────────────────────────────────────
    public async Task<TenantInfo?> LoadAsync(Guid tenantId, CancellationToken ct = default)
    {
        await using var conn = Conn();
        await conn.OpenAsync(ct);
        await using var cmd = conn.CreateCommand();
        cmd.CommandText = """
            SELECT Id, Slug, Name, DisplayName, LogoData, LogoMime,
                   PrimaryColor, SecondaryColor, LicensedModules, MaxUsers,
                   ExpiresAt, IsActive, ContactEmail, Notes
            FROM dbo.Tenants WHERE Id = @Id
            """;
        cmd.Parameters.AddWithValue("@Id", tenantId);
        await using var r = await cmd.ExecuteReaderAsync(ct);
        if (!await r.ReadAsync(ct)) return null;

        return new TenantInfo
        {
            Id             = r.GetGuid(r.GetOrdinal("Id")),
            Slug           = r.GetString(r.GetOrdinal("Slug")),
            Name           = r.GetString(r.GetOrdinal("Name")),
            DisplayName    = r.GetString(r.GetOrdinal("DisplayName")),
            LogoData       = r.IsDBNull(r.GetOrdinal("LogoData"))  ? null : (byte[])r["LogoData"],
            LogoMime       = r.IsDBNull(r.GetOrdinal("LogoMime"))  ? null : r.GetString(r.GetOrdinal("LogoMime")),
            PrimaryColor   = r.GetString(r.GetOrdinal("PrimaryColor")),
            SecondaryColor = r.GetString(r.GetOrdinal("SecondaryColor")),
            LicensedModules = r.GetInt64(r.GetOrdinal("LicensedModules")),
            MaxUsers       = r.GetInt32(r.GetOrdinal("MaxUsers")),
            ExpiresAt      = r.IsDBNull(r.GetOrdinal("ExpiresAt")) ? null : r.GetDateTime(r.GetOrdinal("ExpiresAt")),
            IsActive       = r.GetBoolean(r.GetOrdinal("IsActive")),
            ContactEmail   = r.IsDBNull(r.GetOrdinal("ContactEmail")) ? null : r.GetString(r.GetOrdinal("ContactEmail")),
            Notes          = r.IsDBNull(r.GetOrdinal("Notes"))        ? null : r.GetString(r.GetOrdinal("Notes")),
        };
    }

    // ── Verificar si el tenant está activo y no ha expirado ───────────────────
    public async Task<(bool ok, string? reason)> ValidateAsync(Guid tenantId, CancellationToken ct = default)
    {
        await using var conn = Conn();
        await conn.OpenAsync(ct);
        await using var cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT IsActive, ExpiresAt FROM dbo.Tenants WHERE Id = @Id";
        cmd.Parameters.AddWithValue("@Id", tenantId);
        await using var r = await cmd.ExecuteReaderAsync(ct);
        if (!await r.ReadAsync(ct)) return (false, "Tenant no encontrado.");
        if (!r.GetBoolean(0))       return (false, "Acceso suspendido. Contacte al proveedor.");
        var exp = r.IsDBNull(1) ? (DateTime?)null : r.GetDateTime(1);
        if (exp.HasValue && exp.Value < DateTime.UtcNow)
            return (false, "La licencia de su organización ha expirado. Contacte al proveedor.");
        return (true, null);
    }
}

public class TenantInfo
{
    public Guid     Id              { get; init; }
    public string   Slug            { get; init; } = "";
    public string   Name            { get; init; } = "";
    public string   DisplayName     { get; init; } = "";
    public byte[]?  LogoData        { get; init; }
    public string?  LogoMime        { get; init; }
    public string   PrimaryColor    { get; init; } = "#1A237E";
    public string   SecondaryColor  { get; init; } = "#283593";
    public long     LicensedModules { get; init; } = -1;
    public int      MaxUsers        { get; init; } = 50;
    public DateTime? ExpiresAt      { get; init; }
    public bool     IsActive        { get; init; }
    public string?  ContactEmail    { get; init; }
    public string?  Notes           { get; init; }

    public bool IsExpired => ExpiresAt.HasValue && ExpiresAt.Value < DateTime.UtcNow;

    public string? LogoBase64 => LogoData != null && LogoMime != null
        ? $"data:{LogoMime};base64,{Convert.ToBase64String(LogoData)}"
        : null;
}
