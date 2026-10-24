using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using System.IO.Compression;
using System.Net;
using System.Security.Claims;
using System.Text;
using System.Text.RegularExpressions;

namespace Pandora.API.Controllers;

/// <summary>
/// Actas de entrega / resguardos de equipo (entrega-recepción). Genera el acta en
/// Word (HTML con formato) para imprimir y firmar, y permite subir el documento
/// firmado. Respeta la visibilidad por área: Admin/SuperAdmin ven todas; un usuario
/// normal solo las de su área.
/// </summary>
[ApiController]
[Route("api/inventory/resguardos")]
[Authorize]
public class ResguardosController(IConfiguration config, ILogger<ResguardosController> logger,
    Pandora.API.Services.TenantService tenant) : ControllerBase
{
    private SqlConnection Conn() => new(config.GetConnectionString("PandoraDb"));
    private object TenantParam => (object?)tenant.TenantId ?? DBNull.Value;
    private string? CurrentUsername =>
        User.FindFirstValue(ClaimTypes.Name) ?? User.FindFirstValue("name");
    private bool IsAdmin => User.IsInRole("Admin") || User.IsInRole("SuperAdmin");

    private async Task<string?> GetUserAreaAsync(SqlConnection conn, CancellationToken ct)
    {
        var u = CurrentUsername;
        if (string.IsNullOrWhiteSpace(u)) return null;
        await using var cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT Area FROM dbo.AppUsers WHERE LOWER(Username) = LOWER(@U)";
        cmd.Parameters.AddWithValue("@U", u);
        return (await cmd.ExecuteScalarAsync(ct)) is string s && !string.IsNullOrWhiteSpace(s) ? s : null;
    }

    // ── POST /api/inventory/resguardos ───────────────────────────────────────
    [HttpPost]
    public async Task<IActionResult> Create([FromBody] CreateResguardoDto dto, CancellationToken ct)
    {
        if (string.IsNullOrWhiteSpace(dto.QuienRecibe)) return BadRequest("Indica quién recibe.");
        if (dto.Items is null || dto.Items.Length == 0) return BadRequest("Selecciona al menos un equipo.");

        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);

            // Folio ACT-YYYY-NNNN por tenant/año
            var year = (dto.Fecha ?? DateTime.Today).Year;
            string folio;
            await using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = """
                    SELECT ISNULL(MAX(CAST(RIGHT(Folio,4) AS INT)),0)+1
                    FROM dbo.Resguardos
                    WHERE Folio LIKE @Pfx AND (TenantId = @T OR (@T IS NULL AND TenantId IS NULL))
                    """;
                cmd.Parameters.AddWithValue("@Pfx", $"ACT-{year}-%");
                cmd.Parameters.AddWithValue("@T", TenantParam);
                var next = (int)(await cmd.ExecuteScalarAsync(ct))!;
                folio = $"ACT-{year}-{next:D4}";
            }

            var id = Guid.NewGuid();
            await using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = """
                    INSERT INTO dbo.Resguardos
                        (Id, Folio, Fecha, Area, QuienEntrega, QuienRecibe, JefeInmediato,
                         Observaciones, Estado, CreatedBy, CreatedAt, TenantId)
                    VALUES
                        (@Id, @Folio, @Fecha, @Area, @Entrega, @Recibe, @Jefe,
                         @Obs, 'Pendiente de firma', @By, GETUTCDATE(), @T)
                    """;
                cmd.Parameters.AddWithValue("@Id", id);
                cmd.Parameters.AddWithValue("@Folio", folio);
                cmd.Parameters.AddWithValue("@Fecha", (object?)(dto.Fecha ?? DateTime.Today) ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@Area", (object?)dto.Area?.Trim() ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@Entrega", (object?)dto.QuienEntrega?.Trim() ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@Recibe", dto.QuienRecibe.Trim());
                cmd.Parameters.AddWithValue("@Jefe", (object?)dto.JefeInmediato?.Trim() ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@Obs", (object?)dto.Observaciones?.Trim() ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@By", (object?)CurrentUsername ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@T", TenantParam);
                await cmd.ExecuteNonQueryAsync(ct);
            }

            // Snapshot de los equipos seleccionados (solo del tenant). Accesorios:
            // si el usuario los editó en el acta, se usa ese texto; si no, el del inventario.
            foreach (var it in dto.Items.GroupBy(x => x.ItemId).Select(g => g.First()))
            {
                object acc = string.IsNullOrWhiteSpace(it.Accesorios) ? DBNull.Value : it.Accesorios.Trim();

                await using var cmd = conn.CreateCommand();
                cmd.CommandText = """
                    INSERT INTO dbo.ResguardoItems (Id, ResguardoId, InventoryItemId, Equipo, MarcaModelo, Serie, Inventario, Estado, Accesorios)
                    SELECT NEWID(), @Rid, i.Id, i.Name,
                           LTRIM(RTRIM(ISNULL(i.Brand,'') + ' ' + ISNULL(i.Model,''))),
                           i.SerialNumber, i.InventoryNumber, i.Status, COALESCE(@Acc, i.Accessories)
                    FROM dbo.InventoryItems i
                    WHERE i.Id = @Iid AND i.TenantId = @T
                    """;
                cmd.Parameters.AddWithValue("@Rid", id);
                cmd.Parameters.AddWithValue("@Iid", it.ItemId);
                cmd.Parameters.AddWithValue("@Acc", acc);
                cmd.Parameters.AddWithValue("@T", TenantParam);
                await cmd.ExecuteNonQueryAsync(ct);

                // Asignar el equipo a quien recibe + al área del acta (entrega = asignación).
                // Si se editaron los accesorios, también se actualizan en el inventario.
                if (dto.AsignarEquipos)
                {
                    await using var up = conn.CreateCommand();
                    up.CommandText = """
                        UPDATE dbo.InventoryItems
                        SET AssignedTo = @Recibe,
                            AssignedEmployeeId = NULL,
                            Department = ISNULL(NULLIF(@Area, ''), Department),
                            Accessories = COALESCE(@Acc, Accessories),
                            UpdatedAt = GETUTCDATE()
                        WHERE Id = @Iid AND TenantId = @T
                        """;
                    up.Parameters.AddWithValue("@Recibe", dto.QuienRecibe.Trim());
                    up.Parameters.AddWithValue("@Area", (object?)dto.Area?.Trim() ?? "");
                    up.Parameters.AddWithValue("@Acc", acc);
                    up.Parameters.AddWithValue("@Iid", it.ItemId);
                    up.Parameters.AddWithValue("@T", TenantParam);
                    await up.ExecuteNonQueryAsync(ct);
                }
            }

            logger.LogInformation("Resguardo {Folio} creado por {By} (asignar={Asig})", folio, CurrentUsername, dto.AsignarEquipos);
            return Ok(new { id, folio });
        }
        catch (Exception ex) { logger.LogError(ex, "Resguardos.Create"); return StatusCode(500, ex.Message); }
    }

    // ── GET /api/inventory/resguardos ────────────────────────────────────────
    [HttpGet]
    public async Task<IActionResult> List(CancellationToken ct)
    {
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            string areaFilter = "";
            string? userArea = null;
            if (!IsAdmin)
            {
                userArea = await GetUserAreaAsync(conn, ct);
                areaFilter = userArea is null ? " AND 1=0" : " AND r.Area = @Area";
            }
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = $"""
                SELECT r.Id, r.Folio, r.Fecha, r.Area, r.QuienEntrega, r.QuienRecibe,
                       r.Estado, r.CreatedAt,
                       (SELECT COUNT(*) FROM dbo.ResguardoItems ri WHERE ri.ResguardoId = r.Id) AS NumEquipos,
                       CASE WHEN r.DocumentData IS NULL THEN 0 ELSE 1 END AS TieneDoc
                FROM dbo.Resguardos r
                WHERE (r.TenantId = @T OR (@T IS NULL AND r.TenantId IS NULL)) {areaFilter}
                ORDER BY r.CreatedAt DESC
                """;
            cmd.Parameters.AddWithValue("@T", TenantParam);
            if (userArea is not null) cmd.Parameters.AddWithValue("@Area", userArea);

            var list = new List<object>();
            await using var r = await cmd.ExecuteReaderAsync(ct);
            while (await r.ReadAsync(ct))
                list.Add(new
                {
                    id          = r.GetGuid(0),
                    folio       = r.GetString(1),
                    fecha       = r.GetDateTime(2),
                    area        = r.IsDBNull(3) ? null : r.GetString(3),
                    quienEntrega= r.IsDBNull(4) ? null : r.GetString(4),
                    quienRecibe = r.GetString(5),
                    estado      = r.GetString(6),
                    createdAt   = r.GetDateTime(7),
                    numEquipos  = r.GetInt32(8),
                    tieneDoc    = r.GetInt32(9) == 1,
                });
            return Ok(list);
        }
        catch (Exception ex) { logger.LogError(ex, "Resguardos.List"); return StatusCode(500, ex.Message); }
    }

    // ── GET /api/inventory/resguardos/{id} ───────────────────────────────────
    [HttpGet("{id:guid}")]
    public async Task<IActionResult> Detail(Guid id, CancellationToken ct)
    {
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            var r = await LoadResguardoAsync(conn, id, ct);
            if (r is null) return NotFound();
            if (!await CanAccessAsync(conn, r, ct)) return Forbid();

            var items = new List<object>();
            await using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = "SELECT Equipo, MarcaModelo, Serie, Inventario, Estado, Accesorios FROM dbo.ResguardoItems WHERE ResguardoId = @Id";
                cmd.Parameters.AddWithValue("@Id", id);
                await using var rr = await cmd.ExecuteReaderAsync(ct);
                while (await rr.ReadAsync(ct))
                    items.Add(new
                    {
                        equipo      = rr.IsDBNull(0) ? null : rr.GetString(0),
                        marcaModelo = rr.IsDBNull(1) ? null : rr.GetString(1),
                        serie       = rr.IsDBNull(2) ? null : rr.GetString(2),
                        inventario  = rr.IsDBNull(3) ? null : rr.GetString(3),
                        estado      = rr.IsDBNull(4) ? null : rr.GetString(4),
                        accesorios  = rr.IsDBNull(5) ? null : rr.GetString(5),
                    });
            }
            return Ok(new
            {
                r.Id, r.Folio, r.Fecha, r.Area, r.QuienEntrega, r.QuienRecibe,
                r.JefeInmediato, r.Observaciones, r.Estado, r.TieneDoc, r.DocumentName, items,
            });
        }
        catch (Exception ex) { logger.LogError(ex, "Resguardos.Detail {Id}", id); return StatusCode(500, ex.Message); }
    }

    // ── GET /api/inventory/resguardos/{id}/acta  (descarga Word) ──────────────
    [HttpGet("{id:guid}/acta")]
    public async Task<IActionResult> Acta(Guid id, CancellationToken ct)
    {
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            var r = await LoadResguardoAsync(conn, id, ct);
            if (r is null) return NotFound();
            if (!await CanAccessAsync(conn, r, ct)) return Forbid();

            // Branding del cliente (tenant), no "Pandora"
            string orgName = "Pandora";
            await using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = "SELECT TOP 1 ISNULL(DisplayName, Name) FROM dbo.Tenants WHERE Id = @T";
                cmd.Parameters.AddWithValue("@T", TenantParam);
                if (tenant.TenantId is not null && (await cmd.ExecuteScalarAsync(ct)) is string s && !string.IsNullOrWhiteSpace(s))
                    orgName = s;
            }

            var items = new List<(string eq, string mm, string se, string inv, string es, string acc)>();
            await using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = "SELECT Equipo, MarcaModelo, Serie, Inventario, Estado, Accesorios FROM dbo.ResguardoItems WHERE ResguardoId = @Id";
                cmd.Parameters.AddWithValue("@Id", id);
                await using var rr = await cmd.ExecuteReaderAsync(ct);
                while (await rr.ReadAsync(ct))
                    items.Add((
                        rr.IsDBNull(0) ? "" : rr.GetString(0),
                        rr.IsDBNull(1) ? "" : rr.GetString(1),
                        rr.IsDBNull(2) ? "" : rr.GetString(2),
                        rr.IsDBNull(3) ? "" : rr.GetString(3),
                        rr.IsDBNull(4) ? "" : rr.GetString(4),
                        rr.IsDBNull(5) ? "" : rr.GetString(5)));
            }

            // Si el cliente tiene plantilla Word (.docx) cargada, se rellena TAL CUAL su formato.
            byte[]? template = null;
            await using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = "SELECT Data FROM dbo.ResguardoPlantillas WHERE TenantId = @T";
                cmd.Parameters.AddWithValue("@T", tenant.TenantId ?? Guid.Empty);
                if (await cmd.ExecuteScalarAsync(ct) is byte[] b) template = b;
            }

            if (template is not null)
            {
                var docx = FillDocxTemplate(template, r, items);
                return File(docx, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", $"Acta_{r.Folio}.docx");
            }

            // Fallback sin plantilla: HTML servido como Word.
            var html = BuildActaHtml(orgName, r, items);
            var bytes = Encoding.UTF8.GetBytes(html);
            return File(bytes, "application/msword", $"Acta_{r.Folio}.doc");
        }
        catch (Exception ex) { logger.LogError(ex, "Resguardos.Acta {Id}", id); return StatusCode(500, ex.Message); }
    }

    // ── POST /api/inventory/resguardos/{id}/document  (subir firmado) ─────────
    [HttpPost("{id:guid}/document")]
    public async Task<IActionResult> UploadDocument(Guid id, IFormFile file, CancellationToken ct)
    {
        if (file is null || file.Length == 0) return BadRequest("Archivo requerido.");
        if (file.Length > 20 * 1024 * 1024) return BadRequest("El archivo supera 20 MB.");
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            var r = await LoadResguardoAsync(conn, id, ct);
            if (r is null) return NotFound();
            if (!await CanAccessAsync(conn, r, ct)) return Forbid();

            using var ms = new MemoryStream();
            await file.CopyToAsync(ms, ct);
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = """
                UPDATE dbo.Resguardos
                SET DocumentData = @Data, DocumentMime = @Mime, DocumentName = @Name,
                    Estado = 'Firmado'
                WHERE Id = @Id
                """;
            cmd.Parameters.AddWithValue("@Data", ms.ToArray());
            cmd.Parameters.AddWithValue("@Mime", (object?)file.ContentType ?? "application/octet-stream");
            cmd.Parameters.AddWithValue("@Name", (object?)Path.GetFileName(file.FileName) ?? "documento");
            cmd.Parameters.AddWithValue("@Id", id);
            await cmd.ExecuteNonQueryAsync(ct);
            return Ok(new { id, estado = "Firmado" });
        }
        catch (Exception ex) { logger.LogError(ex, "Resguardos.UploadDocument {Id}", id); return StatusCode(500, ex.Message); }
    }

    // ── GET /api/inventory/resguardos/{id}/document  (descargar firmado) ──────
    [HttpGet("{id:guid}/document")]
    public async Task<IActionResult> GetDocument(Guid id, CancellationToken ct)
    {
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            var r = await LoadResguardoAsync(conn, id, ct);
            if (r is null) return NotFound();
            if (!await CanAccessAsync(conn, r, ct)) return Forbid();

            await using var cmd = conn.CreateCommand();
            cmd.CommandText = "SELECT DocumentData, DocumentMime, DocumentName FROM dbo.Resguardos WHERE Id = @Id";
            cmd.Parameters.AddWithValue("@Id", id);
            await using var rr = await cmd.ExecuteReaderAsync(ct);
            if (!await rr.ReadAsync(ct) || rr.IsDBNull(0)) return NotFound("Sin documento.");
            var data = (byte[])rr[0];
            var mime = rr.IsDBNull(1) ? "application/octet-stream" : rr.GetString(1);
            var name = rr.IsDBNull(2) ? $"Acta_{r.Folio}" : rr.GetString(2);
            return File(data, mime, name);
        }
        catch (Exception ex) { logger.LogError(ex, "Resguardos.GetDocument {Id}", id); return StatusCode(500, ex.Message); }
    }

    // ── DELETE /api/inventory/resguardos/{id} ────────────────────────────────
    [HttpDelete("{id:guid}")]
    public async Task<IActionResult> Delete(Guid id, CancellationToken ct)
    {
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            var r = await LoadResguardoAsync(conn, id, ct);
            if (r is null) return NotFound();
            if (!await CanAccessAsync(conn, r, ct)) return Forbid();
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = "DELETE FROM dbo.Resguardos WHERE Id = @Id AND (TenantId = @T OR (@T IS NULL AND TenantId IS NULL))";
            cmd.Parameters.AddWithValue("@Id", id);
            cmd.Parameters.AddWithValue("@T", TenantParam);
            await cmd.ExecuteNonQueryAsync(ct);
            return Ok();
        }
        catch (Exception ex) { logger.LogError(ex, "Resguardos.Delete {Id}", id); return StatusCode(500, ex.Message); }
    }

    // ── GET /api/inventory/resguardos/personas ───────────────────────────────
    // Usuarios de Pandora (para autocompletar quién entrega / recibe / jefe).
    [HttpGet("personas")]
    public async Task<IActionResult> Personas(CancellationToken ct)
    {
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = """
                SELECT DISTINCT FullName, ISNULL(Email,''), ISNULL(Area,'')
                FROM dbo.AppUsers
                WHERE ISNULL(IsActive,0) = 1
                  AND FullName IS NOT NULL AND LTRIM(RTRIM(FullName)) <> ''
                  AND (TenantId = @T OR (@T IS NULL AND TenantId IS NULL))
                ORDER BY FullName
                """;
            cmd.Parameters.AddWithValue("@T", TenantParam);
            var list = new List<object>();
            await using var r = await cmd.ExecuteReaderAsync(ct);
            while (await r.ReadAsync(ct))
                list.Add(new { fullName = r.GetString(0), email = r.GetString(1), area = r.GetString(2) });
            return Ok(list);
        }
        catch (Exception ex) { logger.LogError(ex, "Resguardos.Personas"); return StatusCode(500, ex.Message); }
    }

    // ── POST /api/inventory/resguardos/plantilla  (subir plantilla Word del cliente) ──
    [HttpPost("plantilla")]
    [Authorize(Roles = "Admin,SuperAdmin")]
    public async Task<IActionResult> UploadPlantilla(IFormFile file, CancellationToken ct)
    {
        if (file is null || file.Length == 0) return BadRequest("Archivo requerido.");
        if (!file.FileName.EndsWith(".docx", StringComparison.OrdinalIgnoreCase))
            return BadRequest("La plantilla debe ser un archivo Word .docx con los tokens {{...}}.");
        if (file.Length > 10 * 1024 * 1024) return BadRequest("La plantilla supera 10 MB.");
        try
        {
            using var ms = new MemoryStream();
            await file.CopyToAsync(ms, ct);
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = """
                MERGE dbo.ResguardoPlantillas AS t
                USING (SELECT @T AS TenantId) AS s ON t.TenantId = s.TenantId
                WHEN MATCHED THEN UPDATE SET Data=@Data, FileName=@Name, UpdatedAt=GETUTCDATE()
                WHEN NOT MATCHED THEN INSERT (TenantId, Data, FileName) VALUES (@T, @Data, @Name);
                """;
            cmd.Parameters.AddWithValue("@T", tenant.TenantId ?? Guid.Empty);
            cmd.Parameters.AddWithValue("@Data", ms.ToArray());
            cmd.Parameters.AddWithValue("@Name", (object?)Path.GetFileName(file.FileName) ?? "plantilla.docx");
            await cmd.ExecuteNonQueryAsync(ct);
            return Ok(new { ok = true, fileName = file.FileName });
        }
        catch (Exception ex) { logger.LogError(ex, "Resguardos.UploadPlantilla"); return StatusCode(500, ex.Message); }
    }

    // ── GET /api/inventory/resguardos/plantilla/info ─────────────────────────
    [HttpGet("plantilla/info")]
    public async Task<IActionResult> PlantillaInfo(CancellationToken ct)
    {
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = "SELECT FileName, UpdatedAt FROM dbo.ResguardoPlantillas WHERE TenantId = @T";
            cmd.Parameters.AddWithValue("@T", tenant.TenantId ?? Guid.Empty);
            await using var r = await cmd.ExecuteReaderAsync(ct);
            if (!await r.ReadAsync(ct)) return Ok(new { has = false });
            return Ok(new { has = true, fileName = r.IsDBNull(0) ? null : r.GetString(0), updatedAt = r.GetDateTime(1) });
        }
        catch (Exception ex) { logger.LogError(ex, "Resguardos.PlantillaInfo"); return StatusCode(500, ex.Message); }
    }

    // ── GET /api/inventory/resguardos/plantilla  (descargar la plantilla actual) ──
    [HttpGet("plantilla")]
    [Authorize(Roles = "Admin,SuperAdmin")]
    public async Task<IActionResult> GetPlantilla(CancellationToken ct)
    {
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = "SELECT Data, FileName FROM dbo.ResguardoPlantillas WHERE TenantId = @T";
            cmd.Parameters.AddWithValue("@T", tenant.TenantId ?? Guid.Empty);
            await using var r = await cmd.ExecuteReaderAsync(ct);
            if (!await r.ReadAsync(ct)) return NotFound("Sin plantilla.");
            var data = (byte[])r[0];
            var name = r.IsDBNull(1) ? "plantilla_acta.docx" : r.GetString(1);
            return File(data, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", name);
        }
        catch (Exception ex) { logger.LogError(ex, "Resguardos.GetPlantilla"); return StatusCode(500, ex.Message); }
    }

    // ── Relleno de la plantilla Word (.docx) ──────────────────────────────────
    // La plantilla tiene tokens {{...}} en runs únicos; clonamos la fila de equipo
    // (con {{NO}}) y su fila de accesorios (con {{ACCESORIOS}}) por cada equipo.
    private static string Xe(string? s) => WebUtility.HtmlEncode(s ?? "");

    private static byte[] FillDocxTemplate(byte[] template,
        ResguardoRow r, List<(string eq, string mm, string se, string inv, string es, string acc)> items)
    {
        using var ms = new MemoryStream();
        ms.Write(template, 0, template.Length);
        ms.Position = 0;
        using (var zip = new ZipArchive(ms, ZipArchiveMode.Update, leaveOpen: true))
        {
            var entry = zip.GetEntry("word/document.xml") ?? throw new InvalidOperationException("document.xml no encontrado");
            string xml;
            using (var s = entry.Open())
            using (var reader = new StreamReader(s, Encoding.UTF8))
                xml = reader.ReadToEnd();

            xml = ExpandEquipmentRows(xml, items);
            xml = xml.Replace("{{FECHA}}", Xe(r.Fecha.ToString("dd/MM/yyyy")))
                     .Replace("{{AREA}}", Xe(r.Area))
                     .Replace("{{ENTREGA}}", Xe(r.QuienEntrega))
                     .Replace("{{RECIBE}}", Xe(r.QuienRecibe))
                     .Replace("{{JEFE}}", Xe(r.JefeInmediato))
                     .Replace("{{FOLIO}}", Xe(r.Folio))
                     .Replace("{{OBSERVACIONES}}", Xe(r.Observaciones));

            using (var s = entry.Open())
            {
                s.SetLength(0);
                using var w = new StreamWriter(s, new UTF8Encoding(false));
                w.Write(xml);
            }
        }
        return ms.ToArray();
    }

    private static string ExpandEquipmentRows(string xml,
        List<(string eq, string mm, string se, string inv, string es, string acc)> items)
    {
        // Fila de equipo (contiene {{NO}}) seguida de la fila de accesorios ({{ACCESORIOS}}).
        var rx = new Regex(
            @"(<w:tr\b(?:(?!</w:tr>).)*?\{\{NO\}\}(?:(?!</w:tr>).)*?</w:tr>)\s*(<w:tr\b(?:(?!</w:tr>).)*?\{\{ACCESORIOS\}\}(?:(?!</w:tr>).)*?</w:tr>)",
            RegexOptions.Singleline);
        var m = rx.Match(xml);
        if (!m.Success) return xml;
        string rowTpl = m.Groups[1].Value, accTpl = m.Groups[2].Value;
        var sb = new StringBuilder();
        int n = 1;
        foreach (var it in items)
        {
            sb.Append(rowTpl
                .Replace("{{NO}}", n.ToString())
                .Replace("{{EQUIPO}}", Xe(it.eq))
                .Replace("{{MARCA}}", Xe(it.mm))
                .Replace("{{SERIE}}", Xe(it.se))
                .Replace("{{INVENTARIO}}", Xe(it.inv))
                .Replace("{{ESTADO}}", Xe(it.es)));
            sb.Append(accTpl.Replace("{{ACCESORIOS}}", Xe(it.acc)));
            n++;
        }
        return xml[..m.Index] + sb + xml[(m.Index + m.Length)..];
    }

    // ── Helpers ───────────────────────────────────────────────────────────────
    private sealed record ResguardoRow(Guid Id, string Folio, DateTime Fecha, string? Area,
        string? QuienEntrega, string QuienRecibe, string? JefeInmediato, string? Observaciones,
        string Estado, bool TieneDoc, string? DocumentName);

    private async Task<ResguardoRow?> LoadResguardoAsync(SqlConnection conn, Guid id, CancellationToken ct)
    {
        await using var cmd = conn.CreateCommand();
        cmd.CommandText = """
            SELECT Id, Folio, Fecha, Area, QuienEntrega, QuienRecibe, JefeInmediato, Observaciones, Estado,
                   CASE WHEN DocumentData IS NULL THEN 0 ELSE 1 END, DocumentName
            FROM dbo.Resguardos
            WHERE Id = @Id AND (TenantId = @T OR (@T IS NULL AND TenantId IS NULL))
            """;
        cmd.Parameters.AddWithValue("@Id", id);
        cmd.Parameters.AddWithValue("@T", TenantParam);
        await using var r = await cmd.ExecuteReaderAsync(ct);
        if (!await r.ReadAsync(ct)) return null;
        return new ResguardoRow(
            r.GetGuid(0), r.GetString(1), r.GetDateTime(2),
            r.IsDBNull(3) ? null : r.GetString(3),
            r.IsDBNull(4) ? null : r.GetString(4),
            r.GetString(5),
            r.IsDBNull(6) ? null : r.GetString(6),
            r.IsDBNull(7) ? null : r.GetString(7),
            r.GetString(8),
            r.GetInt32(9) == 1,
            r.IsDBNull(10) ? null : r.GetString(10));
    }

    // Un usuario normal solo accede a resguardos de su área.
    private async Task<bool> CanAccessAsync(SqlConnection conn, ResguardoRow r, CancellationToken ct)
    {
        if (IsAdmin) return true;
        var area = await GetUserAreaAsync(conn, ct);
        return area is not null && string.Equals(area, r.Area, StringComparison.OrdinalIgnoreCase);
    }

    private static string H(string? s) => WebUtility.HtmlEncode(s ?? "");

    private static string BuildActaHtml(string org, ResguardoRow r,
        List<(string eq, string mm, string se, string inv, string es, string acc)> items)
    {
        var sb = new StringBuilder();
        var rows = new StringBuilder();
        int n = 1;
        foreach (var it in items)
        {
            rows.Append($"""
                <tr>
                  <td style="text-align:center">{n}</td>
                  <td>{H(it.eq)}</td><td>{H(it.mm)}</td><td>{H(it.se)}</td>
                  <td>{H(it.inv)}</td><td style="text-align:center">{H(it.es)}</td>
                </tr>
                <tr><td></td><td colspan="5" style="font-size:9pt;color:#444"><b>Accesorios:</b> {H(it.acc)}</td></tr>
                """);
            n++;
        }

        sb.Append($$"""
        <html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word'>
        <head><meta charset="utf-8"><title>Acta {{H(r.Folio)}}</title>
        <style>
          body { font-family: Calibri, Arial, sans-serif; font-size: 11pt; color:#222; }
          h1 { font-size: 15pt; text-align:center; margin: 6pt 0; }
          .org { text-align:center; font-weight:bold; font-size:12pt; }
          .sub { text-align:center; color:#555; font-size:10pt; margin-bottom:8pt; }
          table { border-collapse: collapse; width: 100%; margin: 6pt 0; }
          td, th { border: 1px solid #999; padding: 4pt 6pt; font-size:10pt; vertical-align: top; }
          th { background:#1a237e; color:#fff; }
          .meta td { border:1px solid #bbb; }
          .lbl { background:#eef; font-weight:bold; width:22%; }
          .sec { font-weight:bold; background:#1a237e; color:#fff; padding:4pt 6pt; margin-top:10pt; }
          .firmas td { height:60pt; vertical-align:bottom; text-align:center; }
          ol { font-size:10pt; }
        </style></head>
        <body>
          <div class="org">{{H(org)}}</div>
          <div class="sub">Dirección de Tecnología</div>
          <h1>ACTA DE ENTREGA DE EQUIPO</h1>

          <table class="meta">
            <tr><td class="lbl">FECHA DE ENTREGA</td><td>{{r.Fecha:dd/MM/yyyy}}</td>
                <td class="lbl">ÁREA / DEPARTAMENTO</td><td>{{H(r.Area)}}</td></tr>
            <tr><td class="lbl">QUIEN ENTREGA</td><td>{{H(r.QuienEntrega)}}</td>
                <td class="lbl">QUIEN RECIBE</td><td>{{H(r.QuienRecibe)}}</td></tr>
            <tr><td class="lbl">JEFE INMEDIATO</td><td>{{H(r.JefeInmediato)}}</td>
                <td class="lbl">No. DE ACTA</td><td>{{H(r.Folio)}}</td></tr>
          </table>

          <div class="sec">1. DETALLE DEL EQUIPO ENTREGADO</div>
          <table>
            <tr><th style="width:5%">No.</th><th>EQUIPO</th><th>MARCA / MODELO</th>
                <th>No. DE SERIE / IMEI</th><th>No. DE INVENTARIO</th><th>ESTADO</th></tr>
            {{rows}}
          </table>

          <div class="sec">2. CONDICIONES DE USO Y RESPONSABILIDAD</div>
          <ol>
            <li>El equipo entregado debe utilizarse exclusivamente para fines laborales de la institución.</li>
            <li>El personal receptor es responsable del cuidado y del buen funcionamiento del equipo mientras esté en su posesión, y se compromete a reportar de inmediato cualquier daño o mal funcionamiento.</li>
            <li>En caso de pérdida, daño o mal uso del equipo, el responsable se compromete a asumir las consecuencias o los costos derivados, según determine la institución.</li>
            <li>El equipo debe devolverse en las mismas condiciones al finalizar el periodo de uso o cuando termine el vínculo laboral con la institución.</li>
          </ol>

          <div class="sec">3. OBSERVACIONES</div>
          <table><tr><td style="height:40pt;vertical-align:top">{{H(r.Observaciones)}}</td></tr></table>

          <div class="sec">CONSTANCIA DE ENTREGA Y FIRMAS</div>
          <table class="firmas">
            <tr><td>{{H(r.QuienEntrega)}}<br>_______________________<br>QUIEN ENTREGA</td>
                <td>{{H(r.QuienRecibe)}}<br>_______________________<br>QUIEN RECIBE</td>
                <td>{{H(r.JefeInmediato)}}<br>_______________________<br>Vo. Bo. JEFE INMEDIATO</td></tr>
          </table>
        </body></html>
        """);
        return sb.ToString();
    }
}

public record CreateResguardoDto(
    DateTime? Fecha, string? Area, string? QuienEntrega, string QuienRecibe,
    string? JefeInmediato, string? Observaciones, ResguardoItemInput[] Items,
    bool AsignarEquipos = true);

public record ResguardoItemInput(Guid ItemId, string? Accesorios);
