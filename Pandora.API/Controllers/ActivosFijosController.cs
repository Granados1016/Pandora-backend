using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using System.Security.Claims;

namespace Pandora.API.Controllers;

/// <summary>
/// Módulo de Activos Fijos — registro y depreciación de activos de la empresa.
/// </summary>
[ApiController]
[Route("api/activos-fijos")]
[Authorize]
public class ActivosFijosController(
    IConfiguration config,
    ILogger<ActivosFijosController> logger) : ControllerBase
{
    private SqlConnection Conn() => new(config.GetConnectionString("PandoraDb"));
    private bool IsAdmin => User.IsInRole("Admin");
    private string CurrentUser => User.FindFirstValue(ClaimTypes.Name) ?? "Desconocido";

    // Tables are created at startup in Program.cs; this flag skips the check after the first request
    private static volatile bool _tablesEnsured;

    private static async Task EnsureTablesAsync(SqlConnection conn, CancellationToken ct)
    {
        await using var cmd = conn.CreateCommand();
        cmd.CommandText = """
            IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'ActivosFijos')
            BEGIN
                CREATE TABLE dbo.ActivosFijos (
                    Id                  UNIQUEIDENTIFIER NOT NULL DEFAULT NEWID() PRIMARY KEY,
                    AssetNumber         NVARCHAR(50)     NOT NULL,
                    Name                NVARCHAR(200)    NOT NULL,
                    Description         NVARCHAR(MAX)    NULL,
                    Category            NVARCHAR(100)    NULL,
                    Brand               NVARCHAR(100)    NULL,
                    Model               NVARCHAR(100)    NULL,
                    SerialNumber        NVARCHAR(100)    NULL,
                    Department          NVARCHAR(100)    NULL,
                    ResponsibleUser     NVARCHAR(100)    NULL,
                    Location            NVARCHAR(200)    NULL,
                    Status              NVARCHAR(50)     NOT NULL DEFAULT 'Activo',
                    PurchaseDate        DATE             NULL,
                    PurchaseCost        DECIMAL(18,2)    NULL,
                    UsefulLifeYears     INT              NULL,
                    DepreciationMethod  NVARCHAR(50)     NULL DEFAULT 'Lineal',
                    ResidualValue       DECIMAL(18,2)    NULL DEFAULT 0,
                    AccumulatedDeprec   DECIMAL(18,2)    NOT NULL DEFAULT 0,
                    CurrentValue        AS (ISNULL(PurchaseCost,0) - AccumulatedDeprec),
                    ImageUrl            NVARCHAR(500)    NULL,
                    Notes               NVARCHAR(MAX)    NULL,
                    CreatedBy           NVARCHAR(100)    NOT NULL,
                    CreatedAt           DATETIME2        NOT NULL DEFAULT GETUTCDATE(),
                    UpdatedAt           DATETIME2        NULL,
                    IsDeleted           BIT              NOT NULL DEFAULT 0
                )
            END;
            IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'ActivoMovimientos')
            BEGIN
                CREATE TABLE dbo.ActivoMovimientos (
                    Id          UNIQUEIDENTIFIER NOT NULL DEFAULT NEWID() PRIMARY KEY,
                    ActivoId    UNIQUEIDENTIFIER NOT NULL,
                    Type        NVARCHAR(50)     NOT NULL,  -- Alta / Baja / Transferencia / Mantenimiento / Depreciacion
                    Description NVARCHAR(500)    NULL,
                    Amount      DECIMAL(18,2)    NULL,
                    FromDept    NVARCHAR(100)    NULL,
                    ToDept      NVARCHAR(100)    NULL,
                    CreatedBy   NVARCHAR(100)    NOT NULL,
                    CreatedAt   DATETIME2        NOT NULL DEFAULT GETUTCDATE()
                )
            END;
            """;
        await cmd.ExecuteNonQueryAsync(ct);
    }

    // ── GET /api/activos-fijos ────────────────────────────────────────────────
    [HttpGet]
    public async Task<IActionResult> GetAll(
        [FromQuery] string? status     = null,
        [FromQuery] string? category   = null,
        [FromQuery] string? department = null,
        [FromQuery] string? search     = null,
        CancellationToken ct = default)
    {
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            if (!_tablesEnsured) { await EnsureTablesAsync(conn, ct); _tablesEnsured = true; }

            var where = new List<string> { "IsDeleted = 0" };
            if (!string.IsNullOrWhiteSpace(status))     where.Add("Status = @Status");
            if (!string.IsNullOrWhiteSpace(category))   where.Add("Category = @Category");
            if (!string.IsNullOrWhiteSpace(department)) where.Add("Department = @Department");
            if (!string.IsNullOrWhiteSpace(search))     where.Add("(Name LIKE @Q OR AssetNumber LIKE @Q OR Brand LIKE @Q OR SerialNumber LIKE @Q)");

            await using var cmd = conn.CreateCommand();
            cmd.CommandText = $"""
                SELECT Id, AssetNumber, Name, Description, Category, Brand, Model, SerialNumber,
                       Department, ResponsibleUser, Location, Status, PurchaseDate,
                       PurchaseCost, UsefulLifeYears, DepreciationMethod, ResidualValue,
                       AccumulatedDeprec, CurrentValue, Notes, CreatedBy, CreatedAt
                FROM dbo.ActivosFijos
                WHERE {string.Join(" AND ", where)}
                ORDER BY Name
                """;
            if (!string.IsNullOrWhiteSpace(status))     cmd.Parameters.AddWithValue("@Status",     status);
            if (!string.IsNullOrWhiteSpace(category))   cmd.Parameters.AddWithValue("@Category",   category);
            if (!string.IsNullOrWhiteSpace(department)) cmd.Parameters.AddWithValue("@Department", department);
            if (!string.IsNullOrWhiteSpace(search))     cmd.Parameters.AddWithValue("@Q",          $"%{search}%");

            var list = new List<object>();
            await using var r = await cmd.ExecuteReaderAsync(ct);
            while (await r.ReadAsync(ct)) list.Add(MapActivo(r));
            return Ok(list);
        }
        catch (Exception ex) { logger.LogError(ex, "GetAll ActivosFijos"); return StatusCode(500, ex.Message); }
    }

    // ── GET /api/activos-fijos/{id} ───────────────────────────────────────────
    [HttpGet("{id:guid}")]
    public async Task<IActionResult> GetById(Guid id, CancellationToken ct)
    {
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = """
                SELECT Id, AssetNumber, Name, Description, Category, Brand, Model, SerialNumber,
                       Department, ResponsibleUser, Location, Status, PurchaseDate,
                       PurchaseCost, UsefulLifeYears, DepreciationMethod, ResidualValue,
                       AccumulatedDeprec, CurrentValue, Notes, CreatedBy, CreatedAt
                FROM dbo.ActivosFijos WHERE Id=@Id AND IsDeleted=0
                """;
            cmd.Parameters.AddWithValue("@Id", id);
            await using var r = await cmd.ExecuteReaderAsync(ct);
            if (!await r.ReadAsync(ct)) return NotFound();
            var activo = MapActivo(r);
            await r.CloseAsync();

            // Movimientos
            await using var mCmd = conn.CreateCommand();
            mCmd.CommandText = "SELECT Id, Type, Description, Amount, FromDept, ToDept, CreatedBy, CreatedAt FROM dbo.ActivoMovimientos WHERE ActivoId=@Id ORDER BY CreatedAt DESC";
            mCmd.Parameters.AddWithValue("@Id", id);
            await using var mr = await mCmd.ExecuteReaderAsync(ct);
            var movs = new List<object>();
            while (await mr.ReadAsync(ct))
                movs.Add(new {
                    id          = mr.GetGuid(0),
                    type        = mr.GetString(1),
                    description = mr.IsDBNull(2) ? null : mr.GetString(2),
                    amount      = mr.IsDBNull(3) ? (decimal?)null : mr.GetDecimal(3),
                    fromDept    = mr.IsDBNull(4) ? null : mr.GetString(4),
                    toDept      = mr.IsDBNull(5) ? null : mr.GetString(5),
                    createdBy   = mr.GetString(6),
                    createdAt   = mr.GetDateTime(7),
                });
            return Ok(new { activo, movimientos = movs });
        }
        catch (Exception ex) { logger.LogError(ex, "GetById Activo {Id}", id); return StatusCode(500, ex.Message); }
    }

    // ── POST /api/activos-fijos ───────────────────────────────────────────────
    [HttpPost]
    [Authorize(Roles = "Admin")]
    public async Task<IActionResult> Create([FromBody] ActivoFijoDto dto, CancellationToken ct)
    {
        if (string.IsNullOrWhiteSpace(dto.Name)) return BadRequest("Nombre requerido.");
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            if (!_tablesEnsured) { await EnsureTablesAsync(conn, ct); _tablesEnsured = true; }
            var newId = Guid.NewGuid();

            // Auto-generate AssetNumber if not provided
            string assetNumber = dto.AssetNumber?.Trim() ?? $"AF-{DateTime.UtcNow:yyyyMMdd}-{newId.ToString()[..4].ToUpper()}";

            await using var cmd = conn.CreateCommand();
            cmd.CommandText = """
                INSERT INTO dbo.ActivosFijos
                    (Id, AssetNumber, Name, Description, Category, Brand, Model, SerialNumber,
                     Department, ResponsibleUser, Location, Status, PurchaseDate, PurchaseCost,
                     UsefulLifeYears, DepreciationMethod, ResidualValue, Notes, CreatedBy, CreatedAt)
                VALUES
                    (@Id, @AN, @Name, @Desc, @Cat, @Brand, @Model, @SN,
                     @Dept, @Resp, @Loc, @Status, @PDate, @PCost,
                     @Life, @Method, @Resid, @Notes, @CreatedBy, GETUTCDATE())
                """;
            cmd.Parameters.AddWithValue("@Id",        newId);
            cmd.Parameters.AddWithValue("@AN",        assetNumber);
            cmd.Parameters.AddWithValue("@Name",      dto.Name.Trim());
            cmd.Parameters.AddWithValue("@Desc",      Nv(dto.Description));
            cmd.Parameters.AddWithValue("@Cat",       Nv(dto.Category));
            cmd.Parameters.AddWithValue("@Brand",     Nv(dto.Brand));
            cmd.Parameters.AddWithValue("@Model",     Nv(dto.Model));
            cmd.Parameters.AddWithValue("@SN",        Nv(dto.SerialNumber));
            cmd.Parameters.AddWithValue("@Dept",      Nv(dto.Department));
            cmd.Parameters.AddWithValue("@Resp",      Nv(dto.ResponsibleUser));
            cmd.Parameters.AddWithValue("@Loc",       Nv(dto.Location));
            cmd.Parameters.AddWithValue("@Status",    dto.Status ?? "Activo");
            cmd.Parameters.AddWithValue("@PDate",     dto.PurchaseDate.HasValue ? (object)dto.PurchaseDate.Value : DBNull.Value);
            cmd.Parameters.AddWithValue("@PCost",     dto.PurchaseCost.HasValue ? (object)dto.PurchaseCost.Value : DBNull.Value);
            cmd.Parameters.AddWithValue("@Life",      dto.UsefulLifeYears.HasValue ? (object)dto.UsefulLifeYears.Value : DBNull.Value);
            cmd.Parameters.AddWithValue("@Method",    dto.DepreciationMethod ?? "Lineal");
            cmd.Parameters.AddWithValue("@Resid",     dto.ResidualValue ?? 0m);
            cmd.Parameters.AddWithValue("@Notes",     Nv(dto.Notes));
            cmd.Parameters.AddWithValue("@CreatedBy", CurrentUser);
            await cmd.ExecuteNonQueryAsync(ct);

            // Registro de movimiento Alta
            await using var mCmd = conn.CreateCommand();
            mCmd.CommandText = """
                INSERT INTO dbo.ActivoMovimientos (Id, ActivoId, Type, Description, Amount, CreatedBy, CreatedAt)
                VALUES (NEWID(), @AId, 'Alta', @Desc, @Amount, @By, GETUTCDATE())
                """;
            mCmd.Parameters.AddWithValue("@AId",    newId);
            mCmd.Parameters.AddWithValue("@Desc",   $"Alta de activo: {dto.Name.Trim()}");
            mCmd.Parameters.AddWithValue("@Amount", dto.PurchaseCost.HasValue ? (object)dto.PurchaseCost.Value : DBNull.Value);
            mCmd.Parameters.AddWithValue("@By",     CurrentUser);
            await mCmd.ExecuteNonQueryAsync(ct);

            return CreatedAtAction(nameof(GetById), new { id = newId }, new { id = newId, assetNumber });
        }
        catch (Exception ex) { logger.LogError(ex, "Create Activo"); return StatusCode(500, ex.Message); }
    }

    // ── PUT /api/activos-fijos/{id} ───────────────────────────────────────────
    [HttpPut("{id:guid}")]
    [Authorize(Roles = "Admin")]
    public async Task<IActionResult> Update(Guid id, [FromBody] ActivoFijoDto dto, CancellationToken ct)
    {
        if (string.IsNullOrWhiteSpace(dto.Name)) return BadRequest("Nombre requerido.");
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = """
                UPDATE dbo.ActivosFijos SET
                    Name=@Name, Description=@Desc, Category=@Cat, Brand=@Brand, Model=@Model,
                    SerialNumber=@SN, Department=@Dept, ResponsibleUser=@Resp, Location=@Loc,
                    Status=@Status, PurchaseDate=@PDate, PurchaseCost=@PCost,
                    UsefulLifeYears=@Life, DepreciationMethod=@Method, ResidualValue=@Resid,
                    Notes=@Notes, UpdatedAt=GETUTCDATE()
                WHERE Id=@Id AND IsDeleted=0
                """;
            cmd.Parameters.AddWithValue("@Id",     id);
            cmd.Parameters.AddWithValue("@Name",   dto.Name.Trim());
            cmd.Parameters.AddWithValue("@Desc",   Nv(dto.Description));
            cmd.Parameters.AddWithValue("@Cat",    Nv(dto.Category));
            cmd.Parameters.AddWithValue("@Brand",  Nv(dto.Brand));
            cmd.Parameters.AddWithValue("@Model",  Nv(dto.Model));
            cmd.Parameters.AddWithValue("@SN",     Nv(dto.SerialNumber));
            cmd.Parameters.AddWithValue("@Dept",   Nv(dto.Department));
            cmd.Parameters.AddWithValue("@Resp",   Nv(dto.ResponsibleUser));
            cmd.Parameters.AddWithValue("@Loc",    Nv(dto.Location));
            cmd.Parameters.AddWithValue("@Status", dto.Status ?? "Activo");
            cmd.Parameters.AddWithValue("@PDate",  dto.PurchaseDate.HasValue ? (object)dto.PurchaseDate.Value : DBNull.Value);
            cmd.Parameters.AddWithValue("@PCost",  dto.PurchaseCost.HasValue ? (object)dto.PurchaseCost.Value : DBNull.Value);
            cmd.Parameters.AddWithValue("@Life",   dto.UsefulLifeYears.HasValue ? (object)dto.UsefulLifeYears.Value : DBNull.Value);
            cmd.Parameters.AddWithValue("@Method", dto.DepreciationMethod ?? "Lineal");
            cmd.Parameters.AddWithValue("@Resid",  dto.ResidualValue ?? 0m);
            cmd.Parameters.AddWithValue("@Notes",  Nv(dto.Notes));
            int rows = await cmd.ExecuteNonQueryAsync(ct);
            if (rows == 0) return NotFound();
            return NoContent();
        }
        catch (Exception ex) { logger.LogError(ex, "Update Activo {Id}", id); return StatusCode(500, ex.Message); }
    }

    // ── DELETE /api/activos-fijos/{id} ────────────────────────────────────────
    [HttpDelete("{id:guid}")]
    [Authorize(Roles = "Admin")]
    public async Task<IActionResult> Delete(Guid id, CancellationToken ct)
    {
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = "UPDATE dbo.ActivosFijos SET IsDeleted=1, UpdatedAt=GETUTCDATE() WHERE Id=@Id AND IsDeleted=0";
            cmd.Parameters.AddWithValue("@Id", id);
            int rows = await cmd.ExecuteNonQueryAsync(ct);
            if (rows == 0) return NotFound();
            return NoContent();
        }
        catch (Exception ex) { logger.LogError(ex, "Delete Activo {Id}", id); return StatusCode(500, ex.Message); }
    }

    // ── POST /api/activos-fijos/{id}/movimientos ──────────────────────────────
    [HttpPost("{id:guid}/movimientos")]
    [Authorize(Roles = "Admin")]
    public async Task<IActionResult> AddMovimiento(Guid id, [FromBody] MovimientoDto dto, CancellationToken ct)
    {
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = """
                INSERT INTO dbo.ActivoMovimientos (Id, ActivoId, Type, Description, Amount, FromDept, ToDept, CreatedBy, CreatedAt)
                VALUES (NEWID(), @AId, @Type, @Desc, @Amount, @From, @To, @By, GETUTCDATE())
                """;
            cmd.Parameters.AddWithValue("@AId",    id);
            cmd.Parameters.AddWithValue("@Type",   dto.Type ?? "Otro");
            cmd.Parameters.AddWithValue("@Desc",   Nv(dto.Description));
            cmd.Parameters.AddWithValue("@Amount", dto.Amount.HasValue ? (object)dto.Amount.Value : DBNull.Value);
            cmd.Parameters.AddWithValue("@From",   Nv(dto.FromDept));
            cmd.Parameters.AddWithValue("@To",     Nv(dto.ToDept));
            cmd.Parameters.AddWithValue("@By",     CurrentUser);

            // Always insert the movement record first
            await cmd.ExecuteNonQueryAsync(ct);

            // Then apply type-specific side-effect on the asset
            if (dto.Type == "Transferencia" && !string.IsNullOrWhiteSpace(dto.ToDept))
            {
                await using var upd = conn.CreateCommand();
                upd.CommandText = "UPDATE dbo.ActivosFijos SET Department=@Dept, UpdatedAt=GETUTCDATE() WHERE Id=@Id";
                upd.Parameters.AddWithValue("@Dept", dto.ToDept);
                upd.Parameters.AddWithValue("@Id",   id);
                await upd.ExecuteNonQueryAsync(ct);
            }
            else if (dto.Type == "Depreciacion" && dto.Amount.HasValue)
            {
                await using var upd = conn.CreateCommand();
                upd.CommandText = "UPDATE dbo.ActivosFijos SET AccumulatedDeprec=AccumulatedDeprec+@Amt, UpdatedAt=GETUTCDATE() WHERE Id=@Id";
                upd.Parameters.AddWithValue("@Amt", dto.Amount.Value);
                upd.Parameters.AddWithValue("@Id",  id);
                await upd.ExecuteNonQueryAsync(ct);
            }
            else if (dto.Type == "Baja")
            {
                await using var upd = conn.CreateCommand();
                upd.CommandText = "UPDATE dbo.ActivosFijos SET Status='Dado de Baja', UpdatedAt=GETUTCDATE() WHERE Id=@Id";
                upd.Parameters.AddWithValue("@Id", id);
                await upd.ExecuteNonQueryAsync(ct);
            }
            return Ok(new { added = true });
        }
        catch (Exception ex) { return StatusCode(500, ex.Message); }
    }

    // ── GET /api/activos-fijos/stats ──────────────────────────────────────────
    [HttpGet("stats")]
    [Authorize(Roles = "Admin")]
    public async Task<IActionResult> GetStats(CancellationToken ct)
    {
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            if (!_tablesEnsured) { await EnsureTablesAsync(conn, ct); _tablesEnsured = true; }
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = """
                SELECT
                    COUNT(*)                                          AS Total,
                    SUM(CASE WHEN Status='Activo' THEN 1 ELSE 0 END) AS Activos,
                    SUM(CASE WHEN Status='Dado de Baja' THEN 1 ELSE 0 END) AS DadosDeBaja,
                    SUM(CASE WHEN Status='En Mantenimiento' THEN 1 ELSE 0 END) AS EnMantenimiento,
                    ISNULL(SUM(PurchaseCost), 0)                    AS TotalCosto,
                    ISNULL(SUM(AccumulatedDeprec), 0)               AS TotalDepreciacion,
                    ISNULL(SUM(CurrentValue), 0)                    AS TotalValorActual
                FROM dbo.ActivosFijos WHERE IsDeleted=0
                """;
            await using var r = await cmd.ExecuteReaderAsync(ct);
            if (!await r.ReadAsync(ct)) return Ok(new {});
            return Ok(new {
                total             = r.GetInt32(0),
                activos           = r.GetInt32(1),
                dadosDeBaja       = r.GetInt32(2),
                enMantenimiento   = r.GetInt32(3),
                totalCosto        = r.GetDecimal(4),
                totalDepreciacion = r.GetDecimal(5),
                totalValorActual  = r.GetDecimal(6),
            });
        }
        catch (Exception ex) { return StatusCode(500, ex.Message); }
    }

    // ── GET /api/activos-fijos/export-excel ───────────────────────────────────
    [HttpGet("export-excel")]
    [Authorize(Roles = "Admin")]
    public async Task<IActionResult> ExportExcel(
        [FromQuery] string? status     = null,
        [FromQuery] string? category   = null,
        [FromQuery] string? department = null,
        CancellationToken ct = default)
    {
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await using var cmd = conn.CreateCommand();
            var where = new List<string> { "IsDeleted=0" };
            if (!string.IsNullOrWhiteSpace(status))     { where.Add("Status=@Status");         cmd.Parameters.AddWithValue("@Status",     status); }
            if (!string.IsNullOrWhiteSpace(category))   { where.Add("Category=@Category");     cmd.Parameters.AddWithValue("@Category",   category); }
            if (!string.IsNullOrWhiteSpace(department)) { where.Add("Department=@Department"); cmd.Parameters.AddWithValue("@Department", department); }
            cmd.CommandText = $"""
                SELECT AssetNumber, Name, Description, Category, Brand, Model, SerialNumber,
                       Department, ResponsibleUser, Location, Status,
                       PurchaseDate, PurchaseCost, UsefulLifeYears, DepreciationMethod,
                       ResidualValue, AccumulatedDeprec, CurrentValue, Notes, CreatedBy, CreatedAt
                FROM dbo.ActivosFijos
                WHERE {string.Join(" AND ", where)}
                ORDER BY CreatedAt DESC
                """;

            using var wb = new ClosedXML.Excel.XLWorkbook();
            var ws = wb.Worksheets.Add("Activos Fijos");

            string[] headers = [
                "No. Activo","Nombre","Descripción","Categoría","Marca","Modelo","No. Serie",
                "Departamento","Responsable","Ubicación","Estado",
                "Fecha Compra","Costo Compra","Vida Útil (años)","Método Depreciación",
                "Valor Residual","Depreciación Acumulada","Valor Actual","Notas","Creado Por","Fecha Alta"
            ];
            for (int i = 0; i < headers.Length; i++)
            {
                var cell = ws.Cell(1, i + 1);
                cell.Value = headers[i];
                cell.Style.Font.Bold = true;
                cell.Style.Fill.BackgroundColor = ClosedXML.Excel.XLColor.FromHtml("#0d2137");
                cell.Style.Font.FontColor       = ClosedXML.Excel.XLColor.White;
            }

            await using var r = await cmd.ExecuteReaderAsync(ct);
            int row = 2;
            while (await r.ReadAsync(ct))
            {
                static string S(SqlDataReader rd, int i) => rd.IsDBNull(i) ? "" : rd.GetValue(i).ToString()!;
                for (int col = 0; col < 21; col++)
                    ws.Cell(row, col + 1).Value = S(r, col);
                row++;
            }

            ws.Columns().AdjustToContents();
            using var ms = new MemoryStream();
            wb.SaveAs(ms);
            return File(ms.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                $"activos_fijos_{DateTime.Now:yyyyMMdd}.xlsx");
        }
        catch (Exception ex) { logger.LogError(ex, "ActivosFijos.ExportExcel"); return StatusCode(500, ex.Message); }
    }

    // ── Helpers ───────────────────────────────────────────────────────────────
    private static object Nv(string? s) => string.IsNullOrWhiteSpace(s) ? DBNull.Value : (object)s;

    private static object MapActivo(SqlDataReader r) => new
    {
        id                 = r.GetGuid(r.GetOrdinal("Id")),
        assetNumber        = r.GetString(r.GetOrdinal("AssetNumber")),
        name               = r.GetString(r.GetOrdinal("Name")),
        description        = r.IsDBNull(r.GetOrdinal("Description"))        ? null : r.GetString(r.GetOrdinal("Description")),
        category           = r.IsDBNull(r.GetOrdinal("Category"))           ? null : r.GetString(r.GetOrdinal("Category")),
        brand              = r.IsDBNull(r.GetOrdinal("Brand"))              ? null : r.GetString(r.GetOrdinal("Brand")),
        model              = r.IsDBNull(r.GetOrdinal("Model"))              ? null : r.GetString(r.GetOrdinal("Model")),
        serialNumber       = r.IsDBNull(r.GetOrdinal("SerialNumber"))       ? null : r.GetString(r.GetOrdinal("SerialNumber")),
        department         = r.IsDBNull(r.GetOrdinal("Department"))         ? null : r.GetString(r.GetOrdinal("Department")),
        responsibleUser    = r.IsDBNull(r.GetOrdinal("ResponsibleUser"))    ? null : r.GetString(r.GetOrdinal("ResponsibleUser")),
        location           = r.IsDBNull(r.GetOrdinal("Location"))           ? null : r.GetString(r.GetOrdinal("Location")),
        status             = r.GetString(r.GetOrdinal("Status")),
        purchaseDate       = r.IsDBNull(r.GetOrdinal("PurchaseDate"))       ? (DateTime?)null : r.GetDateTime(r.GetOrdinal("PurchaseDate")),
        purchaseCost       = r.IsDBNull(r.GetOrdinal("PurchaseCost"))       ? (decimal?)null  : r.GetDecimal(r.GetOrdinal("PurchaseCost")),
        usefulLifeYears    = r.IsDBNull(r.GetOrdinal("UsefulLifeYears"))    ? (int?)null      : r.GetInt32(r.GetOrdinal("UsefulLifeYears")),
        depreciationMethod = r.IsDBNull(r.GetOrdinal("DepreciationMethod")) ? null : r.GetString(r.GetOrdinal("DepreciationMethod")),
        residualValue      = r.IsDBNull(r.GetOrdinal("ResidualValue"))      ? (decimal?)null  : r.GetDecimal(r.GetOrdinal("ResidualValue")),
        accumulatedDeprec  = r.GetDecimal(r.GetOrdinal("AccumulatedDeprec")),
        currentValue       = r.IsDBNull(r.GetOrdinal("CurrentValue"))       ? (decimal?)null  : r.GetDecimal(r.GetOrdinal("CurrentValue")),
        notes              = r.IsDBNull(r.GetOrdinal("Notes"))              ? null : r.GetString(r.GetOrdinal("Notes")),
        createdBy          = r.GetString(r.GetOrdinal("CreatedBy")),
        createdAt          = r.GetDateTime(r.GetOrdinal("CreatedAt")),
    };
}

// ── DTOs ──────────────────────────────────────────────────────────────────────
public record ActivoFijoDto(
    string?   AssetNumber,
    string    Name,
    string?   Description,
    string?   Category,
    string?   Brand,
    string?   Model,
    string?   SerialNumber,
    string?   Department,
    string?   ResponsibleUser,
    string?   Location,
    string?   Status,
    DateTime? PurchaseDate,
    decimal?  PurchaseCost,
    int?      UsefulLifeYears,
    string?   DepreciationMethod,
    decimal?  ResidualValue,
    string?   Notes
);

public record MovimientoDto(
    string?  Type,
    string?  Description,
    decimal? Amount,
    string?  FromDept,
    string?  ToDept
);
