using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using System.Security.Claims;
using System.Text;

namespace Pandora.API.Controllers;

[ApiController]
[Route("api/inventory")]
[Authorize]
public class InventoryController(IConfiguration config, ILogger<InventoryController> logger) : ControllerBase
{
    private SqlConnection Conn() => new(config.GetConnectionString("PandoraDb"));

    private string? CurrentUsername =>
        User.FindFirstValue(ClaimTypes.Name) ??
        User.FindFirstValue("name") ??
        User.Claims.FirstOrDefault(c => c.Type.EndsWith("name", StringComparison.OrdinalIgnoreCase))?.Value;

    // ════════════════════════════════════════════════════════════════════════
    //  TYPES
    // ════════════════════════════════════════════════════════════════════════

    [HttpGet("types")]
    public async Task<IActionResult> GetTypes(CancellationToken ct)
    {
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = """
                SELECT t.Id, t.Name, t.Description, t.Department, t.IsActive, t.CreatedAt,
                       COUNT(i.Id) AS ItemCount
                FROM dbo.InventoryTypes t
                LEFT JOIN dbo.InventoryItems i ON i.InventoryTypeId = t.Id
                GROUP BY t.Id, t.Name, t.Description, t.Department, t.IsActive, t.CreatedAt
                ORDER BY t.Name
                """;
            var list = new List<object>();
            await using var r = await cmd.ExecuteReaderAsync(ct);
            while (await r.ReadAsync(ct))
                list.Add(new
                {
                    id          = r.GetGuid(r.GetOrdinal("Id")),
                    name        = r.GetString(r.GetOrdinal("Name")),
                    description = r.IsDBNull(r.GetOrdinal("Description")) ? null : r.GetString(r.GetOrdinal("Description")),
                    department  = r.IsDBNull(r.GetOrdinal("Department"))  ? null : r.GetString(r.GetOrdinal("Department")),
                    isActive    = r.GetBoolean(r.GetOrdinal("IsActive")),
                    createdAt   = r.GetDateTime(r.GetOrdinal("CreatedAt")),
                    itemCount   = r.GetInt32(r.GetOrdinal("ItemCount")),
                });
            return Ok(list);
        }
        catch (Exception ex) { logger.LogError(ex, "GetTypes"); return StatusCode(500, ex.Message); }
    }

    [HttpGet("types/{id:guid}")]
    public async Task<IActionResult> GetTypeById(Guid id, CancellationToken ct)
    {
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = """
                SELECT Id, Name, Description, Department, IsActive, CreatedAt
                FROM dbo.InventoryTypes WHERE Id = @Id
                """;
            cmd.Parameters.AddWithValue("@Id", id);
            await using var r = await cmd.ExecuteReaderAsync(ct);
            if (!await r.ReadAsync(ct)) return NotFound("Tipo no encontrado.");
            return Ok(new
            {
                id          = r.GetGuid(r.GetOrdinal("Id")),
                name        = r.GetString(r.GetOrdinal("Name")),
                description = r.IsDBNull(r.GetOrdinal("Description")) ? null : r.GetString(r.GetOrdinal("Description")),
                department  = r.IsDBNull(r.GetOrdinal("Department"))  ? null : r.GetString(r.GetOrdinal("Department")),
                isActive    = r.GetBoolean(r.GetOrdinal("IsActive")),
                createdAt   = r.GetDateTime(r.GetOrdinal("CreatedAt")),
            });
        }
        catch (Exception ex) { logger.LogError(ex, "GetTypeById {Id}", id); return StatusCode(500, ex.Message); }
    }

    [HttpPost("types")]
    public async Task<IActionResult> CreateType([FromBody] InventoryTypeDto dto, CancellationToken ct)
    {
        if (string.IsNullOrWhiteSpace(dto.Name)) return BadRequest("Nombre requerido.");
        var id = Guid.NewGuid();
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = """
                INSERT INTO dbo.InventoryTypes (Id, Name, Description, Department, IsActive, CreatedAt)
                VALUES (@Id, @Name, @Desc, @Dept, @IsActive, GETUTCDATE())
                """;
            cmd.Parameters.AddWithValue("@Id",       id);
            cmd.Parameters.AddWithValue("@Name",     dto.Name.Trim());
            cmd.Parameters.AddWithValue("@Desc",     (object?)dto.Description ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@Dept",     (object?)dto.Department  ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@IsActive", dto.IsActive);
            await cmd.ExecuteNonQueryAsync(ct);
            return Ok(new { id });
        }
        catch (Exception ex) { logger.LogError(ex, "CreateType"); return StatusCode(500, ex.Message); }
    }

    [HttpPut("types/{id:guid}")]
    public async Task<IActionResult> UpdateType(Guid id, [FromBody] InventoryTypeDto dto, CancellationToken ct)
    {
        if (string.IsNullOrWhiteSpace(dto.Name)) return BadRequest("Nombre requerido.");
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = """
                UPDATE dbo.InventoryTypes
                SET Name = @Name, Description = @Desc, Department = @Dept,
                    IsActive = @IsActive, UpdatedAt = GETUTCDATE()
                WHERE Id = @Id
                """;
            cmd.Parameters.AddWithValue("@Id",       id);
            cmd.Parameters.AddWithValue("@Name",     dto.Name.Trim());
            cmd.Parameters.AddWithValue("@Desc",     (object?)dto.Description ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@Dept",     (object?)dto.Department  ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@IsActive", dto.IsActive);
            int rows = await cmd.ExecuteNonQueryAsync(ct);
            if (rows == 0) return NotFound("Tipo no encontrado.");
            return Ok(new { id });
        }
        catch (Exception ex) { logger.LogError(ex, "UpdateType {Id}", id); return StatusCode(500, ex.Message); }
    }

    [HttpDelete("types/{id:guid}")]
    public async Task<IActionResult> DeleteType(Guid id, CancellationToken ct)
    {
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = "DELETE FROM dbo.InventoryTypes WHERE Id = @Id";
            cmd.Parameters.AddWithValue("@Id", id);
            int rows = await cmd.ExecuteNonQueryAsync(ct);
            if (rows == 0) return NotFound("Tipo no encontrado.");
            return NoContent();
        }
        catch (Exception ex) { logger.LogError(ex, "DeleteType {Id}", id); return StatusCode(500, ex.Message); }
    }

    // ════════════════════════════════════════════════════════════════════════
    //  ITEMS
    // ════════════════════════════════════════════════════════════════════════

    [HttpGet("items/dashboard")]
    public async Task<IActionResult> GetDashboard(CancellationToken ct)
    {
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = """
                SELECT
                    COUNT(*) AS Total,
                    ISNULL(SUM(CASE WHEN TRY_CAST(Status AS NVARCHAR(50)) IN ('1','Activo')            THEN 1 ELSE 0 END), 0) AS Activos,
                    ISNULL(SUM(CASE WHEN TRY_CAST(Status AS NVARCHAR(50)) IN ('2','Mantenimiento','En mantenimiento') THEN 1 ELSE 0 END), 0) AS EnMantenimiento,
                    ISNULL(SUM(CASE WHEN TRY_CAST(Status AS NVARCHAR(50)) IN ('3','Dado de baja','DadoDeBaja')        THEN 1 ELSE 0 END), 0) AS DadosDeBaja
                FROM dbo.InventoryItems
                """;
            await using var r = await cmd.ExecuteReaderAsync(ct);
            await r.ReadAsync(ct);
            return Ok(new
            {
                total           = r.IsDBNull(0) ? 0 : r.GetInt32(0),
                activos         = r.IsDBNull(1) ? 0 : r.GetInt32(1),
                enMantenimiento = r.IsDBNull(2) ? 0 : r.GetInt32(2),
                dadosDeBaja     = r.IsDBNull(3) ? 0 : r.GetInt32(3),
            });
        }
        catch (Exception ex) { logger.LogError(ex, "GetDashboard Inventory"); return StatusCode(500, ex.Message); }
    }

    [HttpGet("items/next-number")]
    public async Task<IActionResult> GetNextNumber([FromQuery] Guid? departmentId, CancellationToken ct)
    {
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            // Obtener prefijo del departamento
            string prefix = "EQT";
            if (departmentId.HasValue)
            {
                await using var cmdP = conn.CreateCommand();
                cmdP.CommandText = "SELECT InventoryPrefix FROM dbo.Departments WHERE Id = @Id";
                cmdP.Parameters.AddWithValue("@Id", departmentId.Value);
                var p = await cmdP.ExecuteScalarAsync(ct);
                if (p is string s && !string.IsNullOrWhiteSpace(s)) prefix = s;
            }
            // Contar items con ese prefijo
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = "SELECT COUNT(*) FROM dbo.InventoryItems WHERE InventoryNumber LIKE @Prefix + '%'";
            cmd.Parameters.AddWithValue("@Prefix", prefix);
            var count = (int)await cmd.ExecuteScalarAsync(ct)!;
            var next  = $"{prefix}-{(count + 1):D4}";
            return Ok(new { nextNumber = next });
        }
        catch (Exception ex) { logger.LogError(ex, "GetNextNumber"); return StatusCode(500, ex.Message); }
    }

    [HttpGet("items")]
    public async Task<IActionResult> GetItems([FromQuery] Guid? typeId, CancellationToken ct)
    {
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await using var cmd = conn.CreateCommand();
            var sql = """
                SELECT i.Id, i.InventoryNumber, i.Name, i.Brand, i.Model, i.SerialNumber,
                       i.Status, i.Department, i.AssignedTo, i.AssignedEmployeeId,
                       i.InventoryTypeId, i.IsPhone,
                       i.PurchaseDate, i.PurchasePrice, i.Accessories, i.CreatedAt,
                       t.Name AS TypeName,
                       e.FullName AS EmployeeName
                FROM dbo.InventoryItems i
                LEFT JOIN dbo.InventoryTypes t ON i.InventoryTypeId = t.Id
                LEFT JOIN dbo.Employees e ON i.AssignedEmployeeId = e.Id
                """;
            if (typeId.HasValue) { sql += " WHERE i.InventoryTypeId = @TypeId"; }
            sql += " ORDER BY i.InventoryNumber";
            cmd.CommandText = sql;
            if (typeId.HasValue) cmd.Parameters.AddWithValue("@TypeId", typeId.Value);

            var list = new List<object>();
            await using var r = await cmd.ExecuteReaderAsync(ct);
            while (await r.ReadAsync(ct)) list.Add(ReadItem(r));
            return Ok(list);
        }
        catch (Exception ex) { logger.LogError(ex, "GetItems"); return StatusCode(500, ex.Message); }
    }

    [HttpGet("items/{id:guid}")]
    public async Task<IActionResult> GetItemById(Guid id, CancellationToken ct)
    {
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = """
                SELECT i.Id, i.InventoryNumber, i.Name, i.Brand, i.Model, i.SerialNumber,
                       i.Status, i.Department, i.AssignedTo, i.AssignedEmployeeId,
                       i.InventoryTypeId, i.IsPhone,
                       i.PurchaseDate, i.PurchasePrice, i.Accessories, i.CreatedAt,
                       t.Name AS TypeName,
                       e.FullName AS EmployeeName
                FROM dbo.InventoryItems i
                LEFT JOIN dbo.InventoryTypes t ON i.InventoryTypeId = t.Id
                LEFT JOIN dbo.Employees e ON i.AssignedEmployeeId = e.Id
                WHERE i.Id = @Id
                """;
            cmd.Parameters.AddWithValue("@Id", id);
            await using var r = await cmd.ExecuteReaderAsync(ct);
            if (!await r.ReadAsync(ct)) return NotFound("Equipo no encontrado.");
            return Ok(ReadItem(r));
        }
        catch (Exception ex) { logger.LogError(ex, "GetItemById {Id}", id); return StatusCode(500, ex.Message); }
    }

    [HttpPost("items")]
    public async Task<IActionResult> CreateItem([FromBody] InventoryItemDto dto, CancellationToken ct)
    {
        if (string.IsNullOrWhiteSpace(dto.Name))            return BadRequest("Nombre requerido.");
        if (string.IsNullOrWhiteSpace(dto.InventoryNumber)) return BadRequest("Número de inventario requerido.");
        var id = Guid.NewGuid();
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = """
                INSERT INTO dbo.InventoryItems
                    (Id, InventoryNumber, Name, Brand, Model, SerialNumber, Status,
                     Department, AssignedTo, AssignedEmployeeId, InventoryTypeId, IsPhone,
                     PurchaseDate, PurchasePrice, Accessories, CreatedAt)
                VALUES
                    (@Id, @InvNum, @Name, @Brand, @Model, @Serial, @Status,
                     @Dept, @AssignedTo, @EmpId, @TypeId, @IsPhone,
                     @PurchDate, @PurchPrice, @Acc, GETUTCDATE())
                """;
            cmd.Parameters.AddWithValue("@Id",         id);
            cmd.Parameters.AddWithValue("@InvNum",     dto.InventoryNumber.Trim());
            cmd.Parameters.AddWithValue("@Name",       dto.Name.Trim());
            cmd.Parameters.AddWithValue("@Brand",      (object?)dto.Brand          ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@Model",      (object?)dto.Model          ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@Serial",     (object?)dto.SerialNumber   ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@Status",     (object?)dto.Status ?? 1);
            cmd.Parameters.AddWithValue("@Dept",       (object?)dto.Department     ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@AssignedTo", (object?)dto.AssignedTo     ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@EmpId",      (object?)dto.AssignedEmployeeId ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@TypeId",     dto.InventoryTypeId);
            cmd.Parameters.AddWithValue("@IsPhone",    dto.IsPhone);
            cmd.Parameters.AddWithValue("@PurchDate",  (object?)dto.PurchaseDate   ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@PurchPrice", (object?)dto.PurchasePrice  ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@Acc",        (object?)dto.Accessories    ?? DBNull.Value);
            await cmd.ExecuteNonQueryAsync(ct);
            return Ok(new { id });
        }
        catch (Exception ex) { logger.LogError(ex, "CreateItem"); return StatusCode(500, ex.Message); }
    }

    [HttpPut("items/{id:guid}")]
    public async Task<IActionResult> UpdateItem(Guid id, [FromBody] InventoryItemDto dto, CancellationToken ct)
    {
        if (string.IsNullOrWhiteSpace(dto.Name)) return BadRequest("Nombre requerido.");
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = """
                UPDATE dbo.InventoryItems
                SET InventoryNumber = @InvNum, Name = @Name, Brand = @Brand, Model = @Model,
                    SerialNumber = @Serial, Status = @Status, Department = @Dept,
                    AssignedTo = @AssignedTo, AssignedEmployeeId = @EmpId,
                    InventoryTypeId = @TypeId, IsPhone = @IsPhone,
                    PurchaseDate = @PurchDate, PurchasePrice = @PurchPrice,
                    Accessories = @Acc, UpdatedAt = GETUTCDATE()
                WHERE Id = @Id
                """;
            cmd.Parameters.AddWithValue("@Id",         id);
            cmd.Parameters.AddWithValue("@InvNum",     dto.InventoryNumber?.Trim() ?? "");
            cmd.Parameters.AddWithValue("@Name",       dto.Name.Trim());
            cmd.Parameters.AddWithValue("@Brand",      (object?)dto.Brand          ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@Model",      (object?)dto.Model          ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@Serial",     (object?)dto.SerialNumber   ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@Status",     (object?)dto.Status ?? 1);
            cmd.Parameters.AddWithValue("@Dept",       (object?)dto.Department     ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@AssignedTo", (object?)dto.AssignedTo     ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@EmpId",      (object?)dto.AssignedEmployeeId ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@TypeId",     dto.InventoryTypeId);
            cmd.Parameters.AddWithValue("@IsPhone",    dto.IsPhone);
            cmd.Parameters.AddWithValue("@PurchDate",  (object?)dto.PurchaseDate   ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@PurchPrice", (object?)dto.PurchasePrice  ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@Acc",        (object?)dto.Accessories    ?? DBNull.Value);
            int rows = await cmd.ExecuteNonQueryAsync(ct);
            if (rows == 0) return NotFound("Equipo no encontrado.");
            return Ok(new { id });
        }
        catch (Exception ex) { logger.LogError(ex, "UpdateItem {Id}", id); return StatusCode(500, ex.Message); }
    }

    [HttpDelete("items/{id:guid}")]
    public async Task<IActionResult> DeleteItem(Guid id, CancellationToken ct)
    {
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = "DELETE FROM dbo.InventoryItems WHERE Id = @Id";
            cmd.Parameters.AddWithValue("@Id", id);
            int rows = await cmd.ExecuteNonQueryAsync(ct);
            if (rows == 0) return NotFound("Equipo no encontrado.");
            return NoContent();
        }
        catch (Exception ex) { logger.LogError(ex, "DeleteItem {Id}", id); return StatusCode(500, ex.Message); }
    }

    // ── Transfers ─────────────────────────────────────────────────────────────
    [HttpGet("items/{id:guid}/transfers")]
    public async Task<IActionResult> GetTransfers(Guid id, CancellationToken ct)
    {
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = """
                SELECT Id, FromDepartment, FromPerson, ToDepartment, ToPerson,
                       TransferDate, Notes, CreatedBy, CreatedAt
                FROM dbo.EquipmentTransfers
                WHERE InventoryItemId = @Id
                ORDER BY TransferDate DESC
                """;
            cmd.Parameters.AddWithValue("@Id", id);
            var list = new List<object>();
            await using var r = await cmd.ExecuteReaderAsync(ct);
            while (await r.ReadAsync(ct))
                list.Add(new
                {
                    id             = r.GetGuid(r.GetOrdinal("Id")),
                    fromDepartment = r.IsDBNull(r.GetOrdinal("FromDepartment")) ? null : r.GetString(r.GetOrdinal("FromDepartment")),
                    fromPerson     = r.IsDBNull(r.GetOrdinal("FromPerson"))     ? null : r.GetString(r.GetOrdinal("FromPerson")),
                    toDepartment   = r.IsDBNull(r.GetOrdinal("ToDepartment"))   ? null : r.GetString(r.GetOrdinal("ToDepartment")),
                    toPerson       = r.IsDBNull(r.GetOrdinal("ToPerson"))       ? null : r.GetString(r.GetOrdinal("ToPerson")),
                    transferDate   = r.GetDateTime(r.GetOrdinal("TransferDate")),
                    notes          = r.IsDBNull(r.GetOrdinal("Notes"))          ? null : r.GetString(r.GetOrdinal("Notes")),
                    createdBy      = r.IsDBNull(r.GetOrdinal("CreatedBy"))      ? null : r.GetString(r.GetOrdinal("CreatedBy")),
                    createdAt      = r.GetDateTime(r.GetOrdinal("CreatedAt")),
                });
            return Ok(list);
        }
        catch (Exception ex) { logger.LogError(ex, "GetTransfers {Id}", id); return StatusCode(500, ex.Message); }
    }

    [HttpPost("items/{id:guid}/transfers")]
    public async Task<IActionResult> CreateTransfer(Guid id, [FromBody] TransferDto dto, CancellationToken ct)
    {
        var tid = Guid.NewGuid();
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = """
                INSERT INTO dbo.EquipmentTransfers
                    (Id, InventoryItemId, FromDepartment, FromPerson, ToDepartment, ToPerson,
                     TransferDate, Notes, CreatedBy, CreatedAt)
                VALUES
                    (@Id, @ItemId, @FromDept, @FromPerson, @ToDept, @ToPerson,
                     @TransferDate, @Notes, @CreatedBy, GETUTCDATE())
                """;
            cmd.Parameters.AddWithValue("@Id",           tid);
            cmd.Parameters.AddWithValue("@ItemId",       id);
            cmd.Parameters.AddWithValue("@FromDept",     (object?)dto.FromDepartment ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@FromPerson",   (object?)dto.FromPerson     ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@ToDept",       (object?)dto.ToDepartment   ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@ToPerson",     (object?)dto.ToPerson       ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@TransferDate", dto.TransferDate ?? DateTime.UtcNow);
            cmd.Parameters.AddWithValue("@Notes",        (object?)dto.Notes         ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@CreatedBy",    (object?)(CurrentUsername) ?? DBNull.Value);
            await cmd.ExecuteNonQueryAsync(ct);

            // Actualizar campos del equipo
            await using var cmd2 = conn.CreateCommand();
            cmd2.CommandText = """
                UPDATE dbo.InventoryItems
                SET Department = @ToDept, AssignedTo = @ToPerson, UpdatedAt = GETUTCDATE()
                WHERE Id = @ItemId
                """;
            cmd2.Parameters.AddWithValue("@ToDept",  (object?)dto.ToDepartment ?? DBNull.Value);
            cmd2.Parameters.AddWithValue("@ToPerson",(object?)dto.ToPerson     ?? DBNull.Value);
            cmd2.Parameters.AddWithValue("@ItemId",  id);
            await cmd2.ExecuteNonQueryAsync(ct);

            return Ok(new { id = tid });
        }
        catch (Exception ex) { logger.LogError(ex, "CreateTransfer {Id}", id); return StatusCode(500, ex.Message); }
    }

    // ── Excel Export (simple CSV) ─────────────────────────────────────────────
    [HttpGet("excel/export")]
    [AllowAnonymous]
    public async Task<IActionResult> Export([FromQuery] string? access_token, CancellationToken ct)
    {
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = """
                SELECT i.InventoryNumber, i.Name, i.Brand, i.Model, i.SerialNumber,
                       i.Status, i.Department, i.AssignedTo, t.Name AS TypeName,
                       i.PurchaseDate, i.PurchasePrice, i.Accessories
                FROM dbo.InventoryItems i
                LEFT JOIN dbo.InventoryTypes t ON i.InventoryTypeId = t.Id
                ORDER BY i.InventoryNumber
                """;
            var sb = new StringBuilder();
            sb.AppendLine("NumInventario,Nombre,Marca,Modelo,NumSerie,Estado,Departamento,AsignadoA,Tipo,FechaCompra,Precio,Accesorios");
            await using var r = await cmd.ExecuteReaderAsync(ct);
            while (await r.ReadAsync(ct))
            {
                string G(int i) => r.IsDBNull(i) ? "" : r.GetString(i);
                string D(int i) => r.IsDBNull(i) ? "" : r.GetDateTime(i).ToString("yyyy-MM-dd");
                string N(int i) => r.IsDBNull(i) ? "" : r.GetDecimal(i).ToString("F2");
                sb.AppendLine(string.Join(",",
                    Csv(G(0)), Csv(G(1)), Csv(G(2)), Csv(G(3)), Csv(G(4)),
                    Csv(G(5)), Csv(G(6)), Csv(G(7)), Csv(G(8)),
                    Csv(D(9)), Csv(N(10)), Csv(G(11))));
            }
            var bytes = Encoding.UTF8.GetPreamble().Concat(Encoding.UTF8.GetBytes(sb.ToString())).ToArray();
            return File(bytes, "text/csv", "inventario.csv");
        }
        catch (Exception ex) { logger.LogError(ex, "Export Inventory"); return StatusCode(500, ex.Message); }
    }

    private static string Csv(string s) =>
        s.Contains(',') || s.Contains('"') ? $"\"{s.Replace("\"", "\"\"")}\"" : s;

    private static int? ParseStatus(object val) => val switch
    {
        int i    => i,
        long l   => (int)l,
        string s => s switch
        {
            "Activo"        or "1" => 1,
            "Mantenimiento" or "En mantenimiento" or "2" => 2,
            "Dado de baja"  or "DadoDeBaja" or "3"       => 3,
            "En almacén"    or "4"                        => 4,
            _ => int.TryParse(s, out var n) ? n : 1,
        },
        _ => 1,
    };

    private static object ReadItem(SqlDataReader r) => new
    {
        id                  = r.GetGuid(r.GetOrdinal("Id")),
        inventoryNumber     = r.GetString(r.GetOrdinal("InventoryNumber")),
        name                = r.GetString(r.GetOrdinal("Name")),
        brand               = r.IsDBNull(r.GetOrdinal("Brand"))               ? null : r.GetString(r.GetOrdinal("Brand")),
        model               = r.IsDBNull(r.GetOrdinal("Model"))               ? null : r.GetString(r.GetOrdinal("Model")),
        serialNumber        = r.IsDBNull(r.GetOrdinal("SerialNumber"))        ? null : r.GetString(r.GetOrdinal("SerialNumber")),
        status              = r.IsDBNull(r.GetOrdinal("Status"))              ? (int?)null : ParseStatus(r.GetValue(r.GetOrdinal("Status"))),
        department          = r.IsDBNull(r.GetOrdinal("Department"))          ? null : r.GetString(r.GetOrdinal("Department")),
        assignedTo          = r.IsDBNull(r.GetOrdinal("AssignedTo"))          ? null : r.GetString(r.GetOrdinal("AssignedTo")),
        assignedEmployeeId  = r.IsDBNull(r.GetOrdinal("AssignedEmployeeId"))  ? (Guid?)null : r.GetGuid(r.GetOrdinal("AssignedEmployeeId")),
        inventoryTypeId     = r.GetGuid(r.GetOrdinal("InventoryTypeId")),
        isPhone             = r.GetBoolean(r.GetOrdinal("IsPhone")),
        purchaseDate        = r.IsDBNull(r.GetOrdinal("PurchaseDate"))        ? (DateTime?)null : r.GetDateTime(r.GetOrdinal("PurchaseDate")),
        purchasePrice       = r.IsDBNull(r.GetOrdinal("PurchasePrice"))       ? (decimal?)null  : r.GetDecimal(r.GetOrdinal("PurchasePrice")),
        accessories         = r.IsDBNull(r.GetOrdinal("Accessories"))         ? null : r.GetString(r.GetOrdinal("Accessories")),
        createdAt           = r.GetDateTime(r.GetOrdinal("CreatedAt")),
        typeName            = r.IsDBNull(r.GetOrdinal("TypeName"))            ? null : r.GetString(r.GetOrdinal("TypeName")),
        employeeName        = r.IsDBNull(r.GetOrdinal("EmployeeName"))        ? null : r.GetString(r.GetOrdinal("EmployeeName")),
    };
}

// ── DTOs ──────────────────────────────────────────────────────────────────────
public record InventoryTypeDto(string Name, string? Description, string? Department, bool IsActive);

public class InventoryItemDto
{
    public string    InventoryNumber    { get; set; } = "";
    public string    Name               { get; set; } = "";
    public string?   Brand              { get; set; }
    public string?   Model              { get; set; }
    public string?   SerialNumber       { get; set; }
    public int?      Status             { get; set; }
    public string?   Department         { get; set; }
    public string?   AssignedTo         { get; set; }
    public Guid?     AssignedEmployeeId { get; set; }
    public Guid      InventoryTypeId    { get; set; }
    public bool      IsPhone            { get; set; }
    public DateTime? PurchaseDate       { get; set; }
    public decimal?  PurchasePrice      { get; set; }
    public string?   Accessories        { get; set; }
}

public record TransferDto(
    string?  FromDepartment,
    string?  FromPerson,
    string?  ToDepartment,
    string?  ToPerson,
    DateTime? TransferDate,
    string?  Notes
);
