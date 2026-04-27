using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using MiniExcelLibs;
using IConfiguration = Microsoft.Extensions.Configuration.IConfiguration;
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

            int     totalCount = 0, activeCount = 0, maintenanceCount = 0, decommissionedCount = 0, inStorageCount = 0;
            decimal activeValue = 0, totalValue = 0;

            await using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = """
                    SELECT
                        COUNT(*) AS TotalCount,
                        ISNULL(SUM(CASE WHEN TRY_CAST(Status AS NVARCHAR(50)) IN ('1','Activo')                                  THEN 1 ELSE 0 END), 0) AS ActiveCount,
                        ISNULL(SUM(CASE WHEN TRY_CAST(Status AS NVARCHAR(50)) IN ('2','Mantenimiento','En mantenimiento')         THEN 1 ELSE 0 END), 0) AS MaintenanceCount,
                        ISNULL(SUM(CASE WHEN TRY_CAST(Status AS NVARCHAR(50)) IN ('3','Dado de baja','DadoDeBaja')                THEN 1 ELSE 0 END), 0) AS DecommissionedCount,
                        ISNULL(SUM(CASE WHEN TRY_CAST(Status AS NVARCHAR(50)) IN ('4','En almacén','Almacén','EnAlmacen')         THEN 1 ELSE 0 END), 0) AS InStorageCount,
                        ISNULL(SUM(CASE WHEN TRY_CAST(Status AS NVARCHAR(50)) IN ('1','Activo') THEN ISNULL(PurchasePrice,0) ELSE 0 END), 0) AS ActiveValue,
                        ISNULL(SUM(ISNULL(PurchasePrice, 0)), 0) AS TotalValue
                    FROM dbo.InventoryItems
                    """;
                await using var r = await cmd.ExecuteReaderAsync(ct);
                if (await r.ReadAsync(ct))
                {
                    totalCount          = r.IsDBNull(0) ? 0 : r.GetInt32(0);
                    activeCount         = r.IsDBNull(1) ? 0 : r.GetInt32(1);
                    maintenanceCount    = r.IsDBNull(2) ? 0 : r.GetInt32(2);
                    decommissionedCount = r.IsDBNull(3) ? 0 : r.GetInt32(3);
                    inStorageCount      = r.IsDBNull(4) ? 0 : r.GetInt32(4);
                    activeValue         = r.IsDBNull(5) ? 0 : r.GetDecimal(5);
                    totalValue          = r.IsDBNull(6) ? 0 : r.GetDecimal(6);
                }
            }

            var byType = new List<object>();
            await using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = """
                    SELECT t.Name AS TypeName,
                           COUNT(i.Id) AS TotalCount,
                           ISNULL(SUM(CASE WHEN TRY_CAST(i.Status AS NVARCHAR(50)) IN ('1','Activo') THEN 1 ELSE 0 END), 0) AS ActiveCount
                    FROM dbo.InventoryTypes t
                    LEFT JOIN dbo.InventoryItems i ON i.InventoryTypeId = t.Id
                    WHERE t.IsActive = 1
                    GROUP BY t.Id, t.Name
                    ORDER BY COUNT(i.Id) DESC
                    """;
                await using var r = await cmd.ExecuteReaderAsync(ct);
                while (await r.ReadAsync(ct))
                    byType.Add(new { typeName = r.GetString(0), totalCount = r.GetInt32(1), activeCount = r.GetInt32(2) });
            }

            var recentTransfers = new List<object>();
            await using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = """
                    SELECT TOP 5 et.Id, et.TransferDate, et.FromPerson, et.ToPerson,
                                 i.Name AS ItemName, i.InventoryNumber
                    FROM dbo.EquipmentTransfers et
                    JOIN dbo.InventoryItems i ON et.InventoryItemId = i.Id
                    ORDER BY et.TransferDate DESC
                    """;
                await using var r = await cmd.ExecuteReaderAsync(ct);
                while (await r.ReadAsync(ct))
                    recentTransfers.Add(new
                    {
                        id              = r.GetGuid(0),
                        transferDate    = r.GetDateTime(1),
                        fromPerson      = r.IsDBNull(2) ? null : r.GetString(2),
                        toPerson        = r.IsDBNull(3) ? null : r.GetString(3),
                        itemName        = r.GetString(4),
                        inventoryNumber = r.GetString(5),
                    });
            }

            return Ok(new
            {
                totalCount,
                activeCount,
                maintenanceCount,
                decommissionedCount,
                inStorageCount,
                activeValue,
                totalValue,
                byType,
                recentTransfers,
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

    // ── Excel Export ──────────────────────────────────────────────────────────
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
            var rows = new List<Dictionary<string, object?>>();
            await using var r = await cmd.ExecuteReaderAsync(ct);
            while (await r.ReadAsync(ct))
            {
                var statusRaw = r.IsDBNull(5) ? null : r.GetValue(5);
                var statusStr = statusRaw switch
                {
                    int n    => StatusToString(n),
                    long n   => StatusToString((int)n),
                    string s => s,
                    _        => ""
                };
                rows.Add(new Dictionary<string, object?>
                {
                    ["NumInventario"] = r.IsDBNull(0)  ? ""    : r.GetString(0),
                    ["Nombre"]        = r.IsDBNull(1)  ? ""    : r.GetString(1),
                    ["Marca"]         = r.IsDBNull(2)  ? ""    : r.GetString(2),
                    ["Modelo"]        = r.IsDBNull(3)  ? ""    : r.GetString(3),
                    ["NumSerie"]      = r.IsDBNull(4)  ? ""    : r.GetString(4),
                    ["Estado"]        = statusStr,
                    ["Departamento"]  = r.IsDBNull(6)  ? ""    : r.GetString(6),
                    ["AsignadoA"]     = r.IsDBNull(7)  ? ""    : r.GetString(7),
                    ["Categoria"]     = r.IsDBNull(8)  ? ""    : r.GetString(8),
                    ["FechaCompra"]   = r.IsDBNull(9)  ? ""    : r.GetDateTime(9).ToString("yyyy-MM-dd"),
                    ["Precio"]        = r.IsDBNull(10) ? ""    : r.GetDecimal(10).ToString("F2"),
                    ["Accesorios"]    = r.IsDBNull(11) ? ""    : r.GetString(11),
                });
            }
            var ms = new MemoryStream();
            await ms.SaveAsAsync(rows, cancellationToken: ct);
            ms.Position = 0;
            return File(ms, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "inventario.xlsx");
        }
        catch (Exception ex) { logger.LogError(ex, "Export Inventory"); return StatusCode(500, ex.Message); }
    }

    // ── Excel Template ────────────────────────────────────────────────────────
    [HttpGet("excel/template")]
    [AllowAnonymous]
    public async Task<IActionResult> Template(CancellationToken ct)
    {
        var rows = new[]
        {
            new Dictionary<string, object?>
            {
                ["NumInventario"] = "EQT-0001",
                ["Nombre"]        = "Laptop Dell Latitude",
                ["Marca"]         = "Dell",
                ["Modelo"]        = "Latitude 5420",
                ["NumSerie"]      = "SN-ABC12345",
                ["Categoria"]     = "Laptop",
                ["Estado"]        = "Activo",
                ["Departamento"]  = "Coordinación de TI",
                ["AsignadoA"]     = "Juan Pérez",
                ["FechaCompra"]   = "2024-01-15",
                ["Precio"]        = "12000.00",
                ["Accesorios"]    = "Cargador, funda",
            }
        };
        var ms = new MemoryStream();
        await ms.SaveAsAsync(rows, cancellationToken: ct);
        ms.Position = 0;
        return File(ms, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "plantilla-inventario.xlsx");
    }

    // ── Excel Import Preview ──────────────────────────────────────────────────
    [HttpPost("excel/import/preview")]
    public async Task<IActionResult> ImportPreview(IFormFile file, CancellationToken ct)
    {
        if (file is null || file.Length == 0) return BadRequest("Archivo requerido.");
        try
        {
            // Cargar tipos de inventario para validar Categoria
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            var typeMap = new Dictionary<string, Guid>(StringComparer.OrdinalIgnoreCase);
            await using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = "SELECT Id, Name FROM dbo.InventoryTypes WHERE IsActive = 1";
                await using var tr = await cmd.ExecuteReaderAsync(ct);
                while (await tr.ReadAsync(ct))
                    typeMap[tr.GetString(1)] = tr.GetGuid(0);
            }

            await using var stream = file.OpenReadStream();
            var rawRows = await stream.QueryAsync(useHeaderRow: true, cancellationToken: ct);

            var validRows  = new List<object>();
            var errors     = new List<object>();
            int rowNumber  = 1;

            foreach (var rawRow in rawRows)
            {
                rowNumber++;
                var row = (IDictionary<string, object?>)rawRow;

                string Get(string key) =>
                    row.TryGetValue(key, out var v) ? v?.ToString()?.Trim() ?? "" : "";

                var name       = Get("Nombre");
                var category   = Get("Categoria");
                var status     = Get("Estado");

                if (string.IsNullOrWhiteSpace(name))
                {
                    errors.Add(new { rowNumber, field = "Nombre", message = "El nombre es requerido." });
                    continue;
                }
                if (string.IsNullOrWhiteSpace(category) || !typeMap.ContainsKey(category))
                {
                    errors.Add(new { rowNumber, field = "Categoria",
                        message = string.IsNullOrWhiteSpace(category)
                            ? "La categoría es requerida."
                            : $"Categoría '{category}' no existe en el sistema." });
                    continue;
                }

                var normalizedStatus = status switch
                {
                    "Activo"        or "1" => "Activo",
                    "Mantenimiento" or "2" => "Mantenimiento",
                    "Dado de baja"  or "3" => "Dado de baja",
                    "En almacén"    or "Almacén" or "4" => "En almacén",
                    "" => "Activo",
                    _  => null,
                };
                if (normalizedStatus is null)
                {
                    errors.Add(new { rowNumber, field = "Estado",
                        message = $"Estado '{status}' no válido. Use: Activo, Mantenimiento, Dado de baja, En almacén." });
                    continue;
                }

                decimal? price = null;
                var priceStr = Get("Precio");
                if (!string.IsNullOrWhiteSpace(priceStr) && decimal.TryParse(priceStr, out var p))
                    price = p;

                DateTime? purchaseDate = null;
                var dateStr = Get("FechaCompra");
                if (!string.IsNullOrWhiteSpace(dateStr) && DateTime.TryParse(dateStr, out var d))
                    purchaseDate = d;

                validRows.Add(new
                {
                    inventoryNumber = Get("NumInventario"),
                    name,
                    brand           = Get("Marca"),
                    model           = Get("Modelo"),
                    serialNumber    = Get("NumSerie"),
                    category,
                    inventoryTypeId = typeMap[category],
                    status          = normalizedStatus,
                    department      = Get("Departamento"),
                    assignedTo      = Get("AsignadoA"),
                    purchaseDate,
                    purchasePrice   = price,
                    accessories     = Get("Accesorios"),
                });
            }

            return Ok(new { totalRows = rowNumber - 1, validRows, errors });
        }
        catch (Exception ex) { logger.LogError(ex, "ImportPreview"); return StatusCode(500, ex.Message); }
    }

    // ── Excel Import Confirm ──────────────────────────────────────────────────
    [HttpPost("excel/import/confirm")]
    public async Task<IActionResult> ImportConfirm([FromBody] List<ImportRowDto> rows, CancellationToken ct)
    {
        if (rows is null || rows.Count == 0) return BadRequest("Sin filas para importar.");
        int inserted = 0, skipped = 0, autoGenerated = 0;
        var warnings = new List<string>();
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);

            foreach (var row in rows)
            {
                var invNum = row.InventoryNumber?.Trim();

                // Auto-generar número si viene vacío
                if (string.IsNullOrWhiteSpace(invNum))
                {
                    await using var cmdN = conn.CreateCommand();
                    cmdN.CommandText = "SELECT COUNT(*) FROM dbo.InventoryItems";
                    var cnt = (int)await cmdN.ExecuteScalarAsync(ct)!;
                    invNum = $"IMP-{(cnt + 1):D4}";
                    autoGenerated++;
                }

                // Verificar duplicado
                await using var cmdCheck = conn.CreateCommand();
                cmdCheck.CommandText = "SELECT COUNT(1) FROM dbo.InventoryItems WHERE InventoryNumber = @N";
                cmdCheck.Parameters.AddWithValue("@N", invNum);
                var exists = (int)await cmdCheck.ExecuteScalarAsync(ct)! > 0;
                if (exists)
                {
                    warnings.Add($"{invNum}: número ya existe, omitido.");
                    skipped++;
                    continue;
                }

                var statusInt = row.Status switch
                {
                    "Activo"       => 1,
                    "Mantenimiento"=> 2,
                    "Dado de baja" => 3,
                    "En almacén"   => 4,
                    _              => 1,
                };

                await using var cmd = conn.CreateCommand();
                cmd.CommandText = """
                    INSERT INTO dbo.InventoryItems
                        (Id, InventoryNumber, Name, Brand, Model, SerialNumber, Status,
                         Department, AssignedTo, InventoryTypeId, IsPhone,
                         PurchaseDate, PurchasePrice, Accessories, CreatedAt)
                    VALUES
                        (@Id, @InvNum, @Name, @Brand, @Model, @Serial, @Status,
                         @Dept, @AssignedTo, @TypeId, 0,
                         @PurchDate, @PurchPrice, @Acc, GETUTCDATE())
                    """;
                cmd.Parameters.AddWithValue("@Id",        Guid.NewGuid());
                cmd.Parameters.AddWithValue("@InvNum",    invNum);
                cmd.Parameters.AddWithValue("@Name",      row.Name?.Trim() ?? "");
                cmd.Parameters.AddWithValue("@Brand",     (object?)row.Brand      ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@Model",     (object?)row.Model      ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@Serial",    (object?)row.SerialNumber ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@Status",    statusInt);
                cmd.Parameters.AddWithValue("@Dept",      (object?)row.Department ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@AssignedTo",(object?)row.AssignedTo ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@TypeId",    row.InventoryTypeId);
                cmd.Parameters.AddWithValue("@PurchDate", (object?)row.PurchaseDate  ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@PurchPrice",(object?)row.PurchasePrice ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@Acc",       (object?)row.Accessories   ?? DBNull.Value);
                await cmd.ExecuteNonQueryAsync(ct);
                inserted++;
            }

            return Ok(new { inserted, skipped, autoGenerated, warnings });
        }
        catch (Exception ex) { logger.LogError(ex, "ImportConfirm"); return StatusCode(500, ex.Message); }
    }

    private static string StatusToString(int n) => n switch
    {
        1 => "Activo",
        2 => "Mantenimiento",
        3 => "Dado de baja",
        4 => "En almacén",
        _ => n.ToString(),
    };

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

public class ImportRowDto
{
    public string?   InventoryNumber  { get; set; }
    public string    Name             { get; set; } = "";
    public string?   Brand            { get; set; }
    public string?   Model            { get; set; }
    public string?   SerialNumber     { get; set; }
    public string?   Status           { get; set; }
    public string?   Department       { get; set; }
    public string?   AssignedTo       { get; set; }
    public Guid      InventoryTypeId  { get; set; }
    public DateTime? PurchaseDate     { get; set; }
    public decimal?  PurchasePrice    { get; set; }
    public string?   Accessories      { get; set; }
}
