using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;

namespace Pandora.API.Controllers;

[ApiController]
[Route("api/catalogs")]
[Authorize]
public class CatalogsController(IConfiguration config, ILogger<CatalogsController> logger) : ControllerBase
{
    private SqlConnection Conn() => new(config.GetConnectionString("PandoraDb"));

    // ════════════════════════════════════════════════════════════════════════
    //  DEPARTMENTS
    // ════════════════════════════════════════════════════════════════════════

    [HttpGet("departments")]
    public async Task<IActionResult> GetDepartments(CancellationToken ct)
    {
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = """
                SELECT Id, Name, InventoryPrefix, IsActive, CreatedAt
                FROM dbo.Departments
                ORDER BY Name
                """;
            var list = new List<object>();
            await using var r = await cmd.ExecuteReaderAsync(ct);
            while (await r.ReadAsync(ct))
                list.Add(new
                {
                    id              = r.GetGuid(r.GetOrdinal("Id")),
                    name            = r.GetString(r.GetOrdinal("Name")),
                    inventoryPrefix = r.IsDBNull(r.GetOrdinal("InventoryPrefix")) ? null : r.GetString(r.GetOrdinal("InventoryPrefix")),
                    isActive        = r.GetBoolean(r.GetOrdinal("IsActive")),
                    createdAt       = r.GetDateTime(r.GetOrdinal("CreatedAt")),
                });
            return Ok(list);
        }
        catch (Exception ex) { logger.LogError(ex, "GetDepartments"); return StatusCode(500, ex.Message); }
    }

    [HttpPost("departments")]
    public async Task<IActionResult> CreateDepartment([FromBody] DepartmentDto dto, CancellationToken ct)
    {
        if (string.IsNullOrWhiteSpace(dto.Name)) return BadRequest("Nombre requerido.");
        var id = Guid.NewGuid();
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = """
                INSERT INTO dbo.Departments (Id, Name, InventoryPrefix, IsActive, CreatedAt)
                VALUES (@Id, @Name, @Prefix, @IsActive, GETUTCDATE())
                """;
            cmd.Parameters.AddWithValue("@Id",       id);
            cmd.Parameters.AddWithValue("@Name",     dto.Name.Trim());
            cmd.Parameters.AddWithValue("@Prefix",   (object?)dto.InventoryPrefix ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@IsActive", dto.IsActive);
            await cmd.ExecuteNonQueryAsync(ct);
            return Ok(new { id });
        }
        catch (Exception ex) { logger.LogError(ex, "CreateDepartment"); return StatusCode(500, ex.Message); }
    }

    [HttpPut("departments/{id:guid}")]
    public async Task<IActionResult> UpdateDepartment(Guid id, [FromBody] DepartmentDto dto, CancellationToken ct)
    {
        if (string.IsNullOrWhiteSpace(dto.Name)) return BadRequest("Nombre requerido.");
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = """
                UPDATE dbo.Departments
                SET Name = @Name, InventoryPrefix = @Prefix, IsActive = @IsActive, UpdatedAt = GETUTCDATE()
                WHERE Id = @Id
                """;
            cmd.Parameters.AddWithValue("@Id",       id);
            cmd.Parameters.AddWithValue("@Name",     dto.Name.Trim());
            cmd.Parameters.AddWithValue("@Prefix",   (object?)dto.InventoryPrefix ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@IsActive", dto.IsActive);
            int rows = await cmd.ExecuteNonQueryAsync(ct);
            if (rows == 0) return NotFound("Departamento no encontrado.");
            return Ok(new { id });
        }
        catch (Exception ex) { logger.LogError(ex, "UpdateDepartment {Id}", id); return StatusCode(500, ex.Message); }
    }

    [HttpDelete("departments/{id:guid}")]
    public async Task<IActionResult> DeleteDepartment(Guid id, CancellationToken ct)
    {
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = "DELETE FROM dbo.Departments WHERE Id = @Id";
            cmd.Parameters.AddWithValue("@Id", id);
            int rows = await cmd.ExecuteNonQueryAsync(ct);
            if (rows == 0) return NotFound("Departamento no encontrado.");
            return NoContent();
        }
        catch (Exception ex) { logger.LogError(ex, "DeleteDepartment {Id}", id); return StatusCode(500, ex.Message); }
    }

    // ════════════════════════════════════════════════════════════════════════
    //  EMPLOYEES
    // ════════════════════════════════════════════════════════════════════════

    [HttpGet("employees")]
    public async Task<IActionResult> GetEmployees([FromQuery] Guid? departmentId, CancellationToken ct)
    {
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await using var cmd = conn.CreateCommand();
            var sql = """
                SELECT e.Id, e.FullName, e.Email, e.Phone, e.Position,
                       e.DepartmentId, d.Name AS DepartmentName, e.IsActive, e.CreatedAt
                FROM dbo.Employees e
                LEFT JOIN dbo.Departments d ON e.DepartmentId = d.Id
                """;
            if (departmentId.HasValue) { sql += " WHERE e.DepartmentId = @DeptId"; }
            sql += " ORDER BY e.FullName";
            cmd.CommandText = sql;
            if (departmentId.HasValue) cmd.Parameters.AddWithValue("@DeptId", departmentId.Value);

            var list = new List<object>();
            await using var r = await cmd.ExecuteReaderAsync(ct);
            while (await r.ReadAsync(ct))
                list.Add(new
                {
                    id             = r.GetGuid(r.GetOrdinal("Id")),
                    fullName       = r.IsDBNull(r.GetOrdinal("FullName"))       ? null : r.GetString(r.GetOrdinal("FullName")),
                    email          = r.IsDBNull(r.GetOrdinal("Email"))          ? null : r.GetString(r.GetOrdinal("Email")),
                    phone          = r.IsDBNull(r.GetOrdinal("Phone"))          ? null : r.GetString(r.GetOrdinal("Phone")),
                    position       = r.IsDBNull(r.GetOrdinal("Position"))       ? null : r.GetString(r.GetOrdinal("Position")),
                    departmentId   = r.IsDBNull(r.GetOrdinal("DepartmentId"))   ? (Guid?)null : r.GetGuid(r.GetOrdinal("DepartmentId")),
                    departmentName = r.IsDBNull(r.GetOrdinal("DepartmentName")) ? null : r.GetString(r.GetOrdinal("DepartmentName")),
                    isActive       = r.GetBoolean(r.GetOrdinal("IsActive")),
                    createdAt      = r.GetDateTime(r.GetOrdinal("CreatedAt")),
                });
            return Ok(list);
        }
        catch (Exception ex) { logger.LogError(ex, "GetEmployees"); return StatusCode(500, ex.Message); }
    }

    [HttpPost("employees")]
    public async Task<IActionResult> CreateEmployee([FromBody] EmployeeDto dto, CancellationToken ct)
    {
        if (string.IsNullOrWhiteSpace(dto.FullName)) return BadRequest("Nombre requerido.");
        var id = Guid.NewGuid();
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = """
                INSERT INTO dbo.Employees (Id, FullName, Email, Phone, Position, DepartmentId, IsActive, CreatedAt)
                VALUES (@Id, @FullName, @Email, @Phone, @Position, @DeptId, @IsActive, GETUTCDATE())
                """;
            cmd.Parameters.AddWithValue("@Id",       id);
            cmd.Parameters.AddWithValue("@FullName", dto.FullName.Trim());
            cmd.Parameters.AddWithValue("@Email",    (object?)dto.Email      ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@Phone",    (object?)dto.Phone      ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@Position", (object?)dto.Position   ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@DeptId",   (object?)dto.DepartmentId ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@IsActive", dto.IsActive);
            await cmd.ExecuteNonQueryAsync(ct);
            return Ok(new { id });
        }
        catch (Exception ex) { logger.LogError(ex, "CreateEmployee"); return StatusCode(500, ex.Message); }
    }

    [HttpPut("employees/{id:guid}")]
    public async Task<IActionResult> UpdateEmployee(Guid id, [FromBody] EmployeeDto dto, CancellationToken ct)
    {
        if (string.IsNullOrWhiteSpace(dto.FullName)) return BadRequest("Nombre requerido.");
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = """
                UPDATE dbo.Employees
                SET FullName = @FullName, Email = @Email, Phone = @Phone,
                    Position = @Position, DepartmentId = @DeptId, IsActive = @IsActive, UpdatedAt = GETUTCDATE()
                WHERE Id = @Id
                """;
            cmd.Parameters.AddWithValue("@Id",       id);
            cmd.Parameters.AddWithValue("@FullName", dto.FullName.Trim());
            cmd.Parameters.AddWithValue("@Email",    (object?)dto.Email      ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@Phone",    (object?)dto.Phone      ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@Position", (object?)dto.Position   ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@DeptId",   (object?)dto.DepartmentId ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@IsActive", dto.IsActive);
            int rows = await cmd.ExecuteNonQueryAsync(ct);
            if (rows == 0) return NotFound("Empleado no encontrado.");
            return Ok(new { id });
        }
        catch (Exception ex) { logger.LogError(ex, "UpdateEmployee {Id}", id); return StatusCode(500, ex.Message); }
    }

    [HttpDelete("employees/{id:guid}")]
    public async Task<IActionResult> DeleteEmployee(Guid id, CancellationToken ct)
    {
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = "DELETE FROM dbo.Employees WHERE Id = @Id";
            cmd.Parameters.AddWithValue("@Id", id);
            int rows = await cmd.ExecuteNonQueryAsync(ct);
            if (rows == 0) return NotFound("Empleado no encontrado.");
            return NoContent();
        }
        catch (Exception ex) { logger.LogError(ex, "DeleteEmployee {Id}", id); return StatusCode(500, ex.Message); }
    }
}

// ── DTOs ──────────────────────────────────────────────────────────────────────
public record DepartmentDto(string Name, string? InventoryPrefix, bool IsActive);

public record EmployeeDto(
    string  FullName,
    string? Email,
    string? Phone,
    string? Position,
    Guid?   DepartmentId,
    bool    IsActive
);
