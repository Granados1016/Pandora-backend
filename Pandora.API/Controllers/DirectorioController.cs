using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using System.Security.Claims;

namespace Pandora.API.Controllers;

[ApiController]
[Route("api/directorio")]
[Authorize]
public class DirectorioController(
    IConfiguration config,
    ILogger<DirectorioController> logger) : ControllerBase
{
    private SqlConnection Conn() => new(config.GetConnectionString("PandoraDb"));

    private bool IsAdmin =>
        User.IsInRole("Admin") ||
        User.Claims.Any(c => (c.Type.EndsWith("role", StringComparison.OrdinalIgnoreCase) ||
                               c.Type.EndsWith("roles", StringComparison.OrdinalIgnoreCase)) &&
                              c.Value.Equals("Admin", StringComparison.OrdinalIgnoreCase));

    // ── GET /api/directorio ───────────────────────────────────────────────────
    [HttpGet]
    public async Task<IActionResult> GetAll(
        [FromQuery] string? search     = null,
        [FromQuery] int?    department = null,
        CancellationToken ct = default)
    {
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await EnsureTableAsync(conn, ct);

            var where = new List<string> { "d.IsActive = 1" };
            if (!string.IsNullOrWhiteSpace(search))
                where.Add("(d.FullName LIKE @Search OR d.Position LIKE @Search OR d.Email LIKE @Search OR d.Extension LIKE @Search OR dep.Name LIKE @Search)");
            if (department.HasValue)
                where.Add("d.DepartmentId = @Dept");

            await using var cmd = conn.CreateCommand();
            cmd.CommandText = $"""
                SELECT d.Id, d.FullName, d.Position, d.Email, d.Phone, d.Extension,
                       d.PhotoUrl, d.Birthday, d.StartDate, d.IsActive,
                       dep.Name AS Department, dep.Id AS DepartmentId,
                       d.Notes, d.CreatedAt
                FROM dbo.DirectoryEmployees d
                LEFT JOIN dbo.Departments dep ON dep.Id = d.DepartmentId
                WHERE {string.Join(" AND ", where)}
                ORDER BY d.FullName
                """;
            if (!string.IsNullOrWhiteSpace(search))
                cmd.Parameters.AddWithValue("@Search", $"%{search.Trim()}%");
            if (department.HasValue)
                cmd.Parameters.AddWithValue("@Dept", department.Value);

            var items = new List<object>();
            await using var r = await cmd.ExecuteReaderAsync(ct);
            while (await r.ReadAsync(ct))
            {
                items.Add(new
                {
                    id           = r.GetInt32(0),
                    fullName     = r.IsDBNull(1) ? "" : r.GetString(1),
                    position     = r.IsDBNull(2) ? "" : r.GetString(2),
                    email        = r.IsDBNull(3) ? "" : r.GetString(3),
                    phone        = r.IsDBNull(4) ? "" : r.GetString(4),
                    extension    = r.IsDBNull(5) ? "" : r.GetString(5),
                    photoUrl     = r.IsDBNull(6) ? null : r.GetString(6),
                    birthday     = r.IsDBNull(7) ? null : r.GetFieldValue<DateOnly>(7).ToString("yyyy-MM-dd"),
                    startDate    = r.IsDBNull(8) ? null : r.GetFieldValue<DateOnly>(8).ToString("yyyy-MM-dd"),
                    isActive     = r.GetBoolean(9),
                    department   = r.IsDBNull(10) ? "" : r.GetString(10),
                    departmentId = r.IsDBNull(11) ? (int?)null : r.GetInt32(11),
                    notes        = r.IsDBNull(12) ? "" : r.GetString(12),
                    createdAt    = ((DateTime)r.GetValue(13)).ToString("yyyy-MM-dd"),
                });
            }
            return Ok(items);
        }
        catch (Exception ex) { logger.LogError(ex, "Directorio.GetAll"); return StatusCode(500, ex.Message); }
    }

    // ── POST /api/directorio ──────────────────────────────────────────────────
    [HttpPost]
    [Authorize(Roles = "Admin")]
    public async Task<IActionResult> Create([FromBody] DirectorioDto dto, CancellationToken ct)
    {
        if (string.IsNullOrWhiteSpace(dto.FullName)) return BadRequest("Nombre requerido.");
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await EnsureTableAsync(conn, ct);

            await using var cmd = conn.CreateCommand();
            cmd.CommandText = """
                INSERT INTO dbo.DirectoryEmployees
                    (FullName, Position, Email, Phone, Extension, PhotoUrl,
                     Birthday, StartDate, DepartmentId, Notes, IsActive, CreatedAt)
                OUTPUT INSERTED.Id
                VALUES
                    (@FullName, @Position, @Email, @Phone, @Ext, @Photo,
                     @Birthday, @StartDate, @DeptId, @Notes, 1, GETUTCDATE())
                """;
            cmd.Parameters.AddWithValue("@FullName",  dto.FullName.Trim());
            cmd.Parameters.AddWithValue("@Position",  string.IsNullOrWhiteSpace(dto.Position)  ? DBNull.Value : dto.Position.Trim());
            cmd.Parameters.AddWithValue("@Email",     string.IsNullOrWhiteSpace(dto.Email)     ? DBNull.Value : dto.Email.Trim());
            cmd.Parameters.AddWithValue("@Phone",     string.IsNullOrWhiteSpace(dto.Phone)     ? DBNull.Value : dto.Phone.Trim());
            cmd.Parameters.AddWithValue("@Ext",       string.IsNullOrWhiteSpace(dto.Extension) ? DBNull.Value : dto.Extension.Trim());
            cmd.Parameters.AddWithValue("@Photo",     string.IsNullOrWhiteSpace(dto.PhotoUrl)  ? DBNull.Value : dto.PhotoUrl.Trim());
            cmd.Parameters.AddWithValue("@Birthday",  dto.Birthday  == null ? DBNull.Value : (object)dto.Birthday.Value);
            cmd.Parameters.AddWithValue("@StartDate", dto.StartDate == null ? DBNull.Value : (object)dto.StartDate.Value);
            cmd.Parameters.AddWithValue("@DeptId",    dto.DepartmentId == null ? DBNull.Value : (object)dto.DepartmentId.Value);
            cmd.Parameters.AddWithValue("@Notes",     string.IsNullOrWhiteSpace(dto.Notes) ? DBNull.Value : dto.Notes.Trim());

            var newId = (int)(await cmd.ExecuteScalarAsync(ct))!;
            return Ok(new { id = newId });
        }
        catch (Exception ex) { logger.LogError(ex, "Directorio.Create"); return StatusCode(500, ex.Message); }
    }

    // ── PUT /api/directorio/{id} ──────────────────────────────────────────────
    [HttpPut("{id:int}")]
    [Authorize(Roles = "Admin")]
    public async Task<IActionResult> Update(int id, [FromBody] DirectorioDto dto, CancellationToken ct)
    {
        if (string.IsNullOrWhiteSpace(dto.FullName)) return BadRequest("Nombre requerido.");
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = """
                UPDATE dbo.DirectoryEmployees
                SET FullName = @FullName, Position = @Position, Email = @Email,
                    Phone = @Phone, Extension = @Ext, PhotoUrl = @Photo,
                    Birthday = @Birthday, StartDate = @StartDate,
                    DepartmentId = @DeptId, Notes = @Notes, IsActive = @Active
                WHERE Id = @Id
                """;
            cmd.Parameters.AddWithValue("@Id",        id);
            cmd.Parameters.AddWithValue("@FullName",  dto.FullName.Trim());
            cmd.Parameters.AddWithValue("@Position",  string.IsNullOrWhiteSpace(dto.Position)  ? DBNull.Value : dto.Position.Trim());
            cmd.Parameters.AddWithValue("@Email",     string.IsNullOrWhiteSpace(dto.Email)     ? DBNull.Value : dto.Email.Trim());
            cmd.Parameters.AddWithValue("@Phone",     string.IsNullOrWhiteSpace(dto.Phone)     ? DBNull.Value : dto.Phone.Trim());
            cmd.Parameters.AddWithValue("@Ext",       string.IsNullOrWhiteSpace(dto.Extension) ? DBNull.Value : dto.Extension.Trim());
            cmd.Parameters.AddWithValue("@Photo",     string.IsNullOrWhiteSpace(dto.PhotoUrl)  ? DBNull.Value : dto.PhotoUrl.Trim());
            cmd.Parameters.AddWithValue("@Birthday",  dto.Birthday  == null ? DBNull.Value : (object)dto.Birthday.Value);
            cmd.Parameters.AddWithValue("@StartDate", dto.StartDate == null ? DBNull.Value : (object)dto.StartDate.Value);
            cmd.Parameters.AddWithValue("@DeptId",    dto.DepartmentId == null ? DBNull.Value : (object)dto.DepartmentId.Value);
            cmd.Parameters.AddWithValue("@Notes",     string.IsNullOrWhiteSpace(dto.Notes) ? DBNull.Value : dto.Notes.Trim());
            cmd.Parameters.AddWithValue("@Active",    dto.IsActive ?? true);
            await cmd.ExecuteNonQueryAsync(ct);
            return NoContent();
        }
        catch (Exception ex) { logger.LogError(ex, "Directorio.Update {Id}", id); return StatusCode(500, ex.Message); }
    }

    // ── DELETE /api/directorio/{id} ───────────────────────────────────────────
    [HttpDelete("{id:int}")]
    [Authorize(Roles = "Admin")]
    public async Task<IActionResult> Delete(int id, CancellationToken ct)
    {
        try
        {
            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = "DELETE FROM dbo.DirectoryEmployees WHERE Id = @Id";
            cmd.Parameters.AddWithValue("@Id", id);
            await cmd.ExecuteNonQueryAsync(ct);
            return NoContent();
        }
        catch (Exception ex) { logger.LogError(ex, "Directorio.Delete {Id}", id); return StatusCode(500, ex.Message); }
    }

    // ── POST /api/directorio/{id}/photo ───────────────────────────────────────
    [HttpPost("{id:int}/photo")]
    [Authorize(Roles = "Admin")]
    [Consumes("multipart/form-data")]
    public async Task<IActionResult> UploadPhoto(int id, IFormFile file, CancellationToken ct)
    {
        if (file == null || file.Length == 0) return BadRequest("Archivo requerido.");
        var ext  = Path.GetExtension(file.FileName).ToLower();
        var safe = new[] { ".jpg", ".jpeg", ".png", ".webp" };
        if (!safe.Contains(ext)) return BadRequest("Formato no soportado.");
        if (file.Length > 5 * 1024 * 1024) return BadRequest("El archivo no puede superar 5 MB.");

        try
        {
            var dir  = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "uploads", "directorio");
            Directory.CreateDirectory(dir);
            var name = $"emp_{id}_{Guid.NewGuid():N}{ext}";
            var path = Path.Combine(dir, name);
            await using (var stream = System.IO.File.Create(path)) await file.CopyToAsync(stream, ct);
            var url = $"/uploads/directorio/{name}";

            await using var conn = Conn();
            await conn.OpenAsync(ct);
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = "UPDATE dbo.DirectoryEmployees SET PhotoUrl = @Url WHERE Id = @Id";
            cmd.Parameters.AddWithValue("@Url", url);
            cmd.Parameters.AddWithValue("@Id",  id);
            await cmd.ExecuteNonQueryAsync(ct);
            return Ok(new { url });
        }
        catch (Exception ex) { logger.LogError(ex, "Directorio.UploadPhoto {Id}", id); return StatusCode(500, ex.Message); }
    }

    // ── Helpers ───────────────────────────────────────────────────────────────
    private static async Task EnsureTableAsync(SqlConnection conn, CancellationToken ct)
    {
        await using var cmd = conn.CreateCommand();
        cmd.CommandText = """
            IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name='DirectoryEmployees' AND schema_id=SCHEMA_ID('dbo'))
            BEGIN
                CREATE TABLE dbo.DirectoryEmployees (
                    Id           INT          IDENTITY(1,1) PRIMARY KEY,
                    FullName     NVARCHAR(200) NOT NULL,
                    Position     NVARCHAR(150),
                    Email        NVARCHAR(200),
                    Phone        NVARCHAR(50),
                    Extension    NVARCHAR(20),
                    PhotoUrl     NVARCHAR(500),
                    Birthday     DATE,
                    StartDate    DATE,
                    DepartmentId INT,
                    Notes        NVARCHAR(MAX),
                    IsActive     BIT NOT NULL DEFAULT 1,
                    CreatedAt    DATETIME2 NOT NULL DEFAULT GETUTCDATE()
                )
            END
            """;
        await cmd.ExecuteNonQueryAsync(ct);
    }
}

// ── DTOs ──────────────────────────────────────────────────────────────────────
public record DirectorioDto(
    string   FullName,
    string?  Position,
    string?  Email,
    string?  Phone,
    string?  Extension,
    string?  PhotoUrl,
    DateOnly? Birthday,
    DateOnly? StartDate,
    int?     DepartmentId,
    string?  Notes,
    bool?    IsActive);
