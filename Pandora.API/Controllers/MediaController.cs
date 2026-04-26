using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;

namespace Pandora.API.Controllers;

[ApiController]
[Route("api/media")]
[Authorize]
public class MediaController(IWebHostEnvironment env, ILogger<MediaController> logger) : ControllerBase
{
    private static readonly string[] AllowedExtensions = [".jpg", ".jpeg", ".png", ".gif", ".webp"];
    private const long MaxBytes = 5 * 1024 * 1024; // 5 MB

    [HttpPost("upload")]
    public async Task<IActionResult> Upload(IFormFile file, CancellationToken ct)
    {
        if (file is null || file.Length == 0)
            return BadRequest("No se recibió ningún archivo.");

        if (file.Length > MaxBytes)
            return BadRequest("El archivo supera el límite de 5 MB.");

        string ext = Path.GetExtension(file.FileName).ToLowerInvariant();
        if (!AllowedExtensions.Contains(ext))
            return BadRequest("Tipo de archivo no permitido. Use jpg, png, gif o webp.");

        string wwwroot   = env.WebRootPath ?? Path.Combine(Directory.GetCurrentDirectory(), "wwwroot");
        string uploadDir = Path.Combine(wwwroot, "uploads", "templates");
        Directory.CreateDirectory(uploadDir);

        string fileName = $"{Guid.NewGuid()}{ext}";
        string filePath = Path.Combine(uploadDir, fileName);

        await using var stream = System.IO.File.Create(filePath);
        await file.CopyToAsync(stream, ct);

        string url = $"{Request.Scheme}://{Request.Host}/uploads/templates/{fileName}";
        logger.LogInformation("Imagen subida: {FileName}", fileName);

        return Ok(new { url });
    }
}
