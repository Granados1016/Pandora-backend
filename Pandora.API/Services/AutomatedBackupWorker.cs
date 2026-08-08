namespace Pandora.API.Services;

/// <summary>
/// Corre el backup automático una vez al día (hora fija) si está habilitado
/// en dbo.SystemSettings (backup_auto_enabled = 'true'). El "Ejecutar ahora"
/// del panel Admin llama a <see cref="BackupOrchestrator"/> directamente y no
/// pasa por este worker.
/// </summary>
public class AutomatedBackupWorker(
    IServiceScopeFactory scopeFactory,
    ILogger<AutomatedBackupWorker> logger) : BackgroundService
{
    // Hora local (México) a la que corre el backup diario.
    private static readonly TimeSpan RunAtLocalTime = new(3, 0, 0);

    private static readonly TimeZoneInfo MxTz =
        TimeZoneInfo.FindSystemTimeZoneById(
            OperatingSystem.IsWindows() ? "Central Standard Time" : "America/Mexico_City");

    protected override async Task ExecuteAsync(CancellationToken ct)
    {
        // Esperar a que la app y la BD estén listas.
        await Task.Delay(TimeSpan.FromSeconds(60), ct);

        while (!ct.IsCancellationRequested)
        {
            try
            {
                await Task.Delay(TimeUntilNextRun(), ct);
                await RunIfEnabledAsync(ct);
            }
            catch (OperationCanceledException)
            {
                break;
            }
            catch (Exception ex)
            {
                logger.LogError(ex, "AutomatedBackupWorker error inesperado");
                // Evita un loop apretado si algo falla antes del Delay.
                await Task.Delay(TimeSpan.FromMinutes(30), ct);
            }
        }
    }

    private static TimeSpan TimeUntilNextRun()
    {
        var nowMx = TimeZoneInfo.ConvertTime(DateTimeOffset.Now, MxTz);
        var todayRun = nowMx.Date + RunAtLocalTime;
        var nextRun = nowMx.TimeOfDay >= RunAtLocalTime ? todayRun.AddDays(1) : todayRun;
        return nextRun - nowMx.DateTime;
    }

    private async Task RunIfEnabledAsync(CancellationToken ct)
    {
        using var scope = scopeFactory.CreateScope();
        var orchestrator = scope.ServiceProvider.GetRequiredService<BackupOrchestrator>();

        var settings = await orchestrator.LoadSettingsAsync(ct);
        if (settings.GetValueOrDefault("backup_auto_enabled") != "true")
        {
            logger.LogDebug("AutomatedBackupWorker: backup automático deshabilitado, se omite.");
            return;
        }

        logger.LogInformation("AutomatedBackupWorker: iniciando backup diario programado.");
        var summary = await orchestrator.RunAsync(ct);
        logger.LogInformation(
            "AutomatedBackupWorker: backup completado — {File} ({Method}, {Kb} KB), correo a {Emails}, drive={Drive}",
            summary.FileName, summary.Method, summary.SizeBytes / 1024,
            summary.EmailedTo.Count > 0 ? string.Join(", ", summary.EmailedTo) : "(ninguno)",
            summary.DriveUploaded);
    }
}
