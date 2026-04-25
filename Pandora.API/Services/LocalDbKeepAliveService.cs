using System.Diagnostics;
using System.Text.RegularExpressions;
using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;

namespace Pandora.API.Services;

/// <summary>
/// Mantiene viva la instancia de LocalDB enviando un ping cada 3 minutos.
/// Solo se registra en Development para no consumir recursos en Docker.
/// </summary>
public sealed class LocalDbKeepAliveService(
    IConfiguration config,
    ILogger<LocalDbKeepAliveService> logger) : BackgroundService
{
    private static readonly TimeSpan Interval = TimeSpan.FromMinutes(3);

    protected override async Task ExecuteAsync(CancellationToken ct)
    {
        logger.LogInformation("[LocalDB] Keep-alive activo (ping cada {m} min).", Interval.TotalMinutes);
        await PingAsync(ct);

        while (!ct.IsCancellationRequested)
        {
            try   { await Task.Delay(Interval, ct); }
            catch (OperationCanceledException) { break; }

            if (ct.IsCancellationRequested) break;
            await PingAsync(ct);
        }
    }

    private async Task PingAsync(CancellationToken ct)
    {
        string? connStr = config.GetConnectionString("PandoraDb");
        if (string.IsNullOrWhiteSpace(connStr)) return;

        try
        {
            await using var conn = new SqlConnection(connStr);
            await conn.OpenAsync(ct);
            await using var cmd = conn.CreateCommand();
            cmd.CommandText    = "SELECT 1";
            cmd.CommandTimeout = 5;
            await cmd.ExecuteScalarAsync(ct);
            logger.LogDebug("[LocalDB] Ping OK.");
        }
        catch (OperationCanceledException) { }
        catch (SqlException ex) when (ex.Number is -1 or 2 or 53)
        {
            logger.LogWarning("[LocalDB] Pipe expirado (error {n}). Re-iniciando LocalDB...", ex.Number);
            string newConn = ResolveAndRestart(connStr);
            if (newConn != connStr)
            {
                ((IConfigurationRoot)config)["ConnectionStrings:PandoraDb"] = newConn;
                logger.LogInformation("[LocalDB] Connection string actualizado con nuevo pipe.");
            }
        }
        catch (Exception ex)
        {
            logger.LogWarning("[LocalDB] Ping fallido: {msg}", ex.Message);
        }
    }

    private static string ResolveAndRestart(string connStr)
    {
        try
        {
            const string instance = "MSSQLLocalDB";
            Run($"start \"{instance}\"");
            Thread.Sleep(2_000);
            string info = Run($"info \"{instance}\"");

            foreach (string line in info.Split('\n'))
            {
                string t = line.Trim();
                int idx  = t.IndexOf(@"\\.\pipe\", StringComparison.OrdinalIgnoreCase);
                if (idx < 0) idx = t.IndexOf("np:", StringComparison.OrdinalIgnoreCase);
                if (idx < 0) continue;

                string pipe = t[idx..].Trim();
                if (pipe.StartsWith("np:", StringComparison.OrdinalIgnoreCase)) pipe = pipe[3..];

                const string pattern = @"\\\\\.\\pipe\\LOCALDB#[A-F0-9]+\\tsql\\query";
                return Regex.Replace(connStr, pattern, pipe.Replace("\\", "\\\\"), RegexOptions.IgnoreCase);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"[LocalDB] ResolveAndRestart error: {ex.Message}");
        }
        return connStr;
    }

    private static string Run(string args)
    {
        var psi = new ProcessStartInfo("sqllocaldb", args)
        {
            RedirectStandardOutput = true,
            RedirectStandardError  = true,
            UseShellExecute        = false,
            CreateNoWindow         = true
        };
        using var p = Process.Start(psi)!;
        string output = p.StandardOutput.ReadToEnd();
        p.WaitForExit(8_000);
        return output;
    }
}
