using System.Security.Cryptography;
using System.Text;
using MailKit.Net.Smtp;
using MailKit.Security;
using Microsoft.Data.SqlClient;
using MimeKit;

namespace Pandora.API.Services;

/// <summary>
/// Helper estático para envío de correos.
/// Lee la configuración SMTP primero desde dbo.SystemSettings (BD) y hace
/// fallback a appsettings.json. La contraseña se almacena cifrada con AES-256.
/// </summary>
public static class SmtpHelper
{
    // Clave de cifrado derivada de una frase fija — en producción puede venir de env var
    private static readonly byte[] EncKey =
        SHA256.HashData(Encoding.UTF8.GetBytes("Pandora_SMTP_Enc_Key_2024!"));

    // ── Descifrar contraseña ─────────────────────────────────────────────────
    public static string? Decrypt(string? encrypted)
    {
        if (string.IsNullOrWhiteSpace(encrypted)) return null;
        try
        {
            var raw = Convert.FromBase64String(encrypted);
            using var aes = Aes.Create();
            aes.Key = EncKey;
            var iv = new byte[16];
            Buffer.BlockCopy(raw, 0, iv, 0, 16);
            aes.IV = iv;
            using var dec = aes.CreateDecryptor();
            var cipher = new byte[raw.Length - 16];
            Buffer.BlockCopy(raw, 16, cipher, 0, cipher.Length);
            return Encoding.UTF8.GetString(dec.TransformFinalBlock(cipher, 0, cipher.Length));
        }
        catch { return null; }
    }

    // ── Leer configuración SMTP desde BD con fallback a appsettings ──────────
    public static async Task<SmtpConfig> LoadAsync(string connStr, IConfiguration config)
    {
        var smtp = config.GetSection("SmtpSettings");
        var d    = new Dictionary<string, string?>();

        try
        {
            await using var conn = new SqlConnection(connStr);
            await conn.OpenAsync();
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = """
                IF OBJECT_ID('dbo.SystemSettings') IS NOT NULL
                    SELECT SettingKey, SettingValue FROM dbo.SystemSettings
                    WHERE SettingKey IN (
                        'smtp_host','smtp_port','smtp_from_email',
                        'smtp_password','smtp_from_name','smtp_use_ssl')
                """;
            await using var r = await cmd.ExecuteReaderAsync();
            while (await r.ReadAsync())
                d[r.GetString(0)] = r.IsDBNull(1) ? null : r.GetString(1);
        }
        catch { /* fallback a appsettings */ }

        var host     = d.GetValueOrDefault("smtp_host")       ?? smtp["Host"]      ?? "";
        var port     = int.TryParse(d.GetValueOrDefault("smtp_port") ?? smtp["Port"], out var p) ? p : 587;
        var from     = d.GetValueOrDefault("smtp_from_email") ?? smtp["FromEmail"] ?? smtp["Username"] ?? "";
        var encPass  = d.GetValueOrDefault("smtp_password");
        var pass     = (encPass != null ? Decrypt(encPass) : null) ?? smtp["Password"] ?? "";
        var fromName = d.GetValueOrDefault("smtp_from_name")  ?? smtp["FromName"]  ?? "Pandora";
        var useSsl   = (d.GetValueOrDefault("smtp_use_ssl")   ?? smtp["UseSsl"]    ?? "true") == "true";

        return new SmtpConfig(host, port, from, pass, fromName, useSsl);
    }

    // ── Enviar un correo HTML ────────────────────────────────────────────────
    /// <returns>null si OK, mensaje de error si falló</returns>
    public static async Task<string?> SendAsync(
        SmtpConfig cfg,
        string toEmail, string toName,
        string subject, string htmlBody,
        IReadOnlyList<EmailAttachment>? attachments = null)
    {
        if (string.IsNullOrWhiteSpace(cfg.Host) ||
            string.IsNullOrWhiteSpace(cfg.From) ||
            string.IsNullOrWhiteSpace(cfg.Password))
            return "SMTP no configurado (host, usuario o contraseña vacíos).";

        try
        {
            using var client = new SmtpClient();
            await client.ConnectAsync(cfg.Host, cfg.Port, SecureSocketOptions.StartTls);
            await client.AuthenticateAsync(cfg.From, cfg.Password);

            var msg = new MimeMessage();
            msg.From.Add(new MailboxAddress(cfg.FromName, cfg.From));
            msg.To.Add(new MailboxAddress(toName, toEmail));
            msg.Subject = subject;

            var builder = new BodyBuilder { HtmlBody = htmlBody };
            if (attachments is { Count: > 0 })
                foreach (var att in attachments)
                    builder.Attachments.Add(att.FileName, att.Bytes,
                        ContentType.Parse(att.ContentType));

            msg.Body = builder.ToMessageBody();

            await client.SendAsync(msg);
            await client.DisconnectAsync(true);
            return null;
        }
        catch (Exception ex)
        {
            return ex.Message;
        }
    }
}

public record EmailAttachment(string FileName, byte[] Bytes, string ContentType = "application/octet-stream");

public record SmtpConfig(
    string Host,
    int    Port,
    string From,
    string Password,
    string FromName,
    bool   UseSsl = true
);
