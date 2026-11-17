using System.Security.Cryptography;
using System.Text;

namespace Pandora.API.Services;

/// <summary>
/// Cifra/descifra los artefactos de respaldo con AES-256-GCM (cifrado
/// autenticado). Es <b>opt-in</b>: solo se activa si hay una llave configurada
/// en <c>Backup:EncryptionKey</c> (en Railway: variable <c>Backup__EncryptionKey</c>).
/// Sin llave, <see cref="IsEnabled"/> es false y todo se comporta como antes
/// (respaldos en texto plano) — retrocompatible.
///
/// Objetivo de seguridad: que el archivo de respaldo (.bak/.sql) sea ilegible
/// fuera del servidor. Aunque la cuenta de Google Drive o el correo donde se
/// guardan los respaldos se vea comprometida, sin la llave el contenido de la
/// base (usuarios, hashes, datos) no se puede leer.
///
/// Formato del archivo cifrado:
///   [MAGIC 8 bytes "PNDAESG1"][nonce 12][tag 16][ciphertext...]
/// La llave efectiva de 32 bytes se deriva con SHA-256 de la cadena configurada,
/// así que la llave puede ser cualquier texto largo (passphrase).
/// </summary>
public class BackupEncryptionService(IConfiguration config, ILogger<BackupEncryptionService> logger)
{
    private static readonly byte[] Magic = "PNDAESG1"u8.ToArray();
    private const int NonceSize = 12; // AES-GCM estándar
    private const int TagSize   = 16;

    /// <summary>true si hay una llave configurada (cifrado activo).</summary>
    public bool IsEnabled => !string.IsNullOrWhiteSpace(config["Backup:EncryptionKey"]);

    private byte[] DeriveKey()
    {
        var raw = config["Backup:EncryptionKey"]
            ?? throw new InvalidOperationException("No hay llave de cifrado de respaldos configurada (Backup:EncryptionKey).");
        return SHA256.HashData(Encoding.UTF8.GetBytes(raw)); // 32 bytes
    }

    /// <summary>¿El contenido empieza con la cabecera mágica de un respaldo cifrado por Pandora?</summary>
    public static bool IsEncrypted(byte[] data) =>
        data.Length >= Magic.Length && data.AsSpan(0, Magic.Length).SequenceEqual(Magic);

    /// <summary>
    /// Cifra el contenido y agrega el sufijo <c>.enc</c> al nombre. Si el cifrado
    /// no está activo, devuelve el contenido y el nombre sin cambios.
    /// </summary>
    public (byte[] Bytes, string FileName) Encrypt(byte[] plain, string fileName)
    {
        if (!IsEnabled) return (plain, fileName);

        var key   = DeriveKey();
        var nonce = RandomNumberGenerator.GetBytes(NonceSize);
        var cipher = new byte[plain.Length];
        var tag    = new byte[TagSize];

        using (var aes = new AesGcm(key, TagSize))
            aes.Encrypt(nonce, plain, cipher, tag);

        var outBytes = new byte[Magic.Length + NonceSize + TagSize + cipher.Length];
        var pos = 0;
        Buffer.BlockCopy(Magic, 0, outBytes, pos, Magic.Length); pos += Magic.Length;
        Buffer.BlockCopy(nonce, 0, outBytes, pos, NonceSize);    pos += NonceSize;
        Buffer.BlockCopy(tag,   0, outBytes, pos, TagSize);      pos += TagSize;
        Buffer.BlockCopy(cipher, 0, outBytes, pos, cipher.Length);

        logger.LogInformation("Respaldo cifrado (AES-256-GCM): {File} ({Kb} KB)", fileName, plain.Length / 1024);
        return (outBytes, fileName + ".enc");
    }

    /// <summary>
    /// Descifra el contenido si trae la cabecera de Pandora. Si no está cifrado
    /// (respaldo en texto plano o de una versión anterior), lo devuelve tal cual
    /// — así los respaldos viejos se siguen pudiendo restaurar.
    /// </summary>
    public byte[] Decrypt(byte[] data)
    {
        if (!IsEncrypted(data)) return data; // passthrough: no cifrado

        if (!IsEnabled)
            throw new InvalidOperationException(
                "Este respaldo está cifrado pero no hay llave de cifrado configurada en el servidor (Backup:EncryptionKey). " +
                "Configura la misma llave con la que se generó para poder restaurarlo.");

        var key = DeriveKey();
        var pos = Magic.Length;
        var nonce = data.AsSpan(pos, NonceSize).ToArray();          pos += NonceSize;
        var tag   = data.AsSpan(pos, TagSize).ToArray();            pos += TagSize;
        var cipher = data.AsSpan(pos).ToArray();
        var plain = new byte[cipher.Length];

        try
        {
            using var aes = new AesGcm(key, TagSize);
            aes.Decrypt(nonce, cipher, tag, plain);
        }
        catch (CryptographicException ex)
        {
            throw new InvalidOperationException(
                "No se pudo descifrar el respaldo: la llave de cifrado no coincide con la que se usó para generarlo, " +
                "o el archivo está corrupto.", ex);
        }

        return plain;
    }
}
