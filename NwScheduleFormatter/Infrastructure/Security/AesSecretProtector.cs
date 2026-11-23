using System.IO;
using System.Security.Cryptography;
using System.Text;

public static class AesSecretProtector
{
    private const int KeySize = 32;   // AES-256
    private const int NonceSize = 12; // GCM pattern
    private const int TagSize = 16;   // authentication

    private static readonly byte[] KeyInfo =
        Encoding.UTF8.GetBytes("NwScheduleFormatter.WhatsAppConfig.AesKey.v2");

    private static readonly Lazy<byte[]> Key =
        new(() => HKDF.DeriveKey(HashAlgorithmName.SHA256,
                                 LoadOrCreateInstallGuid(),
                                 KeySize,
                                 salt: null,
                                 info: KeyInfo));

    /// <summary>Cryptography. Result format (Base64): nonce | tag | encrypted.</summary>
    public static string? Encrypt(string? plainText)
    {
        if (string.IsNullOrEmpty(plainText))
            return null;

        var nonce = RandomNumberGenerator.GetBytes(NonceSize);
        var data = Encoding.UTF8.GetBytes(plainText);
        var cipher = new byte[data.Length];
        var tag = new byte[TagSize];

        using var aes = new AesGcm(Key.Value, TagSize);
        aes.Encrypt(nonce, data, cipher, tag);

        var payload = new byte[NonceSize + TagSize + cipher.Length];
        nonce.CopyTo(payload, 0);
        tag.CopyTo(payload, NonceSize);
        cipher.CopyTo(payload, NonceSize + TagSize);

        return Convert.ToBase64String(payload);
    }

    public static string? Decrypt(string? cipherText)
    {
        if (string.IsNullOrEmpty(cipherText))
            return null;

        var payload = Convert.FromBase64String(cipherText);
        var nonce = payload.AsSpan(0, NonceSize);
        var tag = payload.AsSpan(NonceSize, TagSize);
        var cipher = payload.AsSpan(NonceSize + TagSize);
        var plain = new byte[cipher.Length];

        using var aes = new AesGcm(Key.Value, TagSize);
        aes.Decrypt(nonce, cipher, tag, plain); // throws CryptographicException if corrupted

        return Encoding.UTF8.GetString(plain);
    }

    /// <summary>
    /// Create a GUID on the first execution and stores protected by DPAPI.
    /// In the next ones, read and decrypt the same GUID.
    /// </summary>
    private static byte[] LoadOrCreateInstallGuid()
    {
        var dir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "NwScheduleFormatter");
        Directory.CreateDirectory(dir);

        var path = Path.Combine(dir, "install.key");

        if (File.Exists(path))
        {
            return ProtectedData.Unprotect(
                File.ReadAllBytes(path), null, DataProtectionScope.CurrentUser);
        }

        var guid = Guid.NewGuid().ToByteArray();
        File.WriteAllBytes(path,
            ProtectedData.Protect(guid, null, DataProtectionScope.CurrentUser));
        return guid;
    }
}