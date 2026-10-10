using System.Buffers.Binary;
using System.Security.Cryptography;
using System.Text;

namespace PivotWright.Internal.Crypto;

/// <summary>
/// The CryptKey.bin file of a data model ([MS-XLDM] 2.1.2.4) and the encryption it is used for:
/// connection strings and passwords in the model metadata are 3DES-CBC encrypted with this key.
/// The key travels with the file (as a SIMPLEBLOB wrapped with an exponent-of-one RSA key, i.e.
/// in the clear), so it provides obfuscation rather than secrecy.
/// </summary>
internal sealed class ModelCryptKey
{
    private static readonly Guid Magic = new("5d21bc98-8d2d-4ee6-a8e5-d038aac94441");
    private readonly byte[] _key;

    private ModelCryptKey(byte[] key) => _key = key;

    public static ModelCryptKey Create()
    {
        var key = new byte[24];
        RandomNumberGenerator.Fill(key);
        return new ModelCryptKey(key);
    }

    /// <summary>Reads the session key back from a CryptKey.bin file.</summary>
    public static ModelCryptKey FromFile(ReadOnlySpan<byte> file)
    {
        if (file.Length != 144 || !new Guid(file[..16]).Equals(Magic))
            throw new InvalidDataException("Not a CryptKey.bin file.");
        var block = file.Slice(60, 64).ToArray();
        Array.Reverse(block);
        if (block[0] != 0 || block[1] != 2) throw new InvalidDataException("Unexpected key block.");
        int sep = Array.IndexOf(block, (byte)0, 2);
        return new ModelCryptKey(block[(sep + 1)..]);
    }

    /// <summary>Serializes CryptKey.bin (144 bytes).</summary>
    public byte[] ToFile()
    {
        var file = new byte[144];
        var s = file.AsSpan();
        Magic.TryWriteBytes(s);
        BinaryPrimitives.WriteUInt32LittleEndian(s[16..], 4);           // m_dwVersion
        BinaryPrimitives.WriteUInt32LittleEndian(s[20..], 48);          // m_dwHeaderSize (incl. padding)
        BinaryPrimitives.WriteUInt32LittleEndian(s[24..], 80);          // m_dwKeyDataSize
        BinaryPrimitives.WriteUInt32LittleEndian(s[28..], 16);          // m_dwTrailerSize
        BinaryPrimitives.WriteUInt32LittleEndian(s[32..], 1);           // m_dwProvider
        BinaryPrimitives.WriteUInt32LittleEndian(s[36..], 7);           // m_dwAlgorithm
        BinaryPrimitives.WriteUInt32LittleEndian(s[40..], 0xFFFFFFFF);  // m_dwFlags

        // SIMPLEBLOB: BLOBHEADER { SIMPLEBLOB, CUR_BLOB_VERSION, 0, CALG_3DES }, ALG_ID CALG_RSA_KEYX,
        // then the PKCS#1 v1.5 type-2 block holding the session key, little-endian.
        var blob = s[48..128];
        blob[0] = 0x01; blob[1] = 0x02;
        BinaryPrimitives.WriteUInt32LittleEndian(blob[4..], 0x6603);
        BinaryPrimitives.WriteUInt32LittleEndian(blob[8..], 0xA400);
        var block = new byte[64];
        block[0] = 0x00; block[1] = 0x02;
        var pad = block.AsSpan(2, 64 - 3 - _key.Length);
        RandomNumberGenerator.Fill(pad);
        for (int i = 0; i < pad.Length; i++) if (pad[i] == 0) pad[i] = 0x5A;
        block[64 - _key.Length - 1] = 0x00;
        _key.CopyTo(block, 64 - _key.Length);
        Array.Reverse(block);
        block.CopyTo(blob[12..]);
        Magic.TryWriteBytes(s[128..]);
        return file;
    }

    /// <summary>
    /// Encrypts a string the way Analysis Services stores secrets: an 8-byte header, then
    /// 3DES-CBC (zero IV, PKCS#7 padding) over 8 random salt bytes followed by the NUL-terminated
    /// UTF-16 text. The result is Base64, wrapped every 76 characters as in Excel-written files.
    /// </summary>
    public string Encrypt(string plaintext, uint headerFlags)
    {
        var plain = new byte[8 + (plaintext.Length + 1) * 2];
        RandomNumberGenerator.Fill(plain.AsSpan(0, 8));
        Encoding.Unicode.GetBytes(plaintext, 0, plaintext.Length, plain, 8);

        using var des = TripleDES.Create();
        des.Key = _key;
        byte[] cipher = des.EncryptCbc(plain, new byte[8], PaddingMode.PKCS7);

        var blob = new byte[8 + cipher.Length];
        BinaryPrimitives.WriteUInt32LittleEndian(blob, 8);
        BinaryPrimitives.WriteUInt32LittleEndian(blob.AsSpan(4), headerFlags);
        cipher.CopyTo(blob, 8);

        string b64 = Convert.ToBase64String(blob);
        var sb = new StringBuilder();
        for (int i = 0; i < b64.Length; i += 76)
            sb.Append(b64, i, Math.Min(76, b64.Length - i)).Append(" \n");
        return sb.ToString();
    }

    /// <summary>Header flags Excel uses for an encrypted connection string.</summary>
    public const uint ConnectionStringFlags = 0x08000001;
    /// <summary>Header flags Excel uses for the database impersonation password.</summary>
    public const uint DatabasePasswordFlags = 0x18000001;
    /// <summary>Header flags Excel uses for an empty data source password.</summary>
    public const uint EmptyPasswordFlags = 0;

    /// <summary>Test hook: decrypts a value produced by <see cref="Encrypt"/>.</summary>
    public string Decrypt(string base64)
    {
        byte[] blob = Convert.FromBase64String(base64.Replace(" ", "").Replace("\n", ""));
        using var des = TripleDES.Create();
        des.Key = _key;
        byte[] plain = des.DecryptCbc(blob.AsSpan(8), new byte[8], PaddingMode.PKCS7);
        return Encoding.Unicode.GetString(plain, 8, plain.Length - 10);
    }
}
