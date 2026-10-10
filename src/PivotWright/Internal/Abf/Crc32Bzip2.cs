namespace PivotWright.Internal.Abf;

/// <summary>
/// CRC-32/BZIP2 (polynomial 0x04C11DB7, MSB-first, init and xor-out 0xFFFFFFFF).
/// This is the checksum the ABF container appends to every stored file ([MS-XLDM] 2.1.2.2.1.1).
/// </summary>
internal struct Crc32Bzip2
{
    private static readonly uint[] Table = BuildTable();
    private uint _state;

    public Crc32Bzip2() => _state = 0xFFFFFFFFu;

    public static Crc32Bzip2 Create() => new();

    private static uint[] BuildTable()
    {
        var t = new uint[256];
        for (uint i = 0; i < 256; i++)
        {
            uint c = i << 24;
            for (int k = 0; k < 8; k++)
                c = (c & 0x80000000u) != 0 ? (c << 1) ^ 0x04C11DB7u : c << 1;
            t[i] = c;
        }
        return t;
    }

    public void Append(ReadOnlySpan<byte> data)
    {
        uint c = _state;
        foreach (byte b in data)
            c = Table[((c >> 24) ^ b) & 0xFF] ^ (c << 8);
        _state = c;
    }

    public readonly uint Value => _state ^ 0xFFFFFFFFu;

    public static uint Compute(ReadOnlySpan<byte> data)
    {
        var c = new Crc32Bzip2();
        c.Append(data);
        return c.Value;
    }
}
