using System.Buffers.Binary;

namespace PivotWright.Internal.Storage;

/// <summary>
/// Append-only buffer for fixed-size column values. Starts in memory and spills to a temporary
/// file (deleted on close) once it grows past a threshold, so very large tables do not need to
/// fit in RAM. After writing, the content is read back sequentially, segment by segment.
/// </summary>
internal sealed class Spool : IDisposable
{
    private readonly long _threshold;
    private readonly string? _tempDirectory;
    private readonly byte[] _buffer = new byte[1 << 16];
    private int _count;
    private Stream _stream = new MemoryStream();
    private bool _spilled;
    private bool _reading;

    public Spool(long threshold, string? tempDirectory)
    {
        _threshold = threshold;
        _tempDirectory = tempDirectory;
    }

    public long Length { get; private set; }

    public void WriteUInt32(uint v)
    {
        if (_count + 4 > _buffer.Length) FlushBuffer();
        BinaryPrimitives.WriteUInt32LittleEndian(_buffer.AsSpan(_count), v);
        _count += 4;
        Length += 4;
    }

    public void WriteInt64(long v)
    {
        if (_count + 8 > _buffer.Length) FlushBuffer();
        BinaryPrimitives.WriteInt64LittleEndian(_buffer.AsSpan(_count), v);
        _count += 8;
        Length += 8;
    }

    public void WriteDouble(double v) => WriteInt64(BitConverter.DoubleToInt64Bits(v));

    private void FlushBuffer()
    {
        if (_count == 0) return;
        if (_reading) throw new InvalidOperationException("Spool is already being read.");
        if (!_spilled && _stream.Length + _count > _threshold)
        {
            string dir = _tempDirectory ?? Path.GetTempPath();
            string path = Path.Combine(dir, "ppw-" + Guid.NewGuid().ToString("N") + ".tmp");
            var fs = new FileStream(path, FileMode.CreateNew, FileAccess.ReadWrite, FileShare.None,
                1 << 16, FileOptions.DeleteOnClose | FileOptions.SequentialScan);
            _stream.Position = 0;
            _stream.CopyTo(fs);
            _stream.Dispose();
            _stream = fs;
            _spilled = true;
        }
        _stream.Write(_buffer, 0, _count);
        _count = 0;
    }

    /// <summary>Returns a reader positioned at the start of the spooled data.</summary>
    public SpoolReader OpenReader()
    {
        FlushBuffer();
        _reading = true;
        _stream.Flush();
        _stream.Position = 0;
        return new SpoolReader(_stream);
    }

    public void Dispose() => _stream.Dispose();
}

internal sealed class SpoolReader
{
    private readonly Stream _stream;
    private byte[] _bytes = Array.Empty<byte>();

    public SpoolReader(Stream stream) => _stream = stream;

    private ReadOnlySpan<byte> ReadBytes(int count)
    {
        if (_bytes.Length < count) _bytes = new byte[count];
        _stream.ReadExactly(_bytes, 0, count);
        return _bytes.AsSpan(0, count);
    }

    public void ReadUInt32(Span<uint> dest)
    {
        var src = ReadBytes(dest.Length * 4);
        for (int i = 0; i < dest.Length; i++) dest[i] = BinaryPrimitives.ReadUInt32LittleEndian(src[(i * 4)..]);
    }

    public void ReadInt64(Span<long> dest)
    {
        var src = ReadBytes(dest.Length * 8);
        for (int i = 0; i < dest.Length; i++) dest[i] = BinaryPrimitives.ReadInt64LittleEndian(src[(i * 8)..]);
    }
}

/// <summary>Growable bit set used to remember which rows of a numeric column are null.</summary>
internal sealed class NullMask
{
    private ulong[] _bits = new ulong[16];

    public bool Any { get; private set; }

    public void Set(long row)
    {
        long word = row >> 6;
        if (word >= _bits.Length)
        {
            long size = _bits.Length;
            while (size <= word) size *= 2;
            Array.Resize(ref _bits, checked((int)size));
        }
        _bits[word] |= 1UL << (int)(row & 63);
        Any = true;
    }

    public bool IsNull(long row)
    {
        long word = row >> 6;
        return word < _bits.Length && (_bits[word] & (1UL << (int)(row & 63))) != 0;
    }
}
