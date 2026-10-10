using System.Buffers.Binary;
using System.Globalization;
using System.Text;

namespace PivotWright.Internal.Abf;

/// <summary>
/// Writes the Spreadsheet Data Model stream (an Analysis Services backup, "ABF") described in
/// [MS-XLDM] 2.1: a one-page header, the files section, the backup log and the virtual directory.
/// Files are written sequentially; every file except the partition list and the log is framed
/// in 4 KiB chunks (u16 uncompressed size, u16 stored size) and followed by a CRC-32/BZIP2.
/// Chunks whose stored size equals their uncompressed size are stored verbatim, exactly as
/// Excel does for chunks that do not compress.
/// </summary>
internal sealed class AbfWriter
{
    public const int PageSize = 4096;
    public const int ChunkSize = 4096;
    private const string Signature = "STREAM_STORAGE_SIGNATURE_)!@#$%^&*(";

    private readonly Stream _stream;
    private readonly long _origin;
    private readonly List<VirtualDirectoryEntry> _entries = new();
    private readonly IChunkCompressor? _compressor;
    private bool _fileOpen;
    private bool _finished;

    public AbfWriter(Stream stream, IChunkCompressor? compressor = null)
    {
        if (!stream.CanSeek || !stream.CanWrite)
            throw new ArgumentException("The ABF stream must be writable and seekable.", nameof(stream));
        _stream = stream;
        _compressor = compressor;
        _origin = stream.Position;
        // Header page placeholder; the real header is written by Finish once offsets are known.
        _stream.Write(new byte[PageSize]);
    }

    public IReadOnlyList<VirtualDirectoryEntry> Entries => _entries;

    private long Position => _stream.Position - _origin;

    /// <summary>Writes a non-chunked file (the PARTITIONS list or the LOG): content followed by its CRC.</summary>
    public void WriteRawFile(string storageName, ReadOnlySpan<byte> content)
    {
        EnsureWritable();
        long offset = Position;
        long created = FileTime.Now();
        _stream.Write(content);
        Span<byte> crc = stackalloc byte[4];
        BinaryPrimitives.WriteUInt32LittleEndian(crc, Crc32Bzip2.Compute(content));
        _stream.Write(crc);
        _entries.Add(new VirtualDirectoryEntry(storageName, offset, content.Length + 4, created, FileTime.Now()));
    }

    /// <summary>Opens a chunked file. Dispose the returned stream to finish the file.</summary>
    public ChunkedFileStream CreateFile(string storageName)
    {
        EnsureWritable();
        _fileOpen = true;
        return new ChunkedFileStream(this, storageName, Position, FileTime.Now());
    }

    private void EnsureWritable()
    {
        if (_finished) throw new InvalidOperationException("The ABF stream has already been finished.");
        if (_fileOpen) throw new InvalidOperationException("Another file is still open; dispose it first.");
    }

    internal void WriteChunk(ReadOnlySpan<byte> chunk, ref Crc32Bzip2 crc, Span<byte> scratch)
    {
        Span<byte> header = stackalloc byte[4];
        ReadOnlySpan<byte> payload = chunk;
        if (_compressor is not null)
        {
            int n = _compressor.TryCompress(chunk, scratch);
            if (n > 0 && n < chunk.Length) payload = scratch[..n];
        }
        BinaryPrimitives.WriteUInt16LittleEndian(header, (ushort)chunk.Length);
        BinaryPrimitives.WriteUInt16LittleEndian(header[2..], (ushort)payload.Length);
        _stream.Write(header);
        _stream.Write(payload);
        crc.Append(header);
        crc.Append(payload);
    }

    internal void CloseFile(string storageName, long offset, long created, uint crc)
    {
        Span<byte> buf = stackalloc byte[4];
        BinaryPrimitives.WriteUInt32LittleEndian(buf, crc);
        _stream.Write(buf);
        _entries.Add(new VirtualDirectoryEntry(storageName, offset, Position - offset, created, FileTime.Now()));
        _fileOpen = false;
    }

    /// <summary>
    /// Writes the backup log (as the last file), the virtual directory and the header page.
    /// </summary>
    public void Finish(string backupLogXml, Guid objectId)
    {
        EnsureWritable();
        var log = new MemoryStream();
        log.Write(new byte[] { 0xFF, 0xFE });
        log.Write(Encoding.Unicode.GetBytes(backupLogXml));
        WriteRawFile("LOG", log.ToArray());

        PadToPage();
        long vdOffset = Position;
        byte[] vd = Encoding.Unicode.GetBytes(BuildVirtualDirectory());
        _stream.Write(vd);
        PadToPage();
        long end = _stream.Position;

        var header = new StringBuilder();
        header.Append("<BackupLog>");
        header.Append("<BackupRestoreSyncVersion>150</BackupRestoreSyncVersion>");
        header.Append("<Fault>false</Fault>");
        header.Append("<faultcode>43981</faultcode>");
        header.Append("<ErrorCode>true</ErrorCode>");
        header.Append("<EncryptionFlag>false</EncryptionFlag>");
        header.Append("<EncryptionKey>6</EncryptionKey>");
        header.Append("<ApplyCompression>true</ApplyCompression>");
        header.Append("<m_cbOffsetHeader>").Append(vdOffset.ToString(CultureInfo.InvariantCulture)).Append("</m_cbOffsetHeader>");
        header.Append("<DataSize>").Append(vd.Length.ToString(CultureInfo.InvariantCulture)).Append("</DataSize>");
        header.Append("<Files>").Append(_entries.Count.ToString(CultureInfo.InvariantCulture)).Append("</Files>");
        header.Append("<ObjectID>").Append(objectId.ToString("D").ToUpperInvariant()).Append("</ObjectID>");
        header.Append("<m_cbOffsetData>").Append(PageSize.ToString(CultureInfo.InvariantCulture)).Append("</m_cbOffsetData>");
        header.Append("</BackupLog>");

        var page = new byte[PageSize];
        page[0] = 0xFF; page[1] = 0xFE;
        int n = 2 + Encoding.Unicode.GetBytes(Signature, page.AsSpan(2));
        byte[] hx = Encoding.Unicode.GetBytes(header.ToString());
        if (n + hx.Length > PageSize) throw new InvalidOperationException("ABF header does not fit in one page.");
        hx.CopyTo(page, n);
        _stream.Position = _origin;
        _stream.Write(page);
        _stream.Position = end;
        _finished = true;
    }

    private void PadToPage()
    {
        long rem = Position % PageSize;
        if (rem != 0) _stream.Write(new byte[PageSize - rem]);
    }

    private string BuildVirtualDirectory()
    {
        var sb = new StringBuilder();
        sb.Append("<VirtualDirectory>");
        foreach (var e in _entries)
        {
            sb.Append("<BackupFile><Path>").Append(e.StorageName).Append("</Path>");
            sb.Append("<Size>").Append(e.Size.ToString(CultureInfo.InvariantCulture)).Append("</Size>");
            sb.Append("<m_cbOffsetHeader>").Append(e.Offset.ToString(CultureInfo.InvariantCulture)).Append("</m_cbOffsetHeader>");
            sb.Append("<Delete>false</Delete>");
            sb.Append("<CreatedTimestamp>").Append(e.Created.ToString(CultureInfo.InvariantCulture)).Append("</CreatedTimestamp>");
            sb.Append("<Access>").Append(e.Written.ToString(CultureInfo.InvariantCulture)).Append("</Access>");
            sb.Append("<LastWriteTime>").Append(e.Written.ToString(CultureInfo.InvariantCulture)).Append("</LastWriteTime>");
            sb.Append("</BackupFile>");
        }
        sb.Append("</VirtualDirectory>");
        return sb.ToString();
    }
}

internal sealed record VirtualDirectoryEntry(string StorageName, long Offset, long Size, long Created, long Written);

/// <summary>Optional per-chunk compressor (Xpress). Returns the compressed length, or 0 when not beneficial.</summary>
internal interface IChunkCompressor
{
    int TryCompress(ReadOnlySpan<byte> input, Span<byte> output);
}

/// <summary>Write-only stream that frames its content into ABF chunks.</summary>
internal sealed class ChunkedFileStream : Stream
{
    private readonly AbfWriter _owner;
    private readonly string _name;
    private readonly long _offset;
    private readonly long _created;
    private readonly byte[] _buffer = new byte[AbfWriter.ChunkSize];
    private readonly byte[] _scratch = new byte[AbfWriter.ChunkSize * 2 + 64];
    private int _count;
    private long _length;
    private Crc32Bzip2 _crc = Crc32Bzip2.Create();
    private bool _closed;

    internal ChunkedFileStream(AbfWriter owner, string name, long offset, long created)
    {
        _owner = owner; _name = name; _offset = offset; _created = created;
    }

    public string StorageName => _name;

    /// <summary>Total number of uncompressed bytes written.</summary>
    public long UncompressedLength => _length;

    public override bool CanRead => false;
    public override bool CanSeek => false;
    public override bool CanWrite => !_closed;
    public override long Length => _length;
    public override long Position { get => _length; set => throw new NotSupportedException(); }
    public override void Flush() { }
    public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();
    public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
    public override void SetLength(long value) => throw new NotSupportedException();

    public override void Write(byte[] buffer, int offset, int count) => Write(buffer.AsSpan(offset, count));

    public override void Write(ReadOnlySpan<byte> data)
    {
        if (_closed) throw new ObjectDisposedException(nameof(ChunkedFileStream));
        while (!data.IsEmpty)
        {
            int n = Math.Min(data.Length, _buffer.Length - _count);
            data[..n].CopyTo(_buffer.AsSpan(_count));
            _count += n;
            _length += n;
            data = data[n..];
            if (_count == _buffer.Length) FlushChunk();
        }
    }

    public override void WriteByte(byte value)
    {
        Span<byte> one = stackalloc byte[1];
        one[0] = value;
        Write(one);
    }

    private void FlushChunk()
    {
        if (_count == 0) return;
        _owner.WriteChunk(_buffer.AsSpan(0, _count), ref _crc, _scratch);
        _count = 0;
    }

    protected override void Dispose(bool disposing)
    {
        if (!_closed)
        {
            _closed = true;
            FlushChunk();
            _owner.CloseFile(_name, _offset, _created, _crc.Value);
        }
        base.Dispose(disposing);
    }
}

internal static class FileTime
{
    public static long Now() => DateTime.UtcNow.ToFileTimeUtc();
}
