using System.Buffers.Binary;
using System.Numerics;

namespace PivotWright.Internal.VertiPaq;

/// <summary>Reserved data identifiers used by the xVelocity engine.</summary>
internal static class DataIds
{
    /// <summary>Data identifier of BLANK/null.</summary>
    public const uint Null = 2;
    /// <summary>First data identifier assigned to a real value.</summary>
    public const uint First = 3;
}

/// <summary>Bit-packing helpers for XMRENoSplitCompressionInfo&lt;N&gt; ([MS-XLDM] 2.7.1).</summary>
internal static class BitPacking
{
    public static readonly int[] SupportedWidths = { 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 12, 16, 21, 32 };

    public static int ChooseWidth(uint maxDelta)
    {
        int need = maxDelta == 0 ? 1 : 32 - BitOperations.LeadingZeroCount(maxDelta);
        foreach (int w in SupportedWidths)
            if (w >= need) return w;
        return 32;
    }

    /// <summary>Number of 64-bit words used by a bit-packed run of <paramref name="count"/> values.</summary>
    public static long WordCount(long count, int width) => count / (64 / width) + 1;

    public static void Pack(ReadOnlySpan<uint> values, uint min, int width, Span<ulong> words)
    {
        int per = 64 / width;
        words.Clear();
        for (int i = 0; i < values.Length; i++)
        {
            ulong v = values[i] - min;
            words[i / per] |= v << (width * (i % per));
        }
    }
}

/// <summary>Result of encoding one column segment with hybrid RLE + bit packing.</summary>
internal sealed class EncodedSegment
{
    /// <summary>Serialized segment (cleared once written to the .idf file).</summary>
    public required byte[] Data { get; set; }
    public long Records { get; init; }
    public long SubRecords { get; init; }
    public int Width { get; init; }
    public uint Min { get; init; }
    public uint MinDataId { get; init; }
    public uint MaxDataId { get; init; }
    public bool HasNulls { get; init; }
    public long RleRuns { get; init; }
    public long BitPackedRuns { get; init; }
    public long BookmarkBits { get; init; }
    public long StorageAllocSize { get; init; }
    public long StorageUsedSize { get; init; }
    public bool IsRowNumber { get; init; }
}

/// <summary>
/// Encodes a segment of data identifiers using XMHybridRLECompressionInfo&lt;XMRENoSplitCompressionInfo&lt;N&gt;&gt;
/// ([MS-XLDM] 2.3.1.1.2 and 2.7.3): a primary segment of RLE entries followed by a bit-packed sub-segment.
/// </summary>
internal static class HybridRleEncoder
{
    /// <summary>Runs at least this long are stored as literal RLE entries.</summary>
    public const int DefaultRleThreshold = 64;

    public static EncodedSegment Encode(ReadOnlySpan<uint> ids, int rleThreshold = DefaultRleThreshold)
    {
        int n = ids.Length;
        uint minNonNull = uint.MaxValue, max = 0;
        bool hasNulls = false;
        for (int i = 0; i < n; i++)
        {
            uint id = ids[i];
            if (id == DataIds.Null) hasNulls = true;
            else if (id < minNonNull) minNonNull = id;
            if (id > max) max = id;
        }
        if (n == 0) { minNonNull = DataIds.First; max = DataIds.Null; }
        if (minNonNull == uint.MaxValue) minNonNull = DataIds.Null; // all-null segment
        uint min = hasNulls || n == 0 ? DataIds.Null : minNonNull;
        int width = BitPacking.ChooseWidth(max >= min ? max - min : 0);

        // Build entries: literal runs >= threshold, everything else bit-packed.
        var entries = new List<(uint value, uint repeat)>();
        var packed = new List<uint>(n);
        long rle = 0, bpRuns = 0;
        int pos = 0;
        int bpStart = -1;
        uint bpOffset = 0;
        void CloseBitPacked(int end)
        {
            if (bpStart < 0) return;
            int len = end - bpStart;
            entries.Add((0xFFFFFFFFu - bpOffset, (uint)len));
            bpOffset += (uint)len;
            bpRuns++;
            bpStart = -1;
        }
        while (pos < n)
        {
            uint v = ids[pos];
            int run = 1;
            while (pos + run < n && ids[pos + run] == v) run++;
            if (run >= rleThreshold)
            {
                CloseBitPacked(pos);
                entries.Add((v, (uint)run));
                rle++;
            }
            else
            {
                if (bpStart < 0) bpStart = pos;
                for (int k = 0; k < run; k++) packed.Add(v);
            }
            pos += run;
        }
        CloseBitPacked(n);

        long slots = 16;
        while (slots < entries.Count + 1) slots *= 2;
        long words = BitPacking.WordCount(packed.Count, width);

        var data = new byte[8 + slots * 8 + 8 + words * 8];
        var span = data.AsSpan();
        BinaryPrimitives.WriteUInt64LittleEndian(span, (ulong)slots);
        int o = 8;
        foreach (var (value, repeat) in entries)
        {
            BinaryPrimitives.WriteUInt32LittleEndian(span[o..], value);
            BinaryPrimitives.WriteUInt32LittleEndian(span[(o + 4)..], repeat);
            o += 8;
        }
        o = (int)(8 + slots * 8);
        BinaryPrimitives.WriteUInt64LittleEndian(span[o..], (ulong)words);
        o += 8;
        var w = new ulong[words];
        BitPacking.Pack(System.Runtime.InteropServices.CollectionsMarshal.AsSpan(packed), min, width, w);
        for (int i = 0; i < w.Length; i++, o += 8)
            BinaryPrimitives.WriteUInt64LittleEndian(span[o..], w[i]);

        return new EncodedSegment
        {
            Data = data,
            Records = n,
            SubRecords = packed.Count,
            Width = width,
            Min = min,
            MinDataId = minNonNull,
            MaxDataId = max,
            HasNulls = hasNulls,
            RleRuns = rle,
            BitPackedRuns = bpRuns,
            BookmarkBits = BookmarkBits(n),
            StorageAllocSize = slots * 2,
            StorageUsedSize = (entries.Count + 1) * 2,
        };
    }

    /// <summary>
    /// Encodes a segment of the implicit RowNumber column with XMHybridRLECompressionInfo&lt;XM123CompressionInfo&gt;:
    /// a single bit-packed entry and an empty sub-segment; values are implied by Min.
    /// </summary>
    public static EncodedSegment EncodeRowNumber(long firstRow, long rows)
    {
        const int slots = 16;
        var data = new byte[8 + slots * 8 + 8];
        BinaryPrimitives.WriteUInt64LittleEndian(data, slots);
        BinaryPrimitives.WriteUInt32LittleEndian(data.AsSpan(8), 0xFFFFFFFFu);
        BinaryPrimitives.WriteUInt32LittleEndian(data.AsSpan(12), (uint)rows);
        uint min = (uint)(DataIds.First + firstRow);
        return new EncodedSegment
        {
            Data = data,
            Records = rows,
            SubRecords = rows,
            Width = 0,
            Min = min,
            MinDataId = min,
            MaxDataId = (uint)(min + rows - 1),
            HasNulls = false,
            RleRuns = 0,
            BitPackedRuns = 0,
            BookmarkBits = 24,
            StorageAllocSize = slots * 2,
            StorageUsedSize = 2,
            IsRowNumber = true,
        };
    }

    public static long BookmarkBits(long rows)
    {
        if (rows <= 0) return 1;
        int lg = 63 - BitOperations.LeadingZeroCount((ulong)rows);
        return Math.Max(1, lg - 6);
    }
}

/// <summary>
/// Writes the 32-bit, non-hybrid columns of system tables (POS_TO_ID, ID_TO_POS):
/// each segment is a u64 word count followed by the bit-packed words.
/// </summary>
internal static class NoSplit32Writer
{
    public static void WriteSegment(Stream s, ReadOnlySpan<uint> values)
    {
        long words = BitPacking.WordCount(values.Length, 32);
        Span<byte> b8 = stackalloc byte[8];
        BinaryPrimitives.WriteUInt64LittleEndian(b8, (ulong)words);
        s.Write(b8);
        var buf = new byte[words * 8];
        for (int i = 0; i < values.Length; i++)
            BinaryPrimitives.WriteUInt32LittleEndian(buf.AsSpan(i * 4), values[i]);
        s.Write(buf);
    }
}

/// <summary>
/// Data file of a calculated column that Excel has not calculated yet (H10): one run of BLANK over the table's rows, in 16
/// reserved run slots, followed by one empty bit-packed word. Excel computes the values when a PivotTable first uses the column.
/// </summary>
internal static class UnprocessedColumn
{
    public const int Slots = 16;

    /// <summary>The file for a table of <paramref name="rows"/> rows in a single segment.</summary>
    public static byte[] Segment(long rows)
    {
        var data = new byte[8 + Slots * 8 + 8 + 8];
        BinaryPrimitives.WriteUInt64LittleEndian(data, Slots);
        BinaryPrimitives.WriteUInt32LittleEndian(data.AsSpan(8), DataIds.Null);
        BinaryPrimitives.WriteUInt32LittleEndian(data.AsSpan(12), (uint)rows);
        BinaryPrimitives.WriteUInt64LittleEndian(data.AsSpan(8 + Slots * 8), 1);
        return data;
    }
}
