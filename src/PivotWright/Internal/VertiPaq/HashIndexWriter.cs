using System.Buffers.Binary;

namespace PivotWright.Internal.VertiPaq;

/// <summary>
/// Writes a column hierarchy hash index (.hidx, [MS-XLDM] 2.3.3) mapping data identifiers to
/// sorted positions for value-encoded columns. The layout (64-byte bins holding six 8-byte
/// entries, overflow entries appended in bin order, hash statistics with a 10-slot histogram)
/// and the bin-count growth rule were verified byte-for-byte against Excel-generated files.
/// </summary>
internal static class HashIndexWriter
{
    private const uint MagicConstant = 0x12B9B6A5;
    private const int EntrySize = 8;
    private const int BinSize = 64;
    private const int EntriesPerBin = 6;
    private const int HistogramSlots = 10;

    public static int BinCount(long records)
    {
        long bins = 16;
        while (records > bins * 4.5) bins *= 2;
        return checked((int)bins);
    }

    private static int Bucket(uint key, int bins)
    {
        // cHashBitsUsed = CEILING(LOG2(cBuckets)); bins is always a power of two here.
        int bits = System.Numerics.BitOperations.Log2((uint)bins);
        return (int)((key * MagicConstant) >> (32 - bits));
    }

    /// <param name="pairs">(data id, position) pairs, in insertion (position) order.</param>
    public static void Write(Stream s, IReadOnlyList<(uint Key, uint Value)> pairs)
    {
        int binCount = BinCount(pairs.Count);
        var counts = new int[binCount];
        foreach (var p in pairs) counts[Bucket(p.Key, binCount)]++;

        // Bucket the entries in insertion order.
        var starts = new int[binCount + 1];
        for (int i = 0; i < binCount; i++) starts[i + 1] = starts[i] + counts[i];
        var fill = (int[])starts.Clone();
        var ordered = new (uint, uint)[pairs.Count];
        foreach (var p in pairs) ordered[fill[Bucket(p.Key, binCount)]++] = p;

        long used = 0, fast = 0, maxChain = 0;
        var hist = new long[HistogramSlots];
        for (int i = 0; i < binCount; i++)
        {
            int c = counts[i];
            if (c == 0) continue;
            used++;
            fast += Math.Min(c, EntriesPerBin);
            maxChain = Math.Max(maxChain, c);
            hist[Math.Min(c, HistogramSlots) - 1]++;
        }

        var w = new BinaryWriter(s, System.Text.Encoding.UTF8, leaveOpen: true);
        w.Write(-1);                       // HashAlgorithm = XM_INVALID
        w.Write((uint)EntrySize);
        w.Write((uint)BinSize);
        w.Write((uint)EntriesPerBin);
        w.Write((long)binCount);           // cBins
        w.Write((ulong)pairs.Count);       // NumberOfRecords
        w.Write((ulong)(binCount - 1));    // CurrentMask
        w.Write((byte)1);                  // HashStats present
        w.Write((ulong)pairs.Count);       // NumberOfElements
        w.Write((ulong)binCount);          // NumberOfBins
        w.Write((ulong)used);              // NumberOfUsedBins
        w.Write((ulong)fast);              // FastAccessElements
        w.Write((ulong)EntriesPerBin);     // LocalsSizePerBin
        w.Write((ulong)maxChain);          // MaximumChain
        w.Write((ulong)HistogramSlots);    // histogram element count
        w.Write((uint)8);                  // histogram element size
        foreach (long h in hist) w.Write((ulong)h);

        var bin = new byte[BinSize];
        for (int i = 0; i < binCount; i++)
        {
            Array.Clear(bin);
            // m_rgChain (8 bytes, NULL on disk), then m_Count.
            BinaryPrimitives.WriteUInt32LittleEndian(bin.AsSpan(8), (uint)counts[i]);
            int local = Math.Min(counts[i], EntriesPerBin);
            for (int j = 0; j < local; j++)
            {
                var (k, v) = ordered[starts[i] + j];
                BinaryPrimitives.WriteUInt32LittleEndian(bin.AsSpan(12 + j * EntrySize), k);
                BinaryPrimitives.WriteUInt32LittleEndian(bin.AsSpan(16 + j * EntrySize), v);
            }
            w.Write(bin);
        }

        long collisions = 0;
        for (int i = 0; i < binCount; i++) collisions += Math.Max(0, counts[i] - EntriesPerBin);
        w.Write((ulong)collisions);
        for (int i = 0; i < binCount; i++)
        {
            for (int j = EntriesPerBin; j < counts[i]; j++)
            {
                var (k, v) = ordered[starts[i] + j];
                w.Write(k);
                w.Write(v);
            }
        }
        w.Flush();
    }
}
