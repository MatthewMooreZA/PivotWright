using System.Text;

namespace PivotWright.Internal.VertiPaq;

/// <summary>Writers for column data dictionaries (.dictionary, [MS-XLDM] 2.3.2).</summary>
internal static class DictionaryWriter
{
    private const int XmTypeLong = 0;
    private const int XmTypeReal = 1;
    private const int XmTypeString = 2;

    /// <summary>Capacity of an uncompressed string page in UTF-16 code units (1 MiB of storage).</summary>
    public const long PageCapacityChars = 1L << 19;

    private static void WriteHashHeader(BinaryWriter w, int type, int algorithm, uint entrySize, uint entriesPerBin)
    {
        w.Write(type);
        w.Write(algorithm);
        w.Write(entrySize);
        w.Write(64u);            // HashBinSize
        w.Write(entriesPerBin);  // LocalEntryCount
        w.Write(-1L);            // cBins = XM_HASH_BIN_VECTOR_INVALID_BIN_COUNT
    }

    /// <summary>XM_TYPE_LONG dictionary whose values all fit in 32 bits (OperatingOn32 = true).</summary>
    public static void WriteLong32(Stream s, IReadOnlyList<long> values)
    {
        using var w = new BinaryWriter(s, Encoding.UTF8, leaveOpen: true);
        WriteHashHeader(w, XmTypeLong, -1, 8, 6);
        w.Write((ulong)values.Count);
        w.Write(4u);
        foreach (long v in values) w.Write(checked((int)v));
    }

    /// <summary>XM_TYPE_LONG dictionary with 64-bit values (OperatingOn32 = false).</summary>
    public static void WriteLong64(Stream s, IReadOnlyList<long> values)
    {
        using var w = new BinaryWriter(s, Encoding.UTF8, leaveOpen: true);
        WriteHashHeader(w, XmTypeLong, -1, 16, 3);
        w.Write((ulong)values.Count);
        w.Write(8u);
        foreach (long v in values) w.Write(v);
    }

    /// <summary>XM_TYPE_REAL dictionary.</summary>
    public static void WriteReal(Stream s, IReadOnlyList<double> values)
    {
        using var w = new BinaryWriter(s, Encoding.UTF8, leaveOpen: true);
        WriteHashHeader(w, XmTypeReal, -1, 16, 3);
        w.Write((ulong)values.Count);
        w.Write(8u);
        foreach (double v in values) w.Write(v);
    }

    /// <summary>XM_TYPE_STRING dictionary with uncompressed pages.</summary>
    public static void WriteString(Stream s, IReadOnlyList<string> values, int hashAlgorithm)
    {
        // Lay out pages.
        var pages = new List<(int start, int count, long usedChars)>();
        var handles = new (uint offset, uint page)[values.Count];
        long longest = 0;
        int pageStart = 0;
        long used = 0;
        for (int i = 0; i < values.Count; i++)
        {
            long len = values[i].Length + 1L; // NUL terminated
            longest = Math.Max(longest, values[i].Length);
            if (used > 0 && used + len > PageCapacityChars)
            {
                pages.Add((pageStart, i - pageStart, used));
                pageStart = i;
                used = 0;
            }
            handles[i] = (checked((uint)used), (uint)pages.Count);
            used += len;
        }
        if (values.Count > 0) pages.Add((pageStart, values.Count - pageStart, used));

        using var w = new BinaryWriter(s, Encoding.UTF8, leaveOpen: true);
        WriteHashHeader(w, XmTypeString, hashAlgorithm, 8, 6);
        w.Write((long)values.Count);  // StoreStringCount
        w.Write((sbyte)1);            // fStoreCompressed (Excel writes 1 even for uncompressed pages)
        w.Write(longest);             // StoreLongestString
        w.Write((long)pages.Count);   // StorePageCount

        var chars = new byte[0];
        for (int p = 0; p < pages.Count; p++)
        {
            var (start, count, usedChars) = pages[p];
            bool last = p == pages.Count - 1;
            long capacity = last ? usedChars : Math.Max(PageCapacityChars, usedChars);
            w.Write(0UL);             // page mask: XM_STRING_STORE_PAGE_OPTION_DEFAULT
            w.Write((byte)0);         // page contains nulls
            w.Write((ulong)start);    // page start index
            w.Write((ulong)count);    // page string count
            w.Write((byte)0);         // page compressed
            w.Write(new byte[] { 0xDD, 0xCC, 0xBB, 0xAA });
            w.Write((ulong)(capacity - usedChars)); // RemainingStoreAvailable
            w.Write((ulong)usedChars);              // BufferUsedCharacters
            w.Write((ulong)(capacity * 2));         // AllocationSize (bytes)
            long written = 0;
            for (int i = start; i < start + count; i++)
            {
                string v = values[i];
                int byteCount = (v.Length + 1) * 2;
                if (chars.Length < byteCount) chars = new byte[Math.Max(byteCount, chars.Length * 2)];
                Encoding.Unicode.GetBytes(v, 0, v.Length, chars, 0);
                chars[byteCount - 2] = 0; chars[byteCount - 1] = 0;
                w.Write(chars, 0, byteCount);
                written += byteCount;
            }
            long slack = capacity * 2 - written;
            if (slack > 0) WriteZeros(w, slack);
            w.Write(new byte[] { 0xCD, 0xAB, 0xCD, 0xAB });
        }

        w.Write((ulong)values.Count);
        w.Write(8u);
        foreach (var (offset, page) in handles)
        {
            w.Write(offset);
            w.Write(page);
        }
    }

    private static void WriteZeros(BinaryWriter w, long count)
    {
        var zero = new byte[Math.Min(count, 65536)];
        while (count > 0)
        {
            int n = (int)Math.Min(count, zero.Length);
            w.Write(zero, 0, n);
            count -= n;
        }
    }
}
