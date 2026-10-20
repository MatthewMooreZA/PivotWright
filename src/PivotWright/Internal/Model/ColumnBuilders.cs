using System.Globalization;
using PivotWright.Internal.Storage;
using PivotWright.Internal.VertiPaq;

namespace PivotWright.Internal.Model;

internal sealed class BuilderContext
{
    public required CultureInfo Culture { get; init; }
    public required StringComparer StringComparer { get; init; }
    public required int StringHashAlgorithm { get; init; }
    public required long SpillThreshold { get; init; }
    public string? TempDirectory { get; init; }
    public int RleThreshold { get; init; } = HybridRleEncoder.DefaultRleThreshold;
}

/// <summary>Fills the data ids of the next segment.</summary>
internal delegate void SegmentFiller(Span<uint> ids);

/// <summary>
/// Accumulates the values of one column (first pass) and encodes them into VertiPaq files (second pass).
/// </summary>
internal abstract class ColumnBuilder : IDisposable
{
    protected readonly BuilderContext Ctx;
    protected readonly Spool Spool;

    protected ColumnBuilder(ColumnDefinition definition, string id, BuilderContext ctx)
    {
        Definition = definition;
        Id = id;
        Ctx = ctx;
        Spool = new Spool(ctx.SpillThreshold, ctx.TempDirectory);
    }

    public ColumnDefinition Definition { get; }
    public string Id { get; }
    /// <summary>Column name in the model (the definition's name without characters Power Pivot rejects).</summary>
    public string ModelName { get; set; } = string.Empty;
    public long Rows { get; protected set; }

    /// <summary>Column whose values order this column (Sort by Column), or null.</summary>
    public ColumnBuilder? SortBy { get; set; }

    /// <summary>Rank of every value of this column under <see cref="SortBy"/>; set before <see cref="Encode"/>.</summary>
    public Dictionary<object, int>? SortRanks { get; set; }

    /// <summary>
    /// Set when the column is the many-side of a relationship. Its data ids then index the relationship's
    /// lookup table, so the id range is capped (wide value-encoded ranges fall back to hash encoding).
    /// </summary>
    public bool IsForeignKey { get; set; }

    /// <summary>Set when the column is the one-side key of a relationship (flagged unique in the model).</summary>
    public bool IsPrimaryKey { get; set; }

    /// <summary>Set when the column is the date column of a table marked as date table.</summary>
    public bool IsDateKey { get; set; }

    /// <summary>Largest span of value-encoded data ids for a foreign key column.</summary>
    protected const long MaxForeignKeyIdRange = 1 << 22;

    /// <summary>Largest value range that is value-encoded for this column.</summary>
    protected long ValueRangeLimit(long normal) => IsForeignKey ? Math.Min(normal, MaxForeignKeyIdRange) : normal;

    /// <summary>The value of every row in order (null for BLANK), as keys comparable with <see cref="KeyComparer"/>.</summary>
    public abstract IEnumerable<object?> Keys();

    /// <summary>After <see cref="Encode"/>: the value that a data id (3 or more) stands for.</summary>
    public Func<uint, object>? IdToKey { get; protected set; }

    /// <summary>After <see cref="Encode"/>: the data ids in sort order (BLANK first), as in the column hierarchy.</summary>
    public uint[]? SortedIds { get; private set; }

    /// <summary>After <see cref="Encode"/> of a foreign key: the number of rows per data id.</summary>
    public long[]? IdRowCounts { get; private set; }

    public abstract void AppendNull();

    /// <summary>Appends a boxed value, converting it to the column type.</summary>
    public abstract void AppendObject(object value);

    /// <summary>Encodes the column data and writes its files.</summary>
    public abstract EncodedColumn Encode(string tableId, IFileSink sink, int segmentSize);

    public void Dispose() => Spool.Dispose();

    protected EncodedColumn NewResult(int dbType, XmType xmType) => new()
    {
        Name = ModelName.Length > 0 ? ModelName : Definition.Name,
        Id = Id,
        DataType = Definition.DataType,
        IsHidden = Definition.IsHidden,
        IsPrimaryKey = IsPrimaryKey,
        IsForeignKey = IsForeignKey,
        IsDateKey = IsDateKey,
        DbType = dbType,
        XmType = xmType,
        Rows = Rows,
    };

    /// <summary>Encodes all segments of a column; <paramref name="fill"/> produces the data ids of each segment.</summary>
    protected void EncodeSegments(EncodedColumn col, string tableId, IFileSink sink, int segmentSize, SegmentFiller fill)
    {
        IdRowCounts = null;
        var ids = new uint[(int)Math.Min(segmentSize, Math.Max(1, Rows))];
        using var idf = sink.Create(FileNames.Idf(tableId, Id));
        long done = 0;
        uint minId = uint.MaxValue, maxId = 0;
        do
        {
            int n = (int)Math.Min(segmentSize, Rows - done);
            var span = ids.AsSpan(0, n);
            fill(span);
            if (IsForeignKey) CountIds(span);
            var seg = HybridRleEncoder.Encode(span, Ctx.RleThreshold);
            idf.Write(seg.Data);
            seg.Data = Array.Empty<byte>();
            col.Segments.Add(seg);
            col.RleRuns += seg.RleRuns;
            col.BitPackedRuns += seg.BitPackedRuns;
            if (n > 0 && seg.MinDataId >= DataIds.First) minId = Math.Min(minId, seg.MinDataId);
            if (n > 0) maxId = Math.Max(maxId, seg.MaxDataId);
            done += n;
        } while (done < Rows);
        col.MinDataId = minId == uint.MaxValue ? DataIds.First : minId;
        col.MaxDataId = maxId == 0 ? DataIds.Null : maxId;
    }

    private void CountIds(ReadOnlySpan<uint> ids)
    {
        var counts = IdRowCounts ?? new long[1024];
        foreach (uint id in ids)
        {
            if (id >= counts.Length)
            {
                long size = counts.Length;
                while (size <= id) size *= 2;
                Array.Resize(ref counts, checked((int)size));
            }
            counts[id]++;
        }
        IdRowCounts = counts;
    }

    /// <summary>
    /// Writes the column hierarchy: POS_TO_ID plus either ID_TO_POS (hash encoding) or a hash index
    /// (value encoding). <paramref name="posToId"/> lists the data ids in sort order, BLANK first.
    /// </summary>
    protected void WriteHierarchy(EncodedColumn col, string tableId, IFileSink sink, uint[] posToId)
    {
        if (SortRanks is not null) posToId = ApplySortRanks(posToId);
        SortedIds = posToId;
        long distinct = posToId.Length;
        long records = col.HierarchyRecords;
        var segments = col.HierarchySegments;

        var allPos = new uint[records];
        posToId.CopyTo(allPos, 0);
        using (var s = sink.Create(FileNames.PosToId(tableId, col.Id, col.AttributeHierarchyVersion)))
            WriteSegments(s, allPos, segments);

        if (col.IsValueEncoded)
        {
            var pairs = new (uint, uint)[distinct];
            for (int p = 0; p < distinct; p++) pairs[p] = (posToId[p], (uint)p);
            using var s = sink.Create(FileNames.Hidx(tableId, col.Id, col.AttributeHierarchyVersion));
            HashIndexWriter.Write(s, pairs);
        }
        else
        {
            var idToPos = new uint[records];
            idToPos[1] = (uint)distinct;
            for (int p = 0; p < distinct; p++) idToPos[posToId[p]] = (uint)p;
            using var s = sink.Create(FileNames.IdToPos(tableId, col.Id, col.AttributeHierarchyVersion));
            WriteSegments(s, idToPos, segments);
        }
    }

    // Re-orders the sorted data ids by the rank of their value under the sort column; BLANK stays first and
    // values with the same rank keep their natural order.
    private uint[] ApplySortRanks(uint[] natural)
    {
        var ranks = SortRanks!;
        var keys = new long[natural.Length];
        var order = new uint[natural.Length];
        for (int i = 0; i < natural.Length; i++)
        {
            order[i] = natural[i];
            long rank = natural[i] == DataIds.Null ? -2 : ranks[IdToKey!(natural[i])];
            keys[i] = (rank << 32) | (uint)i;
        }
        Array.Sort(keys, order);
        return order;
    }

    private static void WriteSegments(Stream s, uint[] values, IReadOnlyList<long> segments)
    {
        long start = 0;
        foreach (long n in segments)
        {
            NoSplit32Writer.WriteSegment(s, values.AsSpan((int)start, (int)n));
            start += n;
        }
    }

    protected static uint[] WithBlank(bool hasNulls, IEnumerable<uint> sortedIds)
    {
        var list = new List<uint>();
        if (hasNulls) list.Add(DataIds.Null);
        list.AddRange(sortedIds);
        return list.ToArray();
    }
}

/// <summary>Text column: hash encoded with an XM_TYPE_STRING dictionary.</summary>
internal sealed class StringColumnBuilder : ColumnBuilder
{
    private readonly Dictionary<string, uint> _exact = new(StringComparer.Ordinal);
    private readonly Dictionary<string, uint> _collated;
    private readonly List<string> _values = new();
    private bool _hasNulls;

    public StringColumnBuilder(ColumnDefinition d, string id, BuilderContext ctx) : base(d, id, ctx)
        => _collated = new Dictionary<string, uint>(ctx.StringComparer);

    public override void AppendNull()
    {
        Spool.WriteUInt32(DataIds.Null);
        _hasNulls = true;
        Rows++;
    }

    public void Append(string value)
    {
        if (!_exact.TryGetValue(value, out uint id))
        {
            if (!_collated.TryGetValue(value, out id))
            {
                // The engine compares strings case-insensitively: the first spelling seen wins.
                id = checked((uint)(DataIds.First + _values.Count));
                _values.Add(value);
                _collated.Add(value, id);
            }
            if (_exact.Count < 4_000_000) _exact[value] = id;
        }
        Spool.WriteUInt32(id);
        Rows++;
    }

    public override void AppendObject(object value)
    {
        string s = value switch
        {
            string str => str,
            IFormattable f => f.ToString(null, CultureInfo.InvariantCulture),
            _ => value.ToString() ?? string.Empty,
        };
        Append(s);
    }

    public override IEnumerable<object?> Keys()
    {
        var reader = Spool.OpenReader();
        var buf = new uint[8192];
        for (long done = 0; done < Rows;)
        {
            int n = (int)Math.Min(buf.Length, Rows - done);
            reader.ReadUInt32(buf.AsSpan(0, n));
            for (int i = 0; i < n; i++) yield return buf[i] == DataIds.Null ? null : _values[(int)(buf[i] - DataIds.First)];
            done += n;
        }
    }

    public override EncodedColumn Encode(string tableId, IFileSink sink, int segmentSize)
    {
        IdToKey = id => _values[(int)(id - DataIds.First)];
        var col = NewResult(130, XmType.String);
        col.Dictionary = DictionaryKind.String;
        col.HasNulls = _hasNulls;
        col.LastId = DataIds.Null + _values.Count;
        col.DistinctInclBlank = _values.Count + (_hasNulls ? 1 : 0);

        var reader = Spool.OpenReader();
        EncodeSegments(col, tableId, sink, segmentSize, span => reader.ReadUInt32(span));

        using (var s = sink.Create(FileNames.Dictionary(tableId, Id)))
            DictionaryWriter.WriteString(s, _values, Ctx.StringHashAlgorithm);

        var order = new int[_values.Count];
        for (int i = 0; i < order.Length; i++) order[i] = i;
        var keys = _values.ToArray();
        Array.Sort(keys, order, Ctx.StringComparer);
        WriteHierarchy(col, tableId, sink, WithBlank(_hasNulls, order.Select(i => (uint)(DataIds.First + i))));
        return col;
    }
}

/// <summary>TRUE/FALSE column: hash encoded with an XM_TYPE_LONG dictionary holding 0 and 1.</summary>
internal sealed class BooleanColumnBuilder : ColumnBuilder
{
    private readonly List<long> _values = new(2);
    private uint _falseId, _trueId;
    private bool _hasNulls;

    public BooleanColumnBuilder(ColumnDefinition d, string id, BuilderContext ctx) : base(d, id, ctx) { }

    public override void AppendNull()
    {
        Spool.WriteUInt32(DataIds.Null);
        _hasNulls = true;
        Rows++;
    }

    public void Append(bool value)
    {
        ref uint id = ref value ? ref _trueId : ref _falseId;
        if (id == 0)
        {
            id = (uint)(DataIds.First + _values.Count);
            _values.Add(value ? 1 : 0);
        }
        Spool.WriteUInt32(id);
        Rows++;
    }

    public override void AppendObject(object value) => Append(value switch
    {
        bool b => b,
        string s => bool.Parse(s),
        _ => Convert.ToBoolean(value, CultureInfo.InvariantCulture),
    });

    public override IEnumerable<object?> Keys()
    {
        var reader = Spool.OpenReader();
        var buf = new uint[8192];
        for (long done = 0; done < Rows;)
        {
            int n = (int)Math.Min(buf.Length, Rows - done);
            reader.ReadUInt32(buf.AsSpan(0, n));
            for (int i = 0; i < n; i++) yield return buf[i] == DataIds.Null ? null : _values[(int)(buf[i] - DataIds.First)];
            done += n;
        }
    }

    public override EncodedColumn Encode(string tableId, IFileSink sink, int segmentSize)
    {
        IdToKey = id => _values[(int)(id - DataIds.First)];
        var col = NewResult(11, XmType.Long);
        col.Dictionary = DictionaryKind.Long32;
        col.HasNulls = _hasNulls;
        col.LastId = DataIds.Null + _values.Count;
        col.DistinctInclBlank = _values.Count + (_hasNulls ? 1 : 0);

        var reader = Spool.OpenReader();
        EncodeSegments(col, tableId, sink, segmentSize, span => reader.ReadUInt32(span));
        using (var s = sink.Create(FileNames.Dictionary(tableId, Id)))
            DictionaryWriter.WriteLong32(s, _values);

        var sorted = new List<uint>();
        if (_falseId != 0) sorted.Add(_falseId);
        if (_trueId != 0) sorted.Add(_trueId);
        WriteHierarchy(col, tableId, sink, WithBlank(_hasNulls, sorted));
        return col;
    }
}

/// <summary>
/// Integer-valued column (Whole Number, and Currency stored as units of 1/10,000). Uses value
/// encoding when the value range fits in 31 bits, otherwise a hash dictionary.
/// </summary>
internal sealed class Int64ColumnBuilder : ColumnBuilder
{
    private const long MaxValueRange = int.MaxValue - 64;
    private readonly HashSet<long> _distinct = new();
    private readonly NullMask _nulls = new();
    private long _min = long.MaxValue, _max = long.MinValue;
    private readonly bool _currency;

    public Int64ColumnBuilder(ColumnDefinition d, string id, BuilderContext ctx) : base(d, id, ctx)
        => _currency = d.DataType == ColumnDataType.Currency;

    public override void AppendNull()
    {
        _nulls.Set(Rows);
        Spool.WriteInt64(0);
        Rows++;
    }

    public void Append(long value)
    {
        Spool.WriteInt64(value);
        _distinct.Add(value);
        if (value < _min) _min = value;
        if (value > _max) _max = value;
        Rows++;
    }

    public void AppendCurrency(decimal value) => Append(ToCurrencyUnits(value));

    public static long ToCurrencyUnits(decimal value)
    {
        decimal scaled = decimal.Round(value * 10000m, 0, MidpointRounding.AwayFromZero);
        if (scaled > long.MaxValue || scaled < long.MinValue)
            throw new OverflowException($"Value {value} is outside the range of the Currency data type.");
        return (long)scaled;
    }

    public override void AppendObject(object value)
    {
        if (_currency)
        {
            AppendCurrency(Convert.ToDecimal(value, CultureInfo.InvariantCulture));
            return;
        }
        Append(value switch
        {
            long l => l,
            int i => i,
            short s => s,
            byte b => b,
            sbyte sb => sb,
            ushort us => us,
            uint ui => ui,
            ulong ul => checked((long)ul),
            bool bo => bo ? 1 : 0,
            _ => Convert.ToInt64(value, CultureInfo.InvariantCulture),
        });
    }

    public override IEnumerable<object?> Keys()
    {
        var reader = Spool.OpenReader();
        var buf = new long[8192];
        for (long done = 0; done < Rows;)
        {
            int n = (int)Math.Min(buf.Length, Rows - done);
            reader.ReadInt64(buf.AsSpan(0, n));
            for (int i = 0; i < n; i++) yield return _nulls.IsNull(done + i) ? null : buf[i];
            done += n;
        }
    }

    public override EncodedColumn Encode(string tableId, IFileSink sink, int segmentSize)
    {
        var col = NewResult(_currency ? 6 : 20, XmType.Long);
        col.HasNulls = _nulls.Any;
        col.DistinctInclBlank = _distinct.Count + (_nulls.Any ? 1 : 0);
        var sorted = _distinct.ToArray();
        Array.Sort(sorted);
        var reader = Spool.OpenReader();
        var raw = new long[(int)Math.Min(segmentSize, Math.Max(1, Rows))];
        long row = 0;

        bool valueEncoding = _distinct.Count > 0 && (decimal)_max - _min <= ValueRangeLimit(MaxValueRange);
        if (_distinct.Count == 0) valueEncoding = false;

        if (valueEncoding)
        {
            // Currency columns use the largest power-of-ten scale that divides every value.
            long scale = 1;
            if (_currency)
            {
                foreach (long p in new long[] { 10000, 1000, 100, 10 })
                {
                    if (sorted.All(v => v % p == 0)) { scale = p; break; }
                }
            }
            long baseId = _min / scale - DataIds.First;
            col.Dictionary = DictionaryKind.Value;
            col.BaseId = baseId;
            col.Magnitude = 1.0 / scale;
            IdToKey = id => ((long)id + baseId) * scale;
            EncodeSegments(col, tableId, sink, segmentSize, span =>
            {
                var buf = raw.AsSpan(0, span.Length);
                reader.ReadInt64(buf);
                for (int i = 0; i < span.Length; i++, row++)
                    span[i] = _nulls.IsNull(row) ? DataIds.Null : (uint)(buf[i] / scale - baseId);
            });
            WriteHierarchy(col, tableId, sink, WithBlank(col.HasNulls, sorted.Select(v => (uint)(v / scale - baseId))));
            return col;
        }

        // Hash encoding: ids assigned in order of first appearance.
        var map = new Dictionary<long, uint>(_distinct.Count);
        var values = new List<long>(_distinct.Count);
        EncodeSegments(col, tableId, sink, segmentSize, span =>
        {
            var buf = raw.AsSpan(0, span.Length);
            reader.ReadInt64(buf);
            for (int i = 0; i < span.Length; i++, row++)
            {
                if (_nulls.IsNull(row)) { span[i] = DataIds.Null; continue; }
                long v = buf[i];
                if (!map.TryGetValue(v, out uint id))
                {
                    id = (uint)(DataIds.First + values.Count);
                    map.Add(v, id);
                    values.Add(v);
                }
                span[i] = id;
            }
        });
        bool fits32 = _distinct.Count == 0 || (_min >= int.MinValue && _max <= int.MaxValue);
        IdToKey = id => values[(int)(id - DataIds.First)];
        col.Dictionary = fits32 ? DictionaryKind.Long32 : DictionaryKind.Long64;
        col.LastId = DataIds.Null + values.Count;
        using (var s = sink.Create(FileNames.Dictionary(tableId, Id)))
        {
            if (fits32) DictionaryWriter.WriteLong32(s, values);
            else DictionaryWriter.WriteLong64(s, values);
        }
        WriteHierarchy(col, tableId, sink, WithBlank(col.HasNulls, sorted.Select(v => map[v])));
        return col;
    }
}

/// <summary>
/// Floating point column (Decimal Number, and Date stored as an OLE Automation date). Uses value
/// encoding when every value is integral and the range fits in 31 bits, otherwise an
/// XM_TYPE_REAL hash dictionary.
/// </summary>
internal sealed class DoubleColumnBuilder : ColumnBuilder
{
    private const double MaxValueRange = int.MaxValue - 64;
    private readonly HashSet<double> _distinct = new();
    private readonly NullMask _nulls = new();
    private double _min = double.MaxValue, _max = double.MinValue;
    private bool _allIntegral = true;
    private readonly bool _date;

    public DoubleColumnBuilder(ColumnDefinition d, string id, BuilderContext ctx) : base(d, id, ctx)
        => _date = d.DataType == ColumnDataType.DateTime;

    public override void AppendNull()
    {
        _nulls.Set(Rows);
        Spool.WriteDouble(0);
        Rows++;
    }

    public void Append(double value)
    {
        if (double.IsNaN(value) || double.IsInfinity(value)) { AppendNull(); return; }
        if (value == 0) value = 0; // normalise -0.0
        Spool.WriteDouble(value);
        _distinct.Add(value);
        if (value < _min) _min = value;
        if (value > _max) _max = value;
        if (_allIntegral && value != Math.Floor(value)) _allIntegral = false;
        Rows++;
    }

    public void AppendDate(DateTime value) => Append(value.ToOADate());

    public static double ToOADate(object value) => value switch
    {
        DateTime dt => dt.ToOADate(),
        DateTimeOffset dto => dto.DateTime.ToOADate(),
        DateOnly d => d.ToDateTime(TimeOnly.MinValue).ToOADate(),
        TimeOnly t => t.ToTimeSpan().TotalDays,
        TimeSpan ts => ts.TotalDays,
        string s => DateTime.Parse(s, CultureInfo.InvariantCulture).ToOADate(),
        _ => Convert.ToDateTime(value, CultureInfo.InvariantCulture).ToOADate(),
    };

    public override void AppendObject(object value)
    {
        if (_date) { Append(ToOADate(value)); return; }
        Append(value switch
        {
            double d => d,
            float f => f,
            decimal m => (double)m,
            _ => Convert.ToDouble(value, CultureInfo.InvariantCulture),
        });
    }

    public override IEnumerable<object?> Keys()
    {
        var reader = Spool.OpenReader();
        var buf = new long[8192];
        for (long done = 0; done < Rows;)
        {
            int n = (int)Math.Min(buf.Length, Rows - done);
            reader.ReadInt64(buf.AsSpan(0, n));
            for (int i = 0; i < n; i++) yield return _nulls.IsNull(done + i) ? null : BitConverter.Int64BitsToDouble(buf[i]);
            done += n;
        }
    }

    public override EncodedColumn Encode(string tableId, IFileSink sink, int segmentSize)
    {
        var col = NewResult(_date ? 7 : 5, XmType.Real);
        col.HasNulls = _nulls.Any;
        col.DistinctInclBlank = _distinct.Count + (_nulls.Any ? 1 : 0);
        var sorted = _distinct.ToArray();
        Array.Sort(sorted);
        var reader = Spool.OpenReader();
        var raw = new long[(int)Math.Min(segmentSize, Math.Max(1, Rows))];
        long row = 0;

        bool valueEncoding = _distinct.Count > 0 && _allIntegral &&
                             Math.Abs(_min) < 9e15 && Math.Abs(_max) < 9e15 && _max - _min <= ValueRangeLimit((long)MaxValueRange);
        if (valueEncoding)
        {
            long baseId = (long)_min - DataIds.First;
            col.Dictionary = DictionaryKind.Value;
            col.BaseId = baseId;
            IdToKey = id => (double)((long)id + baseId);
            EncodeSegments(col, tableId, sink, segmentSize, span =>
            {
                var buf = raw.AsSpan(0, span.Length);
                reader.ReadInt64(buf);
                for (int i = 0; i < span.Length; i++, row++)
                    span[i] = _nulls.IsNull(row) ? DataIds.Null : (uint)((long)BitConverter.Int64BitsToDouble(buf[i]) - baseId);
            });
            WriteHierarchy(col, tableId, sink, WithBlank(col.HasNulls, sorted.Select(v => (uint)((long)v - baseId))));
            return col;
        }

        var map = new Dictionary<double, uint>(_distinct.Count);
        var values = new List<double>(_distinct.Count);
        EncodeSegments(col, tableId, sink, segmentSize, span =>
        {
            var buf = raw.AsSpan(0, span.Length);
            reader.ReadInt64(buf);
            for (int i = 0; i < span.Length; i++, row++)
            {
                if (_nulls.IsNull(row)) { span[i] = DataIds.Null; continue; }
                double v = BitConverter.Int64BitsToDouble(buf[i]);
                if (!map.TryGetValue(v, out uint id))
                {
                    id = (uint)(DataIds.First + values.Count);
                    map.Add(v, id);
                    values.Add(v);
                }
                span[i] = id;
            }
        });
        IdToKey = id => values[(int)(id - DataIds.First)];
        col.Dictionary = DictionaryKind.Real;
        col.LastId = DataIds.Null + values.Count;
        using (var s = sink.Create(FileNames.Dictionary(tableId, Id)))
            DictionaryWriter.WriteReal(s, values);
        WriteHierarchy(col, tableId, sink, WithBlank(col.HasNulls, sorted.Select(v => map[v])));
        return col;
    }
}
