using PivotWright.Internal.VertiPaq;

namespace PivotWright.Internal.Model;

internal sealed class EncodedTable
{
    public required string Name { get; init; }
    public required string Id { get; init; }
    public required long Rows { get; init; }
    public required int SegmentSize { get; init; }
    public required EncodedColumn RowNumber { get; init; }
    public required List<EncodedColumn> Columns { get; init; }
    public int SegmentCount => RowNumber.Segments.Count;
    /// <summary>Relationships in which this table is the many side.</summary>
    public List<EncodedRelationship> Relationships { get; } = new();
    public List<EncodedHierarchy> Hierarchies { get; } = new();
    /// <summary>The table is marked as date table (dimension type Time, its date column is the key).</summary>
    public bool IsDateTable { get; init; }
    /// <summary>Version of the table's .tbl.xml: 2 with hierarchies, plus 1 for a date table and the tables that relate to one.</summary>
    public int Version { get; set; }

    public IEnumerable<EncodedColumn> AllColumns
    {
        get
        {
            yield return RowNumber;
            foreach (var c in Columns) yield return c;
        }
    }
}

/// <summary>Collects the rows of one table.</summary>
internal sealed class TableBuilder : IDisposable
{
    private readonly List<ColumnBuilder> _columns = new();
    private readonly BuilderContext _ctx;

    public TableBuilder(string name, string id, IEnumerable<ColumnDefinition> columns, BuilderContext ctx)
    {
        Name = name;
        Id = id;
        _ctx = ctx;
        var used = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { FileNames.RowNumberId };
        var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var modelNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { FileNames.RowNumberId };
        foreach (var def in columns)
        {
            if (!names.Add(def.Name))
                throw new ArgumentException($"Table '{name}' already has a column named '{def.Name}'.", "columns");
            string modelName = Identifiers.Unique(Identifiers.SanitizeName(def.Name, "Column"), modelNames);
            string colId = Identifiers.Unique(Identifiers.Sanitize(modelName, "Column"), used);
            ColumnBuilder builder = def.DataType switch
            {
                ColumnDataType.String => new StringColumnBuilder(def, colId, ctx),
                ColumnDataType.Boolean => new BooleanColumnBuilder(def, colId, ctx),
                ColumnDataType.Int64 or ColumnDataType.Currency => new Int64ColumnBuilder(def, colId, ctx),
                ColumnDataType.Double or ColumnDataType.DateTime => new DoubleColumnBuilder(def, colId, ctx),
                _ => throw new NotSupportedException(def.DataType.ToString()),
            };
            builder.ModelName = modelName;
            _columns.Add(builder);
        }
        if (_columns.Count == 0) throw new ArgumentException($"Table '{name}' must have at least one column.", "columns");
        foreach (var b in _columns.ToList())
            if (b.Definition.SortByColumn is { } by) SetSortBy(b.Definition.Name, by);
    }

    /// <summary>Finds a column by the name it was defined with or its model name, ignoring case.</summary>
    /// <param name="name">Column name.</param>
    /// <param name="paramName">Parameter of the public method that took the name, reported in the exception.</param>
    public ColumnBuilder FindColumn(string name, string paramName = "column")
    {
        ArgumentNullException.ThrowIfNull(name, paramName);
        return _columns.FirstOrDefault(c => string.Equals(c.Definition.Name, name, StringComparison.OrdinalIgnoreCase)
                                            || string.Equals(c.ModelName, name, StringComparison.OrdinalIgnoreCase))
            ?? throw new ArgumentException(
                $"Table '{Name}' has no column named '{name}'. Its columns are: {Identifiers.List(_columns.Select(c => c.Definition.Name))}.", paramName);
    }

    /// <summary>Set when the table is complete (a generated date table); the text says why rows cannot be added.</summary>
    public string? ReadOnlyReason { get; set; }

    // Set when a row was only partly added (the source failed in the middle of a row); the table can then not be saved.
    private string? _fault;

    // Values of the row being added, converted before any column takes one.
    private object?[]? _scratch;

    /// <summary>Hides a column from client tools, or shows it again.</summary>
    public void SetHidden(string column, bool hidden, string paramName = "columns")
    {
        var c = FindColumn(column, paramName);
        if (c.Definition.IsHidden != hidden) c.Definition = c.Definition.With(isHidden: hidden);
    }

    /// <summary>Orders <paramref name="column"/> by <paramref name="sortColumn"/> (Power Pivot's Sort by Column).</summary>
    public void SetSortBy(string column, string sortColumn)
    {
        var c = FindColumn(column, nameof(column));
        var by = FindColumn(sortColumn, nameof(sortColumn));
        if (c == by) throw new ArgumentException($"Column '{c.Definition.Name}' cannot be sorted by itself.");
        if (by.SortBy is not null)
            throw new ArgumentException($"Column '{by.Definition.Name}' is itself sorted by another column, so it cannot be a sort column.");
        if (_columns.Any(o => o.SortBy == c))
            throw new ArgumentException($"Column '{c.Definition.Name}' is the sort column of another column, so it cannot be sorted itself.");
        c.SortBy = by;
        if (!string.Equals(c.Definition.SortByColumn, by.Definition.Name, StringComparison.Ordinal))
            c.Definition = c.Definition.With(sortByColumn: by.Definition.Name);
    }

    // For every sorted column: checks that each of its values belongs to exactly one value of the sort column
    // and ranks the values by the order of their sort values (ties keep the column's own order).
    private void ComputeSortRanks()
    {
        var cmp = new KeyComparer(_ctx.StringComparer);
        foreach (var c in _columns.Where(c => c.SortBy is not null))
        {
            var by = c.SortBy!;
            var sortOf = new Dictionary<object, object>(cmp);
            object blank = DBNull.Value;
            using (var a = c.Keys().GetEnumerator())
            using (var b = by.Keys().GetEnumerator())
            {
                while (a.MoveNext() && b.MoveNext())
                {
                    if (a.Current is null) continue;
                    object sv = b.Current ?? blank;
                    if (sortOf.TryAdd(a.Current, sv)) continue;
                    if (!cmp.Equals(sortOf[a.Current], sv))
                        throw new InvalidOperationException(
                            $"Column '{c.Definition.Name}' cannot be sorted by '{by.Definition.Name}' in table '{Name}': the value '{a.Current}' appears with more than one '{by.Definition.Name}' value ('{Show(sortOf[a.Current])}' and '{Show(sv)}').");
                }
            }
            var distinct = new Dictionary<object, int>(cmp);
            foreach (var v in sortOf.Values) if (!ReferenceEquals(v, blank)) distinct.TryAdd(v, 0);
            var list = distinct.Keys.ToList();
            list.Sort((x, y) => x is string sx && y is string sy ? _ctx.StringComparer.Compare(sx, sy) : Comparer<object>.Default.Compare(x, y));
            for (int i = 0; i < list.Count; i++) distinct[list[i]] = i;
            var ranks = new Dictionary<object, int>(cmp);
            foreach (var (k, v) in sortOf) ranks[k] = ReferenceEquals(v, blank) ? -1 : distinct[v];
            c.SortRanks = ranks;
        }
    }

    private static string Show(object v) => ReferenceEquals(v, DBNull.Value) ? "BLANK" : Convert.ToString(v, System.Globalization.CultureInfo.InvariantCulture) ?? "";

    public string Name { get; }
    public string Id { get; }
    /// <summary>The date column when the table is marked as date table, else null.</summary>
    public ColumnBuilder? DateColumn { get; set; }
    public long Rows { get; private set; }
    public IReadOnlyList<ColumnBuilder> Columns => _columns;

    /// <summary>Throws when rows cannot be added to the table.</summary>
    public void ThrowIfReadOnly()
    {
        if (ReadOnlyReason is not null) throw new InvalidOperationException(ReadOnlyReason);
    }

    /// <summary>
    /// Adds a row. Every value is converted first, so a value that does not fit its column leaves the table unchanged and
    /// the caller can skip the row and go on.
    /// </summary>
    public void AddRow(ReadOnlySpan<object?> values)
    {
        ThrowIfReadOnly();
        if (values.Length != _columns.Count)
            throw new ArgumentException($"Row {Rows + 1} of table '{Name}' has {values.Length} values, but the table has {_columns.Count} columns ({Identifiers.List(_columns.Select(c => c.Definition.Name))}).", nameof(values));
        var row = _scratch ??= new object?[_columns.Count];
        for (int i = 0; i < values.Length; i++)
        {
            object? v = values[i];
            row[i] = v is null || v is DBNull ? null : Coerce(i, v, Rows + 1);
        }
        for (int i = 0; i < row.Length; i++)
        {
            if (row[i] is { } v) _columns[i].AppendObject(v);
            else _columns[i].AppendNull();
            row[i] = null;
        }
        Rows++;
    }

    /// <summary>Converts a value for the column at <paramref name="ordinal"/>, naming row, column and value when it does not fit.</summary>
    public object Coerce(int ordinal, object value, long row)
    {
        var c = _columns[ordinal];
        try { return c.Coerce(value); }
        catch (Exception e) when (e is FormatException or InvalidCastException or OverflowException)
        {
            throw new ArgumentException(
                $"Row {row} of table '{Name}': column '{c.Definition.Name}' ({c.Definition.DataType}) cannot hold the {value.GetType().Name} value '{value}'. {e.Message}", e);
        }
    }

    /// <summary>Called by typed writers after they appended one value to every column.</summary>
    public void CommitRow() => Rows++;

    /// <summary>Records that a row was only partly added; saving then fails with <paramref name="reason"/>.</summary>
    public void Fault(string reason) => _fault ??= reason;

    public EncodedTable Encode(IFileSink sink, int segmentSize)
    {
        if (_fault is not null) throw new InvalidOperationException(_fault);
        foreach (var c in _columns)
            if (c.Rows != Rows) throw new InvalidOperationException($"Column '{c.Definition.Name}' has {c.Rows} values but table '{Name}' has {Rows} rows.");

        var rowNumber = new EncodedColumn
        {
            Name = FileNames.RowNumberId,
            Id = FileNames.RowNumberId,
            DataType = ColumnDataType.Int64,
            IsHidden = true,
            IsRowNumber = true,
            Dictionary = DictionaryKind.Value,
            XmType = XmType.Long,
            DbType = 3,
            BaseId = -DataIds.First,
            Rows = Rows,
            DistinctInclBlank = Rows,
            MinDataId = DataIds.First,
            MaxDataId = (uint)(DataIds.First + Rows - 1),
        };
        using (var idf = sink.Create(FileNames.Idf(Id, FileNames.RowNumberId)))
        {
            long done = 0;
            do
            {
                long n = Math.Min(segmentSize, Rows - done);
                var seg = HybridRleEncoder.EncodeRowNumber(done, n);
                idf.Write(seg.Data);
                seg.Data = Array.Empty<byte>();
                rowNumber.Segments.Add(seg);
                done += n;
            } while (done < Rows);
        }

        ComputeSortRanks();
        var cols = new List<EncodedColumn>(_columns.Count);
        foreach (var c in _columns)
        {
            var e = c.Encode(Id, sink, segmentSize);
            if (c.SortBy is not null) e.OrderByColumnId = c.SortBy.Id;
            cols.Add(e);
        }
        return new EncodedTable { Name = Name, Id = Id, Rows = Rows, SegmentSize = segmentSize, RowNumber = rowNumber, Columns = cols, IsDateTable = DateColumn is not null };
    }

    public void Dispose()
    {
        foreach (var c in _columns) c.Dispose();
    }
}

internal static class Identifiers
{
    private static readonly char[] Removed = "\\/:*?\"<>|.$()[]{}'`;=,&%#@!^~+".ToCharArray();

    // Characters Power Pivot does not allow in table, column or measure names
    // ("Naming Requirements for Tables and Columns", PowerPivot for Excel).
    private static readonly char[] InvalidInNames = ".,;':/\\*|?&%$!+=()[]{}<>\"".ToCharArray();

    /// <summary>Names quoted and separated by commas for an error message, the first 20 only.</summary>
    public static string List(IEnumerable<string> names)
    {
        var all = names.ToList();
        if (all.Count == 0) return "(none)";
        string shown = string.Join(", ", all.Take(20).Select(n => "'" + n + "'"));
        return all.Count > 20 ? $"{shown} and {all.Count - 20} more" : shown;
    }

    /// <summary>Removes characters that Power Pivot rejects in object names, and surrounding spaces.</summary>
    public static string SanitizeName(string name, string fallback)
    {
        var sb = new System.Text.StringBuilder(name.Length);
        foreach (char c in name)
        {
            if (char.IsControl(c) || Array.IndexOf(InvalidInNames, c) >= 0) continue;
            sb.Append(c);
        }
        string s = sb.ToString().Trim();
        return s.Length == 0 ? fallback : s;
    }

    /// <summary>Builds an object identifier that is safe to use in model file names.</summary>
    public static string Sanitize(string name, string fallback)
    {
        var sb = new System.Text.StringBuilder(name.Length);
        foreach (char c in name)
        {
            if (char.IsControl(c) || Array.IndexOf(Removed, c) >= 0) continue;
            sb.Append(c);
        }
        string s = System.Text.RegularExpressions.Regex.Replace(sb.ToString(), "\\s+", " ").Trim();
        if (s.Length > 100) s = s[..100].TrimEnd();
        return s.Length == 0 ? fallback : s;
    }

    public static string Unique(string id, HashSet<string> used)
    {
        string candidate = id;
        for (int i = 2; !used.Add(candidate); i++) candidate = id + " " + i;
        return candidate;
    }
}
