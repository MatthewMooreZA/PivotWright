using System.Data;
using System.Globalization;
using PivotWright.Internal;
using PivotWright.Internal.Model;
using PivotWright.Internal.Package;

namespace PivotWright;

/// <summary>
/// Builds an Excel workbook (.xlsx) whose Power Pivot data model contains the tables you add.
/// Rows are streamed: each column is encoded into the xVelocity (VertiPaq) column-store format
/// when <see cref="Save(string)"/> is called, and spilled to temporary files while loading so that
/// tables larger than memory can be written.
/// </summary>
/// <remarks>
/// The model owns tables and relationships; a <see cref="ModelTable"/> owns its columns, measures and hierarchies. Every
/// <c>Add…</c> method returns the object it added, and <see cref="Tables"/> finds a table by name.
/// </remarks>
/// <example>
/// <code>
/// using var model = new DataModelWriter();
/// var products = model.AddTable("Products", productsReader);
/// var sales = model.AddTable("Sales", salesReader);
/// model.AddRelationship(sales["ProductKey"], products["ProductKey"]);
/// sales.AddMeasure("Total Qty", "SUM(Sales[Qty])", MeasureFormat.WholeNumber);
/// model.Save("sales.xlsx");
/// </code>
/// </example>
public sealed class DataModelWriter : IDisposable
{
    private readonly DataModelOptions _options;
    private readonly BuilderContext _ctx;
    private readonly List<ModelTable> _tables = new();
    private readonly List<ModelRelationship> _relationships = new();
    private readonly List<RelationshipDefinition> _relationshipDefinitions = new();
    private bool _disposed;

    public DataModelWriter(DataModelOptions? options = null)
    {
        _options = options ?? new DataModelOptions();
        var culture = _options.Culture;
        _ctx = new BuilderContext
        {
            Culture = culture,
            StringComparer = StringComparer.Create(culture,
                CompareOptions.IgnoreCase | CompareOptions.IgnoreKanaType | CompareOptions.IgnoreWidth),
            // XM_HASH_ALGORITHM_FAST_CI for en-US, XM_HASH_ALGORITHM_SQL otherwise ([MS-XLDM] 2.3.3.1.4.2).
            StringHashAlgorithm = _options.Lcid == 1033 ? 1 : 0,
            SpillThreshold = _options.SpillThresholdBytes,
            TempDirectory = _options.TempDirectory,
        };
        Tables = new ModelCollection<ModelTable>(_tables, FindTable,
            name => $"There is no table named '{name}'. The model's tables are: {Identifiers.List(_tables.Select(t => t.Name))}.");
    }

    /// <summary>The tables, in the order added; <c>Tables["Sales"]</c> finds one by the name given or as stored, ignoring case.</summary>
    public ModelCollection<ModelTable> Tables { get; }

    /// <summary>The relationships, in the order added.</summary>
    public IReadOnlyList<ModelRelationship> Relationships => _relationships;

    /// <summary>The table marked with <see cref="ModelTable.MarkAsDateTable"/> or made by <see cref="AddDateTable"/>, or null.</summary>
    public ModelTable? DateTable => _tables.FirstOrDefault(t => t.IsDateTable);

    // Every measure and hierarchy of the model in the order added, as the serializer writes them.
    internal List<MeasureDefinition> MeasureDefinitions { get; } = new();
    internal List<HierarchyDefinition> HierarchyDefinitions { get; } = new();

    /// <summary>Creates an empty table; add rows with <see cref="ModelTable.AddRow"/>.</summary>
    /// <param name="name">Table name. Characters Power Pivot rejects are removed (see <see cref="ModelTable.Name"/>); <see cref="Tables"/>
    /// finds the table by either form.</param>
    /// <param name="columns">The columns, in order.</param>
    /// <exception cref="ArgumentException">The name is empty or taken, there are no columns, or two have the same name.</exception>
    public ModelTable CreateTable(string name, IEnumerable<ColumnDefinition> columns)
    {
        ThrowIfDisposed();
        if (string.IsNullOrWhiteSpace(name)) throw new ArgumentException("Table name is required.", nameof(name));
        ArgumentNullException.ThrowIfNull(columns);
        string modelName = Identifiers.SanitizeName(name, "Table");
        if (_tables.Any(t => string.Equals(t.Name, modelName, StringComparison.OrdinalIgnoreCase)))
            throw new ArgumentException($"A table named '{modelName}' already exists.", nameof(name));
        string id = Identifiers.Sanitize(modelName, "Table") + "_" + Guid.NewGuid().ToString("D");
        var table = new ModelTable(this, new TableBuilder(modelName, id, columns, _ctx));
        _tables.Add(table);
        return table;
    }

    /// <summary>
    /// Adds a table and loads every row of <paramref name="reader"/> into it. Column types are taken
    /// from the reader's field types unless overridden.
    /// </summary>
    /// <param name="name">Table name.</param>
    /// <param name="reader">Source rows; read to the end but not closed.</param>
    /// <param name="typeOverrides">Optional per-column data types, keyed by column name.</param>
    /// <returns>The table, with every row loaded.</returns>
    public ModelTable AddTable(string name, IDataReader reader, IReadOnlyDictionary<string, ColumnDataType>? typeOverrides = null)
    {
        ArgumentNullException.ThrowIfNull(reader);
        var defs = new List<ColumnDefinition>(reader.FieldCount);
        for (int i = 0; i < reader.FieldCount; i++)
        {
            string col = reader.GetName(i);
            if (string.IsNullOrWhiteSpace(col)) col = "Column" + (i + 1);
            var type = typeOverrides is not null && typeOverrides.TryGetValue(col, out var t)
                ? t
                : ColumnDefinition.FromClrType(reader.GetFieldType(i), _options.DecimalAsCurrency);
            defs.Add(new ColumnDefinition(col, type));
        }
        var table = CreateTable(name, defs);
        table.AddRows(reader);
        return table;
    }

    /// <summary>Adds a table from a <see cref="DataTable"/>, named after <see cref="DataTable.TableName"/> (Table1, Table2, ... when empty).</summary>
    /// <returns>The table, with every row loaded.</returns>
    public ModelTable AddTable(DataTable table, IReadOnlyDictionary<string, ColumnDataType>? typeOverrides = null)
    {
        ArgumentNullException.ThrowIfNull(table);
        using var reader = table.CreateDataReader();
        return AddTable(string.IsNullOrEmpty(table.TableName) ? "Table" + (_tables.Count + 1) : table.TableName, reader, typeOverrides);
    }

    /// <summary>
    /// Adds a table from a sequence of objects; every public readable property becomes a column.
    /// </summary>
    /// <returns>The table, with every row loaded.</returns>
    public ModelTable AddTable<T>(string name, IEnumerable<T> rows, IReadOnlyDictionary<string, ColumnDataType>? typeOverrides = null)
    {
        ArgumentNullException.ThrowIfNull(rows);
        var props = typeof(T).GetProperties(System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance)
            .Where(p => p.CanRead && p.GetIndexParameters().Length == 0)
            .ToArray();
        if (props.Length == 0) throw new ArgumentException($"Type {typeof(T).Name} has no public properties.", nameof(rows));
        var defs = props.Select(p => new ColumnDefinition(p.Name,
            typeOverrides is not null && typeOverrides.TryGetValue(p.Name, out var t)
                ? t
                : ColumnDefinition.FromClrType(p.PropertyType, _options.DecimalAsCurrency))).ToList();
        var getters = props.Select(p => (Func<T, object?>)(item => p.GetValue(item))).ToArray();
        var table = CreateTable(name, defs);
        var values = new object?[props.Length];
        foreach (var item in rows)
        {
            for (int i = 0; i < getters.Length; i++) values[i] = item is null ? null : getters[i](item);
            table.AddRow(values);
        }
        return table;
    }

    /// <summary>
    /// Adds a many-to-one relationship, like the one Power Pivot's diagram view creates by dragging a
    /// column of the "many" table onto the key column of the "one" table:
    /// <c>AddRelationship(sales["ProductKey"], products["ProductKey"])</c>.
    /// </summary>
    /// <param name="manyColumn">Foreign key column of the "many" table (for example Sales[ProductKey]).</param>
    /// <param name="oneColumn">Key column of the "one" table (for example Products[ProductKey]); its values must be unique and not
    /// blank, which is checked when saving.</param>
    /// <returns>The relationship.</returns>
    /// <exception cref="ArgumentException">A column is of another model, both are of one table, the column types differ, the tables
    /// are already related, or the relationship would make the model's relationships circular.</exception>
    public ModelRelationship AddRelationship(ModelColumn manyColumn, ModelColumn oneColumn)
    {
        ThrowIfDisposed();
        CheckOwned(manyColumn, nameof(manyColumn));
        CheckOwned(oneColumn, nameof(oneColumn));
        var many = manyColumn.Table;
        var one = oneColumn.Table;
        many.ThrowIfCalculationTrigger();
        one.ThrowIfCalculationTrigger();
        if (many == one) throw new ArgumentException($"A relationship needs two different tables; {manyColumn} and {oneColumn} are both in '{many.Name}'.", nameof(oneColumn));
        if (manyColumn.DataType != oneColumn.DataType)
            throw new ArgumentException(
                $"{manyColumn} is {manyColumn.DataType} but {oneColumn} is {oneColumn.DataType}; related columns need the same data type.", nameof(oneColumn));
        if (_relationships.Any(r => (r.ManyColumn.Table == many && r.OneColumn.Table == one) || (r.ManyColumn.Table == one && r.OneColumn.Table == many)))
            throw new ArgumentException($"'{many.Name}' and '{one.Name}' are already related.", nameof(oneColumn));
        if (Reaches(one, many))
            throw new ArgumentException($"Relating '{many.Name}' to '{one.Name}' would make the relationships circular.", nameof(oneColumn));
        manyColumn.Builder.IsForeignKey = true;
        oneColumn.Builder.IsPrimaryKey = true;
        _relationshipDefinitions.Add(new RelationshipDefinition
        {
            Name = Guid.NewGuid().ToString("D"),
            From = many.Builder, FromColumn = manyColumn.Builder, To = one.Builder, ToColumn = oneColumn.Builder,
        });
        var relationship = new ModelRelationship(manyColumn, oneColumn);
        _relationships.Add(relationship);
        return relationship;
    }

    /// <summary>
    /// Generates a standard calendar table with one row per day and marks it as the model's date table
    /// (<see cref="ModelTable.MarkAsDateTable"/> still works for tables you load yourself).
    /// </summary>
    /// <remarks>
    /// Columns: Date, Year, QuarterNumber, Quarter ("Q1", sorted by QuarterNumber), MonthNumber, Month (name in the model
    /// culture, sorted by MonthNumber), YearMonthNumber (202503), YearMonth ("Mar 2025", sorted by YearMonthNumber, so a month
    /// is not ambiguous across years), Day, DayOfWeekNumber (1 = Monday), DayOfWeek (name, sorted by DayOfWeekNumber).
    /// A hierarchy named "Calendar" (Year, Quarter, Month) is added.
    /// With a <see cref="DateTableOptions.YearEndMonth"/> other than 12 the table also gets FiscalYear (the calendar year in which
    /// the year ends, so a June year end makes July 2024 part of FiscalYear 2025), FiscalQuarterNumber, FiscalQuarter,
    /// FiscalMonthNumber (1 = first month of that year), FiscalMonth (sorted by FiscalMonthNumber) and a hierarchy
    /// "Fiscal Calendar" (FiscalYear, FiscalQuarter, FiscalMonth). <see cref="DateTableOptions.YearLabel"/> replaces "Fiscal" in these
    /// names, for a year that is not a fiscal year (for example "Financial": FinancialYear, "Financial Calendar").
    /// Relate the table to your fact table yourself:
    /// <c>AddRelationship(sales["Date"], dates["Date"])</c>; marking as date table is only accepted when a relationship
    /// uses the Date column, which is checked when saving.
    /// </remarks>
    /// <param name="name">Table name, for example Dates.</param>
    /// <param name="from">A day in the first year to include.</param>
    /// <param name="to">A day in the last year to include, not before <paramref name="from"/>.</param>
    /// <param name="options">Year end, labels and the optional features (whole-year range, YearMonth, hidden sort columns); defaults when null.</param>
    /// <returns>The table, which is complete: adding rows to it throws <see cref="InvalidOperationException"/>.</returns>
    public ModelTable AddDateTable(string name, DateOnly from, DateOnly to, DateTableOptions? options = null)
    {
        ThrowIfDisposed();
        options ??= new DateTableOptions();
        int yearEndMonth = options.YearEndMonth;
        string yearLabel = options.YearLabel;
        bool hideSortColumns = options.HideSortColumns;
        if (to < from) throw new ArgumentException("The last day must not be before the first day.", nameof(to));
        if (DateTable is { } existing)
            throw new InvalidOperationException($"'{existing.Name}' is already marked as date table; only one date table per model is supported.");
        if (options.CompleteYears) (from, to) = CompleteYears(from, to, yearEndMonth);
        if (to.DayNumber - from.DayNumber > 3_660_000) throw new ArgumentException("A date table can hold at most about 10,000 years.", nameof(to));
        bool fiscal = yearEndMonth != 12;
        var culture = _options.Culture;
        var dtf = culture.DateTimeFormat;

        // Every column with the function that computes its value from the date.
        var spec = new List<(ColumnDefinition column, Func<DateOnly, object> value)>();
        void Add(string column, ColumnDataType type, Func<DateOnly, object> value, string? sortBy = null, bool sortKey = false)
            => spec.Add((new ColumnDefinition(column, type) { SortByColumn = sortBy, IsHidden = sortKey && hideSortColumns }, value));
        spec.Add((new ColumnDefinition("Date", ColumnDataType.DateTime) { Format = options.DateFormat }, d => d));
        Add("Year", ColumnDataType.Int64, d => (long)d.Year);
        Add("QuarterNumber", ColumnDataType.Int64, d => (long)((d.Month - 1) / 3 + 1), sortKey: true);
        Add("Quarter", ColumnDataType.String, d => "Q" + ((d.Month - 1) / 3 + 1), "QuarterNumber");
        Add("MonthNumber", ColumnDataType.Int64, d => (long)d.Month, sortKey: true);
        Add("Month", ColumnDataType.String, d => dtf.GetMonthName(d.Month), "MonthNumber");
        if (options.IncludeYearMonth)
        {
            Add("YearMonthNumber", ColumnDataType.Int64, d => (long)(d.Year * 100 + d.Month), sortKey: true);
            Add("YearMonth", ColumnDataType.String, d => dtf.GetAbbreviatedMonthName(d.Month) + " " + d.Year, "YearMonthNumber");
        }
        Add("Day", ColumnDataType.Int64, d => (long)d.Day);
        Add("DayOfWeekNumber", ColumnDataType.Int64, d => d.DayOfWeek == System.DayOfWeek.Sunday ? 7L : (long)d.DayOfWeek, sortKey: true);
        Add("DayOfWeek", ColumnDataType.String, d => dtf.GetDayName(d.DayOfWeek), "DayOfWeekNumber");
        if (fiscal)
        {
            Add(yearLabel + "Year", ColumnDataType.Int64, d => (long)FiscalPeriod(d, yearEndMonth).year);
            Add(yearLabel + "QuarterNumber", ColumnDataType.Int64, d => (long)FiscalPeriod(d, yearEndMonth).quarter, sortKey: true);
            Add(yearLabel + "Quarter", ColumnDataType.String, d => "Q" + FiscalPeriod(d, yearEndMonth).quarter, yearLabel + "QuarterNumber");
            Add(yearLabel + "MonthNumber", ColumnDataType.Int64, d => (long)FiscalPeriod(d, yearEndMonth).month, sortKey: true);
            // Name of the calendar month that is this fiscal month.
            Add(yearLabel + "Month", ColumnDataType.String, d => dtf.GetMonthName(d.Month), yearLabel + "MonthNumber");
        }

        var table = CreateTable(name, spec.Select(c => c.column).ToList());
        var row = new object?[spec.Count];
        for (int n = from.DayNumber; n <= to.DayNumber; n++)
        {
            var d = DateOnly.FromDayNumber(n);
            for (int i = 0; i < spec.Count; i++) row[i] = spec[i].value(d);
            table.AddRow(row);
        }
        table.AddHierarchy("Calendar", "Year", "Quarter", "Month");
        if (fiscal) table.AddHierarchy(yearLabel + " Calendar", yearLabel + "Year", yearLabel + "Quarter", yearLabel + "Month");
        table.MarkAsDateTable("Date");
        table.Builder.ReadOnlyReason =
            $"Table '{table.Name}' was generated by AddDateTable and is complete; widen its range with the from and to arguments instead of adding rows.";
        return table;
    }

    /// <summary>First day of the year that contains <paramref name="from"/> to last day of the year that contains <paramref name="to"/>.</summary>
    internal static (DateOnly from, DateOnly to) CompleteYears(DateOnly from, DateOnly to, int yearEndMonth)
    {
        if (yearEndMonth == 12) return (new DateOnly(from.Year, 1, 1), new DateOnly(to.Year, 12, 31));
        int first = FiscalPeriod(from, yearEndMonth).year, last = FiscalPeriod(to, yearEndMonth).year;
        return (new DateOnly(first - 1, yearEndMonth + 1, 1), new DateOnly(last, yearEndMonth, DateTime.DaysInMonth(last, yearEndMonth)));
    }

    /// <summary>Financial year (the calendar year in which it ends), month (1 = first month) and quarter of a date.</summary>
    internal static (int year, int month, int quarter) FiscalPeriod(DateOnly d, int yearEndMonth)
    {
        int month = (d.Month - yearEndMonth - 1 + 12) % 12 + 1;
        return (d.Month > yearEndMonth ? d.Year + 1 : d.Year, month, (month - 1) / 3 + 1);
    }

    // True if the graph of relationships (treated as undirected) connects the two tables.
    private bool Reaches(ModelTable from, ModelTable target)
    {
        var seen = new HashSet<ModelTable> { from };
        var queue = new Queue<ModelTable>();
        queue.Enqueue(from);
        while (queue.Count > 0)
        {
            var t = queue.Dequeue();
            if (t == target) return true;
            foreach (var r in _relationships)
            {
                var next = r.ManyColumn.Table == t ? r.OneColumn.Table : r.OneColumn.Table == t ? r.ManyColumn.Table : null;
                if (next is not null && seen.Add(next)) queue.Enqueue(next);
            }
        }
        return false;
    }

    // Finds a table by its model name or by the name it was given (before characters Power Pivot rejects were removed).
    private ModelTable? FindTable(string name)
    {
        string modelName = Identifiers.SanitizeName(name, "Table");
        return _tables.FirstOrDefault(t => string.Equals(t.Name, name, StringComparison.OrdinalIgnoreCase))
            ?? _tables.FirstOrDefault(t => string.Equals(t.Name, modelName, StringComparison.OrdinalIgnoreCase));
    }

    private void CheckOwned(ModelColumn column, string paramName)
    {
        ArgumentNullException.ThrowIfNull(column, paramName);
        if (column.Table.Model != this)
            throw new ArgumentException($"Column {column} belongs to another DataModelWriter.", paramName);
    }

    /// <summary>Encodes the model and writes the workbook to a file.</summary>
    /// <remarks>
    /// The workbook is written to a temporary file in the same directory and moved over <paramref name="path"/> only when it is
    /// complete, so a failed save (for example a duplicate key, reported as <see cref="InvalidOperationException"/>) leaves an
    /// existing file untouched.
    /// </remarks>
    public void Save(string path)
    {
        ArgumentException.ThrowIfNullOrEmpty(path);
        string full = Path.GetFullPath(path);
        string temp = Path.Combine(Path.GetDirectoryName(full)!, "~" + Path.GetFileName(full) + "." + Guid.NewGuid().ToString("N")[..8] + ".tmp");
        try
        {
            using (var fs = new FileStream(temp, FileMode.CreateNew, FileAccess.ReadWrite, FileShare.None, 1 << 16))
                Save(fs);
            File.Move(temp, full, overwrite: true);
        }
        finally
        {
            if (File.Exists(temp)) File.Delete(temp);
        }
    }

    /// <summary>
    /// Adds a one-row table that Excel treats as a linked worksheet table, which makes Excel calculate every calculated column when the
    /// workbook opens (<see cref="ModelTable.AddCalculatedColumn"/> only stores the formula). Without it a calculated column works
    /// as a PivotTable value but not as a row, column, filter or hierarchy level until it has been calculated once. The table has one
    /// column, Key, with the value 1, and a worksheet "Sheet2" holds the table it is linked to; Excel refreshes it on open and recalculates
    /// the model, and shows its usual Enable External Content bar for the workbook's data connections. Do not delete the sheet or its table.
    /// </summary>
    /// <param name="name">Table name: letters, digits and underscores, starting with a letter. It is also the worksheet table's name, so it
    /// must not be the name of another table of the model. Default Trigger.</param>
    /// <returns>The table. It is complete (adding rows throws <see cref="InvalidOperationException"/>) and takes no measures, columns,
    /// hierarchies or relationships; <see cref="ModelTable.IsCalculationTrigger"/> is true.</returns>
    /// <exception cref="ArgumentException">The name is not letters, digits and underscores starting with a letter, or a table of that name exists.</exception>
    /// <exception cref="InvalidOperationException">The model already has a calculation trigger.</exception>
    public ModelTable AddCalculationTrigger(string name = "Trigger")
    {
        ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(name);
        if (name.Length == 0 || !char.IsAsciiLetter(name[0]) || !name.All(c => char.IsAsciiLetterOrDigit(c) || c == '_'))
            throw new ArgumentException($"'{name}' is not a valid trigger table name; use letters, digits and underscores, starting with a letter (for example \"Trigger\").", nameof(name));
        if (_tables.FirstOrDefault(t => t.IsCalculationTrigger) is { } existing)
            throw new InvalidOperationException($"'{existing.Name}' is already the calculation trigger; a model needs only one.");
        if (_tables.Any(t => string.Equals(t.Name, name, StringComparison.OrdinalIgnoreCase)))
            throw new ArgumentException($"A table named '{name}' already exists.", nameof(name));
        var builder = new TableBuilder(name, name, new[] { new ColumnDefinition("Key", ColumnDataType.Int64) }, _ctx) { IsLinked = true };
        var table = new ModelTable(this, builder);
        _tables.Add(table);
        table.AddRow(1L);
        builder.ReadOnlyReason = $"Table '{name}' was made by AddCalculationTrigger and is complete.";
        return table;
    }

    /// <summary>Encodes the model and writes the workbook to a stream.</summary>
    /// <remarks>Nothing is written to <paramref name="output"/> unless the model is valid; the stream is left open.</remarks>
    public void Save(Stream output)
    {
        ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(output);
        ThrowIfEmpty();
        string dir = _options.TempDirectory ?? Path.GetTempPath();
        string tempPath = Path.Combine(dir, "ppw-model-" + Guid.NewGuid().ToString("N") + ".tmp");
        using var model = new FileStream(tempPath, FileMode.CreateNew, FileAccess.ReadWrite, FileShare.None,
            1 << 20, FileOptions.DeleteOnClose);
        var serializer = CreateSerializer();
        serializer.Write(model);
        XlsxWriter.Write(output, model, serializer.Info, serializer.EncodedTables, _options);
    }

    /// <summary>
    /// Writes only the data model part (the content of xl/model/item.data) to a stream. Useful for
    /// injecting the model into an existing workbook.
    /// </summary>
    public void SaveModelPart(Stream output)
    {
        ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(output);
        if (!output.CanSeek) throw new ArgumentException("The output stream must be seekable.", nameof(output));
        ThrowIfEmpty();
        CreateSerializer().Write(output);
    }

    private ModelSerializer CreateSerializer()
        => new(_options, _tables.Select(t => t.Builder).ToList(), _relationshipDefinitions, _ctx.StringComparer, MeasureDefinitions, HierarchyDefinitions);

    private void ThrowIfEmpty()
    {
        if (_tables.Count == 0) throw new InvalidOperationException("Add at least one table before saving.");
    }

    internal void ThrowIfDisposed() => ObjectDisposedException.ThrowIf(_disposed, this);

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;
        foreach (var t in _tables) t.Builder.Dispose();
    }
}
