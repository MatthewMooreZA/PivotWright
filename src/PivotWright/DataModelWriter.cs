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
/// <example>
/// <code>
/// using var model = new DataModelWriter();
/// var sales = model.AddTable("Sales", sqlCommand.ExecuteReader());
/// model.AddMeasure(sales, "Total Qty", "SUM(Sales[Qty])");
/// model.Save("sales.xlsx");
/// </code>
/// Every method that takes a table accepts either its name (as given or as stored, ignoring case) or the
/// <see cref="TableWriter"/> returned when the table was added.
/// </example>
public sealed class DataModelWriter : IDisposable
{
    private readonly DataModelOptions _options;
    private readonly BuilderContext _ctx;
    private readonly List<TableBuilder> _tables = new();
    private readonly List<RelationshipDefinition> _relationships = new();
    private readonly List<MeasureDefinition> _measures = new();
    private readonly List<HierarchyDefinition> _hierarchies = new();
    private readonly HashSet<string> _tableIds = new(StringComparer.OrdinalIgnoreCase);
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
    }

    /// <summary>Names of the tables added so far.</summary>
    public IEnumerable<string> TableNames => _tables.Select(t => t.Name);

    /// <summary>Creates an empty table; add rows through the returned <see cref="TableWriter"/>.</summary>
    /// <param name="name">Table name. Characters Power Pivot rejects are removed (see <see cref="TableWriter.Name"/>); later
    /// calls may use either form.</param>
    /// <param name="columns">The columns, in order.</param>
    public TableWriter CreateTable(string name, IEnumerable<ColumnDefinition> columns)
    {
        ThrowIfDisposed();
        if (string.IsNullOrWhiteSpace(name)) throw new ArgumentException("Table name is required.", nameof(name));
        ArgumentNullException.ThrowIfNull(columns);
        string modelName = Identifiers.SanitizeName(name, "Table");
        if (_tables.Any(t => string.Equals(t.Name, modelName, StringComparison.OrdinalIgnoreCase)))
            throw new ArgumentException($"A table named '{modelName}' already exists.", nameof(name));
        string id = Identifiers.Sanitize(modelName, "Table") + "_" + Guid.NewGuid().ToString("D");
        _tableIds.Add(id);
        var builder = new TableBuilder(modelName, id, columns, _ctx);
        _tables.Add(builder);
        return new TableWriter(builder);
    }

    /// <summary>
    /// Adds a table and loads every row of <paramref name="reader"/> into it. Column types are taken
    /// from the reader's field types unless overridden.
    /// </summary>
    /// <param name="name">Table name.</param>
    /// <param name="reader">Source rows; read to the end but not closed.</param>
    /// <param name="typeOverrides">Optional per-column data types, keyed by column name.</param>
    /// <returns>The table, with every row loaded (<see cref="TableWriter.RowCount"/>); pass it to the other methods in place of its name.</returns>
    public TableWriter AddTable(string name, IDataReader reader, IReadOnlyDictionary<string, ColumnDataType>? typeOverrides = null)
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
    public TableWriter AddTable(DataTable table, IReadOnlyDictionary<string, ColumnDataType>? typeOverrides = null)
    {
        ArgumentNullException.ThrowIfNull(table);
        using var reader = table.CreateDataReader();
        return AddTable(string.IsNullOrEmpty(table.TableName) ? "Table" + (_tables.Count + 1) : table.TableName, reader, typeOverrides);
    }

    /// <summary>
    /// Adds a table from a sequence of objects; every public readable property becomes a column.
    /// </summary>
    /// <returns>The table, with every row loaded.</returns>
    public TableWriter AddTable<T>(string name, IEnumerable<T> rows, IReadOnlyDictionary<string, ColumnDataType>? typeOverrides = null)
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

    /// <summary>Number of relationships added so far.</summary>
    public int RelationshipCount => _relationships.Count;

    /// <summary>
    /// Adds a many-to-one relationship, like the one Power Pivot's diagram view creates by dragging a
    /// column of the "many" table onto the key column of the "one" table. Both tables must already exist.
    /// </summary>
    /// <param name="manyTable">Table holding the foreign key (for example Sales).</param>
    /// <param name="manyColumn">Foreign key column (for example ProductKey).</param>
    /// <param name="oneTable">Table holding the primary key (for example Products).</param>
    /// <param name="oneColumn">Primary key column; its values must be unique and not blank, which is checked when saving.</param>
    /// <exception cref="ArgumentException">A table or column is unknown, the column types differ, the tables
    /// are already related, or the relationship would make the model's relationships circular.</exception>
    public void AddRelationship(string manyTable, string manyColumn, string oneTable, string oneColumn)
    {
        ThrowIfDisposed();
        var many = FindTable(manyTable, nameof(manyTable));
        var one = FindTable(oneTable, nameof(oneTable));
        if (many == one) throw new ArgumentException("A relationship needs two different tables.", nameof(oneTable));
        var fk = many.FindColumn(manyColumn, nameof(manyColumn));
        var pk = one.FindColumn(oneColumn, nameof(oneColumn));
        if (fk.Definition.DataType != pk.Definition.DataType)
            throw new ArgumentException(
                $"'{many.Name}[{fk.Definition.Name}]' is {fk.Definition.DataType} but '{one.Name}[{pk.Definition.Name}]' is {pk.Definition.DataType}; related columns need the same data type.");
        if (_relationships.Any(r => (r.From == many && r.To == one) || (r.From == one && r.To == many)))
            throw new ArgumentException($"'{many.Name}' and '{one.Name}' are already related.");
        if (Reaches(one, many))
            throw new ArgumentException($"Relating '{many.Name}' to '{one.Name}' would make the relationships circular.");
        fk.IsForeignKey = true;
        pk.IsPrimaryKey = true;
        _relationships.Add(new RelationshipDefinition
        {
            Name = Guid.NewGuid().ToString("D"),
            From = many, FromColumn = fk, To = one, ToColumn = pk,
        });
    }

    /// <summary>
    /// Orders a column by another column of the same table, like Power Pivot's "Sort by Column": a Month
    /// column sorted by a MonthNumber column lists January to December instead of alphabetically. Every
    /// value of <paramref name="column"/> must belong to exactly one value of <paramref name="sortColumn"/>,
    /// which is checked when saving. Can also be set with <see cref="ColumnDefinition.SortByColumn"/>.
    /// </summary>
    public void SetSortByColumn(string table, string column, string sortColumn)
    {
        ThrowIfDisposed();
        FindTable(table, nameof(table)).SetSortBy(column, sortColumn);
    }

    /// <summary>Overload taking the table returned when it was added.</summary>
    public void SetSortByColumn(TableWriter table, string column, string sortColumn) => SetSortByColumn(NameOf(table), column, sortColumn);

    /// <summary>
    /// Hides columns from client tools, like Power Pivot's "Hide from Client Tools": they stay in the model (relationships,
    /// sorting, hierarchies and DAX can still use them) but are not listed in the PivotTable field list. The usual treatment
    /// for keys and sort-order columns. Same as <see cref="ColumnDefinition.IsHidden"/>, for tables loaded with
    /// <see cref="AddTable(string, IDataReader, IReadOnlyDictionary{string, ColumnDataType}?)"/>.
    /// </summary>
    /// <param name="table">Table that holds the columns.</param>
    /// <param name="columns">One or more column names.</param>
    /// <exception cref="ArgumentException">The table or a column is unknown, or no column is given.</exception>
    public void HideColumns(string table, params string[] columns)
    {
        ThrowIfDisposed();
        var t = FindTable(table, nameof(table));
        ArgumentNullException.ThrowIfNull(columns);
        if (columns.Length == 0) throw new ArgumentException("Name at least one column to hide.", nameof(columns));
        foreach (var c in columns) t.FindColumn(c, nameof(columns));
        foreach (var c in columns) t.SetHidden(c, hidden: true);
    }

    /// <summary>Overload taking the table returned when it was added.</summary>
    public void HideColumns(TableWriter table, params string[] columns) => HideColumns(NameOf(table), columns);

    /// <summary>Number of hierarchies added so far.</summary>
    public int HierarchyCount => _hierarchies.Count;

    /// <summary>
    /// Adds a hierarchy to a table, like the one Power Pivot's diagram view creates with "Create Hierarchy":
    /// an ordered drill-down path such as Category, Subcategory, Product that PivotTables offer as a single field.
    /// </summary>
    /// <param name="table">Table that holds the level columns.</param>
    /// <param name="name">Hierarchy name; unique within the table and different from its column names.</param>
    /// <param name="levels">Two or more columns of the table, from the top level down.</param>
    public void AddHierarchy(string table, string name, params string[] levels)
    {
        ThrowIfDisposed();
        var t = FindTable(table, nameof(table));
        ArgumentNullException.ThrowIfNull(levels);
        if (string.IsNullOrWhiteSpace(name)) throw new ArgumentException("Hierarchy name is required.", nameof(name));
        string modelName = Identifiers.SanitizeName(name, "Hierarchy");
        if (levels.Length < 2) throw new ArgumentException("A hierarchy needs at least two levels.", nameof(levels));
        if (_hierarchies.Any(h => h.Table == t && string.Equals(h.Name, modelName, StringComparison.OrdinalIgnoreCase)))
            throw new ArgumentException($"Table '{t.Name}' already has a hierarchy named '{modelName}'.", nameof(name));
        if (t.Columns.Any(c => string.Equals(c.ModelName, modelName, StringComparison.OrdinalIgnoreCase)))
            throw new ArgumentException($"Table '{t.Name}' already has a column named '{modelName}'; a hierarchy needs its own name.", nameof(name));
        var columns = new List<ColumnBuilder>();
        foreach (var level in levels)
        {
            var c = t.FindColumn(level, nameof(levels));
            if (columns.Contains(c)) throw new ArgumentException($"Column '{c.Definition.Name}' is used twice in the hierarchy.", nameof(levels));
            columns.Add(c);
        }
        _hierarchies.Add(new HierarchyDefinition { Name = modelName, Table = t, Columns = columns });
    }

    /// <summary>Overload taking the table returned when it was added.</summary>
    public void AddHierarchy(TableWriter table, string name, params string[] levels) => AddHierarchy(NameOf(table), name, levels);

    /// <summary>Name of the table marked with <see cref="MarkAsDateTable(string, string)"/>, or null.</summary>
    public string? DateTableName => _tables.FirstOrDefault(t => t.DateColumn is not null)?.Name;

    /// <summary>
    /// Marks a table as the model's date table, like Power Pivot's "Mark as Date Table": the date column becomes the
    /// table's key, which Excel's time intelligence functions (TOTALYTD, SAMEPERIODLASTYEAR, ...) need.
    /// </summary>
    /// <remarks>
    /// Written the way Excel saves a model it just marked; only that case has been seen in an Excel-written file, so
    /// <see cref="Save(string)"/> requires a hierarchy on the table, a relationship that uses the date column as its
    /// primary key, and a date on every row (whole days, no duplicates). One date table per model.
    /// </remarks>
    /// <param name="table">Table that holds the dates (for example Dates).</param>
    /// <param name="dateColumn">Column of type <see cref="ColumnDataType.DateTime"/> with one row per day.</param>
    /// <exception cref="ArgumentException">The table or column is unknown, the column is not a date column, or it is sorted by another column.</exception>
    /// <exception cref="InvalidOperationException">A date table was already marked.</exception>
    public void MarkAsDateTable(string table, string dateColumn)
    {
        ThrowIfDisposed();
        var t = FindTable(table, nameof(table));
        var c = t.FindColumn(dateColumn, nameof(dateColumn));
        if (c.Definition.DataType != ColumnDataType.DateTime)
            throw new ArgumentException($"'{t.Name}[{c.Definition.Name}]' is {c.Definition.DataType}; a date table needs a DateTime column.", nameof(dateColumn));
        if (c.SortBy is not null || t.Columns.Any(o => o.SortBy == c))
            throw new ArgumentException($"'{t.Name}[{c.Definition.Name}]' takes part in Sort by Column, so it cannot be the date column.", nameof(dateColumn));
        if (DateTableName is { } existing)
            throw new InvalidOperationException($"'{existing}' is already marked as date table; only one date table per model is supported.");
        t.DateColumn = c;
        c.IsDateKey = true;
    }

    /// <summary>Overload taking the table returned when it was added.</summary>
    public void MarkAsDateTable(TableWriter table, string dateColumn) => MarkAsDateTable(NameOf(table), dateColumn);

    /// <summary>
    /// Generates a standard calendar table with one row per day and marks it as the model's date table
    /// (<see cref="MarkAsDateTable(string, string)"/> still works for tables you load yourself).
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
    /// <c>AddRelationship("Sales", "Date", name, "Date")</c>; marking as date table is only accepted when a relationship
    /// uses the Date column, which is checked when saving.
    /// </remarks>
    /// <param name="name">Table name, for example Dates.</param>
    /// <param name="from">A day in the first year to include.</param>
    /// <param name="to">A day in the last year to include, not before <paramref name="from"/>.</param>
    /// <param name="options">Year end, labels and the optional features (whole-year range, YearMonth, hidden sort columns); defaults when null.</param>
    /// <returns>The table, which is complete: adding rows to it throws <see cref="InvalidOperationException"/>.</returns>
    public TableWriter AddDateTable(string name, DateOnly from, DateOnly to, DateTableOptions? options = null)
    {
        ThrowIfDisposed();
        options ??= new DateTableOptions();
        int yearEndMonth = options.YearEndMonth;
        string yearLabel = options.YearLabel;
        bool hideSortColumns = options.HideSortColumns;
        if (to < from) throw new ArgumentException("The last day must not be before the first day.", nameof(to));
        if (DateTableName is { } existing)
            throw new InvalidOperationException($"'{existing}' is already marked as date table; only one date table per model is supported.");
        if (options.CompleteYears) (from, to) = CompleteYears(from, to, yearEndMonth);
        if (to.DayNumber - from.DayNumber > 3_660_000) throw new ArgumentException("A date table can hold at most about 10,000 years.", nameof(to));
        bool fiscal = yearEndMonth != 12;
        var culture = _options.Culture;
        var dtf = culture.DateTimeFormat;

        // Every column with the function that computes its value from the date.
        var spec = new List<(ColumnDefinition column, Func<DateOnly, object> value)>();
        void Add(string column, ColumnDataType type, Func<DateOnly, object> value, string? sortBy = null, bool sortKey = false)
            => spec.Add((new ColumnDefinition(column, type) { SortByColumn = sortBy, IsHidden = sortKey && hideSortColumns }, value));
        Add("Date", ColumnDataType.DateTime, d => d);
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
        AddHierarchy(table.Name, "Calendar", "Year", "Quarter", "Month");
        if (fiscal) AddHierarchy(table.Name, yearLabel + " Calendar", yearLabel + "Year", yearLabel + "Quarter", yearLabel + "Month");
        MarkAsDateTable(table.Name, "Date");
        FindTable(table.Name, nameof(name)).ReadOnlyReason =
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

    /// <summary>Number of measures added so far.</summary>
    public int MeasureCount => _measures.Count;

    /// <summary>
    /// Adds a DAX measure, like one created in Power Pivot's "New Measure" dialog. The expression is
    /// stored as given and evaluated by Excel when the model is opened; the library does not parse it,
    /// so a mistake shows up as an error in the pivot table, not here.
    /// </summary>
    /// <param name="table">Table the measure belongs to (it appears under this table in the field list).</param>
    /// <param name="name">Measure name; unique across the model, and must not contain [ or ].</param>
    /// <param name="expression">DAX expression without the leading "=", for example <c>SUM(Sales[Qty])</c>.</param>
    /// <param name="format">Number format; <see cref="MeasureFormat.General"/> when omitted.</param>
    /// <param name="hidden">Hide the measure from client tools.</param>
    public void AddMeasure(string table, string name, string expression, MeasureFormat? format = null, bool hidden = false)
    {
        ThrowIfDisposed();
        var t = FindTable(table, nameof(table));
        if (string.IsNullOrWhiteSpace(name)) throw new ArgumentException("Measure name is required.", nameof(name));
        name = name.Trim();
        if (name.IndexOfAny(new[] { '[', ']' }) >= 0 || name.Any(char.IsControl))
            throw new ArgumentException($"Measure name '{name}' must not contain [, ] or control characters.", nameof(name));
        if (name.StartsWith("__", StringComparison.Ordinal))
            throw new ArgumentException("Measure names starting with '__' are reserved.", nameof(name));
        if (string.IsNullOrWhiteSpace(expression)) throw new ArgumentException("A DAX expression is required.", nameof(expression));
        expression = expression.Trim();
        if (expression.StartsWith('=')) expression = expression[1..].TrimStart();
        if (_measures.Any(m => string.Equals(m.Name, name, StringComparison.OrdinalIgnoreCase)))
            throw new ArgumentException($"A measure named '{name}' already exists.", nameof(name));
        foreach (var tb in _tables)
            if (tb.Columns.Any(c => string.Equals(c.Definition.Name, name, StringComparison.OrdinalIgnoreCase)
                                    || string.Equals(c.ModelName, name, StringComparison.OrdinalIgnoreCase)) && tb == t)
                throw new ArgumentException($"Table '{t.Name}' already has a column named '{name}'.", nameof(name));
        _measures.Add(new MeasureDefinition
        {
            Name = name, Table = t.Name, Expression = expression,
            Format = format ?? MeasureFormat.General, Hidden = hidden,
        });
    }

    /// <summary>Overload taking the table returned when it was added.</summary>
    public void AddMeasure(TableWriter table, string name, string expression, MeasureFormat? format = null, bool hidden = false)
        => AddMeasure(NameOf(table), name, expression, format, hidden);

    /// <summary>Overload taking the tables returned when they were added.</summary>
    public void AddRelationship(TableWriter manyTable, string manyColumn, TableWriter oneTable, string oneColumn)
        => AddRelationship(NameOf(manyTable, nameof(manyTable)), manyColumn, NameOf(oneTable, nameof(oneTable)), oneColumn);

    // True if the graph of relationships (treated as undirected) connects the two tables.
    private bool Reaches(TableBuilder from, TableBuilder target)
    {
        var seen = new HashSet<TableBuilder> { from };
        var queue = new Queue<TableBuilder>();
        queue.Enqueue(from);
        while (queue.Count > 0)
        {
            var t = queue.Dequeue();
            if (t == target) return true;
            foreach (var r in _relationships)
            {
                var next = r.From == t ? r.To : r.To == t ? r.From : null;
                if (next is not null && seen.Add(next)) queue.Enqueue(next);
            }
        }
        return false;
    }

    // Finds a table by its model name or by the name it was given (before characters Power Pivot rejects were removed).
    private TableBuilder FindTable(string name, string paramName)
    {
        ArgumentNullException.ThrowIfNull(name, paramName);
        string modelName = Identifiers.SanitizeName(name, "Table");
        return _tables.FirstOrDefault(t => string.Equals(t.Name, name, StringComparison.OrdinalIgnoreCase))
            ?? _tables.FirstOrDefault(t => string.Equals(t.Name, modelName, StringComparison.OrdinalIgnoreCase))
            ?? throw new ArgumentException($"There is no table named '{name}'. The model's tables are: {Identifiers.List(TableNames)}.", paramName);
    }

    // The name of a table passed as a TableWriter, which must belong to this model.
    private string NameOf(TableWriter table, string paramName = "table")
    {
        ArgumentNullException.ThrowIfNull(table, paramName);
        if (!_tables.Contains(table.Builder))
            throw new ArgumentException($"Table '{table.Name}' belongs to another DataModelWriter.", paramName);
        return table.Name;
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
        var serializer = new ModelSerializer(_options, _tables, _relationships, _ctx.StringComparer, _measures, _hierarchies);
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
        new ModelSerializer(_options, _tables, _relationships, _ctx.StringComparer, _measures, _hierarchies).Write(output);
    }

    private void ThrowIfEmpty()
    {
        if (_tables.Count == 0) throw new InvalidOperationException("Add at least one table before saving.");
    }

    private void ThrowIfDisposed() => ObjectDisposedException.ThrowIf(_disposed, this);

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;
        foreach (var t in _tables) t.Dispose();
    }
}
