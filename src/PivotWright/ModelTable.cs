using System.Data;
using System.Globalization;
using PivotWright.Internal.Model;

namespace PivotWright;

/// <summary>
/// A table of the model, returned by <see cref="DataModelWriter.CreateTable"/>, the <c>AddTable</c> overloads and
/// <see cref="DataModelWriter.AddDateTable"/>, or found with <c>model.Tables["Sales"]</c>. It takes rows, and holds the
/// table's columns, measures and hierarchies.
/// </summary>
/// <example>
/// <code>
/// var sales = model.AddTable("Sales", reader);
/// sales["ProductKey"].IsHidden = true;
/// sales.AddMeasure("Total Qty", "SUM(Sales[Qty])", MeasureFormat.WholeNumber);
/// </code>
/// </example>
public sealed class ModelTable
{
    private readonly List<ModelColumn> _columns;
    private readonly List<ModelMeasure> _measures = new();
    private readonly List<ModelHierarchy> _hierarchies = new();
    private readonly List<ModelCalculatedColumn> _calculatedColumns = new();

    internal ModelTable(DataModelWriter model, TableBuilder builder)
    {
        Model = model;
        Builder = builder;
        _columns = builder.Columns.Select(c => new ModelColumn(this, c)).ToList();
        Columns = new ModelCollection<ModelColumn>(_columns,
            name => _columns.FirstOrDefault(c => string.Equals(c.Builder.Definition.Name, name, StringComparison.OrdinalIgnoreCase)
                                                 || string.Equals(c.Name, name, StringComparison.OrdinalIgnoreCase)),
            name => $"Table '{Name}' has no column named '{name}'. Its columns are: {Identifiers.List(_columns.Select(c => c.Name))}.");
    }

    internal DataModelWriter Model { get; }
    internal TableBuilder Builder { get; }

    /// <summary>Table name as it appears in the model (the name given, without the characters Power Pivot rejects).</summary>
    public string Name => Builder.Name;

    /// <summary>The columns, in order; <c>Columns["Qty"]</c> finds one by name.</summary>
    public ModelCollection<ModelColumn> Columns { get; }

    /// <summary>The column with this name (as given or as stored, ignoring case); same as <c>Columns[name]</c>.</summary>
    /// <exception cref="KeyNotFoundException">The table has no such column; the message lists its columns.</exception>
    public ModelColumn this[string column] => Columns[column];

    /// <summary>Number of rows added so far.</summary>
    public long RowCount => Builder.Rows;

    /// <summary>Measures added to this table, in order.</summary>
    public IReadOnlyList<ModelMeasure> Measures => _measures;

    /// <summary>Calculated columns added to this table, in order. They are not in <see cref="Columns"/>, which holds the columns that take row values.</summary>
    public IReadOnlyList<ModelCalculatedColumn> CalculatedColumns => _calculatedColumns;

    /// <summary>Hierarchies added to this table, in order.</summary>
    public IReadOnlyList<ModelHierarchy> Hierarchies => _hierarchies;

    /// <summary>True when the table is the model's date table (<see cref="MarkAsDateTable"/>).</summary>
    public bool IsDateTable => Builder.DateColumn is not null;

    /// <summary>
    /// Adds a row, one value per column in column order. Values are converted to the column types; <c>null</c> and
    /// <see cref="DBNull"/> become BLANK.
    /// </summary>
    /// <exception cref="ArgumentException">The number of values differs from the number of columns, or a value cannot be
    /// converted to its column's type. The message names the row, column and value; the row is not added, so the table
    /// stays consistent and further rows can be added.</exception>
    /// <exception cref="InvalidOperationException">The table is complete (a table made by <see cref="DataModelWriter.AddDateTable"/>).</exception>
    public void AddRow(params object?[] values)
    {
        Model.ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(values);
        Builder.AddRow(values);
    }

    /// <summary>Adds every row of a data reader. The reader's columns must match this table's columns by position.</summary>
    /// <remarks>
    /// If a value cannot be converted, or the reader fails, in the middle of a row, the exception is passed on and the table
    /// is left incomplete: saving the model then throws, naming the row. Load the table again from a corrected source.
    /// </remarks>
    public void AddRows(IDataReader reader)
    {
        Model.ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(reader);
        Builder.ThrowIfReadOnly();
        var columns = Builder.Columns;
        if (reader.FieldCount != columns.Count)
            throw new ArgumentException($"The reader has {reader.FieldCount} columns but table '{Name}' has {columns.Count} ({Identifiers.List(columns.Select(c => c.Definition.Name))}); they are matched by position.", nameof(reader));
        var appenders = new Action<IDataRecord>[columns.Count];
        for (int i = 0; i < appenders.Length; i++)
            appenders[i] = CreateAppender(Builder, i, reader.GetFieldType(i));
        while (reader.Read())
        {
            int done = 0;
            try
            {
                for (; done < appenders.Length; done++) appenders[done](reader);
            }
            catch (Exception e) when (done > 0)
            {
                Builder.Fault($"Table '{Name}' is incomplete: loading row {Builder.Rows + 1} from the data reader failed at column '{columns[done].Definition.Name}' ({e.Message}). Load the table again from a corrected source.");
                throw;
            }
            Builder.CommitRow();
        }
    }

    /// <summary>Adds rows from objects using a projection to column values, one value per column in column order.</summary>
    public void AddRows<T>(IEnumerable<T> items, Func<T, object?[]> project)
    {
        Model.ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(items);
        ArgumentNullException.ThrowIfNull(project);
        foreach (var item in items) Builder.AddRow(project(item));
    }

    /// <summary>
    /// Adds a DAX measure to this table, like one created in Power Pivot's "New Measure" dialog. The expression is stored as
    /// given and evaluated by Excel when the model is opened; the library does not parse it, so a mistake shows up as an error
    /// in the PivotTable, not here. Hide it with <see cref="ModelMeasure.IsHidden"/>.
    /// </summary>
    /// <param name="name">Measure name; unique across the model, different from this table's column names, and without [ or ].</param>
    /// <param name="expression">DAX expression, with or without the leading "=", for example <c>SUM(Sales[Qty])</c>. Brackets, parentheses, braces and quotes must be paired; nothing else is checked.</param>
    /// <param name="format">Number format; <see cref="MeasureFormat.General"/> when omitted.</param>
    /// <returns>The measure.</returns>
    /// <exception cref="ArgumentException">The name is empty, reserved, taken or contains [, ] or control characters, or the expression is empty or malformed (unpaired brackets, parentheses or quotes).</exception>
    public ModelMeasure AddMeasure(string name, string expression, MeasureFormat? format = null)
    {
        Model.ThrowIfDisposed();
        ThrowIfCalculationTrigger();
        if (string.IsNullOrWhiteSpace(name)) throw new ArgumentException("Measure name is required.", nameof(name));
        name = name.Trim();
        if (name.IndexOfAny(new[] { '[', ']' }) >= 0 || name.Any(char.IsControl))
            throw new ArgumentException($"Measure name '{name}' must not contain [, ] or control characters.", nameof(name));
        if (name.StartsWith("__", StringComparison.Ordinal))
            throw new ArgumentException("Measure names starting with '__' are reserved.", nameof(name));
        if (string.IsNullOrWhiteSpace(expression)) throw new ArgumentException("A DAX expression is required.", nameof(expression));
        expression = expression.Trim();
        if (expression.StartsWith('=')) expression = expression[1..].TrimStart();
        PivotWright.Internal.DaxText.Check(expression, nameof(expression));
        if (Model.Tables.SelectMany(t => t.Measures).FirstOrDefault(m => string.Equals(m.Name, name, StringComparison.OrdinalIgnoreCase)) is { } taken)
            throw new ArgumentException($"A measure named '{name}' already exists ({taken}); measure names are unique across the model.", nameof(name));
        if (Columns.Contains(name) || _calculatedColumns.Any(c => string.Equals(c.Name, name, StringComparison.OrdinalIgnoreCase)))
            throw new ArgumentException($"Table '{Name}' already has a column named '{name}'.", nameof(name));
        if (format is { KindValue: MeasureFormat.Kind.LongDate or MeasureFormat.Kind.Date } or { KindValue: MeasureFormat.Kind.ShortDate, Culture: not null })
            throw new ArgumentException("A measure can only take the short date of the model's culture, MeasureFormat.ShortDate(); other date formats (including MeasureFormat.Date) were never seen for measures.", nameof(format));
        var definition = new MeasureDefinition { Name = name, Table = Name, Expression = expression, Format = format ?? MeasureFormat.General };
        Model.MeasureDefinitions.Add(definition);
        var measure = new ModelMeasure(this, definition);
        _measures.Add(measure);
        return measure;
    }

    /// <summary>
    /// Adds a calculated column, like one created in Power Pivot by typing a formula into the empty column. The file holds only the
    /// formula, not the values: Excel calculates them the first time a PivotTable uses the column. The library does not parse the
    /// expression, so a mistake shows up in Excel, not here.
    /// </summary>
    /// <remarks>
    /// <b>Limitation:</b> until Excel has calculated the column it can be a value in a PivotTable but not a row, column or filter
    /// field. The user puts it in Values once, or clicks Calculate Now in Power Pivot; Refresh All does not calculate it.
    /// <para>
    /// Written the way Excel saves calculated columns it has not calculated yet; only some shapes have been seen in Excel-written
    /// files, so saving requires a model with relationships, no hidden columns, a column on the date table only if it is the one primary-side table with a column, hidden columns (samples exist for the hierarchy model without a date table; the other models extrapolate), and
    /// any number of columns on the many-side table and any number on each table that is the primary side of a relationship (with a hierarchy, or in a model without hierarchies). Other combinations make <c>Save</c> throw.
    /// </para>
    /// </remarks>
    /// <param name="name">Column name; unique among this table's columns, calculated columns, measures and hierarchies, and without [ or ].</param>
    /// <param name="expression">DAX expression, with or without the leading "=", for example <c>[Qty]*[Amount]</c>.</param>
    /// <param name="dataType">Data type of the formula's result, which Excel infers from the formula when the column is typed in; give the
    /// type the formula produces (for example <see cref="ColumnDataType.Currency"/> for <c>[Qty]*[Amount]</c> with a Currency Amount).</param>
    /// <param name="format">Number or date format of the column, like Power Pivot's Format command, for example <c>MeasureFormat.Decimal(3)</c>,
    /// <c>MeasureFormat.Currency(2, new CultureInfo("en-US"))</c> or <c>MeasureFormat.CurrencyWithoutSymbol()</c>; it must fit <paramref name="dataType"/>.
    /// Null keeps the one Excel gives the type (a general date, the model culture's currency, General).</param>
    /// <returns>The calculated column.</returns>
    /// <exception cref="ArgumentException">The name is empty, reserved, taken or contains [, ] or control characters; the expression
    /// is empty; or the data type is not a <see cref="ColumnDataType"/>.</exception>
    public ModelCalculatedColumn AddCalculatedColumn(string name, string expression, ColumnDataType dataType, MeasureFormat? format = null)
    {
        Model.ThrowIfDisposed();
        ThrowIfCalculationTrigger();
        if (string.IsNullOrWhiteSpace(name)) throw new ArgumentException("Column name is required.", nameof(name));
        if (name.IndexOfAny(new[] { '[', ']' }) >= 0)
            throw new ArgumentException($"Column name '{name}' must not contain [ or ].", nameof(name));
        name = Identifiers.SanitizeName(name, "Column");
        if (name.StartsWith("__", StringComparison.Ordinal))
            throw new ArgumentException("Column names starting with '__' are reserved.", nameof(name));
        if (string.IsNullOrWhiteSpace(expression)) throw new ArgumentException("A DAX expression is required.", nameof(expression));
        expression = expression.Trim();
        if (expression.StartsWith('=')) expression = expression[1..].TrimStart();
        if (expression.Length == 0) throw new ArgumentException("A DAX expression is required.", nameof(expression));
        PivotWright.Internal.DaxText.Check(expression, nameof(expression));
        if (!Enum.IsDefined(dataType)) throw new ArgumentException($"'{dataType}' is not a data type.", nameof(dataType));
        MeasureFormat.CheckFits(format, dataType, name, nameof(format));
        if (Columns.Contains(name) || _calculatedColumns.Any(c => string.Equals(c.Name, name, StringComparison.OrdinalIgnoreCase)))
            throw new ArgumentException($"Table '{Name}' already has a column named '{name}'.", nameof(name));
        if (Model.Tables.SelectMany(t => t.Measures).FirstOrDefault(m => string.Equals(m.Name, name, StringComparison.OrdinalIgnoreCase)) is { } taken)
            throw new ArgumentException($"A measure named '{name}' already exists ({taken}); a column cannot share its name.", nameof(name));
        if (_hierarchies.Any(h => string.Equals(h.Name, name, StringComparison.OrdinalIgnoreCase)))
            throw new ArgumentException($"Table '{Name}' already has a hierarchy named '{name}'.", nameof(name));
        var column = new ModelCalculatedColumn(this, Builder.AddCalculatedColumn(name, expression, dataType, format));
        _calculatedColumns.Add(column);
        return column;
    }

    /// <summary>
    /// Adds a hierarchy, like the one Power Pivot's diagram view creates with "Create Hierarchy": an ordered drill-down path
    /// such as Category, Subcategory, Product that PivotTables offer as a single field.
    /// </summary>
    /// <param name="name">Hierarchy name; unique within the table and different from its column names.</param>
    /// <param name="levels">Two or more column names of this table, from the top level down. A level may be a calculated column
    /// (<see cref="AddCalculatedColumn"/>); Excel has not calculated it, so it saves such a hierarchy unprocessed (no tree) and builds it
    /// when the column is calculated. <see cref="ModelHierarchy.IsUnprocessed"/> tells.</param>
    /// <returns>The hierarchy.</returns>
    /// <exception cref="ArgumentException">The name is empty or taken, a column is unknown or repeated, or fewer than two levels are given.</exception>
    public ModelHierarchy AddHierarchy(string name, params string[] levels)
    {
        Model.ThrowIfDisposed();
        ThrowIfCalculationTrigger();
        ArgumentNullException.ThrowIfNull(levels);
        if (string.IsNullOrWhiteSpace(name)) throw new ArgumentException("Hierarchy name is required.", nameof(name));
        string modelName = Identifiers.SanitizeName(name, "Hierarchy");
        if (levels.Length < 2) throw new ArgumentException("A hierarchy needs at least two levels.", nameof(levels));
        if (_hierarchies.Any(h => string.Equals(h.Name, modelName, StringComparison.OrdinalIgnoreCase)))
            throw new ArgumentException($"Table '{Name}' already has a hierarchy named '{modelName}'.", nameof(name));
        if (Columns.Contains(modelName) || _calculatedColumns.Any(c => string.Equals(c.Name, modelName, StringComparison.OrdinalIgnoreCase)))
            throw new ArgumentException($"Table '{Name}' already has a column named '{modelName}'; a hierarchy needs its own name.", nameof(name));
        var columns = new List<ModelColumn>();
        var hierarchyLevels = new List<HierarchyLevel>();
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var level in levels)
        {
            ArgumentNullException.ThrowIfNull(level, nameof(levels));
            // A calculated column is looked up first by name; any other name must be one of the table's data columns.
            var calculated = _calculatedColumns.FirstOrDefault(x => string.Equals(x.Name, level, StringComparison.OrdinalIgnoreCase));
            if (calculated is not null)
            {
                if (!seen.Add(calculated.Name)) throw new ArgumentException($"Column '{calculated.Name}' is used twice in the hierarchy.", nameof(levels));
                hierarchyLevels.Add(new HierarchyLevel(calculated.Name, calculated.Definition.Id, null));
                continue;
            }
            var c = Builder.FindColumn(level, nameof(levels));
            if (!seen.Add(c.ModelName)) throw new ArgumentException($"Column '{c.Definition.Name}' is used twice in the hierarchy.", nameof(levels));
            columns.Add(_columns.First(x => x.Builder == c));
            hierarchyLevels.Add(new HierarchyLevel(c.ModelName, c.Id, c));
        }
        Model.HierarchyDefinitions.Add(new HierarchyDefinition { Name = modelName, Table = Builder, Levels = hierarchyLevels });
        var hierarchy = new ModelHierarchy(this, modelName, columns, hierarchyLevels.Select(l => l.Name).ToList());
        _hierarchies.Add(hierarchy);
        return hierarchy;
    }

    /// <summary>
    /// Marks the table as the model's date table, like Power Pivot's "Mark as Date Table": the date column becomes the
    /// table's key, which Excel's time intelligence functions (TOTALYTD, SAMEPERIODLASTYEAR, ...) need.
    /// </summary>
    /// <remarks>
    /// Written the way Excel saves a model it just marked; only that case has been seen in an Excel-written file, so
    /// saving requires a hierarchy on the table, a relationship that uses the date column as its key, and a date on every row
    /// (whole days, no duplicates). One date table per model.
    /// </remarks>
    /// <param name="dateColumn">Column of type <see cref="ColumnDataType.DateTime"/> with one row per day.</param>
    /// <exception cref="ArgumentException">The column is unknown, not a date column, or takes part in Sort by Column.</exception>
    /// <exception cref="InvalidOperationException">A date table was already marked.</exception>
    public void MarkAsDateTable(string dateColumn)
    {
        Model.ThrowIfDisposed();
        ThrowIfCalculationTrigger();
        var c = Builder.FindColumn(dateColumn, nameof(dateColumn));
        if (c.Definition.DataType != ColumnDataType.DateTime)
            throw new ArgumentException($"'{Name}[{c.Definition.Name}]' is {c.Definition.DataType}; a date table needs a DateTime column.", nameof(dateColumn));
        if (c.SortBy is not null || Builder.Columns.Any(o => o.SortBy == c))
            throw new ArgumentException($"'{Name}[{c.Definition.Name}]' takes part in Sort by Column, so it cannot be the date column.", nameof(dateColumn));
        if (Model.DateTable is { } existing)
            throw new InvalidOperationException($"'{existing.Name}' is already marked as date table; only one date table per model is supported.");
        Builder.DateColumn = c;
        c.IsDateKey = true;
    }

    /// <summary>The table name and row count.</summary>
    /// <summary>True for the table made by <see cref="DataModelWriter.AddCalculationTrigger"/>.</summary>
    public bool IsCalculationTrigger => Builder.IsLinked;

    internal void ThrowIfCalculationTrigger()
    {
        if (Builder.IsLinked)
            throw new InvalidOperationException($"Table '{Name}' was made by AddCalculationTrigger and only triggers the calculation; it cannot take measures, columns, hierarchies, relationships or a date table.");
    }

    public override string ToString() => $"{Name} ({RowCount.ToString("N0", CultureInfo.InvariantCulture)} rows)";

    private static Action<IDataRecord> CreateAppender(TableBuilder table, int ordinal, Type fieldType)
    {
        var column = table.Columns[ordinal];
        Type t = Nullable.GetUnderlyingType(fieldType) ?? fieldType;
        switch (column)
        {
            case StringColumnBuilder s when t == typeof(string):
                return r => { if (r.IsDBNull(ordinal)) s.AppendNull(); else s.Append(r.GetString(ordinal)); };
            case Int64ColumnBuilder l when column.Definition.DataType == ColumnDataType.Int64:
                if (t == typeof(long)) return r => { if (r.IsDBNull(ordinal)) l.AppendNull(); else l.Append(r.GetInt64(ordinal)); };
                if (t == typeof(int)) return r => { if (r.IsDBNull(ordinal)) l.AppendNull(); else l.Append(r.GetInt32(ordinal)); };
                if (t == typeof(short)) return r => { if (r.IsDBNull(ordinal)) l.AppendNull(); else l.Append(r.GetInt16(ordinal)); };
                if (t == typeof(byte)) return r => { if (r.IsDBNull(ordinal)) l.AppendNull(); else l.Append(r.GetByte(ordinal)); };
                break;
            case Int64ColumnBuilder c when column.Definition.DataType == ColumnDataType.Currency && t == typeof(decimal):
                return r => { if (r.IsDBNull(ordinal)) c.AppendNull(); else c.AppendCurrency(r.GetDecimal(ordinal)); };
            case DoubleColumnBuilder d when column.Definition.DataType == ColumnDataType.Double:
                if (t == typeof(double)) return r => { if (r.IsDBNull(ordinal)) d.AppendNull(); else d.Append(r.GetDouble(ordinal)); };
                if (t == typeof(float)) return r => { if (r.IsDBNull(ordinal)) d.AppendNull(); else d.Append(r.GetFloat(ordinal)); };
                if (t == typeof(decimal)) return r => { if (r.IsDBNull(ordinal)) d.AppendNull(); else d.Append((double)r.GetDecimal(ordinal)); };
                break;
            case DoubleColumnBuilder dt when column.Definition.DataType == ColumnDataType.DateTime && t == typeof(DateTime):
                return r => { if (r.IsDBNull(ordinal)) dt.AppendNull(); else dt.AppendDate(r.GetDateTime(ordinal)); };
            case BooleanColumnBuilder b when t == typeof(bool):
                return r => { if (r.IsDBNull(ordinal)) b.AppendNull(); else b.Append(r.GetBoolean(ordinal)); };
        }
        return r =>
        {
            if (r.IsDBNull(ordinal)) { column.AppendNull(); return; }
            object v = r.GetValue(ordinal);
            if (v is DBNull) column.AppendNull();
            else column.AppendObject(table.Coerce(ordinal, v, table.Rows + 1));
        };
    }
}
