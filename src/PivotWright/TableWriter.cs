using System.Data;
using System.Globalization;
using PivotWright.Internal.Model;

namespace PivotWright;

/// <summary>
/// A table of the model, returned by <see cref="DataModelWriter.CreateTable"/>, <see cref="DataModelWriter.AddTable(string, IDataReader, IReadOnlyDictionary{string, ColumnDataType}?)"/>
/// and <see cref="DataModelWriter.AddDateTable"/>. Appends rows, and can be passed to the writer's methods in place of the table name.
/// </summary>
public sealed class TableWriter
{
    private readonly TableBuilder _table;

    internal TableWriter(TableBuilder table) => _table = table;

    internal TableBuilder Builder => _table;

    /// <summary>Table name as it appears in the model (the name given, without the characters Power Pivot rejects).</summary>
    public string Name => _table.Name;

    /// <summary>Column definitions, in order, reflecting later changes such as <see cref="DataModelWriter.HideColumns(string, string[])"/>.</summary>
    public IReadOnlyList<ColumnDefinition> Columns => _table.Columns.Select(c => c.Definition).ToArray();

    /// <summary>Number of rows added so far.</summary>
    public long RowCount => _table.Rows;

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
        ArgumentNullException.ThrowIfNull(values);
        _table.AddRow(values);
    }

    /// <summary>Adds every row of a data reader. The reader's columns must match this table's columns by position.</summary>
    /// <remarks>
    /// If a value cannot be converted, or the reader fails, in the middle of a row, the exception is passed on and the table
    /// is left incomplete: saving the model then throws, naming the row. Load the table again from a corrected source.
    /// </remarks>
    public void AddRows(IDataReader reader)
    {
        ArgumentNullException.ThrowIfNull(reader);
        _table.ThrowIfReadOnly();
        if (reader.FieldCount != _table.Columns.Count)
            throw new ArgumentException($"The reader has {reader.FieldCount} columns but table '{Name}' has {_table.Columns.Count} ({Identifiers.List(_table.Columns.Select(c => c.Definition.Name))}); they are matched by position.", nameof(reader));
        var appenders = new Action<IDataRecord>[_table.Columns.Count];
        for (int i = 0; i < appenders.Length; i++)
            appenders[i] = CreateAppender(_table, i, reader.GetFieldType(i));
        while (reader.Read())
        {
            int done = 0;
            try
            {
                for (; done < appenders.Length; done++) appenders[done](reader);
            }
            catch (Exception e) when (done > 0)
            {
                _table.Fault($"Table '{Name}' is incomplete: loading row {_table.Rows + 1} from the data reader failed at column '{_table.Columns[done].Definition.Name}' ({e.Message}). Load the table again from a corrected source.");
                throw;
            }
            _table.CommitRow();
        }
    }

    /// <summary>Adds rows from objects using a projection to column values, one value per column in column order.</summary>
    public void AddRows<T>(IEnumerable<T> items, Func<T, object?[]> project)
    {
        ArgumentNullException.ThrowIfNull(items);
        ArgumentNullException.ThrowIfNull(project);
        foreach (var item in items) _table.AddRow(project(item));
    }

    /// <summary>The table name and row count.</summary>
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
