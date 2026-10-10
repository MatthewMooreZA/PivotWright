using System.Data;
using System.Globalization;
using PivotWright.Internal.Model;

namespace PivotWright;

/// <summary>Appends rows to a model table.</summary>
public sealed class TableWriter
{
    private readonly TableBuilder _table;
    private readonly object?[] _row;

    internal TableWriter(TableBuilder table)
    {
        _table = table;
        _row = new object?[table.Columns.Count];
        Columns = table.Columns.Select(c => c.Definition).ToArray();
    }

    /// <summary>Table name.</summary>
    public string Name => _table.Name;

    /// <summary>Column definitions, in order.</summary>
    public IReadOnlyList<ColumnDefinition> Columns { get; }

    /// <summary>Number of rows added so far.</summary>
    public long RowCount => _table.Rows;

    /// <summary>
    /// Adds a row. Values are converted to the column types; <c>null</c> and <see cref="DBNull"/> become BLANK.
    /// </summary>
    public void AddRow(params object?[] values) => _table.AddRow(values);

    /// <summary>Adds every row of a data reader. The reader's columns must match this table's columns by position.</summary>
    public void AddRows(IDataReader reader)
    {
        ArgumentNullException.ThrowIfNull(reader);
        if (reader.FieldCount != _table.Columns.Count)
            throw new ArgumentException($"The reader has {reader.FieldCount} columns but table '{Name}' has {_table.Columns.Count}.");
        var appenders = new Action<IDataRecord>[_table.Columns.Count];
        for (int i = 0; i < appenders.Length; i++)
            appenders[i] = CreateAppender(_table.Columns[i], i, reader.GetFieldType(i));
        while (reader.Read())
        {
            foreach (var a in appenders) a(reader);
            _table.CommitRow();
        }
    }

    /// <summary>Adds rows from objects using a projection to column values.</summary>
    public void AddRows<T>(IEnumerable<T> items, Func<T, object?[]> project)
    {
        foreach (var item in items) _table.AddRow(project(item));
    }

    private static Action<IDataRecord> CreateAppender(ColumnBuilder column, int ordinal, Type fieldType)
    {
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
            else column.AppendObject(v);
        };
    }
}
