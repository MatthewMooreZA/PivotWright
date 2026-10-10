namespace PivotWright;

/// <summary>Data types supported in the Power Pivot data model.</summary>
public enum ColumnDataType
{
    /// <summary>Whole Number (64-bit integer).</summary>
    Int64,
    /// <summary>Decimal Number (double-precision floating point).</summary>
    Double,
    /// <summary>Fixed Decimal Number / Currency (64-bit integer scaled by 10,000).</summary>
    Currency,
    /// <summary>Date (stored as an OLE Automation date).</summary>
    DateTime,
    /// <summary>TRUE/FALSE.</summary>
    Boolean,
    /// <summary>Text.</summary>
    String,
}

/// <summary>Describes one column of a model table.</summary>
public sealed class ColumnDefinition
{
    public ColumnDefinition(string name, ColumnDataType dataType)
    {
        if (string.IsNullOrWhiteSpace(name)) throw new ArgumentException("Column name is required.", nameof(name));
        Name = name;
        DataType = dataType;
    }

    /// <summary>Column name as shown in the Power Pivot window and the PivotTable field list.</summary>
    public string Name { get; }

    /// <summary>Data type of the column.</summary>
    public ColumnDataType DataType { get; }

    /// <summary>When true, the column is hidden from client tools.</summary>
    public bool IsHidden { get; init; }

    /// <summary>
    /// Name of another column of the same table that orders this column in client tools, like Power Pivot's
    /// "Sort by Column": for example Month sorted by MonthNumber shows January to December instead of
    /// April, August, ... Every value of this column must belong to exactly one value of the sort column.
    /// </summary>
    public string? SortByColumn { get; init; }

    public override string ToString() => $"{Name} ({DataType})";

    /// <summary>Maps a CLR type to the data model type used for it.</summary>
    public static ColumnDataType FromClrType(Type type, bool decimalAsCurrency = false)
    {
        type = Nullable.GetUnderlyingType(type) ?? type;
        if (type == typeof(long) || type == typeof(int) || type == typeof(short) || type == typeof(sbyte) ||
            type == typeof(byte) || type == typeof(ushort) || type == typeof(uint) || type == typeof(ulong))
            return ColumnDataType.Int64;
        if (type == typeof(double) || type == typeof(float) || type == typeof(Half)) return ColumnDataType.Double;
        if (type == typeof(decimal)) return decimalAsCurrency ? ColumnDataType.Currency : ColumnDataType.Double;
        if (type == typeof(DateTime) || type == typeof(DateTimeOffset) || type == typeof(DateOnly) ||
            type == typeof(TimeSpan) || type == typeof(TimeOnly))
            return ColumnDataType.DateTime;
        if (type == typeof(bool)) return ColumnDataType.Boolean;
        return ColumnDataType.String;
    }
}
