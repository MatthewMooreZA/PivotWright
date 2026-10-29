using PivotWright.Internal.Model;

namespace PivotWright;

/// <summary>A calculated column, added with <see cref="ModelTable.AddCalculatedColumn"/>.</summary>
public sealed class ModelCalculatedColumn
{
    private readonly CalculatedColumnDefinition _definition;

    internal ModelCalculatedColumn(ModelTable table, CalculatedColumnDefinition definition)
    {
        Table = table;
        _definition = definition;
    }

    internal CalculatedColumnDefinition Definition => _definition;

    /// <summary>Table that holds the column.</summary>
    public ModelTable Table { get; }

    /// <summary>Column name as it appears in the model.</summary>
    public string Name => _definition.Name;

    /// <summary>DAX formula, without the leading "=".</summary>
    public string Expression => _definition.Expression;

    /// <summary>Data type of the values Excel calculates.</summary>
    public ColumnDataType DataType => _definition.DataType;

    /// <summary>
    /// Hides the column from client tools, like Power Pivot's "Hide from Client Tools": it stays in the model but is not listed in the
    /// PivotTable field list. Written as in Excel's H35 sample.
    /// </summary>
    public bool IsHidden
    {
        get => _definition.IsHidden;
        set
        {
            Table.Model.ThrowIfDisposed();
            _definition.IsHidden = value;
        }
    }

    /// <summary>
    /// Another calculated column of the same table that orders this one in client tools, like Power Pivot's "Sort by Column": a text
    /// column sorted by a number column. Excel works out the order when it calculates the columns, so each value of this column must
    /// belong to exactly one value of the sort column; that cannot be checked here. Null for the column's own order. Only a calculated
    /// column can sort a calculated column (a data column and a calculated column sorting each other was never seen in Excel).
    /// </summary>
    /// <exception cref="ArgumentException">The column is of another table, is this column, is itself sorted, or this column is
    /// already a sort column.</exception>
    public ModelCalculatedColumn? SortByColumn
    {
        get => _definition.SortBy is { } by ? Table.CalculatedColumns.First(c => c.Definition == by) : null;
        set
        {
            Table.Model.ThrowIfDisposed();
            if (value is null) { _definition.SortBy = null; return; }
            if (value.Table != Table)
                throw new ArgumentException($"{this} can only be sorted by a calculated column of '{Table.Name}', not by {value}; a data column cannot sort a calculated column.", nameof(SortByColumn));
            if (value == this) throw new ArgumentException($"{this} cannot be sorted by itself.", nameof(SortByColumn));
            if (value._definition.SortBy is not null)
                throw new ArgumentException($"{value} is itself sorted by another column, so it cannot be a sort column.", nameof(SortByColumn));
            if (Table.CalculatedColumns.Any(c => c._definition.SortBy == _definition))
                throw new ArgumentException($"{this} is the sort column of another column, so it cannot be sorted itself.", nameof(SortByColumn));
            _definition.SortBy = value._definition;
        }
    }

    /// <summary>The column in DAX notation, for example <c>'Sales'[Line Total]</c>.</summary>
    public override string ToString() => $"'{Table.Name.Replace("'", "''")}'[{Name}]";
}
