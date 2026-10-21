using PivotWright.Internal.Model;

namespace PivotWright;

/// <summary>A column of a model table: <c>sales["Qty"]</c>.</summary>
public sealed class ModelColumn
{
    internal ModelColumn(ModelTable table, ColumnBuilder builder)
    {
        Table = table;
        Builder = builder;
    }

    internal ColumnBuilder Builder { get; }

    /// <summary>Table that holds the column.</summary>
    public ModelTable Table { get; }

    /// <summary>Column name as it appears in the model (the name given, without the characters Power Pivot rejects).</summary>
    public string Name => Builder.ModelName;

    /// <summary>Data type of the column.</summary>
    public ColumnDataType DataType => Builder.Definition.DataType;

    /// <summary>
    /// Hides the column from client tools, like Power Pivot's "Hide from Client Tools": it stays in the model (relationships,
    /// sorting, hierarchies and DAX can still use it) but is not listed in the PivotTable field list. The usual treatment for
    /// keys and sort-order columns. Starts as <see cref="ColumnDefinition.IsHidden"/>.
    /// </summary>
    public bool IsHidden
    {
        get => Builder.Definition.IsHidden;
        set
        {
            Table.Model.ThrowIfDisposed();
            Table.Builder.SetHidden(Builder, value);
        }
    }

    /// <summary>
    /// Another column of the same table that orders this column in client tools, like Power Pivot's "Sort by Column": Month
    /// sorted by MonthNumber lists January to December instead of April, August, ... Every value of this column must belong to
    /// exactly one value of the sort column, which is checked when saving. Null for the column's own order. Starts as
    /// <see cref="ColumnDefinition.SortByColumn"/>.
    /// </summary>
    /// <exception cref="ArgumentException">The column is of another table, is this column, is itself sorted, or this column is
    /// already a sort column; or one of them is the date column of a date table.</exception>
    public ModelColumn? SortByColumn
    {
        get => Builder.SortBy is { } by ? Table.Columns.First(c => c.Builder == by) : null;
        set
        {
            Table.Model.ThrowIfDisposed();
            if (value is null) { Table.Builder.ClearSortBy(Builder); return; }
            if (value.Table != Table)
                throw new ArgumentException($"{this} can only be sorted by a column of '{Table.Name}', not by {value}.", nameof(SortByColumn));
            Table.Builder.SetSortBy(Builder, value.Builder);
        }
    }

    /// <summary>The column in DAX notation, for example <c>'Sales'[Qty]</c>.</summary>
    public override string ToString() => $"'{Table.Name.Replace("'", "''")}'[{Name}]";
}
