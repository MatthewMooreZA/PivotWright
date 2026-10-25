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

    /// <summary>The column in DAX notation, for example <c>'Sales'[Line Total]</c>.</summary>
    public override string ToString() => $"'{Table.Name.Replace("'", "''")}'[{Name}]";
}
