namespace PivotWright.Internal.Model;

/// <summary>A calculated column: a DAX formula that Excel evaluates for every row; the file holds no values.</summary>
internal sealed class CalculatedColumnDefinition
{
    public required string Name { get; init; }
    /// <summary>Identifier used in the model's files and attribute ids.</summary>
    public required string Id { get; init; }
    public required string Expression { get; init; }
    public required ColumnDataType DataType { get; init; }
    /// <summary>Hidden from client tools, like a hidden data column (H35).</summary>
    public bool IsHidden { get; set; }
    /// <summary>The number or date format chosen for the column, or null for the one Excel infers from the type.</summary>
    public MeasureFormat? Format { get; init; }
    /// <summary>Set when the model is saved: the data file prefix, the dictionary data version and the data file version Excel gives the column.</summary>
    public (int Prefix, int Dictionary, int File) Versions { get; set; } = (2, 3, 2);
}
