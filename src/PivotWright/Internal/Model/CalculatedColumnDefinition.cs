namespace PivotWright.Internal.Model;

/// <summary>A calculated column: a DAX formula that Excel evaluates for every row; the file holds no values.</summary>
internal sealed class CalculatedColumnDefinition
{
    public required string Name { get; init; }
    /// <summary>Identifier used in the model's files and attribute ids.</summary>
    public required string Id { get; init; }
    public required string Expression { get; init; }
    public required ColumnDataType DataType { get; init; }
}
