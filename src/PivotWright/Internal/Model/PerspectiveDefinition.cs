namespace PivotWright.Internal.Model;

/// <summary>What a perspective shows, as the serializer needs it (SPEC section 13).</summary>
internal sealed class PerspectiveDefinition
{
    public required string Name { get; init; }
    public List<(TableBuilder Table, ColumnBuilder Column)> Columns { get; } = new();
    public List<(TableBuilder Table, MeasureDefinition Measure)> Measures { get; } = new();
    public List<(TableBuilder Table, string Hierarchy)> Hierarchies { get; } = new();

    public bool IsEmpty => Columns.Count == 0 && Measures.Count == 0 && Hierarchies.Count == 0;
}
