namespace PivotWright.Internal.Model;

internal sealed class MeasureDefinition
{
    public required string Name { get; init; }
    public required string Table { get; init; }
    public required string Expression { get; init; }
    public required MeasureFormat Format { get; init; }
    public bool Hidden { get; set; }
}
