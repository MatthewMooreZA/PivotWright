namespace PivotWright.Internal.Model;

internal sealed class MeasureDefinition
{
    public required string Name { get; init; }
    public required string Table { get; init; }
    public required string Expression { get; init; }
    public required MeasureFormat Format { get; init; }
    public bool Hidden { get; set; }
    /// <summary>The hidden measure the library adds to models with relationships; its output was verified before hidden user measures were seen (H59).</summary>
    public bool Helper { get; init; }
    public KpiDefinition? Kpi { get; set; }
}
