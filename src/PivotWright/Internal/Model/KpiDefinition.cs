namespace PivotWright.Internal.Model;

/// <summary>The KPI of a measure (SPEC section 12): a goal, status thresholds and an icon style.</summary>
internal sealed class KpiDefinition
{
    /// <summary>The target as a number (Excel's "Absolute value"), or null when <see cref="GoalMeasure"/> is the target.</summary>
    public double? GoalValue { get; init; }
    public MeasureDefinition? GoalMeasure { get; init; }
    public required KpiScale Scale { get; init; }
    /// <summary>Thresholds as fractions of the target, ascending.</summary>
    public required IReadOnlyList<double> Thresholds { get; init; }
    /// <summary>Excel's name of the status icon style, for example "Traffic Light - Single".</summary>
    public required string Graphic { get; init; }
}
