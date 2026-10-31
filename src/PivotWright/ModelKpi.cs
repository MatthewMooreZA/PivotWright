using PivotWright.Internal.Model;

namespace PivotWright;

/// <summary>The key performance indicator of a measure, added with <see cref="ModelMeasure.AddKpi(double, KpiOptions?)"/>.</summary>
public sealed class ModelKpi
{
    private readonly KpiDefinition _definition;

    internal ModelKpi(ModelMeasure measure, KpiDefinition definition, ModelMeasure? targetMeasure)
    {
        Measure = measure;
        _definition = definition;
        TargetMeasure = targetMeasure;
    }

    /// <summary>The measure this KPI belongs to; its value is the KPI's value.</summary>
    public ModelMeasure Measure { get; }

    /// <summary>The target as a number, or null when the target is a measure.</summary>
    public double? Target => _definition.GoalValue;

    /// <summary>The measure that holds the target, or null when the target is a number.</summary>
    public ModelMeasure? TargetMeasure { get; }

    /// <summary>Division of the status into bands.</summary>
    public KpiScale Scale => _definition.Scale;

    /// <summary>Where the status changes, as fractions of the target (see <see cref="KpiOptions.Thresholds"/>).</summary>
    public IReadOnlyList<double> Thresholds => _definition.Thresholds;

    /// <summary>The status icons.</summary>
    public KpiIconStyle IconStyle => KpiIconStyle.Named(_definition.Graphic);

    /// <summary>The KPI in Power Pivot notation, for example <c>KPI of 'Sales'[Total Qty]</c>.</summary>
    public override string ToString() => "KPI of " + Measure;
}
