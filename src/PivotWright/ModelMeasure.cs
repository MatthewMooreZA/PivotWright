using PivotWright.Internal.Model;

namespace PivotWright;

/// <summary>A DAX measure, added with <see cref="ModelTable.AddMeasure"/>. Give it a KPI with <see cref="AddKpi(double, KpiOptions?)"/>.</summary>
public sealed class ModelMeasure
{
    private readonly MeasureDefinition _definition;
    private ModelKpi? _kpi;

    internal ModelMeasure(ModelTable table, MeasureDefinition definition)
    {
        Table = table;
        _definition = definition;
    }

    /// <summary>Table the measure belongs to; it appears under this table in the field list.</summary>
    public ModelTable Table { get; }

    /// <summary>Measure name, unique across the model.</summary>
    public string Name => _definition.Name;

    /// <summary>DAX expression, without the leading "=".</summary>
    public string Expression => _definition.Expression;

    /// <summary>Number format.</summary>
    public MeasureFormat Format => _definition.Format;

    /// <summary>Hides the measure from client tools, for a measure that only other measures use.</summary>
    public bool IsHidden
    {
        get => _definition.Hidden;
        set
        {
            Table.Model.ThrowIfDisposed();
            _definition.Hidden = value;
        }
    }

    internal MeasureDefinition Definition => _definition;

    /// <summary>The measure's KPI, or null when it has none (see <see cref="AddKpi(double, KpiOptions?)"/>).</summary>
    public ModelKpi? Kpi => _kpi;

    /// <summary>
    /// Adds a KPI to this measure with a fixed target, like Power Pivot's "Create KPI" with the target set to an absolute value. The
    /// measure is the KPI's value; the status is calculated from it against the target and the thresholds in <paramref name="options"/>.
    /// </summary>
    /// <remarks>
    /// Written as Excel saves a KPI (SPEC section 12). Only the shapes of Excel's samples are known, so <c>Save</c> requires a model
    /// with relationships and no hidden columns or calculated columns. A measure has at most one KPI.
    /// </remarks>
    /// <param name="target">The target value, a positive number, for example <c>1000</c>.</param>
    /// <param name="options">Scale, thresholds and icon style; Power Pivot's defaults when omitted.</param>
    /// <returns>The KPI.</returns>
    /// <exception cref="ArgumentException">The target is not a positive number, or the thresholds do not fit the scale.</exception>
    /// <exception cref="InvalidOperationException">The measure already has a KPI.</exception>
    public ModelKpi AddKpi(double target, KpiOptions? options = null)
    {
        if (!double.IsFinite(target) || target <= 0)
            throw new ArgumentOutOfRangeException(nameof(target), target, $"The target of the KPI of {this} must be a positive number.");
        return CreateKpi(target, null, options);
    }

    /// <summary>
    /// Adds a KPI to this measure whose target is another measure, like Power Pivot's "Create KPI" with the target set to a measure.
    /// The thresholds are fractions of the target measure's value.
    /// </summary>
    /// <param name="target">The measure that holds the target; it belongs to this model and is not this measure.</param>
    /// <param name="options">Scale, thresholds and icon style; Power Pivot's defaults when omitted.</param>
    /// <returns>The KPI.</returns>
    /// <exception cref="ArgumentNullException"><paramref name="target"/> is null.</exception>
    /// <exception cref="ArgumentException">The target is this measure or belongs to another model, or the thresholds do not fit the scale.</exception>
    /// <exception cref="InvalidOperationException">The measure already has a KPI.</exception>
    public ModelKpi AddKpi(ModelMeasure target, KpiOptions? options = null)
    {
        ArgumentNullException.ThrowIfNull(target);
        if (target.Table.Model != Table.Model)
            throw new ArgumentException($"Measure {target} belongs to another DataModelWriter.", nameof(target));
        if (target == this)
            throw new ArgumentException($"The target of the KPI of {this} cannot be the measure itself.", nameof(target));
        return CreateKpi(null, target, options);
    }

    private ModelKpi CreateKpi(double? value, ModelMeasure? targetMeasure, KpiOptions? options)
    {
        Table.Model.ThrowIfDisposed();
        if (_kpi is not null) throw new InvalidOperationException($"{this} already has a KPI; a measure has at most one.");
        options ??= new KpiOptions();
        int count = options.Scale == KpiScale.Centered ? 4 : 2;
        var thresholds = options.Thresholds ?? (options.Scale == KpiScale.Centered ? new[] { 0.34, 0.67, 1.33, 1.66 } : new[] { 0.4, 0.8 });
        if (thresholds.Count != count)
            throw new ArgumentException($"The {options.Scale} scale takes {count} thresholds, but {thresholds.Count} were given.", nameof(options));
        var definition = new KpiDefinition
        {
            GoalValue = value, GoalMeasure = targetMeasure?.Definition, Scale = options.Scale, Thresholds = thresholds, Graphic = options.IconStyle.Name,
        };
        _definition.Kpi = definition;
        _kpi = new ModelKpi(this, definition, targetMeasure);
        return _kpi;
    }

    /// <summary>The measure in DAX notation, for example <c>'Sales'[Total Qty]</c>.</summary>
    public override string ToString() => $"'{Table.Name.Replace("'", "''")}'[{Name}]";
}
