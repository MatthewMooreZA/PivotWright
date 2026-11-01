namespace PivotWright;

/// <summary>How a KPI's status is divided into bands, like the four band layouts of Power Pivot's "Define status thresholds" picker.</summary>
/// <remarks>
/// The number of bands follows the icon style: a three-state style (<see cref="KpiIconStyle.States"/>) has three bands in the linear layouts and
/// five in the centered ones, the five-state <see cref="KpiIconStyle.GaugeAscending"/> five and nine.
/// </remarks>
public enum KpiScale
{
    /// <summary>Higher is better: bad, medium, good (Excel's default).</summary>
    Linear,
    /// <summary>Lower is better: good, medium, bad.</summary>
    LinearReversed,
    /// <summary>The best values lie around the target: bad, medium, good, medium, bad.</summary>
    Centered,
    /// <summary>The worst values lie around the target: good, medium, bad, medium, good.</summary>
    CenteredReversed,
}

/// <summary>The status icons of a KPI, like the "Select icon style" choice in Power Pivot. Excel offers exactly these five.</summary>
public sealed class KpiIconStyle : IEquatable<KpiIconStyle>
{
    private KpiIconStyle(string name, int states) { Name = name; States = states; }

    /// <summary>Excel's name of the style, as it appears in the file.</summary>
    public string Name { get; }

    /// <summary>Number of different statuses the icons show: 3 (bad, medium, good) or 5 (<see cref="GaugeAscending"/>), which fixes the number of thresholds.</summary>
    public int States { get; }

    /// <summary>The plain red, yellow and green circles; Excel's default.</summary>
    public static KpiIconStyle TrafficLightSingle { get; } = new("Traffic Light - Single", 3);

    /// <summary>Circles with a cross, an exclamation mark and a tick.</summary>
    public static KpiIconStyle RoadSigns { get; } = new("Road Signs", 3);

    /// <summary>Red, yellow and green dots in black squares.</summary>
    public static KpiIconStyle TrafficLight { get; } = new("Traffic Light", 3);

    /// <summary>A red diamond, a yellow triangle and a green circle.</summary>
    public static KpiIconStyle Shapes { get; } = new("Shapes", 3);

    /// <summary>Quarter-filled circles, five states.</summary>
    public static KpiIconStyle GaugeAscending { get; } = new("Gauge - Ascending", 5);

    /// <summary>All the styles.</summary>
    public static IReadOnlyList<KpiIconStyle> All { get; } = new[] { TrafficLightSingle, RoadSigns, TrafficLight, Shapes, GaugeAscending };

    internal static KpiIconStyle FromName(string name) => All.First(s => s.Name == name);

    public bool Equals(KpiIconStyle? other) => other is not null && Name == other.Name;
    public override bool Equals(object? obj) => Equals(obj as KpiIconStyle);
    public override int GetHashCode() => Name.GetHashCode();
    public override string ToString() => Name;
}

/// <summary>Optional settings of <see cref="ModelMeasure.AddKpi(double, KpiOptions?)"/>. The defaults are those of Power Pivot's Key Performance Indicator dialog.</summary>
public sealed class KpiOptions
{
    private IReadOnlyList<double>? _thresholds;
    private KpiScale _scale;
    private KpiIconStyle _iconStyle = KpiIconStyle.TrafficLightSingle;

    /// <summary>Division of the status into bands. Default <see cref="KpiScale.Linear"/>.</summary>
    /// <exception cref="ArgumentOutOfRangeException">The value is not a <see cref="KpiScale"/>.</exception>
    public KpiScale Scale
    {
        get => _scale;
        init
        {
            if (!Enum.IsDefined(value)) throw new ArgumentOutOfRangeException(nameof(Scale), value, "Not a KpiScale.");
            _scale = value;
        }
    }

    /// <summary>
    /// Where the status changes, as fractions of the target in ascending order. For the linear layouts one fewer than the icon style has states
    /// (<see cref="KpiIconStyle.States"/>): two for the three-state styles, four for the gauge. For the centered layouts twice that: four, or eight.
    /// For a target of 1000, <c>0.4, 0.8</c> turns the status at 400 and 800. Defaults, the ones Power Pivot proposes: 0.4 and 0.8 (linear, three states),
    /// 0.34, 0.67, 1.33 and 1.66 (centered, three states), 0.2, 0.4, 0.6 and 0.8 (linear gauge) and 0.2 to 0.8 followed by 1.2 to 1.8 (centered gauge).
    /// </summary>
    /// <exception cref="ArgumentException">A fraction is not positive, or they are not ascending.</exception>
    public IReadOnlyList<double>? Thresholds
    {
        get => _thresholds;
        init
        {
            if (value is not null)
            {
                for (int i = 0; i < value.Count; i++)
                {
                    if (!double.IsFinite(value[i]) || value[i] <= 0)
                        throw new ArgumentException($"Threshold {i + 1} ({value[i]}) must be a positive number: a fraction of the target, for example 0.8.", nameof(Thresholds));
                    if (i > 0 && value[i] <= value[i - 1])
                        throw new ArgumentException($"Thresholds must be ascending, but threshold {i + 1} ({value[i]}) is not above threshold {i} ({value[i - 1]}).", nameof(Thresholds));
                }
                value = value.ToArray();
            }
            _thresholds = value;
        }
    }

    /// <summary>The status icons. Default <see cref="KpiIconStyle.TrafficLightSingle"/>.</summary>
    public KpiIconStyle IconStyle
    {
        get => _iconStyle;
        init => _iconStyle = value ?? throw new ArgumentNullException(nameof(IconStyle));
    }
}
