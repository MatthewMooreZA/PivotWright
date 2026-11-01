namespace PivotWright;

/// <summary>How a KPI's status is divided into bands, like the layouts of Power Pivot's "Define status thresholds" section.</summary>
public enum KpiScale
{
    /// <summary>Higher is better, with two thresholds and three bands: bad, medium, good (Excel's default).</summary>
    Linear,
    /// <summary>Lower is better, with two thresholds and three bands: good, medium, bad (Power Pivot's reversed layout, Excel's H51).</summary>
    LinearReversed,
    /// <summary>The best values lie around the target, with four thresholds and five bands: bad, medium, good, medium, bad.</summary>
    Centered,
    /// <summary>
    /// Higher is better, with four thresholds and five bands, for the five-state icon style <see cref="KpiIconStyle.GaugeAscending"/> (Excel's
    /// H61 writes it as status type Linear with four thresholds and the states -2 to 2).
    /// </summary>
    LinearFiveBands,
}

/// <summary>The status icons of a KPI, like the "Select icon style" choice in Power Pivot. Excel offers exactly these five.</summary>
public sealed class KpiIconStyle : IEquatable<KpiIconStyle>
{
    private KpiIconStyle(string name, int states) { Name = name; States = states; }

    /// <summary>Excel's name of the style, as it appears in the file.</summary>
    public string Name { get; }

    /// <summary>Number of different statuses the icons show: 3 (bad, medium, good) or 5 (<see cref="GaugeAscending"/>).</summary>
    public int States { get; }

    /// <summary>The plain red, yellow and green circles; Excel's default.</summary>
    public static KpiIconStyle TrafficLightSingle { get; } = new("Traffic Light - Single", 3);

    /// <summary>Circles with a cross, an exclamation mark and a tick.</summary>
    public static KpiIconStyle RoadSigns { get; } = new("Road Signs", 3);

    /// <summary>Red, yellow and green dots in black squares.</summary>
    public static KpiIconStyle TrafficLight { get; } = new("Traffic Light", 3);

    /// <summary>A red diamond, a yellow triangle and a green circle.</summary>
    public static KpiIconStyle Shapes { get; } = new("Shapes", 3);

    /// <summary>Quarter-filled circles, five states; needs <see cref="KpiScale.LinearFiveBands"/>.</summary>
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
    /// Where the status changes, as fractions of the target in ascending order: two for <see cref="KpiScale.Linear"/> and <see cref="KpiScale.LinearReversed"/>, and four for
    /// <see cref="KpiScale.Centered"/>. For a target of 1000, <c>0.4, 0.8</c> turns the status at 400 and 800. Default 0.4 and 0.8
    /// (Linear and LinearReversed), 0.34, 0.67, 1.33 and 1.66 (Centered) or 0.2, 0.4, 0.6 and 0.8 (LinearFiveBands), the ones Power Pivot proposes.
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
