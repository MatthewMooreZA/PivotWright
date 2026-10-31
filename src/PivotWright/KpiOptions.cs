namespace PivotWright;

/// <summary>How a KPI's status is divided into bands, like the layouts of Power Pivot's "Define status thresholds" section.</summary>
public enum KpiScale
{
    /// <summary>Higher is better, with two thresholds and three bands: bad, medium, good (Excel's default).</summary>
    Linear,
    /// <summary>The best values lie around the target, with four thresholds and five bands: bad, medium, good, medium, bad.</summary>
    Centered,
}

/// <summary>The status icons of a KPI, like the "Select icon style" choice in Power Pivot.</summary>
/// <remarks>
/// Only <see cref="TrafficLightSingle"/> and <see cref="RoadSigns"/> have been seen in Excel-written files. Other styles can be
/// named with <see cref="Named"/>; they are written the same way but were not tested in Excel.
/// </remarks>
public sealed class KpiIconStyle : IEquatable<KpiIconStyle>
{
    private KpiIconStyle(string name) => Name = name;

    /// <summary>Excel's name of the style, as Power Pivot lists it.</summary>
    public string Name { get; }

    /// <summary>"Traffic Light - Single", Excel's default style.</summary>
    public static KpiIconStyle TrafficLightSingle { get; } = new("Traffic Light - Single");

    /// <summary>"Road Signs".</summary>
    public static KpiIconStyle RoadSigns { get; } = new("Road Signs");

    /// <summary>A style by the name Power Pivot shows for it.</summary>
    /// <exception cref="ArgumentException">The name is empty or contains a quote or a control character.</exception>
    public static KpiIconStyle Named(string name)
    {
        if (string.IsNullOrWhiteSpace(name)) throw new ArgumentException("The icon style name is required.", nameof(name));
        name = name.Trim();
        if (name.Contains('\'') || name.Any(char.IsControl))
            throw new ArgumentException($"The icon style name '{name}' must not contain quotes or control characters.", nameof(name));
        return new KpiIconStyle(name);
    }

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
    /// Where the status changes, as fractions of the target in ascending order: two for <see cref="KpiScale.Linear"/> and four for
    /// <see cref="KpiScale.Centered"/>. For a target of 1000, <c>0.4, 0.8</c> turns the status at 400 and 800. Default 0.4 and 0.8
    /// (Linear) or 0.34, 0.67, 1.33 and 1.66 (Centered), the ones Power Pivot proposes.
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
