namespace PivotWright;

/// <summary>Number format of a DAX measure, as chosen in Power Pivot's measure dialog.</summary>
public sealed class MeasureFormat
{
    internal enum Kind { General, WholeNumber, Decimal, Currency, Percentage, Date }

    internal MeasureFormat(Kind kind, int decimals = 2, bool thousands = true)
    {
        KindValue = kind;
        Decimals = decimals;
        ThousandSeparator = thousands;
    }

    internal Kind KindValue { get; }
    internal int Decimals { get; }
    internal bool ThousandSeparator { get; }

    /// <summary>No formatting ("General").</summary>
    public static MeasureFormat General { get; } = new(Kind.General);

    /// <summary>Whole number with thousands separator (for example 1,234).</summary>
    public static MeasureFormat WholeNumber { get; } = new(Kind.WholeNumber);

    /// <summary>Decimal number (for example 1,234.57).</summary>
    public static MeasureFormat Decimal(int decimals = 2, bool thousandSeparator = true)
        => new(Kind.Decimal, Check(decimals), thousandSeparator);

    /// <summary>Currency in the model culture (for example R 1,234.57).</summary>
    public static MeasureFormat Currency(int decimals = 2) => new(Kind.Currency, Check(decimals));

    /// <summary>Percentage (0.256 is shown as 25.60%).</summary>
    public static MeasureFormat Percentage(int decimals = 2) => new(Kind.Percentage, Check(decimals));

    /// <summary>Short date, as the regional settings show it (verified against an Excel-written file).</summary>
    public static MeasureFormat Date { get; } = new(Kind.Date);

    private static int Check(int d)
        => d is >= 0 and <= 9 ? d : throw new ArgumentOutOfRangeException(nameof(d), "Decimals must be between 0 and 9.");
}
