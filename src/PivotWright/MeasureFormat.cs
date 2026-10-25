using System.Globalization;

namespace PivotWright;

/// <summary>
/// Number format of a DAX measure, a data column or a calculated column, as chosen in Power Pivot's Data Type and Format commands.
/// Which formats fit which type: whole number, decimal, currency and percentage fit the numeric types (Whole Number, Decimal Number,
/// Currency); dates fit Date columns; text and TRUE/FALSE columns take no format.
/// </summary>
public sealed class MeasureFormat
{
    internal enum Kind { General, WholeNumber, Decimal, Currency, Percentage, ShortDate, LongDate }

    internal MeasureFormat(Kind kind, int decimals = 2, bool thousands = true, CultureInfo? culture = null, bool noSymbol = false)
    {
        KindValue = kind;
        Decimals = decimals;
        ThousandSeparator = thousands;
        Culture = culture;
        NoSymbol = noSymbol;
    }

    internal Kind KindValue { get; }
    internal int Decimals { get; }
    internal bool ThousandSeparator { get; }
    /// <summary>The culture whose currency symbol or date pattern is used, or null for the model's culture.</summary>
    internal CultureInfo? Culture { get; }
    /// <summary>Currency without a symbol (Power Pivot's Symbol "None").</summary>
    internal bool NoSymbol { get; }

    /// <summary>No formatting ("General").</summary>
    public static MeasureFormat General { get; } = new(Kind.General);

    /// <summary>Whole number with thousands separator (for example 1,234).</summary>
    public static MeasureFormat WholeNumber { get; } = new(Kind.WholeNumber);

    /// <summary>Decimal number (for example 1,234.57).</summary>
    public static MeasureFormat Decimal(int decimals = 2, bool thousandSeparator = true)
        => new(Kind.Decimal, Check(decimals), thousandSeparator);

    /// <summary>
    /// Currency with the symbol and layout of a culture (for example R 1,234.57 for en-ZA, $1,234.57 for en-US). With no culture,
    /// the model's culture (<see cref="DataModelOptions.Culture"/>) is used.
    /// </summary>
    public static MeasureFormat Currency(int decimals = 2, CultureInfo? culture = null)
        => new(Kind.Currency, Check(decimals), culture: culture);

    /// <summary>Currency without a symbol (Power Pivot's Symbol "None"), for example 1,234.57; negatives are in parentheses.</summary>
    public static MeasureFormat CurrencyWithoutSymbol(int decimals = 2) => new(Kind.Currency, Check(decimals), noSymbol: true);

    /// <summary>Percentage (0.256 is shown as 25.60%).</summary>
    public static MeasureFormat Percentage(int decimals = 2) => new(Kind.Percentage, Check(decimals));

    /// <summary>
    /// The short date pattern of a culture, for example 3/14/2001 for en-US, 14/03/2001 for en-GB, 2001/03/14 for en-ZA. With no culture, the
    /// model's culture is used, so the order of day, month and year is never fixed by the library.
    /// </summary>
    public static MeasureFormat ShortDate(CultureInfo? culture = null) => new(Kind.ShortDate, culture: culture);

    /// <summary>The long date pattern of a culture, for example Wednesday, March 14, 2001 for en-US. With no culture, the model's culture is used.</summary>
    public static MeasureFormat LongDate(CultureInfo? culture = null) => new(Kind.LongDate, culture: culture);

    private static int Check(int d)
        => d is >= 0 and <= 9 ? d : throw new ArgumentOutOfRangeException(nameof(d), "Decimals must be between 0 and 9.");

    /// <summary>Throws when the format does not fit the data type of the column it is given to.</summary>
    internal static void CheckFits(MeasureFormat? format, ColumnDataType type, string column, string paramName)
    {
        if (format is null || format.KindValue == Kind.General) return;
        bool numeric = type is ColumnDataType.Int64 or ColumnDataType.Double or ColumnDataType.Currency;
        bool date = format.KindValue is Kind.ShortDate or Kind.LongDate;
        if (date ? type == ColumnDataType.DateTime : numeric) return;
        string fits = date ? "a date format fits Date columns" : "a number format fits Whole Number, Decimal Number and Currency columns";
        throw new ArgumentException($"Column '{column}' is {type}; {fits}, not {format.KindValue}.", paramName);
    }
}
