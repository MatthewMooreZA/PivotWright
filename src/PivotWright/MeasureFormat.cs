using System.Globalization;

namespace PivotWright;

/// <summary>
/// Number format of a DAX measure, a data column or a calculated column, as chosen in Power Pivot's Data Type and Format commands.
/// Which formats fit which type: whole number, decimal, currency and percentage fit the numeric types (Whole Number, Decimal Number,
/// Currency); dates fit Date columns; text and TRUE/FALSE columns take no format.
/// </summary>
public sealed class MeasureFormat
{
    internal enum Kind { General, WholeNumber, Decimal, Currency, Percentage, ShortDate, LongDate, Date }

    internal MeasureFormat(Kind kind, int decimals = 2, bool thousands = true, CultureInfo? culture = null, bool noSymbol = false, string? pattern = null)
    {
        KindValue = kind;
        Pattern = pattern;
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
    /// <summary>The .NET date format string of <see cref="Date"/>, as given.</summary>
    internal string? Pattern { get; }

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
    public static MeasureFormat ShortDate(CultureInfo? culture = null) => new(Kind.ShortDate, culture: culture, pattern: "d");

    /// <summary>The long date pattern of a culture, for example Wednesday, March 14, 2001 for en-US. With no culture, the model's culture is used.</summary>
    public static MeasureFormat LongDate(CultureInfo? culture = null) => new(Kind.LongDate, culture: culture, pattern: "D");

    /// <summary>
    /// A date format given as a .NET date and time format string, the same string <c>DateTime.ToString(format, culture)</c> takes: a custom
    /// pattern such as <c>"dd MMM yyyy"</c> or <c>"yyyy-MM-dd HH:mm"</c>, or a single standard specifier (<c>"d"</c> short date, <c>"D"</c> long
    /// date, <c>"t"</c>, <c>"T"</c>, <c>"f"</c>, <c>"F"</c>, <c>"g"</c>, <c>"G"</c>, <c>"M"</c>, <c>"Y"</c>, <c>"s"</c>, <c>"u"</c>, <c>"o"</c>,
    /// <c>"R"</c>), which is replaced by the culture's own pattern for it: <c>"d"</c> is <c>M/d/yyyy</c> for en-US, <c>dd/MM/yyyy</c> for en-GB.
    /// With no culture, the model's culture is used. The order of day, month and year is whatever the pattern or the culture says.
    /// </summary>
    /// <exception cref="ArgumentException">The format is empty or not a valid .NET date format string.</exception>
    public static MeasureFormat Date(string format, CultureInfo? culture = null)
    {
        ArgumentNullException.ThrowIfNull(format);
        if (format.Length == 0) throw new ArgumentException("A date format string is required.", nameof(format));
        try { new DateTime(2001, 3, 14, 13, 30, 55).ToString(format, culture ?? CultureInfo.InvariantCulture); }
        catch (FormatException e) { throw new ArgumentException($"'{format}' is not a valid date format string: {e.Message}", nameof(format), e); }
        string expanded = ExpandDatePattern(format, culture ?? CultureInfo.InvariantCulture);
        if (format is "o" or "O" || UnquotedContains(expanded, "Kz"))
            throw new ArgumentException($"'{format}' shows a time zone or offset (K, z); the Data Model has no time zones, so it cannot be used.", nameof(format));
        return new MeasureFormat(Kind.Date, culture: culture, pattern: format);
    }

    private static bool UnquotedContains(string pattern, string chars)
    {
        char q = '\0';
        for (int i = 0; i < pattern.Length; i++)
        {
            char c = pattern[i];
            if (q != '\0') { if (c == q) q = '\0'; }
            else if (c == '\\') i++;
            else if (c is '\'' or '"') q = c;
            else if (chars.IndexOf(c) >= 0) return true;
        }
        return false;
    }

    /// <summary>The pattern Excel stores for a date format: a standard specifier expanded through the culture, a custom pattern as given.</summary>
    internal static string ExpandDatePattern(string format, CultureInfo culture)
    {
        if (format.Length != 1) return format;
        // Recent culture data (ICU) puts a narrow no-break space before AM/PM; Windows, and so Excel, has a plain space (H38 screenshots).
        return Standard(format, culture).Replace('\u202f', ' ').Replace('\u00a0', ' ');
    }

    /// <summary>
    /// The engine's own format string for a .NET date pattern: Excel cell-format syntax, where <c>tt</c> is <c>AM/PM</c> and literals are
    /// in double quotes (H38). Fractional seconds <c>fff</c> become <c>000</c>.
    /// </summary>
    internal static string ToEngineDatePattern(string pattern)
    {
        var sb = new System.Text.StringBuilder();
        for (int i = 0; i < pattern.Length;)
        {
            char c = pattern[i];
            if (c == '\'' || c == '"')
            {
                int end = pattern.IndexOf(c, i + 1);
                if (end < 0) end = pattern.Length;
                string text = pattern.Substring(i + 1, end - i - 1);
                if (text.Length > 0) sb.Append('"').Append(text.Replace("\"", "")).Append('"');
                i = end + 1;
            }
            else if (c == '\\' && i + 1 < pattern.Length) { sb.Append('"').Append(pattern[i + 1]).Append('"'); i += 2; }
            else if (c == 't' || c == 'f' || c == 'F')
            {
                int n = 1;
                while (i + n < pattern.Length && pattern[i + n] == c) n++;
                if (c == 't') sb.Append(n >= 2 ? "AM/PM" : "A/P");
                else sb.Append('0', n);
                i += n;
            }
            else { sb.Append(c); i++; }
        }
        return sb.ToString();
    }

    /// <summary>The Group Excel files a date pattern under in its format list (H37, H38); custom patterns are classified by what they show.</summary>
    internal static string DateGroup(string source, string pattern)
    {
        switch (source)
        {
            case "d": return "ShortDate";
            case "D": return "LongDate";
            case "t": return "ShortTime";
            case "T": return "LongTime";
            case "s": return "SortableDateTime";
        }
        bool inQuote = false; char q = '\0';
        bool date = false, time = false, seconds = false;
        foreach (char c in pattern)
        {
            if (inQuote) { if (c == q) inQuote = false; continue; }
            if (c == '\'' || c == '"') { inQuote = true; q = c; }
            else if (c is 'd' or 'M' or 'y') date = true;
            else if (c is 'h' or 'H' or 'm') time = true;
            else if (c == 's') { time = true; seconds = true; }
        }
        if (date && time) return seconds ? "GeneralLongDateTime" : "GeneralShortDateTime";
        if (time) return seconds ? "LongTime" : "ShortTime";
        return source.Length == 1 && source[0] is 'M' or 'm' or 'Y' or 'y' ? "LongDate" : "ShortDate";
    }

    private static string Standard(string format, CultureInfo culture)
    {
        var f = culture.DateTimeFormat;
        return format[0] switch
        {
            'd' => f.ShortDatePattern,
            'D' => f.LongDatePattern,
            't' => f.ShortTimePattern,
            'T' => f.LongTimePattern,
            'f' => f.LongDatePattern + " " + f.ShortTimePattern,
            'F' => f.FullDateTimePattern,
            'g' => f.ShortDatePattern + " " + f.ShortTimePattern,
            'G' => f.ShortDatePattern + " " + f.LongTimePattern,
            'M' or 'm' => f.MonthDayPattern,
            'Y' or 'y' => f.YearMonthPattern,
            's' => f.SortableDateTimePattern,
            'u' => f.UniversalSortableDateTimePattern,
            'o' or 'O' => "yyyy'-'MM'-'dd'T'HH':'mm':'ss'.'fffffffK",
            'R' or 'r' => f.RFC1123Pattern,
            _ => throw new ArgumentException($"'{format}' is not a standard date format specifier; a one-letter custom pattern needs a leading % (for example \"%d\").", nameof(format)),
        };
    }

    private static int Check(int d)
        => d is >= 0 and <= 9 ? d : throw new ArgumentOutOfRangeException(nameof(d), "Decimals must be between 0 and 9.");

    /// <summary>Throws when the format does not fit the data type of the column it is given to.</summary>
    internal static void CheckFits(MeasureFormat? format, ColumnDataType type, string column, string paramName)
    {
        if (format is null || format.KindValue == Kind.General) return;
        bool numeric = type is ColumnDataType.Int64 or ColumnDataType.Double or ColumnDataType.Currency;
        bool date = format.KindValue is Kind.ShortDate or Kind.LongDate or Kind.Date;
        if (date ? type == ColumnDataType.DateTime : numeric) return;
        string fits = date ? "a date format fits Date columns" : "a number format fits Whole Number, Decimal Number and Currency columns";
        throw new ArgumentException($"Column '{column}' is {type}; {fits}, not {format.KindValue}.", paramName);
    }
}
