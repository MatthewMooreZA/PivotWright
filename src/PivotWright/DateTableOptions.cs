namespace PivotWright;

/// <summary>Settings for <see cref="DataModelWriter.AddDateTable"/>. Every setting has a default, and each feature can be switched off on its own.</summary>
public sealed class DateTableOptions
{
    private int _yearEndMonth = 12;
    private string _yearLabel = "Fiscal";

    /// <summary>
    /// Last month of the year, 1 to 12. 12 (December, the default) means the year is the calendar year; any other month adds
    /// the year columns (FiscalYear, FiscalQuarter, FiscalMonth and their numbers) and a hierarchy for that year, for example
    /// a June year end makes July 2024 part of FiscalYear 2025 (the calendar year in which the year ends).
    /// </summary>
    /// <exception cref="ArgumentOutOfRangeException">The month is not 1 to 12.</exception>
    public int YearEndMonth
    {
        get => _yearEndMonth;
        init
        {
            if (value is < 1 or > 12) throw new ArgumentOutOfRangeException(nameof(YearEndMonth), "The year end month must be 1 to 12.");
            _yearEndMonth = value;
        }
    }

    /// <summary>
    /// Prefix of the columns and the hierarchy added for a year end other than December, "Fiscal" by default. Use "Financial" or
    /// "Tax" for a year that is not a fiscal year (FinancialYear, "Financial Calendar"). Letters and digits only, starting with
    /// a letter. Not used when <see cref="YearEndMonth"/> is 12.
    /// </summary>
    /// <exception cref="ArgumentException">The label is empty or contains other characters.</exception>
    public string YearLabel
    {
        get => _yearLabel;
        init
        {
            if (string.IsNullOrEmpty(value) || !char.IsLetter(value[0]) || !value.All(char.IsLetterOrDigit))
                throw new ArgumentException("The year label must be letters and digits only, starting with a letter (for example \"Fiscal\" or \"Financial\").", nameof(YearLabel));
            _yearLabel = value;
        }
    }

    /// <summary>
    /// Extend the range to whole years (calendar years, or years ending in <see cref="YearEndMonth"/>), as DAX's CALENDARAUTO
    /// does and as Microsoft's date table guidance asks, so the first and last day given only need to fall inside the first
    /// and last year. Default true; false uses the days as given.
    /// </summary>
    public bool CompleteYears { get; init; } = true;

    /// <summary>Add YearMonthNumber and YearMonth ("Mar 2025", sorted by the number), so a month can be told apart across years. Default true.</summary>
    public bool IncludeYearMonth { get; init; } = true;

    /// <summary>
    /// Hide the number columns that only exist to sort other columns (QuarterNumber, MonthNumber, YearMonthNumber,
    /// DayOfWeekNumber and the year columns' numbers) from client tools, as is usual practice. Default true.
    /// </summary>
    public bool HideSortColumns { get; init; } = true;

    /// <summary>
    /// How the Date column is displayed, for example <c>MeasureFormat.Date("dd MMM yyyy")</c> or <c>MeasureFormat.LongDate()</c>.
    /// Null (the default) leaves Excel's general date format.
    /// </summary>
    /// <exception cref="ArgumentException">The format is a number format.</exception>
    public MeasureFormat? DateFormat
    {
        get => _dateFormat;
        init
        {
            if (value is not null && value.KindValue is not (MeasureFormat.Kind.General or MeasureFormat.Kind.ShortDate or MeasureFormat.Kind.LongDate or MeasureFormat.Kind.Date))
                throw new ArgumentException("The date format must be a date format (MeasureFormat.ShortDate, LongDate or Date).", nameof(DateFormat));
            _dateFormat = value;
        }
    }

    private MeasureFormat? _dateFormat;
}
