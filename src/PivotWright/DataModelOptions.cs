using System.Globalization;

namespace PivotWright;

/// <summary>Settings for a <see cref="DataModelWriter"/>.</summary>
public sealed class DataModelOptions
{
    private int _segmentRowCount = 1 << 20;

    private CultureInfo _culture = CultureInfo.GetCultureInfo("en-US");

    /// <summary>
    /// Locale of the model, for example <c>CultureInfo.GetCultureInfo("en-ZA")</c>. It drives the collation used to sort and
    /// de-duplicate text values (case-insensitive, as in Power Pivot), the currency format, and the month and day names of
    /// <see cref="DataModelWriter.AddDateTable"/>. It must be a culture with a Windows locale id (LCID), which Excel stores in the
    /// file. Default: en-US.
    /// </summary>
    /// <exception cref="ArgumentException">The culture has no Windows locale id (the invariant culture or a custom culture).</exception>
    public CultureInfo Culture
    {
        get => _culture;
        init
        {
            ArgumentNullException.ThrowIfNull(value, nameof(Culture));
            if (value.LCID is 127 or 4096 or 0)
                throw new ArgumentException($"Culture '{value.Name}' has no Windows locale id; use a named culture such as \"en-US\" or \"en-ZA\".", nameof(Culture));
            _culture = value;
        }
    }

    /// <summary>Windows locale id of <see cref="Culture"/>, as written to the file.</summary>
    internal int Lcid => _culture.LCID;

    /// <summary>
    /// Rows per column segment. Must be a power of two between 16,384 and 16,777,216.
    /// Default 1,048,576, the segment size Excel uses.
    /// </summary>
    public int SegmentRowCount
    {
        get => _segmentRowCount;
        init
        {
            if (value < 16384 || value > 16777216 || (value & (value - 1)) != 0)
                throw new ArgumentOutOfRangeException(nameof(SegmentRowCount),
                    "Segment size must be a power of two between 16,384 and 16,777,216 rows.");
            _segmentRowCount = value;
        }
    }

    /// <summary>Directory for temporary spill files. Default: <see cref="Path.GetTempPath"/>.</summary>
    public string? TempDirectory { get; init; }

    /// <summary>
    /// Bytes of column data kept in memory per column before spilling to a temporary file. Default 32 MiB.
    /// </summary>
    public long SpillThresholdBytes { get; init; } = 32L << 20;

    /// <summary>Map <see cref="decimal"/> source columns to Currency (fixed decimal) instead of Double.</summary>
    public bool DecimalAsCurrency { get; init; }

    /// <summary>Name of the single worksheet in the generated workbook.</summary>
    public string WorksheetName { get; init; } = "Sheet1";

    /// <summary>
    /// Name of the data source / workbook connection that the model tables belong to. It appears in
    /// Excel under Data &gt; Queries &amp; Connections.
    /// </summary>
    public string DataSourceName { get; init; } = "PivotWright";

    /// <summary>
    /// Connection string stored (encrypted, as Excel does) for the data source. The model's data is
    /// written directly, so this is only used if somebody refreshes the model in Excel.
    /// </summary>
    public string ConnectionString { get; init; } =
        "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=PivotWright;Persist Security Info=false";

    /// <summary>
    /// Make Excel calculate calculated columns when the workbook opens. Default true. A column that Excel has not calculated works as
    /// a PivotTable value but not as a row, column, filter or hierarchy level. With a calculated column in the model, <see cref="DataModelWriter.Save(Stream)"/>
    /// adds a one-row linked table, <c>PivotWrightTrigger</c> (column Key, value 1, listed in the field list), and a very hidden
    /// worksheet that holds the table it is linked to; the table refreshes when the workbook opens, and Excel calculates the model after
    /// processing it. Excel shows its usual Enable Editing and Enable Content bars. Set false to write only the formulas (the column
    /// then calculates once a field is put in a PivotTable's Values or after Calculate Now). Not used by <see cref="DataModelWriter.SaveModelPart"/>.
    /// </summary>
    public bool CalculateOnOpen { get; init; } = true;

    /// <summary>
    /// Leave the data source out of Excel's Refresh All (the connection's "Refresh this connection on Refresh All" box is unticked, as in
    /// Excel's H41 sample). The source is a placeholder, so Refresh All otherwise asks for a user name and password. Refresh All then
    /// refreshes the PivotTables only. Default false.
    /// </summary>
    public bool ExcludeSourceFromRefreshAll { get; init; }

    /// <summary>
    /// Compress the files inside the data model with Xpress (LZ77), as Excel does. Default true.
    /// Turning it off stores every chunk verbatim, which Excel also accepts, at the cost of a larger file.
    /// </summary>
    public bool XpressCompression { get; init; } = true;

    /// <summary>
    /// Deflate-compress the xl/model/item.data part inside the .xlsx. Excel itself stores the part
    /// uncompressed; leave this off unless file size matters more than fidelity to Excel's layout.
    /// </summary>
    public bool CompressModelPart { get; init; }

    /// <summary>
    /// When a model has relationships but no measures, add one hidden measure (<c>= 1</c>) to the table that
    /// holds the first relationship. Excel's "Sort by Column" fails on models with relationships unless the
    /// model contains a measure (see the README, "Open questions"). The measure is hidden from client tools.
    /// Default true; turn it off if you add your own measures or never sort in Excel.
    /// </summary>
    public bool AddHelperMeasure { get; init; } = true;
}
