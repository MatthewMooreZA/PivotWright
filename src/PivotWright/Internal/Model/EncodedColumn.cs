using PivotWright.Internal.VertiPaq;

namespace PivotWright.Internal.Model;

internal enum XmType
{
    Long = 0,
    Real = 1,
    String = 2,
}

internal enum DictionaryKind
{
    /// <summary>Value encoding: no dictionary file, XMValueDataDictionary with BaseId/Magnitude.</summary>
    Value,
    String,
    Long32,
    Long64,
    Real,
}

/// <summary>Everything the metadata writers need to know about a column after its data was encoded.</summary>
internal sealed class EncodedColumn
{
    public required string Name { get; init; }
    public required string Id { get; init; }
    public required ColumnDataType DataType { get; init; }
    public bool IsHidden { get; init; }
    /// <summary>Id of the column whose order this column follows (Sort by Column), or null.</summary>
    public string? OrderByColumnId { get; set; }
    public bool IsRowNumber { get; init; }
    /// <summary>Key of the one side of a relationship: ColumnFlags 11 and a unique dictionary, as Excel writes them.</summary>
    public bool IsPrimaryKey { get; init; }
    /// <summary>Many-side key of a relationship (column statistics Usage 1; primary keys use 0, other columns 3).</summary>
    public bool IsForeignKey { get; init; }

    public DictionaryKind Dictionary { get; set; }
    public XmType XmType { get; set; }
    public int DbType { get; set; }
    public long BaseId { get; set; }
    public double Magnitude { get; set; } = 1.0;
    public long LastId { get; set; }

    public long Rows { get; set; }
    public bool HasNulls { get; set; }
    /// <summary>Distinct values including BLANK when the column has nulls.</summary>
    public long DistinctInclBlank { get; set; }
    public uint MinDataId { get; set; }
    public uint MaxDataId { get; set; }
    public long RleRuns { get; set; }
    public long BitPackedRuns { get; set; }
    public List<EncodedSegment> Segments { get; } = new();

    public bool IsValueEncoded => Dictionary == DictionaryKind.Value;

    /// <summary>Records in the column hierarchy (H$) table.</summary>
    public long HierarchyRecords => DistinctInclBlank + (HasNulls ? 2 : 3);

    /// <summary>
    /// Record counts of the hierarchy table's segments: the distinct values, then the reserved
    /// data ids. A column without any value (empty table) gets a single segment.
    /// </summary>
    /// <summary>
    /// Segment sizes of the hierarchy tables. Records are chunked in units of the distinct-value count
    /// (including BLANK): 3 values plus 3 reserved ids give [3, 3], 6 values give [6, 3], but 2 values
    /// give [2, 2, 1] and 1 value gives [1, 1, 1, 1], as Excel and Power BI write them.
    /// </summary>
    public IReadOnlyList<long> HierarchySegments
    {
        get
        {
            long records = HierarchyRecords;
            long unit = DistinctInclBlank;
            if (unit <= 0) return new[] { records };
            var list = new List<long>();
            for (long left = records; left > 0; left -= unit) list.Add(Math.Min(unit, left));
            return list;
        }
    }
}

/// <summary>Names of the files that make up a table and its columns ([MS-XLDM] 2.2.3.7).</summary>
internal static class FileNames
{
    public const string RowNumberId = "__XL_RowNumber";

    public static string Idf(string tableId, string columnId) => $"0.{tableId}.{columnId}.0.idf";
    public static string Dictionary(string tableId, string columnId) => $"0.{tableId}.{columnId}.dictionary";
    public static string HierarchyTable(string tableId, string columnId) => $"H${tableId}${columnId}";
    public static string HierarchyTableXml(string tableId, string columnId) => $"H${tableId}${columnId}.0.tbl.xml";
    public static string Hidx(string tableId, string columnId) => $"0.H${tableId}${columnId}.hidx";
    public static string PosToId(string tableId, string columnId) => $"0.H${tableId}${columnId}.POS_TO_ID.0.idf";
    public static string IdToPos(string tableId, string columnId) => $"0.H${tableId}${columnId}.ID_TO_POS.0.idf";
    public static string TableXml(string tableId, int version = 0) => $"{tableId}.{version}.tbl.xml";
    /// <summary>Data version Excel gives user hierarchy tables (and the tables that hold them) when it creates them.</summary>
    public const int HierarchyVersion = 2;
    public static string UserHierarchyTable(string tableId, string hierarchy) => $"U${tableId}${hierarchy}";
    public static string UserHierarchyTableXml(string tableId, string hierarchy) => $"{UserHierarchyTable(tableId, hierarchy)}.{HierarchyVersion}.tbl.xml";
    public static string UserHierarchyColumn(string tableId, string hierarchy, string column) => $"{HierarchyVersion}.{UserHierarchyTable(tableId, hierarchy)}.{column}.0.idf";
    public static string RelationshipTable(string foreignTableId, string relationship) => $"R${foreignTableId}${relationship}";
    public static string RelationshipTableXml(string foreignTableId, string relationship) => $"{RelationshipTable(foreignTableId, relationship)}.0.tbl.xml";
    public static string RelationshipIndex(string foreignTableId, string relationship) => $"0.{RelationshipTable(foreignTableId, relationship)}.INDEX.0.idf";
}

/// <summary>Destination for the data files of one table (its dimension folder in the model).</summary>
internal interface IFileSink
{
    /// <summary>Creates a file in the table's dimension folder; dispose the stream to commit it.</summary>
    Stream Create(string fileName);
}
