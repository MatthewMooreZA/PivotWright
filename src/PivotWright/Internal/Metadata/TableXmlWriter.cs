using PivotWright.Internal.Model;
using PivotWright.Internal.VertiPaq;

namespace PivotWright.Internal.Metadata;

/// <summary>Writes the XMSimpleTable documents (.tbl.xml) for a table and its column hierarchies ([MS-XLDM] 2.5).</summary>
internal static class TableXmlWriter
{
    private const string LT = "&lt;";

    public static string MainTable(EncodedTable t, int lcid)
    {
        var x = new XmBuilder();
        x.Open("XMSimpleTable", t.Id, root: true);
        x.Properties(p => p.Int("Version", t.Hierarchies.Count > 0 ? FileNames.HierarchyVersion : 0).Long("Settings", 4353).Long("RIViolationCount", 0));
        x.Members(m =>
        {
            m.Member("SegmentMap", s =>
            {
                s.Open("XMMultiPartSegmentMap");
                s.Properties(p => p.Long("FirstPartitionRecordCount", 0).Long("FirstPartitionSegmentCount", 0));
                s.Collections(c => c.Collection("Partitions", pc =>
                {
                    if (t.SegmentCount <= 1)
                    {
                        pc.Open("XMSegment1Map");
                        pc.Properties(p => p.Long("Records", t.Rows));
                    }
                    else
                    {
                        pc.Open($"XMSegmentEqualMapEx{LT}XMSegmentEqualMap_FastInstantiation>");
                        pc.Properties(p => p.Long("Segments", t.SegmentCount).Long("Records", t.Rows).Long("RecordsPerSegment", t.SegmentSize));
                    }
                    pc.Close();
                }));
                s.Close();
            });
            m.Member("TableStats", s => TableStats(s));
        });
        x.Collections(c =>
        {
            c.Collection("Partitions", pc =>
            {
                pc.Open("XMPartition", t.Id);
                pc.Properties(p => p.Bool("IsProcessed", true).Int("Partition", 0));
                pc.Close();
            });
            c.Collection("Columns", cc =>
            {
                foreach (var col in t.AllColumns) RawColumn(cc, t, col, lcid);
            });
            c.Collection("Relationships", rc =>
            {
                foreach (var r in t.Relationships) Relationship(rc, r);
            });
            c.Collection("UserHierarchies", hc =>
            {
                foreach (var h in t.Hierarchies)
                {
                    // TableStore lists every level with the position of its first node: $Level$0$Level2$3$...$
                    var store = new System.Text.StringBuilder("$");
                    for (int i = 0; i < h.Levels.Count; i++) store.Append(h.Levels[i].Id).Append('$').Append(h.LevelStarts[i]).Append('$');
                    hc.Open("XMUserHierarchy", h.Name);
                    hc.Raw("<Properties><IsProcessed xsi:type=\"xsd:boolean\">true</IsProcessed><TableStore>").Raw(Xml.Esc(store.ToString()))
                      .Raw("</TableStore><TableName>").Raw(Xml.Esc(h.TableName)).Raw("</TableName></Properties>");
                    hc.Close();
                }
            });
        });
        x.Close();
        return x.ToString();
    }

    private static void Relationship(XmBuilder x, EncodedRelationship r)
    {
        x.Open("XMRelationship", r.Name);
        x.Raw("<Properties><PrimaryTable>").Raw(Xml.Esc(r.PrimaryTableId)).Raw("</PrimaryTable><PrimaryColumn>")
         .Raw(Xml.Esc(r.PrimaryColumnId)).Raw("</PrimaryColumn><ForeignColumn>").Raw(Xml.Esc(r.ForeignColumnId)).Raw("</ForeignColumn></Properties>");
        x.DataObjects(d => d.DataObject(o =>
        {
            o.Open("XMRelationshipIndexDenseDIDs");
            o.Raw("<Properties>");
            o.Long("Records", r.Records);
            o.Raw("<TableName>").Raw(Xml.Esc(r.TableName)).Raw("</TableName>");
            o.Long("Flags", 0);
            o.Raw("</Properties>");
            o.Close();
        }));
        x.Close();
    }

    /// <summary>The U$ table holding the tree of a user hierarchy (four 32-bit columns, one record per node).</summary>
    public static string HierarchyTable(EncodedHierarchy h)
    {
        var x = new XmBuilder();
        x.Open("XMSimpleTable", h.TableName, root: true);
        x.Properties(p => p.Int("Version", FileNames.HierarchyVersion).Long("Settings", 5).Long("RIViolationCount", 0));
        x.Members(m =>
        {
            m.Member("SegmentMap", s =>
            {
                s.Open("XMSegment1Map");
                s.Properties(p => p.Long("Records", h.Records));
                s.Close();
            });
            m.Member("TableStats", s => TableStats(s));
        });
        x.Collections(c =>
        {
            c.Collection("Partitions", null);
            c.Collection("Columns", cc =>
            {
                var columns = new (string name, int settings, int min)[]
                {
                    ("MULTI_LEVEL_ID", 16, 0), ("PARENT_POS", 8, -1), ("FIRST_CHILD_POS", 9, -1), ("CHILD_COUNT", 17, -1),
                };
                foreach (var (name, settings, min) in columns)
                {
                    cc.Open("XMRawColumn", name, providerVersion: 1);
                    cc.Properties(p => p.Long("Settings", settings).Long("ColumnFlags", 0).Raw("<Collation/><OrderByColumn/>")
                        .Long("Locale", 0).UInt("BinaryCharacters", 10240));
                    cc.Members(m => m.Member("ColumnStats", s =>
                    {
                        s.Open("XMColumnStats");
                        s.Properties(p => p.Int("DistinctStates", 0).Int("MinDataID", 2).Int("MaxDataID", 2)
                            .Int("OriginalMinSegmentDataID", 2).Long("RLESortOrder", -1).Long("RowCount", 0)
                            .Bool("HasNulls", false).Long("RLERuns", 0).Long("OthersRLERuns", 0).Int("Usage", 3)
                            .Short("DBType", 0).Int("XMType", 0).Int("CompressionType", 0).Long("CompressionParam", 0)
                            .Int("EncodingHint", 0).Long("AggCounter", 0).Long("WhereCounter", 0).Long("OrderByCounter", 0));
                        s.Close();
                    }));
                    cc.Collections(sc => sc.Collection("Segments", segs =>
                    {
                        segs.Open("XMColumnSegment");
                        segs.Properties(p => p.Long("Records", h.Records).Long("Mask", 0));
                        segs.Members(sm =>
                        {
                            sm.Member("CompressionInfo", ci => { ci.Open($"XMRENoSplitCompressionInfo{LT}32>"); ci.Properties(p => p.Int("Min", min)); ci.Close(); });
                            sm.Member("ColumnSegmentStats", s =>
                            {
                                s.Open("XMColumnSegmentStats");
                                s.Properties(p => p.Long("DistinctStates", 0).Int("MinDataID", 2).Int("MaxDataID", 2)
                                    .Int("OriginalMinSegmentDataID", 2).Long("RLESortOrder", -1).Long("RowCount", 0)
                                    .Bool("HasNulls", false).Long("RLERuns", 0).Long("OthersRLERuns", 0));
                                s.Close();
                            });
                        });
                        segs.Close();
                    }));
                    cc.DataObjects(d => d.DataObject(o =>
                    {
                        o.Open("XMRawColumnPartitionDataObject", FileNames.UserHierarchyColumn(h.TableId, h.Name, name));
                        o.Properties(p => p.Int("DataVersion", FileNames.HierarchyVersion).Int("Partition", 0).Int("SegmentCount", 1));
                        o.Close();
                    }));
                    cc.Close();
                }
            });
            c.Collection("Relationships", null);
            c.Collection("UserHierarchies", null);
        });
        x.Close();
        return x.ToString();
    }

    /// <summary>The R$ table holding a relationship's INDEX column (data id of the foreign key -> primary row).</summary>
    public static string RelationshipTable(EncodedRelationship r)
    {
        long distinct = r.MaxRow + 2;
        // Excel writes these placeholder statistics for relationship tables while a model is open; Sort by Column needs them.
        // Excel itself writes these placeholder statistics for relationship tables while a model is open. Fully processed
        // statistics made Sort by Column fail in Excel (see README, Open questions).
        long cDistinct = 0; int cMin = 2; int cMax = 2; long cRows = 0; int cType = 0; var cW = r.Width;
        var x = new XmBuilder();
        x.Open("XMSimpleTable", r.TableName, root: true);
        x.Properties(p => p.Int("Version", 0).Long("Settings", 2).Long("RIViolationCount", r.ViolationCount));
        x.Members(m =>
        {
            m.Member("SegmentMap", s =>
            {
                s.Open("XMSegment1Map");
                s.Properties(p => p.Long("Records", r.Records));
                s.Close();
            });
            m.Member("TableStats", s => TableStats(s));
        });
        x.Collections(c =>
        {
            c.Collection("Partitions", null);
            c.Collection("Columns", cc =>
            {
                cc.Open("XMRawColumn", "INDEX", providerVersion: 1);
                cc.Properties(p => p.Long("Settings", 3).Long("ColumnFlags", 0).Raw("<Collation/><OrderByColumn/>")
                    .Long("Locale", 0).UInt("BinaryCharacters", 0));
                cc.Members(m => m.Member("ColumnStats", s =>
                {
                    s.Open("XMColumnStats");
                    s.Properties(p => p.Int("DistinctStates", (int)cDistinct).Int("MinDataID", cMin).Int("MaxDataID", cMax)
                        .Int("OriginalMinSegmentDataID", 2).Long("RLESortOrder", -1).Long("RowCount", cRows)
                        .Bool("HasNulls", false).Long("RLERuns", 0).Long("OthersRLERuns", 0).Int("Usage", 3)
                        .Short("DBType", 0).Int("XMType", cType).Int("CompressionType", 0).Long("CompressionParam", 0)
                        .Int("EncodingHint", 0).Long("AggCounter", 0).Long("WhereCounter", 0).Long("OrderByCounter", 0));
                    s.Close();
                }));
                cc.Collections(sc => sc.Collection("Segments", segs =>
                {
                    segs.Open("XMColumnSegment");
                    segs.Properties(p => p.Long("Records", r.Records).Long("Mask", 0));
                    segs.Members(sm =>
                    {
                        sm.Member("CompressionInfo", ci => { ci.Open($"XMRENoSplitCompressionInfo{LT}{cW}>"); ci.Properties(p => p.Int("Min", -1)); ci.Close(); });
                        sm.Member("ColumnSegmentStats", s =>
                        {
                            s.Open("XMColumnSegmentStats");
                            s.Properties(p => p.Long("DistinctStates", cDistinct).Int("MinDataID", cMin).Int("MaxDataID", cMax)
                                .Int("OriginalMinSegmentDataID", 2).Long("RLESortOrder", -1).Long("RowCount", cRows)
                                .Bool("HasNulls", false).Long("RLERuns", 0).Long("OthersRLERuns", 0));
                            s.Close();
                        });
                    });
                    segs.Close();
                }));
                cc.DataObjects(d => d.DataObject(o =>
                {
                    o.Open("XMRawColumnPartitionDataObject", FileNames.RelationshipIndex(r.ForeignTableId, r.Name));
                    o.Properties(p => p.Int("DataVersion", 0).Int("Partition", 0).Int("SegmentCount", 1));
                    o.Close();
                }));
                cc.Close();
            });
            c.Collection("Relationships", null);
            c.Collection("UserHierarchies", null);
        });
        x.Close();
        return x.ToString();
    }

    private static void TableStats(XmBuilder s)
    {
        s.Open("XMTableStats");
        s.Properties(p => p.Long("SegmentSize", 0).Int("Usage", 0));
        s.Close();
    }

    private static void RawColumn(XmBuilder x, EncodedTable t, EncodedColumn col, int lcid)
    {
        x.Open("XMRawColumn", col.Id, providerVersion: 1);
        x.Properties(p => p
            .Long("Settings", 1025)
            .Long("ColumnFlags", col.IsRowNumber ? 31 : col.IsPrimaryKey ? 11 : 8)
            .Raw(col.OrderByColumnId is null ? "<Collation/><OrderByColumn/>" : "<Collation/><OrderByColumn>" + System.Security.SecurityElement.Escape(col.OrderByColumnId) + "</OrderByColumn>")
            .Long("Locale", lcid)
            .UInt("BinaryCharacters", 0));
        x.Members(m =>
        {
            m.Member("IntrinsicHierarchy", h =>
            {
                h.Open("XMHierarchy", $"[Hierarchy for column {col.Id}]");
                h.Properties(p =>
                {
                    p.Int("SortOrder", 0).Bool("IsProcessed", true);
                    if (col.IsRowNumber)
                        p.Int("TypeMaterialization", 3).Long("ColumnPosition2DataID", -1).Long("ColumnDataID2Position", -1);
                    else if (col.IsValueEncoded)
                        p.Int("TypeMaterialization", 1).Long("ColumnPosition2DataID", 0).Long("ColumnDataID2Position", -1);
                    else
                        p.Int("TypeMaterialization", 0).Long("ColumnPosition2DataID", 0).Long("ColumnDataID2Position", 1);
                    p.Long("DistinctDataIDs", col.DistinctInclBlank);
                    if (col.IsRowNumber) p.Raw("<TableStore/>");
                    else p.Raw("<TableStore>").Raw(Xml.Esc(FileNames.HierarchyTable(t.Id, col.Id))).Raw("</TableStore>");
                });
                if (!col.IsRowNumber && col.IsValueEncoded)
                    h.DataObjects(d => d.DataObject(o => o.Empty("XMHierarchyDataID2PositionHashIndex", FileNames.Hidx(t.Id, col.Id))));
                h.Close();
            });
            m.Member("ColumnStats", s =>
            {
                s.Open("XMColumnStats");
                s.Properties(p => p
                    .Int("DistinctStates", col.IsValueEncoded ? col.DistinctInclBlank : 0)
                    .Int("MinDataID", col.MinDataId)
                    .Int("MaxDataID", col.MaxDataId)
                    .Int("OriginalMinSegmentDataID", DataIds.Null)
                    .Long("RLESortOrder", -1)
                    .Long("RowCount", col.Rows)
                    .Bool("HasNulls", col.HasNulls)
                    .Long("RLERuns", col.RleRuns)
                    .Long("OthersRLERuns", col.BitPackedRuns)
                    .Int("Usage", col.IsPrimaryKey ? 0 : col.IsForeignKey ? 1 : 3)
                    .Short("DBType", col.DbType)
                    .Int("XMType", (int)col.XmType)
                    .Int("CompressionType", 0)
                    .Long("CompressionParam", 0)
                    .Int("EncodingHint", t.Hierarchies.Count == 0 && (col.IsRowNumber || col.IsPrimaryKey) ? 1 : 0)
                    .Long("AggCounter", 0)
                    .Long("WhereCounter", 0)
                    .Long("OrderByCounter", 0));
                s.Close();
            });
        });
        x.Collections(c => c.Collection("Segments", sc =>
        {
            foreach (var seg in col.Segments) Segment(sc, col, seg);
        }));
        x.DataObjects(d =>
        {
            d.DataObject(o => Dictionary(o, t, col));
            d.DataObject(o =>
            {
                o.Open("XMRawColumnPartitionDataObject", FileNames.Idf(t.Id, col.Id));
                o.Properties(p => p.Int("DataVersion", 0).Int("Partition", 0).Int("SegmentCount", col.Segments.Count));
                o.Close();
            });
        });
        x.Close();
    }

    private static void Dictionary(XmBuilder o, EncodedTable t, EncodedColumn col)
    {
        switch (col.Dictionary)
        {
            case DictionaryKind.Value:
                o.Open(col.XmType == XmType.Real ? $"XMValueDataDictionary{LT}XM_Real>" : $"XMValueDataDictionary{LT}XM_Long>");
                o.Properties(p => p.Int("DataVersion", 0).Long("BaseId", col.BaseId).Double("Magnitude", Xml.Magnitude(col.Magnitude)));
                break;
            case DictionaryKind.String:
                o.Open($"XMHashDataDictionary{LT}XM_String>", FileNames.Dictionary(t.Id, col.Id));
                o.Properties(p => p.Int("DataVersion", 0).Int("LastId", col.LastId).Bool("Nullable", col.HasNulls)
                    .Bool("Unique", col.IsPrimaryKey).Long("DictionaryFlags", 3));
                break;
            case DictionaryKind.Long32:
            case DictionaryKind.Long64:
                o.Open($"XMHashDataDictionary{LT}XM_Long>", FileNames.Dictionary(t.Id, col.Id));
                o.Properties(p => p.Int("DataVersion", 0).Int("LastId", col.LastId).Bool("Nullable", col.HasNulls)
                    .Bool("Unique", col.IsPrimaryKey).Bool("OperatingOn32", col.Dictionary == DictionaryKind.Long32));
                break;
            case DictionaryKind.Real:
                o.Open($"XMHashDataDictionary{LT}XM_Real>", FileNames.Dictionary(t.Id, col.Id));
                o.Properties(p => p.Int("DataVersion", 0).Int("LastId", col.LastId).Bool("Nullable", col.HasNulls)
                    .Bool("Unique", col.IsPrimaryKey));
                break;
        }
        o.Close();
    }

    private static void EmptySegmentStats(XmBuilder s)
    {
        s.Open("XMColumnSegmentStats");
        s.Properties(p => p.Long("DistinctStates", 0).Int("MinDataID", 2).Int("MaxDataID", 2).Int("OriginalMinSegmentDataID", 2)
            .Long("RLESortOrder", -1).Long("RowCount", 0).Bool("HasNulls", false).Long("RLERuns", 0).Long("OthersRLERuns", 0));
        s.Close();
    }

    private static void Segment(XmBuilder x, EncodedColumn col, EncodedSegment seg)
    {
        string sub = seg.IsRowNumber ? "XM123CompressionInfo" : $"XMRENoSplitCompressionInfo{LT}{seg.Width}>";
        string hybrid = seg.IsRowNumber
            ? $"XMHybridRLECompressionInfo{LT}class XM123CompressionInfo>"
            : $"XMHybridRLECompressionInfo{LT}class XMRENoSplitCompressionInfo{LT}{seg.Width}>>";

        x.Open("XMColumnSegment");
        x.Properties(p => p.Long("Records", seg.Records).Long("Mask", seg.IsRowNumber ? 0 : 1));
        x.Members(m =>
        {
            m.Member("SubSegment", s =>
            {
                s.Open("XMColumnSegment");
                s.Properties(p => p.Long("Records", seg.SubRecords).Long("Mask", 0));
                s.Members(sm =>
                {
                    sm.Member("CompressionInfo", ci => { ci.Open(sub); ci.Properties(p => p.Int("Min", seg.Min)); ci.Close(); });
                    sm.Member("ColumnSegmentStats", EmptySegmentStats);
                });
                s.Close();
            });
            m.Member("CompressionInfo", ci =>
            {
                ci.Open(hybrid);
                ci.Members(cm =>
                {
                    cm.Member("RLECompression", r =>
                    {
                        r.Open("XMRLECompressionInfo");
                        r.Properties(p => p.Long("BookmarkBits", seg.BookmarkBits).Long("StorageAllocSize", seg.StorageAllocSize)
                            .Long("StorageUsedSize", seg.StorageUsedSize).Bool("SegmentNeedsResizing", false));
                        r.Close();
                    });
                    cm.Member("SubCompression", sc => { sc.Open(sub); sc.Properties(p => p.Int("Min", seg.Min)); sc.Close(); });
                });
                ci.Close();
            });
            m.Member("ColumnSegmentStats", s =>
            {
                s.Open("XMColumnSegmentStats");
                uint omin = seg.IsRowNumber || !col.IsValueEncoded ? DataIds.Null : DataIds.First;
                s.Properties(p => p.Long("DistinctStates", 0)
                    .Int("MinDataID", seg.MinDataId)
                    .Int("MaxDataID", seg.MaxDataId)
                    .Int("OriginalMinSegmentDataID", omin)
                    .Long("RLESortOrder", -1)
                    .Long("RowCount", seg.Records)
                    .Bool("HasNulls", seg.HasNulls)
                    .Long("RLERuns", seg.RleRuns)
                    .Long("OthersRLERuns", seg.BitPackedRuns));
                s.Close();
            });
        });
        x.Close();
    }

    /// <summary>The H$ table of a column's attribute hierarchy.</summary>
    public static string HierarchyTable(EncodedTable t, EncodedColumn col)
    {
        long records = col.HierarchyRecords;
        var segments = col.HierarchySegments;
        long perSegment = segments[0];
        var x = new XmBuilder();
        x.Open("XMSimpleTable", FileNames.HierarchyTable(t.Id, col.Id), root: true);
        x.Properties(p => p.Int("Version", 0).Long("Settings", 4).Long("RIViolationCount", 0));
        x.Members(m =>
        {
            m.Member("SegmentMap", s =>
            {
                s.Open($"XMSegmentEqualMapEx{LT}XMSegmentEqualMap_ComplexInstantiation>");
                s.Properties(p => p.Long("Segments", segments.Count).Long("Records", records).Long("RecordsPerSegment", perSegment));
                s.Close();
            });
            m.Member("TableStats", s => TableStats(s));
        });
        x.Collections(c =>
        {
            c.Collection("Partitions", null);
            c.Collection("Columns", cc =>
            {
                SystemColumn(cc, "POS_TO_ID", 7, FileNames.PosToId(t.Id, col.Id), segments);
                if (!col.IsValueEncoded)
                    SystemColumn(cc, "ID_TO_POS", 5, FileNames.IdToPos(t.Id, col.Id), segments);
            });
            c.Collection("Relationships", null);
            c.Collection("UserHierarchies", null);
        });
        x.Close();
        return x.ToString();
    }

    private static void SystemColumn(XmBuilder x, string name, long settings, string file, IReadOnlyList<long> segments)
    {
        x.Open("XMRawColumn", name, providerVersion: 1);
        x.Properties(p => p.Long("Settings", settings).Long("ColumnFlags", 0).Raw("<Collation/><OrderByColumn/>")
            .Long("Locale", 0).UInt("BinaryCharacters", 0));
        x.Members(m => m.Member("ColumnStats", s =>
        {
            s.Open("XMColumnStats");
            s.Properties(p => p.Int("DistinctStates", 0).Int("MinDataID", 2).Int("MaxDataID", 2).Int("OriginalMinSegmentDataID", 2)
                .Long("RLESortOrder", -1).Long("RowCount", 0).Bool("HasNulls", false).Long("RLERuns", 0).Long("OthersRLERuns", 0)
                .Int("Usage", 3).Short("DBType", 0).Int("XMType", 0).Int("CompressionType", 0).Long("CompressionParam", 0)
                .Int("EncodingHint", 0).Long("AggCounter", 0).Long("WhereCounter", 0).Long("OrderByCounter", 0));
            s.Close();
        }));
        x.Collections(c => c.Collection("Segments", sc =>
        {
            foreach (long records in segments)
            {
                sc.Open("XMColumnSegment");
                sc.Properties(p => p.Long("Records", records).Long("Mask", 0));
                sc.Members(m =>
                {
                    m.Member("CompressionInfo", ci => { ci.Open($"XMRENoSplitCompressionInfo{LT}32>"); ci.Properties(p => p.Int("Min", 0)); ci.Close(); });
                    m.Member("ColumnSegmentStats", EmptySegmentStats);
                });
                sc.Close();
            }
        }));
        x.DataObjects(d => d.DataObject(o =>
        {
            o.Open("XMRawColumnPartitionDataObject", file);
            o.Properties(p => p.Int("DataVersion", 0).Int("Partition", 0).Int("SegmentCount", segments.Count));
            o.Close();
        }));
        x.Close();
    }
}
