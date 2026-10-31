using System.Text;
using System.Xml;
using PivotWright.Internal.Model;

namespace PivotWright.Internal.Metadata;

/// <summary>Identity and settings shared by all metadata documents of one model.</summary>
internal sealed class ModelInfo
{
    public required string DatabaseId { get; init; }
    public string DatabaseName { get; init; } = "Microsoft_SQLServer_AnalysisServices";
    public required string DatabaseObjectId { get; init; }
    public required string CubeObjectId { get; init; }
    public required string MdxScriptObjectId { get; init; }
    public required string DataSourceId { get; init; }
    public required string DataSourceName { get; init; }
    public required string DataSourceObjectId { get; init; }
    public required string DsvObjectId { get; init; }
    public required int Lcid { get; init; }
    public required string CultureName { get; init; }
    public required System.Globalization.CultureInfo Culture { get; init; }
    public required DateTime Created { get; init; }
    public required DateTime Processed { get; init; }
    public required string EncryptedConnectionString { get; init; }
    public required string EncryptedDatabasePassword { get; init; }
    public required string EncryptedDataSourcePassword { get; init; }
    public required IReadOnlyDictionary<string, TableIds> TableIds { get; init; }

    /// <summary>Ids of the tables that hold user hierarchies. Excel saves such models with higher object/structure versions.</summary>
    public IReadOnlySet<string> HierarchyTables { get; init; } = new HashSet<string>();
    public bool AnyHierarchy => HierarchyTables.Count > 0;

    /// <summary>Ids of the tables marked as date table (SPEC section 9). Excel saves such models one more edit further on.</summary>
    public IReadOnlySet<string> DateTables { get; init; } = new HashSet<string>();
    /// <summary>Ids of the tables on the many side of a relationship to a date table.</summary>
    public IReadOnlySet<string> DateReferencingTables { get; init; } = new HashSet<string>();
    public bool AnyDate => DateTables.Count > 0;
    /// <summary>
    /// Measure groups get distinct Ordinals (and the partitions and script 1), as Excel numbers them. With every Ordinal 0 the engine tells
    /// the measure groups apart only by hash order of their ids, so whether Create Hierarchy or Sort by Column worked depended on the ids
    /// (mddetail.cpp, ASDetail::OnClone2). Written for every model.
    /// </summary>
    public bool UniqueOrdinals => true;

    /// <summary>Ids of the tables with at least one column hidden from client tools.</summary>
    public IReadOnlySet<string> HiddenTables { get; init; } = new HashSet<string>();
    /// <summary>
    /// Models with hierarchies and hidden columns are written the way Excel saves them right after a hide operation
    /// (H8_hide_one_column.xlsx, H9_hide_one_hierarchy_level.xlsx; SPEC section 8). Without these counters Excel loaded and displayed
    /// such a model but failed every later edit in ASDetail::OnClone2. Not used together with a date table (verified without).
    /// </summary>
    public bool HiddenProfile => HiddenTables.Count > 0 && AnyHierarchy && !AnyDate;
    /// <summary>
    /// When the hide operation happened, in the hidden profile: Excel keeps CreatedTimestamp at the original creation and stamps the
    /// edited dimensions and the cube with the edit time in LastSchemaUpdate (the cube's LastUpdate a fraction later), as in H8 and H9.
    /// </summary>
    public DateTime EditedAt { get; init; }

    /// <summary>
    /// How Excel numbers a model with calculated columns, or null without any (SPEC section 10). The models are written the way
    /// Excel saves them after columns were added in manual calculation mode.
    /// </summary>
    public CalculatedPlan? Calc { get; init; }
    public bool CalculatedProfile => Calc is not null;

    // Version counters as Excel saves them after "Create Hierarchy" and after "Mark as Date Table" (SPEC sections 5 and 9).
    public int CubeVersion => Calc?.Cube ?? (AnyDate ? 21 : HiddenProfile ? 21 : AnyHierarchy ? 18 : AsslWriter.CubeObjectVersion);
    /// <summary>Object version of the measure groups: the cube's, except after a hide operation, where Excel keeps them one step behind.</summary>
    public int MeasureGroupVersion => Calc?.MeasureGroup ?? (HiddenProfile ? 20 : CubeVersion);
    public int ScriptVersion => Calc?.Script ?? (AnyDate ? 10 : AnyHierarchy ? 9 : AsslWriter.CubeObjectVersion);
    public int PartitionVersion => Calc?.Partition ?? (AnyDate ? 10 : AnyHierarchy ? 9 : 1);
    /// <summary>Object version of a measure group: 1 for the linked trigger table, which Excel creates at version 1 (H42).</summary>
    public int MeasureGroupVersionOf(string tableId) => LinkedTables.Contains(tableId) ? LinkedVersion : MeasureGroupVersion;
    /// <summary>Object version of a partition: 1 for the linked trigger table (H42).</summary>
    public int PartitionVersionOf(string tableId) => LinkedTables.Contains(tableId) ? LinkedVersion : PartitionVersion;
    /// <summary>Ids of the linked worksheet tables (the trigger table of <see cref="DataModelWriter.AddCalculationTrigger"/>).</summary>
    public IReadOnlySet<string> LinkedTables { get; init; } = new HashSet<string>();
    public const int LinkedVersion = 1;
    /// <summary>Dimension version of a linked table: 2, as Excel writes it (H42).</summary>
    public const int LinkedDimensionVersion = 2;
    /// <summary>True when Excel saved the table's dimension again after an edit (H10), which drops the empty annotation values.</summary>
    public bool Resaved(string tableId) => Calc is not null && Calc.Resaved(tableId);
    public int DimensionVersion(string tableId, bool manySide) =>
        Calc is not null && Calc.Tables.TryGetValue(tableId, out var calc) ? calc.DimensionVersion
        : (HierarchyTables.Contains(tableId) ? 5 : manySide ? (AnyHierarchy ? 3 : 2) : 1) + (AnyDate ? 1 : 0)
        + (HiddenProfile && HiddenTables.Contains(tableId) ? (HierarchyTables.Contains(tableId) ? 2 : 1) : 0);
    /// <summary>Version of a table's .tbl.xml: 2 when it holds hierarchies, plus 1 for a date table or a table related to one.</summary>
    public int TableVersion(string tableId) =>
        Calc is not null && Calc.Tables.TryGetValue(tableId, out var calc) && calc.TableVersion is { } version ? version :
        (HierarchyTables.Contains(tableId) ? FileNames.HierarchyVersion : 0) + (DateTables.Contains(tableId) || DateReferencingTables.Contains(tableId) ? 1 : 0);
}

internal sealed class TableIds
{
    public required string DimensionObjectId { get; init; }
    public required string MeasureGroupObjectId { get; init; }
    public required string PartitionObjectId { get; init; }
    /// <summary>1-based position of the table; Excel numbers measure groups with it after a date table was marked.</summary>
    public int Ordinal { get; init; }
}

/// <summary>
/// Writes the Analysis Services (ASSL, compatibility level 1103) object definitions that make up
/// the tabular model ([MS-XLDM] 2.6), mirroring the documents Excel writes.
/// </summary>
internal static class AsslWriter
{
    private const string LoadOpen =
        "<Load xmlns=\"http://schemas.microsoft.com/analysisservices/2003/engine\" xmlns:ddl2=\"http://schemas.microsoft.com/analysisservices/2003/engine/2\" xmlns:ddl2_2=\"http://schemas.microsoft.com/analysisservices/2003/engine/2/2\" xmlns:ddl100=\"http://schemas.microsoft.com/analysisservices/2008/engine/100\" xmlns:ddl100_100=\"http://schemas.microsoft.com/analysisservices/2008/engine/100/100\" xmlns:ddl200=\"http://schemas.microsoft.com/analysisservices/2010/engine/200\" xmlns:ddl200_200=\"http://schemas.microsoft.com/analysisservices/2010/engine/200/200\" xmlns:ddl300=\"http://schemas.microsoft.com/analysisservices/2011/engine/300\" xmlns:ddl300_300=\"http://schemas.microsoft.com/analysisservices/2011/engine/300/300\" xmlns:ddl400=\"http://schemas.microsoft.com/analysisservices/2012/engine/400\" xmlns:ddl400_400=\"http://schemas.microsoft.com/analysisservices/2012/engine/400/400\" xmlns:ddl410=\"http://schemas.microsoft.com/analysisservices/2012/engine/410\" xmlns:ddl410_410=\"http://schemas.microsoft.com/analysisservices/2012/engine/410/410\" xmlns:ddl500=\"http://schemas.microsoft.com/analysisservices/2013/engine/500\" xmlns:ddl500_500=\"http://schemas.microsoft.com/analysisservices/2013/engine/500/500\" xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\">";

    private const string NsA =
        " xmlns:ddl500_500=\"http://schemas.microsoft.com/analysisservices/2013/engine/500/500\" xmlns:ddl500=\"http://schemas.microsoft.com/analysisservices/2013/engine/500\" xmlns:ddl400_400=\"http://schemas.microsoft.com/analysisservices/2012/engine/400/400\" xmlns:ddl400=\"http://schemas.microsoft.com/analysisservices/2012/engine/400\" xmlns:ddl300_300=\"http://schemas.microsoft.com/analysisservices/2011/engine/300/300\" xmlns:ddl300=\"http://schemas.microsoft.com/analysisservices/2011/engine/300\" xmlns:ddl200_200=\"http://schemas.microsoft.com/analysisservices/2010/engine/200/200\" xmlns:ddl200=\"http://schemas.microsoft.com/analysisservices/2010/engine/200\" xmlns:ddl100_100=\"http://schemas.microsoft.com/analysisservices/2008/engine/100/100\" xmlns:ddl2_2=\"http://schemas.microsoft.com/analysisservices/2003/engine/2/2\" xmlns:ddl2=\"http://schemas.microsoft.com/analysisservices/2003/engine/2\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\" xmlns=\"http://schemas.microsoft.com/analysisservices/2003/engine\"";

    private const string NsB = NsA +
        " xmlns:ddl410_410=\"http://schemas.microsoft.com/analysisservices/2012/engine/410/410\" xmlns:ddl410=\"http://schemas.microsoft.com/analysisservices/2012/engine/410\" xmlns:ddl100=\"http://schemas.microsoft.com/analysisservices/2008/engine/100\"";

    private const string ProactiveCaching =
        "<ProactiveCaching><SilenceInterval>-PT1S</SilenceInterval><Latency>-PT1S</Latency><SilenceOverrideInterval>-PT1S</SilenceOverrideInterval><ForceRebuildInterval>-PT1S</ForceRebuildInterval><Enabled>false</Enabled><AggregationStorage>Regular</AggregationStorage><OnlineMode>Immediate</OnlineMode><Source xsi:type=\"ProactiveCachingInheritedBinding\"><NotificationTechnique>Client</NotificationTechnique></Source></ProactiveCaching>";

    private const string ErrorConfiguration =
        "<ErrorConfiguration><KeyErrorLimit>0</KeyErrorLimit><KeyErrorLogFile/><KeyErrorAction>ConvertToUnknown</KeyErrorAction><KeyErrorLimitAction>StopProcessing</KeyErrorLimitAction><KeyDuplicate>ReportAndStop</KeyDuplicate><KeyNotFound>IgnoreError</KeyNotFound><NullKeyConvertedToUnknown>IgnoreError</NullKeyConvertedToUnknown><NullKeyNotAllowed>ReportAndStop</NullKeyNotAllowed></ErrorConfiguration>";

    private const string TableAnnotations =
        "\n            <Annotation>\n              <Name>IsQueryEditorUsed</Name>\n              <Value>False</Value>\n            </Annotation>\n            <Annotation>\n              <Name>QueryEditorSerialization</Name>\n              <Value/>\n            </Annotation>\n            <Annotation>\n              <Name>TableWidgetSerialization</Name>\n              <Value/>\n            </Annotation>\n          ";

    /// <summary>Excel leaves the empty Value elements out when it saves a table again after an edit (H10).</summary>
    private static string Annotations(bool saved) => saved ? TableAnnotations.Replace("\n              <Value/>", "") : TableAnnotations;

    private static string E(string s) => Xml.Esc(s);
    private static string T(DateTime d) => Xml.Time(d);

    /// <summary>
    /// Object version of the cube, its measure groups and the calculation script. Excel writes 1 for a model it saved after
    /// creating relationships; models written at version 0 failed Create Hierarchy in Excel (README, Open questions).
    /// </summary>
    internal const int CubeObjectVersion = 1;

    private static void CommonHeader(StringBuilder sb, string name, string id, DateTime created, int objectVersion, string objectId, int ordinal = 0, DateTime? schemaUpdated = null)
    {
        sb.Append("<Name>").Append(E(name)).Append("</Name><ID>").Append(E(id)).Append("</ID>");
        sb.Append("<CreatedTimestamp>").Append(T(created)).Append("</CreatedTimestamp>");
        sb.Append("<LastSchemaUpdate>").Append(T(schemaUpdated ?? created)).Append("</LastSchemaUpdate>");
        sb.Append("<ObjectVersion>").Append(objectVersion).Append("</ObjectVersion>");
        sb.Append("<ObjectID>").Append(objectId).Append("</ObjectID>");
        sb.Append("<Ordinal>").Append(ordinal).Append("</Ordinal><PersistLocation>0</PersistLocation><System>false</System>");
    }

    public static string Database(ModelInfo m)
    {
        var sb = new StringBuilder(LoadOpen);
        sb.Append("<ObjectDefinition><Database>");
        CommonHeader(sb, m.DatabaseName, m.DatabaseId, m.Created, 1, m.DatabaseObjectId);
        sb.Append("<DataFileList>0.CryptKey.bin</DataFileList><Description/>");
        sb.Append("<Annotations").Append(NsA).Append(">\n        <Annotation>\n          <Name>ClientCompatibilityLevel</Name>\n          <Value>300</Value>\n        </Annotation>\n      </Annotations>");
        sb.Append("<LastProcessed>").Append(T(m.Processed)).Append("</LastProcessed><AggregationPrefix/>");
        sb.Append("<Language>").Append(m.Lcid).Append("</Language><Collation/>");
        sb.Append("<ddl400:DefaultCollationVersion>100</ddl400:DefaultCollationVersion><Visible>true</Visible><MasterDataSourceID/><ProcessingPriority>0</ProcessingPriority>");
        sb.Append("<ddl200_200:StorageEngineUsed>InMemory</ddl200_200:StorageEngineUsed><ddl200:CompatibilityLevel>1103</ddl200:CompatibilityLevel><Translations/>");
        sb.Append("<DataSourceImpersonationInfo><ImpersonationMode>Default</ImpersonationMode><Account/><Password>")
          .Append(m.EncryptedDatabasePassword).Append("</Password></DataSourceImpersonationInfo>");
        sb.Append("<DataVersion>0</DataVersion></Database></ObjectDefinition></Load>");
        return sb.ToString();
    }

    public static string DataSource(ModelInfo m)
    {
        var sb = new StringBuilder(LoadOpen);
        sb.Append("<ParentObject><DatabaseID>").Append(m.DatabaseId).Append("</DatabaseID></ParentObject>");
        sb.Append("<ObjectDefinition><DataSource xsi:type=\"RelationalDataSource\">");
        CommonHeader(sb, m.DataSourceName, m.DataSourceId, m.Created, 0, m.DataSourceObjectId);
        sb.Append("<DataFileList/><Description/>");
        sb.Append("<Annotations").Append(NsA).Append(">\n            <Annotation>\n              <Name>ConnectionEditUISource</Name>\n              <Value>SqlServer</Value>\n            </Annotation>\n          </Annotations>");
        sb.Append("<PermissionFileList/><ConnectionString>[EncryptedDataStart]").Append(m.EncryptedConnectionString).Append("[EncryptedDataEnd]</ConnectionString>");
        sb.Append("<Timeout>P0D</Timeout><ImpersonationInfo><ImpersonationMode>ImpersonateCurrentUser</ImpersonationMode><Account/><Password>")
          .Append(m.EncryptedDataSourcePassword).Append("</Password></ImpersonationInfo>");
        sb.Append("<MaxActiveConnections>10</MaxActiveConnections><Isolation>ReadCommitted</Isolation></DataSource></ObjectDefinition></Load>");
        return sb.ToString();
    }

    private static string DsvType(ColumnDataType t) => t switch
    {
        ColumnDataType.Int64 => "xs:long",
        ColumnDataType.Double => "xs:double",
        ColumnDataType.Currency => "xs:decimal",
        ColumnDataType.DateTime => "xs:dateTime",
        ColumnDataType.Boolean => "xs:boolean",
        _ => "xs:string",
    };

    public static string DataSourceView(ModelInfo m, IReadOnlyList<EncodedTable> tables)
    {
        var sb = new StringBuilder(LoadOpen);
        sb.Append("<ParentObject><DatabaseID>").Append(m.DatabaseId).Append("</DatabaseID></ParentObject>");
        sb.Append("<ObjectDefinition><DataSourceView>");
        CommonHeader(sb, "Sandbox", "Sandbox", m.Created, 0, m.DsvObjectId);
        sb.Append("<DataFileList/><Description/>");
        sb.Append("<Schema").Append(NsA).Append(">\n");
        sb.Append("            <xs:schema id=\"NewDataSet\" xmlns=\"\" xmlns:xs=\"http://www.w3.org/2001/XMLSchema\" xmlns:msdata=\"urn:schemas-microsoft-com:xml-msdata\" xmlns:msprop=\"urn:schemas-microsoft-com:xml-msprop\">\n");
        sb.Append("              <xs:element name=\"NewDataSet\" msdata:IsDataSet=\"true\" msdata:Locale=\"").Append(E(m.CultureName)).Append("\">\n");
        sb.Append("                <xs:complexType>\n                  <xs:choice minOccurs=\"0\" maxOccurs=\"unbounded\">\n");
        foreach (var t in tables)
        {
            string q = "     SELECT [" + t.Name.Replace("]", "]]") + "].*   FROM [" + t.Name.Replace("]", "]]") + "] ";
            sb.Append("                    <xs:element name=\"").Append(E(XmlConvert.EncodeLocalName(t.Id)))
              .Append("\" msdata:Locale=\"\" msprop:IsLogical=\"True\" msprop:Description=\"").Append(E(t.Name))
              .Append("\" msprop:QueryDefinition=\"").Append(E(q))
              .Append("\" msprop:FriendlyName=\"").Append(E(t.Name))
              .Append("\" msprop:TableType=\"View\" msprop:DbTableName=\"").Append(E(t.Name)).Append("\">\n");
            sb.Append("                      <xs:complexType>\n                        <xs:sequence>\n");
            foreach (var c in t.Columns)
            {
                sb.Append("                          <xs:element name=\"").Append(E(XmlConvert.EncodeLocalName(c.Id)))
                  .Append("\" msprop:FriendlyName=\"").Append(E(c.Name))
                  .Append("\" msprop:DbColumnName=\"").Append(E(c.Name)).Append('"');
                if (c.DataType == ColumnDataType.String)
                {
                    sb.Append(" minOccurs=\"0\">\n                            <xs:simpleType>\n                              <xs:restriction base=\"xs:string\">\n                                <xs:maxLength value=\"32766\"/>\n                              </xs:restriction>\n                            </xs:simpleType>\n                          </xs:element>\n");
                }
                else
                {
                    sb.Append(" type=\"").Append(DsvType(c.DataType)).Append("\" minOccurs=\"0\"/>\n");
                }
            }
            sb.Append("                        </xs:sequence>\n                      </xs:complexType>\n                    </xs:element>\n");
        }
        sb.Append("                  </xs:choice>\n                </xs:complexType>\n              </xs:element>\n            </xs:schema>\n            <NewDataSet xmlns=\"\"/>\n          </Schema>");
        sb.Append("<DataSourceID>").Append(m.DataSourceId).Append("</DataSourceID></DataSourceView></ObjectDefinition></Load>");
        return sb.ToString();
    }

    private const string GeneralFormat = "<Format Format=\"General\" xmlns=\"\"/>";

    /// <summary>
    /// The Format annotation and the format string of a dimension attribute: the type's default when <paramref name="f"/> is null (data
    /// columns: General, Text, DateTimeGeneral, Currency of the model culture; calculated columns also Boolean), else the chosen
    /// format, written as in Excel's H37 sample.
    /// </summary>
    internal static (string annotation, string? formatString) AttributeFormat(MeasureFormat? f, ColumnDataType t, ModelInfo m, bool calculated)
    {
        if (f is null)
            return t switch
            {
                ColumnDataType.String => ("<Format Format=\"Text\" xmlns=\"\"/>", null),
                ColumnDataType.DateTime => ("<Format Format=\"DateTimeGeneral\" xmlns=\"\"/>", "General Date"),
                ColumnDataType.Currency => CurrencyFormat(null, m, attribute: true),
                ColumnDataType.Boolean when calculated => ("<Format Format=\"Boolean\" xmlns=\"\"/>", "\"TRUE\";\"TRUE\";\"FALSE\""),
                _ => (GeneralFormat, null),
            };
        string dec(int n) => n == 0 ? "" : "." + new string('0', n);
        string sep(bool on) => on ? " ThousandSeparator=\"True\"" : "";
        switch (f.KindValue)
        {
            case MeasureFormat.Kind.WholeNumber:
                return ("<Format Format=\"NumberWhole\" ThousandSeparator=\"True\" xmlns=\"\"/>", "#,0");
            case MeasureFormat.Kind.Decimal:
                return ("<Format Format=\"NumberDecimal\" Accuracy=\"" + f.Decimals + "\"" + sep(f.ThousandSeparator) + " xmlns=\"\"/>",
                    (f.ThousandSeparator ? "#,0" : "0") + dec(f.Decimals));
            case MeasureFormat.Kind.Percentage:
            {
                string p = "0" + dec(f.Decimals) + "%";
                return ("<Format Format=\"Percentage\" Accuracy=\"" + f.Decimals + "\" xmlns=\"\"/>", p + ";-" + p + ";" + p);
            }
            case MeasureFormat.Kind.Currency:
                return CurrencyFormat(f, m, attribute: true);
            case MeasureFormat.Kind.ShortDate:
            case MeasureFormat.Kind.LongDate:
            case MeasureFormat.Kind.Date:
            {
                var culture = f.Culture ?? m.Culture;
                int lcid = f.Culture?.LCID ?? m.Lcid;
                string source = f.Pattern ?? "d";
                string pattern = MeasureFormat.ExpandDatePattern(source, culture);
                // The annotation keeps the .NET pattern; the engine's format string is the same pattern in Excel's syntax (H38).
                string group = MeasureFormat.DateGroup(source, pattern);
                return ("<Format Format=\"DateTimeCustom\" xmlns=\"\">\n                      <DateTimes>\n                        <DateTime LCID=\"" + lcid
                    + "\" Group=\"" + group + "\" FormatString=\"" + E(pattern) + "\"/>\n                      </DateTimes>\n                    </Format>",
                    MeasureFormat.ToEngineDatePattern(pattern));
            }
            default:
                return (GeneralFormat, null);
        }
    }

    /// <summary>
    /// Power Pivot "Currency" format as Excel writes it: the symbol and layout of the given culture, of the model culture, or none
    /// (<see cref="MeasureFormat.CurrencyWithoutSymbol"/>: locale 1033, empty symbol, parentheses for negatives). In a column's format string
    /// the dollar sign is escaped (<c>\$</c>) and any other symbol quoted, as in H10 and H37.
    /// </summary>
    internal static (string annotation, string formatString) CurrencyFormat(MeasureFormat? f, ModelInfo m, bool attribute)
    {
        int decimals = f?.Decimals ?? 2;
        int lcid = f?.Culture?.LCID ?? m.Lcid;
        var culture = f?.Culture ?? m.Culture;
        string sym, name;
        int pos, neg;
        if (f?.NoSymbol == true)
        {
            lcid = 1033; sym = ""; name = ""; pos = 0; neg = 0;
        }
        else
        {
            var nf = culture.NumberFormat;
            sym = string.IsNullOrEmpty(nf.CurrencySymbol) ? "$" : nf.CurrencySymbol;
            name = sym + " " + (string.IsNullOrEmpty(culture.EnglishName) ? "English (United States)" : culture.EnglishName);
            pos = nf.CurrencyPositivePattern; neg = nf.CurrencyNegativePattern;
            // Windows (and so Excel) lays out negative US dollars in parentheses; the culture data of other platforms (ICU) says -$1.00.
            // Pinned so that the file does not depend on the operating system (H37). Other cultures come from the platform's data.
            if (lcid == 1033) neg = 0;
        }
        string annotation = "<Format Format=\"Currency\" Accuracy=\"" + decimals + "\" ThousandSeparator=\"True\" xmlns=\"\">\n                      <Currency LCID=\""
            + lcid + "\" DisplayName=\"" + E(name) + "\" Symbol=\"" + E(sym) + "\" PositivePattern=\"" + pos
            + "\" NegativePattern=\"" + neg + "\"/>\n                    </Format>";
        string s = attribute && sym == "$" ? "\\$" : "\"" + sym.Replace("\"", "\"\"") + "\"";
        string n = "#,0" + (decimals == 0 ? "" : "." + new string('0', decimals));
        string positive = pos switch { 1 => n + s, 2 => s + " " + n, 3 => n + " " + s, _ => s + n };
        string negative = neg switch
        {
            0 => "(" + s + n + ")", 1 => "-" + s + n, 2 => s + "-" + n, 3 => s + n + "-", 4 => "(" + n + s + ")",
            5 => "-" + n + s, 6 => n + "-" + s, 7 => n + s + "-", 8 => "-" + n + " " + s, 9 => "-" + s + " " + n,
            10 => n + " " + s + "-", 11 => s + " " + n + "-", 12 => s + " -" + n, 13 => n + "- " + s,
            14 => "(" + s + " " + n + ")", 15 => "(" + n + " " + s + ")", _ => "-" + s + n,
        };
        return (annotation, positive + ";" + negative + ";" + positive);
    }

    private static (string type, string size) KeyType(EncodedColumn c) => c.DataType switch
    {
        ColumnDataType.Int64 => ("BigInt", "-1"),
        ColumnDataType.Double => ("Double", "-1"),
        ColumnDataType.Currency => ("Currency", "-1"),
        ColumnDataType.DateTime => ("Date", "-1"),
        ColumnDataType.Boolean => ("Boolean", "-1"),
        _ => ("WChar", "32766"),
    };

    private static void ColumnBinding(StringBuilder sb, string tableId, string columnId)
        => sb.Append("<Source xsi:type=\"ColumnBinding\"><TableID>").Append(E(tableId)).Append("</TableID><ColumnID>")
             .Append(E(columnId)).Append("</ColumnID></Source>");

    public static string Dimension(ModelInfo m, EncodedTable t, IReadOnlyList<string> dataFiles, long estimatedSize, int version = 1)
    {
        var ids = m.TableIds[t.Id];
        var sb = new StringBuilder(LoadOpen);
        sb.Append("<ParentObject><DatabaseID>").Append(m.DatabaseId).Append("</DatabaseID></ParentObject>");
        sb.Append("<ObjectDefinition><Dimension>");
        CommonHeader(sb, t.Name, t.Id, m.Created, version, ids.DimensionObjectId, schemaUpdated: m.HiddenProfile && m.HiddenTables.Contains(t.Id) ? m.EditedAt : null);
        sb.Append("<DataFileList>").Append(E(string.Join(";", dataFiles))).Append("</DataFileList><Description/>");
        sb.Append("<Annotations").Append(NsB).Append('>').Append(Annotations(m.Resaved(t.Id))).Append("</Annotations>");
        sb.Append("<LastProcessed>").Append(T(m.Processed)).Append("</LastProcessed>");
        sb.Append("<EstimatedSize>").Append(estimatedSize).Append("</EstimatedSize>");
        sb.Append("<Source xsi:type=\"DataSourceViewBinding\"><DataSourceViewID>Sandbox</DataSourceViewID></Source><MiningModelID/><Type>").Append(m.DateTables.Contains(t.Id) ? "Time" : "Regular").Append("</Type>");
        sb.Append("<UnknownMember valuens=\"ddl200_200\">AutomaticNull</UnknownMember><MdxMissingMemberMode>Default</MdxMissingMemberMode>");
        sb.Append("<ddl300:StringStoresCompatibilityLevel>1050</ddl300:StringStoresCompatibilityLevel><ddl300:InternalStringStoresCompatibilityLevel>1050</ddl300:InternalStringStoresCompatibilityLevel><ddl300:CurrentStringStoresCompatibilityLevel>1050</ddl300:CurrentStringStoresCompatibilityLevel>");
        sb.Append("<ddl300:ProcessingState>Processed</ddl300:ProcessingState>");
        sb.Append(ErrorConfiguration);
        sb.Append("<StorageMode valuens=\"ddl200_200\">InMemory</StorageMode><CurrentStorageMode valuens=\"ddl200_200\">InMemory</CurrentStorageMode>");
        sb.Append("<ProcessingPriority>0</ProcessingPriority><WriteEnabled>false</WriteEnabled><DependsOnDimensionID/>");
        sb.Append("<Language>").Append(m.Lcid).Append("</Language><Collation/><UnknownMemberName>Unknown</UnknownMemberName>");
        sb.Append("<ProcessingMode>Regular</ProcessingMode><OrderByKey>false</OrderByKey><AttributeAllMemberName/>");
        sb.Append(ProactiveCaching);
        sb.Append("<UnknownMemberTranslations/><AttributeAllMemberTranslations/><Translations/><Attributes>");

        // RowNumber attribute (key of the table), related to every column.
        sb.Append("<Attribute><Name>").Append(FileNames.RowNumberId).Append("</Name><ID>").Append(FileNames.RowNumberId).Append("</ID><Description/>");
        // In a date table the date column takes over as the key of the dimension.
        sb.Append("<Type valuens=\"ddl200_200\">RowNumber</Type><Usage>").Append(m.DateTables.Contains(t.Id) ? "Regular" : "Key").Append("</Usage>");
        AttributeCommon(sb, visible: false);
        sb.Append("<KeyColumns><KeyColumn><DataType>Integer</DataType><DataSize>4</DataSize><NullProcessing>Error</NullProcessing><Collation/><Format/><Source xsi:type=\"ddl200_200:RowNumberBinding\"/></KeyColumn></KeyColumns>");
        sb.Append("<NameColumn><DataType>WChar</DataType><DataSize>4</DataSize><NullProcessing>ZeroOrBlank</NullProcessing><Collation/><Format/><Source xsi:type=\"ddl200_200:RowNumberBinding\"/></NameColumn>");
        sb.Append("<Translations/><NamingTemplateTranslations/><AttributeRelationships>");
        foreach (var c in t.Columns.Where(c => !c.IsDateKey))
        {
            sb.Append("<AttributeRelationship><AttributeID>").Append(E(c.Id)).Append("</AttributeID><Name>").Append(E(c.Name))
              .Append("</Name><RelationshipType>Flexible</RelationshipType><Cardinality>").Append(c.IsPrimaryKey ? "One" : "Many").Append("</Cardinality><Optionality>Mandatory</Optionality><OverrideBehavior>None</OverrideBehavior><Translations/></AttributeRelationship>");
        }
        foreach (var c in t.CalculatedColumns)
        {
            sb.Append("<AttributeRelationship><AttributeID>").Append(E(c.Id)).Append("</AttributeID><Name>").Append(E(c.Name))
              .Append("</Name><RelationshipType>Flexible</RelationshipType><Cardinality>Many</Cardinality><Optionality>Mandatory</Optionality><OverrideBehavior>None</OverrideBehavior><Translations/></AttributeRelationship>");
        }
        sb.Append("</AttributeRelationships></Attribute>");

        foreach (var c in t.Columns)
        {
            var (type, size) = KeyType(c);
            sb.Append("<Attribute><Name>").Append(E(c.Name)).Append("</Name><ID>").Append(E(c.Id)).Append("</ID><Description/>");
            var (annotation, formatString) = AttributeFormat(c.Format, c.DataType, m, calculated: false);
            sb.Append("<Annotations").Append(NsB).Append(">\n                <Annotation>\n                  <Name>Format</Name>\n                  <Value>\n                    ")
              .Append(annotation).Append("\n                  </Value>\n                </Annotation>\n              </Annotations>");
            sb.Append("<Type>Regular</Type><Usage>").Append(c.IsDateKey ? "Key" : "Regular").Append("</Usage>");
            AttributeCommon(sb, visible: !c.IsHidden, orderBy: c.OrderByColumnId);
            sb.Append("<KeyColumns><KeyColumn><DataType>").Append(type).Append("</DataType><DataSize>").Append(size)
              .Append("</DataSize><NullProcessing>").Append(c.IsPrimaryKey ? "Error" : "Preserve")
              .Append("</NullProcessing><InvalidXmlCharacters>Remove</InvalidXmlCharacters><Collation/><Format/>");
            ColumnBinding(sb, t.Id, c.Id);
            sb.Append("</KeyColumn></KeyColumns>");
            sb.Append("<NameColumn><DataType>WChar</DataType><DataSize>").Append(size)
              .Append("</DataSize><NullProcessing>ZeroOrBlank</NullProcessing><InvalidXmlCharacters>Remove</InvalidXmlCharacters><Collation/><Format/>");
            ColumnBinding(sb, t.Id, c.Id);
            sb.Append("</NameColumn><Translations/><NamingTemplateTranslations/>");
            if (c.IsDateKey)
                sb.Append("<AttributeRelationships><AttributeRelationship><AttributeID>").Append(FileNames.RowNumberId).Append("</AttributeID><Name>").Append(FileNames.RowNumberId)
                  .Append("</Name><RelationshipType>Flexible</RelationshipType><Cardinality>One</Cardinality><Optionality>Mandatory</Optionality><OverrideBehavior>None</OverrideBehavior><Translations/></AttributeRelationship></AttributeRelationships>");
            else sb.Append("<AttributeRelationships/>");
            if (formatString is not null) sb.Append("<ddl300_300:FormatString>").Append(E(formatString)).Append("</ddl300_300:FormatString>");
            sb.Append("</Attribute>");
        }
        foreach (var c in t.CalculatedColumns) CalculatedAttribute(sb, m, c);
        sb.Append("</Attributes>");
        if (t.Hierarchies.Count == 0) sb.Append("<Hierarchies/>");
        else
        {
            sb.Append("<Hierarchies>");
            foreach (var h in t.Hierarchies)
            {
                sb.Append("<Hierarchy><Name>").Append(E(h.Name)).Append("</Name><ID>").Append(E(h.Name)).Append("</ID><Description/>")
                  .Append("<ddl300:ProcessingState>").Append(h.Unprocessed ? "Unprocessed" : "Processed").Append("</ddl300:ProcessingState><ddl300:StructureType>")
                  .Append(h.Unprocessed ? "Unknown" : h.Natural ? "Natural" : "Unnatural").Append("</ddl300:StructureType><ddl300:RaggedHoleState>")
                  .Append(h.Unprocessed ? "Unknown" : "None").Append("</ddl300:RaggedHoleState>")
                  .Append("<DisplayFolder/><AllMemberName/><MemberNamesUnique>false</MemberNamesUnique><AllowDuplicateNames>true</AllowDuplicateNames><ddl2:MemberKeysUnique>NotUnique</ddl2:MemberKeysUnique>")
                  .Append("<AllMemberTranslations/><Translations/><Levels>");
                foreach (var lc in h.Levels)
                    sb.Append("<Level><Name>").Append(E(lc.Name)).Append("</Name><ID>").Append(E(lc.Id)).Append("</ID><Description/><SourceAttributeID>")
                      .Append(E(lc.Id)).Append("</SourceAttributeID><HideMemberIf>Never</HideMemberIf><Translations/></Level>");
                sb.Append("</Levels></Hierarchy>");
            }
            sb.Append("</Hierarchies>");
        }
        if (t.Relationships.Count > 0)
        {
            // The many-side dimension owns its relationships (the engine looks them up here by id).
            sb.Append("<ddl300_300:Relationships>");
            foreach (var r in t.Relationships)
            {
                sb.Append("<ddl300_300:Relationship><ID>").Append(r.Name).Append("</ID><Visible>true</Visible><ddl300:ProcessingState>Processed</ddl300:ProcessingState>");
                sb.Append("<ddl300_300:FromRelationshipEnd><Role/><ddl300_300:Multiplicity>Many</ddl300_300:Multiplicity><DimensionID>").Append(E(t.Id))
                  .Append("</DimensionID><Attributes><Attribute><AttributeID>").Append(E(r.ForeignColumnId)).Append("</AttributeID></Attribute></Attributes></ddl300_300:FromRelationshipEnd>");
                sb.Append("<ddl300_300:ToRelationshipEnd><Role/><ddl300_300:Multiplicity>One</ddl300_300:Multiplicity><DimensionID>").Append(E(r.PrimaryTableId))
                  .Append("</DimensionID><Attributes><Attribute><AttributeID>").Append(E(r.PrimaryColumnId)).Append("</AttributeID></Attribute></Attributes></ddl300_300:ToRelationshipEnd>");
                sb.Append("</ddl300_300:Relationship>");
            }
            sb.Append("</ddl300_300:Relationships>");
        }
        int hv = m.HierarchyTables.Contains(t.Id) ? 2 : 0;
        if (m.Calc is not null && m.Calc.Tables.TryGetValue(t.Id, out var calc) && calc.DimensionCounters is var (structure, dataId, cache, data))
            sb.Append("<StructureVersion>" + structure + "</StructureVersion><DataIDVersion>" + dataId + "</DataIDVersion><IncDataIDVersion>" + dataId + "</IncDataIDVersion><CacheVersion>" + cache + "</CacheVersion><DataVersion>" + data + "</DataVersion>");
        else if (m.DateTables.Contains(t.Id))
            sb.Append("<StructureVersion>" + (hv + 1) + "</StructureVersion><DataIDVersion>2</DataIDVersion><IncDataIDVersion>2</IncDataIDVersion><CacheVersion>2</CacheVersion><DataVersion>" + (hv + 1) + "</DataVersion>");
        else
            sb.Append("<StructureVersion>" + hv + "</StructureVersion><DataIDVersion>1</DataIDVersion><IncDataIDVersion>1</IncDataIDVersion><CacheVersion>1</CacheVersion><DataVersion>" + (hv + (m.DateReferencingTables.Contains(t.Id) ? 1 : 0)) + "</DataVersion>");
        sb.Append("<LastBindingChange>1601-01-01T00:00:00</LastBindingChange><PermissionFileList/></Dimension></ObjectDefinition></Load>");
        return sb.ToString();
    }

    private static void AttributeCommon(StringBuilder sb, bool visible, string? orderBy = null, bool processed = true)
    {
        string state = processed ? "Processed" : "Unprocessed";
        sb.Append("<EstimatedCount>0</EstimatedCount><DiscretizationMethod>None</DiscretizationMethod><DiscretizationBucketCount>0</DiscretizationBucketCount>");
        sb.Append(orderBy is null ? "<OrderBy>Key</OrderBy>" : "<OrderBy>AttributeKey</OrderBy>").Append("<InstanceSelection>None</InstanceSelection><DefaultMember/>")
          .Append(orderBy is null ? "<OrderByAttributeID/>" : "<OrderByAttributeID>" + E(orderBy) + "</OrderByAttributeID>").Append("<NamingTemplate/>");
        sb.Append("<MembersWithData>NonLeafDataVisible</MembersWithData><MembersWithDataCaption/><MemberNamesUnique>false</MemberNamesUnique>");
        sb.Append("<KeyUniquenessGuarantee>false</KeyUniquenessGuarantee><IsAggregatable>true</IsAggregatable><AttributeHierarchyEnabled>true</AttributeHierarchyEnabled>");
        sb.Append("<AttributeHierarchyVisible>").Append(visible ? "true" : "false").Append("</AttributeHierarchyVisible>");
        sb.Append("<AttributeHierarchyOrdered>true</AttributeHierarchyOrdered><AttributeHierarchyOptimizedState>FullyOptimized</AttributeHierarchyOptimizedState><AttributeHierarchyDisplayFolder/>");
        sb.Append("<ddl300:AttributeHierarchyProcessingState>").Append(state).Append("</ddl300:AttributeHierarchyProcessingState><GroupingBehavior>EncourageGrouping</GroupingBehavior><ddl200:ProcessingState>").Append(state).Append("</ddl200:ProcessingState>");
    }

    /// <summary>
    /// The dimension attribute of a calculated column as Excel saves it before the column was calculated (H10): the formula is an
    /// ExpressionBinding, both processing states are Unprocessed, and the type is inferred from the formula.
    /// </summary>
    private static void CalculatedAttribute(StringBuilder sb, ModelInfo m, EncodedCalculatedColumn c)
    {
        sb.Append("<Attribute><Name>").Append(E(c.Name)).Append("</Name><ID>").Append(E(c.Id)).Append("</ID><Description/>");
        var (annotation, formatString) = AttributeFormat(c.Format, c.DataType, m, calculated: true);
        sb.Append("<Annotations").Append(NsB).Append(">\n                <Annotation>\n                  <Name>Format</Name>\n                  <Value>\n                    ")
          .Append(annotation).Append("\n                  </Value>\n                </Annotation>\n              </Annotations>");
        sb.Append("<Type>Regular</Type><Usage>Regular</Usage>");
        AttributeCommon(sb, visible: !c.IsHidden, orderBy: c.OrderByColumnId, processed: false);
        sb.Append("<ddl200:InferredDatatype>").Append(InferredDatatype(c.DataType)).Append("</ddl200:InferredDatatype>");
        string expression = "<Source xsi:type=\"ddl200_200:ExpressionBinding\"><Expression>" + E(c.Expression) + "</Expression></Source>";
        sb.Append("<KeyColumns><KeyColumn><DataType>Empty</DataType><DataSize>0</DataSize><NullProcessing>Automatic</NullProcessing><Collation/><Format/>").Append(expression).Append("</KeyColumn></KeyColumns>");
        sb.Append("<NameColumn><DataType>WChar</DataType><DataSize>0</DataSize><NullProcessing>Automatic</NullProcessing><Collation/><Format/>").Append(expression).Append("</NameColumn>");
        sb.Append("<Translations/><NamingTemplateTranslations/><AttributeRelationships/>");
        if (formatString is not null) sb.Append("<ddl300_300:FormatString>").Append(E(formatString)).Append("</ddl300_300:FormatString>");
        sb.Append("</Attribute>");
    }

    /// <summary>The type Excel infers from the formula, as H10 to H15 show it.</summary>
    private static string InferredDatatype(ColumnDataType t) => t switch
    {
        ColumnDataType.Currency => "Currency",
        ColumnDataType.Int64 => "BigInt",
        ColumnDataType.Double => "Double",
        ColumnDataType.DateTime => "Date",
        ColumnDataType.Boolean => "Boolean",
        _ => "WChar",
    };

    /// <summary>Dimension information file (info.N.xml in the .dim folder, [MS-XLDM] 2.6.10.2).</summary>
    public static string DimensionInfo(EncodedTable t)
    {
        var sb = new StringBuilder();
        sb.Append("<Dimension xmlns=\"http://schemas.microsoft.com/analysisservices/2003/ext\" xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\">");
        sb.Append("<DataVersion>0</DataVersion><IndexVersion>0</IndexVersion><DecodeStoreVersion>-1</DecodeStoreVersion><LevelStoreVersion>-1</LevelStoreVersion><Properties>");
        foreach (var _ in t.AllColumns)
        {
            sb.Append("<Property><ParentChild>false</ParentChild><MapDataset><m_cbOffsetHeader>9223372036854775808</m_cbOffsetHeader><m_cbOffsetData>9223372036854775808</m_cbOffsetData><m_cRecord>0</m_cRecord><m_cSegment>0</m_cSegment><m_mskFormat>269484288</m_mskFormat><m_cbHeader>0</m_cbHeader><m_cPath>0</m_cPath><m_cData>0</m_cData><m_cSegmentIndex>4294967295</m_cSegmentIndex><MapDataIndices/><MinMaxValues/></MapDataset><Depth>0</Depth><Balanced>false</Balanced><HasHoles>false</HasHoles></Property>");
        }
        sb.Append("</Properties></Dimension>");
        return sb.ToString();
    }

    /// <summary>Cube information file, empty as Excel writes it.</summary>
    public const string CubeInfo = "<Cube xmlns=\"http://schemas.microsoft.com/analysisservices/2003/ext\" xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\"/>";

    public const string DefaultMemberName = "__XL_Count of Models";

    public static string Cube(ModelInfo m, IReadOnlyList<EncodedTable> tables, bool hasMeasures = false)
    {
        var sb = new StringBuilder(LoadOpen);
        sb.Append("<ParentObject><DatabaseID>").Append(m.DatabaseId).Append("</DatabaseID></ParentObject>");
        sb.Append("<ObjectDefinition><Cube>");
        CommonHeader(sb, "Model", "Model", m.Created, m.CubeVersion, m.CubeObjectId, schemaUpdated: m.HiddenProfile ? m.EditedAt : null);
        sb.Append("<DataFileList/><Description/>");
        sb.Append("<Annotations").Append(NsA).Append(">\n            <Annotation>\n              <Name>DefaultMeasure</Name>\n              <Value>").Append(hasMeasures ? DefaultMemberName : "__No measures defined").Append("</Value>\n            </Annotation>\n          </Annotations>");
        sb.Append("<LastProcessed>1699-12-31T00:00:00</LastProcessed><Language>").Append(m.Lcid).Append("</Language><Collation/><DefaultMeasure/><Visible>true</Visible><AggregationPrefix/>");
        sb.Append("<StorageMode valuens=\"ddl200_200\">InMemory</StorageMode><ProcessingMode>Regular</ProcessingMode><ProcessingPriority>0</ProcessingPriority>");
        sb.Append("<ScriptCacheProcessingMode>Regular</ScriptCacheProcessingMode><ScriptErrorHandlingMode>IgnoreNone</ScriptErrorHandlingMode><StorageLocation/><EstimatedRows>0</EstimatedRows>");
        sb.Append("<Source xsi:type=\"DataSourceViewBinding\"><DataSourceViewID>Sandbox</DataSourceViewID></Source>");
        sb.Append(ProactiveCaching);
        sb.Append("<Translations/><Dimensions>");
        foreach (var t in tables)
        {
            sb.Append("<Dimension><ID>").Append(E(t.Id)).Append("</ID><Name>").Append(E(t.Name)).Append("</Name><DimensionID>").Append(E(t.Id))
              .Append("</DimensionID><Visible>true</Visible><HierarchyUniqueNameStyle>IncludeDimensionName</HierarchyUniqueNameStyle><MemberUniqueNameStyle>Native</MemberUniqueNameStyle><AllMemberAggregationUsage>Default</AllMemberAggregationUsage><Translations/><Attributes>");
            foreach (var c in t.AllColumns)
            {
                sb.Append("<Attribute><AttributeID>").Append(E(c.Id)).Append("</AttributeID><AggregationUsage>Default</AggregationUsage><AttributeHierarchyEnabled>true</AttributeHierarchyEnabled><AttributeHierarchyVisible>")
                  .Append(c.IsRowNumber || c.IsHidden ? "false" : "true")
                  .Append("</AttributeHierarchyVisible><AttributeHierarchyOptimizedState>FullyOptimized</AttributeHierarchyOptimizedState></Attribute>");
            }
            foreach (var c in t.CalculatedColumns)
                sb.Append("<Attribute><AttributeID>").Append(E(c.Id)).Append("</AttributeID><AggregationUsage>Default</AggregationUsage><AttributeHierarchyEnabled>true</AttributeHierarchyEnabled><AttributeHierarchyVisible>")
                  .Append(c.IsHidden ? "false" : "true").Append("</AttributeHierarchyVisible><AttributeHierarchyOptimizedState>FullyOptimized</AttributeHierarchyOptimizedState></Attribute>");
            sb.Append("</Attributes>");
            if (t.Hierarchies.Count == 0) sb.Append("<Hierarchies/>");
            else
            {
                sb.Append("<Hierarchies>");
                foreach (var h in t.Hierarchies)
                    sb.Append("<Hierarchy><HierarchyID>").Append(E(h.Name)).Append("</HierarchyID><Enabled>true</Enabled><Visible>true</Visible><OptimizedState>FullyOptimized</OptimizedState></Hierarchy>");
                sb.Append("</Hierarchies>");
            }
            sb.Append("</Dimension>");
        }
        sb.Append("</Dimensions><Kpis/><Actions/><StructureVersion>" + (m.Calc?.CubeStructure ?? (m.AnyDate || m.HiddenProfile ? 12 : m.AnyHierarchy ? 11 : 0)) + "</StructureVersion><DataVersion>-1</DataVersion><CalcVersion>0</CalcVersion>");
        sb.Append("<LastBindingChange>1601-01-01T00:00:00</LastBindingChange><LastUpdate>").Append(T(m.HiddenProfile ? m.EditedAt.AddMilliseconds(200) : m.Created)).Append("</LastUpdate>");
        sb.Append("<MeasureGroupFileList>").Append(E(string.Join(";", (m.CalculatedProfile ? tables.OrderBy(t => t.Id, StringComparer.Ordinal).ToList() : tables).Select(t => t.Id + "." + m.MeasureGroupVersionOf(t.Id) + ".det.xml")))).Append("</MeasureGroupFileList>");
        sb.Append("<PerspectiveFileList/><AssemblyFileList>MdxScript." + m.ScriptVersion + ".scr.xml</AssemblyFileList><PermissionFileList/></Cube></ObjectDefinition></Load>");
        return sb.ToString();
    }

    public static string MdxScript(ModelInfo m, IReadOnlyList<MeasureDefinition> measures)
    {
        var sb = new StringBuilder(LoadOpen);
        sb.Append("<ParentObject><DatabaseID>").Append(m.DatabaseId).Append("</DatabaseID><CubeID>Model</CubeID></ParentObject>");
        sb.Append("<ObjectDefinition><MdxScript>");
        CommonHeader(sb, "MdxScript", "MdxScript", m.Created, m.ScriptVersion, m.MdxScriptObjectId, m.UniqueOrdinals ? 1 : 0);
        string member = (measures.Count > 0) ? DefaultMemberName : "__No measures defined";
        sb.Append("<DataFileList/><Description/><DefaultScript>true</DefaultScript><Commands><Command><Text>CALCULATE; \n");
        sb.Append("CREATE MEMBER CURRENTCUBE.Measures.[").Append(member).Append("] AS 1, VISIBLE = 0; \n");
        sb.Append("ALTER CUBE CURRENTCUBE UPDATE DIMENSION Measures, Default_Member = [").Append(member).Append("]; </Text></Command>");
        foreach (var me in measures)
        {
            sb.Append("<Command><Text>----------------------------------------------------------\n")
              .Append("-- PowerPivot measures command (do not modify manually) --\n")
              .Append("----------------------------------------------------------\n\n\n")
              .Append(E("CREATE MEASURE '" + me.Table.Replace("'", "''") + "'[" + me.Name + "]=" + me.Expression + ";\n"))
              .Append("</Text><Annotations").Append(NsA).Append(">\n                <Annotation>\n                  <Name>FullName</Name>\n                  <Value>")
              .Append(E(me.Name)).Append("</Value>\n                </Annotation>\n                <Annotation>\n                  <Name>Table</Name>\n                  <Value>")
              .Append(E(me.Table)).Append("</Value>\n                </Annotation>\n              </Annotations></Command>");
        }
        sb.Append("</Commands>");
        if (measures.Count == 0) sb.Append("<CalculationProperties/>");
        else
        {
            sb.Append("<CalculationProperties>");
            foreach (var me in measures)
            {
                var (formatString, formatXml) = MeasureFormatXml(me.Format, m);
                sb.Append("<CalculationProperty><CalculationReference>[").Append(E(me.Name)).Append("]</CalculationReference><CalculationType>Member</CalculationType><Description/><FormatString>")
                  .Append(E("'" + formatString.Replace("'", "''") + "'"))
                  .Append("</FormatString><ForeColor/><BackColor/><FontName/><FontSize/><FontFlags/><NonEmptyBehavior/><AssociatedMeasureGroupID/><DisplayFolder/><Language/><Annotations>")
                  .Append("\n                <Annotation>\n                  <Name>Type</Name>\n                  <Value>User</Value>\n                </Annotation>")
                  .Append("\n                <Annotation>\n                  <Name>IsPrivate</Name>\n                  <Value>").Append(me.Hidden ? "True" : "False").Append("</Value>\n                </Annotation>")
                  .Append("\n                <Annotation>\n                  <Name>Format</Name>\n                  <Value>\n                    ").Append(formatXml)
                  .Append("\n                  </Value>\n                </Annotation>\n              </Annotations><Translations/></CalculationProperty>");
            }
            sb.Append("</CalculationProperties>");
        }
        sb.Append("</MdxScript></ObjectDefinition></Load>");
        return sb.ToString();
    }

    // General and WholeNumber are verified against Excel-written files; the rest follow the same pattern.
    private static (string formatString, string xml) MeasureFormatXml(MeasureFormat f, ModelInfo m)
    {
        string dec(int n) => n == 0 ? "" : "." + new string('0', n);
        string acc(int n) => " Accuracy=\"" + n + "\"";
        switch (f.KindValue)
        {
            case MeasureFormat.Kind.WholeNumber:
                return ("#,0", "<Format Format=\"NumberWhole\" ThousandSeparator=\"True\"/>");
            case MeasureFormat.Kind.Decimal:
                return ((f.ThousandSeparator ? "#,0" : "0") + dec(f.Decimals),
                    "<Format Format=\"NumberDecimal\"" + acc(f.Decimals) + (f.ThousandSeparator ? " ThousandSeparator=\"True\"" : "") + "/>");
            case MeasureFormat.Kind.Percentage:
                return ("0" + dec(f.Decimals) + "%", "<Format Format=\"Percentage\"" + acc(f.Decimals) + "/>");
            case MeasureFormat.Kind.Currency:
            {
                var (annotation, fs) = CurrencyFormat(f, m, attribute: false);
                return (fs, annotation.Replace(" xmlns=\"\"", ""));
            }
            case MeasureFormat.Kind.ShortDate:
                return ("Short Date", "<Format Format=\"DateTimeShortDatePattern\"/>");
            default:
                return ("", "<Format Format=\"General\"/>");
        }
    }

    public static string MeasureGroup(ModelInfo m, EncodedTable t, IReadOnlyDictionary<string, EncodedTable> tablesById)
    {
        var ids = m.TableIds[t.Id];
        var sb = new StringBuilder(LoadOpen);
        sb.Append("<ParentObject><DatabaseID>").Append(m.DatabaseId).Append("</DatabaseID><CubeID>Model</CubeID></ParentObject>");
        sb.Append("<ObjectDefinition><MeasureGroup>");
        CommonHeader(sb, t.Name, t.Id, m.Created, m.MeasureGroupVersion, ids.MeasureGroupObjectId, m.UniqueOrdinals ? ids.Ordinal + (m.CalculatedProfile ? 2 : 0) : 0);
        sb.Append("<DataFileList/><Description/><LastProcessed>").Append(T(m.Processed)).Append("</LastProcessed>");
        sb.Append("<Type>Regular</Type><DataAggregation>DataAndCacheAggregatable</DataAggregation><AggregationPrefix/>");
        sb.Append("<StorageMode valuens=\"ddl200_200\">InMemory</StorageMode><ProcessingPriority>0</ProcessingPriority><StorageLocation/>");
        sb.Append("<IgnoreUnrelatedDimensions>true</IgnoreUnrelatedDimensions><EstimatedRows>0</EstimatedRows>");
        sb.Append(ErrorConfiguration);
        sb.Append("<ProcessingMode>Regular</ProcessingMode>").Append(ProactiveCaching);
        sb.Append("<Measures><Measure><Name>__XL_Count ").Append(E(t.Name)).Append("</Name><ID>").Append(E(t.Id)).Append("</ID><Description/>");
        sb.Append("<AggregateFunction>Count</AggregateFunction><DataType>BigInt</DataType><Visible>false</Visible><MeasureExpression/><DisplayFolder/><FormatString/>");
        sb.Append("<BackColor/><ForeColor/><FontName/><FontSize/><FontFlags/><Source><DataType>BigInt</DataType><DataSize>8</DataSize><NullProcessing>Automatic</NullProcessing><Collation/><Format/>");
        sb.Append("<Source xsi:type=\"RowBinding\"><TableID>").Append(E(t.Id)).Append("</TableID></Source></Source><Translations/></Measure></Measures>");
        sb.Append("<Dimensions><Dimension xsi:type=\"DegenerateMeasureGroupDimension\"><Attributes>");
        sb.Append("<Attribute><AttributeID>").Append(FileNames.RowNumberId).Append("</AttributeID><Type>Granularity</Type><Enabled>true</Enabled><KeyColumns><KeyColumn><DataType>Integer</DataType><DataSize>0</DataSize><NullProcessing>Automatic</NullProcessing><Collation/><Format/>");
        sb.Append("<Source xsi:type=\"ColumnBinding\"><TableID>").Append(E(t.Name)).Append("</TableID><ColumnID>").Append(FileNames.RowNumberId).Append("</ColumnID></Source></KeyColumn></KeyColumns></Attribute>");
        foreach (var c in t.Columns)
        {
            var (type, size) = KeyType(c);
            sb.Append("<Attribute><AttributeID>").Append(E(c.Id)).Append("</AttributeID><Type>Regular</Type><Enabled>true</Enabled><KeyColumns><KeyColumn><DataType>")
              .Append(type).Append("</DataType><DataSize>").Append(size)
              .Append("</DataSize><NullProcessing>Preserve</NullProcessing><InvalidXmlCharacters>Remove</InvalidXmlCharacters><Collation/><Format/>");
            ColumnBinding(sb, t.Id, c.Id);
            sb.Append("</KeyColumn></KeyColumns></Attribute>");
        }
        foreach (var c in t.CalculatedColumns)
            CalculatedGroupAttribute(sb, m, c);
        sb.Append("</Attributes><ddl200_200:ShareDimensionStorage>Shared</ddl200_200:ShareDimensionStorage><CubeDimensionID>").Append(E(t.Id))
          .Append("</CubeDimensionID></Dimension>");
        foreach (var r in t.Relationships) ReferenceDimension(sb, m, t, r, tablesById[r.PrimaryTableId]);
        sb.Append("</Dimensions><Translations/><StructureVersion>" + (m.Calc?.MeasureGroupStructure ?? (m.AnyDate ? 9 : m.AnyHierarchy ? 8 : 0)) + "</StructureVersion><CacheVersion>" + (m.Calc?.MeasureGroupCache ?? (m.AnyDate ? 2 : 1)) + "</CacheVersion><DataVersion>" + (m.Calc is not null ? m.Calc.MeasureGroupData[t.Id] : !m.AnyHierarchy ? 0 : m.HierarchyTables.Contains(t.Id) ? -1 : t.Relationships.Count > 0 ? 1 : 0) + "</DataVersion>");
        sb.Append("<LastBindingChange>1601-01-01T00:00:00</LastBindingChange><AggregationDesignFileList/><PartitionFileList>")
          .Append(E(t.Id)).Append(".").Append(m.PartitionVersion).Append(".prt.xml</PartitionFileList></MeasureGroup></ObjectDefinition></Load>");
        return sb.ToString();
    }

    private static void CalculatedGroupAttribute(StringBuilder sb, ModelInfo m, EncodedCalculatedColumn c)
        => sb.Append("<Attribute><AttributeID>").Append(E(c.Id)).Append("</AttributeID><Type>Regular</Type><Enabled>true</Enabled><KeyColumns><KeyColumn><DataType>Empty</DataType><DataSize>0</DataSize><NullProcessing>Automatic</NullProcessing><Collation/><Format/><Source xsi:type=\"ddl200_200:ExpressionBinding\">")
             .Append(m.Calc?.FormulasInMeasureGroups == true ? "<Expression>" + E(c.Expression) + "</Expression>" : "<Expression/>")
             .Append("</Source></KeyColumn></KeyColumns></Attribute>");

    /// <summary>The measure group dimension that lets a table be sliced by the columns of a related (primary) table.</summary>
    private static void ReferenceDimension(StringBuilder sb, ModelInfo m, EncodedTable fk, EncodedRelationship r, EncodedTable pk)
    {
        sb.Append("<Dimension xsi:type=\"ReferenceMeasureGroupDimension\"><Attributes>");
        sb.Append("<Attribute><AttributeID>").Append(FileNames.RowNumberId).Append("</AttributeID><Type>Regular</Type><Enabled>true</Enabled><KeyColumns><KeyColumn><DataType>Integer</DataType><DataSize>4</DataSize><NullProcessing>Error</NullProcessing><Collation/><Format/><Source xsi:type=\"ddl200_200:RowNumberBinding\"/></KeyColumn></KeyColumns></Attribute>");
        foreach (var c in pk.Columns)
        {
            var (type, size) = KeyType(c);
            bool key = c.Id == r.PrimaryColumnId;
            sb.Append("<Attribute><AttributeID>").Append(E(c.Id)).Append("</AttributeID><Type>").Append(key ? "Granularity" : "Regular")
              .Append("</Type><Enabled>true</Enabled><KeyColumns><KeyColumn><DataType>").Append(type).Append("</DataType><DataSize>").Append(size)
              .Append("</DataSize><NullProcessing>").Append(key ? "Error" : "Preserve")
              .Append("</NullProcessing><InvalidXmlCharacters>Remove</InvalidXmlCharacters><Collation/><Format/>");
            ColumnBinding(sb, pk.Id, c.Id);
            sb.Append("</KeyColumn></KeyColumns></Attribute>");
        }
        if (m.Calc?.FormulasInMeasureGroups == true)
            foreach (var c in pk.CalculatedColumns) CalculatedGroupAttribute(sb, m, c);
        sb.Append("</Attributes><IntermediateCubeDimensionID>").Append(E(fk.Id)).Append("</IntermediateCubeDimensionID><IntermediateGranularityAttributeID>")
          .Append(E(r.ForeignColumnId)).Append("</IntermediateGranularityAttributeID><Materialization>Regular</Materialization><ddl200:ProcessingState>Processed</ddl200:ProcessingState><ddl300:RelationshipID>")
          .Append(r.Name).Append("</ddl300:RelationshipID><CubeDimensionID>").Append(E(pk.Id)).Append("</CubeDimensionID></Dimension>");
    }

    /// <summary>
    /// The dimension of a linked worksheet table as Excel writes it (H42): no data source binding, attributes Key and RowNumber, a short
    /// column id annotation and a relationship from the row number to the column.
    /// </summary>
    public static string LinkedDimension(ModelInfo m, EncodedTable t, IReadOnlyList<string> dataFiles, long estimatedSize)
    {
        var ids = m.TableIds[t.Id];
        var sb = new StringBuilder(LoadOpen);
        sb.Append("<ParentObject><DatabaseID>").Append(m.DatabaseId).Append("</DatabaseID></ParentObject>");
        sb.Append("<ObjectDefinition><Dimension>");
        CommonHeader(sb, t.Name, t.Id, m.Created, ModelInfo.LinkedDimensionVersion, ids.DimensionObjectId, 1, m.Created.AddMilliseconds(320));
        sb.Append("<DataFileList>").Append(E(string.Join(";", dataFiles))).Append("</DataFileList><Description/>");
        sb.Append("<LastProcessed>").Append(T(m.Processed)).Append("</LastProcessed>");
        sb.Append("<EstimatedSize>").Append(estimatedSize).Append("</EstimatedSize><MiningModelID/><Type>Regular</Type>");
        sb.Append("<UnknownMember valuens=\"ddl200_200\">AutomaticNull</UnknownMember><MdxMissingMemberMode>Default</MdxMissingMemberMode>");
        sb.Append("<ddl300:StringStoresCompatibilityLevel>1050</ddl300:StringStoresCompatibilityLevel><ddl300:InternalStringStoresCompatibilityLevel>1050</ddl300:InternalStringStoresCompatibilityLevel><ddl300:CurrentStringStoresCompatibilityLevel>1050</ddl300:CurrentStringStoresCompatibilityLevel>");
        sb.Append("<ddl300:ProcessingState>Processed</ddl300:ProcessingState>").Append(ErrorConfiguration);
        sb.Append("<StorageMode valuens=\"ddl200_200\">InMemory</StorageMode><CurrentStorageMode valuens=\"ddl200_200\">InMemory</CurrentStorageMode>");
        sb.Append("<ProcessingPriority>0</ProcessingPriority><WriteEnabled>false</WriteEnabled><DependsOnDimensionID/>");
        sb.Append("<Language>").Append(m.Lcid).Append("</Language><Collation/><UnknownMemberName>Unknown</UnknownMemberName>");
        sb.Append("<ProcessingMode>Regular</ProcessingMode><OrderByKey>false</OrderByKey><AttributeAllMemberName/>");
        sb.Append(ProactiveCaching);
        sb.Append("<UnknownMemberTranslations/><AttributeAllMemberTranslations/><Translations/><Attributes>");
        var key = t.Columns[0];
        sb.Append("<Attribute><Name>").Append(E(key.Name)).Append("</Name><ID>").Append(E(key.Id)).Append("</ID><Description/>");
        sb.Append("<Annotations").Append(NsB).Append(">\n              <Annotation>\n                <Name>ShortColumnId</Name>\n                <Value>A</Value>\n              </Annotation>\n            </Annotations>");
        sb.Append("<Type>Regular</Type><Usage>Regular</Usage>");
        AttributeCommon(sb, visible: true);
        sb.Append("<KeyColumns><KeyColumn><DataType>BigInt</DataType><DataSize>0</DataSize><NullProcessing>Preserve</NullProcessing><Collation/><Format/></KeyColumn></KeyColumns>");
        sb.Append("<NameColumn><DataType>WChar</DataType><DataSize>0</DataSize><NullProcessing>ZeroOrBlank</NullProcessing><Collation/><Format/></NameColumn>");
        sb.Append("<Translations/><NamingTemplateTranslations/><AttributeRelationships/></Attribute>");
        sb.Append("<Attribute><Name>RowNumber</Name><ID>RowNumber</ID><Description/><Type valuens=\"ddl200_200\">RowNumber</Type><Usage>Key</Usage>");
        AttributeCommon(sb, visible: false);
        sb.Append("<KeyColumns><KeyColumn><DataType>Integer</DataType><DataSize>4</DataSize><NullProcessing>Error</NullProcessing><Collation/><Format/><Source xsi:type=\"ddl200_200:RowNumberBinding\"/></KeyColumn></KeyColumns>");
        sb.Append("<NameColumn><DataType>WChar</DataType><DataSize>4</DataSize><NullProcessing>ZeroOrBlank</NullProcessing><Collation/><Format/><Source xsi:type=\"ddl200_200:RowNumberBinding\"/></NameColumn>");
        sb.Append("<Translations/><NamingTemplateTranslations/><AttributeRelationships><AttributeRelationship><AttributeID>").Append(E(key.Id)).Append("</AttributeID><Name>")
          .Append(E(key.Id)).Append("_</Name><RelationshipType>Flexible</RelationshipType><Cardinality>Many</Cardinality><Optionality>Mandatory</Optionality><OverrideBehavior>None</OverrideBehavior><Translations/></AttributeRelationship></AttributeRelationships></Attribute>");
        sb.Append("</Attributes><Hierarchies/><StructureVersion>0</StructureVersion><DataIDVersion>1</DataIDVersion><IncDataIDVersion>1</IncDataIDVersion><CacheVersion>1</CacheVersion><DataVersion>0</DataVersion>");
        sb.Append("<LastBindingChange>1601-01-01T00:00:00</LastBindingChange><PermissionFileList/></Dimension></ObjectDefinition></Load>");
        return sb.ToString();
    }

    /// <summary>The measure group of a linked worksheet table as Excel writes it (H42): attributes Key and RowNumber without column bindings, and a row binding on the table id "TableID".</summary>
    public static string LinkedMeasureGroup(ModelInfo m, EncodedTable t)
    {
        var ids = m.TableIds[t.Id];
        var sb = new StringBuilder(LoadOpen);
        sb.Append("<ParentObject><DatabaseID>").Append(m.DatabaseId).Append("</DatabaseID><CubeID>Model</CubeID></ParentObject>");
        sb.Append("<ObjectDefinition><MeasureGroup>");
        CommonHeader(sb, t.Name, t.Id, m.Created, ModelInfo.LinkedVersion, ids.MeasureGroupObjectId, ids.Ordinal + (m.CalculatedProfile ? 2 : 0));
        sb.Append("<DataFileList/><Description/><LastProcessed>").Append(T(m.Processed)).Append("</LastProcessed>");
        sb.Append("<Type>Regular</Type><DataAggregation>DataAndCacheAggregatable</DataAggregation><AggregationPrefix/>");
        sb.Append("<StorageMode valuens=\"ddl200_200\">InMemory</StorageMode><ProcessingPriority>0</ProcessingPriority><StorageLocation/>");
        sb.Append("<IgnoreUnrelatedDimensions>true</IgnoreUnrelatedDimensions><EstimatedRows>0</EstimatedRows>");
        sb.Append(ErrorConfiguration);
        sb.Append("<ProcessingMode>Regular</ProcessingMode>").Append(ProactiveCaching);
        sb.Append("<Measures><Measure><Name>__XL_Count ").Append(E(t.Name)).Append("</Name><ID>").Append(E(t.Id)).Append("</ID><Description/>");
        sb.Append("<AggregateFunction>Count</AggregateFunction><DataType>BigInt</DataType><Visible>false</Visible><MeasureExpression/><DisplayFolder/><FormatString/>");
        sb.Append("<BackColor/><ForeColor/><FontName/><FontSize/><FontFlags/><Source><DataType>BigInt</DataType><DataSize>8</DataSize><NullProcessing>Automatic</NullProcessing><Collation/><Format/>");
        sb.Append("<Source xsi:type=\"RowBinding\"><TableID>TableID</TableID></Source></Source><Translations/></Measure></Measures>");
        sb.Append("<Dimensions><Dimension xsi:type=\"DegenerateMeasureGroupDimension\"><Attributes>");
        var key = t.Columns[0];
        sb.Append("<Attribute><AttributeID>").Append(E(key.Id)).Append("</AttributeID><Type>Regular</Type><Enabled>true</Enabled><KeyColumns><KeyColumn><DataType>BigInt</DataType><DataSize>0</DataSize><NullProcessing>Preserve</NullProcessing><Collation/><Format/></KeyColumn></KeyColumns></Attribute>");
        sb.Append("<Attribute><AttributeID>RowNumber</AttributeID><Type>Granularity</Type><Enabled>true</Enabled><KeyColumns><KeyColumn><DataType>Integer</DataType><DataSize>0</DataSize><NullProcessing>Automatic</NullProcessing><Collation/><Format/>");
        sb.Append("<Source xsi:type=\"ColumnBinding\"><TableID>").Append(E(t.Id)).Append("</TableID><ColumnID>RowNumber</ColumnID></Source></KeyColumn></KeyColumns></Attribute>");
        sb.Append("</Attributes><ddl200_200:ShareDimensionStorage>Shared</ddl200_200:ShareDimensionStorage><CubeDimensionID>").Append(E(t.Id)).Append("</CubeDimensionID></Dimension></Dimensions>");
        sb.Append("<Translations/><StructureVersion>0</StructureVersion><CacheVersion>1</CacheVersion><DataVersion>0</DataVersion>");
        sb.Append("<LastBindingChange>1601-01-01T00:00:00</LastBindingChange><AggregationDesignFileList/><PartitionFileList>")
          .Append(E(t.Id)).Append('.').Append(ModelInfo.LinkedVersion).Append(".prt.xml</PartitionFileList></MeasureGroup></ObjectDefinition></Load>");
        return sb.ToString();
    }

    /// <summary>The partition of a linked worksheet table as Excel writes it (H42): no annotations, error configuration or query binding.</summary>
    public static string LinkedPartition(ModelInfo m, string tableName, string tableId)
    {
        var ids = m.TableIds[tableId];
        var sb = new StringBuilder(LoadOpen);
        sb.Append("<ParentObject><DatabaseID>").Append(m.DatabaseId).Append("</DatabaseID><CubeID>Model</CubeID><MeasureGroupID>").Append(E(tableId)).Append("</MeasureGroupID></ParentObject>");
        sb.Append("<ObjectDefinition><Partition>");
        CommonHeader(sb, tableName, tableId, m.Created, ModelInfo.LinkedVersion, ids.PartitionObjectId, 0);
        sb.Append("<DataFileList/><Description/><DataVersion>0</DataVersion><LastBindingChange>1601-01-01T00:00:00</LastBindingChange><LastProcessed>").Append(T(m.Processed)).Append("</LastProcessed>");
        sb.Append("<Type>Data</Type><AggregationPrefix/><StorageMode valuens=\"ddl200_200\">InMemory</StorageMode><CurrentStorageMode valuens=\"ddl200_200\">InMemory</CurrentStorageMode>");
        sb.Append("<ddl300:StringStoresCompatibilityLevel>1050</ddl300:StringStoresCompatibilityLevel><ddl300:InternalStringStoresCompatibilityLevel>1050</ddl300:InternalStringStoresCompatibilityLevel><ddl300:CurrentStringStoresCompatibilityLevel>1050</ddl300:CurrentStringStoresCompatibilityLevel>");
        sb.Append("<ProcessingMode>Regular</ProcessingMode><ProcessingPriority>0</ProcessingPriority>");
        sb.Append("<StorageLocation/><RemoteDataSourceID/><Slice/><EstimatedRows>0</EstimatedRows><AggregationDesignID/>");
        sb.Append(ProactiveCaching).Append("<EstimatedSize>0</EstimatedSize><CacheVersion>1</CacheVersion></Partition></ObjectDefinition></Load>");
        return sb.ToString();
    }

    public static string Partition(ModelInfo m, string tableName, string tableId)
    {
        var ids = m.TableIds[tableId];
        var t = (Name: tableName, Id: tableId);
        var sb = new StringBuilder(LoadOpen);
        sb.Append("<ParentObject><DatabaseID>").Append(m.DatabaseId).Append("</DatabaseID><CubeID>Model</CubeID><MeasureGroupID>").Append(E(t.Id)).Append("</MeasureGroupID></ParentObject>");
        sb.Append("<ObjectDefinition><Partition>");
        CommonHeader(sb, t.Name, t.Id, m.Created, m.PartitionVersion, ids.PartitionObjectId, m.UniqueOrdinals ? 1 : 0);
        sb.Append("<DataFileList/><Description/>");
        sb.Append("<Annotations").Append(NsB).Append('>').Append(Annotations(m.CalculatedProfile).Replace("\n          ", "\n                  ").Replace("\n            ", "\n                    ")).Append("</Annotations>");
        sb.Append(m.Calc is not null
            ? "<DataVersion>" + m.Calc.PartitionData[tableId] + "</DataVersion>"
            : m.HierarchyTables.Contains(tableId) || m.DateReferencingTables.Contains(tableId) ? "<DataVersion>1</DataVersion>" : "<DataVersion>0</DataVersion>").Append("<LastBindingChange>1601-01-01T00:00:00</LastBindingChange><LastProcessed>").Append(T(m.Processed)).Append("</LastProcessed>");
        sb.Append("<Type>Data</Type><AggregationPrefix/><StorageMode valuens=\"ddl200_200\">InMemory</StorageMode><CurrentStorageMode valuens=\"ddl200_200\">InMemory</CurrentStorageMode>");
        sb.Append("<ddl300:StringStoresCompatibilityLevel>1050</ddl300:StringStoresCompatibilityLevel><ddl300:InternalStringStoresCompatibilityLevel>1050</ddl300:InternalStringStoresCompatibilityLevel><ddl300:CurrentStringStoresCompatibilityLevel>1050</ddl300:CurrentStringStoresCompatibilityLevel>");
        sb.Append("<ProcessingMode>Regular</ProcessingMode><ProcessingPriority>0</ProcessingPriority>").Append(ErrorConfiguration);
        sb.Append("<StorageLocation/><RemoteDataSourceID/><Slice/><EstimatedRows>0</EstimatedRows><AggregationDesignID/>");
        string q = "SELECT [" + t.Name.Replace("]", "]]") + "].*   FROM [" + t.Name.Replace("]", "]]") + "]";
        sb.Append("<Source xsi:type=\"QueryBinding\"><DataSourceID>").Append(m.DataSourceId).Append("</DataSourceID><QueryDefinition>").Append(E(q)).Append("</QueryDefinition></Source>");
        sb.Append(ProactiveCaching).Append("<EstimatedSize>0</EstimatedSize><CacheVersion>" + (m.Calc?.PartitionCache ?? 1) + "</CacheVersion></Partition></ObjectDefinition></Load>");
        return sb.ToString();
    }

    /// <summary>Partition information file (info.1.xml in the .prt folder, [MS-XLDM] 2.6.10.1).</summary>
    public const string PartitionInfo =
        "<Partition xmlns=\"http://schemas.microsoft.com/analysisservices/2003/ext\" xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\"><DataVersion>1</DataVersion><RigidAggVersion>-1</RigidAggVersion><FlexAggVersion>-1</FlexAggVersion><DataIndexVersion>-1</DataIndexVersion><RigidIndexVersion>-1</RigidIndexVersion><FlexIndexVersion>-1</FlexIndexVersion></Partition>";

    /// <summary>Content of the PARTITIONS file ([MS-XLDM] 2.1.2.1.1).</summary>
    public static string Partitions(ModelInfo m, string serverName, IReadOnlyList<(string name, string id, long dataSize)> tables)
    {
        var sb = new StringBuilder("<Partitions>");
        foreach (var (name, id, size) in tables)
        {
            sb.Append("<Partition><ObjectPath>").Append(E(serverName)).Append("\\LocalCube.").Append(m.DatabaseId).Append(".Model.")
              .Append(E(id)).Append('.').Append(E(id)).Append("</ObjectPath><Name>").Append(E(name)).Append("</Name><DataSize>")
              .Append(size).Append("</DataSize><Location/><DataSourceID/><DataSourceName/><ConnectionString/></Partition>");
        }
        sb.Append("</Partitions>");
        return sb.ToString();
    }
}
