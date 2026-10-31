using System.IO.Compression;
using System.Text;
using PivotWright.Internal.Metadata;
using PivotWright.Internal.Model;

namespace PivotWright.Internal.Package;

/// <summary>Writes the Open XML workbook that carries the data model part (xl/model/item.data).</summary>
internal static class XlsxWriter
{
    private static readonly UTF8Encoding Utf8 = new(false);
    private const string Decl = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n";

    public static void Write(Stream output, Stream modelData, ModelInfo info, IReadOnlyList<EncodedTable> tables,
                             DataModelOptions options)
    {
        using var zip = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true);
        var linked = tables.FirstOrDefault(t => t.IsLinked);
        Text(zip, "[Content_Types].xml", linked is null ? ContentTypes : ContentTypes.Replace("</Types>", LinkedContentTypes + "</Types>"));
        Text(zip, "_rels/.rels", RootRels);
        Text(zip, "docProps/app.xml", App(options.WorksheetName, linked is null ? null : LinkedSheetName));
        Text(zip, "docProps/core.xml", Core());
        Text(zip, "xl/workbook.xml", Workbook(info, tables, options.WorksheetName, linked));
        Text(zip, "xl/_rels/workbook.xml.rels", linked is null ? WorkbookRels : WorkbookRels.Replace("</Relationships>", LinkedWorkbookRels + "</Relationships>"));
        Text(zip, "xl/worksheets/sheet1.xml", Sheet);
        Text(zip, "xl/styles.xml", Styles);
        Text(zip, "xl/connections.xml", Connections(info, options.ExcludeSourceFromRefreshAll, linked));
        if (linked is not null) WriteLinkedParts(zip, linked);

        // Excel stores the data model part uncompressed; keep it that way unless asked otherwise.
        var entry = zip.CreateEntry("xl/model/item.data",
            options.CompressModelPart ? CompressionLevel.Optimal : CompressionLevel.NoCompression);
        using var es = entry.Open();
        modelData.Position = 0;
        modelData.CopyTo(es, 1 << 20);
    }

    private static void Text(ZipArchive zip, string name, string content)
    {
        var e = zip.CreateEntry(name, CompressionLevel.Optimal);
        using var s = e.Open();
        s.Write(Utf8.GetBytes(content));
    }

    private const string ContentTypes = Decl +
        "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">" +
        "<Default Extension=\"data\" ContentType=\"application/vnd.openxmlformats-officedocument.model+data\"/>" +
        "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>" +
        "<Default Extension=\"xml\" ContentType=\"application/xml\"/>" +
        "<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>" +
        "<Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>" +
        "<Override PartName=\"/xl/connections.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.connections+xml\"/>" +
        "<Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>" +
        "<Override PartName=\"/docProps/core.xml\" ContentType=\"application/vnd.openxmlformats-package.core-properties+xml\"/>" +
        "<Override PartName=\"/docProps/app.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.extended-properties+xml\"/>" +
        "</Types>";

    private const string RootRels = Decl +
        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
        "<Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties\" Target=\"docProps/app.xml\"/>" +
        "<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties\" Target=\"docProps/core.xml\"/>" +
        "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>" +
        "</Relationships>";

    private const string WorkbookRels = Decl +
        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
        "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/>" +
        "<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>" +
        "<Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/connections\" Target=\"connections.xml\"/>" +
        "<Relationship Id=\"rId4\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/powerPivotData\" Target=\"model/item.data\"/>" +
        "</Relationships>";

    private static string Workbook(ModelInfo info, IReadOnlyList<EncodedTable> tables, string sheetName, EncodedTable? linked)
    {
        var sb = new StringBuilder(Decl);
        sb.Append("<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" mc:Ignorable=\"x15\" xmlns:x15=\"http://schemas.microsoft.com/office/spreadsheetml/2010/11/main\">");
        sb.Append("<fileVersion appName=\"xl\" lastEdited=\"7\" lowestEdited=\"7\" rupBuild=\"30228\"/>");
        sb.Append("<workbookPr/>");
        sb.Append("<bookViews><workbookView xWindow=\"-110\" yWindow=\"-110\" windowWidth=\"19420\" windowHeight=\"11500\"/></bookViews>");
        sb.Append("<sheets><sheet name=\"").Append(Xml.Esc(sheetName)).Append("\" sheetId=\"1\" r:id=\"rId1\"/>");
        if (linked is not null) sb.Append("<sheet name=\"").Append(LinkedSheetName).Append("\" sheetId=\"2\" state=\"veryHidden\" r:id=\"rId5\"/>");
        sb.Append("</sheets>");
        if (linked is not null) sb.Append("<definedNames><definedName name=\"").Append(LinkedSourceName(linked)).Append("\" hidden=\"1\">").Append(Xml.Esc(linked.Name)).Append("[]</definedName></definedNames>");
        sb.Append("<calcPr calcId=\"191029\"/>");
        sb.Append("<extLst>");
        sb.Append("<ext uri=\"{140A7094-0E35-4892-8432-C4D2E57EDEB5}\" xmlns:x15=\"http://schemas.microsoft.com/office/spreadsheetml/2010/11/main\"><x15:workbookPr chartTrackingRefBase=\"1\"/></ext>");
        sb.Append("<ext uri=\"{FCE2AD5D-F65C-4FA6-A056-5C36A1767C68}\" xmlns:x15=\"http://schemas.microsoft.com/office/spreadsheetml/2010/11/main\"><x15:dataModel><x15:modelTables>");
        foreach (var t in tables)
        {
            sb.Append("<x15:modelTable id=\"").Append(Xml.Esc(t.Id)).Append("\" name=\"").Append(Xml.Esc(t.Name))
              .Append("\" connection=\"").Append(Xml.Esc(t.IsLinked ? LinkedConnectionName(t) : info.DataSourceName)).Append("\"/>");
        }
        sb.Append("</x15:modelTables>");
        if (tables.Any(t => t.Relationships.Count > 0))
        {
            // The diagram view reads the relationships from here (by table and column name), not from the model.
            var byId = tables.ToDictionary(t => t.Id);
            sb.Append("<x15:modelRelationships>");
            foreach (var t in tables)
                foreach (var r in t.Relationships)
                {
                    var one = byId[r.PrimaryTableId];
                    sb.Append("<x15:modelRelationship fromTable=\"").Append(Xml.Esc(t.Name))
                      .Append("\" fromColumn=\"").Append(Xml.Esc(t.Columns.First(c => c.Id == r.ForeignColumnId).Name))
                      .Append("\" toTable=\"").Append(Xml.Esc(one.Name))
                      .Append("\" toColumn=\"").Append(Xml.Esc(one.Columns.First(c => c.Id == r.PrimaryColumnId).Name)).Append("\"/>");
                }
            sb.Append("</x15:modelRelationships>");
        }
        sb.Append("</x15:dataModel></ext>");
        sb.Append("</extLst></workbook>");
        return sb.ToString();
    }

    private static string Connections(ModelInfo info, bool excludeFromRefreshAll, EncodedTable? linked)
    {
        var sb = new StringBuilder(Decl);
        sb.Append("<connections xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">");
        sb.Append("<connection id=\"1\" name=\"").Append(Xml.Esc(info.DataSourceName))
          .Append("\" type=\"100\" refreshedVersion=\"8\" minRefreshableVersion=\"5\"><extLst><ext uri=\"{DE250136-89BD-433C-8126-D09CA5730AF9}\" xmlns:x15=\"http://schemas.microsoft.com/office/spreadsheetml/2010/11/main\"><x15:connection id=\"")
          .Append(info.DataSourceId).Append(excludeFromRefreshAll ? "\" excludeFromRefreshAll=\"1" : "").Append("\"/></ext></extLst></connection>");
        sb.Append("<connection id=\"2\" keepAlive=\"1\" name=\"ThisWorkbookDataModel\" description=\"Data Model\" type=\"5\" refreshedVersion=\"8\" minRefreshableVersion=\"5\" background=\"1\">");
        sb.Append("<dbPr connection=\"Data Model Connection\" command=\"Model\" commandType=\"1\"/><olapPr sendLocale=\"1\" rowDrillCount=\"1000\"/>");
        sb.Append("<extLst><ext uri=\"{DE250136-89BD-433C-8126-D09CA5730AF9}\" xmlns:x15=\"http://schemas.microsoft.com/office/spreadsheetml/2010/11/main\"><x15:connection id=\"\" model=\"1\"/></ext></extLst></connection>");
        if (linked is not null)
            sb.Append("<connection id=\"3\" name=\"").Append(Xml.Esc(LinkedConnectionName(linked))).Append("\" type=\"102\" refreshOnLoad=\"1\" refreshedVersion=\"8\" minRefreshableVersion=\"5\"><extLst><ext uri=\"{DE250136-89BD-433C-8126-D09CA5730AF9}\" xmlns:x15=\"http://schemas.microsoft.com/office/spreadsheetml/2010/11/main\"><x15:connection id=\"")
              .Append(Xml.Esc(linked.Name)).Append("\"><x15:rangePr sourceName=\"").Append(LinkedSourceName(linked)).Append("\"/></x15:connection></ext></extLst></connection>");
        sb.Append("</connections>");
        return sb.ToString();
    }

    // ----- The linked worksheet table of AddCalculationTrigger (H42) -----

    private const string LinkedSheetName = "Sheet2";
    private const string LinkedFileName = "PivotWright.xlsx";
    private static string LinkedConnectionName(EncodedTable t) => "WorksheetConnection_" + LinkedFileName + "!" + t.Name;
    private static string LinkedSourceName(EncodedTable t) => "_xlcn.WorksheetConnection_" + LinkedFileName.Replace("!", "") + t.Name;

    private const int LinkedFirstCustomXml = 7;   // rId7 and up; rId5 is sheet2, rId6 the shared strings

    // The Power Pivot settings parts of H42 that do not depend on a table: manual calculation off, linked tables update automatically, and the
    // sandbox facts. (Power Pivot's own settings, in the customXml parts Excel keeps with the model.)
    private static readonly (string Name, string Value)[] GeminiParts =
    {
        ("ManualCalcMode", "False"), ("LinkedTableUpdateMode", "True"), ("SandboxNonEmpty", "1"), ("IsSandboxEmbedded", "yes"),
        ("PowerPivotVersion", "2015.130.1606.47"), ("RelationshipAutoDetectionEnabled", "True"),
    };

    private const string LinkedWorkbookRels =
        "<Relationship Id=\"rId5\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet2.xml\"/>" +
        "<Relationship Id=\"rId6\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings\" Target=\"sharedStrings.xml\"/>" +
        "<Relationship Id=\"rId7\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml\" Target=\"../customXml/item1.xml\"/>" +
        "<Relationship Id=\"rId8\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml\" Target=\"../customXml/item2.xml\"/>" +
        "<Relationship Id=\"rId9\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml\" Target=\"../customXml/item3.xml\"/>" +
        "<Relationship Id=\"rId10\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml\" Target=\"../customXml/item4.xml\"/>" +
        "<Relationship Id=\"rId11\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml\" Target=\"../customXml/item5.xml\"/>" +
        "<Relationship Id=\"rId12\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml\" Target=\"../customXml/item6.xml\"/>";

    private const string LinkedContentTypes =
        "<Override PartName=\"/xl/worksheets/sheet2.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>" +
        "<Override PartName=\"/xl/tables/table1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml\"/>" +
        "<Override PartName=\"/xl/sharedStrings.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml\"/>" +
        "<Override PartName=\"/customXml/itemProps1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.customXmlProperties+xml\"/>" +
        "<Override PartName=\"/customXml/itemProps2.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.customXmlProperties+xml\"/>" +
        "<Override PartName=\"/customXml/itemProps3.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.customXmlProperties+xml\"/>" +
        "<Override PartName=\"/customXml/itemProps4.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.customXmlProperties+xml\"/>" +
        "<Override PartName=\"/customXml/itemProps5.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.customXmlProperties+xml\"/>" +
        "<Override PartName=\"/customXml/itemProps6.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.customXmlProperties+xml\"/>";

    private static void WriteLinkedParts(ZipArchive zip, EncodedTable t)
    {
        string name = Xml.Esc(t.Name);
        const string main = "xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"";
        Text(zip, "xl/worksheets/sheet2.xml", Decl +
            "<worksheet " + main + " xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" +
            "<dimension ref=\"A1:A2\"/><sheetViews><sheetView workbookViewId=\"0\"/></sheetViews><sheetFormatPr defaultRowHeight=\"15\"/>" +
            "<sheetData><row r=\"1\"><c r=\"A1\" t=\"s\"><v>0</v></c></row><row r=\"2\"><c r=\"A2\"><v>2</v></c></row></sheetData>" +
            "<pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/><tableParts count=\"1\"><tablePart r:id=\"rId1\"/></tableParts></worksheet>");
        Text(zip, "xl/worksheets/_rels/sheet2.xml.rels", Decl +
            "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/table\" Target=\"../tables/table1.xml\"/></Relationships>");
        Text(zip, "xl/tables/table1.xml", Decl +
            "<table " + main + " id=\"1\" name=\"" + name + "\" displayName=\"" + name + "\" ref=\"A1:A2\" totalsRowShown=\"0\"><autoFilter ref=\"A1:A2\"/>" +
            "<tableColumns count=\"1\"><tableColumn id=\"1\" name=\"" + Xml.Esc(t.Columns[0].Name) + "\"/></tableColumns>" +
            "<tableStyleInfo name=\"TableStyleMedium2\" showFirstColumn=\"0\" showLastColumn=\"0\" showRowStripes=\"1\" showColumnStripes=\"0\"/></table>");
        Text(zip, "xl/sharedStrings.xml", Decl +
            "<sst " + main + " count=\"1\" uniqueCount=\"1\"><si><t>" + Xml.Esc(t.Columns[0].Name) + "</t></si></sst>");
        for (int i = 0; i < GeminiParts.Length; i++)
        {
            int n = i + 1;
            var (part, value) = GeminiParts[i];
            string xml = "<?xml version=\"1.0\" encoding=\"UTF-16\"?><Gemini xmlns=\"http://gemini/pivotcustomization/" + part + "\"><CustomContent><![CDATA[" + value + "]]></CustomContent></Gemini>";
            var e = zip.CreateEntry("customXml/item" + n + ".xml", CompressionLevel.Optimal);
            using (var s = e.Open()) { s.Write(new byte[] { 0xFF, 0xFE }); s.Write(Encoding.Unicode.GetBytes(xml)); }
            Text(zip, "customXml/itemProps" + n + ".xml", "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"no\"?>\r\n" +
                "<ds:datastoreItem ds:itemID=\"{" + Guid.NewGuid().ToString("D").ToUpperInvariant() + "}\" xmlns:ds=\"http://schemas.openxmlformats.org/officeDocument/2006/customXml\"><ds:schemaRefs/></ds:datastoreItem>");
            Text(zip, "customXml/_rels/item" + n + ".xml.rels", Decl +
                "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps\" Target=\"itemProps" + n + ".xml\"/></Relationships>");
        }
    }

    private const string Sheet = Decl +
        "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" +
        "<dimension ref=\"A1\"/><sheetViews><sheetView tabSelected=\"1\" workbookViewId=\"0\"/></sheetViews>" +
        "<sheetFormatPr defaultRowHeight=\"15\"/><sheetData/>" +
        "<pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/></worksheet>";

    private const string Styles = Decl +
        "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">" +
        "<fonts count=\"1\"><font><sz val=\"11\"/><name val=\"Calibri\"/><family val=\"2\"/></font></fonts>" +
        "<fills count=\"2\"><fill><patternFill patternType=\"none\"/></fill><fill><patternFill patternType=\"gray125\"/></fill></fills>" +
        "<borders count=\"1\"><border><left/><right/><top/><bottom/><diagonal/></border></borders>" +
        "<cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs>" +
        "<cellXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/></cellXfs>" +
        "<cellStyles count=\"1\"><cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/></cellStyles>" +
        "<dxfs count=\"0\"/><tableStyles count=\"0\" defaultTableStyle=\"TableStyleMedium2\" defaultPivotStyle=\"PivotStyleLight16\"/>" +
        "</styleSheet>";

    private static string App(string sheetName, string? second = null) => Decl +
        "<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\" xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\">" +
        "<Application>Microsoft Excel</Application><DocSecurity>0</DocSecurity><ScaleCrop>false</ScaleCrop>" +
        "<HeadingPairs><vt:vector size=\"2\" baseType=\"variant\"><vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant><vt:variant><vt:i4>" + (second is null ? 1 : 2) + "</vt:i4></vt:variant></vt:vector></HeadingPairs>" +
        "<TitlesOfParts><vt:vector size=\"" + (second is null ? 1 : 2) + "\" baseType=\"lpstr\"><vt:lpstr>" + Xml.Esc(sheetName) + "</vt:lpstr>" + (second is null ? "" : "<vt:lpstr>" + second + "</vt:lpstr>") + "</vt:vector></TitlesOfParts>" +
        "<LinksUpToDate>false</LinksUpToDate><SharedDoc>false</SharedDoc><HyperlinksChanged>false</HyperlinksChanged><AppVersion>16.0300</AppVersion></Properties>";

    private static string Core()
    {
        string now = DateTime.UtcNow.ToString("yyyy-MM-dd'T'HH:mm:ss'Z'", System.Globalization.CultureInfo.InvariantCulture);
        return Decl +
            "<cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" xmlns:dc=\"http://purl.org/dc/elements/1.1/\" xmlns:dcterms=\"http://purl.org/dc/terms/\" xmlns:dcmitype=\"http://purl.org/dc/dcmitype/\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\">" +
            "<dc:creator>PivotWright</dc:creator><cp:lastModifiedBy>PivotWright</cp:lastModifiedBy>" +
            "<dcterms:created xsi:type=\"dcterms:W3CDTF\">" + now + "</dcterms:created>" +
            "<dcterms:modified xsi:type=\"dcterms:W3CDTF\">" + now + "</dcterms:modified></cp:coreProperties>";
    }
}
