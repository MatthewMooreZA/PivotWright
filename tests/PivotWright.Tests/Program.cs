using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using PivotWright;
using PivotWright.Internal.Abf;
using PivotWright.Internal.Crypto;
using PivotWright.Internal.Model;
using PivotWright.Tests;

// A dependency-free test runner: `dotnet run --project tests/PivotWright.Tests`.
var tests = new (string Name, Action Body)[]
{
    ("CRC-32/BZIP2 check value", Tests.Crc),
    ("Xpress chunks decode back to the input", Tests.XpressRoundTrip),
    ("CryptKey decrypts an Excel-written connection string", Tests.DecryptExcelSecret),
    ("CryptKey round trip", Tests.CryptKeyRoundTrip),
    ("Container: header, directory, CRCs, sizes", Tests.Container),
    ("Value-encoded columns are byte-identical to Excel's", Tests.ExcelGoldenValueColumns),
    ("Text dictionary and hierarchy are byte-identical to Excel's", Tests.ExcelGoldenTextColumn),
    ("Golden files also match with Xpress compression on", Tests.ExcelGoldenCompressed),
    ("Multi-segment table metadata", Tests.MultiSegment),
    ("Names are sanitised the Power Pivot way", Tests.Names),
    ("Empty table and all-BLANK columns", Tests.EmptyAndBlank),
    ("Hierarchy segments of 1- and 2-value columns follow Power BI/Excel", Tests.TinyCardinalityHierarchy),
    ("Relationship index maps value-encoded foreign keys to primary rows", Tests.RelationshipValueKeys),
    ("Relationship index maps text keys case-insensitively and counts orphans", Tests.RelationshipTextKeys),
    ("Relationship metadata: R$ table, XMRelationship, reference dimension", Tests.RelationshipMetadata),
    ("Relationships with a huge key range use hash encoding", Tests.RelationshipWideKeys),
    ("Invalid relationships are rejected", Tests.RelationshipValidation),
    ("workbook.xml lists the relationships for the diagram view", Tests.WorkbookRelationships),
    ("Key columns are flagged the way Excel requires", Tests.KeyColumnFlags),
    ("Hierarchy tree matches the one Excel writes (decoded from an Excel-saved file)", Tests.HierarchyTree),
    ("Hierarchy metadata: dimension, cube and table entries", Tests.HierarchyMetadata),
    ("Hierarchy with BLANK values and case-insensitive duplicates", Tests.HierarchyBlanksAndCase),
    ("Invalid hierarchies are rejected", Tests.HierarchyValidation),
    ("Measures: MDX script commands and calculation properties", Tests.MeasureScript),
    ("Measures: no measures keeps the placeholder member", Tests.NoMeasures),
    ("Sort by Column: order, tree, metadata", Tests.SortByColumn),
    ("Sort by Column: invalid definitions are rejected", Tests.SortByColumnValidation),
    ("Relationship models: helper measure, table version 2, statistics", Tests.RelationshipDefaults),
    ("Measures: invalid definitions are rejected", Tests.MeasureValidation),
    ("Date table: dimension, tables, files and counters as Excel saves them", Tests.DateTableMetadata),
    ("Date table: invalid definitions are rejected", Tests.DateTableValidation),
    ("AddDateTable generates a standard calendar and marks it", Tests.GeneratedDateTable),
    ("AddDateTable: fiscal year end other than December", Tests.FiscalDateTable),
    ("Fiscal Calendar hierarchy lists quarters in fiscal order under every year", Tests.FiscalHierarchyOrder),
    ("Hierarchy StructureType: Natural only when every member has one parent", Tests.HierarchyStructureType),
    ("AddDateTable: yearLabel replaces Fiscal in the names", Tests.DateTableYearLabel),
    ("Options: ExcludeSourceFromRefreshAll writes the H41 attribute", Tests.ExcludeFromRefreshAll),
    ("CalculateOnOpen: a linked worksheet table makes Excel calculate the columns when the workbook opens (H42, e1)", Tests.CalculationTrigger),
    ("Calculated column: Sort by Column with another calculated column (H39)", Tests.CalculatedColumnSortBy),
    ("AddDateTable: DateFormat formats the Date column", Tests.DateTableDateFormat),
    ("AddDateTable: feature flags (complete years, YearMonth, hidden sort columns)", Tests.DateTableFlags),
    ("Hidden columns: plain table keeps the base profile", Tests.HiddenColumnsPlainTable),
    ("Hidden columns: keys and hierarchy levels, as in Excel's H1 sample", Tests.HiddenColumnsKeysAndLevels),
    ("Hidden columns: counters after a hide operation (H8, H9)", Tests.HiddenColumnsProfile),
    ("Measure groups get distinct Ordinals in every model", Tests.DistinctOrdinals),
    ("MarkAsDateTable accepts DateOnly, DateTime and DateTimeOffset sources", Tests.DateTableSourceTypes),
    ("API: tables are found by the name given and returned by AddTable", Tests.ApiTableHandles),
    ("API: a value that does not fit leaves the table consistent", Tests.ApiRowsAreAtomic),
    ("API: ModelColumn.IsHidden writes the same model as ColumnDefinition.IsHidden", Tests.ApiHideColumns),
    ("API: a failed Save leaves an existing file untouched", Tests.ApiSaveIsAtomic),
    ("Calculated column: formula, no values, counters as Excel saves them (H10)", Tests.CalculatedColumnProfile),
    ("Calculated column: the six data types as Excel infers them (H10 to H15)", Tests.CalculatedColumnTypes),
    ("Calculated columns: counters for two columns and for columns on the primary side (H18 to H20)", Tests.CalculatedColumnShapes),
    ("Calculated columns without a hierarchy: the four shapes of H21 to H24", Tests.CalculatedColumnShapesWithoutHierarchy),
    ("Calculated columns: one block of BLANKs per segment of a larger table", Tests.CalculatedColumnSeveralSegments),
    ("Formats: the eight columns of Excel's H37 sample, currency symbols and date patterns", Tests.ColumnFormats),
    ("Date formats are .NET date format strings: custom patterns as given, standard specifiers through the culture", Tests.DateFormatStrings),
    ("DAX text: paired brackets, parentheses and quotes are checked, real DAX is not rejected", Tests.DaxTextCheck),
    ("Calculated columns with hidden columns: hide a data column, a level, or the calculated column (H33 to H36)", Tests.CalculatedColumnsWithHiddenColumns),
    ("Calculated columns: BLANK run-length entry from 64 rows, bit-packed below (H21, H23)", Tests.CalculatedColumnBlankEncoding),
    ("Calculated columns with a date table: columns on Sales and on Dates (H29 to H32)", Tests.CalculatedColumnsWithDateTable),
    ("Calculated column: invalid definitions are rejected", Tests.CalculatedColumnValidation),
    ("KPI: script, annotations and status measures as in Excel's H47, H48 and H49", Tests.KpiScripts),
    ("KPI: two KPIs in one model (H50)", Tests.KpiTwo),
    ("KPI: lower is better, the LinearReversed scale (H51)", Tests.KpiReversed),
    ("KPI: a model without hierarchies (H52)", Tests.KpiPlain),
    ("KPI: a model with a date table (H53)", Tests.KpiDateTable),
    ("KPI: a model with hidden columns (H54)", Tests.KpiHidden),
    ("KPI: a one-table model without relationships (H57)", Tests.KpiOneTable),
    ("KPI: a calculated column on the primary-side table Products (H58)", Tests.KpiCalculatedOnPrimary),
    ("KPI: a hidden target measure (H59)", Tests.KpiHiddenTarget),
    ("KPI: a hidden KPI measure (H60)", Tests.KpiHiddenMeasure),
    ("KPI: the five icon styles, the gauge with five states (H61)", Tests.KpiIconStyles),
    ("KPI: the four band layouts with three-state and five-state icons (H62)", Tests.KpiLayouts),
    ("KPI: own gauge thresholds and Excel's annotation order (H63)", Tests.KpiCustomGaugeThresholds),
    ("KPI: four KPIs with a plain and a hidden measure beside them (extrapolated)", Tests.KpiMany),
    ("KPI: a model with calculated columns, with and without the trigger table (H55, H56)", Tests.KpiCalculated),
    ("KPI: counters and files as Excel saves them after Create KPI (H47)", Tests.KpiProfile),
    ("KPI: invalid definitions and unsupported models are rejected", Tests.KpiValidation),
    ("Perspective: one perspective, files and counters as Excel saves them (H64)", Tests.PerspectiveOne),
    ("Perspective: three perspectives, with and without a measure (H65)", Tests.PerspectiveThree),
    ("Perspective: hierarchies live in the cube annotation, a table shows all its items (H66)", Tests.PerspectiveHierarchiesAndTables),
    ("Perspective: models without perspectives keep their counters", Tests.PerspectiveAbsent),
    ("Perspective: invalid definitions and unsupported models are rejected", Tests.PerspectiveValidation),
    ("Inactive relationship: files and counters as Excel saves them (H73)", Tests.InactiveRelationship),
    ("Inactive relationship: workbook lists only active relationships, models without one keep their counters", Tests.InactiveRelationshipWorkbookAndAbsent),
    ("Inactive relationship: invalid and unsupported shapes are rejected", Tests.InactiveRelationshipValidation),
    ("Inactive relationship: two of them, counters as Excel saves them (H77)", Tests.InactiveRelationshipTwo),
    ("Inactive relationship: the date relationship added first (H78)", Tests.InactiveRelationshipDateFirst),
    ("Inactive relationship: a model without hierarchies (H79)", Tests.InactiveRelationshipPlain),
    ("Inactive relationship: next to a KPI (H83)", Tests.InactiveRelationshipKpi),
    ("Inactive relationship: one in each of two fact tables (H85)", Tests.InactiveRelationshipTwoFacts),
    ("KPI: Excel saves the first table of the model again, not the primary side of the first relationship (H86, H87)", Tests.KpiResavedTableIsTheFirstTable),
    ("Perspective: Excel saves the first table of the model again (H88, H89)", Tests.PerspectiveResavedTableIsTheFirstTable),
    ("KPI with a calculated column: the first table without one is saved again (H90, H91)", Tests.KpiCalcResavedTable),
    ("Inactive relationship: next to a perspective (H84)", Tests.InactiveRelationshipPerspective),
};

int failed = 0;
foreach (var (name, body) in tests)
{
    try
    {
        body();
        Console.WriteLine("PASS  " + name);
    }
    catch (Exception e)
    {
        failed++;
        Console.WriteLine("FAIL  " + name + "\n      " + e.Message);
    }
}
Console.WriteLine(failed == 0 ? $"All {tests.Length} tests passed." : $"{failed} of {tests.Length} tests failed.");
return failed == 0 ? 0 : 1;

namespace PivotWright.Tests
{
    internal static class Check
    {
        public static void True(bool condition, string what)
        {
            if (!condition) throw new Exception("Expected true: " + what);
        }

        public static void Equal<T>(T expected, T actual, string what)
        {
            if (!EqualityComparer<T>.Default.Equals(expected, actual))
                throw new Exception($"{what}: expected {expected}, got {actual}");
        }
    }

    internal static class Tests
    {
        private static string Sha(byte[] b) => Convert.ToHexString(SHA256.HashData(b)).ToLowerInvariant();

        public static void Crc() => Check.Equal(0xFC891918u, Crc32Bzip2.Compute(Encoding.ASCII.GetBytes("123456789")), "CRC");

        public static void XpressRoundTrip()
        {
            var rnd = new Random(3);
            var c = new XpressCompressor();
            var output = new byte[8300];
            int compressed = 0;
            for (int k = 0; k < 400; k++)
            {
                var input = new byte[rnd.Next(16, 4097)];
                for (int i = 0; i < input.Length; i++)
                    input[i] = (k % 3) switch { 0 => (byte)rnd.Next(4), 1 => (byte)(i % 11 == 0 ? rnd.Next(256) : 0), _ => (byte)rnd.Next(256) };
                int n = c.TryCompress(input, output);
                if (n == 0) continue;
                compressed++;
                Check.True(Xpress.Decompress(output.AsSpan(0, n), input.Length).AsSpan().SequenceEqual(input), "round trip " + k);
            }
            Check.True(compressed > 200, "most chunks compress");
        }

        // CryptKey.bin and the ConnectionString element of an Excel 365 workbook (null_data_id.xlsx).
        private const string ExcelCryptKey =
            "98bc215d2d8de64ea8e5d038aac94441040000003000000050000000100000000100000007000000ffffffff00000000" +
            "010200000366000000a40000914ab57fc8574004a2a48667da790ec16e431698ad1cfdae009665e02ca59e6c7f862ee1" +
            "efc38a31e217f3f6c75582fbcef7424dc6e8a0477df44565649f02000000000098bc215d2d8de64ea8e5d038aac94441";

        private const string ExcelConnectionString =
            "CAAAAAEAAAjzVUAc/riJ2SgYTdGZ5rHGYqGApiObaFo8knoug4BilSGuVZ+Z48c8T04P1qzWPvrR vVzhFbjX0/F1GivER0b/h1N8rLGZBgJN6Qei0wGFUmqn3L3QRWKSPJvcNvFASL45QEekVuRWaOOr " +
            "EOOweBVjFc+U3zRV8KNTqokGRKtMxibep6ZYJMZW7XoQeB2yjONEwo1OB6R4mof7bBao8QwJ74yW bvjUxYssSw3+5dhVvZqvIdbgt7c7NUsp8iyh3B5SBZmEeGenguqtmLwVhzHS";

        public static void DecryptExcelSecret()
        {
            var key = ModelCryptKey.FromFile(Convert.FromHexString(ExcelCryptKey));
            string plain = key.Decrypt(ExcelConnectionString).TrimEnd('\0');
            Check.Equal("Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=TheTable;Extended Properties=", plain, "connection string");
        }

        public static void CryptKeyRoundTrip()
        {
            var key = ModelCryptKey.Create();
            var file = key.ToFile();
            Check.Equal(144, file.Length, "CryptKey.bin size");
            var again = ModelCryptKey.FromFile(file);
            string secret = "Provider=SQLOLEDB;Data Source=server;Initial Catalog=Sales";
            Check.Equal(secret, again.Decrypt(key.Encrypt(secret, ModelCryptKey.ConnectionStringFlags)).TrimEnd('\0'), "decrypt");
        }

        private static ModelReader BuildReference(bool sortedByA, bool compressed)
        {
            var types = new[] { ColumnDataType.Int64, ColumnDataType.Int64, ColumnDataType.Currency, ColumnDataType.String, ColumnDataType.Int64 };
            var lines = ExcelReferenceData.Csv.Split('\n', StringSplitOptions.RemoveEmptyEntries).Select(l => l.TrimEnd('\r')).ToList();
            var header = lines[0].Split(',');
            var rows = lines.Skip(1).Select(l => l.Split(',')).ToList();
            if (sortedByA) rows = rows.OrderBy(r => long.Parse(r[0], CultureInfo.InvariantCulture)).ToList();
            using var model = new DataModelWriter(new DataModelOptions { Culture = CultureInfo.GetCultureInfo(2057), XpressCompression = compressed });
            var table = model.CreateTable("TheTable", header.Select((h, i) => new ColumnDefinition(h, types[i])));
            foreach (var r in rows)
            {
                var values = new object?[5];
                for (int i = 0; i < 5; i++)
                {
                    if (r[i].Length == 0) continue;
                    values[i] = types[i] switch
                    {
                        ColumnDataType.Int64 => (long)double.Parse(r[i], CultureInfo.InvariantCulture),
                        ColumnDataType.Currency => decimal.Parse(r[i], NumberStyles.Float, CultureInfo.InvariantCulture),
                        _ => r[i],
                    };
                }
                table.AddRow(values);
            }
            var ms = new MemoryStream();
            model.SaveModelPart(ms);
            return ModelReader.Read(ms.ToArray());
        }

        // SHA-256 of the files Excel wrote for the same data (null_data_id.xlsx, Excel 365, 2026).
        private static readonly (string Pattern, string Sha)[] ValueGolden =
        {
            (@"\.0\.dim\\0\.TheTable_G\.__XL_RowNumber\.0\.idf$", "fd43852412b51c5231cc30bbade1e23f799a04c8cb77cdc3701f925f6228bc63"),
            (@"\.0\.dim\\0\.TheTable_G\.A\.0\.idf$", "5400e218bc2fe1eac8202a874382c4218a3f6063517f42a884d242b442b57b40"),
            (@"\.0\.dim\\0\.TheTable_G\.C\.0\.idf$", "7cc3e99360d84d3a3f801cf2d8303854e88ca55441348372969e87f393d5f3a7"),
            (@"\.0\.dim\\0\.TheTable_G\.K\.0\.idf$", "e883a188fd1a68f70dc80393cb508c1c421991b1603748554fda7c5e8db20603"),
            (@"\.0\.dim\\0\.TheTable_G\.N\.0\.idf$", "e53f737dc523a20dc6b56ae1cf70d343a5e303c539d6b17819ac489c1bfbd299"),
            (@"\\0\.H\$TheTable_G\$A\.hidx$", "d590597bdf21f1c22d5b01ae7bc75ee62be77a48b94c6ea284b2b10e0282da04"),
            (@"\\0\.H\$TheTable_G\$C\.hidx$", "c27a98a2974c84d4b7fac22362f4b855c8789ba12b7dfe2c4b84947a8c50669a"),
            (@"\\0\.H\$TheTable_G\$K\.hidx$", "2f427f4857fd68a65e3af6ae4cb1c2e7d264db181a2a6d5af61b083a3cc2ad21"),
            (@"\\0\.H\$TheTable_G\$N\.hidx$", "74a59ed70264594a409dafdb672ebefdb3c16d9f908e1bf1297351100261b2cd"),
            (@"\\0\.H\$TheTable_G\$A\.POS_TO_ID\.0\.idf$", "2b3dbe7df047e1e0b65d09ad937cdd6191b1e0657c20207f44bcf500177fa572"),
            (@"\\0\.H\$TheTable_G\$C\.POS_TO_ID\.0\.idf$", "54973cf9075c3df3b788e6ddd5ceb6ce42fa13d0ec9d6cf8402ae520e1b00828"),
            (@"\\0\.H\$TheTable_G\$K\.POS_TO_ID\.0\.idf$", "ea52268142c83ace2a7c0af48f6fb849c7dd2ab083eea76d851cc723172df41a"),
            (@"\\0\.H\$TheTable_G\$N\.POS_TO_ID\.0\.idf$", "6b4656c934b38c488e8f9a2eb018d13d773161a7265cd46d76d753c6991ad3d1"),
        };

        // Excel's string dictionary order follows the source row order, so these are compared
        // after loading the rows in that order (sorted by A).
        private static readonly (string Pattern, string Sha)[] TextGolden =
        {
            (@"\\0\.TheTable_G\.S\.dictionary$", "7e6796c5d1b525bf8235c061e9721102ee5846a6c17cb47be5ee6abe387f0edc"),
            (@"\\0\.H\$TheTable_G\$S\.POS_TO_ID\.0\.idf$", "a8ffe622a89e2c80f313a1206ed2b80d54fa528ee21a29fb09a8d9e6530ebd7c"),
            (@"\\0\.H\$TheTable_G\$S\.ID_TO_POS\.0\.idf$", "3ab0cfd04311e941bbafc36c1bc073fb22833d4236712c23074f3b89b80a5aa0"),
        };

        private static void Compare(ModelReader m, (string Pattern, string Sha)[] golden)
        {
            foreach (var (pattern, sha) in golden)
                Check.Equal(sha, Sha(m.Find(pattern)), pattern);
        }

        public static void ExcelGoldenValueColumns() => Compare(BuildReference(sortedByA: false, compressed: false), ValueGolden);
        public static void ExcelGoldenTextColumn() => Compare(BuildReference(sortedByA: true, compressed: false), TextGolden);

        public static void ExcelGoldenCompressed()
        {
            Compare(BuildReference(sortedByA: false, compressed: true), ValueGolden);
            Compare(BuildReference(sortedByA: true, compressed: true), TextGolden);
        }

        public static void Container()
        {
            var m = BuildReference(sortedByA: false, compressed: true);
            Check.True(m.HeaderXml.Contains("<ApplyCompression>true</ApplyCompression>"), "compression flag");
            Check.Equal(8, m.Log.Element("FileGroups")!.Elements("FileGroup").Count(), "file groups");
            Check.True(m.Files.Keys.Any(k => k.EndsWith(".1.db.xml")), "database definition");
            Check.True(m.Files.Keys.Any(k => k.EndsWith("\\0.CryptKey.bin")), "CryptKey.bin");
            foreach (var (path, content) in m.Files.Where(kv => kv.Key.EndsWith(".xml")))
                System.Xml.Linq.XDocument.Parse(Encoding.UTF8.GetString(content)); // every metadata file is well-formed
        }

        public static void MultiSegment()
        {
            using var model = new DataModelWriter(new DataModelOptions { SegmentRowCount = 16384 });
            var t = model.CreateTable("Big", new[] { new ColumnDefinition("N", ColumnDataType.Int64), new ColumnDefinition("S", ColumnDataType.String) });
            for (int i = 0; i < 40000; i++) t.AddRow(i % 1000, "v" + (i % 77));
            var ms = new MemoryStream();
            model.SaveModelPart(ms);
            var m = ModelReader.Read(ms.ToArray());
            string tbl = Encoding.UTF8.GetString(m.Find(@"\.0\.dim\\Big_G\.0\.tbl\.xml$"));
            Check.True(tbl.Contains("XMSegmentEqualMapEx&lt;XMSegmentEqualMap_FastInstantiation>"), "equal segment map");
            Check.True(tbl.Contains("<Segments xsi:type=\"xsd:long\">3</Segments>"), "three segments");
            Check.True(tbl.Contains("<RecordsPerSegment xsi:type=\"xsd:long\">16384</RecordsPerSegment>"), "segment size");
            Check.True(tbl.Contains("<Min xsi:type=\"xsd:int\">32771</Min>"), "row number of third segment starts at 3 + 32768");
        }

        public static void Names()
        {
            Check.Equal("Sales USD", Identifiers.SanitizeName(" Sales (USD) ", "x"), "parentheses removed");
            Check.Equal("ab", Identifiers.SanitizeName("a.b", "x"), "dot removed");
            Check.Equal("Größe", Identifiers.SanitizeName("Größe", "x"), "unicode kept");
            Check.Equal("x", Identifiers.SanitizeName("()", "x"), "fallback");
        }

        public static void EmptyAndBlank()
        {
            using var model = new DataModelWriter();
            model.CreateTable("Empty", new[] { new ColumnDefinition("A", ColumnDataType.Int64) });
            var b = model.CreateTable("Blank", new[] { new ColumnDefinition("A", ColumnDataType.Double), new ColumnDefinition("S", ColumnDataType.String) });
            for (int i = 0; i < 10; i++) b.AddRow(null, null);
            var ms = new MemoryStream();
            model.SaveModelPart(ms);
            var m = ModelReader.Read(ms.ToArray());
            string tbl = Encoding.UTF8.GetString(m.Find(@"\.0\.dim\\Empty_G\.0\.tbl\.xml$"));
            Check.True(tbl.Contains("<Records xsi:type=\"xsd:long\">0</Records>"), "zero rows");
        }

        /// <summary>
        /// Columns with only one or two distinct values (and no BLANKs) must chunk their hierarchy records in
        /// units of the distinct count, as Power BI writes them; Excel refuses to open the model otherwise
        /// (verified in Excel). Expected words come from a real two-valued Boolean and a one-valued date column.
        /// </summary>
        public static void TinyCardinalityHierarchy()
        {
            using var model = new DataModelWriter();
            var t = model.CreateTable("T", new[]
            {
                new ColumnDefinition("Two", ColumnDataType.Int64),
                new ColumnDefinition("One", ColumnDataType.Int64),
                new ColumnDefinition("Three", ColumnDataType.Int64),
            });
            for (int i = 0; i < 100; i++) t.AddRow(1 + i % 2, 7, 1 + i % 3);
            var ms = new MemoryStream();
            model.SaveModelPart(ms);
            var m = ModelReader.Read(ms.ToArray());
            uint[] Words(string col)
            {
                var b = m.Find(@"\.0\.dim\\0\.H\$T_G\$" + col + @"\.POS_TO_ID\.0\.idf$");
                var w = new uint[b.Length / 4];
                for (int i = 0; i < w.Length; i++) w[i] = BitConverter.ToUInt32(b, i * 4);
                return w;
            }
            Check.Equal("2,0,3,4,0,0,2,0,0,0,0,0,1,0,0,0", string.Join(",", Words("Two")), "2 distinct: segments 2,2,1");
            Check.Equal("1,0,3,0,1,0,0,0,1,0,0,0,1,0,0,0", string.Join(",", Words("One")), "1 distinct: segments 1,1,1,1");
            Check.Equal("2,0,3,4,5,0,2,0,0,0,0,0", string.Join(",", Words("Three")), "3 distinct: segments 3,3");
        }

        private static (DataModelWriter model, ModelTable many, ModelTable one) Pair(ColumnDataType type)
        {
            var model = new DataModelWriter();
            var one = model.CreateTable("One", new[] { new ColumnDefinition("K", type), new ColumnDefinition("V", ColumnDataType.String) });
            var many = model.CreateTable("Many", new[] { new ColumnDefinition("K", type), new ColumnDefinition("N", ColumnDataType.Int64) });
            return (model, many, one);
        }

        /// <summary>Decodes the single INDEX column of the model: (values with -1 for no match, width, records, tbl.xml).</summary>
        private static (int[] values, int width, int records, string xml) DecodeIndex(ModelReader m)
        {
            string xml = Encoding.UTF8.GetString(m.Find(@"\.0\.dim\\R\$Many_G\$G\.0\.tbl\.xml$"));
            var b = m.Find(@"\.0\.dim\\0\.R\$Many_G\$G\.INDEX\.0\.idf$");
            int width = int.Parse(System.Text.RegularExpressions.Regex.Match(xml, @"XMRENoSplitCompressionInfo&lt;(\d+)>").Groups[1].Value);
            int records = int.Parse(System.Text.RegularExpressions.Regex.Match(xml, @"<Records xsi:type=""xsd:long"">(\d+)</Records>").Groups[1].Value);
            long words = BitConverter.ToInt64(b, 0);
            Check.Equal(words, (long)(records / (64 / width) + 1), "word count");
            Check.Equal(8 + words * 8, (long)b.Length, "file length");
            var values = new int[records];
            for (int i = 0; i < records; i++)
            {
                int per = 64 / width;
                ulong w = BitConverter.ToUInt64(b, 8 + 8 * (i / per));
                values[i] = (int)((w >> (width * (i % per))) & ((1UL << width) - 1)) - 1;
            }
            return (values, width, records, xml);
        }

        public static void RelationshipValueKeys()
        {
            var (model, many, one) = Pair(ColumnDataType.Int64);
            using (model)
            {
                foreach (var k in new[] { 10, 20, 30, 40 }) one.AddRow(k, "v" + k);
                foreach (var k in new[] { 30, 10, 40, 10, 30 }) many.AddRow(k, 1);
                model.AddRelationship(model.Tables["Many"]["K"], model.Tables["One"]["K"]);
                var ms = new MemoryStream();
                model.SaveModelPart(ms);
                var (values, width, records, xml) = DecodeIndex(ModelReader.Read(ms.ToArray()));
                // Foreign key ids: 3 = 10 ... 33 = 40; ids 0-2 and the unused ids in between do not matter, the used ones must.
                Check.Equal(34, records, "records = highest data id + 1");
                Check.Equal("-1,-1,-1,0,", string.Join(",", values.Take(4)) + ",", "ids 0..3");
                Check.Equal(0, values[3], "10 is the first primary row");
                Check.Equal(2, values[23], "30 is the third primary row");
                Check.Equal(3, values[33], "40 is the fourth primary row");
                Check.Equal(3, width, "3 bits for rows -1..3");
                Check.True(xml.Contains("<RIViolationCount xsi:type=\"xsd:long\">0</RIViolationCount>"), "no violations");
            }
        }

        public static void RelationshipTextKeys()
        {
            var (model, many, one) = Pair(ColumnDataType.String);
            using (model)
            {
                one.AddRow("b", "Bob"); one.AddRow("A", "Ann"); one.AddRow("c", "Cy");
                foreach (var k in new[] { "a", "B", "C", "z", "zz", "a" }) many.AddRow(k, 1);
                many.AddRow(null, 2);
                model.AddRelationship(model.Tables["Many"]["K"], model.Tables["One"]["K"]);
                var ms = new MemoryStream();
                model.SaveModelPart(ms);
                var (values, _, records, xml) = DecodeIndex(ModelReader.Read(ms.ToArray()));
                // ids by first appearance: a=3, B=4, C=5, z=6, zz=7
                Check.Equal(8, records, "records");
                Check.Equal("-1,-1,-1,1,0,2,-1,-1", string.Join(",", values), "index");
                Check.True(xml.Contains("<RIViolationCount xsi:type=\"xsd:long\">2</RIViolationCount>"), "z and zz are orphans");
            }
        }

        public static void RelationshipMetadata()
        {
            var (model, many, one) = Pair(ColumnDataType.Int64);
            using (model)
            {
                one.AddRow(1, "x"); many.AddRow(1, 1);
                model.AddRelationship(many["K"], one["K"]);
                var ms = new MemoryStream();
                model.SaveModelPart(ms);
                var m = ModelReader.Read(ms.ToArray());
                string fk = Encoding.UTF8.GetString(m.Find(@"\.0\.dim\\Many_G\.0\.tbl\.xml$"));
                Check.True(fk.Contains("class=\"XMRelationship\""), "XMRelationship in the many table");
                Check.True(fk.Contains("<PrimaryTable>One_"), "primary table");
                Check.True(fk.Contains("<PrimaryColumn>K</PrimaryColumn><ForeignColumn>K</ForeignColumn>"), "columns");
                Check.True(fk.Contains("XMRelationshipIndexDenseDIDs"), "index object");
                string primary = Encoding.UTF8.GetString(m.Find(@"\.0\.dim\\One_G\.0\.tbl\.xml$"));
                Check.True(!primary.Contains("XMRelationship\""), "the one side has no relationship entry");
                string det = Encoding.UTF8.GetString(m.Find(@"\.0\.cub\\Many_G\.1\.det\.xml$"));
                Check.True(det.Contains("ReferenceMeasureGroupDimension"), "reference dimension in the many table's measure group");
                Check.True(det.Contains("<ddl300:RelationshipID>"), "relationship id");
                string dim = Encoding.UTF8.GetString(m.Find(@"\.0\.db\\Many_G\.2\.dim\.xml$"));
                Check.True(dim.Contains("INDEX.0.idf"), "index file listed in the dimension's data files");
                Check.True(dim.Contains("<ddl300_300:Relationships><ddl300_300:Relationship><ID>"), "relationship declared on the many dimension");
                Check.True(dim.Contains("<ddl300_300:Multiplicity>Many</ddl300_300:Multiplicity>") && dim.Contains("<ddl300_300:Multiplicity>One</ddl300_300:Multiplicity>"), "multiplicities");
            }
        }

        public static void RelationshipWideKeys()
        {
            var (model, many, one) = Pair(ColumnDataType.Int64);
            using (model)
            {
                one.AddRow(1L, "a"); one.AddRow(100_000_000L, "b");
                many.AddRow(100_000_000L, 1); many.AddRow(1L, 2);
                model.AddRelationship(model.Tables["Many"]["K"], model.Tables["One"]["K"]);
                var ms = new MemoryStream();
                model.SaveModelPart(ms);
                var m = ModelReader.Read(ms.ToArray());
                var (values, _, records, _) = DecodeIndex(m);
                Check.Equal(5, records, "hash encoded: two values + 3 reserved ids");
                Check.Equal("-1,-1,-1,1,0", string.Join(",", values), "index");
                m.Find(@"\.0\.dim\\0\.Many_G\.K\.dictionary$"); // hash encoded columns have a dictionary file
            }
        }

        public static void RelationshipValidation()
        {
            void Throws(Action a, string what)
            {
                try { a(); }
                catch (Exception e) when (e is ArgumentException or InvalidOperationException) { return; }
                throw new Exception("expected an error for " + what);
            }
            void Save(DataModelWriter w) => w.SaveModelPart(new MemoryStream());

            var (m1, many1, one1) = Pair(ColumnDataType.Int64);
            using (m1)
            {
                Tests.Throws<KeyNotFoundException>(() => m1.AddRelationship(m1.Tables["Nope"]["K"], m1.Tables["One"]["K"]), "unknown table");
                Tests.Throws<KeyNotFoundException>(() => m1.AddRelationship(m1.Tables["Many"]["Nope"], m1.Tables["One"]["K"]), "unknown column");
                Throws(() => m1.AddRelationship(m1.Tables["Many"]["K"], m1.Tables["Many"]["K"]), "same table");
                Throws(() => m1.AddRelationship(m1.Tables["Many"]["K"], m1.Tables["One"]["V"]), "type mismatch");
                m1.AddRelationship(m1.Tables["Many"]["K"], m1.Tables["One"]["K"]);
                Throws(() => m1.AddRelationship(m1.Tables["One"]["K"], m1.Tables["Many"]["K"]), "already related");
            }

            var (m2, many2, one2) = Pair(ColumnDataType.Int64);
            using (m2)
            {
                one2.AddRow(1, "a"); one2.AddRow(1, "b"); many2.AddRow(1, 1);
                m2.AddRelationship(m2.Tables["Many"]["K"], m2.Tables["One"]["K"]);
                Throws(() => Save(m2), "duplicate primary key");
            }

            var (m3, many3, one3) = Pair(ColumnDataType.String);
            using (m3)
            {
                one3.AddRow("a", "x"); one3.AddRow("A", "y"); many3.AddRow("a", 1);
                m3.AddRelationship(m3.Tables["Many"]["K"], m3.Tables["One"]["K"]);
                Throws(() => Save(m3), "keys that differ only by case are duplicates");
            }

            var (m4, many4, one4) = Pair(ColumnDataType.Int64);
            using (m4)
            {
                one4.AddRow(null, "x"); many4.AddRow(1, 1);
                m4.AddRelationship(m4.Tables["Many"]["K"], m4.Tables["One"]["K"]);
                Throws(() => Save(m4), "blank primary key");
            }

            using var m5 = new DataModelWriter();
            m5.CreateTable("A", new[] { new ColumnDefinition("K", ColumnDataType.Int64) });
            m5.CreateTable("B", new[] { new ColumnDefinition("K", ColumnDataType.Int64) });
            m5.CreateTable("C", new[] { new ColumnDefinition("K", ColumnDataType.Int64) });
            m5.AddRelationship(m5.Tables["A"]["K"], m5.Tables["B"]["K"]);
            m5.AddRelationship(m5.Tables["B"]["K"], m5.Tables["C"]["K"]);
            Throws(() => m5.AddRelationship(m5.Tables["A"]["K"], m5.Tables["C"]["K"]), "A-B-C-A is circular");
        }

        private static DataModelWriter ProductsModel(out ModelTable products)
        {
            var model = new DataModelWriter();
            products = model.CreateTable("Products", new[] { new ColumnDefinition("ProductKey", ColumnDataType.Int64), new ColumnDefinition("Category", ColumnDataType.String), new ColumnDefinition("Subcategory", ColumnDataType.String), new ColumnDefinition("Product", ColumnDataType.String) });
            string[] cats = { "Bikes", "Clothing", "Accessories" };
            string[][] subs = { new[] { "Road", "Mountain" }, new[] { "Jerseys", "Gloves" }, new[] { "Helmets", "Lights" } };
            int k = 1;
            for (int c = 0; c < 3; c++) for (int s = 0; s < 2; s++) for (int i = 1; i <= 3; i++) products.AddRow(k++, cats[c], subs[c][s], $"{subs[c][s]} {i}");
            return model;
        }

        /// <summary>The four columns of a hierarchy tree: (values, minimum stored in the file).</summary>
        private static int[] DecodeTree(ModelReader m, string table, string hierarchy, string column, int min)
        {
            string xml = Encoding.UTF8.GetString(m.Find(@"\.0\.dim\\U\$" + table + "_G\\$" + hierarchy + @"\.2\.tbl\.xml$"));
            int records = int.Parse(System.Text.RegularExpressions.Regex.Match(xml, @"<Records xsi:type=""xsd:long"">(\d+)</Records>").Groups[1].Value);
            var b = m.Find(@"\.0\.dim\\2\.U\$" + table + "_G\\$" + hierarchy + "\\." + column + @"\.0\.idf$");
            long words = BitConverter.ToInt64(b, 0);
            Check.Equal(words, (long)(records / 2 + 1), "word count of " + column);
            Check.Equal(8 + words * 8, (long)b.Length, "file length of " + column);
            var values = new int[records];
            for (int i = 0; i < records; i++)
                values[i] = (int)(BitConverter.ToUInt32(b, 8 + 4 * i)) + min;
            return values;
        }

        /// <summary>Expected values were read from the hierarchy tables of a workbook that Excel itself saved.</summary>
        public static void HierarchyTree()
        {
            using var model = ProductsModel(out _);
            model.Tables["Products"].AddHierarchy("Product Hierarchy", "Category", "Subcategory", "Product");
            var ms = new MemoryStream();
            model.SaveModelPart(ms);
            var m = ModelReader.Read(ms.ToArray());
            string J(int[] a) => string.Join(",", a);
            Check.Equal("5,3,4,7,8,4,3,6,5,15,16,17,18,19,20,6,7,8,3,4,5,12,13,14,9,10,11", J(DecodeTree(m, "Products", "Product Hierarchy", "MULTI_LEVEL_ID", 0)), "MULTI_LEVEL_ID");
            Check.Equal("-1,-1,-1,0,0,1,1,2,2,3,3,3,4,4,4,5,5,5,6,6,6,7,7,7,8,8,8", J(DecodeTree(m, "Products", "Product Hierarchy", "PARENT_POS", -1)), "PARENT_POS");
            Check.Equal("3,5,7,9,12,15,18,21,24,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1", J(DecodeTree(m, "Products", "Product Hierarchy", "FIRST_CHILD_POS", -1)), "FIRST_CHILD_POS");
            Check.Equal("2,2,2,3,3,3,3,3,3,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1", J(DecodeTree(m, "Products", "Product Hierarchy", "CHILD_COUNT", -1)), "CHILD_COUNT");
        }

        public static void HierarchyMetadata()
        {
            using var model = ProductsModel(out _);
            model.Tables["Products"].AddHierarchy("Product Hierarchy", "Category", "Subcategory", "Product");
            Check.Equal(1, model.Tables.Sum(t => t.Hierarchies.Count), "count");
            var ms = new MemoryStream();
            model.SaveModelPart(ms);
            var m = ModelReader.Read(ms.ToArray());
            string dim = Encoding.UTF8.GetString(m.Find(@"Products_G\.0\.dim\.xml$|Products_G\.\d+\.dim\.xml$"));
            Check.True(dim.Contains("<Hierarchy><Name>Product Hierarchy</Name><ID>Product Hierarchy</ID>"), "dimension hierarchy");
            Check.True(dim.Contains("<Level><Name>Subcategory</Name><ID>Subcategory</ID><Description/><SourceAttributeID>Subcategory</SourceAttributeID><HideMemberIf>Never</HideMemberIf>"), "level");
            string cube = Encoding.UTF8.GetString(m.Find(@"Model\.18\.cub\.xml$"));
            Check.True(cube.Contains("<Hierarchy><HierarchyID>Product Hierarchy</HierarchyID><Enabled>true</Enabled><Visible>true</Visible><OptimizedState>FullyOptimized</OptimizedState></Hierarchy>"), "cube hierarchy");
            string tbl = Encoding.UTF8.GetString(m.Find(@"\.0\.dim\\Products_G\.2\.tbl\.xml$"));
            Check.True(tbl.Contains("<TableStore>$Category$0$Subcategory$3$Product$9$</TableStore>"), "TableStore");
            Check.True(tbl.Contains("<TableName>U$Products_G$Product Hierarchy</TableName>") || System.Text.RegularExpressions.Regex.IsMatch(tbl, @"<TableName>U\$Products_[0-9a-f-]{36}\$Product Hierarchy</TableName>"), "TableName");
        }

        public static void HierarchyBlanksAndCase()
        {
            var model = new DataModelWriter();
            using (model)
            {
                var t = model.CreateTable("T", new[] { new ColumnDefinition("A", ColumnDataType.String), new ColumnDefinition("B", ColumnDataType.Int64) });
                t.AddRow("x", 1); t.AddRow("X", 1); t.AddRow("x", 2); t.AddRow(null, 5); t.AddRow("y", null); t.AddRow("y", 7);
                model.Tables["T"].AddHierarchy("H", "A", "B");
                var ms = new MemoryStream();
                model.SaveModelPart(ms);
                var m = ModelReader.Read(ms.ToArray());
                string J(int[] a) => string.Join(",", a);
                // Level 1: BLANK, x (x and X are one value), y. Level 2 below each: BLANK sorts first.
                var parents = DecodeTree(m, "T", "H", "PARENT_POS", -1);
                var counts = DecodeTree(m, "T", "H", "CHILD_COUNT", -1);
                var ids = DecodeTree(m, "T", "H", "MULTI_LEVEL_ID", 0);
                Check.Equal("-1,-1,-1,0,1,1,2,2", J(parents), "parents");
                Check.Equal("1,2,2,-1,-1,-1,-1,-1", J(counts), "child counts");
                Check.Equal(2, ids[0], "BLANK node first");
            }
        }

        public static void HierarchyValidation()
        {
            using var model = ProductsModel(out _);
            void Rejects(Action a, string what)
            {
                try { a(); } catch (ArgumentException) { return; }
                throw new Exception("Expected ArgumentException: " + what);
            }
            Tests.Throws<KeyNotFoundException>(() => model.Tables["Nope"].AddHierarchy("H", "Category", "Product"), "unknown table");
            Rejects(() => model.Tables["Products"].AddHierarchy("H", "Category"), "one level");
            Rejects(() => model.Tables["Products"].AddHierarchy("H", "Category", "Nope"), "unknown column");
            Rejects(() => model.Tables["Products"].AddHierarchy("H", "Category", "category"), "same column twice");
            Rejects(() => model.Tables["Products"].AddHierarchy("Category", "Category", "Product"), "name of a column");
            Rejects(() => model.Tables["Products"].AddHierarchy("", "Category", "Product"), "empty name");
            model.Tables["Products"].AddHierarchy("H", "Category", "Product");
            Rejects(() => model.Tables["Products"].AddHierarchy("h", "Subcategory", "Product"), "duplicate name");
        }

        private static string Script(DataModelWriter model)
        {
            var ms = new MemoryStream();
            model.SaveModelPart(ms);
            var m = ModelReader.Read(ms.ToArray());
            // Parse so that the assertions see the text Excel will see, whatever escaping is used.
            var doc = new System.Xml.XmlDocument();
            doc.LoadXml(Encoding.UTF8.GetString(m.Find(@"MdxScript\.1\.scr\.xml$")).TrimStart('\uFEFF'));
            return doc.OuterXml.Replace(" />", "/>").Replace("&lt;", "<").Replace("&gt;", ">").Replace("&amp;", "&");
        }

        public static void MeasureScript()
        {
            var (model, many, one) = Pair(ColumnDataType.Int64);
            using (model)
            {
                one.AddRow(1, "x"); many.AddRow(1, 1);
                model.Tables["Many"].AddMeasure("Total N", "=SUM(Many[N])", MeasureFormat.WholeNumber);
                many.AddMeasure("Plain", "[Total N] * 2 & \"<x>\"");
                model.Tables["One"].AddMeasure("O'Rate", "DIVIDE([Total N],2)", MeasureFormat.Percentage(1)).IsHidden = true;
                Check.Equal(3, model.Tables.Sum(t => t.Measures.Count), "count");
                string x = Script(model);
                Check.True(x.Contains("[__XL_Count of Models] AS 1, VISIBLE = 0;"), "default member");
                Check.True(!x.Contains("__No measures defined"), "placeholder gone");
                Check.True(x.Contains("CREATE MEASURE 'Many'[Total N]=SUM(Many[N]);\n</Text>"), "first command");
                Check.True(x.Contains("[Plain]=[Total N] * 2 & \"<x>\";"), "escaped expression");
                Check.True(x.Contains("CREATE MEASURE 'One'[O'Rate]="), "name with quote");
                Check.True(System.Text.RegularExpressions.Regex.IsMatch(x, @"<Name>FullName</Name>\s*<Value>Total N</Value>"), "FullName annotation");
                Check.True(x.Contains("<CalculationReference>[Total N]</CalculationReference>"), "calc ref");
                Check.True(x.Contains("<FormatString>'#,0'</FormatString>"), "whole number format string");
                Check.True(x.Contains("<Format Format=\"NumberWhole\" ThousandSeparator=\"True\"/>"), "whole number format");
                Check.True(x.Contains("<FormatString>''</FormatString>"), "general format string");
                Check.True(x.Contains("<Format Format=\"General\"/>"), "general format");
                Check.True(x.Contains("<FormatString>'0.0%'</FormatString>"), "percentage");
                model.Tables["Many"].AddMeasure("When", "MAX(Many[N])", MeasureFormat.ShortDate());
                model.Tables["Many"].AddMeasure("Money", "SUM(Many[N])", MeasureFormat.Currency());
                string y = Script(model);
                Check.True(y.Contains("<FormatString>'Short Date'</FormatString>") && y.Contains("<Format Format=\"DateTimeShortDatePattern\"/>"), "short date");
                Check.True(y.Contains("<FormatString>'\"$\"#,0.00;(\"$\"#,0.00);\"$\"#,0.00'</FormatString>"), "currency (en-US default; negatives in parentheses as Windows has them)");
                Check.True(System.Text.RegularExpressions.Regex.IsMatch(x, @"<Name>IsPrivate</Name>\s*<Value>True</Value>"), "hidden");
                var ms = new MemoryStream();
                model.SaveModelPart(ms);
                string cube = Encoding.UTF8.GetString(ModelReader.Read(ms.ToArray()).Find(@"Model\.1\.cub\.xml$"));
                Check.True(cube.Contains("<Value>__XL_Count of Models</Value>"), "cube default measure");
            }
        }

        public static void NoMeasures()
        {
            var (model, many, one) = Pair(ColumnDataType.Int64);
            using (model)
            {
                one.AddRow(1, "x"); many.AddRow(1, 1);
                string x = Script(model);
                Check.True(x.Contains("__No measures defined") && !x.Contains("CREATE MEASURE"), "placeholder");
            }
        }

        public static void MeasureValidation()
        {
            var (model, many, one) = Pair(ColumnDataType.Int64);
            using (model)
            {
                model.Tables["Many"].AddMeasure("M1", "1");
                void Rejects(Action a, string what)
                {
                    try { a(); } catch (ArgumentException) { return; }
                    throw new Exception("Expected ArgumentException: " + what);
                }
                Tests.Throws<KeyNotFoundException>(() => model.Tables["Nope"].AddMeasure("M2", "1"), "unknown table");
                Rejects(() => model.Tables["Many"].AddMeasure("m1", "1"), "duplicate name (case-insensitive)");
                Rejects(() => model.Tables["One"].AddMeasure("m1", "1"), "duplicate name in another table");
                Rejects(() => model.Tables["Many"].AddMeasure("A[B]", "1"), "brackets");
                Rejects(() => model.Tables["Many"].AddMeasure("", "1"), "empty name");
                Rejects(() => model.Tables["Many"].AddMeasure("M3", "  "), "empty expression");
                Rejects(() => model.Tables["Many"].AddMeasure("N", "1"), "name of a column in the same table");
                Rejects(() => model.Tables["Many"].AddMeasure("__x", "1"), "reserved prefix");
            }
        }

        // Dates(Date, Year, Month) with a Calendar hierarchy, related to Sales[Date]: the shape of Excel's H2_date_table.xlsx.
        private static (DataModelWriter model, ModelTable dates, ModelTable sales) DateModel(Action<ModelTable>? fillDates = null)
        {
            var model = new DataModelWriter(new DataModelOptions { XpressCompression = false });
            var dates = model.CreateTable("Dates", new[]
            {
                new ColumnDefinition("Date", ColumnDataType.DateTime),
                new ColumnDefinition("Year", ColumnDataType.Int64),
                new ColumnDefinition("Month", ColumnDataType.String),
            });
            if (fillDates is not null) fillDates(dates);
            else for (int i = 0; i < 40; i++) { var d = new DateTime(2025, 1, 1).AddDays(i); dates.AddRow(d, d.Year, d.ToString("MMMM", CultureInfo.InvariantCulture)); }
            var sales = model.CreateTable("Sales", new[] { new ColumnDefinition("Date", ColumnDataType.DateTime), new ColumnDefinition("N", ColumnDataType.Int64) });
            sales.AddRow(new DateTime(2025, 1, 2), 1);
            model.AddRelationship(model.Tables["Sales"]["Date"], model.Tables["Dates"]["Date"]);
            model.Tables["Dates"].AddHierarchy("Calendar", "Year", "Month");
            return (model, dates, sales);
        }

        /// <summary>The changes Excel makes when Dates[Date] is marked as date table (SPEC section 9, diffed against H2_date_table.xlsx).</summary>
        public static void DateTableMetadata()
        {
            var (model, _, _) = DateModel();
            using (model)
            {
                model.Tables["Dates"].MarkAsDateTable("Date");
                Check.Equal("Dates", model.DateTable?.Name, "DateTableName");
                var ms = new MemoryStream();
                model.SaveModelPart(ms);
                var m = ModelReader.Read(ms.ToArray());
                string Text(string pattern) => Encoding.UTF8.GetString(m.Find(pattern));

                string dim = Text(@"\.0\.db\\Dates_G\.6\.dim\.xml$");
                Check.True(dim.Contains("<Type>Time</Type>"), "dimension type Time");
                int rowNumber = dim.IndexOf("<ID>__XL_RowNumber</ID>", StringComparison.Ordinal);
                int dateAttr = dim.IndexOf("<ID>Date</ID>", StringComparison.Ordinal);
                Check.True(dim.Substring(rowNumber, dateAttr - rowNumber).Contains("<Usage>Regular</Usage>"), "row number is no longer the key");
                Check.True(!dim.Substring(rowNumber, dateAttr - rowNumber).Contains("<AttributeID>Date</AttributeID>"), "row number no longer relates to Date");
                string dateXml = dim.Substring(dateAttr, dim.IndexOf("</Attribute>", dateAttr, StringComparison.Ordinal) - dateAttr);
                Check.True(dateXml.Contains("<Usage>Key</Usage>"), "Date is the key");
                Check.True(dateXml.Contains("<AttributeRelationships><AttributeRelationship><AttributeID>__XL_RowNumber</AttributeID>"), "Date relates to the row number");
                Check.True(dim.Contains("<StructureVersion>3</StructureVersion><DataIDVersion>2</DataIDVersion><IncDataIDVersion>2</IncDataIDVersion><CacheVersion>2</CacheVersion><DataVersion>3</DataVersion>"), "dimension counters");

                string tbl = Text(@"\.0\.dim\\Dates_G\.3\.tbl\.xml$");
                Check.True(tbl.Contains("<Version xsi:type=\"xsd:int\">3</Version>"), "table version 3");
                Check.True(System.Text.RegularExpressions.Regex.IsMatch(tbl, @"name=""__XL_RowNumber""(?s:.*?)<ColumnFlags xsi:type=""xsd:long"">27</ColumnFlags>"), "row number ColumnFlags 27");
                Check.True(System.Text.RegularExpressions.Regex.IsMatch(tbl, @"name=""Date""(?s:.*?)<ColumnFlags xsi:type=""xsd:long"">15</ColumnFlags>"), "date ColumnFlags 15");
                Check.True(tbl.Contains("name=\"2.H$Dates_"), "hash index name");
                string h = Text(@"Dates_G\.0\.dim\\H\$Dates_G\$Date\.2\.tbl\.xml$");
                Check.True(h.Contains("<Version xsi:type=\"xsd:int\">2</Version>") && h.Contains("name=\"2.H$Dates_"), "date column hierarchy table");
                m.Find(@"Dates_G\.0\.dim\\2\.H\$Dates_G\$Date\.POS_TO_ID\.0\.idf$");
                m.Find(@"Dates_G\.0\.dim\\2\.H\$Dates_G\$Date\.hidx$");
                m.Find(@"Dates_G\.0\.dim\\0\.H\$Dates_G\$Year\.hidx$");

                // Sales is the many side: its table, the relationship table and the dimension move on by one.
                Check.True(Text(@"\.0\.dim\\Sales_G\.1\.tbl\.xml$").Contains("<Version xsi:type=\"xsd:int\">1</Version>"), "Sales table version 1");
                string rel = Text(@"R\$Sales_G\$G\.2\.tbl\.xml$");
                Check.True(rel.Contains("<Version xsi:type=\"xsd:int\">2</Version>") && rel.Contains("name=\"2.R$Sales_"), "relationship table version 2");
                m.Find(@"Sales_G\.0\.dim\\2\.R\$Sales_G\$G\.INDEX\.0\.idf$");
                string sales = Text(@"\.0\.db\\Sales_G\.4\.dim\.xml$");
                Check.True(sales.Contains("<DataVersion>1</DataVersion>"), "Sales dimension data version");
                m.Find(@"Sales_G\.0\.dim\\info\.4\.xml$");

                // Counters of the cube, the measure groups, the script and the partitions.
                string cube = Text(@"Model\.21\.cub\.xml$");
                Check.True(cube.Contains("<StructureVersion>12</StructureVersion>") && cube.Contains("MdxScript.10.scr.xml") && cube.Contains("Dates_") && cube.Contains(".21.det.xml"), "cube");
                string det = Text(@"Dates_G\.21\.det\.xml$");
                Check.True(det.Contains("<StructureVersion>9</StructureVersion><CacheVersion>2</CacheVersion>") && det.Contains(".10.prt.xml"), "measure group");
                m.Find(@"MdxScript\.10\.scr\.xml$");
                Check.True(Text(@"Sales_G\.10\.prt\.xml$").Contains("<DataVersion>1</DataVersion>"), "Sales partition data version 1");
                Check.True(Text(@"Dates_G\.10\.prt\.xml$").Contains("<DataVersion>1</DataVersion>"), "Dates partition data version 1");
                m.Find(@"Sales_G\.0\.prt\\info\.10\.xml$");
            }
        }

        public static void GeneratedDateTable()
        {
            using var model = new DataModelWriter(new DataModelOptions { XpressCompression = false, Culture = CultureInfo.GetCultureInfo(1033) });
            var dates = model.AddDateTable("Dates", new DateOnly(2024, 1, 1), new DateOnly(2024, 12, 31));
            Check.Equal(366L, dates.RowCount, "leap year has 366 rows");
            Check.Equal("Dates", model.DateTable?.Name, "marked by default");
            Check.Equal(1, model.Tables.Sum(t => t.Hierarchies.Count), "Calendar hierarchy");
            Check.Equal("Date,Year,QuarterNumber,Quarter,MonthNumber,Month,YearMonthNumber,YearMonth,Day,DayOfWeekNumber,DayOfWeek", string.Join(",", dates.Columns.Select(c => c.Name)), "columns");
            var sales = model.CreateTable("Sales", new[] { new ColumnDefinition("Date", ColumnDataType.DateTime), new ColumnDefinition("N", ColumnDataType.Int64) });
            sales.AddRow(new DateTime(2024, 3, 5), 1);
            model.AddRelationship(model.Tables["Sales"]["Date"], model.Tables["Dates"]["Date"]);
            var ms = new MemoryStream();
            model.SaveModelPart(ms);
            var m = ModelReader.Read(ms.ToArray());
            string dim = Encoding.UTF8.GetString(m.Find(@"\.0\.db\\Dates_G\.[0-9]+\.dim\.xml$"));
            Check.True(dim.Contains("<Type>Time</Type>"), "dimension type Time");
            Check.Equal(4, System.Text.RegularExpressions.Regex.Matches(dim, "<OrderBy>AttributeKey</OrderBy>").Count, "Quarter, Month, YearMonth and DayOfWeek are sorted by their numbers");
            Check.True(dim.Contains("<Level><Name>Quarter</Name>") && dim.Contains("<Hierarchy><Name>Calendar</Name>"), "hierarchy levels");

            void Rejects<T>(Action a, string what) where T : Exception
            {
                try { a(); } catch (T) { return; }
                throw new Exception("Expected " + typeof(T).Name + ": " + what);
            }
            Rejects<ArgumentException>(() => model.AddDateTable("Bad", new DateOnly(2025, 2, 1), new DateOnly(2025, 1, 1)), "to before from");
            Rejects<InvalidOperationException>(() => model.AddDateTable("Second", new DateOnly(2025, 1, 1), new DateOnly(2025, 1, 2)), "second date table");
            Check.Equal(2, model.Tables.Count, "rejected calls add no table");
        }

        private sealed record DateRow(DateOnly Date, long Year, string Month);

        public static void DateTableSourceTypes()
        {
            void Saves(DataModelWriter model) { using var ms = new MemoryStream(); model.SaveModelPart(ms); }

            // Rows given as DateOnly, DateTime and DateTimeOffset values in the same column: all become model dates.
            var (mixed, _, _) = DateModel(t =>
            {
                t.AddRow(new DateOnly(2025, 1, 1), 2025, "January");
                t.AddRow(new DateTime(2025, 1, 2), 2025, "January");
                t.AddRow(new DateTimeOffset(2025, 1, 3, 0, 0, 0, TimeSpan.FromHours(2)), 2025, "January");
            });
            using (mixed) { mixed.Tables["Dates"].MarkAsDateTable("Date"); Saves(mixed); }

            // A typed source with a DateOnly property maps to the same column type.
            using var typed = new DataModelWriter(new DataModelOptions { XpressCompression = false });
            typed.AddTable("Dates", new[] { new DateRow(new DateOnly(2025, 1, 1), 2025, "January"), new DateRow(new DateOnly(2025, 1, 2), 2025, "January") });
            Check.Equal(ColumnDataType.DateTime, ColumnDefinition.FromClrType(typeof(DateOnly)), "DateOnly maps to DateTime");
            Check.Equal(ColumnDataType.DateTime, ColumnDefinition.FromClrType(typeof(DateTimeOffset)), "DateTimeOffset maps to DateTime");
            var sales = typed.CreateTable("Sales", new[] { new ColumnDefinition("Date", ColumnDataType.DateTime) });
            sales.AddRow(new DateOnly(2025, 1, 1));
            typed.AddRelationship(typed.Tables["Sales"]["Date"], typed.Tables["Dates"]["Date"]);
            typed.Tables["Dates"].AddHierarchy("Calendar", "Year", "Month");
            typed.Tables["Dates"].MarkAsDateTable("Date");
            Saves(typed);
        }

        /// <summary>
        /// February year end, data for 2025: FiscalYear 2025 holds only Q4 (Jan, Feb), FiscalYear 2026 holds Q1..Q4 (Mar..Dec).
        /// Q4 is the first quarter that appears in the data, but under 2026 our tree lists it after Q3. Excel's pivot showed Q4 first
        /// (data-id order), so this pins what we write, not what Excel displays: see SPEC section 7, "Open".
        /// </summary>
        public static void FiscalHierarchyOrder()
        {
            using var model = new DataModelWriter(new DataModelOptions { XpressCompression = false, Culture = CultureInfo.GetCultureInfo(1033) });
            model.AddDateTable("Dates", new DateOnly(2025, 1, 1), new DateOnly(2025, 12, 31), new DateTableOptions { YearEndMonth = 2, CompleteYears = false });
            var sales = model.CreateTable("Sales", new[] { new ColumnDefinition("Date", ColumnDataType.DateTime) });
            sales.AddRow(new DateOnly(2025, 1, 1));
            model.AddRelationship(model.Tables["Sales"]["Date"], model.Tables["Dates"]["Date"]);
            var ms = new MemoryStream();
            model.SaveModelPart(ms);
            var m = ModelReader.Read(ms.ToArray());
            var ids = DecodeTree(m, "Dates", "Fiscal Calendar", "MULTI_LEVEL_ID", 0);
            // Level 0: two years (2 nodes). Level 1: Q4 of 2025, then Q1..Q4 of 2026. Quarter ids by first appearance: Q4=3, Q1=4, Q2=5, Q3=6.
            Check.Equal("3,4,5,6,3", string.Join(",", ids.Skip(2).Take(5)), "quarter nodes: 2025 Q4, then 2026 Q1, Q2, Q3, Q4");
        }

        // Excel writes StructureType Unnatural when a level member has more than one parent (H5_sort_reapplied.xlsx: Q4 under
        // FiscalYear 2025 and 2026) and Natural otherwise; with Natural, Excel's pivot ignored the sort order of such levels.
        public static void HierarchyStructureType()
        {
            string Types(DateOnly from, DateOnly to, int yearEnd)
            {
                using var model = new DataModelWriter(new DataModelOptions { XpressCompression = false, Culture = CultureInfo.GetCultureInfo(1033) });
                model.AddDateTable("Dates", from, to, new DateTableOptions { YearEndMonth = yearEnd, CompleteYears = false });
                var sales = model.CreateTable("Sales", new[] { new ColumnDefinition("Date", ColumnDataType.DateTime) });
                sales.AddRow(from);
                model.AddRelationship(model.Tables["Sales"]["Date"], model.Tables["Dates"]["Date"]);
                var ms = new MemoryStream();
                model.SaveModelPart(ms);
                string dim = Encoding.UTF8.GetString(ModelReader.Read(ms.ToArray()).Find(@"\.0\.db\\Dates_G\.[0-9]+\.dim\.xml$"));
                var found = System.Text.RegularExpressions.Regex.Matches(dim, @"<Hierarchy><Name>([^<]+)</Name>.*?<ddl300:StructureType>(\w+)</ddl300:StructureType>");
                return string.Join(",", found.Select(m => m.Groups[1].Value + "=" + m.Groups[2].Value));
            }
            Check.Equal("Calendar=Natural", Types(new DateOnly(2025, 1, 1), new DateOnly(2025, 12, 31), 12), "one calendar year, December year end");
            Check.Equal("Calendar=Unnatural", Types(new DateOnly(2024, 1, 1), new DateOnly(2025, 12, 31), 12), "Q1 under two years");
            Check.Equal("Calendar=Natural,Fiscal Calendar=Unnatural", Types(new DateOnly(2025, 1, 1), new DateOnly(2025, 12, 31), 2), "Q4 under FiscalYear 2025 and 2026");
            Check.Equal("Calendar=Unnatural,Fiscal Calendar=Natural", Types(new DateOnly(2025, 3, 1), new DateOnly(2026, 2, 28), 2), "one whole fiscal year spans two calendar years: Q1 under 2025 and 2026, but one FiscalYear");
        }

        public static void FiscalDateTable()
        {
            // June year end: July 2024 starts FiscalYear 2025; the fiscal year is named by the year in which it ends.
            (int, int, int) F(int y, int m, int d, int fye) => DataModelWriter.FiscalPeriod(new DateOnly(y, m, d), fye);
            Check.Equal((2025, 1, 1), F(2024, 7, 1, 6), "July 2024, June year end");
            Check.Equal((2025, 3, 1), F(2024, 9, 30, 6), "September is fiscal month 3, quarter 1");
            Check.Equal((2025, 4, 2), F(2024, 10, 1, 6), "October starts quarter 2");
            Check.Equal((2025, 12, 4), F(2025, 6, 30, 6), "June 30 closes the year");
            Check.Equal((2026, 1, 1), F(2025, 7, 1, 6), "next year");
            Check.Equal((2024, 1, 1), F(2024, 1, 15, 12), "December year end equals the calendar year");
            Check.Equal((2025, 1, 1), F(2024, 12, 31, 11), "November year end: December 2024 is fiscal month 1 of 2025");
            Check.Equal((2024, 12, 4), F(2024, 1, 31, 1), "January year end: January is the last month");

            using var model = new DataModelWriter(new DataModelOptions { XpressCompression = false, Culture = CultureInfo.GetCultureInfo(1033) });
            var dates = model.AddDateTable("Dates", new DateOnly(2024, 1, 1), new DateOnly(2025, 12, 31), new DateTableOptions { YearEndMonth = 6 });
            Check.Equal("Date,Year,QuarterNumber,Quarter,MonthNumber,Month,YearMonthNumber,YearMonth,Day,DayOfWeekNumber,DayOfWeek,FiscalYear,FiscalQuarterNumber,FiscalQuarter,FiscalMonthNumber,FiscalMonth",
                string.Join(",", dates.Columns.Select(c => c.Name)), "columns with fiscal ones");
            Check.Equal(2, model.Tables.Sum(t => t.Hierarchies.Count), "Calendar and Fiscal Calendar");
            var sales = model.CreateTable("Sales", new[] { new ColumnDefinition("Date", ColumnDataType.DateTime), new ColumnDefinition("N", ColumnDataType.Int64) });
            sales.AddRow(new DateTime(2024, 8, 5), 1);
            model.AddRelationship(model.Tables["Sales"]["Date"], model.Tables["Dates"]["Date"]);
            var ms = new MemoryStream();
            model.SaveModelPart(ms);   // sort-by columns validate that each name maps to one number
            string dim = Encoding.UTF8.GetString(ModelReader.Read(ms.ToArray()).Find(@"\.0\.db\\Dates_G\.[0-9]+\.dim\.xml$"));
            Check.Equal(6, System.Text.RegularExpressions.Regex.Matches(dim, "<OrderBy>AttributeKey</OrderBy>").Count, "six sorted columns");
            Check.True(dim.Contains("<Hierarchy><Name>Fiscal Calendar</Name>"), "fiscal hierarchy");

            using var dec = new DataModelWriter(new DataModelOptions { XpressCompression = false });
            Check.Equal(11, dec.AddDateTable("D", new DateOnly(2025, 1, 1), new DateOnly(2025, 1, 2)).Columns.Count, "December year end adds no fiscal columns");
            Check.Equal(1, dec.Tables.Sum(t => t.Hierarchies.Count), "only Calendar");
            // Invalid settings are rejected when the options are created, before any table work.
            foreach (int month in new[] { 0, 13, -1 })
            {
                try { _ = new DateTableOptions { YearEndMonth = month }; throw new Exception("Expected ArgumentOutOfRangeException for month " + month); }
                catch (ArgumentOutOfRangeException) { }
            }
        }

        // Names of the attributes marked AttributeHierarchyVisible=false in a dimension (Name) or in the cube's dimension list (AttributeID).
        private static string HiddenAttributes(string xml, string dimensionId = "")
        {
            System.Xml.Linq.XNamespace ns = "http://schemas.microsoft.com/analysisservices/2003/engine";
            var doc = System.Xml.Linq.XDocument.Parse(xml);
            IEnumerable<System.Xml.Linq.XElement> attributes = dimensionId.Length == 0
                ? doc.Descendants(ns + "Attributes").First().Elements(ns + "Attribute")
                : doc.Descendants(ns + "Dimension").First(d => (d.Element(ns + "ID")?.Value ?? "").StartsWith(dimensionId, StringComparison.Ordinal)).Element(ns + "Attributes")!.Elements(ns + "Attribute");
            return string.Join(",", attributes.Where(a => a.Element(ns + "AttributeHierarchyVisible")?.Value == "false")
                .Select(a => (a.Element(ns + "Name") ?? a.Element(ns + "AttributeID"))!.Value));
        }

        // A model without hierarchies or relationships stays in the base profile (SPEC 5.1) when columns are hidden:
        // only AttributeHierarchyVisible changes, in the dimension and in the cube.
        public static void HiddenColumnsPlainTable()
        {
            void Build(bool hide, out string dim, out string cube, out string[] names)
            {
                using var model = new DataModelWriter(new DataModelOptions { XpressCompression = false });
                var t = model.CreateTable("T", new[]
                {
                    new ColumnDefinition("A", ColumnDataType.Int64),
                    new ColumnDefinition("B", ColumnDataType.String) { IsHidden = hide },
                    new ColumnDefinition("C", ColumnDataType.Double) { IsHidden = hide },
                });
                t.AddRow(1, "x", 1.5); t.AddRow(2, "y", 2.5);
                var ms = new MemoryStream();
                model.SaveModelPart(ms);
                var m = ModelReader.Read(ms.ToArray());
                dim = Encoding.UTF8.GetString(m.Find(@"\.0\.db\\T_G\.1\.dim\.xml$"));
                cube = Encoding.UTF8.GetString(m.Find(@"Model\.1\.cub\.xml$"));
                names = m.Files.Keys.Select(k => System.Text.RegularExpressions.Regex.Replace(ModelReader.Normalize(k), "[0-9A-F]{20}", "DB")).OrderBy(x => x, StringComparer.Ordinal).ToArray();
            }
            Build(true, out var dimH, out var cubeH, out var namesH);
            Build(false, out var dimV, out var cubeV, out var namesV);
            Check.Equal("__XL_RowNumber,B,C", HiddenAttributes(dimH), "hidden attributes in the dimension");
            Check.Equal("__XL_RowNumber,B,C", HiddenAttributes(cubeH, "T_"), "hidden attributes in the cube");
            Check.Equal("__XL_RowNumber", HiddenAttributes(dimV), "nothing else hidden by default (dimension)");
            Check.Equal("__XL_RowNumber", HiddenAttributes(cubeV, "T_"), "nothing else hidden by default (cube)");
            Check.True(namesH.SequenceEqual(namesV), "hiding adds, removes or renames no file");
        }

        // The columns Matthew hid in Excel for H1_hidden_columns.xlsx: keys, hierarchy levels and sort columns. In the model they stay
        // usable (relationships, hierarchies, sorting), only the attribute hierarchies are marked invisible.
        public static void HiddenColumnsKeysAndLevels()
        {
            using var model = new DataModelWriter(new DataModelOptions { XpressCompression = false });
            var products = model.CreateTable("Products", new[]
            {
                new ColumnDefinition("ProductKey", ColumnDataType.Int64) { IsHidden = true },
                new ColumnDefinition("Category", ColumnDataType.String) { SortByColumn = "CategoryOrder", IsHidden = true },
                new ColumnDefinition("CategoryOrder", ColumnDataType.Int64) { IsHidden = true },
                new ColumnDefinition("Product", ColumnDataType.String) { IsHidden = true },
            });
            products.AddRow(1, "Bikes", 1, "Road"); products.AddRow(2, "Clothing", 2, "Jersey");
            var sales = model.CreateTable("Sales", new[]
            {
                new ColumnDefinition("ProductKey", ColumnDataType.Int64) { IsHidden = true },
                new ColumnDefinition("Qty", ColumnDataType.Int64),
            });
            sales.AddRow(1, 3); sales.AddRow(2, 4);
            model.AddRelationship(model.Tables["Sales"]["ProductKey"], model.Tables["Products"]["ProductKey"]);
            model.Tables["Products"].AddHierarchy("Product Hierarchy", "Category", "Product");
            model.Tables["Sales"].AddMeasure("Total Qty", "SUM(Sales[Qty])", MeasureFormat.WholeNumber);
            var ms = new MemoryStream();
            model.SaveModelPart(ms);   // saves: hidden keys still pass the relationship checks, hidden levels the hierarchy and sort checks
            var m = ModelReader.Read(ms.ToArray());
            string pDim = Encoding.UTF8.GetString(m.Find(@"\.0\.db\\Products_G\.7\.dim\.xml$"));
            string sDim = Encoding.UTF8.GetString(m.Find(@"\.0\.db\\Sales_G\.4\.dim\.xml$"));
            string cube = Encoding.UTF8.GetString(m.Find(@"Model\.21\.cub\.xml$"));
            Check.Equal("__XL_RowNumber,ProductKey,Category,CategoryOrder,Product", HiddenAttributes(pDim), "Products attributes hidden");
            Check.Equal("__XL_RowNumber,ProductKey", HiddenAttributes(sDim), "Sales attributes hidden");
            Check.Equal("__XL_RowNumber,ProductKey,Category,CategoryOrder,Product", HiddenAttributes(cube, "Products_"), "Products hidden in the cube");
            Check.Equal("__XL_RowNumber,ProductKey", HiddenAttributes(cube, "Sales_"), "Sales hidden in the cube");
            Check.True(pDim.Contains("<Hierarchy><Name>Product Hierarchy</Name>") && cube.Contains("<HierarchyID>Product Hierarchy</HierarchyID><Enabled>true</Enabled><Visible>true</Visible>"), "the hierarchy stays visible");
            Check.True(pDim.Contains("<OrderByAttributeID>CategoryOrder</OrderByAttributeID>"), "sort by a hidden column still works");
            Check.True(sDim.Contains("<ddl300_300:Relationship>"), "relationship on a hidden key");
        }

        // After a hide operation Excel advances the counters (H8_hide_one_column.xlsx, H9_hide_one_hierarchy_level.xlsx):
        // cube +3 (structure +1), measure groups +2, the dimension of a table with hidden columns +1, or +2 if it holds a hierarchy.
        // Without them Excel loaded such a model but failed every later edit in ASDetail::OnClone2.
        public static void HiddenColumnsProfile()
        {
            ModelReader Save(bool hide, bool hierarchy, bool hideLevel = false)
            {
                using var model = new DataModelWriter(new DataModelOptions { XpressCompression = false });
                var products = model.CreateTable("Products", new[]
                {
                    new ColumnDefinition("ProductKey", ColumnDataType.Int64),
                    new ColumnDefinition("Category", ColumnDataType.String),
                    new ColumnDefinition("Product", ColumnDataType.String) { IsHidden = hide && hideLevel },
                });
                products.AddRow(1, "Bikes", "Road"); products.AddRow(2, "Clothing", "Jersey");
                var sales = model.CreateTable("Sales", new[]
                {
                    new ColumnDefinition("ProductKey", ColumnDataType.Int64),
                    new ColumnDefinition("Qty", ColumnDataType.Int64) { IsHidden = hide && !hideLevel },
                });
                sales.AddRow(1, 3); sales.AddRow(2, 4);
                model.AddRelationship(model.Tables["Sales"]["ProductKey"], model.Tables["Products"]["ProductKey"]);
                if (hierarchy) model.Tables["Products"].AddHierarchy("Browse", "Category", "Product");
                model.Tables["Sales"].AddMeasure("Total Qty", "SUM(Sales[Qty])");
                var ms = new MemoryStream();
                model.SaveModelPart(ms);
                return ModelReader.Read(ms.ToArray());
            }
            // One hidden plain column (H8): Sales dimension 3 -> 4 (info file follows), cube 21, measure groups 20, cube structure 12.
            var plain = Save(hide: true, hierarchy: true);
            plain.Find(@"Sales_G\.0\.dim\\info\.4\.xml$");
            plain.Find(@"\.0\.db\\Sales_G\.4\.dim\.xml$");
            plain.Find(@"\.0\.db\\Products_G\.5\.dim\.xml$");   // no hidden column there
            Check.True(Encoding.UTF8.GetString(plain.Find(@"Model\.21\.cub\.xml$")).Contains("<StructureVersion>12</StructureVersion>"), "cube structure 12");
            plain.Find(@"Sales_G\.20\.det\.xml$"); plain.Find(@"Products_G\.20\.det\.xml$");
            Check.True(Encoding.UTF8.GetString(plain.Find(@"Model\.21\.cub\.xml$")).Contains("Sales_") && Encoding.UTF8.GetString(plain.Find(@"Model\.21\.cub\.xml$")).Contains(".20.det.xml"), "cube lists the .20. measure groups");
            // One hidden hierarchy level (H9): Products dimension 5 -> 7, no info file for a table with a hierarchy.
            var level = Save(hide: true, hierarchy: true, hideLevel: true);
            level.Find(@"\.0\.db\\Products_G\.7\.dim\.xml$");
            level.Find(@"\.0\.db\\Sales_G\.3\.dim\.xml$");
            // Timestamps as in H9: only the edited dimension and the cube are stamped after their creation, the cube's LastUpdate a bit later.
            string Stamp(string xml, string tag) => System.Text.RegularExpressions.Regex.Match(xml, "<" + tag + ">([^<]*)</" + tag + ">").Groups[1].Value;
            string pDim = Encoding.UTF8.GetString(level.Find(@"\.0\.db\\Products_G\.7\.dim\.xml$"));
            string sDim = Encoding.UTF8.GetString(level.Find(@"\.0\.db\\Sales_G\.3\.dim\.xml$"));
            string cubeXml = Encoding.UTF8.GetString(level.Find(@"Model\.21\.cub\.xml$"));
            string detXml = Encoding.UTF8.GetString(level.Find(@"Products_G\.20\.det\.xml$"));
            Check.True(string.CompareOrdinal(Stamp(pDim, "LastSchemaUpdate"), Stamp(pDim, "CreatedTimestamp")) > 0, "edited dimension stamped after creation");
            Check.Equal(Stamp(sDim, "CreatedTimestamp"), Stamp(sDim, "LastSchemaUpdate"), "untouched dimension keeps its creation time");
            Check.Equal(Stamp(detXml, "CreatedTimestamp"), Stamp(detXml, "LastSchemaUpdate"), "measure groups are not stamped");
            Check.Equal(Stamp(pDim, "LastSchemaUpdate"), Stamp(cubeXml, "LastSchemaUpdate"), "cube stamped with the edit time");
            Check.True(string.CompareOrdinal(Stamp(cubeXml, "LastUpdate"), Stamp(cubeXml, "LastSchemaUpdate")) > 0, "cube LastUpdate after LastSchemaUpdate");
            Check.Equal(Stamp(pDim, "CreatedTimestamp"), Stamp(pDim, "LastProcessed"), "processed at creation");
            // Nothing hidden, or no hierarchy: the earlier profiles are unchanged.
            Save(hide: false, hierarchy: true).Find(@"Model\.18\.cub\.xml$");
            Save(hide: true, hierarchy: false).Find(@"Model\.1\.cub\.xml$");
        }

        // With every Ordinal 0 the engine told the measure groups apart only by hash order of their ids: whether Create Hierarchy and
        // Sort by Column worked in Excel depended on the random ids (ASDetail::OnClone2). Excel numbers them distinctly.
        // Products(Category, Product) with a hierarchy and Sales(ProductKey, Qty, Amount) related to it: the shape of calcbase_v1, the
        // base of Excel's H6, H7 and H10 samples.
        private static (DataModelWriter model, ModelTable products, ModelTable sales) CalcModel(bool hierarchy = true, int salesRows = 40)
        {
            var model = new DataModelWriter(new DataModelOptions { XpressCompression = false });
            var products = model.CreateTable("Products", new[]
            {
                new ColumnDefinition("ProductKey", ColumnDataType.Int64),
                new ColumnDefinition("Category", ColumnDataType.String),
                new ColumnDefinition("Product", ColumnDataType.String),
            });
            for (int i = 1; i <= 6; i++) products.AddRow(i, i <= 3 ? "A" : "B", "P" + i);
            var sales = model.CreateTable("Sales", new[]
            {
                new ColumnDefinition("ProductKey", ColumnDataType.Int64),
                new ColumnDefinition("Qty", ColumnDataType.Int64),
                new ColumnDefinition("Amount", ColumnDataType.Currency),
            });
            for (int i = 0; i < salesRows; i++) sales.AddRow(1 + i % 6, 1 + i % 5, (decimal)(10 + i % 7));
            model.AddRelationship(sales["ProductKey"], products["ProductKey"]);
            if (hierarchy) products.AddHierarchy("Product Hierarchy", "Category", "Product");
            return (model, products, sales);
        }

        // H10: a calculated column added in manual calculation mode. The file holds the formula only; Excel calculates the values
        // when a PivotTable uses the column. Compare with SPEC section 10 and tools/python/xd.py against H10_calculated_column_manual.xlsx.
        public static void CalculatedColumnProfile()
        {
            var (model, products, sales) = CalcModel();
            using (model)
            {
                var column = sales.AddCalculatedColumn("Line Total", "=[Qty]*[Amount]", ColumnDataType.Currency);
                Check.Equal("[Qty]*[Amount]", column.Expression, "leading = removed");
                Check.Equal("'Sales'[Line Total]", column.ToString(), "DAX notation");
                Check.Equal(0, sales.Columns.Count(c => c.Name == "Line Total"), "not a column that takes values");
                var ms = new MemoryStream();
                model.SaveModelPart(ms);
                var m = ModelReader.Read(ms.ToArray());
                string Text(string pattern) => Encoding.UTF8.GetString(m.Find(pattern));

                string dim = Text(@"Sales_G\.7\.dim\.xml$");
                Check.True(dim.Contains("<Expression>[Qty]*[Amount]</Expression>"), "formula in the dimension");
                Check.Equal(2, System.Text.RegularExpressions.Regex.Matches(dim, "<ddl200:ProcessingState>Unprocessed</ddl200:ProcessingState>|<ddl300:AttributeHierarchyProcessingState>Unprocessed</ddl300:AttributeHierarchyProcessingState>").Count, "unprocessed states");
                Check.True(dim.Contains("<ddl200:InferredDatatype>Currency</ddl200:InferredDatatype>"), "inferred type");
                Check.True(dim.Contains("<DataIDVersion>2</DataIDVersion>"), "data id version");
                Check.True(!dim.Contains("<Value/>"), "empty annotation values dropped after the edit");
                Text(@"Products_G\.6\.dim\.xml$");
                string cube = Text(@"Model\.25\.cub\.xml$");
                Check.True(cube.Contains("<AttributeID>Line Total</AttributeID>"), "cube attribute");
                Check.True(cube.Contains("<StructureVersion>17</StructureVersion>"), "cube structure version");
                Check.True(Text(@"Sales_G\.25\.det\.xml$").Contains("ddl200_200:ExpressionBinding"), "measure group attribute");
                Text(@"MdxScript\.11\.scr\.xml$");
                Text(@"Sales_G\.11\.prt\.xml$");
                Text(@"\.cub\\info\.25\.xml$");

                string tbl = Text(@"\.0\.dim\\Sales_G\.3\.tbl\.xml$");
                Check.True(tbl.Contains("<Settings xsi:type=\"xsd:long\">4097</Settings>"), "table settings");
                Check.True(tbl.Contains("name=\"Line Total\"") && tbl.Contains("<ColumnFlags xsi:type=\"xsd:long\">8</ColumnFlags>"), "raw column");
                Check.True(tbl.Contains("<IsProcessed xsi:type=\"xsd:boolean\">false</IsProcessed>"), "intrinsic hierarchy not processed");

                var idf = m.Find(@"\.0\.dim\\2\.Sales_G\.Line Total\.0\.idf$");
                Check.Equal(152, idf.Length, "stub size");
                Check.Equal(16, (int)BitConverter.ToUInt64(idf, 0), "run slots");
                // 40 rows are fewer than the 64 of a run-length entry, so they are bit-packed (SPEC section 10, H23).
                Check.Equal(0xFFFFFFFFu, BitConverter.ToUInt32(idf, 8), "bit-packed entry");
                Check.Equal(40, (int)BitConverter.ToUInt32(idf, 12), "rows");
                Check.True(!m.Files.Keys.Any(k => k.Contains("Line Total") && (k.Contains("H$") || k.EndsWith(".dictionary", StringComparison.Ordinal))),
                    "no hierarchy table and no dictionary");
                Check.Equal(1, m.Files.Keys.Count(k => System.Text.RegularExpressions.Regex.IsMatch(k, @"\.prt\\info\.\d+\.xml$")), "partition information of the table with the column only");
            }
        }

        // What differs between the types in the samples H10 (currency) and H11 to H15: the inferred type and format of the dimension
        // attribute, and the type code and dictionary class of the unprocessed column in the table XML.
        public static void CalculatedColumnTypes()
        {
            var expected = new (ColumnDataType Type, string Inferred, string Format, string? FormatString, int DbType, int XmType, string Dictionary)[]
            {
                (ColumnDataType.Currency, "Currency", "Currency", null, 6, 0, "XM_Long"),
                (ColumnDataType.Int64, "BigInt", "General", null, 20, 0, "XM_Long"),
                (ColumnDataType.Double, "Double", "General", null, 5, 1, "XM_Real"),
                (ColumnDataType.DateTime, "Date", "DateTimeGeneral", "General Date", 7, 1, "XM_Real"),
                (ColumnDataType.Boolean, "Boolean", "Boolean", "\"TRUE\";\"TRUE\";\"FALSE\"", 11, 0, "XM_Long"),
                (ColumnDataType.String, "WChar", "Text", null, 130, 2, "XM_String"),
            };
            foreach (var e in expected)
            {
                var (model, _, sales) = CalcModel();
                using (model)
                {
                    sales.AddCalculatedColumn("Calc", "1", e.Type);
                    var ms = new MemoryStream();
                    model.SaveModelPart(ms);
                    var m = ModelReader.Read(ms.ToArray());
                    string dim = Encoding.UTF8.GetString(m.Find(@"Sales_G\.7\.dim\.xml$"));
                    string what = e.Type.ToString();
                    Check.True(dim.Contains("<ddl200:InferredDatatype>" + e.Inferred + "</ddl200:InferredDatatype>"), what + ": inferred type");
                    int at = dim.IndexOf("<Attribute><Name>Calc</Name>", StringComparison.Ordinal);
                    string attribute = dim[at..dim.IndexOf("</Attribute>", at, StringComparison.Ordinal)];
                    Check.True(attribute.Contains("<Format Format=\"" + e.Format + "\""), what + ": format annotation");
                    string? formatString = System.Text.RegularExpressions.Regex.Match(attribute, "<ddl300_300:FormatString>([^<]*)</ddl300_300:FormatString>") is { Success: true } f
                        ? System.Net.WebUtility.HtmlDecode(f.Groups[1].Value) : null;
                    if (e.Type == ColumnDataType.Currency) Check.True(formatString is not null && formatString.Contains("#,0.00"), what + ": format string follows the culture");
                    else Check.Equal(e.FormatString, formatString, what + ": format string");
                    string tbl = Encoding.UTF8.GetString(m.Find(@"\.0\.dim\\Sales_G\.3\.tbl\.xml$"));
                    int col = tbl.IndexOf("name=\"Calc\"", StringComparison.Ordinal);
                    string block = tbl[col..tbl.IndexOf("</DataObjects>", col, StringComparison.Ordinal)];
                    Check.True(block.Contains("<DBType xsi:type=\"xsd:short\">" + e.DbType + "</DBType>"), what + ": DBType");
                    Check.True(block.Contains("<XMType xsi:type=\"xsd:int\">" + e.XmType + "</XMType>"), what + ": XMType");
                    Check.True(block.Contains("XMHashDataDictionary&lt;" + e.Dictionary + ">"), what + ": dictionary class");
                    Check.Equal(e.Type == ColumnDataType.String, block.Contains("DictionaryFlags"), what + ": dictionary flags");
                }
            }
        }

        // The counters of Excel's H18 (two columns on Sales), H19 (one on Products) and H20 (one on each); xd.py and the file lists of
        // generated files (tools/datagen calc2, calcp, calcsp) agree with them.
        public static void CalculatedColumnShapes()
        {
            ModelReader Save(Action<ModelTable, ModelTable> add)
            {
                var (model, products, sales) = CalcModel();
                using (model)
                {
                    add(products, sales);
                    var ms = new MemoryStream();
                    model.SaveModelPart(ms);
                    return ModelReader.Read(ms.ToArray());
                }
            }
            string Find(ModelReader m, string pattern) => Encoding.UTF8.GetString(m.Find(pattern));
            string Counters(ModelReader m, string pattern) => System.Text.RegularExpressions.Regex.Matches(Find(m, pattern),
                    "<(StructureVersion|DataIDVersion|CacheVersion|DataVersion)>(-?\\d+)").Select(x => x.Groups[2].Value).Aggregate("", (a, b) => a + b + ",");

            // H18: two columns on the many-side table.
            var two = Save((p, s) =>
            {
                s.AddCalculatedColumn("Line Total", "[Qty]*[Amount]", ColumnDataType.Currency);
                s.AddCalculatedColumn("Double Qty", "[Qty]*2", ColumnDataType.Int64);
            });
            Find(two, @"Model\.27\.cub\.xml$");
            Check.Equal("2,3,3,3,", Counters(two, @"Sales_G\.9\.dim\.xml$"), "Sales dimension counters");
            Find(two, @"Products_G\.6\.dim\.xml$");
            Find(two, @"\.0\.dim\\Sales_G\.3\.tbl\.xml$");
            Check.Equal(152, two.Find(@"\.0\.dim\\2\.Sales_G\.Line Total\.0\.idf$").Length, "first column file");
            Check.Equal(152, two.Find(@"\.0\.dim\\3\.Sales_G\.Double Qty\.0\.idf$").Length, "second column file");
            Check.Equal(0, two.Files.Keys.Count(k => System.Text.RegularExpressions.Regex.IsMatch(k, @"\.prt\\info\.\d+\.xml$")), "no partition information");
            Check.True(Find(two, @"Model\.27\.cub\.xml$").Contains("<StructureVersion>21</StructureVersion>"), "cube structure");

            // H25 and H26: three and four columns on the many-side table; H27: two on it and one on the primary side.
            var three = Save((p, s) =>
            {
                s.AddCalculatedColumn("Line Total", "[Qty]*[Amount]", ColumnDataType.Currency);
                s.AddCalculatedColumn("Double Qty", "[Qty]*2", ColumnDataType.Int64);
                s.AddCalculatedColumn("Triple Qty", "[Qty]*3", ColumnDataType.Int64);
            });
            Find(three, @"Model\.31\.cub\.xml$");
            Check.Equal("3,4,4,4,", Counters(three, @"Sales_G\.12\.dim\.xml$"), "three columns: Sales dimension counters");
            Find(three, @"\.0\.dim\\Sales_G\.4\.tbl\.xml$");
            Check.Equal(152, three.Find(@"\.0\.dim\\4\.Sales_G\.Triple Qty\.0\.idf$").Length, "third column file");
            var four = Save((p, s) =>
            {
                s.AddCalculatedColumn("Line Total", "[Qty]*[Amount]", ColumnDataType.Currency);
                s.AddCalculatedColumn("Double Qty", "[Qty]*2", ColumnDataType.Int64);
                s.AddCalculatedColumn("Triple Qty", "[Qty]*3", ColumnDataType.Int64);
                s.AddCalculatedColumn("Quad Qty", "[Qty]*4", ColumnDataType.Int64);
            });
            Find(four, @"Model\.35\.cub\.xml$");
            Check.Equal("4,5,5,5,", Counters(four, @"Sales_G\.15\.dim\.xml$"), "four columns: Sales dimension counters");
            Find(four, @"\.0\.dim\\Sales_G\.5\.tbl\.xml$");
            var mixed = Save((p, s) =>
            {
                s.AddCalculatedColumn("Line Total", "[Qty]*[Amount]", ColumnDataType.Currency);
                s.AddCalculatedColumn("Double Qty", "[Qty]*2", ColumnDataType.Int64);
                p.AddCalculatedColumn("Name Length", "LEN([Product])", ColumnDataType.Int64);
            });
            Find(mixed, @"Model\.31\.cub\.xml$");
            Check.Equal("2,3,3,3,", Counters(mixed, @"Sales_G\.9\.dim\.xml$"), "mixed: Sales dimension counters");
            Check.Equal("3,2,2,3,", Counters(mixed, @"Products_G\.9\.dim\.xml$"), "mixed: Products dimension counters");

            // Beyond the samples (four columns): the same straight lines.
            var ten = Save((p, s) =>
            {
                for (int i = 0; i < 10; i++) s.AddCalculatedColumn("Calc " + i, "[Qty]*" + i, ColumnDataType.Int64);
            });
            Find(ten, @"Model\.59\.cub\.xml$");
            Check.Equal("10,11,11,11,", Counters(ten, @"Sales_G\.33\.dim\.xml$"), "ten columns: Sales dimension counters");
            Find(ten, @"\.0\.dim\\Sales_G\.11\.tbl\.xml$");
            Check.Equal(152, ten.Find(@"\.0\.dim\\11\.Sales_G\.Calc 9\.0\.idf$").Length, "tenth column file");

            // H19: one column on the hierarchy table that is the primary side; Sales is not touched.
            var primary = Save((p, s) => p.AddCalculatedColumn("Name Length", "LEN([Product])", ColumnDataType.Int64));
            Find(primary, @"Model\.23\.cub\.xml$");
            Check.Equal("3,2,2,3,", Counters(primary, @"Products_G\.9\.dim\.xml$"), "Products dimension counters");
            Find(primary, @"Sales_G\.3\.dim\.xml$");
            Find(primary, @"\.0\.dim\\Products_G\.3\.tbl\.xml$");
            Find(primary, @"\.0\.dim\\Sales_G\.0\.tbl\.xml$");
            Check.Equal(152, primary.Find(@"\.0\.dim\\3\.Products_G\.Name Length\.0\.idf$").Length, "column file");
            Find(primary, @"Products_G\.0\.dim\\info\.9\.xml$");

            // Several columns on the primary side are extrapolated from the first (no sample): dimension +3, counters and table version +1.
            var primaryTwo = Save((p, s) =>
            {
                p.AddCalculatedColumn("Name Length", "LEN([Product])", ColumnDataType.Int64);
                p.AddCalculatedColumn("Name Length x2", "LEN([Product])*2", ColumnDataType.Int64);
            });
            Find(primaryTwo, @"Model\.27\.cub\.xml$");
            Check.Equal("4,3,3,4,", Counters(primaryTwo, @"Products_G\.12\.dim\.xml$"), "two columns on Products: dimension counters");
            Find(primaryTwo, @"\.0\.dim\\Products_G\.4\.tbl\.xml$");
            Find(primaryTwo, @"\.0\.dim\\3\.Products_G\.Name Length\.0\.idf$");
            Find(primaryTwo, @"\.0\.dim\\4\.Products_G\.Name Length x2\.0\.idf$");

            // Columns on both primary-side tables (Products and Dates) are extrapolated too: each table gets its own first-column values.
            {
                var model = new DataModelWriter(new DataModelOptions { XpressCompression = false });
                var products = model.CreateTable("Products", new[] { new ColumnDefinition("ProductKey", ColumnDataType.Int64), new ColumnDefinition("Category", ColumnDataType.String), new ColumnDefinition("Product", ColumnDataType.String) });
                for (int i = 1; i <= 6; i++) products.AddRow(i, i <= 3 ? "A" : "B", "P" + i);
                var dates = model.CreateTable("Dates", new[] { new ColumnDefinition("Date", ColumnDataType.DateTime), new ColumnDefinition("Year", ColumnDataType.Int64), new ColumnDefinition("Month", ColumnDataType.String) });
                for (int i = 0; i < 40; i++) { var day = new DateTime(2025, 1, 1).AddDays(i); dates.AddRow(day, day.Year, day.ToString("MMMM", CultureInfo.InvariantCulture)); }
                var sales = model.CreateTable("Sales", new[] { new ColumnDefinition("ProductKey", ColumnDataType.Int64), new ColumnDefinition("Date", ColumnDataType.DateTime), new ColumnDefinition("Qty", ColumnDataType.Int64) });
                for (int i = 0; i < 40; i++) sales.AddRow(1 + i % 6, new DateTime(2025, 1, 1).AddDays(i), 1 + i % 5);
                model.AddRelationship(sales["ProductKey"], products["ProductKey"]);
                model.AddRelationship(sales["Date"], dates["Date"]);
                products.AddHierarchy("Product Hierarchy", "Category", "Product");
                dates.AddHierarchy("Calendar", "Year", "Month");
                products.AddCalculatedColumn("Name Length", "LEN([Product])", ColumnDataType.Int64);
                dates.AddCalculatedColumn("Year Month", "[Year]*100", ColumnDataType.Int64);
                using (model)
                {
                    var ms = new MemoryStream();
                    model.SaveModelPart(ms);
                    var bothPrimary = ModelReader.Read(ms.ToArray());
                    Find(bothPrimary, @"Model\.27\.cub\.xml$");
                    Check.Equal("3,2,2,3,", Counters(bothPrimary, @"Products_G\.9\.dim\.xml$"), "Products dimension counters");
                    Check.Equal("3,2,2,3,", Counters(bothPrimary, @"Dates_G\.9\.dim\.xml$"), "Dates dimension counters");
                    Find(bothPrimary, @"\.0\.dim\\Products_G\.3\.tbl\.xml$");
                    Find(bothPrimary, @"\.0\.dim\\Dates_G\.3\.tbl\.xml$");
                }
            }

            // A column on Products in a model with a date table has no sample: it takes Products' values without a date table.
            {
                var model = new DataModelWriter(new DataModelOptions { XpressCompression = false });
                var products = model.CreateTable("Products", new[] { new ColumnDefinition("ProductKey", ColumnDataType.Int64), new ColumnDefinition("Category", ColumnDataType.String), new ColumnDefinition("Product", ColumnDataType.String) });
                for (int i = 1; i <= 6; i++) products.AddRow(i, i <= 3 ? "A" : "B", "P" + i);
                var dates = model.CreateTable("Dates", new[] { new ColumnDefinition("Date", ColumnDataType.DateTime), new ColumnDefinition("Year", ColumnDataType.Int64), new ColumnDefinition("Month", ColumnDataType.String) });
                for (int i = 0; i < 40; i++) { var day = new DateTime(2025, 1, 1).AddDays(i); dates.AddRow(day, day.Year, day.ToString("MMMM", CultureInfo.InvariantCulture)); }
                var sales = model.CreateTable("Sales", new[] { new ColumnDefinition("ProductKey", ColumnDataType.Int64), new ColumnDefinition("Date", ColumnDataType.DateTime), new ColumnDefinition("Qty", ColumnDataType.Int64) });
                for (int i = 0; i < 40; i++) sales.AddRow(1 + i % 6, new DateTime(2025, 1, 1).AddDays(i), 1 + i % 5);
                model.AddRelationship(sales["ProductKey"], products["ProductKey"]);
                model.AddRelationship(sales["Date"], dates["Date"]);
                products.AddHierarchy("Product Hierarchy", "Category", "Product");
                dates.AddHierarchy("Calendar", "Year", "Month");
                dates.MarkAsDateTable("Date");
                products.AddCalculatedColumn("Name Length", "LEN([Product])", ColumnDataType.Int64);
                using (model)
                {
                    var ms = new MemoryStream();
                    model.SaveModelPart(ms);
                    var withDate = ModelReader.Read(ms.ToArray());
                    Find(withDate, @"Model\.26\.cub\.xml$");
                    Check.Equal("3,2,2,3,", Counters(withDate, @"Products_G\.9\.dim\.xml$"), "Products dimension counters");
                    Find(withDate, @"\.0\.dim\\Products_G\.3\.tbl\.xml$");
                    Find(withDate, @"Dates_G\.6\.dim\.xml$");
                }
            }

            // Two many-side tables with columns (no sample): each takes the many-side steps for its own count, the totals drive the rest.
            {
                var model = new DataModelWriter(new DataModelOptions { XpressCompression = false });
                var products = model.CreateTable("Products", new[] { new ColumnDefinition("ProductKey", ColumnDataType.Int64), new ColumnDefinition("Category", ColumnDataType.String), new ColumnDefinition("Product", ColumnDataType.String) });
                for (int i = 1; i <= 6; i++) products.AddRow(i, i <= 3 ? "A" : "B", "P" + i);
                var sales = model.CreateTable("Sales", new[] { new ColumnDefinition("ProductKey", ColumnDataType.Int64), new ColumnDefinition("Qty", ColumnDataType.Int64) });
                for (int i = 0; i < 40; i++) sales.AddRow(1 + i % 6, 1 + i % 5);
                var stock = model.CreateTable("Stock", new[] { new ColumnDefinition("ProductKey", ColumnDataType.Int64), new ColumnDefinition("OnHand", ColumnDataType.Int64) });
                for (int i = 0; i < 40; i++) stock.AddRow(1 + i % 6, 100 + i);
                model.AddRelationship(sales["ProductKey"], products["ProductKey"]);
                model.AddRelationship(stock["ProductKey"], products["ProductKey"]);
                products.AddHierarchy("Product Hierarchy", "Category", "Product");
                sales.AddCalculatedColumn("Twice", "[Qty]*2", ColumnDataType.Int64);
                stock.AddCalculatedColumn("Doubled", "[OnHand]*2", ColumnDataType.Int64);
                using (model)
                {
                    var ms = new MemoryStream();
                    model.SaveModelPart(ms);
                    var multi = ModelReader.Read(ms.ToArray());
                    Find(multi, @"Model\.27\.cub\.xml$");
                    Check.Equal("1,2,2,2,", Counters(multi, @"Sales_G\.6\.dim\.xml$"), "Sales dimension counters");
                    Check.Equal("1,2,2,2,", Counters(multi, @"Stock_G\.6\.dim\.xml$"), "Stock dimension counters");
                    Find(multi, @"Products_G\.6\.dim\.xml$");
                    Find(multi, @"\.0\.dim\\Stock_G\.2\.tbl\.xml$");
                }
            }

            // H20: one column on each.
            var both = Save((p, s) =>
            {
                s.AddCalculatedColumn("Line Total", "[Qty]*[Amount]", ColumnDataType.Currency);
                p.AddCalculatedColumn("Name Length", "LEN([Product])", ColumnDataType.Int64);
            });
            Find(both, @"Model\.27\.cub\.xml$");
            Check.Equal("1,2,2,2,", Counters(both, @"Sales_G\.6\.dim\.xml$"), "Sales dimension counters");
            Check.Equal("3,2,2,3,", Counters(both, @"Products_G\.9\.dim\.xml$"), "Products dimension counters");
            Find(both, @"\.0\.dim\\Sales_G\.2\.tbl\.xml$");
            Find(both, @"\.0\.dim\\Products_G\.3\.tbl\.xml$");
        }

        // Excel's H21 (one column on Sales), H22 (two), H23 (one on Products) and H24 (one on each) in a model without a hierarchy.
        public static void CalculatedColumnShapesWithoutHierarchy()
        {
            ModelReader Save(Action<ModelTable, ModelTable> add)
            {
                var (model, products, sales) = CalcModel(hierarchy: false);
                using (model)
                {
                    add(products, sales);
                    var ms = new MemoryStream();
                    model.SaveModelPart(ms);
                    return ModelReader.Read(ms.ToArray());
                }
            }
            string Find(ModelReader m, string pattern) => Encoding.UTF8.GetString(m.Find(pattern));
            string Counters(ModelReader m, string pattern) => System.Text.RegularExpressions.Regex.Matches(Find(m, pattern),
                    "<(StructureVersion|DataIDVersion|CacheVersion|DataVersion)>(-?\\d+)").Select(x => x.Groups[2].Value).Aggregate("", (a, b) => a + b + ",");

            var one = Save((p, s) => s.AddCalculatedColumn("Line Total", "[Qty]*[Amount]", ColumnDataType.Currency));
            Find(one, @"Model\.7\.cub\.xml$");
            Check.Equal("1,2,2,2,", Counters(one, @"Sales_G\.6\.dim\.xml$"), "Sales dimension counters");
            Find(one, @"Products_G\.2\.dim\.xml$");
            Find(one, @"\.0\.dim\\Sales_G\.2\.tbl\.xml$");
            Find(one, @"\.cub\\info\.7\.xml$");
            Check.Equal(0, one.Files.Keys.Count(k => System.Text.RegularExpressions.Regex.IsMatch(k, @"\.prt\\info\.\d+\.xml$")), "no partition information");
            Check.True(Find(one, @"Sales_G\.7\.det\.xml$").Contains("<Expression>[Qty]*[Amount]</Expression>"), "formula in the measure group");

            var two = Save((p, s) =>
            {
                s.AddCalculatedColumn("Line Total", "[Qty]*[Amount]", ColumnDataType.Currency);
                s.AddCalculatedColumn("Double Qty", "[Qty]*2", ColumnDataType.Int64);
            });
            Find(two, @"Model\.10\.cub\.xml$");
            Check.Equal("2,3,3,3,", Counters(two, @"Sales_G\.8\.dim\.xml$"), "Sales dimension counters");
            Find(two, @"\.0\.dim\\Sales_G\.3\.tbl\.xml$");

            // H28: three columns on Sales.
            var three = Save((p, s) =>
            {
                s.AddCalculatedColumn("Line Total", "[Qty]*[Amount]", ColumnDataType.Currency);
                s.AddCalculatedColumn("Double Qty", "[Qty]*2", ColumnDataType.Int64);
                s.AddCalculatedColumn("Triple Qty", "[Qty]*3", ColumnDataType.Int64);
            });
            Find(three, @"Model\.14\.cub\.xml$");
            Check.Equal("3,4,4,4,", Counters(three, @"Sales_G\.11\.dim\.xml$"), "three columns: Sales dimension counters");
            Find(three, @"\.0\.dim\\Sales_G\.4\.tbl\.xml$");

            var primary = Save((p, s) => p.AddCalculatedColumn("Name Length", "LEN([Product])", ColumnDataType.Int64));
            Find(primary, @"Model\.6\.cub\.xml$");
            Check.Equal("1,2,2,1,", Counters(primary, @"Products_G\.5\.dim\.xml$"), "Products dimension counters");
            Find(primary, @"\.0\.dim\\Products_G\.1\.tbl\.xml$");
            Check.True(Find(primary, @"Sales_G\.6\.det\.xml$").Contains("<Expression>LEN([Product])</Expression>"), "the many-side group holds the primary side's formula");

            var both = Save((p, s) =>
            {
                s.AddCalculatedColumn("Line Total", "[Qty]*[Amount]", ColumnDataType.Currency);
                p.AddCalculatedColumn("Name Length", "LEN([Product])", ColumnDataType.Int64);
            });
            Find(both, @"Model\.10\.cub\.xml$");
            Check.Equal("1,2,2,2,", Counters(both, @"Sales_G\.5\.dim\.xml$"), "Sales dimension counters");
            Find(both, @"Products_G\.5\.dim\.xml$");
        }

        // A table of several segments gets one block of BLANKs per segment (a model written with 16,384-row segments, 40,000 rows).
        public static void CalculatedColumnSeveralSegments()
        {
            var model = new DataModelWriter(new DataModelOptions { XpressCompression = false, SegmentRowCount = 16384 });
            var products = model.CreateTable("Products", new[] { new ColumnDefinition("ProductKey", ColumnDataType.Int64), new ColumnDefinition("Category", ColumnDataType.String), new ColumnDefinition("Product", ColumnDataType.String) });
            for (int i = 1; i <= 6; i++) products.AddRow(i, i <= 3 ? "A" : "B", "P" + i);
            var sales = model.CreateTable("Sales", new[] { new ColumnDefinition("ProductKey", ColumnDataType.Int64), new ColumnDefinition("Qty", ColumnDataType.Int64) });
            for (int i = 0; i < 40000; i++) sales.AddRow(1 + i % 6, 1 + i % 5);
            model.AddRelationship(sales["ProductKey"], products["ProductKey"]);
            products.AddHierarchy("Product Hierarchy", "Category", "Product");
            sales.AddCalculatedColumn("Twice", "[Qty]*2", ColumnDataType.Int64);
            using (model)
            {
                var ms = new MemoryStream();
                model.SaveModelPart(ms);
                var m = ModelReader.Read(ms.ToArray());
                var idf = m.Find(@"\.0\.dim\\2\.Sales_G\.Twice\.0\.idf$");
                Check.Equal(3 * 152, idf.Length, "three blocks of 152 bytes");
                Check.Equal(16384, (int)BitConverter.ToUInt32(idf, 12), "first segment rows");
                Check.Equal(16384, (int)BitConverter.ToUInt32(idf, 152 + 12), "second segment rows");
                Check.Equal(40000 - 32768, (int)BitConverter.ToUInt32(idf, 304 + 12), "last segment rows");
                string tbl = Encoding.UTF8.GetString(m.Find(@"\.0\.dim\\Sales_G\.3\.tbl\.xml$"));
                int col = tbl.IndexOf("name=\"Twice\"", StringComparison.Ordinal);
                string block = tbl[col..tbl.IndexOf("</DataObjects>", col, StringComparison.Ordinal)];
                Check.Equal(3, System.Text.RegularExpressions.Regex.Matches(block, "<Name>SubSegment</Name>").Count, "three segments in the table XML");
                Check.True(block.Contains("<SegmentCount xsi:type=\"xsd:int\">3</SegmentCount>"), "segment count of the data file");
            }
        }

        // A run of 64 or more BLANKs is a run-length entry (1,000 rows in H10 to H21), fewer are bit-packed (12 rows in H23, H24).
        public static void CalculatedColumnBlankEncoding()
        {
            foreach (var (rows, rle) in new[] { (12, false), (63, false), (64, true), (1000, true) })
            {
                var (model, _, sales) = CalcModel(salesRows: rows);
                using (model)
                {
                    sales.AddCalculatedColumn("Calc", "1", ColumnDataType.Currency);
                    var ms = new MemoryStream();
                    model.SaveModelPart(ms);
                    var m = ModelReader.Read(ms.ToArray());
                    var idf = m.Find(@"\.0\.dim\\2\.Sales_G\.Calc\.0\.idf$");
                    Check.Equal(152, idf.Length, rows + " rows: file size");
                    Check.Equal(rle ? 2u : 0xFFFFFFFFu, BitConverter.ToUInt32(idf, 8), rows + " rows: entry");
                    Check.Equal(rows, (int)BitConverter.ToUInt32(idf, 12), rows + " rows: count");
                    string tbl = Encoding.UTF8.GetString(m.Find(@"\.0\.dim\\Sales_G\.3\.tbl\.xml$"));
                    int col = tbl.IndexOf("name=\"Calc\"", StringComparison.Ordinal);
                    string block = tbl[col..tbl.IndexOf("</DataObjects>", col, StringComparison.Ordinal)];
                    string runs = rle ? "<RLERuns xsi:type=\"xsd:long\">1</RLERuns><OthersRLERuns xsi:type=\"xsd:long\">0</OthersRLERuns>" : "<RLERuns xsi:type=\"xsd:long\">0</RLERuns><OthersRLERuns xsi:type=\"xsd:long\">1</OthersRLERuns>";
                    Check.True(block.Replace("\r", "").Replace("\n", "").Replace(" ", "").Contains(runs.Replace(" ", "")), rows + " rows: run counts");
                }
            }
        }

        // A date table keeps its own offsets: a column on Sales (H29, H30) or on the date table (H31, H32).
        public static void CalculatedColumnsWithDateTable()
        {
            ModelReader Save(Action<ModelTable, ModelTable> add)
            {
                var (model, dates, sales) = DateModel(d =>
                {
                    for (int i = 0; i < 70; i++) { var day = new DateTime(2025, 1, 1).AddDays(i); d.AddRow(day, day.Year, day.ToString("MMMM", CultureInfo.InvariantCulture)); }
                });
                using (model)
                {
                    dates.MarkAsDateTable("Date");
                    add(dates, sales);
                    var ms = new MemoryStream();
                    model.SaveModelPart(ms);
                    return ModelReader.Read(ms.ToArray());
                }
            }
            string Find(ModelReader m, string pattern) => Encoding.UTF8.GetString(m.Find(pattern));
            string Counters(ModelReader m, string pattern) => System.Text.RegularExpressions.Regex.Matches(Find(m, pattern),
                    "<(StructureVersion|DataIDVersion|CacheVersion|DataVersion)>(-?\\d+)").Select(x => x.Groups[2].Value).Aggregate("", (a, b) => a + b + ",");

            var onSales = Save((d, s) => s.AddCalculatedColumn("Twice", "[N]*2", ColumnDataType.Int64));
            Find(onSales, @"Model\.26\.cub\.xml$");
            Check.Equal("1,2,2,3,", Counters(onSales, @"Sales_G\.7\.dim\.xml$"), "Sales dimension counters");
            Find(onSales, @"\.0\.dim\\Sales_G\.3\.tbl\.xml$");
            Check.Equal(152, onSales.Find(@"\.0\.dim\\3\.Sales_G\.Twice\.0\.idf$").Length, "column file");

            var twoOnSales = Save((d, s) =>
            {
                s.AddCalculatedColumn("Twice", "[N]*2", ColumnDataType.Int64);
                s.AddCalculatedColumn("Thrice", "[N]*3", ColumnDataType.Int64);
            });
            Find(twoOnSales, @"Model\.30\.cub\.xml$");
            Check.Equal("2,3,3,4,", Counters(twoOnSales, @"Sales_G\.10\.dim\.xml$"), "Sales dimension counters, two columns");
            Find(twoOnSales, @"\.0\.dim\\Sales_G\.4\.tbl\.xml$");

            var onDates = Save((d, s) => d.AddCalculatedColumn("Year Month", "[Year]*100", ColumnDataType.Int64));
            Find(onDates, @"Model\.26\.cub\.xml$");
            Check.Equal("4,3,3,4,", Counters(onDates, @"Dates_G\.9\.dim\.xml$"), "Dates dimension counters");
            Find(onDates, @"\.0\.dim\\Dates_G\.4\.tbl\.xml$");
            Check.Equal(152, onDates.Find(@"\.0\.dim\\4\.Dates_G\.Year Month\.0\.idf$").Length, "column file");
            Find(onDates, @"Dates_G\.0\.dim\\info\.9\.xml$");

            // H31B: a hierarchy over the calculated column stays unprocessed (no U$ table, StructureType Unknown, level starts -1).
            ModelHierarchy? over = null;
            var hierarchyOver = Save((d, s) =>
            {
                var yearMonth = d.AddCalculatedColumn("Year Month", "[Year]*100", ColumnDataType.Int64);
                over = d.AddHierarchy("Hierarchy1", "Year", "year month");
                Check.True(over.IsUnprocessed, "unprocessed");
                Check.Equal("Hierarchy1 (Year > Year Month)", over.ToString(), "level names");
                Check.Equal(1, over.Levels.Count, "only the data column is a ModelColumn level");
                Check.Equal("Year Month", yearMonth.Name, "column name");
            });
            Find(hierarchyOver, @"Model\.33\.cub\.xml$");
            Check.Equal("6,3,3,6,", Counters(hierarchyOver, @"Dates_G\.12\.dim\.xml$"), "Dates dimension counters");
            Find(hierarchyOver, @"\.0\.dim\\Dates_G\.6\.tbl\.xml$");
            Check.Equal(0, hierarchyOver.Files.Keys.Count(k => k.Contains("Hierarchy1")), "no hierarchy table or files");
            Check.Equal(0, hierarchyOver.Files.Keys.Count(k => System.Text.RegularExpressions.Regex.IsMatch(k, @"\.cub\\info\.\d+\.xml$")), "no cube information file");
            string overDim = Find(hierarchyOver, @"Dates_G\.12\.dim\.xml$");
            Check.True(overDim.Contains("<ddl300:ProcessingState>Unprocessed</ddl300:ProcessingState><ddl300:StructureType>Unknown</ddl300:StructureType>"), "hierarchy states");
            string overTbl = Find(hierarchyOver, @"\.0\.dim\\Dates_G\.6\.tbl\.xml$");
            Check.True(overTbl.Contains("<TableStore>$Year$-1$Year Month$-1$</TableStore>"), "level starts");
            Check.True(overTbl.Contains("<IsProcessed xsi:type=\"xsd:boolean\">false</IsProcessed><TableStore>$Year"), "hierarchy not processed");

            var both = Save((d, s) =>
            {
                s.AddCalculatedColumn("Twice", "[N]*2", ColumnDataType.Int64);
                d.AddCalculatedColumn("Year Month", "[Year]*100", ColumnDataType.Int64);
            });
            Find(both, @"Model\.30\.cub\.xml$");
            Find(both, @"Sales_G\.7\.dim\.xml$");
            Find(both, @"Dates_G\.9\.dim\.xml$");
        }

        // Only malformed text is rejected; DAX with comments, quotes inside text and brackets in names must pass.
        public static void DaxTextCheck()
        {
            string[] good =
            {
                "SUM(Sales[Qty])", "[Qty]*[Amount]", "RELATED(Products[Category])", "\"Qty \"&[Qty]", "\"a \"\"quoted\"\" (word\"",
                "CALCULATE(SUM('Sales Table'[Qty]), FILTER(ALL('Dates'), 'Dates'[Year] = 2025))", "'O''Rate'[Amount[x]]]",
                "IF([Qty] > 2, \"big\", \"small\") // a comment with ( [ and \"", "1 + /* ( [ \" */ 2", "-- note )\n1", "{1, 2, 3}", "[Total N] * 2 & \"<x>\"",
            };
            foreach (var text in good)
                Check.True(PivotWright.Internal.DaxText.Problem(text) is null, "well-formed: " + text);
            string[] bad = { "SUM(Sales[Qty]", "SUM Sales[Qty])", "[Qty", "Qty]", "\"open", "'open", "(1 + 2", "1 + /* open", "{1, 2", "(1}" };
            foreach (var text in bad)
                Check.True(PivotWright.Internal.DaxText.Problem(text) is not null, "malformed: " + text);
            var (model, products, sales) = CalcModel();
            using (model)
            {
                Tests.Throws<ArgumentException>(() => sales.AddCalculatedColumn("Bad", "SUM(Sales[Qty]", ColumnDataType.Int64), "unclosed parenthesis in a column", "malformed");
                Tests.Throws<ArgumentException>(() => sales.AddMeasure("Bad measure", "SUM(Sales[Qty]"), "unclosed parenthesis in a measure", "malformed");
                Check.Equal(0, sales.CalculatedColumns.Count, "failed calls change nothing");
            }
        }

        // Excel's H37 sample: how a calculated column's format is written for each kind of format; and the data columns take formats too.
        public static void ColumnFormats()
        {
            var us = new CultureInfo("en-US");
            var za = new CultureInfo("en-ZA");
            var gb = new CultureInfo("en-GB");
            var (model, products, sales) = CalcModel();
            using (model)
            {
                sales.AddCalculatedColumn("Whole", "[Qty]*2", ColumnDataType.Int64, MeasureFormat.WholeNumber);
                sales.AddCalculatedColumn("Dec3", "[Qty]/3", ColumnDataType.Double, MeasureFormat.Decimal(3));
                sales.AddCalculatedColumn("DecNoSep", "[Qty]/7", ColumnDataType.Double, MeasureFormat.Decimal(2, thousandSeparator: false));
                sales.AddCalculatedColumn("Pct", "[Qty]/10", ColumnDataType.Double, MeasureFormat.Percentage(1));
                sales.AddCalculatedColumn("CurR", "[Qty]*[Amount]", ColumnDataType.Currency, MeasureFormat.Currency(2, za));
                sales.AddCalculatedColumn("CurUsd", "[Qty]*[Amount]", ColumnDataType.Currency, MeasureFormat.Currency(2, us));
                sales.AddCalculatedColumn("CurNone", "[Qty]*[Amount]", ColumnDataType.Currency, MeasureFormat.CurrencyWithoutSymbol());
                var ms = new MemoryStream();
                model.SaveModelPart(ms);
                var read = ModelReader.Read(ms.ToArray());
                string dim = Encoding.UTF8.GetString(read.Find(@"Sales_G\.\d+\.dim\.xml$"));
                (string annotation, string formatString) Of(string column)
                {
                    int at = dim.IndexOf("<Attribute><Name>" + column + "</Name>", StringComparison.Ordinal);
                    string a = dim[at..dim.IndexOf("</Attribute>", at, StringComparison.Ordinal)];
                    var ann = System.Text.RegularExpressions.Regex.Match(a, "<Name>Format</Name>\\s*<Value>\\s*(.*?)\\s*</Value>", System.Text.RegularExpressions.RegexOptions.Singleline).Groups[1].Value;
                    var fs = System.Text.RegularExpressions.Regex.Match(a, "<ddl300_300:FormatString>(.*?)</ddl300_300:FormatString>").Groups[1].Value;
                    return (System.Text.RegularExpressions.Regex.Replace(ann, "\\s+", " "), System.Net.WebUtility.HtmlDecode(fs));
                }
                Check.Equal("#,0", Of("Whole").formatString, "whole number format string");
                Check.True(Of("Whole").annotation.Contains("Format=\"NumberWhole\" ThousandSeparator=\"True\""), "whole number annotation");
                Check.Equal("#,0.000", Of("Dec3").formatString, "decimal with separator");
                Check.True(Of("Dec3").annotation.Contains("Accuracy=\"3\" ThousandSeparator=\"True\""), "decimal annotation");
                Check.Equal("0.00", Of("DecNoSep").formatString, "decimal without separator");
                Check.True(!Of("DecNoSep").annotation.Contains("ThousandSeparator"), "no separator in the annotation");
                Check.Equal("0.0%;-0.0%;0.0%", Of("Pct").formatString, "percentage");
                Check.Equal("\"R\"#,0.00;-\"R\"#,0.00;\"R\"#,0.00", Of("CurR").formatString, "rand");
                Check.True(Of("CurR").annotation.Contains("LCID=\"7177\"") && Of("CurR").annotation.Contains("DisplayName=\"R English (South Africa)\""), "rand annotation");
                Check.Equal("\\$#,0.00;(\\$#,0.00);\\$#,0.00", Of("CurUsd").formatString, "dollar is escaped, negatives in parentheses");
                Check.Equal("\"\"#,0.00;(\"\"#,0.00);\"\"#,0.00", Of("CurNone").formatString, "no symbol");
                Check.True(Of("CurNone").annotation.Contains("LCID=\"1033\" DisplayName=\"\" Symbol=\"\" PositivePattern=\"0\" NegativePattern=\"0\""), "no symbol annotation");
            }

            // Dates follow the culture's own pattern: en-ZA gives year first, en-US month first, en-GB day first; never the library's choice.
            var dateCases = new (CultureInfo culture, string shortPattern)[] { (us, "M/d/yyyy"), (gb, "dd/MM/yyyy"), (za, "yyyy/MM/dd") };
            foreach (var (culture, expected) in dateCases)
            {
                var m = new DataModelWriter(new DataModelOptions { XpressCompression = false, Culture = culture });
                var d = m.CreateTable("Dates", new[] { new ColumnDefinition("Date", ColumnDataType.DateTime) { Format = MeasureFormat.ShortDate() }, new ColumnDefinition("Year", ColumnDataType.Int64) });
                d.AddRow(new DateTime(2025, 3, 14), 2025);
                using (m)
                {
                    var ms2 = new MemoryStream();
                    m.SaveModelPart(ms2);
                    string dimension = Encoding.UTF8.GetString(ModelReader.Read(ms2.ToArray()).Find(@"Dates_G\.\d+\.dim\.xml$"));
                    Check.True(dimension.Contains("Group=\"ShortDate\" FormatString=\"" + expected + "\""), culture.Name + ": short date " + expected);
                    Check.True(dimension.Contains("LCID=\"" + culture.LCID + "\""), culture.Name + ": locale");
                }
            }
            // A specific culture can be chosen for a column of another model.
            var other = new DataModelWriter(new DataModelOptions { XpressCompression = false, Culture = za });
            var table = other.CreateTable("T", new[] { new ColumnDefinition("D", ColumnDataType.DateTime) { Format = MeasureFormat.LongDate(us) }, new ColumnDefinition("A", ColumnDataType.Currency) { Format = MeasureFormat.Currency(0, us) } });
            table.AddRow(new DateTime(2025, 3, 14), 5m);
            using (other)
            {
                var ms3 = new MemoryStream();
                other.SaveModelPart(ms3);
                string dimension = Encoding.UTF8.GetString(ModelReader.Read(ms3.ToArray()).Find(@"T_G\.\d+\.dim\.xml$"));
                Check.True(dimension.Contains("Group=\"LongDate\" FormatString=\"dddd, MMMM d, yyyy\""), "long date of en-US in an en-ZA model");
                Check.True(dimension.Contains("<ddl300_300:FormatString>\\$#,0;(\\$#,0);\\$#,0</ddl300_300:FormatString>"), "dollars with no decimals on a data column");
            }

            // A format must fit the type.
            void Rejects(Action a, string what)
            {
                try { a(); } catch (ArgumentException) { return; }
                throw new Exception("Expected ArgumentException: " + what);
            }
            var (checkModel, _, checkSales) = CalcModel();
            using (checkModel)
            {
                Rejects(() => checkSales.AddCalculatedColumn("A", "1", ColumnDataType.String, MeasureFormat.WholeNumber), "number format on text");
                Rejects(() => checkSales.AddCalculatedColumn("B", "1", ColumnDataType.Int64, MeasureFormat.ShortDate()), "date format on a number");
                Rejects(() => checkSales.AddCalculatedColumn("C", "1", ColumnDataType.DateTime, MeasureFormat.Percentage()), "number format on a date");
                Rejects(() => checkSales.AddMeasure("M", "1", MeasureFormat.LongDate()), "long date on a measure");
                Rejects(() => new DataModelWriter().CreateTable("X", new[] { new ColumnDefinition("S", ColumnDataType.Boolean) { Format = MeasureFormat.WholeNumber } }), "number format on a TRUE/FALSE data column");
                Check.Equal(0, checkSales.CalculatedColumns.Count, "failed calls change nothing");
            }
        }

        public static void DateFormatStrings()
        {
            var us = new CultureInfo("en-US");
            var gb = new CultureInfo("en-GB");
            var za = new CultureInfo("en-ZA");
            // (format, culture, the pattern Excel gets, the Group written)
            var cases = new (string Format, CultureInfo Culture, string Pattern, string Group)[]
            {
                ("d", us, "M/d/yyyy", "ShortDate"), ("d", gb, "dd/MM/yyyy", "ShortDate"), ("d", za, "yyyy/MM/dd", "ShortDate"),
                ("D", us, "dddd, MMMM d, yyyy", "LongDate"), ("t", us, "h:mm tt", "ShortTime"), ("T", us, "h:mm:ss tt", "LongTime"),
                ("s", us, "yyyy'-'MM'-'dd'T'HH':'mm':'ss", "SortableDateTime"), ("dd MMM yyyy", us, "dd MMM yyyy", "ShortDate"),
                ("yyyy-MM-dd HH:mm", gb, "yyyy-MM-dd HH:mm", "GeneralShortDateTime"), ("g", us, "M/d/yyyy h:mm tt", "GeneralShortDateTime"),
                ("G", gb, "dd/MM/yyyy HH:mm:ss", "GeneralLongDateTime"), ("%d", us, "%d", "ShortDate"),
            };
            using var model = new DataModelWriter(new DataModelOptions { XpressCompression = false });
            var columns = cases.Select((c, i) => new ColumnDefinition("D" + i, ColumnDataType.DateTime) { Format = MeasureFormat.Date(c.Format, c.Culture) }).ToArray();
            var table = model.CreateTable("T", columns);
            table.AddRow(columns.Select(_ => (object?)new DateTime(2025, 3, 14, 13, 30, 55)).ToArray());
            var ms = new MemoryStream();
            model.SaveModelPart(ms);
            string dim = Encoding.UTF8.GetString(ModelReader.Read(ms.ToArray()).Find(@"T_G\.\d+\.dim\.xml$"));
            for (int i = 0; i < cases.Length; i++)
            {
                int at = dim.IndexOf("<Attribute><Name>D" + i + "</Name>", StringComparison.Ordinal);
                string a = dim[at..dim.IndexOf("</Attribute>", at, StringComparison.Ordinal)];
                var c = cases[i];
                Check.True(a.Contains("Group=\"" + c.Group + "\""), c.Format + " " + c.Culture.Name + ": group " + c.Group);
                Check.True(a.Contains("FormatString=\"" + System.Net.WebUtility.HtmlEncode(c.Pattern).Replace("&#39;", "&apos;") + "\"") || a.Contains("FormatString=\"" + c.Pattern + "\""), c.Format + " " + c.Culture.Name + ": pattern " + c.Pattern);
                Check.True(a.Contains("LCID=\"" + c.Culture.LCID + "\""), c.Format + " " + c.Culture.Name + ": locale");
            }
            // Engine strings as Excel wrote them in H38
            Check.Equal("M/d/yyyy h:mm:ss AM/PM", MeasureFormat.ToEngineDatePattern("M/d/yyyy h:mm:ss tt"), "tt is AM/PM");
            Check.Equal("yyyy\"-\"MM\"-\"dd\"T\"HH\":\"mm\":\"ss", MeasureFormat.ToEngineDatePattern("yyyy'-'MM'-'dd'T'HH':'mm':'ss"), "literals in double quotes");
            Check.Equal("dd-MMM-yy h:mm AM/PM", MeasureFormat.ToEngineDatePattern("dd-MMM-yy h:mm tt"), "short date and time");
            Tests.Throws<ArgumentException>(() => MeasureFormat.Date("o"), "round-trip format has a time zone");
            Tests.Throws<ArgumentException>(() => MeasureFormat.Date("yyyy-MM-dd zzz"), "time zone offset");
            Tests.Throws<ArgumentException>(() => MeasureFormat.Date(""), "empty format");
            Tests.Throws<ArgumentException>(() => MeasureFormat.Date("x"), "a letter that is no standard specifier");
            Tests.Throws<ArgumentException>(() => MeasureFormat.Date("dd 'open"), "an unclosed literal");
            Tests.Throws<ArgumentNullException>(() => MeasureFormat.Date(null!), "null format");
            using var other = new DataModelWriter();
            var t2 = other.CreateTable("U", new[] { new ColumnDefinition("N", ColumnDataType.Int64) });
            Tests.Throws<ArgumentException>(() => t2.AddCalculatedColumn("C", "1", ColumnDataType.Int64, MeasureFormat.Date("d")), "a date format on a number column");
            Tests.Throws<ArgumentException>(() => t2.AddMeasure("M", "1", MeasureFormat.Date("d")), "a custom date format on a measure");
        }

        public static void CalculatedColumnValidation()
        {
            void Rejects<T>(Action a, string what) where T : Exception
            {
                try { a(); } catch (T) { return; }
                throw new Exception("Expected " + typeof(T).Name + ": " + what);
            }
            void Saves(DataModelWriter model) { using var ms = new MemoryStream(); model.SaveModelPart(ms); }

            var (model, products, sales) = CalcModel();
            using (model)
            {
                sales.AddMeasure("Total Qty", "SUM(Sales[Qty])");
                Rejects<ArgumentException>(() => sales.AddCalculatedColumn("", "1", ColumnDataType.Currency), "empty name");
                Rejects<ArgumentException>(() => sales.AddCalculatedColumn("A[B]", "1", ColumnDataType.Currency), "brackets");
                Rejects<ArgumentException>(() => sales.AddCalculatedColumn("Qty", "1", ColumnDataType.Currency), "name of a column");
                Rejects<ArgumentException>(() => sales.AddCalculatedColumn("total qty", "1", ColumnDataType.Currency), "name of a measure");
                Rejects<ArgumentException>(() => sales.AddCalculatedColumn("X", "  ", ColumnDataType.Currency), "empty expression");
                Rejects<ArgumentException>(() => sales.AddCalculatedColumn("X", "=", ColumnDataType.Currency), "only =");
                Rejects<ArgumentException>(() => sales.AddCalculatedColumn("X", "1", (ColumnDataType)99), "not a data type");
                Check.Equal(0, sales.CalculatedColumns.Count, "failed calls change nothing");
                sales.AddCalculatedColumn("Line Total", "[Qty]*[Amount]", ColumnDataType.Currency);
                Rejects<ArgumentException>(() => sales.AddCalculatedColumn("line total", "1", ColumnDataType.Currency), "duplicate calculated column");
                Rejects<ArgumentException>(() => sales.AddMeasure("Line Total", "1"), "measure named like a calculated column");
                Rejects<ArgumentException>(() => products.AddHierarchy("Product Hierarchy", "Category", "Product"), "duplicate hierarchy");
                Saves(model);
                // A column on the primary side as well is a shape Excel was seen to write (H20, H27); so are many on the many side (H26).
                products.AddCalculatedColumn("Other", "1", ColumnDataType.Int64);
                Saves(model);
                for (int i = 0; i < 12; i++) sales.AddCalculatedColumn("More " + i, "2", ColumnDataType.Int64);
                Saves(model);
            }
            // A model without a hierarchy is supported (H21 to H24), but not a table without any relationship.
            using (var noHierarchy = new DataModelWriter(new DataModelOptions { XpressCompression = false }))
            {
                var a = noHierarchy.CreateTable("A", new[] { new ColumnDefinition("K", ColumnDataType.Int64), new ColumnDefinition("V", ColumnDataType.Currency) });
                a.AddRow(1, 2m);
                var b = noHierarchy.CreateTable("B", new[] { new ColumnDefinition("K", ColumnDataType.Int64) });
                b.AddRow(1);
                noHierarchy.AddRelationship(b["K"], a["K"]);
                b.AddCalculatedColumn("C", "1", ColumnDataType.Currency);
                Saves(noHierarchy);
                var c = noHierarchy.CreateTable("C", new[] { new ColumnDefinition("K", ColumnDataType.Int64) });
                c.AddRow(1);
                c.AddCalculatedColumn("D", "1", ColumnDataType.Currency);
                Rejects<InvalidOperationException>(() => Saves(noHierarchy), "table without a relationship in a model without hierarchy");
            }
            var (lonely, _, _) = CalcModel();
            using (lonely)
            {
                var other = lonely.CreateTable("Lonely", new[] { new ColumnDefinition("K", ColumnDataType.Int64) });
                other.AddRow(1);
                other.AddCalculatedColumn("C", "1", ColumnDataType.Int64);
                Rejects<InvalidOperationException>(() => Saves(lonely), "table without hierarchy that is not on the many side");
            }
        }

        // H33 (hide Sales[Qty], then add a calculated column), H36 (hide a hierarchy level) and H35 (hide the calculated column itself).
        public static void CalculatedColumnsWithHiddenColumns()
        {
            ModelReader Save(Action<ModelTable, ModelTable> add)
            {
                var (model, products, sales) = CalcModel();
                using (model)
                {
                    add(products, sales);
                    var ms = new MemoryStream();
                    model.SaveModelPart(ms);
                    return ModelReader.Read(ms.ToArray());
                }
            }
            string Find(ModelReader m, string pattern) => Encoding.UTF8.GetString(m.Find(pattern));
            string Counters(ModelReader m, string pattern) => System.Text.RegularExpressions.Regex.Matches(Find(m, pattern),
                    "<(StructureVersion|DataIDVersion|CacheVersion|DataVersion)>(-?\\d+)").Select(x => x.Groups[2].Value).Aggregate("", (a, b) => a + b + ",");

            var salesHidden = Save((p, s) =>
            {
                s["Qty"].IsHidden = true;
                s.AddCalculatedColumn("Line Total", "[Qty]*[Amount]", ColumnDataType.Currency);
            });
            Find(salesHidden, @"Model\.25\.cub\.xml$");
            Find(salesHidden, @"Sales_G\.24\.det\.xml$");
            Check.Equal("1,2,2,2,", Counters(salesHidden, @"Sales_G\.7\.dim\.xml$"), "Sales dimension counters");
            Find(salesHidden, @"Products_G\.6\.dim\.xml$");
            Find(salesHidden, @"\.0\.dim\\Sales_G\.2\.tbl\.xml$");

            var levelHidden = Save((p, s) =>
            {
                p["Product"].IsHidden = true;
                s.AddCalculatedColumn("Line Total", "[Qty]*[Amount]", ColumnDataType.Currency);
            });
            Find(levelHidden, @"Model\.25\.cub\.xml$");
            Find(levelHidden, @"Products_G\.7\.dim\.xml$");
            Find(levelHidden, @"Sales_G\.6\.dim\.xml$");

            // No sample without hierarchies: the same steps on that model's fit (cube 6 + 2, measure groups 7, Sales dimension 5 + 1).
            {
                var (plainModel, _, plainSales) = CalcModel(hierarchy: false);
                using (plainModel)
                {
                    plainSales["Qty"].IsHidden = true;
                    plainSales.AddCalculatedColumn("Line Total", "[Qty]*[Amount]", ColumnDataType.Currency);
                    var ms = new MemoryStream();
                    plainModel.SaveModelPart(ms);
                    var plainRead = ModelReader.Read(ms.ToArray());
                    Find(plainRead, @"Model\.8\.cub\.xml$");
                    Find(plainRead, @"Sales_G\.7\.det\.xml$");
                    Find(plainRead, @"Sales_G\.6\.dim\.xml$");
                }
            }

            var calcHidden = Save((p, s) =>
            {
                var total = s.AddCalculatedColumn("Line Total", "[Qty]*[Amount]", ColumnDataType.Currency);
                total.IsHidden = true;
                Check.True(total.IsHidden, "IsHidden reads back");
            });
            string dim = Find(calcHidden, @"Sales_G\.7\.dim\.xml$");
            int at = dim.IndexOf("<Attribute><Name>Line Total</Name>", StringComparison.Ordinal);
            string attribute = dim[at..dim.IndexOf("</Attribute>", at, StringComparison.Ordinal)];
            Check.True(attribute.Contains("<AttributeHierarchyVisible>false</AttributeHierarchyVisible>"), "hidden attribute in the dimension");
            string cube = Find(calcHidden, @"Model\.25\.cub\.xml$");
            int cubeAt = cube.IndexOf("<AttributeID>Line Total</AttributeID>", StringComparison.Ordinal);
            Check.True(cube[cubeAt..].Contains("<AttributeHierarchyVisible>false</AttributeHierarchyVisible>") &&
                       cube[cubeAt..cube.IndexOf("</Attribute>", cubeAt, StringComparison.Ordinal)].Contains("false"), "hidden attribute in the cube");
        }

        public static void DistinctOrdinals()
        {
            ModelReader Save(bool hierarchy)
            {
                using var model = new DataModelWriter(new DataModelOptions { XpressCompression = false });
                var a = model.CreateTable("A", new[] { new ColumnDefinition("K", ColumnDataType.Int64), new ColumnDefinition("Name", ColumnDataType.String) });
                a.AddRow(1, "x");
                var b = model.CreateTable("B", new[] { new ColumnDefinition("K", ColumnDataType.Int64), new ColumnDefinition("Qty", ColumnDataType.Int64) });
                b.AddRow(1, 2);
                var c = model.CreateTable("C", new[] { new ColumnDefinition("K", ColumnDataType.Int64), new ColumnDefinition("Name", ColumnDataType.String) });
                c.AddRow(1, "y");
                model.AddRelationship(model.Tables["B"]["K"], model.Tables["A"]["K"]);
                if (hierarchy) model.Tables["A"].AddHierarchy("H", "K", "Name");
                var ms = new MemoryStream();
                model.SaveModelPart(ms);
                return ModelReader.Read(ms.ToArray());
            }
            string[] Ordinals(ModelReader m, string pattern) =>
                m.Files.Where(kv => System.Text.RegularExpressions.Regex.IsMatch(kv.Key, pattern))
                    .Select(kv => System.Text.RegularExpressions.Regex.Match(Encoding.UTF8.GetString(kv.Value), "<Ordinal>(\\d+)</Ordinal>").Groups[1].Value)
                    .OrderBy(x => x, StringComparer.Ordinal).ToArray();
            var with = Save(hierarchy: true);
            Check.Equal("1,2,3", string.Join(",", Ordinals(with, @"\.det\.xml$")), "measure group Ordinals are distinct");
            Check.Equal("1,1,1", string.Join(",", Ordinals(with, @"\.prt\.xml$")), "partition Ordinals");
            Check.Equal("1", string.Join(",", Ordinals(with, @"\.scr\.xml$")), "script Ordinal");
            // Models without a hierarchy get them too: a hierarchy can be added in Excel later.
            var without = Save(hierarchy: false);
            Check.Equal("1,2,3", string.Join(",", Ordinals(without, @"\.det\.xml$")), "no hierarchy: distinct Ordinals as well");
        }

        public static void DateTableFlags()
        {
            (DateOnly, DateOnly) C(DateOnly f, DateOnly t, int ye) => DataModelWriter.CompleteYears(f, t, ye);
            Check.Equal((new DateOnly(2025, 1, 1), new DateOnly(2025, 12, 31)), C(new DateOnly(2025, 3, 15), new DateOnly(2025, 9, 10), 12), "calendar year");
            Check.Equal((new DateOnly(2023, 7, 1), new DateOnly(2026, 6, 30)), C(new DateOnly(2024, 2, 1), new DateOnly(2025, 9, 10), 6), "June year end");
            Check.Equal((new DateOnly(2024, 3, 1), new DateOnly(2026, 2, 28)), C(new DateOnly(2025, 1, 15), new DateOnly(2025, 12, 31), 2), "February year end");
            Check.Equal((new DateOnly(2022, 3, 1), new DateOnly(2024, 2, 29)), C(new DateOnly(2022, 3, 1), new DateOnly(2023, 6, 1), 2), "ends on 29 February of a leap year");
            Check.Equal((new DateOnly(2024, 2, 1), new DateOnly(2025, 1, 31)), C(new DateOnly(2024, 2, 1), new DateOnly(2025, 1, 31), 1), "January year end, already whole");

            // completeYears
            Check.Equal(365L, new DataModelWriter().AddDateTable("A", new DateOnly(2025, 3, 15), new DateOnly(2025, 3, 16)).RowCount, "extended to 2025 by default");
            Check.Equal(2L, new DataModelWriter().AddDateTable("B", new DateOnly(2025, 3, 15), new DateOnly(2025, 3, 16), new DateTableOptions { CompleteYears = false }).RowCount, "exact range when off");
            Check.Equal(366L, new DataModelWriter().AddDateTable("C", new DateOnly(2024, 2, 1), new DateOnly(2024, 2, 2), new DateTableOptions { YearEndMonth = 6 }).RowCount, "July 2023 to June 2024 (leap)");

            // includeYearMonth and hideSortColumns: each flag on its own
            string[] Names(ModelTable t, bool hidden) => t.Columns.Where(c => c.IsHidden == hidden).Select(c => c.Name).ToArray();
            var range = (new DateOnly(2024, 1, 1), new DateOnly(2025, 12, 31));
            var all = new DataModelWriter().AddDateTable("D", range.Item1, range.Item2, new DateTableOptions { YearEndMonth = 6 });
            Check.Equal("QuarterNumber,MonthNumber,YearMonthNumber,DayOfWeekNumber,FiscalQuarterNumber,FiscalMonthNumber", string.Join(",", Names(all, true)), "defaults hide the sort keys");
            Check.True(all.Columns.Any(c => c.Name == "YearMonth"), "YearMonth by default");
            var noHide = new DataModelWriter().AddDateTable("D", range.Item1, range.Item2, new DateTableOptions { YearEndMonth = 6, HideSortColumns = false });
            Check.Equal(0, noHide.Columns.Count(c => c.IsHidden), "nothing hidden when off");
            var noYm = new DataModelWriter().AddDateTable("D", range.Item1, range.Item2, new DateTableOptions { IncludeYearMonth = false });
            Check.Equal("Date,Year,QuarterNumber,Quarter,MonthNumber,Month,Day,DayOfWeekNumber,DayOfWeek", string.Join(",", noYm.Columns.Select(c => c.Name)), "no YearMonth when off");
            Check.Equal("QuarterNumber,MonthNumber,DayOfWeekNumber", string.Join(",", Names(noYm, true)), "hidden without YearMonth");

            // The hidden attributes reach the dimension, and the model saves (every YearMonth label maps to one number across years).
            using var model = new DataModelWriter(new DataModelOptions { XpressCompression = false, Culture = CultureInfo.GetCultureInfo(1033) });
            model.AddDateTable("Dates", range.Item1, range.Item2, new DateTableOptions { YearEndMonth = 6 });
            var sales = model.CreateTable("Sales", new[] { new ColumnDefinition("Date", ColumnDataType.DateTime) });
            sales.AddRow(new DateOnly(2024, 1, 1));
            model.AddRelationship(model.Tables["Sales"]["Date"], model.Tables["Dates"]["Date"]);
            var ms = new MemoryStream();
            model.SaveModelPart(ms);
            string dim = Encoding.UTF8.GetString(ModelReader.Read(ms.ToArray()).Find(@"\.0\.db\\Dates_G\.[0-9]+\.dim\.xml$"));
            Check.Equal(1 + 6, System.Text.RegularExpressions.Regex.Matches(dim, "<AttributeHierarchyVisible>false</AttributeHierarchyVisible>").Count, "row number plus six hidden attributes");
        }

        public static void ExcludeFromRefreshAll()
        {
            string Connections(bool exclude)
            {
                using var model = new DataModelWriter(new DataModelOptions { XpressCompression = false, ExcludeSourceFromRefreshAll = exclude });
                model.CreateTable("T", new[] { new ColumnDefinition("N", ColumnDataType.Int64) }).AddRow(1L);
                var ms = new MemoryStream();
                model.Save(ms);
                using var zip = new System.IO.Compression.ZipArchive(new MemoryStream(ms.ToArray()));
                using var r = new StreamReader(zip.GetEntry("xl/connections.xml")!.Open());
                return r.ReadToEnd();
            }
            Check.True(!Connections(false).Contains("excludeFromRefreshAll"), "off by default");
            Check.True(Connections(true).Contains("<x15:connection id=\"fe61cfce-2c03-495c-87ad-e281a1607748\" excludeFromRefreshAll=\"1\"/>") || Connections(true).Contains("excludeFromRefreshAll=\"1\"/>"), "attribute as in H41");
        }

        public static void CalculationTrigger()
        {
            using var model = new DataModelWriter(new DataModelOptions { XpressCompression = false });
            var products = model.CreateTable("Products", new[] { new ColumnDefinition("Key", ColumnDataType.Int64), new ColumnDefinition("Name", ColumnDataType.String) });
            products.AddRow(1L, "A");
            var sales = model.CreateTable("Sales", new[] { new ColumnDefinition("Key", ColumnDataType.Int64), new ColumnDefinition("Qty", ColumnDataType.Int64) });
            for (int i = 0; i < 70; i++) sales.AddRow(1L, (long)i);
            model.AddRelationship(sales["Key"], products["Key"]);
            products.AddHierarchy("H", "Name", "Key");
            sales.AddCalculatedColumn("Double Qty", "[Qty]*2", ColumnDataType.Int64);
            Check.Equal(2, model.Tables.Count, "the trigger table is not in Tables");

            var ms = new MemoryStream();
            model.Save(ms);
            using var zip = new System.IO.Compression.ZipArchive(new MemoryStream(ms.ToArray()));
            string Part(string name)
            {
                var e = zip.GetEntry(name) ?? throw new Exception("Missing part " + name);
                using var r = new StreamReader(e.Open());
                return r.ReadToEnd();
            }
            foreach (var e in zip.Entries.Where(e => e.FullName.EndsWith(".xml") || e.FullName.EndsWith(".rels")))
            {
                using var r = new StreamReader(e.Open());
                try { System.Xml.Linq.XDocument.Parse(r.ReadToEnd()); } catch (Exception x) { throw new Exception("Not well-formed: " + e.FullName + ": " + x.Message); }
            }
            string workbook = Part("xl/workbook.xml"), connections = Part("xl/connections.xml");
            Check.True(workbook.Contains("<sheet name=\"Sheet2\" sheetId=\"2\" state=\"veryHidden\" r:id=\"rId5\"/>"), "second sheet");
            Check.True(workbook.Contains("<definedName name=\"_xlcn.WorksheetConnection_PivotWright.xlsxPivotWrightTrigger\" hidden=\"1\">PivotWrightTrigger[]</definedName>"), "hidden defined name");
            Check.True(workbook.Contains("<x15:modelTable id=\"PivotWrightTrigger\" name=\"PivotWrightTrigger\" connection=\"WorksheetConnection_PivotWright.xlsx!PivotWrightTrigger\"/>"), "model table of the linked connection");
            Check.True(connections.Contains("type=\"102\" refreshOnLoad=\"1\"") && connections.Contains("<x15:rangePr sourceName=\"_xlcn.WorksheetConnection_PivotWright.xlsxPivotWrightTrigger\"/>"), "worksheet connection");
            Check.True(Part("xl/worksheets/sheet1.xml").Contains("tabSelected=\"1\"") && !Part("xl/worksheets/sheet2.xml").Contains("tabSelected"), "only the first sheet is selected");
            Check.True(Part("xl/worksheets/sheet2.xml").Contains("<c r=\"A2\"><v>1</v></c>"), "the sheet holds the same value as the model");
            Check.True(Part("xl/tables/table1.xml").Contains("name=\"PivotWrightTrigger\"") && Part("xl/sharedStrings.xml").Contains("<t>Key</t>"), "worksheet table");
            Check.True(Part("[Content_Types].xml").Contains("/xl/tables/table1.xml") && Part("docProps/app.xml").Contains("<vt:lpstr>Sheet2</vt:lpstr>"), "content types and sheet list");
            string manual = Part("customXml/item1.xml");
            Check.True(manual.Contains("ManualCalcMode") && manual.Contains("[CDATA[False]]"), "automatic calculation");
            Check.True(Enumerable.Range(1, 6).Any(n => Part("customXml/item" + n + ".xml").Contains("LinkedTableUpdateMode")), "linked tables update");

            byte[] itemData;
            using (var es = zip.GetEntry("xl/model/item.data")!.Open()) { var buf = new MemoryStream(); es.CopyTo(buf); itemData = buf.ToArray(); }
            var m = ModelReader.Read(itemData);
            string dim = Encoding.UTF8.GetString(m.Find(@"PivotWrightTrigger\.2\.dim\.xml$"));
            Check.True(!dim.Contains("DataSourceViewBinding") && dim.Contains("<ID>RowNumber</ID>") && dim.Contains("<Value>A</Value>"), "linked dimension");
            string prt = Encoding.UTF8.GetString(m.Find(@"PivotWrightTrigger\.1\.prt\.xml$"));
            Check.True(!prt.Contains("QueryBinding") && !prt.Contains("ErrorConfiguration") && prt.Contains("<CacheVersion>1</CacheVersion>"), "linked partition has no query");
            m.Find(@"PivotWrightTrigger\.1\.det\.xml$"); m.Find(@"PivotWrightTrigger\.0\.dim\\0\.PivotWrightTrigger\.RowNumber\.0\.idf$"); m.Find(@"PivotWrightTrigger\.0\.dim\\PivotWrightTrigger\.0\.tbl\.xml$");
            Check.True(!Encoding.UTF8.GetString(m.Find(@"Sandbox\.0\.dsv\.xml$")).Contains("Trigger"), "the view lists only the source tables");
            // The trigger table is hidden from client tools as a whole (H46): a private measure group and an invisible cube dimension; no other table is.
            Check.True(Encoding.UTF8.GetString(m.Find(@"PivotWrightTrigger\.1\.det\.xml$")).Contains("<Name>IsTablePrivate</Name>"), "private measure group");
            Check.Equal(1, m.Files.Where(kv => kv.Key.EndsWith(".det.xml")).Count(kv => Encoding.UTF8.GetString(kv.Value).Contains("IsTablePrivate")), "only the trigger table is private");
            string cube = Encoding.UTF8.GetString(m.Find(@"Model\.\d+\.cub\.xml$"));
            Check.True(cube.Contains("<Dimension><ID>PivotWrightTrigger</ID><Name>PivotWrightTrigger</Name><DimensionID>PivotWrightTrigger</DimensionID><Visible>false</Visible>"), "invisible cube dimension");
            Check.Equal(1, System.Text.RegularExpressions.Regex.Matches(cube, "<DimensionID>[^<]*</DimensionID><Visible>false</Visible>").Count, "only one invisible dimension");
            // Every group of the log has a file named for its version (SPEC section 5).
            foreach (var g in m.Log.Element("FileGroups")!.Elements("FileGroup"))
            {
                string id = (string)g.Element("ID")!;
                int cls = int.Parse((string)g.Element("Class")!), version = int.Parse((string)g.Element("ObjectVersion")!);
                string ext = cls switch { 100006 => "dim", 100016 => "det", 100021 => "prt", 100010 => "cub", _ => "" };
                if (ext == "") continue;
                var paths = g.Descendants("Path").Select(p => (string)p).ToList();
                Check.True(paths.Any(p => p.EndsWith("\\" + id + "." + version + "." + ext + ".xml", StringComparison.Ordinal)),
                    $"group {id} has a file for version {version}");
            }


            // Only with calculated columns, and not when switched off.
            var ms2 = new MemoryStream();
            using (var off = new DataModelWriter(new DataModelOptions { XpressCompression = false, CalculateOnOpen = false }))
            {
                var t = off.CreateTable("T", new[] { new ColumnDefinition("N", ColumnDataType.Int64) }); t.AddRow(1L);
                off.Save(ms2);
            }
            using (var plain = new System.IO.Compression.ZipArchive(new MemoryStream(ms2.ToArray())))
                Check.True(plain.GetEntry("xl/worksheets/sheet2.xml") is null, "no calculated column, no trigger");
            using var noTrigger = new DataModelWriter(new DataModelOptions { XpressCompression = false, CalculateOnOpen = false });
            var nt = noTrigger.CreateTable("Products", new[] { new ColumnDefinition("Key", ColumnDataType.Int64) }); nt.AddRow(1L);
            var ns = noTrigger.CreateTable("Sales", new[] { new ColumnDefinition("Key", ColumnDataType.Int64), new ColumnDefinition("Qty", ColumnDataType.Int64) }); ns.AddRow(1L, 2L);
            noTrigger.AddRelationship(ns["Key"], nt["Key"]);
            ns.AddCalculatedColumn("Q2", "[Qty]*2", ColumnDataType.Int64);
            var ms3 = new MemoryStream(); noTrigger.Save(ms3);
            using (var z3 = new System.IO.Compression.ZipArchive(new MemoryStream(ms3.ToArray())))
                Check.True(z3.GetEntry("xl/worksheets/sheet2.xml") is null, "switched off");
            // SaveModelPart writes the tables only.
            var part = new MemoryStream(); model.SaveModelPart(part);
            Check.True(!ModelReader.Read(part.ToArray()).Files.Keys.Any(k2 => k2.Contains("PivotWrightTrigger")), "SaveModelPart has no trigger");
            // A table of the same name is reported.
            using var clash = new DataModelWriter(new DataModelOptions { XpressCompression = false });
            var ct = clash.CreateTable("PivotWrightTrigger", new[] { new ColumnDefinition("Key", ColumnDataType.Int64) }); ct.AddRow(1L);
            var cs = clash.CreateTable("Sales", new[] { new ColumnDefinition("Key", ColumnDataType.Int64), new ColumnDefinition("Qty", ColumnDataType.Int64) }); cs.AddRow(1L, 2L);
            clash.AddRelationship(cs["Key"], ct["Key"]);
            cs.AddCalculatedColumn("Q2", "[Qty]*2", ColumnDataType.Int64);
            Tests.Throws<InvalidOperationException>(() => clash.Save(new MemoryStream()), "a table with the trigger's name");
        }

        public static void CalculatedColumnSortBy()
        {
            using var model = new DataModelWriter(new DataModelOptions { XpressCompression = false });
            var dim0 = model.CreateTable("Products", new[] { new ColumnDefinition("Key", ColumnDataType.Int64) });
            dim0.AddRow(1L);
            var t = model.CreateTable("Sales", new[] { new ColumnDefinition("Key", ColumnDataType.Int64), new ColumnDefinition("Amount", ColumnDataType.Currency) });
            for (int i = 0; i < 100; i++) t.AddRow(1L, (decimal)i);
            model.AddRelationship(t["Key"], dim0["Key"]);
            var band = t.AddCalculatedColumn("Band", "IF([Amount]>30,\"High\",\"Low\")", ColumnDataType.String);
            var order = t.AddCalculatedColumn("Band Order", "IF([Amount]>30,2,1)", ColumnDataType.Int64);
            band.SortByColumn = order;
            Check.True(band.SortByColumn == order, "getter returns the sort column");
            Tests.Throws<ArgumentException>(() => band.SortByColumn = band, "itself");
            Tests.Throws<ArgumentException>(() => order.SortByColumn = band, "sort column that is itself sorted");
            var third = t.AddCalculatedColumn("Third", "1", ColumnDataType.Int64);
            Tests.Throws<ArgumentException>(() => third.SortByColumn = band, "a sorted column as sort column");
            Tests.Throws<ArgumentException>(() => order.SortByColumn = third, "a sort column cannot be sorted");
            var ms = new MemoryStream();
            model.SaveModelPart(ms);
            var files = ModelReader.Read(ms.ToArray());
            string dim = Encoding.UTF8.GetString(files.Find(@"Sales_G\.\d+\.dim\.xml$"));
            int at = dim.IndexOf("<Attribute><Name>Band</Name>", StringComparison.Ordinal);
            string a = dim[at..dim.IndexOf("</Attribute>", at, StringComparison.Ordinal)];
            Check.True(a.Contains("<OrderBy>AttributeKey</OrderBy>") && a.Contains("<OrderByAttributeID>" + order.Definition.Id + "</OrderByAttributeID>"), "attribute sorted by the other column: " + a[..300]);
            Check.True(a.Contains("<ddl200:ProcessingState>Unprocessed</ddl200:ProcessingState>"), "still unprocessed");
            Check.True(Encoding.UTF8.GetString(files.Find(@"Sales_G\.\d+\.tbl\.xml$")).Contains("<OrderByColumn>" + order.Definition.Id + "</OrderByColumn>"), "tbl.xml names the sort column");
            band.SortByColumn = null;
            Check.True(band.SortByColumn is null, "cleared");
        }

        public static void DateTableDateFormat()
        {
            using var model = new DataModelWriter(new DataModelOptions { XpressCompression = false });
            model.AddDateTable("Dates", new DateOnly(2025, 1, 1), new DateOnly(2025, 1, 2), new DateTableOptions { DateFormat = MeasureFormat.Date("dd MMM yyyy") });
            var sales = model.CreateTable("Sales", new[] { new ColumnDefinition("Date", ColumnDataType.DateTime) });
            sales.AddRow(new DateOnly(2025, 1, 1));
            model.AddRelationship(model.Tables["Sales"]["Date"], model.Tables["Dates"]["Date"]);
            var ms = new MemoryStream();
            model.SaveModelPart(ms);
            string dim = Encoding.UTF8.GetString(ModelReader.Read(ms.ToArray()).Find(@"Dates_G\.\d+\.dim\.xml$"));
            int at = dim.IndexOf("<Attribute><Name>Date</Name>", StringComparison.Ordinal);
            string a = dim[at..dim.IndexOf("</Attribute>", at, StringComparison.Ordinal)];
            Check.True(a.Contains("FormatString=\"dd MMM yyyy\""), "the Date column carries the format");
            Check.True(!dim[(at + a.Length)..].Contains("dd MMM yyyy"), "no other column does");
            Tests.Throws<ArgumentException>(() => new DateTableOptions { DateFormat = MeasureFormat.WholeNumber }, "a number format");
            Check.True(new DateTableOptions().DateFormat is null, "default is Excel's general date");
        }

        public static void DateTableYearLabel()
        {
            using var model = new DataModelWriter(new DataModelOptions { XpressCompression = false, Culture = CultureInfo.GetCultureInfo(1033) });
            var dates = model.AddDateTable("Dates", new DateOnly(2024, 1, 1), new DateOnly(2025, 12, 31), new DateTableOptions { YearEndMonth = 6, YearLabel = "Financial" });
            Check.Equal("Date,Year,QuarterNumber,Quarter,MonthNumber,Month,YearMonthNumber,YearMonth,Day,DayOfWeekNumber,DayOfWeek,FinancialYear,FinancialQuarterNumber,FinancialQuarter,FinancialMonthNumber,FinancialMonth",
                string.Join(",", dates.Columns.Select(c => c.Name)), "columns use the label");
            Check.True(dates["FinancialQuarter"].SortByColumn?.Name == "FinancialQuarterNumber", "sort-by follows the label");
            var sales = model.CreateTable("Sales", new[] { new ColumnDefinition("Date", ColumnDataType.DateTime) });
            sales.AddRow(new DateOnly(2024, 1, 1));
            model.AddRelationship(model.Tables["Sales"]["Date"], model.Tables["Dates"]["Date"]);
            var ms = new MemoryStream();
            model.SaveModelPart(ms);
            string dim = Encoding.UTF8.GetString(ModelReader.Read(ms.ToArray()).Find(@"\.0\.db\\Dates_G\.[0-9]+\.dim\.xml$"));
            Check.True(dim.Contains("<Hierarchy><Name>Financial Calendar</Name>") && !dim.Contains("Fiscal"), "hierarchy named after the label, no Fiscal left");

            // The default label is unchanged, and an invalid label is rejected before anything is added.
            using var def = new DataModelWriter(new DataModelOptions { XpressCompression = false });
            Check.True(def.AddDateTable("D", new DateOnly(2025, 1, 1), new DateOnly(2025, 1, 2), new DateTableOptions { YearEndMonth = 6 }).Columns.Any(c => c.Name == "FiscalYear"), "default label Fiscal");
            foreach (var label in new[] { "", " ", "1Fiscal", "Fin Year", "Fin-Year", "Year.End" })
            {
                try { _ = new DateTableOptions { YearLabel = label }; throw new Exception("Expected ArgumentException for label '" + label + "'"); }
                catch (ArgumentException) { }
            }
            Check.Equal("Fiscal", new DateTableOptions().YearLabel, "default label");
            Check.Equal(12, new DateTableOptions().YearEndMonth, "default year end");
            var defaults = new DateTableOptions();
            Check.True(defaults.CompleteYears && defaults.IncludeYearMonth && defaults.HideSortColumns, "features on by default");
        }

        public static void DateTableValidation()
        {
            void Rejects<T>(Action a, string what) where T : Exception
            {
                try { a(); } catch (T) { return; }
                throw new Exception("Expected " + typeof(T).Name + ": " + what);
            }
            void Saves(DataModelWriter model) { using var ms = new MemoryStream(); model.SaveModelPart(ms); }

            var (model, _, _) = DateModel();
            using (model)
            {
                Rejects<KeyNotFoundException>(() => model.Tables["Nope"].MarkAsDateTable("Date"), "unknown table");
                Rejects<ArgumentException>(() => model.Tables["Dates"].MarkAsDateTable("Nope"), "unknown column");
                Rejects<ArgumentException>(() => model.Tables["Dates"].MarkAsDateTable("Year"), "not a date column");
                Check.True(model.DateTable?.Name is null, "nothing marked yet");
                model.Tables["Dates"].MarkAsDateTable("Date");
                Rejects<InvalidOperationException>(() => model.Tables["Sales"].MarkAsDateTable("Date"), "second date table");
                Saves(model);
            }

            // Hierarchy and relationship are required: no Excel sample of the other cases exists.
            using (var noHierarchy = new DataModelWriter(new DataModelOptions { XpressCompression = false }))
            {
                var d = noHierarchy.CreateTable("Dates", new[] { new ColumnDefinition("Date", ColumnDataType.DateTime) });
                d.AddRow(new DateTime(2025, 1, 1));
                var s = noHierarchy.CreateTable("Sales", new[] { new ColumnDefinition("Date", ColumnDataType.DateTime) });
                s.AddRow(new DateTime(2025, 1, 1));
                noHierarchy.AddRelationship(noHierarchy.Tables["Sales"]["Date"], noHierarchy.Tables["Dates"]["Date"]);
                noHierarchy.Tables["Dates"].MarkAsDateTable("Date");
                Rejects<InvalidOperationException>(() => Saves(noHierarchy), "no hierarchy on the date table");
            }
            using (var noRelationship = new DataModelWriter(new DataModelOptions { XpressCompression = false }))
            {
                var d = noRelationship.CreateTable("Dates", new[] { new ColumnDefinition("Date", ColumnDataType.DateTime), new ColumnDefinition("Year", ColumnDataType.Int64), new ColumnDefinition("Month", ColumnDataType.String) });
                d.AddRow(new DateTime(2025, 1, 1), 2025, "January");
                noRelationship.Tables["Dates"].AddHierarchy("Calendar", "Year", "Month");
                noRelationship.Tables["Dates"].MarkAsDateTable("Date");
                Rejects<InvalidOperationException>(() => Saves(noRelationship), "no relationship uses the date column");
            }
            var (blank, _, _) = DateModel(t => { t.AddRow(new DateTime(2025, 1, 1), 2025, "January"); t.AddRow(null, 2025, "January"); });
            using (blank)
            {
                blank.Tables["Dates"].MarkAsDateTable("Date");
                Rejects<InvalidOperationException>(() => Saves(blank), "blank date");
            }
            var (partial, _, _) = DateModel(t => { t.AddRow(new DateTime(2025, 1, 1, 6, 0, 0), 2025, "January"); });
            using (partial)
            {
                partial.Tables["Dates"].MarkAsDateTable("Date");
                Rejects<InvalidOperationException>(() => Saves(partial), "date with a time part");
            }
            var (dup, _, _) = DateModel(t => { t.AddRow(new DateTime(2025, 1, 1), 2025, "January"); t.AddRow(new DateTime(2025, 1, 1), 2025, "January"); });
            using (dup)
            {
                dup.Tables["Dates"].MarkAsDateTable("Date");
                Rejects<InvalidOperationException>(() => Saves(dup), "duplicate date");
            }
        }

        public static void WorkbookRelationships()
        {
            var (model, many, one) = Pair(ColumnDataType.Int64);
            using (model)
            {
                one.AddRow(1, "x"); many.AddRow(1, 1);
                model.AddRelationship(model.Tables["Many"]["K"], model.Tables["One"]["K"]);
                var ms = new MemoryStream();
                model.Save(ms);
                ms.Position = 0;
                using var zip = new System.IO.Compression.ZipArchive(ms);
                using var r = new StreamReader(zip.GetEntry("xl/workbook.xml")!.Open());
                string xml = r.ReadToEnd();
                Check.True(xml.Contains("<x15:modelRelationships><x15:modelRelationship fromTable=\"Many\" fromColumn=\"K\" toTable=\"One\" toColumn=\"K\"/></x15:modelRelationships>"), "modelRelationships");
            }
        }

        /// <summary>
        /// Excel refuses to load a model whose relationship keys lack these markers (found by diffing against an Excel-written
        /// model, verified in Excel): ColumnFlags 11, EncodingHint 1 and Usage 0 on the primary key, Usage 1 on the foreign key,
        /// NullProcessing Error on the key attribute and Cardinality One from the row number to it.
        /// </summary>
        public static void KeyColumnFlags()
        {
            var (model, many, one) = Pair(ColumnDataType.Int64);
            using (model)
            {
                one.AddRow(1, "x"); many.AddRow(1, 1);
                model.AddRelationship(model.Tables["Many"]["K"], model.Tables["One"]["K"]);
                var ms = new MemoryStream();
                model.SaveModelPart(ms);
                var m = ModelReader.Read(ms.ToArray());
                string pk = Encoding.UTF8.GetString(m.Find(@"\.0\.dim\\One_G\.0\.tbl\.xml$"));
                string fk = Encoding.UTF8.GetString(m.Find(@"\.0\.dim\\Many_G\.0\.tbl\.xml$"));
                string Col(string xml, string name)
                {
                    int i = xml.IndexOf("class=\"XMRawColumn\" name=\"" + name + "\"", StringComparison.Ordinal);
                    int j = xml.IndexOf("class=\"XMRawColumn\"", i + 10, StringComparison.Ordinal);
                    return xml.Substring(i, (j < 0 ? xml.Length : j) - i);
                }
                string pkCol = Col(pk, "K"), fkCol = Col(fk, "K"), other = Col(pk, "V");
                Check.True(pkCol.Contains("<ColumnFlags xsi:type=\"xsd:long\">11</ColumnFlags>"), "pk ColumnFlags 11");
                Check.True(pkCol.Contains("<EncodingHint xsi:type=\"xsd:int\">1</EncodingHint>"), "pk EncodingHint 1");
                Check.True(pkCol.Contains("<Usage xsi:type=\"xsd:int\">0</Usage>"), "pk Usage 0");
                Check.True(fkCol.Contains("<Usage xsi:type=\"xsd:int\">1</Usage>"), "fk Usage 1");
                Check.True(other.Contains("<Usage xsi:type=\"xsd:int\">3</Usage>") && other.Contains("<ColumnFlags xsi:type=\"xsd:long\">8</ColumnFlags>"), "other columns unchanged");
                string dim = Encoding.UTF8.GetString(m.Find(@"\.0\.db\\One_G\.1\.dim\.xml$"));
                Check.True(dim.Contains("<Cardinality>One</Cardinality>"), "row number -> key cardinality One");
                Check.True(System.Text.RegularExpressions.Regex.IsMatch(dim, @"<DataType>BigInt</DataType><DataSize>-1</DataSize><NullProcessing>Error</NullProcessing>"), "key attribute NullProcessing Error");
            }
        }

        private static DataModelWriter SortedProducts()
        {
            var model = new DataModelWriter(new DataModelOptions { XpressCompression = false });
            var t = model.CreateTable("Products", new[]
            {
                new ColumnDefinition("ProductKey", ColumnDataType.Int64),
                new ColumnDefinition("Category", ColumnDataType.String) { SortByColumn = "CategoryOrder" },
                new ColumnDefinition("CategoryOrder", ColumnDataType.Int64),
                new ColumnDefinition("Product", ColumnDataType.String),
            });
            // First-appearance ids: Bikes 3, Clothing 4, Accessories 5. Alphabetical order would be Accessories, Bikes, Clothing.
            string[] cats = { "Bikes", "Clothing", "Accessories" };
            int k = 1;
            for (int c = 0; c < 3; c++) for (int i = 1; i <= 2; i++) t.AddRow(k++, cats[c], c + 1, $"{cats[c]} {i}");
            return model;
        }

        public static void SortByColumn()
        {
            using var model = SortedProducts();
            model.Tables["Products"].AddHierarchy("Browse", "Category", "Product");
            var ms = new MemoryStream();
            model.SaveModelPart(ms);
            var m = ModelReader.Read(ms.ToArray());
            // The tree lists the top level in the order of CategoryOrder: Bikes (3), Clothing (4), Accessories (5).
            var ids = DecodeTree(m, "Products", "Browse", "MULTI_LEVEL_ID", 0);
            Check.True(ids[0] == 3 && ids[1] == 4 && ids[2] == 5, $"top level order {ids[0]},{ids[1]},{ids[2]}");
            string dim = Encoding.UTF8.GetString(m.Find(@"\.0\.db\\Products_G\.5\.dim\.xml$"));
            Check.Equal(1, System.Text.RegularExpressions.Regex.Matches(dim, "<OrderBy>AttributeKey</OrderBy>").Count, "one attribute ordered by another");
            Check.True(dim.Contains("<OrderByAttributeID>CategoryOrder</OrderByAttributeID>"), "OrderByAttributeID");
            string tbl = Encoding.UTF8.GetString(m.Find(@"\.0\.dim\\Products_G\.2\.tbl\.xml$"));
            Check.True(tbl.Contains("<OrderByColumn>CategoryOrder</OrderByColumn>"), "OrderByColumn");
            Check.Equal(4, System.Text.RegularExpressions.Regex.Matches(tbl, "<OrderByColumn/>").Count, "other columns and the row number keep an empty OrderByColumn");
        }

        public static void SortByColumnValidation()
        {
            void Rejects<T>(Action a, string what) where T : Exception
            {
                try { a(); } catch (T) { return; }
                throw new Exception("Expected " + typeof(T).Name + ": " + what);
            }
            using (var model = SortedProducts())
            {
                Rejects<ArgumentException>(() => model.Tables["Products"]["Product"].SortByColumn = model.Tables["Products"]["Product"], "itself");
                Rejects<KeyNotFoundException>(() => model.Tables["Products"]["Product"].SortByColumn = model.Tables["Products"]["Nope"], "unknown column");
                Rejects<ArgumentException>(() => model.Tables["Products"]["Product"].SortByColumn = model.Tables["Products"]["Category"], "sort column that is itself sorted");
                Rejects<ArgumentException>(() => model.Tables["Products"]["CategoryOrder"].SortByColumn = model.Tables["Products"]["ProductKey"], "sorted column used as sort column");
                Rejects<KeyNotFoundException>(() => model.Tables["Nope"]["Product"].SortByColumn = model.Tables["Nope"]["ProductKey"], "unknown table");
            }
            using (var model = new DataModelWriter())
            {
                var t = model.CreateTable("T", new[] { new ColumnDefinition("Month", ColumnDataType.String) { SortByColumn = "N" }, new ColumnDefinition("N", ColumnDataType.Int64) });
                t.AddRow("Jan", 1); t.AddRow("Jan", 2);
                Rejects<InvalidOperationException>(() => model.SaveModelPart(new MemoryStream()), "one value with two sort values");
            }
        }

        public static void RelationshipDefaults()
        {
            string Build(bool helper, bool measure, out string dimName)
            {
                var (model, many, one) = Pair(ColumnDataType.Int64);
                using (model)
                {
                    one.AddRow(1, "x"); many.AddRow(1, 1);
                    model.AddRelationship(model.Tables["Many"]["K"], model.Tables["One"]["K"]);
                    if (measure) model.Tables["Many"].AddMeasure("Mine", "SUM(Many[N])");
                    var ms = new MemoryStream();
                    var opts = new DataModelOptions { AddHelperMeasure = helper, XpressCompression = false };
                    using var m2 = new DataModelWriter(opts);
                    // Rebuild with the options under test.
                    var o = m2.CreateTable("One", new[] { new ColumnDefinition("K", ColumnDataType.Int64), new ColumnDefinition("V", ColumnDataType.String) });
                    var mn = m2.CreateTable("Many", new[] { new ColumnDefinition("K", ColumnDataType.Int64), new ColumnDefinition("N", ColumnDataType.Int64) });
                    o.AddRow(1, "x"); mn.AddRow(1, 1);
                    m2.AddRelationship(m2.Tables["Many"]["K"], m2.Tables["One"]["K"]);
                    if (measure) m2.Tables["Many"].AddMeasure("Mine", "SUM(Many[N])");
                    m2.SaveModelPart(ms);
                    var r = ModelReader.Read(ms.ToArray());
                    dimName = r.Find(@"\.0\.db\\Many_G\.2\.dim\.xml$").Length > 0 ? "Many v2" : "?";
                    Check.True(r.Find(@"\.0\.db\\One_G\.1\.dim\.xml$").Length > 0, "table without relationships stays at version 1");
                    string stats = Encoding.UTF8.GetString(r.Find(@"\.0\.dim\\R\$Many_G\$G\.0\.tbl\.xml$"));
                    Check.True(stats.Contains("<XMType xsi:type=\"xsd:int\">0</XMType>") && stats.Contains("<RowCount xsi:type=\"xsd:long\">0</RowCount>"), "relationship index carries Excel's placeholder statistics");
                    var doc = new System.Xml.XmlDocument();
                    doc.LoadXml(Encoding.UTF8.GetString(r.Find(@"MdxScript\.1\.scr\.xml$")).TrimStart('\uFEFF'));
                    return doc.OuterXml.Replace("&lt;", "<").Replace("&gt;", ">");
                }
            }
            string s = Build(helper: true, measure: false, out string d);
            Check.Equal("Many v2", d, "relationship holder at version 2");
            Check.True(s.Contains("CREATE MEASURE 'Many'[PivotWright helper]=1;") && s.Contains("[__XL_Count of Models]"), "helper measure added");
            Check.True(s.Contains("<Value>True</Value>"), "helper measure hidden");
            s = Build(helper: false, measure: false, out _);
            Check.True(!s.Contains("CREATE MEASURE") && s.Contains("__No measures defined"), "helper can be turned off");
            s = Build(helper: true, measure: true, out _);
            Check.True(s.Contains("CREATE MEASURE 'Many'[Mine]") && !s.Contains("PivotWright helper"), "no helper next to a real measure");
        }
            private static void Throws<T>(Action a, string what, string? messagePart = null) where T : Exception
        {
            try { a(); }
            catch (T e)
            {
                if (messagePart is not null && !e.Message.Contains(messagePart, StringComparison.Ordinal))
                    throw new Exception($"{what}: message '{e.Message}' does not contain '{messagePart}'");
                return;
            }
            throw new Exception("Expected " + typeof(T).Name + ": " + what);
        }

        // ---- Perspectives (SPEC section 13; Excel's H64, H65, H66) ----

        private static (DataModelWriter model, ModelTable products, ModelTable sales, ModelMeasure total) PerspectiveModel()
        {
            var (model, products, sales) = CalcModel();
            return (model, products, sales, sales.AddMeasure("Total Qty", "SUM(Sales[Qty])", MeasureFormat.WholeNumber));
        }

        private static string PerspectiveText(ModelReader m, string pattern) => Encoding.UTF8.GetString(m.Find(pattern));

        // The attribute ids a perspective file lists, per table id prefix.
        private static List<string> PerspectiveAttributes(string xml, string tablePrefix)
        {
            var doc = new System.Xml.XmlDocument();
            doc.LoadXml(xml.TrimStart('\uFEFF'));
            var dim = doc.GetElementsByTagName("Dimension").Cast<System.Xml.XmlNode>().FirstOrDefault(d => d["CubeDimensionID"]!.InnerText.StartsWith(tablePrefix + "_", StringComparison.Ordinal));
            return dim is null ? new List<string>() : dim["Attributes"]!.ChildNodes.Cast<System.Xml.XmlNode>().Select(a => a["AttributeID"]!.InnerText).OrderBy(x => x, StringComparer.Ordinal).ToList();
        }

        // The XML text of the PerspectiveEmptyHierarchyNames annotation of the cube, or null.
        private static string? PerspectiveAnnotation(ModelReader m)
        {
            var doc = new System.Xml.XmlDocument();
            doc.LoadXml(Encoding.UTF8.GetString(m.Find(@"Model\.\d+\.cub\.xml$")).TrimStart('\uFEFF'));
            foreach (System.Xml.XmlNode a in doc.GetElementsByTagName("Annotation"))
                if (a["Name"]!.InnerText == "PerspectiveEmptyHierarchyNames") return a["Value"]!.InnerText;
            return null;
        }

        // H73: Sales[Ship Date] dragged onto Dates[Date] in Diagram View, next to the active relationships of the standard model.
        private static (DataModelWriter model, ModelTable products, ModelTable dates, ModelTable sales, ModelRelationship inactive) InactiveModel(bool inactive = true, bool measure = true, bool dateFirst = false, bool hierarchies = true)
        {
            var model = new DataModelWriter(new DataModelOptions { XpressCompression = false });
            var products = model.CreateTable("Products", new[] { new ColumnDefinition("ProductKey", ColumnDataType.Int64), new ColumnDefinition("Category", ColumnDataType.String), new ColumnDefinition("Product", ColumnDataType.String) });
            for (int i = 1; i <= 6; i++) products.AddRow(i, i <= 3 ? "A" : "B", "P" + i);
            var dates = model.CreateTable("Dates", new[] { new ColumnDefinition("Date", ColumnDataType.DateTime), new ColumnDefinition("Year", ColumnDataType.Int64), new ColumnDefinition("Month", ColumnDataType.Int64) });
            for (int i = 0; i < 40; i++) { var d = new DateTime(2025, 1, 1).AddDays(i); dates.AddRow(d, d.Year, d.Month); }
            var sales = model.CreateTable("Sales", new[] { new ColumnDefinition("ProductKey", ColumnDataType.Int64), new ColumnDefinition("Date", ColumnDataType.DateTime), new ColumnDefinition("Qty", ColumnDataType.Int64), new ColumnDefinition("Ship Date", ColumnDataType.DateTime) });
            for (int i = 0; i < 40; i++) sales.AddRow(1 + i % 6, new DateTime(2025, 1, 1).AddDays(i), 1 + i % 5, new DateTime(2025, 1, 1).AddDays((i + 3) % 40));
            if (dateFirst) { model.AddRelationship(sales["Date"], dates["Date"]); model.AddRelationship(sales["ProductKey"], products["ProductKey"]); }
            else { model.AddRelationship(sales["ProductKey"], products["ProductKey"]); model.AddRelationship(sales["Date"], dates["Date"]); }
            var rel = inactive ? model.AddRelationship(sales["Ship Date"], dates["Date"], isActive: false) : null!;
            if (hierarchies) { products.AddHierarchy("Product Hierarchy", "Category", "Product"); dates.AddHierarchy("Calendar", "Year", "Month"); }
            if (measure) sales.AddMeasure("Total Qty", "SUM(Sales[Qty])", MeasureFormat.WholeNumber);
            return (model, products, dates, sales, rel);
        }

        public static void InactiveRelationship()
        {
            var (model, _, _, _, rel) = InactiveModel();
            using (model)
            {
                Check.True(!rel.IsActive, "the relationship is inactive");
                Check.True(model.Relationships.Take(2).All(r => r.IsActive), "the first two are active");
                var m = SaveKpiModel(model);
                // Versions of H73: cube and measure groups 21, partitions and script 10, Sales dimension 4, Products 6, Dates unchanged at 5.
                string cube = PerspectiveText(m, @"Model\.21\.cub\.xml$");
                Check.True(cube.Contains("<StructureVersion>12</StructureVersion>"), "cube structure 12");
                PerspectiveText(m, @"MdxScript\.10\.scr\.xml$");
                string sales = PerspectiveText(m, @"Sales_[^\\]+\.4\.dim\.xml$");
                PerspectiveText(m, @"Products_[^\\]+\.6\.dim\.xml$");
                PerspectiveText(m, @"Dates_[^\\]+\.5\.dim\.xml$");
                // Three relationships in the Sales dimension and table, three index tables, but only two measure group dimensions.
                Check.Equal(3, System.Text.RegularExpressions.Regex.Matches(sales, "<ddl300_300:Relationship>").Count, "relationships in the dimension");
                Check.True(sales.Contains("<StructureVersion>1</StructureVersion><DataIDVersion>2</DataIDVersion><IncDataIDVersion>2</IncDataIDVersion><CacheVersion>2</CacheVersion><DataVersion>1</DataVersion>"), "Sales counters");
                Check.Equal(2, System.Text.RegularExpressions.Regex.Matches(sales, "<ColumnID>Ship_x0020_Date</ColumnID>").Count, "escaped column id in the attribute");
                string det = PerspectiveText(m, @"Model\.0\.cub\\Sales_[^\\]+\.21\.det\.xml$");
                Check.Equal(2, System.Text.RegularExpressions.Regex.Matches(det, "ReferenceMeasureGroupDimension\"").Count, "measure group dimensions");
                Check.True(det.Contains("<ColumnID>Ship_x0020_Date</ColumnID>"), "escaped column id in the measure group");
                Check.True(det.Contains("<StructureVersion>9</StructureVersion><CacheVersion>2</CacheVersion><DataVersion>-1</DataVersion>"), "Sales measure group counters");
                string table = PerspectiveText(m, @"Sales_[^\\]+\.1\.tbl\.xml$");
                Check.Equal(3, System.Text.RegularExpressions.Regex.Matches(table, "class=\"XMRelationship\"").Count, "relationships in the table");
                Check.Equal(3, m.Files.Keys.Count(k => k.Contains("INDEX.0.idf") && k.Contains("R$Sales")), "index tables");
                Check.True(PerspectiveText(m, @"Sales_[^\\]+\.0\.dim\\info\.4\.xml$").Contains("<DataVersion>1</DataVersion><IndexVersion>-1</IndexVersion>"), "Sales dimension information");
            }
        }

        public static void InactiveRelationshipWorkbookAndAbsent()
        {
            var (model, _, _, _, _) = InactiveModel();
            using (model)
            {
                var ms = new MemoryStream();
                model.Save(ms);
                ms.Position = 0;
                using var zip = new System.IO.Compression.ZipArchive(ms);
                using var reader = new StreamReader(zip.GetEntry("xl/workbook.xml")!.Open());
                string workbook = reader.ReadToEnd();
                Check.Equal(2, System.Text.RegularExpressions.Regex.Matches(workbook, "<x15:modelRelationship ").Count, "relationships in the workbook");
                Check.True(!workbook.Contains("Ship Date"), "the inactive relationship is not listed");
            }
            // The same model without the inactive relationship keeps the hierarchy profile.
            var (plain, _, _, _, _) = InactiveModel(inactive: false);
            using (plain)
            {
                var m = SaveKpiModel(plain);
                PerspectiveText(m, @"Model\.18\.cub\.xml$");
                PerspectiveText(m, @"MdxScript\.9\.scr\.xml$");
                PerspectiveText(m, @"Sales_[^\\]+\.3\.dim\.xml$");
            }
        }

        public static void InactiveRelationshipPerspective()
        {
            // H84: the perspective base (cube 20, script 10, partitions 10, Products 6): the cube 23, the script and partitions 11, Sales 4, Products 7,
            // and the perspective file one version on (Sales view.1.persp.xml).
            var (model, products, _, sales, _) = InactiveModel();
            using (model)
            {
                model.AddPerspective("Sales view").Include(products.Hierarchies[0]).Include(sales.Measures[0]);
                var m = SaveKpiModel(model);
                PerspectiveText(m, @"Model\.23\.cub\.xml$");
                PerspectiveText(m, @"MdxScript\.11\.scr\.xml$");
                PerspectiveText(m, @"Sales_[^\\]+\.4\.dim\.xml$");
                PerspectiveText(m, @"Products_[^\\]+\.7\.dim\.xml$");
                PerspectiveText(m, @"Sales view\.1\.persp\.xml$");
                Check.True(PerspectiveText(m, @"Model\.23\.cub\.xml$").Contains("<PerspectiveFileList>Sales view.1.persp.xml</PerspectiveFileList>"), "perspective list");
            }
        }

        public static void KpiCalcResavedTable()
        {
            // H90 (tables Dates, Products, Sales) and H91 (the Date relationship first): cube 26, script 12. Both the calculated column step (H92, H93)
            // and the KPI step (H90, H91) save the first table, so it is at 7 and the other table stays at 5 (H90 and H91 started from a base whose
            // calculated column step had saved the other table, which is why they end with both at 6).
            foreach (var (datesFirst, dateRelFirst) in new[] { (true, false), (false, true) })
            {
                using var model = new DataModelWriter(new DataModelOptions { XpressCompression = false, CalculateOnOpen = false });
                ModelTable products = null!, dates = null!;
                void MakeProducts() { products = model.CreateTable("Products", new[] { new ColumnDefinition("ProductKey", ColumnDataType.Int64), new ColumnDefinition("Category", ColumnDataType.String), new ColumnDefinition("Product", ColumnDataType.String) }); for (int i = 1; i <= 6; i++) products.AddRow(i, i <= 3 ? "A" : "B", "P" + i); }
                void MakeDates() { dates = model.CreateTable("Dates", new[] { new ColumnDefinition("Date", ColumnDataType.DateTime), new ColumnDefinition("Year", ColumnDataType.Int64), new ColumnDefinition("Month", ColumnDataType.Int64) }); for (int i = 0; i < 40; i++) { var d = new DateTime(2025, 1, 1).AddDays(i); dates.AddRow(d, d.Year, d.Month); } }
                if (datesFirst) { MakeDates(); MakeProducts(); } else { MakeProducts(); MakeDates(); }
                var sales = model.CreateTable("Sales", new[] { new ColumnDefinition("ProductKey", ColumnDataType.Int64), new ColumnDefinition("Date", ColumnDataType.DateTime), new ColumnDefinition("Qty", ColumnDataType.Int64), new ColumnDefinition("Amount", ColumnDataType.Currency) });
                for (int i = 0; i < 40; i++) sales.AddRow(1 + i % 6, new DateTime(2025, 1, 1).AddDays(i), 1 + i % 5, (decimal)(10 + i % 7));
                if (dateRelFirst) { model.AddRelationship(sales["Date"], dates["Date"]); model.AddRelationship(sales["ProductKey"], products["ProductKey"]); }
                else { model.AddRelationship(sales["ProductKey"], products["ProductKey"]); model.AddRelationship(sales["Date"], dates["Date"]); }
                products.AddHierarchy("Product Hierarchy", "Category", "Product");
                dates.AddHierarchy("Calendar", "Year", "Month");
                sales.AddCalculatedColumn("Line Total", "[Qty]*[Amount]", ColumnDataType.Currency);
                sales.AddMeasure("Total Qty", "SUM(Sales[Qty])", MeasureFormat.WholeNumber).AddKpi(1000);
                var m = SaveKpiModel(model);
                PerspectiveText(m, @"Model\.26\.cub\.xml$");
                PerspectiveText(m, @"MdxScript\.12\.scr\.xml$");
                PerspectiveText(m, @$"{(datesFirst ? "Dates" : "Products")}_[^\\]+\.7\.dim\.xml$");
                PerspectiveText(m, @$"{(datesFirst ? "Products" : "Dates")}_[^\\]+\.5\.dim\.xml$");
            }
        }

        public static void PerspectiveResavedTableIsTheFirstTable()
        {
            // H88: tables created in the order Dates, Products, Sales: Dates saved again (6); H89: the Date relationship first: Products (6).
            foreach (var (datesFirst, dateRelFirst) in new[] { (true, false), (false, true) })
            {
                using var model = new DataModelWriter(new DataModelOptions { XpressCompression = false });
                ModelTable products = null!, dates = null!;
                void MakeProducts() { products = model.CreateTable("Products", new[] { new ColumnDefinition("ProductKey", ColumnDataType.Int64), new ColumnDefinition("Category", ColumnDataType.String), new ColumnDefinition("Product", ColumnDataType.String) }); for (int i = 1; i <= 6; i++) products.AddRow(i, i <= 3 ? "A" : "B", "P" + i); }
                void MakeDates() { dates = model.CreateTable("Dates", new[] { new ColumnDefinition("Date", ColumnDataType.DateTime), new ColumnDefinition("Year", ColumnDataType.Int64), new ColumnDefinition("Month", ColumnDataType.Int64) }); for (int i = 0; i < 40; i++) { var d = new DateTime(2025, 1, 1).AddDays(i); dates.AddRow(d, d.Year, d.Month); } }
                if (datesFirst) { MakeDates(); MakeProducts(); } else { MakeProducts(); MakeDates(); }
                var sales = model.CreateTable("Sales", new[] { new ColumnDefinition("ProductKey", ColumnDataType.Int64), new ColumnDefinition("Date", ColumnDataType.DateTime), new ColumnDefinition("Qty", ColumnDataType.Int64) });
                for (int i = 0; i < 40; i++) sales.AddRow(1 + i % 6, new DateTime(2025, 1, 1).AddDays(i), 1 + i % 5);
                if (dateRelFirst) { model.AddRelationship(sales["Date"], dates["Date"]); model.AddRelationship(sales["ProductKey"], products["ProductKey"]); }
                else { model.AddRelationship(sales["ProductKey"], products["ProductKey"]); model.AddRelationship(sales["Date"], dates["Date"]); }
                products.AddHierarchy("Product Hierarchy", "Category", "Product");
                dates.AddHierarchy("Calendar", "Year", "Month");
                var total = sales.AddMeasure("Total Qty", "SUM(Sales[Qty])", MeasureFormat.WholeNumber);
                model.AddPerspective("Sales view").Include(products.Hierarchies[0]).Include(total);
                var m = SaveKpiModel(model);
                string saved = datesFirst ? "Dates" : "Products", other = datesFirst ? "Products" : "Dates";
                PerspectiveText(m, @$"{saved}_[^\\]+\.6\.dim\.xml$");
                PerspectiveText(m, @$"{other}_[^\\]+\.5\.dim\.xml$");
            }
        }

        public static void KpiResavedTableIsTheFirstTable()
        {
            // H86: the tables created in the order Dates, Products, Sales (Sales to Products first): the KPI saves Dates again (6), Products stays 5.
            // H87: the usual table order with the Date relationship first: Products (6), Dates 5.
            foreach (var (datesFirst, dateRelFirst) in new[] { (true, false), (false, true) })
            {
                using var model = new DataModelWriter(new DataModelOptions { XpressCompression = false });
                ModelTable products = null!, dates = null!;
                void MakeProducts() { products = model.CreateTable("Products", new[] { new ColumnDefinition("ProductKey", ColumnDataType.Int64), new ColumnDefinition("Category", ColumnDataType.String), new ColumnDefinition("Product", ColumnDataType.String) }); for (int i = 1; i <= 6; i++) products.AddRow(i, i <= 3 ? "A" : "B", "P" + i); }
                void MakeDates() { dates = model.CreateTable("Dates", new[] { new ColumnDefinition("Date", ColumnDataType.DateTime), new ColumnDefinition("Year", ColumnDataType.Int64), new ColumnDefinition("Month", ColumnDataType.Int64) }); for (int i = 0; i < 40; i++) { var d = new DateTime(2025, 1, 1).AddDays(i); dates.AddRow(d, d.Year, d.Month); } }
                if (datesFirst) { MakeDates(); MakeProducts(); } else { MakeProducts(); MakeDates(); }
                var sales = model.CreateTable("Sales", new[] { new ColumnDefinition("ProductKey", ColumnDataType.Int64), new ColumnDefinition("Date", ColumnDataType.DateTime), new ColumnDefinition("Qty", ColumnDataType.Int64) });
                for (int i = 0; i < 40; i++) sales.AddRow(1 + i % 6, new DateTime(2025, 1, 1).AddDays(i), 1 + i % 5);
                if (dateRelFirst) { model.AddRelationship(sales["Date"], dates["Date"]); model.AddRelationship(sales["ProductKey"], products["ProductKey"]); }
                else { model.AddRelationship(sales["ProductKey"], products["ProductKey"]); model.AddRelationship(sales["Date"], dates["Date"]); }
                products.AddHierarchy("Product Hierarchy", "Category", "Product");
                dates.AddHierarchy("Calendar", "Year", "Month");
                sales.AddMeasure("Total Qty", "SUM(Sales[Qty])", MeasureFormat.WholeNumber).AddKpi(1000);
                var m = SaveKpiModel(model);
                string saved = datesFirst ? "Dates" : "Products", other = datesFirst ? "Products" : "Dates";
                PerspectiveText(m, @$"{saved}_[^\\]+\.6\.dim\.xml$");
                PerspectiveText(m, @$"{other}_[^\\]+\.5\.dim\.xml$");
            }
        }

        public static void InactiveRelationshipTwoFacts()
        {
            // H85: Sales[Ship Date] and then Stock[Restock Date] onto Dates[Date], Stock related to Dates only: the cube 23 (18 and five), the script and
            // partitions 11, Sales 4, Stock 5, Products and Dates unchanged at 5, and no information file for Stock, the table of the last relationship.
            var (model, _, dates, sales, _) = InactiveModel();
            using (model)
            {
                var stock = model.CreateTable("Stock", new[] { new ColumnDefinition("Date", ColumnDataType.DateTime), new ColumnDefinition("Restock Date", ColumnDataType.DateTime), new ColumnDefinition("OnHand", ColumnDataType.Int64) });
                for (int i = 0; i < 20; i++) stock.AddRow(new DateTime(2025, 1, 1).AddDays(i), new DateTime(2025, 1, 1).AddDays((i + 3) % 40), 100 + i);
                model.AddRelationship(stock["Date"], dates["Date"]);
                var restock = model.AddRelationship(stock["Restock Date"], dates["Date"], isActive: false);
                Check.True(!restock.IsActive, "inactive");
                var m = SaveKpiModel(model);
                PerspectiveText(m, @"Model\.23\.cub\.xml$");
                PerspectiveText(m, @"MdxScript\.11\.scr\.xml$");
                PerspectiveText(m, @"Sales_[^\\]+\.4\.dim\.xml$");
                PerspectiveText(m, @"Stock_[^\\]+\.5\.dim\.xml$");
                PerspectiveText(m, @"Products_[^\\]+\.5\.dim\.xml$");
                PerspectiveText(m, @"Dates_[^\\]+\.5\.dim\.xml$");
                Check.Equal(1, m.Files.Keys.Count(k => System.Text.RegularExpressions.Regex.IsMatch(k, @"Sales_[^\\]+\.0\.dim\\info\.\d+\.xml$")), "Sales keeps its information file");
                Check.Equal(0, m.Files.Keys.Count(k => System.Text.RegularExpressions.Regex.IsMatch(k, @"Stock_[^\\]+\.0\.dim\\info\.\d+\.xml$")), "Stock loses it");
            }
        }

        public static void InactiveRelationshipKpi()
        {
            // H83: the KPI base (cube 19, script 10, Products 6): the cube 22, the script 11, partitions 10, Sales 4, Products 7, no cube information file.
            var (model, _, _, sales, _) = InactiveModel();
            using (model)
            {
                sales.Measures[0].AddKpi(1000);
                var m = SaveKpiModel(model);
                PerspectiveText(m, @"Model\.22\.cub\.xml$");
                PerspectiveText(m, @"MdxScript\.11\.scr\.xml$");
                PerspectiveText(m, @"Sales_[^\\]+\.4\.dim\.xml$");
                PerspectiveText(m, @"Products_[^\\]+\.7\.dim\.xml$");
                PerspectiveText(m, @"Dates_[^\\]+\.5\.dim\.xml$");
                Check.True(PerspectiveText(m, @"MdxScript\.11\.scr\.xml$").Contains("KPIs.[Total Qty]"), "the KPI");
                Check.Equal(0, m.Files.Keys.Count(k => System.Text.RegularExpressions.Regex.IsMatch(k, @"Model\.0\.cub\\info\.\d+\.xml$")), "no cube information file");
            }
        }

        public static void InactiveRelationshipPlain()
        {
            // H79: the model without hierarchies (base 1): the cube 4, the script and partitions 2, Sales 3, Products 2, Dates 1, no cube information file.
            var (model, _, _, _, _) = InactiveModel(hierarchies: false);
            using (model)
            {
                var m = SaveKpiModel(model);
                PerspectiveText(m, @"Model\.4\.cub\.xml$");
                PerspectiveText(m, @"MdxScript\.2\.scr\.xml$");
                string sales = PerspectiveText(m, @"Sales_[^\\]+\.3\.dim\.xml$");
                PerspectiveText(m, @"Products_[^\\]+\.2\.dim\.xml$");
                PerspectiveText(m, @"Dates_[^\\]+\.1\.dim\.xml$");
                Check.True(sales.Contains("<StructureVersion>1</StructureVersion><DataIDVersion>2</DataIDVersion><IncDataIDVersion>2</IncDataIDVersion><CacheVersion>2</CacheVersion><DataVersion>1</DataVersion>"), "Sales counters");
                Check.Equal(0, m.Files.Keys.Count(k => System.Text.RegularExpressions.Regex.IsMatch(k, @"Model\.0\.cub\\info\.\d+\.xml$")), "no cube information file");
            }
            // The same model without the inactive relationship keeps the plain profile.
            var (plain, _, _, _, _) = InactiveModel(inactive: false, hierarchies: false);
            using (plain) PerspectiveText(SaveKpiModel(plain), @"Model\.1\.cub\.xml$");
        }

        public static void InactiveRelationshipDateFirst()
        {
            // H78: the date relationship before the product one. Excel still saved Products again (6) and left Dates at 5.
            var (model, _, _, _, _) = InactiveModel(dateFirst: true);
            using (model)
            {
                var m = SaveKpiModel(model);
                PerspectiveText(m, @"Model\.21\.cub\.xml$");
                PerspectiveText(m, @"Sales_[^\\]+\.4\.dim\.xml$");
                PerspectiveText(m, @"Products_[^\\]+\.6\.dim\.xml$");
                PerspectiveText(m, @"Dates_[^\\]+\.5\.dim\.xml$");
            }
        }

        public static void InactiveRelationshipTwo()
        {
            // H77: Ship Date and then Delivery Date dragged onto Dates[Date].
            var two = new DataModelWriter(new DataModelOptions { XpressCompression = false });
            var products = two.CreateTable("Products", new[] { new ColumnDefinition("ProductKey", ColumnDataType.Int64), new ColumnDefinition("Category", ColumnDataType.String), new ColumnDefinition("Product", ColumnDataType.String) });
            for (int i = 1; i <= 6; i++) products.AddRow(i, i <= 3 ? "A" : "B", "P" + i);
            var d = two.CreateTable("Dates", new[] { new ColumnDefinition("Date", ColumnDataType.DateTime), new ColumnDefinition("Year", ColumnDataType.Int64), new ColumnDefinition("Month", ColumnDataType.Int64) });
            for (int i = 0; i < 40; i++) { var day = new DateTime(2025, 1, 1).AddDays(i); d.AddRow(day, day.Year, day.Month); }
            var s = two.CreateTable("Sales", new[] { new ColumnDefinition("ProductKey", ColumnDataType.Int64), new ColumnDefinition("Date", ColumnDataType.DateTime), new ColumnDefinition("Qty", ColumnDataType.Int64), new ColumnDefinition("Ship Date", ColumnDataType.DateTime), new ColumnDefinition("Delivery Date", ColumnDataType.DateTime) });
            for (int i = 0; i < 40; i++) s.AddRow(1 + i % 6, new DateTime(2025, 1, 1).AddDays(i), 1 + i % 5, new DateTime(2025, 1, 1).AddDays((i + 3) % 40), new DateTime(2025, 1, 1).AddDays((i + 5) % 40));
            two.AddRelationship(s["ProductKey"], products["ProductKey"]);
            two.AddRelationship(s["Date"], d["Date"]);
            two.AddRelationship(s["Ship Date"], d["Date"], isActive: false);
            two.AddRelationship(s["Delivery Date"], d["Date"], isActive: false);
            products.AddHierarchy("Product Hierarchy", "Category", "Product");
            d.AddHierarchy("Calendar", "Year", "Month");
            s.AddMeasure("Total Qty", "SUM(Sales[Qty])", MeasureFormat.WholeNumber);
            using (two)
            {
                var m = SaveKpiModel(two);
                // From 18, 9 and 3 (the hierarchy profile): the cube 23, the script and partitions 11, Sales 5, Products 6, Dates 5, Sales table 2.
                string cube = PerspectiveText(m, @"Model\.23\.cub\.xml$");
                Check.True(cube.Contains("<StructureVersion>13</StructureVersion>"), "cube structure 13");
                PerspectiveText(m, @"MdxScript\.11\.scr\.xml$");
                string sales = PerspectiveText(m, @"Sales_[^\\]+\.5\.dim\.xml$");
                PerspectiveText(m, @"Products_[^\\]+\.6\.dim\.xml$");
                PerspectiveText(m, @"Dates_[^\\]+\.5\.dim\.xml$");
                PerspectiveText(m, @"\\Sales_[^\\]+\.2\.tbl\.xml$");
                Check.True(sales.Contains("<StructureVersion>2</StructureVersion><DataIDVersion>3</DataIDVersion><IncDataIDVersion>3</IncDataIDVersion><CacheVersion>3</CacheVersion><DataVersion>2</DataVersion>"), "Sales counters");
                Check.Equal(4, System.Text.RegularExpressions.Regex.Matches(sales, "<ddl300_300:Relationship>").Count, "relationships in the dimension");
                string det = PerspectiveText(m, @"Model\.0\.cub\\Sales_[^\\]+\.23\.det\.xml$");
                Check.Equal(2, System.Text.RegularExpressions.Regex.Matches(det, "ReferenceMeasureGroupDimension\"").Count, "measure group dimensions");
                Check.True(det.Contains("<StructureVersion>10</StructureVersion><CacheVersion>3</CacheVersion><DataVersion>-1</DataVersion>"), "Sales measure group counters");
                // The many-side dimension keeps no information file with two inactive relationships (one keeps it, H73).
                Check.Equal(0, m.Files.Keys.Count(k => System.Text.RegularExpressions.Regex.IsMatch(k, @"Sales_[^\\]+\.0\.dim\\info\.\d+\.xml$")), "no Sales information file");
            }
        }

        public static void InactiveRelationshipValidation()
        {
            var (model, products, dates, sales, _) = InactiveModel(inactive: false);
            using (model)
            {
                // A second relationship between two tables must say it is inactive.
                Throws<ArgumentException>(() => model.AddRelationship(sales["Ship Date"], dates["Date"]), "a second active relationship", "isActive: false");
                Throws<ArgumentException>(() => model.AddRelationship(sales["Date"], dates["Date"], isActive: false), "the same columns again", "already related");
                // An inactive relationship needs an active one first, with the same many side.
                Throws<ArgumentException>(() => model.AddRelationship(dates["Date"], sales["Date"], isActive: false), "reversed", "Add the active relationship first");
                var lone = model.CreateTable("Lone", new[] { new ColumnDefinition("Day", ColumnDataType.DateTime) });
                lone.AddRow(new DateTime(2025, 1, 1));
                Throws<ArgumentException>(() => model.AddRelationship(lone["Day"], dates["Date"], isActive: false), "no active relationship between the tables", "Add the active relationship first");
                Check.Equal(2, model.Relationships.Count, "failed calls change nothing");
                var second = model.AddRelationship(sales["Ship Date"], dates["Date"], isActive: false);
                Check.True(!second.IsActive && model.Relationships[0].IsActive, "IsActive");
                // Unsupported models are refused when saving.
                // Without a measure of its own the hidden helper measure stands in, with the same counters (H81).
                using (var other = InactiveModel(measure: false).model)
                {
                    var helper = SaveKpiModel(other);
                    PerspectiveText(helper, @"Model\.21\.cub\.xml$");
                    Check.True(PerspectiveText(helper, @"MdxScript\.10\.scr\.xml$").Contains("PivotWright helper"), "the helper measure");
                }
                // Hidden columns are supported with hierarchies and without a date table (H80): the cube 24, the measure groups one behind.
                sales["Qty"].IsHidden = true;
                var hiddenModel = SaveKpiModel(model);
                PerspectiveText(hiddenModel, @"Model\.24\.cub\.xml$");
                PerspectiveText(hiddenModel, @"Sales_[^\\]+\.23\.det\.xml$");
                PerspectiveText(hiddenModel, @"MdxScript\.10\.scr\.xml$");
                // ... but not together with a date table, which was not seen.
                dates.MarkAsDateTable("Date");
                Throws<InvalidOperationException>(() => model.SaveModelPart(new MemoryStream()), "hidden columns with a date table", "without a date table");
                sales["Qty"].IsHidden = false;
                // A date table is supported (H76): the counters go three further, the dimensions as in H73 one step on from the date profile.
                var m = SaveKpiModel(model);
                PerspectiveText(m, @"Model\.24\.cub\.xml$");
                PerspectiveText(m, @"MdxScript\.11\.scr\.xml$");
                PerspectiveText(m, @"Sales_[^\\]+\.5\.dim\.xml$");
                PerspectiveText(m, @"Products_[^\\]+\.7\.dim\.xml$");
                PerspectiveText(m, @"Dates_[^\\]+\.6\.dim\.xml$");
                PerspectiveText(m, @"\\Sales_[^\\]+\.2\.tbl\.xml$");
                // The inactive relationship's index table stays at version 0 while the active one to the date table is at 2.
                Check.Equal(1, m.Files.Keys.Count(k => k.Contains("R$Sales") && k.Contains("\\0.R$") && k.EndsWith("INDEX.0.idf")) - 1, "index tables at version 0: the product key and the inactive one");
                Check.Equal(1, m.Files.Keys.Count(k => k.Contains("2.R$Sales") && k.EndsWith("INDEX.0.idf")), "index table of the active date relationship at version 2");
            }
            var (two, _, dates2, sales2, _) = InactiveModel();
            using (two)
            {
                // A calculated column on the many-side table is supported (H82): the cube 28 (25 and three), the relationship stays unprocessed.
                sales2.AddCalculatedColumn("Line", "[Qty]*2", ColumnDataType.Int64);
                var withColumn = SaveKpiModel(two);
                PerspectiveText(withColumn, @"Model\.28\.cub\.xml$");
                PerspectiveText(withColumn, @"MdxScript\.12\.scr\.xml$");
                string salesDim = PerspectiveText(withColumn, @"Sales_[^\\]+\.8\.dim\.xml$");
                Check.True(salesDim.Contains("<ddl300:ProcessingState>Unprocessed</ddl300:ProcessingState><ddl300_300:FromRelationshipEnd>"), "the inactive relationship is unprocessed");
                Check.Equal(2, withColumn.Files.Keys.Count(k => k.Contains("INDEX.0.idf") && k.Contains("R$Sales")), "index tables of the active relationships only");
                // ... but not on another table.
                dates2.AddCalculatedColumn("Year Text", "\"Y\" & [Year]", ColumnDataType.String);
                Throws<InvalidOperationException>(() => two.SaveModelPart(new MemoryStream()), "calculated column on the date table", "many-side table");
            }
            // The only other relationship pointing to the same table as the inactive one: nothing to compare with in Excel's samples.
            var flipped = new DataModelWriter(new DataModelOptions { XpressCompression = false });
            var d = flipped.CreateTable("Dates", new[] { new ColumnDefinition("Date", ColumnDataType.DateTime), new ColumnDefinition("Year", ColumnDataType.Int64), new ColumnDefinition("Month", ColumnDataType.Int64) });
            d.AddRow(new DateTime(2025, 1, 1), 2025, 1);
            var s = flipped.CreateTable("Sales", new[] { new ColumnDefinition("Date", ColumnDataType.DateTime), new ColumnDefinition("Ship Date", ColumnDataType.DateTime) });
            s.AddRow(new DateTime(2025, 1, 1), new DateTime(2025, 1, 1));
            flipped.AddRelationship(s["Date"], d["Date"]);
            flipped.AddRelationship(s["Ship Date"], d["Date"], isActive: false);
            d.AddHierarchy("Calendar", "Year", "Month");
            s.AddMeasure("Rows", "COUNTROWS(Sales)");
            using (flipped) Throws<InvalidOperationException>(() => flipped.SaveModelPart(new MemoryStream()), "no other active relationship of the table", "needs another active relationship");
        }

        public static void PerspectiveOne()
        {
            var (model, products, sales, total) = PerspectiveModel();
            using (model)
            {
                var view = model.AddPerspective("Sales view");
                Check.True(ReferenceEquals(view, model.Perspectives["sales VIEW"]), "found by name, ignoring case");
                Check.True(ReferenceEquals(view, view.Include(products.Hierarchies[0]).Include(sales["Qty"]).Include(sales["Amount"]).Include(total)), "Include returns the perspective");
                view.Include(sales["Qty"]);
                Check.Equal(2, view.Columns.Count, "including a column twice does nothing");
                Check.Equal("Sales view (2 columns, 1 measure, 1 hierarchy)", view.ToString(), "ToString");
                var m = SaveKpiModel(model);
                // H64: the cube and the measure groups at 20, the partitions and the script at 10, Products saved again at 6, Sales' partition information only.
                string cube = PerspectiveText(m, @"Model\.20\.cub\.xml$");
                Check.True(cube.Contains("<ObjectVersion>20</ObjectVersion>"), "cube version");
                Check.True(cube.Contains("<PerspectiveFileList>Sales view.0.persp.xml</PerspectiveFileList>"), "perspective file list");
                Check.True(cube.Contains("<AssemblyFileList>MdxScript.10.scr.xml</AssemblyFileList>"), "script file");
                Check.True(cube.Contains("<StructureVersion>11</StructureVersion>"), "cube structure as before");
                foreach (var table in new[] { "Products", "Sales" })
                {
                    PerspectiveText(m, table + @"_G\.20\.det\.xml$");
                    PerspectiveText(m, table + @"_G\.10\.prt\.xml$");
                }
                PerspectiveText(m, @"Products_G\.6\.dim\.xml$");
                PerspectiveText(m, @"Sales_G\.3\.dim\.xml$");
                PerspectiveText(m, @"Sales_G\.0\.prt\\info\.10\.xml$");
                Check.Equal(0, m.Files.Keys.Count(k => ModelReader.Normalize(k).Contains(@"Products_G.0.prt\info")), "no partition information for Products");
                Check.True(PerspectiveText(m, @"Products_G\.10\.prt\.xml$").Contains("<DataVersion>-1</DataVersion>"), "Products partition data version");
                Check.True(PerspectiveText(m, @"Sales_G\.10\.prt\.xml$").Contains("<DataVersion>1</DataVersion>"), "Sales partition data version");
                Check.True(PerspectiveText(m, @"Sales_G\.20\.det\.xml$").Contains("<CacheVersion>3</CacheVersion>"), "Sales measure group cache version");
                Check.True(PerspectiveText(m, @"Products_G\.20\.det\.xml$").Contains("<CacheVersion>2</CacheVersion>"), "Products measure group cache version");
                Check.True(PerspectiveText(m, @"Products_G\.20\.det\.xml$").Contains("<Ordinal>4</Ordinal>"), "measure group ordinals follow the three edits");
                Check.True(PerspectiveText(m, @"MdxScript\.10\.scr\.xml$").Contains("<Ordinal>2</Ordinal>"), "script ordinal");
                Check.True(!PerspectiveText(m, @"Products_G\.6\.dim\.xml$").Contains("<Value/>"), "Products is saved again: empty annotation values dropped");
                // The perspective file.
                string file = PerspectiveText(m, @"Model\.0\.cub\\Sales view\.0\.persp\.xml$");
                Check.True(file.Contains("<Name>Sales view</Name><ID>Sales view</ID>"), "name and id");
                Check.True(file.Contains("<ObjectVersion>0</ObjectVersion>") && file.Contains("<Ordinal>0</Ordinal>"), "version and ordinal");
                Check.True(file.Contains("<Hierarchies/>"), "hierarchies are not listed in the file");
                Check.Equal("Amount;Qty", string.Join(";", PerspectiveAttributes(file, "Sales")), "attributes of Sales");
                Check.Equal("", string.Join(";", PerspectiveAttributes(file, "Products")), "Products has no attributes, only its hierarchy");
                Check.True(file.Contains("<Calculation><Name>[Measures].[Total Qty]</Name><Type>Member</Type></Calculation>"), "the measure is a calculation");
                Check.True(System.Text.RegularExpressions.Regex.IsMatch(file, "<MeasureGroups><MeasureGroup><MeasureGroupID>Sales_[0-9a-f-]+</MeasureGroupID><Measures/></MeasureGroup></MeasureGroups>"), "the measure group of Sales");
                Check.Equal("<?xml version=\"1.0\" encoding=\"UTF-16\"?><Perspectives><Perspective Name=\"Sales view\"><Tables><Table Name=\"Products\" EmptyHierarchyNames=\"Product Hierarchy\"/></Tables></Perspective></Perspectives>",
                    PerspectiveAnnotation(m), "the hierarchy is in the cube annotation");
                // The backup log lists the perspective between the partitions and the script, as class 100054.
                var classes = m.Log.Descendants("FileGroup").Select(g => (string)g.Element("Class")!).ToList();
                int at = classes.IndexOf("100054");
                Check.True(at > 0 && classes[at - 1] == "100021" && classes[at + 1] == "100060", "log order: partition, perspective, script");
            }
        }

        public static void PerspectiveThree()
        {
            var (model, products, sales, total) = PerspectiveModel();
            using (model)
            {
                // H65: a perspective without a measure lists the hidden count measure; one without a column has no dimension.
                model.AddPerspective("1 - subset of columns").Include(sales["Amount"]).Include(products["Product"]).Include(products["ProductKey"]);
                model.AddPerspective("2 - measure and hierarchy").Include(products.Hierarchies[0]).Include(total);
                model.AddPerspective("3 - both").Include(products["Product"]).Include(products["ProductKey"]).Include(products.Hierarchies[0]).Include(total);
                var m = SaveKpiModel(model);
                // The counters do not depend on the number of perspectives.
                PerspectiveText(m, @"Model\.20\.cub\.xml$");
                PerspectiveText(m, @"MdxScript\.10\.scr\.xml$");
                string one = PerspectiveText(m, @"\\1 - subset of columns\.0\.persp\.xml$");
                string two = PerspectiveText(m, @"\\2 - measure and hierarchy\.0\.persp\.xml$");
                string three = PerspectiveText(m, @"\\3 - both\.0\.persp\.xml$");
                Check.True(one.Contains("<Ordinal>0</Ordinal>") && two.Contains("<Ordinal>1</Ordinal>") && three.Contains("<Ordinal>2</Ordinal>"), "ordinals follow the order added");
                Check.True(one.Contains("<Name>[Measures].[__XL_Count of Models]</Name>") && one.Contains("<MeasureGroups/>"), "no measure: the count measure, no measure group");
                Check.True(two.Contains("<Dimensions/>"), "only a hierarchy and a measure: no dimension");
                Check.Equal("Product;ProductKey", string.Join(";", PerspectiveAttributes(three, "Products")), "attributes of perspective 3");
                Check.Equal("Amount", string.Join(";", PerspectiveAttributes(one, "Sales")), "attributes of Sales in perspective 1");
                // Only the perspectives that show a hierarchy are in the annotation (the first one is not).
                var annotation = PerspectiveAnnotation(m)!;
                Check.True(!annotation.Contains("1 - subset of columns"), "perspective 1 has no hierarchy");
                Check.True(annotation.Contains("<Perspective Name=\"2 - measure and hierarchy\">") && annotation.Contains("<Perspective Name=\"3 - both\">"), "perspectives 2 and 3");
                var list = System.Text.RegularExpressions.Regex.Match(PerspectiveText(m, @"Model\.20\.cub\.xml$"), "<PerspectiveFileList>(.*?)</PerspectiveFileList>").Groups[1].Value.Split(';');
                Check.Equal(3, list.Length, "three files listed");
                foreach (var name in new[] { "1 - subset of columns", "2 - measure and hierarchy", "3 - both" })
                    Check.True(list.Contains(name + ".0.persp.xml"), "listed: " + name);
            }
        }

        public static void PerspectiveHierarchiesAndTables()
        {
            var (model, products, sales, total) = PerspectiveModel();
            using (model)
            {
                var hidden = sales.AddMeasure("Hidden Total", "SUM(Sales[Amount])");
                hidden.IsHidden = true;
                // H66 A: a hierarchy with every level ticked is still only in the annotation.
                model.AddPerspective("A - all levels").Include(products["Category"]).Include(products["Product"]).Include(products.Hierarchies[0]);
                // H66 B: a table ticks all its columns, its visible measures and its hierarchies.
                model.AddPerspective("B - whole table").Include(products);
                model.AddPerspective("C - sales").Include(sales);
                var m = SaveKpiModel(model);
                string a = PerspectiveText(m, @"\\A - all levels\.0\.persp\.xml$");
                Check.True(a.Contains("<Hierarchies/>"), "no hierarchy in the file, all levels ticked");
                Check.Equal("Category;Product", string.Join(";", PerspectiveAttributes(a, "Products")), "levels as attributes");
                string b = PerspectiveText(m, @"\\B - whole table\.0\.persp\.xml$");
                Check.Equal("Category;Product;ProductKey", string.Join(";", PerspectiveAttributes(b, "Products")), "whole table: every column");
                string c = PerspectiveText(m, @"\\C - sales\.0\.persp\.xml$");
                Check.Equal("Amount;ProductKey;Qty", string.Join(";", PerspectiveAttributes(c, "Sales")), "the Sales columns");
                Check.True(c.Contains("[Measures].[Total Qty]") && !c.Contains("Hidden Total"), "the table's visible measures, not the hidden one");
                var annotation = PerspectiveAnnotation(m)!;
                Check.True(annotation.Contains("<Perspective Name=\"A - all levels\"><Tables><Table Name=\"Products\" EmptyHierarchyNames=\"Product Hierarchy\"/></Tables></Perspective>"), "A");
                Check.True(annotation.Contains("<Perspective Name=\"B - whole table\"><Tables><Table Name=\"Products\" EmptyHierarchyNames=\"Product Hierarchy\"/></Tables></Perspective>"), "B");
                Check.True(!annotation.Contains("C - sales"), "Sales has no hierarchy");
            }
        }

        public static void PerspectiveAbsent()
        {
            // Without a perspective the model keeps the hierarchy profile and an empty perspective list.
            var (model, _, _, _) = PerspectiveModel();
            using (model)
            {
                var m = SaveKpiModel(model);
                string cube = PerspectiveText(m, @"Model\.18\.cub\.xml$");
                Check.True(cube.Contains("<PerspectiveFileList/>"), "empty perspective list");
                Check.True(!cube.Contains("PerspectiveEmptyHierarchyNames"), "no annotation");
                PerspectiveText(m, @"MdxScript\.9\.scr\.xml$");
                Check.Equal(0, m.Files.Keys.Count(k => k.Contains(".persp.xml")), "no perspective files");
                Check.True(m.Log.Descendants("FileGroup").All(g => (string)g.Element("Class")! != "100054"), "no perspective in the log");
            }
        }

        public static void PerspectiveValidation()
        {
            var (model, products, sales, total) = PerspectiveModel();
            using (model)
            {
                Throws<ArgumentException>(() => model.AddPerspective(" "), "blank name", "required");
                Throws<ArgumentException>(() => model.AddPerspective(null!), "null name", "required");
                Throws<ArgumentException>(() => model.AddPerspective("a/b"), "slash in the name", "must not contain");
                Throws<ArgumentException>(() => model.AddPerspective("a|b"), "bar in the name", "'a|b'");
                Throws<ArgumentException>(() => model.AddPerspective(" padded"), "leading space", "start or end with a space");
                Throws<ArgumentException>(() => model.AddPerspective("model"), "the model's own name", "'Model'");
                var view = model.AddPerspective("Sales view");
                Throws<ArgumentException>(() => model.AddPerspective("SALES VIEW"), "duplicate name", "already exists");
                Throws<KeyNotFoundException>(() => _ = model.Perspectives["Nope"], "unknown perspective", "Sales view");
                Throws<ArgumentNullException>(() => view.Include((ModelColumn)null!), "null column");
                Throws<ArgumentNullException>(() => view.Include((ModelMeasure)null!), "null measure");
                Throws<ArgumentNullException>(() => view.Include((ModelHierarchy)null!), "null hierarchy");
                Throws<ArgumentNullException>(() => view.Include((ModelTable)null!), "null table");
                using (var foreign = new DataModelWriter())
                {
                    var t = foreign.AddTable("Sales", new[] { new { Qty = 1L } });
                    var fm = t.AddMeasure("X", "1");
                    Throws<ArgumentException>(() => view.Include(t["Qty"]), "column of another model", "another DataModelWriter");
                    Throws<ArgumentException>(() => view.Include(fm), "measure of another model", "another DataModelWriter");
                    Throws<ArgumentException>(() => view.Include(t), "table of another model", "another DataModelWriter");
                }
                // Save: a perspective must show something.
                Throws<InvalidOperationException>(() => model.SaveModelPart(new MemoryStream()), "empty perspective", "shows nothing");
                view.Include(total);
                model.SaveModelPart(new MemoryStream());
                // Two hierarchies of one table: Excel's samples show one per table.
                products.AddHierarchy("Other", "Product", "Category");
                view.Include(products.Hierarchies[0]).Include(products.Hierarchies[1]);
                Throws<InvalidOperationException>(() => model.SaveModelPart(new MemoryStream()), "two hierarchies of one table", "one hierarchy per table");
            }
            // Shapes no Excel sample shows are refused with a message that says what the model lacks.
            var (plain, _, plainSales) = CalcModel(hierarchy: false);
            using (plain)
            {
                plainSales.AddMeasure("Total Qty", "SUM(Sales[Qty])");
                plain.AddPerspective("P").Include(plainSales["Qty"]);
                Throws<InvalidOperationException>(() => plain.SaveModelPart(new MemoryStream()), "no hierarchy", "needs the model to have a hierarchy");
            }
            var (noMeasure, _, noMeasureSales) = CalcModel();
            using (noMeasure)
            {
                noMeasure.AddPerspective("P").Include(noMeasureSales["Qty"]);
                Throws<InvalidOperationException>(() => noMeasure.SaveModelPart(new MemoryStream()), "no measure", "needs the model to have a measure");
            }
            using (var single = new DataModelWriter())
            {
                var only = single.AddTable("Sales", new[] { new { Category = "A", Product = "P1", Qty = 1L } });
                only.AddHierarchy("H", "Category", "Product");
                single.AddPerspective("P").Include(only.AddMeasure("Total Qty", "SUM(Sales[Qty])"));
                Throws<InvalidOperationException>(() => single.SaveModelPart(new MemoryStream()), "no relationship", "needs the model to have a relationship");
            }
            var (hiddenModel, _, hiddenSales, hiddenTotal) = PerspectiveModel();
            using (hiddenModel)
            {
                hiddenModel.AddPerspective("P").Include(hiddenTotal);
                hiddenSales["Amount"].IsHidden = true;
                Throws<InvalidOperationException>(() => hiddenModel.SaveModelPart(new MemoryStream()), "hidden column", "hidden columns");
            }
            var (calcModel, _, calcSales, calcTotal) = PerspectiveModel();
            using (calcModel)
            {
                calcModel.AddPerspective("P").Include(calcTotal);
                calcSales.AddCalculatedColumn("Line Total", "[Qty]*[Amount]", ColumnDataType.Currency);
                Throws<InvalidOperationException>(() => calcModel.SaveModelPart(new MemoryStream()), "calculated column", "calculated columns");
            }
            var (kpiModel, _, _, kpiTotal) = PerspectiveModel();
            using (kpiModel)
            {
                kpiModel.AddPerspective("P").Include(kpiTotal);
                kpiTotal.AddKpi(1000);
                Throws<InvalidOperationException>(() => kpiModel.SaveModelPart(new MemoryStream()), "KPI", "KPIs");
            }
            // A disposed model refuses new perspectives.
            var (gone, _, _, _) = PerspectiveModel();
            gone.Dispose();
            Throws<ObjectDisposedException>(() => gone.AddPerspective("P"), "disposed model");
        }

        // ---- KPIs (SPEC section 12; Excel's H47, H48, H49) ----

        private const string MeasuresHeader = "----------------------------------------------------------\n-- PowerPivot measures command (do not modify manually) --\n----------------------------------------------------------\n\n\n";

        private static (DataModelWriter model, ModelMeasure total) KpiModel(bool hierarchy = true)
        {
            var (model, _, sales) = CalcModel(hierarchy);
            return (model, sales.AddMeasure("Total Qty", "SUM(Sales[Qty])", MeasureFormat.WholeNumber));
        }

        private static ModelReader SaveKpiModel(DataModelWriter model)
        {
            var ms = new MemoryStream();
            model.SaveModelPart(ms);
            return ModelReader.Read(ms.ToArray());
        }

        // Texts and annotations of the script, parsed so that the assertions see what Excel sees.
        private static (List<string> commands, Dictionary<string, Dictionary<string, string>> annotations, List<string> references) KpiScriptParts(ModelReader m)
        {
            var doc = new System.Xml.XmlDocument();
            doc.LoadXml(Encoding.UTF8.GetString(m.Find(@"MdxScript\.\d+\.scr\.xml$")).TrimStart('\uFEFF'));
            var commands = doc.GetElementsByTagName("Command").Cast<System.Xml.XmlNode>().Select(c => c["Text"]!.InnerText).ToList();
            var annotations = new Dictionary<string, Dictionary<string, string>>();
            var references = new List<string>();
            foreach (System.Xml.XmlNode cp in doc.GetElementsByTagName("CalculationProperty"))
            {
                string reference = cp["CalculationReference"]!.InnerText;
                references.Add(reference);
                var map = new Dictionary<string, string>();
                foreach (System.Xml.XmlNode a in cp["Annotations"]?.ChildNodes ?? doc.CreateElement("x").ChildNodes) map[a["Name"]!.InnerText] = a["Value"]!.InnerText.Trim();
                annotations[reference] = map;
            }
            return (commands, annotations, references);
        }

        public static void KpiScripts()
        {
            // H47: a target of 1000, the default thresholds (400 and 800) and Traffic Light - Single.
            {
                var (model, total) = KpiModel();
                using (model)
                {
                    var kpi = total.AddKpi(1000);
                    Check.True(ReferenceEquals(kpi, total.Kpi), "Kpi property");
                    Check.Equal("KPI of 'Sales'[Total Qty]", kpi.ToString(), "ToString");
                    var (commands, ann, refs) = KpiScriptParts(SaveKpiModel(model));
                    Check.Equal(2, commands.Count, "header command and the measure's");
                    Check.Equal(MeasuresHeader + "CREATE MEASURE 'Sales'[Total Qty]=SUM(Sales[Qty]);\nCREATE MEASURE 'Sales'[_Total Qty Goal] = 1000; \nCREATE MEASURE 'Sales'[_Total Qty Status] = if(ISBLANK('Sales'[Total Qty]),BLANK(),\n                    If('Sales'[Total Qty]<400,-1,\n\t                If('Sales'[Total Qty]<800,0,1)\n                )\n            )\n            ; \nCREATE KPI CURRENTCUBE.[Total Qty] AS Measures.[Total Qty], ASSOCIATED_MEASURE_GROUP = 'Sales', GOAL = Measures.[_Total Qty Goal], STATUS = Measures.[_Total Qty Status], STATUS_GRAPHIC = 'Traffic Light - Single';\n",
                        commands[1], "H47 command");
                    Check.Equal("[Total Qty]|[_Total Qty Goal]|[_Total Qty Status]|KPIs.[Total Qty]|Measures.[__XL_Count of Models]", string.Join("|", refs), "calculation properties");
                    var a = ann["[Total Qty]"];
                    Check.Equal("Kpi", a["Type"], "Type");
                    Check.Equal("StaticValue", a["GoalType"], "GoalType");
                    Check.Equal("1000", a["GoalValue"], "GoalValue");
                    Check.Equal("Linear", a["KpiStatusType"], "KpiStatusType");
                    Check.Equal("Absolute", a["KpiThresholdType"], "KpiThresholdType");
                    Check.Equal("Ascending", a["KpiThresholdOrdering"], "KpiThresholdOrdering");
                    Check.Equal("2", a["KpiThresholdCount"], "KpiThresholdCount");
                    Check.Equal("400", a["KpiThreshold_0"], "threshold 0");
                    Check.Equal("800", a["KpiThreshold_1"], "threshold 1");
                    Check.True(!a.ContainsKey("GoalMeasure"), "no goal measure");
                    Check.Equal("SupportMdx", ann["[_Total Qty Goal]"]["Type"], "goal support measure");
                    Check.Equal("Total Qty", ann["[_Total Qty Status]"]["MainObjectName"], "status support measure");
                    Check.Equal("SupportKpi", ann["KPIs.[Total Qty]"]["Type"], "KPI support entry");
                }
            }
            // H48: the target is the measure Target Qty; value / target against fractions, thresholds stored as percentages.
            {
                var (model, total) = KpiModel();
                using (model)
                {
                    var target = model.Tables["Sales"].AddMeasure("Target Qty", "1000");
                    var kpi = total.AddKpi(target);
                    Check.True(ReferenceEquals(target, kpi.TargetMeasure) && kpi.Target is null, "target measure");
                    var (commands, ann, refs) = KpiScriptParts(SaveKpiModel(model));
                    Check.Equal(3, commands.Count, "header, Target Qty and Total Qty");
                    Check.Equal(MeasuresHeader + "CREATE MEASURE 'Sales'[Target Qty]=1000;\n", commands[1], "measures are listed by name");
                    Check.Equal(MeasuresHeader + "CREATE MEASURE 'Sales'[Total Qty]=SUM(Sales[Qty]);\nCREATE MEASURE 'Sales'[_Total Qty Goal] = 'Sales'[Target Qty]; \nCREATE MEASURE 'Sales'[_Total Qty Status] = if(ISBLANK('Sales'[Total Qty]/'Sales'[_Total Qty Goal]),BLANK(),\n                    If('Sales'[Total Qty]/'Sales'[_Total Qty Goal]<0.4,-1,\n\t                If('Sales'[Total Qty]/'Sales'[_Total Qty Goal]<0.8,0,1)\n                )\n            )\n            ; \nCREATE KPI CURRENTCUBE.[Total Qty] AS Measures.[Total Qty], ASSOCIATED_MEASURE_GROUP = 'Sales', GOAL = Measures.[_Total Qty Goal], STATUS = Measures.[_Total Qty Status], STATUS_GRAPHIC = 'Traffic Light - Single';\n",
                        commands[2], "H48 command");
                    Check.Equal("[Target Qty]|[Total Qty]|[_Total Qty Goal]|[_Total Qty Status]|KPIs.[Total Qty]|Measures.[__XL_Count of Models]", string.Join("|", refs), "calculation properties");
                    var a = ann["[Total Qty]"];
                    Check.Equal("Measure", a["GoalType"], "GoalType");
                    Check.True(!a.ContainsKey("GoalValue"), "no fixed goal");
                    Check.Equal("Percentage", a["KpiThresholdType"], "KpiThresholdType");
                    Check.Equal("40", a["KpiThreshold_0"], "threshold 0");
                    Check.Equal("80", a["KpiThreshold_1"], "threshold 1");
                    Check.Equal("Target Qty", a["GoalMeasure"], "GoalMeasure");
                    Check.Equal("Sales", a["GoalMeasureTableName"], "GoalMeasureTableName");
                }
            }
            // H49: the Centered scale (five bands) with the Road Signs icons; the defaults are 340, 670, 1330 and 1660 for a target of 1000.
            {
                var (model, total) = KpiModel();
                using (model)
                {
                    var kpi = total.AddKpi(1000, new KpiOptions { Scale = KpiScale.Centered, IconStyle = KpiIconStyle.RoadSigns });
                    Check.Equal("Road Signs", kpi.IconStyle.Name, "icon style");
                    var (commands, ann, _) = KpiScriptParts(SaveKpiModel(model));
                    Check.Equal(MeasuresHeader + "CREATE MEASURE 'Sales'[Total Qty]=SUM(Sales[Qty]);\nCREATE MEASURE 'Sales'[_Total Qty Goal] = 1000; \nCREATE MEASURE 'Sales'[_Total Qty Status] = if(ISBLANK('Sales'[Total Qty]),BLANK(),\n                If('Sales'[Total Qty]<670,\n\t                If('Sales'[Total Qty]<340,-1,0),\n\t                If('Sales'[Total Qty]<1330,1,\n\t\t                If('Sales'[Total Qty]<1660,0,-1)\n\t                )\n                )\n            )\n            ; \nCREATE KPI CURRENTCUBE.[Total Qty] AS Measures.[Total Qty], ASSOCIATED_MEASURE_GROUP = 'Sales', GOAL = Measures.[_Total Qty Goal], STATUS = Measures.[_Total Qty Status], STATUS_GRAPHIC = 'Road Signs';\n",
                        commands[1], "H49 command");
                    var a = ann["[Total Qty]"];
                    Check.Equal("Centered", a["KpiStatusType"], "KpiStatusType");
                    Check.Equal("4", a["KpiThresholdCount"], "KpiThresholdCount");
                    Check.Equal("340|670|1330|1660", string.Join("|", Enumerable.Range(0, 4).Select(i => a["KpiThreshold_" + i])), "thresholds");
                }
            }
            // Own thresholds are fractions of the target.
            {
                var (model, total) = KpiModel();
                using (model)
                {
                    total.AddKpi(2500, new KpiOptions { Thresholds = new[] { 0.5, 0.9 } });
                    var (commands, ann, _) = KpiScriptParts(SaveKpiModel(model));
                    Check.Equal("1250", ann["[Total Qty]"]["KpiThreshold_0"], "0.5 of 2500");
                    Check.Equal("2250", ann["[Total Qty]"]["KpiThreshold_1"], "0.9 of 2500");
                    Check.True(commands[1].Contains("= 2500; \n") && commands[1].Contains("<1250,-1,") && commands[1].Contains("<2250,0,1)"), "script");
                }
            }
        }

        // H59: the target measure of the KPI is hidden. The markup is that of any hidden measure; the hide is one more script edit (12, H48 has 11).
        public static void KpiHiddenTarget()
        {
            var (model, total) = KpiModel();
            using (model)
            {
                var target = model.Tables["Sales"].AddMeasure("Target Qty", "1000");
                target.IsHidden = true;
                total.AddKpi(target);
                var m = SaveKpiModel(model);
                var (commands, ann, refs) = KpiScriptParts(m);
                Check.Equal("True", ann["[Target Qty]"]["IsPrivate"], "target hidden");
                Check.True(Encoding.UTF8.GetString(m.Find(@"MdxScript\.12\.scr\.xml$")).Contains("<CalculationReference>[Target Qty]</CalculationReference><CalculationType>Member</CalculationType><Description/><Visible>false</Visible>"), "Visible=false as in H59");
                Check.Equal("False", ann["[Total Qty]"]["IsPrivate"], "KPI measure visible");
                Check.True(commands[2].Contains("_Total Qty Goal] = 'Sales'[Target Qty]; "), "goal refers to the hidden target");
                Check.Equal(6, refs.Count, "calculation properties");
                Encoding.UTF8.GetString(m.Find(@"MdxScript\.12\.scr\.xml$"));
                Encoding.UTF8.GetString(m.Find(@"Model\.19\.cub\.xml$"));
            }
        }

        // H61: one KPI per icon style Excel offers. The four three-state styles are written like H47; the gauge has five states (-2 to 2),
        // four thresholds and the status type Linear.
        public static void KpiIconStyles()
        {
            var (model, total) = KpiModel();
            using (model)
            {
                var sales = model.Tables["Sales"];
                for (int i = 0; i < 4; i++)
                    sales.AddMeasure("Style 0" + (i + 1), (i + 1).ToString()).AddKpi(1000, new KpiOptions { IconStyle = KpiIconStyle.All[i] });
                var gauge = sales.AddMeasure("Style 05", "5").AddKpi(1000, new KpiOptions { IconStyle = KpiIconStyle.GaugeAscending });
                Check.Equal("0.2,0.4,0.6,0.8", string.Join(",", gauge.Thresholds.Select(x => x.ToString(System.Globalization.CultureInfo.InvariantCulture))), "gauge defaults");
                var (commands, ann, _) = KpiScriptParts(SaveKpiModel(model));
                string[] graphics = { "Traffic Light - Single", "Road Signs", "Traffic Light", "Shapes", "Gauge - Ascending" };
                for (int i = 0; i < 5; i++)
                    Check.True(commands.Any(c => c.Contains($"CREATE KPI CURRENTCUBE.[Style 0{i + 1}] AS") && c.Contains($"STATUS_GRAPHIC = '{graphics[i]}';")), "graphic " + graphics[i]);
                Check.Equal("Linear", ann["[Style 05]"]["KpiStatusType"], "the gauge layout is Linear");
                Check.Equal("4", ann["[Style 05]"]["KpiThresholdCount"], "four thresholds");
                Check.Equal("200|400|600|800", string.Join("|", Enumerable.Range(0, 4).Select(i => ann["[Style 05]"]["KpiThreshold_" + i])), "thresholds");
                Check.Equal("2", ann["[Style 04]"]["KpiThresholdCount"], "three-state styles keep two thresholds");
                string gaugeCommand = commands.First(c => c.Contains("[Style 05]"));
                Check.True(gaugeCommand.Contains("CREATE MEASURE 'Sales'[_Style 05 Status] = if(ISBLANK('Sales'[Style 05]),BLANK(),\n                If('Sales'[Style 05]<400,\n\t                If('Sales'[Style 05]<200,-2,-1),\n\t                If('Sales'[Style 05]<600,0,\n\t\t                If('Sales'[Style 05]<800,1,2)\n\t                )\n                )\n            )\n            ; \n"), "H61 gauge status");
            }
        }

        // H63: a gauge from scratch with own thresholds (110, 475, 710, 1030 of 1000); the thresholds are the fractions times the target.
        public static void KpiCustomGaugeThresholds()
        {
            using var model = new DataModelWriter(new DataModelOptions { XpressCompression = false });
            var t = model.CreateTable("Table1", new[] { new ColumnDefinition("Value", ColumnDataType.Int64) });
            for (int i = 1; i <= 5; i++) t.AddRow(i * 100);
            t.AddMeasure("Val", "SUM(Table1[Value])").AddKpi(1000, new KpiOptions { IconStyle = KpiIconStyle.GaugeAscending, Thresholds = new[] { 0.11, 0.475, 0.71, 1.03 } });
            var m = SaveKpiModel(model);
            var (commands, ann, _) = KpiScriptParts(m);
            Check.Equal(MeasuresHeader + "CREATE MEASURE 'Table1'[Val]=SUM(Table1[Value]);\nCREATE MEASURE 'Table1'[_Val Goal] = 1000; \nCREATE MEASURE 'Table1'[_Val Status] = if(ISBLANK('Table1'[Val]),BLANK(),\n                If('Table1'[Val]<475,\n\t                If('Table1'[Val]<110,-2,-1),\n\t                If('Table1'[Val]<710,0,\n\t\t                If('Table1'[Val]<1030,1,2)\n\t                )\n                )\n            )\n            ; \nCREATE KPI CURRENTCUBE.[Val] AS Measures.[Val], ASSOCIATED_MEASURE_GROUP = 'Table1', GOAL = Measures.[_Val Goal], STATUS = Measures.[_Val Status], STATUS_GRAPHIC = 'Gauge - Ascending';\n",
                commands[1], "H63 command");
            Check.Equal("110|475|710|1030", string.Join("|", Enumerable.Range(0, 4).Select(i => ann["[Val]"]["KpiThreshold_" + i])), "thresholds");
            // Excel's order of the annotations: two thresholds, the descriptions, the other thresholds.
            string script = Encoding.UTF8.GetString(m.Find(@"MdxScript\.\d+\.scr\.xml$"));
            var order = System.Text.RegularExpressions.Regex.Matches(script, @"<Name>(KpiThreshold_\d|\w+Description)</Name>").Select(x => x.Groups[1].Value).ToList();
            Check.Equal("KpiThreshold_0|KpiThreshold_1|ValueDescription|GoalDescription|StatusDescription|TrendDescription|KpiThreshold_2|KpiThreshold_3", string.Join("|", order), "annotation order as in H49, H61, H62 and H63");
        }

        // H62: the four layouts not seen before, on measures of the Products table. CenteredReversed with three-state icons negates the Centered values;
        // the gauge reversed negates its linear ones; the centered gauge has eight thresholds and nine bands (-2 up to 2 and back).
        public static void KpiLayouts()
        {
            var (model, _) = KpiModel();
            using (model)
            {
                var products = model.Tables["Products"];
                products.AddMeasure("Layout A", "1").AddKpi(1000, new KpiOptions { Scale = KpiScale.CenteredReversed });
                products.AddMeasure("Layout B", "1").AddKpi(1000, new KpiOptions { Scale = KpiScale.LinearReversed, IconStyle = KpiIconStyle.GaugeAscending });
                products.AddMeasure("Layout C", "1").AddKpi(1000, new KpiOptions { Scale = KpiScale.Centered, IconStyle = KpiIconStyle.GaugeAscending });
                var d = products.AddMeasure("Layout D", "1").AddKpi(1000, new KpiOptions { Scale = KpiScale.CenteredReversed, IconStyle = KpiIconStyle.GaugeAscending });
                Check.Equal("0.2,0.4,0.6,0.8,1.2,1.4,1.6,1.8", string.Join(",", d.Thresholds.Select(x => x.ToString(System.Globalization.CultureInfo.InvariantCulture))), "centered gauge defaults");
                var (commands, ann, _) = KpiScriptParts(SaveKpiModel(model));
                string StatusOf(string name) { string c = commands.First(c => c.Contains("CREATE KPI CURRENTCUBE.[" + name + "]")); int i = c.IndexOf("Status] =", StringComparison.Ordinal); return c[i..c.IndexOf("CREATE KPI", StringComparison.Ordinal)]; }
                Check.Equal("Status] = if(ISBLANK('Products'[Layout A]),BLANK(),\n                If('Products'[Layout A]<670,\n\t                If('Products'[Layout A]<340,1,0),\n\t                If('Products'[Layout A]<1330,-1,\n\t\t                If('Products'[Layout A]<1660,0,1)\n\t                )\n                )\n            )\n            ; \n", StatusOf("Layout A"), "H62 layout A");
                Check.Equal("Status] = if(ISBLANK('Products'[Layout B]),BLANK(),\n                If('Products'[Layout B]<400,\n\t                If('Products'[Layout B]<200,2,1),\n\t                If('Products'[Layout B]<600,0,\n\t\t                If('Products'[Layout B]<800,-1,-2)\n\t                )\n                )\n            )\n            ; \n", StatusOf("Layout B"), "H62 layout B");
                Check.Equal("Status] = if(ISBLANK('Products'[Layout D]),BLANK(),\n                If('Products'[Layout D]<800,\n\t                If('Products'[Layout D]<400,\n\t\t                If('Products'[Layout D]<200,2,1),\n\t\t                    If('Products'[Layout D]<600,0,-1)\n\t                    ),\n\t                    If('Products'[Layout D]<1400,\n\t\t                    If('Products'[Layout D]<1200,-2,-1),\n\t\t                        If('Products'[Layout D]<1600,0,\n\t\t\t                        If('Products'[Layout D]<1800,1,2)\n\t\t                )\n\t                )\n                )\n            )\n            ; \n", StatusOf("Layout D"), "H62 layout D");
                Check.True(StatusOf("Layout C").Contains("If('Products'[Layout C]<200,-2,-1)") && StatusOf("Layout C").Contains("If('Products'[Layout C]<1200,2,1)") && StatusOf("Layout C").Contains("If('Products'[Layout C]<1800,-1,-2)"), "H62 layout C: the values rise to 2 and fall again");
                Check.Equal("CenteredReversed", ann["[Layout A]"]["KpiStatusType"], "type A");
                Check.Equal("LinearReversed", ann["[Layout B]"]["KpiStatusType"], "type B");
                Check.Equal("4", ann["[Layout B]"]["KpiThresholdCount"], "four thresholds for the reversed gauge");
                Check.Equal("8", ann["[Layout C]"]["KpiThresholdCount"], "eight thresholds for the centered gauge");
                Check.Equal("1800", ann["[Layout D]"]["KpiThreshold_7"], "last threshold");
            }
        }

        // Extrapolated: four KPIs with a plain measure and a hidden target beside them (`datagen ... kpimany`). The script version counts KPIs (4),
        // measures beyond the first (5) and hidden measures (1) on top of 9.
        public static void KpiMany()
        {
            var (model, total) = KpiModel();
            using (model)
            {
                var sales = model.Tables["Sales"];
                total.AddKpi(1000);
                var target = sales.AddMeasure("Target Amount", "30000");
                target.IsHidden = true;
                sales.AddMeasure("Total Amount", "SUM(Sales[Amount])").AddKpi(target, new KpiOptions { Scale = KpiScale.Centered });
                sales.AddMeasure("Sales Rows", "COUNTROWS(Sales)").AddKpi(1200, new KpiOptions { Scale = KpiScale.LinearReversed, Thresholds = new[] { 0.5, 0.9 } });
                sales.AddMeasure("Average Qty", "AVERAGE(Sales[Qty])").AddKpi(3, new KpiOptions { Scale = KpiScale.Centered, Thresholds = new[] { 0.5, 0.8, 1.2, 1.5 } });
                sales.AddMeasure("Max Amount", "MAX(Sales[Amount])");
                var m = SaveKpiModel(model);
                var (commands, ann, refs) = KpiScriptParts(m);
                Check.Equal(6, commands.Count - 1, "one command per measure");
                // Measures without a KPI first, then the KPI measures in the order created.
                Check.Equal("[Target Amount]|[Max Amount]|[Total Qty]", string.Join("|", refs.Take(3)), "plain measures first");
                Check.Equal(4, refs.Count(r => r.StartsWith("KPIs.[", StringComparison.Ordinal)), "four KPI entries");
                Check.Equal(1, refs.Count(r => r == "Measures.[__XL_Count of Models]"), "default member once");
                Check.Equal("Percentage", ann["[Total Amount]"]["KpiThresholdType"], "measure target");
                Check.Equal("Absolute", ann["[Sales Rows]"]["KpiThresholdType"], "fixed target");
                Check.Equal("600", ann["[Sales Rows]"]["KpiThreshold_0"], "0.5 of 1200");
                Check.Equal("1080", ann["[Sales Rows]"]["KpiThreshold_1"], "0.9 of 1200");
                Check.Equal("LinearReversed", ann["[Sales Rows]"]["KpiStatusType"], "reversed");
                Check.Equal("1.5", ann["[Average Qty]"]["KpiThreshold_0"], "0.5 of 3");
                Check.Equal("4.5", ann["[Average Qty]"]["KpiThreshold_3"], "1.5 of 3");
                Check.Equal(0, commands.Count(c => c.Contains("CREATE MEASURE") && c.Contains("Max Amount") && c.Contains("CREATE KPI")), "no KPI on the plain measure");
                Encoding.UTF8.GetString(m.Find(@"MdxScript\.19\.scr\.xml$"));
                Encoding.UTF8.GetString(m.Find(@"Model\.19\.cub\.xml$"));
                Encoding.UTF8.GetString(m.Find(@"Products_G\.6\.dim\.xml$"));
            }
        }

        // H60: the KPI measure itself is hidden: the markup of a hidden measure on the base measure, the script one edit further (11).
        public static void KpiHiddenMeasure()
        {
            var (model, total) = KpiModel();
            using (model)
            {
                total.AddKpi(1000);
                total.IsHidden = true;
                var m = SaveKpiModel(model);
                var (commands, ann, refs) = KpiScriptParts(m);
                Check.Equal("True", ann["[Total Qty]"]["IsPrivate"], "measure hidden");
                Check.Equal("Kpi", ann["[Total Qty]"]["Type"], "still a KPI");
                Check.True(Encoding.UTF8.GetString(m.Find(@"MdxScript\.11\.scr\.xml$")).Contains("<CalculationReference>[Total Qty]</CalculationReference><CalculationType>Member</CalculationType><Description/><Visible>false</Visible>"), "Visible=false as in H60");
                Check.Equal(5, refs.Count, "calculation properties");
                Encoding.UTF8.GetString(m.Find(@"Model\.19\.cub\.xml$"));
            }
        }

        // H58: the KPI on top of a calculated column on Products. Products holds the column, so Excel saved the next primary-side table, Dates.
        public static void KpiCalculatedOnPrimary()
        {
            using var model = new DataModelWriter(new DataModelOptions { XpressCompression = false });
            var products = model.CreateTable("Products", new[] { new ColumnDefinition("ProductKey", ColumnDataType.Int64), new ColumnDefinition("Category", ColumnDataType.String), new ColumnDefinition("Product", ColumnDataType.String) });
            for (int i = 1; i <= 6; i++) products.AddRow(i, i <= 3 ? "A" : "B", "P" + i);
            var dates = model.CreateTable("Dates", new[] { new ColumnDefinition("Date", ColumnDataType.DateTime), new ColumnDefinition("Year", ColumnDataType.Int64), new ColumnDefinition("Month", ColumnDataType.String) });
            for (int i = 0; i < 60; i++) { var d = new DateTime(2025, 1, 1).AddDays(i); dates.AddRow(d, d.Year, d.ToString("MMMM", System.Globalization.CultureInfo.InvariantCulture)); }
            var sales = model.CreateTable("Sales", new[] { new ColumnDefinition("ProductKey", ColumnDataType.Int64), new ColumnDefinition("Date", ColumnDataType.DateTime), new ColumnDefinition("Qty", ColumnDataType.Int64) });
            for (int i = 0; i < 40; i++) sales.AddRow(1 + i % 6, new DateTime(2025, 1, 1).AddDays(i % 60), 1 + i % 5);
            model.AddRelationship(sales["ProductKey"], products["ProductKey"]);
            model.AddRelationship(sales["Date"], dates["Date"]);
            products.AddHierarchy("Product Hierarchy", "Category", "Product");
            dates.AddHierarchy("Calendar", "Year", "Month");
            products.AddCalculatedColumn("Name Length", "LEN([Product])", ColumnDataType.Int64);
            sales.AddMeasure("Total Qty", "SUM(Sales[Qty])", MeasureFormat.WholeNumber).AddKpi(1000);
            var m = SaveKpiModel(model);
            string Text(string pattern) => Encoding.UTF8.GetString(m.Find(pattern));
            var (commands, _, refs) = KpiScriptParts(m);
            Check.True(commands[1].Contains("CREATE KPI CURRENTCUBE.[Total Qty]"), "same script");
            Check.Equal(5, refs.Count, "calculation properties");
            Check.True(Text(@"Model\.24\.cub\.xml$").Contains("<AssemblyFileList>MdxScript.11.scr.xml</AssemblyFileList>"), "cube 24 and script 11 (the column alone gives 23 and 10)");
            Text(@"\.cub\\info\.24\.xml$");
            Check.True(!Text(@"Dates_G\.6\.dim\.xml$").Contains("<Value/>"), "Dates is the table saved again");
            Text(@"Products_G\.9\.dim\.xml$");
            Text(@"Sales_G\.3\.dim\.xml$");
        }

        // H57: the same KPI in a one-table model without relationships; every counter one step up from that model's own base.
        public static void KpiOneTable()
        {
            using var model = new DataModelWriter(new DataModelOptions { XpressCompression = false });
            var sales = model.CreateTable("Sales", new[] { new ColumnDefinition("ProductKey", ColumnDataType.Int64), new ColumnDefinition("Qty", ColumnDataType.Int64) });
            for (int i = 0; i < 40; i++) sales.AddRow(1 + i % 6, 1 + i % 5);
            sales.AddMeasure("Total Qty", "SUM(Sales[Qty])", MeasureFormat.WholeNumber).AddKpi(1000);
            var m = SaveKpiModel(model);
            string Text(string pattern) => Encoding.UTF8.GetString(m.Find(pattern));
            var (commands, _, refs) = KpiScriptParts(m);
            Check.True(commands[1].Contains("CREATE KPI CURRENTCUBE.[Total Qty]"), "same script as with relationships");
            Check.Equal(5, refs.Count, "calculation properties");
            Check.True(Text(@"Model\.2\.cub\.xml$").Contains("<AssemblyFileList>MdxScript.2.scr.xml</AssemblyFileList>"), "cube and script version 2");
            Text(@"\.cub\\info\.2\.xml$");
            Text(@"Sales_G\.2\.det\.xml$");
            Check.True(!Text(@"Sales_G\.2\.dim\.xml$").Contains("<Value/>"), "the only table is saved again: empty annotation values dropped");
        }

        // H55 and H56: the KPI on top of a calculated column on the many-side table. Excel calculated the column while it created the
        // KPI; the library writes it uncalculated, so only the KPI step (one up from the calculated-column numbers) is expected.
        public static void KpiCalculated()
        {
            ModelReader Save(bool workbook)
            {
                var (model, total) = KpiModel();
                using (model)
                {
                    model.Tables["Sales"].AddCalculatedColumn("Line Total", "[Qty]*[Amount]", ColumnDataType.Currency);
                    total.AddKpi(1000);
                    if (!workbook) return SaveKpiModel(model);
                    string path = Path.Combine(Path.GetTempPath(), "kpi_calc_" + Guid.NewGuid().ToString("N") + ".xlsx");   // the trigger table is added when a workbook is saved
                    try { model.Save(path); return ModelReader.Read(ModelReader.ItemData(path)); }
                    finally { File.Delete(path); }
                }
            }
            foreach (bool workbook in new[] { false, true })
            {
                var m = Save(workbook);
                string Text(string pattern) => Encoding.UTF8.GetString(m.Find(pattern));
                var (commands, _, refs) = KpiScriptParts(m);
                Check.True(commands[1].Contains("CREATE KPI CURRENTCUBE.[Total Qty]"), "same script as without the column");
                Check.Equal(5, refs.Count, "calculation properties");
                Check.True(Text(@"Model\.26\.cub\.xml$").Contains("<AssemblyFileList>MdxScript.12.scr.xml</AssemblyFileList>"), "cube 26 and script 12 (the column alone gives 25 and 11)");
                Text(@"\.cub\\info\.26\.xml$");
                Text(@"Sales_G\.26\.det\.xml$");
                Text(@"Products_G\.26\.det\.xml$");
                Text(@"Products_G\.7\.dim\.xml$");
                Text(@"Sales_G\.7\.dim\.xml$");
                Check.True(Text(@"Sales_G\.7\.dim\.xml$").Contains("<Expression>[Qty]*[Amount]</Expression>"), "the column stays a formula");
                Check.Equal(workbook ? 1 : 0, m.Files.Keys.Count(k => k.Contains("PivotWrightTrigger") && k.EndsWith(".dim.xml", StringComparison.Ordinal)), "trigger table only in the workbook");
            }
        }

        // H54: the same KPI in a model with a hidden column and hierarchies; every counter one step up from that model's own base.
        public static void KpiHidden()
        {
            var (model, total) = KpiModel();
            using (model)
            {
                model.Tables["Sales"]["Qty"].IsHidden = true;
                total.AddKpi(1000);
                var m = SaveKpiModel(model);
                string Text(string pattern) => Encoding.UTF8.GetString(m.Find(pattern));
                var (commands, _, refs) = KpiScriptParts(m);
                Check.True(commands[1].Contains("CREATE KPI CURRENTCUBE.[Total Qty]"), "same script as without hidden columns");
                Check.Equal(5, refs.Count, "calculation properties");
                Check.True(Text(@"Model\.22\.cub\.xml$").Contains("<AssemblyFileList>MdxScript.10.scr.xml</AssemblyFileList>"), "cube 22 and script 10");
                Text(@"\.cub\\info\.22\.xml$");
                Text(@"Sales_G\.21\.det\.xml$");
                Text(@"Products_G\.21\.det\.xml$");
                Text(@"Products_G\.6\.dim\.xml$");
                Text(@"Sales_G\.4\.dim\.xml$");
                Check.True(!Text(@"Products_G\.6\.dim\.xml$").Contains("<Value/>"), "saved dimension drops the empty annotation values");
            }
        }

        // H53: the same KPI in a model with a date table; every counter one step up from that model's own base.
        public static void KpiDateTable()
        {
            var model = new DataModelWriter(new DataModelOptions { XpressCompression = false });
            using (model)
            {
                var products = model.CreateTable("Products", new[] { new ColumnDefinition("ProductKey", ColumnDataType.Int64), new ColumnDefinition("Category", ColumnDataType.String), new ColumnDefinition("Product", ColumnDataType.String) });
                for (int i = 1; i <= 6; i++) products.AddRow(i, i <= 3 ? "A" : "B", "P" + i);
                var dates = model.CreateTable("Dates", new[] { new ColumnDefinition("Date", ColumnDataType.DateTime), new ColumnDefinition("Year", ColumnDataType.Int64), new ColumnDefinition("Month", ColumnDataType.String) });
                for (int i = 0; i < 60; i++) { var d = new DateTime(2025, 1, 1).AddDays(i); dates.AddRow(d, d.Year, d.ToString("MMMM", System.Globalization.CultureInfo.InvariantCulture)); }
                var sales = model.CreateTable("Sales", new[] { new ColumnDefinition("ProductKey", ColumnDataType.Int64), new ColumnDefinition("Date", ColumnDataType.DateTime), new ColumnDefinition("Qty", ColumnDataType.Int64) });
                for (int i = 0; i < 40; i++) sales.AddRow(1 + i % 6, new DateTime(2025, 1, 1).AddDays(i % 60), 1 + i % 5);
                model.AddRelationship(sales["ProductKey"], products["ProductKey"]);
                model.AddRelationship(sales["Date"], dates["Date"]);
                products.AddHierarchy("Product Hierarchy", "Category", "Product");
                dates.AddHierarchy("Calendar", "Year", "Month");
                dates.MarkAsDateTable("Date");
                sales.AddMeasure("Total Qty", "SUM(Sales[Qty])", MeasureFormat.WholeNumber).AddKpi(1000);
                var m = SaveKpiModel(model);
                string Text(string pattern) => Encoding.UTF8.GetString(m.Find(pattern));
                var (commands, _, refs) = KpiScriptParts(m);
                Check.True(commands[1].Contains("CREATE KPI CURRENTCUBE.[Total Qty]"), "same script as without a date table");
                Check.Equal(5, refs.Count, "calculation properties");
                Check.True(Text(@"Model\.22\.cub\.xml$").Contains("<AssemblyFileList>MdxScript.11.scr.xml</AssemblyFileList>"), "cube 22 and script 11");
                Text(@"\.cub\\info\.22\.xml$");
                Text(@"Products_G\.7\.dim\.xml$");
                Text(@"Dates_G\.6\.dim\.xml$");
                Text(@"Sales_G\.4\.dim\.xml$");
                Text(@"Sales_G\.22\.det\.xml$");
                Text(@"Sales_G\.10\.prt\.xml$");
                Check.True(!Text(@"Products_G\.7\.dim\.xml$").Contains("<Value/>"), "saved dimension drops the empty annotation values");
            }
        }

        // H52: the same KPI in a model without hierarchies; every counter one step up from that model's own base.
        public static void KpiPlain()
        {
            var (model, total) = KpiModel(hierarchy: false);
            using (model)
            {
                total.AddKpi(1000);
                var m = SaveKpiModel(model);
                string Text(string pattern) => Encoding.UTF8.GetString(m.Find(pattern));
                var (commands, _, refs) = KpiScriptParts(m);
                Check.True(commands[1].Contains("CREATE KPI CURRENTCUBE.[Total Qty]"), "same script as with hierarchies");
                Check.Equal("[Total Qty]|[_Total Qty Goal]|[_Total Qty Status]|KPIs.[Total Qty]|Measures.[__XL_Count of Models]", string.Join("|", refs), "calculation properties");
                Check.True(Text(@"Model\.2\.cub\.xml$").Contains("<AssemblyFileList>MdxScript.2.scr.xml</AssemblyFileList>"), "cube and script version 2");
                Check.True(Text(@"Model\.2\.cub\.xml$").Contains("<StructureVersion>0</StructureVersion>"), "cube structure as before");
                Text(@"\.cub\\info\.2\.xml$");
                Text(@"Products_G\.2\.dim\.xml$");
                Text(@"Sales_G\.2\.dim\.xml$");
                Text(@"Sales_G\.2\.det\.xml$");
                Text(@"Sales_G\.1\.prt\.xml$");
                Check.True(!Text(@"Products_G\.2\.dim\.xml$").Contains("<Value/>"), "saved dimension drops the empty annotation values");
            }
        }

        // H51: lower is better. Only KpiStatusType (LinearReversed) and the status values change; thresholds and ordering stay.
        public static void KpiReversed()
        {
            var (model, total) = KpiModel();
            using (model)
            {
                total.AddKpi(1000, new KpiOptions { Scale = KpiScale.LinearReversed });
                var (commands, ann, _) = KpiScriptParts(SaveKpiModel(model));
                Check.Equal(MeasuresHeader + "CREATE MEASURE 'Sales'[Total Qty]=SUM(Sales[Qty]);\nCREATE MEASURE 'Sales'[_Total Qty Goal] = 1000; \nCREATE MEASURE 'Sales'[_Total Qty Status] = if(ISBLANK('Sales'[Total Qty]),BLANK(),\n                    If('Sales'[Total Qty]<400,1,\n\t                If('Sales'[Total Qty]<800,0,-1)\n                )\n            )\n            ; \nCREATE KPI CURRENTCUBE.[Total Qty] AS Measures.[Total Qty], ASSOCIATED_MEASURE_GROUP = 'Sales', GOAL = Measures.[_Total Qty Goal], STATUS = Measures.[_Total Qty Status], STATUS_GRAPHIC = 'Traffic Light - Single';\n",
                    commands[1], "H51 command");
                var a = ann["[Total Qty]"];
                Check.Equal("LinearReversed", a["KpiStatusType"], "KpiStatusType");
                Check.Equal("Ascending", a["KpiThresholdOrdering"], "ordering stays Ascending");
                Check.Equal("2", a["KpiThresholdCount"], "two thresholds");
                Check.Equal("400", a["KpiThreshold_0"], "threshold 0");
                Check.Equal("800", a["KpiThreshold_1"], "threshold 1");
            }
        }

        // H50: a second KPI changes only the script (version 12 with the extra measure); measures keep the order added.
        public static void KpiTwo()
        {
            var (model, total) = KpiModel();
            using (model)
            {
                total.AddKpi(1000);
                model.Tables["Sales"].AddMeasure("Total Amount", "Sum(Sales[Amount])").AddKpi(10000);
                var m = SaveKpiModel(model);
                var (commands, ann, refs) = KpiScriptParts(m);
                Check.Equal(3, commands.Count, "header and two measures");
                Check.True(commands[1].StartsWith(MeasuresHeader + "CREATE MEASURE 'Sales'[Total Qty]=", StringComparison.Ordinal), "first KPI first");
                Check.Equal(MeasuresHeader + "CREATE MEASURE 'Sales'[Total Amount]=Sum(Sales[Amount]);\nCREATE MEASURE 'Sales'[_Total Amount Goal] = 10000; \nCREATE MEASURE 'Sales'[_Total Amount Status] = if(ISBLANK('Sales'[Total Amount]),BLANK(),\n                    If('Sales'[Total Amount]<4000,-1,\n\t                If('Sales'[Total Amount]<8000,0,1)\n                )\n            )\n            ; \nCREATE KPI CURRENTCUBE.[Total Amount] AS Measures.[Total Amount], ASSOCIATED_MEASURE_GROUP = 'Sales', GOAL = Measures.[_Total Amount Goal], STATUS = Measures.[_Total Amount Status], STATUS_GRAPHIC = 'Traffic Light - Single';\n",
                    commands[2], "H50 second command");
                Check.Equal("[Total Qty]|[_Total Qty Goal]|[_Total Qty Status]|KPIs.[Total Qty]|[Total Amount]|[_Total Amount Goal]|[_Total Amount Status]|KPIs.[Total Amount]|Measures.[__XL_Count of Models]", string.Join("|", refs), "properties, default member once");
                Check.Equal("4000", ann["[Total Amount]"]["KpiThreshold_0"], "second thresholds");
                string Text(string pattern) => Encoding.UTF8.GetString(m.Find(pattern));
                Text(@"MdxScript\.12\.scr\.xml$");
                Text(@"Model\.19\.cub\.xml$");
                Text(@"Products_G\.6\.dim\.xml$");
            }
        }

        public static void KpiProfile()
        {
            var (model, total) = KpiModel();
            using (model)
            {
                total.AddKpi(1000);
                var m = SaveKpiModel(model);
                string Text(string pattern) => Encoding.UTF8.GetString(m.Find(pattern));
                // The hierarchy model one edit further on: cube, measure groups and script one step up, the primary side of the first
                // relationship saved again (empty annotation values dropped) at version 6, the other dimensions and the partitions as before.
                Check.True(Text(@"Model\.19\.cub\.xml$").Contains("<ObjectVersion>19</ObjectVersion>"), "cube version");
                Check.True(Text(@"Model\.19\.cub\.xml$").Contains("<AssemblyFileList>MdxScript.10.scr.xml</AssemblyFileList>"), "script file");
                Check.True(Text(@"Model\.19\.cub\.xml$").Contains("<Kpis/>"), "the cube lists no KPIs: they live in the script");
                Check.True(Text(@"Model\.19\.cub\.xml$").Contains("<StructureVersion>11</StructureVersion>"), "cube structure as before");
                Text(@"\.cub\\info\.19\.xml$");
                Text(@"Products_G\.19\.det\.xml$");
                Text(@"Sales_G\.19\.det\.xml$");
                Text(@"Sales_G\.9\.prt\.xml$");
                Text(@"Sales_G\.3\.dim\.xml$");
                string products = Text(@"Products_G\.6\.dim\.xml$");
                Check.True(!products.Contains("<Value/>"), "the saved dimension drops the empty annotation values");
                Check.Equal(0, m.Files.Keys.Count(k => System.Text.RegularExpressions.Regex.IsMatch(k, @"Products_[0-9a-f-]+\.5\.dim\.xml$")), "no old Products dimension");
                var mgOrder = System.Text.RegularExpressions.Regex.Match(Text(@"Model\.19\.cub\.xml$"), "<MeasureGroupFileList>(.*?)</MeasureGroupFileList>").Groups[1].Value.Split(';').Select(x => x[..x.IndexOf('.')]).ToList();
                Check.Equal(string.Join(";", mgOrder.OrderBy(x => x, StringComparer.Ordinal)), string.Join(";", mgOrder), "measure groups listed by id");
            }
            // Without a KPI the counters stay those of the hierarchy model.
            var (plainModel, _) = KpiModel();
            using (plainModel)
            {
                var m = SaveKpiModel(plainModel);
                m.Find(@"Model\.18\.cub\.xml$");
                m.Find(@"MdxScript\.9\.scr\.xml$");
                Check.Equal(0, m.Files.Keys.Count(k => k.Contains("__XL_Count") && k.Contains("KPIs")), "nothing KPI-related");
                Check.True(!Encoding.UTF8.GetString(m.Find(@"MdxScript\.9\.scr\.xml$")).Contains("CalculationReference>Measures.[__XL_Count"), "no default member property without a KPI");
            }
        }

        public static void KpiValidation()
        {
            // AddKpi
            var (model, total) = KpiModel();
            using (model)
            {
                var sales = model.Tables["Sales"];
                var other = model.Tables["Products"].AddMeasure("Rows", "COUNTROWS(Products)");
                Throws<ArgumentOutOfRangeException>(() => total.AddKpi(0), "zero target", "positive");
                Throws<ArgumentOutOfRangeException>(() => total.AddKpi(-5), "negative target", "positive");
                Throws<ArgumentOutOfRangeException>(() => total.AddKpi(double.NaN), "NaN target", "'Sales'[Total Qty]");
                Throws<ArgumentNullException>(() => total.AddKpi((ModelMeasure)null!), "null target measure");
                Throws<ArgumentException>(() => total.AddKpi(total), "the measure itself", "itself");
                using (var foreign = new DataModelWriter())
                {
                    var t = foreign.AddTable("Sales", new[] { new { Qty = 1L } });
                    Throws<ArgumentException>(() => total.AddKpi(t.AddMeasure("X", "1")), "measure of another model", "another DataModelWriter");
                }
                Throws<ArgumentException>(() => total.AddKpi(1000, new KpiOptions { Thresholds = new[] { 0.4, 0.8, 1.2 } }), "three thresholds for Linear", "2 thresholds");
                Throws<ArgumentException>(() => total.AddKpi(1000, new KpiOptions { Scale = KpiScale.LinearReversed, Thresholds = new[] { 0.4 } }), "one threshold for LinearReversed", "2 thresholds");
                Throws<ArgumentException>(() => total.AddKpi(1000, new KpiOptions { Scale = KpiScale.Centered, Thresholds = new[] { 0.4, 0.8 } }), "two thresholds for Centered", "4 thresholds");
                Throws<ArgumentException>(() => new KpiOptions { Thresholds = new[] { 0.8, 0.4 } }, "descending thresholds", "ascending");
                Throws<ArgumentException>(() => new KpiOptions { Thresholds = new[] { 0.0, 0.4 } }, "zero threshold", "positive");
                Throws<ArgumentOutOfRangeException>(() => new KpiOptions { Scale = (KpiScale)7 }, "scale out of range");
                Throws<ArgumentException>(() => total.AddKpi(1000, new KpiOptions { IconStyle = KpiIconStyle.GaugeAscending, Thresholds = new[] { 0.4, 0.8 } }), "the gauge needs four thresholds", "takes 4 thresholds");
                Throws<ArgumentException>(() => total.AddKpi(1000, new KpiOptions { Scale = KpiScale.CenteredReversed, IconStyle = KpiIconStyle.GaugeAscending, Thresholds = new[] { 0.2, 0.4, 0.6, 0.8 } }), "the centered gauge needs eight thresholds", "takes 8 thresholds");
                Check.True(total.Kpi is null, "failed calls change nothing");
                var kpi = total.AddKpi(1000);
                Check.Equal("0.4,0.8", string.Join(",", kpi.Thresholds.Select(x => x.ToString(System.Globalization.CultureInfo.InvariantCulture))), "default thresholds");
                Check.Equal("Traffic Light - Single", kpi.IconStyle.Name, "default icon style");
                Throws<InvalidOperationException>(() => total.AddKpi(2000), "second KPI on a measure", "already has a KPI");
                Check.Equal(1000.0, kpi.Target!.Value, "the first KPI stays");
                // Save: one KPI per model
            }
            // Save: shapes no Excel sample shows
            {
                var (m, t) = KpiModel();
                using (m)
                {
                    t.AddKpi(1000);
                    m.Tables["Products"].AddCalculatedColumn("Name Length", "LEN([Product])", ColumnDataType.Int64);
                    m.Tables["Sales"].AddCalculatedColumn("Line Total", "[Qty]*[Amount]", ColumnDataType.Currency);
                    Throws<InvalidOperationException>(() => m.SaveModelPart(new MemoryStream()), "every primary-side table holds columns", "none");
                }
            }
            {
                var (m, t) = KpiModel();
                using (m)
                {
                    t.AddKpi(1000);
                    t.IsHidden = true;
                    m.SaveModelPart(new MemoryStream());
                }
            }
            {
                var (m, t) = KpiModel();
                using (m)
                {
                    t.AddKpi(1000);
                    m.Tables["Sales"]["Qty"].IsHidden = true;
                    m.SaveModelPart(new MemoryStream());
                }
            }
            {
                var (m, t) = KpiModel();
                using (m)
                {
                    t.AddKpi(1000);
                    m.Tables["Sales"].AddMeasure("_total qty goal", "1");
                    Throws<InvalidOperationException>(() => m.SaveModelPart(new MemoryStream()), "a measure named like the hidden goal measure", "_Total Qty Goal");
                }
            }
        }

        public static void ApiTableHandles()
        {
            using var model = new DataModelWriter();
            var sales = model.AddTable("Sales (USD)", new[] { new { ProductKey = 1L, Qty = 2L } });
            Check.Equal("Sales USD", sales.Name, "AddTable returns the table under its model name");
            Check.Equal(1L, sales.RowCount, "rows loaded");
            Check.True(model.Tables["Sales (USD)"] == sales && model.Tables["sales usd"] == sales && model.Tables[0] == sales, "found by the name given, the model name and position");
            var products = model.CreateTable("Products", new[] { new ColumnDefinition("ProductKey", ColumnDataType.Int64), new ColumnDefinition("Category", ColumnDataType.String), new ColumnDefinition("CategoryOrder", ColumnDataType.Int64) });
            products.AddRow(1, "Bikes", 1);

            var rel = model.AddRelationship(sales["ProductKey"], products["ProductKey"]);
            Check.Equal("'Sales USD'[ProductKey] -> 'Products'[ProductKey]", rel.ToString(), "relationship returned");
            Check.True(model.Relationships.Single() == rel, "listed");
            var measure = sales.AddMeasure("Total Qty", "=SUM('Sales USD'[Qty])", MeasureFormat.WholeNumber);
            Check.Equal("SUM('Sales USD'[Qty])", measure.Expression, "leading = removed");
            measure.IsHidden = true;
            Check.True(sales.Measures.Single().IsHidden, "measure hidden through its property");
            var hierarchy = products.AddHierarchy("Product Hierarchy", "Category", "ProductKey");
            Check.Equal("Product Hierarchy (Category > ProductKey)", hierarchy.ToString(), "hierarchy returned");

            products["Category"].SortByColumn = products["CategoryOrder"];
            Check.True(products["Category"].SortByColumn == products["CategoryOrder"], "sort column set");
            products["Category"].SortByColumn = null;
            Check.True(products["Category"].SortByColumn is null, "sort column cleared");
            Throws<ArgumentException>(() => products["Category"].SortByColumn = sales["Qty"], "sort column of another table", "'Products'");

            Throws<KeyNotFoundException>(() => _ = model.Tables["Sale"], "unknown table", "The model's tables are: 'Sales USD', 'Products'");
            Throws<KeyNotFoundException>(() => _ = products["Nope"], "unknown column", "Its columns are: 'ProductKey', 'Category', 'CategoryOrder'");
            Throws<ArgumentException>(() => products.AddHierarchy("H", "Nope", "ProductKey"), "unknown level", "Its columns are");
            Check.True(model.Tables.Contains("Products") && !model.Tables.Contains("Nope"), "Contains");

            using var other = new DataModelWriter();
            var foreign = other.CreateTable("Customers", new[] { new ColumnDefinition("ProductKey", ColumnDataType.Int64) });
            Throws<ArgumentException>(() => model.AddRelationship(sales["ProductKey"], foreign["ProductKey"]), "a column of another model", "another DataModelWriter");

            var dates = model.AddDateTable("Dates", new DateOnly(2025, 1, 1), new DateOnly(2025, 1, 31));
            Check.True(model.DateTable == dates && dates.IsDateTable, "date table");
            Throws<InvalidOperationException>(() => dates.AddRow(new object?[dates.Columns.Count]), "rows added to a generated date table", "AddDateTable");
            Throws<ArgumentException>(() => dates["Date"].SortByColumn = dates["Day"], "date column sorted", "date column");
            Check.Equal(365L, dates.RowCount, "generated table unchanged");
            Throws<ArgumentException>(() => new DataModelOptions { Culture = CultureInfo.InvariantCulture }, "culture without LCID", "Windows locale id");
        }

        public static void ApiRowsAreAtomic()
        {
            using (var model = new DataModelWriter())
            {
                var t = model.CreateTable("T", new[] { new ColumnDefinition("A", ColumnDataType.Int64), new ColumnDefinition("B", ColumnDataType.DateTime) });
                Throws<ArgumentException>(() => t.AddRow(1, "not a date"), "unconvertible value", "Row 1 of table 'T': column 'B' (DateTime)");
                Throws<ArgumentException>(() => t.AddRow(1), "too few values", "has 1 values, but the table has 2 columns");
                Throws<ArgumentException>(() => t.AddRow(1, new DateTime(50, 1, 1)), "date before the year 100", "column 'B' (DateTime)");
                Throws<ArgumentException>(() => t.AddRow(1L << 40, "x"), "second value fails after a first that needs no conversion", "column 'B'");
                t.AddRow(2, new DateTime(2025, 1, 1));
                Check.Equal(1L, t.RowCount, "only the good row counts");
                string tbl = Encoding.UTF8.GetString(ModelReader.Read(Save(model)).Find(@"\.0\.dim\\T_G\.0\.tbl\.xml$"));
                Check.True(tbl.Contains("<Records xsi:type=\"xsd:long\">1</Records>"), "the model saves with the good row only");
            }
            using (var model = new DataModelWriter())
            {
                var source = new System.Data.DataTable("T");
                source.Columns.Add("A", typeof(long));
                source.Columns.Add("B", typeof(string));
                source.Rows.Add(1L, "1");
                source.Rows.Add(2L, "x");
                var t = model.CreateTable("T", new[] { new ColumnDefinition("A", ColumnDataType.Int64), new ColumnDefinition("B", ColumnDataType.Int64) });
                Throws<ArgumentException>(() => t.AddRows(source.CreateDataReader()), "reader value that does not fit", "Row 2 of table 'T': column 'B'");
                Throws<InvalidOperationException>(() => Save(model), "a table left incomplete cannot be saved", "loading row 2");
            }
        }

        private static byte[] Save(DataModelWriter model)
        {
            var ms = new MemoryStream();
            model.SaveModelPart(ms);
            return ms.ToArray();
        }

        public static void ApiHideColumns()
        {
            (string dim, string[] names) Build(bool late)
            {
                using var model = new DataModelWriter(new DataModelOptions { XpressCompression = false });
                var t = model.CreateTable("T", new[]
                {
                    new ColumnDefinition("A", ColumnDataType.Int64),
                    new ColumnDefinition("B", ColumnDataType.String) { IsHidden = !late },
                    new ColumnDefinition("C", ColumnDataType.Double) { IsHidden = !late },
                });
                t.AddRow(1, "x", 1.5);
                if (late) { t["B"].IsHidden = true; t["c"].IsHidden = true; }
                Check.True(t.Columns[1].IsHidden && t.Columns[2].IsHidden && !t.Columns[0].IsHidden, "Columns reflects the hidden state");
                var m = ModelReader.Read(Save(model));
                return (HiddenAttributes(Encoding.UTF8.GetString(m.Find(@"\.0\.db\\T_G\.1\.dim\.xml$"))),
                        m.Files.Keys.Select(k => System.Text.RegularExpressions.Regex.Replace(ModelReader.Normalize(k), "[0-9A-F]{20}", "DB")).OrderBy(x => x, StringComparer.Ordinal).ToArray());
            }
            var early = Build(late: false);
            var late = Build(late: true);
            Check.Equal("__XL_RowNumber,B,C", late.dim, "hidden attributes");
            Check.Equal(early.dim, late.dim, "same hidden attributes either way");
            Check.True(early.names.SequenceEqual(late.names), "same files either way");
            using var model = new DataModelWriter();
            model.CreateTable("T", new[] { new ColumnDefinition("A", ColumnDataType.Int64) });
            Throws<KeyNotFoundException>(() => model.Tables["T"]["Nope"].IsHidden = true, "unknown column", "'Nope'");
        }

        public static void ApiSaveIsAtomic()
        {
            string dir = Path.Combine(Path.GetTempPath(), "pw-save-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(dir);
            try
            {
                string path = Path.Combine(dir, "model.xlsx");
                File.WriteAllText(path, "previous");
                using (var model = new DataModelWriter())
                {
                    var one = model.CreateTable("One", new[] { new ColumnDefinition("K", ColumnDataType.Int64) });
                    one.AddRow(1); one.AddRow(1);
                    var many = model.CreateTable("Many", new[] { new ColumnDefinition("K", ColumnDataType.Int64) });
                    many.AddRow(1);
                    model.AddRelationship(model.Tables["Many"]["K"], model.Tables["One"]["K"]);
                    Throws<InvalidOperationException>(() => model.Save(path), "duplicate key", "many side");
                }
                Check.Equal("previous", File.ReadAllText(path), "existing file untouched");
                Check.Equal(1, Directory.GetFiles(dir).Length, "no temporary file left behind");
                using (var model = new DataModelWriter())
                {
                    model.CreateTable("T", new[] { new ColumnDefinition("A", ColumnDataType.Int64) }).AddRow(1);
                    model.Save(path);
                }
                Check.True(new FileInfo(path).Length > 1000, "a valid save replaces the file");
                Check.Equal(1, Directory.GetFiles(dir).Length, "no temporary file left behind after success");
            }
            finally { Directory.Delete(dir, recursive: true); }
        }
    }
}
