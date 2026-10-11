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
    ("MarkAsDateTable accepts DateOnly, DateTime and DateTimeOffset sources", Tests.DateTableSourceTypes),
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
            using var model = new DataModelWriter(new DataModelOptions { Lcid = 2057, XpressCompression = compressed });
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

        private static (DataModelWriter model, TableWriter many, TableWriter one) Pair(ColumnDataType type)
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
                model.AddRelationship("Many", "K", "One", "K");
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
                model.AddRelationship("Many", "K", "One", "K");
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
                model.AddRelationship(many, "K", one, "K");
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
                model.AddRelationship("Many", "K", "One", "K");
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
                Throws(() => m1.AddRelationship("Nope", "K", "One", "K"), "unknown table");
                Throws(() => m1.AddRelationship("Many", "Nope", "One", "K"), "unknown column");
                Throws(() => m1.AddRelationship("Many", "K", "Many", "K"), "same table");
                Throws(() => m1.AddRelationship("Many", "K", "One", "V"), "type mismatch");
                m1.AddRelationship("Many", "K", "One", "K");
                Throws(() => m1.AddRelationship("One", "K", "Many", "K"), "already related");
            }

            var (m2, many2, one2) = Pair(ColumnDataType.Int64);
            using (m2)
            {
                one2.AddRow(1, "a"); one2.AddRow(1, "b"); many2.AddRow(1, 1);
                m2.AddRelationship("Many", "K", "One", "K");
                Throws(() => Save(m2), "duplicate primary key");
            }

            var (m3, many3, one3) = Pair(ColumnDataType.String);
            using (m3)
            {
                one3.AddRow("a", "x"); one3.AddRow("A", "y"); many3.AddRow("a", 1);
                m3.AddRelationship("Many", "K", "One", "K");
                Throws(() => Save(m3), "keys that differ only by case are duplicates");
            }

            var (m4, many4, one4) = Pair(ColumnDataType.Int64);
            using (m4)
            {
                one4.AddRow(null, "x"); many4.AddRow(1, 1);
                m4.AddRelationship("Many", "K", "One", "K");
                Throws(() => Save(m4), "blank primary key");
            }

            using var m5 = new DataModelWriter();
            m5.CreateTable("A", new[] { new ColumnDefinition("K", ColumnDataType.Int64) });
            m5.CreateTable("B", new[] { new ColumnDefinition("K", ColumnDataType.Int64) });
            m5.CreateTable("C", new[] { new ColumnDefinition("K", ColumnDataType.Int64) });
            m5.AddRelationship("A", "K", "B", "K");
            m5.AddRelationship("B", "K", "C", "K");
            Throws(() => m5.AddRelationship("A", "K", "C", "K"), "A-B-C-A is circular");
        }

        private static DataModelWriter ProductsModel(out TableWriter products)
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
            model.AddHierarchy("Products", "Product Hierarchy", "Category", "Subcategory", "Product");
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
            model.AddHierarchy("Products", "Product Hierarchy", "Category", "Subcategory", "Product");
            Check.Equal(1, model.HierarchyCount, "count");
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
                model.AddHierarchy("T", "H", "A", "B");
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
            Rejects(() => model.AddHierarchy("Nope", "H", "Category", "Product"), "unknown table");
            Rejects(() => model.AddHierarchy("Products", "H", "Category"), "one level");
            Rejects(() => model.AddHierarchy("Products", "H", "Category", "Nope"), "unknown column");
            Rejects(() => model.AddHierarchy("Products", "H", "Category", "category"), "same column twice");
            Rejects(() => model.AddHierarchy("Products", "Category", "Category", "Product"), "name of a column");
            Rejects(() => model.AddHierarchy("Products", "", "Category", "Product"), "empty name");
            model.AddHierarchy("Products", "H", "Category", "Product");
            Rejects(() => model.AddHierarchy("Products", "h", "Subcategory", "Product"), "duplicate name");
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
                model.AddMeasure("Many", "Total N", "=SUM(Many[N])", MeasureFormat.WholeNumber);
                model.AddMeasure(many, "Plain", "[Total N] * 2 & \"<x>\"");
                model.AddMeasure("One", "O'Rate", "DIVIDE([Total N],2)", MeasureFormat.Percentage(1), hidden: true);
                Check.Equal(3, model.MeasureCount, "count");
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
                model.AddMeasure("Many", "When", "MAX(Many[N])", MeasureFormat.Date);
                model.AddMeasure("Many", "Money", "SUM(Many[N])", MeasureFormat.Currency());
                string y = Script(model);
                Check.True(y.Contains("<FormatString>'Short Date'</FormatString>") && y.Contains("<Format Format=\"DateTimeShortDatePattern\"/>"), "short date");
                Check.True(y.Contains("<FormatString>'\"$\"#,0.00;-\"$\"#,0.00;\"$\"#,0.00'</FormatString>"), "currency (en-US default)");
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
                model.AddMeasure("Many", "M1", "1");
                void Rejects(Action a, string what)
                {
                    try { a(); } catch (ArgumentException) { return; }
                    throw new Exception("Expected ArgumentException: " + what);
                }
                Rejects(() => model.AddMeasure("Nope", "M2", "1"), "unknown table");
                Rejects(() => model.AddMeasure("Many", "m1", "1"), "duplicate name (case-insensitive)");
                Rejects(() => model.AddMeasure("One", "m1", "1"), "duplicate name in another table");
                Rejects(() => model.AddMeasure("Many", "A[B]", "1"), "brackets");
                Rejects(() => model.AddMeasure("Many", "", "1"), "empty name");
                Rejects(() => model.AddMeasure("Many", "M3", "  "), "empty expression");
                Rejects(() => model.AddMeasure("Many", "N", "1"), "name of a column in the same table");
                Rejects(() => model.AddMeasure("Many", "__x", "1"), "reserved prefix");
            }
        }

        // Dates(Date, Year, Month) with a Calendar hierarchy, related to Sales[Date]: the shape of Excel's H2_date_table.xlsx.
        private static (DataModelWriter model, TableWriter dates, TableWriter sales) DateModel(Action<TableWriter>? fillDates = null)
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
            model.AddRelationship("Sales", "Date", "Dates", "Date");
            model.AddHierarchy("Dates", "Calendar", "Year", "Month");
            return (model, dates, sales);
        }

        /// <summary>The changes Excel makes when Dates[Date] is marked as date table (SPEC section 9, diffed against H2_date_table.xlsx).</summary>
        public static void DateTableMetadata()
        {
            var (model, _, _) = DateModel();
            using (model)
            {
                model.MarkAsDateTable("Dates", "Date");
                Check.Equal("Dates", model.DateTableName, "DateTableName");
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
            using var model = new DataModelWriter(new DataModelOptions { XpressCompression = false, Lcid = 1033 });
            var dates = model.AddDateTable("Dates", new DateOnly(2024, 1, 1), new DateOnly(2024, 12, 31));
            Check.Equal(366L, dates.RowCount, "leap year has 366 rows");
            Check.Equal("Dates", model.DateTableName, "marked by default");
            Check.Equal(1, model.HierarchyCount, "Calendar hierarchy");
            Check.Equal("Date,Year,QuarterNumber,Quarter,MonthNumber,Month,Day,DayOfWeekNumber,DayOfWeek", string.Join(",", dates.Columns.Select(c => c.Name)), "columns");
            var sales = model.CreateTable("Sales", new[] { new ColumnDefinition("Date", ColumnDataType.DateTime), new ColumnDefinition("N", ColumnDataType.Int64) });
            sales.AddRow(new DateTime(2024, 3, 5), 1);
            model.AddRelationship("Sales", "Date", "Dates", "Date");
            var ms = new MemoryStream();
            model.SaveModelPart(ms);
            var m = ModelReader.Read(ms.ToArray());
            string dim = Encoding.UTF8.GetString(m.Find(@"\.0\.db\\Dates_G\.[0-9]+\.dim\.xml$"));
            Check.True(dim.Contains("<Type>Time</Type>"), "dimension type Time");
            Check.Equal(3, System.Text.RegularExpressions.Regex.Matches(dim, "<OrderBy>AttributeKey</OrderBy>").Count, "Quarter, Month and DayOfWeek are sorted by their numbers");
            Check.True(dim.Contains("<Level><Name>Quarter</Name>") && dim.Contains("<Hierarchy><Name>Calendar</Name>"), "hierarchy levels");

            void Rejects<T>(Action a, string what) where T : Exception
            {
                try { a(); } catch (T) { return; }
                throw new Exception("Expected " + typeof(T).Name + ": " + what);
            }
            Rejects<ArgumentException>(() => model.AddDateTable("Bad", new DateOnly(2025, 2, 1), new DateOnly(2025, 1, 1)), "to before from");
            Rejects<InvalidOperationException>(() => model.AddDateTable("Second", new DateOnly(2025, 1, 1), new DateOnly(2025, 1, 2)), "second date table");
            Check.Equal(2, model.TableNames.Count(), "rejected calls add no table");
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
            using (mixed) { mixed.MarkAsDateTable("Dates", "Date"); Saves(mixed); }

            // A typed source with a DateOnly property maps to the same column type.
            using var typed = new DataModelWriter(new DataModelOptions { XpressCompression = false });
            typed.AddTable("Dates", new[] { new DateRow(new DateOnly(2025, 1, 1), 2025, "January"), new DateRow(new DateOnly(2025, 1, 2), 2025, "January") });
            Check.Equal(ColumnDataType.DateTime, ColumnDefinition.FromClrType(typeof(DateOnly)), "DateOnly maps to DateTime");
            Check.Equal(ColumnDataType.DateTime, ColumnDefinition.FromClrType(typeof(DateTimeOffset)), "DateTimeOffset maps to DateTime");
            var sales = typed.CreateTable("Sales", new[] { new ColumnDefinition("Date", ColumnDataType.DateTime) });
            sales.AddRow(new DateOnly(2025, 1, 1));
            typed.AddRelationship("Sales", "Date", "Dates", "Date");
            typed.AddHierarchy("Dates", "Calendar", "Year", "Month");
            typed.MarkAsDateTable("Dates", "Date");
            Saves(typed);
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

            using var model = new DataModelWriter(new DataModelOptions { XpressCompression = false, Lcid = 1033 });
            var dates = model.AddDateTable("Dates", new DateOnly(2024, 1, 1), new DateOnly(2025, 12, 31), yearEndMonth: 6);
            Check.Equal("Date,Year,QuarterNumber,Quarter,MonthNumber,Month,Day,DayOfWeekNumber,DayOfWeek,FiscalYear,FiscalQuarterNumber,FiscalQuarter,FiscalMonthNumber,FiscalMonth",
                string.Join(",", dates.Columns.Select(c => c.Name)), "columns with fiscal ones");
            Check.Equal(2, model.HierarchyCount, "Calendar and Fiscal Calendar");
            var sales = model.CreateTable("Sales", new[] { new ColumnDefinition("Date", ColumnDataType.DateTime), new ColumnDefinition("N", ColumnDataType.Int64) });
            sales.AddRow(new DateTime(2024, 8, 5), 1);
            model.AddRelationship("Sales", "Date", "Dates", "Date");
            var ms = new MemoryStream();
            model.SaveModelPart(ms);   // sort-by columns validate that each name maps to one number
            string dim = Encoding.UTF8.GetString(ModelReader.Read(ms.ToArray()).Find(@"\.0\.db\\Dates_G\.[0-9]+\.dim\.xml$"));
            Check.Equal(5, System.Text.RegularExpressions.Regex.Matches(dim, "<OrderBy>AttributeKey</OrderBy>").Count, "five sorted columns");
            Check.True(dim.Contains("<Hierarchy><Name>Fiscal Calendar</Name>"), "fiscal hierarchy");

            using var dec = new DataModelWriter(new DataModelOptions { XpressCompression = false });
            Check.Equal(9, dec.AddDateTable("D", new DateOnly(2025, 1, 1), new DateOnly(2025, 1, 2)).Columns.Count, "December year end adds no fiscal columns");
            Check.Equal(1, dec.HierarchyCount, "only Calendar");
            using var bad = new DataModelWriter();
            try { bad.AddDateTable("D", new DateOnly(2025, 1, 1), new DateOnly(2025, 1, 2), 13); throw new Exception("Expected ArgumentOutOfRangeException"); }
            catch (ArgumentOutOfRangeException) { }
            try { bad.AddDateTable("D", new DateOnly(2025, 1, 1), new DateOnly(2025, 1, 2), 0); throw new Exception("Expected ArgumentOutOfRangeException"); }
            catch (ArgumentOutOfRangeException) { }
            Check.Equal(0, bad.TableNames.Count(), "invalid month adds no table");
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
                Rejects<ArgumentException>(() => model.MarkAsDateTable("Nope", "Date"), "unknown table");
                Rejects<ArgumentException>(() => model.MarkAsDateTable("Dates", "Nope"), "unknown column");
                Rejects<ArgumentException>(() => model.MarkAsDateTable("Dates", "Year"), "not a date column");
                Check.True(model.DateTableName is null, "nothing marked yet");
                model.MarkAsDateTable("Dates", "Date");
                Rejects<InvalidOperationException>(() => model.MarkAsDateTable("Sales", "Date"), "second date table");
                Saves(model);
            }

            // Hierarchy and relationship are required: no Excel sample of the other cases exists.
            using (var noHierarchy = new DataModelWriter(new DataModelOptions { XpressCompression = false }))
            {
                var d = noHierarchy.CreateTable("Dates", new[] { new ColumnDefinition("Date", ColumnDataType.DateTime) });
                d.AddRow(new DateTime(2025, 1, 1));
                var s = noHierarchy.CreateTable("Sales", new[] { new ColumnDefinition("Date", ColumnDataType.DateTime) });
                s.AddRow(new DateTime(2025, 1, 1));
                noHierarchy.AddRelationship("Sales", "Date", "Dates", "Date");
                noHierarchy.MarkAsDateTable("Dates", "Date");
                Rejects<InvalidOperationException>(() => Saves(noHierarchy), "no hierarchy on the date table");
            }
            using (var noRelationship = new DataModelWriter(new DataModelOptions { XpressCompression = false }))
            {
                var d = noRelationship.CreateTable("Dates", new[] { new ColumnDefinition("Date", ColumnDataType.DateTime), new ColumnDefinition("Year", ColumnDataType.Int64), new ColumnDefinition("Month", ColumnDataType.String) });
                d.AddRow(new DateTime(2025, 1, 1), 2025, "January");
                noRelationship.AddHierarchy("Dates", "Calendar", "Year", "Month");
                noRelationship.MarkAsDateTable("Dates", "Date");
                Rejects<InvalidOperationException>(() => Saves(noRelationship), "no relationship uses the date column");
            }
            var (blank, _, _) = DateModel(t => { t.AddRow(new DateTime(2025, 1, 1), 2025, "January"); t.AddRow(null, 2025, "January"); });
            using (blank)
            {
                blank.MarkAsDateTable("Dates", "Date");
                Rejects<InvalidOperationException>(() => Saves(blank), "blank date");
            }
            var (partial, _, _) = DateModel(t => { t.AddRow(new DateTime(2025, 1, 1, 6, 0, 0), 2025, "January"); });
            using (partial)
            {
                partial.MarkAsDateTable("Dates", "Date");
                Rejects<InvalidOperationException>(() => Saves(partial), "date with a time part");
            }
            var (dup, _, _) = DateModel(t => { t.AddRow(new DateTime(2025, 1, 1), 2025, "January"); t.AddRow(new DateTime(2025, 1, 1), 2025, "January"); });
            using (dup)
            {
                dup.MarkAsDateTable("Dates", "Date");
                Rejects<InvalidOperationException>(() => Saves(dup), "duplicate date");
            }
        }

        public static void WorkbookRelationships()
        {
            var (model, many, one) = Pair(ColumnDataType.Int64);
            using (model)
            {
                one.AddRow(1, "x"); many.AddRow(1, 1);
                model.AddRelationship("Many", "K", "One", "K");
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
                model.AddRelationship("Many", "K", "One", "K");
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
            model.AddHierarchy("Products", "Browse", "Category", "Product");
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
                Rejects<ArgumentException>(() => model.SetSortByColumn("Products", "Product", "Product"), "itself");
                Rejects<ArgumentException>(() => model.SetSortByColumn("Products", "Product", "Nope"), "unknown column");
                Rejects<ArgumentException>(() => model.SetSortByColumn("Products", "Product", "Category"), "sort column that is itself sorted");
                Rejects<ArgumentException>(() => model.SetSortByColumn("Products", "CategoryOrder", "ProductKey"), "sorted column used as sort column");
                Rejects<ArgumentException>(() => model.SetSortByColumn("Nope", "Product", "ProductKey"), "unknown table");
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
                    model.AddRelationship("Many", "K", "One", "K");
                    if (measure) model.AddMeasure("Many", "Mine", "SUM(Many[N])");
                    var ms = new MemoryStream();
                    var opts = new DataModelOptions { AddHelperMeasure = helper, XpressCompression = false };
                    using var m2 = new DataModelWriter(opts);
                    // Rebuild with the options under test.
                    var o = m2.CreateTable("One", new[] { new ColumnDefinition("K", ColumnDataType.Int64), new ColumnDefinition("V", ColumnDataType.String) });
                    var mn = m2.CreateTable("Many", new[] { new ColumnDefinition("K", ColumnDataType.Int64), new ColumnDefinition("N", ColumnDataType.Int64) });
                    o.AddRow(1, "x"); mn.AddRow(1, 1);
                    m2.AddRelationship("Many", "K", "One", "K");
                    if (measure) m2.AddMeasure("Many", "Mine", "SUM(Many[N])");
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
    }
}
