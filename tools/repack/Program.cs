using System.IO.Compression;
using System.Xml.Linq;
using PivotWright.Internal.Abf;
using PivotWright.Tests;

internal static class Check
{
    public static void True(bool c, string w) { if (!c) throw new Exception("Expected true: " + w); }
    public static void Equal<T>(T e, T a, string w) { if (!EqualityComparer<T>.Default.Equals(e, a)) throw new Exception($"{w}: expected {e}, got {a}"); }
}

internal static class Program
{
    // Repacks the model of an Excel-written workbook through PivotWright's ABF writer, keeping every file and the backup log verbatim.
    private static int Main(string[] args)
    {
        string input = args[0], output = args[1];
        string mode = args.Length > 2 ? args[2] : "verbatim";
        string swapGroups = args.Length > 4 ? args[4] : "";   // letters: d dim.xml, p prt.xml, c cube, s script, v dsv, t det.xml
        var m = ModelReader.Read(ModelReader.ItemData(input));
        if (mode == "dump")
        {
            foreach (var kv in m.Files)
            {
                string rel = kv.Key.Replace('\\', '/').TrimStart('/');
                string dest = Path.Combine(output, rel);
                Directory.CreateDirectory(Path.GetDirectoryName(dest)!);
                File.WriteAllBytes(dest, kv.Value);
            }
            Console.WriteLine("dumped " + m.Files.Count);
            return 0;
        }
        var model = new MemoryStream();
        var abf = new AbfWriter(model, new XpressCompressor());
        var contentByStorage = m.InLogOrder.ToDictionary(x => x.Storage, x => x.Content);
        string logXml0 = m.RawLogXml;
        if (mode == "dbfolder")
        {
            // Persist the database at location 1 the way Excel does: folder <id>.1.db, PersistLocation 1 in the log group and in the db XML.
            var l = XElement.Parse(System.Text.RegularExpressions.Regex.Replace(m.RawLogXml, "([0-9A-F]{20})\\.0\\.db", "$1.1.db"));
            foreach (var g in l.Element("FileGroups")!.Elements("FileGroup").Where(g => g.Element("Class")!.Value == "100002"))
            {
                g.Element("PersistLocation")!.Value = "1";
                foreach (var bf in g.Descendants("BackupFile").Where(b => b.Element("Path")!.Value.EndsWith(".db.xml")))
                {
                    string st = bf.Element("StoragePath")!.Value;
                    string x = System.Text.Encoding.UTF8.GetString(contentByStorage[st]);
                    x = x.Replace("<PersistLocation>0</PersistLocation>", "<PersistLocation>1</PersistLocation>");
                    contentByStorage[st] = new System.Text.UTF8Encoding(false).GetBytes(x);
                }
            }
            logXml0 = l.ToString(SaveOptions.DisableFormatting);
        }
        if (mode == "tsmap")
        {
            // Our model, but every timestamp element taken from the counterpart file of an Excel-written model (args[3]).
            var ex = ModelReader.Read(ModelReader.ItemData(args[3]));
            string KeyOf(string p) => System.Text.RegularExpressions.Regex.Replace(System.Text.RegularExpressions.Regex.Replace(System.Text.RegularExpressions.Regex.Replace(p, "[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}", "G"), "[0-9A-F]{20}", "DBID"), "DBID\\.\\d+\\.db", "DBID.N.db");
            var exByKey = ex.Files.GroupBy(kv => KeyOf(kv.Key)).ToDictionary(g => g.Key, g => g.First().Value);
            var tsRx2 = new System.Text.RegularExpressions.Regex("<(CreatedTimestamp|LastSchemaUpdate|LastProcessed|LastUpdate|LastBindingChange)>([^<]*)</\\1>");
            var logT = XElement.Parse(m.RawLogXml);
            int done = 0, skipped = 0;
            foreach (var bf in logT.Descendants("BackupFile"))
            {
                string path = bf.Element("Path")!.Value;
                if (!System.Text.RegularExpressions.Regex.IsMatch(path, "\\.(dim|det|prt|cub|scr|db|ds|dsv)\\.xml$")) continue;
                string rel = path.Contains("VertiPaq_") ? path[(path.IndexOf('\\', path.IndexOf("VertiPaq_")) + 1)..] : path;
                if (!exByKey.TryGetValue(KeyOf(rel), out var other)) { skipped++; Console.WriteLine("no Excel counterpart for " + KeyOf(rel)); continue; }
                string storage = bf.Element("StoragePath")!.Value;
                string mine = System.Text.Encoding.UTF8.GetString(contentByStorage[storage]);
                var exTs2 = tsRx2.Matches(System.Text.Encoding.UTF8.GetString(other)).Select(x => x.Value).ToList();
                if (exTs2.Count != tsRx2.Matches(mine).Count) { skipped++; Console.WriteLine("timestamp count differs in " + KeyOf(rel)); continue; }
                int ti = 0;
                mine = tsRx2.Replace(mine, _ => exTs2[ti++]);
                var bytes2 = new System.Text.UTF8Encoding(false).GetBytes(mine);
                contentByStorage[storage] = bytes2;
                bf.Element("Size")!.Value = bytes2.Length.ToString();
                done++;
            }
            logXml0 = logT.ToString(SaveOptions.DisableFormatting);
            Console.WriteLine($"timestamps replaced in {done} files, skipped {skipped}");
        }
        byte[] partitionsRaw = m.PartitionsRaw;
        if (mode == "idmap")
        {
            // Our model with every id renamed to the id Excel's model has for the same object: table ids, relationship ids, database and
            // data source ids, and the ObjectID of every object. args[3] = Excel workbook. Names inside file paths and contents follow.
            var ex = ModelReader.Read(ModelReader.ItemData(args[3]));
            var guid = new System.Text.RegularExpressions.Regex("[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}");
            string Text(byte[] b) => System.Text.Encoding.UTF8.GetString(b);
            string First(Dictionary<string, byte[]> files, string pattern) => files.Keys.First(k => System.Text.RegularExpressions.Regex.IsMatch(k, pattern));
            string KeyOf(string p) => System.Text.RegularExpressions.Regex.Replace(System.Text.RegularExpressions.Regex.Replace(System.Text.RegularExpressions.Regex.Replace(p, "[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}", "G"), "[0-9A-F]{20}", "DBID"), "DBID\\.\\d+\\.db", "DBID.N.db");
            var map = new Dictionary<string, string>();   // ours -> Excel's
            string only = args.Length > 4 ? args[4] : "trsdo";
            foreach (var t in new[] { "Products", "Dates", "Sales" })
            {
                string idOf(Dictionary<string, byte[]> f) => System.Text.RegularExpressions.Regex.Match(First(f, t + "_[0-9a-f-]{36}\\.\\d+\\.dim\\.xml$"), t + "_[0-9a-f-]{36}").Value;
                if (only.Contains('t')) map[idOf(m.Files)] = idOf(ex.Files);
            }
            string Db(Dictionary<string, byte[]> f) => System.Text.RegularExpressions.Regex.Match(First(f, "[0-9A-F]{20}\\.\\d+\\.db\\.xml$"), "[0-9A-F]{20}").Value;
            if (only.Contains('d')) map[Db(m.Files)] = Db(ex.Files);
            string DsId(Dictionary<string, byte[]> f) => guid.Match(First(f, "\\.0\\.ds\\.xml$")).Value;
            if (only.Contains('s')) map[DsId(m.Files)] = DsId(ex.Files);
            Dictionary<string, string> Rels(Dictionary<string, byte[]> f)
            {
                string x = Text(f[First(f, "Sales_[0-9a-f-]{36}\\.\\d+\\.dim\\.xml$")]);
                var d = new Dictionary<string, string>();
                foreach (System.Text.RegularExpressions.Match r in System.Text.RegularExpressions.Regex.Matches(x, "<ddl300_300:Relationship>\\s*<ID>([0-9a-fA-F-]{36})</ID>.*?<ddl300_300:ToRelationshipEnd>.*?<DimensionID>(\\w+?)_[0-9a-f-]{36}</DimensionID>", System.Text.RegularExpressions.RegexOptions.Singleline))
                    d[r.Groups[2].Value] = r.Groups[1].Value;
                return d;
            }
            var rm = Rels(m.Files); var re = Rels(ex.Files);
            if (only.Contains('r')) foreach (var kv in rm) map[kv.Value] = re[kv.Key];
            // ObjectIDs: one per object, matched by file kind and table
            var exByKey = ex.Files.GroupBy(kv => KeyOf(kv.Key)).ToDictionary(g => g.Key, g => g.First().Value);
            int objIds = 0;
            foreach (var kv in m.Files)
            {
                if (!only.Contains('o')) break;
                if (!System.Text.RegularExpressions.Regex.IsMatch(kv.Key, "\\.(dim|det|prt|cub|scr|db|ds|dsv)\\.xml$")) continue;
                if (!exByKey.TryGetValue(KeyOf(kv.Key), out var other)) continue;
                string a = System.Text.RegularExpressions.Regex.Match(Text(kv.Value), "<ObjectID>([^<]*)</ObjectID>").Groups[1].Value;
                string b = System.Text.RegularExpressions.Regex.Match(Text(other), "<ObjectID>([^<]*)</ObjectID>").Groups[1].Value;
                if (a.Length > 0 && b.Length > 0 && a != b) { map[a] = b; objIds++; }
            }
            Console.WriteLine($"id map: {map.Count} entries ({objIds} ObjectIDs)");
            string Apply(string t) { foreach (var kv in map) t = t.Replace(kv.Key, kv.Value); return t; }
            var logI = XElement.Parse(Apply(m.RawLogXml));
            foreach (var bf in logI.Descendants("BackupFile"))
            {
                string storage = bf.Element("StoragePath")!.Value;
                string path = bf.Element("Path")!.Value;
                if (!path.EndsWith(".xml")) continue;
                var bytes = new System.Text.UTF8Encoding(false).GetBytes(Apply(Text(contentByStorage[storage])));
                contentByStorage[storage] = bytes;
                bf.Element("Size")!.Value = bytes.Length.ToString();
            }
            logXml0 = logI.ToString(SaveOptions.DisableFormatting);
            var pt = Apply(System.Text.Encoding.Unicode.GetString(m.PartitionsRaw, 2, m.PartitionsRaw.Length - 2));
            var body = System.Text.Encoding.Unicode.GetBytes(pt);
            partitionsRaw = new byte[2 + body.Length]; partitionsRaw[0] = 0xFF; partitionsRaw[1] = 0xFE; body.CopyTo(partitionsRaw, 2);
        }
        Dictionary<string, string>? renameMap = null;
        if (mode == "rename")
        {
            // Plain string renames (args[3] = "old=new,old=new", same lengths) applied to every XML file, the log, PARTITIONS and the package parts.
            renameMap = args[3].Split(',').Select(p => p.Split('=')).ToDictionary(p => p[0], p => p[1]);
            // lengths may differ: XML sizes are recomputed below
            string Apply(string t) { foreach (var kv in renameMap) t = t.Replace(kv.Key, kv.Value); return t; }
            var logN = XElement.Parse(Apply(m.RawLogXml));
            int changed = 0;
            foreach (var bf in logN.Descendants("BackupFile"))
            {
                string storage = bf.Element("StoragePath")!.Value;
                if (!bf.Element("Path")!.Value.EndsWith(".xml")) continue;
                string before = System.Text.Encoding.UTF8.GetString(contentByStorage[storage]);
                string after = Apply(before);
                if (after != before) changed++;
                var bytes = new System.Text.UTF8Encoding(false).GetBytes(after);
                contentByStorage[storage] = bytes;
                bf.Element("Size")!.Value = bytes.Length.ToString();
            }
            logXml0 = logN.ToString(SaveOptions.DisableFormatting);
            var ptN = Apply(System.Text.Encoding.Unicode.GetString(m.PartitionsRaw, 2, m.PartitionsRaw.Length - 2));
            var bodyN = System.Text.Encoding.Unicode.GetBytes(ptN);
            partitionsRaw = new byte[2 + bodyN.Length]; partitionsRaw[0] = 0xFF; partitionsRaw[1] = 0xFE; bodyN.CopyTo(partitionsRaw, 2);
            Console.WriteLine($"rename: {renameMap.Count} strings, {changed} xml files changed");
        }
        if (mode == "ordinal")
        {
            // args[3] = "dim:Products=0,det:Sales=8,prt:Dates=2,scr:MdxScript=2,..." sets the object's top-level <Ordinal> (first one in the ObjectDefinition).
            var want = args[3].Split(',').Select(p => p.Split('=')).ToDictionary(p => p[0], p => p[1]);
            var logO = XElement.Parse(m.RawLogXml);
            int n = 0;
            foreach (var bf in logO.Descendants("BackupFile"))
            {
                string path = bf.Element("Path")!.Value, storage = bf.Element("StoragePath")!.Value;
                var km = System.Text.RegularExpressions.Regex.Match(path, @"([^\\/]+?)(?:_[0-9a-f-]{36})?\.\d+\.(dim|det|prt|scr|cub)\.xml$");
                if (!km.Success) continue;
                string key = km.Groups[2].Value + ":" + km.Groups[1].Value;
                if (!want.TryGetValue(key, out var val)) continue;
                string t = System.Text.Encoding.UTF8.GetString(contentByStorage[storage]);
                var rx = new System.Text.RegularExpressions.Regex("(<ObjectDefinition><[A-Za-z]+[^>]*>.*?)<Ordinal>\\d+</Ordinal>", System.Text.RegularExpressions.RegexOptions.Singleline);
                string t2 = rx.Replace(t, mm => mm.Groups[1].Value + "<Ordinal>" + val + "</Ordinal>", 1);
                Check.True(t2 != t || t.Contains("<Ordinal>" + val + "</Ordinal>"), "ordinal set " + key);
                var bytes = new System.Text.UTF8Encoding(false).GetBytes(t2);
                contentByStorage[storage] = bytes;
                bf.Element("Size")!.Value = bytes.Length.ToString();
                Console.WriteLine($"ordinal {key} = {val}"); n++;
            }
            Check.Equal(want.Count, n, "ordinals applied");
            logXml0 = logO.ToString(SaveOptions.DisableFormatting);
        }
        if (mode == "detfix")
        {
            // args[3] letters: m = count measure ID set to its name ("__XL_Count T"); r = DegenerateMeasureGroupDimension __XL_RowNumber TableID set to the table id.
            string L = args[3];
            var logD = XElement.Parse(m.RawLogXml);
            foreach (var bf in logD.Descendants("BackupFile"))
            {
                string path = bf.Element("Path")!.Value, storage = bf.Element("StoragePath")!.Value;
                if (!path.EndsWith(".det.xml")) continue;
                string t = System.Text.Encoding.UTF8.GetString(contentByStorage[storage]), t0 = t;
                string mgId = System.Text.RegularExpressions.Regex.Match(t, "<MeasureGroup><Name>[^<]*</Name><ID>([^<]+)</ID>").Groups[1].Value;
                if (L.Contains('m'))
                    t = System.Text.RegularExpressions.Regex.Replace(t, "<Measure><Name>(__XL_Count [^<]*)</Name><ID>[^<]+</ID>", mm => "<Measure><Name>" + mm.Groups[1].Value + "</Name><ID>" + mm.Groups[1].Value + "</ID>");
                if (L.Contains('r'))
                    t = System.Text.RegularExpressions.Regex.Replace(t, "<TableID>[^<_]+</TableID><ColumnID>__XL_RowNumber</ColumnID>", "<TableID>" + mgId + "</TableID><ColumnID>__XL_RowNumber</ColumnID>");
                Console.WriteLine($"detfix {path[(path.LastIndexOf('\\') + 1)..]}: {(t == t0 ? "unchanged" : "changed")}");
                var bytes = new System.Text.UTF8Encoding(false).GetBytes(t);
                contentByStorage[storage] = bytes;
                bf.Element("Size")!.Value = bytes.Length.ToString();
            }
            logXml0 = logD.ToString(SaveOptions.DisableFormatting);
        }
        if (mode == "rswap")
        {
            // Our model with some of Excel's files put in (b = database xml, o = data source xml, k = CryptKey.bin, q = PARTITIONS);
            // Excel's ids are mapped to ours, ObjectIDs kept as in our backup log. args[3] = Excel workbook, args[4] = letters.
            var ex = ModelReader.Read(ModelReader.ItemData(args[3]));
            var guid = new System.Text.RegularExpressions.Regex("[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}");
            string Text(byte[] b) => System.Text.Encoding.UTF8.GetString(b);
            string First(Dictionary<string, byte[]> files, string pattern) => files.Keys.First(k => System.Text.RegularExpressions.Regex.IsMatch(k, pattern));
            var map = new Dictionary<string, string>();   // Excel -> ours
            foreach (var t in new[] { "Products", "Dates", "Sales" })
            {
                string idOf(Dictionary<string, byte[]> f) => System.Text.RegularExpressions.Regex.Match(First(f, t + "_[0-9a-f-]{36}\\.\\d+\\.dim\\.xml$"), t + "_[0-9a-f-]{36}").Value;
                map[idOf(ex.Files)] = idOf(m.Files);
            }
            string Db(Dictionary<string, byte[]> f) => System.Text.RegularExpressions.Regex.Match(First(f, "[0-9A-F]{20}\\.\\d+\\.db\\.xml$"), "[0-9A-F]{20}").Value;
            map[Db(ex.Files)] = Db(m.Files);
            string DsId(Dictionary<string, byte[]> f) => guid.Match(First(f, "\\.0\\.ds\\.xml$")).Value;
            map[DsId(ex.Files)] = DsId(m.Files);
            Dictionary<string, string> Rels(Dictionary<string, byte[]> f)
            {
                string x = Text(f[First(f, "Sales_[0-9a-f-]{36}\\.\\d+\\.dim\\.xml$")]);
                var d = new Dictionary<string, string>();
                foreach (System.Text.RegularExpressions.Match r in System.Text.RegularExpressions.Regex.Matches(x, "<ddl300_300:Relationship>\\s*<ID>([0-9a-fA-F-]{36})</ID>.*?<ddl300_300:ToRelationshipEnd>.*?<DimensionID>(\\w+?)_[0-9a-f-]{36}</DimensionID>", System.Text.RegularExpressions.RegexOptions.Singleline))
                    d[r.Groups[2].Value] = r.Groups[1].Value;
                return d;
            }
            var re = Rels(ex.Files); var ro = Rels(m.Files);
            foreach (var kv in re) map[kv.Value] = ro[kv.Key];
            string G = args.Length > 4 ? args[4] : "";
            var logR = XElement.Parse(m.RawLogXml);
            foreach (var bf in logR.Descendants("BackupFile"))
            {
                string path = bf.Element("Path")!.Value;
                string kind = path.EndsWith(".db.xml") ? "b" : path.EndsWith(".ds.xml") ? "o" : path.EndsWith("CryptKey.bin") ? "k" : "";
                if (kind.Length == 0 || !G.Contains(kind)) continue;
                string exKey = ex.Files.Keys.First(k => kind == "b" ? System.Text.RegularExpressions.Regex.IsMatch(k, "[0-9A-F]{20}\\.\\d+\\.db\\.xml$") : kind == "o" ? k.EndsWith(".ds.xml") : k.EndsWith("CryptKey.bin"));
                string storage = bf.Element("StoragePath")!.Value;
                byte[] bytes;
                if (kind == "k") bytes = ex.Files[exKey];
                else
                {
                    string txt = Text(ex.Files[exKey]);
                    foreach (var kv in map) txt = txt.Replace(kv.Key, kv.Value);
                    string mineOld = Text(contentByStorage[storage]);
                    string exId = System.Text.RegularExpressions.Regex.Match(txt, "<ObjectID>([^<]*)</ObjectID>").Groups[1].Value;
                    string myId = System.Text.RegularExpressions.Regex.Match(mineOld, "<ObjectID>([^<]*)</ObjectID>").Groups[1].Value;
                    if (exId.Length > 0 && myId.Length > 0) txt = txt.Replace("<ObjectID>" + exId + "</ObjectID>", "<ObjectID>" + myId + "</ObjectID>");
                    if (kind == "b") txt = txt.Replace("<PersistLocation>1</PersistLocation>", "<PersistLocation>0</PersistLocation>");   // our log keeps the database at location 0
                    bytes = new System.Text.UTF8Encoding(false).GetBytes(txt);
                }
                contentByStorage[storage] = bytes;
                bf.Element("Size")!.Value = bytes.Length.ToString();
                Console.WriteLine("took Excel's " + kind + " (" + bytes.Length + " bytes)");
            }
            if (G.Contains('q'))
            {
                string pt = System.Text.Encoding.Unicode.GetString(ex.PartitionsRaw, 2, ex.PartitionsRaw.Length - 2);
                foreach (var kv in map) pt = pt.Replace(kv.Key, kv.Value);
                var body = System.Text.Encoding.Unicode.GetBytes(pt);
                partitionsRaw = new byte[2 + body.Length]; partitionsRaw[0] = 0xFF; partitionsRaw[1] = 0xFE; body.CopyTo(partitionsRaw, 2);
                Console.WriteLine("took Excel's PARTITIONS");
            }
            logXml0 = logR.ToString(SaveOptions.DisableFormatting);
        }
        if (mode == "swap")
        {
            // Replace chosen ASSL files of the Excel model by PivotWright's own (args[3] = our workbook), with our ids mapped to Excel's.
            var ours = ModelReader.Read(ModelReader.ItemData(args[3]));
            var guid = new System.Text.RegularExpressions.Regex("[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}");
            string Key(string p) => System.Text.RegularExpressions.Regex.Replace(System.Text.RegularExpressions.Regex.Replace(guid.Replace(p, "G"), "[0-9A-F]{20}", "DBID"), "DBID\\.\\d+\\.db", "DBID.N.db");
            string Text(byte[] b) => System.Text.Encoding.UTF8.GetString(b);
            var map = new Dictionary<string, string>();
            string First(Dictionary<string, byte[]> files, string pattern) => files.Keys.First(k => System.Text.RegularExpressions.Regex.IsMatch(k, pattern));
            foreach (var t in new[] { "Products", "Dates", "Sales" })
            {
                string idOf(Dictionary<string, byte[]> f) => System.Text.RegularExpressions.Regex.Match(First(f, "\\\\?" + t + "_[0-9a-f-]{36}\\.\\d+\\.dim\\.xml$"), t + "_[0-9a-f-]{36}").Value;
                map[idOf(ours.Files)] = idOf(m.Files);
            }
            string Db(Dictionary<string, byte[]> f) => System.Text.RegularExpressions.Regex.Match(First(f, "[0-9A-F]{20}\\.\\d+\\.db\\.xml$"), "[0-9A-F]{20}").Value;
            map[Db(ours.Files)] = Db(m.Files);
            string DsId(Dictionary<string, byte[]> f) => guid.Match(First(f, "\\.0\\.ds\\.xml$")).Value;
            map[DsId(ours.Files)] = DsId(m.Files);
            // relationship ids, matched by the table they point to
            Dictionary<string, string> Rels(Dictionary<string, byte[]> f)
            {
                string x = Text(f[First(f, "Sales_[0-9a-f-]{36}\\.\\d+\\.dim\\.xml$")]);
                var d = new Dictionary<string, string>();
                foreach (System.Text.RegularExpressions.Match r in System.Text.RegularExpressions.Regex.Matches(x, "<ddl300_300:Relationship>\\s*<ID>([0-9a-fA-F-]{36})</ID>.*?<ddl300_300:ToRelationshipEnd>.*?<DimensionID>(\\w+?)_[0-9a-f-]{36}</DimensionID>", System.Text.RegularExpressions.RegexOptions.Singleline))
                    d[r.Groups[2].Value] = r.Groups[1].Value;
                return d;
            }
            var ro = Rels(ours.Files); var rx = Rels(m.Files);
            foreach (var kv in ro) map[kv.Value] = rx[kv.Key];
            var oursByKey = ours.Files.GroupBy(kv => Key(kv.Key)).ToDictionary(g => g.Key, g => g.First().Value);
            string[] pats = { "d:\\.dim\\.xml$", "p:\\.prt\\.xml$", "c:\\.cub\\.xml$", "s:\\.scr\\.xml$", "v:\\.dsv\\.xml$", "t:\\.det\\.xml$", "x:\\.tbl\\.xml$" };
            var log0 = XElement.Parse(m.RawLogXml);
            int swapped = 0;
            foreach (var bf in log0.Descendants("BackupFile"))
            {
                string path = bf.Element("Path")!.Value;
                string storage = bf.Element("StoragePath")!.Value;
                bool take = pats.Any(p => swapGroups.Contains(p[0]) && System.Text.RegularExpressions.Regex.IsMatch(path, p[2..])) && (swapGroups.Contains('x') || !path.EndsWith(".tbl.xml"));
                if (!take) continue;
                string rel = path.Contains("VertiPaq_") ? path[(path.IndexOf('\\', path.IndexOf("VertiPaq_")) + 1)..] : path;
                if (!oursByKey.TryGetValue(Key(rel), out var mine)) { Console.WriteLine("no counterpart for " + Key(rel) + "   [ours e.g. " + oursByKey.Keys.FirstOrDefault(k => k.Contains("Products") && k.EndsWith(".dim.xml")) + "]"); continue; }
                string txt = Text(mine);
                foreach (var kv in map) txt = txt.Replace(kv.Key, kv.Value);
                var excelFile = Text(contentByStorage[storage]);
                string ex1 = System.Text.RegularExpressions.Regex.Match(excelFile, "<ObjectID>([^<]*)</ObjectID>").Groups[1].Value;
                string my1 = System.Text.RegularExpressions.Regex.Match(txt, "<ObjectID>([^<]*)</ObjectID>").Groups[1].Value;
                if (!swapGroups.Contains('O') && ex1.Length > 0 && my1.Length > 0) txt = txt.Replace("<ObjectID>" + my1 + "</ObjectID>", "<ObjectID>" + ex1 + "</ObjectID>");
                // Keep Excel's timestamps, so the objects stay consistent with the ones that are not swapped.
                var tsRx = new System.Text.RegularExpressions.Regex("<(CreatedTimestamp|LastSchemaUpdate|LastProcessed|LastUpdate|LastBindingChange)>([^<]*)</\\1>");
                var exTs = tsRx.Matches(excelFile).Select(x => x.Value).ToList();
                var myTs = tsRx.Matches(txt);
                if (swapGroups.Contains('T')) { }
                else if (exTs.Count == myTs.Count)
                {
                    int ti = 0;
                    txt = tsRx.Replace(txt, _ => exTs[ti++]);
                }
                else Console.WriteLine($"timestamp elements differ in {rel}: Excel {exTs.Count}, ours {myTs.Count}");
                if (swapGroups.Contains('O') && ex1.Length > 0 && my1.Length > 0)
                    foreach (var grp in log0.Descendants("FileGroup"))
                        if (grp.Element("ObjectID")?.Value == ex1) grp.Element("ObjectID")!.Value = my1;
                var bytes = new System.Text.UTF8Encoding(false).GetBytes(txt);
                contentByStorage[storage] = bytes;
                bf.Element("Size")!.Value = bytes.Length.ToString();
                swapped++;
            }
            logXml0 = log0.ToString(SaveOptions.DisableFormatting);
            Console.WriteLine($"swapped {swapped} files ({swapGroups}); id map entries {map.Count}");
        }
        string logXml = logXml0;
        var order = XElement.Parse(logXml0).Descendants("BackupFile").Select(b => b.Element("StoragePath")!.Value).ToList();
        if (mode == "ourlog")
        {
            // Our style of backup log: groups in table order (Products, Dates, Sales; the measure group and partition of a table together,
            // then the script, the dimensions, data source and view), and every file with the same write time.
            var log = XElement.Parse(m.RawLogXml);
            var groups = log.Element("FileGroups")!;
            string Cls(XElement g) => g.Element("Class")!.Value;
            string Nm(XElement g) => System.Text.RegularExpressions.Regex.Replace(g.Element("Name")!.Value, "_[0-9a-f-]{36}", "");
            int Rank(string name) => name switch { "Products" => 0, "Dates" => 1, "Sales" => 2, _ => 9 };
            int Cr(XElement g) => Cls(g) switch { "100002" => 0, "100010" => 1, "100016" => 2, "100021" => 2, "100060" => 3, "100006" => 4, "100003" => 5, "100053" => 6, _ => 9 };
            var sorted = groups.Elements("FileGroup").Select((g, i) => (g, i))
                .OrderBy(t => Cr(t.g) switch { 2 => 2, var c => c })
                .ThenBy(t => Cr(t.g) == 2 || Cr(t.g) == 4 ? Rank(Nm(t.g)) : 0)
                .ThenBy(t => Cls(t.g) == "100021" ? 1 : 0)
                .ThenBy(t => t.i).Select(t => t.g).ToList();
            groups.RemoveAll();
            foreach (var g in sorted) groups.Add(g);
            string tick = log.Descendants("LastWriteTime").Select(e => e.Value).OrderBy(v => v).Last();
            foreach (var e in log.Descendants("LastWriteTime")) e.Value = tick;
            logXml = log.ToString(SaveOptions.DisableFormatting);
            order = log.Descendants("BackupFile").Select(b => b.Element("StoragePath")!.Value).ToList();
        }
        abf.WriteRawFile("PARTITIONS", partitionsRaw);
        foreach (var storage in order)
        {
            using var s = abf.CreateFile(storage);
            s.Write(contentByStorage[storage]);
        }
        var header = XElement.Parse(m.HeaderXml);
        var objectId = Guid.Parse(header.Element("ObjectID")!.Value);
        abf.Finish(logXml, objectId);
        Console.WriteLine($"mode {mode}: files {order.Count}, item.data {model.Length} bytes");
        using var src = ZipFile.OpenRead(input);
        using var dst = new FileStream(output, FileMode.Create);
        using var zip = new ZipArchive(dst, ZipArchiveMode.Create);
        foreach (var e in src.Entries)
        {
            var ne = zip.CreateEntry(e.FullName, e.FullName == "xl/model/item.data" ? CompressionLevel.NoCompression : CompressionLevel.Optimal);
            using var o = ne.Open();
            if (e.FullName == "xl/model/item.data") o.Write(model.ToArray());
            else if (renameMap != null)
            {
                using var i = e.Open(); var ms = new MemoryStream(); i.CopyTo(ms); var b = ms.ToArray();
                bool u16 = b.Length > 1 && b[0] == 0xFF && b[1] == 0xFE;
                var enc = u16 ? (System.Text.Encoding)new System.Text.UnicodeEncoding(false, false) : new System.Text.UTF8Encoding(false);
                int skip = u16 ? 2 : (b.Length > 2 && b[0] == 0xEF && b[1] == 0xBB && b[2] == 0xBF ? 3 : 0);
                string t = enc.GetString(b, skip, b.Length - skip), t2 = t;
                foreach (var kv in renameMap) t2 = t2.Replace(kv.Key, kv.Value);
                if (t2 != t) { Console.WriteLine("package part renamed: " + e.FullName); o.Write(b, 0, skip); o.Write(enc.GetBytes(t2)); }
                else o.Write(b);
            }
            else { using var i = e.Open(); i.CopyTo(o); }
        }
        return 0;
    }
}
