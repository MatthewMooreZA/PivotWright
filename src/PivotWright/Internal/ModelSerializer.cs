using System.Text;
using PivotWright.Internal.Abf;
using PivotWright.Internal.Crypto;
using PivotWright.Internal.Metadata;
using PivotWright.Internal.Model;

namespace PivotWright.Internal;

/// <summary>Serializes the collected tables into a Spreadsheet Data Model stream (xl/model/item.data).</summary>
internal sealed class ModelSerializer
{
    private readonly DataModelOptions _options;
    private readonly IReadOnlyList<TableBuilder> _tables;
    private readonly IReadOnlyList<RelationshipDefinition> _relationships;
    private readonly StringComparer _strings;
    private readonly IReadOnlyList<MeasureDefinition> _measures;
    private readonly IReadOnlyList<HierarchyDefinition> _hierarchies;
    private AbfWriter _abf = null!;
    private BackupLog _log = null!;

    public ModelSerializer(DataModelOptions options, IReadOnlyList<TableBuilder> tables,
        IReadOnlyList<RelationshipDefinition> relationships, StringComparer strings,
        IReadOnlyList<MeasureDefinition>? measures = null, IReadOnlyList<HierarchyDefinition>? hierarchies = null)
    {
        _hierarchies = hierarchies ?? Array.Empty<HierarchyDefinition>();
        _measures = measures ?? Array.Empty<MeasureDefinition>();
        if (_measures.Count == 0 && options.AddHelperMeasure && relationships.Count > 0)
            _measures = new[] { new MeasureDefinition { Name = "PivotWright helper", Table = relationships[0].From.Name, Expression = "1", Format = MeasureFormat.General, Hidden = true } };
        _options = options;
        _tables = tables;
        _relationships = relationships;
        _strings = strings;
    }

    /// <summary>Names of the model tables and the data source they belong to (for the workbook parts).</summary>
    public ModelInfo Info { get; private set; } = null!;
    public IReadOnlyList<EncodedTable> EncodedTables { get; private set; } = Array.Empty<EncodedTable>();

    private static readonly UTF8Encoding Utf8 = new(false);

    // Date tables are written the way Excel saves them right after "Mark as Date Table" on a model with hierarchies and
    // a relationship to the date column (SPEC section 9). Anything else was never seen in an Excel-written file, so it is refused.
    private void ValidateDateTables()
    {
        foreach (var t in _tables.Where(t => t.DateColumn is not null))
        {
            var col = t.DateColumn!;
            string what = $"'{t.Name}[{col.Definition.Name}]'";
            if (!_hierarchies.Any(h => h.Table == t))
                throw new InvalidOperationException($"{what} is marked as date table, which is only supported for tables that have a hierarchy (no Excel sample of the other case exists yet). Add a hierarchy such as Year > Month to '{t.Name}'.");
            if (!_relationships.Any(r => r.To == t && r.ToColumn == col))
                throw new InvalidOperationException($"{what} is marked as date table, which is only supported when a relationship uses it as its primary key (AddRelationship(..., \"{t.Name}\", \"{col.Definition.Name}\")).");
            long row = 0;
            foreach (var key in col.Keys())
            {
                row++;
                if (key is null)
                    throw new InvalidOperationException($"{what} is marked as date table but row {row} has no date; a date table needs a date on every row.");
                double d = (double)key;
                if (d != Math.Floor(d))
                    throw new InvalidOperationException($"{what} is marked as date table but row {row} ({DateTime.FromOADate(d):yyyy-MM-dd HH:mm:ss}) is not a whole day; Excel requires dates without a time part.");
            }
        }
    }

    // A calculated column is written the way Excel saves one added in manual calculation mode, to a model with hierarchies and
    // relationships (H10, SPEC section 10). Any other shape was never seen in an Excel-written file, so it is refused.
    private void ValidateCalculatedColumns()
    {
        var withColumns = _tables.Where(t => t.CalculatedColumns.Count > 0).ToList();
        if (withColumns.Count == 0) return;
        var table = withColumns[0];
        var column = table.CalculatedColumns[0];
        string what = $"'{table.Name}[{column.Name}]'";
        if (withColumns.Count > 1 || table.CalculatedColumns.Count > 1)
            throw new InvalidOperationException($"{what} is a calculated column, but only one calculated column per model is supported so far (Excel's counters were seen for one only).");
        if (_hierarchies.Count == 0)
            throw new InvalidOperationException($"{what} is a calculated column, which is only supported for models that have a hierarchy (no Excel sample of the other case exists yet).");
        if (!_relationships.Any(r => r.From == table))
            throw new InvalidOperationException($"{what} is a calculated column, which is only supported on the many side of a relationship ('{table.Name}' has none).");
        if (_tables.Any(t => t.DateColumn is not null))
            throw new InvalidOperationException($"{what} is a calculated column, which is not supported together with a date table yet.");
        if (_tables.Any(t => t.Columns.Any(c => c.Definition.IsHidden)))
            throw new InvalidOperationException($"{what} is a calculated column, which is not supported together with hidden columns yet.");
    }

    public void Write(Stream output)
    {
        ValidateDateTables();
        ValidateCalculatedColumns();
        var now = DateTime.UtcNow;
        // Excel names the database with 20 upper-case hex digits. A GUID here loads and pivots fine, but Excel
        // then fails to edit the model (Create Hierarchy -> "ASDetail::OnClone2") once a table has two relationships.
        string dbId = Guid.NewGuid().ToString("N").ToUpperInvariant()[..20];
        string serverRoot = @"\\?\C:\Users\PivotWright\AppData\Local\Temp\{" + Guid.NewGuid().ToString("D").ToUpperInvariant()
                            + @"}\VertiPaq_" + BackupLog.NewStorageName();
        string dbFolder = serverRoot + "\\" + dbId + ".0.db";
        string cubeFolder = dbFolder + "\\Model.0.cub";
        var crypt = ModelCryptKey.Create();

        var tableIds = new Dictionary<string, TableIds>();
        foreach (var t in _tables)
            tableIds[t.Id] = new TableIds { DimensionObjectId = Xml.Guid(), MeasureGroupObjectId = Xml.Guid(), PartitionObjectId = Xml.Guid(), Ordinal = tableIds.Count + 1 };

        var hiddenTables = _tables.Where(t => t.Columns.Any(c => c.Definition.IsHidden)).Select(t => t.Id).ToHashSet();
        bool hiddenProfile = hiddenTables.Count > 0 && _hierarchies.Count > 0 && !_tables.Any(t => t.DateColumn is not null);
        // In the hidden profile the model looks like one that was created a little earlier and edited just now (see ModelInfo.EditedAt).
        var created = hiddenProfile ? now.AddMinutes(-1) : now;
        Info = new ModelInfo
        {
            DatabaseId = dbId,
            DatabaseObjectId = Xml.Guid(),
            CubeObjectId = Xml.Guid(),
            MdxScriptObjectId = Xml.Guid(),
            DataSourceId = Guid.NewGuid().ToString("D"),
            DataSourceName = _options.DataSourceName,
            DataSourceObjectId = Xml.Guid(),
            DsvObjectId = Xml.Guid(),
            Lcid = _options.Lcid,
            CultureName = string.IsNullOrEmpty(_options.Culture.Name) ? "en-US" : _options.Culture.Name,
            Culture = _options.Culture,
            Created = created,
            Processed = created,
            EditedAt = now,
            EncryptedConnectionString = crypt.Encrypt(_options.ConnectionString, ModelCryptKey.ConnectionStringFlags),
            EncryptedDatabasePassword = crypt.Encrypt(string.Empty, ModelCryptKey.DatabasePasswordFlags),
            EncryptedDataSourcePassword = crypt.Encrypt(string.Empty, ModelCryptKey.EmptyPasswordFlags),
            TableIds = tableIds,
            HierarchyTables = _hierarchies.Select(h => h.Table.Id).ToHashSet(),
            DateTables = _tables.Where(t => t.DateColumn is not null).Select(t => t.Id).ToHashSet(),
            HiddenTables = hiddenTables,
            DateReferencingTables = _relationships.Where(r => r.To.DateColumn is not null).Select(r => r.From.Id).ToHashSet(),
            CalculatedTables = _tables.Where(t => t.CalculatedColumns.Count > 0).Select(t => t.Id).ToHashSet(),
            ResavedTables = _tables.Where(t => t.CalculatedColumns.Count > 0)
                .Select(t => _relationships.First(r => r.From == t).To.Id).ToHashSet(),
        };

        _abf = new AbfWriter(output, _options.XpressCompression ? new XpressCompressor() : null);
        _log = new BackupLog { ServerRoot = serverRoot, DatabaseId = dbId, DatabaseName = Info.DatabaseName, Lcid = _options.Lcid };

        // Partition definitions do not depend on the data, so the PARTITIONS list can lead the stream.
        var partitionXml = _tables.ToDictionary(t => t.Id, t => AsslWriter.Partition(Info, t.Name, t.Id));
        long partInfoSize = Utf8.GetByteCount(AsslWriter.PartitionInfo);
        var partitions = AsslWriter.Partitions(Info, "LOCALHOST",
            _tables.Select(t => (t.Name, t.Id, (long)Utf8.GetByteCount(partitionXml[t.Id]) + (HasPartitionInfo(t.Id) ? partInfoSize : 0))).ToList());
        _abf.WriteRawFile("PARTITIONS", Bom(partitions));

        // Data: encode every table into its dimension folder.
        var encoded = new List<EncodedTable>();
        var dimensionFiles = new Dictionary<string, (List<LoggedFile> files, long size)>();
        foreach (var t in _tables)
        {
            var sink = new DimensionSink(_abf, dbFolder + "\\" + t.Id + ".0.dim");
            var et = t.Encode(sink, _options.SegmentRowCount);
            et.Version = Info.TableVersion(t.Id);
            encoded.Add(et);
            if (t.DateColumn is not null && !et.Columns.First(c => c.IsDateKey).IsValueEncoded)
                throw new InvalidOperationException($"The date column of '{t.Name}' is not value-encoded (a very wide range of dates); Excel's file for a date table was only seen with value-encoded dates.");

            // Relationships in which this table is the many side: one index table each.
            foreach (var rel in _relationships.Where(r => r.From == t))
            {
                var fk = et.Columns[t.Columns.ToList().IndexOf(rel.FromColumn)];
                et.Relationships.Add(RelationshipEncoder.Write(rel, fk, et.Id, rel.FromColumn.Id, rel.To.Id, rel.ToColumn.Id,
                    _strings, sink, TableXmlWriter.RelationshipTable));
            }

            // User hierarchies of this table: one U$ tree table each.
            foreach (var h in _hierarchies.Where(h => h.Table == t))
                et.Hierarchies.Add(HierarchyEncoder.Write(h, et, _strings, sink, TableXmlWriter.HierarchyTable));

            // Hierarchy tables and the table metadata.
            foreach (var c in et.Columns)
                sink.WriteText(FileNames.HierarchyTableXml(et.Id, c.Id, c.AttributeHierarchyVersion), TableXmlWriter.HierarchyTable(et, c));
            sink.WriteText(FileNames.TableXml(et.Id, et.Version), TableXmlWriter.MainTable(et, _options.Lcid));
            // Excel reports the size of the table's .tbl.xml documents as the dimension's EstimatedSize.
            dimensionFiles[t.Id] = (sink.Files, sink.Files.Where(f => f.Path.EndsWith(".tbl.xml", StringComparison.Ordinal)).Sum(f => f.Size));
        }
        EncodedTables = encoded;

        var encodedById = encoded.ToDictionary(e => e.Id);

        // Database.
        var dbGroup = Group(BackupLog.ClassDatabase, dbId, Info.DatabaseName, 1, dbFolder, Info.DatabaseObjectId);
        AddText(dbGroup, serverRoot + "\\" + dbId + ".1.db.xml", AsslWriter.Database(Info));
        AddBinary(dbGroup, dbFolder + "\\0.CryptKey.bin", crypt.ToFile());

        // Cube, measure groups and partitions.
        var cubeGroup = Group(BackupLog.ClassCube, "Model", "Model", Info.CubeVersion, cubeFolder, Info.CubeObjectId);
        AddText(cubeGroup, dbFolder + "\\Model." + Info.CubeVersion + ".cub.xml", AsslWriter.Cube(Info, encoded, _measures.Count > 0));
        if (Info.CalculatedProfile) AddText(cubeGroup, cubeFolder + "\\info." + Info.CubeVersion + ".xml", AsslWriter.CubeInfo);
        else if (!Info.AnyHierarchy) AddText(cubeGroup, cubeFolder + "\\info.1.xml", "<Cube xmlns=\"http://schemas.microsoft.com/analysisservices/2003/ext\" xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\"/>");
        foreach (var t in encoded)
        {
            var ids = tableIds[t.Id];
            string detFolder = cubeFolder + "\\" + t.Id + ".0.det";
            var mg = Group(BackupLog.ClassMeasureGroup, t.Id, t.Name, Info.MeasureGroupVersion, detFolder, ids.MeasureGroupObjectId);
            AddText(mg, cubeFolder + "\\" + t.Id + "." + Info.MeasureGroupVersion + ".det.xml", AsslWriter.MeasureGroup(Info, t, encodedById));
            string prtFolder = detFolder + "\\" + t.Id + ".0.prt";
            var prt = Group(BackupLog.ClassPartition, t.Id, t.Name, Info.PartitionVersion, prtFolder, ids.PartitionObjectId);
            AddText(prt, detFolder + "\\" + t.Id + "." + Info.PartitionVersion + ".prt.xml", partitionXml[t.Id]);
            if (HasPartitionInfo(t.Id)) AddText(prt, prtFolder + "\\info." + Info.PartitionVersion + ".xml", AsslWriter.PartitionInfo);
        }
        var script = Group(BackupLog.ClassMdxScript, "MdxScript", "MdxScript", Info.ScriptVersion, cubeFolder, Info.MdxScriptObjectId);
        AddText(script, cubeFolder + "\\MdxScript." + Info.ScriptVersion + ".scr.xml", AsslWriter.MdxScript(Info, _measures));

        // Dimensions (tables).
        foreach (var t in encoded)
        {
            string dimFolder = dbFolder + "\\" + t.Id + ".0.dim";
            var (files, size) = dimensionFiles[t.Id];
            // Excel writes the table that holds relationships at version 2 (see README, Open questions).
            int dv = Info.DimensionVersion(t.Id, t.Relationships.Count > 0);
            var dim = Group(BackupLog.ClassDimension, t.Id, t.Name, dv, dimFolder, tableIds[t.Id].DimensionObjectId);
            var names = files.Select(f => f.Path[(dimFolder.Length + 1)..]).ToList();
            AddText(dim, dbFolder + "\\" + t.Id + "." + dv + ".dim.xml", AsslWriter.Dimension(Info, t, names, size, dv));
            if (!Info.HierarchyTables.Contains(t.Id)) AddText(dim, dimFolder + "\\info." + dv + ".xml", AsslWriter.DimensionInfo(t));
            dim.Files.AddRange(files);
        }

        // Data source and data source view.
        var ds = Group(BackupLog.ClassDataSource, Info.DataSourceId, Info.DataSourceName, 0, dbFolder + "\\" + Info.DataSourceId + ".0.ds", Info.DataSourceObjectId);
        AddText(ds, dbFolder + "\\" + Info.DataSourceId + ".0.ds.xml", AsslWriter.DataSource(Info));
        var dsv = Group(BackupLog.ClassDataSourceView, "Sandbox", "Sandbox", 0, dbFolder, Info.DsvObjectId);
        AddText(dsv, dbFolder + "\\Sandbox.0.dsv.xml", AsslWriter.DataSourceView(Info, encoded));

        _abf.Finish(_log.ToXml(), Guid.NewGuid());
    }

    // After a calculated column was added Excel keeps the partition information of the table that holds it only (H10).
    private bool HasPartitionInfo(string tableId) => !Info.CalculatedProfile || Info.CalculatedTables.Contains(tableId);

    private FileGroup Group(int cls, string id, string name, int version, string path, string objectId)
    {
        var g = new FileGroup { Class = cls, Id = id, Name = name, ObjectVersion = version, PersistLocationPath = path, ObjectId = objectId };
        _log.Groups.Add(g);
        return g;
    }

    private void AddText(FileGroup g, string path, string xml) => AddBinary(g, path, Utf8.GetBytes(xml));

    private void AddBinary(FileGroup g, string path, byte[] content)
    {
        string storage = BackupLog.NewStorageName();
        using (var s = _abf.CreateFile(storage)) s.Write(content);
        g.Files.Add(new LoggedFile(path, storage, FileTime.Now(), content.Length));
    }

    private static byte[] Bom(string text)
    {
        var bytes = new byte[2 + Encoding.Unicode.GetByteCount(text)];
        bytes[0] = 0xFF; bytes[1] = 0xFE;
        Encoding.Unicode.GetBytes(text, 0, text.Length, bytes, 2);
        return bytes;
    }

    /// <summary>Writes a table's files into the ABF stream and remembers them for the backup log.</summary>
    private sealed class DimensionSink : IFileSink
    {
        private readonly AbfWriter _abf;
        private readonly string _folder;

        public DimensionSink(AbfWriter abf, string folder)
        {
            _abf = abf;
            _folder = folder;
        }

        public List<LoggedFile> Files { get; } = new();
        public long TotalSize { get; private set; }

        public Stream Create(string fileName) => new Tracked(this, fileName, _abf.CreateFile(BackupLog.NewStorageName()));

        public void WriteText(string fileName, string xml)
        {
            using var s = Create(fileName);
            s.Write(Utf8.GetBytes(xml));
        }

        private void Closed(string fileName, ChunkedFileStream inner)
        {
            Files.Add(new LoggedFile(_folder + "\\" + fileName, inner.StorageName, FileTime.Now(), inner.UncompressedLength));
            TotalSize += inner.UncompressedLength;
        }

        private sealed class Tracked : Stream
        {
            private readonly DimensionSink _owner;
            private readonly string _name;
            private readonly ChunkedFileStream _inner;
            private bool _done;

            public Tracked(DimensionSink owner, string name, ChunkedFileStream inner)
            {
                _owner = owner; _name = name; _inner = inner;
            }

            public override bool CanRead => false;
            public override bool CanSeek => false;
            public override bool CanWrite => true;
            public override long Length => _inner.Length;
            public override long Position { get => _inner.Position; set => throw new NotSupportedException(); }
            public override void Flush() { }
            public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();
            public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
            public override void SetLength(long value) => throw new NotSupportedException();
            public override void Write(byte[] buffer, int offset, int count) => _inner.Write(buffer, offset, count);
            public override void Write(ReadOnlySpan<byte> buffer) => _inner.Write(buffer);
            public override void WriteByte(byte value) => _inner.WriteByte(value);

            protected override void Dispose(bool disposing)
            {
                if (!_done)
                {
                    _done = true;
                    _inner.Dispose();
                    _owner.Closed(_name, _inner);
                }
                base.Dispose(disposing);
            }
        }
    }
}
