using System.Buffers.Binary;
using PivotWright.Internal.VertiPaq;

namespace PivotWright.Internal.Model;

/// <summary>A many-to-one relationship as declared through the public API.</summary>
internal sealed class RelationshipDefinition
{
    public required string Name { get; init; }
    public required TableBuilder From { get; init; }
    public required ColumnBuilder FromColumn { get; init; }
    public required TableBuilder To { get; init; }
    public required ColumnBuilder ToColumn { get; init; }
}

/// <summary>A relationship after its lookup index was written.</summary>
internal sealed class EncodedRelationship
{
    public required string Name { get; init; }
    public required string ForeignTableId { get; init; }
    public required string ForeignColumnId { get; init; }
    public required string PrimaryTableId { get; init; }
    public required string PrimaryColumnId { get; init; }
    /// <summary>Entries of the INDEX column: one per data id of the foreign key column.</summary>
    public required long Records { get; init; }
    /// <summary>Largest row number referenced (-1 if no foreign key matches).</summary>
    public required long MaxRow { get; init; }
    public required int Width { get; init; }
    /// <summary>Foreign key rows with a value that has no match in the primary key column.</summary>
    public required long ViolationCount { get; init; }

    public string TableName => FileNames.RelationshipTable(ForeignTableId, Name);
}

/// <summary>Compares relationship keys: text by the model collation, numbers by value.</summary>
internal sealed class KeyComparer : IEqualityComparer<object>
{
    private readonly StringComparer _strings;
    public KeyComparer(StringComparer strings) => _strings = strings;

    public new bool Equals(object? a, object? b) =>
        a is string sa && b is string sb ? _strings.Equals(sa, sb) : a is not string && b is not string && object.Equals(a, b);

    public int GetHashCode(object o) => o is string s ? _strings.GetHashCode(s) : o.GetHashCode();
}

internal static class RelationshipEncoder
{
    /// <summary>
    /// Computes the relationship index of an encoded foreign key column and writes its R$ table
    /// (INDEX column plus metadata) into the foreign key table's folder.
    /// </summary>
    public static EncodedRelationship Write(RelationshipDefinition rel, EncodedColumn fk, string fkTableId,
        string fkColumnId, string pkTableId, string pkColumnId, StringComparer strings, IFileSink sink,
        Func<EncodedRelationship, string> tableXml)
    {
        // Row number of every primary key value.
        var rows = new Dictionary<object, int>(new KeyComparer(strings));
        int row = 0;
        foreach (var key in rel.ToColumn.Keys())
        {
            if (key is null)
                throw new InvalidOperationException(
                    $"The relationship '{rel.From.Name}[{rel.FromColumn.Definition.Name}]' -> '{rel.To.Name}[{rel.ToColumn.Definition.Name}]' needs a primary key without blanks, but row {row + 1} of '{rel.To.Name}' has none.");
            if (!rows.TryAdd(key, row))
                throw new InvalidOperationException(
                    $"The relationship '{rel.From.Name}[{rel.FromColumn.Definition.Name}]' -> '{rel.To.Name}[{rel.ToColumn.Definition.Name}]' needs unique primary keys, but '{rel.To.Name}[{rel.ToColumn.Definition.Name}]' contains the value '{key}' more than once.");
            row++;
        }

        long records = fk.IsValueEncoded ? fk.MaxDataId + 1 : fk.LastId + 1;
        var index = new int[records];
        Array.Fill(index, -1);
        var lookup = rel.FromColumn.IdToKey!;
        var counts = rel.FromColumn.IdRowCounts;
        long violations = 0, maxRow = -1;
        for (uint id = DataIds.First; id < records; id++)
        {
            int target = rows.TryGetValue(lookup(id), out int r) ? r : -1;
            index[id] = target;
            if (target > maxRow) maxRow = target;
            if (target < 0 && counts is not null && id < counts.Length) violations += counts[id];
        }

        int width = BitPacking.ChooseWidth((uint)(maxRow + 1));
        var values = new uint[records];
        for (int i = 0; i < values.Length; i++) values[i] = (uint)index[i];

        long words = BitPacking.WordCount(records, width);
        var packed = new ulong[words];
        BitPacking.Pack(values, uint.MaxValue, width, packed); // Min = -1
        var result = new EncodedRelationship
        {
            Name = rel.Name,
            ForeignTableId = fkTableId,
            ForeignColumnId = fkColumnId,
            PrimaryTableId = pkTableId,
            PrimaryColumnId = pkColumnId,
            Records = records,
            MaxRow = maxRow,
            Width = width,
            ViolationCount = violations,
        };

        using (var s = sink.Create(FileNames.RelationshipIndex(fkTableId, rel.Name)))
        {
            var buf = new byte[8 + words * 8];
            BinaryPrimitives.WriteUInt64LittleEndian(buf, (ulong)words);
            for (int i = 0; i < packed.Length; i++) BinaryPrimitives.WriteUInt64LittleEndian(buf.AsSpan(8 + i * 8), packed[i]);
            s.Write(buf);
        }
        using (var s = sink.Create(FileNames.RelationshipTableXml(fkTableId, rel.Name)))
            s.Write(new System.Text.UTF8Encoding(false).GetBytes(tableXml(result)));
        return result;
    }
}
