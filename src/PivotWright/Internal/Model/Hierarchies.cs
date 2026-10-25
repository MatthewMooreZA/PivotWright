using System.Buffers.Binary;
using PivotWright.Internal.VertiPaq;

namespace PivotWright.Internal.Model;

/// <summary>One level of a hierarchy: a data column, or a calculated column (then <see cref="Column"/> is null).</summary>
internal sealed record HierarchyLevel(string Name, string Id, ColumnBuilder? Column);

/// <summary>A user hierarchy (for example Category > Subcategory > Product) as declared through the public API.</summary>
internal sealed class HierarchyDefinition
{
    public required string Name { get; init; }
    public required TableBuilder Table { get; init; }
    public required IReadOnlyList<HierarchyLevel> Levels { get; init; }
    /// <summary>The data columns among the levels, from the top down (all of them unless a level is a calculated column).</summary>
    public IReadOnlyList<ColumnBuilder> Columns => Levels.Where(l => l.Column is not null).Select(l => l.Column!).ToList();
    /// <summary>
    /// A level is a calculated column, which has no values yet: Excel keeps such a hierarchy unprocessed, without a tree
    /// (SPEC section 10, H31B).
    /// </summary>
    public bool IsUnprocessed => Levels.Any(l => l.Column is null);
}

/// <summary>A hierarchy after its tree was written.</summary>
internal sealed class EncodedHierarchy
{
    public required string Name { get; init; }
    public required string TableId { get; init; }
    /// <summary>Name of the U$ table that holds the tree.</summary>
    public required string TableName { get; init; }
    public required IReadOnlyList<EncodedColumn> Levels { get; init; }
    /// <summary>Position of the first node of every level in the tree.</summary>
    public required IReadOnlyList<long> LevelStarts { get; init; }
    public required long Records { get; init; }
    /// <summary>
    /// True when every member of a level has exactly one parent (a Product belongs to one Category). Excel writes
    /// StructureType Unnatural otherwise, for example for Quarter Q4 under two years (SPEC section 7).
    /// </summary>
    public bool Natural { get; init; } = true;
    /// <summary>The hierarchy has a calculated level: no tree, level starts -1, StructureType Unknown.</summary>
    public bool Unprocessed { get; init; }

    public static readonly string[] ColumnNames = { "MULTI_LEVEL_ID", "PARENT_POS", "FIRST_CHILD_POS", "CHILD_COUNT" };
}

internal static class HierarchyEncoder
{
    /// <summary>
    /// A hierarchy with a calculated level as Excel saves it before the column was calculated (H31B): no U$ table or files, every level
    /// starts at -1 and the structure is unknown.
    /// </summary>
    public static EncodedHierarchy Unprocessed(HierarchyDefinition def, EncodedTable table)
    {
        var levels = def.Levels.Select(l => l.Column is { } c
            ? table.Columns.First(e => e.Id == c.Id)
            : new EncodedColumn { Name = l.Name, Id = l.Id, DataType = ColumnDataType.Int64 }).ToList();
        return new EncodedHierarchy
        {
            Name = def.Name,
            TableId = table.Id,
            TableName = string.Empty,
            Levels = levels,
            LevelStarts = levels.Select(_ => -1L).ToList(),
            Records = 0,
            Natural = false,
            Unprocessed = true,
        };
    }

    /// <summary>
    /// Builds the tree of distinct value paths of a hierarchy and writes it as the four columns Excel uses:
    /// the data id of the node's value, its parent, its first child and its child count. Nodes are stored level
    /// by level; within a level they follow their parents' order and, below one parent, the sort order of the
    /// level column (BLANK first).
    /// </summary>
    public static EncodedHierarchy Write(HierarchyDefinition def, EncodedTable table, StringComparer strings, IFileSink sink,
        Func<EncodedHierarchy, string> tableXml)
    {
        int levels = def.Columns.Count;
        var keyToId = new Dictionary<object, uint>[levels];
        var rankOfId = new Dictionary<uint, int>[levels];
        for (int l = 0; l < levels; l++)
        {
            var col = def.Columns[l];
            var sorted = col.SortedIds ?? throw new InvalidOperationException("The hierarchy column was not encoded.");
            keyToId[l] = new Dictionary<object, uint>(new KeyComparer(strings));
            rankOfId[l] = new Dictionary<uint, int>(sorted.Length);
            for (int p = 0; p < sorted.Length; p++)
            {
                rankOfId[l][sorted[p]] = p;
                if (sorted[p] != DataIds.Null) keyToId[l].TryAdd(col.IdToKey!(sorted[p]), sorted[p]);
            }
        }

        // One node per distinct (level 1 value, ..., level n value) prefix.
        var nodeOf = new Dictionary<ulong, int>[levels];
        var nodeId = new List<uint>[levels];
        var nodeParent = new List<int>[levels];
        for (int l = 0; l < levels; l++) { nodeOf[l] = new(); nodeId[l] = new(); nodeParent[l] = new(); }
        var iters = def.Columns.Select(c => c.Keys().GetEnumerator()).ToArray();
        var firstParentId = new Dictionary<uint, uint>[levels];
        for (int l = 0; l < levels; l++) firstParentId[l] = new();
        bool natural = true;
        try
        {
            for (long row = 0; row < def.Table.Rows; row++)
            {
                int parent = -1;
                uint parentId = DataIds.Null;
                for (int l = 0; l < levels; l++)
                {
                    if (!iters[l].MoveNext()) throw new InvalidOperationException("A hierarchy column has fewer rows than the table.");
                    object? key = iters[l].Current;
                    uint id = key is null ? DataIds.Null : keyToId[l][key];
                    if (l > 0 && natural && !firstParentId[l].TryAdd(id, parentId) && firstParentId[l][id] != parentId) natural = false;
                    parentId = id;
                    ulong k = ((ulong)(uint)(parent + 1) << 32) | id;
                    if (!nodeOf[l].TryGetValue(k, out int idx))
                    {
                        idx = nodeId[l].Count;
                        nodeOf[l][k] = idx;
                        nodeId[l].Add(id);
                        nodeParent[l].Add(parent);
                    }
                    parent = idx;
                }
            }
        }
        finally { foreach (var it in iters) it.Dispose(); }

        // Final order: by parent position, then by the sort rank of the value.
        var order = new int[levels][];
        var position = new int[levels][];
        var starts = new long[levels + 1];
        for (int l = 0; l < levels; l++)
        {
            int n = nodeId[l].Count;
            var keys = new long[n];
            order[l] = new int[n];
            for (int i = 0; i < n; i++)
            {
                long parentPos = l == 0 ? 0 : position[l - 1][nodeParent[l][i]];
                keys[i] = (parentPos << 32) | (uint)rankOfId[l][nodeId[l][i]];
                order[l][i] = i;
            }
            Array.Sort(keys, order[l]);
            position[l] = new int[n];
            for (int i = 0; i < n; i++) position[l][order[l][i]] = i;
            starts[l + 1] = starts[l] + n;
        }

        long total = starts[levels];
        var multiLevelId = new uint[total];
        var parentPosCol = new uint[total];
        var firstChild = new uint[total];
        var childCount = new uint[total];
        Array.Fill(parentPosCol, uint.MaxValue);
        Array.Fill(firstChild, uint.MaxValue);
        Array.Fill(childCount, uint.MaxValue);
        for (int l = 0; l < levels; l++)
        {
            for (int i = 0; i < order[l].Length; i++)
            {
                int node = order[l][i];
                long global = starts[l] + i;
                multiLevelId[global] = nodeId[l][node];
                if (l == 0) continue;
                long parentGlobal = starts[l - 1] + position[l - 1][nodeParent[l][node]];
                parentPosCol[global] = (uint)parentGlobal;
                if (firstChild[parentGlobal] == uint.MaxValue) { firstChild[parentGlobal] = (uint)global; childCount[parentGlobal] = 0; }
                childCount[parentGlobal]++;
            }
        }

        var encodedLevels = def.Columns.Select(c => table.Columns.First(e => e.Id == c.Id)).ToList();
        var result = new EncodedHierarchy
        {
            Name = def.Name,
            TableId = table.Id,
            TableName = FileNames.UserHierarchyTable(table.Id, def.Name),
            Levels = encodedLevels,
            LevelStarts = starts.Take(levels).ToArray(),
            Records = total,
            Natural = natural,
        };

        void WriteColumn(string name, uint[] values, uint min)
        {
            long words = BitPacking.WordCount(total, 32);
            var packed = new ulong[words];
            BitPacking.Pack(values, min, 32, packed);
            var buf = new byte[8 + words * 8];
            BinaryPrimitives.WriteUInt64LittleEndian(buf, (ulong)words);
            for (int i = 0; i < packed.Length; i++) BinaryPrimitives.WriteUInt64LittleEndian(buf.AsSpan(8 + i * 8), packed[i]);
            using var s = sink.Create(FileNames.UserHierarchyColumn(table.Id, def.Name, name));
            s.Write(buf);
        }
        WriteColumn("MULTI_LEVEL_ID", multiLevelId, 0);
        WriteColumn("PARENT_POS", parentPosCol, uint.MaxValue);
        WriteColumn("FIRST_CHILD_POS", firstChild, uint.MaxValue);
        WriteColumn("CHILD_COUNT", childCount, uint.MaxValue);
        using (var s = sink.Create(FileNames.UserHierarchyTableXml(table.Id, def.Name)))
            s.Write(new System.Text.UTF8Encoding(false).GetBytes(tableXml(result)));
        return result;
    }
}
