using PivotWright.Internal.Model;

namespace PivotWright.Internal.Metadata;

/// <summary>Version counters of one table that Excel changed when a calculated column was added (SPEC section 10).</summary>
internal sealed class CalculatedTableFacts
{
    public required int DimensionVersion { get; init; }
    /// <summary>Counters of the dimension itself, or null when Excel left them as they were (a table saved again without a column).</summary>
    public (int Structure, int DataId, int Cache, int Data)? DimensionCounters { get; init; }
    public int? TableVersion { get; init; }
    public bool HoldsColumns { get; init; }
}

/// <summary>
/// How Excel numbers the objects of a hierarchy model after calculated columns were added in manual calculation mode. Only the
/// shapes of Excel's H10, H12 to H15 and H18 to H20 samples are known; <see cref="Create"/> refuses the others. The counters
/// of one column on the many-side table are those of H10, which are Excel-verified for generated files; the others are fitted to
/// H18 (two columns on one table), H19 (one on a primary-side hierarchy table) and H20 (one on each). A model without a hierarchy has
/// its own numbers (H21 to H24, the same four shapes).
/// </summary>
internal sealed class CalculatedPlan
{
    public required int Cube { get; init; }
    public required int CubeStructure { get; init; }
    public required int MeasureGroupStructure { get; init; }
    public required int MeasureGroupCache { get; init; }
    public required int Partition { get; init; }
    public required int PartitionCache { get; init; }
    public required int Script { get; init; }
    /// <summary>
    /// The measure groups hold the formula of each calculated column, and the many-side table's group the formulas of the primary
    /// tables' columns (every Excel sample except H10; kept off for the shapes of the hierarchy model, which are verified without).
    /// </summary>
    public bool FormulasInMeasureGroups { get; init; }
    /// <summary>Only H10 keeps the partition information of the table with the column; the other samples have none.</summary>
    public required bool PartitionInfo { get; init; }
    public required IReadOnlyDictionary<string, CalculatedTableFacts> Tables { get; init; }
    /// <summary>Data version of each table's partition and measure group dimension (SPEC section 10).</summary>
    public required IReadOnlyDictionary<string, int> PartitionData { get; init; }
    public required IReadOnlyDictionary<string, int> MeasureGroupData { get; init; }
    /// <summary>The version of the data file and dictionary of each column: (file prefix, dictionary data version, file data version).</summary>
    public required IReadOnlyDictionary<(string TableId, string ColumnId), (int Prefix, int Dictionary, int File)> Columns { get; init; }

    public bool Resaved(string tableId) => Tables.ContainsKey(tableId);
    public bool HoldsColumns(string tableId) => Tables.TryGetValue(tableId, out var f) && f.HoldsColumns;

    /// <summary>Most columns on the many-side table; Excel's samples go up to four (H26), and the counters grow in a straight line.</summary>
    private const int MaxOnManySide = 4;

    private const string Seen =
        "The shapes seen in Excel-written files are: one to four columns on the many-side table, one column on a hierarchy table that is the primary side of a relationship, and both together.";

    private const string SeenPlain =
        "The shapes seen in Excel-written files are: one to four columns on the many-side table, one column on a primary-side table, and both together.";

    public static CalculatedPlan? Create(IReadOnlyList<TableBuilder> tables, IReadOnlyList<RelationshipDefinition> relationships,
        IReadOnlyList<HierarchyDefinition> hierarchies)
    {
        var holders = tables.Where(t => t.CalculatedColumns.Count > 0).ToList();
        if (holders.Count == 0) return null;
        int total = holders.Sum(t => t.CalculatedColumns.Count);
        string what = $"'{holders[0].Name}[{holders[0].CalculatedColumns[0].Name}]'";
        if (tables.Any(t => t.DateColumn is not null))
            throw new InvalidOperationException($"{what} is a calculated column, which is not supported together with a date table yet.");
        if (tables.Any(t => t.Columns.Any(c => c.Definition.IsHidden)))
            throw new InvalidOperationException($"{what} is a calculated column, which is not supported together with hidden columns yet.");

        bool plain = hierarchies.Count == 0;
        bool Hierarchy(TableBuilder t) => hierarchies.Any(h => h.Table == t);
        bool ManySide(TableBuilder t) => relationships.Any(r => r.From == t);
        string seen = plain ? SeenPlain : Seen;
        foreach (var t in holders)
        {
            string name = $"'{t.Name}[{t.CalculatedColumns[0].Name}]'";
            if (plain)
            {
                if (!ManySide(t) && !relationships.Any(r => r.To == t))
                    throw new InvalidOperationException($"{name} is a calculated column on a table that has no relationship, which was never seen in an Excel-written file. " + seen);
                if (!ManySide(t) && t.CalculatedColumns.Count > 1)
                    throw new InvalidOperationException($"'{t.Name}' has {t.CalculatedColumns.Count} calculated columns, but only one per primary-side table was seen in an Excel-written file. " + seen);
                continue;
            }
            if (!Hierarchy(t) && !ManySide(t))
                throw new InvalidOperationException($"{name} is a calculated column on a table that has no hierarchy and is not the many side of a relationship, which was never seen in an Excel-written file. " + seen);
            if (Hierarchy(t) && ManySide(t))
                throw new InvalidOperationException($"{name} is a calculated column on a table that has a hierarchy and is the many side of a relationship, which was never seen in an Excel-written file. " + seen);
            if (Hierarchy(t) && t.CalculatedColumns.Count > 1)
                throw new InvalidOperationException($"'{t.Name}' has {t.CalculatedColumns.Count} calculated columns, but only one per hierarchy table was seen in an Excel-written file. " + seen);
        }
        var many = holders.Where(ManySide).ToList();
        var primary = holders.Where(t => !ManySide(t)).ToList();
        if (many.Count > 1 || primary.Count > 1 || (many.Count == 1 && many[0].CalculatedColumns.Count > MaxOnManySide))
            throw new InvalidOperationException($"The model has {total} calculated columns on {holders.Count} tables, which was never seen in an Excel-written file (at most {MaxOnManySide} on the many-side table and one on a primary-side table). " + seen);

        var facts = new Dictionary<string, CalculatedTableFacts>();
        var columns = new Dictionary<(string, string), (int, int, int)>();
        // The primary side of the first relationship of the many-side table is saved again, whether or not it has a column itself.
        var resaved = many.Count == 1 ? relationships.First(r => r.From == many[0]).To : null;
        var partitionData = new Dictionary<string, int>();
        var groupData = new Dictionary<string, int>();
        foreach (var t in tables)
        {
            // The usual values; the shapes below change some of them.
            partitionData[t.Id] = Hierarchy(t) || t.CalculatedColumns.Count > 0 ? -1 : 0;
            groupData[t.Id] = -1;
        }

        // One column on the many-side table in a model with a hierarchy: H10, with the extra step Excel took when the calculation
        // mode was switched in that session.
        if (!plain && total == 1 && many.Count == 1)
        {
            var t = many[0];
            facts[t.Id] = new CalculatedTableFacts { DimensionVersion = 7, DimensionCounters = (1, 2, 2, 3), TableVersion = 3, HoldsColumns = true };
            facts[resaved!.Id] = new CalculatedTableFacts { DimensionVersion = 6 };
            columns[(t.Id, t.CalculatedColumns[0].Id)] = (2, 3, 2);
            partitionData[t.Id] = 1;
            return new CalculatedPlan
            {
                Cube = 25, CubeStructure = 17, MeasureGroupStructure = 10, MeasureGroupCache = 3, Partition = 11, PartitionCache = 2, Script = 11,
                PartitionInfo = true, Tables = facts, PartitionData = partitionData, MeasureGroupData = groupData, Columns = columns,
            };
        }

        foreach (var t in many)
        {
            int n = t.CalculatedColumns.Count;
            // Beside a column on the primary side, the many-side table in a model without a hierarchy is one step behind (H24 against H21).
            int dimension = plain ? (n == 1 ? (primary.Count == 1 ? 5 : 6) : 2 + 3 * n) : 3 + 3 * n;
            facts[t.Id] = new CalculatedTableFacts
            {
                DimensionVersion = dimension, DimensionCounters = (n, 1 + n, 1 + n, 1 + n), TableVersion = 1 + n, HoldsColumns = true,
            };
            for (int j = 1; j <= n; j++) columns[(t.Id, t.CalculatedColumns[j - 1].Id)] = (1 + j, 1 + j, 1 + j);
        }
        foreach (var t in primary)
        {
            // A model with a hierarchy: dimension 9 (3, 2, 2, 3) and table version 3; without one: 5 (1, 2, 2, 1) and 1.
            facts[t.Id] = plain
                ? new CalculatedTableFacts { DimensionVersion = 5, DimensionCounters = (1, 2, 2, 1), TableVersion = 1, HoldsColumns = true }
                : new CalculatedTableFacts { DimensionVersion = 9, DimensionCounters = (3, 2, 2, 3), TableVersion = 3, HoldsColumns = true };
            columns[(t.Id, t.CalculatedColumns[0].Id)] = plain ? (1, 1, 1) : (3, 3, 3);
        }
        if (resaved is not null && !facts.ContainsKey(resaved.Id))
            facts[resaved.Id] = new CalculatedTableFacts { DimensionVersion = plain ? 2 : 6 };

        if (plain)
        {
            // H21 (one on Sales), H22 (two on Sales), H23 (one on Products), H24 (one on each), H28 (three on Sales). A single column on Sales
            // alone is one step ahead of the straight line the others follow, in the cube and in the dimension.
            foreach (var t in primary) { partitionData[t.Id] = 1; groupData[t.Id] = 1; }
            if (total == 1 && many.Count == 1) partitionData[many[0].Id] = 1;
            if (total >= 2) foreach (var t in tables) partitionData[t.Id] = -1;
            return new CalculatedPlan
            {
                Cube = total == 1 && many.Count == 1 ? 7 : 2 + 4 * total, CubeStructure = 5 * total,
                MeasureGroupStructure = total, MeasureGroupCache = 1 + total,
                Partition = 1 + total, PartitionCache = 1 + total, Script = 1 + total,
                PartitionInfo = false, Tables = facts, PartitionData = partitionData, MeasureGroupData = groupData, Columns = columns,
                FormulasInMeasureGroups = true,
            };
        }

        int extra = total - 1;
        return new CalculatedPlan
        {
            Cube = 23 + 4 * extra, CubeStructure = 16 + 5 * extra, MeasureGroupStructure = 9 + extra, MeasureGroupCache = 2 + extra,
            Partition = 10 + extra, PartitionCache = 2 + extra, Script = 10 + extra,
            PartitionInfo = false, Tables = facts, PartitionData = partitionData, MeasureGroupData = groupData, Columns = columns,
        };
    }
}
