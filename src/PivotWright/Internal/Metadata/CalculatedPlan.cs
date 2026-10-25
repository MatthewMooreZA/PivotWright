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
/// H18 (two columns on one table), H19 (one on a primary-side hierarchy table) and H20 (one on each).
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
    /// <summary>Only H10 keeps the partition information of the table with the column; the other samples have none.</summary>
    public required bool PartitionInfo { get; init; }
    public required IReadOnlyDictionary<string, CalculatedTableFacts> Tables { get; init; }
    /// <summary>Tables whose partition keeps a data version of 1 (H10) instead of -1.</summary>
    public required IReadOnlySet<string> PartitionDataOne { get; init; }
    /// <summary>The version of the data file and dictionary of each column: (file prefix, dictionary data version, file data version).</summary>
    public required IReadOnlyDictionary<(string TableId, string ColumnId), (int Prefix, int Dictionary, int File)> Columns { get; init; }

    public bool Resaved(string tableId) => Tables.ContainsKey(tableId);
    public bool HoldsColumns(string tableId) => Tables.TryGetValue(tableId, out var f) && f.HoldsColumns;

    private const string Seen =
        "The shapes seen in Excel-written files are: one column on the many-side table, two columns on the many-side table, one column on a hierarchy table that is the primary side of a relationship, and one column on each of those two tables.";

    public static CalculatedPlan? Create(IReadOnlyList<TableBuilder> tables, IReadOnlyList<RelationshipDefinition> relationships,
        IReadOnlyList<HierarchyDefinition> hierarchies)
    {
        var holders = tables.Where(t => t.CalculatedColumns.Count > 0).ToList();
        if (holders.Count == 0) return null;
        int total = holders.Sum(t => t.CalculatedColumns.Count);
        string what = $"'{holders[0].Name}[{holders[0].CalculatedColumns[0].Name}]'";
        if (hierarchies.Count == 0)
            throw new InvalidOperationException($"{what} is a calculated column, which is only supported for models that have a hierarchy (no Excel sample of the other case exists yet).");
        if (tables.Any(t => t.DateColumn is not null))
            throw new InvalidOperationException($"{what} is a calculated column, which is not supported together with a date table yet.");
        if (tables.Any(t => t.Columns.Any(c => c.Definition.IsHidden)))
            throw new InvalidOperationException($"{what} is a calculated column, which is not supported together with hidden columns yet.");

        bool Hierarchy(TableBuilder t) => hierarchies.Any(h => h.Table == t);
        bool ManySide(TableBuilder t) => relationships.Any(r => r.From == t);
        foreach (var t in holders)
        {
            string name = $"'{t.Name}[{t.CalculatedColumns[0].Name}]'";
            if (!Hierarchy(t) && !ManySide(t))
                throw new InvalidOperationException($"{name} is a calculated column on a table that has no hierarchy and is not the many side of a relationship, which was never seen in an Excel-written file. " + Seen);
            if (Hierarchy(t) && ManySide(t))
                throw new InvalidOperationException($"{name} is a calculated column on a table that has a hierarchy and is the many side of a relationship, which was never seen in an Excel-written file. " + Seen);
            if (Hierarchy(t) && t.CalculatedColumns.Count > 1)
                throw new InvalidOperationException($"'{t.Name}' has {t.CalculatedColumns.Count} calculated columns, but only one per hierarchy table was seen in an Excel-written file. " + Seen);
        }
        var many = holders.Where(t => !Hierarchy(t)).ToList();
        var primary = holders.Where(Hierarchy).ToList();
        if (many.Count > 1 || primary.Count > 1 || total > 2)
            throw new InvalidOperationException($"The model has {total} calculated columns on {holders.Count} tables, which was never seen in an Excel-written file. " + Seen);

        var facts = new Dictionary<string, CalculatedTableFacts>();
        var columns = new Dictionary<(string, string), (int, int, int)>();
        // The primary side of the first relationship of the many-side table is saved again, whether or not it has a column itself.
        var resaved = many.Count == 1 ? relationships.First(r => r.From == many[0]).To : null;

        // One column on the many-side table: H10, with the extra step Excel took when the calculation mode was switched.
        if (total == 1 && many.Count == 1)
        {
            var t = many[0];
            facts[t.Id] = new CalculatedTableFacts { DimensionVersion = 7, DimensionCounters = (1, 2, 2, 3), TableVersion = 3, HoldsColumns = true };
            facts[resaved!.Id] = new CalculatedTableFacts { DimensionVersion = 6 };
            columns[(t.Id, t.CalculatedColumns[0].Id)] = (2, 3, 2);
            return new CalculatedPlan
            {
                Cube = 25, CubeStructure = 17, MeasureGroupStructure = 10, MeasureGroupCache = 3, Partition = 11, PartitionCache = 2, Script = 11,
                PartitionInfo = true, Tables = facts, PartitionDataOne = new HashSet<string> { t.Id }, Columns = columns,
            };
        }

        foreach (var t in many)
        {
            int n = t.CalculatedColumns.Count;
            facts[t.Id] = new CalculatedTableFacts
            {
                DimensionVersion = 3 + 3 * n, DimensionCounters = (n, 1 + n, 1 + n, 1 + n), TableVersion = 1 + n, HoldsColumns = true,
            };
            for (int j = 1; j <= n; j++) columns[(t.Id, t.CalculatedColumns[j - 1].Id)] = (1 + j, 1 + j, 1 + j);
        }
        foreach (var t in primary)
        {
            facts[t.Id] = new CalculatedTableFacts { DimensionVersion = 9, DimensionCounters = (3, 2, 2, 3), TableVersion = 3, HoldsColumns = true };
            columns[(t.Id, t.CalculatedColumns[0].Id)] = (3, 3, 3);
        }
        if (resaved is not null && !facts.ContainsKey(resaved.Id))
            facts[resaved.Id] = new CalculatedTableFacts { DimensionVersion = 6 };

        int extra = total - 1;
        return new CalculatedPlan
        {
            Cube = 23 + 4 * extra, CubeStructure = 16 + 5 * extra, MeasureGroupStructure = 9 + extra, MeasureGroupCache = 2 + extra,
            Partition = 10 + extra, PartitionCache = 2 + extra, Script = 10 + extra,
            PartitionInfo = false, Tables = facts, PartitionDataOne = new HashSet<string>(), Columns = columns,
        };
    }
}
