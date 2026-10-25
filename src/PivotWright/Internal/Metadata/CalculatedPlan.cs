using PivotWright.Internal.Model;

namespace PivotWright.Internal.Metadata;

/// <summary>Version counters of one table that Excel changed when a calculated column was added (SPEC section 10).</summary>
internal sealed record CalculatedTableFacts
{
    public required int DimensionVersion { get; init; }
    /// <summary>Counters of the dimension itself, or null when Excel left them as they were (a table saved again without a column).</summary>
    public (int Structure, int DataId, int Cache, int Data)? DimensionCounters { get; init; }
    public int? TableVersion { get; init; }
    public bool HoldsColumns { get; init; }
    /// <summary>The table keeps a dimension information file; not after a hierarchy over a calculated column was made (H31B).</summary>
    public bool DimensionInfo { get; init; } = true;
}

/// <summary>
/// How Excel numbers the objects of a hierarchy model after calculated columns were added in manual calculation mode. Only the
/// shapes of Excel's H10, H12 to H15 and H18 to H20 samples are known; <see cref="Create"/> refuses the others. The counters
/// of one column on the many-side table are those of H10, which are Excel-verified for generated files; the others are fitted to
/// H18 (two columns on one table), H19 (one on a primary-side hierarchy table) and H20 (one on each), and to H25 to H28 (three and
/// four columns): every counter grows by a fixed step per column on the many-side table, so any number of columns is written by
/// formula (Excel's samples go up to four with a hierarchy and three without). A model without a hierarchy has
/// its own numbers (H21 to H24, the same four shapes).
/// </summary>
internal sealed record CalculatedPlan
{
    public required int Cube { get; init; }
    /// <summary>Object version of the measure groups when it differs from the cube's: one step behind after a hide operation (H33).</summary>
    public int? MeasureGroup { get; init; }
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
    /// <summary>The cube folder holds an empty information file; not after a hierarchy over a calculated column was made (H31B).</summary>
    public bool CubeInfo { get; init; } = true;
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
    public bool HasDimensionInfo(string tableId) => Tables.TryGetValue(tableId, out var f) && f.HoldsColumns && f.DimensionInfo;

    private const string Seen =
        "The shapes seen in Excel-written files are: one or more columns on the many-side table, one column on a hierarchy table that is the primary side of a relationship, and both together; more columns on the primary side, and columns on several primary-side tables, follow the same steps but were not seen.";

    private const string SeenPlain =
        "The shapes seen in Excel-written files are: one or more columns on the many-side table, one or more on a primary-side table, and both together.";

    public static CalculatedPlan? Create(IReadOnlyList<TableBuilder> tables, IReadOnlyList<RelationshipDefinition> relationships,
        IReadOnlyList<HierarchyDefinition> hierarchies)
    {
        var plan = CreateWithoutHierarchies(tables, relationships, hierarchies);
        if (plan is null) return null;
        // A hide operation on top of the calculated columns: the cube two steps ahead and the measure groups one step behind it, and the
        // dimension of every table with a hidden column one step up (H33 to H36, which are the straight-line fit plus these steps).
        var hidden = tables.Where(t => t.Columns.Any(c => c.Definition.IsHidden) || t.CalculatedColumns.Any(c => c.IsHidden)).ToList();
        if (hidden.Count > 0)
        {
            var facts = new Dictionary<string, CalculatedTableFacts>(plan.Tables);
            foreach (var t in hidden)
                if (facts.TryGetValue(t.Id, out var f)) facts[t.Id] = f with { DimensionVersion = f.DimensionVersion + 1 };
            plan = plan with { Cube = plan.Cube + 2, MeasureGroup = plan.Cube + 1, Tables = facts };
        }
        // Hierarchies over calculated columns are a further edit each (H31B): the cube, the measure groups, the partitions and the script
        // step up, the table with the hierarchy takes two more table and data versions and three more dimension versions, and the primary
        // side of the first relationship is saved once more. One such hierarchy was seen; more are taken as repeated edits.
        var unprocessed = hierarchies.Where(h => h.IsUnprocessed).ToList();
        if (unprocessed.Count == 0) return plan;
        int count = unprocessed.Count;
        var tablesFacts = new Dictionary<string, CalculatedTableFacts>(plan.Tables);
        foreach (var group in unprocessed.GroupBy(h => h.Table.Id))
        {
            int k = group.Count();
            var f = tablesFacts[group.Key];
            var (s, d, c, data) = f.DimensionCounters!.Value;
            tablesFacts[group.Key] = f with
            {
                DimensionVersion = f.DimensionVersion + 3 * k, DimensionCounters = (s + 2 * k, d, c, data + 2 * k),
                TableVersion = f.TableVersion + 2 * k, DimensionInfo = false,
            };
        }
        foreach (var (id, f) in tablesFacts.ToList())
            if (!f.HoldsColumns) tablesFacts[id] = f with { DimensionVersion = f.DimensionVersion + count };
        var groupData = new Dictionary<string, int>(plan.MeasureGroupData);
        foreach (var t in tables)
            if (relationships.Any(r => r.From == t) && t.CalculatedColumns.Count == 0) groupData[t.Id] = 1;
        return plan with
        {
            Cube = plan.Cube + 7 * count, MeasureGroup = plan.MeasureGroup is { } previous ? previous + 7 * count : null, CubeStructure = plan.CubeStructure + 5 * count,
            MeasureGroupStructure = plan.MeasureGroupStructure + 3 * count, MeasureGroupCache = plan.MeasureGroupCache + 3 * count,
            Partition = plan.Partition + 3 * count, Script = plan.Script + 3 * count,
            Tables = tablesFacts, MeasureGroupData = groupData, CubeInfo = false,
        };
    }

    private static CalculatedPlan? CreateWithoutHierarchies(IReadOnlyList<TableBuilder> tables, IReadOnlyList<RelationshipDefinition> relationships,
        IReadOnlyList<HierarchyDefinition> hierarchies)
    {
        var holders = tables.Where(t => t.CalculatedColumns.Count > 0).ToList();
        if (holders.Count == 0) return null;
        int total = holders.Sum(t => t.CalculatedColumns.Count);
        string what = $"'{holders[0].Name}[{holders[0].CalculatedColumns[0].Name}]'";
        // Hidden columns (data columns or the calculated column itself) were seen in the hierarchy model without a date table only (H33 to H36).
        var hiddenTables = tables.Where(t => t.Columns.Any(c => c.Definition.IsHidden) || t.CalculatedColumns.Any(c => c.IsHidden)).ToList();
        if (hiddenTables.Count > 0 && (hierarchies.Count == 0 || tables.Any(t => t.DateColumn is not null)))
            throw new InvalidOperationException($"{what} is a calculated column in a model with hidden columns, which was only seen in a model with a hierarchy and no date table (no Excel sample of the other cases exists yet).");

        bool plain = hierarchies.Count == 0;
        var dateTable = tables.FirstOrDefault(t => t.DateColumn is not null);
        bool date = dateTable is not null;
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
                continue;
            }
            if (!Hierarchy(t) && !ManySide(t))
                throw new InvalidOperationException($"{name} is a calculated column on a table that has no hierarchy and is not the many side of a relationship, which was never seen in an Excel-written file. " + seen);
            if (Hierarchy(t) && ManySide(t))
                throw new InvalidOperationException($"{name} is a calculated column on a table that has a hierarchy and is the many side of a relationship, which was never seen in an Excel-written file. " + seen);
        }
        var many = holders.Where(ManySide).ToList();
        var primary = holders.Where(t => !ManySide(t)).ToList();
        if (many.Count > 1)
            throw new InvalidOperationException($"The model has calculated columns on {many.Count} many-side tables, which was never seen in an Excel-written file (at most one many-side table can hold columns). " + seen);

        var facts = new Dictionary<string, CalculatedTableFacts>();
        var columns = new Dictionary<(string, string), (int, int, int)>();
        // The primary side of the first relationship is saved again, unless it holds the column itself: in the samples with a column on
        // the many-side table, and (H31, H32) in a model with a date table whatever table holds the column.
        var resaved = many.Count == 1 || date ? relationships.FirstOrDefault()?.To : null;
        var partitionData = new Dictionary<string, int>();
        var groupData = new Dictionary<string, int>();
        foreach (var t in tables)
        {
            // The usual values; the shapes below change some of them.
            partitionData[t.Id] = date || Hierarchy(t) || t.CalculatedColumns.Count > 0 ? -1 : 0;
            groupData[t.Id] = -1;
        }

        // One column on the many-side table in a model with a hierarchy: H10, with the extra step Excel took when the calculation
        // mode was switched in that session.
        if (!plain && !date && total == 1 && many.Count == 1 && hiddenTables.Count == 0)
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
            // With a date table every counter of the many-side table starts one step higher (its base: 4, table version 1, data 1).
            int shift = date ? 1 : 0;
            if (date) dimension = 4 + 3 * n;
            facts[t.Id] = new CalculatedTableFacts
            {
                DimensionVersion = dimension, DimensionCounters = (n, 1 + n, 1 + n, 1 + n + shift), TableVersion = 1 + n + shift, HoldsColumns = true,
            };
            for (int j = 1; j <= n; j++) columns[(t.Id, t.CalculatedColumns[j - 1].Id)] = (1 + j + shift, 1 + j + shift, 1 + j + shift);
        }
        foreach (var t in primary)
        {
            // One column: dimension 9 (3, 2, 2, 3) and table version 3 in the hierarchy model, 9 (4, 3, 3, 4) and 4 for the date table
            // (H31, H32), 5 (1, 2, 2, 1) and 1 without a hierarchy (H23, H24). Each further column steps the dimension by 3 and the
            // counters and the table version by 1, as on the many-side table; only the first column was seen on this side.
            int n = t.CalculatedColumns.Count;
            // Products in a model with a date table is taken as Products without one (no sample; its base is one dimension version higher).
            bool isDateTable = t == dateTable;
            var (dimension, s, d, c, data, tableVersion) = plain ? (5, 1, 2, 2, 1, 1) : isDateTable ? (9, 4, 3, 3, 4, 4) : (9, 3, 2, 2, 3, 3);
            facts[t.Id] = new CalculatedTableFacts
            {
                DimensionVersion = dimension + 3 * (n - 1), DimensionCounters = (s + n - 1, d + n - 1, c + n - 1, data + n - 1),
                TableVersion = tableVersion + n - 1, HoldsColumns = true,
            };
            int first = plain ? 1 : isDateTable ? 4 : 3;
            for (int j = 1; j <= n; j++) columns[(t.Id, t.CalculatedColumns[j - 1].Id)] = (first + j - 1, first + j - 1, first + j - 1);
        }
        if (resaved is not null && !facts.ContainsKey(resaved.Id))
            facts[resaved.Id] = new CalculatedTableFacts { DimensionVersion = plain ? 2 : date ? 7 : 6 };

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

        if (date)
        {
            // Excel's numbers with a date table are those of the hierarchy model moved up by the date table's own steps (H29 to H32).
            return new CalculatedPlan
            {
                Cube = 22 + 4 * total, CubeStructure = 12 + 5 * total, MeasureGroupStructure = 9 + total, MeasureGroupCache = 2 + total,
                Partition = 10 + total, PartitionCache = 1 + total, Script = 10 + total,
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
