namespace PivotWright;

/// <summary>A hierarchy, added with <see cref="ModelTable.AddHierarchy"/>.</summary>
public sealed class ModelHierarchy
{
    internal ModelHierarchy(ModelTable table, string name, IReadOnlyList<ModelColumn> levels, IReadOnlyList<string> levelNames)
    {
        Table = table;
        Name = name;
        Levels = levels;
        LevelNames = levelNames;
    }

    /// <summary>Table that holds the level columns.</summary>
    public ModelTable Table { get; }

    /// <summary>Hierarchy name, unique within the table.</summary>
    public string Name { get; }

    /// <summary>The level columns that take row values, from the top level down; calculated columns among the levels are only in <see cref="LevelNames"/>.</summary>
    public IReadOnlyList<ModelColumn> Levels { get; }

    /// <summary>The names of all levels, from the top level down, calculated columns included.</summary>
    public IReadOnlyList<string> LevelNames { get; }

    /// <summary>
    /// True when a level is a calculated column. Excel has not calculated it, so the hierarchy is saved unprocessed and Excel builds
    /// it when the column is calculated.
    /// </summary>
    public bool IsUnprocessed => LevelNames.Count != Levels.Count;

    /// <summary>The hierarchy name and its levels, for example <c>Calendar (Year > Quarter > Month)</c>.</summary>
    public override string ToString() => $"{Name} ({string.Join(" > ", LevelNames)})";
}
