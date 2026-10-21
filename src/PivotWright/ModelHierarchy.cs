namespace PivotWright;

/// <summary>A hierarchy, added with <see cref="ModelTable.AddHierarchy"/>.</summary>
public sealed class ModelHierarchy
{
    internal ModelHierarchy(ModelTable table, string name, IReadOnlyList<ModelColumn> levels)
    {
        Table = table;
        Name = name;
        Levels = levels;
    }

    /// <summary>Table that holds the level columns.</summary>
    public ModelTable Table { get; }

    /// <summary>Hierarchy name, unique within the table.</summary>
    public string Name { get; }

    /// <summary>The level columns, from the top level down.</summary>
    public IReadOnlyList<ModelColumn> Levels { get; }

    /// <summary>The hierarchy name and its levels, for example <c>Calendar (Year > Quarter > Month)</c>.</summary>
    public override string ToString() => $"{Name} ({string.Join(" > ", Levels.Select(l => l.Name))})";
}
