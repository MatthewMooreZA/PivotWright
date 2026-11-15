namespace PivotWright;

/// <summary>A many-to-one relationship, added with <see cref="DataModelWriter.AddRelationship"/>.</summary>
public sealed class ModelRelationship
{
    internal ModelRelationship(ModelColumn manyColumn, ModelColumn oneColumn, bool isActive)
    {
        ManyColumn = manyColumn;
        OneColumn = oneColumn;
        IsActive = isActive;
    }

    /// <summary>Foreign key column, of the "many" table (for example Sales[ProductKey]).</summary>
    public ModelColumn ManyColumn { get; }

    /// <summary>Key column of the "one" table (for example Products[ProductKey]); unique and never blank.</summary>
    public ModelColumn OneColumn { get; }

    /// <summary>
    /// True when the relationship filters the "many" table by the "one" table. A second relationship between two tables, for example
    /// Sales[Ship Date] next to Sales[Date] to the date table, is inactive: a measure uses it with <c>USERELATIONSHIP</c>.
    /// </summary>
    public bool IsActive { get; }

    /// <summary>The relationship as Power Pivot's diagram describes it, for example <c>'Sales'[ProductKey] -> 'Products'[ProductKey]</c>.</summary>
    public override string ToString() => $"{ManyColumn} -> {OneColumn}";
}
