using PivotWright.Internal.Model;

namespace PivotWright;

/// <summary>
/// A perspective, added with <see cref="DataModelWriter.AddPerspective"/>: a named subset of the model's columns, measures and
/// hierarchies, like the ones Power Pivot's "Perspectives" dialog (Advanced > Perspectives) creates. Add what the perspective
/// shows with <see cref="Include(ModelColumn)"/> and its overloads.
/// </summary>
/// <remarks>
/// Written as Excel saves perspectives (SPEC section 13, <c>H64</c> to <c>H66</c>). Nothing selects a perspective for a PivotTable yet:
/// no way was found in Excel to bind one to a perspective (SPEC section 13, open question), so for now the perspectives are only listed in the model.
/// </remarks>
public sealed class ModelPerspective
{
    private readonly PerspectiveDefinition _definition;
    private readonly List<ModelColumn> _columns = new();
    private readonly List<ModelMeasure> _measures = new();
    private readonly List<ModelHierarchy> _hierarchies = new();

    internal ModelPerspective(DataModelWriter model, PerspectiveDefinition definition)
    {
        Model = model;
        _definition = definition;
    }

    internal DataModelWriter Model { get; }
    internal PerspectiveDefinition Definition => _definition;

    /// <summary>Perspective name, unique in the model (ignoring case).</summary>
    public string Name => _definition.Name;

    /// <summary>The columns shown, in the order included.</summary>
    public IReadOnlyList<ModelColumn> Columns => _columns;

    /// <summary>The measures shown, in the order included.</summary>
    public IReadOnlyList<ModelMeasure> Measures => _measures;

    /// <summary>The hierarchies shown, in the order included.</summary>
    public IReadOnlyList<ModelHierarchy> Hierarchies => _hierarchies;

    /// <summary>
    /// Shows a column, like ticking it in Power Pivot's Perspectives dialog. Including a column twice does nothing.
    /// </summary>
    /// <param name="column">A column of this model.</param>
    /// <returns>This perspective, so calls can be chained.</returns>
    /// <exception cref="ArgumentNullException"><paramref name="column"/> is null.</exception>
    /// <exception cref="ArgumentException">The column belongs to another model.</exception>
    public ModelPerspective Include(ModelColumn column)
    {
        Model.ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(column);
        if (column.Table.Model != Model)
            throw new ArgumentException($"Column {column} belongs to another DataModelWriter.", nameof(column));
        if (_columns.Contains(column)) return this;
        _columns.Add(column);
        _definition.Columns.Add((column.Table.Builder, column.Builder));
        return this;
    }

    /// <summary>
    /// Shows a measure. A hierarchy's level columns are not included with it, as in Excel's dialog. Including a measure twice does nothing.
    /// </summary>
    /// <param name="measure">A measure of this model.</param>
    /// <returns>This perspective, so calls can be chained.</returns>
    /// <exception cref="ArgumentNullException"><paramref name="measure"/> is null.</exception>
    /// <exception cref="ArgumentException">The measure belongs to another model.</exception>
    public ModelPerspective Include(ModelMeasure measure)
    {
        Model.ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(measure);
        if (measure.Table.Model != Model)
            throw new ArgumentException($"Measure {measure} belongs to another DataModelWriter.", nameof(measure));
        if (_measures.Contains(measure)) return this;
        _measures.Add(measure);
        _definition.Measures.Add((measure.Table.Builder, measure.Definition));
        return this;
    }

    /// <summary>
    /// Shows a hierarchy. Excel's dialog does not add the hierarchy's level columns; include them as well to show them on their own.
    /// Including a hierarchy twice does nothing.
    /// </summary>
    /// <param name="hierarchy">A hierarchy of this model.</param>
    /// <returns>This perspective, so calls can be chained.</returns>
    /// <exception cref="ArgumentNullException"><paramref name="hierarchy"/> is null.</exception>
    /// <exception cref="ArgumentException">The hierarchy belongs to another model.</exception>
    public ModelPerspective Include(ModelHierarchy hierarchy)
    {
        Model.ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(hierarchy);
        if (hierarchy.Table.Model != Model)
            throw new ArgumentException($"Hierarchy {hierarchy} of '{hierarchy.Table.Name}' belongs to another DataModelWriter.", nameof(hierarchy));
        if (_hierarchies.Contains(hierarchy)) return this;
        _hierarchies.Add(hierarchy);
        _definition.Hierarchies.Add((hierarchy.Table.Builder, hierarchy.Name));
        return this;
    }

    /// <summary>
    /// Shows a whole table, like ticking the table in Power Pivot's Perspectives dialog: every column, every visible measure and every
    /// hierarchy of the table (calculated columns are not supported in perspectives yet).
    /// </summary>
    /// <param name="table">A table of this model.</param>
    /// <returns>This perspective, so calls can be chained.</returns>
    /// <exception cref="ArgumentNullException"><paramref name="table"/> is null.</exception>
    /// <exception cref="ArgumentException">The table belongs to another model.</exception>
    public ModelPerspective Include(ModelTable table)
    {
        Model.ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(table);
        if (table.Model != Model)
            throw new ArgumentException($"Table '{table.Name}' belongs to another DataModelWriter.", nameof(table));
        foreach (var column in table.Columns) Include(column);
        foreach (var measure in table.Measures.Where(m => !m.IsHidden)) Include(measure);
        foreach (var hierarchy in table.Hierarchies) Include(hierarchy);
        return this;
    }

    /// <summary>The perspective and what it shows, for example <c>Sales view (3 columns, 1 measure, 0 hierarchies)</c>.</summary>
    /// <returns>The description.</returns>
    public override string ToString() =>
        $"{Name} ({_columns.Count} column{(_columns.Count == 1 ? "" : "s")}, {_measures.Count} measure{(_measures.Count == 1 ? "" : "s")}, {_hierarchies.Count} hierarch{(_hierarchies.Count == 1 ? "y" : "ies")})";
}
