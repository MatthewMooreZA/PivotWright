using PivotWright.Internal.Model;

namespace PivotWright;

/// <summary>A DAX measure, added with <see cref="ModelTable.AddMeasure"/>.</summary>
public sealed class ModelMeasure
{
    private readonly MeasureDefinition _definition;

    internal ModelMeasure(ModelTable table, MeasureDefinition definition)
    {
        Table = table;
        _definition = definition;
    }

    /// <summary>Table the measure belongs to; it appears under this table in the field list.</summary>
    public ModelTable Table { get; }

    /// <summary>Measure name, unique across the model.</summary>
    public string Name => _definition.Name;

    /// <summary>DAX expression, without the leading "=".</summary>
    public string Expression => _definition.Expression;

    /// <summary>Number format.</summary>
    public MeasureFormat Format => _definition.Format;

    /// <summary>Hides the measure from client tools, for a measure that only other measures use.</summary>
    public bool IsHidden
    {
        get => _definition.Hidden;
        set
        {
            Table.Model.ThrowIfDisposed();
            _definition.Hidden = value;
        }
    }

    /// <summary>The measure in DAX notation, for example <c>'Sales'[Total Qty]</c>.</summary>
    public override string ToString() => $"'{Table.Name.Replace("'", "''")}'[{Name}]";
}
