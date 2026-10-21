using System.Collections;

namespace PivotWright;

/// <summary>
/// Objects of the model in the order they were added, also found by name: <c>model.Tables["Sales"]</c>,
/// <c>sales.Columns["Qty"]</c>. Names are compared ignoring case.
/// </summary>
public sealed class ModelCollection<T> : IReadOnlyList<T> where T : class
{
    private readonly IReadOnlyList<T> _items;
    private readonly Func<string, T?> _find;
    private readonly Func<string, string> _notFound;

    internal ModelCollection(IReadOnlyList<T> items, Func<string, T?> find, Func<string, string> notFound)
    {
        _items = items;
        _find = find;
        _notFound = notFound;
    }

    /// <summary>The object at <paramref name="index"/>, in the order added.</summary>
    public T this[int index] => _items[index];

    /// <summary>The object with this name.</summary>
    /// <exception cref="KeyNotFoundException">There is none; the message lists the names that exist.</exception>
    public T this[string name]
    {
        get
        {
            ArgumentNullException.ThrowIfNull(name);
            return _find(name) ?? throw new KeyNotFoundException(_notFound(name));
        }
    }

    /// <summary>True when an object with this name exists.</summary>
    public bool Contains(string name) => name is not null && _find(name) is not null;

    /// <summary>Number of objects.</summary>
    public int Count => _items.Count;

    public IEnumerator<T> GetEnumerator() => _items.GetEnumerator();

    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
}
