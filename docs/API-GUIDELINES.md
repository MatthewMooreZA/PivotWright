# PivotWright API guidelines

How the public API looks and behaves, so every new feature (calculated columns, KPIs, perspectives, the report layer) feels like the existing ones. Read this before adding or changing anything public. `docs/API-USABILITY-REVIEW.md` explains where these rules came from.

The user we design for: a .NET developer who knows Excel and Power Pivot but not VertiPaq. They think in Power Pivot's terms (tables, columns, "Sort by Column", "Mark as Date Table"). They discover the API through IntelliSense on the object in front of them, and they find out about mistakes from exception messages, not from reading our source.

## 1. The object model

```
DataModelWriter                 options, Save, owns tables and relationships
├─ AddCalculationTrigger → ModelTable (IsCalculationTrigger)
├─ Tables: ModelCollection<ModelTable>
│   └─ ModelTable                rows, owns columns, measures, hierarchies
│       ├─ Columns: ModelCollection<ModelColumn>   (table["Qty"])
│       │   └─ ModelColumn       IsHidden, SortByColumn
│       ├─ CalculatedColumns: ModelCalculatedColumn   formula only
│       ├─ Measures: ModelMeasure       IsHidden
│       └─ Hierarchies: ModelHierarchy  Levels
└─ Relationships: ModelRelationship     ManyColumn, OneColumn
```

1. **An operation lives on the object it changes.** If it concerns one table, it is a method of `ModelTable` (`sales.AddMeasure`, `products.AddHierarchy`, `dates.MarkAsDateTable`). If it spans tables or the whole model, it is a method of `DataModelWriter` (`AddRelationship`, `AddDateTable`, `Save`). A new model object gets a public `ModelXxx` class and goes into this tree.
2. **There is one way to do each thing.** No string-table overloads beside handle overloads, and no bulk helpers beside a property. The one sanctioned pair is `ColumnDefinition` (a blueprint, written before the table exists) and the matching `ModelColumn` property (for after it exists).
3. **Every `Add…` returns the object it added.** `CreateTable` and `AddTable` return a `ModelTable`, `AddRelationship` a `ModelRelationship`, `AddMeasure` a `ModelMeasure`, `AddHierarchy` a `ModelHierarchy`. Never return `void`, a count or a bool.
4. **Collections are `ModelCollection<T>`**, in the order added, indexable by position and by name, with `Contains(name)`. Lists that are never looked up by name (`Measures`, `Hierarchies`, `Relationships`) may be `IReadOnlyList<T>`. Do not add `XxxCount` or `XxxNames` properties; the collection already provides them.
5. **Settings that can change after creation are settable properties** that validate in the setter (`column.IsHidden`, `column.SortByColumn`, `measure.IsHidden`). Settings fixed at creation are constructor or method parameters, or `init` properties of an options or definition object.
6. **Calls can come in any order.** The user may add a measure before the rows, or mark a date table before relating it. Whatever depends on other calls or on data is checked in `Save`, not in the call.

## 2. Naming

1. **Verbs match Power Pivot**, and each `<summary>` says "like Power Pivot's …":
   - `Add…` creates a model object: `AddTable`, `AddRelationship`, `AddMeasure`, `AddHierarchy`, `AddDateTable`.
   - `Create…` is only for an empty object the caller fills: `CreateTable`.
   - `Mark…` and the other Power Pivot command verbs change something that already exists: `MarkAsDateTable`.
   - `Save…` writes output.
2. **Public model types are named `ModelXxx`**, to avoid clashes with `System.Data.DataTable` and with spreadsheet libraries' `Table` and `Column`. Options are `XxxOptions`, and blueprints are `XxxDefinition`.
3. **Booleans are `IsXxx` or say what `true` does:** `IsHidden`, `IsDateTable`, `HideSortColumns`, `IncludeYearMonth`. Avoid negatives such as `NoYearMonth`.
4. **Arguments go in the model's order:** name, then what defines the object, then options. For example `AddMeasure(name, expression, format)` and `AddHierarchy(name, levels)`. When two sides are involved, the many side comes first: `AddRelationship(manyColumn, oneColumn)`.
5. `ToString()` gives the DAX or Power Pivot notation: `'Sales'[Qty]`, `'Sales'[ProductKey] -> 'Products'[ProductKey]`, `Calendar (Year > Quarter > Month)`, `Sales (1,000 rows)`.

## 3. References: handles and names

1. **Anything that refers across objects takes a handle:** `AddRelationship(sales["ProductKey"], products["ProductKey"])`, `column.SortByColumn = products["CategoryOrder"]`.
2. **A table method names its own columns by string**, because the table is already fixed: `products.AddHierarchy("H", "Category", "Product")`, `dates.MarkAsDateTable("Date")`. Resolve those strings with `Builder.FindColumn(name, nameof(param))`.
3. **Names are compared ignoring case.** A table or column is found by the name the user gave and by the sanitised model name, so `"Sales (USD)"` keeps working after it is stored as `"Sales USD"`. Never compare names by hand.
4. **A handle must belong to this model and, where the rule requires it, to the same table.** Check with `CheckOwned` (or `value.Table != Table`) and say which model or table it belongs to.
5. **A table that must not change after it is made** (a generated one) sets `TableBuilder.ReadOnlyReason`. The message says what to do instead.

## 4. Errors

Pick the exception by **who can fix it and when**:

| Situation | Exception | Thrown by |
|---|---|---|
| A null argument | `ArgumentNullException` with the caller's parameter name | the call |
| An indexer with an unknown name (`Tables["X"]`, `table["X"]`) | `KeyNotFoundException` listing the names that exist | the indexer |
| A wrong name, type, count or value passed to a method or setter (unknown hierarchy level, duplicate measure, `"abc"` in a Whole Number column) | `ArgumentException` with the caller's parameter name (the property name for a setter) | the call |
| An out-of-range setting | `ArgumentOutOfRangeException` or `ArgumentException` | the `init` accessor |
| A call that conflicts with the model's state (second date table, rows added to a generated table) | `InvalidOperationException` | the call |
| A rule that depends on data or on other calls (duplicate key, a month with two sort numbers, a date table without a relationship) | `InvalidOperationException` | `Save` / `SaveModelPart` |

Rules for every message:

1. **Name the object in Power Pivot notation** (`'Sales'[ProductKey]`, table `'Dates'`, row 12) and give the offending value when there is one.
2. **Say how to fix it.** List the names that exist (`Identifiers.List(...)`) or give the call that fixes it ("AddRelationship takes the many side first…").
3. **A failed call changes nothing.** Check everything first, then apply the change. For example `AddRow` converts every value before any column takes one. Where that is impossible (a reader fails mid-row), call `TableBuilder.Fault(...)` so that `Save` fails with a clear message rather than writing a broken model.
4. **A failed save leaves no damage.** `Save(path)` writes beside the target and moves the file into place. Nothing is written to a caller's stream before the model is valid.

## 5. Options and defaults

1. **Positional parameters for what is required, options for what is not.** A feature with more than two optional settings gets a `sealed class XxxOptions` with `init` properties, passed as a last `XxxOptions? options = null`. One or two optional settings can be optional parameters (`AddMeasure(..., format)`). Never use a `bool` parameter for a setting that is also a property; use the property (`measure.IsHidden = true`).
2. **Defaults reproduce what Excel does**, or the usual practice when Excel has no default. Each default is documented as "Default X" in the summary.
3. **Use .NET types for concepts users already know:** `CultureInfo` (not an LCID), `DateOnly` for whole days, `IDataReader` and `IEnumerable<T>` for rows.
4. **Validate options when they are set** (in the `init` accessor), so that a bad value fails where it is written.
5. **Never add a hidden switch.** No environment variables or static toggles.

## 6. Values and types

1. Accept the obvious CLR types for each `ColumnDataType` (the README's Data types table) and convert others with the invariant culture. Each column builder implements `Coerce` (it converts or throws) and `AppendObject` (it never throws for a value `Coerce` returned).
2. `null` and `DBNull` mean BLANK everywhere.

## 7. State and lifetime

1. `DataModelWriter` is `IDisposable` (spill files). Every public method and setter that changes the model, on any model object, starts with `ThrowIfDisposed()` (`Table.Model.ThrowIfDisposed()` from a child object).
2. Public objects are thin handles over the internal builders (`ModelColumn.Builder`, `ModelTable.Builder`). They hold no state of their own that the serialiser would miss. When a setting changes, update the builder (for a column, its `Definition` through `ColumnDefinition.With(...)`).
3. `Save` can be called more than once and does not change the model.
4. Not thread-safe; one writer per thread.

## 8. Documentation is part of the API

1. Every public member has a `<summary>`. Methods and setters that throw for data or state have `<exception>` tags. Members with Excel-specific limits have a `<remarks>` saying what was verified in Excel.
2. When you add a public member, update in the same commit: the README section and example, `docs/SPEC.md` §2, and a test that covers the normal call and at least one rejected call with its message (`Throws<T>(..., messagePart)`).
3. Examples use realistic names (`sales`, `products`, `dates`, ProductKey, Qty) and the shortest correct code, with variables holding the handles: `sales.AddMeasure(...)`, not `model.Tables["Sales"].AddMeasure(...)`.

## 9. Changing existing API

No stable version has been released yet (only `0.1.0-preview.N`). Until 1.0, breaking changes that make the API more consistent are welcome. Make each one in a single commit, update every caller (tests, `tools/datagen`, `samples/`), and list it in the release notes. From 1.0 on, add overloads or options instead of changing signatures.

## 10. Checklist for a new feature

- [ ] It lives on the object it changes (§1.1) and returns what it adds (§1.3).
- [ ] The name uses a §2.1 verb and the Power Pivot term, and the summary says "like Power Pivot's …".
- [ ] Cross-object references are handles; same-table columns are names (§3).
- [ ] Arguments are checked before any state changes, and messages name the object, the value and the fix.
- [ ] Rules that depend on data are checked in `Save` and throw `InvalidOperationException`.
- [ ] Optional settings use an options class, optional parameters or properties, with defaults that match Excel and are documented.
- [ ] README, SPEC §2 and the tests are updated, and every test passes.
- [ ] Models that don't use the feature produce the same files as before (golden tests; the CLAUDE.md rules on profiles and Ordinals).
