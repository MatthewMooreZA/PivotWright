# PivotWright API guidelines

How the public API looks and behaves, so every new feature (calculated columns, KPIs, perspectives, the report layer) feels like the existing ones. Read this before adding or changing anything public. `docs/API-USABILITY-REVIEW.md` explains where these rules came from.

The user we design for: a .NET developer who knows Excel and Power Pivot but not VertiPaq. They think in Power Pivot's terms (tables, columns, "Sort by Column", "Mark as Date Table") and they find out about mistakes from exception messages, not from reading our source.

## 1. Shape

1. **One entry point.** `DataModelWriter` owns the model. Features are methods on it, not new top-level objects the user has to wire together. `TableWriter` is the only per-object handle, and it exists because tables take rows.
2. **Verbs match Power Pivot.** Name a method after the Power Pivot command it reproduces and say so in its `<summary>` ("like Power Pivot's *Mark as Date Table*"). Use these prefixes:
   - `Add…` creates a model object: `AddTable`, `AddRelationship`, `AddMeasure`, `AddHierarchy`, `AddDateTable`.
   - `Create…` only for an empty object the caller fills: `CreateTable`.
   - `Set…`, `Mark…`, `Hide…` change something that already exists: `SetSortByColumn`, `MarkAsDateTable`, `HideColumns`.
   - `Save…` writes output.
3. **Argument order is the model's order:** table, then the object's name, then what defines it, then options. For example `AddMeasure(table, name, expression, format)` and `AddHierarchy(table, name, levels)`. When two tables are involved, the many side comes first, as in `AddRelationship(manyTable, manyColumn, oneTable, oneColumn)`.
4. **Calls can come in any order.** The user may add a measure before the rows, or mark a date table before relating it. Whatever depends on other calls or on data is checked in `Save`, not in the call (see §3).

## 2. Tables are named or passed

1. Every method that makes a table returns its `TableWriter`: `CreateTable`, `AddTable` (all overloads) and `AddDateTable`. Never return a count or `void` from a table-making method. The row count is `TableWriter.RowCount`.
2. Every public method that takes a table has two overloads: `string table` and `TableWriter table`. The `TableWriter` overload is one line that calls `NameOf(table)` (which rejects a table from another `DataModelWriter`) and forwards to the string overload. All validation lives in the string overload.
3. Look up tables with `FindTable(name, nameof(param))` and columns with `TableBuilder.FindColumn(name, nameof(param))`. They ignore case and accept both the name the user gave and the sanitised model name, so a user who wrote `CreateTable("Sales (USD)", …)` can keep writing `"Sales (USD)"`. Never compare names with `==` or by hand.
4. A table that must not change after it is made (a generated table) sets `TableBuilder.ReadOnlyReason`. The message says what to do instead.

## 3. Errors

Pick the exception by **who can fix it and when**:

| Situation | Exception | Thrown by |
|---|---|---|
| A null argument | `ArgumentNullException` with the caller's parameter name | the call |
| A wrong name, type, count or value that the call can see (unknown table, duplicate measure, `"abc"` in a Whole Number column) | `ArgumentException` with the caller's parameter name | the call |
| An out-of-range setting | `ArgumentOutOfRangeException` | the options `init` accessor |
| A call that conflicts with the model's state (second date table, rows added to a generated table) | `InvalidOperationException` | the call |
| A rule that depends on data or on other calls (duplicate key, a month with two sort numbers, a date table without a relationship) | `InvalidOperationException` | `Save` / `SaveModelPart` |

Rules for every message:

1. **Name the object in Power Pivot's notation:** `'Sales[ProductKey]'`, table `'Dates'`, row 12. Give the offending value when there is one.
2. **Say how to fix it.** List the names that do exist (`Identifiers.List(...)`), or give the call that fixes it ("AddRelationship takes the many side first…").
3. **Pass the parameter name of the public method**, not of an internal helper. `FindTable`, `FindColumn` and `SetHidden` take a `paramName` for this reason.
4. **A failed call changes nothing.** Check everything first, then apply the change. For example `HideColumns` resolves every column before hiding any, and `AddRow` converts every value before any column takes one. Where that is impossible (a reader fails in the middle of a row), call `TableBuilder.Fault(...)` so that `Save` fails with a clear message rather than writing a broken model.
5. **A failed save leaves no damage.** `Save(path)` writes beside the target and moves the file into place. Nothing is written to a caller's stream before the model is valid.

## 4. Options and defaults

1. **Positional parameters for what is required, options for what is not.** A feature with more than two optional settings gets a `sealed class XxxOptions` with `init` properties, passed as a last `XxxOptions? options = null` (`DataModelOptions`, `DateTableOptions`). One or two optional settings can be optional parameters (`AddMeasure(..., format, hidden)`).
2. **Defaults reproduce what Excel does,** or the usual practice when Excel has no default. Each default is documented as "Default X" in the property's summary.
3. **Booleans say what `true` does:** `HideSortColumns`, `IncludeYearMonth`, `CompleteYears`, `IsHidden`. Avoid negatives such as `NoYearMonth`. Each optional feature can be switched off on its own.
4. **Validate options when they are set** (in the `init` accessor), so that a bad value fails where it is written, not at `Save`.
5. **Never add a hidden switch.** No environment variables or static toggles. An experiment either becomes a documented option or stays out of committed code.

## 5. Values and types

1. Accept the obvious CLR types for each `ColumnDataType` (the README's Data types table) and convert others with the invariant culture. Each column builder implements `Coerce` (it converts or throws) and `AppendObject` (it never throws for a value `Coerce` returned).
2. `null` and `DBNull` mean BLANK everywhere.
3. Prefer `DateOnly` for whole days in new APIs (`AddDateTable(name, DateOnly from, DateOnly to)`). Row values may be any date type.

## 6. State and lifetime

1. `DataModelWriter` is `IDisposable` (spill files). Every public method starts with `ThrowIfDisposed()`.
2. Read-only state is exposed as plain properties (`TableNames`, `MeasureCount`, `DateTableName`, `TableWriter.Columns`). Those properties reflect later changes: `TableWriter.Columns` shows a column hidden with `HideColumns`. If you add a setting that can change after creation, update `ColumnBuilder.Definition` through `ColumnDefinition.With(...)`.
3. `Save` can be called more than once and does not change the model.
4. Not thread-safe; one writer per thread. Say so in the docs if a feature could make users expect otherwise.

## 7. Documentation is part of the API

1. Every public member has a `<summary>`. Methods that throw for data or state have `<exception>` tags. Members with Excel-specific limits have a `<remarks>` saying what was verified in Excel.
2. When you add a public member, update in the same commit: the README section and example, `docs/SPEC.md` §2, and a test in `tests/PivotWright.Tests` that covers the normal call and at least one rejected call with its message (`Throws<T>(..., messagePart)`).
3. Examples use realistic names (Sales, Products, Dates, ProductKey, Qty) and the shortest correct code.

## 8. Changing existing API

The package is published (`0.1.0-preview.N` on nuget.org). Before 1.0, a source-breaking change is acceptable when it removes a real usability problem: note it under "Breaking changes" in the release notes and in the review document. Never change behaviour silently: either the compiler flags the change, or the old behaviour stays. From 1.0 on, add overloads or options instead of changing signatures.

## 9. Checklist for a new feature

- [ ] Method name uses a §1.2 verb and the Power Pivot term. The summary says "like Power Pivot's …".
- [ ] Has both a `string table` and a `TableWriter table` overload, and lookups go through `FindTable` and `FindColumn`.
- [ ] Argument checks run before any state changes, and messages name the object, the value and the fix.
- [ ] Rules that depend on data are checked in `Save` with `InvalidOperationException`.
- [ ] Optional settings use options or optional parameters, with defaults that match Excel, documented.
- [ ] README, SPEC §2 and tests are updated, and every test passes.
- [ ] The file output is unchanged for models that don't use the feature (golden tests; CLAUDE.md rules on profiles and Ordinals).
