# API usability review (2026-10-21)

This review looks at the public API of `0.1.0-preview.2` as a user meets it: someone writing a reporting job who knows Power Pivot but not VertiPaq. It was done in two rounds:

1. **Fixes.** Every finding was reproduced with a small program before it was fixed: data loss on a failed save, a table corrupted by one bad value, unhelpful errors.
2. **Redesign.** No stable version has been released, so the API was reshaped around an object model: table-scoped operations on the table, typed column handles, settable properties. Round 1 had held this back to avoid breaking callers.

The file format is unchanged. The golden tests pass, no serialiser code was touched, and the `tools/datagen` models were diffed against the previous commit: only random ids differ. No Excel retest is needed. The conventions that came out of the review are in [`API-GUIDELINES.md`](API-GUIDELINES.md).

## The API before and after

```csharp
// Before (0.1.0-preview.2)
using var model = new DataModelWriter(new DataModelOptions { Lcid = 7177 });
long n = model.AddTable("Sales", salesReader);                 // returns a row count
model.AddTable(productsDataTable);
model.AddRelationship("Sales", "ProductKey", "Products", "ProductKey");   // four strings
model.SetSortByColumn("Products", "Category", "CategoryOrder");
model.AddHierarchy("Products", "Product Hierarchy", "Category", "Product");
model.AddMeasure("Sales", "Total Qty", "SUM(Sales[Qty])", MeasureFormat.WholeNumber);
model.AddMeasure("Sales", "Rate", "0.15", hidden: true);
// no way to hide Sales[ProductKey]: it was loaded from a reader
model.Save("report.xlsx");

// After
using var model = new DataModelWriter(new DataModelOptions { Culture = CultureInfo.GetCultureInfo("en-ZA") });
var sales = model.AddTable("Sales", salesReader);              // returns the table
var products = model.AddTable(productsDataTable);
model.AddRelationship(sales["ProductKey"], products["ProductKey"]);
products["Category"].SortByColumn = products["CategoryOrder"];
products.AddHierarchy("Product Hierarchy", "Category", "Product");
sales.AddMeasure("Total Qty", "SUM(Sales[Qty])", MeasureFormat.WholeNumber);
sales.AddMeasure("Rate", "0.15").IsHidden = true;
sales["ProductKey"].IsHidden = true;
model.Save("report.xlsx");                                     // a failed save leaves the old file intact
```

## Summary

| # | Finding | Severity | Resolution |
|---|---|---|---|
| 1 | A failed `Save(path)` truncated an existing file to 0 bytes | High: data loss | Write beside the target, then move it into place |
| 2 | One bad value in `AddRow` corrupted the table; the error surfaced at `Save` with no row or column | High | Convert the whole row first; the message names the row, column and value |
| 3 | Operations on a table were model methods taking the table name as a string, with duplicate overloads for handles | Medium | Moved onto `ModelTable` |
| 4 | `AddTable` returned a row count, so tables loaded from a reader had no handle | Medium | Every `Add…` returns what it added |
| 5 | `AddRelationship` took four strings, which are easy to mix up | Medium | `AddRelationship(ModelColumn many, ModelColumn one)` |
| 6 | Columns of tables loaded with `AddTable` could not be hidden or re-sorted | Medium | Settable `ModelColumn.IsHidden` and `SortByColumn` |
| 7 | A measure was hidden with a `bool` parameter, a column with a property | Low | `ModelMeasure.IsHidden` |
| 8 | A table created as `"Sales (USD)"` could not be found by that name | Medium | Lookups accept the given and the stored name |
| 9 | "Not found" errors gave no hint and named the wrong parameter | Medium | They list the names that exist |
| 10 | `Lcid = 7177` instead of a culture | Low | `DataModelOptions.Culture` |
| 11 | State was exposed as counts and name lists (`TableNames`, `MeasureCount`, `HierarchyCount`, `RelationshipCount`, `DateTableName`) | Low | Collections of objects |
| 12 | A relationship given the wrong way round failed with a misleading message | Low | The message says which side goes first |
| 13 | Rows could be added to a generated date table, which then failed at `Save` | Low | The table is read-only |
| 14 | A handle from another `DataModelWriter` was accepted silently | Low | Rejected, naming its model |
| 15 | Missing null checks; `SaveModelPart` accepted an empty model | Low | Fixed |

Tests: 42 → 46, all passing. The new `Api*` tests are in `tests/PivotWright.Tests/Program.cs`. The suite, `tools/datagen` and `samples/PivotWright.Demo` were ported to the new API.

## Round 1: fixes

### A failed save no longer destroys the previous file

`Save(path)` used to open the target with `FileMode.Create`, which truncates it, *before* the model was validated. A duplicate key therefore left yesterday's report as an empty file.

```csharp
model.Save("report.xlsx");           // throws: duplicate primary key
// Before: report.xlsx is now 0 bytes.
// After:  report.xlsx is unchanged. The new workbook goes to a temporary file in the same directory and is moved
//         over the old one only when complete. No temporary file is left behind.
```

### Bad values are reported where they happen, and the table stays usable

Values used to be appended column by column. If the second value failed to convert, the first column already had one value more than the others.

```csharp
try { t.AddRow(1, "abc"); } catch (ArgumentException e) { log.Warn(e.Message); }
t.AddRow(2, 3);
model.Save("t.xlsx");
```

| | Before | After |
|---|---|---|
| `AddRow(1, "abc")` | `FormatException: The input string 'abc' was not in a correct format.` | `ArgumentException: Row 1 of table 'T': column 'B' (Int64) cannot hold the String value 'abc'. ...` |
| `Save` | `InvalidOperationException: Column 'A' has 2 values but table 'T' has 1 rows.` | Saves, with the one good row |

`AddRow` now converts the whole row before any column takes a value (`ColumnBuilder.Coerce`). When an `IDataReader` fails in the middle of a row, part of the row is already stored. The table is then marked, and `Save` reports `Table 'T' is incomplete: loading row 2 from the data reader failed at column 'B' (...)`.

The cost is about 9% on a microbenchmark of `AddRow` alone: 2 million rows of 4 columns took 1.97 s before and 2.16 s after (Release build). The typed `IDataReader` path is unchanged.

### Errors say what exists

```text
model.Tables["Sale"]
  KeyNotFoundException: There is no table named 'Sale'. The model's tables are: 'Products', 'Sales'.
sales["Qtty"]
  KeyNotFoundException: Table 'Sales' has no column named 'Qtty'. Its columns are: 'ProductKey', 'Qty'.
products.AddHierarchy("H", "Categry", "Product")
  ArgumentException: Table 'Products' has no column named 'Categry'. Its columns are: ... (Parameter 'levels')
AddRelationship(products["K"], sales["K"])    // swapped; at Save
  ... needs unique primary keys, but 'Sales[K]' contains the value '1' more than once.
  AddRelationship takes the many side (for example Sales) first and the one side ... second.
```

Before, these messages had no list of names, and the parameter named was the internal helper's (`name`).

## Round 2: redesign

### Operations live on the object they change

Measures, hierarchies and the date-table mark belong to one table, so they are now methods of `ModelTable`. IntelliSense on `sales.` shows what you can do with Sales. The model keeps what spans tables: relationships, the generated date table, saving. This replaced 10 overloads on `DataModelWriter` (a string and a `TableWriter` version each of `AddMeasure`, `AddHierarchy`, `MarkAsDateTable`, `SetSortByColumn` and `AddRelationship`) and removed `HideColumns`, all without losing any capability.

| Before | After |
|---|---|
| `model.AddMeasure("Sales", name, dax, format, hidden)` | `sales.AddMeasure(name, dax, format)` → `ModelMeasure` (`.IsHidden`) |
| `model.AddHierarchy("Products", name, levels)` | `products.AddHierarchy(name, levels)` → `ModelHierarchy` |
| `model.MarkAsDateTable("Dates", "Date")` | `dates.MarkAsDateTable("Date")` |
| `model.SetSortByColumn("Products", "Category", "CategoryOrder")` | `products["Category"].SortByColumn = products["CategoryOrder"]` (null clears) |
| `model.HideColumns(sales, "ProductKey")` | `sales["ProductKey"].IsHidden = true` |
| `model.AddRelationship("Sales", "ProductKey", "Products", "ProductKey")` | `model.AddRelationship(sales["ProductKey"], products["ProductKey"])` → `ModelRelationship` |
| `TableWriter` | `ModelTable` |
| `model.TableNames`, `model.DateTableName` | `model.Tables` (by position or name), `model.DateTable` |
| `RelationshipCount`, `MeasureCount`, `HierarchyCount` | `model.Relationships`, `table.Measures`, `table.Hierarchies` |
| `DataModelOptions.Lcid = 7177` | `DataModelOptions.Culture = CultureInfo.GetCultureInfo("en-ZA")` |

### Typed column handles

`sales["ProductKey"]` is a `ModelColumn`. A relationship now takes two columns, each of which already knows its table. So a table and a column from different tables can no longer be paired, a column from another model is rejected, and `ToString()` gives the DAX form `'Sales'[ProductKey]`. Swapping the two arguments is still possible, but now reads visibly wrong (`AddRelationship(products[...], sales[...])`), and `Save` explains the mistake.

### Properties for settings that change

Hidden state and Sort by Column can change after a table is created, so they are settable properties of `ModelColumn`, validated in the setter. They start from `ColumnDefinition.IsHidden` and `SortByColumn` when the table was created from definitions. A measure's hidden state is a property too, so hiding works the same way everywhere:

```csharp
sales["ProductKey"].IsHidden = true;
sales.AddMeasure("Base Rate", "0.15").IsHidden = true;
dates["Date"].SortByColumn = dates["Day"];   // ArgumentException: the date column of a date table cannot take part in Sort by Column
```

### A culture, not a number

`Culture = CultureInfo.GetCultureInfo("en-ZA")` says what it means. The LCID that Excel stores is derived from it. Cultures without an LCID (the invariant culture, custom cultures) are rejected when the option is set.

### Migration from 0.1.0-preview.2

Every change is a compile error at the old call site, so nothing changes behaviour silently. The table above is the mapping. The two changes that need more than renaming a call:

- `long n = model.AddTable(...)` becomes `var t = model.AddTable(...)`, then `t.RowCount`.
- Unknown names in `model.Tables["X"]` and `table["X"]` now throw `KeyNotFoundException` (the .NET convention for indexers) instead of `ArgumentException`.

## Not changed, and why

| Item | Why it stays |
|---|---|
| Date-table structure (a hierarchy, a relationship) checked at `Save` | Keeps calls order-independent; the messages name the call that fixes it |
| No async or `CancellationToken` | `IDataReader` is synchronous and the work is CPU-bound; an async wrapper would only hide a blocking call. Add `SaveAsync` when the report layer needs a real async path |
| `ColumnDefinition` stays a separate blueprint type | `CreateTable` needs column types before a table exists; `ModelColumn` is the live object |
| Verbose `new ColumnDefinition(name, type)` | Explicit and readable; factory shortcuts would add surface without capability |
| Relationship direction is not detected from the data | Rows may arrive after the call; Power Pivot detects the direction because its data is already loaded |
