# API usability review (2026-10-21)

This review looks at the public API of `0.1.0-preview.2` as a user meets it: someone writing a reporting job who knows Power Pivot but not VertiPaq. Each finding was reproduced with a small program against the library before it was fixed. The fixes are on branch `ccr-2e29a1a8-z10ndh`. The file format is unchanged: every golden test still passes and no serialiser code was touched. The conventions that came out of this review are in [`API-GUIDELINES.md`](API-GUIDELINES.md).

## Summary

| # | Finding | Severity | Status |
|---|---|---|---|
| 1 | A failed `Save(path)` truncated an existing file to 0 bytes | High: data loss | Fixed |
| 2 | One bad value in `AddRow` corrupted the table; the error appeared later, at `Save`, with no row or column | High | Fixed |
| 3 | A table created as `"Sales (USD)"` could not be referred to as `"Sales (USD)"` | Medium | Fixed |
| 4 | `AddTable` returned a row count, so tables loaded from a reader had no handle | Medium | Fixed (breaking) |
| 5 | Columns of tables loaded with `AddTable` could not be hidden | Medium | Fixed: `HideColumns` |
| 6 | "Not found" errors gave no hint, and reported the wrong parameter name | Medium | Fixed |
| 7 | A relationship given the wrong way round failed with a misleading message | Low | Fixed: hint in message |
| 8 | Rows could be added to a generated date table; it then failed at `Save` with "row 366 has no date" | Low | Fixed: read-only |
| 9 | A `TableWriter` from another `DataModelWriter` was accepted silently | Low | Fixed |
| 10 | `TableWriter.Columns` did not show later changes (`SetSortByColumn`, hiding) | Low | Fixed |
| 11 | Smaller gaps: no null checks on `columns`, `values` or `output`; `SaveModelPart` with no tables | Low | Fixed |

Tests: 42 → 46 (`Api*` tests in `tests/PivotWright.Tests/Program.cs`), all passing.

## 1. A failed save no longer destroys the previous file

`Save(path)` used to open the target with `FileMode.Create` (which truncates it) *before* the model was validated. A duplicate key therefore left yesterday's report as an empty 0-byte file.

```csharp
File.Exists("report.xlsx");          // true, yesterday's good report
model.Save("report.xlsx");           // throws: duplicate primary key

// Before: report.xlsx is now 0 bytes.
// After:  report.xlsx is unchanged. The new workbook is written to a temporary file in the same
//         directory and moved over the old one only when it is complete. No temporary file is left behind.
```

## 2. Bad values are reported where they happen, and the table stays usable

Values were appended column by column. If the second value of a row failed to convert, the first column already had one value more than the others. The user saw a bare `FormatException` and, if they caught it and went on, `Save` later failed with a message about internals.

```csharp
var t = model.CreateTable("T", new[] { new ColumnDefinition("A", ColumnDataType.Int64),
                                       new ColumnDefinition("B", ColumnDataType.Int64) });
try { t.AddRow(1, "abc"); } catch (ArgumentException e) { log.Warn(e.Message); }
t.AddRow(2, 3);
model.Save("t.xlsx");
```

| | Before | After |
|---|---|---|
| `AddRow(1, "abc")` | `FormatException: The input string 'abc' was not in a correct format.` | `ArgumentException: Row 1 of table 'T': column 'B' (Int64) cannot hold the String value 'abc'. The input string 'abc' was not in a correct format.` |
| `Save` | `InvalidOperationException: Column 'A' has 2 values but table 'T' has 1 rows.` | saves, with the one good row |

`AddRow` now converts the whole row before any column takes a value (`ColumnBuilder.Coerce`). A wrong number of values is reported the same way: `Row 1 of table 'T' has 1 values, but the table has 2 columns ('A', 'B').`

When an `IDataReader` fails in the middle of a row, part of the row is already stored and cannot be taken back. That case is now recorded instead of being left to surface as a confusing error: the exception is passed on, and `Save` refuses the table with `Table 'T' is incomplete: loading row 2 from the data reader failed at column 'B' (...)`.

The cost is about 9% on a microbenchmark of `AddRow` alone (2 million rows of 4 columns: 1.97 s before, 2.16 s after, Release build). Values that already have the column's type are not copied. The typed `IDataReader` fast path is unchanged.

## 3. Tables are found by the name you gave them

Power Pivot rejects some characters in names, so the library removes them: `"Sales (USD)"` becomes `"Sales USD"`. Lookups only knew the stored name, so a user's own string stopped working right after it was used:

```csharp
model.CreateTable("Sales (USD)", columns);
model.AddMeasure("Sales (USD)", "Total", "SUM('Sales USD'[Qty])");
// Before: ArgumentException: There is no table named 'Sales (USD)'.
// After:  works. Either "Sales (USD)" or "Sales USD" finds the table.
```

Columns already accepted both forms. Tables now do too.

## 4. `AddTable` returns the table (breaking change)

`CreateTable` and `AddDateTable` returned a `TableWriter`, but the three `AddTable` overloads returned `long` (the row count). So the most common way to load data, from a database reader, was the one that gave you no handle for the `TableWriter` overloads:

```csharp
// Before
long n = model.AddTable("Sales", reader);
model.AddTable(productsDataTable);
model.AddRelationship("Sales", "ProductKey", productsDataTable.TableName, "ProductKey");

// After
var sales = model.AddTable("Sales", reader);
var products = model.AddTable(productsDataTable);
model.AddRelationship(sales, "ProductKey", products, "ProductKey");
long n = sales.RowCount;
```

This change is source-breaking for code that stored the result in a `long`; the compiler reports it, and `.RowCount` is the fix. It was made before 1.0 because it is the only way to make "every table-making method returns the table" true (guideline §2.1). `TableWriter.ToString()` now gives `Sales (1,000 rows)`.

## 5. `HideColumns` for any table

Hidden columns could only be set with `ColumnDefinition.IsHidden`, which you can't use with `AddTable(reader)`, `AddTable(DataTable)` or `AddTable<T>`. Keys loaded from a database could therefore never be hidden, even though hiding keys is standard practice.

```csharp
var sales = model.AddTable("Sales", reader);
model.HideColumns(sales, "ProductKey", "DateKey");    // or HideColumns("Sales", ...)
```

It writes the same model as `IsHidden = true`: the test `ApiHideColumns` compares the hidden attributes and the file list of both, so the Excel verification of hidden columns carries over. It is all or nothing: if one of the names is wrong, no column is hidden. `TableWriter.Columns` then reports `IsHidden = true` for those columns.

## 6. "Not found" errors say what exists

```text
AddMeasure("Sale", ...)
  Before: There is no table named 'Sale'. (Parameter 'name')
  After:  There is no table named 'Sale'. The model's tables are: 'Products', 'Sales'. (Parameter 'table')

AddHierarchy("Sales", "H", "Qtty", "Qty")
  Before: Table 'Sales' has no column named 'Qtty'. (Parameter 'name')
  After:  Table 'Sales' has no column named 'Qtty'. Its columns are: 'Qty'. (Parameter 'levels')
```

The parameter name used to be the internal helper's (`name`), which pointed at the wrong argument of `AddRelationship(manyTable, manyColumn, oneTable, oneColumn)`. It is now the public method's own parameter name. The lists are capped at 20 names.

## 7. Wrong-way-round relationships get a hint

Four string arguments are easy to swap. Swapping them used to produce a message that was accurate but did not tell the user what they had done wrong:

```text
AddRelationship("Products", "K", "Sales", "K")   // swapped
  ... needs unique primary keys, but 'Sales[K]' contains the value '1' more than once.
  AddRelationship takes the many side (for example Sales) first and the one side, whose key is unique (for example Products), second.
```

A stronger fix would be a typed column reference (`sales["ProductKey"]`) or named arguments in every example. That is left for a later design pass (see below).

## 8–11. Smaller fixes

```csharp
var dates = model.AddDateTable("Dates", new DateOnly(2025, 1, 1), new DateOnly(2025, 12, 31));
dates.AddRow(...);
// Before: accepted; Save later failed with "row 366 has no date".
// After:  InvalidOperationException: Table 'Dates' was generated by AddDateTable and is complete;
//         widen its range with the from and to arguments instead of adding rows.

model.HideColumns(tableFromOtherModel, "K");
// Before: silently acted on whichever table of this model had the same name, or failed with "no table named".
// After:  ArgumentException: Table 'Products' belongs to another DataModelWriter.
```

- `TableWriter.Columns` is computed from the live column state, so `SetSortByColumn` and `HideColumns` show in it.
- Null `columns`, `values`, `items`, `project` and `output` arguments throw `ArgumentNullException` at the call.
- `SaveModelPart` refuses an empty model, as `Save` already did.

## Not changed, and why

- **`AddMeasure(..., bool hidden)`** differs from `ColumnDefinition.IsHidden`. It is an optional parameter that is clear when used as a named argument (`hidden: true`). Renaming it would break callers for little gain. Guideline §4.1 allows one or two optional parameters.
- **Validation at `Save` for date-table structure** (needs a hierarchy and a relationship). This keeps calls order-independent (guideline §1.4), and the messages already name the call that fixes the problem.
- **No async or `CancellationToken`.** Loading is CPU- and stream-bound, and `IDataReader` is synchronous. Worth adding (`SaveAsync`, `AddTableAsync(DbDataReader)`) when the report layer needs it.
- **Typed column references** (`sales["Qty"]`) instead of strings. This would remove swapped-argument mistakes, but it doubles the overload count. It should be designed together with calculated columns, which need column references anyway.
- **`Lcid` instead of a `CultureInfo`.** Matches the file format and Excel's settings, and stays as it is. A `Culture` convenience option could be added later without a break.
- **Verbose `new ColumnDefinition(name, type)`.** Readable and explicit. Factory shortcuts would add surface without adding capability.

## Excel testing

No Excel test is needed for this change set. The output of every existing call is unchanged (the golden tests pass and serialisation was not touched), and `HideColumns` produces the same files as the Excel-verified `IsHidden`.
