# PivotWright

A .NET 8 library that writes Excel workbooks (`.xlsx`) containing a **Power Pivot data model**. It writes the model directly in the xVelocity (VertiPaq) storage format that Excel's in-memory engine loads. It needs no Excel, Analysis Services, Power Query or COM interop at write time, so it runs on any OS.

You stream rows in from an `IDataReader`, a `DataTable`, objects or plain `object[]` rows. The library encodes each column the way the engine stores it: hash or value encoding, bit-packing plus RLE, dictionaries and attribute hierarchies. It then packages the model into `xl/model/item.data`. Open the file in Excel and the tables are already loaded: insert a PivotTable from the data model or open *Data → Manage Data Model*.

The format follows Microsoft's open specification [MS-XLDM] (Spreadsheet Data Model File Format). Every structure was checked against real Excel-generated workbooks (see [Verification](#verification)).

## Quick start

```csharp
using System.Globalization;
using PivotWright;

using var model = new DataModelWriter();   // en-US and Excel's defaults; or set options:
// using var model = new DataModelWriter(new DataModelOptions
// {
//     Culture = CultureInfo.GetCultureInfo("en-GB"),   // collation, currency format, month names
//     DecimalAsCurrency = true,                        // decimal -> Fixed Decimal (Currency) instead of Double
// });

// 1. Straight from a database: rows are streamed, never materialised.
ModelTable sales;
using (var cmd = new SqlCommand("SELECT * FROM dbo.FactSales", connection))
using (var reader = cmd.ExecuteReader())
    sales = model.AddTable("Sales", reader);

// 2. From a DataTable.
var products = model.AddTable(productsDataTable);

// 3. From objects (public properties become columns).
var regions = model.AddTable("Regions", regionList);

// 4. Row by row, with explicit column types.
var fx = model.CreateTable("FX", new[]
{
    new ColumnDefinition("Date", ColumnDataType.DateTime),
    new ColumnDefinition("Currency", ColumnDataType.String),
    new ColumnDefinition("Rate", ColumnDataType.Double),
});
fx.AddRow(new DateTime(2026, 10, 1), "ZAR", 17.42);
fx.AddRow(new DateTime(2026, 10, 1), "USD", 1.0);

model.AddRelationship(sales["ProductKey"], products["ProductKey"]);
sales["ProductKey"].IsHidden = true;
sales.AddMeasure("Total Qty", "SUM(Sales[Qty])", MeasureFormat.WholeNumber);
model.Save("Sales model.xlsx");
```

### How the API is organised

- `DataModelWriter` owns the **tables** and the **relationships**. A `ModelTable` owns its **columns**, **measures** and **hierarchies**, and takes rows.
- Every `Add…` method returns the object it added (`ModelTable`, `ModelRelationship`, `ModelMeasure`, `ModelHierarchy`).
- Find objects by name with `model.Tables["Sales"]` and `sales["Qty"]` (short for `sales.Columns["Qty"]`). A table can be found by the name you gave it or by its stored name, ignoring case. A wrong name throws `KeyNotFoundException`, which lists the names that do exist.
- Settings that can change after an object is added are properties: `column.IsHidden`, `column.SortByColumn`, `measure.IsHidden`.
- Other mistakes in a call throw `ArgumentException` at once. Problems that depend on the data (a duplicate key, a month with two sort numbers) are reported by `Save`, as `InvalidOperationException`. A failed `Save(path)` leaves an existing file at that path untouched. The conventions behind the API are written down in [`docs/API-GUIDELINES.md`](docs/API-GUIDELINES.md).

### Measures

```csharp
sales.AddMeasure("Total Qty", "SUM(Sales[Qty])", MeasureFormat.WholeNumber);
sales.AddMeasure("Revenue", "SUMX(Sales, Sales[Qty] * Sales[Price])", MeasureFormat.Currency());
sales.AddMeasure("Base Rate", "0.15").IsHidden = true;   // used only by other measures
```

A measure is a DAX expression attached to a table, shown in the PivotTable field list under that table. Names must be unique across the model and must not contain `[` or `]`. The expression is stored as written and evaluated by Excel; the library does not parse DAX, so a mistake appears as an error in the PivotTable. Formats: `General`, `WholeNumber`, `Decimal(n)`, `Currency(n)`, `Percentage(n)`, `Date`. General, WholeNumber, Currency and Date match files written by Excel; Decimal and Percentage were confirmed to display correctly.

### Calculated columns

```csharp
sales.AddCalculatedColumn("Line Total", "[Qty] * [Amount]", ColumnDataType.Currency);
```

The file holds only the formula; Excel calculates the values the first time a PivotTable uses the column, so you do not supply them. The library does not parse the DAX, so give the type the formula produces (`Currency`, `Int64`, `Double`, `DateTime`, `Boolean` or `String`), as Excel infers it when the column is typed in.

> **Limitation:** until Excel has calculated the column, it can be used as a value in a PivotTable but **not as a row, column or filter field**. Put it in *Values* once, or click *Calculate Now* in the Power Pivot window, and it works everywhere. *Refresh All* does not calculate it. Excel's own files behave the same way.

Written the way Excel saves a calculated column it has not calculated yet (`reference/excel-samples/H10_calculated_column_manual.xlsx`, and H11 to H15 for the other types); only some shapes have been seen, so saving requires a model with a hierarchy, no date table, no hidden columns, a table of one segment, and one column on the many-side table, two columns on the many-side table, one column on a hierarchy table that is the primary side of a relationship, or one column on each of those two tables (H10, H18, H19 and H20 are the samples; all four shapes are verified in Excel as generated files). Other combinations make `Save` throw a message that lists these shapes. Verified in Excel for all six types: the model loads, the PivotTable shows the values, and adding a hierarchy or a Sort by Column, saving and refreshing afterwards all work.

### Hierarchies

```csharp
products.AddHierarchy("Product Hierarchy", "Category", "Subcategory", "Product");
dates.AddHierarchy("Calendar", "Year", "Month");
```

A hierarchy is an ordered drill-down path over two or more columns of one table; PivotTables offer it as a single field with expandable levels. The name must be unique within the table and different from its column names. The tree is built and sorted when saving (levels sort like the columns do, BLANK first), so memory use grows with the number of distinct value paths. The file layout matches what Excel writes byte for byte for the same data. BLANK values inside a level (sorted first) and values that differ only in case (merged) were confirmed in Excel.

### Date table

```csharp
model.AddRelationship(sales["Date"], dates["Date"]);
dates.AddHierarchy("Calendar", "Year", "Month");
dates.MarkAsDateTable("Date");
```

Marks a table as the model's date table, like Power Pivot's *Mark as Date Table*, which time intelligence functions such as `TOTALYTD` and `SAMEPERIODLASTYEAR` rely on. The column must be a date column (`ColumnDataType.DateTime`, which is what `DateOnly`, `DateTime` and `DateTimeOffset` sources all map to) with one whole day per row (no BLANK, no time part, no duplicates; checked when saving). The file layout follows what Excel writes after marking a date table in a model like the one above; that is the only case seen so far, so for now saving requires **a hierarchy on the date table, a relationship that uses the date column as its primary key, and a single date table per model**. Verified in Excel: the model loads, the date table works in a PivotTable, and adding a hierarchy or a Sort by Column afterwards works.

#### Generated date table

```csharp
var dates = model.AddDateTable("Dates", new DateOnly(2020, 1, 1), new DateOnly(2030, 12, 31));   // also marks it as date table
// or, for a financial year ending in June:
// var dates = model.AddDateTable("Dates", new DateOnly(2020, 7, 1), new DateOnly(2030, 6, 30), new DateTableOptions { YearEndMonth = 6 });
model.AddRelationship(sales["Date"], dates["Date"]);
```

Alternatively `AddDateTable` builds a standard calendar for you, one row per day, with the columns Date, Year, QuarterNumber, Quarter, MonthNumber, Month, Day, DayOfWeekNumber (1 = Monday) and DayOfWeek. Quarter, Month and DayOfWeek are sorted by their number columns, a `Calendar` hierarchy (Year, Quarter, Month) is added, and the table is marked as date table. Settings go in a `DateTableOptions` object. Set `YearEndMonth` (1 to 12, default 12 = December) when the year does not end in December: the table then also gets FiscalYear (the calendar year in which the fiscal year ends, so with a June year end July 2024 belongs to FiscalYear 2025), FiscalQuarterNumber, FiscalQuarter, FiscalMonthNumber (1 = first month of the fiscal year) and FiscalMonth, sorted by their numbers, plus a `Fiscal Calendar` hierarchy (FiscalYear, FiscalQuarter, FiscalMonth). The year does not have to be a fiscal one: set `YearLabel = "Financial"` (letters and digits, default `"Fiscal"`) to get FinancialYear, FinancialQuarter, FinancialMonth and a `Financial Calendar` hierarchy instead. The calendar columns stay in the table either way, because Month and FiscalMonth need different sort orders and reports often need both.

Three more behaviours are on by default and each has its own flag in `DateTableOptions`: `CompleteYears` extends the range to whole years (calendar years, or years ending in `yearEndMonth`) as DAX's `CALENDARAUTO` does and as Microsoft's date table guidance asks, so `from` and `to` only need to be a day inside the first and last year; `IncludeYearMonth` adds `YearMonth` ("Mar 2025", sorted by `YearMonthNumber`) so a month can be told apart across years; `HideSortColumns` hides the number columns that exist only to sort other columns (`QuarterNumber`, `MonthNumber`, `YearMonthNumber`, `DayOfWeekNumber` and the fiscal ones) from the field list, as is usual practice. Set a flag to `false` to switch it off, for example `new DateTableOptions { HideSortColumns = false }`. Invalid values (a year end outside 1 to 12, a label with other than letters and digits) are rejected as soon as the options are created. Month and day names use the model culture. You still add the relationship to your fact table yourself, and `MarkAsDateTable` keeps working for tables you load yourself. The Excel-verified case is a hand-built table (above); the generated table uses the same mechanism with extra sort columns and a three-level hierarchy and was verified in Excel (loads, works as date table in a PivotTable, accepts a new hierarchy and a new Sort by Column). The non-December `yearEndMonth` variant (February year end) was verified in Excel the same way.

### Hidden columns

```csharp
new ColumnDefinition("MonthNumber", ColumnDataType.Int64) { IsHidden = true }
// or, for any table, including one loaded with AddTable:
sales["ProductKey"].IsHidden = true;
```

A hidden column stays in the model (relationships, sorting, hierarchies and DAX can still use it) but does not appear in the PivotTable field list, the usual treatment for keys and sort-order columns. `AddDateTable` hides its sort-order columns this way. Confirmed in Excel: the model loads, the fields are hidden, and hierarchies, sorting and measures keep working. Confirmed in Excel on freshly generated files, including all ten columns of Excel's own H1 sample (keys, hierarchy levels and sort columns) hidden: after hiding a column, adding a hierarchy and a Sort by Column still work.

### Sort by column

```csharp
var dates = model.CreateTable("Dates", new[]
{
    new ColumnDefinition("Date", ColumnDataType.DateTime),
    new ColumnDefinition("Month", ColumnDataType.String) { SortByColumn = "MonthNumber" },
    new ColumnDefinition("MonthNumber", ColumnDataType.Int64),
});
// or, after creating the table (null restores the column's own order):
products["Category"].SortByColumn = products["CategoryOrder"];
```

Like Power Pivot's *Sort by Column*: Month lists January to December instead of April, August, ... in PivotTables, slicers and hierarchies. Every value of the column must belong to exactly one value of the sort column (several months may share a quarter number, but one month may not have two numbers); this is checked when saving. Values with equal sort values keep their own order and BLANK stays first. A column that is sorted cannot itself be a sort column. Hierarchy trees follow the sorted order.

Confirmed in Excel: sorted columns show in the right order in a PivotTable, and a further *Sort by Column* or *Create Hierarchy* on the same model works. In models with relationships this only works because of three details the library handles for you (see "Open questions" below): a hidden helper measure is added when a model has relationships but no measures (`DataModelOptions.AddHelperMeasure`), the table holding the relationships is written at version 2 (3 once hierarchies exist), the relationship tables carry the statistics Excel writes while a model is open, and the cube is written at the version Excel gives a saved model.

### Relationships

Relate tables the way Power Pivot's diagram view does: many-to-one, from a foreign key column to a unique key column. Add both tables first, then:

```csharp
model.AddRelationship(sales["ProductKey"], products["ProductKey"]);     // many side first, then the one side
model.AddRelationship(sales["CustomerKey"], customers["CustomerKey"]);
```

PivotTables can then slice `Sales` by any column of `Products` and `Customers`. Rules, all enforced:

- Both columns need the same data type, and the tables must differ and not already be related (directly or through a cycle).
- The key column of the "one" table must have no blanks and no duplicates (text is compared case-insensitively, like the engine does). This is checked when you call `Save`, which throws an `InvalidOperationException` naming the offending value.
- Foreign key values with no match are allowed (as in Excel); they show up as BLANK.

`SaveModelPart(stream)` writes just the `item.data` part, if you want to inject the model into a workbook of your own.

## Data types

| `ColumnDataType` | Power Pivot type | CLR source types (automatic mapping) | Storage |
|---|---|---|---|
| `Int64` | Whole Number | `long`, `int`, `short`, `byte`, `sbyte`, `uint`, `ushort`, `ulong` | Value encoding (31-bit range) or 32/64-bit hash dictionary |
| `Double` | Decimal Number | `double`, `float`, `decimal` (default) | Value encoding if all values are integral, else hash dictionary |
| `Currency` | Fixed Decimal Number | `decimal` with `DecimalAsCurrency = true` | Units of 1/10,000; value encoding with power-of-ten magnitude |
| `DateTime` | Date | `DateTime`, `DateTimeOffset`, `DateOnly`, `TimeOnly`, `TimeSpan` | OLE Automation date; value encoding for whole days |
| `Boolean` | TRUE/FALSE | `bool` | Hash dictionary of 0/1 |
| `String` | Text | `string`, `char`, `Guid`, anything else (`ToString()` with invariant culture) | Hash dictionary, uncompressed 1 MiB pages |

Behaviour to be aware of:

- **Bad values.** A value that cannot be converted to its column's type (`"abc"` in a Whole Number column) makes `AddRow` throw an `ArgumentException` naming the row, column and value; the row is not added, so you can log it and carry on. When the same happens while loading from an `IDataReader`, the load stops and the table cannot be saved (the error names the row), because part of the row was already taken.
- **Nulls.** `null` and `DBNull` become BLANK. `NaN` and ±∞ in `Double` columns also become BLANK, because the engine cannot store them.
- **Text is case-insensitive.** As in Power Pivot, text values that differ only by case or width are the same value: `"Beta"` and `"beta"` become one value, and the first spelling seen wins. Accents are significant. Sorting and de-duplication use `DataModelOptions.Culture`.
- **Names.** Table and column names lose the characters Power Pivot forbids (`. , ; ' : / \ * | ? & % $ ! + = ( ) [ ] { } < > "` and control characters) and surrounding spaces. `"Sales (USD)"` becomes `"Sales USD"`.

## Options (`DataModelOptions`)

| Option | Default | Meaning |
|---|---|---|
| `Culture` | en-US | Model locale: collation, currency format, month and day names, string hash algorithm. Needs a culture with a Windows LCID (not the invariant culture). See below. |
| `SegmentRowCount` | 1,048,576 | Rows per column segment. Must be a power of two from 16,384 to 16,777,216. Excel uses 2²⁰. |
| `SpillThresholdBytes` | 32 MiB | In-memory buffer per column before values spill to a temp file. |
| `TempDirectory` | system temp | Location of spill files and the temporary model stream. |
| `XpressCompression` | `true` | Xpress-compress the files inside the model, as Excel does. `false` stores them verbatim, which Excel also supports, and gives a larger file. |
| `CompressModelPart` | `false` | Deflate the `item.data` part in the zip. Excel stores it uncompressed. |
| `DecimalAsCurrency` | `false` | Map `decimal` to Currency instead of Double. |
| `DataSourceName`, `ConnectionString` | placeholder | The model's data source and connection. See [Limitations](#limitations). |
| `WorksheetName` | `Sheet1` | Name of the single, empty worksheet. |
| `AddHelperMeasure` | `true` | Add one hidden measure (`= 1`) when a model has relationships but no measures; Excel's *Sort by Column* needs it (see [Sort by column](#sort-by-column)). |

Every option has a default, so `new DataModelWriter()` is enough when en-US suits you. The culture is not taken from the machine by default, so that the same code writes the same file everywhere: the culture decides how text is sorted and de-duplicated, the currency symbol of `MeasureFormat.Currency` and the month and day names of `AddDateTable`. To follow the machine's locale, ask for it explicitly:

```csharp
using var model = new DataModelWriter(new DataModelOptions { Culture = CultureInfo.CurrentCulture });
```

This throws `ArgumentException` where the current culture is the invariant culture, which is usual in containers and CI with `LANG=C`; Excel needs a locale with a Windows LCID.

## Large data

- **Two passes, streaming.** Values are buffered per column, in memory up to `SpillThresholdBytes` and then in a temp file. At `Save` each column is encoded one segment (1M rows) at a time and streamed into the model file. Memory holds the distinct values of each column, plus one segment per column during encoding. The engine needs the same.
- **Throughput.** On a modest Linux VM, 3,000,000 rows × 8 mixed columns (including million-distinct floats, dates with times and 64-bit integers) take about 10 s. That's a Release build with compression on. Low-cardinality data is much faster and much smaller.
- **Excel's own limits still apply.** Up to 1,999,999,997 rows per table and a 2 GB workbook file. 32-bit Excel has a much smaller memory ceiling for data models.

## Verification

Correctness was established four ways. The byte-level comparison and decoder checks were done without Excel, and the final check was done in Excel:

1. **Byte-for-byte against Excel.** An Excel 365 workbook (2026 build) was decoded. The same rows were written with this library and every file of the model compared:
   - Identical: the value-encoded column data (`.idf`), the hash indexes (`.hidx`), the position/ID hierarchy maps, the RowNumber column, the string dictionary and both text hierarchy maps.
   - Identical after normalising generated IDs and timestamps: the XML metadata (table/segment statistics, compression descriptors, ASSL object definitions).
   - Different, as expected: the storage IDs, the random encryption key, the data source name, and the order of the `DataFileList` entries, which Excel writes in hash order.
   - The SHA-256 hashes of Excel's files are pinned in the test suite.
2. **Independent decoder.** Generated workbooks were read back with [pbixray](https://github.com/Hugoberry/pbixray), an open-source VertiPaq reader that uses Microsoft's Xpress decoder. Every value matched the source for:
   - Tables of up to 3,000,000 rows, in multiple segments.
   - All six data types, each encoding path, nulls, case merging, and multi-page string dictionaries (300,000 distinct strings).
   - Empty tables and all-BLANK columns.
3. **Unit tests.** The suite (`tests/`) covers:
   - The CRC check value.
   - An Xpress round trip through a decoder written separately from [MS-XCA].
   - Decrypting a connection string written by Excel.
   - Container integrity: header, directory, every CRC and size.
   - The golden-file comparisons above, with compression both on and off.
   - Multi-segment metadata, naming rules, and empty/BLANK edge cases.

4. **Excel itself.** Workbooks from this library were opened in desktop Excel (Windows, Microsoft 365) and the data model loaded:
   - Manage Data Model lists the tables with their rows.
   - Covered: a three-table model (Whole Number, Decimal, Fixed Decimal, Date, TRUE/FALSE and Text columns), en-US and en-ZA locales, and compression on and off.
   - Covered: columns with one or two distinct values, with and without BLANKs.
   - This testing found one bug that no decoder had caught. See [Lessons from testing in Excel](#lessons-from-testing-in-excel).

If Excel ever reports "An issue with the Data Model is preventing Microsoft Excel from opening this workbook", try `XpressCompression = false` first to rule out the compressor. Then please report the exact column types and distinct-value counts involved.

### Lessons from testing in Excel

- **Database id.** Excel names the model database with 20 upper-case hex digits. A GUID loads and pivots fine, but with two relationships from one table Excel then fails to edit the model ("Create Hierarchy" gives `ASDetail::OnClone2`). The library uses Excel's form.

Independent decoders (pbixray, the test suite's reader) are more forgiving than Excel's engine. The engine rejected the model when any column had only one or two distinct values and no BLANKs, because the hierarchy tables of such columns were segmented differently from the files Excel and Power BI write:

| Distinct values (incl. BLANK) | Hierarchy segments |
|---|---|
| 1 | 1, 1, 1, 1 |
| 2 | 2, 2, 1 |
| 3 | 3, 3 |
| 6 | 6, 3 |

The rule is that the records (distinct values plus the three reserved ids, or two when there are BLANKs) are chunked in units of the distinct count. `Tests.TinyCardinalityHierarchy` pins this layout.

Relationships needed several more details before Excel accepted them. Excel gave no useful message for any of them: it either loaded the model without the relationship, or failed with a generic data model error. Each was found by rebuilding Excel's own sample model and diffing the output against it (`Tests.KeyColumnFlags`, `Tests.WorkbookRelationships` and `Tests.RelationshipMetadata` pin them):

- `xl/workbook.xml` lists every relationship in `<x15:modelRelationships>`; the diagram view reads it from there.
- The many-side dimension declares each relationship in a `<ddl300_300:Relationships>` block, and its measure group has a matching `ReferenceMeasureGroupDimension`. Without the block, a PivotTable fails with "the relationship with the ID ... cannot be found in the ... dimension".
- The primary key column is flagged unique: `ColumnFlags` 11, `EncodingHint` 1, column statistics `Usage` 0, `NullProcessing=Error` on its attribute key, and `Cardinality=One` from the row number to it. The foreign key column has `Usage` 1.

## Open questions

### Why do Sort by Column and Create Hierarchy in Excel need these odd details?

In a model with relationships, Excel's *Sort by Column* and *Create Hierarchy* fail with `An unexpected error occurred (file 'mddetail.cpp', line 1670, function 'ASDetail::OnClone2')` unless the details below hold. They were found by bisecting about forty test files in Excel; each item was confirmed by a file that had the others and failed without it. The library writes all of them whenever a model has relationships:

1. **Placeholder statistics for the relationship tables (`R$...`)**: DistinctStates 0, MinDataID/MaxDataID 2, RowCount 0, XMType 0, the compression width stays real. This is what Excel writes for relationships created while a model is open. Fully processed statistics (the state of a refreshed model) make *Sort by Column* fail.
2. **The table holding the relationships (the many side, for example Sales) has object version 2** (`Sales_<id>.2.dim.xml`, `info.2.xml`). With two relationships version 2 works and version 3 fails.
3. **The model contains at least one real measure command.** A hidden measure is enough (`DataModelOptions.AddHelperMeasure` adds one when there are relationships and no measures). The `__No measures defined` placeholder and the `__XL_Count of Models` default member alone are not.
4. **For *Create Hierarchy*: the cube, the measure groups and the calculation script are at object version 1** (as in a model Excel saved), or the partitions have data version 0. Models with cube version 0 and partition data version 1 sorted but could not create hierarchies; the library writes both.
5. **Models with user hierarchies are written the way Excel saves them after *Create Hierarchy*** (confirmed: further hierarchies and Sort by Column work, and saving and refreshing the PivotTable works). Version counters must match the file names: hierarchy (`U$`) files and the table that holds them are version 2 (`2.U$...idf`, `<Table>.2.tbl.xml`, dimension data/structure version 2, object version 5, no `info` file), the relationship (many-side) table is version 3, measure groups and the cube are version 18 (structure 8 and 11), the script and the partitions are version 9 (partition data version 1 on tables with hierarchies), key columns of those tables have EncodingHint 0, and each measure group's `PartitionFileList` must name the renamed partition file (otherwise Excel reports "Measure group must have at least one partition"). Setting counters without renaming files makes the model fail to load.

6. **Every measure group has its own `Ordinal` (1, 2, 3 in table order; partitions and the calculation script 1)** in models with hierarchies or a date table. With all Ordinals 0 the engine tells the measure groups apart only by the hash order of their ids, so whether *Create Hierarchy* and *Sort by Column* worked depended on the random table ids, and one changed character in an id could flip the result.

Models without relationships work without any of this. Several other combinations that looked plausible did not: normal or partly changed statistics behaved inconsistently (changing only XMType worked on one file and failed on another that already had sort columns), so the library copies the complete state Excel writes.

What is not understood:

- Why a measure matters at all. The measure does not reference the sorted column, and every other file written by Excel differs from ours only in version counters, compression and the database id.
- Why the statistics behaved inconsistently when changed one field group at a time.
- Why the table version has to be exactly 2 and not the number of relationships plus one.
- Whether the same conditions matter for other operations that make Excel clone a measure group (changing a data type or format, renaming, adding a calculated column, hiding columns, marking a date table). Hidden columns and date tables are the next features to be tested.

A way to settle it would be a disassembly or trace of `ASDetail::OnClone2` (the in-process engine used by Power Pivot), or Excel-written files created by adding a relationship, a measure, a sort and a hierarchy in different orders and diffing each save.

See also `docs/SPEC.md` (full format specification) and `docs/HANDOFF.md` (working notes for contributors).

## Limitations

These are not supported yet:

- **Calculated columns cannot be a row or column field until Excel has calculated them.** `AddCalculatedColumn` writes only the formula, and Excel calculates the values on demand. A PivotTable user can put the column in *Values* straight away, but cannot drag it into *Rows*, *Columns* or a filter until it has been calculated once: use it as a value first, or click *Calculate Now* in the Power Pivot window. *Refresh All* does not calculate it. Excel's own files behave the same way. Plan reports accordingly: a calculated column is fine as a value, but a label column that people group by needs that extra step. There is no workaround in the file yet.
- **KPIs and perspectives.** You can add them in Excel after opening the file. Measures, relationships, hierarchies and sort by column are supported (see above, and verified in Excel), but only many-to-one with a single column on each side, no inactive relationships, and one relationship per pair of tables.
- **Refresh.** The data is written directly. The model's data source is a placeholder OLE DB connection, stored encrypted the way Excel stores it, so *Refresh* in Excel reports a connection error and keeps the existing data. Set `ConnectionString` if you want it to point somewhere real.
- **Excel 2010 PowerPivot add-in.** The model uses compatibility level 1103, the format of Excel 2013 and later.

These parts are inferred rather than copied from an Excel sample:

- **64-bit integer dictionaries.** The layout is used only when a whole-number column has a value range over 2³¹ *and* values outside 32 bits. It is consistent with the spec and decodes correctly, but I haven't seen one written by Excel.
- **Empty tables (0 rows).** The hierarchy layout was inferred, and no Excel sample was available.
- **Format strings.** Annotations are General / Text / General Date / Currency (from the locale), as Excel writes them.

## How it works

```
Workbook (.xlsx, OPC zip)
├─ xl/workbook.xml           x15:dataModel → model tables
├─ xl/connections.xml        model connection + ThisWorkbookDataModel
└─ xl/model/item.data        Spreadsheet Data Model = Analysis Services backup (ABF)
   ├─ header page            BOM, signature, BackupLogHeader XML
   ├─ files                  4 KiB chunks [u16 size, u16 stored size] + Xpress, CRC-32/BZIP2 per file
   │  ├─ PARTITIONS          partition list
   │  ├─ <db>.1.db.xml, Model.0.cub.xml, *.det.xml, *.prt.xml, Sandbox.0.dsv.xml, *.ds.xml (ASSL 1103)
   │  ├─ 0.CryptKey.bin      3DES session key (exponent-of-one RSA wrapping)
   │  └─ <table>.0.dim\      per table:
   │     ├─ <table>.0.tbl.xml                XMSimpleTable: columns, segments, statistics
   │     ├─ 0.<table>.<col>.0.idf            column data: hybrid RLE + bit-packed sub-segments
   │     ├─ 0.<table>.<col>.dictionary       hash dictionaries (string / long / real)
   │     ├─ H$<table>$<col>.0.tbl.xml        attribute hierarchy table
   │     ├─ 0.H$…POS_TO_ID.0.idf / ID_TO_POS sort-order maps
   │     └─ 0.H$…hidx                        data-id → position hash index (value encoding)
   ├─ LOG                    backup log: logical paths → stored files
   └─ virtual directory      offsets and sizes of every stored file
```

Source layout:

- `src/PivotWright/`
  - `DataModelWriter`, `ModelTable`, `ModelColumn`, `ModelMeasure`, `ModelHierarchy`, `ModelRelationship`, `ModelCollection`, `ColumnDefinition`, `DataModelOptions`: the public API.
  - `Internal/Abf`: container, chunking, CRC, Xpress compressor, backup log.
  - `Internal/VertiPaq`: segment encoder, bit packing, dictionaries, hash index.
  - `Internal/Model`: per-type column builders, encoding decisions, hierarchies.
  - `Internal/Metadata`: XMObject (`.tbl.xml`) and ASSL writers.
  - `Internal/Crypto`: CryptKey.bin and connection-string encryption.
  - `Internal/Package`: the `.xlsx` parts.
- `tests/PivotWright.Tests/`: a dependency-free test runner (`dotnet run --project tests/PivotWright.Tests`).
- `samples/PivotWright.Demo/`: synthetic, CSV and multi-table examples.
- `tools/verify_with_pbixray.py`: reads a generated workbook back with pbixray and optionally compares it to a CSV.

### Relationships in the file

Each relationship adds three things, copied from the structure Excel writes: an `R$<table>$<guid>` table in the many table's folder whose single `INDEX` column maps every data id of the foreign key to the 0-based row of the matching primary row (-1 for none; the first three ids are reserved); an `XMRelationship` entry in the many table's `.tbl.xml`; and a `ReferenceMeasureGroupDimension` in its measure group. Value-encoded foreign keys spanning more than 4M ids fall back to hash encoding so the index stays small.

## Building

```
dotnet build
dotnet run --project tests/PivotWright.Tests
dotnet run --project samples/PivotWright.Demo -c Release -- synthetic 1000000 big.xlsx
```

The library has no package dependencies.

## References

- [MS-XLDM] Spreadsheet Data Model File Format: https://learn.microsoft.com/openspecs/office_file_formats/ms-xldm/
- [MS-XCA] Xpress Compression Algorithm: https://learn.microsoft.com/openspecs/windows_protocols/ms-xca/
- Power Pivot naming requirements: https://learn.microsoft.com/previous-versions//ff452213(v=technet.10)

## License

MIT. See [LICENSE](LICENSE).
