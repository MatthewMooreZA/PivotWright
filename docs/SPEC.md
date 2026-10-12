# PivotWright — specification as known so far

Status of this document: written at the end of the first working session (October 2026) from the library source, the README, and about 100 Excel-tested files. It has two kinds of statement, and each section says which it is:

- **Verified in Excel** (desktop Excel 365, Windows): a generated or edited file loaded and behaved. "Verified" never means "understood": see the open questions.
- **Observed in Excel-written files** (diffed against samples in `reference/excel-samples/`) but not yet implemented or not yet tested.

Excel gives only generic errors (`ASDetail::OnClone2`, "Measure group ... must have at least one partition", "Needs Calculation" warnings, a credentials prompt on refresh). Progress has come from bisecting Excel-tested files and diffing against files Excel wrote itself.

Normative sources: [MS-XLDM] (Spreadsheet Data Model file format), [MS-XCA] (Xpress compression), Analysis Services ASSL at compatibility level 1103. Everything below that is not in those documents was learned empirically.

---

## 1. Goal and scope

`PivotWright` is a pure-managed .NET 8 library that writes `.xlsx` workbooks containing a Power Pivot data model (`xl/model/item.data`, an Analysis Services backup "ABF" holding VertiPaq files) with no Excel, SSAS or COM at write time. Input is an `IDataReader`, `DataTable`, objects or rows. Long-term aim: an Excel report modeller/writer. The model is the first layer.

Supported and verified in Excel: tables, six data types, relationships (many-to-one), measures, user hierarchies, Sort by Column, hidden columns, date tables (marking an existing table and the generated `AddDateTable`). Backlog: calculated columns (next), KPIs, perspectives, refresh with a real data source.

## 2. Public API (all in namespace `PivotWright`)

| Member | Notes |
|---|---|
| `DataModelWriter(DataModelOptions)` | entry point, `IDisposable` |
| `AddTable(name, IDataReader / DataTable / IEnumerable<T>)`, `CreateTable(name, ColumnDefinition[])` → `TableWriter.AddRow(params object?[])` | tables |
| `AddRelationship(manyTable, manyColumn, oneTable, oneColumn)` | many-to-one, single column each side, no inactive relationships, one per pair of tables, no cycles |
| `AddMeasure(table, name, daxExpression, MeasureFormat)` | DAX stored as written; names unique across model, no `[` `]` |
| `AddHierarchy(table, name, params string[] levels)` | 2+ columns of one table; name unique within table and different from column names |
| `SetSortByColumn(table, column, sortColumn)` (+ `TableWriter` overload); `ColumnDefinition.SortByColumn` | see section 7 |
| `ColumnDefinition { Name, DataType, SortByColumn, IsHidden }` | `IsHidden` writes `AttributeHierarchyVisible=false` |
| `DataModelOptions` | `Lcid` (1033 default; the user's files use 7177 = en-ZA), `SegmentRowCount` (2^20), `SpillThresholdBytes`, `TempDirectory`, `XpressCompression`, `CompressModelPart`, `DecimalAsCurrency`, `DataSourceName`, `ConnectionString`, `WorksheetName`, `AddHelperMeasure` (default true) |
| `Save(path)`, `Save(Stream)`, `SaveModelPart(Stream)` | |
| `MeasureFormat` | General, WholeNumber, Decimal(n), Currency(n), Percentage(n), Date |

Data types: Int64, Double, Currency, DateTime, Boolean, String (see README table for encodings). Names lose `. , ; ' : / \ * | ? & % $ ! + = ( ) [ ] { } < > "` and control characters.

## 3. Container (ABF) — verified, byte-identical to Excel for column data

- `item.data`: header page (BOM, signature, BackupLogHeader XML), files in 4 KiB chunks `[u16 size, u16 stored size]` with Xpress compression and a CRC-32/BZIP2 per file, `PARTITIONS` list first, then ASSL object files, `0.CryptKey.bin`, per-table dimension folders, then `LOG` (backup log) and the virtual directory.
- The backup log (`FileGroup` per ASSL object) records `Class`, `Id`, `Name`, `ObjectVersion`, `PersistLocationPath`, `ObjectId`, and each file with `Path`, `StoragePath`, `Size`, time. **The log's ObjectVersion and the file names must agree** (section 5).
- Database id: 20 upper-case hex characters (a GUID loads but breaks editing with two relationships).
- Encrypted connection string/password use a 3DES session key (`CryptKey.bin`, exponent-of-one RSA wrapping). A refresh with the placeholder connection reports a credentials error; that is expected.

## 4. Objects and files per table

Folder `<Table_GUID>.0.dim\` holds per-table data; the ASSL dimension is `<Table_GUID>.<v>.dim.xml` at database level, the measure group `Model.0.cub\<Table_GUID>.<v>.det.xml`, the partition `...\<Table_GUID>.0.det\<Table_GUID>.<v>.prt.xml` (+ `...0.prt\info.<v>.xml`).

Per table: `<T>.<v>.tbl.xml` (XMSimpleTable: columns, segments, statistics, `XMUserHierarchy` entries), per column `0.<T>.<col>.0.idf`, optional `.dictionary`, `H$<T>$<col>.0.tbl.xml` plus `0.H$...POS_TO_ID.0.idf`, `ID_TO_POS`, `hidx` (attribute hierarchy), relationship tables `R$<T>$<guid>.<v>.tbl.xml` + `<v>.R$...INDEX.0.idf`, hierarchy tables `U$<T>$<name>.<v>.tbl.xml` + 4 idf files.

Statistics, segmenting, compression: see README ("Lessons from testing in Excel") and `Internal/VertiPaq`. One rule worth repeating: hierarchy tables of columns with 1–2 distinct values must be segmented like Excel (1,1,1,1 / 2,2,1 / 3,3 / 6,3).

## 5. Version counters (the central, empirical part)

Rule that explains the earlier failures: **a version counter written inside a file must match the version in its file name and the backup-log ObjectVersion.** Setting counters without renaming files makes the model fail to load.

### 5.1 Base profile (no hierarchies, verified)

| Object | Version | File name |
|---|---|---|
| database | 1 | `<db>.1.db.xml` |
| cube / measure groups / script | 1 | `Model.1.cub.xml`, `<T>.1.det.xml`, `MdxScript.1.scr.xml` |
| partitions | 1 | `<T>.1.prt.xml`, `info.1.xml` (partition DataVersion 0 in prt.xml) |
| dimension (no relationships) | 1 | `<T>.1.dim.xml`, `info.1.xml` |
| dimension holding relationships (many side) | 2 | `<T>.2.dim.xml`, `info.2.xml` |
| tbl.xml | 0 | `<T>.0.tbl.xml` |
| cube info | 1 | `Model.0.cub\info.1.xml` |

### 5.2 Hierarchy profile (verified in Excel: load, pivot, add hierarchies, add sort columns, save and refresh)

Applies when the model has at least one user hierarchy. `H` = tables holding a hierarchy.

| Object | Version / value |
|---|---|
| U$ tables | `U$<T>$<name>.2.tbl.xml`, Version 2, column DataVersion 2; data files `2.U$<T>$<name>.<col>.0.idf` |
| main tbl.xml of tables in H | `<T>.2.tbl.xml`, Version 2 |
| dimension of tables in H | ObjectVersion 5 (`.5.dim.xml`), StructureVersion 2, DataVersion 2, **no `info` file** |
| many-side dimension not in H | ObjectVersion 3 (`.3.dim.xml`, `info.3.xml`) (2 fails once hierarchies exist; 3 works) |
| measure groups | ObjectVersion 18, StructureVersion 8 (`.18.det.xml`); DataVersion -1 for tables in H, 1 for the many side, else 0 |
| cube | ObjectVersion 18, StructureVersion 11 (`Model.18.cub.xml`), **no cube info file** |
| script | ObjectVersion 9 (`MdxScript.9.scr.xml`) |
| partitions | ObjectVersion 9 (`.9.prt.xml`, `info.9.xml`); partition DataVersion 1 for tables in H, else 0 |
| measure group `PartitionFileList` and cube `MeasureGroupFileList`/`AssemblyFileList` | must name the renamed files (a stale `.1.prt.xml` here produced "Measure group must have at least one partition", "Needs Calculation" and a refresh password prompt) |
| EncodingHint on `__XL_RowNumber` and primary-key columns of tables in H | 0 (1 otherwise) |
| H$ column tables and normal column data | stay `.0.` |

Excel then keeps bumping counters on each operation (a second operation gave dims 8/3/3, dets 31/14, cube 31/20, script 15). The library does not need to follow that; the profile above is enough for Excel to carry on editing.

### 5.3 Conditions for Sort by Column / Create Hierarchy on models with relationships (verified; found by bisecting ~40 files)

1. Relationship (`R$`) statistics are the placeholder form: DistinctStates 0, Min/MaxDataID 2, RowCount 0, XMType 0, real compression width.
2. The many-side table is at object version 2 (3 with hierarchies).
3. At least one real measure command exists. A hidden helper measure is auto-added (`AddHelperMeasure`) when there are relationships but no measures. The `__No measures defined` placeholder and `__XL_Count of Models` default member alone do not satisfy it.
4. Cube, measure groups, script at object version 1 (file names `.1.`) or partitions at DataVersion 0 (for Create Hierarchy). Superseded by 5.2 when hierarchies exist.

Why these matter is not understood (README, "Open questions").

## 6. Relationships (verified)

- `xl/workbook.xml` lists them in `<x15:modelRelationships>`.
- Many-side dimension declares `<ddl300_300:Relationships>`; its measure group has a `ReferenceMeasureGroupDimension`.
- Primary key column: `ColumnFlags` 11, `EncodingHint` 1 (0 in H tables), stats `Usage` 0, `NullProcessing=Error`, `Cardinality=One` from the row number; foreign key `Usage` 1.
- Index table `R$<T>$<guid>` (single `INDEX` column): maps each foreign-key data id to the 0-based primary row, -1 for none; first three ids reserved. Value-encoded keys spanning more than 4M ids use hash encoding.
- Orphan foreign key values are allowed (they show as BLANK). Primary key must be unique and non-blank (checked at `Save`).

## 7. Measures, hierarchies, Sort by Column

- **Measures**: MdxScript commands (`CREATE MEASURE 'Table'[Name]=...;`) with calculation properties and annotations (`Type=User`, `IsPrivate`, `Format`). With measures present the cube default member is `__XL_Count of Models`, otherwise the placeholder. Hidden measures use `Visible=false`.
- **User hierarchies**: dim.xml `<Hierarchies>`; tbl.xml `XMUserHierarchy` (`TableStore` like `$Year$0$Month$1$`, `TableName` `U$<T>$<name>`); U$ table is an XMSimpleTable with columns `MULTI_LEVEL_ID`, `PARENT_POS`, `FIRST_CHILD_POS`, `CHILD_COUNT`. Tree is built from sorted values (BLANK first, case-insensitive merge) and follows sort-by-column order. Byte-identical to Excel for the same data (unit-tested).
- **Sort by Column**: dim.xml attribute `<OrderBy>AttributeKey</OrderBy>` + `<OrderByAttributeID>SortColId</OrderByAttributeID>`; tbl.xml raw column `<OrderByColumn>SortColId</OrderByColumn>`. The sorted column's `POS_TO_ID`/`ID_TO_POS`/`hidx` are ordered by the sort column (BLANK first, ties keep their own order). Each value of the sorted column must map to exactly one sort value (checked at `Save`); a sorted column cannot be a sort column.

**Hierarchy StructureType: Natural or Unnatural (found 2026-10-12 with `datetbl_gen_ye2_v1.xlsx`, fixed with the help of `H5_sort_reapplied.xlsx`)**. A user hierarchy is *natural* when every member of a level belongs to exactly one member of the level above (many-to-one from child to parent; a Product has one Category). It is *unnatural* when a member has several parents: with a February year end, quarter "Q4" sits under FiscalYear 2025 and under FiscalYear 2026. Analysis Services decides this by examining the data when the model is processed and stores it as `<ddl300:StructureType>` in the dimension's `<Hierarchy>` (enumeration `Natural`, `Unnatural`, `Unknown`). Excel's PivotTable orders the members of an unnatural hierarchy by the stored tree (our U$ table, which follows Sort by Column); with `Natural` written for such data the pivot ordered a level by data id instead (Fiscal Calendar showed Q4, Q1, Q2, Q3 under FY2026). We used to write `Natural` always, which only looked right when the sort order equalled the order of first appearance in the data.

Evidence: after Excel re-applied Sort by Column on FiscalQuarter, the only substantive change in `H5` was `Natural` -> `Unnatural` on the Fiscal Calendar hierarchy (the Calendar hierarchy, one parent per member, stayed Natural); every data file (idf, dictionary, hidx, U$ columns) was byte-identical to ours, the rest being counter bumps. Rule implemented in `HierarchyEncoder` (`EncodedHierarchy.Natural`): for each level after the first, each distinct value (BLANK included) must always occur under the same value of the previous level. The rule is inferred from this one sample plus the documentation below; levels further apart than neighbours are not checked. A multi-year Calendar hierarchy (Year, Quarter, Month) is therefore Unnatural, which was verified in Excel with `datetbl_gen_multiyear_v1` (2024 to 2026, December year end: loads, correct order, add hierarchy and Sort by Column work) and with `datetbl_gen_multiyear_ye6_v1` (June year end, both Calendar and Fiscal Calendar Unnatural: same result).

References: Microsoft, [HierarchyStructureType enumeration](https://dotnet.territoriali.olinfo.it/en-us/previous-versions/sql/sql-server-2012/gg502673(v=sql.110)) (mirror of the SQL Server 2012 AMO documentation: Natural, Unnatural, Unknown); Microsoft, [Attribute Relationships](https://technet.microsoft.com/en-us/ms174557) (a relationship requires a many-to-one relationship between child and parent members); Chris Webb, [Natural and unnatural hierarchies in the SSAS 2012 Tabular model](https://blog.crossjoin.co.uk/2012/02/28/natural-and-unnatural-hierarchies-in-the-ssas-2012-tabular-model/) (in Tabular the engine determines this from the data while processing).

## 8. Hidden columns (observed; implemented and Excel-verified on a generated file, 2026-10-12)

Excel hides a column by setting `<AttributeHierarchyVisible>false</AttributeHierarchyVisible>` on its attribute in the table's dim.xml **and** in the cube's dimension attribute list (`Model.cub.xml`). No tbl.xml change. `ColumnDefinition.IsHidden` already writes both. Excel also bumps counters (cube 24, measure group structure 9) and writes a new `Sales info` file; **not needed**: `datetbl_gen_defaults_v1` (six hidden sort columns, relationships, two hierarchies, sort-by, date table, no counter bumps) loaded, hid the fields and accepted a new hierarchy and Sort by Column in Excel. Untested: a hidden column that is a hierarchy level, hidden columns in a model without hierarchies. In `H1_hidden_columns.xlsx` the user hid four attributes per table.

**Test result 2026-10-12 (`hidden_h1equiv_v1.xlsx`, V2b with the same ten columns hidden as `H1`, no counter bumps): loads, columns hidden, hierarchies and sort order work, but every edit (add hierarchy, hide another column, add Sort by Column) fails with `ASDetail::OnClone2`.** Excel's `H1` and ours are identical in every ASSL file apart from counters and list order, in every `.tbl.xml` (GUIDs normalised) and in every data file, so the counters are the only candidate. (The date-table file `datetbl_gen_defaults_v1`, which hides six sort columns and carries the date-table counters, did accept edits; its hidden columns were sort-number columns only, no keys or hierarchy levels, which may or may not matter.)

Counter deltas Excel wrote for the same ten hides (Dates: Date, Year, Month, MonthNumber; Products: ProductKey, Category, CategoryOrder, Product; Sales: ProductKey, Date):

| Object | H1 vs hierarchy profile (V2b) | H3 vs H2 (date-table state) |
|---|---|---|
| dimension ObjectVersion Dates / Products / Sales | 5 to 9 / 5 to 10 / 3 to 5 (+4/+5/+2) | 6 to 10 / 6 to 11 / 4 to 6 (+4/+5/+2) |
| measure group ObjectVersion (all three) | 18 to 24 (+6) | 21 to 27 (+6) |
| measure group StructureVersion Dates and Products (Sales unchanged) | 8 to 9 (+1) | 9 to 11 (+2) |
| cube ObjectVersion | 18 to 34 (+16) | 21 to 37 (+16) |
| cube StructureVersion | 11 to 16 (+5) | 12 to 19 (+7) |
| script, partitions, Ordinals | unchanged (9, 9, all 0) | unchanged |

The ObjectVersion deltas are the same in both samples; the StructureVersion deltas differ, so they depend on the history of operations. Products gets one more than its four hidden columns, which is unexplained. A per-operation rule needs smaller samples (HANDOFF).

**Single-operation samples (received 2026-10-12)**, relative to V2b: `H8_hide_one_column.xlsx` (Sales[Qty], a plain column) and `H9_hide_one_hierarchy_level.xlsx` (Products[Product]). Both: cube 18 to 21 (structure 11 to 12), every measure group 18 to 20 (structure unchanged), script and partitions unchanged, no tbl.xml or data file renamed, only the attribute visibility in the dimension and the cube changed. Dimension of the table with the hidden column: Sales (not in a hierarchy) 3 to 4 with `info.4.xml`; Products (holds a hierarchy) 5 to 7. In H8 Products' dimension also went 5 to 6 although only Sales was touched (Dates did not move); unexplained, not reproduced. Ten hides in `H1` were not ten times this (measure groups +6, cube +16), so the counters follow Excel's save sequence, not the number of columns.

Implementation (`ModelInfo.HiddenProfile`: hidden columns in a model with hierarchies and no date table): the state after one hide operation; measure groups get their own version (`MeasureGroupVersion` 20, the cube 21 and cube structure 12); the dimension of each table with hidden columns +1, or +2 when it holds a hierarchy. Reproduces `H9` exactly and `H8` apart from the unexplained Products bump. **Excel test 2026-10-12 (`hidden_h1equiv_v2_counters.xlsx`, ten hidden columns with these counters): same result as without them: loads, hidden, hierarchies and sort order work, every edit fails in `OnClone2`. The counters are therefore not (or not alone) the cause; the profile is kept only because it reproduces H8/H9, and is not shown to be needed.** What passes: `datetbl_gen_defaults_v1` (date-table state, six hidden sort-number columns, edits worked). What fails: hierarchy state with hidden keys and hierarchy levels. Open: whether Excel itself can edit its own `H1`/`H8`/`H9` after reloading them, and which kind of hidden column (plain, sort-number, key, hierarchy level) triggers it.

## 9. Date table, marking an existing table (observed in `H2_date_table.xlsx`; implemented, **not yet tested as a generated file in Excel**)

Marking `Dates[Date]` as the date table, compared with `V2b_pivotwright_hierarchies_verified.xlsx`:

- Dates dim.xml: dimension `<Type>` Regular → **Time**. `__XL_RowNumber` attribute `Usage` Key → Regular and its `AttributeRelationship` to Date is dropped (it keeps Year, Month, MonthNumber). The `Date` attribute becomes `Usage` Key and gets an `AttributeRelationship` to `__XL_RowNumber`.
- Dates tbl.xml: `Version` 2 → 3 (`Dates_G.3.tbl.xml`); `ColumnFlags` of `__XL_RowNumber` 31 → 27 and of `Date` 11 → 15; the Date column's XMHierarchy hidx name becomes `2.H$Dates_G$Date.hidx`.
- Date column attribute-hierarchy table renamed to version 2: `H$Dates_G$Date.2.tbl.xml` (Version 2), `2.H$Dates_G$Date.POS_TO_ID.0.idf`, `2.H$Dates_G$Date.hidx`, column DataVersion 2.
- The relationship whose primary side is Dates: `R$Sales_G$<guid>.2.tbl.xml` (Version 2, DataVersion 2) and `2.R$Sales_G$<guid>.INDEX.0.idf`. The other relationship stays at 0.
- Sales tbl.xml: `Version` 0 → 1 (`Sales_G.1.tbl.xml`); Sales dimension `info.4.xml` (was `info.3`).
- Counters: cube 18 → 21 (StructureVersion 11 → 12), script 9 → 10, partitions 9 → 10 (`info.10.xml`, `.10.prt.xml`), measure groups 18 → 21 (StructureVersion 8 → 9, CacheVersion 1 → 2), `PartitionFileList` updated.
- Every dimension object version goes up by one (Products 5 → 6, Dates 5 → 6, Sales 3 → 4). Dates dimension: StructureVersion 3, DataIDVersion 2, IncDataIDVersion 2, CacheVersion 2, DataVersion 3. Sales dimension: DataVersion 0 → 1; Sales partition DataVersion 0 → 1. Products is unchanged apart from the object version.
- H2 vs H3 (date table + hidden columns) differ only in the hidden attributes.

Implementation: `DataModelWriter.MarkAsDateTable(table, dateColumn)` extends the hierarchy profile (`ModelInfo.DateTables`, `DateReferencingTables`, `TableVersion`; `EncodedColumn.IsDateKey`/`AttributeHierarchyVersion`; `EncodedRelationship.Version`). The output was diffed against `H2` (counters, file names, tbl.xml and dim.xml all equal; only list order, `<Value/>` elements and data-dependent sizes differ).

Only the configuration of `H2` has been seen, so `ModelSerializer.ValidateDateTables` refuses everything else: the date table needs a hierarchy, a relationship must use the date column as primary key, one date table per model, dates must be whole days without BLANK or duplicates, and the date column must be value-encoded. Unknown: a date table without hierarchy or relationship (ColumnFlags of a date key that is no primary key, probably 12), a model without hierarchies, several date tables, per-table counters for tables that are neither in a hierarchy nor on the many side.

**Test result (2026-10-11, `datetbl.xlsx` from `tools/datagen ... date`)**: the model loads and the date table works in a PivotTable, but Create Hierarchy and Sort by Column both fail with `ASDetail::OnClone2`. The generated file equals `H2` in every comparison made: all ASSL files (xd.py/dd.py, counters included), every `.tbl.xml` (GUIDs normalised), every binary data file (idf, dictionary, hidx, byte for byte; same dataset as V2b), the backup log (groups, object versions, file counts) and the PARTITIONS file. Remaining differences: `Ordinal` (Excel writes 1/2/3 on measure groups, 1 on partitions and script; we write 0; Excel's own samples show erratic ordinals, 0 in H1, 22-24 in S6xe), the order of entries in `DataFileList`/`MeasureGroupFileList`, `<Value/>` elements in table annotations (Excel drops them), the 17 `customXml` Gemini parts (UI widget settings) and workbook parts. Open question: does Excel itself manage to add a hierarchy / Sort by Column on `H2_date_table.xlsx` after reloading it? If not, this is Excel's limitation for reloaded date-table models; if yes, bisect the remaining differences (Ordinal first).

**Fix (verified 2026-10-11, second `datetbl.xlsx`)**: after marking a date table Excel numbers `Ordinal` (measure groups 1, 2, 3 in table order; partitions and script 1). Writing 0 made Create Hierarchy and Sort by Column fail in `OnClone2` even though everything else equalled `H2`; with the Ordinals set, the model loads, pivots, accepts a new hierarchy and a new Sort by Column. The remaining differences listed above (file-list order, `<Value/>`, `customXml`) do not matter. Ordinal is therefore a model requirement, not noise; check it first in any future profile.

**No "create from scratch" case**: Excel cannot create a date table by itself; a date table is always an existing table that was marked (Matthew, 2026-10-11).

## 10. Not understood (open)

1. Why a real measure is required before Sort by Column/Create Hierarchy on relationship models.
2. Why statistics changed one field group at a time behaved inconsistently; why many-side version must be exactly 2 (3 with hierarchies).
3. Whether the same conditions apply to other operations that clone a measure group (rename, change type/format, calculated column, hide, date table). Hidden columns and date tables are the next probes.
4. Counter semantics in general; Excel increases them on every operation, the library freezes a "just edited once" state.
5. Possible ways to settle it: disassembly or trace of `ASDetail::OnClone2`, or Excel-written files created by adding a relationship, measure, sort and hierarchy in different orders and diffing each save.
