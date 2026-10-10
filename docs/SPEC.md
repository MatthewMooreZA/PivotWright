# PivotWright — specification as known so far

Status of this document: written at the end of the first working session (October 2026) from the library source, the README, and about 100 Excel-tested files. It has two kinds of statement, and each section says which it is:

- **Verified in Excel** (desktop Excel 365, Windows): a generated or edited file loaded and behaved. "Verified" never means "understood": see the open questions.
- **Observed in Excel-written files** (diffed against samples in `reference/excel-samples/`) but not yet implemented or not yet tested.

Excel gives only generic errors (`ASDetail::OnClone2`, "Measure group ... must have at least one partition", "Needs Calculation" warnings, a credentials prompt on refresh). Progress has come from bisecting Excel-tested files and diffing against files Excel wrote itself.

Normative sources: [MS-XLDM] (Spreadsheet Data Model file format), [MS-XCA] (Xpress compression), Analysis Services ASSL at compatibility level 1103. Everything below that is not in those documents was learned empirically.

---

## 1. Goal and scope

`PivotWright` is a pure-managed .NET 8 library that writes `.xlsx` workbooks containing a Power Pivot data model (`xl/model/item.data`, an Analysis Services backup "ABF" holding VertiPaq files) with no Excel, SSAS or COM at write time. Input is an `IDataReader`, `DataTable`, objects or rows. Long-term aim: an Excel report modeller/writer. The model is the first layer.

Supported and verified in Excel: tables, six data types, relationships (many-to-one), measures, user hierarchies, Sort by Column. Backlog: hidden columns (implemented, not verified as a feature), date tables, perspectives, calculated columns, KPIs, refresh with a real data source.

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

## 8. Hidden columns (observed, implementation exists)

Excel hides a column by setting `<AttributeHierarchyVisible>false</AttributeHierarchyVisible>` on its attribute in the table's dim.xml **and** in the cube's dimension attribute list (`Model.cub.xml`). No tbl.xml change. `ColumnDefinition.IsHidden` already writes both. Excel also bumps counters (cube 24, measure group structure 9) and writes a new `Sales info` file; whether those are needed is untested. In `H1_hidden_columns.xlsx` the user hid four attributes per table.

## 9. Date table (observed in `H2_date_table.xlsx`, **not yet implemented**)

Marking `Dates[Date]` as the date table, compared with `V2b_pivotwright_hierarchies_verified.xlsx`:

- Dates dim.xml: dimension `<Type>` Regular → **Time**. `__XL_RowNumber` attribute `Usage` Key → Regular and its `AttributeRelationship` to Date is dropped (it keeps Year, Month, MonthNumber). The `Date` attribute becomes `Usage` Key and gets an `AttributeRelationship` to `__XL_RowNumber`.
- Dates tbl.xml: `Version` 2 → 3 (`Dates_G.3.tbl.xml`); `ColumnFlags` of `__XL_RowNumber` 31 → 27 and of `Date` 11 → 15; the Date column's XMHierarchy hidx name becomes `2.H$Dates_G$Date.hidx`.
- Date column attribute-hierarchy table renamed to version 2: `H$Dates_G$Date.2.tbl.xml` (Version 2), `2.H$Dates_G$Date.POS_TO_ID.0.idf`, `2.H$Dates_G$Date.hidx`, column DataVersion 2.
- The relationship whose primary side is Dates: `R$Sales_G$<guid>.2.tbl.xml` (Version 2, DataVersion 2) and `2.R$Sales_G$<guid>.INDEX.0.idf`. The other relationship stays at 0.
- Sales tbl.xml: `Version` 0 → 1 (`Sales_G.1.tbl.xml`); Sales dimension `info.4.xml` (was `info.3`).
- Counters: cube 18 → 21, script 9 → 10, partitions 9 → 10 (`info.10.xml`, `.10.prt.xml`), measure group `PartitionFileList` updated.
- H2 vs H3 (date table + hidden columns) differ only in the hidden attributes.

Implementation plan: `DataModelWriter.MarkAsDateTable(table, dateColumn)` (validate: DateTime column, no BLANKs, no duplicates, whole days, one per model is not required but Excel allows several); extend the hierarchy profile with the above; one most-likely test file first; user then checks load, pivot, add hierarchy, add Sort by Column, save and refresh.

## 10. Not understood (open)

1. Why a real measure is required before Sort by Column/Create Hierarchy on relationship models.
2. Why statistics changed one field group at a time behaved inconsistently; why many-side version must be exactly 2 (3 with hierarchies).
3. Whether the same conditions apply to other operations that clone a measure group (rename, change type/format, calculated column, hide, date table). Hidden columns and date tables are the next probes.
4. Counter semantics in general; Excel increases them on every operation, the library freezes a "just edited once" state.
5. Possible ways to settle it: disassembly or trace of `ASDetail::OnClone2`, or Excel-written files created by adding a relationship, measure, sort and hierarchy in different orders and diffing each save.
