# Handoff — moving from the claude.ai session to Claude Code

Owner: Matthew (accounting/actuarial practice, Johannesburg; tests in desktop Excel 365 on Windows, locale en-ZA, `Lcid = 7177`). Repo: `MatthewMooreZA/PivotWright` (renamed from `PowerPivotWriter`; the name was settled on 2026-10-10).

Read first: `docs/SPEC.md` (the full specification as known), `README.md` (API, verification, open questions). This file is the working context.

## 1. State at handoff

| Feature | Status |
|---|---|
| Tables, six data types, multi-segment, streaming, spill to disk | done, byte-compared with Excel |
| Relationships | done, verified in Excel |
| Measures (+ hidden helper measure) | done, verified |
| User hierarchies | done, verified, including adding more hierarchies and Sort by Column in Excel afterwards and save + refresh |
| Sort by Column | done, verified |
| Hidden columns | `ColumnDefinition.IsHidden`; **Excel-verified 2026-10-19/20**: the ten columns of Excel's `H1` (keys, hierarchy levels, sort columns) and a single hidden hierarchy level, in freshly generated models, load, display correctly, and accept hide, add hierarchy and add Sort by Column. The long-standing `OnClone2` failure after a hide was the all-zero `Ordinal` values (SPEC §8): every model now writes distinct Ordinals (measure groups 1..n, partitions and script 1), rechecked on a plain model, a two-table model without a measure, a generated date table and the standard hierarchy model. The hidden-column counters/timestamps of `ModelInfo.HiddenProfile` reproduce H8/H9 but were never shown to be needed. Not tested in Excel: hiding columns in a model with no hierarchy and no relationships beyond the `plain` and `twotable` cases |
| Generated date table (`AddDateTable`) | **Excel-verified 2026-10-12**: standard calendar with `Calendar` hierarchy, marked as date table; one calendar year (December), several years (December and June year ends), February year end, and the defaults (whole-year range, YearMonth, hidden sort columns) with `datetbl_gen_defaults_v1`. Options `yearEndMonth`, `yearLabel` (default "Fiscal"), `completeYears`, `includeYearMonth`, `hideSortColumns`. Needed a fix found with `H5`: hierarchy StructureType must be Unnatural when a member has several parents (SPEC §7). `datagen <dir> <name> gendate [yearEnd firstYear lastYear label flags]`, flags `exact,noym,nohide`. Not tested in Excel: a non-default `yearLabel` |
| Date table, mark an existing table | `DataModelWriter.MarkAsDateTable`; output diffed against `H2_date_table.xlsx` (matches except for list order and data size); **Excel-tested 2026-10-11: loads, date table works in the PivotTable, adding a hierarchy and a Sort by Column work** (the first attempt failed in `OnClone2` until `Ordinal` matched Excel, SPEC §9). Requires a hierarchy on the date table and a relationship on the date column (the only case seen in Excel) |
| Calculated columns, KPIs, perspectives | not started |
| Refresh with a real data source | not supported (placeholder connection; credentials prompt is expected) |
| CI (GitHub Actions), editorconfig, global.json, session hook | done |
| License | MIT (LICENSE, csproj) |
| NuGet packaging | **`0.1.0-preview.1` (2026-10-11) and `0.1.0-preview.2` (2026-10-20, tag on `830111f`: hidden columns, the date table, `AddDateTable`, distinct Ordinals) published on nuget.org**, id `PivotWright`. Metadata, README, snupkg symbols and SourceLink set up. `release.yml` publishes on `v*` tags via nuget.org trusted publishing (environment `nuget` with required reviewer, secret `NUGET_USER` = nuget.org username; no API key) |

Build and test: `dotnet build`, `dotnet run --project tests/PivotWright.Tests` — 42 tests, all pass. Nothing but the .NET 8 SDK is needed.

## 2. Immediate next steps, in order

1. **Calculated columns (next)**: needs Excel-written samples first (nothing is known about the file format). Ask Matthew for `H6_calculated_column.xlsx` (start from `recheck3_hierarchies_no_hidden` or `datagen <dir> <name>`; Sales: new column `=[Qty]*[Amount]` named Line Total) and `H7_calculated_column_related.xlsx` (Sales: new column `=RELATED(Products[Category])` named Category, made from the same base, not from H6). Questions the samples must answer: how the column is described in tbl.xml and the dimension, whether its values are stored (idf) or recomputed by Excel on load, and what counters and `Ordinal`s change. The library does not parse DAX, so the main design question is who supplies the values.
2. **Date table restrictions**: marking an existing table and the generated calendar are done and Excel-verified. Lift the restrictions in `ModelSerializer.ValidateDateTables` one at a time, each needing an Excel sample: date table without a hierarchy, without a relationship, in a model without hierarchies, several date tables, a hash-encoded date column. Excel cannot create a date table from scratch; it only marks an existing table.
3. **Releasing**: push a tag `vX.Y.Z` (stable) or `vX.Y.Z-preview.N` (prerelease) from `main`. The tag sets the package version; `release.yml` builds, tests, packs, waits for Matthew's approval of the `publish` job (check the commit and version on the run page), pushes to nuget.org, then creates the GitHub Release. Never re-tag a published version; use the next `-preview.N`. Released so far: `0.1.0-preview.1`, and `0.1.0-preview.2` (published on nuget.org 2026-10-20 from tag `v0.1.0-preview.2` on `830111f`; it contains the date table, `AddDateTable`, hidden columns and the Ordinal fix). Hidden columns and the date table are Excel-tested, so `v0.1.0` can be cut when Matthew agrees.
4. Other backlog features: KPIs, perspectives, the report-writer layer on top (the long-term aim: an Excel report modeller/writer), refresh against a real connection. Not tested in Excel: a non-default `yearLabel` of `AddDateTable`.

## 3. How testing works (important)

There is no Excel in the sandbox. The loop that produced every result:

1. Generate one or a few small files with the library (see `tools/datagen`).
2. Diff them against Excel-written reference files (tools below).
3. Give Matthew the file(s); he tests in Excel and reports **pass/fail per file**. The standard checks for a model with hierarchies: it loads; both hierarchies work in a PivotTable and respect sort columns; adding another hierarchy on each table works; adding a new Sort by Column works; saving and refreshing the PivotTable works.
4. Matthew can also take a generated file, make an edit in Excel (hide columns, mark a date table, add a hierarchy) and save it; **those Excel-written files are the ground truth**. Ask for them when a feature is missing. They live in `reference/excel-samples/`.

Matthew's preferences: give every test file a new, unique name on each iteration (e.g. `datetbl_v2_ordinal.xlsx`), never reuse a name, so he can refer to a specific version; start with "the most likely" file so he has fewer to test; do not send new files while he is mid-test ("hold on for feedback" — he said this once); practical, ready-to-use deliverables; he is precise about citation. Keep replies short.

**Run every Excel test in a fresh session** (close Excel completely, open only that file). Results from a session that had already opened related files were wrong more than once (a failing id set "passed" after Excel had cached a related sample). Failures are reliable, passes in a used session are not. Repeat a surprising result with a byte-identical copy under a new name.

When a failure is baffling, bisect with hybrids of an Excel sample and a generated model (`tools/repack`: `dump`, `idmap`, `swap`, `rename`, `ordinal`, ...), changing one thing per file, and consider asking Matthew for a model Excel builds itself (a plain workbook with tables added to the Data Model, or a text-file import through Power Pivot): the native samples showed what Excel writes and what it does not. Every sample in `reference/excel-samples/` except the native ones descends from our own output.

### Tools

- `python3 tools/python/extract_tree.py <xlsx> <outdir>` unpacks `item.data` into the logical file tree (needs `pip install xpress8`, from Hugoberry/xpress8-python; `pbixray` for `tools/verify_with_pbixray.py`).
- `tools/python/dd.py rootA rootB "<regex filter>" <maxlines>` diffs two extracted trees with GUIDs normalised (note: its key regex collapses versioned names but not `.tbl.xml`).
- `tools/python/xd.py rootA rootB <maxlines>` diffs the ASSL XML files (dim/det/cub/scr/prt) ignoring namespaces, version counters, timestamps and sizes; good for seeing what Excel changed.
- `dotnet run --project tools/datagen -- <outdir> <name>` writes the standard 3-table model (Products/Dates/Sales, 2 relationships, 2 hierarchies, sort columns, one measure).
- `dump_abf.py`, `xmldiff.py`: lower-level helpers.
- `dotnet run --project tools/datagen -- <outdir> <name> <mode>`: modes `date` (mark Dates), `gendate ...`, `hidden` (H1 ten columns), `h8`, `h9`, `plain` (no hierarchies), `twotable` (two tables, one relationship, no measure, no hierarchy).
- `tools/repack`: ABF repack/dump/hybrid builder used for the bisect (see its README).

## 4. Rules and pitfalls learned the hard way

- **Version counters must match file names and the backup-log versions** (`SPEC.md` §5). Changing counters without renaming files makes the model fail to load. Do not run counter-only experiments again.
- Excel's failure messages are generic. Never trust a decoder: pbixray and the test reader are more forgiving than Excel.
- The hierarchy profile applies only when the model has hierarchies; models without them keep the earlier profile, with one exception: every model gets distinct `Ordinal`s (measure groups 1..n, partitions and script 1). **`Ordinal` is a model requirement, not noise** (SPEC §8, §9): with all Ordinals 0 the engine tells the measure groups apart only by hash order of their ids, so Create Hierarchy and Sort by Column failed or passed depending on the random table ids. Check it first in any new profile; compare Ordinals of Excel's native files before assuming 0.
- Keep temporary env-var toggles out of committed code (a `PPW_*` experiment was removed once).
- After touching names, run all tests: several regexes pin file names (`Many_G\.2\.dim\.xml`, `Products_G\.5\.dim\.xml`, `Model\.18\.cub\.xml`, `U$...\.2\.tbl\.xml`).
- Small output only: when diffing Excel files, filter by object type; the full dd output is very large.

## 5. Source map

- `DataModelWriter.cs`, `TableWriter.cs`, `ColumnDefinition.cs`, `DataModelOptions.cs`, `MeasureFormat.cs`: public API.
- `Internal/ModelSerializer.cs`: the ABF assembly: file order, backup-log groups, version selection (`Info.CubeVersion` etc.), helper measure.
- `Internal/Metadata/AsslWriter.cs`: ASSL (database, cube, dimension, measure group, partition, script, data source view); `ModelInfo` with the version counters.
- `Internal/Metadata/TableXmlWriter.cs`: tbl.xml for main tables, `H$` column tables, `R$` relationship tables, `U$` hierarchy tables.
- `Internal/Model/`: `TableBuilder` (sort ranks, encoding), `ColumnBuilders`, `Hierarchies`, `Relationships`, `EncodedColumn` (`FileNames`, `HierarchyVersion = 2`).
- `Internal/VertiPaq`, `Internal/Abf`, `Internal/Crypto`, `Internal/Package`: encoding, container, crypto, xlsx parts.
- `tests/PivotWright.Tests`: own tiny runner, golden hashes of Excel files in `ExcelReferenceData.cs`, `ModelReader.cs` reads models back.

## 6. History in brief

- Early sessions: ABF writer, VertiPaq encoders, golden-file match with Excel, relationships (several undocumented details), measures, hierarchies.
- Sort by Column: loaded fine from the start, but Excel's Sort by Column / Create Hierarchy then failed on relationship models. About forty files were bisected; the four conditions in `SPEC.md` §5.3 fixed it (S22 passed).
- Hierarchies: a base model with hierarchies loaded but further edits failed (`OnClone2`). Diffing Excel's own save showed everything bumped to specific versions with renamed files; the hierarchy profile (§5.2) reproduces it. The last bug was a measure group still pointing at `.1.prt.xml`; after the fix all of Matthew's checks passed.
- H1/H2/H3 samples (hidden columns, date table, both) were received on 2026-10-10 and analysed. Marking a date table (part 1) was implemented on 2026-10-11; the other dimension versions also move up by one (Products 5→6, Dates 5→6, Sales 3→4), which the first analysis had missed.

## 7. Naming

Names considered: Pivotwright (chosen; repo `PivotWright`), Modelwright, Sheetwright, Reportwright, Vertiwright. Avoid "-foundry" (Palantir etc.) and FactLoom (repo exists). Check NuGet availability before publishing the first package; web searches found no exact conflicts.
