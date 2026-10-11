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
| Hidden columns | `ColumnDefinition.IsHidden` writes the same XML Excel writes (diffed against `H1`); **not yet Excel-tested as a generated file** |
| Date table, mark an existing table | `DataModelWriter.MarkAsDateTable`; output diffed against `H2_date_table.xlsx` (matches except for list order and data size); **not yet Excel-tested as a generated file**. Requires a hierarchy on the date table and a relationship on the date column (the only case seen in Excel) |
| Date table, Excel creates it from scratch (Date Table > New) | **not started**; needs an Excel-written sample (see §2) |
| Calculated columns, KPIs, perspectives | not started |
| Refresh with a real data source | not supported (placeholder connection; credentials prompt is expected) |
| CI (GitHub Actions), editorconfig, global.json, session hook | done |
| License | MIT (LICENSE, csproj) |
| NuGet packaging | **`0.1.0-preview.1` published on nuget.org (2026-10-11)**, id `PivotWright`. Metadata, README, snupkg symbols and SourceLink set up. `release.yml` publishes on `v*` tags via nuget.org trusted publishing (environment `nuget` with required reviewer, secret `NUGET_USER` = nuget.org username; no API key) |

Build and test: `dotnet build`, `dotnet run --project tests/PivotWright.Tests` — 31 tests, all pass. Nothing but the .NET 8 SDK is needed.

## 2. Immediate next steps, in order

1. **Date table, part 1 (mark an existing table)**: implemented (`SPEC.md` §9). Send Matthew the file from `dotnet run --project tools/datagen -- <dir> datetbl date` and ask for the standard checks (loads; Dates marked in the Design tab; time intelligence works; add a hierarchy, add Sort by Column, save and refresh). If it fails, bisect against `H2`. Then lift the restrictions in `ModelSerializer.ValidateDateTables` one at a time, each needing an Excel sample: date table without a hierarchy, without a relationship, in a model without hierarchies, several date tables, a hash-encoded date column.
   **Date table, part 2 (Excel creates the table from scratch)**: ask Matthew for a sample: open `V2b`, Power Pivot > Design > Date Table > New, save as `H4_date_table_new.xlsx`; ideally also one made in an empty workbook. It most likely adds a calculated table with calculated columns (not implemented), so diff it with `extract_tree.py`/`xd.py` before designing.
2. **Hidden columns**: generate a file with hidden columns (and one with hidden columns plus hierarchies), ask Matthew to check that it loads, PivotTable fields hide, and that Create Hierarchy / Sort by Column / marking a date table still work afterwards. Use `H1` as the reference. Whether Excel's counter bumps (cube 24, measure-group structure 9) are needed is unknown; add them only if a test file fails without them.
3. **Releasing**: push a tag `vX.Y.Z` (stable) or `vX.Y.Z-preview.N` (prerelease) from `main`. The tag sets the package version; `release.yml` builds, tests, packs, waits for Matthew's approval of the `publish` job (check the commit and version on the run page), pushes to nuget.org, then creates the GitHub Release. Never re-tag a published version (nuget.org cannot replace one); use the next `-preview.N`. Stay on previews until hidden columns and the date table are Excel-tested, then cut `v0.1.0`. Rename, license, CI, `CLAUDE.md` and packaging setup are done.
4. Backlog features: calculated columns, KPIs, perspectives, the report-writer layer on top (the long-term aim: an Excel report modeller/writer), refresh against a real connection.

## 3. How testing works (important)

There is no Excel in the sandbox. The loop that produced every result:

1. Generate one or a few small files with the library (see `tools/datagen`).
2. Diff them against Excel-written reference files (tools below).
3. Give Matthew the file(s); he tests in Excel and reports **pass/fail per file**. The standard checks for a model with hierarchies: it loads; both hierarchies work in a PivotTable and respect sort columns; adding another hierarchy on each table works; adding a new Sort by Column works; saving and refreshing the PivotTable works.
4. Matthew can also take a generated file, make an edit in Excel (hide columns, mark a date table, add a hierarchy) and save it; **those Excel-written files are the ground truth**. Ask for them when a feature is missing. They live in `reference/excel-samples/`.

Matthew's preferences: start with "the most likely" file so he has fewer to test; do not send new files while he is mid-test ("hold on for feedback" — he said this once); practical, ready-to-use deliverables; he is precise about citation. Keep replies short.

### Tools

- `python3 tools/python/extract_tree.py <xlsx> <outdir>` unpacks `item.data` into the logical file tree (needs `pip install xpress8`, from Hugoberry/xpress8-python; `pbixray` for `tools/verify_with_pbixray.py`).
- `tools/python/dd.py rootA rootB "<regex filter>" <maxlines>` diffs two extracted trees with GUIDs normalised (note: its key regex collapses versioned names but not `.tbl.xml`).
- `tools/python/xd.py rootA rootB <maxlines>` diffs the ASSL XML files (dim/det/cub/scr/prt) ignoring namespaces, version counters, timestamps and sizes; good for seeing what Excel changed.
- `dotnet run --project tools/datagen -- <outdir> <name>` writes the standard 3-table model (Products/Dates/Sales, 2 relationships, 2 hierarchies, sort columns, one measure).
- `dump_abf.py`, `xmldiff.py`: lower-level helpers.

## 4. Rules and pitfalls learned the hard way

- **Version counters must match file names and the backup-log versions** (`SPEC.md` §5). Changing counters without renaming files makes the model fail to load. Do not run counter-only experiments again.
- Excel's failure messages are generic. Never trust a decoder: pbixray and the test reader are more forgiving than Excel.
- The hierarchy profile applies only when the model has hierarchies; models without them must stay byte-for-byte as before (tested by `RelationshipDefaults` and others).
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
