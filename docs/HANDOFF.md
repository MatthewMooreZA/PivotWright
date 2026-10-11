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
| Date table | analysed from Excel samples (`SPEC.md` §9); **not implemented** |
| Calculated columns, KPIs, perspectives | not started |
| Refresh with a real data source | not supported (placeholder connection; credentials prompt is expected) |
| CI (GitHub Actions), editorconfig, global.json, session hook | done |
| License | MIT (LICENSE, csproj) |
| NuGet packaging | metadata, README, snupkg symbols and SourceLink set up; `release.yml` publishes on `v*` tags via nuget.org trusted publishing (environment `nuget`, secret `NUGET_USER` = nuget.org username; no API key). Not yet published; check id `PivotWright` on nuget.org |

Build and test: `dotnet build`, `dotnet run --project tests/PivotWright.Tests` — 29 tests, all pass at handoff. Nothing but the .NET 8 SDK is needed.

## 2. Immediate next steps, in order

1. **Date table** (`MarkAsDateTable`). Everything needed is in `SPEC.md` §9. Version counters must match file names (§5), so this is a profile change touching `FileNames` (`Internal/Model/EncodedColumn.cs`), `TableXmlWriter`, `AsslWriter` (`ModelInfo` has the version helpers: `CubeVersion`, `ScriptVersion`, `PartitionVersion`, `DimensionVersion`), and `ModelSerializer` (file naming and backup-log groups). Generate one most-likely file, diff against `reference/excel-samples/H2_date_table.xlsx` with the tools below, give it to Matthew.
2. **Hidden columns**: generate a file with hidden columns (and one with hidden columns plus hierarchies), ask Matthew to check that it loads, PivotTable fields hide, and that Create Hierarchy / Sort by Column / marking a date table still work afterwards. Use `H1` as the reference. Whether Excel's counter bumps (cube 24, measure-group structure 9) are needed is unknown; add them only if a test file fails without them.
3. **First NuGet release**: confirm the id `PivotWright` is free on nuget.org (unreachable from the sandbox), make the repo public, create the `nuget` environment (required reviewer), register the trusted publishing policy on nuget.org (workflow `release.yml`, environment `nuget`), add the `NUGET_USER` repo secret, then push a tag `vX.Y.Z` (stable) or `vX.Y.Z-preview.N` (prerelease); the tag sets the package version and `release.yml` also creates the GitHub Release. The csproj `<Version>` only affects local builds. Rename, license, CI, `CLAUDE.md` and packaging setup are done.
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
- H1/H2/H3 samples (hidden columns, date table, both) were received on 2026-10-10 and analysed; implementation is the next task.

## 7. Naming

Names considered: Pivotwright (chosen; repo `PivotWright`), Modelwright, Sheetwright, Reportwright, Vertiwright. Avoid "-foundry" (Palantir etc.) and FactLoom (repo exists). Check NuGet availability before publishing the first package; web searches found no exact conflicts.
