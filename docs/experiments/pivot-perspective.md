# Experiment: a PivotTable on the default cube and on a perspective (2026-11-01)

Not part of the library. Question: can a PivotTable be bound to a perspective, using what was learnt about connections (SPEC section 13)?

`tools/python/pivot_experiment.py <generated.xlsx> <out.xlsx> <variant> [perspective]` puts a hand-written, values-only PivotTable (Total Qty) into a generated workbook (`datagen ... persp3`). Its parts are adapted from Excel's own `H43_pivot_calc_value.xlsx` (same three tables). What an Excel-written PivotTable on the data model contains (H43, H5, H63): `xl/pivotCache/pivotCacheDefinition1.xml` with `<cacheSource type="external" connectionId="2"/>` (the connection `ThisWorkbookDataModel`, command `Model`), a `cacheHierarchies` list of every hierarchy and measure of the cube, `dimensions`, `measureGroups` and `maps`; `xl/pivotTables/pivotTable1.xml` with one `pivotHierarchy` per cache hierarchy; the `pivotCaches` element in `workbook.xml`. Neither the cache nor the table names a cube or a perspective: the cube comes from the connection, so a perspective can only be expressed by the connection's command and by what the cache lists.

| Variant | Files made | What it tries |
|---|---|---|
| `default` | `pivot_default_v1.xlsx` | control: PivotTable on `Model`, connection 2 unchanged |
| `conn` | `pivot_conn_v1.xlsx` | a second model connection with `command="3 - both"` and a cache that lists only the perspective's hierarchies and measure |
| `conn-full` | `pivot_conn-full_v1.xlsx` | as `conn`, but the cache lists the whole model (the connection alone decides) |
| `two` | `pivot_two_v1.xlsx` | two PivotTables in one workbook: PivotTable1 at A3 on `Model` (connection 2, whole-model cache), PivotTable2 at C3 on perspective `3 - both` (connection 3, trimmed cache); each with its own cache (ids 5 and 6) |
| `string` | `pivot_string_v1.xlsx` | as `conn`, with an explicit MSOLAP connection string `...;Data Source=$Embedded$;...;Cube=3 - both` (the advice found on the web) |

Earlier result (H68, H70): extra connections in `connections.xml` are kept by Excel but their command is rewritten to `Model` on save.

## Results

Order of testing: `default` first (if a hand-written PivotTable does not load, the others mean nothing), then the most likely perspective variant.

| Variant | Result |
|---|---|
| `default` v1 | opened cleanly, the PivotTable showed 3000, then Excel asked for a username and password. Cause (by comparison with H43): the script set `refreshOnLoad="1"` in the cache, so Excel refreshed on opening and reached for the placeholder data source's credentials |
| `default` v2 | `refreshOnLoad` removed (Excel's own cache does not have it); Excel: opens without a prompt, the PivotTable shows 3000, the Fields pane lists the full model. Right-click > Refresh asks for a username and password. The copy Excel saved afterwards is `reference/excel-samples/H71_pivot_default_saved_by_excel.xlsx`: Excel kept the hand-written cache definition unchanged (no difference at all) and the PivotTable except `cacheId` 5 became 0, so the hand-written PivotTable is valid. But the saved copy is **after the refresh attempt**, and the model in it is one edit further on: cube and measure groups 20 → 21 (with `info.21.xml`), Products dimension 6 → 7, the Sales partition's `DataVersion` 1 → 0. Settled by `conn` below: the same step appears when nothing was refreshed, so it comes from Excel saving a workbook that has a PivotTable, not from the refresh |

Size note: the generated files are 70 KB when this script repacks them (it compresses `item.data`), 270 KB from the library (stored uncompressed, 266,240 bytes, as in Excel's own files); Excel's re-save is 279 KB. The model part is the same length in all of them; the bytes differ because Excel re-encrypts and renames objects when it saves.
| `conn` v2 (`pivot_conn_v2.xlsx`, saved copy `reference/excel-samples/H72_pivot_conn_saved_by_excel.xlsx`) | Opens cleanly, A4 shows 3000, and **the Fields pane lists only the perspective's items** (Sales: Total Qty; Products: Product Hierarchy, ProductKey, Product). Excel kept the second connection (`ThisWorkbookDataModel - 3 - both`, id 3), kept the cache pointing at it (`connectionId="3"`) and kept the hand-written cache unchanged, **but rewrote its `command` from `3 - both` to `Model`**, as it did in H70. The Queries & Connections pane lists only `PivotWright`. The model in the saved copy stepped like H71 (cube 21, Products dimension 7) without any refresh |

## What this shows so far

- A PivotTable whose cache lists only a perspective's hierarchies and measures *looks* bound to that perspective: the Fields pane is filtered. The pane is built from the cache definition.
- The connection command does not survive: Excel normalises every model connection to `Model`. So nothing in the file ties the PivotTable to the perspective except the list in its cache. Whether a refresh keeps the filtering (rebuilds the cache from `Model`, which would show the whole model again) is not yet known; the first refresh of a generated file asks for credentials, see below.
- Excel saving a workbook that has a PivotTable steps the model one edit (cube 20 to 21, Products dimension 6 to 7, `info.21.xml`, the Sales partition's `DataVersion` 1 to 0). H71 and H72 are identical in that respect.
- Refresh asks for a username and password in the original files and also in `H72`, the copy Excel saved (tested 2026-11-01). The hypothesis that only unsaved generated files prompt is refuted. Every refresh in these hand-made PivotTable files reaches for the placeholder data source, so refresh cannot show whether a perspective filter survives; that stays untested. (Earlier checks of `perspective_three_v1` passed refresh with a PivotTable made in Excel; the difference between that and the hand-written PivotTable is not explained.)

## Two PivotTables, one on each (2026-11-01)

`pivot_two_v1.xlsx` (`pivot_experiment.py ... two "3 - both"`): requested to show that one workbook can hold a PivotTable on the default cube and another on a custom perspective. What to look for: both open, A4 and C4 show 3000, clicking each PivotTable shows its own Fields pane (the whole model for the first, only the perspective's items for the second). No refresh (it prompts for credentials). Result (2026-11-01, reported by Matthew): the file loads, both PivotTables show 3000, the PivotTable on the left (default cube) shows the whole model in its Fields pane and the one on the right shows only the perspective's subset. No saved copy was taken, so what Excel keeps of the two caches is not recorded. Nothing was refreshed.

## Conclusion of the experiment

One workbook can hold PivotTables on different views of the model: each PivotTable's cache lists the hierarchies and measures its Fields pane offers (the cache holds no data: `saveData="0"`, no records part, a few kilobytes). Writing a cache that lists only what a perspective shows gives a PivotTable that offers only that subset. Not shown: that the PivotTable is *bound* to the perspective in the engine. Excel rewrites the connection command to `Model`, so a refresh, which would rebuild the field list from the connection, may show the whole model again; refresh prompts for credentials in these files, so this could not be tested.

## Routes tried in the Excel UI to make a PivotTable on a perspective (no sample came of any)

All on `H66_perspective_hierarchies_manual.xlsx` (the Perspectives dialog lists the three perspectives, so Excel knows them):
- Insert PivotTable > Use an external data source > Existing Connections: the Tables tab lists "This Workbook Data Model" (3 tables) and the three placeholder tables, no perspective; the Show box offers All Tables, Connection Tables, Workbook Tables, none lists a perspective; the Connections tab lists only `PivotWright`.
- The PivotTable Fields pane of a PivotTable on the data model has no perspective picker (the Active/All tabs and a gear menu only).
- Data > Get Data > From Analysis Services with server name `$Embedded$`: "Unable to locate database server".
- Connection Properties of the `PivotWright` connection (the placeholder source): Connection type Model OLE DB Query, command type Table Collection, command text `"Products","Dates","Sales"`; editing it asks for credentials. A web summary suggested adding `Cube=<perspective>` to a connection string or a command text; it describes Analysis Services connections and did not apply here.
- `H67` (a PivotTable made by Excel on a perspective) was therefore never obtained.
- Power Pivot window > Home > PivotTable: reported by Matthew (2026-11-01) that Excel never offers a choice of perspective, there or anywhere else for a PivotTable on the workbook's data model. The screenshots of that Fields pane and its gear menu were not taken.

## Open questions (left open 2026-11-01)

1. **Why does refresh ask for credentials in these files?** The prompt is for the placeholder source (`Provider=Microsoft.ACE.OLEDB.12.0;Data Source=PivotWright;Persist Security Info=false`). `docs/research-refresh.md` (Excel treats a changed model as dirty and reprocesses it from source) explains the case after a model edit, but not these files: Refresh prompted in the files as generated (`pivot_default_v2`) and in the copy Excel saved (H72, H71), all with hand-written PivotTables, while a PivotTable made in Excel on `perspective_three_v1` refreshed without a prompt. Decisive test, not run: insert a PivotTable from the data model in H72 and refresh it.
2. **Does a refresh keep a perspective-only cache?** Untestable while refresh prompts. If Excel rebuilds the field list from the connection (`Model`), the whole model comes back.
3. **What does Excel keep of two caches?** `pivot_two_v1.xlsx` was not saved by Excel.
4. **Is there any way to make Excel itself bind a PivotTable to a perspective?** Matthew reports that Excel offers no choice of perspective for a PivotTable on the workbook's data model (all routes above, including the Power Pivot window). A PivotTable with a perspective-only cache is therefore something Excel does not write itself, and the experiment files here are the only ones of their kind; there is no Excel-written sample (H67) to compare with, and none can be made through the UI.
5. Variants never run: `conn-full`, `string`, and a connection without the `x15:connection model="1"` marker.

## Tools made for this

- `tools/datagen` modes `persp1`, `persp3`, `perspH` (H64, H65, H66 on the standard model).
- `tools/python/cmp_trees.py`: compares two extracted model trees keeping the version counters (xd.py ignores them).
- `tools/python/add_perspective_connections.py`: H68.
- `tools/python/pivot_experiment.py`: the PivotTable variants.
- Samples: `reference/excel-samples/H64`, `H65`, `H66`, `H70`, `H71`, `H72`; files given to Excel: `reference/experiments/`.
