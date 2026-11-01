# Experiment: a PivotTable on the default cube and on a perspective (2026-11-01)

Not part of the library. Question: can a PivotTable be bound to a perspective, using what was learnt about connections (SPEC section 13)?

`tools/python/pivot_experiment.py <generated.xlsx> <out.xlsx> <variant> [perspective]` puts a hand-written, values-only PivotTable (Total Qty) into a generated workbook (`datagen ... persp3`). Its parts are adapted from Excel's own `H43_pivot_calc_value.xlsx` (same three tables). What an Excel-written PivotTable on the data model contains (H43, H5, H63): `xl/pivotCache/pivotCacheDefinition1.xml` with `<cacheSource type="external" connectionId="2"/>` (the connection `ThisWorkbookDataModel`, command `Model`), a `cacheHierarchies` list of every hierarchy and measure of the cube, `dimensions`, `measureGroups` and `maps`; `xl/pivotTables/pivotTable1.xml` with one `pivotHierarchy` per cache hierarchy; the `pivotCaches` element in `workbook.xml`. Neither the cache nor the table names a cube or a perspective: the cube comes from the connection, so a perspective can only be expressed by the connection's command and by what the cache lists.

| Variant | Files made | What it tries |
|---|---|---|
| `default` | `pivot_default_v1.xlsx` | control: PivotTable on `Model`, connection 2 unchanged |
| `conn` | `pivot_conn_v1.xlsx` | a second model connection with `command="3 - both"` and a cache that lists only the perspective's hierarchies and measure |
| `conn-full` | `pivot_conn-full_v1.xlsx` | as `conn`, but the cache lists the whole model (the connection alone decides) |
| `string` | `pivot_string_v1.xlsx` | as `conn`, with an explicit MSOLAP connection string `...;Data Source=$Embedded$;...;Cube=3 - both` (the advice found on the web) |

Earlier result (H68, H70): extra connections in `connections.xml` are kept by Excel but their command is rewritten to `Model` on save.

## Results

Order of testing: `default` first (if a hand-written PivotTable does not load, the others mean nothing), then the most likely perspective variant.

| Variant | Result |
|---|---|
| `default` v1 | opened cleanly, the PivotTable showed 3000, then Excel asked for a username and password. Cause (by comparison with H43): the script set `refreshOnLoad="1"` in the cache, so Excel refreshed on opening and reached for the placeholder data source's credentials |
| `default` v2 | `refreshOnLoad` removed (Excel's own cache does not have it); sent 2026-11-01, awaiting Excel |
