# Files used in the perspective and PivotTable experiments (2026-11-01)

Hand-made or library-made, **not Excel-written** (the Excel-written files are in `../excel-samples/`: H64, H65, H66, H70, H71, H72). They are kept because they are the exact bytes Excel was given; the scripts that build them are in `tools/python/` and `tools/datagen`, but ids and timestamps differ on every build. Notes and results: `docs/experiments/pivot-perspective.md`, `docs/SPEC.md` section 13.

| File | What it is | How it was made | Excel result |
|---|---|---|---|
| `perspective_base_v1.xlsx` | the standard model, no perspective; the base H64 to H66 were made from | `datagen <dir> perspective_base_v1` | loads |
| `perspective_three_v1.xlsx` | the three perspectives of H65, written by the library | `datagen <dir> perspective_three_v1 persp3` | passes: loads, the Perspectives dialog shows all three as made, editing one, a new hierarchy, a Sort by Column, save and PivotTable refresh (with a PivotTable made in Excel) |
| `H68_secondary_connections_v1.xlsx` | H66 plus one connection per perspective, command = the perspective name | `tools/python/add_perspective_connections.py` | opens cleanly; Excel rewrote every command to `Model` on saving (H70) |
| `pivot_default_v1.xlsx` | `perspective_three_v1` plus a hand-written PivotTable on `Model`, **with `refreshOnLoad`** | an earlier version of `pivot_experiment.py ... default` | opens, shows 3000, then asks for credentials (the attribute made Excel refresh on opening) |
| `pivot_default_v2.xlsx` | as v1 without `refreshOnLoad` | `pivot_experiment.py ... default` | opens, shows 3000, the Fields pane lists the whole model; Refresh asks for credentials; Excel kept the cache unchanged (H71) |
| `pivot_conn_v2.xlsx` | a PivotTable on perspective `3 - both`: a second connection and a cache that lists only that perspective's items | `pivot_experiment.py ... conn "3 - both"` | opens, shows 3000, the Fields pane lists only the perspective's items; Excel kept the connection and the cache and rewrote the command to `Model` (H72) |
| `pivot_two_v1.xlsx` | two PivotTables: A3 on `Model`, C3 on `3 - both` | `pivot_experiment.py ... two "3 - both"` | opens, both show 3000, the left one lists the whole model, the right one a subset; no saved copy taken |

Built but never sent: the variants `conn-full` (a second connection with a whole-model cache) and `string` (an explicit MSOLAP connection string with `Cube=<perspective>`, the advice found on the web); `pivot_experiment.py` makes them.
