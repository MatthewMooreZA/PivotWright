# Research: making Excel calculate calculated columns on open or refresh (todo item 5)

Status: research only, 2026-10-29. Nothing here is Excel-tested yet. Each statement is marked **[E]** (evidenced in this repo: samples, generated files, or Matthew's reports), **[D]** (public documentation, seen via search snippets; the doc sites could not be fetched from this sandbox) or **[S]** (speculation or recollection, unverified).

## 1. Summary

- **Where the calculation mode is stored.** [E] It is stored outside the model, in a Power Pivot ("Gemini") customXml part, `http://gemini/pivotcustomization/ManualCalcMode`. Its value is `True` in H10, H15 and H21 and `False` in H6, H7, H37, H2, H5 and S1x. PivotWright writes no customXml at all. The engine itself has no calculation mode. The Power Pivot client decides whether to send a recalc after a change.
- **No workbook-level flag in any sample asks for a refresh or recalc on open.** [E] None of `refreshOnLoad`, `fullCalcOnLoad`, `calcMode` or `excludeFromRefreshAll` appears anywhere. `calcPr` is just `calcId="191029"`, the same in Excel's files and ours. Our `connections.xml` equals Excel's apart from the `xr16:uid` attributes. Between H6 (calculated) and H10 (unprocessed), the parts outside the model differ only in the window position and the customXml values (`ManualCalcMode`, the `ErrorCache` timestamp, a grid-focus flag).
- **Inside the model, H10 differs from H6 only in the column itself and the counters.** [E] The differences are the attribute's `ddl200:ProcessingState` and `ddl300:AttributeHierarchyProcessingState` (Unprocessed), the tbl.xml intrinsic hierarchy (`IsProcessed` false, -1 values, no `TableStore`), and no dictionary or `H$` table. No dimension, measure group, partition or cube carries a "needs recalculation" state. The engine on load sees one unprocessed column and nothing else.
- **New finding: H7, the processed form, sits on the straight-line counter fit.** [E] H7 is the calculated `RELATED` column, added in automatic mode. Its counters are cube 23 (structure 16), measure groups 23/9/cache 2, partitions and script 10, Sales dimension 6 (1, 2, 2, 2), table version 2, data file prefix 2, `H$` table at version 0. These are exactly `CalculatedPlan`'s n = 1 formulas. H10 is the outlier (cube 25) because of its extra manual-mode step. So writing the calculated form needs no new counter profile, only per-column content. This makes the fallback in section 6 much cheaper than assumed.
- **Most likely explanation of "works as a value, then everywhere".** [S, testable] Dropping a column into Values creates an implicit measure ("Sum of Line Total"), which is a model schema change. In automatic mode Power Pivot follows every schema change with a recalculation of the database (ProcessRecalc semantics), so the column gets calculated. A plain PivotTable refresh changes nothing in the schema, so nothing is calculated. Nothing calculates on open either, because nothing changes. Diagnostic D2 below tests this in one minute. If it holds, the fix is to cause a recalc command or a schema change automatically. No file flag found so far does that.
- **Best candidate.** A ProcessRecalc command sent through `Workbook.Model.DataModelConnection.ModelConnection.ADOConnection`, first by hand in the VBA Immediate window (no new file format needed). If that works, run it from `Workbook_Open` in an .xlsm. The best macro-free candidate is a linked "trigger" table refreshed on load, with the placeholder source excluded from Refresh All. It needs one Excel-made sample and is speculative.
- **Fallback is viable.** Either the library evaluates a DAX subset, or callers pass their own values. Either way the column is written in H7's calculated form under the counters `CalculatedPlan` already produces. Five Excel samples (H6, H7, H37, H38, H39) contain Excel-computed values to test against.

## 2. What is known from the repo [E]

- Calculated columns are written formula-only, as in H10 (SPEC §10). Excel computes the values when the column is used as a PivotTable value. Until then the column cannot be in Rows, Columns or a filter, or be a level of a hierarchy that is in use (`calchier_v1`). Excel's own H15 behaves the same way.
- Right-click > Refresh on the PivotTable does not calculate it. Refresh All prompts for credentials, because the placeholder source is `Provider=Microsoft.ACE.OLEDB.12.0;Data Source=PivotWright` (`DataModelOptions.ConnectionString`).
- After a model change in Excel, right-click > Refresh also prompts for credentials (`calccol_eight_v1`). On the unchanged file it does not. So in a session, Excel tracks whether the model is "dirty" and then refreshes it from source instead of just re-querying.
- Power Pivot's Calculate Now works (SPEC §10, "Uncalculated limits").
- HANDOFF says the automatic-mode generated file "behaves as H10 did". Our files have no `ManualCalcMode` part, so they presumably run in the default automatic mode, and they still do not calculate on open. **So automatic mode alone does not trigger a recalc on open.**
- The README note for H21 says manual mode "is not remembered", but the part does record `True`. Either Power Pivot reads it only when its window or add-in loads, or the note is wrong. This does not matter for us unless experiment X6 shows otherwise.
- `H17_text_calculated_after_ours.xlsx`, named in HANDOFF item 5, is **not in `reference/excel-samples/`** and not anywhere on disk. Ask Matthew for it again, or produce an equivalent as a side effect of X1.

## 3. Workbook-level parts inspected

Inspected: H6, H7, H10, H15, H21, H37, H2, H5, S1x, V2b, and a fresh `datagen ... calc` file. Only H5 has a PivotTable. The H6–H39 samples were saved without one.

| Part / flag | Excel samples | PivotWright | Could it trigger a model refresh or recalc? |
|---|---|---|---|
| `workbook.xml` `<calcPr>` | `calcId="191029"` only | same | `fullCalcOnLoad="1"` recalculates worksheet formulas only. It could re-query CUBE functions, but a query is not a schema change [S]. `calcMode` is the worksheet mode, not Power Pivot's. Unlikely. |
| `workbook.xml` `x15:dataModel` (`modelTables`, `modelRelationships`) | present | same | No refresh semantics. |
| `connections.xml` id 1 `PivotWright`, `type="100"`, `x15:connection id=<data source id>` | `refreshedVersion="8" minRefreshableVersion="5"` | same | Base attributes `refreshOnLoad`, `background`, `saveData`, `credentials` [D, ECMA-376 CT_Connection]. On the x15 extension, `excludeFromRefreshAll` [S, my recollection of [MS-XLSX] §2.6.x, not re-checked; its object-model counterpart `WorkbookConnection.RefreshWithRefreshAll` is documented [D]]. Refreshing this connection hits the placeholder source and prompts for credentials. |
| `connections.xml` id 2 `ThisWorkbookDataModel`, `type="5"`, `keepAlive="1" background="1"`, `dbPr command="Model" commandType="1"`, `x15:connection id="" model="1"` | as written | same | `refreshOnLoad="1"` here is untested. I expect it to equal "refresh the model", meaning a source refresh and a credentials prompt [S]. |
| `pivotCacheDefinition` (H5 only) | `saveData="0" backgroundQuery="1"`, no `refreshOnLoad` | library writes no pivots | `refreshOnLoad="1"` is a cache refresh on open. It should behave like right-click > Refresh, which is known not to calculate [E for Refresh, S for on-load]. |
| `pivotTableDefinition` | no refresh flag (the flag is on the cache) | none | none |
| customXml Gemini parts (16–20 per Excel file) | `ManualCalcMode` True/False, `ShowHidden`, `ShowImplicitMeasures`, `SandboxNonEmpty`=1, `IsSandboxEmbedded`=yes, `PowerPivotVersion`=2015.130.1606.47, `LinkedTableUpdateMode`=True, `RelationshipAutoDetectionEnabled`, `ErrorCache` (`LastProcessedTime`), UI layout parts, per-PivotTable `Settings` parts (H5) | **none** | `ManualCalcMode` is the only calculation-related item. `LinkedTableUpdateMode` governs linked (worksheet) tables (see X3). |
| `xl/model/item.data` | ABF | ABF | No processing state above column level (section 1). |

## 4. How processing works (background)

- **ProcessRecalc** [D, Microsoft's "processing tabular models" guide and KB2880071]: for every table in the database it recalculates calculated columns and rebuilds relationships and user hierarchies. It does not read the data source. ProcessFull and ProcessDefault include it, and ProcessData leaves the database needing it. KB2880071 quotes the engine error for a column in this state: "... does not hold any data because it needs to be recalculated or refreshed". That is probably what Matthew sees when the column is dropped into Rows (unconfirmed: ask for the exact text).
- **Power Pivot calculation options** [D, "Recalculate Formulas in Power Pivot"]: automatic mode is the default and recalculates "as required". Triggers listed include adding or changing measures and calculated columns, changing relationships, and refreshing data. In manual mode, Calculate Now runs a full update, and the docs advise switching back to automatic before publishing. Nothing in that material describes a recalc on open.
- **Refresh vs. recalc** [D]: refresh reprocesses from the source and is followed by recalculation. A refresh needs the source, which is what produces our credentials prompt.
- **Unprocessed objects in 1103 metadata** [E/S]: Excel stores state per attribute (`ddl200:ProcessingState`, `ddl300:AttributeHierarchyProcessingState`) and per user hierarchy (`ddl300:ProcessingState`; StructureType Unknown in H31B). Partitions and dimensions have no persisted state element. Analysis Services never processes on its own (no proactive caching against an embedded source) [S].
- **PivotTable refresh on a model cache** [E]: on an unchanged model it only re-queries. After an in-session model change it refreshes the model from source.
- **VBA object model** [D]: `Workbook.Model`, `Model.DataModelConnection` (the `ThisWorkbookDataModel` WorkbookConnection), `ModelConnection.ADOConnection` (documented as giving add-ins such as Power View a direct ADO connection to the engine), `Model.Refresh` and `ModelTable.Refresh` (these refresh from source and would prompt here). SQLBI documented injecting XMLA commands on that connection to process the embedded engine. They called it unsupported and advised backing up the workbook [D, search snippet only].

## 5. Ranked experiments

General rules (HANDOFF §3): a fresh Excel session, unique file names, the most likely file first, and wait for the result before sending more.

Each "file" is `dotnet run --project tools/datagen -- <dir> <name> calc` (the standard Sales `Line Total` model), plus a post-processing step where stated. Post-processing edits only package parts, never `item.data`, so the version-counter rules are unaffected. A stdlib-only helper is enough:

```python
# python3 -I patch.py in.xlsx out.xlsx part 'old' 'new'  (string replace in one zip part)
import sys, zipfile
src, dst, part, old, new = sys.argv[1:6]
with zipfile.ZipFile(src) as zi, zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zo:
    for it in zi.infolist():
        b = zi.read(it.filename)
        if it.filename == part:
            t = b.decode("utf-8"); assert old in t; b = t.replace(old, new, 1).encode("utf-8")
        zo.writestr(it, b)
```

Every check below ends with the same **standard check**: Insert > PivotTable > From Data Model. Drag `Line Total` straight into **Rows**, without putting anything in Values first. Success means the distinct values are listed.

### Step 0: diagnostics (all files are plain `datagen ... calc` output under new names; about 1 minute each)

These decide which branch is worth building. They cost nothing to generate.

- **D1 `refresh_diag_d1.xlsx`**: put the explicit measure **Total Qty** in Values. It already exists in the model, so no schema change happens. Then drag `Line Total` into Rows. Expected: it fails. Note the exact error text.
- **D2 `refresh_diag_d2.xlsx`**: put the column **Qty** in Values. This creates an implicit measure (a schema change) that does not touch `Line Total`. Then drag `Line Total` into Rows. **If it works, any model change in automatic mode recalculates the whole model.** That confirms the theory in section 1 and gives users a simpler workaround than the current README text.
- **D3 `refresh_diag_d3.xlsx`**: Data > Manage Data Model (open the Power Pivot window). Does `Line Total` show values or blanks with a warning? Close the window, then run the standard check. This shows whether merely loading the Power Pivot client in automatic mode recalculates.
- **D4 `refresh_diag_meas_v1.xlsx`**: needs a one-line datagen addition, `s.AddMeasure("Line Total Sum", "SUM(Sales[Line Total])", ...)`. Put only `Line Total Sum` in Values. Does it show R90500, blank, or an error? Then run the standard check. If it shows values **and** Rows then works, queries alone calculate, and the query-on-open routes (X5) become worth trying. Otherwise drop X5.

### X1 (rank 1): ProcessRecalc through the model's ADO connection, by hand, then from `Workbook_Open`

- **Why it might work.** [S, plausible] ProcessRecalc is exactly the operation needed and never touches the source [D]. The ADO connection reaches the in-process engine [D], and SQLBI processed the embedded engine with XMLA this way [D].
- **File.** `refresh_vba_x1.xlsx` (plain `calc`). Excel steps: open it, press Alt+F11, press Ctrl+G, and run one line at a time:

```
Set c = ThisWorkbook.Model.DataModelConnection.ModelConnection.ADOConnection
Set r = c.Execute("SELECT * FROM $SYSTEM.DBSCHEMA_CATALOGS"): For i = 0 To r.Fields.Count - 1: Debug.Print r.Fields(i).Name, r.Fields(i).Value: Next
c.Execute "<Process xmlns=""http://schemas.microsoft.com/analysisservices/2003/engine""><Type>ProcessRecalc</Type><Object><DatabaseID>PUT_ID_HERE</DatabaseID></Object></Process>"
```

  Use the `DATABASE_ID` printed by the second line. If there is none, use the 20-hex id of the file's `<id>.0.db`, which we send with the file. Then close VBA and run the standard check.
- **If the first line errors.** The model may not be loaded yet. Insert an empty PivotTable from the Data Model first, without dragging anything.
- **If the XMLA errors.** Retry with `<Batch ...><Process>...</Process></Batch>`, or with `ProcessDefault` in place of `ProcessRecalc`. ProcessDefault processes only what is unprocessed. If it decides to read the source it will fail on credentials, which is harmless.
- **Expected sign.** `Line Total` goes to Rows straight away, with no credentials prompt.
- **Bonus.** Save the file as `refresh_vba_x1_saved.xlsx` and send it back. It is Excel's own calculated form of our file, replacing the missing H17.
- **If it works.** Next is `refresh_vba_x1b.xlsm`: Matthew pastes `Sub Workbook_Open()` with the same code (database id from the DMV) into ThisWorkbook, saves, closes and reopens. Success means the standard check passes with no manual step. A library feature would then ship a prebuilt `vbaProject.bin` (from Matthew's file) and write .xlsm.
- **Costs.** Macro security prompts, an .xlsm extension, and an unsupported route. Offer it as an opt-in.

### X2 (rank 2): exclude the placeholder source from Refresh All, then Refresh All

- **Why it might work.** [S] Refresh All would then refresh only `ThisWorkbookDataModel` and the pivots, with no source and no prompt. If Excel follows a model "refresh" with a recalc, the column is calculated. Even if it does not, this removes the misleading credentials prompt, which is worth having on its own.
- **No-file pre-test, on `refresh_xl_x2a.xlsx` (plain `calc`).** Data > Queries & Connections > Connections. Right-click PivotWright > Properties > Usage, untick "Refresh this connection on Refresh All", then OK. Data > Refresh All. Note whether there is a prompt, then run the standard check. If the box is greyed out, report that, and the file version decides it.
- **File `refresh_excl_x2b.xlsx`.** Patch `xl/connections.xml`: `<x15:connection id="<ds id>"/>` becomes `<x15:connection id="<ds id>" excludeFromRefreshAll="1"/>` (connection 1 only). Run the same check. If the attribute name is wrong, Excel will report a repair. Then ask Matthew to save the pre-test file and read off the attribute Excel writes.
- **Expected signs.** No prompt is the minimum. A calculated column is the goal.

### X3 (rank 3): a linked "trigger" table refreshed on open (macro-free; needs one Excel-made sample)

- **Why it might work.** [S] A linked table, meaning an Excel table on a sheet added to the Data Model, refreshes from the workbook with no credentials. Refreshing any table makes Power Pivot recalculate in automatic mode [D: refresh is followed by recalc], and ProcessRecalc is database-wide [D]. Its worksheet connection could carry `refreshOnLoad="1"`. Combined with X2's exclusion, opening the file would refresh a one-row table and calculate every calculated column.
- **Sample request `S_trigger_manual.xlsx`.** Open `refresh_trig_base.xlsx` (plain `calc`). In Power Pivot: Design > Calculation Options > Manual, so the column stays unprocessed. On a new sheet, enter one cell, Insert > Table, name it `Trigger`, then Power Pivot > Add to Data Model. Save.
- **Then we send `refresh_trig_x3.xlsx`.** It is Matthew's sample with the `ManualCalcMode` part set to `False`, `refreshOnLoad="1"` on the `Trigger` worksheet connection, and `excludeFromRefreshAll` (X2) on PivotWright.
- **Check.** Open it. There should be no prompt, then run the standard check. Second check: Refresh All alone (`refresh_trig_x3b.xlsx` without `refreshOnLoad`).
- **Cost if it works.** The library must learn to write a linked table: no data source and plain ids, like Excel's native model. That is real work, plus a hidden sheet. Rank it after X1 and X2 because it depends on two speculative steps.

### X4 (rank 4): `refreshOnLoad="1"` on `ThisWorkbookDataModel`

- **File `refresh_onload_x4.xlsx`.** Patch connection 2 in `connections.xml` by adding `refreshOnLoad="1"` after `background="1"`.
- **Check.** Open it. Is there a prompt? Then run the standard check.
- **Expectation.** Probably a credentials prompt on open, the same as Refresh All [S]. Worth one file only together with X2's exclusion (`refresh_onload_x4b.xlsx`, both patches), because "refresh the model" with every source excluded might reduce to a recalc.

### X5 (rank 5, only if D4 showed that queries calculate): query on open

- `refresh_cube_x5.xlsx`: D4's model, with `Sheet1!A1` set to `=CUBEVALUE("ThisWorkbookDataModel","[Measures].[Line Total Sum]")` and `<calcPr calcId="191029" fullCalcOnLoad="1"/>`.
- **Check.** Does A1 show 90500 on open? Then run the standard check.
- A pre-built PivotTable with `refreshOnLoad` belongs to the same family, but writing pivot caches is a separate feature. Do not build it for this.

### X6 (rank 6, cheap, low odds): Excel's Gemini parts with `ManualCalcMode=False`

- **File `refresh_gemini_x6.xlsx`.** Add the table-independent customXml parts copied from H6 (`ManualCalcMode`=False, `SandboxNonEmpty`, `IsSandboxEmbedded`, `PowerPivotVersion`, `LinkedTableUpdateMode`, `RelationshipAutoDetectionEnabled`, `ErrorCache`), with their `itemProps` and rels and `[Content_Types]` entries.
- **Check.** Run the standard check.
- **Odds.** The odds are low, because our files already behave as automatic mode [E]. It is included only because it is the one stored calculation setting, and it costs nothing.

**Not proposed.** Faking processing states in `item.data` (marking the column Processed without values, or marking tables unprocessed to force a refresh). Either breaks the counters and content rules of SPEC §5/§10, or makes Excel want the source, which brings the credentials prompt back.

## 6. Fallback: write the calculated form

- **The evidence makes this cheap on the file side** [E]. H7, a calculated column added in automatic mode, has the straight-line counters already in `CalculatedPlan`. The column content is a normal data column: `.idf`, `.dictionary`, an `H$<table>$<col>.0.tbl.xml` with `POS_TO_ID`/`ID_TO_POS`, both processing states `Processed`, data file prefix = table version. It keeps the `ExpressionBinding`, `ColumnFlags` 8, `Settings` 3073 and `InferredDatatype` already written. The encoders for that content already exist (TableBuilder, ColumnBuilders).
- **First Excel test.** `calcvalues_x7.xlsx`: one currency column on Sales, values computed in the test harness, written in H7's form. The check is the standard check plus the usual hierarchy, Sort by Column, save and refresh checks. It needs an internal code path. Keep it on a branch with no env-var toggle (CLAUDE.md).
- **Two ways to get the values.**
  1. **Caller supplies the values**, for example `AddCalculatedColumn(name, dax, type, values)` or a `Func<row, object>`, with no DAX evaluation. This is the simplest option. The risk is that the values and the formula disagree until Excel next recalculates that column, which happens only if something it depends on changes, so wrong values can persist silently. Document it as "you assert these are the formula's results". It reverses the earlier "callers never supply values" decision, so it needs API-GUIDELINES review.
  2. **The library evaluates a DAX subset**: column references in the same row, `RELATED` (the relationship index already exists), numeric and text constants, `+ - * /`, comparisons, `&`, `IF`, `BLANK()`, `DIVIDE`, `AND`/`OR`/`NOT`, `&&`/`||`. Anything else stays formula-only, as today.
- **DAX semantics a subset must get right** [S, from DAX knowledge; to be pinned with tests]:
  - BLANK + n = n, but BLANK * n = BLANK.
  - n / 0 = ±Infinity and 0/0 = NaN. Use `DIVIDE` for BLANK.
  - BLANK = 0 and BLANK = "" are TRUE.
  - Text comparison is case-insensitive.
  - `&` turns numbers into text using DAX's own formatting, not the culture's.
  - Currency is a fixed four-decimal type, so check the rounding of Currency*Currency.
  - Integer / integer yields Double.
  - Date + n adds days and gives DateTime.
  - Booleans in arithmetic become 1/0.
  - Type inference has to match the declared `ColumnDataType`.
- **Free oracles.** H6 (`[Qty]*[Amount]`), H7 (`RELATED`), H37 (nine formatted columns), H38 (date columns) and H39 (a band and order) hold Excel's computed values for known formulas on the known `calcbase_v1` data. Unit tests can compare the evaluator against them byte for byte, because the encoding is deterministic.
- **Assessment.** Option 2 for that subset is moderate work (a parser and evaluator of a few hundred lines plus tests) with clear oracles. Combined with formula-only columns for everything else, it removes the limitation for the label-type columns that matter most (`RELATED`, text concatenation, simple bands with `IF`). It does not remove the limitation for arbitrary DAX. Recommendation: run step 0 and X1/X2 first. If no macro-free trigger exists, implement option 2, keeping option 1 as an advanced escape hatch.

## 7. Suggested order to send

1. Step 0, `refresh_diag_d2.xlsx` (D2): it answers the core question with zero risk.
2. `refresh_vba_x1.xlsx` (X1, Immediate window).
3. `refresh_xl_x2a.xlsx` (X2, UI pre-test).
4. Then D1, D3 and D4 as needed, and X3 only if X2 shows that Refresh All without the source recalculates.

Each takes under two minutes in a fresh Excel session.

## 8. Limits of this research

- WebFetch was blocked (learn.microsoft.com, support.microsoft.com, sqlbi.com, blog.crossjoin.co.uk). Points marked [D] come from search-result summaries of:
  - Microsoft "Recalculate Formulas in Power Pivot"
  - SQL Server 2012 "Troubleshoot Recalculation" and "Calculated Columns"
  - "Processing tabular models 101"
  - KB2880071
  - VBA reference pages for `ModelConnection.ADOConnection`, `Model.DataModelConnection` and `WorkbookConnection.RefreshWithRefreshAll`
  - SQLBI "Automate PowerPivot data refresh in Excel"
- The attribute name `excludeFromRefreshAll` and the exact XMLA accepted by `ADOConnection.Execute` are unverified.
- No experiment here has been run in Excel.
