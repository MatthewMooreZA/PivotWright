# Second opinion: a file-only trigger for calculating calculated columns (todo item 5)

Status: research only, 2026-10-31. Nothing here has been tested in Excel. Marks: **[E]** seen in a sample or a result in `docs/research-refresh.md`; **[D]** public documentation (search snippets only: `learn.microsoft.com` does not resolve from this sandbox, so no page could be read in full); **[S]** speculation.

Samples were unpacked with `tools/repack dump` into the scratchpad and compared with small `python3 -I` scripts (state fields, timestamps, raw diffs). H41 and H43 matter most. Both were saved by Excel from the same generated file, `refresh_xl_x2a`, so they share every id. H41 is still uncalculated. H43 is the same model after Excel calculated it in the session (Sum of Line Total put in Values). Together they are an exact before and after pair.

## 1. Summary

- **The engine never processes on its own, so model state cannot do the job by itself.** Every calculation seen so far followed a command from a client: Calculate Now, VBA ProcessRecalc, an implicit measure created from the field list, or a table being added. A file that loads with mismatched versions or unprocessed objects has never been seen to calculate on open. H10 opens with the Sales partition at DataVersion 1 and the Sales dimension at DataVersion 3, and X7 did not calculate on open [E]. An unprocessed user hierarchy (H31B, `calchier_v1`) also does not calculate on open or on refresh [E]. Section 2 has the detail.
- **New, and the strongest lead: processing any table's data calculates the whole model, even in Power Pivot's manual mode** [E, H42 timestamps]. Excel added the linked table `Trigger` with `ManualCalcMode` True. The Trigger partition was processed at 18:28:24.82. In the same transaction, 30 ms later, the Sales dimension (`Line Total`) was processed, at 18:28:24.853, which is also the database's `LastProcessed`. Section 18 of the earlier research correctly withdrew H42 as evidence about *opening*. It is, however, direct evidence that a **data refresh of an unrelated table** calculates every pending calculated column.
- **What was never validly tested:** refreshing an *existing* model table whose source needs no credentials. On open, x11 showed that Excel does not sync a linked table by itself. Nobody has tried:
  - Refresh All on a file whose linked table is out of date;
  - `refreshOnLoad` on the linked table's own type 102 connection, in a file whose column is still unprocessed (x3 had this patch, but on H42, which was already calculated);
  - a Power Query table built from a constant (no source, so no credentials) with "Refresh data when opening the file".

  These are experiments E1 to E3 below.
- **Model-state levers are weak but cheap to probe once.** There is one untried state flag with a plausible client reaction: a relationship's `ddl300:ProcessingState` (E4). Every other state field either has no effect on load (shown above) or would send Excel to the placeholder source (credentials prompt).
- **Honest conclusion.** I found no flag in the model or the workbook that makes *loading* the model calculate. A file-only trigger most likely has to be a table refresh that needs no credentials, either on open (`refreshOnLoad` on that table's own connection) or on Refresh All, with the PivotWright source excluded (H41). The earlier research dropped that route because of x11, but x11 tested only opening without `refreshOnLoad`.

## 2. Point 1: what Excel changes when it calculates (H41 to H43, file by file) [E]

Everything else in the two models is the same (`tools/python/xd.py` plus raw diffs with GUIDs normalised).

| Object | H41 (uncalculated) | H43 (calculated in session) |
|---|---|---|
| Database `<id>.1.db.xml` | `LastProcessed` 10-29 18:22:05, DataVersion 0 | `LastProcessed` 10-31 10:49:10.12, DataVersion 0; ObjectVersion and `LastSchemaUpdate` unchanged |
| Sales dimension | `.7.dim.xml`, DataVersion 3, `LastProcessed` = creation time; Line Total `ddl200:ProcessingState` and `ddl300:AttributeHierarchyProcessingState` Unprocessed | `.8.dim.xml`, DataVersion 4, `LastProcessed` 10:49:10.12, both states Processed. **`LastSchemaUpdate` unchanged**, so this is a data-only change. StructureVersion 1, DataIDVersion 2 and CacheVersion 2 are unchanged |
| Sales dimension `info` file | `info.7.xml`, DataVersion 0 | `info.8.xml`, DataVersion 3 (the dimension's previous data version) |
| Sales `tbl.xml` | Version 3; intrinsic hierarchy `IsProcessed` false, -1 values, no `TableStore`; 1 BLANK run | Version 4; `IsProcessed` true, `TableStore` `H$Sales…$Line Total`, 37 distinct values, real compression |
| Line Total data | `2.Sales….Line Total.0.idf` (152-byte stub) | `4.….Line Total.0.idf`, `4.….Line Total.dictionary`, `H$Sales…$Line Total.0.tbl.xml` and its files |
| Sales partition `.11.prt.xml` | DataVersion 3 | DataVersion 4 (file name unchanged: data, not schema) |
| Sales measure group | `.26.det.xml`, DataVersion -1, CacheVersion 3 | `.27.det.xml`, **DataVersion 4**, CacheVersion 4, `LastProcessed` 10:49:10.12 |
| MdxScript | `.11.scr.xml` | `.12.scr.xml`, new `CREATE MEASURE [Model].'Sales'[Sum of Line Total]=SUM('Sales'[Line Total])` with annotations `Type`=Implicit, `Application`=Excel, Column, Aggregation, FullName, Table, and a hidden `CalculationProperty` with `IsSimpleMeasure` (`LastSchemaUpdate` 10:49:10.08, 40 ms before the processing) |
| Cube | `.26.cub.xml` | still `.26`; only `MeasureGroupFileList` and `AssemblyFileList` name the new files |
| Products, Dates, their partitions and measure groups | unchanged | unchanged (a recalculation only touches tables that hold calculated columns) |

Where the state lives, and what does not exist [E, from all ASSL files of H6, H10, H41 and H43]:

- **Attribute.** `ddl200:ProcessingState` and `ddl300:AttributeHierarchyProcessingState`. This is the only place "needs calculation" is recorded.
- **Dimension.** A dimension-level `ddl300:ProcessingState`, Processed even in H10. This corrects section 4 of the earlier research, which said dimensions have no persisted state element.
- **Relationship.** `ddl300:ProcessingState` on each relationship in the many-side dimension (Processed).
- **User hierarchy.** `ddl300:ProcessingState` on user hierarchies (Unprocessed in H31B).
- **Measure group.** `ddl200:ProcessingState` on its dimensions (Processed).
- **Nothing persisted** on cube, partition, measure group or database. No "dirty" flag, no ProcessRecalc marker, no pending-command list.
- **Persisted enumeration** [D, AMO 2012 `ProcessingState`: Processed, Unprocessed, CalculationError, DependencyError, InvalidExpression]. The documentation describes Unprocessed as "the attribute still needs processing". The object-level `AnalysisState` (Processed, PartiallyProcessed, Unprocessed) is computed at run time, not stored [D].

Judgement on the ideas in the brief:

- **Stale or mismatching DataVersion (dimension vs partition vs measure group).** Already present in H10 (partition 1, dimension 3, measure group -1) and did not calculate on open (X7) [E]. In multidimensional SSAS a version mismatch invalidates partitions. Invalid partitions need *source* processing, so the best case is the credentials prompt [S].
- **`LastProcessed` timestamps.** Informational only [S]. H6, H10 and H41 have every pattern (`LastProcessed` before, equal to or after `LastSchemaUpdate`) without any effect.
- **Dimension, cube, measure group or partition marked unprocessed.** The engine would then refuse queries ("not processed"), or Excel would process from the source and prompt. The SPEC §5.2 note supports this: a measure group without a valid partition produced "Needs Calculation" **and a refresh password prompt** [E]. So when Excel sees an unprocessed object at refresh time, it processes from the source, not with ProcessRecalc.
- **The one state with a source-free repair is a relationship.** ProcessRecalc is what rebuilds relationships [D]. An unprocessed relationship is untested (E4).
- **The stub `.idf` in `DataFileList`.** It is listed in H10 exactly like a real file. Removing it would only risk a load error [S].
- **A "calculated but empty" column that Excel recomputes on load.** Not possible [S, strong]. A Processed column with BLANK or placeholder values is a valid processed column. Neither ProcessRecalc after an unrelated table refresh nor a query would recompute it, because a recalculation after a refresh redoes only invalidated or unprocessed objects [D, "Processing tabular models": ProcessRecalc recalculates objects that need it; S for the exact invalidation rule]. Calculate Now does a full recalculation [D]. **So the column must stay in the unprocessed form, and the trigger must be a command.**

**Side finding: Excel alters the first relationship's primary-side dimension in every session, and that alter does not calculate** [E].

- **What H41 shows.** H41's session made no model edit, only the connection box. Even so, Products went from 6 to 7 (`LastSchemaUpdate` 10-29 18:24:38.5) and the cube and every measure group went from 25 to 26, while the content was unchanged apart from the `DataFileList` order and `EstimatedSize`.
- **What H43 shows.** H43's session, from the same x2a, did the same 5 s before the implicit measure (Products `LastSchemaUpdate` 10:49:05.14).
- **The same in older samples.** H6, H10, H2 and H8 show it too (Products `LastSchemaUpdate` 17 to 70 s before the user's edit).
- **Consequences.**
  1. An engine-level schema change alone does not trigger a recalculation. What follows an edit is a client decision.
  2. SPEC §10's guess, "the primary side of the first relationship is saved again", is almost certainly this per-session alter, not part of adding a column. No action is needed: the Excel-verified counters stand.

## 3. Point 2: version and upgrade paths

- **The Gemini parts are read only by the Power Pivot add-in, not by Excel core** [E]. H43 and H41 were saved by Excel with no customXml at all: the add-in did not load in those sessions. Yet in H43 Excel core calculated the column and created the implicit measure. Excel core therefore does not need these parts, and probably never reads them. `PowerPivotVersion`, `ManualCalcMode` and `ErrorCache` (`LastProcessedTime`) can at most affect the Power Pivot window. D3 (open the window) and X6 (the part alone) were negative [E].
- **An older `PowerPivotVersion` (for example 10.50.x, Power Pivot 2010).** It might show the add-in's "upgrade this data model" prompt when the Power Pivot window opens [S]. A 2010 model is stored differently (a different part and compatibility level), so a 2013-format model labelled 2010 is more likely to be rejected or ignored. An upgrade also needs a user click and, as far as I recall, refreshes from the sources [S]. That breaks the "file alone" requirement and brings back the prompt. Not recommended (E8 only if everything else fails).
- **`ddl200:CompatibilityLevel` 1103 in `db.xml`** [E]. Lowering it (1100) is not known to start an upgrade on load. In SSAS, changing the level needs an explicit Alter from a client [D/S]. Not proposed.

## 4. Point 3: writing the values

- **Placeholder or wrong values written as Processed** satisfy "usable as Rows" immediately, but show wrong labels. No later trigger corrects them, because the engine considers them up to date (section 2). Only Calculate Now or an edit to a column they depend on would. Rejected.
- **The narrow case** is a bare column reference (`[Qty]`), `RELATED(T[c])` or a constant. The values are copies, not evaluation. The owner's "no DAX evaluation" rule covers it in spirit, so it is only a fallback to raise with the owner. It is not part of the experiments.

## 5. Point 4: Excel-side mechanisms outside the model

| Mechanism | Reaches the process path? | Evidence |
|---|---|---|
| `calcPr fullCalcOnLoad`, `forceFullCalc`, `calcMode`, iterate | Recalculates worksheet formulas only | [D] ECMA-376 CT_CalcPr |
| CUBEVALUE or CUBEMEMBER on open | A query only. A query does not calculate: D2 shows the Values case works through an Alter. The decisive diagnostic is E5 | [S] |
| Pivot cache `refreshOnLoad` (x12, x13) | Prompts even with the source excluded. A model-cache refresh refreshes the whole model, and excluding a connection from Refresh All does not apply to a pivot refresh | [E] |
| `ThisWorkbookDataModel` `refreshOnLoad` (X4/X4b, never sent) | Probably the same as x13 | [S] |
| A type 102 (worksheet) connection with `refreshOnLoad` on an **unprocessed** file | **Untested** (x3 was on calculated H42). Refreshes one linked table, and H42 shows that a table data process calculates everything | [E] mechanism / [S] honoured on open |
| A Power Query (Mashup) connection with `refreshOnLoad`, loaded to the model only | **Untested**. Standard "Refresh data when opening the file". A query with no source (`#table`) needs no credentials and has no privacy levels. Format documented in [MS-QDEFF] (DataMashup customXml part) | [D] format / [S] behaviour |
| Refresh All with a linked or Power Query table and the PivotWright source excluded (H41) | **Untested**. X2 (exclusion alone) had nothing to refresh, so nothing was processed | [E] for X2 |
| Slicer or timeline cache on Line Total | Queries members of the hierarchy, which fails like Rows. No Alter | [S] |
| Sparklines, query tables bound to sheets | No link to the model's process path | [S] |
| `x15:dataModel` (`modelTables`, `modelRelationships`) | A mirror. A dangling table entry is ignored (X8) | [E] |
| `xlpdsc:dataSourceClean` (pivot cache ext, namespace `…/2026/pivotDataSourceClean`, value 1 in H43) | New and undocumented (search finds nothing). Plausibly Excel's persisted "model is clean" bit for a pivot. As 0 it might make the first pivot refresh treat the model as changed. That probably means a source refresh and a prompt (the in-session behaviour after an edit), but it might mean a recalculation. E6b | [E] it exists / [S] meaning |

## 6. Point 5: why a value calculates everything

- **The sequence in H43** [E]: (1) the Products alter on first load, with no calculation; (2) an Alter of the MdxScript adding the implicit measure; (3) 40 ms later, every pending calculated column is processed in the same commit. The calculation therefore follows the client's *implicit-measure creation* (Excel core, not the add-in). It is not caused by the query that fills the value: the query comes after the measure exists.
- **Shipping the implicit measure already in the model** (MdxScript command with the six annotations) [S, likely harmful]:
  - Excel would *reuse* it when the user drops Line Total into Values, with no Alter and so no calculation.
  - The values would then depend on what a query over an unprocessed column does (E5). If that query returns an error, shipping the measure makes things worse.
  - Shipping it is not a trigger.
- **An explicit measure over the column** (`SUM(Sales[Line Total])`) in Values was never tested. D1 (the existing `Total Qty`) and D4 were proposed but never reported.
  - If D4 shows R90500 **and** Rows then works, the engine calculates on query. CUBEVALUE with `fullCalcOnLoad` then becomes a macro-free trigger.
  - If it shows blank or an error, every query-based route (cube functions, slicers, pivots without an Alter) is dead.

  This is E5: cheap, and it settles a whole family.

## 7. Point 6: documented semantics

- **ProcessRecalc** [D]: recalculates calculated columns, rebuilds relationships and hierarchies, and does not touch the source. ProcessFull and ProcessData of a table are followed by a recalculation of what they invalidate. ProcessDefault processes only unprocessed objects. For partitions that means the source.
- **"Recalculate formulas in Power Pivot"** [D]: refreshing data is listed among the operations followed by a recalculation in automatic mode. H42 shows that adding a table calculated pending columns even in manual mode [E].
- **Load or attach in SSAS.** I know of no documented flag that processes on load (no "process on attach"; proactive caching needs a source and is disabled in every sample: `ProactiveCaching/Enabled` false) [S].
- **Not verified (blocked or not found):**
  - The [MS-XLDM] description of the per-object state fields.
  - The [MS-XLSX] page for `x15:connection` attributes.
  - Anything on `xlpdsc:dataSourceClean`.
  - Whether Excel honours `refreshOnLoad` on a type 102 connection.

## 8. Ranked experiments (most likely first)

Common rules:

- A fresh Excel session for each file.
- A new name for each file.
- One file at a time.
- The **standard check**: Insert > PivotTable > From Data Model, then drag `Line Total` straight into Rows. Success means its distinct values are listed.
- "ZIP patch" means the stdlib helper of research §5 (string replace in one package part).
- "ABF text patch" means a scratch copy of `tools/repack` with one extra mode, modelled on `ordinal` and `detfix`, run outside the repo:

```csharp
// textpatch: args[3] = path substring, args[4] = old, args[5] = new (exactly one occurrence per matching file)
if (mode == "textpatch") {
    var lg = XElement.Parse(m.RawLogXml); int n = 0;
    foreach (var bf in lg.Descendants("BackupFile")) {
        string path = bf.Element("Path")!.Value, st = bf.Element("StoragePath")!.Value;
        if (!path.EndsWith(".xml") || !path.Contains(args[3])) continue;
        string t = System.Text.Encoding.UTF8.GetString(contentByStorage[st]);
        int i = t.IndexOf(args[4]); if (i < 0) continue; Check.True(t.IndexOf(args[4], i + 1) < 0, "unique in " + path);
        var b = new System.Text.UTF8Encoding(false).GetBytes(t.Remove(i, args[4].Length).Insert(i, args[5]));
        contentByStorage[st] = b; bf.Element("Size")!.Value = b.Length.ToString(); n++;
    }
    Check.True(n > 0, "patched"); logXml0 = lg.ToString(SaveOptions.DisableFormatting);
}
```

Versions are untouched, so SPEC §5 is not affected (only text inside one file changes).

### E1: linked table, out of date, refreshed on open and by Refresh All (library-built, no sample needed)

- **Build.**
  1. `git worktree add <scratch>/wt 5482b41` (the reverted `AddCalculationTrigger`).
  2. `dotnet run --project <scratch>/wt/tools/datagen -- <out> refresh_trig_e1_src calc trigger`.
  3. ZIP patches:
     - `xl/worksheets/sheet2.xml`: `<c r="A2"><v>1</v></c>` becomes `<c r="A2"><v>2</v></c>` (as x11).
     - `xl/connections.xml`: `type="102"` becomes `type="102" refreshOnLoad="1"`.
     - `xl/connections.xml`: connection 1's `<x15:connection id="<ds id>"/>` becomes `<x15:connection id="<ds id>" excludeFromRefreshAll="1"/>` (H41).
  4. Name the result `refresh_trig_e1.xlsx`.
- **Why it might work.** H42 proves that a linked-table data process calculates everything [E]. The linked table needs no credentials, and the exclusion keeps Refresh All away from the placeholder source. x11 lacked `refreshOnLoad` and was never refreshed.
- **Check (two minutes).**
  1. Open the file and click Enable Content if the bar appears. Note any prompt.
  2. New PivotTable: put `Trigger[Key]` in Rows. Does it show 2 or 1? Then run the standard check in a fresh PivotTable.
  3. If either check failed: Data > Refresh All. Note any prompt, then `Trigger[Key]` again and the standard check.
- **Reading the result.**
  - Key shows 2 and Rows works on open: a file-only on-open trigger.
  - Key is 2 only after Refresh All, and Rows works: an on-refresh trigger.
  - Key stays 1 after Refresh All: Excel does not refresh this library-written linked table. Go to E2.

### E2: the same on an Excel-made linked table (sample request H44)

- **Sample H44.** Use this order so that the column stays unprocessed (adding a table calculates, even in manual mode, as H42 shows):
  1. Open `datagen … calcbase_e2` (the base model, with no calculated column).
  2. On Sheet2 make a one-cell table `Trigger` (Key = 1) and add it to the Data Model.
  3. Set Power Pivot to Manual calculation.
  4. Add `Line Total` = `[Qty]*[Amount]` and do **not** calculate.
  5. Save as H44.
- **Test file `refresh_trig_e2.xlsx`.** H44 with these ZIP patches:
  - `ManualCalcMode` set to False;
  - the sheet cell set to 2;
  - `refreshOnLoad="1"` on the type 102 connection;
  - the H41 exclusion on PivotWright.
- **Check.** As E1.
- **Why.** It removes all doubt about the library's linked-table markup (lib_v1 counters were never matched to H42). It is the experiment §18 implied but never ran.

### E3: a constant Power Query table refreshed on open (sample request H45)

- **Sample H45.** Same order as E2, but the table comes from Data > Get Data > Blank Query `= #table({"Key"}, {{1}})`, named `Trigger`, loaded with Close & Load To > Only Create Connection + Add to the Data Model. In Query Properties, tick "Refresh data when opening the file" and untick the PivotWright connection's "Refresh this connection on Refresh All". Then set manual mode, add `Line Total` and save without calculating.
- **Test file `refresh_pq_e3.xlsx`.** H45 with `ManualCalcMode` False.
- **Check.** Open, Enable Content, run the standard check. If that fails, Refresh All and run the standard check again.
- **Why.**
  - `refreshOnLoad` is certainly honoured for Power Query connections [S, mainstream feature]. The query has no source, so there are no credentials and no privacy prompt.
  - The table data process then calculates everything (as in H42).
  - Use it if E1 and E2 show that a linked table is never refreshed automatically.
- **Library cost if it works.**
  - A `customXml` DataMashup part ([MS-QDEFF], documented).
  - A Mashup data source and connection.
  - A small model table: H45 gives the exact markup.

### E4: model state, an unprocessed relationship (ABF text patch, no sample)

- **File `refresh_state_e4.xlsx`.** Start from `datagen … calc`. In the Sales dimension, set the first relationship's state to Unprocessed: `textpatch "Sales_" "<ddl300:ProcessingState>Processed</ddl300:ProcessingState></ddl300_300:Relationship>" "<ddl300:ProcessingState>Unprocessed</ddl300:ProcessingState></ddl300_300:Relationship>"`. First check the exact neighbouring tag in the dump, because the state may not be the last child, and make the patch unique.
- **Why it might work.** Only ProcessRecalc rebuilds a relationship [D], it needs no source, and any recalculation is database-wide [E D2/H42]. If Excel, or the engine on first query, repairs a relationship it needs, the column is calculated as a side effect.
- **Check.** Open. Put `Products[Category]` in Rows and `Total Qty` in Values. Note values, an error or a prompt. Then run the standard check.
- **Expected.** Probably an error or a source refresh [S]. A negative result closes the model-state angle.
- **Variant E4b** (only if E4 loads cleanly but does nothing): the Sales dimension-level `<ddl300:ProcessingState>Processed</ddl300:ProcessingState>` (the first occurrence, directly under `<Dimension>`) set to Unprocessed.

### E5: does a query calculate? (diagnostics D1/D4, never reported)

- **File `refresh_diag_d4b.xlsx`.** A scratch datagen line after the calculated column: `s.AddMeasure("Line Total Sum", "SUM(Sales[Line Total])", MeasureFormat.WholeNumber)`.
- **Check.**
  1. Put only `Line Total Sum` in Values. Does it show 90500, blank or an error?
  2. Then run the standard check.
  3. In a second fresh session, put only `Total Qty` in Values, then run the standard check (D1).
- **Use.**
  - If D4 passes both: the engine calculates on query, and `=CUBEVALUE("ThisWorkbookDataModel","[Measures].[Line Total Sum]")` with `fullCalcOnLoad="1"` (research X5) is the trigger.
  - Otherwise, drop every query-based idea.

### E6: a PivotTable whose implicit measure is missing from the model, refreshed by hand

- **File `refresh_pivot_e6.xlsx`.** x12 without `refreshOnLoad`, with the H41 exclusion.
- **Check.**
  1. Open the file.
  2. Right-click the existing PivotTable > Refresh. Note any prompt; if one appears, enter nothing and press Cancel.
  3. Run the standard check.
- **Why.** The cache names `[Measures].[Sum of Line Total]` (`x15:cacheHierarchy aggregatedColumn`), but our model has no such measure. Excel may re-create it, which is exactly the D2 Alter, and then recalculate. The prompt in x12 and x13 even with the exclusion fits "the model changed, so the refresh goes to the source" [S].
- **Variant E6b.** In the same file, `<xlpdsc:dataSourceClean>1<` becomes `0<`.

### E7: `refreshOnLoad` on `ThisWorkbookDataModel` with the exclusion (X4b, never sent)

- **File.** `datagen … calc`, with the H41 exclusion on connection 1 and `refreshOnLoad="1"` on connection 2.
- **Check.** Open, note any prompt, run the standard check.
- **Odds.** Low (x13 analogue). It completes the list of connection-level flags.

### E8 (last resort): `PowerPivotVersion` 10.50.1600.1 and the Power Pivot window

- **File.** `datagen … calc` with the H6 Gemini parts, `PowerPivotVersion` set to `10.50.1600.1`.
- **Check.** Open Data > Manage Data Model. Note any upgrade prompt and accept it. Then run the standard check.
- **Odds.** Low, and it needs user clicks. It only tells whether an upgrade path recalculates.

## 9. Recommended order

1. E1, then E2 if Excel ignored the library's linked table.
2. E5 at any point (it is free and settles the query family).
3. E3 if E1 and E2 show that linked tables never refresh automatically.
4. E4 to close the model-state angle.
5. E6 and E7 only if everything above fails.

If E1 to E3 all fail, the conclusion "no file-only trigger" stands with better grounds than before: no stored state (section 2) and no workbook flag (section 5) reaches the engine's process path.
