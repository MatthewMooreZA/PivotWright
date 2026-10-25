# PivotWright — instructions for Claude Code

Start with `docs/HANDOFF.md` (state, next steps, testing loop) and `docs/SPEC.md` (format specification).

- Build/test: `dotnet build` and `dotnet run --project tests/PivotWright.Tests` (all 58 must pass before committing).
- No Excel here. The owner (Matthew) tests generated files in desktop Excel and reports pass/fail. Send the single most-likely file first, and give every file sent a new unique name (never reuse a name across iterations); do not send more while he is mid-test.
- Excel-written samples in `reference/excel-samples/` are ground truth. Compare with `tools/python/extract_tree.py`, `xd.py`, `dd.py`.
- Version counters must match file names and backup-log versions (SPEC §5). Never change one without the other.
- Models without hierarchies must stay unchanged by hierarchy-profile work (the one exception: every model gets distinct `Ordinal`s, SPEC §8).
- No temporary env-var toggles in committed code.
- Public API changes follow `docs/API-GUIDELINES.md` (naming, table handles, errors, defaults); update README, SPEC §2 and tests with them.
- Commit messages end with the attribution lines the session's system reminder gives.
