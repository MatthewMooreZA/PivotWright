# repack

Scratch tool used for the hidden-column bisect (SPEC §8): reads the model of an Excel workbook with `ModelReader`, changes parts of it and writes it back through `AbfWriter`, so that hybrids of Excel's files and PivotWright's can be tested in Excel. It is built as `PivotWright.Tests` to see the library internals.

`dotnet run --project tools/repack -- <in.xlsx> <out.xlsx|outdir> <mode> [args]`

Modes: `verbatim`, `dump` (unpack to a folder), `ourlog`, `dbfolder`, `tsmap`, `idmap` (rename ids to those of another workbook; letters t r s d o), `swap` (take ASSL files from another workbook; letters d p c s v t x, O keep ObjectIDs, T keep timestamps; args[5] limits it to paths containing that text), `rswap`, `rename old=new,...`, `ordinal kind:Name=value,...`, `detfix`. Not part of the library or its tests.
