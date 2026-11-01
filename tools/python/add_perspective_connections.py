"""Experiment H68 (2026-11-01): add one model connection per perspective to xl/connections.xml of a workbook.

    python3 tools/python/add_perspective_connections.py <in.xlsx> <out.xlsx> <perspective name> [<perspective name> ...]

Each added connection is a copy of `ThisWorkbookDataModel` named `ThisWorkbookDataModel - <name>` whose command is the perspective name
instead of `Model`. Result (H70, the file after Excel saved it): Excel keeps the connections and rewrites every command to `Model`, so the
connection command cannot select a perspective. The model part is copied unchanged. See docs/experiments/pivot-perspective.md.
"""
import sys
import zipfile


def main():
    src, dst, names = sys.argv[1], sys.argv[2], sys.argv[3:]
    if not names:
        sys.exit(__doc__)
    zin = zipfile.ZipFile(src)
    xml = zin.read("xl/connections.xml").decode("utf-8")
    extra = ""
    for i, name in enumerate(names):
        extra += (f'<connection id="{3 + i}" xr16:uid="{{00000000-0015-0000-FFFF-FFFF0{2 + i}000000}}" keepAlive="1" name="ThisWorkbookDataModel - {name}" '
                  f'description="Data Model perspective {name}" type="5" refreshedVersion="8" minRefreshableVersion="5" background="1">'
                  f'<dbPr connection="Data Model Connection" command="{name}" commandType="1"/><olapPr sendLocale="1" rowDrillCount="1000"/>'
                  '<extLst><ext uri="{DE250136-89BD-433C-8126-D09CA5730AF9}" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main">'
                  '<x15:connection id="" model="1"/></ext></extLst></connection>')
    assert xml.count("</connections>") == 1
    xml = xml.replace("</connections>", extra + "</connections>")
    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            zout.writestr(info, xml.encode("utf-8") if info.filename == "xl/connections.xml" else zin.read(info.filename))


if __name__ == "__main__":
    main()
