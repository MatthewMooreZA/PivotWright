"""Experiment (2026-11-01): put a hand-written PivotTable into a generated workbook, on the default cube and on a perspective.

    python3 tools/python/pivot_experiment.py <generated.xlsx> <out.xlsx> <variant> [perspective]

The generated workbook is the standard model (Products, Dates, Sales, measure Total Qty; `datagen ... persp3`). The PivotTable parts are
adapted from Excel's own `H43_pivot_calc_value.xlsx` (reference/excel-samples), which holds the same three tables.

Variants (see docs/experiments/pivot-perspective.md):
  default      control: a values-only PivotTable (Total Qty) on the cube "Model", connection 2 as Excel writes it.
  conn         a PivotTable on a perspective: a second model connection whose command is the perspective name, and a cache that lists
               only what the perspective shows.
  conn-full    as conn, but the cache still lists the whole model (the connection alone decides).
  string       as conn, with an explicit MSOLAP connection string that has Cube=<perspective> (the advice found on the web).

refreshOnLoad is off: with it Excel asked for the placeholder data source's credentials on opening (pivot_default_v1), and Excel's own PivotTables (H43) do not have it.
Nothing here is part of the library; it only produces files for Excel to open.
"""
import re
import sys
import zipfile
from pathlib import Path

TEMPLATE = Path(__file__).resolve().parents[2] / "reference/excel-samples/H43_pivot_calc_value.xlsx"

# What each perspective of datagen's persp3 shows: hierarchy unique names and measures.
PERSPECTIVES = {
    "1 - subset of columns": (["[Products].[Product]", "[Products].[ProductKey]", "[Sales].[Amount]"], []),
    "2 - measure and hierarchy": (["[Products].[Product Hierarchy]"], ["[Measures].[Total Qty]"]),
    "3 - both": (["[Products].[Product]", "[Products].[ProductKey]", "[Products].[Product Hierarchy]"], ["[Measures].[Total Qty]"]),
}

HIERARCHY = re.compile(r'<cacheHierarchy [^>]*/>|<cacheHierarchy [^>]*>.*?</cacheHierarchy>', re.S)
PIVOT_HIERARCHY = re.compile(r'<pivotHierarchy[^>]*/>')


def unique_name(element):
    return re.search(r'uniqueName="([^"]*)"', element).group(1)


def build_cache(template, perspective, connection_id, refresh_on_load):
    """The pivotCacheDefinition for a values-only PivotTable on Total Qty; with a perspective only its hierarchies and measures stay."""
    head, rest = template.split("<cacheFields", 1)
    hierarchies_xml = re.search(r'<cacheHierarchies count="\d+">(.*?)</cacheHierarchies>', rest, re.S).group(1)
    elements = HIERARCHY.findall(hierarchies_xml)
    # Total Qty is used on the Values area: explicit measure, as Excel writes it (H63).
    total_qty = ('<cacheHierarchy uniqueName="[Measures].[Total Qty]" caption="Total Qty" measure="1" displayFolder="" measureGroup="Sales" '
                 'count="0" oneField="1"><fieldsUsage count="1"><fieldUsage x="0"/></fieldsUsage></cacheHierarchy>')
    kept = []
    for e in elements:
        name = unique_name(e)
        if "Line Total" in name:
            continue                      # H43 has a calculated column that our model does not
        if name == "[Measures].[Total Qty]":
            e = total_qty
        if perspective is not None:
            shown, measures = PERSPECTIVES[perspective]
            if name not in shown + measures:
                continue
        kept.append(e)
    index = {unique_name(e): i for i, e in enumerate(kept)}
    fields = ('<cacheFields count="1"><cacheField name="[Measures].[Total Qty]" caption="Total Qty" numFmtId="0" hierarchy="%d" level="32767"/></cacheFields>'
              % index["[Measures].[Total Qty]"])
    dims = sorted({re.match(r'\[(\w+)\]', unique_name(e)).group(1) for e in kept if not unique_name(e).startswith("[Measures]")})
    dim_names = sorted(dims + ["Measures"])
    dimensions = "".join(
        ('<dimension measure="1" name="Measures" uniqueName="[Measures]" caption="Measures"/>' if d == "Measures"
         else f'<dimension name="{d}" uniqueName="[{d}]" caption="{d}"/>') for d in dim_names)
    groups = ["Sales"] if perspective is not None else ["Dates", "Products", "Sales"]
    group_xml = "".join(f'<measureGroup name="{g}" caption="{g}"/>' for g in groups)
    if perspective is None:
        maps = [("Dates", "Dates"), ("Products", "Products"), ("Sales", "Dates"), ("Sales", "Products"), ("Sales", "Sales")]
    else:
        maps = [("Sales", d) for d in dims if d != "Sales"]
    map_xml = "".join(f'<map measureGroup="{groups.index(g)}" dimension="{dim_names.index(d)}"/>' for g, d in maps if g in groups)
    tail = rest[rest.index("</maps>") + len("</maps>"):]   # the extLst after the maps, not the one inside a cacheHierarchy
    body = ('<cacheHierarchies count="%d">%s</cacheHierarchies><kpis count="0"/><dimensions count="%d">%s</dimensions>'
            '<measureGroups count="%d">%s</measureGroups><maps count="%d">%s</maps>'
            % (len(kept), "".join(kept), len(dim_names), dimensions, len(groups), group_xml, len(maps), map_xml))
    head = head.replace('connectionId="2"', f'connectionId="{connection_id}"') if 'connectionId="2"' in head else head
    out = head + fields + body + tail
    out = out.replace('<cacheSource type="external" connectionId="2"/>', f'<cacheSource type="external" connectionId="{connection_id}"/>')
    if refresh_on_load:
        out = out.replace(' saveData="0"', ' saveData="0" refreshOnLoad="1"', 1)
    return out, index, kept


def build_pivot(template, index, kept):
    """The PivotTable: only the data field Total Qty, in the layout of H43's single value."""
    pivot_hierarchies = []
    # H43 lists a pivotHierarchy per cache hierarchy, in the same order; look them up by position in the template.
    template_hierarchies = PIVOT_HIERARCHY.findall(template)
    template_cache_names = TEMPLATE_CACHE_NAMES
    for e in kept:
        name = unique_name(e)
        i = template_cache_names.index(name)
        pivot_hierarchies.append(template_hierarchies[i])
    out = re.sub(r'<pivotHierarchies count="\d+">.*?</pivotHierarchies>',
                 '<pivotHierarchies count="%d">%s</pivotHierarchies>' % (len(pivot_hierarchies), "".join(pivot_hierarchies)), template, flags=re.S)
    out = out.replace('<dataField name="Sum of Line Total"', '<dataField name="Total Qty"')
    return out.replace('cacheId="5"', f'cacheId="{CACHE_ID}"')


CACHE_ID = 5
TEMPLATE_CACHE_NAMES = []


def main():
    base, out_path, variant = sys.argv[1], sys.argv[2], sys.argv[3]
    perspective = sys.argv[4] if len(sys.argv) > 4 else "3 - both"
    if variant not in ("default", "conn", "conn-full", "string"):
        sys.exit("variant must be default, conn, conn-full or string")
    if variant != "default" and perspective not in PERSPECTIVES:
        sys.exit("perspective must be one of: " + "; ".join(PERSPECTIVES))
    with zipfile.ZipFile(TEMPLATE) as t:
        cache_template = t.read("xl/pivotCache/pivotCacheDefinition1.xml").decode("utf-8")
        pivot_template = t.read("xl/pivotTables/pivotTable1.xml").decode("utf-8")
    TEMPLATE_CACHE_NAMES.extend(unique_name(e) for e in HIERARCHY.findall(cache_template))
    use_perspective = perspective if variant in ("conn", "string") else None
    connection_id = 2 if variant == "default" else 3
    cache, index, kept = build_cache(cache_template, use_perspective, connection_id, refresh_on_load=False)
    pivot = build_pivot(pivot_template, index, kept)

    with zipfile.ZipFile(base) as z:
        parts = {i.filename: z.read(i.filename) for i in z.infolist()}
        order = [i.filename for i in z.infolist()]

    def text(name):
        return parts[name].decode("utf-8")

    # Connections: variants add connection 3, a copy of the data model connection.
    if variant != "default":
        connections = text("xl/connections.xml")
        if variant == "string":
            command = (f'connection="Provider=MSOLAP.8;Integrated Security=SSPI;Persist Security Info=True;'
                       f'Initial Catalog=Microsoft_SQLServer_AnalysisServices;Data Source=$Embedded$;MDX Compatibility=1;Safety Options=2;'
                       f'MDX Missing Member Mode=Error;Optimize Response=3;Cell Error Mode=TextValue;Cube={perspective}" command="{perspective}" commandType="1"')
        else:
            command = f'connection="Data Model Connection" command="{perspective}" commandType="1"'
        extra = (f'<connection id="3" keepAlive="1" name="ThisWorkbookDataModel - {perspective}" description="Data Model" type="5" '
                 f'refreshedVersion="8" minRefreshableVersion="5" background="1"><dbPr {command}/><olapPr sendLocale="1" rowDrillCount="1000"/>'
                 '<extLst><ext uri="{DE250136-89BD-433C-8126-D09CA5730AF9}" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main">'
                 '<x15:connection id="" model="1"/></ext></extLst></connection>')
        parts["xl/connections.xml"] = connections.replace("</connections>", extra + "</connections>").encode("utf-8")

    # Workbook: the cache, its relationship and the content types.
    workbook = text("xl/workbook.xml")
    assert "<pivotCaches>" not in workbook
    workbook = workbook.replace("<calcPr calcId=\"191029\"/>", f'<calcPr calcId="191029"/><pivotCaches><pivotCache cacheId="{CACHE_ID}" r:id="rId9"/></pivotCaches>', 1)
    assert "<pivotCaches>" in workbook
    parts["xl/workbook.xml"] = workbook.encode("utf-8")
    rels = text("xl/_rels/workbook.xml.rels").replace(
        "</Relationships>", '<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition" Target="pivotCache/pivotCacheDefinition1.xml"/></Relationships>')
    parts["xl/_rels/workbook.xml.rels"] = rels.encode("utf-8")
    types = text("[Content_Types].xml").replace(
        "</Types>",
        '<Override PartName="/xl/pivotCache/pivotCacheDefinition1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml"/>'
        '<Override PartName="/xl/pivotTables/pivotTable1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml"/></Types>')
    parts["[Content_Types].xml"] = types.encode("utf-8")

    # Sheet: the two cells of a values-only PivotTable (Total Qty of 1000 rows with Qty 1..5 is 3000), and the relationship to the table.
    sheet = text("xl/worksheets/sheet1.xml")
    sheet = sheet.replace('<dimension ref="A1"/>', '<dimension ref="A3:A4"/>').replace(
        "<sheetData/>", '<sheetData><row r="3"><c r="A3" t="inlineStr"><is><t>Total Qty</t></is></c></row><row r="4"><c r="A4"><v>3000</v></c></row></sheetData>')
    parts["xl/worksheets/sheet1.xml"] = sheet.encode("utf-8")
    parts["xl/worksheets/_rels/sheet1.xml.rels"] = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable" Target="../pivotTables/pivotTable1.xml"/></Relationships>').encode("utf-8")
    parts["xl/pivotTables/_rels/pivotTable1.xml.rels"] = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition" Target="../pivotCache/pivotCacheDefinition1.xml"/></Relationships>').encode("utf-8")
    parts["xl/pivotCache/pivotCacheDefinition1.xml"] = cache.encode("utf-8")
    parts["xl/pivotTables/pivotTable1.xml"] = pivot.encode("utf-8")
    for name in ("xl/worksheets/_rels/sheet1.xml.rels", "xl/pivotTables/_rels/pivotTable1.xml.rels", "xl/pivotCache/pivotCacheDefinition1.xml", "xl/pivotTables/pivotTable1.xml"):
        if name not in order:
            order.append(name)

    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as z:
        for name in order:
            z.writestr(name, parts[name])
    # Every XML part must be well-formed.
    import xml.dom.minidom
    for name in order:
        if name.endswith((".xml", ".rels")):
            xml.dom.minidom.parseString(parts[name])
    print(f"{out_path}: {variant}, perspective={use_perspective}, cache hierarchies={len(kept)}")


if __name__ == "__main__":
    main()
