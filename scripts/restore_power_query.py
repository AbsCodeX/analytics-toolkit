# =========================================================================
# SANITIZED TEMPLATE — org-specific values were replaced with placeholders.
# Search for: YourOrg, YourUser, Lastname01..NN, C:\path\to\...  and
# fill in your own values before running. See README.md for the full map.
# =========================================================================
"""
restore_power_query.py

openpyxl cannot round-trip Power Query ("Get & Transform"). On save it silently
drops every Power Query part: xl/connections.xml, xl/queryTables/*, the
DataMashup in customXml/*, the queryTable-backed table parts, and (separately)
xl/metadata.xml. Excel then "repairs" the file on open, removing the orphaned
"External data range" parts and killing the ability to refresh those queries.

This module re-grafts those parts from a Power-Query-intact source workbook
(e.g. the pre-run backup the update script already makes) onto the
openpyxl-written output, keeping the openpyxl data edits while restoring
refreshable Power Query.

Everything written is taken VERBATIM from the source workbook. The only
constructed XML is package wiring ([Content_Types].xml entries, workbook and
worksheet relationships, the worksheet <tableParts> overlay) — and every
content-type string / relationship Type used is lifted byte-for-byte out of
the source package itself, never inferred. Tables are matched by their
internal `name`/`tableType`, never by file number (openpyxl renumbers table
parts — matching by filename is exactly the mistake that corrupts this file).

Usage:
    from restore_power_query import restore_power_query
    restore_power_query(pq_source_path, openpyxl_output_path)

Idempotent: if the target already contains the query tables, it is left as-is.
"""

from __future__ import annotations

import re
import shutil
import zipfile
from pathlib import Path

REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
QT_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/queryTable"
TABLE_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table"


def _read_all(path: Path) -> dict[str, bytes]:
    with zipfile.ZipFile(path, "r") as zf:
        return {i.filename: zf.read(i.filename) for i in zf.infolist()}


def _sheet_name_to_part(parts: dict[str, bytes]) -> dict[str, str]:
    """Map worksheet display name -> 'xl/worksheets/sheetN.xml' for a package."""
    wb = parts["xl/workbook.xml"].decode("utf-8")
    rels = parts["xl/_rels/workbook.xml.rels"].decode("utf-8")
    rid_to_target = dict(
        re.findall(r'Id="(rId\d+)"[^>]*Target="(worksheets/sheet\d+\.xml)"', rels)
    )
    out = {}
    for m in re.finditer(r'<sheet name="([^"]*)"[^>]*r:id="(rId\d+)"', wb):
        tgt = rid_to_target.get(m.group(2))
        if tgt:
            out[m.group(1)] = "xl/" + tgt
    return out


def _content_type_override(ct_xml: str, part_name: str) -> str:
    """Return the exact <Override> element for part_name from a Content_Types xml."""
    m = re.search(
        r'<Override PartName="%s"[^>]*/>' % re.escape(part_name), ct_xml
    )
    if not m:
        raise ValueError(f"No Content_Types Override for {part_name} in source package")
    return m.group(0)


def _table_content_type(ct_xml: str) -> str:
    """Lift the spreadsheetml.table ContentType string from an existing override."""
    m = re.search(
        r'<Override PartName="/xl/tables/table\d+\.xml" ContentType="([^"]*)"/>',
        ct_xml,
    )
    if not m:
        raise ValueError("Could not find an existing table Override to copy ContentType from")
    return m.group(1)


def restore_power_query(pq_source: Path, target: Path) -> dict:
    """
    Graft Power Query parts from pq_source onto target (an openpyxl output).
    Modifies target in place (via a temp file + atomic replace).

    Returns a summary dict. Raises on any structural inconsistency rather than
    writing a half-wired package.
    """
    pq_source = Path(pq_source)
    target = Path(target)

    src = _read_all(pq_source)
    tgt = _read_all(target)

    # --- Identify the queryTable-backed tables in the SOURCE -----------------
    # Match by tableType="queryTable" (internal attribute), never by filename.
    src_qt_tables = {}  # src table part -> (table_xml_bytes, name)
    for name in src:
        if re.match(r"xl/tables/table\d+\.xml$", name):
            txt = src[name].decode("utf-8")
            if 'tableType="queryTable"' in txt:
                tname = re.search(r'\bname="([^"]*)"', txt).group(1)
                src_qt_tables[name] = (src[name], tname)

    if not src_qt_tables:
        return {"status": "noop", "reason": "no Power Query tables in source"}

    # Already restored? (target has a queryTable table) -> idempotent no-op.
    for name in tgt:
        if re.match(r"xl/tables/table\d+\.xml$", name):
            if 'tableType="queryTable"' in tgt[name].decode("utf-8"):
                return {"status": "noop", "reason": "target already has Power Query"}

    # For each source query table: find the worksheet that references it and
    # the queryTable part it is bound to (via xl/tables/_rels/<table>.rels).
    src_sheet_by_part = _sheet_name_to_part(src)
    src_part_to_sheetname = {v: k for k, v in src_sheet_by_part.items()}
    tgt_sheet_by_name = _sheet_name_to_part(tgt)

    grafts = []  # list of dicts describing each query table to transplant
    for src_tbl_part, (tbl_bytes, tname) in src_qt_tables.items():
        tbl_file = src_tbl_part.split("/")[-1]                     # tableN.xml
        tbl_rels = f"xl/tables/_rels/{tbl_file}.rels"
        if tbl_rels not in src:
            raise ValueError(f"Source missing {tbl_rels} for query table {tname}")
        qt_target = re.search(
            r'Target="\.\./(queryTables/queryTable\d+\.xml)"', src[tbl_rels].decode()
        )
        if not qt_target:
            raise ValueError(f"{tbl_rels} has no queryTable relationship")
        qt_part = "xl/" + qt_target.group(1)

        # Which worksheet references this table?
        owning_sheet_part = None
        for ws_rels_name, blob in src.items():
            if re.match(r"xl/worksheets/_rels/sheet\d+\.xml\.rels$", ws_rels_name):
                if f"tables/{tbl_file}" in blob.decode("utf-8"):
                    owning_sheet_part = "xl/worksheets/" + (
                        ws_rels_name.split("/")[-1].replace(".rels", "")
                    )
                    break
        if not owning_sheet_part:
            raise ValueError(f"No worksheet references table {tname} in source")

        sheet_name = src_part_to_sheetname.get(owning_sheet_part)
        if sheet_name not in tgt_sheet_by_name:
            raise ValueError(
                f"Source sheet '{sheet_name}' (for query {tname}) not found in target"
            )

        grafts.append({
            "name": tname,
            "src_table_part": src_tbl_part,
            "table_bytes": tbl_bytes,
            "qt_part": qt_part,
            "qt_bytes": src[qt_part],
            "tgt_sheet_part": tgt_sheet_by_name[sheet_name],
            "sheet_name": sheet_name,
        })

    # --- Allocate new table numbers above target's current max --------------
    existing_nums = [
        int(re.search(r"table(\d+)\.xml$", n).group(1))
        for n in tgt
        if re.match(r"xl/tables/table\d+\.xml$", n)
    ]
    next_num = (max(existing_nums) + 1) if existing_nums else 1

    src_ct = src["[Content_Types].xml"].decode("utf-8")
    tgt_ct = tgt["[Content_Types].xml"].decode("utf-8")
    table_ct = _table_content_type(tgt_ct)

    new_parts: dict[str, bytes] = {}
    ct_overrides: list[str] = []
    seen_qt_parts: set[str] = set()

    for g in grafts:
        new_tbl = f"xl/tables/table{next_num}.xml"
        new_tbl_rels = f"xl/tables/_rels/table{next_num}.xml.rels"
        next_num += 1

        # Table part + its rel -> queryTable (queryTable keeps its source name,
        # so the verbatim relative target stays valid).
        new_parts[new_tbl] = g["table_bytes"]
        qt_file = g["qt_part"].split("/")[-1]
        new_parts[new_tbl_rels] = (
            f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
            f'<Relationships xmlns="{REL_NS}">'
            f'<Relationship Id="rId1" Type="{QT_REL_TYPE}" '
            f'Target="../queryTables/{qt_file}"/></Relationships>'
        ).encode("utf-8")
        ct_overrides.append(
            f'<Override PartName="/{new_tbl}" ContentType="{table_ct}"/>'
        )

        # queryTable part (verbatim) + its Content_Types override (verbatim).
        if g["qt_part"] not in seen_qt_parts:
            new_parts[g["qt_part"]] = g["qt_bytes"]
            ct_overrides.append(_content_type_override(src_ct, "/" + g["qt_part"]))
            seen_qt_parts.add(g["qt_part"])

        # Worksheet rels: add (or extend) a relationship to the new table.
        ws_rels_name = (
            "xl/worksheets/_rels/"
            + g["tgt_sheet_part"].split("/")[-1]
            + ".rels"
        )
        if ws_rels_name in tgt:
            rels_xml = tgt[ws_rels_name].decode("utf-8")
            used = [int(x) for x in re.findall(r'Id="rId(\d+)"', rels_xml)]
            rid = f"rId{(max(used) + 1) if used else 1}"
            rels_xml = rels_xml.replace(
                "</Relationships>",
                f'<Relationship Id="{rid}" Type="{TABLE_REL_TYPE}" '
                f'Target="../tables/table{next_num - 1}.xml"/></Relationships>',
            )
            tgt[ws_rels_name] = rels_xml.encode("utf-8")
        else:
            rid = "rId1"
            new_parts[ws_rels_name] = (
                f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
                f'<Relationships xmlns="{REL_NS}">'
                f'<Relationship Id="{rid}" Type="{TABLE_REL_TYPE}" '
                f'Target="../tables/table{next_num - 1}.xml"/></Relationships>'
            ).encode("utf-8")

        # Inject <tableParts> into the target worksheet (after pageMargins,
        # before </worksheet>; append to an existing block if present).
        ws_xml = tgt[g["tgt_sheet_part"]].decode("utf-8")
        if "<tableParts" in ws_xml:
            ws_xml = re.sub(
                r'<tableParts count="(\d+)">',
                lambda m: f'<tableParts count="{int(m.group(1)) + 1}">',
                ws_xml,
                count=1,
            )
            ws_xml = ws_xml.replace(
                "</tableParts>", f'<tablePart r:id="{rid}"/></tableParts>', 1
            )
        else:
            ws_xml = ws_xml.replace(
                "</worksheet>",
                f'<tableParts count="1"><tablePart r:id="{rid}"/>'
                f"</tableParts></worksheet>",
            )
        tgt[g["tgt_sheet_part"]] = ws_xml.encode("utf-8")

    # --- Copy supporting parts verbatim: connections, DataMashup, metadata --
    workbook_rel_adds = []  # (Type, Target) lifted verbatim from source rels
    src_wb_rels = src["xl/_rels/workbook.xml.rels"].decode("utf-8")

    def _src_wb_rel(target_suffix: str):
        m = re.search(
            r'<Relationship [^>]*Type="([^"]*)"[^>]*Target="([^"]*%s)"[^>]*/>'
            % re.escape(target_suffix),
            src_wb_rels,
        )
        return (m.group(1), m.group(2)) if m else None

    for part, ct_partname, wb_rel_suffix in [
        ("xl/connections.xml", "/xl/connections.xml", "connections.xml"),
        ("xl/metadata.xml", "/xl/metadata.xml", "metadata.xml"),
        ("customXml/item1.xml", None, "../customXml/item1.xml"),
        ("customXml/itemProps1.xml", "/customXml/itemProps1.xml", None),
        ("customXml/_rels/item1.xml.rels", None, None),
    ]:
        if part in src:
            new_parts[part] = src[part]
            if ct_partname:
                ct_overrides.append(_content_type_override(src_ct, ct_partname))
            if wb_rel_suffix:
                rel = _src_wb_rel(wb_rel_suffix)
                if rel:
                    workbook_rel_adds.append(rel)

    # --- Patch [Content_Types].xml -----------------------------------------
    # De-dupe against anything already declared, then insert before </Types>.
    add_ct = [o for o in ct_overrides if o not in tgt_ct]
    tgt_ct = tgt_ct.replace("</Types>", "".join(add_ct) + "</Types>")
    tgt["[Content_Types].xml"] = tgt_ct.encode("utf-8")

    # --- Patch workbook.xml.rels (fresh, unused rIds) -----------------------
    wb_rels = tgt["xl/_rels/workbook.xml.rels"].decode("utf-8")
    used_ids = [int(x) for x in re.findall(r'Id="rId(\d+)"', wb_rels)]
    rid_n = (max(used_ids) + 1) if used_ids else 1
    additions = []
    for rtype, rtarget in workbook_rel_adds:
        if f'Target="{rtarget}"' in wb_rels:
            continue
        additions.append(
            f'<Relationship Id="rId{rid_n}" Type="{rtype}" Target="{rtarget}"/>'
        )
        rid_n += 1
    if additions:
        wb_rels = wb_rels.replace(
            "</Relationships>", "".join(additions) + "</Relationships>"
        )
        tgt["xl/_rels/workbook.xml.rels"] = wb_rels.encode("utf-8")

    # --- Write the repaired package ----------------------------------------
    tmp = target.with_suffix(".pqfix.xlsx")
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for fname, data in tgt.items():
            zout.writestr(fname, data)
        for fname, data in new_parts.items():
            zout.writestr(fname, data)

    import time
    for attempt in range(5):
        try:
            shutil.move(str(tmp), str(target))
            break
        except PermissionError:
            if attempt < 4:
                time.sleep(1)
            else:
                tmp.unlink(missing_ok=True)
                raise

    return {
        "status": "restored",
        "queries": [g["name"] for g in grafts],
        "parts_added": sorted(new_parts),
        "content_type_overrides_added": len(add_ct),
        "workbook_rels_added": len(additions),
    }


if __name__ == "__main__":
    import sys

    if len(sys.argv) != 3:
        print("usage: python restore_power_query.py <pq_source.xlsx> <openpyxl_output.xlsx>")
        sys.exit(1)
    result = restore_power_query(Path(sys.argv[1]), Path(sys.argv[2]))
    for k, v in result.items():
        print(f"{k}: {v}")
