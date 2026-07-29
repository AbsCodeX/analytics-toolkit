# =========================================================================
# SANITIZED TEMPLATE — org-specific values were replaced with placeholders.
# Search for: YourOrg, YourUser, Lastname01..NN, C:\path\to\...  and
# fill in your own values before running. See README.md for the full map.
# =========================================================================
"""
surgical_xlsx.py — in-place, formula-safe editor for the Master Wave File.

WHY THIS EXISTS
---------------
openpyxl cannot round-trip the Master Wave File. Its DATA sheet contains
~7,599 shared formulas (t="shared"), ~15,198 single-cell array formulas
(t="array", using _xlfn.LET / _xlpm parameters and Table1[[#This Row]]
structured refs), 8 Table1 calculated columns, fragmented conditional-
formatting sqref ranges, and a list data-validation. openpyxl.save() rewrites
the ENTIRE workbook from its object model and mangles or drops every one of
those, so Excel "repairs" the file on open, the table ref balloons, and empty
rows render unbanded ("white at the bottom").

This module never instantiates openpyxl on the Master. It loads the .xlsx as
a zip of byte parts, edits ONLY the specific value cells the update logic
changes inside the DATA worksheet XML, and writes every other part
(formulas, styles, CF, validation, theme, drawings, other sheets) back
byte-for-byte identical. The only constructed XML is the handful of cells we
intentionally change plus narrowly-scoped wiring (Table1 ref, dimension,
Run Log row, DASHBOARD date). calcChain.xml is dropped cleanly (part +
[Content_Types] override + workbook rel) — Excel rebuilds it silently with
no repair prompt; leaving a stale calcChain is itself a repair trigger.

SCOPE / SAFETY DECISIONS
------------------------
* Value cells only. Columns carrying a formula are never written — set_value
  refuses to touch any cell that contains <f>. Formula columns are detected
  per-cell, not by letter (letters shift when DATA columns are added/removed,
  as on 2026-07-07).
* Bottom-append (promoting New Users) is supported and safe: appending rows
  past the last one needs no renumbering of existing rows. Appended rows get
  the row-2 formula cells cloned so the calc columns populate.
* Mid-sheet row deletion is NOT performed in XML. The file's fragmented CF
  sqref ranges and shared-formula masters cannot be renumbered safely by
  inference (this is exactly the class of edit that corrupted this file on
  2026-05-13). In practice the maintained Master has zero blank rows; if any
  are found the caller is warned to delete them in Excel instead.
"""

from __future__ import annotations

import re
import shutil
import time
import zipfile
from html import escape, unescape
from pathlib import Path

# ---------------------------------------------------------------------------
# Column-letter helpers
# ---------------------------------------------------------------------------

def col_to_num(col: str) -> int:
    n = 0
    for ch in col:
        n = n * 26 + (ord(ch) - 64)
    return n


def num_to_col(n: int) -> str:
    s = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


_CELL_RE_TMPL = r'<c r="{ref}"(?: [^>]*)?(?:/>|>.*?</c>)'
_ANYCELL_RE = re.compile(r'<c r="([A-Z]+)(\d+)"((?: [^>]*)?)(?:/>|>(.*?)</c>)', re.S)
_ROW_RE = re.compile(r'<row r="(\d+)"([^>]*)>(.*?)</row>', re.S)
_STYLE_RE = re.compile(r'\bs="(\d+)"')


def _xml_text(s: str) -> str:
    """Escape a Python string for use inside an XML <t> element."""
    return escape(s, quote=False)


# ---------------------------------------------------------------------------
# Package
# ---------------------------------------------------------------------------

class SurgicalWorkbook:
    """A Master Wave File loaded as raw byte parts, edited in place."""

    def __init__(self, path: Path):
        self.path = Path(path)
        with zipfile.ZipFile(self.path, "r") as zf:
            # Preserve original member order so the repackaged zip stays as
            # close to the original as possible.
            self._order = [i.filename for i in zf.infolist()]
            self.parts: dict[str, bytes] = {n: zf.read(n) for n in self._order}

        self._shared: list[str] = self._read_shared_strings()
        # Reverse lookup for interning new shared strings (first index wins).
        # _new_shared holds strings appended this session (indices continue
        # past the original table); _added_shared_refs counts shared-string
        # cell references written so the sst <count> can be bumped on save.
        self._shared_lookup: dict[str, int] = {}
        for _i, _s in enumerate(self._shared):
            self._shared_lookup.setdefault(_s, _i)
        self._new_shared: list[str] = []
        self._added_shared_refs: int = 0
        self.sheet_part: dict[str, str] = self._map_sheets()

        data_part = self.sheet_part.get("DATA")
        if not data_part:
            raise ValueError("Workbook has no sheet named 'DATA'")
        self._data_part = data_part

        xml = self.parts[data_part].decode("utf-8")
        m = re.search(r"<sheetData>(.*)</sheetData>", xml, re.S)
        if not m:
            raise ValueError("DATA sheet has no <sheetData> block")
        self._pre = xml[: m.start()]
        self._post = xml[m.end():]

        # Ordered list of [row_num, attrs, body]; body is the raw inner XML.
        self.rows: list[list] = []
        self._row_index: dict[int, int] = {}
        for rm in _ROW_RE.finditer(m.group(1)):
            rn = int(rm.group(1))
            self._row_index[rn] = len(self.rows)
            self.rows.append([rn, rm.group(2), rm.group(3)])

        self.max_row = max((r[0] for r in self.rows), default=1)
        self.headers, self.col_of = self._read_headers()
        self._col_style = self._infer_col_styles()

    # -- package wiring ----------------------------------------------------

    def _read_shared_strings(self) -> list[str]:
        raw = self.parts.get("xl/sharedStrings.xml")
        if not raw:
            return []
        txt = raw.decode("utf-8")
        out: list[str] = []
        for si in re.finditer(r"<si>(.*?)</si>", txt, re.S):
            inner = si.group(1)
            # Concatenate every <t> (handles rich-text runs).
            parts = re.findall(r"<t[^>]*>(.*?)</t>", inner, re.S)
            out.append(unescape("".join(parts)))
        return out

    def _map_sheets(self) -> dict[str, str]:
        wb = self.parts["xl/workbook.xml"].decode("utf-8")
        rels = self.parts["xl/_rels/workbook.xml.rels"].decode("utf-8")
        rid_to_target = dict(
            re.findall(
                r'Id="(rId\d+)"[^>]*Target="(worksheets/sheet\d+\.xml)"', rels
            )
        )
        out: dict[str, str] = {}
        for m in re.finditer(r'<sheet name="([^"]*)"[^>]*r:id="(rId\d+)"', wb):
            tgt = rid_to_target.get(m.group(2))
            if tgt:
                out[unescape(m.group(1))] = "xl/" + tgt
        return out

    # -- cell primitives ---------------------------------------------------

    def _decode_cell(self, attrs: str, inner: str):
        tm = re.search(r'\bt="(\w+)"', attrs)
        t = tm.group(1) if tm else "n"
        vm = re.search(r"<v>(.*?)</v>", inner, re.S)
        if t == "s":
            if vm is None:
                return None
            return self._shared_at(int(vm.group(1)))
        if t == "b":
            return vm is not None and vm.group(1) == "1"
        if t in ("str",):
            return unescape(vm.group(1)) if vm else None
        if t == "inlineStr":
            im = re.search(r"<t[^>]*>(.*?)</t>", inner, re.S)
            return unescape(im.group(1)) if im else None
        if vm is None:
            return None
        s = vm.group(1)
        try:
            f = float(s)
            return int(f) if f.is_integer() else f
        except ValueError:
            return s

    def _shared_at(self, i: int):
        """Resolve a shared-string index across the original table and any
        strings interned this session."""
        if i < len(self._shared):
            return self._shared[i]
        j = i - len(self._shared)
        if 0 <= j < len(self._new_shared):
            return self._new_shared[j]
        return None

    def _intern_shared(self, s: str) -> int:
        """Return the shared-string index for `s`, appending a new entry to the
        table if it is not already present."""
        if s in self._shared_lookup:
            return self._shared_lookup[s]
        idx = len(self._shared) + len(self._new_shared)
        self._new_shared.append(s)
        self._shared_lookup[s] = idx
        return idx

    def _find_cell(self, body: str, ref: str):
        """Return (match_obj) for <c r="ref" ...> in body, or None."""
        return re.search(_CELL_RE_TMPL.format(ref=re.escape(ref)), body, re.S)

    def get_value(self, row: int, col: str):
        idx = self._row_index.get(row)
        if idx is None:
            return None
        body = self.rows[idx][2]
        cm = self._find_cell(body, f"{col}{row}")
        if not cm:
            return None
        full = cm.group(0)
        am = re.match(r'<c r="[A-Z]+\d+"((?: [^>]*)?)(?:/>|>(.*)</c>)', full, re.S)
        attrs = am.group(1) or ""
        inner = am.group(2) or ""
        return self._decode_cell(attrs, inner)

    def cell_has_formula(self, row: int, col: str) -> bool:
        idx = self._row_index.get(row)
        if idx is None:
            return False
        cm = self._find_cell(self.rows[idx][2], f"{col}{row}")
        return bool(cm and "<f" in cm.group(0))

    def _build_cell(self, ref: str, value, style: str) -> str:
        s_attr = style or ""
        if value is None or (isinstance(value, str) and value == ""):
            return ""  # caller removes the cell entirely
        if isinstance(value, bool):
            return f'<c r="{ref}"{s_attr} t="b"><v>{1 if value else 0}</v></c>'
        if isinstance(value, (int, float)):
            num = int(value) if isinstance(value, float) and value.is_integer() else value
            return f'<c r="{ref}"{s_attr}><v>{num}</v></c>'
        txt = _xml_text(str(value))
        return (
            f'<c r="{ref}"{s_attr} t="inlineStr">'
            f'<is><t xml:space="preserve">{txt}</t></is></c>'
        )

    def set_value(self, row: int, col: str, value, shared: bool = False) -> bool:
        """
        Set a DATA value cell. Refuses formula cells. Returns True if the
        stored value actually changed.

        When `shared=True` and the value is a non-empty string, the cell is
        written as a shared-string reference (t="s") instead of an inline
        string (t="inlineStr"). This matches the native format Excel uses for
        text and is required for cells inserted into previously-empty columns:
        Excel's open-time repair has been observed to drop or mis-map large
        numbers of inline-string cells, whereas shared-string cells survive
        untouched. Numbers, booleans and None/blank are written the same way
        regardless of this flag.
        """
        idx = self._row_index.get(row)
        if idx is None:
            raise KeyError(f"row {row} not present on DATA")
        ref = f"{col}{row}"
        body = self.rows[idx][2]
        cm = self._find_cell(body, ref)

        if cm and "<f" in cm.group(0):
            raise ValueError(
                f"refusing to overwrite formula cell {ref} (formula column)"
            )

        # No-op guard: compare decoded current value to target.
        if cm:
            am = re.match(
                r'<c r="[A-Z]+\d+"((?: [^>]*)?)(?:/>|>(.*)</c>)',
                cm.group(0), re.S,
            )
            cur = self._decode_cell(am.group(1) or "", am.group(2) or "")
        else:
            cur = None
        norm_new = value if value not in ("",) else None
        if cur == norm_new:
            return False

        # Preserve an existing cell's style; otherwise use the column's.
        if cm:
            sm = _STYLE_RE.search(cm.group(0).split(">", 1)[0])
            style = f' s="{sm.group(1)}"' if sm else self._col_style.get(col, "")
        else:
            style = self._col_style.get(col, "")

        if shared and isinstance(norm_new, str) and norm_new != "":
            sidx = self._intern_shared(norm_new)
            self._added_shared_refs += 1
            new_cell = f'<c r="{ref}"{style} t="s"><v>{sidx}</v></c>'
        else:
            new_cell = self._build_cell(ref, norm_new, style)

        if cm:
            new_body = body[: cm.start()] + new_cell + body[cm.end():]
        else:
            new_body = self._insert_cell(body, col, new_cell)

        self.rows[idx][2] = new_body
        return True

    def _insert_cell(self, body: str, col: str, new_cell: str) -> str:
        if not new_cell:
            return body
        target = col_to_num(col)
        pos = None
        for m in _ANYCELL_RE.finditer(body):
            if col_to_num(m.group(1)) > target:
                pos = m.start()
                break
        if pos is None:
            return body + new_cell
        return body[:pos] + new_cell + body[pos:]

    # -- headers / styles --------------------------------------------------

    def _read_headers(self):
        idx = self._row_index.get(1)
        headers: dict[str, str] = {}  # stripped name -> col letter
        if idx is None:
            return headers, {}
        body = self.rows[idx][2]
        for m in _ANYCELL_RE.finditer(body):
            col = m.group(1)
            val = self._decode_cell(m.group(3) or "", m.group(4) or "")
            if isinstance(val, str) and val.strip():
                headers[val.strip()] = col
        col_of = {name: c for name, c in headers.items()}
        return headers, col_of

    def iter_header_cells(self):
        """Yield (col_letter, value) for every populated cell in row 1, in
        column order. Public accessor so callers can validate the header
        (e.g. detect duplicate/empty names) without reaching into internals."""
        idx = self._row_index.get(1)
        if idx is None:
            return
        for m in _ANYCELL_RE.finditer(self.rows[idx][2]):
            yield m.group(1), self._decode_cell(m.group(3) or "", m.group(4) or "")

    def _infer_col_styles(self) -> dict[str, str]:
        """Most representative ' s="N"' per column, sampled from data rows."""
        styles: dict[str, str] = {}
        for _, _, body in self.rows[1:60]:
            for m in _ANYCELL_RE.finditer(body):
                col = m.group(1)
                if col in styles:
                    continue
                sm = _STYLE_RE.search(m.group(3) or "")
                if sm:
                    styles[col] = f' s="{sm.group(1)}"'
        return styles

    # -- row append --------------------------------------------------------

    def append_row(self, values: dict[str, object]) -> int:
        """
        Append a new DATA row at the bottom. `values` maps column letter ->
        value. Formula cells (detected per-cell) are cloned from row 2 with the
        row number substituted so the calc columns populate. Returns new row num.
        """
        new_rn = self.max_row + 1
        tmpl_idx = self._row_index.get(2)
        tmpl_body = self.rows[tmpl_idx][2] if tmpl_idx is not None else ""

        cells: dict[int, str] = {}

        # Clone formula cells verbatim (structured refs are row-relative;
        # array refs and r= just need the new row number). Cached <v> dropped
        # so Excel recomputes (calcChain is removed anyway).
        for fm in _ANYCELL_RE.finditer(tmpl_body):
            seg = fm.group(0)
            if "<f" not in seg:
                continue
            col = fm.group(1)
            cnum = col_to_num(col)
            attrs = fm.group(3) or ""
            inner = fm.group(4) or ""
            fmt = re.search(r"<f.*?</f>|<f[^>]*/>", inner, re.S)
            if not fmt:
                continue
            ftxt = fmt.group(0)
            if 't="shared"' in ftxt:
                # A shared formula cloned as-is would create a second master
                # for the same si= (and keep the template row's cell refs),
                # which corrupts the group. Detach it: drop t/ref/si, keep the
                # literal formula, and renumber the template-row (row 2) cell
                # refs to this row. Structured Table1[...] refs are row-
                # relative and need no change; bare A1-style refs do.
                if "</f>" not in ftxt:
                    continue  # shared child with no body — nothing to clone
                inner_f = re.search(r"<f[^>]*>(.*?)</f>", ftxt, re.S).group(1)
                inner_f = re.sub(
                    r"(?<![A-Za-z0-9_$])([A-Z]{1,3})2(?![0-9])",
                    rf"\g<1>{new_rn}",
                    inner_f,
                )
                ftxt = f"<f>{inner_f}</f>"
            else:
                # Re-anchor a single-cell array ref="COL2" to this row.
                ftxt = re.sub(r'ref="[A-Z]+2"', f'ref="{col}{new_rn}"', ftxt)
            cells[cnum] = f'<c r="{col}{new_rn}"{attrs}>{ftxt}</c>'

        # Value cells.
        for col, val in values.items():
            if val is None or (isinstance(val, str) and not str(val).strip()):
                continue
            cnum = col_to_num(col)
            if cnum in cells:  # never clobber a formula column
                continue
            style = self._col_style.get(col, "")
            cells[cnum] = self._build_cell(f"{col}{new_rn}", val, style)

        body = "".join(cells[k] for k in sorted(cells))
        spans = f'spans="1:{col_to_num(max(self.col_of.values(), key=col_to_num))}"'
        attrs = f' {spans}'
        self.rows.append([new_rn, attrs, body])
        self._row_index[new_rn] = len(self.rows) - 1
        self.max_row = new_rn
        return new_rn

    def count_blank_rows(self, uid_col: str) -> list[int]:
        """Data rows whose UID cell is empty/whitespace (should be 0)."""
        out = []
        for rn, _, _ in self.rows:
            if rn == 1:
                continue
            v = self.get_value(rn, uid_col)
            if not str(v or "").strip():
                out.append(rn)
        return out

    # -- read-only helpers on other sheets ---------------------------------

    def read_column(self, sheet_name: str, col: str, min_row: int = 1) -> list:
        part = self.sheet_part.get(sheet_name)
        if not part:
            return []
        xml = self.parts[part].decode("utf-8")
        out = []
        for rm in _ROW_RE.finditer(xml):
            rn = int(rm.group(1))
            if rn < min_row:
                continue
            cm = re.search(
                _CELL_RE_TMPL.format(ref=re.escape(f"{col}{rn}")),
                rm.group(3), re.S,
            )
            if not cm:
                continue
            am = re.match(
                r'<c r="[A-Z]+\d+"((?: [^>]*)?)(?:/>|>(.*)</c>)',
                cm.group(0), re.S,
            )
            val = self._decode_cell(am.group(1) or "", am.group(2) or "")
            if val is not None:
                out.append(val)
        return out

    def set_other_cell(self, sheet_name: str, ref: str, value) -> None:
        """Set a single cell on a non-DATA sheet (e.g. DASHBOARD!C3)."""
        part = self.sheet_part.get(sheet_name)
        if not part:
            return
        xml = self.parts[part].decode("utf-8")
        col = re.match(r"[A-Z]+", ref).group(0)
        rn = int(re.match(r"[A-Z]+(\d+)", ref).group(1))
        cell = self._build_cell(ref, value, self._col_style.get(col, ""))

        cm = re.search(_CELL_RE_TMPL.format(ref=re.escape(ref)), xml, re.S)
        if cm:
            xml = xml[: cm.start()] + cell + xml[cm.end():]
        else:
            rm = re.search(rf'<row r="{rn}"([^>]*)>(.*?)</row>', xml, re.S)
            if rm:
                new_body = self._insert_cell(rm.group(2), col, cell)
                xml = (
                    xml[: rm.start()]
                    + f'<row r="{rn}"{rm.group(1)}>{new_body}</row>'
                    + xml[rm.end():]
                )
            else:
                return  # row absent — skip rather than guess structure
        self.parts[part] = xml.encode("utf-8")

    # -- simple (formula-free) sheets: staging, Run Log -------------------
    #
    # These helpers parse/mutate/reserialize an entire worksheet's rows. They
    # are ONLY safe on sheets with no formulas, conditional formatting, merged
    # cells, or data validation tied to row positions. The Master's
    # 'New Users to Add' and 'Run Log' sheets qualify (verified: zero
    # formulas / CF / validation). Never call these on DATA.

    def _parse_simple(self, sheet_name: str):
        part = self.sheet_part.get(sheet_name)
        if not part:
            return None
        xml = self.parts[part].decode("utf-8")
        m = re.search(r"<sheetData>(.*)</sheetData>", xml, re.S)
        if not m:
            return None
        pre, post = xml[: m.start()], xml[m.end():]
        rows = [
            [int(rm.group(1)), rm.group(2), rm.group(3)]
            for rm in _ROW_RE.finditer(m.group(1))
        ]
        return part, pre, post, rows

    def _serialize_simple(self, part, pre, post, rows) -> None:
        last = max((r[0] for r in rows), default=1)
        body = "".join(f'<row r="{rn}"{a}>{b}</row>' for rn, a, b in rows)
        xml = pre + "<sheetData>" + body + "</sheetData>" + post
        xml = re.sub(
            r'(<dimension ref="[A-Z]+1:)([A-Z]+)\d+("/>)',
            lambda m: f"{m.group(1)}{m.group(2)}{last}{m.group(3)}",
            xml, count=1,
        )
        self.parts[part] = xml.encode("utf-8")

    def read_simple_sheet(self, sheet_name: str):
        """Return (header{name->col}, [(row_num, {col: value})]) for a sheet."""
        parsed = self._parse_simple(sheet_name)
        if not parsed:
            return {}, []
        _, _, _, rows = parsed
        header: dict[str, str] = {}
        data = []
        for rn, _, b in rows:
            cells = {}
            for m in _ANYCELL_RE.finditer(b):
                cells[m.group(1)] = self._decode_cell(
                    m.group(3) or "", m.group(4) or ""
                )
            if rn == 1:
                for col, v in cells.items():
                    if isinstance(v, str) and v.strip():
                        header[v.strip()] = col
            else:
                data.append((rn, cells))
        return header, data

    def delete_simple_rows(self, sheet_name: str, row_nums: set[int]) -> int:
        """Delete rows from a formula-free sheet and renumber the rest."""
        parsed = self._parse_simple(sheet_name)
        if not parsed:
            return 0
        part, pre, post, rows = parsed
        kept = [r for r in rows if r[0] not in row_nums]
        removed = len(rows) - len(kept)
        if not removed:
            return 0
        new_rows = []
        next_rn = 1
        for rn, attrs, body in kept:
            if rn == 1:
                new_rows.append([1, attrs, body])
                next_rn = 2
                continue
            tgt = next_rn
            next_rn += 1
            if rn != tgt:
                body = re.sub(
                    r'(<c r=")([A-Z]+)\d+(")',
                    lambda m: f"{m.group(1)}{m.group(2)}{tgt}{m.group(3)}",
                    body,
                )
            new_rows.append([tgt, attrs, body])
        self._serialize_simple(part, pre, post, new_rows)
        return removed

    def rewrite_simple_sheet(
        self, sheet_name: str, headers: list[str],
        data_rows: list[dict[str, object]],
    ) -> None:
        """
        Replace a formula-free sheet's contents: row 1 = `headers` in columns
        A, B, C…; rows 2+ = `data_rows` (each a {col-letter: value} dict).
        Used to self-heal the Run Log into a clean canonical layout. Preserves
        the verbatim pre/post (sheetViews, pageMargins, etc.).
        """
        parsed = self._parse_simple(sheet_name)
        if not parsed:
            raise KeyError(f"sheet '{sheet_name}' not found")
        part, pre, post, old_rows = parsed

        # Preserve the existing header style (the Run Log header is bold) so a
        # self-heal rewrite doesn't strip its formatting.
        hdr_style = ""
        for rn, _, b in old_rows:
            if rn == 1:
                sm = re.search(r"<c r=\"[A-Z]+1\"([^>]*)>", b)
                if sm:
                    s = _STYLE_RE.search(sm.group(1))
                    if s:
                        hdr_style = f' s="{s.group(1)}"'
                break

        rows: list[list] = []
        hdr_cells = "".join(
            self._build_cell(f"{num_to_col(i + 1)}1", h, hdr_style)
            for i, h in enumerate(headers)
        )
        rows.append([1, "", hdr_cells])
        for ridx, dr in enumerate(data_rows, start=2):
            ordered = sorted(dr.items(), key=lambda kv: col_to_num(kv[0]))
            body = "".join(
                self._build_cell(f"{c}{ridx}", v, "")
                for c, v in ordered
                if not (v is None or (isinstance(v, str) and v == ""))
            )
            rows.append([ridx, "", body])
        self._serialize_simple(part, pre, post, rows)

    def append_simple_row(self, sheet_name: str, cells: dict[str, object]) -> int:
        """Append one row (col-letter -> value) to a formula-free sheet."""
        parsed = self._parse_simple(sheet_name)
        if not parsed:
            raise KeyError(f"sheet '{sheet_name}' not found")
        part, pre, post, rows = parsed
        new_rn = max((r[0] for r in rows), default=1) + 1
        ordered = sorted(cells.items(), key=lambda kv: col_to_num(kv[0]))
        body = "".join(
            self._build_cell(f"{c}{new_rn}", v, self._col_style.get(c, ""))
            for c, v in ordered
            if not (v is None or (isinstance(v, str) and v == ""))
        )
        rows.append([new_rn, "", body])
        self._serialize_simple(part, pre, post, rows)
        return new_rn

    def has_sheet(self, name: str) -> bool:
        return name in self.sheet_part

    # -- Table1 ref / dimension -------------------------------------------

    def sync_table_and_dimension(self) -> None:
        """Extend Table1 ref, its autoFilter, and the sheet dimension to the
        true last data row. Only the trailing row number is rewritten."""
        last = self.max_row

        # dimension lives in the verbatim _pre block.
        self._pre = re.sub(
            r'(<dimension ref="[A-Z]+1:)([A-Z]+)\d+("/>)',
            lambda m: f"{m.group(1)}{m.group(2)}{last}{m.group(3)}",
            self._pre, count=1,
        )

        for name, raw in list(self.parts.items()):
            if not re.match(r"xl/tables/table\d+\.xml$", name):
                continue
            txt = raw.decode("utf-8")
            if 'name="Table1"' not in txt:
                continue

            def _ref_sub(m):
                start, endcol = m.group(1), m.group(2)
                return f'ref="{start}:{endcol}{last}"'

            txt2 = re.sub(
                r'ref="([A-Z]+\d+):([A-Z]+)\d+"', _ref_sub, txt, count=1
            )
            txt2 = re.sub(
                r'(<autoFilter ref="[A-Z]+\d+:)([A-Z]+)\d+(")',
                lambda m: f"{m.group(1)}{m.group(2)}{last}{m.group(3)}",
                txt2, count=1,
            )
            self.parts[name] = txt2.encode("utf-8")

    # -- calcChain drop ----------------------------------------------------

    def drop_calc_chain(self) -> None:
        """Remove calcChain.xml + its [Content_Types] override + workbook rel.
        Excel rebuilds the chain silently; a stale one triggers repair."""
        if "xl/calcChain.xml" not in self.parts:
            return
        del self.parts["xl/calcChain.xml"]

        ct = self.parts["[Content_Types].xml"].decode("utf-8")
        ct = re.sub(
            r'<Override PartName="/xl/calcChain\.xml"[^>]*/>', "", ct
        )
        self.parts["[Content_Types].xml"] = ct.encode("utf-8")

        rels_name = "xl/_rels/workbook.xml.rels"
        rels = self.parts[rels_name].decode("utf-8")
        rels = re.sub(
            r'<Relationship [^>]*Target="calcChain\.xml"[^>]*/>', "", rels
        )
        self.parts[rels_name] = rels.encode("utf-8")

    # -- shared strings ----------------------------------------------------

    def _append_new_shared_strings(self) -> None:
        """Append any strings interned this session to xl/sharedStrings.xml,
        preserving every existing <si> byte-for-byte. Updates uniqueCount to
        the exact <si> total and bumps count by the shared refs written.
        Per ISO/IEC 29500: count = total string-cell references (advisory,
        Excel recomputes), uniqueCount = number of <si> entries."""
        if not self._new_shared:
            return
        name = "xl/sharedStrings.xml"
        raw = self.parts.get(name)
        if not raw:
            # No shared string table to extend; leave interned strings unused
            # rather than fabricate a part on inference.
            raise ValueError(
                "Workbook has no sharedStrings.xml to append shared strings to."
            )
        txt = raw.decode("utf-8")
        new_sis = "".join(
            f'<si><t xml:space="preserve">{_xml_text(s)}</t></si>'
            for s in self._new_shared
        )
        close = txt.rfind("</sst>")
        if close == -1:
            raise ValueError("sharedStrings.xml has no </sst> closing tag.")
        txt = txt[:close] + new_sis + txt[close:]

        unique_total = len(self._shared) + len(self._new_shared)
        txt = re.sub(
            r'(<sst\b[^>]*\buniqueCount=")\d+(")',
            rf"\g<1>{unique_total}\g<2>", txt, count=1,
        )
        txt = re.sub(
            r'(<sst\b[^>]*\bcount=")(\d+)(")',
            lambda m: f"{m.group(1)}{int(m.group(2)) + self._added_shared_refs}{m.group(3)}",
            txt, count=1,
        )
        self.parts[name] = txt.encode("utf-8")

    # -- serialize ---------------------------------------------------------

    def _rebuild_data_xml(self) -> bytes:
        buf = [self._pre, "<sheetData>"]
        for rn, attrs, body in self.rows:
            buf.append(f'<row r="{rn}"{attrs}>{body}</row>')
        buf.append("</sheetData>")
        buf.append(self._post)
        return "".join(buf).encode("utf-8")

    def save(self) -> None:
        self._append_new_shared_strings()
        self.parts[self._data_part] = self._rebuild_data_xml()

        tmp = self.path.with_suffix(".surgical.xlsx")
        ordered = list(self._order)
        # Keep any newly added parts (none expected) appended at the end;
        # skip parts we deleted (calcChain).
        for n in self.parts:
            if n not in ordered:
                ordered.append(n)
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
            for n in ordered:
                if n in self.parts:
                    zout.writestr(n, self.parts[n])

        for attempt in range(5):
            try:
                shutil.move(str(tmp), str(self.path))
                return
            except PermissionError:
                if attempt < 4:
                    time.sleep(1)
                else:
                    tmp.unlink(missing_ok=True)
                    raise
