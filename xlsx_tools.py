"""
xlsx_tools.py
=============

Shared worksheet styling helpers and the Excel formula-cache engine.

Previously copy-pasted (and drifted) across make_tables.py,
make_services_tables.py and make_quarterly_tables.py; this module is the
single source of truth.  The formula-cache engine saves a workbook with
live formulas via openpyxl, then patches cached ``<v>`` results into the
sheet XML so ``data_only=True`` readers (the report generators) see plain
numbers while Excel still shows recalculating formulas.
"""

import re
import zipfile

import openpyxl
from lxml import etree
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.utils.cell import coordinate_to_tuple

# ---------------------------------------------------------------------------
# Constants / styles
# ---------------------------------------------------------------------------
FONT = "Century Gothic"
FONT_SIZE = 10
HEADER_FONT_SIZE = 10
FONT_GRAY = "595959"

# Header styles - deep navy with white text
HDR_FILL = PatternFill(fill_type="solid", fgColor="1F3864")
HDR_FONT_COLOR = "FFFFFF"

# Alternating row fills (very subtle gray tones)
ROW_FILL_EVEN = PatternFill(fill_type="solid", fgColor="F2F2F2")
ROW_FILL_ODD = None  # No fill for odd rows (white)

# Kenya highlight (soft gold)
KENYA_FILL = PatternFill(fill_type="solid", fgColor="FFF2CC")

# Band/summary row (light blue-gray)
BAND_FILL = PatternFill(fill_type="solid", fgColor="D6DCE4")

# Border styles - thin bottom only for headers, no side borders
THIN = Side(style="thin", color="B4C6E7")  # Light blue-gray
MEDIUM = Side(style="medium", color="1F3864")  # Navy accent
HEADER_BORDER = Border(bottom=MEDIUM)
DATA_BORDER = Border(bottom=Side(style="thin", color="D6DCE4"))
TOTAL_BORDER = Border(top=Side(style="thin", color="808080"),
                      bottom=Side(style="double", color="808080"))

# Legacy border (kept for compatibility)
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

# Number formats
FMT_VALUE = "#,##0.0"
FMT_VALUE_INT = "#,##0"
FMT_SHARE = "0.0%"
FMT_GROWTH = "0.0%"


# ---------------------------------------------------------------------------
# Small helpers
# ---------------------------------------------------------------------------
def to_float(v):
    """Convert a raw cell value to float; None for blanks / '...' / garbage."""
    if v is None:
        return None
    if isinstance(v, (int, float)):
        f = float(v)
        return f if f == f else None  # NaN-safe
    s = str(v).strip().replace(",", "").replace(" ", "")
    if not s or s in ("...", "..", "-", "–", "—", "n/a", "N/A"):
        return None
    try:
        return float(s)
    except ValueError:
        return None


def set_widths(ws, widths):
    for col, w in widths.items():
        ws.column_dimensions[col].width = w


def style_cell(cell, bold=False, size=FONT_SIZE, fill=None, numfmt=None,
               wrap=False, align="center", border_style="data", italic=False):
    """Apply professional styling to a cell.

    border_style: 'header', 'data', 'total', or 'none'
    """
    cell.font = Font(name=FONT, size=size, bold=bold, italic=italic,
                     color="000000" if not bold else "1F3864")

    if border_style == "header":
        cell.border = HEADER_BORDER
    elif border_style == "total":
        cell.border = TOTAL_BORDER
    elif border_style == "none":
        cell.border = Border()
    else:
        cell.border = DATA_BORDER

    if fill is not None:
        cell.fill = fill
    if numfmt is not None:
        cell.number_format = numfmt
    cell.alignment = Alignment(horizontal=align, vertical="center",
                               wrap_text=wrap,
                               indent=1 if align == "left" else 0)


def style_header(cell, text, size=HEADER_FONT_SIZE, fill=HDR_FILL,
                 font_color=HDR_FONT_COLOR, align="center", wrap=True):
    """Style a header cell with professional dark header."""
    cell.value = text
    cell.font = Font(name=FONT, size=size, bold=True, color=font_color)
    cell.fill = fill
    cell.border = HEADER_BORDER
    cell.alignment = Alignment(horizontal=align, vertical="center",
                               wrap_text=wrap)


def get_row_fill(row_index):
    """Get alternating row fill for data rows (0-indexed)."""
    return ROW_FILL_EVEN if row_index % 2 == 0 else ROW_FILL_ODD


def put(ws, row, col, value, **kw):
    cell = ws.cell(row=row, column=col, value=value)
    style_cell(cell, **kw)
    return cell


def put_text(ws, row, col, text, bold=False, size=FONT_SIZE, fill=None,
             wrap=False, align="left", use_row_fill=True, italic=False):
    """Write text with optional alternating row fill."""
    if fill is None and use_row_fill:
        fill = get_row_fill(row)
    return put(ws, row, col, text, bold=bold, size=size, fill=fill,
               wrap=wrap, align=align, italic=italic)


def put_val(ws, row, col, value, bold=False, fill=None, fmt=FMT_VALUE,
            use_row_fill=True):
    """Write numeric value with optional alternating row fill."""
    if fill is None and use_row_fill:
        fill = get_row_fill(row)
    return put(ws, row, col, value, bold=bold, fill=fill, numfmt=fmt)


def put_share(ws, row, col, value, bold=False, fill=None, use_row_fill=True):
    """Write share percentage with optional alternating row fill."""
    if fill is None and use_row_fill:
        fill = get_row_fill(row)
    return put(ws, row, col, value, bold=bold, fill=fill, numfmt=FMT_SHARE)


def merge(ws, r1, c1, r2, c2):
    ws.merge_cells(start_row=r1, start_column=c1, end_row=r2, end_column=c2)


def _cell_ref(row, col):
    return f"{get_column_letter(col)}{row}"


def _put_formula(ws, cache, row, col, formula, cached, bold=False, fill=None,
                 numfmt=None):
    """Write an Excel formula and register the numeric result for caching.

    ``cache`` is a list of (cell_ref, value) pairs (single-sheet workbook).
    """
    put(ws, row, col, formula, bold=bold, fill=fill, numfmt=numfmt)
    cache.append((_cell_ref(row, col), cached))


def _put_formula_sheet(ws, cache, row, col, formula, cached, **kw):
    """Like _put_formula but registers the cache per sheet title (dict form)."""
    put(ws, row, col, formula, **kw)
    cache.setdefault(ws.title, {})[_cell_ref(row, col)] = cached


# ---------------------------------------------------------------------------
# Formula-cache engine (patch cached <v> results into the saved workbook)
# ---------------------------------------------------------------------------
def _sheet_order(path):
    """Sheet names in document order (openpyxl names the xml files
    sheet1.xml, sheet2.xml ... matching that order)."""
    with zipfile.ZipFile(path) as zin:
        xml = zin.read("xl/workbook.xml").decode("utf-8", "ignore")
    return re.findall(r'<sheet[^>]*name="([^"]+)"', xml)


def _cache_to_refs(order, cache):
    """Normalise a formula cache to {sheet_title: {cell_ref: value}}.

    A dict cache is used as-is (minus None values).  A list cache holds pairs
    for one sheet and is mapped to the workbook's first sheet -- whatever its
    title -- so callers never depend on the default "Sheet1" name.
    """
    if isinstance(cache, dict):
        return {k: {r: v for r, v in m.items() if v is not None}
                for k, m in cache.items()}
    refs = {r: v for r, v in cache if v is not None}
    return {order[0]: refs} if refs and order else {}


def _inject_cached_values(path, cache):
    """Add the cached <v> results next to the <f> formulas in every sheet.

    ``cache`` is either a list of (cell_ref, value) pairs for a single-sheet
    workbook or a dict {sheet_name: {cell_ref: value}}.  openpyxl writes a
    formula without a cached value, so a plain data_only=True read would
    return None.  This patches the saved workbook (which is exactly how Excel
    itself stores a recalculated formula) so the report pipeline can read the
    numbers while Excel still shows live formulas that recalculate when a
    value is edited.
    """
    order = _sheet_order(path)
    per_sheet = _cache_to_refs(order, cache)
    if not any(per_sheet.values()):
        return
    with zipfile.ZipFile(path) as zin:
        names = zin.namelist()
        blobs = {n: zin.read(n) for n in names}
    ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    patched = 0
    for i, sheet in enumerate(order):
        refs = per_sheet.get(sheet)
        if not refs:
            continue
        fname = f"xl/worksheets/sheet{i + 1}.xml"
        if fname not in blobs:
            continue
        root = etree.fromstring(blobs[fname])
        for cell in root.iter("{%s}c" % ns):
            ref = cell.get("r")
            if ref not in refs:
                continue
            if cell.find("{%s}f" % ns) is None:
                continue
            v = cell.find("{%s}v" % ns)
            if v is None:
                v = etree.SubElement(cell, "{%s}v" % ns)
            v.text = repr(refs[ref])
        blobs[fname] = etree.tostring(root, xml_declaration=True,
                                      encoding="UTF-8")
        patched += 1
    if not patched:
        raise RuntimeError("no worksheet patched with cached values")
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zout:
        for n in names:
            zout.writestr(n, blobs[n])


def _missing_cached_refs(path, refs):
    """Verify patched-in formula caches without paying a full workbook load.

    Streams the file in read-only mode and skips raw record sheets ('Data…')
    -- they never hold formulas but dominate the quarterly workbooks' size.
    Returns [(sheet, ref)] for every cached ref still reading as None.
    """
    wb2 = openpyxl.load_workbook(path, data_only=True, read_only=True)
    try:
        bad = []
        for sh, m in refs.items():
            if sh not in wb2.sheetnames:
                continue
            want = {coordinate_to_tuple(r): r for r, v in m.items()
                    if v is not None}
            if not want or sh.lower().startswith("data"):
                continue
            ws2 = wb2[sh]
            seen = set()
            for row in ws2.iter_rows():
                for c in row:
                    # read-only sheets emit EmptyCell placeholders that carry
                    # no coordinates
                    r_i = getattr(c, "row", None)
                    c_i = getattr(c, "column", None)
                    if r_i is None or c_i is None:
                        continue
                    key = (r_i, c_i)
                    if key in want and c.value is not None:
                        seen.add(key)
                if len(seen) == len(want):
                    break
            bad.extend((sh, want[k]) for k in want.keys() - seen)
        return bad
    finally:
        wb2.close()


def _save_workbook(wb, path, cache):
    """Save the workbook, patch in cached formula values, and fall back to
    plain values if anything goes wrong so the pipeline never breaks.

    ``cache`` is a list of (cell_ref, value) pairs (first sheet) or a dict
    {sheet_name: {cell_ref: value}}.
    """
    try:
        wb.calculation.fullCalcOnLoad = True
    except Exception:
        pass
    wb.save(path)
    if not cache:
        return
    refs = _cache_to_refs(_sheet_order(path), cache)
    try:
        _inject_cached_values(path, cache)
        bad = _missing_cached_refs(path, refs)
        if bad:
            raise RuntimeError(f"cached values missing for {bad}")
    except Exception:
        wb3 = openpyxl.load_workbook(path)
        for sh, m in refs.items():
            if sh not in wb3.sheetnames:
                continue
            ws3 = wb3[sh]
            for ref, val in m.items():
                if val is not None:
                    ws3[ref] = val
        wb3.save(path)


def _finalize(ws, path, cache):
    """Convenience wrapper for _save_workbook with a single worksheet."""
    _save_workbook(ws.parent, path,
                   {ws.title: {r: v for r, v in cache if v is not None}})
