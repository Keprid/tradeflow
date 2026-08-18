#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
make_tables.py
==============

Turn the raw ITC (International Trade Centre) Excel downloads for Kenya's
trade with any partner country into the six "Table 1 .. Table 6" workbooks
that the report generator (`generate_report.py`) expects.

Raw source files expected in ``--excel-dir`` (located by keywords in the
filename, case-insensitive):

    <country>s-imports-from-world-by-exporter_all.xlsx   -> Table 1  (import source markets)
    <country>s-imports-from-world-by-product_all.xlsx    -> Table 2  (import products)
    <country>s-exports-to-world-by-importer_all.xlsx     -> Table 3  (export destinations)
    <country>s-exports-to-world-by-product_all.xlsx      -> Table 4  (export products)
    kenyas-exports-to-<country>-by-product_all.xlsx      -> Table 5  (Kenya exports to the country)
    kenyas-imports-from-<country>-by-product_all.xlsx    -> Table 6  (Kenya imports from the country)

Data manipulation (mirrors how the reference tables were prepared):

* Source values are in USD Thousand.
* Tables 1-4  : divide by 1,000,000  ->  "Value in USD Billion".
* Tables 5-6  : divide by 1,000      ->  "Value in USD Million".
* The first column ranks rows by the latest year, highest first.
* Market tables (Tables 1 and 3) always keep the top 25 exporters/partners;
  product tables (Tables 2, 4, 5, 6) keep the top ``--top`` rows
  (default 20).
* For the partner-country tables (1 and 3): if Kenya is not inside the top N
  it is still included with its actual rank and the whole row is highlighted.
* The values of every row not shown are summed into an
  "All other countries/products" row; the last row is the World / All products
  total.
* The final column is the share of the latest-year total
  (row value / total), shown as a percentage.

Runs on Windows, Linux and macOS (Python 3.9+, openpyxl only).

Usage
-----
    python make_tables.py --excel-dir sourcefiles --out-dir output/tables

Then feed the generated folder to ``generate_report.py``. The script also
writes a ``Figure 1 Trade Balance.xlsx`` derived from the Table 5 / Table 6
totals (Kenya's exports minus imports with the partner country).
"""

import argparse
import os
import re
import sys
import zipfile

import openpyxl
from lxml import etree
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------------------
# Constants / styles
# ---------------------------------------------------------------------------
FONT = "Times New Roman"

TOP_MARKETS = 25

HDR_FILL = PatternFill(fill_type="solid", fgColor="5D7B9D")
KENYA_FILL = PatternFill(fill_type="solid", fgColor="FFF2CC")
BAND_FILL = PatternFill(fill_type="solid", fgColor="F7F6F3")

THIN = Side(style="thin", color="808080")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

FMT_VALUE = "0.0"
FMT_SHARE = "0.0%"

# Column widths for Figure 1 (label + up to ten year columns).
BALANCE_WIDTHS = {"A": 16, "B": 30, "C": 12, "D": 12, "E": 12, "F": 12,
                  "G": 12, "H": 12, "I": 12, "J": 12, "K": 12}


# ---------------------------------------------------------------------------
# Small helpers
# ---------------------------------------------------------------------------
def to_float(v):
    """Convert a raw cell value to float; None for blanks / '...' / garbage."""
    if v is None:
        return None
    if isinstance(v, str):
        v = v.strip().replace(",", "")
        if v in ("", "...", "..", "-"):
            return None
    try:
        f = float(v)
        return f if f == f else None
    except (TypeError, ValueError):
        return None


def truncate_label(label, limit=80):
    """Shorten a long product label for the product tables.

    Labels are long HS descriptions built from ``;``-separated clauses. The
    result keeps whole clauses as long as they fit within ``limit`` chars and
    never appends a trailing ellipsis. Any trailing dots/punctuation (which
    the raw ITC download sometimes leaves on a cut description) are removed so
    every label reads as a clean, complete sentence.
    """
    if not label:
        return ""
    label = str(label).strip().rstrip(".,;: ")
    if not label:
        return ""
    parts = re.split(r";\s*", label)
    out = parts[0]
    for p in parts[1:]:
        candidate = out + "; " + p
        if len(candidate) <= limit:
            out = candidate
        else:
            break
    if len(out) > limit:
        head = out[:limit]
        cut = head.rsplit(" ", 1)[0]
        out = (cut or head).rstrip(".,;: ")
    return out


def find_source_files(excel_dir):
    """Locate the six raw source files by filename keywords."""
    found = {}
    for fname in sorted(os.listdir(excel_dir)):
        low = fname.lower()
        if not low.endswith((".xlsx", ".xlsm")):
            continue
        if "imports-from-world" in low and "by-exporter" in low:
            found["table1"] = fname
        elif "imports-from-world" in low and "by-product" in low:
            found["table2"] = fname
        elif "exports-to-world" in low and "by-importer" in low:
            found["table3"] = fname
        elif "exports-to-world" in low and "by-product" in low:
            found["table4"] = fname
        elif "imports-from" in low and "by-product" in low:
            found["table6"] = fname
        elif "exports-to" in low and "by-product" in low:
            found["table5"] = fname
    missing = [k for k in ("table1", "table2", "table3", "table4", "table5", "table6")
               if k not in found]
    if missing:
        sys.exit(
            "[ERROR] Could not find all six raw source files in '%s'.\n"
            "Missing: %s\n"
            "Expected names like:\n"
            "  <country>s-imports-from-world-by-exporter_all.xlsx\n"
            "  <country>s-imports-from-world-by-product_all.xlsx\n"
            "  <country>s-exports-to-world-by-importer_all.xlsx\n"
            "  <country>s-exports-to-world-by-product_all.xlsx\n"
            "  kenyas-exports-to-<country>-by-product_all.xlsx\n"
            "  kenyas-imports-from-<country>-by-product_all.xlsx"
            % (excel_dir, ", ".join(missing)))
    return {k: os.path.join(excel_dir, v) for k, v in found.items()}


# ---------------------------------------------------------------------------
# Source parsing
# ---------------------------------------------------------------------------
def parse_source(path, years=None):
    """Read an ITC workbook. Return (rows, year_columns, years, labels).

    year_columns is a list of (column_index, year) for the years shown.
    When `years` is given, only those years' columns are kept (all tables in
    one report must share the same year set); otherwise the years are chosen
    automatically via _pick_years.
    """
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb[wb.sheetnames[0]]
    grid = list(ws.iter_rows(values_only=True))
    wb.close()

    header = [str(c).strip() if c is not None else "" for c in grid[0]]
    ycols = []
    for i, h in enumerate(header):
        m = re.match(r"^(\d{4})", h)
        if m:
            ycols.append((i, int(m.group(1))))
    ycols.sort(key=lambda t: t[1])

    if years is None:
        chosen = _pick_year_cols(ycols)
    else:
        allowed = set(years)
        chosen = [(i, y) for i, y in ycols if y in allowed]
    years_out = [y for _, y in chosen]

    rows = [r for r in grid[1:]
            if r and any(c is not None and str(c).strip() != "" for c in r)]
    labels = {
        "reporter": str(grid[1][1]).strip() if len(grid) > 1 and grid[1][1] else "",
        "partner": str(grid[1][3]).strip() if len(grid) > 1 and grid[1][3] else "",
    }
    return rows, chosen, years_out, labels


def _pick_years(years):
    """Choose the report years: the longest run of 4..10 consecutive years
    (ties broken by the most recent run). Falls back to the last five years
    if no valid run exists.
    """
    if not years:
        return []
    runs = []
    cur = [years[0]]
    for y in years[1:]:
        if y == cur[-1] + 1:
            cur.append(y)
        else:
            runs.append(cur)
            cur = [y]
    runs.append(cur)
    valid = [r for r in runs if 4 <= len(r) <= 10]
    if not valid:
        return years[-min(5, len(years)):]
    return max(valid, key=lambda r: (len(r), r[-1]))


def _pick_year_cols(ycols):
    """Choose the year columns the report will cover (see _pick_years)."""
    selected = set(_pick_years([y for _, y in ycols]))
    return [(i, y) for i, y in ycols if y in selected]


def _vals(row, ycols):
    return [to_float(row[i]) if i < len(row) else None for i, _ in ycols]


def extract_markets(rows, ycols):
    """Partner-country tables (Table 1 / Table 3): one row per partner."""
    total = None
    items = []
    for r in rows:
        if len(r) < 6:
            continue
        if str(r[4] or "").strip().upper() != "TOTAL":
            continue
        vals = _vals(r, ycols)
        if str(r[2] or "").strip() == "000":
            total = {"label": str(r[3]).strip(), "vals": vals}
        else:
            items.append({"label": str(r[3]).strip(), "vals": vals})
    return total, items


def extract_products(rows, ycols):
    """Product tables (Table 2 / Table 4): World rows only."""
    total = None
    items = []
    for r in rows:
        if str(r[2] or "").strip() != "000":
            continue
        code = str(r[4] or "").strip()
        vals = _vals(r, ycols)
        if code.upper() == "TOTAL":
            total = {"code": code, "label": str(r[5]).strip(), "vals": vals}
        else:
            items.append({"code": code, "label": str(r[5]).strip(), "vals": vals})
    return total, items


def extract_kenya(rows, ycols):
    """Bilateral product tables (Table 5 / Table 6): every row is Kenya's
    exports/imports with the partner."""
    total = None
    items = []
    for r in rows:
        code = str(r[4] or "").strip()
        vals = _vals(r, ycols)
        if code.upper() == "TOTAL":
            total = {"code": code, "label": str(r[5]).strip(), "vals": vals}
        else:
            items.append({"code": code, "label": str(r[5]).strip(), "vals": vals})
    return total, items


# ---------------------------------------------------------------------------
# Ranking / aggregation
# ---------------------------------------------------------------------------
def rank_and_top(items, top_n):
    """Sort by latest-year value (desc), assign ranks to every item and
    return the top `top_n` items."""
    items.sort(
        key=lambda it: (it["vals"][-1] if it["vals"][-1] is not None else float("-inf")),
        reverse=True,
    )
    for i, it in enumerate(items, 1):
        it["rank"] = i
    return items[:top_n]


def find_kenya(items):
    for it in items:
        if str(it.get("label", "")).strip().lower() == "kenya":
            return it
    return None


def sum_shown(shown, n_years):
    out = []
    for k in range(n_years):
        s = 0.0
        for it in shown:
            v = it["vals"][k]
            s += v if v is not None else 0.0
        out.append(s)
    return out


def all_other_vals(total, shown, n_years):
    """Total minus the sum of the displayed rows, per year."""
    if total is None:
        return [None] * n_years
    sums = sum_shown(shown, n_years)
    out = []
    for k in range(n_years):
        t = total["vals"][k]
        out.append(t - sums[k] if t is not None else None)
    return out


def share(v, total_v):
    if v is None or total_v is None or total_v == 0:
        return None
    return v / total_v


def prepare(extracted, years, top_n, divisor, is_markets):
    """Build the display structure for one table.

    extracted : dict with 'total' and 'items' (raw USD Thousand values)
    divisor   : 1e6 for billion tables, 1e3 for million tables
    """
    total = extracted["total"]
    items = extracted["items"]

    for it in items:
        it["vals"] = [v / divisor if v is not None else None for v in it["vals"]]
    if total:
        total["vals"] = [v / divisor if v is not None else None for v in total["vals"]]

    shown = list(rank_and_top(items, top_n))
    if is_markets:
        kenya = find_kenya(items)
        if kenya is not None and kenya not in shown:
            shown.append(kenya)

    n = len(years)
    ao_vals = all_other_vals(total, shown, n)
    t_latest = total["vals"][-1] if total else None
    for it in shown:
        it["share"] = share(it["vals"][-1], t_latest)
    total["share"] = share(t_latest, t_latest) if total else None
    all_other = {"vals": ao_vals,
                 "share": share(ao_vals[-1], t_latest) if total else None}

    return {"years": years, "total": total, "shown": shown,
            "all_other": all_other}


# ---------------------------------------------------------------------------
# Worksheet styling helpers
# ---------------------------------------------------------------------------
def set_widths(ws, widths):
    for col, w in widths.items():
        ws.column_dimensions[col].width = w


def style_cell(cell, bold=False, size=11, fill=None, numfmt=None,
               wrap=False, align="center"):
    cell.font = Font(name=FONT, size=size, bold=bold)
    cell.border = BORDER
    if fill is not None:
        cell.fill = fill
    if numfmt is not None:
        cell.number_format = numfmt
    cell.alignment = Alignment(horizontal=align, vertical="center", wrap_text=wrap)


def put(ws, row, col, value, **kw):
    cell = ws.cell(row=row, column=col, value=value)
    style_cell(cell, **kw)
    return cell


def put_text(ws, row, col, text, bold=False, size=11, fill=None, wrap=False,
             align="left"):
    return put(ws, row, col, text, bold=bold, size=size, fill=fill,
               wrap=wrap, align=align)


def put_val(ws, row, col, value, bold=False, fill=None, fmt=FMT_VALUE):
    return put(ws, row, col, value, bold=bold, fill=fill, numfmt=fmt)


def put_share(ws, row, col, value, bold=False, fill=None):
    return put(ws, row, col, value, bold=bold, fill=fill, numfmt=FMT_SHARE)


def merge(ws, r1, c1, r2, c2):
    ws.merge_cells(start_row=r1, start_column=c1, end_row=r2, end_column=c2)


def _cell_ref(row, col):
    return f"{get_column_letter(col)}{row}"


def _put_formula(ws, cache, row, col, formula, cached, bold=False, fill=None,
                 numfmt=None):
    """Write an Excel formula and register the numeric result for caching."""
    put(ws, row, col, formula, bold=bold, fill=fill, numfmt=numfmt)
    cache.append((_cell_ref(row, col), cached))


def _sheet_order(path):
    """Sheet names in document order (openpyxl names the xml files
    sheet1.xml, sheet2.xml ... matching that order)."""
    with zipfile.ZipFile(path) as zin:
        xml = zin.read("xl/workbook.xml").decode("utf-8", "ignore")
    return re.findall(r'<sheet[^>]*name="([^"]+)"', xml)


def _inject_cached_values(path, cache):
    """Add the cached <v> results next to the <f> formulas in every sheet.

    ``cache`` is either a list of (cell_ref, value) pairs for a single-sheet
    workbook or a dict {sheet_name: {cell_ref: value}}. openpyxl writes a
    formula without a cached value, so a plain data_only=True read would
    return None. This patches the saved workbook (which is exactly how Excel
    itself stores a recalculated formula) so the report pipeline can read the
    numbers while Excel still shows live formulas that recalculate when a
    value is edited.
    """
    if isinstance(cache, dict):
        per_sheet = {k: {r: v for r, v in m.items() if v is not None}
                     for k, m in cache.items()}
    else:
        per_sheet = {"Sheet1": {r: v for r, v in cache if v is not None}}
    if not any(per_sheet.values()):
        return
    with zipfile.ZipFile(path) as zin:
        names = zin.namelist()
        blobs = {n: zin.read(n) for n in names}
    ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    order = _sheet_order(path)
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


def _save_workbook(wb, path, cache):
    """Save the workbook, patch in cached formula values, and fall back to
    plain values if anything goes wrong so the pipeline never breaks."""
    try:
        wb.calculation.fullCalcOnLoad = True
    except Exception:
        pass
    wb.save(path)
    try:
        _inject_cached_values(path, cache)
        wb2 = openpyxl.load_workbook(path, data_only=True)
        refs = cache if isinstance(cache, dict) else \
            {"Sheet1": {r: v for r, v in cache if v is not None}}
        bad = []
        for sh, m in refs.items():
            if sh not in wb2.sheetnames:
                continue
            ws2 = wb2[sh]
            for ref, val in m.items():
                if ws2[ref].value is None:
                    bad.append((sh, ref))
        if bad:
            raise RuntimeError(f"cached values missing for {bad}")
    except Exception:
        wb3 = openpyxl.load_workbook(path)
        refs = cache if isinstance(cache, dict) else \
            {"Sheet1": {r: v for r, v in cache if v is not None}}
        for sh, m in refs.items():
            if sh not in wb3.sheetnames:
                continue
            ws3 = wb3[sh]
            for ref, val in m.items():
                ws3[ref] = val
        wb3.save(path)


def _finalize(ws, path, cache):
    """Convenience wrapper for _save_workbook with a single worksheet."""
    _save_workbook(ws.parent, path, {ws.title: {r: v for r, v in cache
                                                if v is not None}})


def _set_table_filter(ws, first_row, last_row, last_col):
    """Add Excel AutoFilter (column dropdown arrows) over the data range.
    AutoFilter is used instead of a native Excel Table object because the
    header rows contain merged cells, which Excel Tables reject."""
    try:
        ws.auto_filter.ref = f"A{first_row}:{_cell_ref(last_row, last_col)}"
    except Exception:
        pass


# ---------------------------------------------------------------------------
# Table writers
# ---------------------------------------------------------------------------
def write_market_table(ws, data, rep, is_exports, unit_row, kenya_highlight,
                       row1_label=None, row1_title=None, cache=None):
    """Table 1 (import source markets) and Table 3 (export destinations).

    The share column, the "All other countries" row and the "World" row are
    written as live Excel formulas (with cached results), so the workbook
    recalculates when a value is edited.
    """
    cache = [] if cache is None else cache
    years = data["years"]
    n = len(years)
    latest = years[-1]
    total = data["total"]
    shown = data["shown"]

    label_col = "Importers" if is_exports else "Exporters"
    verb = ("Importing Countries from " if is_exports else "List of Exporters to ")
    verb += rep
    if not unit_row:
        verb += "\nValue in USD Billion"

    # optional title row (Table 1 only)
    if row1_label or row1_title:
        if row1_label:
            put_text(ws, 1, 1, row1_label, bold=True, size=11, align="center")
        merge(ws, 1, 2, 1, 2 + n)
        put_text(ws, 1, 2, row1_title, bold=True, size=11, wrap=True, align="center")
        title_rows = 1
    else:
        title_rows = 0

    h = 1 + title_rows
    share_col = 3 + n
    put_text(ws, h, 1, f"Rank in {latest}", bold=True, size=11, align="center")
    merge(ws, h, 1, h + 1, 1)
    put_text(ws, h, 2, label_col, bold=True, size=8, fill=HDR_FILL, align="center")
    merge(ws, h, 3, h, 2 + n)
    put_text(ws, h, 3, verb, bold=True, size=11 if unit_row else 8,
             fill=None if unit_row else HDR_FILL, wrap=True, align="center")
    merge(ws, h, share_col, h + 1, share_col)
    put_text(ws, h, share_col, f"Share in {latest} %", bold=True, size=11, align="center")

    # years row
    for k, y in enumerate(years):
        put_text(ws, h + 1, 3 + k, y, bold=True, size=8, fill=HDR_FILL, align="center")

    if unit_row:
        merge(ws, h + 2, 3, h + 2, 2 + n)
        put_text(ws, h + 2, 3, "Value in USD Billion", bold=True, size=8,
                 fill=HDR_FILL, align="center")
        put_text(ws, h + 2, 2, "", bold=True, size=8, fill=HDR_FILL, align="center")
        row0 = h + 3
    else:
        row0 = h + 2

    # data rows
    r_first = row0
    r_last = row0 + len(shown) - 1
    r_ao = row0 + len(shown)
    r_world = r_ao + 1
    for i, it in enumerate(shown):
        r = row0 + i
        is_kenya = kenya_highlight and str(it["label"]).strip().lower() == "kenya"
        fill = KENYA_FILL if is_kenya else None
        put_text(ws, r, 1, it["rank"], size=11, fill=fill)
        put_text(ws, r, 2, it["label"], size=8, fill=fill, wrap=True)
        for k in range(n):
            put_val(ws, r, 3 + k, it["vals"][k], fill=fill)
        if it["share"] is not None and total and total["vals"][-1] is not None:
            _put_formula(ws, cache, r, share_col,
                         f"={_cell_ref(r, 3 + n - 1)}/{_cell_ref(r_world, 3 + n - 1)}",
                         it["share"], fill=fill, numfmt=FMT_SHARE)
        else:
            put_share(ws, r, share_col, it["share"], fill=fill)

    # all other row
    r = r_ao
    put_text(ws, r, 2, "All other countries", size=11, wrap=True)
    for k in range(n):
        c = 3 + k
        val = data["all_other"]["vals"][k]
        if (len(shown) and val is not None
                and total and total["vals"][k] is not None):
            _put_formula(ws, cache, r, c,
                         f"={_cell_ref(r_world, c)}"
                         f"-SUM({_cell_ref(r_first, c)}:{_cell_ref(r_last, c)})",
                         val, numfmt=FMT_VALUE)
        else:
            put_val(ws, r, c, val)
    ao_share = data["all_other"]["share"]
    if ao_share is not None and total and total["vals"][-1] is not None:
        _put_formula(ws, cache, r, share_col,
                     f"={_cell_ref(r_ao, 3 + n - 1)}/{_cell_ref(r_world, 3 + n - 1)}",
                     ao_share, numfmt=FMT_SHARE)
    else:
        put_share(ws, r, share_col, ao_share)

    # world row
    r = r_world
    put_text(ws, r, 2, "World", bold=True, size=8, fill=BAND_FILL)
    for k in range(n):
        c = 3 + k
        val = total["vals"][k] if total else None
        put_val(ws, r, c, val, bold=True, fill=BAND_FILL)
    put_share(ws, r, share_col, total["share"] if total else None, bold=True)
    ws.freeze_panes = f"A{row0}"
    _set_table_filter(ws, h, r_world, share_col)


def write_product_table(ws, data, hdr, unit_row,
                        row1_label=None, row1_title=None,
                        prefix_apostrophe=True, cache=None):
    """Tables 2, 4, 5, 6 (product tables).

    The share column, the "All other products" row and the "All products"
    total row are written as live Excel formulas (with cached results).
    """
    cache = [] if cache is None else cache
    years = data["years"]
    n = len(years)
    latest = years[-1]
    total = data["total"]
    shown = data["shown"]

    # optional title row (Tables 2, 4, 6)
    if row1_label or row1_title:
        if row1_label:
            put_text(ws, 1, 1, row1_label, bold=True, size=11, align="center")
        merge(ws, 1, 2, 1, 2 + n)
        put_text(ws, 1, 2, row1_title, bold=True, size=11, wrap=True, align="center")
        title_rows = 1
    else:
        title_rows = 0

    h = 1 + title_rows
    put_text(ws, h, 1, f"Rank in {latest}", bold=True, size=11, align="center")
    merge(ws, h, 2, h + 1, 2)
    put_text(ws, h, 2, "Code", bold=True, size=11, align="center")
    merge(ws, h, 3, h + 1, 3)
    put_text(ws, h, 3, "Product label", bold=True, size=11, align="center")
    merge(ws, h, 4, h, 3 + n)
    put_text(ws, h, 4, hdr, bold=True, size=11, wrap=True, align="center")
    merge(ws, h, 4 + n, h + 1, 4 + n)
    put_text(ws, h, 4 + n, f"Share in {latest} %", bold=True, size=11, align="center")

    # years row
    for k, y in enumerate(years):
        put_text(ws, h + 1, 4 + k, y, bold=True, size=11, align="center")

    if unit_row:
        merge(ws, h + 2, 4, h + 2, 3 + n)
        put_text(ws, h + 2, 4, "Value in USD Billion", bold=True, size=11,
                 fill=HDR_FILL, align="center")
        row0 = h + 3
    else:
        row0 = h + 2

    # data rows
    r_first = row0
    r_last = row0 + len(shown) - 1
    r_ao = row0 + len(shown)
    r_total = r_ao + 1
    for i, it in enumerate(shown):
        r = row0 + i
        put_text(ws, r, 1, it["rank"], size=11)
        code = it["code"]
        if prefix_apostrophe and code:
            code = "'" + code
        put_text(ws, r, 2, code, size=11)
        put_text(ws, r, 3, truncate_label(it["label"]), size=11, wrap=True)
        for k in range(n):
            put_val(ws, r, 4 + k, it["vals"][k])
        if it["share"] is not None and total and total["vals"][-1] is not None:
            _put_formula(ws, cache, r, 4 + n,
                         f"={_cell_ref(r, 4 + n - 1)}/{_cell_ref(r_total, 4 + n - 1)}",
                         it["share"], numfmt=FMT_SHARE)
        else:
            put_share(ws, r, 4 + n, it["share"])

    # all other row
    r = r_ao
    put_text(ws, r, 3, "All other products", size=11, wrap=True)
    for k in range(n):
        c = 4 + k
        val = data["all_other"]["vals"][k]
        if (len(shown) and val is not None
                and total and total["vals"][k] is not None):
            _put_formula(ws, cache, r, c,
                         f"={_cell_ref(r_total, c)}"
                         f"-SUM({_cell_ref(r_first, c)}:{_cell_ref(r_last, c)})",
                         val, numfmt=FMT_VALUE)
        else:
            put_val(ws, r, c, val)
    ao_share = data["all_other"]["share"]
    if ao_share is not None and total and total["vals"][-1] is not None:
        _put_formula(ws, cache, r, 4 + n,
                     f"={_cell_ref(r_ao, 4 + n - 1)}/{_cell_ref(r_total, 4 + n - 1)}",
                     ao_share, numfmt=FMT_SHARE)
    else:
        put_share(ws, r, 4 + n, ao_share)

    # total row
    r = r_total
    if total and total.get("code"):
        put_text(ws, r, 2, ("'" if prefix_apostrophe else "") + str(total["code"]),
                 bold=True, size=11, fill=BAND_FILL, wrap=True)
    put_text(ws, r, 3, "All products", bold=True, size=11, fill=BAND_FILL, wrap=True)
    for k in range(n):
        c = 4 + k
        val = total["vals"][k] if total else None
        put_val(ws, r, c, val, bold=True, fill=BAND_FILL)
    put_share(ws, r, 4 + n, total["share"] if total else None, bold=True)
    ws.freeze_panes = f"A{row0}"
    _set_table_filter(ws, h, r_total, 4 + n)


def write_balance(ws, krep, kpartner, years, exports, imports, cache=None, skip_chart=False):
    """Figure 1: bilateral trade balance derived from the Table 5/6 totals.

    exports = Kenya's exports to the partner (Table 5 total, USD Million)
    imports = Kenya's imports from the partner (Table 6 total, USD Million)
    Balance of Trade = exports - imports (a live Excel formula).
    """
    cache = [] if cache is None else cache
    hdr = HDR_FILL
    put_text(ws, 1, 1, "Figure 1", bold=True, size=11, align="center")
    put_text(ws, 1, 2, f"BOT Kenya- {kpartner}", bold=True, size=11, align="center")
    put_text(ws, 2, 2, f"Kenya's exports to {kpartner}", size=11, fill=hdr, align="center")

    put_text(ws, 3, 2, "", bold=True, size=11, fill=hdr, align="center")
    for k, y in enumerate(years):
        put_text(ws, 3, 2 + k, y, bold=True, size=11, fill=hdr, align="center")

    def row(label, values):
        put_text(ws, r, 1, label, bold=True, size=11, fill=hdr, align="center")
        for k, v in enumerate(values):
            put_val(ws, r, 2 + k, v)

    r = 4
    row("Exports", exports)
    r += 1
    row("Imports", imports)
    r += 1
    put_text(ws, r, 1, "Balance of Trade", bold=True, size=11, fill=hdr, align="center")
    for k, (e, i) in enumerate(zip(exports, imports)):
        c = 2 + k
        if e is not None and i is not None:
            _put_formula(ws, cache, r, c, f"={_cell_ref(4, c)}-{_cell_ref(5, c)}",
                         e - i)
        else:
            put_val(ws, r, c, (e - i) if (e is not None and i is not None) else None)
    _set_table_filter(ws, 3, 6, 2 + len(years))
    if not skip_chart:
        _add_balance_chart(ws, krep, kpartner, years)


def _add_balance_chart(ws, krep, kpartner, years):
    """Embed a clustered bar chart of exports/imports/balance below the table."""
    try:
        from openpyxl.chart import BarChart, Reference
        n = len(years)
        chart = BarChart()
        chart.type = "col"
        chart.style = 10
        chart.title = f"{krep}-{kpartner} Balance of Trade"
        chart.y_axis.title = "USD Million"
        chart.x_axis.title = "Year"
        chart.y_axis.crosses = "autoZero"
        chart.height = 9
        chart.width = 20
        cats = Reference(ws, min_col=2, max_col=1 + n, min_row=3)
        chart.add_data(Reference(ws, min_col=1, max_col=1 + n,
                                 min_row=4, max_row=6),
                       titles_from_data=True, from_rows=True)
        chart.set_categories(cats)
        ws.add_chart(chart, f"A{8 if len(years) <= 5 else 9}")
    except Exception:
        pass


def read_existing_balance(path):
    """Read (years, exports, imports) from an existing balance workbook.

    Handles both the layout written by write_balance and the reference
    "Figure 1 Trade Balance.xlsx" (which may carry two series blocks and extra
    columns). Returns None if the file is unreadable or has no valid series.
    """
    if not os.path.exists(path):
        return None
    try:
        wb = openpyxl.load_workbook(path, data_only=True)
        ws = wb.worksheets[0]
        grid = [[c.value for c in row] for row in ws.iter_rows()]
    except Exception:
        return None

    for ri, row in enumerate(grid):
        if not row or not isinstance(row[0], str):
            continue
        lbl = row[0].strip()
        if lbl == "Exports" and ri + 2 < len(grid):
            l1 = grid[ri + 1][0].strip() if isinstance(grid[ri + 1][0], str) else ""
            l2 = grid[ri + 2][0].strip() if isinstance(grid[ri + 2][0], str) else ""
            if l1 == "Imports" and l2 == "Balance of Trade":
                years = [int(v) for v in grid[ri - 1][1:]
                         if isinstance(v, (int, float)) and not isinstance(v, bool)
                         and 1990 <= v <= 2100]
                if not years:
                    continue
                exports = [grid[ri][1 + i] if isinstance(grid[ri][1 + i], (int, float)) else None
                           for i in range(len(years))]
                imports = [grid[ri + 1][1 + i] if isinstance(grid[ri + 1][1 + i], (int, float)) else None
                           for i in range(len(years))]
                return {"years": years, "exports": exports, "imports": imports}
    return None


def merge_balance_series(existing, years, exports, imports):
    """Merge new data over an existing series, keeping older history."""
    if not existing:
        return years, exports, imports
    ymap = dict(zip(existing["years"], existing["exports"]))
    imap = dict(zip(existing["years"], existing["imports"]))
    for y, e, i in zip(years, exports, imports):
        ymap[y] = e
        imap[y] = i
    ys = sorted(ymap)
    return ys, [ymap[y] for y in ys], [imap[y] for y in ys]


# ---------------------------------------------------------------------------
# Bilateral export composition analysis
# ---------------------------------------------------------------------------
def compute_bilateral_alignment(kenya_items, kenya_total, partner_items,
                                partner_total, years):
    """Compare Kenya's export products to the partner vs partner's import demand.

    Matches by HS code. For each Kenya product, computes:
    - Kenya's share of its total exports
    - Partner's import share for the same product
    - Alignment score: how well Kenya's exports match partner's import demand
    """
    partner_by_code = {}
    for it in partner_items:
        partner_by_code[it["code"]] = it

    latest = -1
    results = []
    for kit in kenya_items:
        code = kit["code"]
        pit = partner_by_code.get(code)
        if pit is None:
            continue

        kv = kit["vals"][latest] if kit["vals"] and len(kit["vals"]) > abs(latest) else None
        kt = kenya_total["vals"][latest] if kenya_total and kenya_total["vals"] else None
        pv = pit["vals"][latest] if pit["vals"] and len(pit["vals"]) > abs(latest) else None
        pt = partner_total["vals"][latest] if partner_total and partner_total["vals"] else None

        if kv is None or kt is None or kt == 0:
            continue
        k_share = kv / kt

        p_share = None
        if pv is not None and pt is not None and pt != 0:
            p_share = pv / pt

        # Alignment: how relevant is this product to the partner's imports
        alignment = p_share if p_share is not None else 0

        # Growth
        growth = None
        if (len(kit["vals"]) >= 2 and kit["vals"][-1] is not None
                and kit["vals"][-2] is not None and kit["vals"][-2] != 0):
            growth = (kit["vals"][-1] - kit["vals"][-2]) / abs(kit["vals"][-2])

        results.append({
            "code": code,
            "label": kit["label"],
            "kenya_val": kv,
            "kenya_share": k_share,
            "partner_val": pv,
            "partner_share": p_share,
            "alignment": alignment,
            "growth": growth,
        })

    results.sort(key=lambda x: x["kenya_val"] if x["kenya_val"] is not None else -1,
                 reverse=True)
    results = results[:30]
    for i, it in enumerate(results, 1):
        it["rank"] = i
    return results


def write_bilateral_table(ws, items, years, krep, kpartner,
                          row1_label=None, row1_title=None):
    """Write Kenya's bilateral export composition + market alignment table.

    Columns: Rank | Code | Product | Kenya Value | Kenya Share | Partner Import Share | Growth %
    """
    latest = years[-1] if years else ""

    if row1_label or row1_title:
        if row1_label:
            put_text(ws, 1, 1, row1_label, bold=True, size=11, align="center")
        merge(ws, 1, 2, 1, 7)
        put_text(ws, 1, 2, row1_title, bold=True, size=11, wrap=True, align="center")
        title_rows = 1
    else:
        title_rows = 0

    h = 1 + title_rows
    cols = ["Rank", "HS Code", "Product",
            f"Kenya Value USD M ({latest})",
            f"Kenya Share (%)",
            f"{kpartner} Import Share (%)",
            "Annual Growth %"]
    for c, hdr in enumerate(cols, 1):
        put_text(ws, h, c, hdr, bold=True, size=8, fill=HDR_FILL, align="center",
                 wrap=True)

    row0 = h + 1
    for it in items:
        r = row0 + it["rank"] - 1
        put_text(ws, r, 1, it["rank"], size=11)
        put_text(ws, r, 2, str(it["code"]), size=11)
        put_text(ws, r, 3, it["label"], size=11, wrap=True, align="left")
        put_val(ws, r, 4, it["kenya_val"], fmt="0.0")
        put_share(ws, r, 5, it["kenya_share"])
        put_share(ws, r, 6, it["partner_share"])
        if it["growth"] is not None:
            put_share(ws, r, 7, it["growth"])
        else:
            put_share(ws, r, 7, None)

    set_widths(ws, {"A": 7, "B": 10, "C": 48, "D": 18, "E": 14, "F": 20, "G": 14})
    ws.freeze_panes = f"A{row0}"


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
def generate_tables(excel_dir, out_dir, top_n):
    """Build the Table 1-6 workbooks.

    top_n : number of top product rows kept for Tables 2, 4, 5 and 6.
            Market tables (Tables 1 and 3) always keep TOP_MARKETS rows.
    """
    os.makedirs(out_dir, exist_ok=True)
    files = find_source_files(excel_dir)

    # ---- common year set: all six tables must share the same years ---------
    keys = ("table1", "table2", "table3", "table4", "table5", "table6")
    years_sets = []
    for k in keys:
        _, _, ys, _ = parse_source(files[k])
        years_sets.append(set(ys))
    full_years = _pick_years(sorted(set.intersection(*years_sets)))
    # Tables 1-6 show only the last five years; Figure 1 (the trade balance)
    # uses the full common year run so the analysis covers e.g. 2016-2025.
    table_years = full_years[-5:] if len(full_years) > 5 else full_years

    # ---- Table 1: import source markets --------------------------------
    rows, ycols, years, labels = parse_source(files["table1"], years=table_years)
    total, items = extract_markets(rows, ycols)
    rep = labels["reporter"]
    d1 = prepare({"total": total, "items": items}, years, TOP_MARKETS, 1e6, is_markets=True)

    # ---- Table 3: export destinations -----------------------------------
    rows, ycols, years, labels = parse_source(files["table3"], years=table_years)
    total, items = extract_markets(rows, ycols)
    d3 = prepare({"total": total, "items": items}, years, TOP_MARKETS, 1e6, is_markets=True)

    # ---- Table 2: import products ---------------------------------------
    rows, ycols, years, labels = parse_source(files["table2"], years=table_years)
    total, items = extract_products(rows, ycols)
    d2 = prepare({"total": total, "items": items}, years, top_n, 1e6, is_markets=False)
    p_total_raw = total   # partner import total (raw, USD Thousand)
    p_items_raw = items   # partner import items (raw, USD Thousand)

    # ---- Table 4: export products ---------------------------------------
    rows, ycols, years, labels = parse_source(files["table4"], years=table_years)
    total, items = extract_products(rows, ycols)
    d4 = prepare({"total": total, "items": items}, years, top_n, 1e6, is_markets=False)

    # ---- Table 5: Kenya exports to partner ------------------------------
    rows, ycols, years, labels = parse_source(files["table5"], years=table_years)
    total, items = extract_kenya(rows, ycols)
    krep = labels["reporter"] or "Kenya"
    kpartner = labels["partner"]
    d5 = prepare({"total": total, "items": items}, years, top_n, 1e3, is_markets=False)
    k_total_raw = total   # Kenya export total (raw, USD Thousand)
    k_items_raw = items   # Kenya export items (raw, USD Thousand)

    # ---- Table 6: Kenya imports from partner ----------------------------
    rows, ycols, years, labels = parse_source(files["table6"], years=table_years)
    total, items = extract_kenya(rows, ycols)
    d6 = prepare({"total": total, "items": items}, years, top_n, 1e3, is_markets=False)

    # ---- Table 7: Kenya's export composition + market alignment -----------
    bilateral_items = compute_bilateral_alignment(
        k_items_raw, k_total_raw, p_items_raw, p_total_raw, years)

    rep_possessive = rep if rep.endswith("s") else rep + "'s"

    out = {}

    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Table 1"
    cache = []
    write_market_table(ws, d1, rep, is_exports=False, unit_row=False,
                       kenya_highlight=True, cache=cache,
                       row1_label="Table 1:",
                       row1_title=f"{rep} Import Source Markets")
    set_widths(ws, {"A": 7, "B": 32, "C": 12, "D": 12, "E": 12, "F": 12,
                    "G": 12, "H": 12})
    out["t1"] = os.path.join(out_dir, f"Table 1 {rep} Import Source Markets.xlsx")
    _finalize(ws, out["t1"], cache)

    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Table 2"
    cache = []
    write_product_table(ws, d2, "Imported Products from the World",
                        unit_row=True, cache=cache,
                        row1_title=f"List of products imported by {rep}")
    set_widths(ws, {"A": 7, "B": 9, "C": 45, "D": 12, "E": 12, "F": 12,
                    "G": 12, "H": 12, "I": 12})
    out["t2"] = os.path.join(out_dir, f"Table 2 {rep} Lead Imports.xlsx")
    _finalize(ws, out["t2"], cache)

    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Table 3"
    cache = []
    write_market_table(ws, d3, rep, is_exports=True, unit_row=True,
                       kenya_highlight=True, cache=cache)
    set_widths(ws, {"A": 7, "B": 32, "C": 12, "D": 12, "E": 12, "F": 12,
                    "G": 12, "H": 12})
    out["t3"] = os.path.join(out_dir, f"Table 3 {rep} Lead Export Destinations.xlsx")
    _finalize(ws, out["t3"], cache)

    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Table 4"
    cache = []
    write_product_table(
        ws, d4, "Exported Products to the World\nValue in USD Billions",
        unit_row=False, cache=cache,
        row1_title=f"{rep_possessive} Lead Export Products to the World")
    set_widths(ws, {"A": 7, "B": 9, "C": 45, "D": 12, "E": 12, "F": 12,
                    "G": 12, "H": 12, "I": 12})
    out["t4"] = os.path.join(out_dir, f"Table 4 {rep} Lead Export Products.xlsx")
    _finalize(ws, out["t4"], cache)

    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Table 5"
    cache = []
    write_product_table(
        ws, d5, f"{krep}'s exports to {kpartner}\nValue in USD Million",
        unit_row=False, cache=cache)
    set_widths(ws, {"A": 7, "B": 9, "C": 45, "D": 12, "E": 12, "F": 12,
                    "G": 12, "H": 12, "I": 12})
    out["t5"] = os.path.join(out_dir, f"Table 5 {krep} Exports to {kpartner}.xlsx")
    _finalize(ws, out["t5"], cache)

    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Table 6"
    cache = []
    write_product_table(
        ws, d6, f"{krep}'s imports from {kpartner}\nValue in USD Million",
        unit_row=False, cache=cache,
        row1_label="Table 6",
        row1_title=f"{krep}'s Top imports from {kpartner}")
    set_widths(ws, {"A": 7, "B": 9, "C": 45, "D": 12, "E": 12, "F": 12,
                    "G": 12, "H": 12, "I": 12})
    out["t6"] = os.path.join(out_dir, f"Table 6 {krep} Top Imports from {kpartner}.xlsx")
    _finalize(ws, out["t6"], cache)

    # Table 7: Kenya's export composition + market alignment
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Table 7"
    cache = []
    write_bilateral_table(ws, bilateral_items, years, krep, kpartner,
                          row1_label="Table 7:",
                          row1_title=f"{krep}'s Export Composition and Market Alignment with {kpartner}")
    set_widths(ws, {"A": 7, "B": 10, "C": 48, "D": 18, "E": 14, "F": 20, "G": 14})
    out["t7"] = os.path.join(out_dir, f"Table 7 {krep} Export Composition and Market Alignment.xlsx")
    _finalize(ws, out["t7"], cache)

    # ---- Figure 1: trade balance derived from Table 5 / Table 6 ----------
    # The balance uses the full common year run (up to 10 years) even though
    # Tables 1-6 only show the last five years.
    bal_path = os.path.join(out_dir, "Figure 1 Trade Balance.xlsx")
    rows5f, ycols5f, years5f, _ = parse_source(files["table5"], years=full_years)
    rows6f, ycols6f, years6f, _ = parse_source(files["table6"], years=full_years)
    tot5f, _ = extract_kenya(rows5f, ycols5f)
    tot6f, _ = extract_kenya(rows6f, ycols6f)
    exports_full = [v / 1e3 if v is not None else None for v in tot5f["vals"]]
    imports_full = [v / 1e3 if v is not None else None for v in tot6f["vals"]]
    years_m, exports_m, imports_m = merge_balance_series(
        read_existing_balance(bal_path),
        years5f, exports_full, imports_full)
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Figure 1"
    cache = []
    write_balance(ws, krep, kpartner, years_m, exports_m, imports_m, cache=cache)
    set_widths(ws, BALANCE_WIDTHS)
    out["balance"] = bal_path
    _finalize(ws, out["balance"], cache)

    # ---- Combined workbook: all seven sheets in one file -------------------
    wb_all = openpyxl.Workbook()
    wb_all.remove(wb_all.active)
    for s in ("Table 1", "Table 2", "Table 3", "Table 4", "Table 5",
              "Table 6", "Table 7", "Figure 1"):
        wb_all.create_sheet(s)
    cache_all = {}
    widths_mkt = {"A": 7, "B": 32, "C": 12, "D": 12, "E": 12, "F": 12,
                  "G": 12, "H": 12}
    widths_prd = {"A": 7, "B": 9, "C": 45, "D": 12, "E": 12, "F": 12,
                  "G": 12, "H": 12, "I": 12}
    ws = wb_all["Table 1"]; c = []
    write_market_table(ws, d1, rep, is_exports=False, unit_row=False,
                       kenya_highlight=True, cache=c,
                       row1_label="Table 1:",
                       row1_title=f"{rep} Import Source Markets")
    set_widths(ws, widths_mkt); cache_all["Table 1"] = dict(c)
    ws = wb_all["Table 2"]; c = []
    write_product_table(ws, d2, "Imported Products from the World",
                        unit_row=True, cache=c,
                        row1_title=f"List of products imported by {rep}")
    set_widths(ws, widths_prd); cache_all["Table 2"] = dict(c)
    ws = wb_all["Table 3"]; c = []
    write_market_table(ws, d3, rep, is_exports=True, unit_row=True,
                       kenya_highlight=True, cache=c)
    set_widths(ws, widths_mkt); cache_all["Table 3"] = dict(c)
    ws = wb_all["Table 4"]; c = []
    write_product_table(
        ws, d4, "Exported Products to the World\nValue in USD Billions",
        unit_row=False, cache=c,
        row1_title=f"{rep_possessive} Lead Export Products to the World")
    set_widths(ws, widths_prd); cache_all["Table 4"] = dict(c)
    ws = wb_all["Table 5"]; c = []
    write_product_table(
        ws, d5, f"{krep}'s exports to {kpartner}\nValue in USD Million",
        unit_row=False, cache=c)
    set_widths(ws, widths_prd); cache_all["Table 5"] = dict(c)
    ws = wb_all["Table 6"]; c = []
    write_product_table(
        ws, d6, f"{krep}'s imports from {kpartner}\nValue in USD Million",
        unit_row=False, cache=c,
        row1_label="Table 6",
        row1_title=f"{krep}'s Top imports from {kpartner}")
    set_widths(ws, widths_prd); cache_all["Table 6"] = dict(c)
    ws = wb_all["Table 7"]; c = []
    write_bilateral_table(ws, bilateral_items, years, krep, kpartner,
                          row1_label="Table 7:",
                          row1_title=f"{krep}'s Export Composition and Market Alignment with {kpartner}")
    set_widths(ws, {"A": 7, "B": 10, "C": 48, "D": 18, "E": 14, "F": 20, "G": 14})
    cache_all["Table 7"] = dict(c)
    ws = wb_all["Figure 1"]; c = []
    write_balance(ws, krep, kpartner, years_m, exports_m, imports_m, cache=c, skip_chart=True)
    set_widths(ws, BALANCE_WIDTHS)
    cache_all["Figure 1"] = dict(c)
    out["all"] = os.path.join(out_dir, "All Tables.xlsx")
    _save_workbook(wb_all, out["all"], cache_all)

    return out, (rep, kpartner)


def main():
    ap = argparse.ArgumentParser(
        description="Build Table 1-6 workbooks from raw ITC source files.")
    ap.add_argument("--excel-dir", default="sourcefiles",
                    help="Folder with the six raw ITC source files (default: sourcefiles)")
    ap.add_argument("--out-dir", default=os.path.join("output", "tables"),
                    help="Where to write the generated Table 1-6 files "
                         "(default: output/tables)")
    ap.add_argument("--top", type=int, default=20,
                    help="Number of top product rows to keep (default: 20). "
                         "Tables 1 and 3 always keep the top 25 exporters")
    args = ap.parse_args()

    if args.top < 1:
        sys.exit("[ERROR] --top must be >= 1")

    try:
        out, (rep, partner) = generate_tables(args.excel_dir, args.out_dir, args.top)
    except SystemExit:
        raise
    except Exception as e:
        sys.exit(f"[ERROR] Failed to generate tables: {e}")

    print("[1/2] Read raw source files from :", os.path.abspath(args.excel_dir))
    print("[2/2] Wrote tables for           :", rep, "<-> Kenya")
    for key in ("t1", "t2", "t3", "t4", "t5", "t6", "t7", "balance", "all"):
        print(f"      - {os.path.basename(out[key])}")
    print()
    print("The Figure 1 file is derived from the Table 5 and Table 6 totals")
    print("(Kenya's exports to / imports from the partner country). The")
    print("'All Tables.xlsx' workbook bundles every sheet into one file.")
    print()
    print("Next step: run")
    print(f'  python generate_report.py --excel-dir "{os.path.abspath(args.out_dir)}"')


if __name__ == "__main__":
    main()
