#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
export_promotion_analysis.py
============================

OPTIONAL, ADDITIVE KEPROBA export-promotion module.

This module produces a standalone Excel workbook that supports Kenya's export
promotion effort (market selection, competitor positioning, price/quality
positioning).  It is a superset of the analyses already printed in the report:

  * it NEVER modifies the .docx report or the companion "TABLES.xlsx"
    deliverable (those stay byte-for-byte identical with or without it);
  * it is only invoked when the ``--promotion-analysis <out.xlsx>`` flag is
    passed to ``generate_report.py``.

Data requirements
-----------------
Tier 1 (works with the standard six ITC Trade Map downloads already in use)::

    Tier 1                             | Needs
    ----------------------------------|----------------------------------------
    Kenya -> partner product mix      | Table 5 (Kenya's exports to partner)
    Partner import-demand structure   | Table 1 (partner's import source mmrkts)
    Competitor share concentration    | Table 1 (who supplies the partner today)

Tier 2 (optional extra downloads unlock price positioning)::

    Unit-value / price positioning    | any Trade Map workbook that carries
                                      | the Quantity columns, e.g. the
                                      | Kenya-<partner> bilateral download
                                      | (already in Tier 1) or a "Time Series"
                                      | download from either side, dropped
                                      | into the same --excel-dir folder.

When Tier-2 data is absent the workbook still builds and points out exactly
which download would unlock each blank tab (no import failure).
"""

import argparse
import os
import re
import shutil
import sys
import tempfile

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE_DIR)

from generate_report import clean_label, num, pct, to_float  # noqa: E402
from xlsx_compat import convert_to_xlsx  # noqa: E402

NAVY = "1F3864"
GRID_COLOR = "D6DCE4"
HDR_FILL = PatternFill(fill_type="solid", fgColor=NAVY)
HDR_FONT = Font(name="Century Gothic", size=11, bold=True, color="FFFFFF")
THIN = Side(style="thin", color=GRID_COLOR)
CENTER = Alignment(horizontal="center", vertical="center")
LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)
RIGHT = Alignment(horizontal="right", vertical="center")

THEME = ["4472C4", "ED7D31", "A5A5A5", "FFC000", "5B9BD5",
         "70AD47", "FF0000", "7030A0", "00B0F0", "F7A7A7"]

QTY_KEYWORDS = ("quantity", "qty", "tons", "tonnes", "kg", " unit")
UV_REQUIREMENT_NOTE = (
    "Unit-value (price) positioning needs a workbook that carries the "
    "Quantity columns - the Trade Map \"Time Series\" download for Kenya's "
    "exported products or the Kenya-<partner> bilateral workbook. "
    "Drop that file into the --excel-dir folder and re-run.")


def _hhi(shares):
    """Herfindahl-Hirschman index on a list of percentage shares (0-100)."""
    total = sum(s for s in shares if s is not None)
    if total <= 0:
        return None
    norm = [s / total * 100 for s in shares if s is not None]
    return round(sum(v * v for v in norm), 1)


def _growth(row, kind="cagr"):
    vals = [to_float(v) for v in row]
    vals = [v for v in vals if v is not None and v != 0]
    if len(vals) < 2:
        return None
    first, last = vals[0], vals[-1]
    if first <= 0 or last <= 0:
        return None
    n = len(vals) - 1
    return (last / first) ** (1.0 / n) - 1.0 if kind == "cagr" else last / first - 1


def _style_header(ws, row, ncols, labels):
    for c, lab in zip(range(1, ncols + 1), labels):
        cell = ws.cell(row, c)
        cell.value = lab
        cell.font = HDR_FONT
        cell.fill = HDR_FILL
        cell.alignment = CENTER


def _fit(ws, widths):
    for c, w in enumerate(widths, start=1):
        try:
            ws.column_dimensions[get_column_letter(c)].width = w
        except Exception:
            pass


def _legend(ws, r, lines):
    """Recurring interpretation block, styled like the reading guides."""
    for line in lines:
        cell = ws.cell(r, 1, line)
        cell.alignment = LEFT
        cell.font = Font(italic=True, color="595959")
        r += 1
    return r


# ---------------------------------------------------------------------------
# Tier 1: Kenya -> partner product mix (Table 5)
# ---------------------------------------------------------------------------
def _kenya_partner_mix(ws, a, partner):
    t5 = a.table5
    items = t5.get("items") if t5 else []
    if not items:
        ws.cell(1, 1, "Table 5 (Kenya's exports to %s) not found." % partner)
        return
    years = t5.get("years") or []
    r = 1
    ws.cell(r, 1, "Kenya -> %s product mix (drivers & laggards)" % partner,
            ).font = Font(bold=True, size=13)
    r += 2
    latest_col = max(0, len(years) - 1)
    y = years[latest_col] if latest_col < len(years) else None
    latest_label = ("%s (USD m)" % y) if y else "Latest (USD m)"
    _style_header(ws, r, 8, ["Rank", "HS", "Product", latest_label,
                              "Share", "5-yr CAGR", "Status", "Promotion note"])
    r += 1
    rows = []
    for d in items:
        yrs = d.get("years") or []
        latest = to_float(yrs[latest_col]) if latest_col < len(yrs) else None
        rows.append((d, latest))
    rows.sort(key=lambda x: -(x[1] or 0))
    total = sum(latest for _, latest in rows) or 0.0
    for i, (d, latest) in enumerate(rows, start=1):
        if latest is None or latest == 0:
            continue
        cagr = _growth(d.get("years") or [], "cagr")
        share = (latest / total * 100) if total else 0.0
        if cagr is not None and cagr > 0.05 and share >= 5:
            status, note = "Rising star", "Promote at priority fairs"
        elif cagr is not None and cagr < -0.05 and share >= 5:
            status, note = "Declining", "Match price / quality vs rivals"
        elif cagr is not None and cagr > 0.05:
            status, note = "Niche momentum", "B2B buyer-seller targeting"
        else:
            status, note = "Mature", "Defend volume; upgrade packaging"
        vals = [i, d.get("code") or "", clean_label(d.get("name") or ""),
                None if latest is None else round(latest, 3),
                round(share, 1),
                None if cagr is None else round(cagr * 100, 1),
                status, note]
        for c, v in enumerate(vals, start=1):
            cell = ws.cell(r, c, v)
            if c in (4, 5, 6):
                cell.number_format = "#,##0.0" if c == 4 else "0.0"
                cell.alignment = RIGHT
            elif c >= 7:
                cell.alignment = LEFT
        r += 1
    hhi = _hhi([s for _, s in rows if s is not None])
    r += 1
    if hhi is not None:
        ws.cell(r, 1, "Concentration of Kenya's exports to %s (HHI): %.0f  %s"
                % (partner, hhi,
                   "(basket is DIVERSIFIED)" if hhi < 1500
                   else "(basket is CONCENTRATED - diversify)"))
        ws.cell(r, 1).font = Font(bold=True, color="1F4E79")
    r += 1
    r = _legend(ws, r, [
        "How to read the 'Status' column:",
        "  Rising star (>= 5% of basket and growing >= 5%/yr) - the priority: promote at "
        "the big fairs and allocate the exporter slots.",
        "  Niche momentum (growing >= 5%/yr but small share) - grow quietly via B2B "
        "buyer-seller targeting before competitors notice.",
        "  Declining (>= 5% of basket, shrinking <= -5%/yr) - something is beating you: "
        "match the rival on price and quality or reposition the offer.",
        "  Mature (everything else) - defend volume and defend the moat with packaging "
        "and positioning upgrades.",
        "Share = product's share (in %%) of Kenya's exports to %s; 5-yr CAGR = compound "
        "annual growth of Kenya's exports to %s over the years shown." % (partner,
                                                                          partner),
        "HHI < 1500 = diversified basket (healthy). HHI >= 1500 = concentrated - a few "
        "products carry the country, so spread the promotion."])
    return {"hhi": hhi, "rows": len(rows)}


# ---------------------------------------------------------------------------
# Tier 1: partner import-demand structure & competitor concentration (Table 1)
# ---------------------------------------------------------------------------
def _partner_import_structure(ws, a, partner):
    t1 = a.table1
    items = t1.get("items") if t1 else []
    if not items:
        ws.cell(1, 1, "Table 1 (%s's import source markets) not found." % partner)
        return
    years = t1.get("years") or []
    r = 1
    ws.cell(r, 1, "Who supplies %s today (competitor map)" % partner,
            ).font = Font(bold=True, size=13)
    r += 2
    latest_col = max(0, len(years) - 1)
    y = years[latest_col] if latest_col < len(years) else None
    latest_label = ("%s (USD bn)" % y) if y else "Latest (USD bn)"
    _style_header(ws, r, 7, ["Rank", "Supplier", latest_label, "Share",
                             "5-yr CAGR", "Position", "Action"])
    r += 1
    rows = []
    for d in items:
        yrs = d.get("years") or []
        latest = to_float(yrs[latest_col]) if latest_col < len(yrs) else None
        rows.append((d, latest))
    rows.sort(key=lambda x: -(x[1] or 0))
    total = (t1.get("total") or {}).get("years") or []
    total_latest = to_float(total[latest_col]) if latest_col < len(total) else None
    for i, (d, latest) in enumerate(rows[:25], start=1):
        if latest is None:
            continue
        cagr = _growth(d.get("years") or [], "cagr")
        share = (latest / total_latest * 100) if total_latest else None
        if share is not None and share >= 15 and (cagr or 0) > 0.02:
            pos, act = "Dominant & growing", "Study their price/terms playbook"
        elif share is not None and share >= 10:
            pos, act = "Established", "Target adjacent products they ignore"
        else:
            pos, act = "Contestable", "Feasible share-gain target for Kenya"
        for c, v in enumerate([i, clean_label(d.get("name") or ""),
                               None if latest is None else round(latest, 3),
                               None if share is None else round(share, 1),
                               None if cagr is None else round(cagr * 100, 1),
                               pos, act], start=1):
            cell = ws.cell(r, c, v)
            if c in (3, 4, 5):
                cell.number_format = "#,##0.0" if c == 3 else "0.0"
                cell.alignment = RIGHT
            elif c >= 6:
                cell.alignment = LEFT
        r += 1
    r += 1
    r = _legend(ws, r, [
        "How to read the 'Position' column:",
        "  Dominant & growing (share >= 15% and rising) - entrenched: study and match "
        "their price/terms/delivery, or attack niches they under-serve.",
        "  Established (share >= 10%) - credible but beatable: target adjacent products "
        "they ignore and use Kenya's EAC and logistics advantage.",
        "  Contestable (below the thresholds, or falling) - NOT entrenched. This is the "
        "feasible share-gain target for Kenya: win share with price, terms, packaging "
        "and delivery - no head-on fight needed.",
        "Share = supplier's share (in %%) of %s's total imports in the latest year "
        "shown; 5-yr CAGR = compound growth of that supplier's exports to %s over the "
        "years shown." % (partner, partner),
        "The 'Action' column turns each position into a promotion task for the KEPROBA "
        "country desk."])
    _fit(ws, [8, 34, 13, 9, 10, 24, 44])


# ---------------------------------------------------------------------------
# Tier 2: unit-value / price positioning (needs Quantity data)
# ---------------------------------------------------------------------------
def _find_quantity_workbooks(excel_dir):
    """Return bonuses: files whose name suggests a Trade Map download that
    carries the Quantity columns - the \"Time Series\" workbooks for Kenya's
    exported products and the Kenya-<partner> bilateral workbooks."""
    return _scan_files(excel_dir, None, _qty_matcher)


def _qty_matcher(fname):
    low = (fname or "").lower()
    if not low.endswith((".xlsx", ".xls")):
        return False
    return ("time" in low or "series" in low or "quantity" in low
            or "bilateral_trade_between_kenya_and" in low
            or "exported_by_kenya" in low and "product" in low
            or "by-product" in low)


def _scan_files(excel_dir, scan_dirs, matcher):
    """Collect absolute paths (sorted, deduped) from a base dir and any extra
    scan dirs whose filename matches *matcher*."""
    paths, seen = [], set()
    for d in [excel_dir] + [x for x in (scan_dirs or [])]:
        if not d or not os.path.isdir(d):
            continue
        for f in sorted(os.listdir(d)):
            if not matcher(f):
                continue
            key = os.path.abspath(os.path.join(d, f))
            if key not in seen:
                seen.add(key)
                paths.append(key)
    return paths


def _open_scan(path):
    """Open any supported spreadsheet, converting .xls/.csv/.xlsb to a temp
    .xlsx first.  Returns (workbook, tmpdir_to_clean) or (None, None)."""
    tmpdir = None
    try:
        if path.lower().endswith((".xls", ".xlsb", ".csv", ".tsv")):
            tmpdir = tempfile.mkdtemp(prefix="tradeflow_scan_")
            wb, _ = convert_to_xlsx(path, out_dir=tmpdir)
        else:
            wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
        return wb, tmpdir
    except Exception:
        if tmpdir:
            shutil.rmtree(tmpdir, ignore_errors=True)
        return None, None


def _close_scan(wb, tmpdir):
    try:
        wb.close()
    except Exception:
        pass
    if tmpdir:
        shutil.rmtree(tmpdir, ignore_errors=True)


def _bilateral_rows(excel_dir, scan_dirs=None):
    """(rows, source, year) from a 'Trade Map - Bilateral trade between Kenya
    and <P>' download found in the scanned folders."""
    def m(f):
        low = f.lower()
        return low.endswith((".xls", ".xlsx")) \
            and "bilateral_trade_between_kenya_and" in low
    for path in _scan_files(excel_dir, scan_dirs, m):
        wb, tmp = _open_scan(path)
        if wb is None:
            continue
        try:
            rows, year = _bilateral_products(wb.worksheets[0])
            if rows:
                return rows, os.path.basename(path), year
        finally:
            _close_scan(wb, tmp)
    return None, None, None


def _champions_rows(excel_dir, scan_dirs=None):
    """(rows, source, year) from the Trade Map 'List of products exported by
    Kenya' download (Kenya's world-export champions) in the scanned folders."""
    def m(f):
        low = f.lower()
        return low.endswith((".xls", ".xlsx")) \
            and "exported_by_kenya" in low and "product" in low
    for path in _scan_files(excel_dir, scan_dirs, m):
        wb, tmp = _open_scan(path)
        if wb is None:
            continue
        try:
            rows, year = _kenya_world_exports(wb.worksheets[0])
            if rows:
                return rows, os.path.basename(path), year
        finally:
            _close_scan(wb, tmp)
    return None, None, None


def _header_index(ws):
    """Map each column letter-cell to (label, row).  Returns {label: col}."""
    labels = {}
    for row in ws.iter_rows(min_row=1, max_row=8, max_col=min(ws.max_column, 20)):
        for cell in row:
            if cell.value is None:
                continue
            lab = str(cell.value)
            low = lab.lower()
            if "value" in low or "quantity" in low or low in ("qty", "unit"):
                if lab not in labels:
                    labels.setdefault(low[:8], cell.column)
    out = {}
    for key, col in labels.items():
        if key.startswith("quantity") or key.strip() in ("qty",):
            out.setdefault("quantity", col)
        elif key.startswith("value"):
            out.setdefault("value", col)
    return out


def _unit_value_rows(excel_dir, n=20, scan_dirs=None):
    """Best-effort (product, unit_value) from a quantity-bearing workbook.

    Scans *excel_dir* (the tables folder in the webapp raw pipeline) and,
    when given, the extra *scan_dirs* (the raw uploads folder that holds the
    bilateral download)."""
    for path in _scan_files(excel_dir, scan_dirs, _qty_matcher):
        bilateral = "bilateral_trade_between_kenya_and" in path.lower()
        wb, tmp = _open_scan(path)
        if wb is None:
            continue
        try:
            ws = wb.worksheets[0]
            out = (_classic_bilateral_unit_values(ws, n) if bilateral
                   else _generic_unit_value_rows(ws, n))
            if out:
                return out, os.path.basename(path)
        finally:
            _close_scan(wb, tmp)
    return None, None


def _generic_unit_value_rows(ws, n):
    cols = _header_index(ws)
    if not cols.get("value") or not cols.get("quantity"):
        return None
    vcol, qcol = cols["value"], cols["quantity"]
    out = []
    for row in ws.iter_rows(min_row=2):
        label = None
        for cell in row:
            if isinstance(cell.value, str) and len(cell.value.strip()) > 3 \
                    and any(ch.isalpha() for ch in cell.value) \
                    and not cell.value.strip().startswith(("Reporter",
                                                           "Product")):
                if label is None or len(cell.value.strip()) > len(label):
                    label = cell.value.strip()
        v = None
        q = None
        for cell in row:
            if cell.column == vcol:
                v = to_float(cell.value)
            elif cell.column == qcol:
                q = to_float(cell.value)
        if label is None or not v or not q:
            continue
        out.append((label, v / q, None))
    return out[:n] or None


def _s(v):
    return str(v).strip() if v is not None else ""


_VALUE_IN_RE = re.compile(r"value\s+in\s+((?:19|20)\d{2})", re.IGNORECASE)
_QTY_EXPORTED_RE = re.compile(r"quantity\s+exported", re.IGNORECASE)
_UNIT_VALUE_RE = re.compile(r"unit\s+value", re.IGNORECASE)
_GROWTH_RE = re.compile(r"annual\s+growth\s+in\s+value", re.IGNORECASE)
_SHARE_RE = re.compile(r"share\s+in", re.IGNORECASE)
_TARIFF_RE = re.compile(r"tariff", re.IGNORECASE)
_WORLD_VALUE_RE = re.compile(
    r"value\s+exported\s+in\s+((?:19|20)\d{2})", re.IGNORECASE)
_WORLD_BALANCE_RE = re.compile(
    r"trade\s+balance\s+in\s+((?:19|20)\d{2})", re.IGNORECASE)
_WORLD_GROWTH_VALUE_RE = re.compile(
    r"annual\s+growth\s+in\s+value\s+between", re.IGNORECASE)
_WORLD_GROWTH_QTY_RE = re.compile(
    r"annual\s+growth\s+in\s+quantity\s+between", re.IGNORECASE)
_WORLD_IMP_GROWTH_RE = re.compile(
    r"growth\s+of\s+world\s+imports", re.IGNORECASE)
_WORLD_SHARE_RE = re.compile(r"share\s+in\s+world\s+exports", re.IGNORECASE)
_WORLD_RANK_RE = re.compile(r"ranking\s+of\s+", re.IGNORECASE)


def _bilateral_products(ws):
    """Parse a classic 'Trade Map - Bilateral trade between Kenya and
    <Partner>' workbook (or the mirror 'Kenya imports from <Partner>') into
    per-product records.

    The workbook's indicator sub-row sits two columns left of its data (the
    'Product code'/'Product label' headers only span the first header row), so
    columns are re-anchored on the first data row - exactly the trick
    make_tables uses for Tables 5/6.  The three value blocks (Kenya's exports
    to the partner, the partner's imports from the world, Kenya's exports to
    the world) are split via the block titles in the row above.

    Returns (rows, year) with fields per product: label, k_* (Kenya's exports
    to the partner), u_* (partner's imports from the world), w_* (Kenya's
    exports to the world).  Values are USD thousand, growths/shares %, unit
    values USD/unit (unit value is computed from value/quantity when Trade Map
    does not print it)."""
    grid = [[c.value for c in row]
            for row in ws.iter_rows(min_row=1, max_row=min(ws.max_row, 1300),
                                    max_col=min(ws.max_column, 30))]
    sub_idx = None
    for ri, row in enumerate(grid):
        if any(_VALUE_IN_RE.search(_s(v)) for v in row):
            sub_idx = ri
            break
    if sub_idx is None:
        return None, None
    sub = grid[sub_idx]
    title_row = grid[sub_idx - 1] if sub_idx else []
    value_hdrs = [i for i, v in enumerate(sub)
                  if isinstance(v, str) and _VALUE_IN_RE.search(v)]
    if not value_hdrs:
        return None, None
    year = _VALUE_IN_RE.search(_s(sub[value_hdrs[0]])).group(1)
    year = int(year) if year else None

    def _hdr_cols(pat):
        return {i for i, v in enumerate(sub)
                if isinstance(v, str) and pat.search(v)}

    growth_hdrs, share_hdrs = _hdr_cols(_GROWTH_RE), _hdr_cols(_SHARE_RE)
    tariff_hdrs, qty_hdrs = _hdr_cols(_TARIFF_RE), _hdr_cols(_QTY_EXPORTED_RE)
    uv_hdrs = _hdr_cols(_UNIT_VALUE_RE)

    blocks = []
    for i, vi in enumerate(value_hdrs):
        nxt = value_hdrs[i + 1] if i + 1 < len(value_hdrs) else len(sub)
        span = list(range(vi + 1, nxt))
        blk = {"value": vi}
        for key, hset in (("growth", growth_hdrs), ("share", share_hdrs),
                          ("tariff", tariff_hdrs), ("qty", qty_hdrs),
                          ("uv", uv_hdrs)):
            blk[key] = next((x for x in span if x in hset), None)
        blk["title"] = _s(title_row[i]).lower() if i < len(title_row) else ""
        blocks.append(blk)
    data_row = next((r for r in grid[sub_idx + 1:] if any(_s(c) for c in r)),
                    None)
    if data_row is None:
        return None, None
    data_cols = [i for i, c in enumerate(data_row) if i >= 2 and _s(c)]
    if not data_cols:
        return None, None
    off = data_cols[0] - blocks[0]["value"]

    kenya = uganda = world = None
    for b in blocks:
        t = b["title"]
        if kenya is None and "kenya" in t and "export" in t \
                and "to" in t and "world" not in t:
            kenya = {k: v + off if isinstance(v, int) else v
                     for k, v in b.items()}
        elif uganda is None and "world" in t and "import" in t:
            uganda = {k: v + off if isinstance(v, int) else v
                      for k, v in b.items()}
        elif world is None and "kenya" in t and "export" in t and "world" in t:
            world = {k: v + off if isinstance(v, int) else v
                     for k, v in b.items()}
    if kenya is None:
        return None, None

    def _at(r, block, key):
        col = block.get(key)
        if col is None or len(r) <= col:
            return None
        return to_float(r[col])

    rows = []
    for r in grid[sub_idx + 1:]:
        if not any(_s(c) for c in r):
            continue
        label = _s(r[1]) if len(r) > 1 else ""
        low = label.lower()
        if len(label) <= 3 or low.startswith(("total", "all products", "world")):
            continue
        rec = {"label": label}
        for prefix, blk in (("k", kenya), ("u", uganda), ("w", world)):
            if blk is None:
                continue
            uv = _at(r, blk, "uv")
            if not uv:
                v, q = _at(r, blk, "value"), _at(r, blk, "qty")
                if v and q:
                    uv = v * 1000.0 / q
            rec[prefix + "_value"] = _at(r, blk, "value")
            rec[prefix + "_growth"] = _at(r, blk, "growth")
            rec[prefix + "_share"] = _at(r, blk, "share")
            rec[prefix + "_tariff"] = _at(r, blk, "tariff")
            rec[prefix + "_uv"] = uv
        rows.append(rec)
    return (rows or None), year


def _classic_bilateral_unit_values(ws, n=20):
    """Kenya->partner (product, market_uv, world_uv) tuples from a bilateral
    download - see _bilateral_products for the parsing; *world_uv* is the same
    product's unit value in Kenya's exports-to-world block, a like-for-like
    within-file benchmark for the premium/discount reading."""
    rows, _year = _bilateral_products(ws)
    if not rows:
        return None
    out = [(r["label"], r.get("k_uv"), r.get("w_uv")) for r in rows
           if r.get("k_uv")]
    return out[:n] or None


def _kenya_world_exports(ws):
    """Parse the Trade Map 'List of products exported by Kenya in <YEAR>'
    workbook (Kenya's world exports, one row per HS 4-digit product) into
    per-product records.  Same 2-column header offset as the bilateral file.
    Returns (rows, year): fields code, label, value (USD thousand, the top
    export year), balance (USD thousand), growth_recent/growth_long (Kenya's
    export value growth %), growth_qty (%), world_import_growth (%),
    world_share (%), ranking (3-digit index)."""
    grid = [[c.value for c in row]
            for row in ws.iter_rows(min_row=1, max_row=min(ws.max_row, 1300),
                                    max_col=min(ws.max_column, 20))]
    hdr = None
    for ri, row in enumerate(grid[:8]):
        if any(_WORLD_VALUE_RE.search(_s(v)) for v in row):
            hdr = ri
            break
    if hdr is None:
        return None, None
    sub = grid[hdr]
    shared = {i for i, v in enumerate(sub) if isinstance(v, str)}

    def _first(pat, start=0):
        for v in shared:
            if v + 1 < start:
                continue
            if isinstance(sub[v], str) and pat.search(sub[v]):
                return v
        return None

    value_cols = [i for i in shared if _WORLD_VALUE_RE.search(_s(sub[i]))]
    growth_value_cols = [i for i in shared
                         if _WORLD_GROWTH_VALUE_RE.search(_s(sub[i]))]
    year = None
    mv = _WORLD_VALUE_RE.search(_s(sub[value_cols[0]])) if value_cols else None
    if mv:
        year = int(mv.group(1))

    cols = {
        "value": value_cols[0] if value_cols else None,
        "balance": _first(_WORLD_BALANCE_RE),
        "growth_recent": growth_value_cols[0] if growth_value_cols else None,
        "growth_long": (growth_value_cols[1] if len(growth_value_cols) > 1
                        else None),
        "growth_qty": _first(_WORLD_GROWTH_QTY_RE),
        "world_import_growth": _first(_WORLD_IMP_GROWTH_RE),
        "world_share": _first(_WORLD_SHARE_RE),
        "ranking": _first(_WORLD_RANK_RE),
    }
    data_row = next((r for r in grid[hdr + 1:] if any(_s(c) for c in r)),
                    None)
    if data_row is None:
        return None, None
    data_cols = [i for i, c in enumerate(data_row) if i >= 2 and _s(c)]
    if not data_cols:
        return None, None
    off = data_cols[0] - cols["value"]

    def _at(r, key):
        c = cols.get(key)
        if c is None or not isinstance(c, int) or len(r) <= c + off:
            return None
        return to_float(r[c + off])

    rows = []
    for r in grid[hdr + 1:]:
        code = _s(r[0]) if r else ""
        label = _s(r[1]) if len(r) > 1 else ""
        if not code or not label or label.lower().startswith(("all products",
                                                              "total")):
            continue
        rec = {"code": code, "label": label,
               "value": _at(r, "value"), "balance": _at(r, "balance"),
               "growth_recent": _at(r, "growth_recent"),
               "growth_long": _at(r, "growth_long"),
               "growth_qty": _at(r, "growth_qty"),
               "world_import_growth": _at(r, "world_import_growth"),
               "world_share": _at(r, "world_share"),
               "ranking": _at(r, "ranking")}
        if rec["value"] is None:
            continue
        rows.append(rec)
    return (rows or None), year


def _price_positioning(ws, excel_dir, scan_dirs=None):
    r = 1
    ws.cell(r, 1, "Unit-value (price) positioning").font = Font(bold=True, size=13)
    r += 2
    rows, src = _unit_value_rows(excel_dir, scan_dirs=scan_dirs)
    if not rows:
        ws.cell(r, 1, UV_REQUIREMENT_NOTE)
        ws.cell(r, 1).font = Font(italic=True, color="9A1F1F")
        return
    ws.cell(r, 2, "Source workbook: %s" % src).font = Font(italic=True, size=9)
    r += 1
    _style_header(ws, r, 4, ["Product", "Unit value (USD/unit)",
                             "vs World", "Reading"])
    r += 1
    for label, uv, w_uv in rows:
        vs = "n/a - needs world UV to compare"
        note = "Premium quality signal - brand & protect"
        if w_uv and uv:
            delta = (uv / w_uv - 1.0) * 100.0
            if delta >= 15:
                vs, note = "%+.0f%% vs world" % delta, (
                    "Premium over Kenya's world unit value - brand, protect "
                    "and upgrade packaging/positioning")
            elif delta <= -15:
                vs, note = "%+.0f%% vs world" % delta, (
                    "Discounted vs Kenya's world unit value - price-led entry; "
                    "verify buyers don't read lower price as lower quality")
            else:
                vs = "~world parity" if abs(delta) < 1 else "%+.0f%% vs world" % delta
                note = ("Near Kenya's world unit value - standard positioning; "
                        "defend volume and commercial terms")
        cell = ws.cell(r, 1, clean_label(label))
        cell.alignment = LEFT
        cell = ws.cell(r, 2, round(uv, 2))
        cell.number_format = "#,##0.00"
        cell.alignment = RIGHT
        ws.cell(r, 3, vs).alignment = CENTER
        ws.cell(r, 4, note).alignment = LEFT
        r += 1
    r += 1
    ws.cell(r, 1, "How to read this tab: unit value = Kenya's average selling "
                  "price per unit in this market.  'vs World' compares it with "
                  "Kenya's own world-average unit value for the same product "
                  "(from the same bilateral workbook).").alignment = LEFT
    ws.cell(r, 1).font = Font(italic=True, color="595959")
    r += 1
    ws.cell(r, 1, "  Premium (>= +15%)  -> quality-led: brand & protect, push "
                  "packaging/positioning.").font = Font(italic=True, color="595959")
    r += 1
    ws.cell(r, 1, "  Parity (+/-15%)    -> volume-led: defend market share, "
                  "offer terms/reliability.").font = Font(italic=True, color="595959")
    r += 1
    ws.cell(r, 1, "  Discount (<= -15%) -> price-led: entry point, but confirm "
                  "the price isn't read as lower quality.").font = Font(italic=True, color="595959")


def _price_competitiveness(ws, rows, partner, year):
    r = 1
    ws.cell(r, 1, "Kenya's price edge vs %s's world suppliers" % partner,
            ).font = Font(bold=True, size=13)
    r += 2
    data = [d for d in rows if d.get("k_uv") and d.get("u_uv")]
    data.sort(key=lambda x: -(x.get("k_value") or 0))
    if not data:
        ws.cell(r, 1, "Needs the bilateral workbook with the Unit value "
                      "columns for both Kenya and the world block.").font = \
            Font(italic=True, color="9A1F1F")
        return
    year_sfx = (" in %s" % year) if year else ""
    _style_header(ws, r, 5, ["Product", "Kenya's unit value (USD/unit)",
                             "%s's world-import unit value (USD/unit)" % partner,
                             "Kenya price edge", "Reading"])
    r += 1
    for d in data[:25]:
        edge = (d["k_uv"] / d["u_uv"] - 1.0) * 100.0
        if edge <= -15:
            reading = ("Kenya is price-competitive - lead with price, win volume "
                       "while rivals are costlier")
        elif edge >= 15:
            reading = ("Kenya prices above the market - win on quality/brand/"
                       "service, not on price")
        else:
            reading = ("At parity with %s's world import prices - compete on "
                       "terms, reliability and delivery" % partner)
        cell = ws.cell(r, 1, clean_label(d["label"]))
        cell.alignment = LEFT
        cell = ws.cell(r, 2, round(d["k_uv"], 2))
        cell.number_format = "#,##0.00"
        cell.alignment = RIGHT
        cell = ws.cell(r, 3, round(d["u_uv"], 2))
        cell.number_format = "#,##0.00"
        cell.alignment = RIGHT
        cell = ws.cell(r, 4, "%+.0f%%" % edge)
        cell.alignment = CENTER
        cell.fill = PatternFill(fill_type="solid",
                                fgColor="C6EFCE" if edge <= -15 else
                                ("FFEB9C" if edge >= 15 else "D6DCE4"))
        ws.cell(r, 5, reading).alignment = LEFT
        r += 1
    r += 1
    r = _legend(ws, r, [
        "How to read this tab (Kenya's unit value vs the unit value %s's "
        "imports pay the rest of the world%s):" % (partner, year_sfx),
        "  Price edge <= -15%  -> Kenya undercuts the world supply: exploit it "
        "now to win volume and shelf space.",
        "  Price edge between -15% and +15% -> parity: differentiate on terms, "
        "reliability, packaging and delivery.",
        "  Price edge >= +15%  -> Kenya is premium: defend with brand, quality "
        "and service; avoid a price war.",
        "Unit value = trade value / quantity, so it is a true per-unit price "
        "benchmark, not an average ticket size."])
    _fit(ws, [42, 18, 20, 13, 44])


def _market_access(ws, rows, partner, year):
    r = 1
    ws.cell(r, 1, "Market access for Kenya's exports to %s" % partner,
            ).font = Font(bold=True, size=13)
    r += 2
    data = [d for d in rows if d.get("k_value")]
    data.sort(key=lambda x: -(x.get("k_value") or 0))
    if not data:
        ws.cell(r, 1, "Needs the bilateral workbook (Kenya -> %s block)." % partner).font = \
            Font(italic=True, color="9A1F1F")
        return
    year_sfx = (", %s" % year) if year else ""
    _style_header(ws, r, 5, ["Product", "Kenya's exports (USD m%s)" % year_sfx,
                             "Kenya's tariff in %s" % partner, "Demand = %s's "
                             "world imports (USD m)" % partner, "Reading"])
    r += 1
    for d in data[:25]:
        t = d.get("k_tariff")
        if t is not None and t <= 0:
            reading = "Duty-free (EAC) - no tariff barrier, push volume"
        elif t:
            reading = ("Tariff barrier (%.0f%%) - flag for customs/advocacy: "
                       "verify HS classification and zero-rating options" % t)
        else:
            reading = "Tariff n/a in source - confirm at customs"
        cell = ws.cell(r, 1, clean_label(d["label"]))
        cell.alignment = LEFT
        cell = ws.cell(r, 2, None if d.get("k_value") is None
                       else round(d["k_value"] / 1000.0, 1))
        if d.get("k_value") is not None:
            cell.number_format = "#,##0.0"
        cell.alignment = RIGHT
        cell = ws.cell(r, 3, None if t is None else round(t, 1))
        if t is not None:
            cell.number_format = "0.0"
        cell.alignment = RIGHT
        cell.fill = PatternFill(fill_type="solid",
                                fgColor="C6EFCE" if (t is not None and t <= 0)
                                else ("FFC7CE" if t else "D6DCE4"))
        cell = ws.cell(r, 4, None if d.get("u_value") is None
                       else round(d["u_value"] / 1000.0, 1))
        if d.get("u_value") is not None:
            cell.number_format = "#,##0.0"
        cell.alignment = RIGHT
        ws.cell(r, 5, reading).alignment = LEFT
        r += 1
    r += 1
    r = _legend(ws, r, [
        "How to read this tab: under the EAC common market, %s's tariff on "
        "Kenya's exported goods is usually 0%% - those rows are 'push volume' "
        "products." % partner,
        "A tariff > 0% flags a product where market access is NOT automatic: "
        "check the HS classification, claim the EAC preferential rate or "
        "escalate to trade advocacy.",
        "Demand = %s's total imports of that product from the WORLD (USD m) - "
        "it sizes the prize Kenya is currently under-served in." % partner])
    _fit(ws, [42, 15, 14, 18, 44])


def _global_champions(ws, excel_dir, scan_dirs, partner, year_hint):
    r = 1
    ws.cell(r, 1, "Kenya's world-export champions").font = Font(bold=True, size=13)
    r += 2
    rows, src, year = _champions_rows(excel_dir, scan_dirs)
    if not rows:
        ws.cell(r, 1, "Needs the Trade Map 'List of products exported by "
                      "Kenya' download (Kenya's exports to the world).").font = \
            Font(italic=True, color="9A1F1F")
        return
    ws.cell(r, 1, "Source workbook: %s" % src).font = Font(italic=True, size=9)
    r += 1
    y = year or year_hint
    year_sfx = (", %s" % y) if y else ""
    _style_header(ws, r, 8, ["HS", "Product", "Kenya's exports (USD m%s)"
                             % year_sfx, "Kenya 5-yr growth",
                             "World import growth (%)",
                             "Share in world exports", "World rank", "Reading"])
    r += 1
    data = [d for d in rows if d.get("value")]
    data.sort(key=lambda x: -(x.get("value") or 0))
    for d in data[:25]:
        share, rank = d.get("world_share"), d.get("ranking")
        g = d.get("growth_long") or d.get("growth_recent")
        wg = d.get("world_import_growth")
        if share is not None and rank is not None and rank <= 10:
            reading = ("World top-%d exporter (%.1f%% share) - global champion; "
                       "anchor the %s campaign on it" % (int(rank), share,
                                                          partner))
        elif g is not None and wg is not None:
            diff = g - wg
            if diff >= 5:
                reading = ("Gaining world share (Kenya %+.0f pts vs world "
                           "demand) - scale up" % diff)
            elif diff <= -5:
                reading = ("Losing ground (%+.0f pts vs world demand) - check "
                           "price, quality and costs" % diff)
            else:
                reading = "Tracking world demand (%+.0f pts) - hold position" % diff
        elif share is not None:
            reading = ("%.1f%% of world exports - established supplier" % share)
        else:
            reading = "Review using world-trade data"
        cell = ws.cell(r, 1, d.get("code") or "")
        cell.alignment = CENTER
        ws.cell(r, 2, clean_label(d["label"])).alignment = LEFT
        cell = ws.cell(r, 3, round(d["value"] / 1000.0, 1))
        cell.number_format = "#,##0.0"
        cell.alignment = RIGHT
        for c, v in (4, g), (5, wg):
            cell = ws.cell(r, c, None if v is None else round(v, 1))
            if v is not None:
                cell.number_format = "0.0"
            cell.alignment = RIGHT
        cell = ws.cell(r, 6, None if share is None else round(share, 1))
        if share is not None:
            cell.number_format = "0.0"
        cell.alignment = RIGHT
        cell = ws.cell(r, 7, None if rank is None else int(rank))
        if rank is not None:
            cell.number_format = "0"
        cell.alignment = CENTER
        cell = ws.cell(r, 8, reading)
        cell.alignment = LEFT
        cell.fill = PatternFill(fill_type="solid",
                                fgColor="C6EFCE" if (rank is not None
                                                    and rank <= 10) else
                                ("FFEB9C" if (g and wg and g - wg >= 5)
                                 else "D6DCE4"))
        r += 1
    r += 1
    r = _legend(ws, r, [
        "How to read this tab: Kenya's exports of each product to the WORLD, "
        "not just %s." % partner,
        "Comparing Kenya's own export growth with %s's market growth shows "
        "whether Kenya is gaining or losing global share in the product."
        % partner,
        "Share and rank measure world leadership; a top-10 rank means Kenya has "
        "real pricing and supply credibility to leverage in %s." % partner,
        "Trade balance < 0 (not shown if blank in source) means Kenya also "
        "imports the product - a supply gap worth exploring."])
    _fit(ws, [7, 40, 15, 13, 16, 14, 9, 44])


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------
def build_promotion_analysis(a, cfg, excel_dir, out_path, extra_dirs=None):
    """Write the standalone promotion-positioning workbook (additive only)."""
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    partner = cfg["country"]["name"]

    suffix = " Mix"
    room = 31 - len("Kenya-") - len(suffix)      # Excel sheet-name limit
    mix_title = "Kenya-%s%s" % (partner[:room], suffix)
    ws = wb.create_sheet(mix_title)
    _kenya_partner_mix(ws, a, partner)
    _fit(ws, [6, 8, 42, 13, 9, 10, 20, 44])

    ws = wb.create_sheet("Competitor Map")
    _partner_import_structure(ws, a, partner)

    brows, bsrc, byear = _bilateral_rows(excel_dir, extra_dirs)
    if brows:
        ws = wb.create_sheet("Price Competitiveness")
        _price_competitiveness(ws, brows, partner, byear)

        ws = wb.create_sheet("Market Access")
        _market_access(ws, brows, partner, byear)

    ws = wb.create_sheet("Global Champions")
    _global_champions(ws, excel_dir, extra_dirs, partner, byear)

    ws = wb.create_sheet("Price Positioning")
    _price_positioning(ws, excel_dir, extra_dirs)
    _fit(ws, [44, 16, 22, 40])

    wb.save(out_path)
    return out_path


def build_from_dir(cfg, excel_dir, out_path, extra_dirs=None):
    """Load the goods Analysis and write the promotion workbook (webapp path)."""
    from generate_report import Analysis
    a = Analysis(cfg).load(excel_dir)
    return build_promotion_analysis(a, cfg, excel_dir, out_path, extra_dirs)


def main(argv=None):
    ap = argparse.ArgumentParser(
        description="Build the additive KEPROBA export-promotion workbook.")
    ap.add_argument("--excel-dir", default=os.path.join(BASE_DIR, "sample_data"))
    ap.add_argument("--config", default=os.path.join(BASE_DIR, "config",
                                                     "saudi_arabia.json"))
    ap.add_argument("--output", default=None)
    args = ap.parse_args(argv)

    import json
    from generate_report import Analysis
    with open(args.config, "r", encoding="utf-8") as f:
        cfg = json.load(f)
    a = Analysis(cfg).load(args.excel_dir)
    out = args.output or os.path.join(
        BASE_DIR, "output", "Export_Promotion_Positioning.xlsx")
    out_dir = os.path.dirname(out)
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)
    build_promotion_analysis(a, cfg, args.excel_dir, out)
    print("Promotion workbook saved to :", out)


if __name__ == "__main__":
    main()