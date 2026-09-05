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

    Unit-value / price positioning    | any Trade Map "Time Series" workbook
                                      | (carries the Quantity columns) for the
                                      | products of interest, e.g.
                                      |   "List of products exported by Kenya"
                                      |   "List of products imported by <partner>"
                                      | dropped into the same --excel-dir folder.

When Tier-2 data is absent the workbook still builds and points out exactly
which download would unlock each blank tab (no import failure).
"""

import argparse
import os
import sys

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE_DIR)

from generate_report import clean_label, num, pct, to_float  # noqa: E402

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
    "Unit-value (price) positioning needs the Trade Map \"Time Series\" "
    "download for Kenya's exported products (it carries Quantity columns). "
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
    _style_header(ws, r, 8, ["Rank", "HS", "Product", "Latest (USD m)",
                              "Share", "5-yr CAGR", "Status", "Promotion note"])
    r += 1
    latest_col = max(0, len(years) - 1)
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
    _style_header(ws, r, 7, ["Rank", "Supplier", "Latest (USD bn)", "Share",
                             "5-yr CAGR", "Position", "Action"])
    r += 1
    latest_col = max(0, len(years) - 1)
    rows = []
    for d in items:
        yrs = d.get("years") or []
        latest = to_float(yrs[latest_col]) if latest_col < len(yrs) else None
        rows.append((d, latest))
    rows.sort(key=lambda x: -(x[1] or 0))
    total = (t1.get("total") or {}).get("values") or []
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
    _fit(ws, [8, 34, 13, 9, 10, 24, 44])


# ---------------------------------------------------------------------------
# Tier 2: unit-value / price positioning (needs Quantity data)
# ---------------------------------------------------------------------------
def _find_quantity_workbooks(excel_dir):
    """Return bonuses: files whose name suggests a Trade Map Time Series
    download with Quantity columns."""
    hits = []
    if not excel_dir or not os.path.isdir(excel_dir):
        return hits
    for f in sorted(os.listdir(excel_dir)):
        low = f.lower()
        if low.endswith((".xlsx", ".xls")) and (
                "time" in low or "series" in low or "quantity" in low
                or "exported_by_kenya" in low and "product" in low
                or "by-product" in low):
            hits.append(os.path.join(excel_dir, f))
    return hits


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


def _unit_value_rows(excel_dir, n=20):
    """Best-effort (product, unit_value) from a Time Series workbook."""
    for path in _find_quantity_workbooks(excel_dir):
        try:
            wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
        except Exception:
            continue
        ws = wb.worksheets[0]
        cols = _header_index(ws)
        if not cols.get("value") or not cols.get("quantity"):
            continue
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
            out.append((label, v / q))
        if out:
            return out[:n], os.path.basename(path)
    return None, None


def _price_positioning(ws, excel_dir):
    r = 1
    ws.cell(r, 1, "Unit-value (price) positioning").font = Font(bold=True, size=13)
    r += 2
    rows, src = _unit_value_rows(excel_dir)
    if not rows:
        ws.cell(r, 1, UV_REQUIREMENT_NOTE)
        ws.cell(r, 1).font = Font(italic=True, color="9A1F1F")
        return
    ws.cell(r, 2, "Source workbook: %s" % src).font = Font(italic=True, size=9)
    r += 1
    _style_header(ws, r, 4, ["Product", "Unit value (USD/unit)",
                             "vs World", "Reading"])
    r += 1
    for label, uv in rows:
        note = "Premium quality signal - brand & protect"
        cell = ws.cell(r, 1, clean_label(label))
        cell.alignment = LEFT
        cell = ws.cell(r, 2, round(uv, 2))
        cell.number_format = "#,##0.00"
        cell.alignment = RIGHT
        ws.cell(r, 3, "n/a - needs world UV to compare").alignment = CENTER
        ws.cell(r, 4, note).alignment = LEFT
        r += 1
    ws.cell(r, 1, "Note: compare Kenya's unit values against the same product's "
                  "world-average unit value (Trade Map Time Series for world or "
                  "the partner) to classify price-vs-quality positioning.")
    ws.cell(r, 1).font = Font(italic=True, color="595959")


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------
def build_promotion_analysis(a, cfg, excel_dir, out_path):
    """Write the standalone promotion-positioning workbook (additive only)."""
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    partner = cfg["country"]["name"]

    ws = wb.create_sheet("Kenya-%s Mix" % partner[:24])
    _kenya_partner_mix(ws, a, partner)
    _fit(ws, [6, 8, 42, 13, 9, 10, 20, 44])

    ws = wb.create_sheet("Competitor Map")
    _partner_import_structure(ws, a, partner)

    ws = wb.create_sheet("Price Positioning")
    _price_positioning(ws, excel_dir)
    _fit(ws, [44, 16, 22, 40])

    wb.save(out_path)
    return out_path


def build_from_dir(cfg, excel_dir, out_path):
    """Load the goods Analysis and write the promotion workbook (webapp path)."""
    from generate_report import Analysis
    a = Analysis(cfg).load(excel_dir)
    return build_promotion_analysis(a, cfg, excel_dir, out_path)


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