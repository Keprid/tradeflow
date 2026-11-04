#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
generate_competitor_analysis.py
===============================

Competitor analysis for the Goods Trade Flow pipeline.

The module reads Trade Map "imports from world by exporter" downloads -- the
"List of supplying markets for a product imported by <market>" query -- and
builds the KEPROBA competitor-analysis table in one of two layouts, chosen
automatically from the data supplied:

* **Per product** -- one importing market, one workbook per HS product, e.g.
  ``nigerias-imports-from-world-by-exporter_0902.xlsx``.  Products are ranked
  by the value imported and each block lists the product's top exporters.

* **Per Kenya export market** -- one HS product line, one workbook per
  importing market, e.g. ``netherlands-imports-from-world-by-exporter_0603.xlsx``,
  ``united-kingdoms-imports-from-world-by-exporter_0603.xlsx`` and so on for
  cut flowers (HS 0603).  Only the markets Kenya actually supplies are kept,
  ranked by Kenya's exports to them; each block then shows the top exporters
  competing with Kenya in that market.  This is the layout shown in
  ``Samplefiles/Competitor Analysis/competitoranalysis.jpg``.

Deliverables (matching the picture template):

    COMPETITOR ANALYSIS.png   paste-ready table image (the report embeds its
                              tables as pictures)
    COMPETITOR ANALYSIS.docx  editable Word table
    COMPETITOR ANALYSIS.xlsx  editable Excel workbook

Conventions (as in the template picture):
  * in the per-market layout the importing markets are the destinations Kenya
    supplies, ranked 1..n by Kenya's exports to them (descending); in the
    per-product layout products are ranked by the value imported;
  * each block shows its top exporters by value with their rank, value and
    share of the analysed market's imports of the product;
  * Kenya's row is always shown and printed in red: when she is outside the
    top N her row replaces the last slot and carries her true rank;
  * the per-product layout keeps the "Tariff Applied %" column, left empty;
    the per-market layout follows the sample picture and omits it.

When the review year is not available in a download (some markets report a
year later than others) the latest year present in that workbook is used; the
table header shows the most recent year across the set and the note explains
the mixed coverage.

Usage:
    python generate_competitor_analysis.py --data TEKTAT --out TEKTAT
    python generate_competitor_analysis.py --data "CA files" --out output
"""

import argparse
import os
import re
import sys
from collections import Counter

import openpyxl
from openpyxl.styles import Alignment, Border, Font, Side
from openpyxl.utils import get_column_letter

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE_DIR)

from country_names import display_name, fix_label  # noqa: E402

KENYA_RED = "FF0000"
GRID = "808080"

# Trade Map naming: "<market>-imports-from-world-by-exporter_<HS>[ (n)].xlsx"
_SUPPLIER_RE = re.compile(
    r"imports-from-world-by-exporter[_\s-]*'?(\d{4,6})(?:\s*\(\d+\))?\.xlsx$",
    re.IGNORECASE)
_YEAR_HDR_RE = re.compile(r"^\s*'?(\d{4})\b")
# Reporter prefix in the same naming convention, used as a fallback when a
# workbook does not carry a reporter label.
_REPORTER_SLUG_RE = re.compile(
    r"^(?P<slug>.+?)-imports-from-world-by-exporter", re.IGNORECASE)


def is_supplier_download(filename):
    """True when ``filename`` is an ITC supplying-markets download."""
    return bool(_SUPPLIER_RE.search(os.path.basename(str(filename))))


# ---------------------------------------------------------------------------
# Input
# ---------------------------------------------------------------------------
def find_supplier_paths(data_dir):
    """All supplying-markets workbooks under ``data_dir``."""
    hits = []
    for root, _dirs, names in os.walk(data_dir):
        for name in names:
            if _SUPPLIER_RE.search(name):
                hits.append(os.path.join(root, name))
    return sorted(set(hits))


def find_product_supplier_files(data_dir):
    """All per-product supplying-markets workbooks under ``data_dir``.

    Returns a list of ``(path, hs_code)`` sorted by HS code, one entry per
    product code (duplicate downloads -- e.g. ``_0902 (1).xlsx`` -- are
    collapsed, the longer workbook winning).

    Note: this helper is product-oriented and collapses files that share an HS
    code.  When the set is one product line across many importing markets use
    :func:`find_supplier_paths` / :func:`load_markets` instead.
    """
    hits = {}
    for path in find_supplier_paths(data_dir):
        m = _SUPPLIER_RE.search(os.path.basename(path))
        if not m:
            continue
        code = m.group(1)
        prev = hits.get(code)
        if prev is None or _row_count(path) > _row_count(prev[0]):
            hits[code] = (path, code)
    return [hits[c] for c in sorted(hits)]


def _row_count(path):
    try:
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        ws = wb.worksheets[0]
        n = ws.max_row or 0
        wb.close()
        return n
    except Exception:
        return 0


def _header_map(ws):
    """Locate the header row of a Trade Map 'Data' sheet and map columns."""
    for ri, row in enumerate(ws.iter_rows(min_row=1, max_row=10, values_only=True),
                            start=1):
        labels = [str(v).strip().lower() if v is not None else ""
                  for v in row]
        if "productcd" in labels and "partnerlabel" in labels:
            cols = {lab: i for i, lab in enumerate(labels) if lab}
            years = {}
            for i, v in enumerate(row):
                if v is None:
                    continue
                m = _YEAR_HDR_RE.match(str(v))
                if m:
                    years[int(m.group(1))] = i
            return ri, cols, years
    raise ValueError("no Trade Map header row found")


def _slug_to_name(path):
    m = _REPORTER_SLUG_RE.search(os.path.basename(str(path)))
    if not m:
        return ""
    return " ".join(w.capitalize() for w in re.split(r"[-_]+", m.group("slug")))


def _parse_supplier(path, year=None):
    """Parse one supplying-markets workbook.

    Returns ``reporter`` metadata plus one record per product found in the
    workbook.  Each product record carries ``code``/``label``/``world`` and an
    ``exporters`` list (rank, label, value, share, kenya).
    """
    wb = openpyxl.load_workbook(path, data_only=True)
    try:
        ws = wb.worksheets[0]
        hdr_row, cols, years = _header_map(ws)
        if not years:
            raise ValueError("no year columns found in %s"
                             % os.path.basename(path))
        rev = year if (year in years) else max(years)
        ycol = years[rev]
        c_partner = cols.get("partnerlabel")
        c_code = cols.get("productcd")
        c_label = cols.get("productlabel")
        c_pcode = cols.get("partnercd")
        c_reporter = cols.get("reporterlabel")
        c_rep_cd = cols.get("reportercd")
        reporter = ""
        reporter_cd = ""
        rows_by_product = {}
        for row in ws.iter_rows(min_row=hdr_row + 1, values_only=True):
            if row is None or all(v is None for v in row):
                continue
            if not reporter and c_reporter is not None:
                reporter = (str(row[c_reporter]).strip()
                            if row[c_reporter] is not None else "")
            if not reporter_cd and c_rep_cd is not None:
                reporter_cd = (str(row[c_rep_cd]).strip()
                               if row[c_rep_cd] is not None else "")
            pcode = (str(row[c_code]).strip()
                     if c_code is not None and row[c_code] is not None
                     else _SUPPLIER_RE.search(os.path.basename(path)).group(1))
            val = row[ycol]
            val = float(val) if isinstance(val, (int, float)) else None
            entry = {
                "partner": (str(row[c_partner]).strip()
                            if c_partner is not None and row[c_partner] is not None
                            else ""),
                "partner_cd": (str(row[c_pcode]).strip()
                               if c_pcode is not None and row[c_pcode] is not None
                               else ""),
                "label": (str(row[c_label]).strip()
                          if c_label is not None and row[c_label] is not None
                          else ""),
                "value": val,
                "reporter": reporter,
                "reporter_cd": reporter_cd,
            }
            rows_by_product.setdefault(pcode, []).append(entry)
    finally:
        wb.close()

    if not reporter:
        reporter = _slug_to_name(path)
    records = [_product_record(pcode, entries, rev, path)
               for pcode, entries in rows_by_product.items()]
    return {"path": path, "reporter": reporter, "reporter_cd": reporter_cd,
            "year": rev, "years": years, "products": records}


def _product_record(pcode, entries, year, path):
    world_rows = [e for e in entries if _is_world(e)]
    exporters = [e for e in entries
                 if not _is_world(e) and (e["value"] or 0.0) > 0.0]
    exporters.sort(key=lambda e: -(e["value"] or 0.0))
    world = world_rows[0]["value"] if world_rows else None
    if world is None:
        world = sum((e["value"] or 0.0) for e in exporters) or None
    rows = []
    for i, e in enumerate(exporters, start=1):
        share = (e["value"] / world) if world else None
        rows.append({"rank": i,
                     "label": display_name(fix_label(e["partner"])),
                     "value": e["value"], "share": share,
                     "kenya": _is_kenya(e)})
    label = next((e["label"] for e in entries if e["label"]), "")
    reporter = next((e["reporter"] for e in entries if e.get("reporter")), "")
    reporter_cd = next((e["reporter_cd"] for e in entries
                        if e.get("reporter_cd")), "")
    return {"code": pcode, "label": label, "year": year,
            "world": world, "exporters": rows, "source_file": path,
            "reporter": reporter, "reporter_cd": reporter_cd}


def _is_world(entry):
    return entry["partner_cd"] == "000" or \
        entry["partner"].strip().lower() in ("world", "world total")


def _is_kenya(entry):
    return "kenya" in entry["partner"].strip().lower()


def load_products(files, year=None):
    """Parse the per-product workbooks into product records.

    ``files`` may be ``(path, hs_code)`` pairs (the historical form) or plain
    paths.  Each record: ``code``, ``label``, ``world`` (the market's imports
    of the product from the world), ``exporters`` (sorted by value, each with
    ``rank``/``label``/``value``/``share``/``kenya``) and ``source_file``.
    ``year`` is the review year; when omitted each workbook's latest year is
    used.
    """
    products = []
    seen = {}
    for item in files:
        if isinstance(item, (tuple, list)):
            path, code = item[0], item[1]
        else:
            path, code = item, None
        parsed = _parse_supplier(path, year=year)
        candidates = parsed["products"]
        if code is not None:
            matched = [p for p in candidates if p["code"] == str(code)]
            candidates = matched or candidates
        for rec in candidates:
            key = (rec.get("reporter_cd") or rec.get("reporter", ""),
                   rec["code"])
            prev = seen.get(key)
            if prev is None or (rec["world"] or 0.0) > (prev["world"] or 0.0):
                seen[key] = rec
    products = list(seen.values())
    products.sort(key=lambda p: -(p["world"] or 0.0))
    for i, p in enumerate(products, start=1):
        p["rank"] = i
    return products


def load_markets(files, year=None, product_code=None, kenya_only=True):
    """Parse one product line across the markets Kenya supplies.

    ``files`` is a list of paths (or ``(path, hs_code)`` pairs).  Each market
    record contains ``country``/``country_cd``, the analysed ``code`` and
    ``label``, the value imported from the world (``world``), Kenya's exports
    to that market (``kenya_value`` / ``kenya_share``), the review ``year`` and
    the ``exporters`` list.  By default only markets where Kenya exported the
    product are kept, ranked by Kenya's exports to them (the same ordering as
    the lead export-destination table).  Set ``kenya_only=False`` to keep every
    market and rank by total imports instead.
    """
    markets = []
    seen = {}
    for item in files:
        path = item[0] if isinstance(item, (tuple, list)) else item
        parsed = _parse_supplier(path, year=year)
        candidates = parsed["products"]
        if product_code is not None:
            candidates = [p for p in candidates
                          if p["code"] == str(product_code)]
        if not candidates:
            continue
        rec = max(candidates, key=lambda p: p["world"] or 0.0)
        country = _clean_market(parsed["reporter"] or rec.get("reporter", ""))
        key = parsed["reporter_cd"] or country.lower()
        kenya = next((r for r in rec["exporters"] if r["kenya"]), None)
        market = {
            "country": country,
            "country_raw": parsed["reporter"],
            "country_cd": parsed["reporter_cd"],
            "code": rec["code"],
            "label": rec["label"],
            "year": rec["year"],
            "world": rec["world"],
            "kenya_value": kenya["value"] if kenya else None,
            "kenya_share": kenya["share"] if kenya else None,
            "exporters": rec["exporters"],
            "source_file": path,
        }
        prev = seen.get(key)
        if prev is None or ((market["kenya_value"] or 0.0),
                            (market["world"] or 0.0)) > (
                                (prev["kenya_value"] or 0.0),
                                (prev["world"] or 0.0)):
            seen[key] = market
    markets = list(seen.values())
    if kenya_only:
        markets = [m for m in markets if (m["kenya_value"] or 0.0) > 0.0]
        markets.sort(key=lambda m: -(m["kenya_value"] or 0.0))
    else:
        markets.sort(key=lambda m: -(m["world"] or 0.0))
    for i, m in enumerate(markets, start=1):
        m["rank"] = i
    return markets


def _clean_market(name):
    """Tidy a reporter label without shortening it (the sample picture keeps
    "United Kingdom" / "United Arab Emirates" spelled out)."""
    text = " ".join(str(name or "").split())
    return text or "Unknown"


# ---------------------------------------------------------------------------
# Table model
# ---------------------------------------------------------------------------
HEADER = ["Rank in {Y}", "Product code", "Product label",
          "Value imported in {Y}", "Rank in {Y}", "Top exporters",
          "Value exported in {Y}", "Share in {Y} %", "Tariff Applied %"]

COUNTRY_HEADER = ["Rank in {Y}", "Importers", "Value imported in {Y}",
                  "Rank in {Y}", "Top Exporters", "Value exported in {Y}",
                  "Share in {Y} %"]


def product_blocks(products, top=5):
    """One block per product: ``top`` exporter rows, with Kenya guaranteed.

    When Kenya is outside the top ``top`` her row replaces the last slot and
    keeps her true rank.
    """
    blocks = []
    for p in products:
        blocks.append({"product": p,
                       "rows": _top_exporter_rows(p["exporters"], top)})
    return blocks


def market_blocks(markets, top=5):
    """One block per market Kenya supplies, with the same Kenya guarantee."""
    blocks = []
    for m in markets:
        blocks.append({"market": m,
                       "rows": _top_exporter_rows(m["exporters"], top)})
    return blocks


def _top_exporter_rows(exporters, top):
    rows = list(exporters[:top])
    kenya = next((e for e in exporters if e["kenya"]), None)
    if kenya is not None and not any(r["kenya"] for r in rows):
        if rows:
            rows = rows[:top - 1] + [kenya]
        else:
            rows = [kenya]
    return rows


def _fmt_value(v):
    return "" if v is None else "{:,.0f}".format(v)


def _fmt_share(s):
    return "" if s is None else "{:.1f}%".format(s * 100.0)


def _shorten(text, n):
    text = " ".join(str(text).split())
    return text if len(text) <= n else text[:n - 1].rstrip() + "\u2026"


def _product_layout():
    return {
        "headers": list(HEADER),
        "widths": [0.055, 0.065, 0.235, 0.100, 0.055, 0.150, 0.105, 0.075,
                   0.070],
        "wrap_chars": [3, 5, 34, 9, 3, 16, 9, 6, 6],
        "aligns": ["center", "center", "left", "right", "center", "left",
                   "right", "right", "right"],
        "group_cols": 4,
        "group_label_col": 2,
        "row_label_col": 5,
        "group_valign": ["center", "center", "top", "center"],
        "number_cols": [3, 6],
        "percent_cols": [7],
        "docx_widths_cm": [1.15, 1.35, 5.6, 1.9, 1.15, 3.6, 2.05, 1.45, 1.5],
        "xlsx_widths": [8, 11, 46, 14, 8, 24, 15, 11, 11],
    }


def _country_layout():
    return {
        "headers": list(COUNTRY_HEADER),
        "widths": [0.055, 0.240, 0.135, 0.055, 0.190, 0.135, 0.090],
        "wrap_chars": [3, 30, 11, 3, 18, 11, 7],
        "aligns": ["center", "left", "right", "center", "left", "right",
                   "right"],
        "group_cols": 3,
        "group_label_col": 1,
        "row_label_col": 4,
        "group_valign": ["center", "top", "center"],
        "number_cols": [2, 5],
        "percent_cols": [6],
        "docx_widths_cm": [1.15, 5.0, 2.5, 1.15, 4.8, 2.5, 1.7],
        "xlsx_widths": [8, 36, 17, 8, 32, 17, 13],
    }


def _product_table(products, top=5):
    blocks = []
    for b in product_blocks(products, top=top):
        p = b["product"]
        group = [p["rank"], "'%s" % p["code"],
                 _shorten(p["label"], 95), _fmt_value(p["world"])]
        group_values = [p["rank"], "'%s" % p["code"], p["label"], p["world"]]
        rows = [{"cells": [r["rank"], r["label"], _fmt_value(r["value"]),
                           _fmt_share(r["share"]), ""],
                 "values": [r["rank"], r["label"], r["value"], r["share"],
                            None],
                 "red": bool(r["kenya"])}
                for r in b["rows"]]
        if not rows:
            rows = [{"cells": [""] * 5, "values": [""] * 5, "red": False}]
        blocks.append({"group": group, "group_values": group_values,
                       "rows": rows})
    return blocks, _product_layout()


def _country_table(markets, top=5):
    blocks = []
    for b in market_blocks(markets, top=top):
        m = b["market"]
        group = [m["rank"], m["country"], _fmt_value(m["world"])]
        group_values = [m["rank"], m["country"], m["world"]]
        rows = [{"cells": [r["rank"], r["label"], _fmt_value(r["value"]),
                           _fmt_share(r["share"])],
                 "values": [r["rank"], r["label"], r["value"], r["share"]],
                 "red": bool(r["kenya"])}
                for r in b["rows"]]
        if not rows:
            rows = [{"cells": [""] * 4, "values": [""] * 4, "red": False}]
        blocks.append({"group": group, "group_values": group_values,
                       "rows": rows})
    return blocks, _country_layout()


def narrative_bullets(products, partner, top=5):
    """Short auto-generated reading notes for a per-product analysis."""
    out = []
    blocks = product_blocks(products, top=top)
    lead = []
    for b in blocks:
        rows = b["rows"]
        if rows and rows[0]["kenya"]:
            lead.append(b["product"]["code"])
    if lead:
        out.append("Kenya is the leading supplier of %s to %s in the review "
                   "year (HS %s)." % (_product_phrase(lead), partner,
                                      ", ".join(lead)))
    ranked = []
    for b in blocks:
        kenya = next((r for r in b["rows"] if r["kenya"]), None)
        if kenya and kenya["share"] is not None:
            ranked.append((kenya["share"], b["product"], kenya))
    if ranked:
        ranked.sort(key=lambda t: -t[0])
        share, prod, kenya = ranked[0]
        out.append("Kenya's strongest position is HS %s (%s), where it "
                   "supplies %.1f%% of %s's imports (rank %d)."
                   % (prod["code"], _shorten(prod["label"], 45), share * 100.0,
                      partner, kenya["rank"]))
    biggest = max(blocks, key=lambda b: b["product"]["world"] or 0.0) \
        if blocks else None
    if biggest:
        p = biggest["product"]
        kenya = next((r for r in biggest["rows"] if r["kenya"]), None)
        out.append("The largest import market in the set is HS %s (%s), worth "
                   "%s USD thousand; Kenya %s." %
                   (p["code"], _shorten(p["label"], 45), _fmt_value(p["world"]),
                    ("ranks %d with %.1f%%" % (kenya["rank"],
                                               (kenya["share"] or 0.0) * 100.0))
                    if kenya else "is not among the suppliers listed"))
    return out


def narrative_bullets_markets(markets, product_label, product_code, top=5):
    """Short auto-generated reading notes for a per-market analysis."""
    out = []
    blocks = market_blocks(markets, top=top)
    if not blocks:
        return out
    biggest = blocks[0]
    m = biggest["market"]
    kenya = next((r for r in biggest["rows"] if r["kenya"]), None)
    if kenya:
        out.append("Kenya's largest market for %s (HS %s) is %s: Kenya "
                   "exported %s USD thousand there in %s, %.1f%% of the "
                   "market's %s USD thousand imports (supplier rank %d)." %
                   (_shorten(product_label, 45), product_code, m["country"],
                    _fmt_value(kenya["value"]), m["year"],
                    (kenya["share"] or 0.0) * 100.0, _fmt_value(m["world"]),
                    kenya["rank"]))
    else:
        out.append("The largest market in the set for %s (HS %s) is %s, "
                   "worth %s USD thousand in %s." %
                   (_shorten(product_label, 45), product_code, m["country"],
                    _fmt_value(m["world"]), m["year"]))
    lead = [b["market"]["country"] for b in blocks
            if b["rows"] and b["rows"][0]["kenya"]]
    if lead:
        out.append("Kenya is the leading supplier in %s."
                   % ", ".join(lead))
    ranked = []
    for b in blocks:
        kenya = next((r for r in b["rows"] if r["kenya"]), None)
        if kenya and kenya["share"] is not None:
            ranked.append((kenya["share"], b["market"], kenya))
    if ranked:
        ranked.sort(key=lambda t: -t[0])
        share, market, kenya = ranked[0]
        out.append("Kenya's strongest market is %s, where it supplies %.1f%% "
                   "of imports (rank %d)." %
                   (market["country"], share * 100.0, kenya["rank"]))
    return out


def _product_phrase(codes):
    return "the product" if len(codes) == 1 else "the products"


# ---------------------------------------------------------------------------
# PNG (paste-ready picture, as in the template)
# ---------------------------------------------------------------------------
def render_table_png(blocks, layout, out_png, year=None, dpi=220):
    """Render any block table described by ``layout`` as a PNG image."""
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    from matplotlib.patches import Rectangle
    import textwrap

    headers = [h.format(Y=year) for h in layout["headers"]]
    widths = list(layout["widths"])
    total = sum(widths)
    widths = [w / total for w in widths]
    aligns = layout["aligns"]
    wrap_chars = layout["wrap_chars"]
    group_cols = layout["group_cols"]
    group_label_col = layout["group_label_col"]
    row_label_col = layout["row_label_col"]
    group_valign = layout.get("group_valign") or ["center"] * group_cols
    n_cols = len(headers)

    def wrap(text, chars):
        lines = textwrap.wrap(str(text), chars, break_long_words=False,
                              break_on_hyphens=False) or [""]
        return lines

    row_h = 0.24
    block_metrics = []
    for b in blocks:
        label_lines = len(wrap(_shorten(b["group"][group_label_col], 95),
                               wrap_chars[group_label_col]))
        name_lines = max([len(wrap(r["cells"][row_label_col - group_cols],
                                   wrap_chars[row_label_col]))
                          for r in b["rows"]] or [1])
        lines = max(label_lines, name_lines)
        block_h = max(row_h * len(b["rows"]),
                      0.20 * lines + 0.06 * (len(b["rows"]) - 1))
        block_metrics.append(block_h)

    fig_h = 0.34 + sum(block_metrics) + 0.18
    fig_w = 10.2
    fig = plt.figure(figsize=(fig_w, fig_h), dpi=dpi)
    ax = fig.add_axes([0, 0, 1, 1])
    ax.set_xlim(0, fig_w)
    ax.set_ylim(0, fig_h)
    ax.axis("off")

    font = _serif()
    xs = [0.0]
    for w in widths:
        xs.append(xs[-1] + w)

    def cell_rect(c, y_top, h):
        x0 = xs[c] * fig_w
        w = (xs[c + 1] - xs[c]) * fig_w
        y0 = fig_h - y_top - h
        return x0, y0, w, h

    def draw_cell(c, y_top, h, text, bold=False, red=False,
                  align="left", valign="center", size=6.6):
        x0, y0, w, h = cell_rect(c, y_top, h)
        ax.add_patch(Rectangle((x0, y0), w, h, fill=False,
                               edgecolor="black", linewidth=0.5))
        if text == "" or text is None:
            return
        lines = wrap(text, wrap_chars[c])
        n = len(lines)
        lh = size / 72.0 * 1.25
        block_h = lh * n
        if valign == "top":
            ty = y0 + h - 0.045 - lh * 0.85
        else:
            ty = y0 + (h + block_h) / 2.0 - lh * 0.85
        tx = x0 + 0.035
        ha = "left"
        if align == "center":
            tx = x0 + w / 2.0
            ha = "center"
        elif align == "right":
            tx = x0 + w - 0.035
            ha = "right"
        color = "#%s" % KENYA_RED if red else "black"
        for line in lines:
            ax.text(tx, ty, line, ha=ha, va="baseline", fontsize=size,
                    fontfamily=font, fontweight="bold" if bold else "normal",
                    color=color)
            ty -= lh

    def draw_cell_merged(c, y_top, h, text, align, bold=False, size=6.6,
                         valign="center"):
        x0, y0, w, h = cell_rect(c, y_top, h)
        if text == "" or text is None:
            return
        lines = wrap(text, wrap_chars[c])
        lh = size / 72.0 * 1.25
        if valign == "top":
            ty = y0 + h - 0.045 - lh * 0.85
        else:
            ty = y0 + (h + lh * len(lines)) / 2.0 - lh * 0.85
        tx = x0 + 0.035
        ha = "left"
        if align == "center":
            tx = x0 + w / 2.0
            ha = "center"
        elif align == "right":
            tx = x0 + w - 0.035
            ha = "right"
        for line in lines:
            ax.text(tx, ty, line, ha=ha, va="baseline", fontsize=size,
                    fontfamily=font, fontweight="bold" if bold else "normal",
                    color="black")
            ty -= lh

    # header
    y = 0.0
    h0 = 0.34
    for c, text in enumerate(headers):
        draw_cell(c, y, h0, text, bold=True, align="center", size=6.4)
    y += h0

    for bi, b in enumerate(blocks):
        rows = b["rows"]
        top_y = y
        block_h = block_metrics[bi]
        row_h_each = block_h / max(len(rows), 1)

        for c in range(group_cols):
            x0, y0, w, hh = cell_rect(c, top_y, block_h)
            ax.add_patch(Rectangle((x0, y0), w, hh, fill=False,
                                   edgecolor="black", linewidth=0.5))
            draw_cell_merged(c, top_y, block_h, b["group"][c], aligns[c],
                             valign=group_valign[c])

        yy = top_y
        for r in rows:
            red = r.get("red", False)
            for j, c in enumerate(range(group_cols, n_cols)):
                draw_cell(c, yy, row_h_each, r["cells"][j], red=red,
                          align=aligns[c], size=6.6)
            yy += row_h_each
        y += block_h

    os.makedirs(os.path.dirname(os.path.abspath(out_png)) or ".", exist_ok=True)
    fig.savefig(out_png, dpi=dpi, facecolor="white")
    plt.close(fig)
    return out_png


def render_png(products, out_png, year=None, partner_label="", top=5):
    """Render the per-product competitor-analysis table as a PNG image."""
    blocks, layout = _product_table(products, top=top)
    year = year or (products[0]["year"] if products else "")
    return render_table_png(blocks, layout, out_png, year=year)


def render_png_markets(markets, out_png, year=None, product_label="", top=5):
    """Render the per-importing-market competitor-analysis table as PNG."""
    blocks, layout = _country_table(markets, top=top)
    year = year or (markets[0]["year"] if markets else "")
    return render_table_png(blocks, layout, out_png, year=year)


def _serif():
    import matplotlib.font_manager as fm
    installed = {f.name for f in fm.fontManager.ttflist}
    for name in ("Liberation Serif", "Times New Roman", "DejaVu Serif"):
        if name in installed:
            return name
    return "DejaVu Serif"


# ---------------------------------------------------------------------------
# DOCX
# ---------------------------------------------------------------------------
def write_table_docx(blocks, layout, out_docx, year=None, title="Competitor Analysis",
                     caption="", note="", source=None):
    from docx import Document
    from docx.enum.table import WD_TABLE_ALIGNMENT
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Cm, Pt, RGBColor

    headers = [h.format(Y=year) for h in layout["headers"]]
    aligns = layout["aligns"]
    group_cols = layout["group_cols"]
    n_cols = len(headers)
    source = source or ("Source: International Trade Centre (ITC) "
                        "Database; Compiled by KEPROBA.")

    doc = Document()
    sec = doc.sections[0]
    sec.left_margin = sec.right_margin = Cm(1.2)
    sec.top_margin = sec.bottom_margin = Cm(1.2)

    title_p = doc.add_paragraph(title)
    title_p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = title_p.runs[0] if title_p.runs else title_p.add_run()
    run.bold = True
    run.font.size = Pt(12)

    cap = doc.add_paragraph(caption or ("Table: Competitor Analysis"
                                        + (" - %s" % year if year else "")))
    cap.runs[0].italic = True
    cap.runs[0].font.size = Pt(9)

    n_data_rows = sum(len(b["rows"]) for b in blocks)
    table = doc.add_table(rows=1 + n_data_rows, cols=n_cols)
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    def set_cell(cell, text, bold=False, red=False, align="left", size=7):
        cell.text = ""
        para = cell.paragraphs[0]
        para.paragraph_format.space_before = Pt(0)
        para.paragraph_format.space_after = Pt(0)
        if align == "center":
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        elif align == "right":
            para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        run = para.add_run(str(text))
        run.font.size = Pt(size)
        run.font.name = "Times New Roman"
        run.bold = bold
        if red:
            run.font.color.rgb = RGBColor(0xFF, 0x00, 0x00)

    def vmerge(cell_a, cell_b):
        tc_a = cell_a._tc
        tc_b = cell_b._tc
        tc_prop = tc_a.get_or_add_tcPr()
        merge = OxmlElement("w:vMerge")
        merge.set(qn("w:val"), "restart")
        tc_prop.append(merge)
        tc_prop_b = tc_b.get_or_add_tcPr()
        merge_b = OxmlElement("w:vMerge")
        tc_prop_b.append(merge_b)

    for c, text in enumerate(headers):
        set_cell(table.rows[0].cells[c], text, bold=True, align="center",
                 size=6.5)

    ri = 1
    for b in blocks:
        start = ri
        for r in b["rows"]:
            red = r.get("red", False)
            for j, c in enumerate(range(group_cols, n_cols)):
                set_cell(table.rows[ri].cells[c], r["cells"][j], red=red,
                         align=aligns[c])
            ri += 1
        for c in range(group_cols):
            cell = table.rows[start].cells[c]
            set_cell(cell, b["group"][c], align=aligns[c])
            for rr in range(start + 1, start + len(b["rows"])):
                vmerge(cell, table.rows[rr].cells[c])

    for row in table.rows:
        for c, w in enumerate(layout["docx_widths_cm"]):
            row.cells[c].width = Cm(w)

    note_p = doc.add_paragraph()
    note_run = note_p.add_run(note)
    note_run.italic = True
    note_run.font.size = Pt(8)

    src = doc.add_paragraph()
    src_run = src.add_run(source)
    src_run.italic = True
    src_run.font.size = Pt(8)

    os.makedirs(os.path.dirname(os.path.abspath(out_docx)) or ".", exist_ok=True)
    doc.save(out_docx)
    return out_docx


def write_docx(products, out_docx, year=None, partner_label="", top=5):
    blocks, layout = _product_table(products, top=top)
    year = year or (products[0]["year"] if products else "")
    caption = "Table: Competitor Analysis" + (
        " - %s, %s" % (partner_label, year) if partner_label else "")
    note = ("Note: values are in USD thousand. Kenya's row is highlighted in "
            "red; where Kenya is outside the top exporters her row shows her "
            "true rank. Tariff data not provided in the source downloads.")
    return write_table_docx(blocks, layout, out_docx, year=year, caption=caption,
                            note=note)


def write_docx_markets(markets, out_docx, year=None, product_label="", top=5):
    blocks, layout = _country_table(markets, top=top)
    year = year or (markets[0]["year"] if markets else "")
    if product_label:
        caption = "Table: Competitor Analysis - %s (HS %s), %s" % (
            _shorten(product_label, 80), markets[0]["code"] if markets else "",
            year)
    else:
        caption = "Table: Competitor Analysis - %s" % year
    note = ("Note: values are in USD thousand and cover the latest available "
            "year for each market. Only markets where Kenya exported the "
            "product are shown, ranked by Kenya's exports to them. Kenya's "
            "row is highlighted in red; where Kenya is outside the top "
            "exporters her row shows her true rank.")
    return write_table_docx(blocks, layout, out_docx, year=year, caption=caption,
                            note=note)


# ---------------------------------------------------------------------------
# XLSX
# ---------------------------------------------------------------------------
def write_table_xlsx(blocks, layout, out_xlsx, year=None, note="", source=None):
    headers = [h.format(Y=year) for h in layout["headers"]]
    aligns = layout["aligns"]
    group_cols = layout["group_cols"]
    n_cols = len(headers)
    number_cols = set(layout.get("number_cols") or [])
    percent_cols = set(layout.get("percent_cols") or [])
    source = source or ("Source: International Trade Centre (ITC) "
                        "Database; Compiled by KEPROBA.")

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Competitor Analysis"

    thin = Side(style="thin", color=GRID)
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    hdr_font = Font(name="Century Gothic", bold=True)
    body_font = Font(name="Century Gothic")
    red_font = Font(name="Century Gothic", color=KENYA_RED)
    left = Alignment(horizontal="left", vertical="center", wrap_text=True)
    right = Alignment(horizontal="right", vertical="center")
    center = Alignment(horizontal="center", vertical="center")
    xl_aligns = {"left": left, "right": right, "center": center}

    for c, text in enumerate(headers, start=1):
        cell = ws.cell(1, c, text)
        cell.font = hdr_font
        cell.border = border
        cell.alignment = center

    ri = 2
    for b in blocks:
        start = ri
        for r in b["rows"]:
            font = red_font if r.get("red") else body_font
            values = r.get("values") or r["cells"]
            for j, c in enumerate(range(group_cols, n_cols)):
                val = values[j]
                cell = ws.cell(ri, c + 1, val)
                cell.font = font
                cell.border = border
                cell.alignment = xl_aligns[aligns[c]]
                if isinstance(val, (int, float)):
                    if c in percent_cols:
                        cell.number_format = "0.0%"
                    elif c in number_cols:
                        cell.number_format = "#,##0"
            ri += 1
        for c in range(group_cols):
            values = b.get("group_values") or b["group"]
            val = values[c]
            cell = ws.cell(start, c + 1, val)
            cell.font = body_font
            cell.border = border
            cell.alignment = xl_aligns[aligns[c]]
            if isinstance(val, (int, float)):
                if c in percent_cols:
                    cell.number_format = "0.0%"
                elif c in number_cols:
                    cell.number_format = "#,##0"
            if c == 1 and n_cols == 9:
                cell.number_format = "@"
            if len(b["rows"]) > 1:
                ws.merge_cells(start_row=start, start_column=c + 1,
                               end_row=start + len(b["rows"]) - 1,
                               end_column=c + 1)
                for rr in range(start + 1, start + len(b["rows"])):
                    ws.cell(rr, c + 1).border = border

    for c, w in enumerate(layout["xlsx_widths"], start=1):
        ws.column_dimensions[get_column_letter(c)].width = w

    nr = ws.max_row + 2
    ws.cell(nr, 1, note).font = Font(name="Century Gothic", italic=True,
                                     size=9)
    ws.cell(nr + 1, 1, source).font = Font(name="Century Gothic", italic=True,
                                           size=9)

    os.makedirs(os.path.dirname(os.path.abspath(out_xlsx)) or ".", exist_ok=True)
    wb.save(out_xlsx)
    return out_xlsx


def write_xlsx(products, out_xlsx, year=None, partner_label="", top=5):
    blocks, layout = _product_table(products, top=top)
    year = year or (products[0]["year"] if products else "")
    note = ("Note: values in USD thousand; Kenya's row in red; where Kenya is "
            "outside the top exporters her row carries her true rank. Tariff "
            "data not provided in the source downloads.")
    return write_table_xlsx(blocks, layout, out_xlsx, year=year, note=note)


def write_xlsx_markets(markets, out_xlsx, year=None, product_label="", top=5):
    blocks, layout = _country_table(markets, top=top)
    year = year or (markets[0]["year"] if markets else "")
    note = ("Note: values in USD thousand, latest available year per market. "
            "Only markets where Kenya exported the product are shown, ranked "
            "by Kenya's exports to them. Kenya's row is in red; where Kenya is "
            "outside the top exporters her row carries her true rank.")
    return write_table_xlsx(blocks, layout, out_xlsx, year=year, note=note)


# ---------------------------------------------------------------------------
# Orchestration
# ---------------------------------------------------------------------------
def _reporter_key(info):
    return (info.get("reporter_cd")
            or (info.get("reporter") or "").strip().lower())


def _info_weight(info):
    return sum((r["world"] or 0.0) for r in info["products"])


def _dedupe_parsed(parsed):
    """Collapse duplicate downloads (same reporter + same product list)."""
    by_key = {}
    for info in parsed:
        codes = tuple(sorted(r["code"] for r in info["products"]))
        key = (_reporter_key(info), codes)
        prev = by_key.get(key)
        if prev is None or _info_weight(info) > _info_weight(prev):
            by_key[key] = info
    return list(by_key.values())


def detect_mode(parsed):
    """Choose 'product' or 'country' from the shape of the uploaded data."""
    reporters = {_reporter_key(i) for i in parsed}
    codes = {r["code"] for i in parsed for r in i["products"]}
    if len(reporters) == 1:
        return "product"
    if len(codes) == 1:
        return "country"
    return "country"


def generate(data_dir, out_dir, year=None, top=5,
             basename="COMPETITOR ANALYSIS", mode="auto"):
    paths = find_supplier_paths(data_dir)
    if not paths:
        raise SystemExit("no supplying-markets files "
                         "(*-imports-from-world-by-exporter_<HS>.xlsx) "
                         "found under %s" % data_dir)
    parsed = _dedupe_parsed([_parse_supplier(p, year=year) for p in paths])
    parsed = [i for i in parsed if i["products"]]
    if not parsed:
        raise SystemExit("no product rows found in the supplying-markets "
                         "files under %s" % data_dir)
    if mode == "auto":
        mode = detect_mode(parsed)

    os.makedirs(out_dir, exist_ok=True)
    if mode == "product":
        # Keep the reporter with the most products (a valid product-mode set
        # comes from one importing market).
        counts = Counter(_reporter_key(i) for i in parsed)
        target_rep = counts.most_common(1)[0][0]
        chosen = [i for i in parsed if _reporter_key(i) == target_rep]
        products = []
        seen = {}
        for info in chosen:
            rec = max(info["products"], key=lambda r: r["world"] or 0.0)
            key = rec["code"]
            prev = seen.get(key)
            if prev is None or (rec["world"] or 0.0) > (prev["world"] or 0.0):
                seen[key] = rec
        products = list(seen.values())
        products.sort(key=lambda p: -(p["world"] or 0.0))
        for i, p in enumerate(products, start=1):
            p["rank"] = i
        rev = max((p["year"] for p in products if p.get("year")),
                  default=year or "")
        partner = chosen[0]["reporter"] if chosen else ""
        png = render_png(products, os.path.join(out_dir, basename + ".png"),
                         year=rev, partner_label=partner, top=top)
        docx_p = write_docx(products,
                            os.path.join(out_dir, basename + ".docx"),
                            year=rev, partner_label=partner, top=top)
        xlsx_p = write_xlsx(products,
                            os.path.join(out_dir, basename + ".xlsx"),
                            year=rev, partner_label=partner, top=top)
        return {"mode": "product", "products": products, "markets": [],
                "product": ({"code": products[0]["code"],
                             "label": products[0]["label"]}
                            if len(products) == 1 else None),
                "year": rev, "partner": partner,
                "png": png, "docx": docx_p, "xlsx": xlsx_p}

    # Per Kenya export market (one product line across the markets Kenya serves).
    code_counts = Counter()
    for info in parsed:
        for code in {r["code"] for r in info["products"]}:
            code_counts[code] += 1
    product_code = code_counts.most_common(1)[0][0]
    markets = load_markets([i["path"] for i in parsed], year=year,
                           product_code=product_code, kenya_only=True)
    if not markets:
        raise SystemExit("no markets where Kenya exports HS %s were found "
                         "under %s" % (product_code, data_dir))
    product_label = next((m["label"] for m in markets if m.get("label")), "")
    rev = max((m["year"] for m in markets if m.get("year")),
              default=year or "")
    png = render_png_markets(markets, os.path.join(out_dir, basename + ".png"),
                             year=rev, product_label=product_label, top=top)
    docx_p = write_docx_markets(markets,
                                os.path.join(out_dir, basename + ".docx"),
                                year=rev, product_label=product_label, top=top)
    xlsx_p = write_xlsx_markets(markets,
                                os.path.join(out_dir, basename + ".xlsx"),
                                year=rev, product_label=product_label, top=top)
    return {"mode": "country", "markets": markets, "products": markets,
            "product": {"code": product_code, "label": product_label},
            "year": rev, "partner": "",
            "png": png, "docx": docx_p, "xlsx": xlsx_p}


def main():
    ap = argparse.ArgumentParser(
        description="Generate the competitor analysis table from supplying-"
                    "markets Trade Map downloads (per product or per importing "
                    "market, detected automatically).")
    ap.add_argument("--data", required=True,
                    help="folder holding *-imports-from-world-by-exporter_<HS>.xlsx")
    ap.add_argument("--out", default=None,
                    help="output folder (default: same as --data)")
    ap.add_argument("--year", type=int, default=None,
                    help="review year (default: latest year in the files)")
    ap.add_argument("--top", type=int, default=5,
                    help="exporters shown per block (default 5)")
    ap.add_argument("--mode", choices=["auto", "product", "country"],
                    default="auto",
                    help="table layout (default: detect from the data)")
    ap.add_argument("--basename", default="COMPETITOR ANALYSIS",
                    help="output file basename")
    args = ap.parse_args()
    result = generate(args.data, args.out or args.data, year=args.year,
                      top=args.top, basename=args.basename, mode=args.mode)
    if result["mode"] == "country":
        p = result["product"] or {}
        print("Competitor analysis by Kenya's export market "
              "(%d markets, HS %s, %s):"
              % (len(result["markets"]), p.get("code", ""), result["year"]))
        for m in result["markets"]:
            print("  %d. %s  Kenya exports %s (%.1f%%); market imports %s  (%s)"
                  % (m["rank"], m["country"], _fmt_value(m["kenya_value"]),
                     (m["kenya_share"] or 0.0) * 100.0,
                     _fmt_value(m["world"]), _shorten(m["label"], 50)))
    else:
        print("Competitor analysis by product (%d products, %s):" %
              (len(result["products"]), result["year"]))
        for p in result["products"]:
            print("  %d. HS %s  imported %s  (%s)" %
                  (p["rank"], p["code"], _fmt_value(p["world"]),
                   _shorten(p["label"], 50)))
    for key in ("png", "docx", "xlsx"):
        print("  saved: %s" % result[key])


if __name__ == "__main__":
    main()
