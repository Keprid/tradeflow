#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
generate_quarterly_report.py
============================

Generate the KEPROBA "Kenya Export Performance" quarterly Word report
(.docx) from the two pivot workbooks built by ``make_quarterly_tables.py``
(or the equivalent manually produced ones):

  - ``Exports.xlsx``  -- Table 1 Top Destination Markets,
                         Table 2 Top Export Products (+ pivots)
  - ``Imports.xlsx``  -- Annex 1 Imports by Partner,
                         Annex 2 Imports by Products (+ pivots)

The report carries the same features as the goods reports:

  * KEPROBA letterhead on every page;
  * title page, table of contents, list of tables, list of figures --
    none of them numbered;
  * page numbering restarts at 1 where the report body begins;
  * a data-driven narrative (overview, comparative perspectives and the
    Section 5 "Deduction") computed from the quarter's numbers.

Usage
-----
    python3 generate_quarterly_tables.py --excel-dir "EXPORT PERFORMANCE FOR Q2" \
        --out-dir output/quarterly
    python3 generate_quarterly_report.py --excel-dir output/quarterly \
        --output output/KENYA EXPORT PERFORMANCE IN APRIL-JUNE 2025-2026.docx

Dependencies: python-docx, openpyxl, matplotlib.
"""

import argparse
import os
import re
import sys

from pipeline_errors import PipelineError
from datetime import date

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import openpyxl

from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt, RGBColor

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE_DIR)

from charts import consolidate, draw_share_pie, slice_callouts, new_fig, finish
from country_names import display_name, short_product_name
from generate_report import (
    ReportBuilder, to_float, num, chart_font, THEME_ACCENTS, lighten,
)

SRC_KRA = "Source: KRA Database; Compiled by KEPROBA"

MONTH_NAMES = {1: "January", 2: "February", 3: "March", 4: "April",
               5: "May", 6: "June", 7: "July", 8: "August", 9: "September",
               10: "October", 11: "November", 12: "December"}


def _staple_name(name):
    """Concise, prose-friendly noun for a staple-food import label.

    The raw HS descriptions are verbose and punctuated (e.g. ``Rice.
    - Semi-milled or wholly milled rice, whether or not polished or
    glazed``), which reads poorly inside a narrative sentence. Map the
    known staple patterns to short common names instead.
    """
    low = name.lower()
    if "palm oil" in low:
        return "palm oil"
    if "wheat" in low or "meslin" in low:
        return "wheat and meslin"
    if "maize" in low or "corn" in low:
        return "maize (corn)"
    if "rice" in low:
        return "rice"
    return short_product_name(name, maxlen=20)

# Regional groupings used by the narrative -------------------------------
EAC = {"kenya", "uganda", "tanzania", "rwanda", "burundi",
       "democratic rep of congo", "democratic republic of congo",
       "south sudan", "somalia"}
AFRICA = EAC | {"egypt", "south africa", "malawi", "zambia", "zimbabwe",
                "nigeria", "ghana", "sudan", "ethiopia", "djibouti",
                "libya", "morocco", "tunisia", "algeria", "angola",
                "mozambique", "mauritius", "madagascar", "senegal",
                "cote d ivoire", "ivory coast", "cameroon", "gabon",
                "equatorial guinea", "congo", "namibia", "botswana",
                "lesotho", "eswatini", "swaziland", "seychelles",
                "sao tome & principe", "cape verde", "comoros",
                "mauritania", "mali", "niger", "chad", "benin", "togo",
                "sierra leone", "liberia", "guinea", "guinea bissau",
                "gambia", "burkina faso", "central african republic"}
ASIA = {"china", "india", "japan", "malaysia", "indonesia", "thailand",
        "vietnam", "singapore", "south korea", "republic of korea",
        "taiwan", "hong kong", "pakistan", "bangladesh", "sri lanka",
        "philippines", "myanmar(burma)", "myanmar", "cambodia", "laos",
        "nepal", "bhutan", "afghanistan", "uzbekistan", "kazakhstan"}
GULF = {"saudi arabia", "united arab emirates", "oman", "qatar", "kuwait",
        "bahrain", "yemen", "iraq", "iran"}
WEST = {"united states of america", "usa", "united kingdom", "japan",
        "belgium", "netherlands", "germany", "france", "italy", "canada",
        "spain", "sweden", "denmark", "finland", "austria", "ireland",
        "switzerland", "norway", "portugal", "greece", "australia",
        "new zealand"}

# Product keyword groups used in the deduction narrative ------------------
HORTICULTURE_KW = ("cut flower", "rose", "avocado", "mango", "vegetable",
                   "macadamia", "pineapple", "passion fruit", "legulous",
                   "beans,", "peas,", "courgette", "capsicum", "chilli")
PETROLEUM_KW = ("petroleum oils", "motor spirit", "gasoline",
                "butanes", "liquefied petroleum", "gas oil", "diesel",
                "jet fuel", "kerosene", "fuel")


# ---------------------------------------------------------------------------
# Workbook locating / parsing
# ---------------------------------------------------------------------------
def find_workbooks(excel_dir):
    """Locate Exports.xlsx / Imports.xlsx in ``excel_dir``."""
    found = {}
    for fname in sorted(os.listdir(excel_dir)):
        low = fname.lower()
        if not low.endswith(".xlsx"):
            continue
        path = os.path.join(excel_dir, fname)
        if low.startswith("export"):
            found.setdefault("exports", path)
        elif low.startswith("import"):
            found.setdefault("imports", path)
    missing = [k for k in ("exports", "imports") if k not in found]
    if missing:
        names = ", ".join(sorted(os.listdir(excel_dir)))
        raise PipelineError("[ERROR] Missing %s workbook(s) in '%s'.\nFiles present: %s"
                 % ("/".join(missing), excel_dir, names))
    return found


def _grid(ws):
    return [[c.value for c in row] for row in ws.iter_rows()]


def parse_rank_sheet(ws):
    """Parse one comparison rank-table sheet (current vs previous year).

    Expected generated layout::

        row1  (Value in Ksh. Billion)
        row2  Rank | ITEM | <Y-1> (merged)        | <Y> (merged)         | Change in <Y-1>-<Y>
        row3  ...  | April May June Total         | April May June Total | Total change | %
        data rows, All-other row, Grand Total

    Returns::

        {"sheet", "months": [...], "years": [yp, yc],
         "items": [{"name", "rank", "months": [[yp vals], [yc vals]],
                    "totals": [tp, tc], "change", "pct"}, ...],
         "all_other": {...} | None,
         "grand": {...}}
    """
    grid = _grid(ws)

    def cellv(row, ci):
        return row[ci] if ci < len(row) else None

    # -- locate the label row: first row holding >= 3 month names ----------
    month_set = set(MONTH_NAMES.values())
    lbl = None
    for ri, row in enumerate(grid):
        cnt = 0
        for v in row:
            if v is not None and str(v).strip().capitalize() in month_set:
                cnt += 1
        if cnt >= 3:
            lbl = ri
            break
    if lbl is None:
        return None

    # -- classify label-row columns into month runs ------------------------
    # runs[i] = [first_col, last_col, [month names]]
    runs = []
    for ci, v in enumerate(grid[lbl]):
        s = str(v).strip() if v is not None else ""
        if s.capitalize() in month_set:
            if runs and ci == runs[-1][1] + 1:
                runs[-1][1] = ci
                runs[-1][2].append(s)
            else:
                runs.append([ci, ci, [s]])
    if not runs:
        return None

    comparison = len(runs) >= 2
    if comparison:
        prev_run, cur_run = runs[0], runs[1]
    else:
        prev_run = cur_run = runs[0]      # single-year layout
    months = [m.capitalize() for m in cur_run[2]]

    # totals sit immediately after each run
    p_total = (prev_run[1] + 1) if comparison else None
    c_total = cur_run[1] + 1
    hdr_end = max(c_total, p_total or 0) + 3
    hdr_lbl = [str(cellv(grid[lbl], c)).strip()
               if cellv(grid[lbl], c) is not None else ""
               for c in range(hdr_end)]
    if "total" not in str(hdr_lbl[c_total]).lower():
        return None
    if comparison and "total" not in str(hdr_lbl[p_total]).lower():
        return None
    if not comparison:
        p_total = None

    chg_col = pct_col = share_col = None
    if comparison:
        # change / % columns come after the current-year Total
        for ci in range(c_total + 1, len(hdr_lbl)):
            h = hdr_lbl[ci].lower()
            if chg_col is None and "change" in h:
                chg_col = ci
            elif pct_col is None and h in ("%", "% change"):
                pct_col = ci
        if chg_col is None:
            chg_col = c_total + 1
            pct_col = c_total + 2
        elif pct_col is None:
            pct_col = chg_col + 1
    else:
        # single-year layouts carry a Share column after the Total
        for ci in range(c_total + 1, len(hdr_lbl)):
            h = hdr_lbl[ci].lower()
            if h.startswith("%") or "share" in h:
                share_col = ci
                break

    # -- years from the group header row ------------------------------------
    grp = lbl - 1
    grp_row = grid[grp] if grp >= 0 else []

    def year_at(ci):
        for c in range(min(ci, len(grp_row) - 1), -1, -1):
            v = grp_row[c]
            if isinstance(v, (int, float)) and 2015 <= int(v) <= 2100:
                return int(v)
        return None

    if comparison:
        y_prev = year_at(prev_run[0])
        y_cur = year_at(cur_run[0])
        if y_cur is None:
            y_cur = (y_prev or date.today().year - 1) + 1
        if y_prev is None:
            y_prev = y_cur - 1
        years = [y_prev, y_cur]
    else:
        years = [None, year_at(cur_run[0])]   # review year resolved later

    # -- data rows -----------------------------------------------------------
    def to_f(v):
        try:
            if v is None or (isinstance(v, str) and not v.strip()):
                return None
            return float(str(v).replace(",", ""))
        except (TypeError, ValueError):
            return None

    items, all_other = [], None
    grand = None
    n_m = len(months)
    for row in grid[lbl + 1:]:
        if not row or all(v is None or str(v).strip() == ""
                          for v in row[:c_total + 1]):
            continue
        label = cellv(row, 1)
        if label is None:
            continue
        label = str(label).strip()
        low = label.lower()
        if low.startswith("value in") or low in ("%", "% change"):
            continue
        cur_vals = [to_f(cellv(row, cur_run[0] + k)) for k in range(n_m)]
        prev_vals = ([to_f(cellv(row, prev_run[0] + k))
                      for k in range(n_m)] if comparison
                     else [None] * n_m)
        mvals = [prev_vals, cur_vals]
        t_prev = to_f(cellv(row, p_total)) if comparison else None
        t_cur = to_f(cellv(row, c_total))
        totals = [t_prev, t_cur]
        if t_cur is None:
            t_cur = sum(v for v in cur_vals if v is not None)
            totals[1] = t_cur
        if comparison and t_prev is None:
            totals[0] = sum(v for v in prev_vals if v is not None)
        share = to_f(cellv(row, share_col)) if share_col else None
        if comparison:
            change = to_f(cellv(row, chg_col))
            if change is None:
                change = (totals[1] or 0.0) - (totals[0] or 0.0)
            pct = to_f(cellv(row, pct_col))
            if pct is None and totals[0]:
                pct = change / totals[0]
        else:
            change, pct = None, None
        entry = {"name": label,
                 "rank": int(to_f(cellv(row, 0)))
                 if to_f(cellv(row, 0)) is not None else None,
                 "months": mvals, "totals": totals,
                 "share": share, "change": change, "pct": pct}
        if low.startswith("grand"):
            grand = entry
            break
        if entry["rank"] is None and low.startswith("all other"):
            all_other = entry
        else:
            items.append(entry)
    if grand is None:
        return None
    return {"sheet": ws.title, "months": months, "years": years,
            "comparison": comparison,
            "items": items, "all_other": all_other, "grand": grand}


def pick_sheet(wb, prefer, keywords):
    """Find the best-matching rank sheet by name, then by content.

    Raw-record sheets ('Data') are skipped: they hold the whole extract
    (tens of thousands of rows), can never be rank tables, and streaming
    them repeatedly is what made report builds slow / memory-hungry.
    """
    candidates = [n for n in wb.sheetnames if not n.lower().startswith("data")]
    for name in candidates:
        if prefer and prefer in name.lower():
            parsed = parse_rank_sheet(wb[name])
            if parsed:
                return parsed
    for name in candidates:
        low = name.lower()
        if any(k in low for k in keywords):
            parsed = parse_rank_sheet(wb[name])
            if parsed:
                return parsed
    for name in candidates:
        parsed = parse_rank_sheet(wb[name])
        if parsed:
            return parsed
    return None


def detect_year(wbs, fallback=None):
    """Review year: workbook title property, then Data-sheet Year column."""
    for wb in wbs:
        title = (wb.properties.title or "")
        m = re.search(r"(20\d{2})", title)
        if m:
            return int(m.group(1))
    for wb in wbs:
        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            for row in ws.iter_rows(min_row=1, max_row=5, max_col=6):
                for cell in row:
                    if isinstance(cell.value, str) \
                            and cell.value.strip().lower() == "year":
                        for r2 in ws.iter_rows(
                                min_row=cell.row + 1,
                                max_row=cell.row + 200,
                                min_col=cell.column, max_col=cell.column):
                            for c2 in r2:
                                if isinstance(c2.value, (int, float)) \
                                        and 2015 <= int(c2.value) <= 2100:
                                    return int(c2.value)
    return fallback


# ---------------------------------------------------------------------------
# Data loading and analysis
# ---------------------------------------------------------------------------
class QuarterAnalysis:
    """Loads the quarterly workbooks and derives every narrative figure.

    Tables compare the review quarter with the same quarter of the previous
    year; every table parsed carries ``months``/``totals`` lists ordered
    [previous year, current year].
    """

    def __init__(self, year=None):
        self.year = year          # review (current) year, detected if None
        self.month_names = []
        self.quarter = ""
        self.t1 = None   # top destination markets (exports by partner)
        self.t2 = None   # top export products
        self.a1 = None   # imports by supplying countries
        self.a2 = None   # imports by products
        self.cross = {"exports": {}, "imports": {}}   # (partner, product) -> [prev, cur]

    # -- loading ------------------------------------------------------------
    def load(self, excel_dir):
        files = find_workbooks(excel_dir)
        # read_only mode streams sheets instead of materialising every cell
        # of the big raw Data sheets in RAM
        wb_e = openpyxl.load_workbook(files["exports"], data_only=True,
                                      read_only=True)
        wb_i = openpyxl.load_workbook(files["imports"], data_only=True,
                                      read_only=True)
        try:
            self.t1 = pick_sheet(wb_e, "table 1", ("destination", "partner"))
            self.t2 = pick_sheet(wb_e, "table 2", ("product",))
            self.a1 = pick_sheet(wb_i, "annex 1", ("partner",))
            self.a2 = pick_sheet(wb_i, "annex 2", ("product",))
            for name, t in (("Table 1", self.t1), ("Table 2", self.t2),
                            ("Annex 1", self.a1), ("Annex 2", self.a2)):
                if t is None:
                    raise PipelineError("[ERROR] Could not locate %s in the quarterly "
                             "workbooks." % name)
            comparison = bool(self.t1.get("comparison")
                              and self.a1.get("comparison"))
            if self.year is None:
                years = [t["years"][1] for t in (self.t1, self.t2,
                                                 self.a1, self.a2)
                         if t["years"][1]]
                if len(years) == 4:
                    self.year = max(years)
                else:
                    # single-year sheets: fall back to workbook metadata / data
                    self.year = detect_year([wb_e, wb_i],
                                            fallback=date.today().year)
        finally:
            wb_e.close()
            wb_i.close()
        self.has_comparison = comparison
        self.year_prev = (self.year - 1) if comparison else None
        self.month_names = self.t1["months"]
        self.quarter = quarter_label(self.month_names)
        return self

    def load_crosstabs(self, excel_dir):
        """Aggregate the Data sheets into (partner, product) totals.

        Gives the narrative the partner-product cross view (which product
        drove a market's rise/fall) that the rank tables cannot provide.
        Missing Data sheets degrade silently to empty cross-tabs.
        """
        files = find_workbooks(excel_dir)
        for side in ("exports", "imports"):
            agg = {}
            wb = openpyxl.load_workbook(files[side], data_only=True,
                                        read_only=True)
            try:
                if "Data" not in wb.sheetnames:
                    self.cross[side] = agg
                    continue
                ws = wb["Data"]
                rows = ws.iter_rows(values_only=True)
                next(rows, None)          # header row
                for row in rows:
                    if row is None or len(row) < 5:
                        continue
                    yr, _mo, partner, product, value = row[:5]
                    if yr is None or partner is None or product is None:
                        continue
                    try:
                        yr = int(yr)
                    except (TypeError, ValueError):
                        continue
                    idx = 1 if yr == self.year else 0
                    key = (str(partner).strip().upper(),
                           str(product).strip())
                    cell = agg.setdefault(key, [0.0, 0.0])
                    try:
                        cell[idx] += float(value or 0.0)
                    except (TypeError, ValueError):
                        pass
            finally:
                wb.close()
            self.cross[side] = agg
        return self

    # -- partner x product driver analysis -----------------------------------
    def partner_drivers(self, side, partner, topn=3, min_abs=0.05):
        """Products behind a partner's change: [(product, prev, cur, chg)]."""
        out = []
        for (p, prod), (prev, cur) in self.cross[side].items():
            if p != str(partner).strip().upper():
                continue
            chg = cur - prev
            if abs(chg) < min_abs:
                continue
            out.append((prod, prev, cur, chg))
        out.sort(key=lambda t: abs(t[3]), reverse=True)
        return out[:topn]

    def product_drivers(self, side, product, topn=3, min_abs=0.05):
        """Partners behind a product's change: [(partner, prev, cur, chg)]."""
        key = str(product).strip().upper()
        out = []
        for (p, prod), (prev, cur) in self.cross[side].items():
            if prod.strip().upper() != key:
                continue
            chg = cur - prev
            if abs(chg) < min_abs:
                continue
            out.append((p, prev, cur, chg))
        out.sort(key=lambda t: abs(t[3]), reverse=True)
        return out[:topn]

    def top_movers(self, table, topn=3, min_abs=1.0):
        """Items with the largest absolute YoY change in ``table``."""
        rows = [d for d in table["items"]
                if d["change"] is not None and abs(d["change"]) >= min_abs]
        rows.sort(key=lambda d: abs(d["change"]), reverse=True)
        return rows[:topn]

    # -- basic aggregates ---------------------------------------------------
    @property
    def exp_total(self):
        return self.t1["grand"]["totals"][1] or 0.0

    @property
    def exp_total_prev(self):
        v = self.t1["grand"]["totals"][0]
        return v if v is not None else None

    @property
    def imp_total(self):
        return self.a1["grand"]["totals"][1] or 0.0

    @property
    def imp_total_prev(self):
        v = self.a1["grand"]["totals"][0]
        return v if v is not None else None

    @property
    def deficit(self):
        return self.exp_total - self.imp_total

    def exp_months(self):
        return list(self.t1["grand"]["months"][1])

    def exp_months_prev(self):
        return list(self.t1["grand"]["months"][0])

    def imp_months(self):
        return list(self.a1["grand"]["months"][1])

    def imp_months_prev(self):
        return list(self.a1["grand"]["months"][0])

    def balance_months(self):
        return [e - i for e, i in zip(self.exp_months(), self.imp_months())]

    def exp_growth_through_quarter(self):
        m = self.exp_months()
        if len(m) >= 2 and m[0]:
            return (m[-1] - m[0]) / abs(m[0])
        return None

    def yoy_growth(self):
        """Total export growth vs the same quarter of the previous year."""
        prev = self.exp_total_prev
        return (self.exp_total - prev) / prev if prev else None

    def yoy_import_growth(self):
        prev = self.imp_total_prev
        return (self.imp_total - prev) / prev if prev else None

    def import_multiple(self):
        if self.exp_total:
            return self.imp_total / self.exp_total
        return None

    # -- ranking helpers ----------------------------------------------------
    def shares(self, table, n=None, year="cur"):
        """(name, value, share) per ranked item for one of the two years."""
        idx = 1 if year in ("cur", "current") else 0
        total = table["grand"]["totals"][idx]
        rows = table["items"][:n] if n else table["items"]
        if total is None:                      # year not present in data
            return []
        out = []
        for d in rows:
            v = d["totals"][idx] or 0.0
            out.append((d["name"], v, (v / total) if total else 0.0))
        return out

    def cum_share(self, table, n, year="cur"):
        return sum(s for _, _, s in self.shares(table, n, year))

    def find_item(self, table, pattern, year="cur"):
        rx = re.compile(pattern, re.IGNORECASE)
        for name, v, s in self.shares(table, year=year):
            if rx.search(name):
                return (name, v, s)
        return None

    def group_share(self, table, group, year="cur"):
        """Combined share of items whose normalised name is inside group."""
        idx = 1 if year in ("cur", "current") else 0
        total = table["grand"]["totals"][idx] or 0.0
        acc = 0.0
        for d in table["items"]:
            key = re.sub(r"\s+", " ", str(d["name"]).strip().lower())
            if key in group:
                acc += d["totals"][idx] or 0.0
        return (acc / total) if total else 0.0

    def keyword_share(self, table, keywords, exclude=(), year="cur"):
        """Share of items matching any keyword (searched in the label)."""
        idx = 1 if year in ("cur", "current") else 0
        total = table["grand"]["totals"][idx] or 0.0
        acc = 0.0
        ex = [re.compile(p, re.IGNORECASE) for p in exclude]
        for d in table["items"]:
            name = str(d["name"])
            if any(rx.search(name) for rx in ex):
                continue
            low = name.lower()
            hit = any(kw in low for kw in keywords) \
                or any(re.search(kw, name, re.IGNORECASE)
                       for kw in keywords if kw.startswith("^"))
            if hit:
                acc += d["totals"][idx] or 0.0
        return (acc / total) if total else 0.0

    def growth_decline(self, table, threshold=1.0):
        """Items whose YoY change exceeds +/- ``threshold`` Ksh. Billion.

        Returns (growers, decliners); both sorted by absolute change,
        largest first.
        """
        if not self.has_comparison:
            return [], []
        growers = [d for d in table["items"]
                   if d["change"] is not None and d["change"] > threshold]
        decliners = [d for d in table["items"]
                     if d["change"] is not None and d["change"] < -threshold]
        growers.sort(key=lambda d: d["change"], reverse=True)
        decliners.sort(key=lambda d: d["change"])
        return growers, decliners


def quarter_label(month_names):
    """'April-June' from parsed month header names."""
    if len(month_names) == 1:
        return month_names[0]
    return "%s-%s" % (month_names[0], month_names[-1])


def fmt_b(v, dec=1):
    return ("Ksh %.*f billion" % (dec, v)) if v is not None else ""


def pct_s(share, dec=1):
    return ("%.*f%%" % (dec, share * 100)) if share is not None else ""


def short_prod(label, maxlen=52):
    txt = short_product_name(label, maxlen=maxlen)
    return txt.rstrip(".") if len(txt) <= maxlen else short_product_name(
        label, maxlen=maxlen - 1)


def _share_doughnut_data(a, table, other_label, year="cur"):
    """(label, value) pairs for an editable share doughnut, consolidated the
    same way as the reference PNG (make_chart_share)."""
    labels, values = [], []
    for name, v, _s in a.shares(table, year=year):
        if v:
            labels.append(short_product_name(name, maxlen=40))
            values.append(v)
    return consolidate(labels, values, max_slices=8, min_pct=2.0,
                       other_label=other_label)


# ---------------------------------------------------------------------------
# Table-as-image renderer (reference-report style)
# ---------------------------------------------------------------------------
# The reference report embeds the big rank tables as full-page images
# (pasted from Excel) instead of Word tables, so each table always fits a
# single page and keeps the Century Gothic 12 workbook formatting. The
# renderer below reproduces that: the table is drawn at its natural size
# (Century Gothic 12 everywhere) and inserted into Word scaled to the
# text width, exactly like the pasted EMF tables in the reference.
import textwrap

TBL_FONT = "Century Gothic"
TBL_FONT_SIZE = 12                      # pt, as in the workbook tables
TBL_HDR_FILL = "#FFFFFF"
TBL_HDR_TEXT = "#000000"
TBL_TOTAL_FILL = "#FFFFFF"
TBL_ROW_EVEN = "#FFFFFF"
TBL_ROW_ODD = "#FFFFFF"
TBL_GRID = "#B0B0B0"
TBL_GREEN = "#00B050"
TBL_RED = "#FF0000"
TBL_UNIT_NOTE = "(Value in Ksh. Billion)"

_CHAR_W = 0.088          # approx inches per character, Century Gothic 12
_ROW_H = 0.28            # inches per text line
_ROW_PAD = 0.10


def _wrap_label(text, width_chars):
    lines = []
    for chunk in str(text).split("\n"):
        lines.extend(textwrap.wrap(chunk, width_chars) or [""])
    return lines


def _col_w(headers, values, min_w=0.55, pad=0.28):
    """Width for a numeric column: widest of header / formatted values."""
    longest = max([len(str(h)) for h in headers] +
                  [len(str(v)) for v in values] + [4])
    return max(min_w, longest * _CHAR_W + pad)


def _draw_cell(ax, x, y_top, w, h, text, align="center", bold=False,
               color="#000000", fontsize=TBL_FONT_SIZE):
    if text is None or str(text) == "":
        return
    tx = x + w / 2 if align == "center" else x + 0.06
    ha = "center" if align == "center" else "left"
    ax.text(tx, y_top - h / 2, str(text), ha=ha, va="center",
            fontsize=fontsize, color=color, family=TBL_FONT,
            fontweight="bold" if bold else "normal")


def render_table_image(parsed, out_path, item_header="ITEM", year=None):
    """Render one rank table (comparison or single-year) to a PNG.

    Layout mirrors the workbook sheets: dark-blue header block, alternating
    row fills, light-blue All-other / Grand-total rows and green-up /
    red-down arrows in the change columns. Returns the PNG path.
    """
    comp = bool(parsed.get("comparison"))
    months = [m[:3] for m in parsed["months"]]
    n = len(months)
    years = parsed.get("years") or [None, None]
    Yc = years[1] or year or ""
    Yp = years[0] if comp else None

    def fmt(v):
        return num(v) if v is not None else ""

    def pct_fmt(v):
        return ("%.1f%%" % (v * 100)) if v is not None else ""

    data_rows = list(parsed["items"])
    if parsed.get("all_other"):
        data_rows.append(parsed["all_other"])
    grand = parsed["grand"]

    # ---- column widths -----------------------------------------------------
    if comp:
        chg_vals = [d["change"] for d in data_rows] + [grand["change"]]
        pct_vals = [d["pct"] for d in data_rows] + [grand["pct"]]
        w_rank = 0.45
        w_item = 3.0
        w_month = _col_w(months, [fmt(v) for d in data_rows
                                  for v in list(d["months"][0]) +
                                  list(d["months"][1])])
        w_total = _col_w(["Total"], [fmt(d["totals"][0])
                                     for d in data_rows] +
                         [fmt(grand["totals"][0])], min_w=0.75)
        w_chg = _col_w(["Total change"], [fmt(v) for v in chg_vals],
                       min_w=0.95) + 0.15        # room for the arrow
        w_pct = _col_w(["%", "% change"], [pct_fmt(v) for v in pct_vals],
                       min_w=0.7)
        widths = ([w_rank, w_item] + [w_month] * n + [w_total] +
                  [w_month] * n + [w_total, w_chg, w_pct])
        ncols = len(widths)
        c_chg, c_pct = ncols - 2, ncols - 1
        p_first, p_total = 2, 2 + n
        c_first, c_total = p_total + 1, p_total + 1 + n
    else:
        w_rank = 0.45
        w_item = 3.4
        w_month = _col_w(months, [fmt(v) for d in data_rows
                                  for v in d["months"][1]])
        w_total = _col_w(["Total"], [fmt(d["totals"][1])
                                     for d in data_rows] +
                         [fmt(grand["totals"][1])], min_w=0.8)
        w_share = 0.75
        widths = [w_rank, w_item] + [w_month] * n + [w_total, w_share]
        ncols = len(widths)
        c_share = ncols - 1
        c_total = 2 + n

    total_w = sum(widths)

    # ---- row heights (wrapped item labels) ---------------------------------
    item_chars = max(12, int(w_item / _CHAR_W))
    row_lines = []
    for d in data_rows:
        row_lines.append(max(1, len(_wrap_label(d["name"], item_chars))))
    grand_lines = max(1, len(_wrap_label(grand["name"], item_chars)))

    hdr_h0, hdr_h1 = 0.30, 0.30
    note_h = 0.22
    row_hs = [ln * _ROW_H + _ROW_PAD for ln in row_lines]
    grand_h = grand_lines * _ROW_H + _ROW_PAD
    fig_h = note_h + hdr_h0 + hdr_h1 + sum(row_hs) + grand_h + 0.1

    fig = plt.figure(figsize=(total_w, fig_h), dpi=200)
    ax = fig.add_axes([0, 0, 1, 1])
    ax.set_xlim(0, total_w)
    ax.set_ylim(0, fig_h)
    ax.axis("off")
    ax.set_facecolor("white")

    def rect(x, y, w, h, fill):
        from matplotlib.patches import Rectangle
        ax.add_patch(Rectangle((x, y), w, h, facecolor=fill,
                               edgecolor="none", zorder=0))

    # ---- column x positions -------------------------------------------------
    xs = [0.0]
    for w in widths:
        xs.append(xs[-1] + w)

    # unit note
    ax.text(0, fig_h - note_h / 2, TBL_UNIT_NOTE, ha="left", va="center",
            fontsize=9, style="italic", color="#555555", family=TBL_FONT)

    y = fig_h - note_h
    y_hdr0_bot = y - hdr_h0
    y_hdr1_bot = y_hdr0_bot - hdr_h1

    # header band fills
    rect(0, y_hdr1_bot, total_w, y - y_hdr1_bot, TBL_HDR_FILL)

    def hcenter(c0, c1, text, row_top, row_h, fontsize=TBL_FONT_SIZE):
        x0, x1 = xs[c0], xs[c1]
        _draw_cell(ax, x0, row_top, x1 - x0, row_h, text,
                   bold=True, color=TBL_HDR_TEXT, fontsize=fontsize)

    # row 0 (group headers) - merged cells drawn as one label per span
    # rank + item span both header rows
    _draw_cell(ax, xs[0], y, widths[0], hdr_h0 + hdr_h1, "Rank",
               bold=True, color=TBL_HDR_TEXT)
    _draw_cell(ax, xs[1], y, widths[1], hdr_h0 + hdr_h1, item_header,
               bold=True, color=TBL_HDR_TEXT)
    if comp:
        hcenter(p_first, p_total + 1, str(Yp), y, hdr_h0)
        hcenter(c_first, c_total + 1, str(Yc), y, hdr_h0)
        hcenter(c_chg, c_pct + 1, "Change in\n%d-%d" % (Yp, Yc),
                y, hdr_h0, fontsize=TBL_FONT_SIZE - 1)
        for base, tcol in ((p_first, p_total), (c_first, c_total)):
            for k, m in enumerate(months):
                hcenter(base + k, base + k + 1, m, y_hdr0_bot, hdr_h1)
            hcenter(tcol, tcol + 1, "Total", y_hdr0_bot, hdr_h1)
        hcenter(c_chg, c_chg + 1, "Total change", y_hdr0_bot, hdr_h1,
                fontsize=TBL_FONT_SIZE - 1)
        hcenter(c_pct, c_pct + 1, "%", y_hdr0_bot, hdr_h1)
    else:
        for k, m in enumerate(months):
            hcenter(2 + k, 3 + k, m, y, hdr_h0)
        hcenter(c_total, c_total + 1, "Total", y, hdr_h0)
        # share header spans both header rows (year needs the space)
        _draw_cell(ax, xs[c_share], y, widths[c_share],
                   hdr_h0 + hdr_h1, "Share\nin %s" % Yc,
                   bold=True, color=TBL_HDR_TEXT,
                   fontsize=TBL_FONT_SIZE - 1)
        # unit label spanning the month+total columns
        hcenter(2, c_total + 1, TBL_UNIT_NOTE, y_hdr0_bot, hdr_h1,
                fontsize=10)

    y = y_hdr1_bot
    # ---- data rows -----------------------------------------------------------
    def arrow_cell(ri, c, value, kind, h):
        if value is None:
            txt = "\u2013"
            color = "#808080"
        else:
            up = value >= 0
            txt = ("\u25b2 " if up else "\u25bc ") + \
                (("%.1f%%" % (value * 100)) if kind == "pct" else fmt(value))
            color = TBL_GREEN if up else TBL_RED
        _draw_cell(ax, xs[c], y, widths[c], h, txt,
                   bold=True, color=color, fontsize=TBL_FONT_SIZE - 1)

    def emit(d, ri, is_total=False, is_grand=False):
        nonlocal y
        h = grand_h if is_grand else row_hs[ri]
        fill = TBL_TOTAL_FILL if (is_total or is_grand) else \
            (TBL_ROW_EVEN if ri % 2 == 0 else TBL_ROW_ODD)
        rect(0, y - h, total_w, h, fill)
        bold = is_total or is_grand
        name_lines = _wrap_label(d["name"], item_chars)
        _draw_cell(ax, xs[0], y, widths[0], h,
                   str(d["rank"]) if d.get("rank") is not None else "",
                   bold=bold)
        _draw_cell(ax, xs[1], y, widths[1], h,
                   "\n".join(name_lines), align="left", bold=bold)
        if comp:
            for base, tcol, mvals, tot in (
                    (p_first, p_total, d["months"][0], d["totals"][0]),
                    (c_first, c_total, d["months"][1], d["totals"][1])):
                for k, v in enumerate(mvals):
                    _draw_cell(ax, xs[base + k], y, widths[base + k], h,
                               fmt(v), bold=bold)
                _draw_cell(ax, xs[tcol], y, widths[tcol], h, fmt(tot),
                           bold=bold)
            arrow_cell(ri, c_chg, d.get("change"), "chg", h)
            arrow_cell(ri, c_pct, d.get("pct"), "pct", h)
        else:
            for k, v in enumerate(d["months"][1]):
                _draw_cell(ax, xs[2 + k], y, widths[2 + k], h, fmt(v),
                           bold=bold)
            _draw_cell(ax, xs[c_total], y, widths[c_total], h,
                       fmt(d["totals"][1]), bold=bold)
            sh = d.get("share")
            _draw_cell(ax, xs[c_share], y, widths[c_share], h,
                       ("%.1f%%" % (sh * 100)) if sh is not None else "",
                       bold=bold)
        y -= h

    ri = 0
    for d in data_rows:
        is_total = parsed.get("all_other") is not None and d is data_rows[-1] \
            and d.get("rank") is None
        emit(d, ri, is_total=is_total)
        ri += 1
    emit(grand, ri, is_grand=True)

    # grid lines
    for yb in [fig_h - note_h, y_hdr1_bot] + \
            [y_hdr1_bot - sum(row_hs[:i + 1]) for i in range(len(row_hs))] + \
            [0.05]:
        ax.plot([0, total_w], [yb, yb], color=TBL_GRID, linewidth=0.6,
                zorder=1)
    for xv in xs:
        ax.plot([xv, xv], [0.05, fig_h - note_h], color=TBL_GRID,
                linewidth=0.6, zorder=1)

    fig.savefig(out_path, facecolor="white")
    plt.close(fig)
    return out_path


# ---------------------------------------------------------------------------
# Narrative generation (every figure derived from the analysed data)
# ---------------------------------------------------------------------------
_KEEP_CAPS = {"USA", "UAE", "UK", "US", "EU", "DRC", "USSR", "GCC"}


def _mkt(name):
    """Narration-friendly market name ('JAPAN' -> 'Japan', 'USA' kept)."""
    disp = display_name(name)
    if disp and disp.isupper():
        words = [w if w.strip(".,") in _KEEP_CAPS else
                 "-".join(p.capitalize() for p in w.split("-"))
                 for w in disp.split()]
        disp = " ".join(words)
    return disp


def _chg_phrase(prev, cur, y_prev, Y):
    """'an increase of Ksh 2.2 billion (5.0%), from Ksh 44.0 billion in
    2025 to Ksh 46.3 billion in 2026' -- reference-report wording."""
    chg = cur - prev
    word = "an increase" if chg >= 0 else "a decline"
    pct = (" (%+.1f%%)" % (abs(chg) / prev * 100)) if prev else ""
    return ("%s of %s%s, from %s in %d to %s in %d"
            % (word, fmt_b(abs(chg)), pct, fmt_b(prev), y_prev,
               fmt_b(cur), Y))


def _driver_clause(drivers, y_prev, Y, maxn=2):
    """', from Ksh A in <yp> to Ksh B in <Y>' clauses for the top drivers."""
    parts = []
    for prod, prev, cur, chg in drivers[:maxn]:
        parts.append("%s by %s, from %s in %d to %s in %d"
                     % (short_prod(prod, 46).lower(), fmt_b(abs(chg)),
                        fmt_b(prev), y_prev, fmt_b(cur), Y))
    return ", and ".join(parts)


def _up_list(items):
    """'Name (up 54.7%, from Ksh A to Ksh B)' list for the overview."""
    return ", ".join(
        "%s (up %.1f%%, from %s to %s)"
        % (short_prod(d["name"], 34), d["pct"] * 100,
           fmt_b(d["totals"][0]), fmt_b(d["totals"][1]))
        for d in items)


def _interventions():
    """The standing raft of proposed interventions (reference report)."""
    return [
        "**Market Intelligence** \u2013 Set up a trade intelligence unit "
        "and real-time dashboard; commission diagnostic studies on the "
        "declines in strategic markets; strengthen embassy trade "
        "attach\u00e9s.",
        "**Diplomatic Engagement** \u2013 Reactivate trade dialogues with "
        "markets registering declines; align with the EU on compliance "
        "requirements; use AfCFTA and trade missions to reinforce ties "
        "with underperforming markets.",
        "**Diversification** \u2013 Expand into new markets, building on "
        "the gains recorded by the fastest-growing destinations, and "
        "promote higher-value product lines in traditional sectors.",
        "**Trade Facilitation** \u2013 Streamline customs and logistics; "
        "expand export credit/insurance and trade finance to manage "
        "currency risk.",
    ]


def build_narratives(a: QuarterAnalysis):
    """Build every narrative block from the analysed quarter.

    All figures are derived from the parsed tables and the partner-product
    cross-tabs; the phrasing follows the reference (Q2) export performance
    report. Single-year extracts (no prior-year quarter in the data) fall
    back to share-based narratives.
    """
    Y = a.year
    q = a.quarter
    m_first, m_last = a.month_names[0], a.month_names[-1]
    period = "%s %d" % (q, Y)
    comp = a.has_comparison
    Yp = a.year_prev
    t1, t2 = a.t1, a.t2
    a1, a2 = a.a1, a.a2

    exp_t, imp_t = a.exp_total, a.imp_total
    def_t = a.deficit
    yoy_e = a.yoy_growth() if comp else None
    yoy_i = a.yoy_import_growth() if comp else None

    mkts = a.shares(t1)
    prods = [x for x in a.shares(t2)
             if not str(x[0]).lower().startswith("all others")]
    imp_mkts = a.shares(a1)
    imp_prods = [x for x in a.shares(a2)
                 if not str(x[0]).lower().startswith("all others")]

    n = {}

    # ---- Section 1: overview ---------------------------------------------
    p1 = ("Kenya's trade sector recorded total merchandise exports of %s "
          "in %s (between %s and %s), according to data released by the "
          "Kenya Revenue Authority (KRA). "
          % (fmt_b(exp_t), period, m_first, m_last))
    if comp and yoy_e is not None:
        p1 += ("This was %s of %s (%+.1f%%) over the same period in %d. "
               % ("an increase" if yoy_e >= 0 else "a decline",
                  fmt_b(abs(exp_t - (a.exp_total_prev or 0))),
                  yoy_e * 100, Yp))
    top_m = [x for x in mkts[:3]]
    if top_m:
        p1 += ("Export growth during this period was driven by stronger "
               "demand from destination countries including %s. "
               % ", ".join(_mkt(x[0]) for x in top_m))
    if comp:
        growers_p = sorted(
            [d for d in t2["items"]
             if d["change"] is not None and d["change"] > 0.3
             and d["pct"] is not None],
            key=lambda d: d["change"], reverse=True)[:5]
        if growers_p:
            p1 += ("Significant contributions also came from key sectors "
                   "such as %s, all of which recorded export growth in "
                   "%s compared to %s. "
                   % (_up_list(growers_p), Y, Yp))
        if yoy_i is not None and yoy_e is not None and yoy_i >= yoy_e:
            p1 += ("However, the pace of import growth outpaced that of "
                   "exports, resulting in a widening trade deficit \u2013 a "
                   "key indicator of the country's persistent reliance on "
                   "foreign goods. ")
    p1 += ("The leading buyers of Kenyan goods were %s (%s), %s (%s) and "
           "%s (%s). "
           % (_mkt(mkts[0][0]), pct_s(mkts[0][2]),
              _mkt(mkts[1][0]), pct_s(mkts[1][2]),
              _mkt(mkts[2][0]), pct_s(mkts[2][2])))
    if prods:
        p1 += "%s was the top export at %s (%s of total exports). " \
              % (short_prod(prods[0][0], 30), fmt_b(prods[0][1]),
                 pct_s(prods[0][2]))
    n["overview"] = [p1]

    p2 = ("Total imports were valued at %s in the period %s to %s %d. "
          % (fmt_b(imp_t), m_first, m_last, Y))
    if comp and yoy_i is not None:
        p2 += ("This was %s of %s (%+.1f%%) over the same quarter of %d. "
               % ("an increase" if yoy_i >= 0 else "a decline",
                  fmt_b(abs(imp_t - (a.imp_total_prev or 0))),
                  yoy_i * 100, Yp))
    if imp_mkts:
        p2 += ("The leading sources of imports were %s (%s), %s (%s) and "
               "%s (%s). "
               % (_mkt(imp_mkts[0][0]), pct_s(imp_mkts[0][2]),
                  _mkt(imp_mkts[1][0]), pct_s(imp_mkts[1][2]),
                  _mkt(imp_mkts[2][0]), pct_s(imp_mkts[2][2])))
    if imp_prods:
        p2 += "%s remained the largest import at %s (%s of the bill). " \
              % (short_prod(imp_prods[0][0], 30), fmt_b(imp_prods[0][1]),
                 pct_s(imp_prods[0][2]))
    n["overview"].append(p2)

    if comp:
        p3 = ("This report provides a comprehensive comparative analysis "
              "of Kenya's trade performance between the %s quarters of %d "
              "and %d. It examines critical trends in export destinations, "
              "and the performance of major export commodities."
              % (q, Yp, Y))
    else:
        p3 = ("This report provides a comprehensive analysis of Kenya's "
              "trade performance for %s. It examines critical trends in "
              "export destinations, and the performance of major export "
              "commodities." % period)
    n["overview"].append(p3)

    # ---- Section 2: comparative perspectives (plain paragraphs) -----------
    if comp:
        s2 = []
        chg_e = exp_t - (a.exp_total_prev or 0)
        s2.append("In %s %d, total exports were valued at %s compared to "
                  "%s in the same period in %d; and therefore, total "
                  "exports %s by %s (%+.1f%%)."
                  % (q, Y, fmt_b(exp_t), fmt_b(a.exp_total_prev), Yp,
                     "increased" if chg_e >= 0 else "decreased",
                     fmt_b(abs(chg_e)), yoy_e * 100))
        chg_i = imp_t - (a.imp_total_prev or 0)
        s2.append("On the other hand, total imports were valued at %s in "
                  "the period %s %d, which %s by %s (a %s of %+.1f%%) from "
                  "%s in the same period in %d."
                  % (fmt_b(imp_t), q, Y,
                     "increased" if chg_i >= 0 else "decreased",
                     fmt_b(abs(chg_i)),
                     "increase" if chg_i >= 0 else "decline",
                     yoy_i, fmt_b(a.imp_total_prev), Yp))
        if def_t < 0:
            s2.append("The balance of trade deficit for merchandised trade "
                      "for the period %s %d %s by %s (%+.1f%%) to record "
                      "%s in %d from %s in %d."
                      % (q, Y,
                         "widened" if abs(def_t) > abs(a.exp_total_prev -
                                                       a.imp_total_prev or 0)
                         else "narrowed",
                         fmt_b(abs(def_t) -
                               abs((a.exp_total_prev or 0) -
                                   (a.imp_total_prev or 0))),
                         (abs(def_t) / abs((a.exp_total_prev or 0) -
                                           (a.imp_total_prev or 0)) - 1)
                         * 100 if a.exp_total_prev and a.imp_total_prev
                         else 0.0,
                         fmt_b(abs(def_t)), Y,
                         fmt_b(abs((a.exp_total_prev or 0) -
                                   (a.imp_total_prev or 0))), Yp))
        else:
            s2.append("The balance of trade for merchandised trade for the "
                      "period %s %d recorded a surplus of %s."
                      % (q, Y, fmt_b(def_t)))
        n["s2"] = s2
    else:
        n["s2"] = [
            "In the period %s %d (between %s and %s %d), total exports "
            "were valued at %s."
            % (q, Y, m_first, m_last, Y, fmt_b(exp_t)),
            "On the other hand, total imports were valued at %s in the "
            "same period." % fmt_b(imp_t),
            "The balance of trade for merchandised trade for the period "
            "%s %d was %s."
            % (q, Y,
               ("a deficit of %s" % fmt_b(abs(def_t))) if def_t < 0
               else ("a surplus of %s" % fmt_b(def_t))),
        ]

    # ---- Section 3: markets narrative + driver bullets ---------------------
    if comp:
        n["markets_note"] = (
            "The two last columns in the above table represent change in "
            "value and % respectively between the two years. In %d "
            "compared to the same period %s %d, Kenyan export destination "
            "markets that showed significant change were as follows."
            % (Y, q, Yp))
        bullets = []
        decliners = [d for d in a.top_movers(t1, topn=4, min_abs=1.0)
                     if d["change"] < 0]
        growers = [d for d in a.top_movers(t1, topn=4, min_abs=1.0)
                   if d["change"] > 0]
        for d in decliners[:3]:
            drv = a.partner_drivers("exports", d["name"], topn=2)
            if drv:
                bullets.append(
                    "Decline in exports to %s for the reviewed period was "
                    "as a result of decline in export of %s."
                    % (_mkt(d["name"]),
                       _driver_clause(drv, Yp, Y)))
            else:
                bullets.append(
                    "Exports to %s declined by %s (%+.1f%%), from %s in "
                    "%d to %s in %d."
                    % (_mkt(d["name"]), fmt_b(abs(d["change"])),
                       (d["pct"] or 0) * 100, fmt_b(d["totals"][0]), Yp,
                       fmt_b(d["totals"][1]), Y))
        for d in growers[:2]:
            drv = a.partner_drivers("exports", d["name"], topn=2)
            if drv:
                bullets.append(
                    "Export growth to %s was driven by increase in export "
                    "of %s."
                    % (_mkt(d["name"]), _driver_clause(drv, Yp, Y)))
        n["market_bullets"] = bullets
    else:
        n["markets_note"] = None
        n["market_bullets"] = [
            "%s was the leading destination market, absorbing %s of "
            "exports (%s of the total)."
            % (_mkt(mkts[0][0]), fmt_b(mkts[0][1]), pct_s(mkts[0][2])),
            "The top ten destination markets accounted for %s of total "
            "exports, led by %s."
            % (pct_s(a.cum_share(t1, 10)),
               ", ".join(_mkt(x[0]) for x in mkts[1:6])),
        ]

    # ---- Section 4: products narrative --------------------------------------
    if comp:
        n["products_note"] = (
            "Compared to the period %s %d and %d, below is a table with "
            "top export products whose exports experienced significant "
            "change." % (q, Yp, Y))
        paras = []
        decliners_p = [d for d in a.top_movers(t2, topn=3, min_abs=0.5)
                       if d["change"] < 0]
        for d in decliners_p[:2]:
            drv = a.product_drivers("exports", d["name"], topn=3)
            if drv:
                paras.append(
                    "The decline in %s by %s was mainly driven by reduced "
                    "exports to %s."
                    % (short_prod(d["name"], 46), fmt_b(abs(d["change"])),
                       ", ".join(
                           "%s (down %s, from %s to %s)"
                           % (_mkt(x[0]), fmt_b(abs(x[3])), fmt_b(x[1]),
                              fmt_b(x[2]))
                           for x in drv[:3])))
        growers_p = sorted(
            [d for d in t2["items"]
             if d["change"] is not None and d["change"] > 0.5
             and d["pct"] is not None],
            key=lambda d: d["change"], reverse=True)[:4]
        if growers_p:
            paras.append("On the growth side, %s registered the largest "
                         "gains in the review period."
                         % _up_list(growers_p))
        n["product_paras"] = paras
    else:
        n["products_note"] = None
        n["product_paras"] = [
            "%s, %s and %s were the leading export products, together "
            "accounting for %s of total exports."
            % (short_prod(prods[0][0], 30), short_prod(prods[1][0], 30),
               short_prod(prods[2][0], 30),
               pct_s(sum(x[2] for x in prods[:3]), 0))
            if len(prods) >= 3 else
            "The leading export products accounted for the bulk of "
            "earnings in the review period.",
            "The top ten export products accounted for %s of total "
            "exports." % pct_s(a.cum_share(t2, 10)),
        ]

    # ---- Section 5: deduction ----------------------------------------------
    if comp:
        g_mkt, d_mkt = a.growth_decline(t1, threshold=1.0)
        g_prd, d_prd = a.growth_decline(t2, threshold=1.0)
        d1 = ("The %s quarter of %d marked %s for Kenya's merchandise "
              "trade. Total exports %s by %s (%+.1f%%), from %s in %d to "
              "%s in %d, signalling %s demand for Kenyan products in key "
              "international markets."
              % (q, Y,
                 "a period of robust expansion" if yoy_e and yoy_e > 0
                 else "a challenging period",
                 "rose" if yoy_e >= 0 else "fell",
                 fmt_b(abs(exp_t - (a.exp_total_prev or 0))),
                 yoy_e * 100, fmt_b(a.exp_total_prev), Yp, fmt_b(exp_t), Y,
                 "improved" if yoy_e >= 0 else "weakened"))
        n["deduction"] = [d1]

        d2 = ("This growth was broad, with %d export destination(s) "
              "registering increases exceeding Kshs 1 billion in the top "
              "25 category, led by %s (%s)"
              % (len(g_mkt),
                 _mkt(g_mkt[0]["name"]) if g_mkt else mkts[0][0],
                 fmt_b(g_mkt[0]["change"]) if g_mkt else fmt_b(0)))
        if len(g_mkt) > 1:
            d2 += (" and regional neighbours %s"
                   % ", ".join("%s (%s)" % (_mkt(x["name"]), fmt_b(x["change"]))
                               for x in g_mkt[1:4]))
        d2 += "."
        if d_mkt:
            d2 += (" However, this positive export performance was "
                   "partially offset by declines in shipments to %s, "
                   "reflecting shifting global demand patterns and "
                   "competitive pressures in traditional markets."
                   % ", ".join("%s (%s)" % (_mkt(x["name"]),
                                           fmt_b(x["change"]))
                               for x in d_mkt[:3]))
        n["deduction"].append(d2)

        d3 = ""
        if g_prd:
            lead = g_prd[0]
            d3 = ("On the product performance front, Kenya's export basket "
                  "demonstrated significant dynamism. The standout "
                  "performer was %s, which recorded an increase of %s "
                  "(%+.1f%%), from %s in %d to %s in %d. "
                  % (short_prod(lead["name"], 40), fmt_b(lead["change"]),
                     (lead["pct"] or 0) * 100, fmt_b(lead["totals"][0]),
                     Yp, fmt_b(lead["totals"][1]), Y))
            others = [x for x in g_prd[1:5]]
            if others:
                d3 += ("Traditional mainstays such as %s continued to "
                       "register steady growth. "
                       % ", ".join("%s (%s)" % (short_prod(x["name"], 26),
                                                fmt_b(x["change"]))
                                   for x in others))
        if d_prd:
            d3 += ("The decline in %s was the notable contraction, "
                   "suggesting that Kenya's export growth was largely "
                   "resilient in the review period."
                   % ", ".join("%s (%s)" % (short_prod(x["name"], 30),
                                            fmt_b(x["change"]))
                               for x in d_prd[:2]))
        if d3:
            n["deduction"].append(d3)

        chg_i = imp_t - (a.imp_total_prev or 0)
        d4 = ("The combined effect of strong import growth outpacing "
              "export expansion resulted in a significantly widened trade "
              "deficit. The merchandise trade deficit %s by %s (%+.1f%%), "
              "from %s in %d to %s in %d. This widening imbalance "
              "underscores Kenya's continued reliance on imported capital "
              "goods, machinery, vehicles and energy products to support "
              "its economic activities, even as its export sector shows "
              "encouraging growth and diversification."
              % ("ballooned" if chg_i > 0 else "narrowed",
                 fmt_b(abs(abs(def_t) -
                           abs((a.exp_total_prev or 0) -
                               (a.imp_total_prev or 0)))),
                 (abs(def_t) / abs((a.exp_total_prev or 0) -
                                   (a.imp_total_prev or 0)) - 1) * 100
                 if a.exp_total_prev and a.imp_total_prev else 0.0,
                 fmt_b(abs((a.exp_total_prev or 0) -
                           (a.imp_total_prev or 0))), Yp,
                 fmt_b(abs(def_t)), Y))
        n["deduction"].append(d4)
    else:
        n["deduction"] = [
            "Kenya's merchandise trade in the %s quarter of %d recorded "
            "total exports of %s against an import bill of %s, yielding a "
            "%s of %s. Trade remains concentrated in a narrow set of "
            "markets and products: the top ten destinations absorbed %s "
            "of exports and the top ten products earned %s of the total."
            % (q, Y, fmt_b(exp_t), fmt_b(imp_t),
               "deficit" if def_t < 0 else "surplus",
               fmt_b(abs(def_t)), pct_s(a.cum_share(t1, 10), 0),
               pct_s(a.cum_share(t2, 10), 0)),
            "The leading destination markets were %s, with %s alone "
            "taking %s of total exports. On the product front, %s led "
            "export earnings at %s (%s), followed by %s. This "
            "concentration in both products and markets exposes earnings "
            "to commodity price volatility, weather shocks and external "
            "demand fluctuations."
            % (", ".join(_mkt(x[0]) for x in mkts[:5]), _mkt(mkts[0][0]),
               pct_s(mkts[0][2]),
               short_prod(prods[0][0], 30) if prods else "-",
               fmt_b(prods[0][1]) if prods else "",
               pct_s(prods[0][2]) if prods else "",
               ", ".join(short_prod(x[0], 24) for x in prods[1:4])),
            "Imports were valued at %s, with the leading sources being %s. "
            "%s was the largest import item at %s (%s of the bill). The "
            "import bill of %s against exports of %s %s the need to "
            "broaden exports and address competitiveness and supply-side "
            "constraints in key sectors."
            % (fmt_b(imp_t),
               ", ".join(_mkt(x[0]) for x in imp_mkts[:5]),
               short_prod(imp_prods[0][0], 30) if imp_prods else "-",
               fmt_b(imp_prods[0][1]) if imp_prods else "",
               pct_s(imp_prods[0][2]) if imp_prods else "",
               fmt_b(imp_t), fmt_b(exp_t),
               "underscores" if def_t < 0 else "despite the surplus, "
               "highlights"),
        ]

    # ---- Implication for the country ---------------------------------------
    if comp:
        imp_p1 = ("The %s quarter of %d demonstrated that Kenya's export "
                  "sector remains resilient. Growth in exports to %s "
                  "suggests that there exist diverse opportunities in "
                  "these markets that require targeted market-development "
                  "interventions. %s remain among Kenya's strongest export "
                  "performers, confirming the need for ongoing "
                  "productivity improvements, compliance with "
                  "international standards, climate-resilience "
                  "interventions and support for value addition."
                  % (q, Y,
                     ", ".join(_mkt(x["name"]) for x in g_mkt[:5])
                     if comp and g_mkt else
                     ", ".join(_mkt(x[0]) for x in mkts[:5]),
                     ", ".join(short_prod(x["name"], 22)
                               for x in g_prd[:4])
                     if comp and g_prd else
                     ", ".join(short_prod(x[0], 22) for x in prods[:4])))
        imp_p2 = ("The trade balance moved from %s in %d to %s in %d, "
                  "with total imports %s by %+.1f%%. This raises concerns "
                  "on long-term trade balance sustainability and "
                  "reinforces the need for policies that promote domestic "
                  "production, export competitiveness and value addition. "
                  "The country should prioritise support for emerging "
                  "export sectors while sustaining competitiveness in "
                  "traditional sectors such as tea, coffee and "
                  "horticulture."
                  % (fmt_b(abs((a.exp_total_prev or 0) -
                               (a.imp_total_prev or 0))), Yp,
                     fmt_b(abs(def_t)), Y,
                     "growing" if yoy_i >= 0 else "declining",
                     yoy_i * 100))
        imp_p3 = ("The concentration of export growth in a limited number "
                  "of markets highlights the need for continued efforts "
                  "to diversify Kenya's export destinations and products. "
                  "Overreliance on a few destination markets exposes "
                  "exporters to external shocks, policy changes and "
                  "market disruptions. The reported declines in exports "
                  "to strategic markets suggest that greater diplomatic "
                  "and trade engagement may be required to address "
                  "market-specific challenges affecting Kenyan exports.")
        n["implication"] = [imp_p1, imp_p2, imp_p3]
    else:
        n["implication"] = [
            "The %s quarter of %d demonstrated that Kenya's export sector "
            "remains anchored on traditional products and markets. %s and "
            "the wider East African region remain among Kenya's strongest "
            "export performers, confirming the need for ongoing "
            "productivity improvements, compliance with international "
            "standards, climate-resilience interventions and support for "
            "value addition."
            % (q, Y, _mkt(mkts[0][0]) if mkts else "Kenya's traditional "
               "markets"),
            "The trade balance recorded a %s of %s in the review period, "
            "with total imports of %s against exports of %s. This "
            "reinforces the need for policies that promote domestic "
            "production, export competitiveness and value addition. The "
            "country should prioritise support for emerging export "
            "sectors while sustaining competitiveness in traditional "
            "sectors such as tea, coffee and horticulture."
            % ("deficit" if def_t < 0 else "surplus", fmt_b(abs(def_t)),
               fmt_b(imp_t), fmt_b(exp_t)),
            "The concentration of export growth in a limited number of "
            "markets highlights the need for continued efforts to "
            "diversify Kenya's export destinations and products. "
            "Overreliance on a few destination markets exposes exporters "
            "to external shocks, policy changes and market disruptions.",
        ]
    n["interventions"] = _interventions()

    # ---- Annex 1: imports by supplying countries ---------------------------
    annex1 = []
    if comp:
        movers = a.top_movers(a1, topn=6, min_abs=1.0)
        lead_name = (a1["items"][0]["name"].strip().upper()
                     if a1["items"] else None)
        for d in movers:
            drv_txt = _driver_clause(
                a.partner_drivers("imports", d["name"], topn=3),
                Yp, Y, maxn=3)
            if d["name"].strip().upper() == lead_name:
                annex1.append(
                    "In the reviewed period, %s remains the leading source "
                    "of imports, with imports %s (%+.1f%%, from %s to %s) "
                    "due to %s in import of %s."
                    % (_mkt(d["name"]),
                       _chg_phrase(d["totals"][0], d["totals"][1], Yp, Y),
                       (d["pct"] or 0) * 100, fmt_b(d["totals"][0]),
                       fmt_b(d["totals"][1]),
                       "an increase" if d["change"] >= 0 else "a decline",
                       drv_txt))
            else:
                annex1.append(
                    "In the reviewed period, %s's imports %s (%+.1f%%, "
                    "from %s to %s), driven by %s in import of %s."
                    % (_mkt(d["name"]),
                       _chg_phrase(d["totals"][0], d["totals"][1], Yp, Y),
                       (d["pct"] or 0) * 100, fmt_b(d["totals"][0]),
                       fmt_b(d["totals"][1]),
                       "an increase" if d["change"] >= 0 else "a decline",
                       drv_txt))
        annex1.append(
            "Total imports to Kenya %s by %+.1f%% (from %s to %s) over "
            "the review period."
            % ("grew" if yoy_i >= 0 else "contracted", yoy_i * 100,
               fmt_b(a.imp_total_prev), fmt_b(imp_t)))
    else:
        annex1.append(
            "Kenya imported goods worth %s in the %s quarter (%s to %s "
            "%d). Import sourcing is concentrated in a small group of "
            "partners: the top three sources \u2013 %s \u2013 supplied %s "
            "(%s of the total import bill)."
            % (fmt_b(imp_t), q, m_first, m_last, Y,
               ", ".join(_mkt(x[0]) for x in imp_mkts[:3]),
               fmt_b(sum(x[1] for x in imp_mkts[:3])),
               pct_s(sum(x[2] for x in imp_mkts[:3]))))
        annex1.append(
            "The top five suppliers accounted for %s of imports and the "
            "top ten for %s. This concentration shows how dependent Kenya "
            "is on a narrow group of supply markets, particularly in Asia "
            "and the Gulf."
            % (pct_s(a.cum_share(a1, 5)), pct_s(a.cum_share(a1, 10))))
        annex1.append(
            "The leading suppliers were %s. %s alone accounted for %s of "
            "the import bill."
            % (", ".join("%s (%s)" % (_mkt(x[0]), fmt_b(x[1]))
                         for x in imp_mkts[:6]),
               _mkt(imp_mkts[0][0]), pct_s(imp_mkts[0][2])))
    n["annex1"] = annex1

    # ---- Annex 2: imports by products ---------------------------------------
    annex2 = []
    if comp:
        movers_p = a.top_movers(a2, topn=6, min_abs=0.5)
        if movers_p:
            annex2.append(
                "During the reviewed period, %s."
                % ", ".join(
                    "%s %s by %+.1f%% (from %s to %s)"
                    % (short_prod(d["name"], 36),
                       "grew" if d["change"] >= 0 else "declined",
                       (d["pct"] or 0) * 100, fmt_b(d["totals"][0]),
                       fmt_b(d["totals"][1]))
                    for d in movers_p[:5]))
        annex2.append(
            "Total imports to Kenya %s by %+.1f%% (from %s to %s), "
            "representing %s during this period."
            % ("grew" if yoy_i >= 0 else "contracted", yoy_i * 100,
               fmt_b(a.imp_total_prev), fmt_b(imp_t),
               fmt_b(imp_t - (a.imp_total_prev or 0))))
    else:
        annex2.append(
            "During the reviewed period, the leading import products were "
            "%s."
            % ", ".join("%s (%s; %s of the bill)"
                        % (short_prod(x[0], 30), fmt_b(x[1]), pct_s(x[2]))
                        for x in imp_prods[:6]))
        annex2.append(
            "The top ten import products accounted for %s of the total "
            "import bill of %s."
            % (pct_s(a.cum_share(a2, 10)), fmt_b(imp_t)))
    n["annex2"] = annex2
    return n


# Charts (matplotlib, house style shared with the other report generators)
# ---------------------------------------------------------------------------
PIE_PALETTE = list(THEME_ACCENTS) + [lighten(c, 0.45) for c in THEME_ACCENTS]


def make_chart_balance(a: QuarterAnalysis, out_path):
    """Clustered columns of monthly Exports / Imports / Balance."""
    plt.rcParams["font.family"] = chart_font()
    fig, ax = plt.subplots(figsize=(7.6, 4.4), dpi=160)
    months = a.month_names
    x = list(range(len(months)))
    w = 0.26
    cols = [lighten(THEME_ACCENTS[0], 0.30), lighten(THEME_ACCENTS[1], 0.30),
            lighten(THEME_ACCENTS[2], 0.30)]
    exp_m, imp_m = a.exp_months(), a.imp_months()
    bal_m = a.balance_months()
    ax.bar([i - w for i in x], exp_m, width=w, label="Exports",
           color=cols[0])
    ax.bar([i for i in x], imp_m, width=w, label="Imports", color=cols[1])
    ax.bar([i + w for i in x], bal_m, width=w, label="Balance of Trade",
           color=cols[2])
    ax.axhline(0, color="#555555", linewidth=0.8)
    ax.set_xticks(x)
    ax.set_xticklabels(months, fontsize=9)
    ax.set_ylabel("Value in Ksh. Billion", fontsize=10)
    ax.set_title("Kenya Balance of Trade %s %d" % (a.quarter, a.year),
                 fontsize=12, weight="bold")
    ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, _: f"{x:,.1f}"))
    ax.yaxis.grid(True, linestyle="--", alpha=0.35)
    ax.set_axisbelow(True)
    ax.legend(loc="upper center", bbox_to_anchor=(0.5, -0.10), ncol=3,
              frameon=False, fontsize=10)
    for spine in ("top", "right"):
        ax.spines[spine].set_visible(False)
    fig.tight_layout()
    fig.savefig(out_path, bbox_inches="tight")
    plt.close(fig)


def make_chart_share(a: QuarterAnalysis, table, title, out_path,
                     other_label="Others"):
    """Share doughnuts.

    Comparison mode: two pies stacked vertically ('SHARE IN <Y>' above
    'SHARE IN <Y-1>'), sized to fill a full A4 page. Single-year mode:
    one centred doughnut.
    """
    Yc, Yp = a.year, a.year_prev

    def side(idx):
        labels, values = [], []
        for name, v, _s in a.shares(table, year=("cur" if idx == 1
                                                 else "prev")):
            labels.append(short_prod(name, 40))
            values.append(v if v else 0.0)
        return consolidate(labels, values, max_slices=8, min_pct=2.0,
                           other_label=other_label)

    plt.rcParams["font.family"] = chart_font()
    if a.has_comparison:
        # full A4 portrait content page: two pies stacked vertically,
        # each labelled with a right-hand call-out legend. Saved directly
        # (no tight bbox) so the image keeps its full-page dimensions.
        fig = plt.figure(figsize=(7.5, 10.4), dpi=160)
        fig.patch.set_facecolor("white")
        for y0, idx, yr in ((0.545, 1, Yc), (0.055, 0, Yp)):
            ax = fig.add_axes([0.03, y0, 0.56, 0.36])
            labels, values = side(idx)
            labels2, shares2, wedges = draw_share_pie(
                ax, labels, values, PIE_PALETTE[:len(values)],
                style="3d_exploded", other_label=other_label)
            total = sum(shares2) or 1.0
            entries = ["%s \u2013 %.1f%%" % (l, v / total * 100)
                       for l, v in zip(labels2, shares2)]
            leg = ax.legend(wedges, entries, loc="center left",
                            bbox_to_anchor=(1.04, 0.5), frameon=False,
                            fontsize=10, handlelength=1.3,
                            labelspacing=1.0, borderaxespad=0.0)
            for t in leg.get_texts():
                t.set_fontsize(10)
            ax.set_title("SHARE IN %d" % yr, fontsize=14, weight="bold",
                         y=1.06)
        fig.savefig(out_path, facecolor="white")
        plt.close(fig)
        return

    labels, values = side(1)
    fig, ax = new_fig()
    labels2, shares2, wedges = draw_share_pie(
        ax, labels, values, PIE_PALETTE[:len(values)],
        style="3d_exploded", other_label=other_label)
    slice_callouts(fig, wedges, labels2, shares2)
    ax.set_title("Share in %d" % Yc, fontsize=12, weight="bold")
    finish(fig, out_path)


# ---------------------------------------------------------------------------
# Quarterly Word tables
# ---------------------------------------------------------------------------
GREEN = RGBColor(0x00, 0xB0, 0x50)
RED = RGBColor(0xFF, 0x00, 0x00)
GREY = RGBColor(0x80, 0x80, 0x80)


def _arrow_text(value, kind):
    """(arrow, text, color) for a change / % change value."""
    if value is None:
        return "", "–", GREY
    up = value >= 0
    arrow = "\u25b2" if up else "\u25bc"
    color = GREEN if up else RED
    txt = ("%.1f%%" % (value * 100)) if kind == "pct" else num(value)
    return arrow, txt, color


def _fill_arrow_cell(b, cell, value, kind):
    """Write an arrow + coloured value into a table cell."""
    from docx.enum.text import WD_ALIGN_PARAGRAPH as _AL
    p = cell.paragraphs[0]
    p.alignment = _AL.CENTER
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.space_before = Pt(2)
    arrow, txt, color = _arrow_text(value, kind)
    if arrow:
        r = p.add_run(arrow + " ")
        r.font.name = "Century Gothic"
        r.font.size = Pt(9)
        r.font.bold = True
        r.font.color.rgb = color
    r2 = p.add_run(txt)
    r2.font.name = "Century Gothic"
    r2.font.size = Pt(10)
    r2.font.bold = True
    r2.font.color.rgb = color


def _add_q_table_simple(b: ReportBuilder, parsed, item_header):
    """Single-year table: Rank | Item | months | Total | Share."""
    Y = b.year
    n = len(parsed["months"])
    ncols = n + 4
    data_rows = len(parsed["items"]) + (1 if parsed["all_other"] else 0)
    rows = 3 + data_rows
    table = b.doc.add_table(rows=rows, cols=ncols)
    from docx.enum.table import WD_TABLE_ALIGNMENT
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    widths = [700, 2750] + [900] * n + [1000, 950]
    b._set_table_widths(table, widths)

    b._cell_text(table.cell(0, 0), "Rank \nin\n%d" % Y, bold=True,
                 align=WD_ALIGN_PARAGRAPH.CENTER)
    b._cell_text(table.cell(0, 1), item_header, bold=True, wrap=True,
                 align=WD_ALIGN_PARAGRAPH.CENTER)
    for k, m in enumerate(parsed["months"]):
        b._cell_text(table.cell(0, 2 + k), m, bold=True,
                     align=WD_ALIGN_PARAGRAPH.CENTER)
    b._cell_text(table.cell(0, 2 + n), "%d\nTotal" % Y, bold=True,
                 align=WD_ALIGN_PARAGRAPH.CENTER)
    b._cell_text(table.cell(0, 3 + n), "Share \nin\n%d" % Y, bold=True,
                 align=WD_ALIGN_PARAGRAPH.CENTER)
    b._cell_text(table.cell(1, 3 + n), "%", bold=True,
                 align=WD_ALIGN_PARAGRAPH.CENTER)

    ucell = table.cell(1, 2)
    b._cell_text(ucell, "Value in Ksh. Billion", bold=True,
                 align=WD_ALIGN_PARAGRAPH.CENTER)
    ucell.merge(table.cell(1, 2 + n))

    def emit(d, ri, bold=False):
        b._cell_text(table.cell(ri, 0),
                     str(d["rank"]) if d.get("rank") is not None else "",
                     bold=bold, align=WD_ALIGN_PARAGRAPH.CENTER)
        b._cell_text(table.cell(ri, 1), str(d["name"]), bold=bold, wrap=True)
        for k, v in enumerate(d["months"][1]):
            b._cell_text(table.cell(ri, 2 + k),
                         num(v) if v is not None else "",
                         bold=bold, align=WD_ALIGN_PARAGRAPH.CENTER)
        tot = d["totals"][1]
        b._cell_text(table.cell(ri, 2 + n),
                     num(tot) if tot is not None else "",
                     bold=bold, align=WD_ALIGN_PARAGRAPH.CENTER)
        sh = d.get("share")
        b._cell_text(table.cell(ri, 3 + n),
                     ("%.1f%%" % (sh * 100)) if sh is not None else "",
                     bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)

    ri = 2
    for d in parsed["items"]:
        emit(d, ri)
        ri += 1
    if parsed["all_other"]:
        emit(parsed["all_other"], ri)
        ri += 1
    grand = {"name": parsed["grand"].get("name", "Grand Total"),
             "rank": None,
             "months": [None, parsed["grand"]["months"][1]],
             "totals": [None, parsed["grand"]["totals"][1]],
             "share": parsed["grand"].get("share", 1.0)}
    emit(grand, ri, bold=True)
    b._fit_table_on_page(table)
    return table


def add_q_table(b: ReportBuilder, parsed, item_header):
    """Two-year comparison table (Rank | Item | <Y-1> | <Y> | Change).

    Columns: rank, item, prev months+Total, current months+Total,
    Total change and % with green up / red down arrows.
    Falls back to the single-year layout when no prior-year data exists.
    """
    if not parsed.get("comparison"):
        return _add_q_table_simple(b, parsed, item_header)
    Yc = b.year
    Yp = Yc - 1
    n = len(parsed["months"])
    ncols = 4 + 2 * n + 2                 # rank,item,(n+1)*2,chg,pct
    data_rows = len(parsed["items"]) + (1 if parsed["all_other"] else 0)
    rows = 2 + data_rows + 1              # header x2 + data + grand
    table = b.doc.add_table(rows=rows, cols=ncols)
    from docx.enum.table import WD_TABLE_ALIGNMENT
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    c_rank, c_item = 0, 1
    p_first = 2
    p_total = p_first + n
    c_first = p_total + 1
    c_total = c_first + n
    c_chg = c_total + 1
    c_pct = c_chg + 1

    widths = [430, 2350] + [640] * n + [760] + [640] * n + [760] \
        + [820, 700]
    b._set_table_widths(table, widths)

    def hcell(r, c, text):
        b._cell_text(table.cell(r, c), text, bold=True,
                     align=WD_ALIGN_PARAGRAPH.CENTER)

    # header row 0: merged group headers
    hcell(0, c_rank, "Rank")
    hcell(0, c_item, item_header)
    table.cell(0, c_rank).merge(table.cell(1, c_rank))
    table.cell(0, c_item).merge(table.cell(1, c_item))
    table.cell(0, p_first).merge(table.cell(0, p_total))
    hcell(0, p_first, str(Yp))
    table.cell(0, c_first).merge(table.cell(0, c_total))
    hcell(0, c_first, str(Yc))
    table.cell(0, c_chg).merge(table.cell(0, c_pct))
    hcell(0, c_chg, "Change in\n%d-%d" % (Yp, Yc))

    # header row 1: month labels / totals / change labels
    for base, tcol in ((p_first, p_total), (c_first, c_total)):
        for k, m in enumerate(parsed["months"]):
            hcell(1, base + k, m[:3])
        hcell(1, tcol, "Total")
    hcell(1, c_chg, "Total change")
    hcell(1, c_pct, "%")

    def emit(d, ri, bold=False):
        b._cell_text(
            table.cell(ri, c_rank),
            str(d["rank"]) if d.get("rank") is not None else "",
            bold=bold, align=WD_ALIGN_PARAGRAPH.CENTER)
        b._cell_text(table.cell(ri, c_item), str(d["name"]), bold=bold,
                     wrap=True)
        vals = ((d["months"][0], d["totals"][0], p_first, p_total),
                (d["months"][1], d["totals"][1], c_first, c_total))
        for mvals, tot, first, tcol in vals:
            for k, v in enumerate(mvals):
                b._cell_text(table.cell(ri, first + k),
                             num(v) if v is not None else "",
                             bold=bold, align=WD_ALIGN_PARAGRAPH.CENTER)
            b._cell_text(table.cell(ri, tcol),
                         num(tot) if tot is not None else "",
                         bold=bold, align=WD_ALIGN_PARAGRAPH.CENTER)
        _fill_arrow_cell(b, table.cell(ri, c_chg), d.get("change"), "chg")
        _fill_arrow_cell(b, table.cell(ri, c_pct), d.get("pct"), "pct")

    ri = 2
    for d in parsed["items"]:
        emit(d, ri)
        ri += 1
    if parsed["all_other"]:
        emit(parsed["all_other"], ri)
        ri += 1
    emit(parsed["grand"], ri, bold=True)
    b._fit_table_on_page(table)
    return table


def add_growth_decline_table(b: ReportBuilder, growers, decliners,
                             left_title, right_title,
                             threshold_label="Ksh. 1 billion"):
    """Two-column exhibit: markets/products that grew vs declined.

    Each entry reads ``Name  <arrow> Ksh X.X billion`` with the arrow and
    amount coloured green (growth) or red (decline).
    """

    def shade(cell, hexcolor):
        from docx.oxml.ns import qn as _qn
        from docx.oxml import OxmlElement as _El
        tcPr = cell._tc.get_or_add_tcPr()
        shd = _El("w:shd")
        shd.set(_qn("w:val"), "clear")
        shd.set(_qn("w:fill"), hexcolor)
        tcPr.append(shd)

    nrows = 1 + max(len(growers), len(decliners), 1)
    table = b.doc.add_table(rows=nrows, cols=2)
    from docx.enum.table import WD_TABLE_ALIGNMENT
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    b._set_table_widths(table, [5300, 5300])

    for c, title in ((0, left_title), (1, right_title)):
        shade(table.cell(0, c), "D9E1F2")
        b._cell_text(table.cell(0, c), title, bold=True, wrap=True, size=10)

    def fill_entry(cell, d):
        p = cell.paragraphs[0]
        p.paragraph_format.space_after = Pt(3)
        r = p.add_run(str(d["name"]))
        r.font.name = "Century Gothic"
        r.font.size = Pt(10)
        arrow, txt, color = _arrow_text(d["change"], "chg")
        r2 = p.add_run("  %s %s" % (arrow, txt))
        r2.font.name = "Century Gothic"
        r2.font.size = Pt(10)
        r2.font.bold = True
        r2.font.color.rgb = color

    def fill_none(cell):
        b._cell_text(cell, "None registered change of over %s"
                     % threshold_label, italic=True, size=10)

    for i in range(nrows - 1):
        if i < len(growers):
            fill_entry(table.cell(i + 1, 0), growers[i])
        else:
            fill_none(table.cell(i + 1, 0))
        if i < len(decliners):
            fill_entry(table.cell(i + 1, 1), decliners[i])
        else:
            fill_none(table.cell(i + 1, 1))
    return table


DEFAULT_ORG = "KENYA EXPORT PROMOTION & BRANDING AGENCY"
DEFAULT_DIRECTORATE = "Research and Innovation Directorate (RID)"


def _q_short(a):
    """'Apr-Jun' style abbreviated quarter label (reference-report style)."""
    ms = a.month_names
    if len(ms) == 1:
        return ms[0][:3].capitalize()
    return "%s-%s" % (ms[0][:3].capitalize(), ms[-1][:3].capitalize())


def _table_image(b: ReportBuilder, parsed, tmp_dir, name, item_header,
                 caption=None, source=SRC_KRA):
    """Caption + full-width table image + source, reference-report style.

    The big rank tables are embedded as images (like the pasted Excel
    tables in the reference report) so each table always occupies a single
    page with its Century Gothic 12 formatting intact.
    """
    if caption:
        b.add_table_caption(caption)
    img = render_table_image(parsed, os.path.join(tmp_dir, name),
                             item_header=item_header, year=b.year)
    b.add_figure(img, width_in=6.27)
    if source:
        b.add_source(source)


def build_quarterly_report(excel_dir, out_path, tmp_dir,
                           year=None, org=DEFAULT_ORG,
                           directorate=DEFAULT_DIRECTORATE,
                           month_year=None):
    os.makedirs(tmp_dir, exist_ok=True)
    a = QuarterAnalysis(year=year).load(excel_dir)
    a.load_crosstabs(excel_dir)
    narr = build_narratives(a)

    Y = a.year
    y_prev = Y - 1
    q_up = a.quarter.upper()
    qs = _q_short(a)
    qs_up = qs.upper()
    comp = a.has_comparison
    y_span = "(%d-%d)" % (y_prev, Y) if comp else "(%d)" % Y

    cfg = {
        "country": {"name": "Kenya", "short": "Kenya",
                    "possessive": "Kenya's", "title": "Overview"},
        "report": {"year": Y},
    }
    b = ReportBuilder(cfg, narr)
    b.year = Y
    b.has_comparison = comp

    # ============================== TITLE PAGE ==============================
    b.add_letterhead()
    b.add_para("KENYA", size=24, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER,
               space_after=0)
    b.add_para("EXPORT PERFORMANCE IN %s %s" % (q_up, y_span),
               size=24, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER,
               space_after=24)
    b.add_para("")
    b.add_para("REPORT", size=24, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER,
               space_after=30)
    for _ in range(3):
        b.add_para("")
    b.add_para("Prepared", size=16, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER,
               space_after=6)
    b.add_para("By", size=16, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER,
               space_after=14)
    b.add_para(org, size=22, bold=True, italic=True,
               align=WD_ALIGN_PARAGRAPH.CENTER, space_after=28)
    if not month_year:
        month_year = date.today().strftime("%B %Y")
    b.add_para(month_year, size=22, bold=True,
               align=WD_ALIGN_PARAGRAPH.CENTER)
    b.page_break()

    # ============ LIST OF TABLES / LIST OF FIGURES (one page) ===============
    b.add_list_placeholder("List of Tables", "table")
    b.add_para("")
    b.add_list_placeholder("List of Figures", "figure")
    # page numbers start at 1 from here (title/lists carry none)
    b.start_body_section()

    # ============================== SECTION 1 ===============================
    p = b.doc.add_paragraph()
    r = p.add_run("KENYA\u2019S TRADE PERFORMANCE IN %s %s"
                  % (qs_up, ("(%d & %d)" % (y_prev, Y)) if comp
                     else ("(%d)" % Y)))
    b._style_run(r, bold=True)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    for para in narr["overview"]:
        b.add_para(para)
    b.page_break()

    # ============================== SECTION 2 ===============================
    if comp:
        b.add_heading("Comparative Perspectives of Kenya\u2019s Trade in "
                      "%s (%d/%d)" % (qs, y_prev, Y))
    else:
        b.add_heading("Comparative Perspectives of Kenya\u2019s Trade in "
                      "%s %d" % (qs, Y))
    b.add_source(SRC_KRA)
    for line in narr["s2"]:
        b.add_para(line)
    b.page_break()

    # ============================== SECTION 3 ===============================
    b.add_heading("Lead Destination Markets %s %s" % (qs, y_span))
    _table_image(b, a.t1, tmp_dir, "table1_markets.png", "PARTNER",
                 caption="Table 1: Top Destination Markets in %s %s"
                         % (qs, y_span))
    if comp:
        # exhibit: markets that grew vs declined by over Ksh 1 billion
        g_mkt, d_mkt = a.growth_decline(a.t1, threshold=1.0)
        if g_mkt or d_mkt:
            add_growth_decline_table(
                b, g_mkt, d_mkt,
                "Export markets that Registered Growth (by more than 1 "
                "billion)",
                "Economies whose imports from Kenya Declined (by more "
                "than Kshs 1 billion)")
            b.add_source(SRC_KRA)
        b.add_para(narr["markets_note"])
    for bullet in narr["market_bullets"]:
        b.add_bullet(bullet)

    b.add_heading("Comparative Perspective on Shares of Lead Export "
                  "Markets %s %s"
                  % (qs, ("(%d & %d)" % (y_prev, Y)) if comp
                     else ("(%d)" % Y)), level=2)
    b.add_table_caption("Figure 1: Share of Exports by Markets, %s %s"
                        % (qs, y_span))
    for _yr in (a.year, a.year_prev):
        if _yr is None:
            continue
        _cur = (_yr == a.year)
        if not _cur and not comp:
            break
        _labels, _values = _share_doughnut_data(a, a.t1, "Other markets",
                                                year=("cur" if _cur else "prev"))
        if not _labels or not _values:
            b.add_para("[Market share chart data not available]", italic=True,
                       color=RGBColor(0x9A, 0x1F, 0x1F))
            break
        b.add_word_chart(
            "doughnut", "Share of Exports by Markets: %d" % _yr,
            [short_prod(l, 42) for l in _labels], list(_values),
            width_in=6.3, height_in=(4.6 if _cur else 3.6),
            name="Figure 1 - Export Shares %d" % _yr,
            colors=[c.lstrip("#") for c in PIE_PALETTE])
    b.add_source(SRC_KRA)
    b.page_break()

    # ============================== SECTION 4 ===============================
    b.add_heading("Composition of Kenya\u2019s Export Commodities")
    _table_image(b, a.t2, tmp_dir, "table2_products.png", "Product label",
                 caption="Table 2: Top Export Products in %s %s"
                         % (qs, y_span))
    if comp:
        # exhibit: products that grew vs declined by over Ksh 1 billion
        g_prd, d_prd = a.growth_decline(a.t2, threshold=1.0)
        if g_prd or d_prd:
            add_growth_decline_table(
                b, g_prd, d_prd,
                "Higher Performing Products / (Growth by over Ksh. 1 "
                "billion)",
                "Weaker Performing Products (Decline by over Ksh. 1 "
                "billion)")
            b.add_source(SRC_KRA)
        b.add_para(narr["products_note"])
    for para in narr["product_paras"]:
        b.add_para(para)

    b.add_heading("Share of Lead Export Products in %s %s"
                  % (qs, ("(%d & %d)" % (y_prev, Y)) if comp
                     else ("(%d)" % Y)), level=2)
    b.add_table_caption("Figure 2: Share of Top Exports by Products, %s %s"
                        % (qs, y_span))
    for _yr in (a.year, a.year_prev):
        if _yr is None:
            continue
        _cur = (_yr == a.year)
        if not _cur and not comp:
            break
        _labels, _values = _share_doughnut_data(a, a.t2, "All others products",
                                                year=("cur" if _cur else "prev"))
        if not _labels or not _values:
            b.add_para("[Product share chart data not available]", italic=True,
                       color=RGBColor(0x9A, 0x1F, 0x1F))
            break
        b.add_word_chart(
            "doughnut", "Share of Top Exports by Products: %d" % _yr,
            [short_prod(l, 42) for l in _labels], list(_values),
            width_in=6.3, height_in=(4.6 if _cur else 3.6),
            name="Figure 2 - Product Shares %d" % _yr,
            colors=[c.lstrip("#") for c in PIE_PALETTE])
    b.add_source(SRC_KRA)
    b.page_break()

    # ============================== SECTION 5 ===============================
    b.add_heading("Deduction from the Quarter performance %s %d." % (qs, Y))
    for para in narr["deduction"]:
        b.add_para(para)
    b.add_heading("Implication for the country", level=2)
    for para in narr["implication"]:
        b.add_para(para)
    b.add_para("To respond to the reported declines in exports to "
               "strategic markets and reinforce overall export "
               "competitiveness, a coordinated raft of interventions is "
               "proposed across the following areas.")
    for bullet in narr["interventions"]:
        b.add_bullet(bullet)
    b.page_break()

    # ============================== SECTION 6 ===============================
    b.add_heading("Annex")
    b.add_heading("Annex 1: Kenya\u2019s imports by supplying countries.",
                  level=1)
    _table_image(b, a.a1, tmp_dir, "annex1_import_partners.png", "PARTNER")
    for para in narr["annex1"]:
        b.add_bullet(para)
    b.page_break()

    b.add_heading("Annex 2: Kenya\u2019s imports by products.", level=1)
    _table_image(b, a.a2, tmp_dir, "annex2_import_products.png",
                 "Product Label")
    for para in narr["annex2"]:
        b.add_bullet(para)

    b.add_footer()
    return a, b.doc


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------
def main():
    ap = argparse.ArgumentParser(
        description="Generate the Kenya quarterly export performance report.")
    ap.add_argument("--excel-dir",
                    default=os.path.join(BASE_DIR, "output", "quarterly"),
                    help="Folder containing Exports.xlsx and Imports.xlsx")
    ap.add_argument("--output", default=None, help="Output .docx path")
    ap.add_argument("--tmp", default=os.path.join(BASE_DIR, "output", ".tmp"),
                    help="Temporary directory for chart images")
    ap.add_argument("--year", type=int, default=None,
                    help="Review year when it cannot be detected from data")
    ap.add_argument("--org", default=DEFAULT_ORG,
                    help="Organisation line on the title page")
    ap.add_argument("--directorate", default=DEFAULT_DIRECTORATE)
    ap.add_argument("--month-year", default=None,
                    help="Date line on the title page (e.g. 'August 2026')")
    args = ap.parse_args()

    out = args.output
    out = os.path.abspath(out) if out else os.path.join(
        BASE_DIR, "output", "KENYA EXPORT PERFORMANCE REPORT.docx")
    os.makedirs(os.path.dirname(out), exist_ok=True)

    print("[1/3] Reading quarterly workbooks from : %s"
          % os.path.abspath(args.excel_dir))
    a, doc = build_quarterly_report(
        args.excel_dir, out, args.tmp, year=args.year, org=args.org,
        directorate=args.directorate, month_year=args.month_year)
    print("[2/3] Review period                    : %s %d"
          % (a.quarter, a.year))
    doc.save(out)
    print("[3/3] Report saved to                  : %s" % out)

    # Companion editable Excel workbook (additive; never breaks the build).
    try:
        import excel_deliverable
        xlsx_out = os.path.splitext(out)[0] + " TABLES.xlsx"
        os.makedirs(os.path.dirname(out), exist_ok=True)
        excel_deliverable.build_quarterly_deliverable(
            a, {"report": {"title_line2": ""}}, xlsx_out)
        print("      Excel deliverable saved to   : %s" % xlsx_out)
    except Exception as _e:
        print("      [warn] Excel deliverable skipped: %s" % _e)


if __name__ == "__main__":
    try:
        main()
    except PipelineError as e:
        sys.exit(str(e))
