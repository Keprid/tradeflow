#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
generate_product_profile.py
===========================

General, data-driven "product profile" report generator.

Reads a folder of ITC (International Trade Centre / Trade Map) matrix
exports -- the same eight files downloaded for the Coffee product profile
(plus an optional export-potential file) -- and produces:

  * a styled Word report (.docx) that mirrors the "Product Profile for
    Textile and apparel 2024" template (front matter, trend table, Kenya's
    destination markets, global exporter/importer rankings, global trend
    tables and an optional export-potential section), rendered with native
    auto-fitting tables instead of screenshots; and
  * an editable Excel workbook ("... TABLES.xlsx") with the same tables and
    editable doughnut charts.

The tool is fully generic: the product family, its member HS codes, the
anchor product and the review years are all detected from the data, so the
same script produces a profile for any product family the ITC files are
downloaded for.

Input file contract (a plain folder, names matched by prefix):
    exporting-economies_*.xlsx            world exports by economy (anchor)
    importing-economies_*.xlsx            world imports by economy (anchor)
    kenyas-exports-to-world-by-importer_*.xlsx   Kenya exports by partner (anchor)
    kenyas-exports-to-world-by-product_*.xlsx    Kenya exports by 4/6-digit product
    kenyas-imports-from-world-by-exporter_*.xlsx Kenya imports by partner (anchor)
    kenyas-imports-from-world-by-product_*.xlsx  Kenya imports by product
    products-exported-globally_*.xlsx     world exports by product (family)
    products-imported-globally_*.xlsx     world imports by product (family)
    export_potential_*.xlsx  [optional]   markets x {potential, actual}

Usage:
    python generate_product_profile.py \
        --config config/product_profile_coffee.json [--output path] [--tmp dir]
"""

import argparse
import glob
import json
import os
import re

import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

import matplotlib

matplotlib.use("Agg")

from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_SECTION_START
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor

from generate_report import (ReportBuilder, num, pct, clean_label, short_label,
                             to_float)
import charts
from excel_deliverable import THEME

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

NAVY = "1F3864"
THIN = Side(style="thin", color="999999")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

FILE_PREFIXES = {
    "exporting-economies": "world_exports_by_economy",
    "importing-economies": "world_imports_by_economy",
    "kenyas-exports-to-world-by-importer": "kenya_exports_by_partner",
    "kenyas-exports-to-world-by-product": "kenya_exports_by_product",
    "kenyas-imports-from-world-by-exporter": "kenya_imports_by_partner",
    "kenyas-imports-from-world-by-product": "kenya_imports_by_product",
    "products-exported-globally": "world_exports_by_product",
    "products-imported-globally": "world_imports_by_product",
    "export_potential": "export_potential",
}


# --------------------------------------------------------------------------
# ITC matrix loading
# --------------------------------------------------------------------------
def _find_file(data_dir, prefix):
    hits = [p for p in glob.glob(os.path.join(data_dir, "*")) if
            os.path.basename(p).lower().startswith(prefix)]
    return sorted(hits)[0] if hits else None


def load_matrix(path):
    """Parse an ITC all-countries / all-products matrix.

    Returns ``(years, records)`` where ``years`` is the ordered list of
    review years picked up from the column headers and each record is:
    ``{"reporter", "reporter_label", "partner", "partner_label",
        "product", "product_label", "years": {year: value}}``.
    """
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    ws = wb.worksheets[0]
    rows = list(ws.iter_rows(values_only=True))
    wb.close()
    header = rows[0] if rows else []
    years, year_cols = [], []
    for i, h in enumerate(header or []):
        m = re.match(r"\s*(20\d\d)\s*\(", str(h or ""))
        if m:
            years.append(int(m.group(1)))
            year_cols.append(i)
    records = []
    for r in rows[1:]:
        if not r or r[0] is None:
            continue
        records.append({
            "reporter": str(r[0]),
            "reporter_label": str(r[1] or ""),
            "partner": str(r[2]),
            "partner_label": str(r[3] or ""),
            "product": str(r[4]),
            "product_label": clean_label(str(r[5] or "")),
            "years": {y: to_float(r[c]) for y, c in zip(years, year_cols)},
        })
    return years, records


class ProfileData:
    """Loads every ITC file in the folder and derives the analysis tables."""

    def __init__(self, data_dir):
        self.data_dir = os.path.abspath(data_dir)
        self.files = {}
        missing = []
        for prefix, key in FILE_PREFIXES.items():
            path = _find_file(self.data_dir, prefix)
            if path is None:
                if key == "export_potential":
                    continue
                missing.append(prefix)
                continue
            years, records = load_matrix(path)
            self.files[key] = {"path": path, "years": years,
                               "records": records}
        if missing:
            raise OSError("Missing ITC files in %s: %s"
                          % (self.data_dir, ", ".join(missing)))

        self.all_years = sorted({y for k in self.files.values()
                                 for y in k["years"]})

        # Anchor product (the product of the by-importer download)
        partner_rows = self._rows("kenya_exports_by_partner")
        anchor_product = None
        if partner_rows:
            counts = {}
            for r in partner_rows:
                counts[r["product"]] = counts.get(r["product"], 0) + 1
            anchor_product = max(counts, key=counts.get)
        if anchor_product is None:
            raise OSError("Could not detect the anchor product from the "
                          "by-importer file.")
        self.anchor_hs = anchor_product
        self.anchor_label = next(
            (r["product_label"] for r in partner_rows
             if r["product"] == anchor_product), "")

        # Total rows: a product code whose series equals the sum of all the
        # other codes in the same file (the basket aggregate of a selection
        # download). They are header rows, not family members.
        self._file_totals = {}
        for key in ("kenya_exports_by_product", "kenya_imports_by_product",
                    "world_exports_by_product", "world_imports_by_product"):
            if key in self.files:
                self._file_totals[key] = self._total_codes(key)

        # Family members: product detail rows of the by-product download
        # (its total row, if any, is excluded).
        total = self._file_totals.get("kenya_exports_by_product", set())
        members = []
        for r in self._rows("kenya_exports_by_product"):
            if r["partner"] != "000" or r["product"] in total:
                continue
            members.append({"code": r["product"], "label": r["product_label"],
                            "years": r["years"]})
        members.sort(key=lambda m: m["years"].get(self.review_year) or 0.0,
                     reverse=True)
        self.members = members
        self.anchor_is_total = anchor_product in total

    def _total_codes(self, key):
        """Codes in ``key`` whose review-year value equals the sum of all the
        other codes in the same file (the selection total row)."""
        by_code = {}
        for r in self._rows(key):
            if r["partner"] != "000":
                continue
            by_code.setdefault(r["product"], r["years"])
        codes = [c for c in by_code if by_code[c].get(self.review_year)]
        if len(codes) < 3:
            return set()
        grand = sum(by_code[c].get(self.review_year) or 0.0 for c in codes)
        out = set()
        for c in codes:
            v = by_code[c].get(self.review_year) or 0.0
            if grand and abs(v - (grand - v)) / grand < 0.01:
                out.add(c)
        return out

    # -- accessors ----------------------------------------------------------
    def _rows(self, key):
        return self.files.get(key, {}).get("records", [])

    @property
    def review_year(self):
        return self.all_years[-1] if self.all_years else None

    @property
    def start_year(self):
        return self.all_years[0] if self.all_years else None

    @property
    def years(self):
        return self.all_years

    def destinations(self):
        rows = sorted((r for r in self._rows("kenya_exports_by_partner")
                       if r["partner"] != "000"),
                      key=lambda r: r["years"].get(self.review_year) or 0.0,
                      reverse=True)
        return [{"label": r["partner_label"], "years": r["years"]}
                for r in rows]

    def exporters(self):
        rows = sorted((r for r in self._rows("world_exports_by_economy")
                       if r["reporter"] != "000"),
                      key=lambda r: r["years"].get(self.review_year) or 0.0,
                      reverse=True)
        return [{"label": r["reporter_label"], "years": r["years"]}
                for r in rows]

    def importers(self):
        rows = sorted((r for r in self._rows("world_imports_by_economy")
                       if r["reporter"] != "000"),
                      key=lambda r: r["years"].get(self.review_year) or 0.0,
                      reverse=True)
        return [{"label": r["reporter_label"], "years": r["years"]}
                for r in rows]

    def kenya_import_sources(self):
        rows = sorted((r for r in self._rows("kenya_imports_by_partner")
                       if r["partner"] != "000"),
                      key=lambda r: r["years"].get(self.review_year) or 0.0,
                      reverse=True)
        return [{"label": r["partner_label"], "years": r["years"]}
                for r in rows]

    def kenya_import_products(self):
        total = self._file_totals.get("kenya_imports_by_product", set())
        rows = sorted((r for r in self._rows("kenya_imports_by_product")
                       if r["partner"] == "000" and r["product"] not in total),
                      key=lambda r: r["years"].get(self.review_year) or 0.0,
                      reverse=True)
        return [{"label": r["product_label"], "years": r["years"]}
                for r in rows]

    def global_export_products(self):
        total = self._file_totals.get("world_exports_by_product", set())
        rows = sorted((r for r in self._rows("world_exports_by_product")
                       if r["partner"] == "000" and r["product"] not in total),
                      key=lambda r: r["years"].get(self.review_year) or 0.0,
                      reverse=True)
        return [{"label": r["product_label"], "years": r["years"]}
                for r in rows]

    def global_import_products(self):
        total = self._file_totals.get("world_imports_by_product", set())
        rows = sorted((r for r in self._rows("world_imports_by_product")
                       if r["partner"] == "000" and r["product"] not in total),
                      key=lambda r: r["years"].get(self.review_year) or 0.0,
                      reverse=True)
        return [{"label": r["product_label"], "years": r["years"]}
                for r in rows]


# --------------------------------------------------------------------------
# Numeric / phrasing helpers
# --------------------------------------------------------------------------
def cagr(values, years):
    """Average annual growth over the first..last populated year."""
    ys = {y: v for y, v in zip(years, values) if v}
    if len(ys) < 2:
        return None
    yrs = sorted(ys)
    first, last = ys[yrs[0]], ys[yrs[-1]]
    if first <= 0 or last <= 0:
        return None
    return (last / first) ** (1.0 / (len(yrs) - 1)) - 1.0


def yoy_change(values, years):
    """Growth between the last two populated years."""
    ys = {y: v for y, v in zip(years, values) if v is not None}
    yrs = sorted(ys)
    if len(yrs) < 2:
        return None
    base, cur = ys[yrs[-2]], ys[yrs[-1]]
    if base is None or base <= 0:
        return None
    return (cur / base) - 1.0


def display(value, scale=1e6):
    """ITC values arrive in USD thousands; scale to USD millions."""
    if value is None:
        return None
    return value * 1000.0 / scale


def fmt(v, decimals=1):
    d = display(v)
    return "" if d is None else num(d, decimals)


def usd_phrase(v):
    """'USD 353.8 Million' style phrasing for narratives."""
    if v is None:
        return ""
    d = display(v)
    word = "Million"
    if d >= 1000.0:
        d, word = d / 1000.0, "Billion"
    return "USD %s %s" % (num(d, 1), word)


def growth_phrase(g, period):
    if g is None:
        return None
    verb = "contracted" if g < 0 else "grew"
    return "%s at an average annual rate of %.1f%% %s" % (
        verb, abs(g) * 100.0, period)


def yoy_phrase(g, y1, y2):
    if g is None:
        return None
    verb = "declined by" if g < 0 else "rose by"
    return "%s %.1f%% between %d and %d" % (verb, abs(g) * 100.0, y1, y2)


def period_phrase(y1, y2):
    return "between %d and %d" % (y1, y2)


def short_anchor(label):
    """Human-friendly product name for headings and prose, e.g. 'Coffee'."""
    clean = clean_label(label)
    head = re.split(r"[,;]", clean)[0].strip()
    return head or short_label(clean, 24)


def ordinal_list(names, sep=", ", last=" and "):
    if not names:
        return ""
    if len(names) == 1:
        return names[0]
    return sep.join(names[:-1]) + last + names[-1]


def _shares(rows, years):
    """(label, share_fraction) pairs from the review year of ``rows``."""
    last = years[-1]
    total = sum((r["years"].get(last) or 0.0) for r in rows)
    out = []
    for r in rows:
        v = r["years"].get(last)
        out.append((r["label"], (v or 0.0) / total if total else 0.0))
    return out


def top_rows(rows, n, years, residual="All other markets"):
    """Keep ``n`` rows, consolidating the rest into a residual row."""
    rows = list(rows)
    if len(rows) <= n:
        return rows
    top = rows[:n]
    residual_years = {}
    for y in years:
        shown = sum((r["years"].get(y) or 0.0) for r in top)
        total = sum((r["years"].get(y) or 0.0) for r in rows)
        residual_years[y] = max(0.0, total - shown)
    return top + [{"label": residual, "years": residual_years}]


def _year_totals(rows):
    years = set()
    for r in rows:
        years |= set(r["years"])
    return {y: sum((r["years"].get(y) or 0.0) for r in rows) for y in years}


# --------------------------------------------------------------------------
# Word builder
# --------------------------------------------------------------------------
class ProfileBuilder(ReportBuilder):
    def __init__(self, cfg, narratives=None):
        cfg = dict(cfg)
        cfg.setdefault("country", {"name": cfg.get("family_title", "")})
        super().__init__(cfg, narratives or {})
        self.cfg = cfg
        self.tcap = 0
        self.fcap = 0

    def _next_table(self, title, source):
        self.tcap += 1
        self.add_table_caption("Table %d: %s" % (self.tcap, title))
        self.add_source(source)

    def _next_figure(self, title, source):
        self.fcap += 1
        self.add_table_caption("Figure %d: %s" % (self.fcap, title))
        self.add_source(source)

    def add_value_table(self, first_col_header, rows, years, share_header,
                        title, source, total_label=None, rank=False,
                        unit_label="USD Million", widths=None):
        """Structure a value matrix table.

        ``rows`` = list of ``{"label", "years": {y: v}}`` already truncated
        and consolidated. The share column shows each row's share of the
        total in the review year (bold); a ``total_label`` row (optional)
        sums the displayed rows.
        """
        n = len(years)
        label_cols = 2 if rank else 1
        cols = label_cols + n + 1
        nrows = 2 + len(rows) + (1 if total_label else 0)
        table = self.doc.add_table(rows=nrows, cols=cols, style="Table Grid")
        table.alignment = WD_TABLE_ALIGNMENT.CENTER

        hdr = table.rows[0]
        if rank:
            hdr.cells[0].text = "#"
        hdr.cells[rank].text = first_col_header
        if n > 1:
            hdr.cells[label_cols].merge(hdr.cells[label_cols + n - 1])
        hdr.cells[label_cols].text = "Value in %s" % unit_label
        hdr.cells[label_cols + n].text = share_header

        r2 = table.rows[1]
        for c in range(cols):
            r2.cells[c].text = ""
        if rank:
            r2.cells[1].text = first_col_header
        for i, y in enumerate(years):
            r2.cells[label_cols + i].text = str(y)

        # shares are relative to the total of the review year across rows
        rev = years[-1]
        rev_total = sum((r["years"].get(rev) or 0.0) for r in rows)
        for ri, row in enumerate(rows, start=2):
            if rank:
                table.rows[ri].cells[0].text = str(ri - 1)
            table.rows[ri].cells[rank].text = row["label"]
            for i, y in enumerate(years):
                table.rows[ri].cells[label_cols + i].text = fmt(row["years"].get(y))
            cur = row["years"].get(rev)
            share = ((cur or 0.0) / rev_total * 100.0) if rev_total else None
            table.rows[ri].cells[label_cols + n].text = \
                "" if share is None else "%.1f%%" % share

        if total_label:
            t = table.rows[2 + len(rows)]
            t.cells[rank].text = total_label
            for i, y in enumerate(years):
                tot = sum((r["years"].get(y) or 0.0) for r in rows)
                t.cells[label_cols + i].text = fmt(tot)
            t.cells[label_cols + n].text = "100.0%"

        if widths is None:
            widths = ([420] if rank else []) + [3000] + [700] * n + [800]
        self._set_table_widths(table, widths)
        self._style_table(table, rank=rank, label_cols=label_cols, n=n,
                          total_label=total_label)
        self._fit_table_on_page(table)
        return table

    def _style_table(self, table, rank=False, label_cols=1, n=1,
                     total_label=None):
        cols = label_cols + n + 1
        share_col = cols - 1
        last_row = len(table.rows) - 1
        for ri, row in enumerate(table.rows):
            for ci, cell in enumerate(row.cells):
                if ci >= cols:
                    continue
                numeric = ci >= label_cols and ci != share_col
                align = WD_ALIGN_PARAGRAPH.CENTER if numeric or ci == 0 \
                    else WD_ALIGN_PARAGRAPH.LEFT
                for p in cell.paragraphs:
                    p.alignment = align
                    if ri < 2:
                        continue
                    for r in p.runs:
                        if ci == share_col or ri == last_row:
                            r.font.bold = True
        self._fill_header(table)

    def _fill_header(self, table):
        for ri in range(2):
            for cell in table.rows[ri].cells:
                tcPr = cell._tc.get_or_add_tcPr()
                for el in tcPr.findall(qn("w:shd")):
                    tcPr.remove(el)
                shd = OxmlElement("w:shd")
                shd.set(qn("w:val"), "clear")
                shd.set(qn("w:color"), "auto")
                shd.set(qn("w:fill"), NAVY)
                tcPr.append(shd)
                for p in cell.paragraphs:
                    for r in p.runs:
                        r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    # -- front matter -------------------------------------------------------
    def title_page(self, cfg):
        self.add_letterhead()
        self.doc.add_paragraph()
        self.doc.add_paragraph()
        t = self.add_para(cfg.get("title", "Product Profile"),
                          align=WD_ALIGN_PARAGRAPH.CENTER)
        for r in t.runs:
            r.font.size = Pt(24)
            r.font.bold = True
            r.font.color.rgb = RGBColor(0x15, 0x60, 0x82)
        self.doc.add_paragraph()
        s = self.add_para(cfg.get("subtitle", ""),
                          align=WD_ALIGN_PARAGRAPH.CENTER)
        for r in s.runs:
            r.font.size = Pt(14)
            r.font.italic = True
        self.doc.add_paragraph()
        for _ in range(6):
            self.doc.add_paragraph()
        d = self.add_para(cfg.get("date_line", ""),
                          align=WD_ALIGN_PARAGRAPH.CENTER)
        for r in d.runs:
            r.font.size = Pt(13)
        self.doc.add_paragraph()
        prep = self.add_para("Prepared by",
                             align=WD_ALIGN_PARAGRAPH.CENTER)
        for r in prep.runs:
            r.font.size = Pt(12)
        o = self.add_para(cfg.get("org", ""),
                          align=WD_ALIGN_PARAGRAPH.CENTER)
        for r in o.runs:
            r.font.size = Pt(13)
            r.font.bold = True
        self.page_break()
        self.add_toc()
        self.add_list_placeholder("List of Tables", "table")
        self.page_break()
        self.add_list_placeholder("List of Figures", "figure")
        self.start_body_section()
        self.add_footer()


# --------------------------------------------------------------------------
# Sections
# --------------------------------------------------------------------------
def make_donut(pairs, tmp_dir, name, title):
    """Render a doughnut of (label, share) pairs to a PNG and return its path."""
    pairs = [(l, s) for l, s in pairs if s > 0.0]
    if not pairs:
        return None
    labels = [l for l, _ in pairs]
    values = [s * 100.0 for _, s in pairs]
    colors = [c if c.startswith("#") else "#" + c for c in THEME]
    fig, ax = charts.new_fig(width=6.6, height=4.4)
    labels, values, wedges = charts.draw_share_pie(
        ax, labels, values, colors, style="donut", min_pct=1.0, max_slices=8)
    charts.share_legend(fig, wedges, labels, values, ncol=2)
    path = os.path.join(tmp_dir, name)
    charts.finish(fig, path)
    return path


def family_members_line(data, top=12):
    members = data.members
    if not members:
        return ""
    by_code = {m["code"]: m["label"] for m in members}
    shown = members[:top]
    items = ["%s (%s)" % (m["code"], short_label(by_code[m["code"]], 60))
             for m in shown]
    tail = ""
    if len(members) > top:
        tail = ", among %d product headings in total" % len(members)
    return ("The family covers the following product headings, ranked by the "
            "value of Kenya's exports in %d: %s%s."
            % (data.review_year, ordinal_list(items), tail))


def section_trade_family(b, cfg, data, source, tmp_dir):
    family = cfg.get("family_title", "the product family")
    years = data.years
    rev = data.review_year

    b.add_heading("TRADE IN %s" % family.upper(), level=1)
    for par in cfg.get("intro", []):
        b.add_para(par)
    line = family_members_line(data)
    if line:
        b.add_para(line)

    members = data.members
    if not members:
        return
    b._next_table("Trend on %s: Kenya's Exports by Product, %d" % (family, rev),
                  source)
    members_tbl = top_rows(members, cfg.get("top_n", 10), years,
                           "All other products")
    b.add_value_table("Product", members_tbl, years, "Share in %d" % rev,
                      "Kenya's Exports of %s by Product" % family, source,
                      total_label="Total")

    lead, follows = ([], [])
    pairs = _shares(members, years)
    pairs.sort(key=lambda p: p[1], reverse=True)
    if pairs:
        lead = pairs[0]
        follows = pairs[1:3]
    if lead:
        txt = ("The leading Kenyan export product of %s in %d was %s, which "
               "accounted for %.1f%% of Kenya's exports of the family."
               % (family.lower(), rev, short_label(lead[0], 70), lead[1] * 100))
        if follows:
            txt += " It was followed by %s." % ordinal_list(
                ["%s (%.1f%%)" % (short_label(l, 50), s * 100)
                 for l, s in follows])
        b.add_bullet(txt)

    totals = _year_totals(members)
    g = growth_phrase(cagr([totals.get(y) for y in years], years),
                      period_phrase(years[0], rev))
    yo = yoy_phrase(yoy_change([totals.get(y) for y in years], years),
                    years[-2], years[-1])
    if g:
        sentence = "Kenya's total exports of %s %s." % (family.lower(), g)
        if yo:
            sentence = sentence[:-1] + ", and %s." % yo
        b.add_bullet(sentence)

    if len(members) >= 2:
        img = make_donut(pairs, tmp_dir, "f1_share.png",
                         "Share of %s" % family)
        if img:
            b._next_figure("Share of Kenya's Exports of %s by Product, %d"
                           % (family, rev), source)
            b.add_figure(img)


def section_kenya_exports(b, cfg, data, source, tmp_dir):
    anchor = short_anchor(data.anchor_label)
    years = data.years
    rev = data.review_year

    b.add_heading("KENYA'S EXPORTS OF %s TO THE WORLD" % anchor.upper(), level=1)

    destinations = top_rows(data.destinations(), cfg.get("top_n", 10), years,
                            "All other markets")
    if not destinations:
        return
    b._next_table("Destination Markets for Kenya's %s, %d" % (anchor, rev),
                  source)
    b.add_value_table("Destination market", destinations, years,
                      "Share in %d" % rev,
                      "Kenya's Exports of %s by Destination" % anchor, source,
                      total_label="Total")

    totals = _year_totals(destinations)
    total_last = totals.get(rev)
    g = growth_phrase(cagr([totals.get(y) for y in years], years),
                      period_phrase(years[0], rev))
    yo = yoy_phrase(yoy_change([totals.get(y) for y in years], years),
                    years[-2], years[-1])

    parts = []
    if total_last:
        parts.append("Kenya's exports of %s were valued at %s in %d."
                     % (anchor.lower(), usd_phrase(total_last), rev))
    if g:
        parts.append("They %s." % g)
    if yo:
        parts.append("They %s." % yo)
    for p_ in parts:
        b.add_bullet(p_)

    first = destinations[0] if destinations else None
    if first and first["years"].get(rev):
        share = (first["years"].get(rev) or 0.0) / (total_last or 1.0) * 100
        b.add_bullet("The leading destination for Kenya's %s exports was %s "
                     "(%s; %.1f%% of Kenya's exports of the product)."
                     % (anchor.lower(), first["label"],
                        usd_phrase(first["years"].get(rev)), share))

    pairs = _shares(destinations, years)
    if len(pairs) >= 2:
        img = make_donut(pairs, tmp_dir, "f2_dest.png",
                         "Kenya's %s Exports by Destination" % anchor)
        if img:
            b._next_figure("Kenya's %s Exports by Destination, %d"
                           % (anchor, rev), source)
            b.add_figure(img)


def section_global(b, cfg, data, source, tmp_dir):
    anchor = short_anchor(data.anchor_label)
    family = cfg.get("family_title", "the product family")
    years = data.years
    rev = data.review_year

    b.add_heading("EXPORT OF %s GLOBALLY" % anchor.upper(), level=1)

    exporters = top_rows(data.exporters(), cfg.get("top_n", 10), years,
                         "All other economies")
    if exporters:
        b._next_table("World Exports of %s by Economy, %d" % (anchor, rev),
                      source)
        b.add_value_table("Exporting economy", exporters, years,
                          "Share in %d" % rev,
                          "Countries Exporting %s" % anchor, source,
                          total_label="Total")
        geo_bullets(b, exporters, years, anchor, "exporter")
        pairs = _shares(exporters, years)
        if len(pairs) >= 2:
            img = make_donut(pairs, tmp_dir, "f3_exporters.png",
                             "Share of World Exports of %s" % anchor)
            if img:
                b._next_figure("World Exports of %s by Economy, %d"
                               % (anchor, rev), source)
                b.add_figure(img)

    importers = top_rows(data.importers(), cfg.get("top_n", 10), years,
                         "All other economies")
    if importers:
        b._next_table("World Imports of %s by Economy, %d" % (anchor, rev),
                      source)
        b.add_value_table("Importing economy", importers, years,
                          "Share in %d" % rev,
                          "Countries Importing %s" % anchor, source,
                          total_label="Total")
        geo_bullets(b, importers, years, anchor, "importer")
        pairs = _shares(importers, years)
        if len(pairs) >= 2:
            img = make_donut(pairs, tmp_dir, "f4_importers.png",
                             "Share of World Imports of %s" % anchor)
            if img:
                b._next_figure("World Imports of %s by Economy, %d"
                               % (anchor, rev), source)
                b.add_figure(img)

    g_exp = top_rows(data.global_export_products(),
                     cfg.get("top_n", 10), years, "All other products")
    if g_exp:
        b._next_table("Trend on %s Globally - Export, %d" % (family, rev),
                      source)
        b.add_value_table("Product", g_exp, years, "Share in %d" % rev,
                          "Global Exports of %s by Product" % family, source,
                          total_label="Total")
        trend_bullets(b, g_exp, years, family, "exports")

    g_imp = top_rows(data.global_import_products(),
                     cfg.get("top_n", 10), years, "All other products")
    if g_imp:
        b._next_table("Trend on %s Globally - Import, %d" % (family, rev),
                      source)
        b.add_value_table("Product", g_imp, years, "Share in %d" % rev,
                          "Global Imports of %s by Product" % family, source,
                          total_label="Total")
        trend_bullets(b, g_imp, years, family, "imports")


RESIDUE = {"exporter": "All other economies",
           "importer": "All other economies",
           "source": "All other sources"}


def geo_bullets(b, rows, years, anchor_short, role):
    """'Brazil was the world's leading exporter of X in 2023...' bullets."""
    last = years[-1]
    pairs = _shares(rows, years)
    pairs.sort(key=lambda p: p[1], reverse=True)
    if not pairs:
        return
    group = {"exporter": "exporters", "importer": "importers",
             "source": "sources"}[role]
    word = {"exporter": "exporter", "importer": "importer",
            "source": "source"}[role]
    residue = RESIDUE.get(role)
    residual_share = 0.0
    if residue:
        for l, s in pairs:
            if l == residue:
                residual_share = s
                break
    real = [p for p in pairs if p[0] != residue]
    if not real:
        return
    if role == "source":
        b.add_bullet("The leading source of Kenya's imports of %s in %d was "
                     "%s (%s; %.1f%% of Kenya's imports)."
                     % (anchor_short.lower(), last, real[0][0],
                        usd_phrase(next((r["years"].get(last) for r in rows
                                         if r["label"] == real[0][0]), None)),
                        real[0][1] * 100))
        denom = "Kenya's imports in %d" % last
    else:
        b.add_bullet("%s was the world's leading %s of %s in %d, with %s "
                     "(%.1f%% of the world total)."
                     % (real[0][0], word, anchor_short.lower(), last,
                        usd_phrase(next((r["years"].get(last) for r in rows
                                         if r["label"] == real[0][0]), None)),
                        real[0][1] * 100))
        denom = "the world total in %d" % last
    names = ["%s (%s; %.1f%%)" % (l, usd_phrase(next(
        (r["years"].get(last) for r in rows if r["label"] == l), None)), s * 100)
        for l, s in real[:5]]
    b.add_bullet("The top five %s were %s." % (group, ordinal_list(names)))
    top5 = sum(s for _, s in real[:5]) * 100
    b.add_bullet("Together, the top five %s accounted for %.1f%% of %s."
                 % (group, top5, denom))
    if residual_share > 0:
        b.add_bullet("%s together accounted for %.1f%% of %s."
                     % (residue, residual_share * 100, denom))


def trend_bullets(b, rows, years, family, noun, scope="World",
                  residual="All other products"):
    last = years[-1]
    pairs = _shares(rows, years)
    pairs.sort(key=lambda p: p[1], reverse=True)
    if not pairs:
        return
    real = [p for p in pairs if p[0] != residual]
    if not real:
        return
    subj = ("%s %s" % (scope, noun)).strip()
    denom = "Kenya's imports" if scope == "Kenya's" else "the world total"
    lead, share = real[0]
    txt = ("%s of %s in %d were led by %s (%.1f%% of %s)"
           % (subj, family.lower(), last, short_label(lead, 60), share * 100,
              denom))
    follows = real[1:3]
    if follows:
        txt += ", followed by %s" % ordinal_list(
            ["%s (%.1f%%)" % (short_label(l, 50), s * 100)
             for l, s in follows])
    b.add_bullet(txt + ".")
    lead_share = sum(s for _, s in real)
    b.add_bullet("Together, the leading product headings accounted for %.1f%% "
                 "of %s." % (lead_share * 100, denom))
    totals = _year_totals(rows)
    g = growth_phrase(cagr([totals.get(y) for y in years], years),
                      period_phrase(years[0], last))
    yo = yoy_phrase(yoy_change([totals.get(y) for y in years], years),
                    years[-2], years[-1])
    if g:
        sentence = "%s of %s %s." % (subj, family.lower(), g)
        if yo:
            sentence = sentence[:-1] + ", and %s." % yo
        b.add_bullet(sentence)


def section_kenya_imports(b, cfg, data, source):
    anchor = short_anchor(data.anchor_label)
    family = cfg.get("family_title", "the product family")
    years = data.years
    rev = data.review_year

    sources = top_rows(data.kenya_import_sources(), cfg.get("top_n", 10),
                       years, "All other sources")
    products = data.kenya_import_products()
    if not sources and not products:
        return
    b.add_heading("KENYA'S IMPORTS OF %s" % anchor.upper(), level=1)
    b.add_para("For completeness, the profile also covers Kenya's imports of "
               "the family, by source market and by product, over the same "
               "review period.")

    if sources:
        b._next_table("Kenya's Imports of %s by Source, %d" % (anchor, rev),
                      source)
        b.add_value_table("Source market", sources, years, "Share in %d" % rev,
                          "Kenya's Imports of %s by Source" % anchor, source,
                          total_label="Total")
        geo_bullets(b, sources, years, anchor, "source")

    if products:
        products = top_rows(products, cfg.get("top_n", 10), years,
                            "All other products")
        b._next_table("Kenya's Imports of %s by Product, %d" % (family, rev),
                      source)
        b.add_value_table("Product", products, years, "Share in %d" % rev,
                          "Kenya's Imports of %s by Product" % family, source,
                          total_label="Total")
        trend_bullets(b, products, years, family, "imports", scope="Kenya's")


def section_potential(b, cfg, data, pot, source):
    """Optional export-potential section (ITC Export Potential Map file)."""
    records = pot.get("records") or []
    years = pot.get("years") or []
    if not records or not years:
        return
    markets = {}
    for r in records:
        if r["partner"] == "000":
            continue
        markets.setdefault(r["partner_label"], r["years"])
    last = years[-1]
    market_labels = sorted(markets,
                           key=lambda m: markets[m].get(last) or 0.0,
                           reverse=True)[:10]
    if not market_labels:
        return
    b.add_heading("EXPORT POTENTIAL FOR %s"
                  % cfg.get("family_title", "the product").upper(), level=1)
    b.add_para("The markets below show the largest gap between Kenya's "
               "current exports and the potential demand estimated by the "
               "ITC Export Potential Map.")
    table = b.doc.add_table(rows=1 + len(market_labels), cols=2,
                            style="Table Grid")
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    hdr = table.rows[0]
    hdr.cells[0].text = "Market"
    hdr.cells[1].text = "Export potential (%s)" % cfg.get("unit_label", "USD Million")
    for i, m in enumerate(market_labels, start=1):
        table.rows[i].cells[0].text = m
        table.rows[i].cells[1].text = fmt(markets[m].get(last) or 0.0)
    b._set_table_widths(table, [4200, 3000])
    b._style_table(table, rank=False, label_cols=1, n=1)
    b._fit_table_on_page(table)
    b.add_bullet("Together, these ten markets offer the greatest headroom for "
                 "additional exports of %s." % cfg.get("family_title", ""))


# --------------------------------------------------------------------------
# Top-level build
# --------------------------------------------------------------------------
def build_profile_document(cfg, data, tmp_dir):
    os.makedirs(tmp_dir, exist_ok=True)
    source = cfg.get("source", "International Trade Centre Database")
    b = ProfileBuilder(cfg, {})
    b.title_page(cfg)
    section_trade_family(b, cfg, data, source, tmp_dir)
    section_kenya_exports(b, cfg, data, source, tmp_dir)
    section_global(b, cfg, data, source, tmp_dir)
    if cfg.get("include_imports", True):
        section_kenya_imports(b, cfg, data, source)
    pot = data.files.get("export_potential")
    if pot:
        section_potential(b, cfg, data, pot, source)
    return b.doc


# --------------------------------------------------------------------------
# Excel deliverable
# --------------------------------------------------------------------------
def _xc(ws, r, c, value, bold=False, number_format=None, fill=None,
        align=None):
    cell = ws.cell(r, c, value)
    cell.font = Font(name="Century Gothic", bold=bold)
    cell.border = BORDER
    if number_format:
        cell.number_format = number_format
    if fill:
        cell.fill = fill
    if align:
        cell.alignment = align
    return cell


def write_excel_deliverable(cfg, data, out_path):
    """Companion workbook: same tables + editable doughnut charts."""
    wb = Workbook()
    years = data.years
    rev = years[-1]
    hdr_fill = PatternFill("solid", fgColor=NAVY)
    cm = Alignment(horizontal="center")
    lm = Alignment(horizontal="left")
    val_fmt = "#,##0.0"

    def sheet_name(prefix, label="", cap=31):
        name = (prefix + " - " + label) if label else prefix
        return name[:31]

    def value_sheet(name, first_col, rows, doughnut=None):
        ws = wb.create_sheet(name)
        _xc(ws, 1, 1, first_col, bold=True, fill=hdr_fill, align=cm)
        for i, y in enumerate(years):
            _xc(ws, 1, 2 + i, y, bold=True, fill=hdr_fill, align=cm)
        _xc(ws, 1, 2 + len(years), "Share in %d" % rev, bold=True,
            fill=hdr_fill, align=cm)
        rev_total = sum(display(r["years"].get(rev)) or 0.0 for r in rows)
        for ri, row in enumerate(rows, start=2):
            _xc(ws, ri, 1, row["label"], align=lm)
            for i, y in enumerate(years):
                v = display(row["years"].get(y))
                c = _xc(ws, ri, 2 + i,
                        None if v is None else round(v, 1),
                        number_format=val_fmt, align=cm)
            v = display(row["years"].get(rev)) or 0.0
            _xc(ws, ri, 2 + len(years),
                (v / rev_total if rev_total else None),
                bold=True, number_format="0.0%", align=cm)
        width = max([30] + [len(str(r["label"])) for r in rows])
        ws.column_dimensions["A"].width = min(width, 60)
        for i in range(len(years)):
            ws.column_dimensions[chr(ord("B") + i)].width = 11
        if doughnut:
            anchor = ws.cell(2 + len(rows), 1).row + 2
            _add_doughnut_here(ws, anchor, doughnut[0], doughnut[1])
        return ws

    def _add_doughnut_here(ws, top_row, title, labels_values):
        from openpyxl.chart import DoughnutChart
        from openpyxl.chart.label import DataLabelList
        from openpyxl.chart.legend import Legend
        from openpyxl.chart.series import DataPoint
        from openpyxl.chart.reference import Reference
        ws.cell(top_row, 1).value = "Category"
        ws.cell(top_row, 2).value = "Share"
        for cc in (1, 2):
            c = ws.cell(top_row, cc)
            c.font = Font(bold=True, color="FFFFFF")
            c.fill = hdr_fill
            c.alignment = cm
        for i, (lab, v) in enumerate(labels_values, start=top_row + 1):
            ws.cell(i, 1).value = lab
            cv = ws.cell(i, 2)
            cv.value = v
            cv.number_format = "0.0%"
            cv.border = BORDER
            cv.alignment = cm
        last = top_row + len(labels_values)
        data = Reference(ws, min_col=2, min_row=top_row + 1, max_row=last)
        cats = Reference(ws, min_col=1, min_row=top_row + 1, max_row=last)
        chart = DoughnutChart()
        chart.title = title
        chart.width = 14
        chart.height = 9
        chart.dataLabels = DataLabelList()
        chart.dataLabels.showPercent = True
        chart.dataLabels.numFmt = "0.0%"
        chart.dataLabels.dLblPos = "outEnd"
        chart.dataLabels.showLeaderLines = True
        chart.add_data(data, titles_from_data=False)
        chart.set_categories(cats)
        for i in range(len(labels_values)):
            dp = DataPoint(idx=i)
            dp.graphicalProperties.solidFill = THEME[i % len(THEME)]
            chart.series[0].data_points.append(dp)
        chart.legend = Legend()
        chart.legend.position = "r"
        chart.legend.overlay = False
        ws.add_chart(chart, "C%d" % top_row)

    # ---- sheets -----------------------------------------------------------
    if data.members:
        members = top_rows(data.members, cfg.get("top_n", 10), years,
                           "All other products")
        value_sheet("Kenya Exports by Product", "Product", members,
                    doughnut=("Share of Kenya's %s by Product"
                              % cfg.get("family_title"),
                              [(m["label"], _share01(m, rev, members))
                               for m in members]))
    destinations = top_rows(data.destinations(), cfg.get("top_n", 10), years,
                            "All other markets")
    if destinations:
        value_sheet(sheet_name("Destinations", data.anchor_label),
                    "Destination market", destinations,
                    doughnut=("Kenya's exports by destination",
                              [(d["label"], _share01(d, rev, destinations))
                               for d in destinations]))
    exporters = top_rows(data.exporters(), cfg.get("top_n", 10), years,
                         "All other economies")
    importers = top_rows(data.importers(), cfg.get("top_n", 10), years,
                         "All other economies")
    if exporters:
        value_sheet(sheet_name("World Exporters"), "Exporting economy", exporters)
    if importers:
        value_sheet(sheet_name("World Importers"), "Importing economy", importers)
    g_exp = top_rows(data.global_export_products(), cfg.get("top_n", 10),
                     years, "All other products")
    g_imp = top_rows(data.global_import_products(), cfg.get("top_n", 10),
                     years, "All other products")
    if g_exp:
        value_sheet("Global Exports by Product", "Product", g_exp)
    if g_imp:
        value_sheet("Global Imports by Product", "Product", g_imp)
    sources = top_rows(data.kenya_import_sources(), cfg.get("top_n", 10),
                       years, "All other sources")
    if sources:
        value_sheet("Kenya Imports by Source", "Source market", sources)
    prods = top_rows(data.kenya_import_products(), cfg.get("top_n", 10),
                     years, "All other products")
    if prods:
        value_sheet("Kenya Imports by Product", "Product", prods)

    # default "Sheet" removed by the first create_sheet call? keep membership
    if "Sheet" in wb.sheetnames:
        del wb["Sheet"]
    wb.save(out_path)
    return wb


def _share01(row, rev, rows):
    total = sum(display(r["years"].get(rev)) or 0.0 for r in rows)
    v = display(row["years"].get(rev)) or 0.0
    return (v / total) if total else None


def main():
    ap = argparse.ArgumentParser(
        description="Generate a general product-profile report.")
    ap.add_argument("--config",
                    default=os.path.join(BASE_DIR, "config",
                                         "product_profile_coffee.json"))
    ap.add_argument("--excel-dir", default=None,
                    help="Folder of ITC matrix files (default: config data_dir)")
    ap.add_argument("--output", default=None, help="Output .docx path")
    ap.add_argument("--tmp",
                    default=os.path.join(BASE_DIR, "output", ".tmp"),
                    help="Temporary directory for chart images")
    args = ap.parse_args()

    cfg = json.load(open(args.config, encoding="utf-8"))
    data_dir = args.excel_dir or os.path.join(
        BASE_DIR, cfg.get("data_dir", "Coffee"))
    out = args.output
    if not out:
        out = os.path.join(BASE_DIR, cfg.get("output",
                                             "output/Product Profile.docx"))
    out = os.path.abspath(out)

    data = ProfileData(data_dir)
    if data.anchor_is_total:
        data.anchor_label = cfg.get("family_title", "")
    print("[1/4] Loading ITC files from      : %s" % data_dir)
    print("      family    = %s" % cfg.get("family_title"))
    print("      anchor    = %s (%s)" % (data.anchor_hs, data.anchor_label))
    print("      period    = %d - %d" % (data.start_year, data.review_year))
    print("      members   = %s (product detail rows)"
          % len(data.members))

    print("[2/4] Building report             : %s" % out)
    doc = build_profile_document(cfg, data, args.tmp)
    doc.save(out)
    print("[3/4] Report saved to             : %s" % out)

    try:
        t = os.path.splitext(out)[0] + " TABLES.xlsx"
        os.makedirs(os.path.dirname(t), exist_ok=True)
        write_excel_deliverable(cfg, data, t)
        print("[4/4] Excel deliverable saved to : %s" % t)
    except Exception as e:
        print("      [warn] Excel deliverable skipped: %s" % e)


if __name__ == "__main__":
    main()