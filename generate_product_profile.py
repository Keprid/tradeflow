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
from openpyxl.utils.cell import get_column_letter

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
from country_names import display_name, fix_label, is_africa
import charts
from excel_deliverable import THEME
import xlsx_compat

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
    if not hits:
        hits = [p for p in glob.glob(os.path.join(data_dir, "*")) if
                prefix in os.path.basename(p).lower()]
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


def load_html_matrix(path):
    """Parse an ITC HTML ``.xls`` matrix (label columns + year columns).

    These files (e.g. ``List of products exported by Kenya``, or the
    product-group listings) are HTML tables saved under an ``.xls`` name.
    Returns ``(years, records)`` with ``{"code", "label", "years"}`` rows.
    Returns ``(None, None)`` when the file has no usable year columns.
    """
    for loader in ("read_text", "csv", "html"):
        try:
            if loader == "read_text":
                with open(path, "r", encoding="utf-8", errors="replace") as fh:
                    text = fh.read()
            elif loader == "html":
                text = xlsx_compat._read_text_any(path)
            else:
                continue
            if not text:
                continue
            parser = xlsx_compat.HTMLTableParser()
            parser.feed(text)
            if not parser.tables:
                return None, None
            table = max(parser.tables, key=len)
            break
        except Exception:
            continue
    else:
        return None, None
    rows = [["" if c is None else str(c).strip() for c in row]
            for row in table if row]
    header = next((r for r in rows
                   if any(re.search(r"20\d\d", c) for c in r)), None)
    if header is None:
        return None, None
    years, year_cols = [], []
    for i, h in enumerate(header):
        m = re.search(r"(20\d\d)", h)
        if m:
            y = int(m.group(1))
            if 2000 <= y <= 2100:
                years.append(y)
                year_cols.append(i)
    if not years:
        return None, None
    start = rows.index(header) + 1
    records = []
    for r in rows[start:]:
        if not r or not any(c for c in r):
            continue
        code = r[0].lstrip("'").strip() if r[0] else ""
        label = next((c for c in r[1:4] if c and not re.match(r"^20\d\d", c)),
                     "") if len(r) > 1 else ""
        if "Product label" in label or "Code" in label.lower():
            continue
        val = {y: to_float(r[c]) if c < len(r) else None
               for y, c in zip(years, year_cols)}
        if not code and not label:
            continue
        if not any(v for v in val.values()) and not code:
            continue
        records.append({"code": code, "label": label, "years": val})
    return years, records


class ProfileData:
    """Loads every ITC file in the folder and derives the analysis tables."""

    def __init__(self, data_dir, include_codes=None, family_title=None):
        self.data_dir = os.path.abspath(data_dir)
        self.include_codes = list(include_codes or [])
        self.family_title = family_title
        self.warnings = []
        self.files = {}
        for prefix, key in FILE_PREFIXES.items():
            path = _find_file(self.data_dir, prefix)
            if path is None:
                if key not in ("export_potential", "kenya_exports_by_partner"):
                    self.warnings.append(
                        "missing %s file(s) in %s - related sections skipped"
                        % (prefix, self.data_dir))
                continue
            years, records = load_matrix(path)
            self.files[key] = {"path": path, "years": years,
                               "records": records}

        # ---- supplementary ITC HTML .xls downloads (optional) -------------
        # Kenya's total merchandise exports (RCA / specialization) and the
        # world / Kenya product-group breakout come from the "List of
        # products..." HTML workbooks.  Each is optional: if the file is not
        # present the related section simply uses whatever is available.
        self.kenya_total_exports = {}          # {year: USD thous}
        self.supplement = {}                   # lowercased stem -> {"years","records"}
        for stem, key in (("list_of_products_exported_by_kenya",
                           "kenya_total_exports"),
                          ("list_of_exported_products_for_the_selected",
                           "world_group_products")):
            path = _find_file(self.data_dir, stem)
            if path is None:
                continue
            years, records = load_html_matrix(path)
            if records is None:
                continue
            self.supplement[key] = {"years": years, "records": records}
        if "kenya_total_exports" in self.supplement:
            for r in self.supplement["kenya_total_exports"]["records"]:
                code = (r.get("code") or "").upper()
                label = (r.get("label") or "").lower()
                if code.startswith("TOTAL") or "all products" in label:
                    self.kenya_total_exports = dict(r["years"])
                    break

        # Projection years carried only by the export-potential file must not
        # shift the traded-review window used by the rest of the profile.
        self.all_years = sorted({y for k, f in self.files.items()
                                 if k != "export_potential"
                                 for y in f["years"]})

        # Anchor product (the product of the by-importer download)
        partner_rows = self._rows("kenya_exports_by_partner")
        anchor_product = None
        if partner_rows:
            counts = {}
            for r in partner_rows:
                counts[r["product"]] = counts.get(r["product"], 0) + 1
            anchor_product = max(counts, key=counts.get)
        if anchor_product is None:
            raise OSError("Missing the by-importer file; cannot detect the "
                          "anchor product in %s." % self.data_dir)
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
        # (its total row, if any, is excluded), filtered by include_codes.
        total = self._file_totals.get("kenya_exports_by_product", set())
        members = []
        for r in self._rows("kenya_exports_by_product"):
            if r["partner"] != "000" or r["product"] in total:
                continue
            if not self._code_ok(r["product"]):
                continue
            members.append({"code": r["product"], "label": r["product_label"],
                            "years": r["years"]})
        members.sort(key=lambda m: m["years"].get(self.review_year) or 0.0,
                     reverse=True)
        self.members = members
        self.anchor_is_total = anchor_product in total
        if self.anchor_is_total and self.family_title:
            # The anchor is the ITC selection group (e.g. "coffee one"), whose
            # group name is not meaningful on its own.  Use the configured
            # family title ("Coffee") for headings and tables instead.
            self.anchor_label = self.family_title

    def _code_ok(self, code):
        """True when ``code`` matches the configured include_codes.

        HS codes are hierarchical: a 6-digit code belongs to its 4-digit
        heading and its 2-digit chapter, so the check is prefix-based.

        Each entry in ``include_codes`` is either
          * a plain code at any digit depth (``"0901"``, ``"4420.10"``), which
            matches that heading together with every code under it, or
          * a closed range of siblings (``"4419.11-4419.90"``), which matches
            every code numerically in [lo, hi] plus any heading that covers
            a part of that range.

        Codes are normalised to digits (dots/spaces removed), so ``"4420.10"``
        and ``"442010"`` are equivalent.  This lets the config express a whole
        product group (e.g. commercial crafts spanning many HS chapters) by
        listing its headings; all revised HS 2022 codes beneath them are
        picked up automatically.
        """
        if not self.include_codes:
            return True
        c = self._norm_code(code)
        if not c:
            return False
        for spec in self.include_codes:
            spec = str(spec).strip()
            if "-" in spec:
                lo, hi = spec.split("-", 1)
                lo, hi = self._norm_code(lo), self._norm_code(hi)
                if not (lo and hi):
                    continue
                if lo <= c <= hi:
                    return True
                # A 4-digit heading that contains an endpoint is inside the
                # range (e.g. "5701" with "5701.10-5702.99").  Chapter-level
                # (2-digit) codes never match a heading range, so they can't
                # drag whole chapters into a group that spans only some of it.
                if len(c) >= 4 and (lo.startswith(c) or hi.startswith(c)):
                    return True
            else:
                p = self._norm_code(spec)
                if not p:
                    continue
                if c == p or c.startswith(p):
                    return True
                # A deeper spec (e.g. "4420.10") also covers its 4-digit
                # parent heading (e.g. "4420"), but never the 2-digit chapter.
                if len(c) >= 4 and p.startswith(c):
                    return True
        return False

    @staticmethod
    def _norm_code(code):
        """Normalise an HS code to its digits (e.g. "4420.10" -> "442010")."""
        return re.sub(r"[^0-9]", "", str(code or ""))

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
        return [{"label": fix_label(r["partner_label"]), "years": r["years"]}
                for r in rows]

    def exporters(self):
        rows = sorted((r for r in self._rows("world_exports_by_economy")
                       if r["reporter"] != "000"),
                      key=lambda r: r["years"].get(self.review_year) or 0.0,
                      reverse=True)
        return [{"label": fix_label(r["reporter_label"]), "years": r["years"]}
                for r in rows]

    def importers(self):
        rows = sorted((r for r in self._rows("world_imports_by_economy")
                       if r["reporter"] != "000"),
                      key=lambda r: r["years"].get(self.review_year) or 0.0,
                      reverse=True)
        return [{"label": fix_label(r["reporter_label"]), "years": r["years"]}
                for r in rows]

    def kenya_import_sources(self):
        rows = sorted((r for r in self._rows("kenya_imports_by_partner")
                       if r["partner"] != "000"),
                      key=lambda r: r["years"].get(self.review_year) or 0.0,
                      reverse=True)
        return [{"label": fix_label(r["partner_label"]), "years": r["years"]}
                for r in rows]

    def kenya_import_products(self):
        total = self._file_totals.get("kenya_imports_by_product", set())
        rows = sorted((r for r in self._rows("kenya_imports_by_product")
                       if r["partner"] == "000" and r["product"] not in total
                       and self._code_ok(r["product"])),
                      key=lambda r: r["years"].get(self.review_year) or 0.0,
                      reverse=True)
        return [{"code": r["product"], "label": r["product_label"],
                 "years": r["years"]} for r in rows]

    def global_export_products(self):
        total = self._file_totals.get("world_exports_by_product", set())
        rows = sorted((r for r in self._rows("world_exports_by_product")
                       if r["partner"] == "000" and r["product"] not in total
                       and self._code_ok(r["product"])),
                      key=lambda r: r["years"].get(self.review_year) or 0.0,
                      reverse=True)
        return [{"code": r["product"], "label": r["product_label"],
                 "years": r["years"]} for r in rows]

    def global_import_products(self):
        total = self._file_totals.get("world_imports_by_product", set())
        rows = sorted((r for r in self._rows("world_imports_by_product")
                       if r["partner"] == "000" and r["product"] not in total
                       and self._code_ok(r["product"])),
                      key=lambda r: r["years"].get(self.review_year) or 0.0,
                      reverse=True)
        return [{"code": r["product"], "label": r["product_label"],
                 "years": r["years"]} for r in rows]

    @staticmethod
    def _is_kenya_label(label):
        text = " ".join(str(label or "").lower().split())
        return text == "kenya" or display_name(text).lower() == "kenya"

    def kenya_exporters(self):
        """Export rows of ``world_exports_by_economy`` that reference Kenya."""
        return [r for r in self._rows("world_exports_by_economy")
                if r["reporter"] != "000" and self._is_kenya_label(
                    r["reporter_label"])]

    def kenya_global_metrics(self):
        """Kenya's standing in the world exports of the family.

        Returns a dict with Kenya's export value, world total, Kenya's share
        of the world total and Kenya's rank among *all* exporters and among
        *African* exporters, in the review year:
        ``{"value", "world_total", "share", "global_rank", "n_exporters",
           "africa_rank", "n_africa"}`` (ranks are 1-based, ties keep the
        highest rank).  ``None`` when Kenya is not present in the file.
        """
        rows = [r for r in self._rows("world_exports_by_economy")
                if r["reporter"] != "000"]
        rev = self.review_year
        if not rows:
            return None
        key = lambda r: r["years"].get(rev) or 0.0  # noqa: E731
        rows.sort(key=key, reverse=True)
        world_total = sum(key(r) for r in rows)
        ranked = [r for r in rows if key(r) > 0]
        kenya = next((r for r in rows if self._is_kenya_label(
            r["reporter_label"])), None)
        if kenya is None or not world_total:
            return None
        value = key(kenya)
        global_rank = next((i for i, r in enumerate(ranked, 1)
                            if self._is_kenya_label(r["reporter_label"])), None)
        # The African standing uses the latest available export value (review
        # year, or the most recent year with data) so economies whose latest
        # ITC figure predates the review year - e.g. Ethiopia - still rank.
        africa = [r for r in rows
                  if is_africa(r["reporter_label"]) and _latest_value(r, rev) > 0]
        africa.sort(key=lambda r: _latest_value(r, rev), reverse=True)
        africa_rank = next((i for i, r in enumerate(africa, 1)
                            if self._is_kenya_label(r["reporter_label"])), None)
        return {
            "value": value,
            "world_total": world_total,
            "share": value / world_total if world_total else None,
            "global_rank": global_rank,
            "n_exporters": len(ranked),
            "africa_rank": africa_rank,
            "n_africa": len(africa),
        }

    def market_share_series(self):
        """Kenya's share of world exports of the family, year by year.

        ``None`` if the world-exports-by-economy file is missing. Returns
        a list of ``(year, kenya, world, share)`` tuples for the years where
        both values are available.
        """
        rows = [r for r in self._rows("world_exports_by_economy")
                if r["reporter"] != "000"]
        if not rows:
            return None
        kenya = next((r for r in rows if self._is_kenya_label(
            r["reporter_label"])), None)
        if kenya is None:
            return None
        years = sorted({y for r in rows for y in r["years"]})
        world_by_year = {}
        for r in rows:
            for y, v in r["years"].items():
                world_by_year[y] = (world_by_year.get(y) or 0.0) + (v or 0.0)
        out = []
        for y in sorted(years):
            w = world_by_year.get(y) or 0.0
            k = kenya["years"].get(y) or 0.0
            if w and k:
                out.append((y, k, w, k / w))
        return out or None

    def specialization_metrics(self):
        """Export specialization for the family (RCA-style).

        Computes the family's share of Kenya's *total* merchandise exports
        each year from the optional Kenya "List of products exported" total
        row.  ``None`` when that total is unavailable.

        Returns ``{"years": [(year, family_value, kenya_total,
                             family_share)], "review_family": ...}``.
        The full RCA ratio needs world total exports which are not part of
        the standard downloads, so the share of Kenya's own exports is used
        as the specialization measure (the ratio would require the world
        total-exports denominator).
        """
        if not self.kenya_total_exports:
            return None
        rows = [r for r in self._rows("world_exports_by_economy")
                if self._is_kenya_label(r["reporter_label"])]
        if not rows:
            return None
        kenya = rows[0]
        rev = self.review_year
        years = sorted(self.kenya_total_exports)
        out = []
        for y in years:
            total = self.kenya_total_exports.get(y)
            fam = kenya["years"].get(y)
            if total and fam is not None:
                out.append({"year": y, "family": fam, "total": total,
                            "share": fam / total})
        if not out:
            return None
        return {"years": out,
                "review_family": next((r["family"] for r in out
                                       if r["year"] == rev), None),
                "review_total": next((r["total"] for r in out
                                      if r["year"] == rev), None)}

    def african_peers(self, n=5):
        """``n`` leading African exporters of the family plus Kenya, ranked
        by export value - largest to smallest (review year, falling back to
        the most recent year with data, so e.g. Ethiopia - whose latest ITC
        figure predates the review year - still ranks on its newest value).
        Returns ``None`` when no African economy other than Kenya has data.
        """
        rows = [r for r in self._rows("world_exports_by_economy")
                if r["reporter"] != "000"]
        rev = self.review_year
        if not rows:
            return None
        kenya = next((r for r in rows
                      if self._is_kenya_label(r["reporter_label"])), None)
        peers = [r for r in rows
                 if is_africa(r["reporter_label"]) and not self._is_kenya_label(
                     r["reporter_label"]) and _latest_value(r, rev) > 0]
        if not peers:
            return None
        combined = ([kenya] if kenya else []) + peers
        seen, out = set(), []
        for r in combined:
            key = fix_label(r["reporter_label"])
            if key in seen:
                continue
            seen.add(key)
            out.append({"label": key, "years": r["years"],
                        "src_year": _latest_year(r, rev)})
        out.sort(key=lambda r: _latest_value(r, rev), reverse=True)
        return out[:n + 1]

    def kenya_world_shares(self):
        """Kenya's exports vs the world by 6-digit HS code.

        Joins Kenya's member rows (``kenya_exports_by_product``) with the
        corresponding world rows (``world_exports_by_product``) on product
        code.  Returns a list (already ranked by Kenya's review-year value) of
        ``{"code", "label", "kenya": {y: v}, "world": {y: v}}`` plus
        ``_kenya_total`` / ``_world_total`` keys holding the review-year
        totals, so callers can compute Kenya's share of the world by code and
        overall.
        """
        kenya_total = self._file_totals.get("kenya_exports_by_product", set())
        world_total = self._file_totals.get("world_exports_by_product", set())
        kenya_rows = {r["product"]: r["years"] for r in
                      self._rows("kenya_exports_by_product")
                      if r["partner"] == "000" and r["product"] not in kenya_total
                      and self._code_ok(r["product"])}
        world_rows = {r["product"]: r["years"] for r in
                      self._rows("world_exports_by_product")
                      if r["partner"] == "000" and r["product"] not in world_total
                      and self._code_ok(r["product"])}
        rev = self.review_year
        out = []
        for code, ky in kenya_rows.items():
            out.append({"code": code,
                        "label": next((r["product_label"] for r in
                                       self._rows("kenya_exports_by_product")
                                       if r["product"] == code), ""),
                        "kenya": ky,
                        "world": world_rows.get(code, {})})
        out.sort(key=lambda r: (r["kenya"].get(rev) or 0.0), reverse=True)
        k_tot = sum(r["kenya"].get(rev) or 0.0 for r in out)
        w_tot = sum(r["world"].get(rev) or 0.0 for r in out)
        return {"rows": out, "_kenya_total": k_tot, "_world_total": w_tot,
                "year": rev}


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


def _series_unit(values):
    """Display unit for a series of raw (USD-thousand) values.

    Returns ``"USD Million"`` or ``"USD Thousand"``.  Thousands is chosen
    whenever any nonzero value would otherwise round to 0.0 in millions
    (i.e. sit below about USD 50,000), so small sub-category values never
    "disappear" in a table."""
    vs = [v for v in values if v]
    if not vs:
        return "USD Million"
    if max(abs(v) for v in vs) < 100.0 or min(abs(v) for v in vs) < 50.0:
        return "USD Thousand"
    return "USD Million"


def fmt_for_unit(v, unit, decimals=1):
    """Format a raw (USD-thousand) value in the selected ``unit``."""
    if v is None:
        return ""
    if unit == "USD Thousand":
        return num(v, decimals)
    return num(display(v), decimals)


def usd_phrase(v):
    """'USD 353.8 Million' style phrasing for narratives.  Amounts below
    USD 100,000 are shown in thousands so a small-but-real value never
    reads as 'USD 0.0 Million' next to its percentage share."""
    if v is None:
        return ""
    d = display(v)
    word = "Million"
    if d >= 1000.0:
        d, word = d / 1000.0, "Billion"
    elif d < 0.1:
        d, word = d * 1000.0, "Thousand"
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


def _latest_value(row, rev):
    """Review-year export value, falling back to the most recent year that
    has data.  Used so economies whose latest figures ITC has not yet
    published (e.g. Ethiopia for 2024-2025) still rank fairly."""
    v = row["years"].get(rev)
    if v:
        return v
    for y in sorted(row["years"], reverse=True):
        if row["years"].get(y):
            return row["years"][y]
    return 0.0


def _latest_year(row, rev):
    """Year behind ``_latest_value`` (the review year, or an earlier one)."""
    if row["years"].get(rev):
        return rev
    for y in sorted(row["years"], reverse=True):
        if row["years"].get(y):
            return y
    return None


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


def _ranked_rows(rows, n, years, ensure_label=None,
                 residual="All other economies"):
    """Like ``top_rows`` but records each row's true overall position and
    keeps ``ensure_label`` (e.g. Kenya) in the list even when it falls
    outside the top ``n`` -- so it is reported with its overall rank, before
    the remainder is consolidated into the residual row.

    ``rows`` must already be sorted descending by review-year value (as the
    ``data.exporters()`` / ``data.importers()`` accessors return).  Rows are
    assigned a shared ``rank`` when their review-year values are tied.
    """
    rows = [dict(r) for r in rows]
    rev = years[-1]
    prev_val, prev_rank = None, 0
    for idx, r in enumerate(rows, 1):
        v = r["years"].get(rev) or 0.0
        rank = prev_rank if (prev_val is not None and v == prev_val) else idx
        prev_val, prev_rank = v, rank
        r["rank"] = rank

    extra = 1 if ensure_label and ensure_label not in \
        {r["label"] for r in rows[:n]} else 0
    if len(rows) <= n + extra:
        return rows

    top = rows[:n]
    if ensure_label and ensure_label not in {r["label"] for r in top}:
        k = next((r for r in rows if r["label"] == ensure_label), None)
        if k is not None:
            top = top + [k]
    residual_years = {}
    for y in years:
        shown = sum((r["years"].get(y) or 0.0) for r in top)
        total = sum((r["years"].get(y) or 0.0) for r in rows)
        residual_years[y] = max(0.0, total - shown)
    return top + [{"label": residual, "years": residual_years}]


def _africa_rank(rows, ensure_label):
    """Position of ``ensure_label`` among African rows (1-based), by
    latest available export value (review year, falling back to the most
    recent year with data).  ``rows`` must be sorted descending by review
    year.  Returns ``None`` when there is no Africa match for
    ``ensure_label``.
    """
    rev = max((y for r in rows for y in r["years"]), default=None)
    africa = [r for r in rows
              if is_africa(r.get("label", "")) and
              _latest_value(r, rev) > 0]
    africa_sorted = sorted(africa,
                           key=lambda r: _latest_value(r, rev),
                           reverse=True)
    for i, r in enumerate(africa_sorted, 1):
        if r.get("label") == ensure_label:
            return i
    return None


def _africa_count(rows):
    """Number of distinct African rows (latest available value > 0)."""
    rev = max((y for r in rows for y in r["years"]), default=None)
    return sum(1 for r in rows
               if is_africa(r.get("label", "")) and
               _latest_value(r, rev) > 0)


def _year_totals(rows):
    years = set()
    for r in rows:
        years |= set(r["years"])
    return {y: sum((r["years"].get(y) or 0.0) for r in rows) for y in years}


# --------------------------------------------------------------------------
# Growth-strategy analytics (decompositions, rankings, share shifts)
# --------------------------------------------------------------------------
def margin_decomposition(rows, start, rev):
    """Split the net change in a series into its constituent margins.

    ``rows`` is a list of ``{"label" or "code", "years": {y: v}}`` (the full,
    un-truncated download).  Existing headings that kept exporting are the
    *intensive* margin (deepening); headings that only appear in the review
    year are the *extensive* margin (entering the basket); headings that
    disappear are exits.  Returns a dict of USD-thousand contributions plus
    their share of the net change (``None`` when the net change is zero).
    """
    tracked = {}
    for r in rows:
        key = r.get("label") or r.get("code")
        if not key:
            continue
        tracked[key] = (r["years"].get(start) or 0.0,
                        r["years"].get(rev) or 0.0)
    deepening, entering, exiting = [], [], []
    for key, (sv, rv) in tracked.items():
        if sv and rv:
            deepening.append((key, sv, rv))
        elif not sv and rv:
            entering.append((key, rv))
        elif sv and not rv:
            exiting.append((key, sv))
    intens_value = sum(rv - sv for _, sv, rv in deepening)
    enter_value = sum(rv for _, rv in entering)
    exits_value = -sum(sv for _, sv in exiting)
    total_start = sum(sv for _, sv, _ in deepening) + sum(sv for _, sv in exiting)
    total_rev = sum(rv for _, _, rv in deepening) + sum(rv for _, rv in entering)
    net = total_rev - total_start

    def share_of(v):
        return (v / net) if net else None

    return {
        "intensive": intens_value, "intensive_share": share_of(intens_value),
        "entering": enter_value, "entering_share": share_of(enter_value),
        "exiting": exits_value, "exiting_share": share_of(exits_value),
        "net": net, "total_start": total_start, "total_rev": total_rev,
        "n_existing": len(deepening), "n_entering": len(entering),
        "n_exiting": len(exiting),
        "top_existing": sorted(deepening, key=lambda t: t[2], reverse=True)[:8],
        "top_entering": sorted(entering, key=lambda t: t[1], reverse=True)[:8],
        "top_exiting": sorted(exiting, key=lambda t: t[1], reverse=True)[:8],
    }


def _potential_markets(pot):
    """Per-market potential-vs-actual from an ITC Export Potential Map file.

    Returns ``(markets, potential_year)`` where ``markets[label]`` is
    ``{"actual", "actual_year", "potential", "potential_year"}`` (USD
    thousands).  The earliest year column carries the actual base; the last
    carries the potential projection.
    """
    records = pot.get("records") or []
    pyears = pot.get("years") or []
    if not records or not pyears:
        return {}, None
    markets = {}
    for r in records:
        if str(r["partner"]) == "000":
            continue
        d = markets.setdefault(r["partner_label"], {})
        for y in pyears:
            v = r["years"].get(y)
            if v is None:
                continue
            d.setdefault(y, v)
    last = pyears[-1]
    out = {}
    for label, d in markets.items():
        if not d:
            continue
        actual_year = min(d)
        out[label] = {"actual": d.get(actual_year),
                      "actual_year": actual_year,
                      "potential": d.get(last),
                      "potential_year": last}
    return out, last


def market_attractiveness(rows, years, pot_by_market=None, weights=None,
                          access=()):
    """Rank destination markets as targets for export expansion.

    Score = weighted blend of growth momentum (CAGR of Kenya's exports to the
    market), headroom (potential gap from the Export Potential Map, when
    available), size (share of Kenya's exports) and market access (explicit
    ``access`` list, or the Africa tier; 0 otherwise).  Markets without a
    potential figure get a neutral headroom score rather than being
    penalised.  Returns records sorted by descending score (highest first).
    """
    rev = years[-1] if years else None
    active = [r for r in rows if (r["years"].get(rev) or 0.0) > 0]
    if not active:
        return []
    total = sum(r["years"].get(rev) or 0.0 for r in active)
    w = {"growth": 0.35, "gap": 0.35, "size": 0.15, "access": 0.15}
    if weights:
        w.update(dict(weights))
    access_set = set(access)
    recs = []
    for r in active:
        g = cagr([r["years"].get(y) for y in years], years)
        rec = {
            "label": r["label"],
            "value_review": r["years"].get(rev) or 0.0,
            "share": (r["years"].get(rev) or 0.0) / total if total else 0.0,
            "growth": g,
            "gap": None,
            "tier": 1 if r["label"] in access_set else
                    (2 if is_africa(r["label"]) else 0),
            "has_potential": False,
        }
        if pot_by_market and r["label"] in pot_by_market:
            p = pot_by_market[r["label"]]
            rec["has_potential"] = True
            if (p.get("potential") or 0.0) > 0:
                rec["gap"] = max(0.0, p["potential"] - (p.get("actual") or 0.0))
        recs.append(rec)

    def bounds(key):
        known = [x[key] for x in recs if x[key] is not None]
        if not known:
            return None
        lo, hi = min(known), max(known)
        return None if lo == hi else (lo, hi)

    gb, sb, gab = bounds("growth"), bounds("share"), bounds("gap")

    def nscore(v, bnd, default=0.5):
        if v is None or not bnd:
            return default
        return (v - bnd[0]) / (bnd[1] - bnd[0])

    for x in recs:
        access_score = 1.0 if x["tier"] == 1 else (0.65 if x["tier"] == 2
                                                   else 0.2)
        x["score"] = (w["growth"] * nscore(x["growth"], gb)
                      + w["gap"] * nscore(x["gap"], gab)
                      + w["size"] * nscore(x["share"], sb)
                      + w["access"] * access_score)
    recs.sort(key=lambda x: x["score"], reverse=True)
    return recs


def share_change(rows, years):
    """Per-economy change in world-export share between ``years[0]`` and
    ``years[-1]`` (share points).  Returns ``{"rows", "world_start",
    "world_rev"}`` sorted by biggest share gain first, or ``None`` when the
    world totals for either endpoint are missing."""
    start, rev = years[0], years[-1]
    world = {}
    for r in rows:
        for y, v in r["years"].items():
            world[y] = (world.get(y) or 0.0) + (v or 0.0)
    w0, w1 = world.get(start), world.get(rev)
    if not w0 or not w1:
        return None
    recs = []
    for r in rows:
        sv = r["years"].get(start) or 0.0
        rv = r["years"].get(rev) or 0.0
        if not sv and not rv:
            continue
        recs.append({"label": r["label"],
                     "start_share": sv / w0, "rev_share": rv / w1,
                     "delta_pp": (rv / w1 - sv / w0) * 100.0})
    recs.sort(key=lambda x: x["delta_pp"], reverse=True)
    return {"rows": recs, "world_start": w0, "world_rev": w1}


# --------------------------------------------------------------------------
# Word builder
# --------------------------------------------------------------------------
def _current_month_year():
    from datetime import date
    return date.today().strftime("%B %Y")


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
                        unit_label="USD Million", widths=None,
                        adaptive_unit=True):
        """Structure a value matrix table.

        ``rows`` = list of ``{"label", "years": {y: v}}`` already truncated
        and consolidated. The share column shows each row's share of the
        total in the review year (bold); a ``total_label`` row (optional)
        sums the displayed rows.

        When *adaptive_unit* is ``True`` (default) the whole table is shown in
        USD Thousand (instead of Millions) if any nonzero value would
        otherwise round to 0.0 in millions, so small sub-category values
        never disappear.
        """
        n = len(years)
        code_cols = 1 if any(row.get("code") for row in rows) else 0
        label_cols = (1 if rank else 0) + code_cols + 1
        cols = label_cols + n + 1
        c_label = (1 if rank else 0) + code_cols
        nrows = 2 + len(rows) + (1 if total_label else 0)
        unit = unit_label
        if adaptive_unit:
            unit = _series_unit([v for r in rows for y in years
                                 if (v := r["years"].get(y)) is not None])
        table = self.doc.add_table(rows=nrows, cols=cols, style="Table Grid")
        table.alignment = WD_TABLE_ALIGNMENT.CENTER

        hdr = table.rows[0]
        if rank:
            hdr.cells[0].text = "#"
        if code_cols:
            hdr.cells[1 if rank else 0].text = "Code"
        hdr.cells[c_label].text = first_col_header
        if n > 1:
            hdr.cells[label_cols].merge(hdr.cells[label_cols + n - 1])
        hdr.cells[label_cols].text = "Value in %s" % unit
        hdr.cells[label_cols + n].text = share_header

        r2 = table.rows[1]
        for c in range(cols):
            r2.cells[c].text = ""
        if rank:
            r2.cells[c_label].text = first_col_header
        for i, y in enumerate(years):
            r2.cells[label_cols + i].text = str(y)

        # shares are relative to the total of the review year across rows
        rev = years[-1]
        rev_total = sum((r["years"].get(rev) or 0.0) for r in rows)
        for ri, row in enumerate(rows, start=2):
            c = 0
            if rank:
                rk = row.get("rank")
                table.rows[ri].cells[0].text = "" if rk is None else str(rk)
                c = 1
            if code_cols:
                table.rows[ri].cells[c].text = str(row.get("code") or "")
                c += 1
            table.rows[ri].cells[c].text = row["label"]
            for i, y in enumerate(years):
                table.rows[ri].cells[c + 1 + i].text = \
                    fmt_for_unit(row["years"].get(y), unit)
            cur = row["years"].get(rev)
            if cur is None:
                share = None
            else:
                share = (cur / rev_total * 100.0) if rev_total else None
            table.rows[ri].cells[c + 1 + n].text = \
                "" if share is None else "%.1f%%" % share

        if total_label:
            t = table.rows[2 + len(rows)]
            t.cells[c_label].text = total_label
            for i, y in enumerate(years):
                tot = sum((r["years"].get(y) or 0.0) for r in rows)
                t.cells[label_cols + i].text = fmt_for_unit(tot, unit)
            t.cells[label_cols + n].text = "100.0%"

        if widths is None:
            widths = ([420] if rank else []) \
                + ([1100] if code_cols else []) + [3000] + [700] * n + [1300]
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
                for p in cell.paragraphs:
                    for r in p.runs:
                        r.font.bold = True
                        r.font.color.rgb = RGBColor(0x26, 0x26, 0x26)

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
        d = self.add_para(
            cfg.get("date_line") if cfg.get("pinned_date")
            else _current_month_year(),
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
    """Render a doughnut of (label, share) pairs to a PNG and return its path.

    Keeps only the top 5 slices; every remaining slice is folded into an
    "Others" slice so the chart stays readable.
    """
    pairs = [p for p in pairs if p[1] > 0.0]
    if not pairs:
        return None
    pairs.sort(key=lambda p: p[1], reverse=True)
    if len(pairs) > 5:
        top, rest = pairs[:5], pairs[5:]
        pairs = top + [("Others", sum(s for _, s in rest))]
    labels = [l for l, _ in pairs]
    values = [s * 100.0 for _, s in pairs]
    colors = [c if c.startswith("#") else "#" + c for c in THEME]
    fig, ax = charts.new_fig(width=6.6, height=4.4)
    labels, values, wedges = charts.draw_share_pie(
        ax, labels, values, colors, style="donut", min_pct=0.0, max_slices=6)
    charts.share_legend(fig, wedges, labels, values, ncol=2)
    path = os.path.join(tmp_dir, name)
    charts.finish(fig, path)
    return path


def family_members_line(data, top=12):
    members = data.members
    if len(members) <= 1:
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
    b._next_table("Trend on %s: Kenya's Exports by Product" % family, source)
    members_tbl = top_rows(members, cfg.get("top_n", 10), years,
                           "All other products")
    b.add_value_table("Product", members_tbl, years, "Share in %d" % rev,
                      "Kenya's Exports of %s by Product" % family, source,
                      total_label="Total", adaptive_unit=True)

    lead, follows = ([], [])
    pairs = _shares(members, years)
    pairs.sort(key=lambda p: p[1], reverse=True)
    if pairs and len(members) > 1:
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

    destinations = _ranked_rows(data.destinations(), 25, years,
                                ensure_label="Kenya",
                                residual="All other markets")
    if not destinations:
        return
    b._next_table("Destination Markets for Kenya's %s" % anchor, source)
    b.add_value_table("Destination market", destinations, years,
                      "Share in %d" % rev,
                      "Kenya's Exports of %s by Destination" % anchor, source,
                      total_label="Total", rank=True)

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

    # Kenya's standing among its destination markets: leading African market
    # and its position within the overall destination list.
    all_dests = data.destinations()
    afr_dests = [d for d in all_dests if is_africa(d.get("label", ""))
                 and (d["years"].get(rev) or 0.0) > 0]
    if afr_dests:
        afr_dests.sort(key=lambda d: d["years"].get(rev) or 0.0, reverse=True)
        lead_afr = afr_dests[0]
        rev_val = lead_afr["years"].get(rev) or 0.0
        rank = next((i for i, d in enumerate(
            sorted(all_dests,
                   key=lambda d: d["years"].get(rev) or 0.0, reverse=True), 1)
                     if d["label"] == lead_afr["label"]), None)
        text = ("Kenya's leading African destination for %s was %s "
                "(%s; %.1f%% of Kenya's exports of the product)"
                % (anchor.lower(), lead_afr["label"],
                   usd_phrase(rev_val),
                   rev_val / (total_last or 1.0) * 100))
        if rank:
            text += ", ranked %s among all destination markets" % _ordinal(rank)
        b.add_bullet(text + ".")

    # Market-risk callout: heavy reliance on a single destination.
    if first and first["years"].get(rev):
        top_share = (first["years"].get(rev) or 0.0) / (total_last or 1.0) * 100
        if top_share >= 15.0:
            b.add_bullet("Concentration warning: Kenya's leading destination "
                         "for %s accounts for %.1f%% of exports, leaving "
                         "exports exposed to that market's demand." %
                         (anchor.lower(), top_share))

    pairs = _shares(destinations, years)
    if len(pairs) >= 2:
        img = make_donut(pairs, tmp_dir, "f2_dest.png",
                         "Kenya's %s Exports by Destination" % anchor)
        if img:
            b._next_figure("Kenya's %s Exports by Destination, %d"
                           % (anchor, rev), source)
            b.add_figure(img)


def section_competitiveness(b, cfg, data, source):
    """Competitive-intelligence paragraphs: Kenya's market-share trend,
    export specialization, product diversification within the family and a
    comparison with leading African peers.  Each sub-feature silently
    degrades when the underlying data is not available."""
    anchor = short_anchor(data.anchor_label)
    family = cfg.get("family_title", "the product family")
    years = data.years
    rev = data.review_year

    parts = []

    # -- Kenya's share of world exports over time -------------------------
    share_series = data.market_share_series()
    if share_series:
        first_y, first_k, first_w, first_s = share_series[0]
        last_y, last_k, last_w, last_s = share_series[-1]
        txt = ("Kenya's share of world exports of %s %s between %d and %d "
               "(%.2f%% in %d; %.2f%% in %d)."
               % (anchor.lower(),
                  "rose" if last_s >= first_s else "fell",
                  first_y, last_y,
                  first_s * 100, first_y, last_s * 100, last_y))
        if last_s >= first_s:
            txt += " Kenya exported %s of %s in %d." % (
                usd_phrase(last_k), family.lower(), last_y)
        parts.append(txt)

    # -- Export specialization --------------------------------------------
    spec = data.specialization_metrics()
    if spec and spec["years"]:
        speclast = spec["years"][-1]
        fam_share = speclast["share"]
        txt = ("%s accounted for %.2f%% of Kenya's total merchandise exports "
               "in %d (%s of %s), reflecting Kenya's export specialization."
               % (family.title(), fam_share * 100, speclast["year"],
                  usd_phrase(speclast["family"]),
                  usd_phrase(speclast["total"])))
        if len(spec["years"]) >= 2:
            first_s = spec["years"][0]["share"]
            txt += " This share was %.2f%% in %d." % (
                first_s * 100, spec["years"][0]["year"])
        parts.append(txt)

    # -- Product diversification within the family ------------------------
    members = [m for m in data.members
               if (m["years"].get(rev) or 0.0) > 0]
    if len(members) >= 2:
        total = sum(m["years"].get(rev) or 0.0 for m in members)
        lead = max(members, key=lambda m: m["years"].get(rev) or 0.0)
        lead_share = (lead["years"].get(rev) or 0.0) / total * 100
        hhi = sum(((m["years"].get(rev) or 0.0) / total * 100) ** 2
                  for m in members)
        label = ""
        if len(members) > 1:
            second = sorted(members,
                            key=lambda m: m["years"].get(rev) or 0.0,
                            reverse=True)[1]
            second_share = (second["years"].get(rev) or 0.0) / total * 100
            if len(members) == 2:
                label = " %s was the runner-up with %.1f%%." % (
                    short_label(second["label"], 50), second_share)
            else:
                label = " %s was the runner-up with %.1f%%; the remaining %d " \
                    "headings together accounted for %.1f%%." % (
                        short_label(second["label"], 50), second_share,
                        len(members) - 2,
                        max(0.0, 100 - lead_share - second_share))
        concentration = ("highly concentrated" if hhi >= 2500 else
                         "moderately concentrated" if hhi >= 1500 else
                         "diversified")
        txt = ("Kenya's exports of %s are %s across %d product headings: "
               "the leading heading (%s) accounted for %.1f%% of the family "
               "total in %d.%s"
               % (family.lower(), concentration, len(members),
                  short_label(lead["label"], 50), lead_share, rev, label))
        parts.append(txt)

    if parts:
        b.add_heading("KENYA'S COMPETITIVE POSITION IN %s" % anchor.upper(),
                      level=1)
        for p_ in parts:
            b.add_bullet(p_)

    # -- Kenya vs leading African peers -----------------------------------
    peers = data.african_peers(5)
    if peers:
        older = [p for p in peers if p.get("src_year") and p["src_year"] != rev]
        b._next_table("Kenya vs Leading African Exporters of %s"
                      % anchor, source)
        b.add_value_table("Exporting economy", peers, years,
                          "Share in %d" % rev,
                          "African Exporters of %s" % anchor, source,
                          total_label="Total")
        if older:
            b.add_para(
                "Note: %s did not report a %d figure in the source download; "
                "the value shown is for their most recent available year."
                % (ordinal_list([p["label"] for p in older]), rev))


def section_global(b, cfg, data, source, tmp_dir):
    anchor = short_anchor(data.anchor_label)
    family = cfg.get("family_title", "the product family")
    years = data.years
    rev = data.review_year

    b.add_heading("EXPORT OF %s GLOBALLY" % anchor.upper(), level=1)

    all_exporters = data.exporters()
    exporters = _ranked_rows(all_exporters, cfg.get("top_n", 10), years,
                             ensure_label="Kenya", residual="All other economies")
    if exporters:
        b._next_table("World Exports of %s by Economy" % anchor, source)
        b.add_value_table("Exporting economy", exporters, years,
                          "Share in %d" % rev,
                          "Countries Exporting %s" % anchor, source,
                          total_label="Total", rank=True)
        _kenya_standing_bullets(b, all_exporters, years, anchor, "exporter")
        geo_bullets(b, exporters, years, anchor, "exporter")
        pairs = _shares(exporters, years)
        if len(pairs) >= 2:
            img = make_donut(pairs, tmp_dir, "f3_exporters.png",
                             "Share of World Exports of %s" % anchor)
            if img:
                b._next_figure("World Exports of %s by Economy, %d"
                               % (anchor, rev), source)
                b.add_figure(img)

    all_importers = data.importers()
    importers = _ranked_rows(all_importers, cfg.get("top_n", 10), years,
                             ensure_label="Kenya", residual="All other economies")
    if importers:
        b._next_table("World Imports of %s by Economy" % anchor, source)
        b.add_value_table("Importing economy", importers, years,
                          "Share in %d" % rev,
                          "Countries Importing %s" % anchor, source,
                          total_label="Total", rank=True)
        _kenya_standing_bullets(b, all_importers, years, anchor, "importer")
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
        b._next_table("Trend on %s Globally - Export" % family, source)
        b.add_value_table("Product", g_exp, years, "Share in %d" % rev,
                          "Global Exports of %s by Product" % family, source,
                          total_label="Total", adaptive_unit=True)
        trend_bullets(b, g_exp, years, family, "exports")

    g_imp = top_rows(data.global_import_products(),
                     cfg.get("top_n", 10), years, "All other products")
    if g_imp:
        b._next_table("Trend on %s Globally - Import" % family, source)
        b.add_value_table("Product", g_imp, years, "Share in %d" % rev,
                          "Global Imports of %s by Product" % family, source,
                          total_label="Total", adaptive_unit=True)
        trend_bullets(b, g_imp, years, family, "imports")


def _kenya_standing_bullets(b, rows, years, anchor_short, role):
    """Kenya's standing bullets: position in the overall list and within
    Africa, drawn from the *full* (un-truncated) ``rows`` so the rank refers
    to the complete list of economies, not just the displayed top-N."""
    rev = years[-1]
    anchor = anchor_short.lower()
    group = {"exporter": "world exporters", "importer": "world importers"}[role]
    africa_group = {"exporter": "African exporters",
                    "importer": "African importers"}[role]
    action = {"exporter": "exports", "importer": "imports"}[role]
    kenya_row = next((r for r in rows if r.get("label") == "Kenya"), None)
    if kenya_row is None:
        return
    value = kenya_row["years"].get(rev) or 0.0
    ranked = sorted(rows, key=lambda r: r["years"].get(rev) or 0.0, reverse=True)
    global_rank = next((i for i, r in enumerate(ranked, 1)
                        if r.get("label") == "Kenya"), None)
    n_total = sum(1 for r in ranked if (r["years"].get(rev) or 0.0) > 0)
    africa_rank = _africa_rank(ranked, "Kenya")
    n_africa = _africa_count(ranked)

    if value <= 0:
        return
    parts = ["Kenya ranked"]
    if global_rank and n_total:
        parts.append("%s of %d %s" % (_ordinal(global_rank), n_total, group))
    if africa_rank and n_africa:
        if len(parts) > 1:
            parts.append("and")
        parts.append("%s of %d %s"
                     % (_ordinal(africa_rank), n_africa, africa_group))
    if len(parts) > 1:
        b.add_bullet(" ".join(parts) + " of %s in %d, with %s valued at "
                     "%s." % (anchor, rev, action, usd_phrase(value)))


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


def section_kenya_global_position(b, cfg, data, source):
    """Kenya's exports of the family vs the world, by 6-digit HS code.

    Centers the profile on the export side: Kenya's share of world exports of
    each member code, Kenya's overall share of the world market, Kenya's rank
    among world exporters and among African exporters.
    """
    family = cfg.get("family_title", "the product family")
    years = data.years
    rev = data.review_year
    shares = data.kenya_world_shares()
    rows = shares["rows"]
    metrics = data.kenya_global_metrics()

    if not rows:
        return

    b.add_heading("KENYA'S EXPORTS OF %s VS THE WORLD" % family.upper(),
                  level=1)
    b.add_para("This section compares Kenya's exports of %s with world "
               "exports, heading by heading at the six-digit HS code level, "
               "and positions Kenya within the global and African markets in "
               "%d." % (family.lower(), rev))

    # -- Kenya vs World by 6-digit HS code ------------------------------
    k_tot = shares.get("_kenya_total") or 0.0
    w_tot = shares.get("_world_total") or 0.0

    b._next_table("Kenya's Exports of %s vs the World by Product, %d"
                  % (family, rev), source)
    display_rows = []
    for r in rows[: cfg.get("top_n", 10)]:
        k = r["kenya"].get(rev) or 0.0
        w = r["world"].get(rev) or 0.0
        display_rows.append({
            "label": "%s %s" % (r["code"], short_label(r["label"], 44)),
            "years": {"Kenya": k, "World": w},
            "code": r["code"],
            "share": (k / w * 100.0) if w else None,
        })
    _kenya_world_table(b, display_rows, rev, k_tot, w_tot,
                       first_col="Six-digit HS code")

    # -- bullet point on Kenya's share and position ----------------------
    overall = (k_tot / w_tot * 100.0) if w_tot else None
    if overall is not None:
        b.add_bullet("Kenya accounted for %.1f%% of world exports of %s in "
                     "%d (%s of %s)."
                     % (overall, family.lower(), rev, usd_phrase(k_tot),
                        usd_phrase(w_tot)))
    if rows:
        code, k, w = rows[0]["code"], rows[0]["kenya"].get(rev) or 0.0, \
            rows[0]["world"].get(rev) or 0.0
        share = (k / w * 100.0) if w else None
        txt = ("Kenya's leading export heading of %s in %d was %s, valued "
               "at %s." % (family.lower(), rev, code, usd_phrase(k)))
        if share is not None:
            txt += (" That heading accounted for %.1f%% of Kenya's exports "
                    "of the family." % share)
        b.add_bullet(txt)
    if metrics:
        gr = metrics.get("global_rank")
        nr = metrics.get("n_exporters")
        ar = metrics.get("africa_rank")
        na = metrics.get("n_africa")
        parts = ["Kenya ranked"]
        if gr and nr:
            parts.append("%s among %d world exporters of %s"
                         % (_ordinal(gr), nr, family.lower()))
        if ar and na:
            if len(parts) > 1:
                parts.append("and")
            parts.append("%s among %d African exporters" % (_ordinal(ar), na))
        if len(parts) > 1:
            b.add_bullet(" ".join(parts) + " in %d." % rev)


def _ordinal(n):
    return "%d%s" % (n, "th" if 10 <= n % 100 <= 20 else
                     {1: "st", 2: "nd", 3: "rd"}.get(n % 10, "th"))


def _kenya_world_table(b, rows, year, kenya_total, world_total, first_col):
    """Table with a Kenya / World value pair per row plus overall totals.

    Columns: label | Kenya exports | World exports | share.  Reuses
    ``add_value_table``'s styling by treating this as a 2 data-column
    matrix (n=2) plus a share column.  Each column picks its own display
    unit (USD Million / USD Thousand) so that small sub-category values
    never round to 0.0 in millions.
    """
    k_unit = _series_unit([r["years"].get("Kenya") for r in rows]
                          + [kenya_total or 0.0])
    w_unit = _series_unit([r["years"].get("World") for r in rows]
                          + [world_total or 0.0])
    table = b.doc.add_table(rows=2 + len(rows) + 1, cols=4, style="Table Grid")
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    hdr = table.rows[0]
    hdr.cells[0].text = "%s, %d" % (first_col, year)
    hdr.cells[1].text = "Kenya exports (%s)" % k_unit
    hdr.cells[2].text = "World exports (%s)" % w_unit
    hdr.cells[3].text = "Kenya share of world"
    r2 = table.rows[1]
    r2.cells[1].text = str(year)
    r2.cells[2].text = str(year)
    r2.cells[3].text = str(year)
    for ri, r in enumerate(rows, start=2):
        table.rows[ri].cells[0].text = r["label"]
        table.rows[ri].cells[1].text = fmt_for_unit(
            r["years"].get("Kenya"), k_unit)
        table.rows[ri].cells[2].text = fmt_for_unit(
            r["years"].get("World"), w_unit)
        table.rows[ri].cells[3].text = (
            "%.1f%%" % r["share"] if r["share"] is not None else "")
    t = table.rows[2 + len(rows)]
    t.cells[0].text = "Total, %s" % first_col
    t.cells[1].text = fmt_for_unit(kenya_total if kenya_total else None,
                                   k_unit)
    t.cells[2].text = fmt_for_unit(world_total if world_total else None,
                                   w_unit)
    t.cells[3].text = (
        "%.1f%%" % (kenya_total / world_total * 100.0) if world_total else "")
    b._set_table_widths(table, [3500, 1600, 1600, 1500])
    b._style_table(table, rank=False, label_cols=1, n=2, total_label=True)
    b._fit_table_on_page(table)


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
        b._next_table("Kenya's Imports of %s by Source" % anchor, source)
        b.add_value_table("Source market", sources, years, "Share in %d" % rev,
                          "Kenya's Imports of %s by Source" % anchor, source,
                          total_label="Total")
        geo_bullets(b, sources, years, anchor, "source")

    if products:
        products = top_rows(products, cfg.get("top_n", 10), years,
                            "All other products")
        b._next_table("Kenya's Imports of %s by Product" % family, source)
        b.add_value_table("Product", products, years, "Share in %d" % rev,
                          "Kenya's Imports of %s by Product" % family, source,
                          total_label="Total", adaptive_unit=True)
        trend_bullets(b, products, years, family, "imports", scope="Kenya's")


def section_potential(b, cfg, data, pot, source):
    """Optional export-potential section (ITC Export Potential Map file).

    Shows the largest unrealised opportunities: each market's potential
    exports vs the actual base embedded in the download, ranked by the gap."""
    markets, pot_year = _potential_markets(pot)
    if not markets:
        return
    family = cfg.get("family_title", "the product")
    records = []
    for label, m in markets.items():
        potential = m["potential"] or 0.0
        actual = m["actual"] or 0.0
        if potential <= 0:
            continue
        gap = max(0.0, potential - actual) if m["actual"] is not None else None
        records.append({"label": label, "potential": potential,
                        "actual": actual, "gap": gap})
    if not records:
        return
    records.sort(key=lambda r: r["gap"] if r["gap"] is not None else 0.0,
                 reverse=True)
    top = records[:10]
    unit = _series_unit([r["potential"] for r in records]
                        + [r.get("gap") or 0.0 for r in records])
    gap_known = [r for r in top if r["gap"] is not None]
    b.add_heading("EXPORT POTENTIAL FOR %s" % family.upper(), level=1)
    b.add_para("The markets below show the gap between Kenya's current "
               "exports and the potential demand estimated by the ITC Export "
               "Potential Map (projection year %s)."
               % (pot_year if pot_year else "unknown"))
    cols = 3 if gap_known else 2
    table = b.doc.add_table(rows=1 + len(top), cols=cols, style="Table Grid")
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    hdr = table.rows[0]
    hdr.cells[0].text = "Market"
    hdr.cells[1].text = "Export potential (%s)" % unit
    if gap_known:
        hdr.cells[2].text = "Unrealised gap (%s)" % unit
    for i, r in enumerate(top, start=1):
        table.rows[i].cells[0].text = r["label"]
        table.rows[i].cells[1].text = fmt_for_unit(r["potential"], unit)
        if gap_known:
            table.rows[i].cells[2].text = fmt_for_unit(r["gap"], unit)
    b._set_table_widths(table, [3000, 2800, 2800] if gap_known else [4200, 3000])
    b._style_table(table, rank=False, label_cols=1, n=2 if gap_known else 1)
    b._fit_table_on_page(table)
    b.add_bullet("Together, these %d markets offer the greatest headroom for "
                 "additional exports of %s." % (len(top), family.lower()))
    if gap_known:
        lead = gap_known[0]
        b.add_bullet("The largest unrealised opportunity is estimated to be "
                     "%s: potential exports of %s versus current exports of "
                     "%s."
                     % (lead["label"], fmt_for_unit(lead["potential"], unit),
                        fmt_for_unit(lead["actual"], unit)))


def section_growth_decomposition(b, cfg, data, source):
    """Growth decomposition: how much of Kenya's export change came from
    existing product headings (intensive margin), newly-exported headings
    (extensive margin) and headings dropped from the basket."""
    years = data.years
    if len(years) < 2:
        return
    start, rev = years[0], years[-1]
    anchor = short_anchor(data.anchor_label)
    family = cfg.get("family_title", "the product")
    prod_m = margin_decomposition(data.members, start, rev) if data.members \
        else None
    dest_m = margin_decomposition(data.destinations(), start, rev) \
        if data.destinations() else None
    if prod_m is None and dest_m is None:
        return
    if not (prod_m["total_rev"] if prod_m else 0
            or dest_m["total_rev"] if dest_m else 0):
        return
    b.add_heading("WHERE KENYA'S %s EXPORT GROWTH CAME FROM"
                  % anchor.upper(), level=1)
    if prod_m and prod_m["total_rev"]:
        b.add_bullet(_margin_headline(prod_m, family, start, rev,
                                      "product headings"))
    if dest_m and dest_m["total_rev"]:
        b.add_bullet(_margin_headline(dest_m, family, start, rev,
                                      "destination markets"))
    if prod_m and prod_m["total_rev"]:
        _margin_table(b, "Growth of Kenya's exports of %s by product "
                         "heading, %d-%d" % (family, start, rev),
                      prod_m, source, "product heading")
    if dest_m and dest_m["total_rev"]:
        _margin_table(b, "Growth of Kenya's exports of %s by destination "
                         "market, %d-%d" % (family, start, rev),
                      dest_m, source, "destination market")
    for m, noun in ((prod_m, "product heading"), (dest_m, "destination market")):
        if not m or m["net"] <= 0:
            continue
        if m["top_entering"]:
            names = " ".join(short_label(t[0], 45) for t in m["top_entering"][:3])
            b.add_bullet("%d new %s(s) entered the export basket after %d "
                         "(led by %s)."
                         % (m["n_entering"], noun, start, names))
        if m["top_exiting"]:
            names = " ".join(short_label(t[0], 45) for t in m["top_exiting"][:3])
            b.add_bullet("%d %s(s) ceased being exported; the largest were %s."
                         % (m["n_exiting"], noun, names))


def _margin_headline(m, family, start, rev, noun):
    net = m["net"]
    parts = ["Between %d and %d, Kenya's exports of %s %s from %s to %s, "
             "a net change of %s."
             % (start, rev, family.lower(),
                "grew" if net >= 0 else "shrank",
                usd_phrase(m["total_start"]), usd_phrase(m["total_rev"]),
                usd_phrase(abs(net)))]
    if net:
        parts.append("%.0f%% of that change came from existing %s deepening "
                     "their sales, %.0f%% from %s new to the basket, and "
                     "%.0f%% was offset by %s that stopped being exported."
                     % (m["intensive_share"] * 100, noun,
                        m["entering_share"] * 100, noun,
                        abs(m["exiting_share"]) * 100, noun))
    return " ".join(parts)


def _margin_table(b, title, m, source, noun):
    b._next_table(title, source)
    values = [m["intensive"], m["entering"], m["exiting"]]
    unit = _series_unit([abs(v) for v in values] + [abs(m["net"])])
    rows = [("Existing %s deepening" % noun, m["intensive"],
             m["intensive_share"]),
            ("%s(s) new to the export basket" % noun.capitalize(),
             m["entering"], m["entering_share"]),
            ("%s(s) no longer exported" % noun.capitalize(),
             m["exiting"], m["exiting_share"]),
            ("Net change", m["net"], None)]
    table = b.doc.add_table(rows=1 + len(rows), cols=3, style="Table Grid")
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    hdr = table.rows[0]
    for c, text in enumerate(("Component", "Contribution (%s)" % unit,
                              "Share of net change")):
        hdr.cells[c].text = text
    for i, (label, v, share) in enumerate(rows, start=1):
        table.rows[i].cells[0].text = label
        table.rows[i].cells[1].text = fmt_for_unit(v, unit)
        table.rows[i].cells[2].text = "" if share is None \
            else "%.1f%%" % (share * 100)
    widths = [3300, 2600, 2200]
    b._set_table_widths(table, widths)
    b._style_table(table, rank=False, label_cols=1, n=2)
    b._fit_table_on_page(table)


def section_market_attractiveness(b, cfg, data, source, pot):
    """Rank Kenya's destination markets as targets for export expansion:
    weighted blend of momentum, demand headroom, size and market access."""
    years = data.years
    rev = years[-1] if years else None
    if not rev:
        return
    access = cfg.get("access_markets") or []
    weights = cfg.get("attractiveness_weights")
    pot_markets, _ = _potential_markets(pot) if pot else ({}, None)
    recs = market_attractiveness(data.destinations(), years, pot_markets,
                                 weights, access)
    if len(recs) < 3:
        return
    anchor = short_anchor(data.anchor_label)
    family = cfg.get("family_title", "the product")
    b.add_heading("PRIORITY MARKETS FOR EXPANDING KENYA'S %s EXPORTS"
                  % anchor.upper(), level=1)
    b.add_para("Markets are ranked by an attractiveness score that blends "
               "growth momentum (CAGR of Kenya's sales), demand headroom "
               "(unrealised export potential), market size and preferential "
               "access.")
    top = recs[:10]
    unit = _series_unit([r["value_review"] for r in top])
    gap_present = any(r["has_potential"] and r["gap"] is not None for r in top)
    gap_unit = _series_unit([r["gap"] for r in top if r["gap"]])
    headers = ["Market", "CAGR", "Exports %d (%s)" % (rev, unit),
               "Export share"]
    if gap_present:
        headers.append("Potential gap (%s)" % gap_unit)
    headers.append("Score")
    cols = len(headers)
    table = b.doc.add_table(rows=1 + len(top), cols=cols, style="Table Grid")
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    hdr = table.rows[0]
    for c, text in enumerate(headers):
        hdr.cells[c].text = text
    for i, r in enumerate(top, start=1):
        table.rows[i].cells[0].text = r["label"]
        g = r["growth"]
        table.rows[i].cells[1].text = "" if g is None else "%.1f%%" % (g * 100)
        table.rows[i].cells[2].text = fmt_for_unit(r["value_review"], unit)
        table.rows[i].cells[3].text = "%.1f%%" % (r["share"] * 100)
        c = 4
        if gap_present:
            table.rows[i].cells[c].text = fmt_for_unit(r["gap"], gap_unit) \
                if r["gap"] is not None else ""
            c += 1
        table.rows[i].cells[c].text = "%.2f" % r["score"]
    b._set_table_widths(table, [1600, 800, 2440, 1200]
                        + ([1300] if gap_present else []) + [1300])
    b._style_table(table, rank=False, label_cols=1,
                   n=cols - 1)
    b._fit_table_on_page(table)
    leaders = " ".join(short_label(r["label"], 40) for r in top[:3])
    b.add_bullet("Priority markets for the next phase of export growth: %s. "
                 "%s tops the attractiveness ranking on current momentum "
                 "and headroom." % (leaders, top[0]["label"]))
    fast = [r for r in top if r["growth"] and r["growth"] > 0]
    if len(fast) >= 2:
        b.add_bullet("Fast-growth destinations worth deepening: %s."
                     % ", ".join(short_label(r["label"], 30)
                                 for r in fast[:3]))
    big_gaps = [r for r in top if r["gap"] is not None]
    if len(big_gaps) >= 2:
        b.add_bullet("Markets with the largest unrealised headroom: %s."
                     % ", ".join(short_label(r["label"], 30)
                                 for r in big_gaps[:3]))
    b.add_bullet("Focus sectors should be aligned with these ranked markets "
                 "and the preferential routes open to Kenya's exports of %s."
                 % family.lower())


def section_competitor_watch(b, cfg, data, source):
    """Competitor share-shift watch: economies that gained or lost world
    export share in the family over the review period."""
    years = data.years
    if len(years) < 2:
        return
    sc = share_change(data.exporters(), years)
    if not sc or len(sc["rows"]) < 3:
        return
    anchor = short_anchor(data.anchor_label)
    start, rev = years[0], years[-1]
    gainers = [r for r in sc["rows"] if r["delta_pp"] > 0.05][:6]
    losers = [r for r in sc["rows"] if r["delta_pp"] < -0.05][-6:]
    if not gainers and not losers:
        return
    b.add_heading("WHO IS GAINING MARKET SHARE IN %s?"
                  % anchor.upper(), level=1)
    b.add_para("Share of world exports of %s by economy, %d vs %d "
               "(percentage points)." % (anchor.lower(), start, rev))
    for group, title in ((gainers, "Economies gaining world export share"),
                         (losers, "Economies losing world export share")):
        if not group:
            continue
        b._next_table("%s, %d-%d" % (title, start, rev), source)
        table = b.doc.add_table(rows=1 + len(group), cols=4,
                                style="Table Grid")
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        for c, text in enumerate(("Economy", "Share in %d" % start,
                                  "Share in %d" % rev, "Change (pp)")):
            table.rows[0].cells[c].text = text
        for i, r in enumerate(group, start=1):
            table.rows[i].cells[0].text = r["label"]
            table.rows[i].cells[1].text = "%.1f%%" % (r["start_share"] * 100)
            table.rows[i].cells[2].text = "%.1f%%" % (r["rev_share"] * 100)
            table.rows[i].cells[3].text = "%+.1f" % r["delta_pp"]
        b._set_table_widths(table, [2600, 1300, 1300, 1300])
        b._style_table(table, rank=False, label_cols=1, n=3)
        b._fit_table_on_page(table)
    krow = next((r for r in sc["rows"] if r["label"] == "Kenya"), None)
    if krow and abs(krow["delta_pp"]) >= 0.05:
        b.add_bullet("Kenya's share of world exports %s by %.1f percentage "
                     "points between %d and %d (%.1f%% to %.1f%%)."
                     % ("rose" if krow["delta_pp"] > 0 else "fell",
                        abs(krow["delta_pp"]), start, rev,
                        krow["start_share"] * 100, krow["rev_share"] * 100))
    if gainers:
        b.add_bullet("The main share-gainers were %s."
                     % " ".join(short_label(r["label"], 40)
                                for r in gainers[:3]))


def section_strategy(b, cfg, data, source, pot):
    """Strategic read-out: one-line recommendation per leading product
    (growth momentum, family share, priority markets, concentration risk)."""
    years = data.years
    rev = years[-1] if years else None
    if not rev or len(years) < 2:
        return
    anchor = short_anchor(data.anchor_label)
    family = cfg.get("family_title", "the product")
    members = [m for m in data.members if (m["years"].get(rev) or 0.0) > 0]
    if len(members) < 1:
        return
    members.sort(key=lambda m: m["years"].get(rev) or 0.0, reverse=True)
    top = members[:5]
    total = sum(m["years"].get(rev) or 0.0 for m in members)
    access = cfg.get("access_markets") or []
    weights = cfg.get("attractiveness_weights")
    pot_markets, _ = _potential_markets(pot) if pot else ({}, None)
    att = market_attractiveness(data.destinations(), years, pot_markets,
                                weights, access)
    priority = [short_label(r["label"], 30) for r in att[:3]] if att else []
    dests = data.destinations()
    top_dest = dests[0] if dests else None
    dest_total = sum(r["years"].get(rev) or 0.0 for r in dests)

    b.add_heading("STRATEGIC READ-OUT FOR %s EXPORT GROWTH"
                  % anchor.upper(), level=1)
    b.add_para("One-line takeaway per leading product heading, drawn from the "
               "tables throughout this profile.")
    for i, m in enumerate(top, start=1):
        code = str(m.get("code") or "")
        label = m["label"]
        share = (m["years"].get(rev) or 0.0) / total * 100 if total else 0.0
        g = cagr([m["years"].get(y) for y in years], years)
        growth = ("grew at %.1f%% CAGR between %d and %d"
                  % (g * 100, years[0], rev) if g is not None
                  else "showed no clear trend over %d-%d" % (years[0], rev))
        growth_verb = "grew" if (g is None or g <= 0) else "accelerated"
        action = []
        if priority:
            action.append("prioritise %s" % " and ".join(priority))
        if i == 1:
            action.append("anchor heading - defend and deepen")
        bullet = ("%d. %s (%s) - %.1f%% of Kenya's exports of %s in %d; %s."
                   % (i, short_label(label, 55), code, share, family.lower(),
                      rev, growth))
        if action:
            bullet += " Recommended action: %s." % "; ".join(action)
        b.add_bullet(bullet)
    if top_dest and dest_total:
        dshare = (top_dest["years"].get(rev) or 0.0) / dest_total * 100
        if dshare >= 15.0:
            b.add_bullet("Risk to manage: %.1f%% of Kenya's %s exports go to "
                         "a single destination (%s); deepen a second market."
                         % (dshare, family.lower(), top_dest["label"]))
    if att:
        b.add_bullet("Net takeaway: concentrate promotion on %s where "
                     "momentum, headroom and preferential access reinforce "
                     "one another."
                     % " and ".join(short_label(r["label"], 30)
                                    for r in att[:3]))


# --------------------------------------------------------------------------
# Top-level build
# --------------------------------------------------------------------------
def build_profile_document(cfg, data, tmp_dir):
    os.makedirs(tmp_dir, exist_ok=True)
    source = cfg.get("source", "International Trade Centre Database")
    b = ProfileBuilder(cfg, {})
    b.title_page(cfg)
    section_trade_family(b, cfg, data, source, tmp_dir)
    section_kenya_global_position(b, cfg, data, source)
    section_kenya_exports(b, cfg, data, source, tmp_dir)
    section_growth_decomposition(b, cfg, data, source)
    section_market_attractiveness(b, cfg, data, source,
                                  data.files.get("export_potential"))
    section_competitiveness(b, cfg, data, source)
    section_global(b, cfg, data, source, tmp_dir)
    section_competitor_watch(b, cfg, data, source)
    if cfg.get("include_imports", True):
        section_kenya_imports(b, cfg, data, source)
    pot = data.files.get("export_potential")
    if pot:
        section_potential(b, cfg, data, pot, source)
    section_strategy(b, cfg, data, source, pot)
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
    hdr_fill = PatternFill(fill_type=None)
    edit_fill = PatternFill("solid", fgColor="FFF2CC")
    cm = Alignment(horizontal="center")
    lm = Alignment(horizontal="left")
    val_fmt = "#,##0.0"

    def sheet_name(prefix, label="", cap=31):
        name = (prefix + " - " + label) if label else prefix
        return name[:31]

    def value_sheet(name, first_col, rows, doughnut=None, unit=None):
        ws = wb.create_sheet(name)
        if unit is None:
            unit = _series_unit([v for r in rows for y in years
                                 if (v := r["years"].get(y)) is not None])
        has_code = any(r.get("code") for r in rows)
        first = 1 + (1 if has_code else 0)
        if has_code:
            _xc(ws, 1, 1, "Code", bold=True, fill=hdr_fill, align=cm)
        hdr_col = ("%s (USD Thousand)" % first_col if unit == "USD Thousand"
                   else first_col)
        _xc(ws, 1, first, hdr_col, bold=True, fill=hdr_fill, align=cm)
        for i, y in enumerate(years):
            _xc(ws, 1, first + 1 + i, y, bold=True, fill=hdr_fill, align=cm)
        share_col = first + 1 + len(years)
        share_letter = get_column_letter(share_col)
        _xc(ws, 1, share_col, "Share in %d" % rev,
            bold=True, fill=hdr_fill, align=cm)
        first_data, last_data = 2, 1 + len(rows)
        for ri, row in enumerate(rows, start=first_data):
            if has_code:
                _xc(ws, ri, 1, str(row.get("code") or ""), align=lm)
            _xc(ws, ri, first, row["label"], align=lm)
            for i, y in enumerate(years):
                raw = row["years"].get(y)
                if unit == "USD Thousand":
                    v_out = None if raw is None else round(raw, 1)
                else:
                    v_out = None if raw is None else round(display(raw), 1)
                _xc(ws, ri, first + 1 + i, v_out,
                    number_format=val_fmt, align=cm)
            # Share: live formula against the Total row below, so editing a
            # year cell recalculates the whole workbook.
            rev_letter = get_column_letter(first + len(years))
            _xc(ws, ri, share_col, "=%s%d/$%s%d"
                % (rev_letter, ri, rev_letter, last_data + 1),
                bold=True, number_format="0.0%", align=cm)
        # Total row: SUM formulas per year, "100.0%" as the share total.
        trow = last_data + 1
        _xc(ws, trow, first, "Total", bold=True, align=lm)
        for i, y in enumerate(years):
            c_letter = get_column_letter(first + 1 + i)
            _xc(ws, trow, first + 1 + i,
                "=SUM(%s%d:%s%d)" % (c_letter, first_data, c_letter,
                                     last_data),
                bold=True, number_format=val_fmt, align=cm)
        _xc(ws, trow, share_col, 1.0, bold=True, number_format="0.0%",
            align=cm)
        width = max([30] + [len(str(r["label"])) for r in rows])
        if has_code:
            ws.column_dimensions["A"].width = 12
            ws.column_dimensions["B"].width = min(width, 60)
        else:
            ws.column_dimensions["A"].width = min(width, 60)
        for i in range(len(years)):
            ws.column_dimensions[chr(ord("A") + first + i)].width = 11
        if doughnut and len(doughnut[1]) >= 2:
            anchor = ws.cell(2 + len(rows), 1).row + 2
            _add_doughnut_here(ws, anchor, doughnut[0], doughnut[1],
                               share_letter, first_data)
        return ws

    def _add_doughnut_here(ws, top_row, title, labels_values,
                           share_letter="", ref_row=2):
        from openpyxl.chart import DoughnutChart
        from openpyxl.chart.label import DataLabelList
        from openpyxl.chart.legend import Legend
        from openpyxl.chart.series import DataPoint
        from openpyxl.chart.reference import Reference
        ws.cell(top_row, 1).value = "Category"
        ws.cell(top_row, 2).value = "Share"
        for cc in (1, 2):
            c = ws.cell(top_row, cc)
            c.font = Font(bold=True)
            c.fill = hdr_fill
            c.alignment = cm
        for i, (lab, v) in enumerate(labels_values, start=top_row + 1):
            ws.cell(i, 1).value = lab
            cv = ws.cell(i, 2)
            # Live reference into the main table's share column (same sheet),
            # so editing the table moves the chart.
            cv.value = ("=%s%d" % (share_letter, ref_row + i)
                        if share_letter and ref_row + i != top_row else v)
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
    destinations = _ranked_rows(data.destinations(), 25, years,
                                ensure_label="Kenya",
                                residual="All other markets")
    if destinations:
        value_sheet(sheet_name("Destinations", data.anchor_label),
                    "Destination market", destinations,
                    doughnut=("Kenya's exports by destination",
                              [(d["label"], _share01(d, rev, destinations))
                               for d in destinations]))
    exporters = _ranked_rows(data.exporters(), cfg.get("top_n", 10), years,
                             ensure_label="Kenya", residual="All other economies")
    importers = _ranked_rows(data.importers(), cfg.get("top_n", 10), years,
                             ensure_label="Kenya", residual="All other economies")
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

    # Kenya vs the World, by six-digit HS code (export focus)
    shares = data.kenya_world_shares()
    if shares["rows"]:
        ws = wb.create_sheet(sheet_name("Kenya vs World"))
        k_unit = _series_unit([r["kenya"].get(rev) for r in shares["rows"]]
                              + [shares.get("_kenya_total") or 0.0])
        w_unit = _series_unit([r["world"].get(rev) for r in shares["rows"]]
                              + [shares.get("_world_total") or 0.0])
        _xc(ws, 1, 1, "Code", bold=True, fill=hdr_fill, align=cm)
        _xc(ws, 1, 2, "Product", bold=True, fill=hdr_fill, align=cm)
        _xc(ws, 1, 3, "Kenya exports (%s)" % k_unit, bold=True,
            fill=hdr_fill, align=cm)
        _xc(ws, 1, 4, "World exports (%s)" % w_unit, bold=True,
            fill=hdr_fill, align=cm)
        _xc(ws, 1, 5, "Kenya share of world", bold=True, fill=hdr_fill,
            align=cm)
        for ri, r in enumerate(shares["rows"], start=2):
            k_raw = r["kenya"].get(rev) or 0.0
            w_raw = r["world"].get(rev) or 0.0
            _xc(ws, ri, 1, r["code"], align=lm)
            _xc(ws, ri, 2, r["label"], align=lm)
            if k_unit == "USD Thousand":
                _xc(ws, ri, 3, round(k_raw, 1), number_format=val_fmt,
                    align=cm)
            else:
                _xc(ws, ri, 3, round(display(k_raw), 1),
                    number_format=val_fmt, align=cm)
            if w_unit == "USD Thousand":
                _xc(ws, ri, 4, round(w_raw, 1), number_format=val_fmt,
                    align=cm)
            else:
                _xc(ws, ri, 4, round(display(w_raw), 1),
                    number_format=val_fmt, align=cm)
            _xc(ws, ri, 5, "=C%d/D%d" % (ri, ri), bold=True,
                number_format="0.0%", align=cm)
        ti = ri + 1
        _xc(ws, ti, 2, "Total", bold=True)
        for c, letter in ((3, "C"), (4, "D")):
            _xc(ws, ti, c, "=SUM(%s2:%s%d)" % (letter, letter, ri),
                bold=True, number_format=val_fmt, align=cm)
        _xc(ws, ti, 5, "=C%d/D%d" % (ti, ti), bold=True,
            number_format="0.0%", align=cm)
        ws.column_dimensions["A"].width = 12
        ws.column_dimensions["B"].width = 60

    # Kenya vs leading African exporters
    peers = data.african_peers(5)
    if peers:
        value_sheet("African Peers", "Exporting economy", peers)

    # Kenya's share of world exports over time + specialization
    share_series = data.market_share_series()
    spec = data.specialization_metrics()
    if share_series or (spec and spec["years"]):
        ws = wb.create_sheet(sheet_name("Kenya Standing"))
        _xc(ws, 1, 1, "Year", bold=True, fill=hdr_fill, align=cm)
        if share_series:
            _xc(ws, 1, 2, "Kenya exports (USD Million)", bold=True,
                fill=hdr_fill, align=cm)
            _xc(ws, 1, 3, "World exports (USD Million)", bold=True,
                fill=hdr_fill, align=cm)
            _xc(ws, 1, 4, "Kenya share of world", bold=True,
                fill=hdr_fill, align=cm)
        if spec and spec["years"]:
            _xc(ws, 1, 5, "Family share of Kenya exports", bold=True,
                fill=hdr_fill, align=cm)
        ri = 2
        for y in years:
            _xc(ws, ri, 1, y, align=cm)
            if share_series:
                m = next((t for t in share_series if t[0] == y), None)
                if m:
                    _xc(ws, ri, 2, round(display(m[1]), 1),
                        number_format=val_fmt, align=cm)
                    _xc(ws, ri, 3, round(display(m[2]), 1),
                        number_format=val_fmt, align=cm)
                    _xc(ws, ri, 4, "=B%d/C%d" % (ri, ri), bold=True,
                        number_format="0.0%", align=cm)
            if spec and spec["years"]:
                sp = next((s for s in spec["years"] if s["year"] == y), None)
                if sp:
                    _xc(ws, ri, 5, sp["share"], number_format="0.0%", align=cm)
            ri += 1
        ws.column_dimensions["A"].width = 8

    if cfg.get("include_imports", True):
        sources = top_rows(data.kenya_import_sources(), cfg.get("top_n", 10),
                           years, "All other sources")
        if sources:
            value_sheet("Kenya Imports by Source", "Source market", sources)
        prods = top_rows(data.kenya_import_products(), cfg.get("top_n", 10),
                         years, "All other products")
        if prods:
            value_sheet("Kenya Imports by Product", "Product", prods)

    # ---- new analysis sheets ----------------------------------------------
    start_year = years[0]

    # Export Potential gap sheet (ITC Export Potential Map file, if present)
    pot = data.files.get("export_potential")
    if pot:
        markets, pot_year = _potential_markets(pot)
        pot_recs = [{"label": l, "actual": m["actual"],
                     "potential": m["potential"],
                     "gap": None if m["actual"] is None else
                            max(0.0, (m["potential"] or 0.0)
                                - (m["actual"] or 0.0))}
                    for l, m in markets.items()
                    if (m["potential"] or 0.0) > 0]
        pot_recs.sort(key=lambda r: r["gap"] if r["gap"] is not None else -1.0,
                      reverse=True)
        if pot_recs:
            ws = wb.create_sheet("Export Potential")
            unit = _series_unit([r["potential"] for r in pot_recs]
                                + [r.get("gap") or 0.0 for r in pot_recs])
            headers = ["Market", "Current exports (%s)" % unit,
                       "Potential exports (%s)" % unit,
                       "Unrealised gap (%s)" % unit]
            for c, h in enumerate(headers, 1):
                _xc(ws, 1, c, h, bold=True, fill=hdr_fill, align=cm)
            for i, r in enumerate(pot_recs, start=2):
                _xc(ws, i, 1, r["label"], align=lm)
                for j, key in ((2, "actual"), (3, "potential")):
                    raw = r[key]
                    v = None if raw is None else (
                        round(raw, 1) if unit == "USD Thousand"
                        else round(display(raw), 1))
                    _xc(ws, i, j, v, number_format=val_fmt, align=cm)
                _xc(ws, i, 4, "=C%d-B%d" % (i, i)
                    if r["actual"] is not None else None,
                    number_format=val_fmt, bold=True, align=cm)
            ws.column_dimensions["A"].width = 40

    # Growth Decomposition sheet
    if len(years) >= 2 and data.members:
        margins = []
        for label, pred in (("Existing product headings (intensive)",
                             lambda s, r: bool(s and r)),
                            ("Newly exported product headings (extensive)",
                             lambda s, r: bool(not s and r)),
                            ("Headings no longer exported (exits)",
                             lambda s, r: bool(s and not r))):
            s0 = s1 = 0.0
            for m in data.members:
                sv = m["years"].get(start_year) or 0.0
                rv = m["years"].get(rev) or 0.0
                if pred(bool(sv), bool(rv)):
                    s0 += sv
                    s1 += rv
            margins.append({"label": label, "start": s0, "rev": s1})
        if any(r["rev"] or r["start"] for r in margins):
            ws = wb.create_sheet("Growth Decomposition")
            headers = ["Margin", "Exports %d" % start_year,
                       "Exports %d" % rev, "Change",
                       "Share of net change"]
            for c, h in enumerate(headers, 1):
                _xc(ws, 1, c, h, bold=True, fill=hdr_fill, align=cm)
            for i, r in enumerate(margins, start=2):
                _xc(ws, i, 1, r["label"], align=lm)
                for j, key in ((2, "start"), (3, "rev")):
                    _xc(ws, i, j, round(display(r[key]), 1),
                        number_format=val_fmt, align=cm)
                _xc(ws, i, 4, "=C%d-B%d" % (i, i),
                    number_format=val_fmt, align=cm)
                _xc(ws, i, 5, "=D%d/$D$%d" % (i, len(margins) + 2),
                    number_format="0.0%", align=cm)
            trow = len(margins) + 2
            _xc(ws, trow, 1, "Net change", bold=True, align=lm)
            _xc(ws, trow, 2, "=SUM(B2:B%d)" % (trow - 1),
                bold=True, number_format=val_fmt, align=cm)
            _xc(ws, trow, 3, "=SUM(C2:C%d)" % (trow - 1), bold=True,
                number_format=val_fmt, align=cm)
            _xc(ws, trow, 4, "=C%d-B%d" % (trow, trow), bold=True,
                number_format=val_fmt, align=cm)
            _xc(ws, trow, 5, "=D%d/D%d" % (trow, trow), bold=True,
                number_format="0.0%", align=cm)
            ws.column_dimensions["A"].width = 50

    # Market Attractiveness sheet
    access = cfg.get("access_markets") or []
    weights = cfg.get("attractiveness_weights")
    pot_markets, _ = _potential_markets(pot) if pot else ({}, None)
    att = market_attractiveness(data.destinations(), years,
                                pot_markets, weights, access)
    if len(att) >= 3:
        ws = wb.create_sheet("Market Attractiveness")
        gap_any = any(x["has_potential"] and x["gap"] is not None
                      for x in att)
        headers = ["Rank", "Market", "Kenya exports %d" % rev, "Share",
                   "CAGR %d-%d" % (start_year, rev), "Access tier"] \
            + (["Potential gap"] if gap_any else []) \
            + ["Attractiveness score"]
        for c, h in enumerate(headers, 1):
            _xc(ws, 1, c, h, bold=True, fill=hdr_fill, align=cm)
        unit = _series_unit([x["value_review"] for x in att])
        tier_names = {1: "Priority", 2: "Africa", 0: "Other"}
        first_data, last_data = 2, 1 + len(att)
        for i, x in enumerate(att, start=first_data):
            _xc(ws, i, 1, i - 1, align=cm)
            _xc(ws, i, 2, x["label"], align=lm)
            raw = x["value_review"]
            _xc(ws, i, 3, round(raw, 1) if unit == "USD Thousand"
                else round(display(raw), 1), number_format=val_fmt,
                align=cm)
            _xc(ws, i, 4, "=C%d/$C$%d" % (i, last_data + 1),
                number_format="0.0%", align=cm)
            _xc(ws, i, 5, x["growth"], number_format="0.0%", align=cm)
            _xc(ws, i, 6, tier_names.get(x["tier"], "Other"), align=cm)
            c = 7
            if gap_any:
                gap = x["gap"] if x["has_potential"] else None
                if gap is not None:
                    gap = round(gap, 1) if unit == "USD Thousand" \
                        else round(display(gap), 1)
                _xc(ws, i, c, gap, number_format=val_fmt, align=cm)
                c += 1
            _xc(ws, i, c, x["score"], number_format="0.00", bold=True,
                align=cm)
        trow = last_data + 1
        _xc(ws, trow, 2, "Total", bold=True, align=lm)
        _xc(ws, trow, 3, "=SUM(C%d:C%d)" % (first_data, last_data),
            bold=True, number_format=val_fmt, align=cm)
        ws.column_dimensions["B"].width = 40

    # Competitor Share Shift sheet
    sc = share_change(data.exporters(), years)
    if sc and sc["rows"]:
        ws = wb.create_sheet("Competitor Share Shift")
        headers = ["Economy", "World share %d" % start_year,
                   "World share %d" % rev, "Change (share points)"]
        for c, h in enumerate(headers, 1):
            _xc(ws, 1, c, h, bold=True, fill=hdr_fill, align=cm)
        for i, x in enumerate(sc["rows"], start=2):
            _xc(ws, i, 1, x["label"], align=lm)
            _xc(ws, i, 2, x["start_share"], number_format="0.0%",
                align=cm)
            _xc(ws, i, 3, x["rev_share"], number_format="0.0%",
                align=cm)
            _xc(ws, i, 4, x["delta_pp"], number_format="0.00",
                bold=True, align=cm)
        ws.column_dimensions["A"].width = 40

    # Scenario sheet: editable target-share column with live formulas
    members_top = (sorted(data.members,
                          key=lambda r: r["years"].get(rev) or 0.0,
                          reverse=True)[:10]
                   if data.members else [])
    if members_top:
        ws = wb.create_sheet("Scenario")
        headers = ["Product", "Exports %d" % rev, "Current share",
                   "Target share (editable)", "Target exports %d" % rev,
                   "Implied export growth"]
        for c, h in enumerate(headers, 1):
            _xc(ws, 1, c, h, bold=True, fill=hdr_fill, align=cm)
        total = sum(m["years"].get(rev) or 0.0 for m in data.members)
        for i, m in enumerate(members_top, start=2):
            _xc(ws, i, 1, m["label"], align=lm)
            cur = (m["years"].get(rev) or 0.0) / total if total else 0.0
            _xc(ws, i, 2, round(display(m["years"].get(rev) or 0.0), 1),
                number_format=val_fmt, align=cm)
            _xc(ws, i, 3, cur, number_format="0.0%", align=cm)
            _xc(ws, i, 4, cur, number_format="0.0%", align=cm,
                fill=edit_fill)
            _xc(ws, i, 5, "=B%d*(D%d/C%d)" % (i, i, i),
                number_format=val_fmt, align=cm)
            _xc(ws, i, 6, "=E%d/B%d-1" % (i, i), number_format="0.0%",
                bold=True, align=cm)
        note = len(members_top) + 3
        _xc(ws, note, 1, "Edit the yellow Target share cells to model "
                         "raising each product's share; Target exports and "
                         "Implied growth recalculate automatically.",
            align=lm)
        ws.column_dimensions["A"].width = 70

    # default "Sheet" removed by the first create_sheet call? keep membership
    if "Sheet" in wb.sheetnames:
        del wb["Sheet"]
    wb.save(out_path)
    return wb


def _share01(row, rev, rows):
    total = sum(display(r["years"].get(rev)) or 0.0 for r in rows)
    v = display(row["years"].get(rev))
    if v is None:
        return None
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

    data = ProfileData(data_dir, cfg.get("include_codes"),
                       cfg.get("family_title"))
    print("[1/4] Loading ITC files from      : %s" % data_dir)
    print("      family    = %s" % cfg.get("family_title"))
    print("      anchor    = %s (%s)" % (data.anchor_hs, data.anchor_label))
    print("      period    = %d - %d" % (data.start_year, data.review_year))
    print("      members   = %s (product detail rows)"
          % len(data.members))
    for w in data.warnings:
        print("      [warn] %s" % w)

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