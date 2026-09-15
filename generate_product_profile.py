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

# Ordered upload manifest the web app can surface: every required download for
# a product-profile analysis (phase 1) followed by the optional ones (phase 2)
# that fill the quantity, Africa, supplier and Export-Potential tables.  The
# "name" uses Trade Map's slug with a placeholder for the product/anchor.
REQUIRED_UPLOADS = [
    {"prefix": "exporting-economies", "key": "world_exports_by_economy",
     "label": "Top exporting economies of the product (value, USD Thousand)",
     "measure": "value"},
    {"prefix": "products-exported-globally", "key": "world_exports_by_product",
     "label": "Exports of the product by product detail (value, USD Thousand)",
     "measure": "value"},
    {"prefix": "importing-economies", "key": "world_imports_by_economy",
     "label": "Top importing economies of the product (value, USD Thousand)",
     "measure": "value"},
    {"prefix": "products-imported-globally", "key": "world_imports_by_product",
     "label": "Imports of the product by product detail (value, USD Thousand)",
     "measure": "value"},
    {"prefix": "kenyas-exports-to-world-by-importer",
     "key": "kenya_exports_by_partner",
     "label": "Where does Kenya export the product (destinations, value)",
     "measure": "value"},
    {"prefix": "kenyas-exports-to-world-by-product",
     "key": "kenya_exports_by_product",
     "label": "Kenya's exports of the product by product detail (value)",
     "measure": "value"},
    {"prefix": "kenyas-imports-from-world-by-exporter",
     "key": "kenya_imports_by_partner",
     "label": "Kenya's imports of the product by source market (value)",
     "measure": "value"},
    {"prefix": "kenyas-imports-from-world-by-product",
     "key": "kenya_imports_by_product",
     "label": "Kenya's imports of the product by product detail (value)",
     "measure": "value"},
]

OPTIONAL_UPLOADS = [
    {"prefix": "products-exported-globally", "key": "world_exports_by_product",
     "kind": "quantity",
     "label": "Exports of the product by product detail (Quantity, Tonnes) - "
              "same file, exported with the Tonnes option",
     "measure": "quantity",
     "note": "fills Table 3 (global export quantities)"},
    {"prefix": "products-imported-globally", "key": "world_imports_by_product",
     "kind": "quantity",
     "label": "Imports of the product by product detail (Quantity, Tonnes)",
     "measure": "quantity",
     "note": "fills Table 6 (global import quantities)"},
    {"prefix": "kenyas-exports-to-world-by-product",
     "key": "kenya_exports_by_product", "kind": "quantity",
     "label": "Kenya's exports by product (Quantity, Tonnes)",
     "measure": "quantity", "note": "fills Table 14"},
    {"prefix": "kenyas-imports-from-world-by-product",
     "key": "kenya_imports_by_product", "kind": "quantity",
     "label": "Kenya's imports by product (Quantity, Tonnes)",
     "measure": "quantity", "note": "fills Table 17"},
    {"prefix": "africas-exports-to-world-by-product",
     "key": "africa_exports_by_product",
     "label": "Exports of the product by product detail from Africa (value)",
     "measure": "value", "note": "fills Table 8"},
    {"prefix": "africas-imports-from-world-by-product",
     "key": "africa_imports_by_product",
     "label": "Imports of the product by product detail into Africa (value)",
     "measure": "value", "note": "fills Table 10"},
    {"prefix": "south-sudan-imports-from-world-by-exporter",
     "key": "supplier_south_sudan", "kind": "supplier",
     "label": "List of supplying markets for a product imported by <Top "
              "Destination> (one file per market, value)",
     "measure": "value",
     "note": "one file per top destination - fills Table 18 and the "
             "competitor analysis; download after phase 1 identifies the "
             "markets",
     "repeat": "per top destination market"},
    {"prefix": "export_potential", "key": "export_potential",
     "kind": "potential",
     "label": "Export Potential Map for the product (markets x potential)",
     "measure": "value",
     "note": "one map per product sub-family - fills the Export Potential "
             "figures and table",
     "repeat": "one per product sub-family"},
]

DEFAULT_MAX_YEARS = 5


def upload_manifest(anchor="<product>"):
    """Ordered list of files the web app should ask the user to upload for a
    product-profile analysis.  Each entry has ``required``, ``name``,
    ``label`` and (for optionals) ``phase``/``note`` so the UI can render the
    checklist and surface click-to-download hints."""
    out = []
    for spec in REQUIRED_UPLOADS:
        out.append({"required": True, "phase": 1, "name":
                    "%s_%s.xlsx" % (spec["prefix"], anchor),
                    "measure": spec["measure"], "label": spec["label"]})
    for i, spec in enumerate(OPTIONAL_UPLOADS, start=1):
        out.append({
            "required": False,
            "phase": 2 if spec.get("kind") in ("supplier",) else 2,
            "name": "%s_%s.xlsx" % (spec["prefix"], anchor),
            "measure": spec["measure"], "label": spec["label"],
            "note": spec["note"], "repeat": spec.get("repeat", "")})
    return out


# --------------------------------------------------------------------------
# ITC matrix loading
# --------------------------------------------------------------------------
def _walk_files(data_dir):
    """Every regular file under ``data_dir``, recursing into sub-folders.

    Web-app uploads arrive as individually-uniqued objects (their original
    names may collide across uploads - e.g. a value and a quantity twin that
    share the exact same filename) so the loader must not assume the files
    sit flat in one directory.  Hidden files (``~``, leading ``.``) are
    skipped.
    """
    out = []
    for dirpath, _dirs, files in os.walk(data_dir):
        for f in files:
            base = f.lower()
            if base.startswith("~") or base.startswith("."):
                continue
            out.append(os.path.join(dirpath, f))
    return sorted(out)


def _find_files(data_dir, prefix):
    """All files under ``data_dir`` (recursively) whose basename begins with
    ``prefix``; when nothing matches, files that merely contain ``prefix``.
    Sorted so the browser "(1)" copies never win over the primary file."""
    all_files = _walk_files(data_dir)
    hits = [p for p in all_files if
            os.path.basename(p).lower().startswith(prefix)]
    if not hits:
        hits = [p for p in all_files if
                prefix in os.path.basename(p).lower()]
    return hits


def _find_file(data_dir, prefix):
    hits = _find_files(data_dir, prefix)
    return hits[0] if hits else None


_MEASURE_HDR_RE = re.compile(
    r"\s*(20\d\d)\s*\(\s*(?P<unit>[^)]{1,40}?)\s*\)")
_QTY_SLUG_RE = re.compile(r"(tonne|tonnes|quantity| kg)", re.IGNORECASE)


def _matrix_unit(path):
    """Classify an ITC matrix download as ``"value"`` or ``"quantity"`` by the
    unit of its year-column headers.

    Trade Map lets every by-product download be exported either in *value*
    (``2016 (USD Thousand)``) or in *quantity* (``2016 (Tonnes)`` /
    ``2016 (Quantity in tonnes)``).  The two look identical otherwise - and
    can even share the same filename when uploaded through the web app - so
    the measure is read from the ``path``'s header before the file is
    assigned to the value pipeline or the tonnes pipeline.  Classic HTML
    ``.xls`` downloads (an "Exported value"/"Tons" unit row in the table
    body) are read textually; if no unit is found the filename slug decides,
    then ``"value"``."""
    try:
        wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
        ws = wb.worksheets[0]
        header = next(ws.iter_rows(values_only=True))
        wb.close()
    except Exception:
        header = ()
    for h in header:
        m = _MEASURE_HDR_RE.match(str(h or ""))
        if m:
            unit = m.group("unit").lower()
            if ("ton" in unit or "kg" in unit or "quantity" in unit):
                return "quantity"
            return "value"
    unit = _classic_unit(path)
    if unit is not None:
        return unit
    if _QTY_SLUG_RE.search(os.path.basename(path)):
        return "quantity"
    return "value"


# Classic HTML ``.xls`` files carry their unit as a table-body row
# ("Exported quantity ... Tons") rather than in a parenthetical header, so
# the measure is read from a short textual scan of the file head.
def _classic_unit(path):
    try:
        with open(path, "r", encoding="utf-8", errors="replace") as fh:
            head = fh.read(12000)
    except Exception:
        return None
    head = head.lower()
    has_tons = bool(re.search(r"\b(tonnes?|tons?)\b", head))
    has_usd = bool(re.search(r"\b(usd\b|usd thousand|exchanged value|thousand us)",
                             head))
    if has_tons and not has_usd:
        return "quantity"
    return "value" if has_usd else None


# Classic Trade Map HTML ``.xls`` downloads are named after the report they
# are ("Trade_Map_-_List_of_exporters_for_the_selected_product_(...).xls")
# rather than ITC's beta-product slugs.  Map those stems back to the same
# loader keys used by the slugged files.
_CLASSIC_STEMS = {
    "world_exports_by_economy":
        ("list_of_exporters_for_the_selected_product",),
    "world_imports_by_economy":
        ("list_of_importers_for_the_selected_product",),
    "world_exports_by_product":
        ("list_of_exported_products_for_the_selected_product",),
    "world_imports_by_product":
        ("list_of_imported_products_for_the_selected_product",),
    "kenya_exports_by_product":
        ("list_of_products_exported_by_kenya",),
    "kenya_imports_by_product":
        ("list_of_products_imported_by_kenya",),
    "kenya_exports_by_partner":
        ("list_of_importing_markets_for_a_product_exported_by_kenya",
         "list_of_destination_countries_for_a_product_exported_by_kenya"),
    "kenya_imports_by_partner":
        ("list_of_exporting_markets_for_a_product_imported_by_kenya",
         "list_of_supplying_markets_for_a_product_imported_by_kenya"),
    "export_potential": ("export_potential",),
}
_BY_ECONOMY_KEYS = {"world_exports_by_economy", "world_imports_by_economy"}
_BY_PARTNER_KEYS = {"kenya_exports_by_partner", "kenya_imports_by_partner"}
_BY_PRODUCT_KEYS = {
    "world_exports_by_product", "world_imports_by_product",
    "kenya_exports_by_product", "kenya_imports_by_product",
    "africa_exports_by_product", "africa_imports_by_product",
}


def _classic_group(path):
    """The product group of a classic filename: ``(..._((Meat_and_edible_
    meat_offal)).xls`` -> ``Meat and edible meat offal``."""
    m = re.search(r"\(([^()]+)\)\.?\s*xls?$",
                  os.path.basename(path), re.IGNORECASE)
    if not m:
        return ""
    return " ".join(
        re.sub(r"[-_]+", " ", m.group(1)).lower().split())


def _classic_candidates(data_dir, key):
    """Classic ``.xls`` downloads under ``data_dir`` that supply ``key``.

    A family economy download ("List of exporters for the selected product")
    often sits next to per-sub-product twins; when several match, only the
    candidate whose group matches the family's by-product download is kept
    (otherwise the aggregate row would be shadowed by a sub-product file)."""
    probes = _CLASSIC_STEMS.get(key) or ()
    if not probes:
        return []
    by_name = []
    for p in _walk_files(data_dir):
        base = os.path.basename(p).lower()
        if any(stem in base for stem in probes):
            by_name.append(p)
    if not by_name:
        return []
    if key in _BY_ECONOMY_KEYS and len(by_name) > 1:
        family = ""
        for stem in ("list_of_exported_products_for_the_selected_product",
                     "list_of_imported_products_for_the_selected_product"):
            mate = _classic_candidates(data_dir, "world_exports_by_product" if
                                       "export" in stem else
                                       "world_imports_by_product")
            if mate:
                family = _classic_group(mate[0])
                break
        if family:
            picks = [p for p in by_name if _classic_group(p) == family]
            if picks:
                by_name = picks
    return by_name


def _classic_to_matrix(key, records):
    """Convert classic ``{code, label, years}`` records to the matrix shape
    the profile accessors expect (``reporter``/``partner``/``product``)."""
    out = []
    for r in records:
        code = str(r.get("code") or "").strip()
        label = r.get("label")
        years = r.get("years") or {}
        if not code and not label:
            continue
        is_world = code.lower() in ("world", "total", "all")
        if key in _BY_ECONOMY_KEYS:
            out.append({"reporter": "000" if is_world else code,
                        "reporter_label": "World" if is_world else code,
                        "partner": "000", "partner_label": "World",
                        "product": "", "product_label": "", "years": years})
        elif key in _BY_PARTNER_KEYS:
            out.append({"reporter": "404", "reporter_label": "Kenya",
                        "partner": "000" if is_world else code,
                        "partner_label": "World" if is_world else code,
                        "product": "", "product_label": "", "years": years})
        elif key in _BY_PRODUCT_KEYS:
            out.append({"reporter": "000", "reporter_label": "World",
                        "partner": "000", "partner_label": "World",
                        "product": code, "product_label": str(label or ""),
                        "years": years})
    return out


def _load_any(path, key=None):
    """Load an ITC matrix workbook, falling back to the HTML-table parser for
    Trade Map's classic ``.xls`` downloads (BIFF/HTML saved under ``.xls``)."""
    try:
        return load_matrix(path, key)
    except Exception:
        return load_html_matrix(path)


def _find_region_product_file(data_dir, flow):
    """Locate a 'List of products exported/imported by <region>' workbook.

    Trade Map regions report their trade as a pseudo-economy (e.g. Africa), so
    the by-product download for a region carries the same shape as the world
    by-product matrix but with ``africa`` in the filename.  Both the beta slug
    (``africas-exports-to-world-by-product_all.xlsx``) and the classic name
    (``Trade_Map_-_List_of_products_exported_by_Africa.xls``) are recognised.
    """
    verbs = {"export": ("export", "exported", "exports"),
             "import": ("import", "imported", "imports")}[flow]
    for path in _walk_files(data_dir):
        base = os.path.basename(path).lower()
        if "africa" in base:
            if any(v in base for v in verbs) and \
                    ("by-product" in base or "list_of_products" in base):
                return path
    return None


def _supplier_market_label(path):
    """Best-effort market name from a per-market supplying-markets download.

    ``united-arab-emirates-imports-from-world-by-exporter_all.xlsx`` ->
    ``United Arab Emirates``; ``Trade_Map_-_List_of_supplying_markets_for_a_
    product_imported_by_Saudi_Arabia.xls`` -> ``Saudi Arabia``."""
    stem = os.path.splitext(os.path.basename(path))[0].lower()
    m = re.match(r"^([a-z0-9\-_ ]+?)-?imports-from-world-by-exporter", stem) \
        or re.search(r"imported_by[_\s]+(.+?)\.", stem)
    if not m:
        return None
    name = re.sub(r"[-_]+", " ", m.group(1)).strip().title()
    if not name or name.lower() in ("kenya", "world"):
        return None
    return display_name(fix_label(name))


def _header_year(cell):
    """Extract a review year from a Trade Map column header.

    Classic headers look like ``2024 (USD thousand)``; the current workbook
    downloads use ``Export value in 2024 (USD thousand)`` or a bare ``2024``.
    Indicator columns (shares, growth, indices, totals, unit values) are
    ignored so only genuine annual trade columns are treated as years.
    """
    low = str(cell or "").lower().strip()
    if not low:
        return None
    if any(m in low for m in ("share", "growth", "index", "trend",
                              "concentration", "tariff", "unit value",
                              "total")):
        return None
    m = re.search(r"(20\d\d)", low)
    return int(m.group(1)) if m else None


def _matrix_roles(header, key):
    """Map column positions from the header names of current Trade Map exports.

    ``List of exporters ...`` workbooks name their columns (``Reporter ISO``,
    ``Reporter``, ``Partner``, ``Product code``, ``Product label``, ...) while
    the classic layout is exactly ``[reporter, reporter_label, partner,
    partner_label, product, product_label]``.  ISO / code columns are skipped
    and the economy / market / product label columns are detected per loader
    key.  Returns ``{}`` when the header carries no useful names."""
    if not header:
        return {}
    words = [" ".join(x for x in re.split(r"[^a-z0-9]+",
                                           str(h or "").lower()) if x)
             for h in header]
    roles = {}
    name_cols = []
    for i, low in enumerate(words):
        if not low:
            continue
        if "iso" in low or ("code" in low and "product" not in low
                            and "hs" not in low and "ntl" not in low):
            continue  # identifier column (Reporter ISO, Partner ISO, ...)
        if "product" in low or "hs" in low or "ntl" in low:
            if "label" in low or "descrip" in low:
                roles.setdefault("product_label", i)
            elif "code" in low:
                roles.setdefault("product_code", i)
            continue
        if any(w in low for w in ("report", "export", "import", "econom",
                                  "countr", "market", "partner",
                                  "destination", "source")):
            name_cols.append(i)
    # Name columns appear in the same order as the table reads them: for the
    # by-market (Kenya) files the first is the reporter, the second the
    # partner; economy lists carry a single name column.
    if key in _BY_PARTNER_KEYS and name_cols:
        roles["partner_label"] = (name_cols[1] if len(name_cols) >= 2
                                  else name_cols[0])
    elif key in _BY_ECONOMY_KEYS and name_cols:
        roles["reporter_label"] = name_cols[0]
    return roles


def load_matrix(path, key=None):
    """Parse an ITC all-countries / all-products matrix workbook.

    Accepts both the classic ``.xls`` exports (whose first table row is the
    header) and the current Trade Map ``.xlsx`` downloads, which start with a
    few metadata lines above the actual column header.  Column roles are taken
    from the header names (``key`` tells which label - reporter, partner or
    product - drives the table).  Returns ``(years, records)`` where
    ``years`` is the ordered list of review years picked up from the year
    columns and each record is ``{"reporter", "reporter_label", "partner",
    "partner_label", "product", "product_label", "years": {year: value}}``.
    """
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    ws = wb.worksheets[0]
    rows = list(ws.iter_rows(values_only=True))
    wb.close()

    # Locate the true header row: modern Downloads carry a metadata preamble
    # (source note, unit, reporting period) before the column names.
    header_idx = 0
    year_cols = {}
    for i, r in enumerate(rows):
        cols = {c: y for c, y in
                ((c, _header_year(v)) for c, v in enumerate(r))
                if y is not None}
        if len(cols) >= 2:
            header_idx, year_cols = i, cols
            break
    if not year_cols:
        return [], []
    years = sorted(set(year_cols.values()))
    year_ix = {y: next(c for c, yy in year_cols.items() if yy == y)
               for y in years}

    header = rows[header_idx] if header_idx < len(rows) else ()
    roles = _matrix_roles(header, key)
    pos = {"reporter": 0, "reporter_label": 1, "partner": 2,
           "partner_label": 3, "product": 4, "product_label": 5}
    if "reporter_label" in roles:
        pos["reporter_label"] = roles["reporter_label"]
    if "partner_label" in roles:
        pos["partner_label"] = roles["partner_label"]
    if "product_code" in roles:
        pos["product"] = roles["product_code"]
    if "product_label" in roles:
        pos["product_label"] = roles["product_label"]
    non_year = [c for c in range(len(header)) if c not in year_cols]

    def cell(r, col):
        return r[col] if 0 <= col < len(r) else None

    def name_of(r, pref, default):
        """The label value at ``pref``, or the first text-bearing cell before
        the year columns when ``pref`` holds a bare code / nothing."""
        if pref is not None:
            v = cell(r, pref)
            if v is not None and str(v).strip() and \
                    re.search(r"[A-Za-z]", str(v)):
                return v
        for c in ([pref] if pref is not None else []) + non_year:
            v = cell(r, c)
            if v is None:
                continue
            s = str(v).strip()
            if s and re.search(r"[A-Za-z]", s):
                return v
        v = cell(r, pref) if pref is not None else cell(r, default)
        return v if v is not None else ""

    records = []
    for r in rows[header_idx + 1:]:
        if not r or r[0] is None or str(r[0]).strip() == "":
            continue
        if isinstance(r[0], str) and r[0].strip().lower().startswith("source"):
            continue
        if key in (_BY_ECONOMY_KEYS | _BY_PRODUCT_KEYS):
            # World-partnered matrices (economy or product view) never carry a
            # partner column of their own: narrow modern downloads would even
            # put a year figure in r[2], so force the world partner here.
            partner, partner_label = "000", "World"
        else:
            partner = str(cell(r, 2) or "")
            partner_label = str(name_of(r, pos["partner_label"], 3) or "")
        records.append({
            "reporter": str(cell(r, 0) or ""),
            "reporter_label": str(
                name_of(r, pos["reporter_label"], 1) or ""),
            "partner": partner,
            "partner_label": partner_label,
            "product": str(cell(r, pos["product"]) or ""),
            "product_label": clean_label(str(
                name_of(r, pos["product_label"], 5) or "")),
            "years": {y: to_float(cell(r, c)) for y, c in year_ix.items()},
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

    def __init__(self, data_dir, include_codes=None, family_title=None,
                 max_years=None):
        self.data_dir = os.path.abspath(data_dir)
        self.include_codes = list(include_codes or [])
        self.family_title = family_title
        self.max_years = max_years or DEFAULT_MAX_YEARS
        self.warnings = []
        self.files = {}                        # value (USD Thousand) matrices
        self.qty_files = {}                    # parallel Quantity (Tonnes)
        self.potential_files = []              # one Export Potential Map each

        # Every known download is matched by prefix; value and quantity twins
        # of a product table are told apart by their header measure, so both
        # can live in one folder.
        qty_keys = set(FILE_PREFIXES.values()) - {"export_potential"}
        for prefix, key in FILE_PREFIXES.items():
            candidates = _find_files(self.data_dir, prefix)
            # A slugged download and a classic Trade Map .xls for the same
            # report can coexist (e.g. value as .xlsx + the Tonnes twin as the
            # classic HTML .xls): keep every candidate, each is then sorted
            # into the value / quantity pipelines by its header measure.
            for p in _classic_candidates(self.data_dir, key):
                if p not in candidates:
                    candidates.append(p)
            if not candidates:
                continue  # surfaced as a structured alert in file_status
            if key == "export_potential":
                for p in candidates:
                    years, records = _load_any(p)
                    if not records:
                        continue
                    label = next((r["product_label"] for r in records
                                  if r.get("product_label")), "") or \
                        os.path.splitext(os.path.basename(p))[0]
                    self.potential_files.append(
                        {"path": p, "years": years, "records": records,
                         "label": label})
                continue
            value_hits = [p for p in candidates if _matrix_unit(p) == "value"]
            qty_hits = [p for p in candidates if _matrix_unit(p) == "quantity"]
            if len(value_hits) > 1:
                self.warnings.append(
                    "%d value candidates for '%s' in %s - using %s"
                    % (len(value_hits), prefix, self.data_dir,
                       os.path.basename(value_hits[0])))
            if len(qty_hits) > 1:
                self.warnings.append(
                    "%d quantity candidates for '%s' in %s - using %s"
                    % (len(qty_hits), prefix, self.data_dir,
                       os.path.basename(qty_hits[0])))

            def _store(path, store_self):
                try:
                    years, records = _load_any(path, key)
                except Exception:
                    return False
                if records is None:
                    return False
                if records and "product" not in records[0] and \
                        key in (_BY_ECONOMY_KEYS | _BY_PARTNER_KEYS |
                                _BY_PRODUCT_KEYS):
                    records = _classic_to_matrix(key, records)
                (self.files if store_self else self.qty_files)[key] = \
                    {"path": path, "years": years, "records": records}
                return True

            if value_hits:
                _store(value_hits[0], True)
            if qty_hits and key in qty_keys:
                # A quantity-only download fills the Tonnes table on its own;
                # the value table is left empty (surfaced as an alert).
                if not value_hits:
                    _store(qty_hits[0], False)
                else:
                    qy, qr = _load_any(qty_hits[0])
                    if qr:
                        self.qty_files[key] = {"path": qty_hits[0],
                                               "years": qy, "records": qr}
            elif qty_hits:
                if not value_hits:
                    self.warnings.append(
                        "unsupported measure in %s file(s) in %s" %
                        (prefix, self.data_dir))

        # ---- region by-product downloads (e.g. Africa) --------------------
        # An ITC region reports its trade as a pseudo-economy, so a "List of
        # products exported/imported by Africa" download drives the template's
        # Africa sub-section tables.
        for flow, key in (("export", "africa_exports_by_product"),
                          ("import", "africa_imports_by_product")):
            path = _find_region_product_file(self.data_dir, flow)
            if path is None:
                continue
            years, records = _load_any(path)
            if records is None:
                continue
            if records and "product" not in records[0]:
                records = _classic_to_matrix(key, records)
            self.files[key] = {"path": path, "years": years,
                               "records": records}

        # ---- per-market supplying-markets downloads ------------------------
        # For the template's Competitor Analysis section (Kenya's market share
        # within its top destination markets) each market's own
        # "List of supplying markets for a product imported by <Market>"
        # download is mapped back to the destination label.
        self.suppliers = {}
        for path in _walk_files(self.data_dir):
            base = os.path.basename(path).lower()
            if ("kenyas" in base or
                    ("imports-from-world-by-exporter" not in base and
                     "list_of_supplying_markets" not in base)):
                continue
            label = _supplier_market_label(path)
            if not label:
                continue
            years, records = _load_any(path)
            if not records:
                continue
            self.suppliers[label] = {"path": path, "years": years,
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
        # Keep the tables compact: analyse only the most recent years, even
        # when the portfolio carries a decade of history (ITC data runs the
        # full profile window but a 10-year matrix no longer fits a page).
        full_years = self.all_years
        if self.max_years and len(self.all_years) > self.max_years:
            self.all_years = self.all_years[-self.max_years:]
        self._full_years = full_years

        if not self.all_years:
            loaded = ", ".join(sorted(self.files)) or "none"
            raise ValueError(
                "No yearly columns could be read from the uploaded matrices "
                "under %s (parsed: %s). Trade Map downloads must be Excel "
                "workbooks whose header row contains year columns such as "
                "'2024 (USD thousand)' - re-download the files from Trade Map "
                "and try again." % (self.data_dir, loaded))

        # Anchor product (the product of the by-importer download).  A missing
        # by-importer download must not abort the profile: it falls back to
        # the configured family so every other table still generates and the
        # data-availability note flags the file.
        partner_rows = self._rows("kenya_exports_by_partner")
        anchor_product = None
        if partner_rows:
            counts = {}
            for r in partner_rows:
                product = r.get("product") or r.get("code")
                if product:
                    counts[product] = counts.get(product, 0) + 1
            if counts:
                anchor_product = max(counts, key=counts.get)
        if anchor_product is None:
            candidate = self.include_codes[0] if self.include_codes else None
            if candidate is None:
                records = self._rows("kenya_exports_by_product")
                candidate = next((r.get("product") or r.get("code") or None
                                  for r in records), None)
            if candidate is not None:
                anchor_product = candidate
                self.warnings.append(
                    "by-importer file not found - anchor inferred from the "
                    "by-product download (%s)" % anchor_product)
        if anchor_product is None:
            raise OSError("Missing the by-importer file; cannot detect the "
                          "anchor product in %s." % self.data_dir)
        self.anchor_hs = anchor_product
        self.anchor_label = next(
            (r["product_label"] for r in partner_rows
             if (r.get("product") or r.get("code")) == anchor_product
             and r.get("product_label")), "")
        if not self.anchor_label and self.family_title:
            self.anchor_label = self.family_title

        # Total rows: a product code whose series equals the sum of all the
        # other codes in the same file (the basket aggregate of a selection
        # download). They are header rows, not family members.
        self._file_totals = {}
        for key in ("kenya_exports_by_product", "kenya_imports_by_product",
                    "world_exports_by_product", "world_imports_by_product"):
            if key in self.files:
                self._file_totals[key] = self._total_codes(key)

        # Family members: product detail rows of Kenya's by-product download
        # (its total row, if any, is excluded), filtered by include_codes.
        self.members = self._family_product_rows(self.files,
                                                 "kenya_exports_by_product")
        self.anchor_is_total = anchor_product in \
            self._file_totals.get("kenya_exports_by_product", set())
        if self.anchor_is_total and self.family_title:
            # The anchor is the ITC selection group (e.g. "coffee one"), whose
            # group name is not meaningful on its own.  Use the configured
            # family title ("Coffee") for headings and tables instead.
            self.anchor_label = self.family_title

        # ---- data availability manifest ------------------------------------
        # Presence of every file the reports draw on, plus human-readable
        # alerts. The web app shows these alerts as info/warning banners and
        # the profile prints them as a data-availability note, so a missing
        # file never silently blanks a table.
        self.file_status = {}
        self.alerts = []
        for spec in REQUIRED_UPLOADS:
            present = spec["key"] in self.files
            if not present:
                self.alerts.append({"level": "warning",
                                    "topic": spec["key"],
                                    "message": "Missing required download "
                                               "\"%s\": %s" %
                                               (spec["prefix"], spec["label"])})
            self.file_status[spec["key"]] = {"present": present,
                                             "required": True}
        for spec in OPTIONAL_UPLOADS:
            if spec.get("kind") == "supplier":
                present = bool(self.suppliers)
            elif spec.get("kind") == "potential":
                present = bool(self.potential_files)
            elif spec.get("measure") == "quantity":
                present = spec["key"] in self.qty_files
            else:
                present = spec["key"] in self.files
            self.file_status.setdefault(spec["key"], {})["present"] = present
            if not present:
                self.alerts.append({"level": "info", "topic": spec["key"],
                                    "message": "Optional download \"%s\" not "
                                               "provided (%s) - the related "
                                               "tables are skipped until it "
                                               "is added."
                                               % (spec["prefix"],
                                                  spec["note"])})

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

    def _total_codes(self, key, store=None):
        """Codes in ``key`` (or ``store[key]``) whose review-year value equals
        the sum of all the other codes in the same file (the selection total
        row)."""
        store = self.files if store is None else store
        by_code = {}
        for r in store.get(key, {}).get("records", []):
            if r.get("partner") != "000" or "product" not in r:
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
    def _rows(self, key, store=None):
        store = self.files if store is None else store
        return store.get(key, {}).get("records", [])

    def _family_product_rows(self, store, key):
        """Product-detail rows of a by-product download.

        ``store`` is the value or quantity pipeline; rows are filtered to the
        configured family, its selection-total row is dropped and the rest are
        ranked by their review-year value.  Also accepts the classic HTML
        ``.xls`` records (code/label/years) used for region downloads.
        """
        records = store.get(key, {}).get("records", [])
        if not records:
            return []
        total = self._total_codes(key, store) if "product" in records[0] \
            else set()
        rows = []
        for r in records:
            code = r.get("product") or r.get("code")
            if code is None:
                continue
            code = str(code)
            if "product" in r:
                label = r["product_label"]
                if r["partner"] != "000" or code in total:
                    continue
            else:
                label = r.get("label") or ""
                stem = code.upper()
                if stem in ("TOTAL", "WORLD") or \
                        str(label).lower().strip() in ("total", "world",
                                                       "all products", "all"):
                    continue
            if not self._code_ok(code):
                continue
            rows.append({"code": code, "label": label, "years": r["years"]})
        rows.sort(key=lambda m: m["years"].get(self.review_year) or 0.0,
                  reverse=True)
        return rows

    def sub_family_members(self, codes):
        """Members of ``self.members`` that fall under one of ``codes`` -- the
        sub-family (e.g. processed meat = HS 16) of the template's Table 13."""
        if not codes:
            return []
        wanted = [self._norm_code(c) for c in codes]
        out = []
        for m in self.members:
            c = self._norm_code(m["code"])
            if any(c.startswith(w) for w in wanted if w):
                out.append(m)
        return out

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
        return self._family_product_rows(self.files, "kenya_imports_by_product")

    def global_export_products(self):
        return self._family_product_rows(self.files, "world_exports_by_product")

    def global_import_products(self):
        return self._family_product_rows(self.files, "world_imports_by_product")

    def global_export_products_qty(self):
        """World exports of the family by product in Quantity (Tonnes)."""
        return self._family_product_rows(self.qty_files,
                                         "world_exports_by_product")

    def global_import_products_qty(self):
        """World imports of the family by product in Quantity (Tonnes)."""
        return self._family_product_rows(self.qty_files,
                                         "world_imports_by_product")

    def kenya_export_products_qty(self):
        """Kenya's exports of the family by product in Quantity (Tonnes)."""
        return self._family_product_rows(self.qty_files,
                                         "kenya_exports_by_product")

    def kenya_import_products_qty(self):
        """Kenya's imports of the family by product in Quantity (Tonnes)."""
        return self._family_product_rows(self.qty_files,
                                         "kenya_imports_by_product")

    def africa_exporters(self):
        """All African exporters of the family, ranked (template Table 7)."""
        return [e for e in self.exporters() if is_africa(e["label"])]

    def africa_importers(self):
        """All African importers of the family, ranked (template Table 9)."""
        return [i for i in self.importers() if is_africa(i["label"])]

    def africa_export_products(self):
        """Africa's exports of the family by product (template Table 8)."""
        return self._family_product_rows(self.files, "africa_exports_by_product")

    def africa_import_products(self):
        """Africa's imports of the family by product (template Table 10)."""
        return self._family_product_rows(self.files, "africa_imports_by_product")

    def market_suppliers(self):
        """Per-market importing-side data from the supplying-markets downloads.

        Returns ``{market_label: {"years", "total", "total_label", "rows"}}``
        where ``total`` is the market's total imports of the family by year and
        ``rows`` are the ranked supplier rows (Kenya included).  Markets whose
        download is not present are simply absent from the dict.
        """
        out = {}
        for label, d in self.suppliers.items():
            records = d["records"]
            total = {}
            total_year = None
            rows = []
            for r in records:
                if "product" in r:
                    # By-exporter matrix: the importing market is the reporter
                    # and the total (suppliers = World) row is partner '000'.
                    if r["partner"] == "000":
                        total = r["years"]
                        total_year = d["years"]
                    else:
                        rows.append({"label": fix_label(r["partner_label"]),
                                     "years": r["years"]})
                else:
                    stem = (r.get("code") or "").upper()
                    lab = str(r.get("label") or "")
                    if stem in ("TOTAL", "WORLD") or lab.lower().strip() in (
                            "world", "total"):
                        total = r["years"]
                        total_year = d["years"]
                    else:
                        rows.append({"label": fix_label(r.get("label") or ""),
                                     "years": r["years"]})
            rows.sort(key=lambda r: r["years"].get(self.review_year) or 0.0,
                      reverse=True)
            out[label] = {"years": d["years"], "total": total,
                          "total_year": total_year, "rows": rows,
                          "label": label}
        return out

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


def _qty_unit(values):
    """Display unit for a raw quantity (tonnes) series."""
    vs = [abs(v) for v in values if v]
    if not vs:
        return "Tonnes"
    if max(vs) >= 1_000_000.0:
        return "Thousand Tonnes"
    return "Tonnes"


def fmt_qty(v, unit):
    """Format a raw quantity (tonnes) value in ``unit``."""
    if v is None:
        return ""
    if unit == "Thousand Tonnes":
        return num(v / 1000.0, 1)
    return num(v, 1)


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
    items = list(names)
    if not items:
        return ""
    if len(items) == 1:
        return items[0]
    return sep.join(items[:-1]) + last + items[-1]


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
        if r.get("label") and r.get("code"):
            key = "%s (%s)" % (r["label"], r["code"])
        else:
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
        super().__init__(cfg, narratives or {}, cfg.get("template_doc"))
        self.cfg = cfg
        self.tcap = 0
        self.fcap = 0
        self.done_tables = set()

    # The source line must sit *below* the table / figure body.  These
    # methods therefore emit the caption only; every table / figure builder
    # calls ``add_source`` last, after the body has been written.  A mirror
    # table may supply an explicit ``num`` so its number never shifts when
    # another table's source download is missing.
    def _next_table(self, title, source, num=None):
        if num is not None:
            self.tcap = max(self.tcap, num)
        else:
            self.tcap += 1
        self.done_tables.add(self.tcap)
        self.add_table_caption("Table %d: %s" % (self.tcap, title))

    def _missing_required_table(self, num, title, source):
        # Keeps the mirror-table contract: every required table number
        # always appears in the report even when its source download was
        # not among the uploaded files.
        if num in self.done_tables:
            return
        self.done_tables.add(num)
        self.tcap = max(self.tcap, num)
        self.add_table_caption("Table %d: %s" % (num, title))
        p = self.doc.add_paragraph()
        r = p.add_run(
            "Table data not available for the uploaded files. "
            "Please supply the matching Trade Map download "
            "and re-run the report."
        )
        self._style_run(r, italic=True)
        p.paragraph_format.space_before = Pt(6)
        p.paragraph_format.space_after = Pt(6)
        self.add_source(source)

    def _next_figure(self, title, source):
        self.fcap += 1
        self.add_table_caption("Figure %d: %s" % (self.fcap, title))

    def add_value_table(self, first_col_header, rows, years, share_header,
                        title, source, total_label=None, rank=False,
                        unit_label="USD Million", widths=None,
                        adaptive_unit=True, quantity=False):
        """Structure a value matrix table.

        ``rows`` = list of ``{"label", "years": {y: v}}`` already truncated
        and consolidated. The share column shows each row's share of the
        total in the review year (bold); a ``total_label`` row (optional)
        sums the displayed rows.

        When *adaptive_unit* is ``True`` (default) the whole table is shown in
        USD Thousand (instead of Millions) if any nonzero value would
        otherwise round to 0.0 in millions, so small sub-category values
        never disappear.  With ``quantity=True`` the table instead measures a
        tons-series and uses ``Tonnes`` / ``Thousand Tonnes`` units.
        """
        n = len(years)
        code_cols = 1 if any(row.get("code") for row in rows) else 0
        label_cols = (1 if rank else 0) + code_cols + 1
        cols = label_cols + n + 1
        c_label = (1 if rank else 0) + code_cols
        nrows = 2 + len(rows) + (1 if total_label else 0)
        unit = unit_label
        if quantity:
            unit = _qty_unit([v for r in rows for y in years
                              if (v := r["years"].get(y)) is not None])
        elif adaptive_unit:
            unit = _series_unit([v for r in rows for y in years
                                 if (v := r["years"].get(y)) is not None])
        fmt = (lambda v: fmt_qty(v, unit)) if quantity else \
            (lambda v: fmt_for_unit(v, unit))
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
                    fmt(row["years"].get(y))
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
                t.cells[label_cols + i].text = fmt(tot)
            t.cells[label_cols + n].text = "100.0%"

        if widths is None:
            widths = ([420] if rank else []) \
                + ([1100] if code_cols else []) + [3000] + [700] * n + [1300]
        self._set_table_widths(table, widths)
        self._style_table(table, rank=rank, label_cols=label_cols, n=n,
                          total_label=total_label)
        self._fit_table_on_page(table)
        self.add_source(source)
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


def _add_share_donut(b, pairs, title, name, width_in=6.3, height_in=4.2,
                     hole_size=58):
    """Native, Word-editable doughnut for the share figures.

    Keeps the top 5 slices and folds the remainder into an ``Others`` slice,
    mirroring the static PNG doughnuts this replaces; ``add_word_chart``
    installs a real chart the reader can click and edit (Chart Design / Edit
    Data).  Returns the chart r-id, or ``None`` when there is nothing to plot.
    """
    pairs = [p for p in pairs if p[1] > 0.0]
    if not pairs:
        return None
    pairs.sort(key=lambda p: p[1], reverse=True)
    if len(pairs) > 5:
        top, rest = pairs[:5], pairs[5:]
        pairs = top + [("Others", sum(s for _, s in rest))]
    return b.add_word_chart(
        "doughnut", title, [l for l, _ in pairs], [s for _, s in pairs],
        width_in=width_in, height_in=height_in, name=name,
        colors=[c.lstrip("#") for c in THEME], hole_size=hole_size)


def family_members_line(data, top=12):
    members = data.members
    if len(members) <= 1:
        return ""
    by_code = {m["code"]: m["label"] for m in members}
    shown = members[:top]
    items = ["%s (%s)" % (m["code"], short_label(by_code[m["code"]], 100))
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
    b._next_table("Trend in %s: Kenya's Exports by Product" % family, source)
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
               "accounted for **%.1f%%** of Kenya's exports of the family."
               % (family.lower(), rev, short_label(lead[0], 100),
                  lead[1] * 100))
        if follows:
            txt += " It was followed by %s." % ordinal_list(
                ["%s (**%.1f%%**)" % (short_label(l, 100), s * 100)
                 for l, s in follows])
        b.add_bullet(txt)

    totals = _year_totals(members)
    for p_ in _growth_paragraphs(years, totals, "Kenya's total exports of",
                                 family.lower(), rows=members):
        b.add_bullet(p_)

    if len(members) >= 2:
        _add_share_donut(
            b, pairs, "Share of Kenya's Exports of %s by Product, %d"
            % (family, rev), "Share of %s by Product" % family)
        b._next_figure("Share of Kenya's Exports of %s by Product, %d"
                       % (family, rev), source)
        b.add_source(source)


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

    if total_last:
        b.add_bullet("Kenya's exports of %s were valued at **%s** in **%d**."
                     % (anchor.lower(), usd_phrase(total_last), rev))
    for p_ in _growth_paragraphs(years, totals, "Kenya's exports of",
                                 anchor.lower(), rows=destinations):
        b.add_bullet(p_)

    first = destinations[0] if destinations else None
    if first and first["years"].get(rev):
        share = (first["years"].get(rev) or 0.0) / (total_last or 1.0) * 100
        b.add_bullet("The leading destination for Kenya's %s exports was %s "
                     "(**%s**; **%.1f%%** of Kenya's exports of the product)."
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
                "(**%s**; **%.1f%%** of Kenya's exports of the product)"
                % (anchor.lower(), lead_afr["label"],
                   usd_phrase(rev_val),
                   rev_val / (total_last or 1.0) * 100))
        if rank:
            text += (", ranked %s among all destination markets"
                     % _ordinal(rank, bold=True))
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
        _add_share_donut(
            b, pairs,
            "Kenya's %s Exports by Destination, %d" % (anchor, rev),
            "Kenya's %s Exports by Destination" % anchor)
        b._next_figure("Kenya's %s Exports by Destination, %d"
                       % (anchor, rev), source)
        b.add_source(source)


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
               "(**%.2f%%** in %d; **%.2f%%** in %d)."
               % (anchor.lower(),
                  "rose" if last_s >= first_s else "fell",
                  first_y, last_y,
                  first_s * 100, first_y, last_s * 100, last_y))
        if last_s >= first_s:
            txt += " Kenya exported **%s** of %s in %d." % (
                usd_phrase(last_k), family.lower(), last_y)
        parts.append(txt)

    # -- Export specialization --------------------------------------------
    spec = data.specialization_metrics()
    if spec and spec["years"]:
        speclast = spec["years"][-1]
        fam_share = speclast["share"]
        txt = ("%s accounted for **%.2f%%** of Kenya's total merchandise "
               "exports in **%d** (**%s** of **%s**), reflecting Kenya's "
               "export specialization."
               % (family.title(), fam_share * 100, speclast["year"],
                  usd_phrase(speclast["family"]),
                  usd_phrase(speclast["total"])))
        if len(spec["years"]) >= 2:
            first_s = spec["years"][0]["share"]
            txt += " This share was **%.2f%%** in %d." % (
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
                label = " %s was the runner-up with **%.1f%%**." % (
                    short_label(second["label"], 100), second_share)
            else:
                label = " %s was the runner-up with **%.1f%%**; the remaining " \
                    "%d headings together accounted for **%.1f%%**." % (
                        short_label(second["label"], 100), second_share,
                        len(members) - 2,
                        max(0.0, 100 - lead_share - second_share))
        concentration = ("highly concentrated" if hhi >= 2500 else
                         "moderately concentrated" if hhi >= 1500 else
                         "diversified")
        txt = ("Kenya's exports of %s are %s across **%d** product headings: "
               "the leading heading (%s) accounted for **%.1f%%** of the "
               "family total in %d.%s"
               % (family.lower(), concentration, len(members),
                  short_label(lead["label"], 100), lead_share, rev, label))
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

    b.add_heading("GLOBAL EXPORTS OF %s" % anchor.upper(), level=1)

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
            _add_share_donut(
                b, pairs, "World Exports of %s by Economy, %d" % (anchor, rev),
                "World Exports of %s by Economy" % anchor)
            b._next_figure("World Exports of %s by Economy, %d"
                           % (anchor, rev), source)
            b.add_source(source)

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
            _add_share_donut(
                b, pairs, "World Imports of %s by Economy, %d" % (anchor, rev),
                "World Imports of %s by Economy" % anchor)
            b._next_figure("World Imports of %s by Economy, %d"
                           % (anchor, rev), source)
            b.add_source(source)

    g_exp = top_rows(data.global_export_products(),
                     cfg.get("top_n", 10), years, "All other products")
    if g_exp:
        b._next_table("Trend in %s Globally - Export" % family, source)
        b.add_value_table("Product", g_exp, years, "Share in %d" % rev,
                          "Global Exports of %s by Product" % family, source,
                          total_label="Total", adaptive_unit=True)
        trend_bullets(b, g_exp, years, family, "exports")

    g_imp = top_rows(data.global_import_products(),
                     cfg.get("top_n", 10), years, "All other products")
    if g_imp:
        b._next_table("Trend in %s Globally - Import" % family, source)
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
        parts.append("%s among **%d** %s"
                     % (_ordinal(global_rank, bold=True), n_total, group))
    if africa_rank and n_africa:
        if len(parts) > 1:
            parts.append("and")
        parts.append("%s among **%d** %s"
                     % (_ordinal(africa_rank, bold=True), n_africa,
                        africa_group))
    if len(parts) > 1:
        b.add_bullet(" ".join(parts) + " of %s in **%d**, with %s valued at "
                     "**%s**." % (anchor, rev, action, usd_phrase(value)))


RESIDUE = {"exporter": "All other economies",
           "importer": "All other economies",
           "source": "All other sources"}


def geo_bullets(b, rows, years, anchor_short, role, region=None):
    """'Brazil was the world's leading exporter of X in 2023...' bullets.
    ``region`` ("Africa") re-frames the leader phrase and denominator to the
    region instead of the world."""
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
    subject = "%s in Africa" % anchor_short.lower() if region else \
        anchor_short.lower()
    if role == "source":
        b.add_bullet("The leading source of Kenya's imports of %s in %d was "
                     "%s (**%s**; **%.1f%%** of Kenya's imports)."
                     % (anchor_short.lower(), last, real[0][0],
                        usd_phrase(next((r["years"].get(last) for r in rows
                                         if r["label"] == real[0][0]), None)),
                        real[0][1] * 100))
        denom = "Kenya's imports in %d" % last
    elif region:
        b.add_bullet("%s was the leading %s of %s in %d, with %s "
                     "(**%.1f%%** of Africa's total)."
                     % (real[0][0], word, subject, last,
                        usd_phrase(next((r["years"].get(last) for r in rows
                                         if r["label"] == real[0][0]), None)),
                        real[0][1] * 100))
        denom = "Africa's total in %d" % last
    else:
        b.add_bullet("%s was the world's leading %s of %s in %d, with %s "
                     "(**%.1f%%** of the world total)."
                     % (real[0][0], word, anchor_short.lower(), last,
                        usd_phrase(next((r["years"].get(last) for r in rows
                                         if r["label"] == real[0][0]), None)),
                        real[0][1] * 100))
        denom = "the world total in %d" % last
    names = ["%s (**%s**; **%.1f%%**)" % (l, usd_phrase(next(
        (r["years"].get(last) for r in rows if r["label"] == l), None)), s * 100)
        for l, s in real[:5]]
    b.add_bullet("The top five %s were %s." % (group, ordinal_list(names)))
    top5 = sum(s for _, s in real[:5]) * 100
    b.add_bullet("Together, the top five %s accounted for **%.1f%%** of %s."
                 % (group, top5, denom))
    if residual_share > 0:
        b.add_bullet("%s together accounted for **%.1f%%** of %s."
                     % (residue, residual_share * 100, denom))


def trend_bullets(b, rows, years, family, noun, scope="World",
                  residual="All other products", denom=None):
    last = years[-1]
    pairs = _shares(rows, years)
    pairs.sort(key=lambda p: p[1], reverse=True)
    if not pairs:
        return
    real = [p for p in pairs if p[0] != residual]
    if not real:
        return
    subj = ("%s %s" % (scope, noun)).strip()
    if denom is None:
        denom = ("%s total" % scope if scope.endswith("'s")
                 else "the world total")
    lead, share = real[0]
    txt = ("%s of %s in %d were led by %s (**%.1f%%** of %s)"
           % (subj, family.lower(), last, short_label(lead, 100), share * 100,
              denom))
    follows = real[1:3]
    if follows:
        txt += ", followed by %s" % ordinal_list(
            ["%s (**%.1f%%**)" % (short_label(l, 100), s * 100)
             for l, s in follows])
    b.add_bullet(txt + ".")
    lead_share = sum(s for _, s in real)
    b.add_bullet("Together, the leading product headings accounted for "
                 "**%.1f%%** of %s." % (lead_share * 100, denom))
    totals = _year_totals(rows)
    for p_ in _growth_paragraphs(years, totals, "%s of" % subj,
                                 family.lower(), rows=rows):
        b.add_bullet(p_)


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
            "label": short_label(r["label"], 120),
            "years": {"Kenya": k, "World": w},
            "code": r["code"],
            "share": (k / w * 100.0) if w else None,
        })
    _kenya_world_table(b, display_rows, rev, k_tot, w_tot, source,
                       first_col="HS code")

    # -- bullet point on Kenya's share and position ----------------------
    overall = (k_tot / w_tot * 100.0) if w_tot else None
    if overall is not None:
        b.add_bullet("Kenya accounted for **%.1f%%** of world exports of %s "
                     "in **%d** (**%s** of **%s**)."
                     % (overall, family.lower(), rev, usd_phrase(k_tot),
                        usd_phrase(w_tot)))
    if rows:
        row0 = rows[0]
        code = row0["code"]
        k, w = row0["kenya"].get(rev) or 0.0, row0["world"].get(rev) or 0.0
        fam_share = (k / k_tot * 100.0) if k_tot else None
        world_share = (k / w * 100.0) if w else None
        txt = ("Kenya's leading export heading of %s in %d was %s, valued "
               "at **%s**."
               % (family.lower(), rev, product_ref(row0["label"], code),
                  usd_phrase(k)))
        if fam_share is not None:
            txt += (" That heading accounted for **%.1f%%** of Kenya's "
                    "exports of the family" % fam_share)
            if world_share is not None:
                txt += (" and **%.1f%%** of world exports of the heading."
                        % world_share)
            else:
                txt += "."
        elif world_share is not None:
            txt += (" That heading accounted for **%.1f%%** of world exports "
                    "of the heading." % world_share)
        b.add_bullet(txt)
    if metrics:
        gr = metrics.get("global_rank")
        nr = metrics.get("n_exporters")
        ar = metrics.get("africa_rank")
        na = metrics.get("n_africa")
        parts = ["Kenya ranked"]
        if gr and nr:
            parts.append("%s among **%d** world exporters of %s"
                         % (_ordinal(gr, bold=True), nr, family.lower()))
        if ar and na:
            if len(parts) > 1:
                parts.append("and")
            parts.append("%s among **%d** African exporters"
                         % (_ordinal(ar, bold=True), na))
        if len(parts) > 1:
            b.add_bullet(" ".join(parts) + " in **%d**." % rev)


def _ordinal(n, bold=False):
    """Ordinal with a superscript suffix for Word: 34 -> '34^{th}'.

    ``bold`` wraps the numeral in ``**...**`` so the figure reads bold in
    ranking sentences; the suffix stays superscript either way.
    """
    suffix = "th" if 10 <= n % 100 <= 20 else \
        {1: "st", 2: "nd", 3: "rd"}.get(n % 10, "th")
    if bold:
        return "**%d**^{%s}" % (n, suffix)
    return "%d^{%s}" % (n, suffix)


def product_ref(label, code=None, maxlen=100):
    """Product reference for narrative text: full-enough label plus code.

    Labels are kept long enough that HS descriptions are never left cut off
    mid-clause ('...of a kind suitable'), while the six-digit code is shown in
    parentheses so products are always identifiable.
    """
    name = short_label(label, maxlen)
    if code:
        return "%s (%s)" % (name, code)
    return name


def _unit_mult(unit):
    """Display multiplier for an ITC raw (USD-thousand) value in ``unit``."""
    return 0.001 if unit == "USD Million" else 1.0


def _share_suffix(share_cf):
    """Excel formula fragment correcting a cross-column unit mismatch, e.g.
    Kenya in USD Thousand vs World in USD Million."""
    return "" if share_cf == 1.0 else "*%.4f" % share_cf


def _decline_cause(rows, year, years):
    """Attribute a yearly decline in a total to its biggest negative driver.

    Returns ``None`` or a prose fragment such as 'a 12.4% fall in Fresh cut
    roses & buds (060311)' for the row whose USD drop was the largest that
    year.
    """
    if not rows or year not in years:
        return None
    idx = years.index(year)
    if idx < 1:
        return None
    base_year = years[idx - 1]
    drops = []
    for r in rows:
        bv = r["years"].get(base_year) or 0.0
        cv = r["years"].get(year) or 0.0
        if bv > 0 and cv < bv:
            drops.append((cv - bv, bv, r))
    if not drops:
        return None
    drops.sort(key=lambda t: t[0])
    drop, bv, r = drops[0]
    name = short_label(r.get("label") or r.get("code") or "", 96)
    code = r.get("code")
    label = "%s (%s)" % (name, code) if code else name
    return "a %.1f%% fall in %s" % (abs(drop) / bv * 100.0, label)


def _growth_paragraphs(years, totals, subject, family, rows=None):
    """Granular, year-on-year growth narrative bullets.

    Reports the average annual rate and then each year's year-on-year change,
    so a mid-period decline is no longer hidden inside a single CAGR figure.
    When ``rows`` is supplied, every declining year is attributed to its
    biggest contributor (the largest negative driver in the data).
    """
    paras = []
    if len(years) < 2:
        return paras
    g = cagr([totals.get(y) for y in years], years)
    if g is not None:
        verb = "grew" if g >= 0 else "contracted"
        paras.append("%s %s %s at an average annual rate of **%.1f%%** %s."
                     % (subject, family, verb, abs(g) * 100.0,
                        period_phrase(years[0], years[-1])))
    yoy = []
    for i in range(1, len(years)):
        prev, cur = totals.get(years[i - 1]), totals.get(years[i])
        if prev and cur:
            yoy.append((years[i], cur / prev - 1.0))
    if yoy:
        seq = ", ".join("**%d** %s%.1f%%"
                        % (y, "+" if c >= 0 else "", abs(c) * 100.0)
                        for y, c in yoy)
        detail = "Year on year, the path was: %s." % seq
        declines = [(y, c) for y, c in yoy if c < 0]
        for y, c in declines[:3]:
            cause = _decline_cause(rows, y, years) if rows else None
            if cause:
                detail += " The decline in %d mainly reflected %s." % (y, cause)
        paras.append(detail)
    return paras


def _margin_headline(m, family, start, rev, noun, dimension):
    net = m["net"]
    direction = "grew" if net >= 0 else "shrank"
    parts = ["Between %d and %d, Kenya's exports of %s %s by %s from %s to "
             "%s, a net change of **%s**."
             % (start, rev, family.lower(), direction, dimension,
                usd_phrase(m["total_start"]), usd_phrase(m["total_rev"]),
                usd_phrase(abs(net)))]
    if net:
        if noun == "product heading":
            enter_txt = "%s new to the export basket" % noun
            exit_txt = "%s no longer exported" % noun
        else:
            enter_txt = "destination markets newly served"
            exit_txt = "destination markets no longer served"
        parts.append(
            "**%.0f%%** of that change came from existing %s deepening their "
            "sales, **%.0f%%** from %s, and **%.0f%%** was offset by %s."
            % (m["intensive_share"] * 100, noun, m["entering_share"] * 100,
               enter_txt, abs(m["exiting_share"]) * 100, exit_txt))
    return " ".join(parts)


def _kenya_world_table(b, rows, year, kenya_total, world_total, source,
                       first_col="HS code"):
    """Table with a Kenya / World value pair per row plus overall totals.

    Columns: HS code | Product | Kenya exports | World exports | Kenya share
    of world.  The code and the product description are separate columns so a
    heading is never presented as a single run-together 'code + label' cell.
    Each value column picks its own display unit (USD Million / USD Thousand)
    so that small sub-category values never round to 0.0 in millions; the
    share is computed from the raw (USD-thousand) values so it is unit-safe.
    """
    k_unit = _series_unit([r["years"].get("Kenya") for r in rows]
                          + [kenya_total or 0.0])
    w_unit = _series_unit([r["years"].get("World") for r in rows]
                          + [world_total or 0.0])
    table = b.doc.add_table(rows=2 + len(rows) + 1, cols=5, style="Table Grid")
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    hdr = table.rows[0]
    hdr.cells[0].text = "Code"
    hdr.cells[1].text = "Product"
    hdr.cells[2].text = "Kenya exports (%s)" % k_unit
    hdr.cells[3].text = "World exports (%s)" % w_unit
    hdr.cells[4].text = "Kenya share of world"
    r2 = table.rows[1]
    r2.cells[2].text = str(year)
    r2.cells[3].text = str(year)
    r2.cells[4].text = str(year)
    for ri, r in enumerate(rows, start=2):
        table.rows[ri].cells[0].text = str(r.get("code") or "")
        table.rows[ri].cells[1].text = r["label"]
        table.rows[ri].cells[2].text = fmt_for_unit(
            r["years"].get("Kenya"), k_unit)
        table.rows[ri].cells[3].text = fmt_for_unit(
            r["years"].get("World"), w_unit)
        table.rows[ri].cells[4].text = (
            "%.1f%%" % r["share"] if r["share"] is not None else "")
    t = table.rows[2 + len(rows)]
    t.cells[0].text = ""
    t.cells[1].text = "Total, %s" % first_col
    t.cells[2].text = fmt_for_unit(kenya_total if kenya_total else None,
                                   k_unit)
    t.cells[3].text = fmt_for_unit(world_total if world_total else None,
                                   w_unit)
    t.cells[4].text = (
        "%.1f%%" % (kenya_total / world_total * 100.0) if world_total else "")
    b._set_table_widths(table, [900, 3000, 1500, 1500, 1500])
    b._style_table(table, rank=False, label_cols=2, n=2, total_label=True)
    b._fit_table_on_page(table)
    b.add_source(source)


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


def section_potential(b, cfg, data, pot, source, tmp_dir=None, _figure=None):
    """Optional export-potential section (ITC Export Potential Map file).

    Shows the largest unrealised opportunities: each market's potential
    exports vs the actual base embedded in the download, ranked by the gap.
    When ``_figure`` is set the block is framed as a Figure (template's
    Figures 3-6) with a share doughnut before the ranking table.
    """
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
    if _figure is None:
        b.add_heading("EXPORT POTENTIAL FOR %s" % family.upper(), level=1)
    title = "Export Potential for %s" % family
    if gap_known and _figure:
        totals = {r["label"]: (r["gap"] if r.get("gap") is not None
                               else r["actual"] or 0.0) for r in top}
        total = sum(totals.values())
        pairs = [(lab, (v / total if total else 0.0))
                 for lab, v in totals.items() if v > 0.0]
        img = make_donut(pairs, tmp_dir, "pot_%d.png" % (_figure or 0), title)
        if img:
            b._next_figure(title, source)
            b.add_figure(img)
    b.add_para("Markets ranked by the gap between Kenya's current exports and "
               "the potential demand estimated by the ITC Export Potential Map "
               "(projection year %s)."
               % (pot_year if pot_year else "unknown"))
    cols = 3 if gap_known else 2
    b._next_table("Unrealised export potential for %s, by destination market"
                  % family, source)
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
    b.add_source(source)
    b.add_bullet("Together, these **%d** markets offer the greatest headroom "
                 "for additional exports of %s." % (len(top), family.lower()))
    if gap_known:
        lead = gap_known[0]
        b.add_bullet("The largest unrealised opportunity is estimated to be "
                     "%s: potential exports of **%s** versus current exports "
                     "of **%s**."
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
                                      "product heading", "product heading"))
    if dest_m and dest_m["total_rev"]:
        b.add_bullet(_margin_headline(dest_m, family, start, rev,
                                      "destination market",
                                      "destination market"))
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
            names = ordinal_list(short_label(t[0], 90)
                                 for t in m["top_entering"][:3])
            names = names if names else ""
            if noun == "product heading":
                b.add_bullet("**%d** product heading(s) entered Kenya's "
                             "export basket after **%d** (led by **%s**)."
                             % (m["n_entering"], start, names))
            else:
                b.add_bullet("**%d** destination market(s) were newly served "
                             "after **%d** (the largest: **%s**)."
                             % (m["n_entering"], start, names))
        if m["top_exiting"]:
            names = ordinal_list(short_label(t[0], 90)
                                 for t in m["top_exiting"][:3])
            names = names if names else ""
            if noun == "product heading":
                b.add_bullet("**%d** product heading(s) ceased being exported "
                             "after **%d** (the largest were **%s**)."
                             % (m["n_exiting"], start, names))
            else:
                b.add_bullet("**%d** destination market(s) were no longer "
                             "served in **%d** (the largest were **%s**)."
                             % (m["n_exiting"], rev, names))


def _margin_table(b, title, m, source, noun):
    b._next_table(title, source)
    values = [m["intensive"], m["entering"], m["exiting"]]
    unit = _series_unit([abs(v) for v in values] + [abs(m["net"])])
    if noun == "product heading":
        enter_label = "Product heading(s) new to the export basket"
        exit_label = "Product heading(s) no longer exported"
    else:
        enter_label = "Destination market(s) newly served"
        exit_label = "Destination market(s) no longer served"
    rows = [("Existing %s(s) deepening" % noun, m["intensive"],
             m["intensive_share"]),
            (enter_label, m["entering"], m["entering_share"]),
            (exit_label, m["exiting"], m["exiting_share"]),
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
    b.add_source(source)


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
    b._next_table("Attractiveness ranking of Kenya's destination markets, %d"
                  % rev, source)
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
    b.add_source(source)
    leaders = ordinal_list([r["label"] for r in top[:3]])
    b.add_bullet("Priority markets for the next phase of export growth: "
                 "**%s**. %s tops the attractiveness ranking on current "
                 "momentum and headroom." % (leaders, top[0]["label"]))
    fast = [r for r in top if r["growth"] and r["growth"] > 0]
    if len(fast) >= 2:
        b.add_bullet("Fast-growth destinations worth deepening: **%s**."
                     % ordinal_list([r["label"] for r in fast[:3]]))
    big_gaps = [r for r in top if r["gap"] is not None]
    if len(big_gaps) >= 2:
        b.add_bullet("Markets with the largest unrealised headroom: **%s**."
                     % ordinal_list([r["label"] for r in big_gaps[:3]]))
    b.add_bullet("Focus sectors should be aligned with these ranked markets "
                 "and the preferential routes open to Kenya's exports of %s."
                 % family.lower())
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
        b.add_source(source)
    krow = next((r for r in sc["rows"] if r["label"] == "Kenya"), None)
    if krow and abs(krow["delta_pp"]) >= 0.05:
        b.add_bullet("Kenya's share of world exports %s by **%.1f** percentage "
                     "points between %d and %d (**%.1f%%** to **%.1f%%**)."
                     % ("rose" if krow["delta_pp"] > 0 else "fell",
                        abs(krow["delta_pp"]), start, rev,
                        krow["start_share"] * 100, krow["rev_share"] * 100))
    if gainers:
        b.add_bullet("The main share-gainers were %s."
                     % ordinal_list([r["label"] for r in gainers[:3]]))


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
        if g is not None:
            growth = ("grew at **%.1f%%** CAGR between %d and %d"
                      % (g * 100, years[0], rev) if g > 0 else
                      "declined at **%.1f%%** CAGR between %d and %d"
                      % (abs(g) * 100, years[0], rev))
        else:
            growth = "showed no clear trend over %d-%d" % (years[0], rev)
        action = []
        if priority:
            action.append("prioritise %s" % ordinal_list(priority))
        if i == 1:
            action.append("anchor heading: defend and deepen")
        bullet = "%d. %s (%s): **%.1f%%** of Kenya's exports of %s in **%d**; %s." \
            % (i, short_label(label, 100), code, share, family.lower(),
               rev, growth)
        if action:
            bullet += " Recommended action: %s." % "; ".join(action)
        b.add_bullet(bullet)
    if top_dest and dest_total:
        dshare = (top_dest["years"].get(rev) or 0.0) / dest_total * 100
        if dshare >= 15.0:
            b.add_bullet("Risk to manage: **%.1f%%** of Kenya's %s exports go "
                         "to a single destination (%s); deepen a second market."
                         % (dshare, family.lower(), top_dest["label"]))
    if att:
        b.add_bullet("Net takeaway: concentrate promotion on **%s** where "
                     "momentum, headroom and preferential access reinforce "
                     "one another."
                     % ordinal_list([r["label"] for r in att[:3]]))


# --------------------------------------------------------------------------
# Template-mirror sections (Meat Product Profile August-2026 structure)
# --------------------------------------------------------------------------
def section_background(b, cfg, data, source):
    """1.0 BACKGROUND."""
    family = cfg.get("family_title", "the product family")
    b.add_heading("1.0 BACKGROUND", level=1)
    for par in cfg.get("intro", []):
        b.add_para(par)
    g = data.kenya_global_metrics()
    rev = data.review_year
    if g:
        b.add_para(
            "Kenya, the country profiled in this report, ranked %s among the "
            "world's exporters of %s in %d, with %s of the world total "
            "(%s) and a %s position among African exporters."
            % (("No. %s" % _ordinal(g["global_rank"])) if g["global_rank"]
               else "-",
               family.lower(), rev,
               pct(g["share"]) if g["share"] is not None else "-",
               usd_phrase(g["value"]),
               ("No. %s of %d African exporters"
                % (_ordinal(g["africa_rank"]), g["n_africa"]))
               if g["africa_rank"] else "-"))
    line = family_members_line(data)
    if line:
        b.add_para(line)

    # Data-availability note: reflect every missing download so it never
    # silently blanks a table. Optional gaps are info-level; a required file
    # missing ends the note with a warning so the reader re-uploads it.
    alerts = [a for a in data.alerts if a["level"] == "warning"] + \
        [a for a in data.alerts if a["level"] == "info"]
    if alerts:
        b.add_para("")
        b.add_para("Data availability. This profile is built from the ITC "
                   "files uploaded for \"%s\" and covers %d to %d (the most "
                   "recent %d years of %d available). Where a download is "
                   "missing, the related table is skipped and the analysis "
                   "continues with what is available." %
                   (family, min(data.years), max(data.years), len(data.years),
                    len(data._full_years)), italic=True)
        for a in alerts:
            b.add_bullet("%s: %s" %
                         ("Missing required file" if a["level"] == "warning"
                          else "Info", a["message"]))


def _mirror_global(b, cfg, data, source, tmp_dir):
    """2.0 GLOBAL <FAMILY> SECTOR: tables 1-6 plus the Africa sub-section
    (tables 7-10) and Figures 1-2."""
    family = cfg.get("family_title", "the product family")
    prod = (family.lower() + " products")
    top = cfg.get("global_top_n", 25)
    years = data.years
    rev = data.review_year

    b.add_heading("2.0 GLOBAL %s SECTOR" % family.upper(), level=1)
    for par in cfg.get("global_intro") or []:
        b.add_para(par)

    # Table 1 / Figure 1 - world exports
    exporters = _ranked_rows(data.exporters(), top, years,
                             ensure_label="Kenya")
    if exporters:
        b._next_table("Global exports of %s by top %d countries"
                      % (prod, top), source, num=1)
        b.add_value_table("Exporting economy", exporters, years,
                          "Share in %d" % rev,
                          "Exports of %s by Economy" % family, source,
                          total_label="World", rank=True)
        geo_bullets(b, exporters, years, family, "exporter")
        img = make_donut(_shares(exporters, years), tmp_dir, "g1_share.png",
                         "Share of world exports")
        if img:
            b._next_figure("Share of World %s Exports, %d" % (family, rev),
                           source)
            b.add_figure(img)

    # Table 2 - world exports by product (value)
    g_exp = top_rows(data.global_export_products(), top, years,
                     "All other products")
    if g_exp:
        b._next_table("Global exports of %s by values" % prod, source, num=2)
        b.add_value_table("Product", g_exp, years, "Share in %d" % rev,
                          "Exports of %s by Product" % family, source,
                          total_label="World")
        trend_bullets(b, g_exp, years, family, "exports")

    # Table 3 - world exports by product (quantity)
    g_exp_q = top_rows(data.global_export_products_qty(), top, years,
                       "All other products")
    if g_exp_q:
        b._next_table("Global exports of %s by quantity (tonnes)" % prod,
                      source, num=3)
        b.add_value_table("Product", g_exp_q, years, "Share in %d" % rev,
                          "Exports of %s by Product (Tonnes)" % family, source,
                          total_label="World", quantity=True)
        trend_bullets(b, g_exp_q, years, family, "exports")

    # Table 4 / Figure 2 - world imports
    importers = _ranked_rows(data.importers(), top, years,
                             ensure_label="Kenya")
    if importers:
        b._next_table("Global imports of %s by top %d countries"
                      % (prod, top), source, num=4)
        b.add_value_table("Importing economy", importers, years,
                          "Share in %d" % rev,
                          "Imports of %s by Economy" % family, source,
                          total_label="World", rank=True)
        geo_bullets(b, importers, years, family, "importer")
        img = make_donut(_shares(importers, years), tmp_dir, "g2_share.png",
                         "Share of world imports")
        if img:
            b._next_figure("Share of World %s Imports, %d" % (family, rev),
                           source)
            b.add_figure(img)

    # Table 5 - world imports by product (value)
    g_imp = top_rows(data.global_import_products(), top, years,
                     "All other products")
    if g_imp:
        b._next_table("Global imports of %s by values" % prod, source, num=5)
        b.add_value_table("Product", g_imp, years, "Share in %d" % rev,
                          "Imports of %s by Product" % family, source,
                          total_label="World")
        trend_bullets(b, g_imp, years, family, "imports")

    # Table 6 - world imports by product (quantity)
    g_imp_q = top_rows(data.global_import_products_qty(), top, years,
                       "All other products")
    if g_imp_q:
        b._next_table("Global imports of %s by quantity (tonnes)" % prod,
                      source, num=6)
        b.add_value_table("Product", g_imp_q, years, "Share in %d" % rev,
                          "Imports of %s by Product (Tonnes)" % family, source,
                          total_label="World", quantity=True)
        trend_bullets(b, g_imp_q, years, family, "imports")

    # ---- Global exports from Africa ---------------------------------------
    b.page_break()
    b.add_heading("Global %s Exports from Africa" % family, level=2)
    for par in cfg.get("africa_intro") or []:
        b.add_para(par)

    africa_exp = _ranked_rows(data.africa_exporters(), top, years,
                              ensure_label="Kenya")
    if africa_exp:  # Table 7
        b._next_table("Global exports of %s by top %d African countries"
                      % (prod, top), source, num=7)
        b.add_value_table("African exporting economy", africa_exp, years,
                          "Share in %d" % rev,
                          "Exports of %s from Africa" % family, source,
                          total_label="Total", rank=True)
        geo_bullets(b, africa_exp, years, family.lower(), "exporter",
                    region="Africa")

    africa_exp_p = top_rows(data.africa_export_products(), top, years,
                            "All other products")
    if africa_exp_p:  # Table 8
        b._next_table("Africa exports of %s by values" % prod, source, num=8)
        b.add_value_table("Product", africa_exp_p, years, "Share in %d" % rev,
                          "Exports of %s from Africa by Product" % family,
                          source, total_label="Total")
        trend_bullets(b, africa_exp_p, years, family, "exports",
                      scope="Africa's", residual="All other products",
                      denom="Africa's total")

    africa_imp = _ranked_rows(data.africa_importers(), top, years,
                              ensure_label="Kenya")
    if africa_imp:  # Table 9
        b._next_table("Africa's top %d importers of %s" % (top, prod),
                      source, num=9)
        b.add_value_table("African importing economy", africa_imp, years,
                          "Share in %d" % rev,
                          "Imports of %s into Africa" % family, source,
                          total_label="Total", rank=True)
        geo_bullets(b, africa_imp, years, family.lower(), "importer",
                    region="Africa")

    africa_imp_p = top_rows(data.africa_import_products(), top, years,
                            "All other products")
    if africa_imp_p:  # Table 10
        b._next_table("Africa imports of %s by values" % prod, source, num=10)
        b.add_value_table("Product", africa_imp_p, years, "Share in %d" % rev,
                          "Imports of %s into Africa by Product" % family,
                          source, total_label="Total")
        trend_bullets(b, africa_imp_p, years, family, "imports",
                      scope="Africa's", residual="All other products",
                      denom="Africa's total")

    if not (africa_exp or africa_imp or africa_exp_p or africa_imp_p):
        b.add_bullet("Africa-level download(s) not available; the Africa "
                     "sub-section is limited to the by-economy tables above.")

    # Keep the mirror contract: every required table number must appear (with
    # an explicit notice) even if its source download was not uploaded.
    required = {
        1: "Global exports of %s by top %d countries" % (prod, top),
        2: "Global exports of %s by values" % prod,
        3: "Global exports of %s by quantity (tonnes)" % prod,
        4: "Global imports of %s by top %d countries" % (prod, top),
        5: "Global imports of %s by values" % prod,
        6: "Global imports of %s by quantity (tonnes)" % prod,
        7: "Global exports of %s by top %d African countries" % (prod, top),
        8: "Africa exports of %s by values" % prod,
        9: "Africa's top %d importers of %s" % (top, prod),
        10: "Africa imports of %s by values" % prod,
    }
    for _n, _t in sorted(required.items()):
        if _n not in b.done_tables:
            b._missing_required_table(_n, _t, source)


def section_mirror_kenya(b, cfg, data, source):
    """3.0 KENYA'S EXPORTS AND IMPORTS MARKET TRENDS: tables 11-17 plus the
    sub-family table 13."""
    family = cfg.get("family_title", "the product family")
    prod = (family.lower() + " products")
    top = cfg.get("top_n", 10)
    years = data.years
    rev = data.review_year

    b.add_heading("3.0 KENYA'S %s EXPORTS AND IMPORTS MARKET TRENDS"
                  % family.upper(), level=1)
    for par in cfg.get("kenya_intro") or []:
        b.add_para(par)

    # Table 11 - Kenya's exports by destination
    dests = _ranked_rows(data.destinations(), cfg.get("global_top_n", 25),
                         years)
    if dests:
        b._next_table("Kenya's exports of %s by country" % prod, source, num=11)
        b.add_value_table("Destination market", dests, years,
                          "Share in %d" % rev,
                          "Kenya's %s Exports by Destination" % family,
                          source, total_label="Total", rank=True)
        lead = next((d for d in dests if d["label"] != "All other markets"),
                    None)
        if lead:
            lead_share = _shares(dests, years)
            top_share = next((s for l, s in lead_share
                              if l == lead["label"]), 0.0)
            b.add_bullet("The leading destination of Kenya's %s exports in %d "
                         "was %s, absorbing %s (%.1f%% of the total)."
                         % (family.lower(), rev, lead["label"],
                            usd_phrase(lead["years"].get(rev)),
                            top_share * 100.0))
        for p_ in _growth_paragraphs(years, _year_totals(dests),
                                     "Kenya's exports of", family.lower()):
            b.add_bullet(p_)

    # Table 12 - Kenya's exports by product
    members = top_rows(data.members, top, years, "All other products")
    if members:
        b._next_table("Kenya's exports of %s by values" % prod, source, num=12)
        b.add_value_table("Product", members, years, "Share in %d" % rev,
                          "Kenya's %s Exports by Product" % family, source,
                          total_label="Total", adaptive_unit=True)
        trend_bullets(b, members, years, family, "exports",
                      scope="Kenya's", residual="All other products",
                      denom="Kenya's total")

    # Table 13 - Kenya's exports of the sub-family (e.g. processed meat)
    for i, sf in enumerate(cfg.get("sub_families") or []):
        title = sf.get("title") or family
        rows = data.sub_family_members(sf.get("codes") or [])
        rows = top_rows(rows, top, years, "All other products")
        if not rows:
            continue
        b._next_table("Kenya's Exports of %s by Product" % title, source, num=13)
        b.add_value_table("Product", rows, years, "Share in %d" % rev,
                          "Kenya's Exports of %s by Product" % title, source,
                          total_label="Total", adaptive_unit=True)
        lead = next((r for r in rows if r["label"] != "All other products"),
                    None)
        if lead and i == 0:
            share = ((lead["years"].get(rev) or 0.0) /
                     sum((r["years"].get(rev) or 0.0) for r in rows)
                     if rows else None)
            b.add_bullet("Within %s, the leading export product heading is %s "
                         "(%.1f%% of Kenya's exports of the sub-family in %d)."
                         % (title.lower(), short_label(lead["label"], 60),
                            (share or 0.0) * 100, rev))

    # Table 14 - Kenya's exports by product (quantity)
    ken_q = top_rows(data.kenya_export_products_qty(), top, years,
                     "All other products")
    if ken_q:
        b._next_table("Kenya's exports of %s by quantities (tonnes)" % prod,
                      source, num=14)
        b.add_value_table("Product", ken_q, years, "Share in %d" % rev,
                          "Quantity of Kenya's %s Exports by Product (Tonnes)"
                          % family, source, total_label="Total",
                          quantity=True)
        trend_bullets(b, ken_q, years, family, "exports",
                      scope="Kenya's", residual="All other products",
                      denom="Kenya's total")

    # Table 15 - Kenya's imports by source
    sources = _ranked_rows(data.kenya_import_sources(),
                           cfg.get("global_top_n", 25), years)
    if sources:
        b._next_table("Kenya's imports of %s by country" % prod, source, num=15)
        b.add_value_table("Supplying market", sources, years,
                          "Share in %d" % rev,
                          "Kenya's %s Imports by Source" % family, source,
                          total_label="Total", rank=True)
        geo_bullets(b, sources, years, family, "source")

    # Table 16 - Kenya's imports by product (value)
    imports = top_rows(data.kenya_import_products(), top, years,
                       "All other products")
    if imports:
        b._next_table("Kenya's imports of %s by values" % prod, source, num=16)
        b.add_value_table("Product", imports, years, "Share in %d" % rev,
                          "Kenya's %s Imports by Product" % family, source,
                          total_label="Total", adaptive_unit=True)
        trend_bullets(b, imports, years, family, "imports",
                      scope="Kenya's", residual="All other products",
                      denom="Kenya's total")

    # Table 17 - Kenya's imports by product (quantity)
    imp_q = top_rows(data.kenya_import_products_qty(), top, years,
                     "All other products")
    if imp_q:
        b._next_table("Kenya's imports of %s by quantities (tonnes)" % prod,
                      source, num=17)
        b.add_value_table("Product", imp_q, years, "Share in %d" % rev,
                          "Quantity of Kenya's %s Imports by Product (Tonnes)"
                          % family, source, total_label="Total",
                          quantity=True)
        trend_bullets(b, imp_q, years, family, "imports",
                      scope="Kenya's", residual="All other products",
                      denom="Kenya's total")

    # Keep the mirror contract: every required table number must appear (with
    # an explicit notice) even if its source download was not uploaded.
    required = {
        11: "Kenya's exports of %s by country" % prod,
        12: "Kenya's exports of %s by values" % prod,
        14: "Kenya's exports of %s by quantities (tonnes)" % prod,
        15: "Kenya's imports of %s by country" % prod,
        16: "Kenya's imports of %s by values" % prod,
        17: "Kenya's imports of %s by quantities (tonnes)" % prod,
    }
    if cfg.get("sub_families"):
        required[13] = "Kenya's Exports of %s by Product" % (
            cfg["sub_families"][0].get("title") or family)
    for _n, _t in sorted(required.items()):
        if _n not in b.done_tables:
            b._missing_required_table(_n, _t, source)


def section_mirror_potential(b, cfg, data, source, tmp_dir):
    """3.1 EXPORT POTENTIAL: one figure + ranking table per Export Potential
    Map download (the template's Figures 3-6)."""
    family = cfg.get("family_title", "the product family")
    b.add_heading("EXPORT POTENTIAL (EXPANSION SEEKING)", level=1)
    for i, pot in enumerate(data.potential_files):
        section_potential(b, cfg, data, pot, source, tmp_dir, _figure=(i + 1))
    if not data.potential_files:
        b.add_bullet("An Export Potential Map download is not available; the "
                     "section is limited to the tables above.")


def _competitor_rows(data, n=5):
    """Kenya's leading destinations joined with each market's own
    supplying-markets download.

    Returns rows with ``market``, ``kenya_exports`` (Kenya's exports to the
    market in the review year), ``market_imports`` (the market's total imports
    of the family), ``share`` (Kenya's share of those imports) and
    ``leader`` (the largest supplying economy other than Kenya) for every
    destination that has a matching download.
    """
    rev = data.review_year
    suppliers = {k: v for k, v in data.market_suppliers().items()}
    rows = []
    for d in data.destinations()[:n]:
        key = display_name(fix_label(d["label"]))
        s = suppliers.get(key)
        if s is None:
            continue
        kenya_val = d["years"].get(rev)
        total = s["total"].get(rev)
        if total is None and s["rows"]:
            total = sum((r["years"].get(rev) or 0.0) for r in s["rows"])
        kenya_row = next((r for r in s["rows"]
                          if data._is_kenya_label(r["label"])), None)
        kenya_in = (kenya_row["years"].get(rev) or 0.0) if kenya_row else None
        share = None
        if total:
            base = kenya_in if kenya_in is not None else kenya_val
            if base is not None:
                share = base / total
        others = [r for r in s["rows"] if not data._is_kenya_label(r["label"])]
        leader = None
        if others:
            leader = max(others, key=lambda r: r["years"].get(rev) or 0.0)
        rows.append({"market": d["label"], "kenya_exports": kenya_val,
                     "market_imports": total, "share": share,
                     "leader": leader})
    return rows


def section_mirror_competitor(b, cfg, data, source):
    """4.0 COMPETITOR ANALYSIS - Kenya's share inside its top destination
    markets (the template's Table 18)."""
    family = cfg.get("family_title", "the product family")
    years = data.years
    rev = data.review_year
    b.add_heading("4.0 COMPETITOR ANALYSIS", level=1)
    rows = _competitor_rows(data, cfg.get("competitor_top", 5))
    if not rows:
        top_markets = [d["label"] for d in
                       data.destinations()[:cfg.get("competitor_top", 5)]]
        b.add_bullet(
            "The per-market supplying-markets downloads needed for this "
            "table are not yet available. Download the \"List of supplying "
            "markets for a product imported by ...\" file for each of "
            "Kenya's top destinations (%s), drop them into the data folder "
            "and re-run the profile to populate the competitor analysis."
            % (" / ".join(top_markets) if top_markets
               else "once destination data is loaded"))
        return
    b.add_para("The table below benchmarks Kenya against its competitors "
               "within the destination markets that absorb the largest value "
               "of Kenya's %s exports: Kenya's own exports to each market, "
               "the market's total imports of the family, Kenya's resulting "
               "share of that market, and the market's leading supplier."
               % family.lower())
    unit = _series_unit([r["kenya_exports"] or 0.0 for r in rows]
                        + [r["market_imports"] or 0.0 for r in rows])
    table = b.doc.add_table(rows=1 + len(rows), cols=5, style="Table Grid")
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    hdr = table.rows[0]
    for c, t in enumerate(("Market",
                           "Kenya's exports (%s)" % unit,
                           "Market's total imports (%s)" % unit,
                           "Kenya's share of market",
                           "Leading competitor")):
        hdr.cells[c].text = t
    for i, r in enumerate(rows, start=1):
        row = table.rows[i]
        row.cells[0].text = r["market"]
        row.cells[1].text = fmt_for_unit(r["kenya_exports"], unit)
        row.cells[2].text = fmt_for_unit(r["market_imports"], unit)
        row.cells[3].text = ("" if r["share"] is None
                             else "%.1f%%" % (r["share"] * 100.0))
        lead = r["leader"]
        row.cells[4].text = (short_label(lead["label"], 45)
                             if lead and (lead["years"].get(rev) or 0.0) > 0
                             else "")
    b._set_table_widths(table, [2400, 2200, 2400, 1900, 2200])
    b._style_table(table, rank=False, label_cols=1, n=4,
                   total_label=None)
    b._fit_table_on_page(table)

    dated = [r for r in rows if r["share"] is not None]
    if dated:
        dated.sort(key=lambda r: r["share"], reverse=True)
        top_r = dated[0]
        b.add_bullet("Kenya commands the largest share of %s within %s "
                     "(%.1f%% of that market's %s imports in %d).%s"
                     % (family.lower(), top_r["market"],
                        top_r["share"] * 100.0, family.lower(), rev,
                        (" Its closest rival is %s."
                         % short_label(top_r["leader"]["label"], 45)
                         if top_r["leader"] else "")))
        if len(dated) > 1:
            second = dated[1]
            b.add_bullet("By comparison, Kenya holds only %.1f%% of the %s "
                         "market - a market that imports %s of the family - "
                         "indicating room to grow."
                         % (second["share"] * 100.0, second["market"],
                            usd_phrase(second["market_imports"])))


def section_mirror_policy(b, cfg, source):
    """5.0 POLICY RECOMMENDATIONS + references."""
    b.add_heading("5.0 POLICY RECOMMENDATIONS", level=1)
    recs = cfg.get("recommendations") or []
    if not recs:
        b.add_para("No policy recommendations were supplied in the "
                   "configuration for this profile.")
    for r in recs:
        b.add_bullet(r)
    b.page_break()
    b.add_heading("References", level=1)
    refs = cfg.get("references") or [
        "International Trade Centre, Trade Map database (source: %s)."
        % source]
    for r in refs:
        b.add_bullet(r)


# --------------------------------------------------------------------------
# Top-level build
# --------------------------------------------------------------------------
def build_profile_document(cfg, data, tmp_dir):
    os.makedirs(tmp_dir, exist_ok=True)
    source = cfg.get("source", "International Trade Centre Database")
    b = ProfileBuilder(cfg, {})
    b.title_page(cfg)

    # Template structure: Background -> Global (Africa sub) -> Kenya exports
    # and imports -> Export Potential -> Competitor Analysis -> Policy.
    section_background(b, cfg, data, source)
    _mirror_global(b, cfg, data, source, tmp_dir)
    section_mirror_kenya(b, cfg, data, source)
    section_mirror_potential(b, cfg, data, source, tmp_dir)
    section_mirror_competitor(b, cfg, data, source)
    section_mirror_policy(b, cfg, source)

    # Optional deep-dive sections kept behind a config flag.
    if cfg.get("include_analysis"):
        section_kenya_global_position(b, cfg, data, source)
        section_growth_decomposition(b, cfg, data, source)
        section_market_attractiveness(b, cfg, data, source,
                                      data.potential_files[0]
                                      if data.potential_files else None)
        section_competitiveness(b, cfg, data, source)
        section_global(b, cfg, data, source, tmp_dir)
        section_competitor_watch(b, cfg, data, source)
        if cfg.get("include_imports", True):
            section_kenya_imports(b, cfg, data, source)
        for pot in data.potential_files:
            section_potential(b, cfg, data, pot, source, None)
        section_strategy(b, cfg, data, source,
                         data.potential_files[0] if data.potential_files
                         else None)
    return b.doc


# --------------------------------------------------------------------------
# Excel deliverable
# --------------------------------------------------------------------------
def _xc(ws, r, c, value, bold=False, number_format=None, fill=None,
        align=None, color=None):
    cell = ws.cell(r, c, value)
    cell.font = Font(name="Century Gothic", bold=bold,
                     color=color if color else None)
    cell.border = BORDER
    if number_format:
        cell.number_format = number_format
    if fill:
        cell.fill = fill
    if align:
        cell.alignment = align
    return cell


def _excel_data_notes(wb, cfg, data):
    """Data Notes sheet: the review window and which uploaded files are in
    use, so analysts can see at a glance what a missing download would have
    filled."""
    ws = wb.create_sheet("Data Notes")
    hdr_fill = PatternFill("solid", fgColor="1F4E79")
    cm = Alignment(horizontal="center", vertical="center")
    lm = Alignment(horizontal="left", vertical="center", wrap_text=True)
    _xc(ws, 1, 1, "PRODUCT PROFILE - DATA NOTES", bold=True,
        fill=hdr_fill, color="FFFFFF", align=lm)
    _xc(ws, 2, 1, "Family", bold=True); _xc(ws, 2, 2,
        cfg.get("family_title", ""), align=lm)
    _xc(ws, 3, 1, "Review period", bold=True)
    _xc(ws, 3, 2, "%d to %d (%d of %d available years used)"
        % (min(data.years), max(data.years), len(data.years),
           len(data._full_years)), align=lm)
    _xc(ws, 4, 1, "Included HS codes", bold=True)
    _xc(ws, 4, 2, ", ".join(data.include_codes) or "(all)", align=lm)
    _xc(ws, 5, 1, "Products analysed", bold=True)
    _xc(ws, 5, 2, "%d product detail rows" % len(data.members), align=lm)
    _xc(ws, 7, 1, "UPLOADED FILES", bold=True, fill=hdr_fill,
        color="FFFFFF")
    _xc(ws, 8, 1, "File (prefix)", bold=True, fill=hdr_fill, color="FFFFFF")
    _xc(ws, 8, 2, "Required", bold=True, fill=hdr_fill, color="FFFFFF")
    _xc(ws, 8, 3, "Present", bold=True, fill=hdr_fill, color="FFFFFF")
    _xc(ws, 8, 4, "Fills", bold=True, fill=hdr_fill, color="FFFFFF")
    row = 9
    for spec in REQUIRED_UPLOADS:
        st = data.file_status.get(spec["key"], {})
        _xc(ws, row, 1, spec["prefix"] + "_<product>.xlsx", align=lm)
        _xc(ws, row, 2, "yes", align=cm)
        _xc(ws, row, 3, "yes" if st.get("present") else "no", align=cm)
        _xc(ws, row, 4, spec["label"], align=lm)
        row += 1
    for spec in OPTIONAL_UPLOADS:
        st = data.file_status.get(spec["key"], {})
        _xc(ws, row, 1, spec["prefix"] + "_<product>.xlsx", align=lm)
        _xc(ws, row, 2, "no", align=cm)
        _xc(ws, row, 3, "yes" if st.get("present") else "no", align=cm)
        _xc(ws, row, 4, spec["label"], align=lm)
        row += 1
    row += 1
    _xc(ws, row, 1, "ALERTS", bold=True, fill=hdr_fill, color="FFFFFF")
    row += 1
    _xc(ws, row, 1, "Level", bold=True, fill=hdr_fill, color="FFFFFF")
    _xc(ws, row, 2, "Message", bold=True, fill=hdr_fill, color="FFFFFF")
    row += 1
    for a in data.alerts:
        _xc(ws, row, 1, a["level"], align=cm)
        _xc(ws, row, 2, a["message"], align=lm)
        row += 1
    ws.column_dimensions["A"].width = 46
    ws.column_dimensions["B"].width = 12
    ws.column_dimensions["C"].width = 12
    ws.column_dimensions["D"].width = 78
    return ws


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

    def value_sheet(name, first_col, rows, doughnut=None, unit=None,
                    quantity=False):
        ws = wb.create_sheet(name)
        if unit is None:
            if quantity:
                unit = _qty_unit([v for r in rows for y in years
                                  if (v := r["years"].get(y)) is not None])
            else:
                unit = _series_unit([v for r in rows for y in years
                                     if (v := r["years"].get(y)) is not None])
        has_code = any(r.get("code") for r in rows)
        first = 1 + (1 if has_code else 0)
        if has_code:
            _xc(ws, 1, 1, "Code", bold=True, fill=hdr_fill, align=cm)
        hdr_col = "%s (%s)" % (first_col, unit)
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
                if quantity:
                    v_out = None if raw is None else round(raw, 1)
                elif unit == "USD Thousand":
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
    top = cfg.get("global_top_n", 25)
    # Template mirror: Table 1..18 sheets in the template's exact order.
    t1 = _ranked_rows(data.exporters(), top, years, ensure_label="Kenya",
                      residual="All other economies")
    if t1:
        value_sheet("Table 1 - Top 25 Global Exporters", "Exporting economy", t1,
                    doughnut=("Share of world exports",
                              [(d["label"], _share01(d, rev, t1))
                               for d in t1]))
    t2 = top_rows(data.global_export_products(), top, years,
                  "All other products")
    if t2:
        value_sheet("Table 2 - Global Exports by Value", "Product", t2)
    t3 = top_rows(data.global_export_products_qty(), top, years,
                  "All other products")
    if t3:
        value_sheet("Table 3 - Global Exports by Quantity (Tonnes)", "Product", t3,
                    quantity=True)
    t4 = _ranked_rows(data.importers(), top, years, ensure_label="Kenya",
                      residual="All other economies")
    if t4:
        value_sheet("Table 4 - Top 25 Global Importers", "Importing economy", t4,
                    doughnut=("Share of world imports",
                              [(d["label"], _share01(d, rev, t4))
                               for d in t4]))
    t5 = top_rows(data.global_import_products(), top, years,
                  "All other products")
    if t5:
        value_sheet("Table 5 - Global Imports by Value", "Product", t5)
    t6 = top_rows(data.global_import_products_qty(), top, years,
                  "All other products")
    if t6:
        value_sheet("Table 6 - Global Imports by Quantity (Tonnes)", "Product", t6,
                    quantity=True)
    t7 = _ranked_rows(data.africa_exporters(), top, years,
                      ensure_label="Kenya", residual="All other economies")
    if t7:
        value_sheet("Table 7 - Top 25 African Exporters",
                    "African exporting economy", t7)
    t8 = top_rows(data.africa_export_products(), top, years,
                  "All other products")
    if t8:
        value_sheet("Table 8 - Africa Exports by Value", "Product", t8)
    t9 = _ranked_rows(data.africa_importers(), top, years,
                      ensure_label="Kenya", residual="All other economies")
    if t9:
        value_sheet("Table 9 - Top 25 African Importers",
                    "African importing economy", t9)
    t10 = top_rows(data.africa_import_products(), top, years,
                   "All other products")
    if t10:
        value_sheet("Table 10 - Africa Imports by Value", "Product", t10)
    t11 = _ranked_rows(data.destinations(), top, years,
                       residual="All other markets")
    if t11:
        value_sheet("Table 11 - Kenya Exports by Country", "Country", t11,
                    doughnut=("Kenya's exports by destination",
                              [(d["label"], _share01(d, rev, t11))
                               for d in t11]))
    t12 = top_rows(data.members, cfg.get("top_n", 10), years,
                   "All other products")
    if t12:
        value_sheet("Table 12 - Kenya Exports by Value", "Product", t12)
    t13 = None
    for sf in cfg.get("sub_families") or []:
        rows = data.sub_family_members(sf.get("codes") or [])
        rows = top_rows(rows, cfg.get("top_n", 10), years,
                        "All other products")
        if not rows:
            continue
        t13 = rows
        value_sheet("Table 13 - Kenya %s Exports" % (sf.get("title") or "Sub")
                    [:31], "Product", rows)
    t14 = top_rows(data.kenya_export_products_qty(), cfg.get("top_n", 10),
                   years, "All other products")
    if t14:
        value_sheet("Table 14 - Kenya Exports by Quantity (Tonnes)",
                    "Product", t14, quantity=True)
    t15 = _ranked_rows(data.kenya_import_sources(), top, years,
                       residual="All other markets")
    if t15:
        value_sheet("Table 15 - Kenya Imports by Country", "Country", t15)
    t16 = top_rows(data.kenya_import_products(), cfg.get("top_n", 10), years,
                   "All other products")
    if t16:
        value_sheet("Table 16 - Kenya Imports by Value", "Product", t16)
    t17 = top_rows(data.kenya_import_products_qty(), cfg.get("top_n", 10),
                   years, "All other products")
    if t17:
        value_sheet("Table 17 - Kenya Imports by Quantity (Tonnes)",
                    "Product", t17, quantity=True)
    comp = _competitor_rows(data, cfg.get("competitor_top", 5))
    if comp:
        ws = wb.create_sheet("Table 18 - Competitor Analysis")
        cu = _series_unit([r["kenya_exports"] or 0.0 for r in comp]
                          + [r["market_imports"] or 0.0 for r in comp])
        headers = ["Market", "Kenya exports (%s)" % cu,
                   "Market total imports (%s)" % cu,
                   "Kenya share of market", "Leading competitor"]
        for c, h in enumerate(headers, 1):
            _xc(ws, 1, c, h, bold=True, fill=hdr_fill, align=cm)
        for i, r in enumerate(comp, start=2):
            _xc(ws, i, 1, r["market"], align=lm)
            for j, k in ((2, "kenya_exports"), (3, "market_imports")):
                raw = r[k]
                v = None if raw is None else (
                    round(raw, 1) if cu == "USD Thousand"
                    else round(display(raw), 1))
                _xc(ws, i, j, v, number_format=val_fmt, align=cm)
            _xc(ws, i, 4, r["share"] if r["share"] is not None else None,
                number_format="0.0%", bold=True, align=cm)
            lead = r["leader"]
            _xc(ws, i, 5, short_label(lead["label"], 45) if lead else "",
                align=lm)
        ws.column_dimensions["A"].width = 32
        ws.column_dimensions["E"].width = 45

    # Keep the mirror contract in the workbook too: a placeholder sheet is
    # added for every required Table 1..17 whose source download was missing.
    required_xl = {
        "Table 1 - Top 25 Global Exporters",
        "Table 2 - Global Exports by Value",
        "Table 3 - Global Exports by Quantity (Tonnes)",
        "Table 4 - Top 25 Global Importers",
        "Table 5 - Global Imports by Value",
        "Table 6 - Global Imports by Quantity (Tonnes)",
        "Table 7 - Top 25 African Exporters",
        "Table 8 - Africa Exports by Value",
        "Table 9 - Top 25 African Importers",
        "Table 10 - Africa Imports by Value",
        "Table 11 - Kenya Exports by Country",
        "Table 12 - Kenya Exports by Value",
        "Table 14 - Kenya Exports by Quantity (Tonnes)",
        "Table 15 - Kenya Imports by Country",
        "Table 16 - Kenya Imports by Value",
        "Table 17 - Kenya Imports by Quantity (Tonnes)",
    }
    if cfg.get("sub_families"):
        required_xl.add(("Table 13 - Kenya %s Exports"
                         % (cfg["sub_families"][0].get("title") or "Sub"))[:31])
    for _name in sorted(required_xl):
        if _name in wb.sheetnames:
            continue
        ws = wb.create_sheet(_name)
        _xc(ws, 1, 1,
            "Table data not available for the uploaded files. Please supply "
            "the matching Trade Map download and re-run the report.",
            bold=True)
        ws.column_dimensions["A"].width = 120

    # Existing deliverable sheets (analysis extras kept below the mirror;
    # the simple value tables above are the mirror's Table 1..18 and are no
    # longer duplicated here).
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
        share_cf = _unit_mult(w_unit) / _unit_mult(k_unit)
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
            _xc(ws, ri, 5,
                "=C%d/D%d%s" % (ri, ri, _share_suffix(share_cf)),
                bold=True, number_format="0.0%", align=cm)
        ti = ri + 1
        _xc(ws, ti, 2, "Total", bold=True)
        for c, letter in ((3, "C"), (4, "D")):
            _xc(ws, ti, c, "=SUM(%s2:%s%d)" % (letter, letter, ri),
                bold=True, number_format=val_fmt, align=cm)
        _xc(ws, ti, 5,
            "=C%d/D%d%s" % (ti, ti, _share_suffix(share_cf)),
            bold=True, number_format="0.0%", align=cm)
        ws.column_dimensions["A"].width = 12
        ws.column_dimensions["B"].width = 60

    # Kenya vs leading African exporters
    peers = data.african_peers(5)
    if peers:
        value_sheet("African Peers", "Exporting economy", peers)

    # Kenya's share of world exports over time + specialization
    share_series = data.market_share_series()
    spec = data.specialization_metrics()
    ss_unit = _series_unit(
        [m[1] for m in (share_series or [])] + [m[2] for m in (share_series or [])])
    if share_series or (spec and spec["years"]):
        ws = wb.create_sheet(sheet_name("Kenya Standing"))
        _xc(ws, 1, 1, "Year", bold=True, fill=hdr_fill, align=cm)
        if share_series:
            _xc(ws, 1, 2, "Kenya exports (%s)" % ss_unit, bold=True,
                fill=hdr_fill, align=cm)
            _xc(ws, 1, 3, "World exports (%s)" % ss_unit, bold=True,
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
    pot = data.potential_files[0] if data.potential_files else None
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
            headers = ["Margin", "Exports %d (USD Million)" % start_year,
                       "Exports %d (USD Million)" % rev, "Change",
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
        unit = _series_unit([x["value_review"] for x in att])
        gap_any = any(x["has_potential"] and x["gap"] is not None
                      for x in att)
        headers = ["Rank", "Market", "Kenya exports %d (%s)" % (rev, unit),
                   "Share",
                   "CAGR %d-%d" % (start_year, rev), "Access tier"] \
            + (["Potential gap"] if gap_any else []) \
            + ["Attractiveness score"]
        for c, h in enumerate(headers, 1):
            _xc(ws, 1, c, h, bold=True, fill=hdr_fill, align=cm)
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
        headers = ["Product", "Exports %d (USD Million)" % rev,
                   "Current share",
                   "Target share (editable)",
                   "Target exports %d (USD Million)" % rev,
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

    _excel_data_notes(wb, cfg, data)

    # default "Sheet" removed by the first create_sheet call? keep membership
    if "Sheet" in wb.sheetnames and len(wb.sheetnames) > 1:
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
    ap.add_argument("--manifest", action="store_true",
                    help="Print the JSON upload checklist and exit")
    args = ap.parse_args()

    cfg = json.load(open(args.config, encoding="utf-8"))

    if args.manifest:
        print(json.dumps(upload_manifest(cfg.get("anchor", "<product>")),
                         indent=2))
        return

    data_dir = args.excel_dir or os.path.join(
        BASE_DIR, cfg.get("data_dir", "Coffee"))
    out = args.output
    if not out:
        out = os.path.join(BASE_DIR, cfg.get("output",
                                             "output/Product Profile.docx"))
    out = os.path.abspath(out)

    data = ProfileData(data_dir, cfg.get("include_codes"),
                       cfg.get("family_title"),
                       max_years=cfg.get("max_years"))
    print("[1/4] Loading ITC files from      : %s" % data_dir)
    print("      family    = %s" % cfg.get("family_title"))
    print("      anchor    = %s (%s)" % (data.anchor_hs, data.anchor_label))
    print("      period    = %d - %d (of %d available)"
          % (data.start_year, data.review_year, len(data._full_years)))
    print("      members   = %s (product detail rows)"
          % len(data.members))
    for w in data.warnings:
        print("      [warn] %s" % w)
    for a in data.alerts:
        print("      [%s] %s" % (a["level"], a["message"]))

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