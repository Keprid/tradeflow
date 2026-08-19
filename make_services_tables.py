#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
make_services_tables.py
=======================

Turn the raw ITC (International Trade Centre) services Excel downloads
into styled Table 1-4 + Figure 1 workbooks that the services report
generator (``generate_services_report.py``) expects.

Raw source files expected in ``--excel-dir`` (located by keywords in the
filename, case-insensitive):

    List_of_exported_services_for_the_selected_service_*.xls   -> global service exports by category
    List_of_exporters_for_the_selected_service_*.xls           -> global top exporters by country
    List_of_imported_services_for_the_selected_service_*.xls   -> global service imports by category
    List_of_importers_for_the_selected_service_*.xls           -> global top importers by country
    List_of_services_exported_by_Kenya_*.xls                   -> Kenya service exports by category
    Data manipulation:
    * Source values are in USD Thousand.
    * Global tables (1, 2): divide by 1,000,000 -> "Value in USD Billion".
    * Kenya tables (3, 4):  divide by 1,000     -> "Value in USD Million".
    * Figure 1 (balance):   divide by 1,000     -> "Value in USD Million".
"""

import argparse
import os
import re
import sys
import zipfile

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import numpy as np

import openpyxl
from lxml import etree
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from html.parser import HTMLParser

# ---------------------------------------------------------------------------
# Constants / styles (shared with make_tables.py)
# ---------------------------------------------------------------------------
FONT = "Times New Roman"

TOP_EXPORTERS = 25

HDR_FILL = PatternFill(fill_type="solid", fgColor="5D7B9D")
KENYA_FILL = PatternFill(fill_type="solid", fgColor="FFF2CC")
BAND_FILL = PatternFill(fill_type="solid", fgColor="F7F6F3")
DEV_FILL = PatternFill(fill_type="solid", fgColor="D6E4F0")
DEVEL_FILL = PatternFill(fill_type="solid", fgColor="E2EFDA")

THIN = Side(style="thin", color="808080")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

FMT_VALUE = "0.0"
FMT_SHARE = "0.0%"
FMT_GROWTH = "0.0%"

BALANCE_WIDTHS = {"A": 16, "B": 30, "C": 12, "D": 12, "E": 12, "F": 12,
                  "G": 12, "H": 12, "I": 12, "J": 12, "K": 12}

# ---------------------------------------------------------------------------
# Development status classification (HDI-based / World Bank / UN LDC)
# ---------------------------------------------------------------------------
CLASSIFICATION_PATH = os.path.join(os.path.dirname(__file__),
                                   "services", "CLASS_2026_07_15.xlsx")
HDI_PATH = os.path.join(os.path.dirname(__file__),
                        "services", "devcountries-2026.csv")

# HDI thresholds (UNDP standard)
HDI_DEVELOPED_THRESHOLD = 0.800

# UN Least Developed Countries (December 2024)
LDC_COUNTRIES = {
    "angola", "benin", "burkina faso", "burundi", "central african republic",
    "chad", "comoros", "congo, democratic republic of the", "djibouti",
    "eritrea", "ethiopia", "gambia", "guinea", "guinea-bissau", "lesotho",
    "liberia", "madagascar", "malawi", "mali", "mauritania", "mozambique",
    "niger", "rwanda", "senegal", "sierra leone", "somalia", "south sudan",
    "sudan", "togo", "uganda", "united republic of tanzania", "zambia",
    "afghanistan", "bangladesh", "cambodia", "lao people's democratic republic",
    "myanmar", "nepal", "timor-leste", "yemen",
    "haiti",
    "kiribati", "solomon islands", "tuvalu",
}

# LDCs that have graduated to developing status
LDC_GRADUATED = {
    "botswana", "cabo verde", "maldives", "samoa",
    "equatorial guinea", "vanuatu", "bhutan", "sao tome and principe",
}

# ITC name -> HDI / World Bank name aliases (for mismatches between data sources)
NAME_ALIASES = {
    "czech republic": "czechia",
    "hong kong, china": "hong kong",
    "macao, china": "macao sar, china",
    "iran, islamic republic of": "iran",
    "tanzania, united republic of": "tanzania",
    "macedonia, north": "north macedonia",
    "moldova, republic of": "moldova",
    "bolivia, plurinational state of": "bolivia",
    "palestine, state of": "palestine",
    "venezuela, bolivarian republic of": "venezuela",
    "libya, state of": "libya",
    "congo, democratic republic of the": "dr congo",
    "congo": "republic of the congo",
    "korea, republic of": "south korea",
    "russian federation": "russia",
    "türkiye": "turkey",
    "viet nam": "vietnam",
    "lao people's democratic republic": "lao pdr",
    "anguilla": None,
    "montserrat": None,
    "cook islands": None,
    "turks and caicos islands": None,
    "sint maarten (dutch part)": None,
    "cayman islands": None,
    "bermuda": None,
    "curaçao": None,
    "aruba": None,
    "french polynesia": None,
    "new caledonia": None,
    "faroe islands": None,
    "marshall islands": None,
}

# Non-country aggregation entries to exclude from dev-status ranking
NON_COUNTRY_ENTRIES = {
    "european union nes",
    "chinese taipei",
}

# ITC Regional classifications (based on tradebriefs.intracen.org)
ITC_REGIONS = {
    "Africa": {
        "algeria", "angola", "benin", "botswana", "burkina faso", "burundi",
        "cameroon", "cape verde", "central african republic", "chad", "comoros",
        "congo", "congo, democratic republic of the", "cote d'ivoire",
        "djibouti", "egypt", "equatorial guinea", "eritrea", "eswatini",
        "ethiopia", "gabon", "gambia", "ghana", "guinea", "guinea-bissau",
        "kenya", "lesotho", "liberia", "libya", "madagascar", "malawi",
        "mali", "mauritania", "mauritius", "morocco", "mozambique",
        "namibia", "niger", "nigeria", "rwanda", "sao tome and principe",
        "senegal", "seychelles", "sierra leone", "somalia", "south africa",
        "south sudan", "sudan", "tanzania", "togo", "tunisia", "uganda",
        "zambia", "zimbabwe",
    },
    "Asia": {
        "afghanistan", "armenia", "azerbaijan", "bahrain", "bangladesh",
        "bhutan", "brunei", "cambodia", "china", "cyprus", "georgia",
        "hong kong", "india", "indonesia", "iran", "iraq", "israel",
        "japan", "jordan", "kazakhstan", "korea", "kuwait", "kyrgyzstan",
        "lao people's democratic republic", "lebanon", "macao",
        "malaysia", "maldives", "mongolia", "myanmar", "nepal",
        "oman", "pakistan", "palestine", "philippines", "qatar",
        "russia", "saudi arabia", "singapore", "sri lanka", "syria",
        "taiwan", "tajikistan", "thailand", "timor-leste", "turkey",
        "türkiye", "turkmenistan", "united arab emirates", "uzbekistan",
        "viet nam", "vietnam", "yemen",
    },
    "Americas": {
        "antigua and barbuda", "argentina", "bahamas", "barbados",
        "belize", "bolivia", "brazil", "canada", "chile", "colombia",
        "costa rica", "cuba", "dominica", "dominican republic",
        "ecuador", "el salvador", "grenada", "guatemala", "guyana",
        "haiti", "honduras", "jamaica", "mexico", "nicaragua",
        "panama", "paraguay", "peru", "saint kitts and nevis",
        "saint lucia", "saint vincent and the grenadines",
        "suriname", "trinidad and tobago", "united states",
        "united states of america", "uruguay", "venezuela",
    },
    "Pacific": {
        "australia", "fiji", "kiribati", "marshall islands",
        "micronesia", "nauru", "new zealand", "papua new guinea",
        "samoa", "solomon islands", "tonga", "tuvalu", "vanuatu",
    },
    "Europe": {
        "albania", "belarus", "belgium", "bosnia and herzegovina",
        "bulgaria", "croatia", "czech republic", "czechia", "denmark",
        "estonia", "finland", "france", "germany", "greece", "hungary",
        "iceland", "ireland", "italy", "latvia", "lithuania",
        "luxembourg", "malta", "montenegro", "netherlands",
        "north macedonia", "norway", "poland", "portugal", "romania",
        "serbia", "slovakia", "slovenia", "spain", "sweden",
        "switzerland", "ukraine", "united kingdom",
    },
}


def load_hdi_data(path=None):
    """Load HDI data from devcountries CSV.

    Returns dict mapping country name (lowercase) -> HDI score.
    """
    import csv
    path = path or HDI_PATH
    if not os.path.exists(path):
        return {}
    result = {}
    with open(path, encoding="utf-8") as f:
        reader = csv.DictReader(f)
        for row in reader:
            name = row.get("country", "").strip()
            hdi_str = row.get("HumanDevelopmentIndex_2023", "").strip()
            if name and hdi_str:
                try:
                    result[name.lower()] = float(hdi_str)
                except ValueError:
                    pass
    return result


def load_development_classification(path=None):
    """Load World Bank income group classification.

    Returns dict mapping country name (lowercase) -> income group.
    """
    path = path or CLASSIFICATION_PATH
    if not os.path.exists(path):
        return {}
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb["List of economies"]
    result = {}
    for r in range(2, ws.max_row + 1):
        name = ws.cell(r, 1).value
        code = ws.cell(r, 2).value
        income = ws.cell(r, 4).value
        if name and income:
            result[name.strip().lower()] = income.strip()
    wb.close()

    # Add common name variations for better matching
    aliases = {
        "united states of america": "High income",
        "usa": "High income",
        "u.s.a.": "High income",
        "u.s.": "High income",
        "uk": "High income",
        "u.k.": "High income",
        "uae": "High income",
        "u.a.e.": "High income",
        "south korea": "High income",
        "korea, republic of": "High income",
        "russia": "Upper middle income",
        "russian federation": "Upper middle income",
        "turkey": "Upper middle income",
        "türkiye": "Upper middle income",
        "democratic republic of the congo": "Low income",
        "congo, democratic republic of the": "Low income",
        "ivory coast": "Lower middle income",
        "côte d'ivoire": "Lower middle income",
        "egypt": "Lower middle income",
        "egypt, arab rep.": "Lower middle income",
        "iran": "Upper middle income",
        "iran, islamic rep.": "Upper middle income",
        "venezuela": "Upper middle income",
        "venezuela, rb": "Upper middle income",
        "viet nam": "Lower middle income",
        "vietnam": "Lower middle income",
        "bolivia": "Lower middle income",
        "bolivia (plurinational state of)": "Lower middle income",
        "tanzania": "Lower middle income",
        "united republic of tanzania": "Lower middle income",
    }
    for alias, income in aliases.items():
        if alias not in result:
            result[alias] = income

    return result


def get_development_status(name, classification, hdi_data=None):
    """Return 'Developed', 'Developing', or 'LDC' for a country name.

    Classification priority:
    1. UN LDC list (highest priority)
    2. HDI score (primary): >= 0.800 = Developed, < 0.800 = Developing
    3. World Bank income group (fallback)

    LDC classification:
    - UN LDC list (not graduated) -> LDC
    - HDI < 0.550 -> LDC (very low development)
    """
    name_lower = name.strip().lower()
    # Resolve ITC name -> HDI / World Bank name
    lookup = NAME_ALIASES.get(name_lower, name_lower)

    # Check LDC first (overrides other classifications)
    if name_lower in LDC_COUNTRIES and name_lower not in LDC_GRADUATED:
        return "LDC"

    # Try HDI data first (most accurate)
    if hdi_data and lookup in hdi_data:
        hdi = hdi_data[lookup]
        if hdi >= HDI_DEVELOPED_THRESHOLD:
            return "Developed"
        elif hdi < 0.550:
            return "LDC"
        else:
            return "Developing"

    # Fallback to World Bank income group
    ig = classification.get(lookup, "")
    if ig == "High income":
        return "Developed"
    return "Developing"


def get_region(name):
    """Return ITC region for a country name.

    Returns one of: Africa, Asia, Americas, Pacific, Europe, or 'Other'.
    """
    name_lower = name.strip().lower()
    lookup = NAME_ALIASES.get(name_lower, name_lower)
    for region, countries in ITC_REGIONS.items():
        if lookup in countries or name_lower in countries:
            return region
    return "Other"


def _dev_fill(status):
    if status == "Developed":
        return DEV_FILL
    elif status == "LDC":
        return PatternFill(fill_type="solid", fgColor="FCE4D6")
    return DEVEL_FILL


# ---------------------------------------------------------------------------
# HTML table parser (for ITC .xls files that are actually HTML)
# ---------------------------------------------------------------------------
class _HTMLTableParser(HTMLParser):
    """Extract all <table> elements from an HTML file."""

    def __init__(self):
        super().__init__()
        self.tables = []
        self._cur = None
        self._row = None
        self._data = []
        self._in_cell = False

    def handle_starttag(self, tag, attrs):
        if tag == "table":
            self._cur = []
        elif tag == "tr" and self._cur is not None:
            self._row = []
        elif tag in ("td", "th") and self._row is not None:
            self._data = []
            self._in_cell = True

    def handle_endtag(self, tag):
        if tag == "table" and self._cur is not None:
            if self._cur:
                self.tables.append(self._cur)
            self._cur = None
        elif tag == "tr" and self._row is not None:
            if self._row:
                self._cur.append(self._row)
            self._row = None
        elif tag in ("td", "th") and self._in_cell:
            self._row.append(" ".join(self._data).replace("\xa0", " ").strip())
            self._data = []
            self._in_cell = False

    def handle_data(self, data):
        if self._in_cell:
            self._data.append(data)


def _read_html(path):
    """Read an HTML file, returning the parsed table list."""
    for enc in ("utf-8-sig", "utf-8", "latin-1"):
        try:
            text = open(path, encoding=enc).read()
            break
        except (UnicodeDecodeError, OSError):
            continue
    else:
        text = open(path, encoding="latin-1", errors="replace").read()
    parser = _HTMLTableParser()
    parser.feed(text)
    return parser.tables


# ---------------------------------------------------------------------------
# Formula caching (from make_tables.py)
# ---------------------------------------------------------------------------
def _sheet_order(path):
    with zipfile.ZipFile(path) as zin:
        xml = zin.read("xl/workbook.xml").decode("utf-8", "ignore")
    return re.findall(r'<sheet[^>]*name="([^"]+)"', xml)


def _inject_cached_values(path, cache):
    """Add cached <v> results next to <f> formulas so data_only=True works."""
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
        return
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zout:
        for n in names:
            zout.writestr(n, blobs[n])


def _finalize(ws, path, cache):
    """Save workbook and inject cached formula values.

    ``cache`` is a list of (cell_ref, value) pairs for this worksheet.
    """
    try:
        ws.parent.calculation.fullCalcOnLoad = True
    except Exception:
        pass
    ws.parent.save(path)
    if cache:
        try:
            _inject_cached_values(path, {ws.title: {r: v for r, v in cache
                                                     if v is not None}})
        except Exception:
            pass


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


def _parse_years_from_header(header_row):
    """Extract year integers from a header row like
    ['Code', 'Service label', 'Exported value in 2020', ...]."""
    years = []
    for cell in header_row:
        m = re.search(r"(\d{4})", str(cell))
        if m:
            y = int(m.group(1))
            if 2000 <= y <= 2100:
                years.append(y)
    return years


def _header_year_indices(header_row):
    """Return list of (col_index, year) for year columns in the header."""
    out = []
    for i, cell in enumerate(header_row):
        m = re.search(r"(\d{4})", str(cell))
        if m:
            y = int(m.group(1))
            if 2000 <= y <= 2100:
                out.append((i, y))
    return out


# ---------------------------------------------------------------------------
# File detection
# ---------------------------------------------------------------------------
def find_service_files(excel_dir):
    """Locate the six required raw service source files by filename keywords.

    The ``kenya_commercialized`` file is optional; when present its data is
    loaded but the balance is always computed from export − import totals.
    """
    found = {}
    for fname in sorted(os.listdir(excel_dir)):
        low = fname.lower()
        if not low.endswith((".xls", ".xlsx", ".xlsm")):
            continue
        if "exported_services_for" in low or "list_of_exported_services" in low:
            found["exported_services"] = fname
        elif "exporters_for" in low or "list_of_exporters_for" in low:
            found["exporters"] = fname
        elif "imported_services_for" in low or "list_of_imported_services" in low:
            found["imported_services"] = fname
        elif "importers_for" in low or "list_of_importers_for" in low:
            found["importers"] = fname
        elif "services_exported_by" in low:
            found["kenya_exports"] = fname
        elif "services_imported_by" in low:
            found["kenya_imports"] = fname
        elif "services_commercialized_by" in low:
            found["kenya_commercialized"] = fname

    required = ("exported_services", "exporters",
                 "imported_services", "importers",
                 "kenya_exports", "kenya_imports")
    missing = [k for k in required if k not in found]
    if missing:
        sys.exit(
            "[ERROR] Could not find all six required raw service source files in '%s'.\n"
            "Missing: %s\n"
            "Expected names containing keywords like:\n"
            "  List_of_exported_services_for_the_selected_service\n"
            "  List_of_exporters_for_the_selected_service\n"
            "  List_of_imported_services_for_the_selected_service\n"
            "  List_of_importers_for_the_selected_service\n"
            "  List_of_services_exported_by_Kenya\n"
            "  List_of_services_imported_by_Kenya\n"
            "Optional (balance is always computed from export − import):\n"
            "  List_of_services_commercialized_by_Kenya"
            % (excel_dir, ", ".join(missing)))
    return {k: os.path.join(excel_dir, v) for k, v in found.items()}


# ---------------------------------------------------------------------------
# Source parsing
# ---------------------------------------------------------------------------
def parse_service_html_table(path):
    """Parse an ITC services file.

    Handles both HTML-in-.xls format and proper .xlsx/.xlsm files.

    Returns (data_table, years) where data_table is the data
    as a list of row-lists, and years is the list of year ints
    found in the header.
    """
    ext = os.path.splitext(path)[1].lower()

    if ext in (".xlsx", ".xlsm"):
        # Proper Excel file — read directly with openpyxl
        wb = openpyxl.load_workbook(path, data_only=True)
        ws = wb.active
        data_table = []
        for row in ws.iter_rows(values_only=True):
            data_table.append([str(c) if c is not None else "" for c in row])
        wb.close()
    else:
        # .xls — assume HTML-wrapped (ITC default format)
        tables = _read_html(path)
        if not tables:
            raise ValueError(f"No HTML tables found in {path}")
        data_table = max(tables, key=len)

    if len(data_table) < 2:
        raise ValueError(f"Data table too small in {path}")
    header = data_table[0]
    years = _parse_years_from_header(header)
    return data_table, years


def parse_exported_services(path):
    """Global service exports by category. Returns (total, items).

    total = {'label': ..., 'vals': [...]}  (the 'All services' row)
    items = [{'code': ..., 'label': ..., 'vals': [...]}, ...]
    """
    table, years = parse_service_html_table(path)
    ycols = _header_year_indices(table[0])
    total = None
    items = []
    for row in table[1:]:
        if len(row) < 3:
            continue
        code = str(row[0]).strip()
        label = str(row[1]).strip()
        vals = [to_float(row[ci]) if ci < len(row) else None for ci, _ in ycols]
        if code.upper() == "S" or label.lower() == "all services":
            total = {"code": code, "label": label, "vals": vals}
        else:
            items.append({"code": code, "label": label, "vals": vals})
    return total, items, years


def parse_exporters(path):
    """Global exporters by country. Returns (total_sum, items).

    total_sum = {'label': 'World', 'vals': [...]}  (computed sum)
    items = [{'label': ..., 'vals': [...]}, ...]
    """
    table, years = parse_service_html_table(path)
    ycols = _header_year_indices(table[0])
    items = []
    for row in table[1:]:
        if len(row) < 2:
            continue
        label = str(row[0]).strip()
        vals = [to_float(row[ci]) if ci < len(row) else None for ci, _ in ycols]
        if not label:
            continue
        items.append({"label": label, "vals": vals})
    # Compute world total
    n_years = len(ycols)
    world_vals = []
    for k in range(n_years):
        s = 0.0
        for it in items:
            v = it["vals"][k] if k < len(it["vals"]) else None
            s += v if v is not None else 0.0
        world_vals.append(s)
    total = {"label": "World", "vals": world_vals}
    return total, items, years


def parse_importers(path):
    """Global importers by country. Returns (total_sum, items)."""
    return parse_exporters(path)  # identical structure


def parse_imported_services(path):
    """Global service imports by category. Returns (total, items, years)."""
    return parse_exported_services(path)  # identical structure


def parse_kenya_services(path):
    """Kenya services exports or imports by category.

    Returns (total, items, years).
    total = {'code': ..., 'label': ..., 'vals': [...]}
    items = [{'code': ..., 'label': ..., 'vals': [...]}, ...]
    """
    table, years = parse_service_html_table(path)
    ycols = _header_year_indices(table[0])
    total = None
    items = []
    for row in table[1:]:
        if len(row) < 3:
            continue
        code = str(row[0]).strip()
        label = str(row[1]).strip()
        vals = [to_float(row[ci]) if ci < len(row) else None for ci, _ in ycols]
        if code.upper() == "S" or label.lower() == "all services":
            total = {"code": code, "label": label, "vals": vals}
        else:
            items.append({"code": code, "label": label, "vals": vals})
    return total, items, years


def parse_kenya_commercialized(path):
    """Kenya services balance by category.

    Returns (total, items, years) where items have 'balance', 'export_val',
    'import_val' keys alongside 'vals' (balance values).
    """
    table, years = parse_service_html_table(path)
    # Header: Code, Service label, Balance in value 2020, ..., Exported Value 2024, Imported Value 2024
    header = table[0]
    ycols = _header_year_indices(header)
    # Find export/import 2024 columns
    export_col = None
    import_col = None
    for i, cell in enumerate(header):
        s = str(cell).lower()
        if "exported value" in s:
            export_col = i
        elif "imported value" in s:
            import_col = i

    total = None
    items = []
    for row in table[1:]:
        if len(row) < 3:
            continue
        code = str(row[0]).strip()
        label = str(row[1]).strip()
        vals = [to_float(row[ci]) if ci < len(row) else None for ci, _ in ycols]
        export_val = to_float(row[export_col]) if export_col and export_col < len(row) else None
        import_val = to_float(row[import_col]) if import_col and import_col < len(row) else None
        entry = {"code": code, "label": label, "vals": vals,
                 "export_val": export_val, "import_val": import_val}
        if code.upper() == "S" or label.lower() == "all services":
            total = entry
        else:
            items.append(entry)
    return total, items, years


# ---------------------------------------------------------------------------
# Ranking / aggregation
# ---------------------------------------------------------------------------
def rank_and_top(items, top_n):
    """Sort by latest-year value (desc), assign ranks, return top N."""
    items.sort(
        key=lambda it: (it["vals"][-1] if it["vals"] and it["vals"][-1] is not None
                        else float("-inf")),
        reverse=True,
    )
    for i, it in enumerate(items, 1):
        it["rank"] = i
    return items[:top_n]


def calc_growth_rates(items, years):
    """Calculate year-over-year and latest annual growth rates for each item."""
    if len(years) < 2:
        for it in items:
            it["growth"] = None
        return
    n = len(years)
    for it in items:
        vals = it.get("vals", [])
        if vals and len(vals) >= 2 and vals[-1] is not None and vals[-2] is not None and vals[-2] != 0:
            it["growth"] = (vals[-1] - vals[-2]) / abs(vals[-2])
        else:
            it["growth"] = None


def rank_by_dev_status(items, years, classification, top_n=10):
    """Split items by development status, rank within each group, return top N per group.

    Returns (dev_items, devel_items, ldc_items) where each is a list sorted by latest-year value.
    """
    hdi_data = load_hdi_data()
    # Filter out non-country aggregation entries
    items = [it for it in items if it.get("label", "").strip().lower()
             not in NON_COUNTRY_ENTRIES]
    for it in items:
        it["dev_status"] = get_development_status(it.get("label", ""), classification, hdi_data)
        it["region"] = get_region(it.get("label", ""))

    calc_growth_rates(items, years)

    dev_items = [it for it in items if it["dev_status"] == "Developed"]
    devel_items = [it for it in items if it["dev_status"] == "Developing"]
    ldc_items = [it for it in items if it["dev_status"] == "LDC"]

    dev_items.sort(
        key=lambda it: (it["vals"][-1] if it["vals"] and it["vals"][-1] is not None
                        else float("-inf")),
        reverse=True,
    )
    for i, it in enumerate(dev_items, 1):
        it["dev_rank"] = i

    devel_items.sort(
        key=lambda it: (it["vals"][-1] if it["vals"] and it["vals"][-1] is not None
                        else float("-inf")),
        reverse=True,
    )
    for i, it in enumerate(devel_items, 1):
        it["dev_rank"] = i

    ldc_items.sort(
        key=lambda it: (it["vals"][-1] if it["vals"] and it["vals"][-1] is not None
                        else float("-inf")),
        reverse=True,
    )
    for i, it in enumerate(ldc_items, 1):
        it["dev_rank"] = i

    return dev_items[:top_n], devel_items[:top_n], ldc_items[:top_n]


def rank_by_region(items, years, top_n=5):
    """Group items by region, rank within each, return top N per region.

    Returns dict mapping region -> list of items.
    """
    calc_growth_rates(items, years)

    by_region = {}
    for it in items:
        region = it.get("region", "Other")
        if region not in by_region:
            by_region[region] = []
        by_region[region].append(it)

    for region, region_items in by_region.items():
        region_items.sort(
            key=lambda it: (it["vals"][-1] if it["vals"] and it["vals"][-1] is not None
                            else float("-inf")),
            reverse=True,
        )
        for i, it in enumerate(region_items, 1):
            it["region_rank"] = i

    return {r: items[:top_n] for r, items in by_region.items()}


def sum_shown(shown, n_years):
    out = []
    for k in range(n_years):
        s = 0.0
        for it in shown:
            v = it["vals"][k] if k < len(it["vals"]) else None
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
        t = total["vals"][k] if k < len(total["vals"]) else None
        out.append(t - sums[k] if t is not None else None)
    return out


def share(v, total_v):
    if v is None or total_v is None or total_v == 0:
        return None
    return v / total_v


# ---------------------------------------------------------------------------
# Worksheet styling helpers (shared with make_tables.py)
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


# ---------------------------------------------------------------------------
# Table writers
# ---------------------------------------------------------------------------
def write_market_table(ws, data, hdr_label, verb, unit_row, kenya_highlight,
                       row1_label=None, row1_title=None):
    """Market-style table: rank | country | year values | share.

    Used for Tables 1 (exporters) and 2 (importers).
    """
    years = data["years"]
    n = len(years)
    latest = years[-1]
    total = data["total"]
    shown = data["shown"]

    share_col = 3 + n

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
    merge(ws, h, 1, h + 1, 1)
    put_text(ws, h, 2, hdr_label, bold=True, size=8, fill=HDR_FILL, align="center")
    merge(ws, h, 3, h, 2 + n)
    put_text(ws, h, 3, verb, bold=True, size=11 if unit_row else 8,
             fill=None if unit_row else HDR_FILL, wrap=True, align="center")
    merge(ws, h, share_col, h + 1, share_col)
    put_text(ws, h, share_col, f"Share in {latest} %", bold=True, size=11, align="center")

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
            put_val(ws, r, 3 + k, it["vals"][k] if k < len(it["vals"]) else None, fill=fill)
        t_latest = total["vals"][-1] if total and len(total["vals"]) >= n else None
        val_latest = it["vals"][-1] if it["vals"] else None
        if val_latest is not None and t_latest is not None:
            put_share(ws, r, share_col, share(val_latest, t_latest), fill=fill)
        else:
            put_share(ws, r, share_col, None, fill=fill)

    # All other countries
    r = r_ao
    put_text(ws, r, 2, "All other countries", size=11, wrap=True)
    ao_vals = data["all_other"]["vals"]
    for k in range(n):
        put_val(ws, r, 3 + k, ao_vals[k] if k < len(ao_vals) else None)
    ao_share_val = data["all_other"]["share"]
    put_share(ws, r, share_col, ao_share_val)

    # World total
    r = r_world
    put_text(ws, r, 2, "World", bold=True, size=8, fill=BAND_FILL)
    for k in range(n):
        val = total["vals"][k] if k < len(total["vals"]) else None
        put_val(ws, r, 3 + k, val, bold=True, fill=BAND_FILL)
    put_share(ws, r, share_col, total.get("share"), bold=True)
    ws.freeze_panes = f"A{row0}"

    try:
        ws.auto_filter.ref = f"A{h}:{_cell_ref(r_world, share_col)}"
    except Exception:
        pass


def write_product_table(ws, data, hdr, unit_row,
                        row1_label=None, row1_title=None):
    """Product-style table: rank | code | label | year values | share.

    Used for Tables 3 (Kenya exports) and 4 (Kenya imports).
    """
    years = data["years"]
    n = len(years)
    latest = years[-1]
    total = data["total"]
    shown = data["shown"]

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
    put_text(ws, h, 3, "Service label", bold=True, size=11, align="center")
    merge(ws, h, 4, h, 3 + n)
    put_text(ws, h, 4, hdr, bold=True, size=11, wrap=True, align="center")
    share_col = 4 + n
    merge(ws, h, share_col, h + 1, share_col)
    put_text(ws, h, share_col, f"Share in {latest} %", bold=True, size=11, align="center")

    for k, y in enumerate(years):
        put_text(ws, h + 1, 4 + k, y, bold=True, size=11, align="center")

    if unit_row:
        merge(ws, h + 2, 4, h + 2, 3 + n)
        put_text(ws, h + 2, 4, unit_row, bold=True, size=11,
                 fill=HDR_FILL, align="center")
        row0 = h + 3
    else:
        row0 = h + 2

    r_first = row0
    r_last = row0 + len(shown) - 1
    r_ao = row0 + len(shown)
    r_total = r_ao + 1

    for i, it in enumerate(shown):
        r = row0 + i
        put_text(ws, r, 1, it["rank"], size=11)
        put_text(ws, r, 2, it["code"], size=11)
        put_text(ws, r, 3, it["label"], size=11, wrap=True)
        for k in range(n):
            put_val(ws, r, 4 + k, it["vals"][k] if k < len(it["vals"]) else None)
        t_latest = total["vals"][-1] if total and len(total["vals"]) >= n else None
        val_latest = it["vals"][-1] if it["vals"] else None
        if val_latest is not None and t_latest is not None:
            put_share(ws, r, share_col, share(val_latest, t_latest))
        else:
            put_share(ws, r, share_col, None)

    # All other
    r = r_ao
    put_text(ws, r, 3, "All other services", size=11, wrap=True)
    ao_vals = data["all_other"]["vals"]
    for k in range(n):
        put_val(ws, r, 4 + k, ao_vals[k] if k < len(ao_vals) else None)
    put_share(ws, r, share_col, data["all_other"]["share"])

    # Total
    r = r_total
    if total and total.get("code"):
        put_text(ws, r, 2, str(total["code"]), bold=True, size=11,
                 fill=BAND_FILL, wrap=True)
    put_text(ws, r, 3, "All services", bold=True, size=11, fill=BAND_FILL, wrap=True)
    for k in range(n):
        val = total["vals"][k] if k < len(total["vals"]) else None
        put_val(ws, r, 4 + k, val, bold=True, fill=BAND_FILL)
    put_share(ws, r, share_col, total.get("share") if total else None, bold=True)
    ws.freeze_panes = f"A{row0}"

    try:
        ws.auto_filter.ref = f"A{h}:{_cell_ref(r_total, share_col)}"
    except Exception:
        pass


def write_balance(ws, years, exports, imports, cache=None):
    """Figure 1: Kenya services balance of trade.

    exports = Kenya's service exports per year (USD Million)
    imports = Kenya's service imports per year (USD Million)
    """
    cache = [] if cache is None else cache
    hdr = HDR_FILL
    put_text(ws, 1, 1, "Figure 1", bold=True, size=11, align="center")
    put_text(ws, 1, 2, "Kenya Services Balance of Trade", bold=True, size=11, align="center")

    put_text(ws, 2, 2, "", bold=True, size=11, fill=hdr, align="center")
    for k, y in enumerate(years):
        put_text(ws, 3, 2 + k, y, bold=True, size=11, fill=hdr, align="center")

    r_exp = 4
    r_imp = 5
    r_bal = 6
    put_text(ws, r_exp, 1, "Exports", bold=True, size=11, fill=hdr, align="center")
    for k, v in enumerate(exports):
        put_val(ws, r_exp, 2 + k, v)

    put_text(ws, r_imp, 1, "Imports", bold=True, size=11, fill=hdr, align="center")
    for k, v in enumerate(imports):
        put_val(ws, r_imp, 2 + k, v)

    put_text(ws, r_bal, 1, "Balance of Trade", bold=True, size=11, fill=hdr, align="center")
    for k, (e, i) in enumerate(zip(exports, imports)):
        c = 2 + k
        if e is not None and i is not None:
            cache.append((_cell_ref(r_bal, c), e - i))
            put(ws, r_bal, c, f"={_cell_ref(r_exp, c)}-{_cell_ref(r_imp, c)}",
                bold=False, numfmt=FMT_VALUE)
        else:
            put_val(ws, r_bal, c, (e - i) if (e is not None and i is not None) else None)

    try:
        ws.auto_filter.ref = f"A3:{_cell_ref(r_bal, 2 + len(years))}"
    except Exception:
        pass


# ---------------------------------------------------------------------------
# Development status table writer
# ---------------------------------------------------------------------------
def write_dev_status_table(ws, dev_items, devel_items, ldc_items, years,
                           verb, unit_label, row1_label=None, row1_title=None):
    """Write a table split by development status.

    Columns: Dev Status | Rank | Country | Region | Value (latest year) | Growth %
    The Development Status label is written once per group and merged across
    all rows in that group.
    """
    latest = years[-1] if years else ""
    n = len(years)

    if row1_label or row1_title:
        if row1_label:
            put_text(ws, 1, 1, row1_label, bold=True, size=11, align="center")
        merge(ws, 1, 2, 1, 6)
        put_text(ws, 1, 2, row1_title, bold=True, size=11, wrap=True, align="center")
        title_rows = 1
    else:
        title_rows = 0

    h = 1 + title_rows
    cols = ["Development Status", f"Rank in {latest}", verb, "Region",
            f"Value in USD Billion ({latest})", "Annual Growth %"]
    for c, hdr in enumerate(cols, 1):
        put_text(ws, h, c, hdr, bold=True, size=8, fill=HDR_FILL, align="center",
                 wrap=True)

    row0 = h + 1
    r = row0
    for status, items in [("Developed Economies", dev_items),
                          ("Developing Economies", devel_items),
                          ("Least Developed Countries (LDCs)", ldc_items)]:
        if not items:
            continue
        group_start = r
        fill = _dev_fill(status.split()[0] if status.startswith("Developed") else
                         ("LDC" if "LDC" in status else "Developing"))
        for it in items:
            put_text(ws, r, 2, it.get("dev_rank", ""), size=11, fill=fill)
            put_text(ws, r, 3, it.get("label", ""), size=11, fill=fill, wrap=True,
                     align="left")
            put_text(ws, r, 4, it.get("region", ""), size=11, fill=fill, wrap=True,
                     align="left")
            val = it["vals"][-1] if it.get("vals") and it["vals"][-1] is not None else None
            put_val(ws, r, 5, val, fill=fill)
            growth = it.get("growth")
            if growth is not None:
                put_share(ws, r, 6, growth, fill=fill)
            else:
                put_share(ws, r, 6, None, fill=fill)
            r += 1
        # Write status label in first row and merge down for the group
        put_text(ws, group_start, 1, status, bold=True, size=11, fill=fill,
                 wrap=True, align="center")
        if r - group_start > 1:
            merge(ws, group_start, 1, r - 1, 1)

    set_widths(ws, {"A": 18, "B": 14, "C": 32, "D": 28, "E": 22, "F": 18})
    ws.freeze_panes = f"A{row0}"


def write_region_table(ws, by_region, years, verb, unit_label,
                       row1_label=None, row1_title=None):
    """Write a table grouped by region.

    Columns: Region | Rank | Country | Value (latest year) | Growth %
    """
    latest = years[-1] if years else ""

    if row1_label or row1_title:
        if row1_label:
            put_text(ws, 1, 1, row1_label, bold=True, size=11, align="center")
        merge(ws, 1, 2, 1, 5)
        put_text(ws, 1, 2, row1_title, bold=True, size=11, wrap=True, align="center")
        title_rows = 1
    else:
        title_rows = 0

    h = 1 + title_rows
    cols = ["Region", f"Rank in {latest}", verb,
            f"Value in USD Billion ({latest})", "Annual Growth %"]
    for c, hdr in enumerate(cols, 1):
        put_text(ws, h, c, hdr, bold=True, size=8, fill=HDR_FILL, align="center",
                 wrap=True)

    row0 = h + 1
    r = row0
    region_fills = {
        "Africa": PatternFill(fill_type="solid", fgColor="D6E4F0"),
        "Asia": PatternFill(fill_type="solid", fgColor="E2EFDA"),
        "Americas": PatternFill(fill_type="solid", fgColor="FFF2CC"),
        "Pacific": PatternFill(fill_type="solid", fgColor="BDD7EE"),
        "Europe": PatternFill(fill_type="solid", fgColor="FCE4D6"),
        "Other": PatternFill(fill_type="solid", fgColor="F2F2F2"),
    }
    for region in ["Africa", "Asia", "Americas", "Pacific", "Europe", "Other"]:
        items = by_region.get(region, [])
        if not items:
            continue
        fill = region_fills.get(region, BAND_FILL)
        for it in items:
            put_text(ws, r, 1, region, size=11, fill=fill, wrap=True, align="left")
            put_text(ws, r, 2, it.get("region_rank", ""), size=11, fill=fill)
            put_text(ws, r, 3, it.get("label", ""), size=11, fill=fill, wrap=True,
                     align="left")
            val = it["vals"][-1] if it.get("vals") and it["vals"][-1] is not None else None
            put_val(ws, r, 4, val, fill=fill)
            growth = it.get("growth")
            if growth is not None:
                put_share(ws, r, 5, growth, fill=fill)
            else:
                put_share(ws, r, 5, None, fill=fill)
            r += 1

    set_widths(ws, {"A": 28, "B": 14, "C": 32, "D": 22, "E": 18})
    ws.freeze_panes = f"A{row0}"


# ---------------------------------------------------------------------------
# RCA (Revealed Comparative Advantage) analysis
# ---------------------------------------------------------------------------
def compute_services_rca(kenya_items, kenya_total, world_items, world_total, years):
    """Compute Balassa-style RCA for each service category.

    RCA_i = (Kenya_cat_i / Kenya_total) / (World_cat_i / World_total)

    Uses the latest year where Kenya has per-category data (not None).
    Returns list of dicts sorted by RCA (desc), each with:
      code, label, kenya_val, world_val, kenya_share, world_share, rca, rca_values
    """
    # Index world items by code
    world_by_code = {}
    for it in world_items:
        world_by_code[it["code"]] = it

    # Find the latest year with actual Kenya category data
    best_year_idx = None
    for k in range(len(years) - 1, -1, -1):
        has_data = any(
            k < len(it["vals"]) and it["vals"][k] is not None
            for it in kenya_items
        )
        if has_data:
            best_year_idx = k
            break
    if best_year_idx is None:
        best_year_idx = len(years) - 1

    results = []
    for kit in kenya_items:
        code = kit["code"]
        wit = world_by_code.get(code)
        if wit is None:
            continue

        entry = {
            "code": code,
            "label": kit["label"],
            "rca_values": [],
        }

        for k in range(len(years)):
            kv = kit["vals"][k] if k < len(kit["vals"]) else None
            kt = kenya_total["vals"][k] if kenya_total and k < len(kenya_total["vals"]) else None
            wv = wit["vals"][k] if k < len(wit["vals"]) else None
            wt = world_total["vals"][k] if world_total and k < len(world_total["vals"]) else None

            if kv is not None and kt is not None and kt != 0 and wv is not None and wt is not None and wt != 0:
                k_share = kv / kt
                w_share = wv / wt
                rca = k_share / w_share if w_share != 0 else None
            else:
                k_share = None
                w_share = None
                rca = None

            if k == best_year_idx:
                entry["kenya_val"] = kv
                entry["world_val"] = wv
                entry["kenya_share"] = k_share
                entry["world_share"] = w_share
                entry["rca"] = rca
            entry["rca_values"].append(rca)

        entry["best_year"] = years[best_year_idx]
        results.append(entry)

    results.sort(key=lambda x: x["rca"] if x["rca"] is not None else -1, reverse=True)
    for i, it in enumerate(results, 1):
        it["rank"] = i
    return results, years[best_year_idx]


def write_rca_table(ws, rca_items, years, row1_label=None, row1_title=None):
    """Write an RCA analysis table.

    Columns: Rank | Code | Service Category | Kenya Share % | World Share % | RCA | Status
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
    cols = ["Rank", "Code", f"Service Category",
            f"Kenya Share ({latest}) %", f"World Share ({latest}) %",
            f"RCA ({latest})", "Classification"]
    for c, hdr in enumerate(cols, 1):
        put_text(ws, h, c, hdr, bold=True, size=8, fill=HDR_FILL, align="center",
                 wrap=True)

    STRONG_FILL = PatternFill(fill_type="solid", fgColor="C6EFCE")
    MODERATE_FILL = PatternFill(fill_type="solid", fgColor="E2EFDA")
    DISADVANTAGE_FILL = PatternFill(fill_type="solid", fgColor="FCE4D6")

    row0 = h + 1
    for i, it in enumerate(rca_items):
        r = row0 + i
        rca_val = it.get("rca")
        if rca_val is not None and rca_val >= 2.5:
            fill = STRONG_FILL
            status = "Strong Advantage"
        elif rca_val is not None and rca_val >= 1.0:
            fill = MODERATE_FILL
            status = "Moderate Advantage"
        else:
            fill = DISADVANTAGE_FILL
            status = "Disadvantage"

        put_text(ws, r, 1, it.get("rank", ""), size=11, fill=fill)
        put_text(ws, r, 2, str(it.get("code", "")), size=11, fill=fill)
        put_text(ws, r, 3, it.get("label", ""), size=11, fill=fill, wrap=True,
                 align="left")
        put_share(ws, r, 4, it.get("kenya_share"), fill=fill)
        put_share(ws, r, 5, it.get("world_share"), fill=fill)
        if rca_val is not None:
            put(ws, r, 6, round(rca_val, 2), fill=fill, numfmt="0.00")
        else:
            put(ws, r, 6, "", fill=fill)
        put_text(ws, r, 7, status, size=11, fill=fill, align="center")

    # Legend row
    r = row0 + len(rca_items) + 1
    put(ws, r, 1, "RCA > 2.5: Strong comparative advantage", size=9)
    r += 1
    put(ws, r, 1, "RCA 1.0-2.5: Moderate comparative advantage", size=9)
    r += 1
    put(ws, r, 1, "RCA < 1.0: Comparative disadvantage", size=9)

    set_widths(ws, {"A": 8, "B": 8, "C": 48, "D": 18, "E": 18, "F": 14, "G": 22})
    ws.freeze_panes = f"A{row0}"


# ---------------------------------------------------------------------------
# Export Concentration Index (HHI) analysis
# ---------------------------------------------------------------------------
def compute_concentration_index(kenya_items, kenya_total, world_items, world_total, years):
    """Compute HHI and concentration ratios for Kenya and World service exports.

    HHI = sum of squared shares across categories.
    Lower HHI = more diversified.
    """
    latest_idx = -1

    def _shares(items, total):
        vals = []
        for it in items:
            v = it["vals"][latest_idx] if it.get("vals") and len(it["vals"]) > abs(latest_idx) else None
            if v is not None:
                vals.append(v)
        t = total["vals"][latest_idx] if total and total.get("vals") and len(total["vals"]) > abs(latest_idx) else None
        if t is None or t == 0:
            return []
        return [v / t for v in vals]

    ke_shares = _shares(kenya_items, kenya_total)
    w_shares = _shares(world_items, world_total)

    def _metrics(shares):
        if not shares:
            return {"hhi": 0, "top3": 0, "top5": 0, "eff_cats": 0}
        sorted_s = sorted(shares, reverse=True)
        hhi = sum(s * s for s in sorted_s)
        top3 = sum(sorted_s[:3])
        top5 = sum(sorted_s[:5])
        eff = 1.0 / hhi if hhi > 0 else 0
        return {"hhi": hhi, "top3": top3, "top5": top5, "eff_cats": eff}

    ke_m = _metrics(ke_shares)
    w_m = _metrics(w_shares)

    return {
        "kenya": ke_m,
        "world": w_m,
        "kenya_n_categories": len(ke_shares),
        "world_n_categories": len(w_shares),
    }


def write_concentration_table(ws, concentration, row1_label=None, row1_title=None):
    """Write export concentration index comparison table."""
    if row1_label or row1_title:
        if row1_label:
            put_text(ws, 1, 1, row1_label, bold=True, size=11, align="center")
        merge(ws, 1, 2, 1, 4)
        put_text(ws, 1, 2, row1_title, bold=True, size=11, wrap=True, align="center")
        title_rows = 1
    else:
        title_rows = 0

    h = 1 + title_rows
    cols = ["Metric", "Kenya", "World", "Assessment"]
    for c, hdr in enumerate(cols, 1):
        put_text(ws, h, c, hdr, bold=True, size=8, fill=HDR_FILL, align="center", wrap=True)

    ke = concentration["kenya"]
    w = concentration["world"]

    def _assess(ke_val, w_val, lower_is_better=True):
        if lower_is_better:
            if ke_val < w_val * 0.8:
                return "More diversified"
            elif ke_val > w_val * 1.2:
                return "More concentrated"
        else:
            if ke_val > w_val * 1.2:
                return "More diversified"
            elif ke_val < w_val * 0.8:
                return "More concentrated"
        return "Similar"

    rows = [
        ("Herfindahl-Hirschman Index (HHI)", ke["hhi"], w["hhi"],
         _assess(ke["hhi"], w["hhi"]), "0.000"),
        ("Top-3 Category Share", ke["top3"], w["top3"],
         _assess(ke["top3"], w["top3"]), "0.0%"),
        ("Top-5 Category Share", ke["top5"], w["top5"],
         _assess(ke["top5"], w["top5"]), "0.0%"),
        ("Effective No. of Categories (1/HHI)", ke["eff_cats"], w["eff_cats"],
         _assess(ke["eff_cats"], w["eff_cats"], lower_is_better=False), "0.0"),
        ("Number of Service Categories", concentration["kenya_n_categories"],
         concentration["world_n_categories"], "", "0"),
    ]

    row0 = h + 1
    for i, (label, ke_val, w_val, assess, fmt) in enumerate(rows):
        r = row0 + i
        fill = BAND_FILL if i % 2 == 0 else None
        put_text(ws, r, 1, label, size=11, fill=fill, align="left")
        if fmt == "0.0%":
            put_share(ws, r, 2, ke_val, fill=fill)
            put_share(ws, r, 3, w_val, fill=fill)
        elif fmt == "0.000":
            put(ws, r, 2, ke_val, fill=fill, numfmt="0.000")
            put(ws, r, 3, w_val, fill=fill, numfmt="0.000")
        else:
            put(ws, r, 2, ke_val, fill=fill, numfmt=fmt)
            put(ws, r, 3, w_val, fill=fill, numfmt=fmt)
        put_text(ws, r, 4, assess, size=11, fill=fill, align="center")

    r = row0 + len(rows) + 1
    put(ws, r, 1, "HHI < 0.15: Low concentration (diversified)", size=9)
    r += 1
    put(ws, r, 1, "HHI 0.15-0.25: Moderate concentration", size=9)
    r += 1
    put(ws, r, 1, "HHI > 0.25: High concentration", size=9)

    set_widths(ws, {"A": 42, "B": 18, "C": 18, "D": 24})
    ws.freeze_panes = f"A{row0}"


def make_concentration_chart(concentration, out_path):
    """Bar chart comparing Kenya vs World concentration metrics."""
    metrics = ["HHI", "Top-3 Share", "Top-5 Share"]
    ke_vals = [concentration["kenya"]["hhi"], concentration["kenya"]["top3"],
               concentration["kenya"]["top5"]]
    w_vals = [concentration["world"]["hhi"], concentration["world"]["top3"],
              concentration["world"]["top5"]]

    fig, ax = plt.subplots(figsize=(8, 5))
    x = np.arange(len(metrics))
    width = 0.35
    bars1 = ax.bar(x - width/2, ke_vals, width, label="Kenya", color="#2E75B6")
    bars2 = ax.bar(x + width/2, w_vals, width, label="World", color="#ED7D31")

    ax.set_xticks(x)
    ax.set_xticklabels(metrics, fontsize=10)
    ax.set_ylabel("Value", fontsize=10)
    ax.set_title("Kenya vs World: Export Concentration Metrics", fontsize=12, weight="bold")
    ax.legend(fontsize=10)
    ax.yaxis.grid(True, linestyle="--", alpha=0.35)
    ax.set_axisbelow(True)
    for spine in ("top", "right"):
        ax.spines[spine].set_visible(False)

    for bar in bars1:
        ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.005,
                f"{bar.get_height():.3f}", ha="center", va="bottom", fontsize=8)
    for bar in bars2:
        ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.005,
                f"{bar.get_height():.3f}", ha="center", va="bottom", fontsize=8)

    fig.tight_layout()
    fig.savefig(out_path, dpi=150, bbox_inches="tight")
    plt.close(fig)


# ---------------------------------------------------------------------------
# Diversification Potential analysis
# ---------------------------------------------------------------------------
def compute_diversification_potential(kenya_items, kenya_total, world_items, world_total,
                                       kenya_imp_items, kenya_imp_total, years):
    """Identify categories where Kenya could expand exports.

    Opportunity score = global_growth * (1 - kenya_global_share) * import_penetration
    """
    # Find latest year with actual Kenya category data
    latest_idx = None
    for k in range(len(years) - 1, -1, -1):
        has_data = any(
            k < len(it["vals"]) and it["vals"][k] is not None
            for it in kenya_items
        )
        if has_data:
            latest_idx = k
            break
    if latest_idx is None:
        latest_idx = len(years) - 1

    prev_idx = latest_idx - 1 if latest_idx > 0 else None

    world_by_code = {it["code"]: it for it in world_items}
    kenya_imp_by_code = {it["code"]: it for it in kenya_imp_items}

    ke_total_latest = kenya_total["vals"][latest_idx] if kenya_total and kenya_total.get("vals") and latest_idx < len(kenya_total["vals"]) else None
    ke_total_prev = kenya_total["vals"][prev_idx] if kenya_total and kenya_total.get("vals") and prev_idx is not None and prev_idx < len(kenya_total["vals"]) else None
    w_total_latest = world_total["vals"][latest_idx] if world_total and world_total.get("vals") and latest_idx < len(world_total["vals"]) else None
    w_total_prev = world_total["vals"][prev_idx] if world_total and world_total.get("vals") and prev_idx is not None and prev_idx < len(world_total["vals"]) else None

    results = []
    for kit in kenya_items:
        code = kit["code"]
        if code.upper() == "S" or kit.get("label", "").lower() == "all services":
            continue

        wit = world_by_code.get(code)
        kit_imp = kenya_imp_by_code.get(code)

        ke_val = kit["vals"][latest_idx] if kit.get("vals") and len(kit["vals"]) > abs(latest_idx) else None
        w_val = wit["vals"][latest_idx] if wit and wit.get("vals") and len(wit["vals"]) > abs(latest_idx) else None

        if ke_val is None or w_val is None or w_total_latest is None or w_total_latest == 0:
            continue

        ke_global_share = ke_val / w_total_latest

        ke_val_prev = kit["vals"][prev_idx] if kit.get("vals") and prev_idx is not None and prev_idx < len(kit["vals"]) else None
        w_val_prev = wit["vals"][prev_idx] if wit and wit.get("vals") and prev_idx is not None and prev_idx < len(wit["vals"]) else None

        if ke_val_prev is not None and ke_val_prev > 0:
            ke_growth = (ke_val - ke_val_prev) / ke_val_prev
        else:
            ke_growth = None

        if w_val_prev is not None and w_val_prev > 0:
            w_growth = (w_val - w_val_prev) / w_val_prev
        else:
            w_growth = None

        imp_penetration = 0
        if kit_imp and kenya_imp_total:
            imp_val = kit_imp["vals"][latest_idx] if kit_imp.get("vals") and len(kit_imp["vals"]) > abs(latest_idx) else None
            imp_total = kenya_imp_total["vals"][latest_idx] if kenya_imp_total.get("vals") and len(kenya_imp_total["vals"]) > abs(latest_idx) else None
            if imp_val is not None and imp_total is not None and imp_total > 0:
                imp_penetration = imp_val / imp_total

        growth_factor = max(0, w_growth) if w_growth is not None else 0
        gap_factor = 1.0 - ke_global_share
        opportunity_score = growth_factor * gap_factor * (1 + imp_penetration)

        results.append({
            "code": code,
            "label": kit.get("label", ""),
            "ke_val": ke_val,
            "ke_global_share": ke_global_share,
            "ke_growth": ke_growth,
            "w_growth": w_growth,
            "imp_penetration": imp_penetration,
            "opportunity_score": opportunity_score,
        })

    results.sort(key=lambda x: x["opportunity_score"], reverse=True)
    for i, it in enumerate(results, 1):
        it["rank"] = i
    return results


def write_diversification_table(ws, items, row1_label=None, row1_title=None):
    """Write diversification potential table."""
    if row1_label or row1_title:
        if row1_label:
            put_text(ws, 1, 1, row1_label, bold=True, size=11, align="center")
        merge(ws, 1, 2, 1, 7)
        put_text(ws, 1, 2, row1_title, bold=True, size=11, wrap=True, align="center")
        title_rows = 1
    else:
        title_rows = 0

    h = 1 + title_rows
    cols = ["Rank", "Code", "Service Category", "Kenya Export\n(USD Mn)",
            "Kenya Global\nShare %", "Global\nGrowth %", "Import\nShare %",
            "Opportunity\nScore"]
    for c, hdr in enumerate(cols, 1):
        put_text(ws, h, c, hdr, bold=True, size=8, fill=HDR_FILL, align="center", wrap=True)

    HIGH_FILL = PatternFill(fill_type="solid", fgColor="C6EFCE")
    MED_FILL = PatternFill(fill_type="solid", fgColor="E2EFDA")
    LOW_FILL = PatternFill(fill_type="solid", fgColor="F2F2F2")

    row0 = h + 1
    for i, it in enumerate(items[:15]):
        r = row0 + i
        score = it.get("opportunity_score", 0)
        if score >= 0.05:
            fill = HIGH_FILL
        elif score >= 0.02:
            fill = MED_FILL
        else:
            fill = LOW_FILL

        put_text(ws, r, 1, it.get("rank", ""), size=11, fill=fill)
        put_text(ws, r, 2, str(it.get("code", "")), size=11, fill=fill)
        put_text(ws, r, 3, it.get("label", ""), size=11, fill=fill, wrap=True, align="left")
        put_val(ws, r, 4, it.get("ke_val"), fill=fill)
        put_share(ws, r, 5, it.get("ke_global_share"), fill=fill)
        put_share(ws, r, 6, it.get("w_growth"), fill=fill)
        put_share(ws, r, 7, it.get("imp_penetration"), fill=fill)
        put(ws, r, 8, round(score, 4), fill=fill, numfmt="0.0000")

    r = row0 + min(len(items), 15) + 1
    put(ws, r, 1, "Higher opportunity score = greater diversification potential", size=9)

    set_widths(ws, {"A": 7, "B": 8, "C": 42, "D": 16, "E": 16, "F": 14, "G": 14, "H": 14})
    ws.freeze_panes = f"A{row0}"


def make_diversification_chart(items, out_path):
    """Horizontal bar chart of top diversification opportunities."""
    top = [it for it in items[:10] if it.get("opportunity_score", 0) > 0]
    if not top:
        return
    top.sort(key=lambda x: x["opportunity_score"])

    labels = [it["label"][:35] for it in top]
    scores = [it["opportunity_score"] for it in top]

    fig, ax = plt.subplots(figsize=(10, max(4, len(top) * 0.5)))
    colors = ["#C6EFCE" if s >= 0.05 else "#E2EFDA" if s >= 0.02 else "#F2F2F2" for s in scores]
    bars = ax.barh(labels, scores, color=colors, edgecolor="#808080", linewidth=0.5)
    ax.set_xlabel("Opportunity Score", fontsize=10)
    ax.set_title("Kenya's Top Diversification Opportunities in Services", fontsize=12, weight="bold")
    for bar, val in zip(bars, scores):
        ax.text(bar.get_width() + 0.001, bar.get_y() + bar.get_height() / 2,
                f"{val:.4f}", va="center", fontsize=8)
    fig.tight_layout()
    fig.savefig(out_path, dpi=150, bbox_inches="tight")
    plt.close(fig)


# ---------------------------------------------------------------------------
# Value Composition Trajectory analysis
# ---------------------------------------------------------------------------
HIGH_VALUE_CODES = {"7", "8", "9", "10"}   # Financial, IP, ICT, Other business
TRADITIONAL_CODES = {"3", "4", "5"}         # Transport, Travel, Construction


def compute_value_trajectory(kenya_items, kenya_total, years):
    """Group categories into high-value, traditional, other and compute shares over time."""
    results = {"years": years, "high_value": [], "traditional": [], "other": []}
    n = len(years)

    hv_totals = [0.0] * n
    trad_totals = [0.0] * n
    other_totals = [0.0] * n

    for kit in kenya_items:
        code = kit.get("code", "")
        for k in range(n):
            v = kit["vals"][k] if kit.get("vals") and k < len(kit["vals"]) and kit["vals"][k] is not None else 0
            if code in HIGH_VALUE_CODES:
                hv_totals[k] += v
            elif code in TRADITIONAL_CODES:
                trad_totals[k] += v
            else:
                other_totals[k] += v

    total_vals = []
    for k in range(n):
        t = kenya_total["vals"][k] if kenya_total and kenya_total.get("vals") and k < len(kenya_total["vals"]) and kenya_total["vals"][k] is not None else 0
        total_vals.append(t)

    for k in range(n):
        t = total_vals[k] if total_vals[k] > 0 else 1
        results["high_value"].append(hv_totals[k] / t)
        results["traditional"].append(trad_totals[k] / t)
        results["other"].append(other_totals[k] / t)
        results.setdefault("abs_high_value", []).append(hv_totals[k])
        results.setdefault("abs_traditional", []).append(trad_totals[k])
        results.setdefault("abs_other", []).append(other_totals[k])
        results.setdefault("abs_total", []).append(total_vals[k])

    return results


def write_value_trajectory_table(ws, trajectory, row1_label=None, row1_title=None):
    """Write value composition trajectory table."""
    years = trajectory["years"]
    if row1_label or row1_title:
        if row1_label:
            put_text(ws, 1, 1, row1_label, bold=True, size=11, align="center")
        merge(ws, 1, 2, 1, 3 + len(years))
        put_text(ws, 1, 2, row1_title, bold=True, size=11, wrap=True, align="center")
        title_rows = 1
    else:
        title_rows = 0

    h = 1 + title_rows
    cols = ["Category Group", "Service Codes"] + [str(y) for y in years]
    for c, hdr in enumerate(cols, 1):
        put_text(ws, h, c, hdr, bold=True, size=8, fill=HDR_FILL, align="center", wrap=True)

    HV_FILL = PatternFill(fill_type="solid", fgColor="D6E4F0")
    TRAD_FILL = PatternFill(fill_type="solid", fgColor="E2EFDA")
    OTHER_FILL = PatternFill(fill_type="solid", fgColor="F2F2F2")

    row0 = h + 1
    groups = [
        ("High-Value Services", "7, 8, 9, 10", trajectory["high_value"], HV_FILL,
         "Financial, IP, ICT, Other business"),
        ("Traditional Services", "3, 4, 5", trajectory["traditional"], TRAD_FILL,
         "Transport, Travel, Construction"),
        ("Other Services", "1, 2, 6, 11, 12, SN", trajectory["other"], OTHER_FILL,
         "Manufacturing, Maintenance, Insurance, Personal, Government"),
    ]

    for i, (label, codes, shares, fill, desc) in enumerate(groups):
        r = row0 + i
        put_text(ws, r, 1, label, size=11, fill=fill, align="left", bold=True)
        put_text(ws, r, 2, codes, size=9, fill=fill, align="center")
        for k, sh in enumerate(shares):
            put_share(ws, r, 3 + k, sh, fill=fill)

    r = row0 + len(groups)
    put_text(ws, r, 1, "Total Exports (USD Mn)", size=11, bold=True, align="left")
    put_text(ws, r, 2, "", size=9)
    for k, t in enumerate(trajectory.get("abs_total", [])):
        put_val(ws, r, 3 + k, t)

    r += 1
    put_text(ws, r, 1, "High-Value Exports (USD Mn)", size=11, align="left")
    put_text(ws, r, 2, "", size=9)
    for k, v in enumerate(trajectory.get("abs_high_value", [])):
        put_val(ws, r, 3 + k, v, fill=HV_FILL)

    r += 1
    put_text(ws, r, 1, "Traditional Exports (USD Mn)", size=11, align="left")
    put_text(ws, r, 2, "", size=9)
    for k, v in enumerate(trajectory.get("abs_traditional", [])):
        put_val(ws, r, 3 + k, v, fill=TRAD_FILL)

    widths = {"A": 26, "B": 18}
    for k in range(len(years)):
        widths[get_column_letter(3 + k)] = 14
    set_widths(ws, widths)
    ws.freeze_panes = f"A{row0}"


def make_value_trajectory_chart(trajectory, out_path):
    """Stacked area chart of value composition over time."""
    years = trajectory["years"]
    hv = [s * 100 for s in trajectory["high_value"]]
    trad = [s * 100 for s in trajectory["traditional"]]
    other = [s * 100 for s in trajectory["other"]]

    fig, ax = plt.subplots(figsize=(10, 5))
    ax.stackplot(years, hv, trad, other,
                 labels=["High-Value Services", "Traditional Services", "Other Services"],
                 colors=["#2E75B6", "#70AD47", "#A5A5A5"], alpha=0.85)
    ax.set_xlabel("Year", fontsize=10)
    ax.set_ylabel("Share of Total Exports (%)", fontsize=10)
    ax.set_title("Kenya's Service Export Composition Over Time", fontsize=12, weight="bold")
    ax.legend(loc="upper right", fontsize=9)
    ax.set_ylim(0, 100)
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda v, _: f"{v:.0f}%"))
    ax.yaxis.grid(True, linestyle="--", alpha=0.35)
    ax.set_axisbelow(True)
    for spine in ("top", "right"):
        ax.spines[spine].set_visible(False)
    fig.tight_layout()
    fig.savefig(out_path, dpi=150, bbox_inches="tight")
    plt.close(fig)


# ---------------------------------------------------------------------------
# Kenya vs Peer comparison
# ---------------------------------------------------------------------------
PEER_COUNTRIES = {
    "African Peers": ["south africa", "egypt", "mauritius", "rwanda"],
    "Aspirational Peers": ["singapore", "malaysia"],
}


def write_peer_comparison_table(ws, items_exp, years_exp, row1_label=None, row1_title=None):
    """Write Kenya vs peer countries total service exports comparison."""
    if row1_label or row1_title:
        if row1_label:
            put_text(ws, 1, 1, row1_label, bold=True, size=11, align="center")
        merge(ws, 1, 2, 1, 6)
        put_text(ws, 1, 2, row1_title, bold=True, size=11, wrap=True, align="center")
        title_rows = 1
    else:
        title_rows = 0

    h = 1 + title_rows
    cols = ["Group", "Country", f"Exports\n({years_exp[-1]}) USD Bn",
            "Global\nShare %", "Growth\n(YoY) %", "Rank\n(Global)"]
    for c, hdr in enumerate(cols, 1):
        put_text(ws, h, c, hdr, bold=True, size=8, fill=HDR_FILL, align="center", wrap=True)

    items_by_name = {it["label"].strip().lower(): it for it in items_exp}

    row0 = h + 1
    r = row0

    KENYA_FILL = PatternFill(fill_type="solid", fgColor="FFF2CC")
    GROUP_FILL = PatternFill(fill_type="solid", fgColor="D6E4F0")
    PEER_FILL = PatternFill(fill_type="solid", fgColor="E2EFDA")

    for group_name, peers in PEER_COUNTRIES.items():
        group_start = r
        for peer in peers:
            it = items_by_name.get(peer)
            if it is None:
                continue
            fill = PEER_FILL
            put_text(ws, r, 1, group_name if r == group_start else "", size=11, fill=fill,
                     align="left", bold=(r == group_start))
            put_text(ws, r, 2, it.get("label", ""), size=11, fill=fill, align="left")
            val = it["vals"][-1] if it.get("vals") and it["vals"][-1] is not None else None
            put_val(ws, r, 3, val, fill=fill)
            put_share(ws, r, 4, it.get("share"), fill=fill)
            growth = it.get("growth")
            if growth is not None:
                put_share(ws, r, 5, growth, fill=fill)
            else:
                put_share(ws, r, 5, None, fill=fill)
            put_text(ws, r, 6, it.get("rank", ""), size=11, fill=fill)
            r += 1

    ke_it = items_by_name.get("kenya")
    if ke_it:
        put_text(ws, r, 1, "Focus Country", size=11, fill=KENYA_FILL, align="left", bold=True)
        put_text(ws, r, 2, ke_it.get("label", ""), size=11, fill=KENYA_FILL, align="left", bold=True)
        val = ke_it["vals"][-1] if ke_it.get("vals") and ke_it["vals"][-1] is not None else None
        put_val(ws, r, 3, val, fill=KENYA_FILL, bold=True)
        put_share(ws, r, 4, ke_it.get("share"), fill=KENYA_FILL, bold=True)
        growth = ke_it.get("growth")
        if growth is not None:
            put_share(ws, r, 5, growth, fill=KENYA_FILL, bold=True)
        else:
            put_share(ws, r, 5, None, fill=KENYA_FILL, bold=True)
        put_text(ws, r, 6, ke_it.get("rank", ""), size=11, fill=KENYA_FILL, bold=True)
        r += 1

    set_widths(ws, {"A": 20, "B": 28, "C": 20, "D": 14, "E": 14, "F": 12})
    ws.freeze_panes = f"A{row0}"


def make_rca_radar_chart(rca_items, years, out_path):
    """Create a horizontal bar chart of RCA values for service categories."""
    items = [it for it in rca_items if it.get("rca") is not None]
    items.sort(key=lambda x: x["rca"])

    labels = [it["label"][:35] for it in items]
    rca_vals = [it["rca"] for it in items]

    fig, ax = plt.subplots(figsize=(10, max(4, len(items) * 0.5)))
    colors = ["#C6EFCE" if v >= 2.5 else "#E2EFDA" if v >= 1.0 else "#FCE4D6"
              for v in rca_vals]
    bars = ax.barh(labels, rca_vals, color=colors, edgecolor="#808080", linewidth=0.5)
    ax.axvline(x=1.0, color="#808080", linestyle="--", linewidth=0.8, label="RCA = 1.0")
    ax.axvline(x=2.5, color="#2E75B6", linestyle="--", linewidth=0.8, label="RCA = 2.5")
    ax.set_xlabel("Revealed Comparative Advantage (RCA)", fontsize=10)
    ax.set_title("Kenya's Service Export RCA by Category", fontsize=12, weight="bold")
    ax.legend(fontsize=8)
    for bar, val in zip(bars, rca_vals):
        ax.text(bar.get_width() + 0.05, bar.get_y() + bar.get_height() / 2,
                f"{val:.2f}", va="center", fontsize=8)
    fig.tight_layout()
    fig.savefig(out_path, dpi=150, bbox_inches="tight")
    plt.close(fig)


# ---------------------------------------------------------------------------
# Chart generators
# ---------------------------------------------------------------------------
def make_pie_chart(exports_by_category, years, out_path):
    """Create a pie chart of global service exports by category.

    exports_by_category: list of {'label': str, 'vals': [float]}
    years: list of year ints
    """
    latest_idx = -1
    data = []
    for item in exports_by_category:
        label = item.get("label", "")
        if label.lower() == "all services":
            continue
        val = item["vals"][latest_idx] if item.get("vals") and len(item["vals"]) > abs(latest_idx) else None
        if val is not None and val > 0:
            data.append((label, val))

    data.sort(key=lambda x: x[1], reverse=True)

    # Group small slices (< 3% of total) into "Other"
    total = sum(v for _, v in data)
    main_slices = []
    other_total = 0
    for label, val in data:
        if val / total >= 0.03:
            main_slices.append((label, val))
        else:
            other_total += val
    if other_total > 0:
        main_slices.append(("Other services", other_total))

    labels = [s[0] for s in main_slices]
    sizes = [s[1] for s in main_slices]

    colors = ["#2E75B6", "#ED7D31", "#A5A5A5", "#FFC000", "#4472C4",
              "#70AD47", "#264478", "#9B57A0", "#636363", "#EB7E30"]

    fig, ax = plt.subplots(figsize=(8, 6))
    wedges, texts, autotexts = ax.pie(
        sizes, labels=None, autopct="%1.1f%%", startangle=90,
        colors=colors[:len(sizes)], pctdistance=0.85,
        textprops={"fontsize": 9})
    for t in autotexts:
        t.set_fontsize(8)

    ax.legend(labels, loc="center left", bbox_to_anchor=(1.0, 0.5), fontsize=8)
    ax.set_title("Global Service Exports by Category", fontsize=12, weight="bold")
    fig.tight_layout()
    fig.savefig(out_path, dpi=150, bbox_inches="tight")
    plt.close(fig)


def make_stacked_bar_chart(category_items, years, out_path):
    """Create a stacked bar chart of service exports by category across years.

    category_items: list of {'label': str, 'vals': [float]}
    years: list of year ints
    """
    # Filter out "All services" total
    items = [it for it in category_items if it.get("label", "").lower() != "all services"]
    items.sort(key=lambda it: it["vals"][-1] if it.get("vals") and it["vals"][-1] is not None else 0,
               reverse=True)

    # Take top 8 categories, group rest as "Other"
    top = items[:8]
    other_items = items[8:]

    fig, ax = plt.subplots(figsize=(10, 6))

    x = np.arange(len(years))
    width = 0.6
    colors = ["#2E75B6", "#ED7D31", "#A5A5A5", "#FFC000", "#4472C4",
              "#70AD47", "#264478", "#9B57A0", "#636363"]
    bottom = np.zeros(len(years))

    for i, item in enumerate(top):
        vals = []
        for k in range(len(years)):
            v = item["vals"][k] if k < len(item.get("vals", [])) and item["vals"][k] is not None else 0
            vals.append(v)
        vals = np.array(vals) / 1e3  # Convert to billion
        ax.bar(x, vals, width, bottom=bottom, label=item["label"],
               color=colors[i % len(colors)])
        bottom += vals

    if other_items:
        other_vals = np.zeros(len(years))
        for item in other_items:
            for k in range(len(years)):
                v = item["vals"][k] if k < len(item.get("vals", [])) and item["vals"][k] is not None else 0
                other_vals[k] += v / 1e3
        ax.bar(x, other_vals, width, bottom=bottom, label="Other services",
               color=colors[-1])

    ax.set_xticks(x)
    ax.set_xticklabels(years, fontsize=10)
    ax.set_ylabel("Value (USD Billion)", fontsize=10)
    ax.set_title("Global Service Exports by Category", fontsize=12, weight="bold")
    ax.legend(loc="upper left", bbox_to_anchor=(1.0, 1.0), fontsize=8)
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda v, _: f"{v:,.0f}"))
    fig.tight_layout()
    fig.savefig(out_path, dpi=150, bbox_inches="tight")
    plt.close(fig)


def make_regional_bar_chart(by_region, years, out_path, title="Service Exports by Region"):
    """Create a grouped bar chart showing top countries per region.

    by_region: dict mapping region -> list of items
    years: list of year ints
    """
    # Get top 3 per region, order regions by total value
    region_totals = {}
    for region, items in by_region.items():
        total = sum(
            (it["vals"][-1] if it.get("vals") and it["vals"][-1] is not None else 0)
            for it in items
        )
        region_totals[region] = total

    ordered_regions = sorted(region_totals.keys(), key=lambda r: region_totals[r], reverse=True)

    fig, ax = plt.subplots(figsize=(12, 6))
    x = np.arange(len(ordered_regions))
    width = 0.2
    colors = ["#2E75B6", "#ED7D31", "#70AD47"]

    for rank in range(3):
        vals = []
        for region in ordered_regions:
            items = by_region.get(region, [])
            if rank < len(items):
                v = items[rank]["vals"][-1] if items[rank]["vals"] and items[rank]["vals"][-1] is not None else 0
                vals.append(v / 1e6)  # Convert to billion
            else:
                vals.append(0)
        offset = (rank - 1) * width
        label = f"Top {rank + 1}"
        ax.bar(x + offset, vals, width, label=label, color=colors[rank])

    ax.set_xticks(x)
    ax.set_xticklabels(ordered_regions, fontsize=9, rotation=15, ha="right")
    ax.set_ylabel("Value (USD Billion)", fontsize=10)
    ax.set_title(title, fontsize=12, weight="bold")
    ax.legend(fontsize=9)
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda v, _: f"{v:,.0f}"))
    fig.tight_layout()
    fig.savefig(out_path, dpi=150, bbox_inches="tight")
    plt.close(fig)


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
def generate_service_tables(excel_dir, out_dir, top_n):
    """Build the service Table 1-4 + Figure 1 workbooks.

    top_n : number of top exporter/importer rows kept for Tables 1 and 2.
    """
    os.makedirs(out_dir, exist_ok=True)
    files = find_service_files(excel_dir)

    # ---- Table 1: Global service exporters by country --------------------
    total_exp, items_exp, years_exp = parse_exporters(files["exporters"])
    # Convert to billions
    for it in items_exp:
        it["vals"] = [v / 1e6 if v is not None else None for v in it["vals"]]
    if total_exp:
        total_exp["vals"] = [v / 1e6 if v is not None else None for v in total_exp["vals"]]

    n_years = len(years_exp)
    shown_exp = rank_and_top(items_exp, top_n)
    ao_exp = all_other_vals(total_exp, shown_exp, n_years)
    t_latest_exp = total_exp["vals"][-1] if total_exp else None
    for it in shown_exp:
        it["share"] = share(it["vals"][-1], t_latest_exp)
    if total_exp:
        total_exp["share"] = share(t_latest_exp, t_latest_exp)
    ao_share_exp = share(ao_exp[-1], t_latest_exp) if ao_exp else None
    d1 = {"years": years_exp, "total": total_exp, "shown": shown_exp,
          "all_other": {"vals": ao_exp, "share": ao_share_exp}}

    # ---- Table 2: Global service importers by country --------------------
    total_imp, items_imp, years_imp = parse_importers(files["importers"])
    for it in items_imp:
        it["vals"] = [v / 1e6 if v is not None else None for v in it["vals"]]
    if total_imp:
        total_imp["vals"] = [v / 1e6 if v is not None else None for v in total_imp["vals"]]

    n_years_i = len(years_imp)
    shown_imp = rank_and_top(items_imp, top_n)
    ao_imp = all_other_vals(total_imp, shown_imp, n_years_i)
    t_latest_imp = total_imp["vals"][-1] if total_imp else None
    for it in shown_imp:
        it["share"] = share(it["vals"][-1], t_latest_imp)
    if total_imp:
        total_imp["share"] = share(t_latest_imp, t_latest_imp)
    ao_share_imp = share(ao_imp[-1], t_latest_imp) if ao_imp else None
    d2 = {"years": years_imp, "total": total_imp, "shown": shown_imp,
          "all_other": {"vals": ao_imp, "share": ao_share_imp}}

    # ---- Global service exports by category (for pie chart) ---------------
    total_gexp, items_gexp, years_gexp = parse_exported_services(files["exported_services"])
    for it in items_gexp:
        it["vals"] = [v / 1e6 if v is not None else None for v in it["vals"]]
    if total_gexp:
        total_gexp["vals"] = [v / 1e6 if v is not None else None for v in total_gexp["vals"]]

    # ---- Table 3: Kenya service exports by category ----------------------
    total_kexp, items_kexp, years_kexp = parse_kenya_services(files["kenya_exports"])
    for it in items_kexp:
        it["vals"] = [v / 1e3 if v is not None else None for v in it["vals"]]
    if total_kexp:
        total_kexp["vals"] = [v / 1e3 if v is not None else None for v in total_kexp["vals"]]

    n_years_ke = len(years_kexp)
    shown_kexp = rank_and_top(items_kexp, len(items_kexp))  # show all service categories
    ao_kexp = all_other_vals(total_kexp, shown_kexp, n_years_ke)
    t_latest_kexp = total_kexp["vals"][-1] if total_kexp else None
    for it in shown_kexp:
        it["share"] = share(it["vals"][-1], t_latest_kexp)
    if total_kexp:
        total_kexp["share"] = share(t_latest_kexp, t_latest_kexp)
    ao_share_kexp = share(ao_kexp[-1], t_latest_kexp) if ao_kexp else None
    d3 = {"years": years_kexp, "total": total_kexp, "shown": shown_kexp,
          "all_other": {"vals": ao_kexp, "share": ao_share_kexp}}

    # ---- Table 4: Kenya service imports by category ----------------------
    total_kimp, items_kimp, years_kimp = parse_kenya_services(files["kenya_imports"])
    for it in items_kimp:
        it["vals"] = [v / 1e3 if v is not None else None for v in it["vals"]]
    if total_kimp:
        total_kimp["vals"] = [v / 1e3 if v is not None else None for v in total_kimp["vals"]]

    n_years_ki = len(years_kimp)
    shown_kimp = rank_and_top(items_kimp, len(items_kimp))
    ao_kimp = all_other_vals(total_kimp, shown_kimp, n_years_ki)
    t_latest_kimp = total_kimp["vals"][-1] if total_kimp else None
    for it in shown_kimp:
        it["share"] = share(it["vals"][-1], t_latest_kimp)
    if total_kimp:
        total_kimp["share"] = share(t_latest_kimp, t_latest_kimp)
    ao_share_kimp = share(ao_kimp[-1], t_latest_kimp) if ao_kimp else None
    d4 = {"years": years_kimp, "total": total_kimp, "shown": shown_kimp,
          "all_other": {"vals": ao_kimp, "share": ao_share_kimp}}

    # ---- Figure 1: Kenya services balance of trade ----------------------
    # Use the Kenya export/import years (consistent set) for the balance
    bal_years = years_kexp
    n_bal = len(bal_years)
    # Pad exports/imports to same length
    exports_full = [v / 1e3 if v is not None else None for v in total_kexp["vals"]] if total_kexp else []
    imports_full = [v / 1e3 if v is not None else None for v in total_kimp["vals"]] if total_kimp else []
    exports_full += [None] * (n_bal - len(exports_full))
    imports_full += [None] * (n_bal - len(imports_full))

    # ---- RCA: Kenya services revealed comparative advantage ---------------
    rca_items, rca_year = compute_services_rca(
        items_kexp, total_kexp, items_gexp, total_gexp, years_gexp)

    # ---- Development status classification ------------------------------
    classification = load_development_classification()
    calc_growth_rates(items_exp, years_exp)
    calc_growth_rates(items_imp, years_imp)

    # Add region data to items
    for it in items_exp:
        it["region"] = get_region(it.get("label", ""))
    for it in items_imp:
        it["region"] = get_region(it.get("label", ""))

    # Top exporters by development status
    exp_dev, exp_devel, exp_ldc = rank_by_dev_status(
        [it for it in items_exp if it.get("label", "").lower() != "world"],
        years_exp, classification, top_n=10)
    # Top importers by development status
    imp_dev, imp_devel, imp_ldc = rank_by_dev_status(
        [it for it in items_imp if it.get("label", "").lower() != "world"],
        years_imp, classification, top_n=10)

    # Top exporters by region (top 5 per region)
    exp_by_region = rank_by_region(
        [it for it in items_exp if it.get("label", "").lower() != "world"],
        years_exp, top_n=5)
    # Top importers by region (top 5 per region)
    imp_by_region = rank_by_region(
        [it for it in items_imp if it.get("label", "").lower() != "world"],
        years_imp, top_n=5)

    # ---- Write Excel files -----------------------------------------------
    out = {}

    # Table 1
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Table 1"
    cache = []
    write_market_table(ws, d1, "Exporters",
                       f"List of Top Service Exporters\nValue in USD Billion",
                       unit_row=True, kenya_highlight=False,
                       row1_label="Table 1:",
                       row1_title="Top Global Service Exporters")
    set_widths(ws, {"A": 7, "B": 32, "C": 12, "D": 12, "E": 12, "F": 12,
                    "G": 12, "H": 12})
    out["t1"] = os.path.join(out_dir, "Table 1 Top Global Service Exporters.xlsx")
    _finalize(ws, out["t1"], cache)

    # Table 2
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Table 2"
    cache = []
    write_market_table(ws, d2, "Importers",
                       f"List of Top Service Importers\nValue in USD Billion",
                       unit_row=True, kenya_highlight=False,
                       row1_label="Table 2:",
                       row1_title="Top Global Service Importers")
    set_widths(ws, {"A": 7, "B": 32, "C": 12, "D": 12, "E": 12, "F": 12,
                    "G": 12, "H": 12})
    out["t2"] = os.path.join(out_dir, "Table 2 Top Global Service Importers.xlsx")
    _finalize(ws, out["t2"], cache)

    # Table 3
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Table 3"
    cache = []
    write_product_table(ws, d3,
                        "Kenya's Service Exports by Category\nValue in USD Million",
                        unit_row="Value in USD Million",
                        row1_label="Table 3",
                        row1_title="Kenya's Service Exports by Category")
    set_widths(ws, {"A": 7, "B": 9, "C": 45, "D": 12, "E": 12, "F": 12,
                    "G": 12, "H": 12, "I": 12})
    out["t3"] = os.path.join(out_dir, "Table 3 Kenya Service Exports.xlsx")
    _finalize(ws, out["t3"], cache)

    # Table 4
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Table 4"
    cache = []
    write_product_table(ws, d4,
                        "Kenya's Service Imports by Category\nValue in USD Million",
                        unit_row="Value in USD Million",
                        row1_label="Table 4",
                        row1_title="Kenya's Service Imports by Category")
    set_widths(ws, {"A": 7, "B": 9, "C": 45, "D": 12, "E": 12, "F": 12,
                    "G": 12, "H": 12, "I": 12})
    out["t4"] = os.path.join(out_dir, "Table 4 Kenya Service Imports.xlsx")
    _finalize(ws, out["t4"], cache)

    # Table 5: Top Service Exporters by Development Status
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Table 5"
    cache = []
    write_dev_status_table(ws, exp_dev, exp_devel, exp_ldc, years_exp,
                           "Exporters", "Value in USD Billion",
                           row1_label="Table 5:",
                           row1_title="Top Global Service Exporters by Development Status")
    out["t5"] = os.path.join(out_dir, "Table 5 Service Exporters by Dev Status.xlsx")
    _finalize(ws, out["t5"], cache)

    # Table 6: Top Service Importers by Development Status
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Table 6"
    cache = []
    write_dev_status_table(ws, imp_dev, imp_devel, imp_ldc, years_imp,
                           "Importers", "Value in USD Billion",
                           row1_label="Table 6:",
                           row1_title="Top Global Service Importers by Development Status")
    out["t6"] = os.path.join(out_dir, "Table 6 Service Importers by Dev Status.xlsx")
    _finalize(ws, out["t6"], cache)

    # Table 7: Top Service Exporters by Region
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Table 7"
    cache = []
    write_region_table(ws, exp_by_region, years_exp,
                       "Exporters", "Value in USD Billion",
                       row1_label="Table 7:",
                       row1_title="Top Service Exporters by Region")
    out["t7"] = os.path.join(out_dir, "Table 7 Service Exporters by Region.xlsx")
    _finalize(ws, out["t7"], cache)

    # Table 8: Top Service Importers by Region
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Table 8"
    cache = []
    write_region_table(ws, imp_by_region, years_imp,
                       "Importers", "Value in USD Billion",
                       row1_label="Table 8:",
                       row1_title="Top Service Importers by Region")
    out["t8"] = os.path.join(out_dir, "Table 8 Service Importers by Region.xlsx")
    _finalize(ws, out["t8"], cache)

    # Table 9: Kenya's Revealed Comparative Advantage in Services
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Table 9"
    cache = []
    write_rca_table(ws, rca_items, [rca_year],
                    row1_label="Table 9:",
                    row1_title="Kenya's Revealed Comparative Advantage (RCA) in Services")
    out["t9_rca"] = os.path.join(out_dir, "Table 9 Kenya Services RCA.xlsx")
    _finalize(ws, out["t9_rca"], cache)

    # Figure 6: RCA bar chart
    rca_chart_path = os.path.join(out_dir, "Figure 6 Kenya Services RCA.png")
    try:
        make_rca_radar_chart(rca_items, [rca_year], rca_chart_path)
        out["rca_chart"] = rca_chart_path
    except Exception as e:
        print(f"Warning: Could not create RCA chart: {e}")

    # ---- Table 10: Export Concentration Index (HHI) -----------------------
    concentration = compute_concentration_index(
        items_kexp, total_kexp, items_gexp, total_gexp, years_kexp)

    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Table 10"
    cache = []
    write_concentration_table(ws, concentration,
                              row1_label="Table 10:",
                              row1_title="Export Concentration Index (HHI)")
    out["t10_conc"] = os.path.join(out_dir, "Table 10 Export Concentration Index.xlsx")
    _finalize(ws, out["t10_conc"], cache)

    # Figure 7: Concentration comparison chart
    conc_chart_path = os.path.join(out_dir, "Figure 7 Concentration Comparison.png")
    try:
        make_concentration_chart(concentration, conc_chart_path)
        out["conc_chart"] = conc_chart_path
    except Exception as e:
        print(f"Warning: Could not create concentration chart: {e}")

    # ---- Table 11: Diversification Potential -----------------------------
    div_items = compute_diversification_potential(
        items_kexp, total_kexp, items_gexp, total_gexp,
        items_kimp, total_kimp, years_kexp)

    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Table 11"
    cache = []
    write_diversification_table(ws, div_items,
                                row1_label="Table 11:",
                                row1_title="Diversification Potential in Services")
    out["t11_div"] = os.path.join(out_dir, "Table 11 Diversification Potential.xlsx")
    _finalize(ws, out["t11_div"], cache)

    # Figure 8: Diversification chart
    div_chart_path = os.path.join(out_dir, "Figure 8 Diversification Opportunities.png")
    try:
        make_diversification_chart(div_items, div_chart_path)
        out["div_chart"] = div_chart_path
    except Exception as e:
        print(f"Warning: Could not create diversification chart: {e}")

    # ---- Table 12: Value Composition Trajectory --------------------------
    trajectory = compute_value_trajectory(items_kexp, total_kexp, years_kexp)

    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Table 12"
    cache = []
    write_value_trajectory_table(ws, trajectory,
                                 row1_label="Table 12:",
                                 row1_title="Service Export Composition Trajectory")
    out["t12_traj"] = os.path.join(out_dir, "Table 12 Value Composition Trajectory.xlsx")
    _finalize(ws, out["t12_traj"], cache)

    # Figure 9: Value composition chart
    traj_chart_path = os.path.join(out_dir, "Figure 9 Value Composition Trajectory.png")
    try:
        make_value_trajectory_chart(trajectory, traj_chart_path)
        out["traj_chart"] = traj_chart_path
    except Exception as e:
        print(f"Warning: Could not create trajectory chart: {e}")

    # ---- Table 13: Kenya vs Peer Total Exports ---------------------------
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Table 13"
    cache = []
    write_peer_comparison_table(ws, items_exp, years_exp,
                                row1_label="Table 13:",
                                row1_title="Kenya vs Peer Countries: Service Exports")
    out["t13_peer"] = os.path.join(out_dir, "Table 13 Kenya vs Peers Service Exports.xlsx")
    _finalize(ws, out["t13_peer"], cache)

    # Figure 1: Kenya Services Balance
    bal_path = os.path.join(out_dir, "Figure 1 Services Balance.xlsx")
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Figure 1"
    cache = []
    write_balance(ws, bal_years, exports_full, imports_full, cache=cache)
    set_widths(ws, BALANCE_WIDTHS)
    out["balance"] = bal_path
    _finalize(ws, out["balance"], cache)

    # Figure 2: Pie chart of global service exports by category
    pie_path = os.path.join(out_dir, "Figure 2 Service Exports Structure.png")
    try:
        make_pie_chart(items_gexp, years_gexp, pie_path)
        out["pie"] = pie_path
    except Exception as e:
        print(f"Warning: Could not create pie chart: {e}")

    # Figure 3: Stacked bar chart of service exports by category
    bar_path = os.path.join(out_dir, "Figure 3 Service Exports by Category.png")
    try:
        make_stacked_bar_chart(items_gexp, years_gexp, bar_path)
        out["stacked_bar"] = bar_path
    except Exception as e:
        print(f"Warning: Could not create stacked bar chart: {e}")

    # Figure 4: Regional bar chart of exporters
    reg_exp_path = os.path.join(out_dir, "Figure 4 Service Exports by Region.png")
    try:
        make_regional_bar_chart(exp_by_region, years_exp, reg_exp_path,
                                title="Top Service Exporters by Region")
        out["reg_exp_chart"] = reg_exp_path
    except Exception as e:
        print(f"Warning: Could not create regional exporter chart: {e}")

    # Figure 5: Regional bar chart of importers
    reg_imp_path = os.path.join(out_dir, "Figure 5 Service Imports by Region.png")
    try:
        make_regional_bar_chart(imp_by_region, years_imp, reg_imp_path,
                                title="Top Service Importers by Region")
        out["reg_imp_chart"] = reg_imp_path
    except Exception as e:
        print(f"Warning: Could not create regional importer chart: {e}")

    # ---- Combined workbook: all sheets in one file -----------------------
    wb_all = openpyxl.Workbook()
    wb_all.remove(wb_all.active)
    widths_mkt = {"A": 7, "B": 32, "C": 12, "D": 12, "E": 12, "F": 12,
                  "G": 12, "H": 12}
    widths_prd = {"A": 7, "B": 9, "C": 45, "D": 12, "E": 12, "F": 12,
                  "G": 12, "H": 12, "I": 12}
    cache_all = {}

    for s in ("Table 1", "Table 2", "Table 3", "Table 4", "Table 5", "Table 6",
              "Table 7", "Table 8", "Table 9", "Table 10", "Table 11", "Table 12",
              "Table 13", "Figure 1"):
        wb_all.create_sheet(s)

    ws = wb_all["Table 1"]; c = []
    write_market_table(ws, d1, "Exporters",
                       "List of Top Service Exporters\nValue in USD Billion",
                       unit_row=True, kenya_highlight=False,
                       row1_label="Table 1:",
                       row1_title="Top Global Service Exporters")
    set_widths(ws, widths_mkt); cache_all["Table 1"] = dict(c)

    ws = wb_all["Table 2"]; c = []
    write_market_table(ws, d2, "Importers",
                       "List of Top Service Importers\nValue in USD Billion",
                       unit_row=True, kenya_highlight=False,
                       row1_label="Table 2:",
                       row1_title="Top Global Service Importers")
    set_widths(ws, widths_mkt); cache_all["Table 2"] = dict(c)

    ws = wb_all["Table 3"]; c = []
    write_product_table(ws, d3,
                        "Kenya's Service Exports by Category\nValue in USD Million",
                        unit_row="Value in USD Million",
                        row1_label="Table 3",
                        row1_title="Kenya's Service Exports by Category")
    set_widths(ws, widths_prd); cache_all["Table 3"] = dict(c)

    ws = wb_all["Table 4"]; c = []
    write_product_table(ws, d4,
                        "Kenya's Service Imports by Category\nValue in USD Million",
                        unit_row="Value in USD Million",
                        row1_label="Table 4",
                        row1_title="Kenya's Service Imports by Category")
    set_widths(ws, widths_prd); cache_all["Table 4"] = dict(c)

    ws = wb_all["Table 5"]; c = []
    write_dev_status_table(ws, exp_dev, exp_devel, exp_ldc, years_exp,
                           "Exporters", "Value in USD Billion",
                           row1_label="Table 5:",
                           row1_title="Top Global Service Exporters by Development Status")
    cache_all["Table 5"] = dict(c)

    ws = wb_all["Table 6"]; c = []
    write_dev_status_table(ws, imp_dev, imp_devel, imp_ldc, years_imp,
                           "Importers", "Value in USD Billion",
                           row1_label="Table 6:",
                           row1_title="Top Global Service Importers by Development Status")
    cache_all["Table 6"] = dict(c)

    ws = wb_all["Table 7"]; c = []
    write_region_table(ws, exp_by_region, years_exp,
                       "Exporters", "Value in USD Billion",
                       row1_label="Table 7:",
                       row1_title="Top Service Exporters by Region")
    cache_all["Table 7"] = dict(c)

    ws = wb_all["Table 8"]; c = []
    write_region_table(ws, imp_by_region, years_imp,
                       "Importers", "Value in USD Billion",
                       row1_label="Table 8:",
                       row1_title="Top Service Importers by Region")
    cache_all["Table 8"] = dict(c)

    ws = wb_all["Table 9"]; c = []
    write_rca_table(ws, rca_items, [rca_year],
                    row1_label="Table 9:",
                    row1_title="Kenya's Revealed Comparative Advantage (RCA) in Services")
    cache_all["Table 9"] = dict(c)

    ws = wb_all["Table 10"]; c = []
    write_concentration_table(ws, concentration,
                              row1_label="Table 10:",
                              row1_title="Export Concentration Index (HHI)")
    cache_all["Table 10"] = dict(c)

    ws = wb_all["Table 11"]; c = []
    write_diversification_table(ws, div_items,
                                row1_label="Table 11:",
                                row1_title="Diversification Potential in Services")
    cache_all["Table 11"] = dict(c)

    ws = wb_all["Table 12"]; c = []
    write_value_trajectory_table(ws, trajectory,
                                 row1_label="Table 12:",
                                 row1_title="Service Export Composition Trajectory")
    cache_all["Table 12"] = dict(c)

    ws = wb_all["Table 13"]; c = []
    write_peer_comparison_table(ws, items_exp, years_exp,
                                row1_label="Table 13:",
                                row1_title="Kenya vs Peer Countries: Service Exports")
    cache_all["Table 13"] = dict(c)

    ws = wb_all["Figure 1"]; c = []
    write_balance(ws, bal_years, exports_full, imports_full, cache=c)
    set_widths(ws, BALANCE_WIDTHS); cache_all["Figure 1"] = dict(c)

    out["all"] = os.path.join(out_dir, "All Service Tables.xlsx")
    try:
        wb_all.calculation.fullCalcOnLoad = True
    except Exception:
        pass
    wb_all.save(out["all"])
    try:
        _inject_cached_values(out["all"], cache_all)
    except Exception:
        pass

    # Extract country name from Kenya exports title
    rep = "Kenya"
    partner = "the World"

    return out, (rep, partner)


def main():
    ap = argparse.ArgumentParser(
        description="Build service Table 1-4 + Figure 1 from raw ITC source files.")
    ap.add_argument("--excel-dir", default="services",
                    help="Folder with the seven raw ITC service files (default: services)")
    ap.add_argument("--out-dir", default=os.path.join("output", "service_tables"),
                    help="Where to write the generated tables (default: output/service_tables)")
    ap.add_argument("--top", type=int, default=20,
                    help="Number of top exporter/importer rows to keep (default: 20)")
    args = ap.parse_args()

    if args.top < 1:
        sys.exit("[ERROR] --top must be >= 1")

    try:
        out, (rep, partner) = generate_service_tables(args.excel_dir, args.out_dir, args.top)
    except SystemExit:
        raise
    except Exception as e:
        sys.exit(f"[ERROR] Failed to generate service tables: {e}")

    print("[1/2] Read raw service source files from :", os.path.abspath(args.excel_dir))
    print("[2/2] Wrote service tables for            :", rep, "<->", partner)
    for key in ("t1", "t2", "t3", "t4", "balance", "all"):
        print(f"      - {os.path.basename(out[key])}")
    print()
    print("Next step: run")
    print(f'  python generate_services_report.py --excel-dir "{os.path.abspath(args.out_dir)}"')


if __name__ == "__main__":
    main()
