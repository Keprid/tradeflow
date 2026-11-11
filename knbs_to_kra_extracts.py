#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
knbs_to_kra_extracts.py
=======================

Convert the KNBS IMTS workbook into KRA-layout quarterly extracts.

The quarterly "Kenya Export Performance" report compares the review quarter
with the same quarter of the previous year. KRA customs extracts normally
carry only the review year; the KNBS IMTS workbook ("KNBS Data for the
Economic Survey ...") holds the previous year's customs records.

This tool slices the IMTS sheet down to the required months and writes one
workbook per flow in exactly the layout ``make_quarterly_tables.py`` reads:

    Year | Month | Destination Name / Origin Name | SHORT_DESC | CIF Value

so the generated previous-year files can sit in the same input folder as
the KRA extracts and be merged automatically.

Notes
-----
* Product descriptions are recovered by mapping the IMTS HS code onto the
  SHORT_DESC labels of the KRA extracts (HS codes are normalised as
  integers, which reconciles KRA's 7-digit codes with KNBS's 8-digit
  codes). Unmapped codes fall back to an HS-chapter label, then to the
  pipeline's "All others" label.
* IMTS imports are FOB-like ``Value`` plus separate ``Insurance`` and
  ``Freight`` columns; the KRA layout is CIF, so the three are summed.
* Both "Domestic exports" and "Re-exports" count as exports; only
  "Imports" counts as imports.

Usage
-----
    python3 knbs_to_kra_extracts.py \
        --knbs "KNBS Data for the Economic Survey 2025.xlsx" \
        --kra-dir "3rd qur" --out-dir output/knbs_2025 --year 2025 \
        --months 7,8,9

Dependencies: openpyxl, pandas, xlrd (for the legacy .xls KRA files).
"""

import argparse
import os
import sys
import zipfile
import xml.etree.ElementTree as ET

from pipeline_errors import PipelineError

import openpyxl
import pandas as pd

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE_DIR)

from hs_chapters import CHAPTER_NAMES  # noqa: E402

EXPORT_FLOWS = {"domestic exports", "re-exports", "exports", "export"}
IMPORT_FLOWS = {"imports", "import"}

UNDEFINED_EXPORT_LABEL = "All others (non-defined)"
UNDEFINED_IMPORT_LABEL = "All others"

NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
NS_R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"


def _col_index(ref):
    """Zero-based column index for a cell reference such as 'AB12'."""
    n = 0
    for ch in ref:
        if ch.isalpha():
            n = n * 26 + (ord(ch.upper()) - 64)
        else:
            break
    return n - 1


def iter_xlsx_rows(path, sheet_name):
    """Stream a worksheet as lists of cell values without openpyxl.

    The KNBS workbook holds 325k rows; openpyxl's read-only mode needs
    several minutes for it, while a direct XML stream takes seconds.
    Cached formula results are read from the cell <v> elements.
    """
    with zipfile.ZipFile(path) as z:
        names = set(z.namelist())
        wb = ET.fromstring(z.read("xl/workbook.xml"))
        rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
        targets = {r.get("Id"): r.get("Target") for r in rels}
        sheet_path = None
        for sh in wb.iter(NS + "sheet"):
            if sh.get("name") == sheet_name:
                sheet_path = targets.get(sh.get(NS_R + "id"))
                break
        if not sheet_path:
            raise PipelineError(
                "[ERROR] Sheet '%s' not found in %s."
                % (sheet_name, os.path.basename(path)))
        if sheet_path.startswith("/"):
            sheet_path = sheet_path.lstrip("/")
        elif not sheet_path.startswith("xl/"):
            sheet_path = "xl/" + sheet_path

        shared = []
        if "xl/sharedStrings.xml" in names:
            with z.open("xl/sharedStrings.xml") as fh:
                for _, el in ET.iterparse(fh):
                    if el.tag == NS + "si":
                        shared.append("".join(
                            t.text or "" for t in el.iter(NS + "t")))
                        el.clear()

        with z.open(sheet_path) as fh:
            row = []
            for _, el in ET.iterparse(fh):
                if el.tag == NS + "c":
                    idx = _col_index(el.get("r") or "")
                    ctype = el.get("t")
                    if ctype == "inlineStr":
                        val = "".join(
                            t.text or "" for t in el.iter(NS + "t"))
                    else:
                        v = el.find(NS + "v")
                        val = None if v is None else v.text
                        if val is not None and ctype == "s":
                            try:
                                val = shared[int(val)]
                            except (ValueError, IndexError):
                                pass
                    while len(row) <= idx:
                        row.append(None)
                    row[idx] = val
                    el.clear()
                elif el.tag == NS + "row":
                    yield row
                    row = []
                    el.clear()
            if row:
                yield row

# IMTS column positions (0-based) in the standard KNBS export
COL_YEAR, COL_MONTH, COL_PARTNER_CODE, COL_PARTNER_NAME = 0, 2, 3, 4
COL_HS, COL_VALUE, COL_INSURANCE, COL_FREIGHT, COL_FLOW, COL_SOURCE = \
    5, 11, 12, 13, 14, 15


def _to_float(v):
    try:
        if v is None or (isinstance(v, str) and not v.strip()):
            return 0.0
        return float(str(v).replace(",", ""))
    except (TypeError, ValueError):
        return 0.0


def _hs_key(v):
    """Normalise an HS code to a comparable integer string ('01012100'
    and '1012100' both -> '1012100')."""
    if v is None:
        return None
    s = str(v).strip()
    if not s:
        return None
    try:
        return str(int(float(s)))
    except (TypeError, ValueError):
        return None


def _chapter_label(hs_key):
    """'Chapter 27 - Mineral fuels, ...' fallback label for an HS code."""
    if not hs_key:
        return None
    chapter = hs_key.zfill(8)[:2]
    name = CHAPTER_NAMES.get(chapter)
    if name:
        return "Chapter %s - %s" % (chapter, name)
    return None


def find_kra_files(kra_dir):
    """Locate the KRA exports / imports extracts in ``kra_dir``."""
    found = {"exports": None, "imports": None}
    for fname in sorted(os.listdir(kra_dir)):
        low = fname.lower()
        if not low.endswith((".xls", ".xlsx", ".xlsm")):
            continue
        path = os.path.join(kra_dir, fname)
        if "export" in low and "destin" in low:
            found.setdefault("exports", path)
            found["exports"] = found["exports"] or path
        elif "import" in low and "origin" in low:
            found["imports"] = found["imports"] or path
    missing = [k for k in ("exports", "imports") if not found[k]]
    if missing:
        raise PipelineError(
            "[ERROR] Missing KRA extract(s) in '%s' to source product "
            "descriptions: %s.\nFiles present: %s"
            % (kra_dir, ", ".join(missing),
               ", ".join(sorted(os.listdir(kra_dir)))))
    return found


def build_hs_maps(kra_files):
    """HS code -> SHORT_DESC maps from the KRA extracts (per side + union)."""
    maps = {}
    for side, path in kra_files.items():
        df = pd.read_excel(path, usecols=lambda c: str(c).strip() in
                           ("HS_C", "SHORT_DESC"))
        df.columns = [str(c).strip() for c in df.columns]
        mapping = {}
        if "HS_C" in df.columns and "SHORT_DESC" in df.columns:
            for hs, desc in zip(df["HS_C"], df["SHORT_DESC"]):
                key = _hs_key(hs)
                if key is None:
                    continue
                if isinstance(desc, str) and desc.strip():
                    mapping.setdefault(key, desc.strip())
        maps[side] = mapping
    union = {}
    for side in ("imports", "exports"):
        for k, v in maps.get(side, {}).items():
            union.setdefault(k, v)
    maps["all"] = union
    return maps


def convert(knbs_path, kra_dir, out_dir, year, months):
    """Slice the IMTS sheet into KRA-layout previous-year extracts."""
    kra_files = find_kra_files(kra_dir)
    hs_maps = build_hs_maps(kra_files)

    os.makedirs(out_dir, exist_ok=True)

    rows = iter_xlsx_rows(knbs_path, "IMTS")
    # (side, month, partner, product_label) -> value in KES
    agg = {}
    stats = {"read": 0, "kept": 0, "unmapped": 0, "unmapped_value": 0.0,
             "skipped_source": 0}
    totals = {"exports": 0.0, "imports": 0.0}

    for row in rows:
        if row is None or len(row) <= COL_FLOW:
            continue
        if row[COL_YEAR] is None:
            continue
        stats["read"] += 1
        try:
            if int(float(row[COL_YEAR])) != year:
                continue
            month = int(float(row[COL_MONTH]))
        except (TypeError, ValueError):
            continue
        if month not in months:
            continue

        flow = str(row[COL_FLOW] or "").strip().lower()
        if flow in EXPORT_FLOWS:
            side = "exports"
        elif flow in IMPORT_FLOWS:
            side = "imports"
        else:
            continue

        source = str(row[COL_SOURCE] or "").strip()
        if source and source.lower() != "customs":
            stats["skipped_source"] += 1
            continue

        partner = (row[COL_PARTNER_NAME]
                   or row[COL_PARTNER_CODE] or "").strip()
        if not partner:
            continue
        partner = str(partner).upper()

        value = _to_float(row[COL_VALUE])
        if side == "imports":
            value += _to_float(row[COL_INSURANCE])
            value += _to_float(row[COL_FREIGHT])
        if value == 0.0:
            continue

        hs = _hs_key(row[COL_HS])
        desc = hs_maps["all"].get(hs) or hs_maps[side].get(hs)
        if not desc:
            desc = _chapter_label(hs)
        if not desc:
            stats["unmapped"] += 1
            stats["unmapped_value"] += value
            desc = (UNDEFINED_EXPORT_LABEL if side == "exports"
                    else UNDEFINED_IMPORT_LABEL)

        key = (side, month, partner, desc)
        agg[key] = agg.get(key, 0.0) + value
        stats["kept"] += 1
        totals[side] += value

    # ---- write the two KRA-layout workbooks --------------------------------
    out_paths = {}
    specs = (
        ("exports", "Destination Name",
         "EPC %d JUL SEPT data Exports by HS Destin.xlsx"),
        ("imports", "Origin Name",
         "EPC %d JUL SEPT data Imports BY Origin.xlsx"),
    )
    for side, partner_col, name_fmt in specs:
        path = os.path.join(out_dir, name_fmt % year)
        out = openpyxl.Workbook(write_only=True)
        sheet = out.create_sheet("Data")
        sheet.append(["Year", "Month", partner_col, "SHORT_DESC",
                      "CIF Value"])
        side_rows = [(month, partner, desc, value)
                     for (s, month, partner, desc), value in agg.items()
                     if s == side]
        side_rows.sort(key=lambda r: (r[0], r[1], -r[3]))
        for month, partner, desc, value in side_rows:
            sheet.append([year, month, partner, desc, round(value, 2)])
        out.save(path)
        out_paths[side] = path

    stats["totals"] = totals
    stats["rows_written"] = {
        side: sum(1 for s, *_ in agg if s == side) for side in totals}
    return stats, out_paths


def main():
    ap = argparse.ArgumentParser(
        description="Convert the KNBS IMTS workbook into KRA-layout "
                    "previous-year quarterly extracts.")
    ap.add_argument("--knbs", required=True,
                    help="Path to the KNBS IMTS workbook (.xlsx)")
    ap.add_argument("--kra-dir", required=True,
                    help="Folder holding the KRA extracts (product names)")
    ap.add_argument("--out-dir", required=True,
                    help="Folder to write the previous-year extracts into")
    ap.add_argument("--year", type=int, required=True,
                    help="Previous year held in the KNBS workbook")
    ap.add_argument("--months", default="7,8,9",
                    help="Comma-separated review months (default: 7,8,9)")
    args = ap.parse_args()

    months = {int(m) for m in args.months.split(",") if m.strip()}
    print("[1/2] Reading KNBS IMTS from : %s" % os.path.abspath(args.knbs))
    stats, paths = convert(args.knbs, args.kra_dir, args.out_dir,
                           args.year, months)
    print("[2/2] Extracts written to    : %s"
          % os.path.abspath(args.out_dir))
    print("      rows read              : %d" % stats["read"])
    print("      rows kept              : %d" % stats["kept"])
    print("      exports (Ksh billion)  : %.1f"
          % (stats["totals"]["exports"] / 1e9))
    print("      imports (Ksh billion)  : %.1f"
          % (stats["totals"]["imports"] / 1e9))
    if stats["unmapped"]:
        print("      unmapped-product rows  : %d (Ksh %.1f billion)"
              % (stats["unmapped"], stats["unmapped_value"] / 1e9))
    for side, path in paths.items():
        print("      %-7s                : %s" % (side, path))


if __name__ == "__main__":
    try:
        main()
    except PipelineError as e:
        sys.exit(str(e))
