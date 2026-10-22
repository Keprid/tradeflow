#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
webapp/main.py
==============

FastAPI web app for the Trade Flow Report Generator.

One click, end to end:
  1. upload the six raw ITC downloads (or a zip of them), or the seven
     ready-made tables;
  2. pick a country config;
  3. the server builds Table 1-6 + Figure 1 (make_tables) and the Word report
     (generate_report), then serves the .docx for download.

Run (from the project root):
    python -m uvicorn webapp.main:app --host 127.0.0.1 --port 8000
"""

import io
import json
import logging
import os
import re
import shutil
import sys
import threading
import time
import traceback
import uuid
import zipfile
from contextlib import asynccontextmanager
from pathlib import Path

import openpyxl
from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse, JSONResponse, Response
from fastapi.staticfiles import StaticFiles

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(name)s: %(message)s")
LOG = logging.getLogger("tradeflow.webapp")

BASE_DIR = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(BASE_DIR))

import make_tables                       # noqa: E402
import input_validation                  # noqa: E402
import generate_report as gr             # noqa: E402
import make_config                       # noqa: E402
import make_services_tables              # noqa: E402
import generate_services_report as gsr   # noqa: E402
import make_period_tables as mpt             # noqa: E402
import generate_period_report as gpr         # noqa: E402
import period_data                           # noqa: E402
import generate_product_profile as gpp   # noqa: E402
import generate_crafts_report as gcr     # noqa: E402
import generate_competitor_analysis as gca  # noqa: E402
from pipeline_errors import PipelineError   # noqa: E402
from xlsx_compat import (                   # noqa: E402
    _write_csv_as_xlsx, _write_xls_as_xlsx, _write_xlsb_as_xlsx,
)

WEBAPP_DIR = Path(__file__).resolve().parent
STATIC_DIR = WEBAPP_DIR / "static"
JOBS_DIR = WEBAPP_DIR / "jobs"
CONFIG_DIR = BASE_DIR / "config"
SAMPLE_DIR = BASE_DIR / "Samplefiles"
JOB_TTL_SECONDS = 24 * 3600              # old jobs are cleaned up after 1 day

ALLOWED_EXT = (".xlsx", ".xlsm", ".xls", ".xlsb", ".csv")

# Upload guards: the hosted (Render free-tier) disk is small and one pipeline
# runs at a time, so cap what a single job may consume.  KRA quarterly
# extracts run to ~400 MB, hence the generous per-file ceiling.
MAX_UPLOAD_BYTES = 500 * 1024 * 1024        # 500 MB per file
MAX_TOTAL_UPLOAD_BYTES = 1024 * 1024 * 1024  # 1 GB per job

# Job ids are uuid4().hex; anything else must never reach the filesystem.
_JOB_ID_RE = re.compile(r"[0-9a-f]{32}")
# Config stems come from filenames we control, but reject traversal anyway.
_CFG_ID_RE = re.compile(r"[A-Za-z0-9_\-]+")

# Trade Map classic downloads carry long product-group names in the filename
# ("...List_of_exported_products_for_the_selected_product_(Preparations_of_meat_"
# "...).xls").  Together with the job dir the full path can exceed Windows'
# 260-char MAX_PATH limit, so stored basenames are clipped to this length.
_MAX_UPLOAD_NAME_CHARS = 120


def _clip_name(name, cap=_MAX_UPLOAD_NAME_CHARS):
    """Shorten a stored basename to at most ``cap`` chars, keeping the
    leading Trade Map stem (role detection) and the file extension."""
    name = os.path.basename(str(name))
    if len(name) <= cap:
        return name
    stem, ext = os.path.splitext(name)
    return stem[: cap - len(ext)] + ext

# The six raw ITC Trade Map downloads the Goods pipeline needs, in the order
# make_tables expects (table1..table6).  Trade Map has no public API (it is an
# Angular SPA), so "auto-fetch" means: we describe the exact per-partner query
# to run and the file-naming patterns auto-detection will pick up once the
# analyst downloads each workbook.
ITC_GOODS_QUERIES = [
    {
        "role": "table1",
        "query": "Imports from World, by Market, for {partner} "
                 "(all products, top supplying markets)",
        "short": "Partner's imports from world - by market",
        "patterns": ("imports-from-world-by-exporter",
                     "list_of_supplying_markets_for_a_product_imported_by_{p}"),
    },
    {
        "role": "table2",
        "query": "Imports from World, by Product, for {partner} "
                 "(all products imported by {partner})",
        "short": "Partner's imports from world - by product",
        "patterns": ("imports-from-world-by-product",
                     "list_of_products_imported_by_{p}"),
    },
    {
        "role": "table3",
        "query": "Exports to World, by Market, for {partner} "
                 "(all products, top importing markets)",
        "short": "Partner's exports to world - by market",
        "patterns": ("exports-to-world-by-importer",
                     "list_of_importing_markets_for_a_product_exported_by_{p}"),
    },
    {
        "role": "table4",
        "query": "Exports to World, by Product, for {partner} "
                 "(all products exported by {partner})",
        "short": "Partner's exports to world - by product",
        "patterns": ("exports-to-world-by-product",
                     "list_of_products_exported_by_{p}"),
    },
    {
        "role": "table5",
        "query": "Bilateral trade between Kenya and {partner} - Kenya's "
                 "EXPORTS to {partner}, by product",
        "short": "Kenya's exports to {partner} - by product (bilateral)",
        "patterns": ("kenyas-exports-to-{p}-by-product",
                     "bilateral_trade_between_kenya_and_{p}"),
    },
    {
        "role": "table6",
        "query": "Bilateral trade between Kenya and {partner} - Kenya's "
                 "IMPORTS from {partner}, by product",
        "short": "Kenya's imports from {partner} - by product (bilateral)",
        "patterns": ("kenyas-imports-from-{p}-by-product",
                     "bilateral_trade_between_kenya_and_{p}"),
    },
]
RAW_KEYWORDS = ("imports-from-world", "exports-to-world",
                "kenyas-imports-from", "kenyas-exports-to")
# classic ("previous") Trade Map download names -- same data, other naming
CLASSIC_RAW_KEYWORDS = ("list_of_supplying_markets_for_a_product_imported",
                        "list_of_products_imported",
                        "list_of_importing_markets_for_a_product_exported",
                        "list_of_products_exported",
                        "bilateral_trade_between_kenya_and")
# Classic query-level product-profile downloads carry a family in the filename
# ("List of exported products for the selected product (Meat and edible meat
# offal)").  These map to the product-profile loaders; a bare economy/partner
# matrix (e.g. a lone "List of products exported by Kenya") is NOT a family
# signal on its own.
CLASSIC_PRODUCT_KEYWORDS = (
    "list_of_exported_products_for_the_selected_product",
    "list_of_imported_products_for_the_selected_product",
    "list_of_exporters_for_the_selected_product",
    "list_of_importers_for_the_selected_product")
SERVICE_RAW_KEYWORDS = ("exported_services_for", "imported_services_for",
                        "services_exported_by", "services_imported_by",
                        "services_commercialized", "list_of_exporters_for",
                        "list_of_importers_for")
QUARTERLY_RAW_KEYWORDS = ("exports by hs destination", "imports by hs origin")
# Product-profile (matrix) downloads -- distinctive Trade Map matrix names.
PRODUCT_RAW_KEYWORDS = (
    "exporting-economies", "importing-economies",
    "products-exported-globally", "products-imported-globally",
    "kenyas-exports-to-world-by-product", "kenyas-exports-to-world-by-importer",
    "kenyas-imports-from-world-by-product", "kenyas-imports-from-world-by-exporter",
    "export_potential")
# Crafts "product group" downloads (Trade Map query-level workbooks).  A few of
# these names overlap the services/classic goods keywords (e.g.
# "list_of_exporters_for" is also in SERVICE_RAW_KEYWORDS), so the crafts check
# must run before the services/matrix auto-detects in _detect_mode.
CRAFTS_RAW_KEYWORDS = (
    "list_of_exported_products_for_the_selected_product_group",
    "list_of_exporters_for_the_selected_product_group",
    "list_of_importing_markets_for_a_product_group_exported_by_kenya")

# Classic Trade Map downloads ("... for the selected product (Family)") carry
# the product family in the filename itself - that is the leading, and most
# reliable, auto-detection signal for query-level uploads.
_CLASSIC_FAMILY_STEMS = (
    "list_of_exported_products_for_the_selected_product",
    "list_of_imported_products_for_the_selected_product",
    "list_of_exporters_for_the_selected_product",
    "list_of_importers_for_the_selected_product",
)
# Modern slugged names <-> classic stems for each by-product source the
# value-based scorer reads (the modern names are ITC's beta-product slugs).
_BY_PRODUCT_PROBES = {
    "world exports": ("products-exported-globally",
                      "list_of_exported_products_for_the_selected_product"),
    "world imports": ("products-imported-globally",
                      "list_of_imported_products_for_the_selected_product"),
    "kenya exports": ("kenyas-exports-to-world-by-product",
                      "list_of_products_exported_by_kenya"),
    "kenya imports": ("kenyas-imports-from-world-by-product",
                      "list_of_products_imported_by_kenya"),
}
_ANCHOR_PROBES = ("kenyas-exports-to-world-by-importer",
                  "list_of_importing_markets_for_a_product_exported_by_kenya")


def _norm_tokens(s):
    """Lower-cased word tokens of a family phrase."""
    return set(re.sub(r"[^a-z0-9 ]+", " ", (s or "").lower()).split())


def _family_from_filenames(uploads):
    """The product family named in classic Trade Map download filenames, or
    ``""``.  ``..._List_of_exported_products_for_the_selected_product_
    (Meat_and_edible_meat_offal).xlsx`` -> ``"Meat and edible meat offal"``."""
    counts = {}
    for p in gpp._walk_files(str(uploads)):
        base = os.path.basename(p).lower()
        if not any(stem in base for stem in _CLASSIC_FAMILY_STEMS):
            continue
        groups = re.findall(r"\(([^()]+)\)", os.path.splitext(base)[0])
        if not groups:
            continue
        phrase = " ".join(
            re.sub(r"[-_]+", " ", groups[-1]).lower().split())
        if phrase:
            counts[phrase] = counts.get(phrase, 0) + 1
    if not counts:
        return ""
    return max(counts.items(), key=lambda kv: (kv[1], len(kv[0])))[0]


def _family_candidates(cfg):
    """Phrases a config knows as its product family (detect_names first)."""
    out = [n for n in (cfg.get("detect_names") or []) if n]
    if cfg.get("family_title"):
        out.append(cfg["family_title"])
    if cfg.get("title"):
        out.append(cfg["title"])
    out += [sf.get("title") for sf in cfg.get("sub_families") or []
            if sf.get("title")]
    return out


def _family_overlap(fam_tokens, name):
    """Share of tokens the config name carries within the filename family
    phrase; a single-word family inside a long Trade Map name scores 1.0."""
    nt = _norm_tokens(name)
    if not fam_tokens or not nt:
        return 0.0
    inter = len(fam_tokens & nt)
    return inter / min(len(fam_tokens), len(nt))

@asynccontextmanager
async def _lifespan(app: FastAPI):
    _cleanup_old_jobs()
    yield

app = FastAPI(title="Trade Flow Report Generator", lifespan=_lifespan)


# ---------------------------------------------------------------------------
# Job storage helpers
# ---------------------------------------------------------------------------
def _new_job_dir():
    jd = JOBS_DIR / uuid.uuid4().hex
    (jd / "uploads").mkdir(parents=True)
    return jd


def _job_dir(job_id):
    """Resolve a validated job id to its directory (404 when unknown)."""
    if not _JOB_ID_RE.fullmatch(job_id or ""):
        raise HTTPException(404, "Job not found (or expired).")
    job_dir = JOBS_DIR / job_id
    if not job_dir.is_dir():
        raise HTTPException(404, "Job not found (or expired).")
    return job_dir


def _read_manifest(job_dir):
    """Load the job manifest, 404 when missing/corrupt."""
    p = Path(job_dir) / "manifest.json"
    if not p.exists():
        raise HTTPException(404, "Job not found (or expired).")
    try:
        return json.loads(p.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        raise HTTPException(404, "Job manifest unreadable (or expired).")


def _cfg_path(cfg_id):
    """Resolve a validated config id to its path (404 when unknown)."""
    if not cfg_id or cfg_id == "__auto__":
        raise HTTPException(404, "Config not found")
    if not _CFG_ID_RE.fullmatch(cfg_id):
        raise HTTPException(400, f"Invalid config id: {cfg_id!r}")
    p = CONFIG_DIR / f"{cfg_id}.json"
    if not p.exists():
        raise HTTPException(404, f"Config '{cfg_id}' not found")
    return p


def _cleanup_old_jobs():
    now = time.time()
    if not JOBS_DIR.exists():
        return
    for d in JOBS_DIR.iterdir():
        if d.is_dir():
            try:
                if now - d.stat().st_mtime > JOB_TTL_SECONDS:
                    shutil.rmtree(d, ignore_errors=True)
            except OSError:
                pass


def _extract_zip(zpath, dest):
    """Extract a zip safely (no path traversal), then flatten spreadsheets."""
    try:
        with zipfile.ZipFile(zpath) as z:
            for member in z.namelist():
                if member.endswith("/"):
                    continue
                safe = os.path.normpath(member)
                if safe.startswith("..") or os.path.isabs(safe):
                    raise HTTPException(400, f"Unsafe path inside zip: {member}")
                safe = os.path.join(os.path.dirname(safe),
                                    _clip_name(os.path.basename(safe)))
                out = dest / safe
                out.parent.mkdir(parents=True, exist_ok=True)
                with z.open(member) as src, out.open("wb") as dst:
                    shutil.copyfileobj(src, dst)
    except zipfile.BadZipFile:
        raise HTTPException(400, f"{zpath.name} is not a valid zip file")
    _flatten_xlsx(dest)
    os.remove(zpath)


def _flatten_xlsx(dest):
    """Move every supported spreadsheet from nested folders up to `dest` (zip case)."""
    for root, _, files in os.walk(dest):
        for f in files:
            if not f.lower().endswith(ALLOWED_EXT):
                continue
            src = os.path.join(root, f)
            if os.path.dirname(src) == str(dest):
                continue
            base, ext = os.path.splitext(f)
            dst = os.path.join(dest, _clip_name(f))
            i = 1
            while os.path.exists(dst):
                dst = os.path.join(dest, f"{base}__{i}{ext}")
                i += 1
            shutil.move(src, dst)
    for root, dirs, _ in os.walk(dest, topdown=False):
        for d in dirs:
            p = os.path.join(root, d)
            if not os.listdir(p):
                os.rmdir(p)
    for f in list(dest.iterdir()):
        if f.suffix.lower() not in (".xlsx", ".xlsm"):
            _convert_to_xlsx(f)


# ---------------------------------------------------------------------------
# Spreadsheet normalisation: anything -> .xlsx (openpyxl can only read .xlsx)
# ---------------------------------------------------------------------------
# _safe_sheet_name / _read_text_any / _sniff_spreadsheet_kind / HTMLTableParser
# / the per-format writers all live in xlsx_compat.py (single source of truth).


def _convert_to_xlsx(src):
    """Convert an .xls/.xlsb/.csv upload into a .xlsx openpyxl can read.

    Returns the new .xlsx path (the original file is deleted). Files that are
    already .xlsx/.xlsm are returned untouched.
    """
    if src.suffix.lower() in (".xlsx", ".xlsm"):
        return src
    ext = src.suffix.lower()
    if ext not in (".xls", ".xlsb", ".csv"):
        return src
    # The UNCTAD quarterly services file is ~60 MB and is only used for a
    # handful of aggregates.  Pre-process it once at upload time into a small
    # sidecar (dropping every unneeded entry); the pipeline reads that instead
    # of a giant converted .xlsx (which previously crashed strict UTF-8 reads).
    if ext in (".csv", ".xls", ".xlsb") and "totandcomservices" in src.name.lower():
        try:
            sidecar = make_services_tables.preprocess_unctad_services(
                src, src.with_name(src.stem + ".unctad.json"))
            if ext == ".csv":
                src.unlink()
            else:
                src.with_suffix(".xlsx").unlink(missing_ok=True)
            LOG.info("Pre-processed UNCTAD services file %s -> %s",
                     src.name, sidecar.name)
            return sidecar
        except Exception as e:  # noqa: BLE001 - keep raw if pre-processing fails
            LOG.warning("UNCTAD pre-processing failed for %s: %s; keeping raw",
                        src.name, e)
            return src
    out = src.with_suffix(".xlsx")
    i = 1
    while out.exists():
        out = src.with_name(f"{src.stem}__{i}.xlsx")
        i += 1
    try:
        if ext == ".xls":
            _write_xls_as_xlsx(src, out)
        elif ext == ".xlsb":
            _write_xlsb_as_xlsx(src, out)
        else:
            _write_csv_as_xlsx(src, out)
    except Exception as e:
        if out.exists():
            out.unlink()
        LOG.warning("Spreadsheet conversion failed for %s: %s", src.name, e)
        raise HTTPException(400, f"Could not read spreadsheet file {src.name} ({ext})")
    src.unlink()
    return out


# ---------------------------------------------------------------------------
# API
# ---------------------------------------------------------------------------
@app.get("/", include_in_schema=False)
def index():
    return FileResponse(STATIC_DIR / "index.html",
                        headers={"Cache-Control": "no-cache"})


app.mount("/static", StaticFiles(directory=str(STATIC_DIR)), name="static")


@app.get("/api/configs")
def api_configs():
    out = [{
        "id": "__auto__",
        "name": "Auto-detect country from data",
        "title": "",
        "year": 0,
    }]
    for p in sorted(CONFIG_DIR.glob("*.json")):
        if p.name == "config_template.json":
            continue
        try:
            cfg = gr.load_config(str(p))
            out.append({
                "id": p.stem,
                "name": cfg["country"]["name"],
                "title": cfg["country"]["title"],
                "year": cfg["report"].get("year", 0),
            })
        except Exception:
            continue
    return out


@app.get("/api/product-profiles")
def api_product_profiles():
    """List the standalone product-profile configs (coffee, crafts, ...)."""
    out = []
    for p, cfg in _product_profile_configs():
        out.append({
            "id": p.stem,
            "name": cfg.get("title") or cfg.get("family_title") or p.stem,
        })
    return out


def _partner_slug(name):
    return re.sub(r"[^a-z0-9]+", "_", (name or "").lower()).strip("_")


def _build_goods_checklist(partner):
    """Return the six ITC download steps for a partner (see ITC_GOODS_QUERIES)."""
    slug = _partner_slug(partner)
    items = []
    for q in ITC_GOODS_QUERIES:
        p = re.sub(r"[^a-z0-9]+", "_", (partner or "").lower()).strip("_")
        items.append({
            "role": q["role"],
            "query": q["query"].format(partner=partner),
            "short": q["short"].format(partner=partner),
            "patterns": [pat.format(p=p) for pat in q["patterns"]],
        })
    return items


@app.get("/api/goods/partner-setup")
def api_goods_partner_setup(partner: str = "", year: int = 0):
    """Prep a Goods flow for a partner by name (e.g. 'Saudi Arabia').

    Because ITC Trade Map has no public API (it is an Angular SPA), this does
    not download anything.  Instead it returns: (1) a ready-to-use config for
    the partner (created from the name, or the existing one reused) and (2) the
    exact six Trade Map queries to run plus the file-name patterns auto-detection
    expects of the downloaded workbooks.  Remaining requirement is a manual
    download from trademap.org."""

    partner = (partner or "").strip()
    if not partner:
        raise HTTPException(400, "Partner name is required (e.g. 'Saudi Arabia').")

    slug = _partner_slug(partner)
    cfg_path = CONFIG_DIR / f"{slug}.json"
    created = False
    if not cfg_path.exists():
        try:
            cfg, s = make_config.build_config_from_name(
                partner, year or None)
        except ValueError as e:
            raise HTTPException(400, str(e))
        cfg_path.write_text(json.dumps(cfg, ensure_ascii=False, indent=2),
                            encoding="utf-8")
        created = True

    cfg = gr.load_config(str(cfg_path))
    checklist = _build_goods_checklist(partner)

    # Best-effort: which of the six files already exist in the sample-data dir.
    sample_dir = BASE_DIR / "sample_data"
    found = {}
    if sample_dir.is_dir():
        lowlist = [f.lower().replace(" ", "_") for f in os.listdir(sample_dir)]
        for it in checklist:
            hit = next(
                (f for f in lowlist
                 if any(pat in f and _partner_slug(partner) not in f
                        for pat in it["patterns"])
                 or any(pat.format(p=_partner_slug(partner)) in f
                        for pat in it["patterns"])),
                None)
            found[it["role"]] = bool(hit)

    return {
        "ok": True,
        "partner": cfg["country"]["name"],
        "config": {"id": slug, "path": str(cfg_path), "created": created,
                   "year": cfg["report"].get("year", 0)},
        "checklist": checklist,
        "sample_found": found,
        "note": (
            "ITC Trade Map has no public API. Run each query below in the "
            "Trade Map web app, download the workbook, then upload them together "
            "for the Trade Brief. Filenames matching the patterns below are "
            "auto-detected and the partner config is reused automatically."),
    }


# ---------------------------------------------------------------------------
# Bundled sample datasets (Samplefiles/), grouped by app feature.
# ---------------------------------------------------------------------------
_SAMPLE_LABELS = {
    "services": "Services Brief",
    "quarterly": "Export Performance",
    "product": "Product Profile",
    "goods": "Trade Brief",
}

# Review periods of the Export Performance feature (report_type keeps its
# internal "quarterly" name for manifest/JS routing). Each entry maps to
# (period, quarter, half) for make_period_tables.build(); the half-year
# sub-picker sends semi1 (Jan-Jun) or semi2 (Jul-Dec).
_EXPORT_PERIODS = {
    "q1": ("quarter", 1, 1),
    "q2": ("quarter", 2, 1),
    "q3": ("quarter", 3, 1),
    "q4": ("quarter", 4, 1),
    "semi1": ("semi", 1, 1),
    "semi2": ("semi", 1, 2),
    "annual": ("year", 1, 1),
}

_EXPORT_PERIOD_LABELS = {
    "q1": "Q1 (Jan-Mar)",
    "q2": "Q2 (Apr-Jun)",
    "q3": "Q3 (Jul-Sep)",
    "q4": "Q4 (Oct-Dec)",
    "semi1": "Semi-Annual H1 (Jan-Jun)",
    "semi2": "Semi-Annual H2 (Jul-Dec)",
    "annual": "Annual (Jan-Dec)",
}


def _iter_spreadsheets(directory):
    """Yield every spreadsheet in a directory, skipping Office lock files
    (``~$``), hidden files and anything whose name is not a real upload."""
    if not directory.is_dir():
        return
    for f in sorted(directory.iterdir()):
        if not f.is_file():
            continue
        low = f.name.lower()
        if low.startswith(("~$", ".")):
            continue
        if not low.endswith(ALLOWED_EXT):
            continue
        yield f


def _sniff_sample_mode(files):
    """Map sample filenames to the app feature they describe.

    The services / quarterly / product keywords are tested first because some
    of them overlap the generic goods markers (e.g. ``kenyas-imports-from``
    also appears in product-profile matrix names).
    """
    low = [f.lower() for f in files]
    if any(any(k in n for k in SERVICE_RAW_KEYWORDS) for n in low):
        return "services"
    if any(any(k in n for k in QUARTERLY_RAW_KEYWORDS) for n in low):
        return "quarterly"
    if any(any(k in n for k in PRODUCT_RAW_KEYWORDS) for n in low):
        return "product"
    if any(any(k in n for k in RAW_KEYWORDS + CLASSIC_RAW_KEYWORDS) for n in low):
        return "goods"
    return "goods"


def _product_sample_config(child):
    """Pick the product-profile config a sample subfolder belongs to.

    Derived directly from the folder name (""crafts"" -> product_profile_crafts,
    ""Coffee"" -> product_profile_coffee), so the stored sample set always
    builds without relying on value-based family auto-detection (which some
    Trade Map baskets do not satisfy).
    """
    base = re.sub(r"[^a-z0-9]+", "_", (child or "").lower()).strip("_")
    if base:
        candidate = CONFIG_DIR / f"product_profile_{base}.json"
        if candidate.exists():
            return candidate.stem
    return "__auto__"


def _scan_sample_groups():
    """Enumerate the bundled sample datasets under ``Samplefiles/``.

    Every directory that directly holds spreadsheets becomes a sample group;
    nested folders (e.g. ``Product_Profiles/crafts`` and
    ``Product_Profiles/Coffee``) are each their own group.  Each group is
    pinned to the report type its filenames describe, so goods samples are
    only ever offered for the goods pipeline, services samples for the
    services pipeline, and so on.
    """
    groups = []

    def _walk(directory, trail):
        for child in sorted(directory.iterdir()):
            if not child.is_dir() or child.name.startswith((".", "~$")):
                continue
            names = [f.name for f in _iter_spreadsheets(child)]
            if names:
                mode = _sniff_sample_mode(names)
                config = "services_world" if mode == "services" else (
                    _product_sample_config(child.name) if mode == "product"
                    else "__auto__")
                groups.append({
                    "id": ("_".join(trail + [child.name]) if trail
                           else child.name),
                    "label": child.name.replace("_", " ").strip(),
                    "feature": mode,
                    "report_type": mode,
                    "description": _SAMPLE_LABELS.get(mode, mode),
                    "dir": str(child.relative_to(SAMPLE_DIR)).replace(os.sep, "/"),
                    "config": config,
                    "files": names,
                })
            _walk(child, trail + [child.name])

    if SAMPLE_DIR.is_dir():
        _walk(SAMPLE_DIR, [])
    return [g for g in groups if g["files"]]


@app.get("/api/samples")
def api_samples():
    """List the bundled sample datasets, grouped by the app feature they demo."""
    return JSONResponse(_scan_sample_groups())


@app.get("/api/sample-file/{group_id}/{filename:path}")
def api_sample_file(group_id: str, filename: str):
    """Stream a single file of a bundled sample dataset.

    The filename must be an exact member of the group's file list (plain
    basenames only), which also makes path-traversal impossible.
    """
    group = next((g for g in _scan_sample_groups() if g["id"] == group_id),
                 None)
    if group is None:
        raise HTTPException(404, f"Sample dataset '{group_id}' not found")
    if filename not in group["files"]:
        raise HTTPException(
            404, f"File '{filename}' not found in sample dataset '{group_id}'")
    p = (SAMPLE_DIR / group["dir"] / filename).resolve()
    if not p.is_file() or SAMPLE_DIR.resolve() not in p.parents:
        raise HTTPException(404, "File not found")
    return FileResponse(str(p), filename=os.path.basename(p))


def _detect_mode(uploads_dir, report_type="goods"):
    names = [f.lower() for f in os.listdir(uploads_dir) if f.lower().endswith(ALLOWED_EXT)]
    if not names:
        return None, "No Excel files were uploaded."
    if report_type == "quarterly":
        stems = {os.path.splitext(n)[0] for n in names}
        if {"exports", "imports"} <= stems:
            return None, (
                "The generated Exports.xlsx / Imports.xlsx workbooks use the "
                "old quarterly layout and cannot be read back as monthly "
                "data.\nUpload the raw monthly extracts instead - the KNBS "
                "workbook and/or the KRA extracts (...Exports by HS "
                "Destination... / ...Imports by HS Origin...). Every sheet "
                "is inspected individually; summary and pivot sheets are "
                "ignored.")
        # Any spreadsheet may carry the monthly data (KNBS workbook, KRA
        # extracts, workbooks mixing raw and pivot sheets): the loader
        # deciphers the file sheet by sheet and errors only when no sheet
        # yields usable Year/Month rows.
        return "quarterly_raw", ""
    # When the user explicitly selects a report type, honour that choice
    # before trying to auto-detect from filenames.
    # Craft product-group downloads are a product-profile dataset even though
    # a couple of their names overlap the services/classic keywords, so
    # recognise them first (works for "Product Profile" and defaults alike).
    if any(any(k in n for k in CRAFTS_RAW_KEYWORDS) for n in names):
        return "crafts", ""
    if report_type == "product":
        # Explicitly selected: require at least one recognisable matrix file,
        # or a classic query-level download (e.g. "List of exported products
        # for the selected product (...)").
        if any(any(k in n for k in PRODUCT_RAW_KEYWORDS) for n in names):
            return "product", ""
        if any(any(k in n for k in CLASSIC_PRODUCT_KEYWORDS) for n in names):
            return "product", ""
        uploaded = ", ".join(sorted(names))
        kw_list = ", ".join(PRODUCT_RAW_KEYWORDS + CLASSIC_PRODUCT_KEYWORDS)
        return None, (
            f"Could not recognise the product-profile upload set.\n"
            f"Uploaded files: {uploaded}\n"
            f"Expected ITC matrix filenames containing any of: {kw_list}")
    if report_type == "services":
        if any(any(k in n for k in SERVICE_RAW_KEYWORDS) for n in names):
            return "services_raw", ""
        uploaded = ", ".join(sorted(names))
        kw_list = ", ".join(SERVICE_RAW_KEYWORDS)
        return None, (
            f"Could not recognise the services upload set.\n"
            f"Uploaded files: {uploaded}\n"
            f"Expected filenames containing any of: {kw_list}")
    if report_type == "competitor":
        # Explicitly selected: need the per-product supplying-markets
        # downloads ("...-imports-from-world-by-exporter_<HS>.xlsx").
        if any(gca.is_supplier_download(n) for n in names):
            return "competitor", ""
        uploaded = ", ".join(sorted(names))
        return None, (
            f"Could not recognise the competitor-analysis upload set.\n"
            f"Uploaded files: {uploaded}\n"
            f"Expected per-product supplying-market downloads named like "
            f"'<partner>-imports-from-world-by-exporter_<HS>.xlsx' (ITC "
            f"Trade Map: List of supplying markets for a product imported "
            f"by <partner>, one file per product).")
    # Competitor files are recognised even without an explicit selection
    # (they would otherwise look like a goods "raw" set): when every workbook
    # is a per-product supplying-markets download the set is a competitor
    # analysis, not a goods report.
    if names and all(gca.is_supplier_download(n) for n in names):
        return "competitor", ""
    # The raw services files are recognised by their own distinctive keywords,
    # so the upload set can be detected even when the client did not explicitly
    # select the "Services Brief" report type.
    if any(any(k in n for k in SERVICE_RAW_KEYWORDS) for n in names):
        return "services_raw", ""
    # Auto-detect: distinctive all-countries/all-products matrix names mean a
    # product-profile dataset (coffee/crafts) rather than a goods report.
    if any(k in n for k in ("products-exported-globally",
                            "products-imported-globally",
                            "exporting-economies", "importing-economies")
           for n in names):
        return "product", ""
    if any(any(k in n for k in RAW_KEYWORDS + CLASSIC_RAW_KEYWORDS)
           for n in names):
        return "raw", ""
    if sum(1 for n in names if re.search(r"table\s*\d", n)) >= 6:
        return "ready", ""
    uploaded = ", ".join(sorted(names))
    raw_kw = ", ".join(RAW_KEYWORDS + CLASSIC_RAW_KEYWORDS)
    return None, (
        f"Could not recognise the upload set.\n"
        f"Uploaded files: {uploaded}\n"
        f"Expected raw filenames containing any of: {raw_kw}\n"
        f"Or ready-made files named 'Table 1' through 'Table 6' plus "
        f"'Figure 1 Trade Balance'.")


def _product_profile_configs():
    """Yield (path, cfg) for config/*.json files that describe a product profile.

    Product profile configs are standalone (no ``country`` key) and carry the
    ``family_title`` field used by the product-profile builder.
    """
    for p in sorted(CONFIG_DIR.glob("*.json")):
        if p.name == "config_template.json":
            continue
        try:
            cfg = gr.load_config(str(p))
        except Exception:
            continue
        if "family_title" in cfg and isinstance(cfg.get("country"), dict) is False:
            yield p, cfg


def _autodetect_product_profile(uploads, configs):
    """Pick the product-profile config whose family matches what was uploaded.

    Two independent signals are used, in order:

    1. *Filename family.*  Classic Trade Map downloads name the family in the
       file itself ("... for the selected product (Meat and edible meat
       offal).xlsx").  Matching that phrase against each config's family
       titles is authoritative - the family is literally in the download
       names - and it is what makes a "product" upload build without any
       manual profile selection.

    2. *Value coverage.*  When the filenames carry no family (modern slugged
       matrix names), the value of every product code in the by-product
       matrices is collected and each config is scored by the *share of trade
       value* it covers (the same HS 2022 hierarchical matching used by the
       generator).  Value, rather than code count, keeps a download that
       mixes in a few unrelated headings (e.g. rubber rows inside a dairy
       group) from hiding the real family.  Selection-total rows (whose value
       equals the sum of the other codes) are excluded, and a config that
       claims far more HS chapters than the upload actually contains is
       penalised, so a catch-all config cannot shadow a specific one.

    Returns ``(config_path, cfg)`` for the best match, or None when neither
    signal resolves to a config covering the uploaded family.
    """
    # ---- 1) family named in classic Trade Map filenames --------------------
    fam = _family_from_filenames(uploads)
    if fam:
        fam_tokens = _norm_tokens(fam)
        best = None
        for path, cfg in configs:
            ov = max((_family_overlap(fam_tokens, n)
                      for n in _family_candidates(cfg) or [""]), default=0.0)
            if ov >= float(cfg.get("detect_min_frac", 0.5)) and (
                    best is None or ov > best[0]):
                best = (ov, path, cfg)
        if best is not None:
            return best[1], best[2]
        return None

    # ---- 2) value coverage across the by-product matrices ------------------
    codes = set()
    # The selection anchor (e.g. "Horticulture Group") is the product code
    # that appears in the by-importer matrix; its row is the family aggregate,
    # not a real product, so it must not count toward the value covered by a
    # config.  Detect it the same way ProfileData does.
    anchor_codes = set()
    anchor_path = None
    for probe in _ANCHOR_PROBES:
        anchor_path = gpp._find_file(str(uploads), probe)
        if anchor_path is not None:
            break
    if anchor_path is not None:
        _, anchor_rows = gpp.load_matrix(anchor_path)
        counts = {}
        for r in anchor_rows:
            c = gpp.ProfileData._norm_code(r["product"])
            counts[c] = counts.get(c, 0) + 1
        if counts:
            anchor_codes.add(max(counts, key=counts.get))
    for _label, probes in _BY_PRODUCT_PROBES.items():
        path = None
        for probe in probes:
            path = gpp._find_file(str(uploads), probe)
            if path is not None:
                break
        if path is None:
            continue
        _, records = gpp.load_matrix(path)
        by_code = {}
        for r in records:
            c = gpp.ProfileData._norm_code(r["product"])
            if not c:
                continue
            v = max((y for y in r["years"].values() if y is not None),
                    default=0.0)
            by_code[c] = max(by_code.get(c, 0.0), v)
        # Drop selection-total rows: a code whose value equals the sum of all
        # the other codes in the same file (the basket aggregate).
        grand = sum(by_code.values())
        for c, v in list(by_code.items()):
            if grand and abs(v - (grand - v)) / grand < 0.01:
                by_code.pop(c)
        for c in list(by_code):
            if c in anchor_codes:
                by_code.pop(c)
        codes.update(by_code.items())
    if not codes:
        return None
    total_value = sum(v for _, v in codes)

    def _chapters(entries):
        """Distinct 2-digit HS chapters spanned by include-code entries."""
        out = set()
        for e in entries:
            e = str(e or "").strip()
            for part in (e.split("-") if "-" in e else [e]):
                c = gpp.ProfileData._norm_code(part)
                if len(c) >= 2:
                    out.add(c[:2])
        return out

    upload_chapters = {c[:2] for c, _ in codes}

    best = None  # (score, value_frac, matched_value, config_path, cfg)
    for path, cfg in configs:
        inc = cfg.get("include_codes")
        if not inc:
            continue
        tester = gpp.ProfileData.__new__(gpp.ProfileData)
        tester.include_codes = list(inc)
        matched_value = sum(v for c, v in codes if tester._code_ok(c))
        frac = matched_value / total_value if total_value else 0.0
        min_frac = float(cfg.get("detect_min_frac", 0.5))
        if frac < min_frac:
            continue
        cfg_chapters = _chapters(inc)
        if upload_chapters and cfg_chapters:
            union = upload_chapters | cfg_chapters
            jaccard = len(upload_chapters & cfg_chapters) / len(union)
        else:
            jaccard = 1.0
        score = frac * jaccard
        if best is None or score > best[0] or (
                score == best[0] and matched_value > best[2]):
            best = (score, frac, matched_value, path, cfg)
    if best is None:
        return None
    return best[3], best[4]


def _detect_reporter(excel_dir):
    """Return the country the data actually describes, from Table 1's title."""
    return make_config.read_table1_meta(excel_dir)[0]


def _norm(s):
    return re.sub(r"[^a-z0-9]+", "", (s or "").lower())


def _check_reporter_matches_config(cfg, excel_dir):
    detected = _detect_reporter(excel_dir)
    if not detected:
        return
    expected = cfg["country"]["name"]
    if _norm(detected) != _norm(expected):
        raise HTTPException(
            400,
            f"Country mismatch: the uploaded data is for '{detected}' but you "
            f"selected the '{expected}' configuration. Choose the matching "
            f"country in the dropdown (or pick 'Auto-detect country from "
            f"data') and try again.")


def _resolve_cfg(excel_dir, cfg_id, logs):
    """Resolve the config to use.

    ``cfg_id`` is normally a config/*.json stem. When it is empty or
    ``"__auto__"`` the country is detected from the data: an existing config
    for that country is reused, otherwise ``config/<country>.json`` is
    generated automatically from the data.
    """
    if cfg_id and cfg_id != "__auto__":
        cfg_path = _cfg_path(cfg_id)
        gr.cfg_path = os.path.abspath(str(cfg_path))
        cfg = gr.load_config(gr.cfg_path)
        _check_reporter_matches_config(cfg, excel_dir)
        return cfg, cfg_id

    detected = _detect_reporter(excel_dir)
    if not detected:
        raise HTTPException(
            400,
            "Could not detect the country from the uploaded data. Pick a "
            "country from the Country configuration dropdown (or upload the "
            "standard Table 1 with a country title).")
    for p in sorted(CONFIG_DIR.glob("*.json")):
        if p.name == "config_template.json":
            continue
        try:
            existing = gr.load_config(str(p))
            name = existing["country"]["name"]
        except Exception:
            continue
        if _norm(name) == _norm(detected):
            gr.cfg_path = os.path.abspath(str(p))
            logs.append(f"Country detected from data: {detected} -> using existing config '{p.stem}'")
            return gr.load_config(gr.cfg_path), p.stem
    new_id = make_config.create_config_file(excel_dir, str(CONFIG_DIR))[0]
    cfg_path = CONFIG_DIR / f"{new_id}.json"
    gr.cfg_path = os.path.abspath(str(cfg_path))
    logs.append(f"Country detected from data: {detected} -> generated config/{new_id}.json automatically")
    return gr.load_config(gr.cfg_path), new_id


def _run_pipeline(job_dir, cfg_id, top_n, mode, logs, promotion=False):
    uploads = job_dir / "uploads"
    tables = job_dir / "tables"
    charts = job_dir / "charts"

    if mode == "raw":
        logs.append("Mode: raw ITC downloads -> building Table 1-6 + Figure 1 ...")
        checks, problems = input_validation.validate_goods_inputs(str(uploads))
        for level, msg in checks:
            logs.append(input_validation.log_line(level, msg))
        if problems:
            raise PipelineError(
                "[ERROR] Input validation failed -- fix the source files "
                "listed above and re-run:\n"
                + "\n".join(f"  - {p}" for p in problems))
        try:
            make_tables.generate_tables(str(uploads), str(tables), top_n)
        except SystemExit as e:
            raise HTTPException(400, str(e))
        excel_dir = str(tables)
        logs.append(f"Tables written to {tables}")
    else:
        excel_dir = str(uploads)
        logs.append("Mode: ready-made tables -> skipping make_tables.py")

    cfg, cfg_id = _resolve_cfg(excel_dir, cfg_id, logs)

    os.makedirs(charts, exist_ok=True)
    report_name = f"KENYA-{gr.title_partner(cfg['country']['name'])} TRADE FLOW.docx"
    report_path = job_dir / report_name
    logs.append(f"Building report for {cfg['country']['name']} (charts -> {charts})")
    try:
        gr.build_report(cfg, excel_dir, str(report_path), str(charts))
    except SystemExit as e:
        raise HTTPException(400, str(e))
    logs.append(f"Report saved as {report_name}")

    manifest = {
        "mode": mode,
        "report_name": report_name,
        "excel_dir": os.path.relpath(excel_dir, job_dir),
        "config": cfg_id,
    }
    if promotion:
        promo_name = (os.path.splitext(report_name)[0]
                      + " Export Promotion.xlsx")
        promo_path = job_dir / promo_name
        try:
            import export_promotion_analysis
            export_promotion_analysis.build_from_dir(
                cfg, excel_dir, str(promo_path), extra_dirs=[str(uploads)])
            manifest["promotion_name"] = promo_name
            logs.append(f"Export promotion workbook saved as {promo_name}")
        except Exception as exc:
            logs.append(f"Warning: promotion workbook skipped: {exc}")
    (job_dir / "manifest.json").write_text(
        json.dumps(manifest), encoding="utf-8")
    return report_path, manifest


def _run_services_pipeline(job_dir, cfg_id, top_n, logs):
    uploads = job_dir / "uploads"
    tables = job_dir / "tables"
    charts = job_dir / "charts"

    logs.append("Mode: raw ITC services downloads -> building service tables ...")
    try:
        make_services_tables.generate_service_tables(
            str(uploads), str(tables), top_n)
    except SystemExit as e:
        raise HTTPException(400, str(e))
    excel_dir = str(tables)
    logs.append(f"Service tables written to {tables}")

    if not cfg_id or cfg_id == "__auto__":
        # Auto-select services_world config when report_type is services
        cfg_id = "services_world"

    cfg_path = _cfg_path(cfg_id)
    gr.cfg_path = os.path.abspath(str(cfg_path))
    cfg = gr.load_config(gr.cfg_path)

    os.makedirs(charts, exist_ok=True)
    report_name = f"KENYA-{gr.title_partner(cfg['country']['name'])} SERVICES TRADE FLOW.docx"
    report_path = job_dir / report_name
    logs.append(f"Building services report for {cfg['country']['name']} (charts -> {charts})")
    try:
        gsr.build_services_report(cfg, excel_dir, str(report_path), str(charts))
    except SystemExit as e:
        raise HTTPException(400, str(e))
    logs.append(f"Report saved as {report_name}")

    manifest = {
        "mode": "services_raw",
        "report_type": "services",
        "report_name": report_name,
        "excel_dir": os.path.relpath(excel_dir, job_dir),
        "config": cfg_id,
    }
    (job_dir / "manifest.json").write_text(
        json.dumps(manifest), encoding="utf-8")
    return report_path, manifest


def _run_product_pipeline(job_dir, cfg_id, top_n, logs):
    uploads = job_dir / "uploads"
    charts = job_dir / "charts"

    if not cfg_id or cfg_id == "__auto__":
        auto = _autodetect_product_profile(
            uploads, list(_product_profile_configs()))
        if auto is None:
            raise HTTPException(
                400, "Could not recognise which product family the files "
                     "describe. Select a product profile (coffee, crafts, ...) "
                     "manually.")
        cfg_path, cfg = auto
        cfg_id = cfg_path.stem
        logs.append(f"Auto-detected product profile: "
                    f"{cfg.get('family_title', cfg_id)}")
    else:
        cfg_path = _cfg_path(cfg_id)
        cfg = gr.load_config(str(cfg_path))

    os.makedirs(charts, exist_ok=True)
    data = gpp.ProfileData(str(uploads), cfg.get("include_codes"),
                           cfg.get("family_title"))
    logs.append(f"Loading ITC matrix files from {uploads}")
    logs.append(f"Profile: {cfg.get('family_title', cfg_id)} "
                f"| anchor {data.anchor_hs} | period {data.start_year}-{data.review_year}")
    for w in data.warnings:
        logs.append(f"  [warn] {w}")

    base = cfg.get("family_title") or cfg_id
    base = re.sub(r"[^A-Za-z0-9]+", " ", base).strip().upper()
    report_name = f"{base} PRODUCT PROFILE.docx"
    report_path = job_dir / report_name
    logs.append(f"Building product profile report (charts -> {charts})")
    try:
        doc = gpp.build_profile_document(cfg, data, str(charts))
        doc.save(str(report_path))
    except SystemExit as e:
        raise HTTPException(400, str(e))
    logs.append(f"Report saved as {report_name}")

    tables_name = f"{base} PRODUCT PROFILE TABLES.xlsx"
    tables_path = job_dir / tables_name
    try:
        gpp.write_excel_deliverable(cfg, data, str(tables_path))
        logs.append(f"Excel deliverable saved as {tables_name}")
    except Exception as e:
        tables_name = ""
        logs.append(f"Warning: Excel deliverable skipped: {e}")

    manifest = {
        "mode": "product",
        "report_type": "product",
        "report_name": report_name,
        "tables_name": tables_name,
        "config": cfg_id,
    }
    (job_dir / "manifest.json").write_text(
        json.dumps(manifest), encoding="utf-8")
    return report_path, manifest


def _run_crafts_pipeline(job_dir, cfg_id, top_n, logs):
    uploads = job_dir / "uploads"
    charts = job_dir / "charts"

    if not cfg_id or cfg_id == "__auto__":
        cfg_id = "product_profile_crafts"

    cfg_path = _cfg_path(cfg_id)
    cfg = gr.load_config(str(cfg_path))
    if not cfg.get("family_title"):
        raise HTTPException(400, "The selected config is not a product-profile "
                                 "config (no family_title).")
    logs.append(f"Profile: {cfg.get('family_title', cfg_id)}")

    os.makedirs(charts, exist_ok=True)
    data = gcr.CraftsData(str(uploads))
    logs.append(f"Loading craft product-group files from {uploads}")
    logs.append(f"Period {data.start_year}-{data.review_year} "
                f"| categories {len(data.order)}")
    for w in data.warnings:
        logs.append(f"  [warn] {w}")

    base = cfg.get("family_title") or cfg_id
    base = re.sub(r"[^A-Za-z0-9]+", " ", base).strip().upper()
    report_name = f"{base} PRODUCT PROFILE.docx"
    report_path = job_dir / report_name
    logs.append(f"Building crafts product profile report (charts -> {charts})")
    try:
        doc = gcr.build_crafts_report(cfg, str(uploads), str(report_path),
                                      str(charts))
        doc.save(str(report_path))
    except SystemExit as e:
        raise HTTPException(400, str(e))
    logs.append(f"Report saved as {report_name}")

    tables_name = f"{base} PRODUCT PROFILE TABLES.xlsx"
    tables_path = job_dir / tables_name
    try:
        wb = gcr.write_crafts_excel(cfg, str(uploads), str(tables_path))
        if wb is None:
            raise RuntimeError("no usable data loaded for the workbook")
        logs.append(f"Excel deliverable saved as {tables_name}")
    except Exception as e:
        tables_name = ""
        logs.append(f"Warning: Excel deliverable skipped: {e}")

    manifest = {
        "mode": "crafts",
        "report_type": "crafts",
        "report_name": report_name,
        "tables_name": tables_name,
        "config": cfg_id,
    }
    (job_dir / "manifest.json").write_text(
        json.dumps(manifest), encoding="utf-8")
    return report_path, manifest


def _run_competitor_pipeline(job_dir, logs):
    """Standalone competitor analysis from per-product supplying-markets
    downloads. Independent of the goods flow on purpose: the products to
    analyse are only known after a first trade-flow run."""
    uploads = job_dir / "uploads"
    logs.append("Mode: per-product supplying-market downloads -> "
                "building competitor analysis ...")
    try:
        result = gca.generate(str(uploads), str(job_dir))
    except SystemExit as e:
        raise HTTPException(400, str(e))
    except Exception as e:
        raise HTTPException(400, f"Competitor analysis failed: {e}")
    products = result["products"]
    logs.append("Products analysed: %d (review year %s)"
                % (len(products), result["year"]))
    for p in products:
        logs.append("  %d. HS %s - imported %s"
                    % (p["rank"], p["code"],
                       "{:,.0f}".format(p["world"]) if p["world"] is not None
                       else "n/a"))
    report_name = os.path.basename(result["docx"])
    manifest = {
        "mode": "competitor",
        "report_type": "competitor",
        "report_name": report_name,
        "png_name": os.path.basename(result["png"]),
        "xlsx_name": os.path.basename(result["xlsx"]),
        "year": result["year"],
        "partner": result["partner"],
    }
    logs.append("Saved %s (+ .png, .xlsx) in the job folder" % report_name)
    (job_dir / "manifest.json").write_text(
        json.dumps(manifest), encoding="utf-8")
    return job_dir / report_name, manifest


def _run_export_pipeline(job_dir, top_n, mode, period_key, logs):
    """Export Performance: raw monthly extracts -> period tables + report.

    Every sheet of every upload is inspected individually (period_data);
    summary/pivot sheets without Year/Month columns are ignored, and the
    review year comes from the data (the newest year present), compared
    with the same sub-period one year earlier.
    """
    uploads = job_dir / "uploads"
    tables = job_dir / "tables"
    period, quarter, half = _EXPORT_PERIODS[period_key]
    logs.append("Mode: Export Performance - review period: %s"
                % _EXPORT_PERIOD_LABELS[period_key])

    # Load once up front so the job log shows which file/sheets were read;
    # make_period_tables.build() reuses this disk cache for its own call.
    df = period_data.load_monthly(str(uploads))
    for unit, n in sorted(df.groupby("SourceUnit").size().items()):
        logs.append("  read %s: %s rows" % (unit, format(n, ",")))
    for w in period_data.WARNINGS:
        logs.append("  [warn] %s" % w)
    year = int(df["Year"].max())
    chk = period_data.period_spec(period, year, half=half, quarter=quarter)
    covered = set(df.loc[df["Year"] == year, "Month"])
    missing = [m for m in chk["months"] if m not in covered]
    if len(missing) == len(chk["months"]):
        # Every month of the review period is absent (e.g. a quarter beyond
        # the months the uploads cover) - fail with guidance instead of
        # building an all-zero report.
        raise HTTPException(
            400,
            "No data for {} {} in the uploaded files (the {} rows cover {}). "
            "Choose a review period covered by the data, or upload extracts "
            "that include it.".format(
                chk["span"], year, year,
                ", ".join(period_data.MONTH_NAMES[m] for m in sorted(covered))))
    if missing:
        logs.append("Note: no %s rows for part of %s %s" % (
            year, chk["span"], year))
    logs.append("Data years: %s; review year %d vs %d"
                % (", ".join(str(y) for y in sorted(set(df["Year"]))),
                   year, year - 1))

    try:
        xlsx_path, spec = mpt.build(
            str(uploads), str(tables), period, year=year,
            half=half, quarter=quarter,
            top_products=top_n, top_markets=top_n)
    except SystemExit as e:
        raise HTTPException(400, str(e))
    logs.append("Period workbooks written to %s" % tables.name)

    report_name = "KENYA EXPORT PERFORMANCE IN %s %d-%d.docx" % (
        spec["span"].upper(), spec["prev_year"], spec["year"])
    try:
        gpr.build_report(xlsx_path, str(job_dir / report_name))
    except SystemExit as e:
        raise HTTPException(400, str(e))
    logs.append("Report saved as %s" % report_name)

    manifest = {
        "mode": mode,
        "report_type": "quarterly",
        "report_name": report_name,
        "excel_dir": os.path.relpath(str(tables), job_dir),
        "config": "",
        "period": period_key,
    }
    (job_dir / "manifest.json").write_text(
        json.dumps(manifest), encoding="utf-8")
    return job_dir / report_name, manifest


_SERVER_START = time.time()
# One pipeline at a time: the builds are CPU/RAM heavy (quarterly KRA
# extracts peak around 400 MB), so parallel jobs on a small host would
# exhaust memory and get the service killed mid-request.
_PIPELINE_LOCK = threading.Lock()


class _JobLog(list):
    """A list that also appends every line to the job's log.txt so the
    browser can follow progress through /api/status while the worker runs."""

    def __init__(self, job_dir):
        super().__init__()
        self._path = Path(job_dir) / "log.txt"

    def append(self, line):
        super().append(line)
        with self._path.open("a", encoding="utf-8") as f:
            f.write(str(line) + "\n")


def _read_job_log(job_dir):
    p = Path(job_dir) / "log.txt"
    if not p.exists():
        return []
    return p.read_text(encoding="utf-8", errors="replace").splitlines()


def _write_job_error(job_dir, message, traceback_text=None):
    """Persist a job failure: a short user-facing message plus (optionally)
    the full traceback.  /api/status only ever returns the short message."""
    (job_dir / "error.txt").write_text(
        message if not traceback_text else f"{message}\n\n{traceback_text}",
        encoding="utf-8")
    (job_dir / "error_public.txt").write_text(message, encoding="utf-8")


def _job_worker(job_dir, job_fn):
    logs = _JobLog(job_dir)
    try:
        # Log before taking the lock: if another job is still running the
        # worker blocks here, and without this line the UI would show a job
        # that is "running" forever with an empty log.
        logs.append("Job received; waiting for the renderer to become free ...")
        with _PIPELINE_LOCK:
            logs.append("Renderer acquired; build starting.")
            report_path, manifest = job_fn(logs)
        (job_dir / "manifest.json").write_text(
            json.dumps(manifest), encoding="utf-8")
        logs.append(f"Report ready: {manifest['report_name']}")
        LOG.info("Background job %s finished -> %s",
                 job_dir.name, manifest["report_name"])
    except HTTPException as e:
        LOG.error("Background job %s failed: %s", job_dir.name, e.detail)
        _write_job_error(job_dir, str(e.detail))
    except PipelineError as e:
        LOG.error("Background job %s failed: %s", job_dir.name, e)
        _write_job_error(job_dir, str(e))
    except SystemExit as e:
        LOG.error("Background job %s aborted: %s", job_dir.name, e)
        _write_job_error(job_dir, f"Pipeline aborted: {e}")
    except Exception as e:
        tb = traceback.format_exc()
        LOG.error("Background job %s failed:", job_dir.name)
        LOG.error("%s", tb)
        _write_job_error(job_dir, f"Pipeline failed: {e}", tb)


@app.get("/api/status/{job_id}")
def api_status(job_id: str):
    """Poll a background job started by POST /api/run.

    States: running | done | error. A job whose meta.json predates this
    server process must have been interrupted by a restart -> error.
    """
    job_dir = _job_dir(job_id)

    log_lines = _read_job_log(job_dir)
    err_path = job_dir / "error_public.txt"
    if not err_path.exists():
        err_path = job_dir / "error.txt"
    man_path = job_dir / "manifest.json"
    if err_path.exists():
        return {"state": "error", "detail": err_path.read_text(encoding="utf-8"),
                "log": log_lines}
    if man_path.exists():
        m = json.loads(man_path.read_text(encoding="utf-8"))
        resp = {"state": "done", "job_id": job_id,
                "report_name": m.get("report_name"),
                "report_url": f"/api/download/{job_id}",
                "tables_url": f"/api/tables/{job_id}",
                "mode": m.get("mode"), "log": log_lines}
        if m.get("report_type") in ("product", "crafts"):
            resp["product_tables_url"] = f"/api/product-tables/{job_id}"
            resp.pop("tables_url", None)
        if m.get("report_type") == "competitor":
            resp["png_url"] = f"/api/competitor-file/{job_id}/png"
            resp["xlsx_url"] = f"/api/competitor-file/{job_id}/xlsx"
            resp.pop("tables_url", None)
        if m.get("promotion_name"):
            resp["promotion_url"] = f"/api/promotion/{job_id}"
        return resp
    try:
        started = json.loads((job_dir / "meta.json").read_text())["started"]
    except Exception:
        started = None
    if started is not None and started < _SERVER_START - 2:
        return {"state": "error", "log": log_lines,
                "detail": "The server restarted while this job was "
                          "running. Please upload the files again."}
    return {"state": "running", "log": log_lines}


def _write_upload(dst, upload, budget):
    """Stream an upload to disk, enforcing the per-job byte budget.

    ``budget`` is a one-item list holding the remaining allowed bytes; it is
    decremented as bytes arrive so the whole job (files + zip) shares one cap.
    """
    remaining = budget[0]
    written = 0
    with dst.open("wb") as f:
        while True:
            chunk = upload.file.read(1024 * 1024)
            if not chunk:
                break
            written += len(chunk)
            remaining -= len(chunk)
            if written > MAX_UPLOAD_BYTES:
                raise HTTPException(
                    400, f"File too large: {upload.filename or 'upload'} "
                         f"(limit {MAX_UPLOAD_BYTES // (1024 * 1024)} MB "
                         f"per file).")
            if remaining < 0:
                raise HTTPException(
                    400, "Upload too large in total "
                         f"(limit {MAX_TOTAL_UPLOAD_BYTES // (1024 * 1024)} "
                         "MB per job).")
            f.write(chunk)
    budget[0] = remaining
    return written


# NOTE: sync def on purpose — FastAPI runs sync endpoints in a threadpool, so
# the (potentially hundreds of MB) upload/conversion work here never blocks
# the event loop that serves /api/status polling for other clients.
@app.post("/api/run")
def api_run(config: str = Form("__auto__"), top: int = Form(20),
            report_type: str = Form("goods"),
            period: str = Form(""),
            promotion: bool = Form(False),
            files: list[UploadFile] = File(default=None),
            zip_upload: UploadFile | None = File(default=None, alias="zip")):
    files = files or []
    if not files and zip_upload is None:
        raise HTTPException(400, "Please select files (or a zip) to upload.")
    if top < 1 or top > 100:
        raise HTTPException(400, "top must be between 1 and 100")
    if config and config != "__auto__" and not _CFG_ID_RE.fullmatch(config):
        raise HTTPException(400, f"Invalid config id: {config!r}")
    if report_type not in ("goods", "services", "quarterly", "product",
                           "competitor"):
        raise HTTPException(400, "report_type must be 'goods', 'services', "
                                 "'quarterly', 'product' or 'competitor'")
    if report_type == "quarterly" and period not in _EXPORT_PERIODS:
        # Force an explicit choice: a default would silently run a full-year
        # report when only one quarter was wanted.
        raise HTTPException(400, "Please choose a review period: Q1, Q2, "
                                 "Q3, Q4, Semi-Annual or Annual.")

    job_dir = _new_job_dir()
    budget = [MAX_TOTAL_UPLOAD_BYTES]
    try:
        for up in files:
            name = _clip_name(up.filename or "")
            if not name.lower().endswith(ALLOWED_EXT):
                raise HTTPException(400,
                                    f"Unsupported file type: {up.filename}")
            dst = job_dir / "uploads" / name
            if dst.exists():
                base, ext = os.path.splitext(name)
                i = 1
                while dst.exists():
                    dst = job_dir / "uploads" / f"{base}__{i}{ext}"
                    i += 1
            _write_upload(dst, up, budget)
            try:
                _convert_to_xlsx(dst)
            except HTTPException as exc:
                # Keep the raw file so the loader can still try it (an
                # unneeded or unsupported extra file must not abort the
                # whole upload set).
                LOG.warning("Keeping raw upload %s after conversion "
                            "failure: %s", name, exc.detail)
        if zip_upload is not None:
            zname = os.path.basename(zip_upload.filename or "upload.zip")
            zpath = job_dir / "uploads" / zname
            _write_upload(zpath, zip_upload, budget)
            _extract_zip(zpath, job_dir / "uploads")

        mode, err = _detect_mode(job_dir / "uploads", report_type)
        if mode is None:
            raise HTTPException(400, err)
        # The mode is derived from the actual uploads: trust it over the
        # client-supplied report_type so a raw services upload set is always
        # routed to the services pipeline even if the caller forgot to select
        # the services report type.
        if mode.startswith("services"):
            report_type = "services"
        elif mode.startswith("quarterly"):
            report_type = "quarterly"
        elif mode == "product":
            report_type = "product"
        elif mode == "crafts":
            report_type = "crafts"
        elif mode == "competitor":
            report_type = "competitor"
        elif mode in ("raw", "ready"):
            report_type = "goods"

        LOG.info("POST /api/run config=%s top=%s mode=%s report_type=%s "
                 "period=%s job=%s",
                 config, top, mode, report_type, period, job_dir.name)

        if report_type == "quarterly":
            job_fn = lambda logs: _run_export_pipeline(       # noqa: E731
                job_dir, top, mode, period, logs)
        elif report_type == "services":
            job_fn = lambda logs: _run_services_pipeline(    # noqa: E731
                job_dir, config, top, logs)
        elif report_type == "product":
            job_fn = lambda logs: _run_product_pipeline(     # noqa: E731
                job_dir, config, top, logs)
        elif report_type == "crafts":
            job_fn = lambda logs: _run_crafts_pipeline(      # noqa: E731
                job_dir, config, top, logs)
        elif report_type == "competitor":
            job_fn = lambda logs: _run_competitor_pipeline(   # noqa: E731
                job_dir, logs)
        else:
            job_fn = lambda logs: _run_pipeline(             # noqa: E731
                job_dir, config, top, mode, logs, promotion=promotion)

        # The heavy build runs in a background thread; the request answers
        # immediately and the page polls /api/status/<job>. Long synchronous
        # responses are what got killed by proxies / restarts before.
        (job_dir / "meta.json").write_text(
            json.dumps({"started": time.time()}), encoding="utf-8")
        threading.Thread(target=_job_worker, args=(job_dir, job_fn),
                         daemon=True).start()
        return {
            "job_id": job_dir.name,
            "status_url": f"/api/status/{job_dir.name}",
            "mode": mode,
        }
    except HTTPException:
        shutil.rmtree(job_dir, ignore_errors=True)
        raise
    except Exception as e:
        LOG.error("Upload handling failed for config=%s top=%s job=%s",
                  config, top, job_dir.name)
        LOG.error("%s", traceback.format_exc())
        shutil.rmtree(job_dir, ignore_errors=True)
        raise HTTPException(500, f"Upload failed: {e}")


@app.get("/api/download/{job_id}")
def api_download(job_id: str):
    job_dir = _job_dir(job_id)
    manifest = _read_manifest(job_dir)
    report = job_dir / manifest["report_name"]
    if not report.exists():
        raise HTTPException(404, "Report not found.")
    return FileResponse(str(report), filename=report.name,
                        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document")


@app.get("/api/promotion/{job_id}")
def api_promotion(job_id: str):
    """Download the additive export-promotion workbook (if requested)."""
    job_dir = _job_dir(job_id)
    manifest = _read_manifest(job_dir)
    if "promotion_name" not in manifest:
        raise HTTPException(404, "This job has no export-promotion workbook.")
    promo = job_dir / manifest["promotion_name"]
    if not promo.exists():
        raise HTTPException(404, "Promotion workbook not found.")
    return FileResponse(str(promo), filename=promo.name,
                        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.get("/api/tables/{job_id}")
def api_tables(job_id: str):
    job_dir = _job_dir(job_id)
    manifest = _read_manifest(job_dir)
    src = job_dir / manifest["excel_dir"]
    if not src.is_dir():
        raise HTTPException(404, "Tables not found.")
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for f in sorted(src.glob("*.xlsx")):
            z.write(str(f), arcname=f.name)
    buf.seek(0)
    return Response(content=buf.getvalue(),
                    media_type="application/zip",
                    headers={"Content-Disposition": f"attachment; filename=tables-{job_id}.zip"})


@app.get("/api/product-tables/{job_id}")
def api_product_tables(job_id: str):
    """Download the Excel deliverable of a product-profile job."""
    job_dir = _job_dir(job_id)
    manifest = _read_manifest(job_dir)
    tables = job_dir / manifest.get("tables_name", "")
    if not tables.exists():
        raise HTTPException(404, "Excel deliverable not found.")
    return FileResponse(
        str(tables), filename=tables.name,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.get("/api/competitor-file/{job_id}/{kind}")
def api_competitor_file(job_id: str, kind: str):
    """Download a competitor-analysis deliverable (png / xlsx / docx)."""
    job_dir = _job_dir(job_id)
    manifest = _read_manifest(job_dir)
    if manifest.get("report_type") != "competitor":
        raise HTTPException(404, "Not a competitor-analysis job.")
    key = {"png": "png_name", "xlsx": "xlsx_name", "docx": "report_name"}.get(kind)
    if key is None:
        raise HTTPException(400, "kind must be png, xlsx or docx.")
    path = job_dir / manifest.get(key, "")
    if not path.exists():
        raise HTTPException(404, "Deliverable not found.")
    media = {
        "png": "image/png",
        "xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    }[kind]
    return FileResponse(str(path), filename=path.name, media_type=media)


# ---------------------------------------------------------------------------
# Convenience: run with `python webapp/main.py`
# ---------------------------------------------------------------------------
def main():
    import uvicorn
    host = os.environ.get("HOST", "127.0.0.1")
    port = int(os.environ.get("PORT", "8000"))
    print(f"Trade Flow Report Generator -> http://{host}:{port}")
    print("Upload the six raw ITC files (or a zip), or the seven ready tables.")
    uvicorn.run(app, host=host, port=port, log_level="info")


if __name__ == "__main__":
    main()
