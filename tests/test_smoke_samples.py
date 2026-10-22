"""Golden-file smoke tests: run every pipeline end-to-end on Samplefiles/.

Each test drives the same entry points the webapp uses (library calls or the
CLI) against the bundled demo datasets and checks the outputs a report
consumes: the expected files exist, and every formula in the workbooks has a
readable cached value (the report readers use data_only=True).

These are the slow tests (~2-3 min total).  Run just the fast unit tests with
``pytest -m "not smoke"`` after registering the marker (see pytest.ini).
"""

import re
import subprocess
import sys
from pathlib import Path

import openpyxl
import pytest

ROOT = Path(__file__).resolve().parent.parent
SAMPLES = ROOT / "Samplefiles"

_CELL_REF = re.compile(r"\b([A-Z]{1,3})(\d+)\b")

pytestmark = pytest.mark.smoke


def run_cli(args, timeout=600):
    r = subprocess.run([sys.executable, *map(str, args)], cwd=ROOT,
                       capture_output=True, text=True, timeout=timeout)
    assert r.returncode == 0, (
        f"command failed: {args}\n--- stdout ---\n{r.stdout[-3000:]}"
        f"\n--- stderr ---\n{r.stderr[-3000:]}")
    return r


def assert_formula_caches_readable(path):
    """Formula cells must carry cached values (readers use data_only=True).

    A formula may legitimately be uncached when one of the cells it
    references is itself empty -- e.g. the services share column while the
    latest year's per-category detail is not published yet (the cached
    result would be None).  A formula whose inputs are all present but has
    no cached value is a cache-injection failure.
    """
    wf = openpyxl.load_workbook(str(path))
    wd = openpyxl.load_workbook(str(path), data_only=True)
    checked = 0
    bad = []
    for sheet in wf.sheetnames:
        ws_f, ws_d = wf[sheet], wd[sheet]
        for row in ws_f.iter_rows():
            for cell in row:
                if not (isinstance(cell.value, str)
                        and cell.value.startswith("=")):
                    continue
                checked += 1
                if ws_d[cell.coordinate].value is not None:
                    continue
                inputs = _CELL_REF.findall(cell.value)
                inputs_empty = any(
                    ws_d[f"{col}{r}"].value is None for col, r in inputs)
                if not inputs_empty:
                    bad.append((sheet, cell.coordinate, cell.value))
    wf.close()
    wd.close()
    assert checked > 0, f"no formulas found in {path}"
    assert not bad, (
        f"{path}: formulas with present inputs but no cached value "
        f"(first: {bad[:10]})")


def assert_docx(path, min_bytes=50_000):
    p = Path(path)
    assert p.exists(), f"missing report: {p}"
    assert p.stat().st_size >= min_bytes, f"suspiciously small: {p}"


# --------------------------------------------------------------------------
# Goods flow (USA sample): Table 1-6 + Figure 1 + combined workbook
# --------------------------------------------------------------------------
def test_golden_goods_tables(tmp_path):
    out = tmp_path / "tables"
    run_cli(["make_tables.py", "--excel-dir", SAMPLES / "USA_Goods_Flow",
             "--out-dir", out, "--top", "20"])
    expected = [
        "Table 1 United States of America Import Source Markets.xlsx",
        "Table 2 United States of America Lead Imports.xlsx",
        "Table 3 United States of America Lead Export Destinations.xlsx",
        "Table 4 United States of America Lead Export Products.xlsx",
        "Table 5 Kenya Exports to United States of America.xlsx",
        "Table 6 Kenya Top Imports from United States of America.xlsx",
        "Table 7 Kenya Export Composition and Market Alignment.xlsx",
        "Figure 1 Trade Balance.xlsx",
        "All Tables.xlsx",
    ]
    for name in expected:
        assert (out / name).exists(), f"missing {name}"
    assert_formula_caches_readable(out / "All Tables.xlsx")


def test_golden_goods_report(tmp_path):
    tables = tmp_path / "tables"
    run_cli(["make_tables.py", "--excel-dir", SAMPLES / "USA_Goods_Flow",
             "--out-dir", tables, "--top", "20"])
    docx = tmp_path / "report.docx"
    run_cli(["generate_report.py", "--excel-dir", tables,
             "--config", ROOT / "config" / "united_states_of_america.json",
             "--output", docx, "--tmp", tmp_path / "tmp"])
    assert_docx(docx)
    assert (tmp_path / "report TABLES.xlsx").exists()


# --------------------------------------------------------------------------
# Services flow (world sample): 14 tables + balance + combined workbook
# --------------------------------------------------------------------------
def test_golden_services_tables(tmp_path):
    out = tmp_path / "services"
    run_cli(["make_services_tables.py",
             "--excel-dir", SAMPLES / "Services_Flow",
             "--out-dir", out, "--top", "20"])
    for name in [
        "Table 1 Top Global Service Exporters.xlsx",
        "Table 14 Global Service Exports by Category.xlsx",
        "Figure 1 Services Balance.xlsx",
        "All Service Tables.xlsx",
    ]:
        assert (out / name).exists(), f"missing {name}"
    assert_formula_caches_readable(out / "All Service Tables.xlsx")


def test_golden_services_report(tmp_path):
    out = tmp_path / "services"
    run_cli(["make_services_tables.py",
             "--excel-dir", SAMPLES / "Services_Flow",
             "--out-dir", out, "--top", "20"])
    docx = tmp_path / "services.docx"
    run_cli(["generate_services_report.py", "--excel-dir", out,
             "--config", ROOT / "config" / "services_world.json",
             "--output", docx, "--tmp", tmp_path / "tmp"])
    assert_docx(docx)


# --------------------------------------------------------------------------
# Export performance (period) flow: KRA .xls extracts -> tables -> report
# --------------------------------------------------------------------------
def test_golden_period_tables_and_report(tmp_path):
    tables = tmp_path / "period"
    run_cli(["make_period_tables.py",
             "--data-dir", SAMPLES / "Quartely_Performance",
             "--out-dir", tables, "--period", "semi",
             "--year", "2026", "--half", "1"])
    xlsx = list(tables.glob("PERFORMANCE TABLES*.xlsx"))
    assert xlsx, f"no performance tables written to {tables}"
    docx = tmp_path / "period.docx"
    run_cli(["generate_period_report.py", "--tables", xlsx[0],
             "--output", docx])
    assert_docx(docx)


# --------------------------------------------------------------------------
# Product profile flow (crafts matrices) -> report + Excel deliverable
# --------------------------------------------------------------------------
def test_golden_product_profile(tmp_path):
    docx = tmp_path / "crafts.docx"
    run_cli(["generate_product_profile.py",
             "--config", ROOT / "config" / "product_profile_crafts.json",
             "--excel-dir", SAMPLES / "Product_Profiles" / "crafts",
             "--output", docx, "--tmp", tmp_path / "tmp"])
    assert_docx(docx)
    assert (tmp_path / "crafts TABLES.xlsx").exists()
