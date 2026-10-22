"""Tests for input_validation: pre-flight checks on the raw ITC downloads."""

import shutil
from pathlib import Path

import openpyxl
import pytest

import input_validation as iv
from pipeline_errors import PipelineError

ROOT = Path(__file__).resolve().parent.parent

BETA = ("reporterCd", "reporterLabel", "partnerCd", "partnerLabel",
        "productCd", "productLabel")
KCD, KENYA = "404", "Kenya"
FCD, FRANCE = "251", "France"
DCD, GERMANY = "276", "Germany"
YEARS = [2021, 2022, 2023, 2024, 2025]

FILENAMES = {
    "table1": "frances-imports-from-world-by-exporter_all.xlsx",
    "table2": "frances-imports-from-world-by-product_all.xlsx",
    "table3": "frances-exports-to-world-by-importer_all.xlsx",
    "table4": "frances-exports-to-world-by-product_all.xlsx",
    "table5": "kenyas-exports-to-france-by-product_all.xlsx",
    "table6": "kenyas-imports-from-france-by-product_all.xlsx",
}

SAMPLE = str(ROOT / "Samplefiles" / "USA_Goods_Flow")


def _vals(base, n):
    return [base + i for i in range(n)]


def _market_rows(rep_cd, rep_label, n, partner_cd="000",
                 partner_label="World"):
    return [
        [rep_cd, rep_label, partner_cd, partner_label, "TOTAL",
         "All Products", *_vals(1_000_000, n)],
        [rep_cd, rep_label, "276", GERMANY, "TOTAL", "All Products",
         *_vals(500_000, n)],
        [rep_cd, rep_label, "056", "Belgium", "TOTAL", "All Products",
         *_vals(300_000, n)],
    ]


def _product_rows(rep_cd, rep_label, partner_cd, partner_label, n):
    return [
        [rep_cd, rep_label, partner_cd, partner_label, "TOTAL",
         "All Products", *_vals(1_000_000, n)],
        [rep_cd, rep_label, partner_cd, partner_label, "8471",
         "Data processing machines", *_vals(400_000, n)],
        [rep_cd, rep_label, partner_cd, partner_label, "0901",
         "Coffee, whether or not roasted", *_vals(200_000, n)],
    ]


def _defaults(role, years):
    n = len(years)
    if role in ("table1", "table3"):
        return _market_rows(FCD, FRANCE, n)
    if role in ("table2", "table4"):
        return _product_rows(FCD, FRANCE, "000", "World", n)
    return _product_rows(KCD, KENYA, FCD, FRANCE, n)


def build(tmp_path, rows=None, role_years=None, units=None):
    """Write a complete six-file source set; per-role dicts override."""
    rows = rows or {}
    role_years = role_years or {}
    units = {} if units is None else units
    tmp_path.mkdir(parents=True, exist_ok=True)
    for role, fname in FILENAMES.items():
        years = role_years.get(role, YEARS)
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Sheet1"
        unit = units.get(role, "USD Thousand")
        ws.append(list(BETA) + [f"{y} ({unit})" if unit else str(y)
                                for y in years])
        for r in rows.get(role) or _defaults(role, years):
            ws.append(list(r))
        wb.save(str(tmp_path / fname))
    return tmp_path


def levels(checks):
    return [lv for lv, _ in checks]


def test_happy_set_passes(tmp_path):
    build(tmp_path)
    checks, problems = iv.validate_goods_inputs(str(tmp_path))
    assert problems == []
    assert "error" not in levels(checks)
    assert "warn" not in levels(checks)
    assert any("Report window: 2021-2025 (5 years)" in m for _, m in checks)


def test_missing_source_file_is_an_error(tmp_path):
    build(tmp_path)
    (tmp_path / FILENAMES["table4"]).unlink()
    checks, problems = iv.validate_goods_inputs(str(tmp_path))
    assert problems and "Missing required source files" in problems[0]
    assert checks[0][0] == "error"


def test_tables1_4_partner_mismatch(tmp_path):
    rows = {"table3": _market_rows(DCD, GERMANY, len(YEARS))}
    build(tmp_path, rows=rows)
    _, problems = iv.validate_goods_inputs(str(tmp_path))
    assert any("Tables 1-4 disagree on the partner country" in p
               for p in problems)


def test_bilateral_partner_mismatch(tmp_path):
    rows = {"table6": _product_rows(KCD, KENYA, DCD, GERMANY, len(YEARS))}
    build(tmp_path, rows=rows)
    _, problems = iv.validate_goods_inputs(str(tmp_path))
    assert any("Tables 5-6 disagree on the partner" in p for p in problems)


def test_bilateral_reporter_must_be_kenya(tmp_path):
    rows = {"table5": _product_rows(FCD, FRANCE, KCD, KENYA, len(YEARS))}
    build(tmp_path, rows=rows)
    _, problems = iv.validate_goods_inputs(str(tmp_path))
    assert any("reporter reads 'France' -- expected 'Kenya'" in p
               for p in problems)


def test_bilateral_partner_must_match_tables1_4(tmp_path):
    rows = {"table6": _product_rows(KCD, KENYA, DCD, GERMANY, len(YEARS)),
            "table5": _product_rows(KCD, KENYA, DCD, GERMANY, len(YEARS))}
    build(tmp_path, rows=rows)
    _, problems = iv.validate_goods_inputs(str(tmp_path))
    assert any("mixed country downloads" in p for p in problems)


def test_disjoint_years_is_an_error(tmp_path):
    older = [2010, 2011, 2012, 2013, 2014]
    build(tmp_path, role_years={"table5": older, "table6": older})
    _, problems = iv.validate_goods_inputs(str(tmp_path))
    assert any("no common report years" in p for p in problems)
    assert any("table5=2010-2014" in p for p in problems)


def test_partial_year_overlap_warns_and_shrinks(tmp_path):
    long_span = list(range(2016, 2026))
    build(tmp_path, role_years={"table5": long_span,
                                "table6": long_span})
    checks, problems = iv.validate_goods_inputs(str(tmp_path))
    assert problems == []
    assert "warn" in levels(checks)
    assert any("report window shrinks to 2021-2025" in m for lv, m in checks
               if lv == "warn")


def test_duplicate_grand_total_row(tmp_path):
    n = len(YEARS)
    dup = [[FCD, FRANCE, "000", "World", "TOTAL", "All Products",
            *_vals(900_000, n)]]
    rows = {"table1": _market_rows(FCD, FRANCE, n) + dup}
    build(tmp_path, rows=rows)
    _, problems = iv.validate_goods_inputs(str(tmp_path))
    assert any("table1: 2 grand-total rows" in p for p in problems)


def test_missing_grand_total_row(tmp_path):
    rows = {"table2": _product_rows(FCD, FRANCE, "000", "World",
                                    len(YEARS))[1:]}
    build(tmp_path, rows=rows)
    _, problems = iv.validate_goods_inputs(str(tmp_path))
    assert any("table2: no grand-total row" in p for p in problems)


def test_non_numeric_value_cell(tmp_path):
    rows = _product_rows(FCD, FRANCE, "000", "World", len(YEARS))
    rows[1][-1] = "banana"
    build(tmp_path, rows={"table4": rows})
    _, problems = iv.validate_goods_inputs(str(tmp_path))
    assert any("table4: 1 non-numeric value cell(s)" in p
               and "'banana'" in p for p in problems)


def test_placeholder_blanks_are_not_garbage(tmp_path):
    rows = _product_rows(FCD, FRANCE, "000", "World", len(YEARS))
    rows[1][-1] = "..."
    rows[2][-2] = "-"
    build(tmp_path, rows={"table4": rows})
    checks, problems = iv.validate_goods_inputs(str(tmp_path))
    assert problems == []


def test_missing_units_warns(tmp_path):
    build(tmp_path, units={"table1": None})
    checks, problems = iv.validate_goods_inputs(str(tmp_path))
    assert problems == []
    assert any(lv == "warn" and "no unit on year headers" in m
               for lv, m in checks)


def test_unreadable_file_is_reported(tmp_path):
    build(tmp_path)
    (tmp_path / FILENAMES["table1"]).write_bytes(b"not a zip archive")
    checks, problems = iv.validate_goods_inputs(str(tmp_path))
    assert any("table1: cannot read" in p for p in problems)
    assert any(lv == "warn" and "Cross-file checks" in m
               for lv, m in checks)


def test_real_usa_sample_passes_with_year_warn():
    checks, problems = iv.validate_goods_inputs(SAMPLE)
    assert problems == []
    assert "error" not in levels(checks)
    assert any(lv == "warn" and "report window shrinks to 2021-2025" in m
               for lv, m in checks)


def test_log_line_symbols():
    assert iv.log_line("ok", "fine").startswith("\u2713 ")
    assert iv.log_line("warn", "hmm").startswith("\u26a0 ")
    assert iv.log_line("error", "bad").startswith("\u2717 ")
    assert iv.log_line("ok", "multi\n line") == "\u2713 multi line"


def test_validation_runs_in_webapp_pipeline(tmp_path):
    from webapp.main import _run_pipeline

    src = build(tmp_path / "src")
    uploads = tmp_path / "uploads"
    uploads.mkdir()
    shutil.copy(src / FILENAMES["table1"], uploads)
    logs = []
    with pytest.raises(PipelineError) as exc:
        _run_pipeline(tmp_path, None, 10, "raw", logs)
    assert "Input validation failed" in str(exc.value)
    assert any(line.startswith("\u2717") for line in logs)
