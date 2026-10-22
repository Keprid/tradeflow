"""Unit tests for the shared xlsx style / formula-cache engine."""

import openpyxl
import pytest

from xlsx_tools import (
    _finalize,
    _inject_cached_values,
    _missing_cached_refs,
    _put_formula,
    _put_formula_sheet,
    _save_workbook,
    _sheet_order,
    get_row_fill,
    put,
    put_text,
    put_val,
    style_cell,
    to_float,
)


def _wb_with_formulas():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Report"
    put_val(ws, 1, 1, 10.0)
    put_val(ws, 1, 2, 32.0)
    cache = []
    _put_formula(ws, cache, 1, 3, "=SUM(A1:B1)", 42.0)
    return wb, ws, cache


def test_to_float_variants():
    assert to_float("1,234.5") == 1234.5
    assert to_float("...") is None
    assert to_float("..") is None
    assert to_float("") is None
    assert to_float(None) is None
    assert to_float(3) == 3.0
    assert to_float(float("nan")) is None


def test_put_helpers_write_values(tmp_path):
    wb, ws, cache = _wb_with_formulas()
    assert ws["C1"].value == "=SUM(A1:B1)"
    assert cache == [("C1", 42.0)]
    assert get_row_fill(0) is not None
    assert get_row_fill(1) is None  # odd rows are white


def test_formula_cache_roundtrip(tmp_path):
    wb, ws, cache = _wb_with_formulas()
    path = tmp_path / "book.xlsx"
    _save_workbook(wb, str(path), cache)

    # data_only readers must see the cached number, not None
    wb2 = openpyxl.load_workbook(str(path), data_only=True)
    assert wb2["Report"]["C1"].value == 42.0
    wb2.close()

    # the live formula must still be there for Excel
    wb3 = openpyxl.load_workbook(str(path))
    assert wb3["Report"]["C1"].value == "=SUM(A1:B1)"
    wb3.close()


def test_inject_is_idempotent_and_validates(tmp_path):
    wb, ws, cache = _wb_with_formulas()
    path = tmp_path / "book.xlsx"
    wb.save(str(path))
    _inject_cached_values(str(path), {"Report": {"C1": 42.0}})
    assert _missing_cached_refs(str(path), {"Report": {"C1": 42.0}}) == []
    assert _missing_cached_refs(str(path), {"Report": {"C9": 1.0}}) == [
        ("Report", "C9")]


def test_inject_raises_for_unknown_sheet(tmp_path):
    wb, ws, cache = _wb_with_formulas()
    path = tmp_path / "book.xlsx"
    wb.save(str(path))
    # a dict keyed by a sheet that does not exist must fail loudly
    # (so _save_workbook's fallback kicks in instead of a silent None)
    with pytest.raises(RuntimeError):
        _inject_cached_values(path, {"Nope": {"C1": 42.0}})


def test_sheet_order_matches_creation_order(tmp_path):
    wb = openpyxl.Workbook()
    wb.active.title = "Alpha"
    wb.create_sheet("Beta")
    path = tmp_path / "order.xlsx"
    wb.save(str(path))
    assert _sheet_order(str(path)) == ["Alpha", "Beta"]


def test_finalize_dict_form(tmp_path):
    wb, ws, cache = _wb_with_formulas()
    path = tmp_path / "fin.xlsx"
    _finalize(ws, str(path), cache)
    wb2 = openpyxl.load_workbook(str(path), data_only=True)
    assert wb2["Report"]["C1"].value == 42.0
    wb2.close()


def test_put_formula_sheet_dict_cache():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "S1"
    cache = {}
    _put_formula_sheet(ws, cache, 2, 2, "=1+1", 2.0, numfmt="0.00")
    assert cache == {"S1": {"B2": 2.0}}


def test_style_cell_border_styles():
    wb = openpyxl.Workbook()
    c = wb.active["A1"]
    style_cell(c, border_style="none")
    style_cell(c, border_style="header")
    style_cell(c, border_style="total")
    style_cell(c, border_style="data")
    put_text(wb.active, 3, 1, "hello")
    assert wb.active["A3"].value == "hello"
