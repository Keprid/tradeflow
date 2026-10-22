"""Tests for the shared spreadsheet sniffing / conversion helpers."""

from pathlib import Path

import openpyxl
import pytest

from xlsx_compat import (
    HTMLTableParser,
    _read_text_any,
    _safe_sheet_name,
    _sniff_spreadsheet_kind,
    _write_csv_as_xlsx,
    _write_html_as_xlsx,
    convert_to_xlsx,
    is_spreadsheet,
)


def test_sniff_html_in_disguise(tmp_path):
    p = tmp_path / "fake.xls"
    p.write_text("<html><table><tr><td>a</td></tr></table></html>",
                 encoding="utf-8")
    assert _sniff_spreadsheet_kind(p) == "html"


def test_sniff_csv(tmp_path):
    p = tmp_path / "data.xls"
    p.write_text("a,b,c\n1,2,3\n", encoding="utf-8")
    assert _sniff_spreadsheet_kind(p) == "csv"


def test_sniff_real_binary_xls(tmp_path):
    p = tmp_path / "bin.xls"
    p.write_bytes(b"\x09\x00\x05\x00\x00\x00" + b"\x00" * 64)
    assert _sniff_spreadsheet_kind(p) == "xls"


def test_safe_sheet_name():
    assert _safe_sheet_name("bad:name/with*chars", "X") == "bad_name_with_chars"
    assert _safe_sheet_name("", "Sheet1") == "Sheet1"
    assert len(_safe_sheet_name("x" * 60, "X")) == 31


def test_read_text_any_encodings(tmp_path):
    p = tmp_path / "u.txt"
    p.write_bytes("héllo".encode("utf-8-sig"))
    assert "héllo" in _read_text_any(p)


def test_html_parser_keeps_nbsp_and_tables():
    parser = HTMLTableParser()
    parser.feed("<table><tr><td>a&nbsp;b</td><td>2</td></tr></table>"
                "<table><tr><td>1</td></tr><tr><td>2</td></tr></table>")
    # largest table wins
    assert len(parser.tables) == 2
    assert parser.tables[1] == [["1"], ["2"]]
    assert parser.tables[0][0][0] == "a b"


def test_html_writer_keeps_product_group_label(tmp_path):
    src = tmp_path / "craft.xls"
    src.write_text(
        "<html><body>Product group: Wood craft items"
        "<table><tr><td>Code</td><td>Value</td></tr>"
        "<tr><td>4421</td><td>12</td></tr></table></body></html>",
        encoding="utf-8")
    out = tmp_path / "craft.xlsx"
    _write_html_as_xlsx(src, out)
    ws = openpyxl.load_workbook(str(out)).active
    assert ws["A1"].value == "Code"
    assert ws["B2"].value == "12"
    # the crafts loader reads the family from T1
    assert ws.cell(row=1, column=20).value == "Product group: Wood craft items"


def test_csv_writer(tmp_path):
    src = tmp_path / "t.csv"
    src.write_text("a;b\n1;2\n", encoding="utf-8")
    out = tmp_path / "t.xlsx"
    _write_csv_as_xlsx(src, out)
    ws = openpyxl.load_workbook(str(out)).active
    assert ws["A2"].value == "1"
    assert ws["B2"].value == "2"


def test_convert_xlsx_passthrough(tmp_path):
    p = tmp_path / "a.xlsx"
    wb = openpyxl.Workbook()
    wb.save(str(p))
    used, path = convert_to_xlsx(p)
    used.close()
    assert Path(path) == p
    assert is_spreadsheet("foo.XLSB")
    assert not is_spreadsheet("foo.docx")
