"""Competitor analysis: per-product and per-importing-market layouts."""

from pathlib import Path

import openpyxl
import pytest

import generate_competitor_analysis as gca

ROOT = Path(__file__).resolve().parent.parent
CA_FILES = ROOT / "Samplefiles" / "Competitor Analysis" / "CA files"
COFFEE = ROOT / "Samplefiles" / "Product_Profiles" / "Coffee"


def test_detects_per_market_layout_from_ca_files(tmp_path):
    result = gca.generate(str(CA_FILES), str(tmp_path), top=5)

    assert result["mode"] == "country"
    assert result["product"]["code"] == "0603"
    assert len(result["markets"]) == 10
    assert [m["country"] for m in result["markets"]][:3] == [
        "Netherlands", "United Kingdom", "Germany"]
    assert result["markets"][0]["rank"] == 1
    assert result["markets"][0]["world"] == pytest.approx(1_270_866)
    assert result["markets"][0]["kenya_value"] == pytest.approx(397_884)
    assert all((m["kenya_value"] or 0) > 0 for m in result["markets"])

    germany = next(m for m in result["markets"] if m["country"] == "Germany")
    assert germany["rank"] == 3
    assert germany["world"] > result["markets"][0]["world"]

    for key in ("png", "docx", "xlsx"):
        assert Path(result[key]).exists()
        assert Path(result[key]).stat().st_size > 0


def test_per_market_workbook_matches_sample_columns(tmp_path):
    result = gca.generate(str(CA_FILES), str(tmp_path), top=5)
    wb = openpyxl.load_workbook(result["xlsx"], data_only=True)
    ws = wb["Competitor Analysis"]
    headers = [ws.cell(1, c).value for c in range(1, 8)]
    assert headers == [
        "Rank in 2025", "Importers", "Value imported in 2025",
        "Rank in 2025", "Top Exporters", "Value exported in 2025",
        "Share in 2025 %"]
    assert ws.cell(2, 2).value == "Netherlands"
    assert ws.cell(2, 3).value == pytest.approx(1_270_866)
    assert ws.cell(2, 3).number_format == "#,##0"
    assert ws.cell(2, 6).value == pytest.approx(397_884)
    assert ws.cell(2, 7).number_format == "0.0%"
    wb.close()


def test_per_market_analysis_guarantees_kenya(tmp_path):
    result = gca.generate(str(CA_FILES), str(tmp_path), top=5)
    for market in result["markets"]:
        assert any(r["kenya"] for r in market["exporters"]), market["country"]


def test_coffee_single_market_uses_per_product_layout(tmp_path):
    result = gca.generate(str(COFFEE), str(tmp_path), top=5)

    assert result["mode"] == "product"
    assert result["product"]["code"] == "0901"
    assert len(result["products"]) == 1

    wb = openpyxl.load_workbook(result["xlsx"], data_only=True)
    ws = wb["Competitor Analysis"]
    headers = [ws.cell(1, c).value for c in range(1, 10)]
    assert headers[1] == "Product code"
    assert headers[8] == "Tariff Applied %"
    assert isinstance(ws.cell(2, 4).value, (int, float))
    assert ws.cell(2, 4).number_format == "#,##0"
    assert ws.cell(2, 8).number_format == "0.0%"
    wb.close()
