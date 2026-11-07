"""Tests for the KNBS regional decomposition (``knbs_regions``).

The 2024 figures asserted here come from Economic Survey 2026 tables 6.6 and
6.7 (KSh million, divided by the 134.82 annual average rate), so a failure
means the cube changed shape or the tree in ``knbs_regions`` no longer matches
what KNBS prints.
"""

import os

import knbs_regions as R
import knbs_reference as K
import pytest

SNAPSHOT = os.path.join(os.path.dirname(K.__file__), "data", "knbs_reference.json")

pytestmark = pytest.mark.skipif(
    not os.path.exists(SNAPSHOT),
    reason="run `python knbs_reference.py refresh` to create the snapshot")


def _row(d, code):
    return next(r for r in d["rows"] if r["code"] == code)


def test_exports_by_region_match_the_published_2024_totals():
    d = R.decomposition("exports_by_destination", 2024)
    assert d["total"] == pytest.approx(8250.3, abs=0.1)
    assert _row(d, "F1")["value"] == pytest.approx(3157.1, abs=0.1)
    assert _row(d, "F1")["share"] == pytest.approx(0.3827, abs=1e-4)
    assert _row(d, "R19")["bloc"] is True
    assert _row(d, "R19")["value"] == pytest.approx(2383.5, abs=0.1)


def test_the_tree_adds_up_for_every_year():
    for year in R.years("exports_by_destination"):
        d = R.decomposition("exports_by_destination", year)
        assert d["reconciles"], (year, d["checks"])
        assert d["max_residual_pct"] < 0.01, (year, d["checks"])


def test_continents_and_their_tail_account_for_the_total():
    d = R.decomposition("exports_by_destination", 2024)
    regions = [r for r in d["rows"]
               if r["level"] == 1 and not r.get("bloc") and r["value"]]
    assert sum(r["share"] for r in regions) == pytest.approx(1.0, abs=1e-6)
    assert [r["code"] for r in regions] == \
        ["F1", "S1", "E1", "A1", "O3A", "_TOTHER"]


def test_exports_split_their_tail_further_than_imports_do():
    """KNBS books aircraft and ship stores inside exports, not imports."""
    exports = R.decomposition("exports_by_destination", 2024)
    imports = R.decomposition("imports_by_origin", 2024)
    assert _row(exports, "ACRT_SHIP")["level"] == 2
    assert all(r["code"] != "ACRT_SHIP" for r in imports["rows"])
    assert _row(imports, "ALL_OT_CNT")["level"] == 1
    assert imports["total"] == pytest.approx(20073.2, abs=0.1)
    assert _row(imports, "S1")["share"] == pytest.approx(0.6641, abs=1e-4)
    assert imports["reconciles"]


def test_sub_regions_nest_under_their_continent():
    d = R.decomposition("exports_by_destination", 2024)
    order = [r["code"] for r in d["rows"]]
    assert order.index("F97") > order.index("S1")
    assert order.index("B6") > order.index("E5")
    assert _row(d, "B6")["level"] == _row(d, "E5")["level"] + 1


def test_latest_year_is_used_and_flagged_provisional():
    d = R.decomposition("exports_by_destination")
    assert d["year"] == R.years("exports_by_destination")[-1] == 2025
    assert d["provisional"] is True


def test_blocs_can_be_left_out():
    d = R.decomposition("exports_by_destination", 2024, blocs=False)
    assert not any(r.get("bloc") for r in d["rows"])


def test_a_year_without_a_rate_raises_instead_of_inventing_one():
    assert 2018 not in R.years("exports_by_destination")
    with pytest.raises(K.KnbsError):
        R.decomposition("exports_by_destination", 2018)


def test_an_unknown_dataset_is_refused_not_guessed():
    with pytest.raises(K.KnbsError):
        R.decomposition("services_by_destination")


def test_gb1_duplicate_partner_is_never_a_row():
    d = R.decomposition("exports_by_destination", 2024)
    assert all(r["code"] != "GB1" for r in d["rows"])


# ---------------------------------------------------------------------------
# Rendering
# ---------------------------------------------------------------------------
def _render(d, caption="Table 7a (KNBS): Kenya's Exports by Region in 2024"):
    from docx import Document
    from generate_report import ReportBuilder
    cfg = {"country": {"name": "Kenya", "short": "KE", "possessive": "Kenya's"}}
    b = ReportBuilder(cfg, {})
    b.doc = Document()
    b.add_knbs_region_table(d, caption=caption)
    rows = [[c.text for c in r.cells] for r in b.doc.tables[0].rows]
    return rows, [p.text for p in b.doc.paragraphs if p.text.strip()]


def test_region_table_indents_and_labels_blocs():
    d = R.decomposition("exports_by_destination", 2024)
    rows, paras = _render(d)
    assert rows[0] == ["Region", "Value (USD million)", "Share of total"]
    assert rows[1] == ["Total", "8,250.3", "100.0%"]
    indent = " " * 4
    assert rows[2] == [indent + "Africa", "3,157.1", "38.3%"]
    bloc = next(r for r in rows if "East African Community" in r[0])
    assert bloc[0] == (indent * _row(d, "R19")["level"]
                       + "of which: East African Community")
    assert any("Table 6.13" in p for p in paras)


def test_region_table_warns_when_the_tree_stops_adding_up():
    d = R.decomposition("exports_by_destination", 2024)
    d = dict(d, reconciles=False, max_residual_pct=4.2)
    _rows, paras = _render(d)
    assert any("do not add up to the total" in p for p in paras)


def test_region_table_is_not_built_from_a_lone_row():
    from docx import Document
    from generate_report import ReportBuilder
    d = dict(R.decomposition("exports_by_destination", 2024),
             rows=[{"code": "_T", "name": "Total", "level": 0,
                    "value": 1.0, "share": 1.0}])
    cfg = {"country": {"name": "Kenya", "short": "KE", "possessive": "Kenya's"}}
    b = ReportBuilder(cfg, {})
    b.doc = Document()
    assert b.add_knbs_region_table(d) is None
    assert b.doc.tables == []


def test_region_bullets_read_the_split_without_double_counting_blocs():
    from generate_report import knbs_region_bullets
    d = R.decomposition("exports_by_destination", 2024)
    previous = R.decomposition("exports_by_destination", 2023)
    bullets = " ".join(knbs_region_bullets(d, previous))
    assert "Africa first" in bullets
    assert "By sub-region" in bullets
    assert "Against 2023" in bullets
    assert "East African Community" not in bullets   # a bloc, not a region


# --------------------------------------------------------------------------
# the partners KNBS names separately (what a bilateral check can use)
# --------------------------------------------------------------------------
def test_partners_are_the_countries_knbs_names_one_by_one():
    partners = R.partners("exports_by_destination")
    assert partners["US"] == "USA"
    assert partners["GB"] == "United Kingdom"
    assert len(partners) < 68          # the cube holds far more rows than names


def test_partners_exclude_regions_blocs_and_second_lines():
    partners = R.partners("exports_by_destination")
    for code in ("_T", "F1", "E5", "F97", "B6", "R19", "ACRT_SHIP",
                 "ALL_OT_CNT", "OTHER_F1", "GB1", "CD1", "SO1"):
        assert code not in partners, code


def test_a_partner_can_appear_on_one_leg_only():
    # KNBS names Egypt as a destination but not as an origin.
    assert "EG" in R.partners("exports_by_destination")
    assert "EG" not in R.partners("imports_by_origin")


def test_region_rows_ignore_the_second_lines_so_uganda_is_counted_once():
    # R19 is exactly the EAC members KNBS names, taken from the canonical code.
    values = R._year_values("exports_by_destination", 2024)
    members = ["UG", "TZ", "RW", "BI", "SS", "SO", "CD"]
    assert sum(values[c] for c in members) == pytest.approx(values["R19"], abs=0.1)
