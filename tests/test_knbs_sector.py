"""Tests for the KNBS sectoral-context layer.

Every figure asserted here is the one Economic Survey 2026 prints in the table
named by the view, so a failure means the portal changed a series or the Survey
renamed a dimension - not that a tolerance moved.
"""

import os

import knbs_reference as K
import knbs_sector as S
import pytest

SNAPSHOT = os.path.join(os.path.dirname(K.__file__), "data", "knbs_reference.json")

pytestmark = pytest.mark.skipif(
    not os.path.exists(SNAPSHOT),
    reason="run `python knbs_reference.py refresh` to create the snapshot")


def _value(ctx, *parts):
    row = next(r for r in ctx["rows"]
               if all(p.lower() in r["label"].lower() for p in parts))
    return row


def test_the_window_is_the_five_years_the_survey_prints():
    ctx = S.context(["0901"])
    assert ctx["years"] == [2021, 2022, 2023, 2024, 2025]
    assert ctx["provisional"] == [2025]


def test_coffee_rows_match_survey_table_7_10_and_7_4():
    ctx = S.context(["0901"])
    v = _value(ctx, "area under coffee")
    assert (v["unit"], v["values"][2024]) == ("Hectares", pytest.approx(113_500))
    v = _value(ctx, "produced, all growers")
    assert v["values"][2021] == pytest.approx(34_500)      # crop year 2020/21
    assert v["values"][2025] == pytest.approx(51_400)
    assert _value(ctx, "co-operatives", "produced")["values"][2025] == \
        pytest.approx(36_800)
    assert _value(ctx, "estates", "produced")["values"][2025] == \
        pytest.approx(14_600)
    # the portal keeps the full quotient; the Survey prints one decimal
    assert _value(ctx, "yield, co-operatives")["values"][2025] == \
        pytest.approx(439.1, abs=0.05)
    price = _value(ctx, "auction price")
    assert (price["unit"], price["values"][2025]) == ("KSh per 100 Kg",
                                                      pytest.approx(93_614.6))
    assert ctx["sources"] == ["ES 2026 Table 7.10", "ES 2026 Table 7.4"]


def test_tea_separates_green_leaf_from_made_tea():
    ctx = S.context(["0902"])
    assert _value(ctx, "green leaf")["values"][2024] == \
        pytest.approx(2_687_200)
    assert _value(ctx, "made tea produced, all")["values"][2024] == \
        pytest.approx(598_500)
    assert _value(ctx, "made tea produced by smallholders")["values"][2025] == \
        pytest.approx(271_770)
    # the portal's green-leaf rows before 2020 hold made-tea weights
    assert min(_value(ctx, "green leaf")["values"]) >= 2021


def test_cut_flower_value_starts_where_the_portal_can_be_trusted():
    ctx = S.context(["0603"])
    volume = _value(ctx, "cut flowers exported")
    value = _value(ctx, "cut flowers exported, value")
    assert volume["values"][2021] == pytest.approx(210_100)
    assert 2021 not in value["values"]                     # mis-filed upstream
    assert value["unit"] == "KSh billion"
    assert value["values"][2024] == pytest.approx(72.1)
    assert _value(ctx, "all fresh horticulture produce exported", "value")[
        "values"][2021] == pytest.approx(157.7)
    assert any("mis-filed" in n for n in ctx["notes"])


def test_horticulture_family_reads_the_basket_it_sells():
    ctx = S.context(["06", "07", "08"])
    assert [r["label"] for r in ctx["rows"]] == [
        "Cut flowers exported", "Fruits exported",
        "All fresh horticulture produce",
        "All fresh horticulture produce, value"]
    assert _value(ctx, "fruits")["values"][2024] == pytest.approx(225_400)


def test_mining_reads_only_what_the_cube_weighs_in_tonnes():
    ctx = S.context(["25", "26"])
    labels = " ".join(r["label"] for r in ctx["rows"])
    # gold is in kilograms and gemstones in thousand carats, both under a
    # tonne label, so they stay out of the table
    assert "Gold" not in labels and "emstone" not in labels
    assert _value(ctx, "value of all minerals")["values"][2024] == \
        pytest.approx(25_529.3)
    assert _value(ctx, "soda ash produced")["values"][2025] == \
        pytest.approx(289_610.8)
    assert _value(ctx, "titanium ore minerals produced")["values"][2025] == \
        pytest.approx(101_000.5)


def test_meat_falls_back_to_the_price_the_farmer_received():
    ctx = S.context(["02", "16"])
    assert _value(ctx, "beef")["values"][2025] == pytest.approx(67_097.0)
    assert _value(ctx, "pig")["values"][2025] == pytest.approx(26_596.9,
                                                               abs=0.05)
    assert ctx["sources"] == ["ES 2026 Table 7.4"]


def test_products_without_a_knbs_industry_print_nothing():
    for codes in (["04"],                                   # dairy
                  ["4420.10", "6214.10-6214.90", "6307.90"],  # crafts
                  ["09"],                                    # chapter, no view
                  []):
        assert S.context(codes) is None, codes


def test_a_heading_is_matched_before_its_chapter():
    assert S.view_for(["0603"]) == "0603"
    assert S.view_for(["06", "07", "08"]) == "06"
    assert S.view_for(["06", "0603"]) == "0603"
    assert S.view_for(["26"]) == "26"
    assert S.VIEWS["26"] is S.VIEWS["25"]


def test_years_no_row_can_fill_are_dropped():
    # mining and coffee both reach 2025, so the block always carries it;
    # a view whose rows stop early loses the empty column rather than printing
    # a row of dashes.
    ctx = S.context(["0901"], columns=3)
    assert ctx["years"] == [2023, 2024, 2025]
    assert all(r["values"] for r in ctx["rows"])


# --------------------------------------------------------------------------
# rendering
# --------------------------------------------------------------------------
def _render(codes):
    from docx import Document
    from generate_report import ReportBuilder
    cfg = {"country": {"name": "Kenya", "short": "KE", "possessive": "Kenya's"}}
    b = ReportBuilder(cfg, {})
    b.doc = Document()
    ctx = S.context(codes)
    table = b.add_knbs_sector_table(ctx, caption="Table 12a (KNBS): Coffee Context")
    body = [[c.text for c in r.cells] for r in table.rows]
    paras = [p.text for p in b.doc.paragraphs if p.text.strip()]
    return body, paras


def test_sector_table_carries_a_unit_column_and_the_provisional_year():
    body, paras = _render(["0901"])
    assert body[0] == ["Indicator", "Unit", "2021", "2022", "2023", "2024",
                       "2025*"]
    row = next(r for r in body[1:] if r[0].startswith("Area under coffee"))
    assert row[1] == "Hectares"
    assert row[2:] == ["108,200", "109,400", "111,900", "113,500", "113,900"]
    price = next(r for r in body[1:] if "auction price" in r[0])
    assert price[2:] == ["65,864.9", "48,871.4", "61,416.2", "65,881.0",
                         "93,614.6"]
    assert any("ES 2026 Table 7.10" in p and "ES 2026 Table 7.4" in p
               for p in paras)
    assert any("2024/25 season" in p for p in paras)
    assert any("* KNBS publishes this year as provisional." in p for p in paras)


def test_sector_table_leaves_a_blank_instead_of_a_wrong_number():
    body, paras = _render(["0603"])
    value = next(r for r in body[1:] if r[0].endswith("value")
                 and r[0].startswith("Cut flowers"))
    assert value[2] == "n/a"
    assert value[-1] == "81.3"
    assert any("mis-filed" in p for p in paras)


def test_profile_letters_the_block_after_the_cross_check():
    from generate_product_profile import add_knbs_sector_context
    b = _builder([12])
    cfg = {"include_codes": ["0901"], "family_title": "Coffee"}
    assert add_knbs_sector_context(b, cfg, "b") is True
    assert any(p.startswith("Table 12b (KNBS): Kenya's Coffee")
               for p in [q.text for q in b.doc.paragraphs if q.text.strip()])
    # without a cross-check the same block is the first letter
    b = _builder([12])
    assert add_knbs_sector_context(b, cfg, "a") is True
    assert any(p.startswith("Table 12a (KNBS):")
               for p in [q.text for q in b.doc.paragraphs if q.text.strip()])


def test_profile_prints_nothing_without_the_table_it_leaves_a_gap_for():
    from generate_product_profile import add_knbs_sector_context
    b = _builder([])                       # no Table 12 in this profile
    cfg = {"include_codes": ["0901"], "family_title": "Coffee"}
    assert add_knbs_sector_context(b, cfg) is False
    assert b.doc.tables == []
    b = _builder([12])
    assert add_knbs_sector_context(b, {"include_codes": ["04"]}) is False
    assert b.doc.tables == []


def _builder(done):
    from docx import Document
    from generate_report import ReportBuilder
    cfg = {"country": {"name": "Kenya", "short": "KE", "possessive": "Kenya's"}}
    b = ReportBuilder(cfg, {})
    b.doc = Document()
    b.done_tables = set(done)
    return b
