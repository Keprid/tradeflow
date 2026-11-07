"""Tests for the KNBS reference layer.

The numbers asserted here are the ones printed in Economic Survey 2026, so a
failure means either the portal changed its shape or a unit scale in
``knbs_reference.DATASETS`` was edited without checking the published table.
"""

import json
import os

import knbs_check
import knbs_reference as K
import pytest

SNAPSHOT = os.path.join(os.path.dirname(K.__file__), "data", "knbs_reference.json")

pytestmark = pytest.mark.skipif(
    not os.path.exists(SNAPSHOT),
    reason="run `python knbs_reference.py refresh` to create the snapshot")


def test_exchange_rates_match_survey_table_6_19():
    assert K.rate(2021) == 109.65
    assert K.rate(2024) == 134.82
    assert K.rate(2025) == 129.30
    assert K.is_provisional(2025) and not K.is_provisional(2024)


def test_missing_rate_raises_rather_than_guessing():
    with pytest.raises(K.KnbsError):
        K.rate(1999)


def test_unit_scale_registry_is_applied():
    # ES 2026 Table 6.7: KSh million.
    kes = K.series("principal_domestic_exports", COMMODITY="OW_CT_FLWR",
                   UNIT_MEASURE="KES")
    assert kes[2024] == pytest.approx(96_979.7e6, rel=1e-6)
    # ES 2026 Table 7.14: the cube labels the column KES but prints KSh
    # billion, and volume as TONNES when the table is in '000 tonnes.
    value = K.series("fresh_horticulture", ITEM="CP0931", AGR_INDIC="VALUE",
                     UNIT_MEASURE="KES")
    assert value[2024] == pytest.approx(72.1e9, rel=1e-6)
    volume = K.series("fresh_horticulture", ITEM="CP0931", AGR_INDIC="VOLUME",
                      UNIT_MEASURE="TONNES")
    assert volume[2024] == pytest.approx(102.5e3, rel=1e-6)
    # ES 2026 Table 12.7: '000 bed-nights.
    nights = K.series("hotel_bed_nights", COUNTRY_RES="EUR", FREQ="A",
                      UNIT_MEASURE="NUMBER")
    assert nights[2024] == pytest.approx(2363.0e3, rel=1e-3)


def test_empty_observations_stay_absent():
    kes = K.series("principal_domestic_exports", COMMODITY="OW_CT_FLWR",
                   UNIT_MEASURE="KES")
    assert 2018 not in kes and 2019 not in kes
    assert min(kes) == 2020


def test_balance_of_payments_services_line():
    credit = K.series("bop_annual", BOP_INDIC="S", ACC_ENTRY="C",
                      UNIT_MEASURE="KES", UNIT_MULT="6")
    assert credit[2024] == pytest.approx(1_084_289.6e6, rel=1e-6)
    assert set(credit) >= {2021, 2022, 2023, 2024, 2025}


def test_merchandise_trade_totals_are_internally_consistent():
    exports = K.series("merchandise_trade", TRADE_FL="EXPORT",
                       INTD_INDIC="BOMT", UNIT_MEASURE="KES")
    by_destination = K.series("exports_by_destination", TRADE_PARTNER="_T",
                              UNIT_MEASURE="KES")
    assert exports[2025] == pytest.approx(by_destination[2025], rel=1e-9)


def test_services_extension_covers_the_year_unctad_does_not():
    services = knbs_check.services_2025()
    assert max(services["exports"]) == 2025
    assert services["exports"][2025] == pytest.approx(8394.0, rel=1e-3)
    assert services["balance"][2025] > 0


def test_product_check_agrees_with_published_itc_series():
    ours = {2021: 725.5, 2022: 628.6, 2023: 667.5, 2024: 722.9, 2025: 799.0}
    result = knbs_check.check_product(ours, hs_code="0603")
    assert result["status"] == "consistent"
    assert result["max_abs_diff_pct"] < 1.0
    assert [r["year"] for r in result["rows"]] == [2021, 2022, 2023, 2024, 2025]
    assert result["rows"][-1]["provisional"] is True


def test_check_without_a_mapping_is_unavailable_not_wrong():
    result = knbs_check.check_product({2024: 100.0}, hs_code="0401")
    assert result["status"] == "unavailable"
    assert result["rows"] == []


def test_check_flags_a_divergent_series():
    result = knbs_check.check_services_exports({2024: 5000.0})
    assert result["status"] == "review"
    assert result["max_abs_diff_pct"] > knbs_check.REVIEW_THRESHOLD_PCT


def test_snapshot_records_its_own_provenance():
    with open(SNAPSHOT, encoding="utf-8") as fh:
        snap = json.load(fh)
    assert snap["generated_at"]
    assert snap["api_root"].startswith("https://opendata.knbs.or.ke")
    for name, entry in snap["datasets"].items():
        assert entry["flow"], name
        assert entry["survey_table"], name
        assert entry["rows"], name


def test_services_extension_stops_at_the_years_unctad_lacks():
    ext = knbs_check.services_extension(2024)
    assert ext["years"] == [2025]
    assert ext["provisional"] == [True]
    assert ext["exports"] == [pytest.approx(8394.0, rel=1e-3)]
    assert ext["balance"] == [pytest.approx(2505.1, rel=1e-3)]
    assert "ES 2026" in ext["source"]
    # nothing is left to add once UNCTAD catches up
    assert knbs_check.services_extension(ext["years"][-1]) is None


def _render_checks(checks):
    from docx import Document
    from generate_report import ReportBuilder
    cfg = {"country": {"name": "Kenya", "short": "KE", "possessive": "Kenya's"}}
    b = ReportBuilder(cfg, {})
    b.doc = Document()
    b.add_knbs_check_table(checks, caption="KNBS Cross-Check")
    rows = [[c.text for c in r.cells] for r in b.doc.tables[0].rows] \
        if b.doc.tables else []
    return rows, [p.text for p in b.doc.paragraphs if p.text.strip()]


def test_check_table_prints_the_gap_and_asks_for_a_review():
    check = {"label": "KNBS services credit vs UNCTAD exports",
             "rows": [{"year": 2024, "ours": 5531.1, "knbs": 8042.5,
                       "diff_pct": 45.4, "provisional": False}],
             "max_abs_diff_pct": 45.4, "status": "review",
             "source": "KNBS Open Data Portal, Balance of Payments, Annual",
             "review_threshold_pct": 2.0,
             "rate_source": "Economic Survey 2026 Table 6.19"}
    rows, paras = _render_checks([check])
    assert rows[0] == ["Series", "Year", "This report (USD m)",
                       "KNBS (USD m)", "Difference"]
    assert rows[1] == [check["label"], "2024", "5,531.1", "8,042.5", "+45.40%"]
    assert any("above the review tolerance" in p for p in paras)
    assert any("Economic Survey 2026 Table 6.19" in p for p in paras)


def test_check_table_marks_provisional_years():
    check = {"label": "KNBS services credit vs UNCTAD exports",
             "rows": [{"year": 2025, "ours": 8394.0, "knbs": 8394.0,
                       "diff_pct": 0.0, "provisional": True}],
             "max_abs_diff_pct": 0.0, "status": "consistent",
             "source": "KNBS Open Data Portal", "review_threshold_pct": 2.0,
             "rate_source": "Economic Survey 2026 Table 6.19"}
    rows, paras = _render_checks([check])
    assert rows[1][3] == "8,394.0*"
    assert any("provisional" in p for p in paras)
    assert not any("above the review tolerance" in p for p in paras)


def test_check_table_names_series_it_cannot_check():
    missing = {"label": "KNBS cross-check", "rows": [], "status": "unavailable",
               "max_abs_diff_pct": None, "source": "no KNBS commodity mapping",
               "review_threshold_pct": 2.0,
               "rate_source": "Economic Survey 2026 Table 6.19"}
    rows, paras = _render_checks([missing])
    assert rows == []                                   # no table, no zeros
    assert any("KNBS cross-check" in p for p in paras)



# --------------------------------------------------------------------------
# bilateral checks: one partner's line, not the world total
# --------------------------------------------------------------------------
def test_partner_code_follows_knbs_names_not_ours():
    code = knbs_check.partner_code
    assert code("United States of America") == "US"      # KNBS prints 'USA'
    assert code("Iran Islamic Republic of") == "IR"
    assert code("South Korea") == "KR"
    assert code("United Rep. of Tanzania") == "TZ"
    assert code("Dem. Rep. Congo") == "CD"


def test_partner_code_is_none_when_knbs_groups_the_market():
    assert knbs_check.partner_code("Nigeria") is None
    assert knbs_check.partner_code("Serbia") is None


def test_partner_check_reconciles_the_usa_export_line():
    check = knbs_check.check_partner("United States of America", {2024: 662.5})
    assert check["label"] == "KNBS exports (USA) vs our USD figures"
    assert check["status"] == "consistent"
    row = check["rows"][0]
    assert row["knbs"] == pytest.approx(659.1, abs=0.1)
    assert row["diff_pct"] == pytest.approx(-0.51, abs=0.02)
    assert "Table 6.13" in check["source"]


def test_partner_check_reads_the_origin_cube_for_imports():
    check = knbs_check.check_partner("United States of America", {2024: 1163.2},
                                     direction="imports")
    assert check["rows"][0]["knbs"] == pytest.approx(1154.4, abs=0.1)
    assert "Table 6.14" in check["source"]


def test_partner_checks_cover_only_the_legs_knbs_publishes():
    both = knbs_check.partner_checks("United States of America",
                                     {2024: 662.5}, {2024: 1163.2})
    assert [c["label"] for c in both] == [
        "KNBS exports (USA) vs our USD figures",
        "KNBS imports (USA) vs our USD figures"]
    # Egypt is a named destination but not a named origin.
    assert len(knbs_check.partner_checks("Egypt", {2024: 1.0}, {2024: 1.0})) == 1
    assert knbs_check.partner_checks("Nigeria", {2024: 1.0}, {2024: 1.0}) == []


def _bilateral_analysis(country, exports, imports, years):
    from generate_report import Analysis
    cfg = {"country": {"name": country, "short": country,
                       "possessive": country + "'s"},
           "report": {"year": years[-1]}}
    a = Analysis(cfg)
    a.balance = {"years": years, "exports": exports, "imports": imports,
                 "balance": [None] * len(years)}
    a.knbs_checks = a._load_knbs_checks()
    return a


def test_goods_brief_holds_the_bilateral_checks_for_its_own_partner():
    a = _bilateral_analysis("United States of America",
                            [662.5, 616.2], [1163.2, 1056.0], [2024, 2025])
    assert [c["label"] for c in a.knbs_checks] == [
        "KNBS exports (USA) vs our USD figures",
        "KNBS imports (USA) vs our USD figures"]
    assert [r["year"] for r in a.knbs_checks[0]["rows"]] == [2024, 2025]


def test_goods_brief_carries_no_check_for_a_market_knbs_does_not_name():
    assert _bilateral_analysis("Nigeria", [1.0], [1.0], [2024]).knbs_checks == []


def test_bilateral_check_table_is_lettered_against_figure_one():
    from docx import Document
    from generate_report import ReportBuilder, add_knbs_bilateral_check
    cfg = {"country": {"name": "United States of America", "short": "USA",
                       "possessive": "the USA's"}}
    b = ReportBuilder(cfg, {})
    b.doc = Document()
    a = _bilateral_analysis("United States of America",
                            [662.5], [1163.2], [2024])
    assert add_knbs_bilateral_check(b, a) is True
    captions = [p.text for p in b.doc.paragraphs if "Table 4a" in p.text]
    assert any("KNBS Customs Data" in c for c in captions)
    body = [[c.text for c in row.cells] for row in b.doc.tables[0].rows]
    assert body[1][:2] == ["KNBS exports (USA) vs our USD figures", "2024"]
    assert body[1][3] == "659.1"


def test_bilateral_check_table_is_skipped_when_there_is_nothing_to_check():
    from docx import Document
    from generate_report import ReportBuilder, add_knbs_bilateral_check
    cfg = {"country": {"name": "Nigeria", "short": "Nigeria",
                       "possessive": "Nigeria's"}}
    b = ReportBuilder(cfg, {})
    b.doc = Document()
    a = _bilateral_analysis("Nigeria", [1.0], [1.0], [2024])
    assert add_knbs_bilateral_check(b, a) is False
    assert b.doc.tables == []


# --------------------------------------------------------------------------
# product profile check: one heading, one KNBS commodity line
# --------------------------------------------------------------------------
COFFEE = os.path.join(os.path.dirname(K.__file__),
                      "Samplefiles", "Product_Profiles", "Coffee")


class _NoKenyaSeries:
    def market_share_series(self):
        return None


def _profile_builder(done=()):
    from docx import Document
    from generate_report import ReportBuilder
    cfg = {"country": {"name": "Kenya", "short": "KE", "possessive": "Kenya's"}}
    b = ReportBuilder(cfg, {})
    b.doc = Document()
    b.done_tables = set(done)
    return b


def test_product_check_needs_the_single_heading_knbs_publishes():
    from generate_product_profile import add_knbs_product_check
    b = _profile_builder([12])
    for codes in (["06"], ["0603", "0604"], []):
        cfg = {"include_codes": codes, "family_title": "Coffee"}
        assert add_knbs_product_check(b, cfg, _NoKenyaSeries()) is False
    # a four-digit heading, but the profile has no Kenya series to compare
    assert add_knbs_product_check(
        b, {"include_codes": ["0901"], "family_title": "Coffee"},
        _NoKenyaSeries()) is False
    assert b.doc.tables == []


def test_product_check_sits_with_the_table_it_is_lettered_against():
    from generate_product_profile import add_knbs_product_check
    b = _profile_builder()                 # Table 12 was never built
    assert add_knbs_product_check(
        b, {"include_codes": ["0901"], "family_title": "Coffee"},
        _NoKenyaSeries()) is False


@pytest.mark.skipif(not os.path.isdir(COFFEE), reason="coffee sample not bundled")
def test_product_check_prints_the_commodity_line_against_kenya_total():
    from generate_product_profile import ProfileData, add_knbs_product_check
    data = ProfileData(COFFEE, ["0901"], "Coffee")
    b = _profile_builder([12])
    cfg = {"include_codes": ["0901"], "family_title": "Coffee"}
    assert add_knbs_product_check(b, cfg, data) is True
    captions = [p.text for p in b.doc.paragraphs if "Table 12a" in p.text]
    assert any("Coffee Export Cross-Check" in c for c in captions)
    body = [[c.text for c in row.cells] for row in b.doc.tables[0].rows]
    row = next(r for r in body if r[1] == "2024")
    assert row[0] == "KNBS 'Coffee, unroasted' vs our USD figures"
    assert row[2] == "297.8"               # ITC's Kenya coffee exports
    assert row[3] == "285.0"               # KNBS's domestic exports
    assert row[4].startswith("-")          # the gap is re-exported coffee


class _TotalExportsProfile:
    """A coffee portfolio that also carries the 'products exported by Kenya'
    workbook, i.e. the one profile file that reports total merchandise trade."""

    kenya_total_exports = {2024: 8_196_000.0}      # USD thousands

    def market_share_series(self):
        return [(2024, 297_800.0, 2_300_000.0, 0.129)]


def test_product_check_adds_kenya_total_merchandise_when_the_file_is_there():
    from generate_product_profile import add_knbs_product_check
    b = _profile_builder([12])
    cfg = {"include_codes": ["0901"], "family_title": "Coffee"}
    assert add_knbs_product_check(b, cfg, _TotalExportsProfile()) is True
    body = [[c.text for c in row.cells] for row in b.doc.tables[0].rows]
    assert {r[0] for r in body[1:]} == {
        "KNBS 'Coffee, unroasted' vs our USD figures",
        "KNBS merchandise exports vs our USD figures"}
    total = next(r for r in body[1:] if r[0].startswith("KNBS merchandise"))
    assert total[2] == "8,196.0"
    assert total[3] == "8,250.3"        # Survey table 6.1, KSh at the 2024 rate
