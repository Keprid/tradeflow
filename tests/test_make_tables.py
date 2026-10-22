"""Tests for shared table-building helpers and PipelineError behaviour."""

import os

import pytest

import make_tables
from country_names import display_name, short_product_name, title_case
from pipeline_errors import PipelineError


def test_truncate_label_keeps_whole_clauses():
    label = "Live sheep; live goats; live horses; live asses; live mules"
    out = make_tables.truncate_label(label, limit=40)
    # keeps every clause that fits, never splits or trailing-punctuates
    assert out == "Live sheep; live goats; live horses"
    assert not out.endswith(";")


def test_truncate_label_empty():
    assert make_tables.truncate_label("") == ""
    assert make_tables.truncate_label(None) == ""


def test_find_source_files_missing_raises_pipeline_error(tmp_path):
    with pytest.raises(PipelineError) as ei:
        make_tables.find_source_files(str(tmp_path))
    assert "Missing required source files" in str(ei.value)


def test_shared_helpers_reexported_from_xlsx_tools():
    import xlsx_tools
    assert make_tables._save_workbook is xlsx_tools._save_workbook
    assert make_tables.to_float is xlsx_tools.to_float
    assert make_tables.put is xlsx_tools.put


def test_country_names_basics():
    assert display_name("united states of america") == "USA"
    # already-uppercase words pass through unchanged (acronym protection)
    assert title_case("SAUDI ARABIA") == "SAUDI ARABIA"
    assert title_case("imports from the world") == "Imports from the World"
    assert short_product_name(code="0902") == "Tea"


def test_cli_guard_wired():
    # The __main__ guard converts PipelineError into a clean sys.exit.
    src = open(os.path.join(os.path.dirname(make_tables.__file__),
                            "make_tables.py")).read()
    assert "except PipelineError as e:" in src
