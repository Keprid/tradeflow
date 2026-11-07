"""UNCTADstat services connector.

Uses the credentialed OData endpoint already documented in
``fetch_unctad_tradeserv.py`` and assembles the six Trade Map-style service
workbooks the Services Brief pipeline expects, so a UNCTAD-only workflow can
build the report without manual Trade Map exports.
"""

import io
import sys
import time
from datetime import date
from pathlib import Path

import openpyxl

from webapp.platform import settings

sys.path.insert(0, str(Path(__file__).resolve().parents[2]))

from fetch_unctad_tradeserv import CATEGORIES, fetch as fetch_unctad  # noqa: E402

from .base import Connector, ConnectorError, FetchResult, Param, as_float, \
    as_int  # noqa: E402

ITC_SERVICES = (
    ("1", "Manufacturing services on physical inputs owned by others"),
    ("2", "Maintenance and repair services n.i.e."),
    ("3", "Transport"),
    ("4", "Travel"),
    ("5", "Construction"),
    ("6", "Insurance and pension services"),
    ("7", "Financial services"),
    ("8", "Charges for the use of intellectual property n.i.e."),
    ("9", "Telecommunications, computer, and information services"),
    ("10", "Other business services"),
    ("11", "Personal, cultural, and recreational services"),
    ("12", "Government goods and services n.i.e."),
)
UNCTAD_TO_ITC = {
    "SA": "1", "SB": "2", "SC": "3", "SD": "4", "SE": "5", "SF": "6",
    "SG": "7", "SH": "8", "SI": "9", "SJ": "10", "SK": "11", "SL": "12",
}
ITC_LABEL = dict(ITC_SERVICES)
CATEGORY_CODES = ("S",) + tuple(UNCTAD_TO_ITC)


class UnctadConnector(Connector):
    id = "unctad"
    label = "UNCTADstat (services)"
    description = ("UNCTAD-WTO trade in services by EBOPS category. Builds "
                   "the six Services Brief source workbooks for Kenya and "
                   "the world.")
    report_types = ("services",)
    requires_credentials = True
    params = (
        Param("start_year", "First year", type="number", required=True,
              default=2020),
        Param("end_year", "Last year", type="number", required=True,
              default=2024),
    )

    def validate(self, values):
        values = super().validate(values)
        if not settings.UNCTAD_CLIENT_ID or not settings.UNCTAD_API_KEY:
            raise ConnectorError(
                "UNCTADstat needs account credentials. Set UNCTAD_CLIENT_ID "
                "and UNCTAD_API_KEY on the server (free from "
                "unctadstat.unctad.org/datacentre).")
        start = as_int(values["start_year"], 0)
        end = as_int(values["end_year"], 0)
        current = date.today().year
        if start < 2005 or end > current or start > end:
            raise ConnectorError(f"Choose a valid year range (2005-{current}).")
        if end - start > 15:
            raise ConnectorError("Choose a range of 15 years or fewer.")
        values["start_year"] = start
        values["end_year"] = end
        return values

    def fetch(self, values):
        values = self.validate(values)
        start, end = values["start_year"], values["end_year"]
        try:
            rows = fetch_unctad(
                settings.UNCTAD_CLIENT_ID, settings.UNCTAD_API_KEY,
                economy=None, flows=("01", "02"), start=start, end=end,
                timeout=300, categories=CATEGORY_CODES)
        except SystemExit as exc:
            raise ConnectorError(str(exc), status_code=502) from exc
        except Exception as exc:  # noqa: BLE001
            raise ConnectorError(
                f"UNCTADstat request failed: {exc}", status_code=502) from exc
        parsed = _parse(rows)
        files = _workbooks(parsed, start, end)
        provenance = {
            "source": "UNCTADstat",
            "dataset": "Services: Trade by category - Annual "
                       "(UNCTAD-WTO Trade in Services)",
            "years": list(range(start, end + 1)),
            "rows": len(rows),
            "license": "UNCTADstat terms of use",
            "retrieved_at": time.time(),
            "url": "https://unctadstat-user-api.unctad.org/"
                   "US.TradeServCatTotal/cur/Facts",
        }
        return FetchResult(
            files=files, provenance=provenance,
            notes=["Values converted from UNCTAD USD million to USD "
                   "thousand for the Trade Map-style headers."])


def _parse(rows):
    world = {"Exports": {}, "Imports": {}}
    kenya = {"Exports": {}, "Imports": {}}
    countries = {"Exports": {}, "Imports": {}}
    for row in rows:
        flow = (row.get("Flow_Label") or "").strip().lower()
        if flow not in ("exports", "imports"):
            continue
        flow = flow.capitalize()
        category = (row.get("Category_Code") or "").strip()
        if category not in CATEGORY_CODES:
            continue
        missing = (row.get("US_at_current_prices_MissingValue") or "").strip()
        value = as_float(row.get("Millions_of_US_at_current_prices_Value"),
                         None)
        if missing or value is None:
            continue
        year = as_int(row.get("Year"), 0)
        if not year:
            continue
        value *= 1000.0  # USD million -> USD thousand
        eco_code = str(row.get("Economy_Code") or "").strip()
        eco_label = (row.get("Economy_Label") or "").strip()
        is_world = eco_code == "0000" or eco_label.lower() == "world"
        is_country = eco_code.isdigit() and len(eco_code) == 3
        if is_world:
            world[flow].setdefault(category, {})[year] = value
            continue
        if is_country:
            countries[flow].setdefault(category, {}).setdefault(
                eco_label, {})[year] = value
        if eco_code == "404" or eco_label.lower() == "kenya":
            kenya[flow].setdefault(category, {})[year] = value

    # Fill missing world category totals by summing countries.
    for flow in ("Exports", "Imports"):
        for category in CATEGORY_CODES:
            if category in world[flow]:
                continue
            totals = {}
            for name, series in countries[flow].get(category, {}).items():
                for year, value in series.items():
                    totals[year] = totals.get(year, 0.0) + value
            if totals:
                world[flow][category] = totals
        # Country ranking uses the total-services (S) series.
        ranking = countries[flow].get("S", {})
        if not ranking:
            ranking = {}
            for category, names in countries[flow].items():
                if category == "S":
                    continue
                for name, series in names.items():
                    target = ranking.setdefault(name, {})
                    for year, value in series.items():
                        target[year] = target.get(year, 0.0) + value
        countries[flow]["S"] = ranking
    return {"world": world, "kenya": kenya, "countries": countries}


def _years_from(start, end):
    return list(range(start, end + 1))


def _category_rows(flow, data, years, kenya=False):
    label_key = "Imported" if flow == "Imports" else "Exported"
    rows = []
    total = data.get("S", {})
    rows.append(["S", "All services"] +
                [round(total.get(year, 0.0)) if total.get(year) is not None
                 else None for year in years])
    for code, label in ITC_SERVICES:
        un_code = next((u for u, itc in UNCTAD_TO_ITC.items()
                        if itc == code), None)
        series = data.get(un_code, {}) if un_code else {}
        if not series:
            continue
        rows.append([code, label] +
                    [round(series.get(year, 0.0))
                     if series.get(year) is not None else None
                     for year in years])
    rows[1:] = sorted(
        rows[1:],
        key=lambda row: next((v for v in reversed(row[2:]) if v is not None),
                             0), reverse=True)
    header = ["Service code", "Service label"] + \
        [f"{label_key} value in {year}" for year in years]
    return [header] + rows


def _ranking_rows(flow, ranking, years):
    label = "Importer" if flow == "Imports" else "Exporter"
    rows = []
    for name, series in ranking.items():
        rows.append([name] +
                    [round(series.get(year, 0.0))
                     if series.get(year) is not None else None
                     for year in years])
    rows.sort(key=lambda row: next(
        (v for v in reversed(row[1:]) if v is not None), 0), reverse=True)
    header = [label + "s"] + [
        f"{'Imported' if flow == 'Imports' else 'Exported'} value in {year}"
        for year in years]
    return [header] + rows


def _kenya_rows(flow, data, years):
    label_key = "Imported" if flow == "Imports" else "Exported"
    rows = []
    total = data.get("S", {})
    rows.append(["S", "All services"] +
                [round(total.get(year, 0.0)) if total.get(year) is not None
                 else None for year in years])
    for code, label in ITC_SERVICES:
        un_code = next((u for u, itc in UNCTAD_TO_ITC.items()
                        if itc == code), None)
        series = data.get(un_code, {}) if un_code else {}
        if not series:
            continue
        rows.append([code, label] +
                    [round(series.get(year, 0.0))
                     if series.get(year) is not None else None
                     for year in years])
    rows[1:] = sorted(
        rows[1:],
        key=lambda row: next((v for v in reversed(row[2:]) if v is not None),
                             0), reverse=True)
    header = ["Code", "Service label"] + \
        [f"{label_key} Value in {year}" for year in years]
    return [header] + rows


def _workbook(rows):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Data"
    for row in rows:
        ws.append(row)
    buffer = io.BytesIO()
    wb.save(buffer)
    return buffer.getvalue()


def _workbooks(parsed, start, end):
    years = _years_from(start, end)
    world = parsed["world"]
    kenya = parsed["kenya"]
    countries = parsed["countries"]
    return {
        "Trade_Map_-_List_of_exported_services_for_the_selected_service_"
        "(All_services).xlsx":
            _workbook(_category_rows("Exports", world["Exports"], years)),
        "Trade_Map_-_List_of_imported_services_for_the_selected_service_"
        "(All_services).xlsx":
            _workbook(_category_rows("Imports", world["Imports"], years)),
        "Trade_Map_-_List_of_exporters_for_the_selected_service_"
        "(All_services).xlsx":
            _workbook(_ranking_rows("Exports", countries["Exports"]["S"],
                                    years)),
        "Trade_Map_-_List_of_importers_for_the_selected_service_"
        "(All_services).xlsx":
            _workbook(_ranking_rows("Imports", countries["Imports"]["S"],
                                    years)),
        "Trade_Map_-_List_of_services_exported_by_Kenya_"
        "(All_services).xlsx":
            _workbook(_kenya_rows("Exports", kenya["Exports"], years)),
        "Trade_Map_-_List_of_services_imported_by_Kenya_"
        "(All_services).xlsx":
            _workbook(_kenya_rows("Imports", kenya["Imports"], years)),
    }
