"""World Bank WITS connector (open SDMX API, no key).

WITS publishes reporter/partner totals and broad product groups (29 sectors)
from its TradeStats dataset.  It is ideal as an independent reference or a
quick country profile, but it does not expose the HS4 product detail the full
Trade Brief needs - UN Comtrade is the source for that.
"""

import csv
import io
import time
from datetime import date

import openpyxl

from .base import (Connector, ConnectorError, FetchResult, Param, as_float,
                   as_int, http_json, slugify)
from .references import resolve_country

BASE = ("https://wits.worldbank.org/API/V1/SDMX/V21/rest/data/"
        "DF_WITS_TradeStats_Trade/")
INDICATORS = {"exports": "XPRT-TRD-VL", "imports": "MPRT-TRD-VL"}


class WitsConnector(Connector):
    id = "wits"
    label = "World Bank WITS"
    description = ("Open World Bank trade statistics: totals by partner and "
                   "broad product groups. No key required; useful as a "
                   "reference dataset or quick validation.")
    report_types = ("reference",)
    params = (
        Param("reporter", "Reporter economy", required=True,
              default="Kenya"),
        Param("start_year", "First year", type="number", required=True,
              default=2020),
        Param("end_year", "Last year", type="number", required=True,
              default=2024),
    )

    def validate(self, values):
        values = super().validate(values)
        reporter = resolve_country(values["reporter"], "reporter")
        if not reporter.get("iso3"):
            raise ConnectorError(
                f"No ISO3 code is available for {reporter['name']}.")
        start = as_int(values["start_year"], 0)
        end = as_int(values["end_year"], 0)
        current = date.today().year
        if start < 1990 or end > current or start > end:
            raise ConnectorError(f"Choose a valid year range (1990-{current}).")
        if end - start > 20:
            raise ConnectorError("Choose a range of 20 years or fewer.")
        values["reporter_info"] = reporter
        values["start_year"] = start
        values["end_year"] = end
        return values

    def fetch(self, values):
        values = self.validate(values)
        reporter = values["reporter_info"]
        start, end = values["start_year"], values["end_year"]
        queries = {
            "exports_partner": (f"A.{reporter['iso3']}..Total."
                               f"{INDICATORS['exports']}", "Exports",
                               "Partner"),
            "imports_partner": (f"A.{reporter['iso3']}..Total."
                               f"{INDICATORS['imports']}", "Imports",
                               "Partner"),
            "exports_product": (f"A.{reporter['iso3']}.WLD.."
                               f"{INDICATORS['exports']}", "Exports",
                               "Product"),
            "imports_product": (f"A.{reporter['iso3']}.WLD.."
                               f"{INDICATORS['imports']}", "Imports",
                               "Product"),
        }
        records = []
        urls = []
        for name, (key, flow, dimension) in queries.items():
            url = f"{BASE}{key}?startPeriod={start}&endPeriod={end}&format=JSON"
            urls.append(url)
            payload = http_json(url, headers={
                "User-Agent": "Mozilla/5.0 (compatible; tradeflow/1.0)",
                "Accept": "application/json",
            }, timeout=120, retries=2)
            records.extend(_parse_sdmx(payload, flow, dimension))

        files = _build_files(records, reporter["name"], start, end)
        available_years = sorted({rec["year"] for rec in records})
        notes = [
            "WITS values are already in USD thousand.",
            "Product rows use WITS sector codes, not HS4.",
            "Reference data only: these files cannot be uploaded as a Trade "
            "Brief source set.",
        ]
        if available_years:
            notes.append(
                f"WITS returned observations for {available_years[0]}-"
                f"{available_years[-1]}; requested years without published "
                f"data are blank.")
        provenance = {
            "source": "World Bank WITS",
            "dataset": "WITS TradeStats Trade (DF_WITS_TradeStats_Trade)",
            "reporter": reporter["name"],
            "years": list(range(start, end + 1)),
            "available_years": available_years,
            "records": len(records),
            "urls": urls,
            "license": "World Bank Open Data terms",
            "retrieved_at": time.time(),
            "note": ("WITS product groups are broad sectors (29), not HS4; "
                     "use UN Comtrade for the full Trade Brief."),
        }
        return FetchResult(files=files, provenance=provenance, notes=notes)


def _parse_sdmx(payload, flow, dimension):
    datasets = payload.get("dataSets") or []
    if not datasets:
        return []
    structure = payload.get("structure") or {}
    series_dims = structure.get("dimensions", {}).get("series", [])
    obs_dims = structure.get("dimensions", {}).get("observation", [])
    if not series_dims or not obs_dims:
        return []
    times = obs_dims[0].get("values", [])
    out = []
    series = datasets[0].get("series") or {}
    for key, body in series.items():
        indexes = [int(part) for part in str(key).split(":")]
        dims = {}
        for dim, index in zip(series_dims, indexes):
            values = dim.get("values", [])
            if 0 <= index < len(values):
                dims[dim.get("id")] = values[index]
        for obs_index, observation in (body.get("observations") or {}).items():
            idx = int(obs_index)
            if idx >= len(times):
                continue
            value = as_float(observation[0] if observation else None, None)
            if value is None:
                continue
            code_dim = dims.get("PARTNER" if dimension == "Partner"
                                else "PRODUCTCODE", {})
            out.append({
                "flow": flow,
                "dimension": dimension,
                "code": code_dim.get("id", ""),
                "name": code_dim.get("name") or code_dim.get("id", ""),
                "year": as_int(times[idx].get("id"), 0),
                "value": value,
            })
    return [row for row in out if row["year"]]


def _build_files(records, reporter_name, start, end):
    years = list(range(start, end + 1))
    slug = slugify(reporter_name)
    files = {}
    for flow in ("exports", "imports"):
        wb = openpyxl.Workbook()
        wb.remove(wb.active)
        for dimension in ("Partner", "Product"):
            ws = wb.create_sheet(dimension)
            ws.append(["Partner code" if dimension == "Partner"
                       else "Product code",
                       "Partner" if dimension == "Partner" else "Product"] +
                      [f"{year} (USD Thousand)" for year in years])
            rows = {}
            for rec in records:
                if rec["flow"] != flow.capitalize() or \
                        rec["dimension"] != dimension:
                    continue
                entry = rows.setdefault(
                    rec["code"], {"name": rec["name"], "values": {}})
                entry["values"][rec["year"]] = entry["values"].get(
                    rec["year"], 0.0) + rec["value"]
            ordered = sorted(rows.items(), key=lambda item: next(
                (v for v in reversed([item[1]["values"].get(y) for y in years])
                 if v is not None), 0), reverse=True)
            for code, entry in ordered:
                ws.append([code, entry["name"]] +
                          [round(entry["values"].get(year))
                           if entry["values"].get(year) is not None else None
                           for year in years])
        buffer = io.BytesIO()
        wb.save(buffer)
        files[f"WITS_TradeStats_{slug}_{flow}_{start}-{end}.xlsx"] = \
            buffer.getvalue()

    stream = io.StringIO()
    writer = csv.writer(stream)
    writer.writerow(["reporter", "flow", "dimension", "code", "name",
                     "year", "value_usd_thousand"])
    for rec in sorted(records, key=lambda r: (r["flow"], r["dimension"],
                                              r["year"], r["code"])):
        writer.writerow([reporter_name, rec["flow"], rec["dimension"],
                         rec["code"], rec["name"], rec["year"],
                         round(rec["value"])])
    files[f"WITS_TradeStats_{slug}_long_{start}-{end}.csv"] = \
        stream.getvalue().encode("utf-8")

    available_years = sorted({rec["year"] for rec in records})
    available = (f"{available_years[0]}-{available_years[-1]}"
                 if available_years else "no observations")
    readme = f"""WORLD BANK WITS - REFERENCE DATA ONLY
=====================================

Reporter: {reporter_name}
Requested years: {start}-{end}
WITS observations: {available}
Retrieved: {time.strftime('%Y-%m-%d')}
Dataset: WITS TradeStats Trade (DF_WITS_TradeStats_Trade)
Source: https://wits.worldbank.org/
License: World Bank Open Data terms

Files in this download
----------------------
* WITS_TradeStats_{slug}_exports_{start}-{end}.xlsx
  Kenya exports with sheets 'Partner' and 'Product' (USD thousand).
* WITS_TradeStats_{slug}_imports_{start}-{end}.xlsx
  Kenya imports with sheets 'Partner' and 'Product' (USD thousand).
* WITS_TradeStats_{slug}_long_{start}-{end}.csv
  One observation per row: reporter, flow, dimension, code, year, value.

IMPORTANT
---------
These are not ITC Trade Map downloads and are not a Trade Brief source set.
WITS product rows use broad WITS sector groups, not HS4 product codes, and the
data describes Kenya's own trade rather than the partner market's rankings.
Uploading this bundle to the report generator will be rejected.

For a full goods Trade Brief use UN Comtrade (COMTRADE_API_KEY) or manually
downloaded ITC Trade Map workbooks. For services use UNCTADstat. Keep these
WITS files for reference, validation or a quick country profile.

Data availability
-----------------
WITS often lags the current year by one or two years. Columns for requested
years with no published observation are blank.
"""
    files[f"WITS_TradeStats_{slug}_README_{start}-{end}.txt"] = \
        readme.encode("utf-8")
    return files
