"""UN Comtrade connector.

With a free subscription key the connector downloads all partners and products
for one bilateral market and assembles the six Trade Map-style workbooks the
existing Trade Brief pipeline already understands:

    <market>s-imports-from-world-by-exporter_all.xlsx   (Table 1)
    <market>s-imports-from-world-by-product_all.xlsx    (Table 2)
    <market>s-exports-to-world-by-importer_all.xlsx     (Table 3)
    <market>s-exports-to-world-by-product_all.xlsx      (Table 4)
    kenyas-exports-to-<market>-by-product_all.xlsx      (Table 5)
    kenyas-imports-from-<market>-by-product_all.xlsx    (Table 6)

No key, no full download: the public preview endpoint caps responses at 500
records and rejects ``partnerCode=all``, so it cannot build a complete market.
"""

import hashlib
import io
import json
import time
from datetime import date
from pathlib import Path

import openpyxl

from webapp.platform import settings

from .base import (Connector, ConnectorError, Param, FetchResult, as_int,
                   http_json, slugify)
from .references import country_name, hs_label, hs_labels, resolve_country

KENYA_CODE = "404"
FLOW_EXPORT = "X"
FLOW_IMPORT = "M"
MAX_RECORDS = 100000
CACHE_TTL = 24 * 3600


class ComtradeConnector(Connector):
    id = "comtrade"
    label = "UN Comtrade"
    description = ("Official customs trade values. With a free Comtrade "
                   "subscription key this builds all six Trade Brief "
                   "workbooks for a bilateral market.")
    report_types = ("goods", "reference")
    requires_credentials = True
    params = (
        Param("partner", "Partner market", required=True,
              default="United States of America",
              help="Country name, ISO3 code or UN M49 code."),
        Param("start_year", "First year", type="number", required=True,
              default=2021),
        Param("end_year", "Last year", type="number", required=True,
              default=2025),
    )

    def validate(self, values):
        values = super().validate(values)
        if not settings.COMTRADE_API_KEY:
            raise ConnectorError(
                "UN Comtrade needs a free subscription key for full "
                "downloads. Set COMTRADE_API_KEY on the server.")
        partner = resolve_country(values["partner"], "partner")
        start = as_int(values["start_year"], 0)
        end = as_int(values["end_year"], 0)
        current = date.today().year
        if start < 1990 or end > current or start > end:
            raise ConnectorError(
                f"Choose a valid year range (1990-{current}).")
        if end - start > 15:
            raise ConnectorError("Choose a range of 15 years or fewer.")
        values["partner_info"] = partner
        values["start_year"] = start
        values["end_year"] = end
        return values

    def fetch(self, values):
        values = self.validate(values)
        builder = TradeBriefBuilder(settings.COMTRADE_API_KEY)
        files, provenance = builder.build(
            values["partner_info"], values["start_year"], values["end_year"])
        return FetchResult(
            files=files,
            provenance=provenance,
            notes=["Downloaded from UN Comtrade; values converted to USD "
                   "thousand and labelled All Products / TOTAL where needed."],
        )


class TradeBriefBuilder:
    def __init__(self, api_key):
        self.api_key = api_key
        self.cache_dir = Path(settings.CACHE_DIR) / "comtrade"
        self.cache_dir.mkdir(parents=True, exist_ok=True)
        self.queries = 0

    def _query(self, reporter, partner, cmd, flow, year):
        params = {
            "reporterCode": reporter,
            "period": str(year),
            "partnerCode": partner,
            "cmdCode": cmd,
            "flowCode": flow,
            "includeDesc": "true",
            "maxrecords": str(MAX_RECORDS),
        }
        key = hashlib.sha256(
            json.dumps(params, sort_keys=True).encode()).hexdigest()
        cache = self.cache_dir / f"{key}.json"
        try:
            if cache.exists() and \
                    time.time() - cache.stat().st_mtime < CACHE_TTL:
                return json.loads(cache.read_text(encoding="utf-8"))
        except (OSError, ValueError):
            pass
        url = "https://comtradeapi.un.org/data/v1/get/C/A/HS?" + \
            "&".join(f"{k}={v}" for k, v in params.items())
        payload = http_json(url, headers={
            "subscription-key": self.api_key,
            "Accept": "application/json",
        }, timeout=120, retries=2)
        self.queries += 1
        if payload.get("error"):
            raise ConnectorError(
                "UN Comtrade API error: " +
                json.dumps(payload["error"])[:400], status_code=502)
        try:
            cache.write_text(json.dumps(payload), encoding="utf-8")
        except OSError:
            pass
        time.sleep(0.15)
        return payload

    @staticmethod
    def _records(payload):
        data = payload.get("data") or payload.get("dataset") or []
        return [row for row in data if isinstance(row, dict)]

    def build(self, partner, start_year, end_year):
        years = list(range(start_year, end_year + 1))
        partner_code = partner["code"]
        partner_name = partner["name"]
        kenya = {"code": KENYA_CODE, "name": "Kenya"}
        labels = None
        try:
            labels = hs_labels()
        except ConnectorError:
            labels = {}

        queries = {
            "table1": (partner_code, "all", "TOTAL", FLOW_IMPORT),
            "table2": (partner_code, "0", "AG4", FLOW_IMPORT),
            "table3": (partner_code, "all", "TOTAL", FLOW_EXPORT),
            "table4": (partner_code, "0", "AG4", FLOW_EXPORT),
            "table5": (KENYA_CODE, partner_code, "AG4", FLOW_EXPORT),
            "table6": (KENYA_CODE, partner_code, "AG4", FLOW_IMPORT),
        }
        data = {}
        for name, (reporter, partner_dim, cmd, flow) in queries.items():
            data[name] = {}
            for year in years:
                payload = self._query(reporter, partner_dim, cmd, flow, year)
                data[name][year] = self._records(payload)

        slug = slugify(partner_name)
        files = {
            f"{slug}s-imports-from-world-by-exporter_all.xlsx":
                _market_workbook(data["table1"], years, partner_name,
                                 partner_code, "imports"),
            f"{slug}s-imports-from-world-by-product_all.xlsx":
                _product_workbook(data["table2"], years, partner_code,
                                  partner_name, "000", "World", labels),
            f"{slug}s-exports-to-world-by-importer_all.xlsx":
                _market_workbook(data["table3"], years, partner_name,
                                 partner_code, "exports"),
            f"{slug}s-exports-to-world-by-product_all.xlsx":
                _product_workbook(data["table4"], years, partner_code,
                                  partner_name, "000", "World", labels),
            f"kenyas-exports-to-{slug}-by-product_all.xlsx":
                _product_workbook(data["table5"], years, KENYA_CODE, "Kenya",
                                  partner_code, partner_name, labels),
            f"kenyas-imports-from-{slug}-by-product_all.xlsx":
                _product_workbook(data["table6"], years, KENYA_CODE, "Kenya",
                                  partner_code, partner_name, labels),
        }
        provenance = {
            "source": "UN Comtrade",
            "dataset": "UN Comtrade Database (HS, annual, C/A/HS)",
            "reporter": partner_name,
            "partner": partner_name,
            "years": years,
            "queries": self.queries,
            "license": "UN Comtrade terms of use",
            "retrieved_at": time.time(),
            "url": "https://comtradeapi.un.org/data/v1/get/C/A/HS",
            "mirror_note": ("Reporter-side values only. Where a reporter does "
                            "not publish a year the column is left blank."),
        }
        return files, provenance


def _num(value):
    try:
        return int(round(float(value) / 1000.0))
    except (TypeError, ValueError):
        return None


def _latest(row_years):
    for value in reversed(list(row_years.values())):
        if value is not None:
            return value
    return 0


def _market_workbook(by_year, years, reporter_name, reporter_code, direction):
    """Rows: partner market x year, product TOTAL. World row first."""
    rows = {}
    for year, records in by_year.items():
        for rec in records:
            if rec.get("partner2Code") not in (0, None, "0"):
                continue
            code = str(rec.get("partnerCode", "")).strip()
            if not code:
                continue
            try:
                code = str(int(code))
            except ValueError:
                pass
            if code == "0":
                code = "000"
            value = _num(rec.get("primaryValue"))
            name = (rec.get("partnerDesc") or "").strip() or \
                country_name(code, "World")
            entry = rows.setdefault(code, {"name": name, "values": {}})
            if value is not None:
                entry["values"][year] = value
    world = rows.get("000")
    if world is None:
        totals = {}
        for year in years:
            values = [entry["values"].get(year) for code, entry in rows.items()
                      if code != "000"]
            values = [v for v in values if v is not None]
            if values:
                totals[year] = sum(values)
        world = {"name": "World", "values": totals}
        rows["000"] = world
    ordered = [("000", world)] + sorted(
        ((code, entry) for code, entry in rows.items() if code != "000"),
        key=lambda item: _latest(item[1]["values"]), reverse=True)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Data"
    header = ["reporterCd", "reporterLabel", "partnerCd", "partnerLabel",
              "productCd", "productLabel"] + \
        [f"{year} (USD Thousand)" for year in years]
    ws.append(header)
    for code, entry in ordered:
        ws.append([reporter_code, reporter_name, code, entry["name"],
                   "TOTAL", "All Products"] +
                  [entry["values"].get(year) for year in years])
    return _xlsx_bytes(wb)


def _product_workbook(by_year, years, reporter_code, reporter_label,
                      partner_code, partner_label, labels):
    rows = {}
    for year, records in by_year.items():
        for rec in records:
            if rec.get("partner2Code") not in (0, None, "0"):
                continue
            code = str(rec.get("cmdCode", "")).strip()
            if not code:
                continue
            value = _num(rec.get("primaryValue"))
            label = (rec.get("cmdDesc") or "").strip() or \
                hs_label(code, labels)
            entry = rows.setdefault(code, {"name": label, "values": {}})
            if value is not None:
                entry["values"][year] = value
    if "TOTAL" not in rows:
        totals = {}
        for year in years:
            values = [entry["values"].get(year) for entry in rows.values()]
            values = [v for v in values if v is not None]
            if values:
                totals[year] = sum(values)
        rows["TOTAL"] = {"name": "All Products", "values": totals}
    ordered = [("TOTAL", rows["TOTAL"])] + sorted(
        ((code, entry) for code, entry in rows.items() if code != "TOTAL"),
        key=lambda item: _latest(item[1]["values"]), reverse=True)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Data"
    header = ["reporterCd", "reporterLabel", "partnerCd", "partnerLabel",
              "productCd", "productLabel"] + \
        [f"{year} (USD Thousand)" for year in years]
    ws.append(header)
    for code, entry in ordered:
        ws.append([reporter_code, reporter_label, partner_code, partner_label,
                   code, entry["name"]] +
                  [entry["values"].get(year) for year in years])
    return _xlsx_bytes(wb)


def _xlsx_bytes(workbook):
    buffer = io.BytesIO()
    workbook.save(buffer)
    return buffer.getvalue()
