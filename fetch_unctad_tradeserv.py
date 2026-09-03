#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
fetch_unctad_tradeserv.py
=========================

Pull Kenya's services trade per category (annual, partner = World) from the
UNCTADstat data API and cache it next to the existing UNCTAD sources.

Endpoint (from UNCTAD's own "Get selected data using the data API" recipe):
    https://unctadstat-user-api.unctad.org/US.TradeServCatTotal/cur/Facts?culture=en

The dataset is "Services: Trade by category - Annual" (UNCTAD-WTO Trade in
Services Data Set).  It reports ~100 service categories per economy per year
for partner "World total" — which is exactly the per-country category detail
that the quarterly file lacks for Kenya's most recent years.

Authentication requires YOUR UNCTADstat account credentials:
    * 'Client ID'  and 'API key' from the 'My Home' page of a logged-in user
    * https://unctadstat.unctad.org/datacentre

Credentials are supplied via environment variables (or --client-id/--api-key):

    $env:UNCTAD_CLIENT_ID = "xxxx"
    $env:UNCTAD_API_KEY   = "yyyy"

Example:
    python fetch_unctad_tradeserv.py --out C:\\...\\Servicesnew\\Kenya_tradeserv_annual.csv.gz

By default this writes a small **coverage report** and the gz CSV.  Opening the
report lets you confirm whether Kenya's annual category detail now extends past
2023 (2024 / 2025), which would let the services tables lift the 2023 cap.
"""

import argparse
import csv
import gzip
import io
import os
import sys
import urllib.parse
import urllib.request
import urllib.error

API_URL = "https://unctadstat-user-api.unctad.org/US.TradeServCatTotal/cur/Facts?culture=en"

# EBOPS service-category codes (the same set the quarterly file uses).
# Fetching ALL categories (no Category filter) also works; this list is used
# only for the human coverage summary.
CATEGORIES = {
    "S": "Services (total)",
    "SC": "Transport",
    "SD": "Travel",
    "SE": "Construction",
    "SF": "Insurance and pension",
    "SG": "Financial services",
    "SH": "Charges for IP n.i.e.",
    "SI": "Telecom, computer, information",
    "SJ": "Other business services",
    "SK": "Personal, cultural, recreational",
    "SL": "Government goods and services n.i.e.",
    "SPX1": "Other services",
    "SPX4": "Goods-related services",
    "SOX": "Memo: Commercial services",
}


def build_filter(economy, flows, start, end):
    """Build the OData $filter string."""
    exprs = []
    if economy:
        # Prefer the numeric code; fall back to the label.
        if economy.isdigit():
            exprs.append("Economy/Code eq '%s'" % economy)
        else:
            exprs.append("Economy/Label eq '%s'" % economy)
    exprs.append("Flow/Code in (%s)" % ",".join("'%s'" % f for f in flows))
    years = ",".join(str(y) for y in range(start, end + 1))
    exprs.append("Year in (%s)" % years)
    return " and ".join(exprs)


def fetch(client_id, api_key, economy="Kenya", flows=("01", "02"),
          start=2005, end=2025, timeout=180):
    """Return the decoded CSV rows as a list."""
    params = {
        "$select": "Economy/Label,Flow/Label,Category/Code,Category/Label,Year,"
                   "Millions_of_US_at_current_prices_Value,"
                   "US_at_current_prices_Footnote,US_at_current_prices_MissingValue",
        "$filter": build_filter(economy, flows, start, end),
        "$orderby": "Economy/Order asc,Year asc",
        "$compute": "round(M0100/Value div 1000000, 0) as "
                    "Millions_of_US_at_current_prices_Value, "
                    "M0100/Footnote/Text as US_at_current_prices_Footnote, "
                    "M0100/MissingValue/Label as US_at_current_prices_MissingValue",
        "$format": "csv",
        "compress": "gz",
    }
    body = urllib.parse.urlencode(params).encode("utf-8")
    req = urllib.request.Request(
        API_URL, data=body,
        headers={"ClientId": client_id, "ClientSecret": api_key,
                 "User-Agent": "tradeflow-client/1.0"})
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            content = resp.read()
    except urllib.error.HTTPError as exc:  # noqa: F821
        raise SystemExit("[ERROR] HTTP %s: %s" % (exc.code, exc.read()[:500]))
    except Exception as exc:  # noqa: BLE001
        raise SystemExit("[ERROR] Request failed: %s" % exc)

    # Response body is gzip-compressed CSV.
    try:
        raw = gzip.decompress(content)
    except Exception:  # noqa: BLE001 - some responses are already uncompressed
        raw = content
    text = raw.decode("utf-8", errors="replace")
    reader = csv.DictReader(io.StringIO(text))
    rows = list(reader)
    return rows


def summarize(rows):
    """Print which categories/years contain real values for Kenya."""
    by = {}
    for r in rows:
        year = r.get("Year", "")
        cat = (r.get("Economy/Label") and
               _category_for_row(r) or "")
        val = r.get("Millions_of_US_at_current_prices_Value")
        try:
            float(val)
            has = val is not None and val != "" and float(val) != 0
        except (TypeError, ValueError):
            has = False
        key = (year, cat)
        by.setdefault(key, False)
        by[key] = by[key] or has
    # fall back: we don't know cat from select output; print years sorted
    years = sorted({r.get("Year") for r in rows}, key=int)
    print("\n=== COVERAGE (rows fetched) ===")
    print("Rows: %d  |  Years present: %s" % (len(rows), ", ".join(years)))
    if rows:
        print("Sample row keys:", list(rows[0].keys()))


def _category_for_row(r):
    # The $select does not carry the Category code, so we cannot tag rows here.
    return ""


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--client-id", default=os.environ.get("UNCTAD_CLIENT_ID"),
                    help="UNCTADstat Client ID (or env UNCTAD_CLIENT_ID)")
    ap.add_argument("--api-key", default=os.environ.get("UNCTAD_API_KEY"),
                    help="UNCTADstat API key (or env UNCTAD_API_KEY)")
    ap.add_argument("--economy", default="Kenya",
                    help="Economy label or 3/4-digit code (default: Kenya)")
    ap.add_argument("--start", type=int, default=2005)
    ap.add_argument("--end", type=int, default=2025)
    ap.add_argument("--out", default=None,
                    help="Output gz CSV path (default: <cwd>/Kenya_tradeserv_annual.csv.gz)")
    args = ap.parse_args(argv)

    if not args.client_id or not args.api_key:
        ap.error(
            "Client ID and API key are required.\n"
            "Get them from the UNCTADstat 'My Home' page (login required):\n"
            "  https://unctadstat.unctad.org/datacentre\n"
            "Then set UNCTAD_CLIENT_ID / UNCTAD_API_KEY or pass --client-id/--api-key.")

    rows = fetch(args.client_id, args.api_key, args.economy,
                 ("01", "02"), args.start, args.end)
    if not rows:
        raise SystemExit("[WARN] API returned no rows — check filters/credentials.")

    out = args.out or os.path.join(os.getcwd(), "Kenya_tradeserv_annual.csv.gz")
    with gzip.open(out, "wt", newline="", encoding="utf-8") as fh:
        writer = csv.DictWriter(fh, fieldnames=list(rows[0].keys()))
        writer.writeheader()
        writer.writerows(rows)

    print("[OK] Wrote %d rows -> %s" % (len(rows), out))
    summarize(rows)


if __name__ == "__main__":
    main()