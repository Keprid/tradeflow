#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
fetch_unctad_tradeserv.py
=========================

Pull annual services-trade-by-category data from the UNCTADstat data API and
cache it as a gzipped CSV next to the other service sources.

Endpoint (from UNCTAD's own "Get selected data using the data API" recipe):
    https://unctadstat-user-api.unctad.org/US.TradeServCatTotal/cur/Facts?culture=en

The dataset is "Services: Trade by category - Annual" (UNCTAD-WTO Trade in
Services Data Set).  It reports ~100 service categories per economy per year
for partner "World total" — which is exactly the per-country category detail
that the quarterly bulk file lacks for the most recent years.

By default this fetches **one economy** (Kenya) — enough for the Kenya category
tables (3/4).  With ``--all-economies`` it fetches **every economy** and becomes
a single canonical annual source for:
  * economy-level total services (``S``) → Tables 1/2 (global exporters/importers)
    and Table 13 (Kenya vs peers), together with the true published **World**
    total (which the quarterly bulk never carried for recent years);
  * per-category world totals (pie chart, RCA/concentration/diversification);
  * Kenya's per-category detail (Tables 3/4).

Authentication requires YOUR UNCTADstat account credentials:
    * 'Client ID'  and 'API key' from the 'My Home' page of a logged-in user
    * https://unctadstat.unctad.org/datacentre

Credentials are supplied via environment variables (or --client-id/--api-key):

    $env:UNCTAD_CLIENT_ID = "xxxx"
    $env:UNCTAD_API_KEY   = "yyyy"

Examples:
    # Kenya-only (Tables 3/4 category detail)
    python fetch_unctad_tradeserv.py --out C:\\...\\services\\UNCTAD_Kenya_tradeserv_annual.csv.gz

    # All economies (canonical annual source replacing the quarterly bulk file)
    python fetch_unctad_tradeserv.py --all-economies \
        --categories S,SA,SB,SC,SD,SE,SF,SG,SH,SI,SJ,SK,SL \
        --out C:\\...\\services\\UNCTAD_tradeserv_annual_all.csv.gz
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


def build_filter(economy, flows, start, end, categories=None):
    """Build the OData $filter string.

    ``economy`` may be ``None`` (fetch **all** economies) or a label / numeric
    code.  ``flows`` is a tuple of ``Flow/Code`` values.  ``categories`` is an
    optional tuple of ``Category/Code`` values; when omitted the endpoint's
    full category set is returned (thousands of rows per economy).
    """
    exprs = []
    if economy:
        # Prefer the numeric code; fall back to the label.
        if economy.isdigit():
            exprs.append("Economy/Code eq '%s'" % economy)
        else:
            exprs.append("Economy/Label eq '%s'" % economy)
    exprs.append("Flow/Code in (%s)" % ",".join("'%s'" % f for f in flows))
    if categories:
        exprs.append("Category/Code in (%s)"
                     % ",".join("'%s'" % c for c in categories))
    years = ",".join(str(y) for y in range(start, end + 1))
    exprs.append("Year in (%s)" % years)
    return " and ".join(exprs)


def _fetch_once(client_id, api_key, filt, orderby, timeout, select=None,
                compute=None):
    """Issue one OData query and return the decoded CSV rows as a list.

    Raises ``SystemExit`` on HTTP errors (the raw body is echoed) and ``IOError``
    on other network failures.
    """
    params = {
        "$select": select or (
            "Economy/Label,Economy/Code,Flow/Label,Category/Code,Category/Label,Year,"
            "Millions_of_US_at_current_prices_Value,"
            "US_at_current_prices_Footnote,US_at_current_prices_MissingValue"),
        "$filter": filt,
        "$orderby": orderby,
        "$compute": compute or (
            "round(M0100/Value div 1000000, 0) as "
            "Millions_of_US_at_current_prices_Value, "
            "M0100/Footnote/Text as US_at_current_prices_Footnote, "
            "M0100/MissingValue/Label as US_at_current_prices_MissingValue"),
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
        raise IOError("[ERROR] Request failed: %s" % exc)

    # Response body is gzip-compressed CSV.
    try:
        raw = gzip.decompress(content)
    except Exception:  # noqa: BLE001 - some responses are already uncompressed
        raw = content
    text = raw.decode("utf-8", errors="replace")
    reader = csv.DictReader(io.StringIO(text))
    return list(reader)


def fetch(client_id, api_key, economy="Kenya", flows=("01", "02"),
          start=2005, end=2025, timeout=300, categories=None):
    """Return the decoded CSV rows as a list.

    When ``economy`` is ``None`` (the *all economies* case) the response for the
    full year span may exceed the OData endpoint's size cap, so the request is
    split **one year at a time** and the per-year pages are concatenated.  A
    single economy/category combo is returned in one shot.
    """
    orderby = "Economy/Order asc,Year asc"
    if economy is None:
        rows = []
        for y in range(start, end + 1):
            filt = build_filter(None, flows, y, y, categories)
            rows.extend(_fetch_once(client_id, api_key, filt, orderby, timeout))
        return rows
    filt = build_filter(economy, flows, start, end, categories)
    return _fetch_once(client_id, api_key, filt, orderby, timeout)


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
                    help="Economy label or 3/4-digit code (default: Kenya")
    ap.add_argument("--all-economies", action="store_true",
                    help="Fetch every economy (overrides --economy). Use for the "
                         "global by-category file that drives country rankings, "
                         "world category totals and Kenya's category detail.")
    ap.add_argument("--categories", default=None,
                    help="Comma/space separated Category/Code list to fetch "
                         "(e.g. 'S,SC,SD').  Omit to fetch the full category set.")
    ap.add_argument("--start", type=int, default=2005)
    ap.add_argument("--end", type=int, default=2025)
    ap.add_argument("--out", default=None,
                    help="Output gz CSV path (default: <cwd>/<name>_tradeserv_annual.csv.gz)")
    args = ap.parse_args(argv)

    if not args.client_id or not args.api_key:
        ap.error(
            "Client ID and API key are required.\n"
            "Get them from the UNCTADstat 'My Home' page (login required):\n"
            "  https://unctadstat.unctad.org/datacentre\n"
            "Then set UNCTAD_CLIENT_ID / UNCTAD_API_KEY or pass --client-id/--api-key.")

    economy = None if args.all_economies else args.economy
    categories = tuple(c.strip() for c in (args.categories or "").split(",")
                       if c.strip()) or None
    rows = fetch(args.client_id, args.api_key, economy,
                 ("01", "02"), args.start, args.end, categories=categories)
    if not rows:
        raise SystemExit("[WARN] API returned no rows — check filters/credentials.")

    if args.out:
        out = args.out
    else:
        scope = "all" if args.all_economies else (economy or "")
        out = os.path.join(os.getcwd(), "%s_tradeserv_annual.csv.gz" % scope)
    with gzip.open(out, "wt", newline="", encoding="utf-8") as fh:
        writer = csv.DictWriter(fh, fieldnames=list(rows[0].keys()))
        writer.writeheader()
        writer.writerows(rows)

    print("[OK] Wrote %d rows -> %s" % (len(rows), out))
    summarize(rows)


if __name__ == "__main__":
    main()