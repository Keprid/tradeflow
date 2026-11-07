"""KNBS reference series: portal client, unit registry, exchange rates.

The Kenya National Bureau of Statistics publishes the machine-readable
equivalent of Economic Survey chapter 6 (external sector) and chapter 7-13
(sectoral tables) through the Hub-and-Node API behind
``opendata.knbs.or.ke/beta/databrowser``.  Responses are JSON-stat 2.0 cubes.

Two properties of that portal shape everything here:

* **The portal drops the scale from its unit labels.**  ``UNIT_MEASURE``
  says ``TONNES`` for a column published in '000 tonnes and ``KES`` for a
  column published in KSh billion.  Every dataset therefore carries an
  explicit multiplier in :data:`DATASETS`, verified against the printed
  Economic Survey table, and the label is never trusted.
* **Values are sparse.**  Missing years come back as ``null`` (for example
  the cut-flower line starts in 2020).  They stay ``None`` here; nothing is
  zero-filled.

Reports read the vendored snapshot in ``data/knbs_reference.json`` so a
build never depends on the portal being up.  Refresh it with
``python knbs_reference.py refresh``, or from the web app's **Or fetch open
data** card (KNBS reference series -> Install snapshot); the snapshot records
when it was taken and the vintage of every series.
"""

import json
import os
import sys
import time
import urllib.request

API_ROOT = "https://opendata.knbs.or.ke/beta/databrowserhub/api"
NODE = "1"

HERE = os.path.dirname(os.path.abspath(__file__))
SNAPSHOT_PATH = os.path.join(HERE, "data", "knbs_reference.json")

# Annual average KSh per USD, Economic Survey Table 6.19 (source: Central
# Bank of Kenya).  2025 is marked provisional by KNBS.
EXCHANGE_RATES = {
    2020: 106.47,
    2021: 109.65,
    2022: 117.87,
    2023: 139.85,
    2024: 134.82,
    2025: 129.30,
}
PROVISIONAL_YEARS = {2025}
RATE_SOURCE = ("Economic Survey 2025/2026 Table 6.19, annual average, "
               "Central Bank of Kenya")

# dataset id -> portal flow id, published label and the multiplier that
# turns a raw observation into base units (KSh, tonnes, headcount).
# `scales` keys on the UNIT_MEASURE code carried by the cube.
DATASETS = {
    "bop_annual": {
        "flow": "KE1,DF_BOP_A,1.0",
        "label": "Balance of Payments, Annual",
        "survey_table": "ES 2026 Table 6.16",
        # UNIT_MULT is carried in the cube (0=units, 6=millions); KES rows
        # are KSh million.
        "scales": {"KES": 1e6, "PT_GDP": 1.0},
    },
    "exports_by_destination": {
        "flow": "KE1,DF_INTD_VALUE_TEXPT_DSTN,1.0",
        "label": "Values of Total Exports by Destination",
        "survey_table": "ES 2026 Table 6.13",
        "scales": {"KES": 1e6},
    },
    "imports_by_origin": {
        "flow": "KE1,DF_INTD_VALUE_IMPT_ORGN,1.0",
        "label": "Values of Imports by Origin",
        "survey_table": "ES 2026 Table 6.14",
        "scales": {"KES": 1e6},
    },
    "merchandise_trade": {
        "flow": "KE1,DF_INTD_BOMT,1.0",
        "label": "Balance of Merchandise Trade",
        "survey_table": "ES 2026 Table 6.1",
        "scales": {"KES": 1e6, "PT": 1.0},
    },
    "principal_domestic_exports": {
        "flow": "KE1,DF_INTD_VALUE_PDOM_EXPRT,1.0",
        "label": "Values of Principal Domestic Exports",
        "survey_table": "ES 2026 Table 6.7",
        "scales": {"KES": 1e6},
    },
    "principal_imports": {
        "flow": "KE1,DF_INTD_VALUE_PIMPRT,1.0",
        "label": "Values of Principal Imports",
        "survey_table": "ES 2026 Table 6.8",
        "scales": {"KES": 1e6},
    },
    "fresh_horticulture": {
        "flow": "KE1,DF_AGR_EXP_FRE_HOR_PRO,1.0",
        "label": "Exports of Fresh Horticultural Produce",
        "survey_table": "ES 2026 Table 7.14",
        # Published as '000 tonnes and KSh billion although the cube labels
        # them TONNES and KES.  The cut-flower VALUE rows before 2022 hold the
        # cells of the printed 2024 line, so knbs_sector reads that commodity's
        # volume only.
        "scales": {"TONNES": 1e3, "KES": 1e9},
    },
    "transport_output": {
        "flow": "KE1,DF_TAS_VOT,1.0",
        "label": "Transportation and Storage - Value of Output",
        "survey_table": "ES 2026 Table 13.1",
        "scales": {"KES": 1e6},
    },
    "hotel_bed_nights": {
        "flow": "KE1,DF_TOURISM_HBN_OCC_COR,1.0",
        "label": "Hotel Bed-Nights Occupancy by Country of Residence",
        "survey_table": "ES 2026 Table 12.7",
        # Column header reads '000.
        "scales": {"NUMBER": 1e3},
    },
    "epz_agoa": {
        "flow": "KE1,DF_MFG_EPZ_GARMENT_APPAREL_AGOA,1.0",
        "label": "Selected EPZ Garment/Apparel Indicators under AGOA",
        "survey_table": "ES 2026 Table 10.12",
        "scales": {"NUMBER": 1.0, "KES": 1e6},
    },
    "coffee_growers": {
        "flow": "KE1,DF_AGR_PRO_ARE_AVE_YIE_COF_TYP_GRO,1.0",
        "label": "Production, Area and Average Yield of Coffee by Grower",
        "survey_table": "ES 2026 Table 7.10",
        # Survey prints '000 Ha and '000 tonnes; the cube labels them HA and
        # TONNES.  Verified: 2024/25 total production 51.4 -> 51,400 tonnes.
        # A period is the October-September crop year that ENDS in it, so the
        # cube's 2025 is the Survey's 2024/25 column.
        "scales": {"HA": 1e3, "TONNES": 1e3, "KG_HA": 1.0},
    },
    "tea_growers": {
        "flow": "KE1,DF_AGR_PRO_ARE_AVE_YIE_TEA_TYP_GRO,1.0",
        "label": "Production, Area and Average Yield of Tea by Grower",
        "survey_table": "ES 2026 Table 7.11",
        "scales": {"HA": 1e3, "TONNES": 1e3, "KG_HA": 1.0},
    },
    "mining_output": {
        "flow": "KE1,DF_ENV_NAT_MINING_MINING,1.0",
        "label": "Quantity and Value of Mineral Production",
        "survey_table": "ES 2026 Table 8.9",
        # Quantities are plain tonnes and values KSh million, both verified
        # against the printed table.  Three rows break that rule - gold is
        # kilograms and the two gemstone lines '000 carats, although all three
        # are labelled T - so knbs_sector never reads them.
        "scales": {"T": 1.0, "KES": 1e6},
    },
    "farm_gate_prices": {
        "flow": "KE1,DF_AGR_AVE_GRO_COM_PRI_FAR,1.0",
        "label": "Average Gross Commodity Prices to Farmers",
        "survey_table": "ES 2026 Table 7.4",
        # 1.0: the Survey prices each crop in the unit farmers are paid in
        # (100 Kg, 100 litres, or a tonne) and the cube drops that column, so
        # the published number is kept as it is and every row carries its own
        # unit label.  The portal has refreshed some series since the Survey
        # went to press: coffee, tea, maize, wheat, beef and pig meat match
        # Table 7.4 for 2021-2025, milk, sisal, seed cotton and cane do not, so
        # knbs_sector reads the six that match.
        "scales": {"KES": 1.0},
    },
}


class KnbsError(RuntimeError):
    pass


# ---------------------------------------------------------------------------
# portal client
# ---------------------------------------------------------------------------
def _request(url, body=None, timeout=120):
    headers = {"User-Agent": "tradeflow/1.0 (KNBS reference series)"}
    data = None
    if body is not None:
        data = json.dumps(body).encode("utf-8")
        headers["Content-Type"] = "application/json"
    req = urllib.request.Request(url, data=data, headers=headers)
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        return json.loads(resp.read().decode("utf-8"))


def datacube(flow, criteria=(), last_n=None):
    """POST the dataset cube.  An empty criteria list returns everything."""
    url = f"{API_ROOT}/Core/nodes/{NODE}/datasets/{flow}/data"
    if last_n:
        url += f"?lastNObservations={last_n}"
    return _request(url, list(criteria))


def decode(cube):
    """JSON-stat 2.0 cube -> list of ({dimension: code}, scaled value)."""
    dims = cube["id"]
    sizes = cube["size"]
    index = {d: cube["dimension"][d]["category"]["index"] for d in dims}
    rows = []
    for flat, raw in (cube.get("value") or {}).items():
        remaining = int(flat)
        coords = []
        for name, size in zip(reversed(dims), reversed(sizes)):
            remaining, pos = divmod(remaining, int(size))
            coords.append(index[name][pos])
        coords.reverse()
        rows.append((dict(zip(dims, coords)), raw))
    return rows


def labels(cube):
    return {d: cube["dimension"][d]["category"].get("label", {})
            for d in cube["id"]}


def fetch_dataset(dataset_id):
    """Return {coords: value} for one dataset, values in base units."""
    spec = DATASETS[dataset_id]
    cube = datacube(spec["flow"])
    out = {}
    for coords, raw in decode(cube):
        if raw is None or raw == "":
            continue
        unit = coords.get("UNIT_MEASURE")
        scale = spec["scales"].get(unit)
        if scale is None:
            raise KnbsError(f"{dataset_id}: no unit scale for {unit!r}")
        key = "|".join(f"{k}={coords[k]}" for k in sorted(coords))
        out[key] = float(raw) * scale
    return {"flow": spec["flow"], "label": spec["label"],
            "survey_table": spec["survey_table"],
            "units": spec["scales"], "rows": out,
            "labels": {d: l for d, l in labels(cube).items()}}


def build_snapshot():
    """Every registered dataset as one snapshot dict, without writing it.

    The CLI writes it to :data:`SNAPSHOT_PATH`; the web app stages it as a
    download and installs it on request.
    """
    snapshot = {"generated_at": time.strftime("%Y-%m-%dT%H:%M:%S"),
                "source": "KNBS Open Data Portal (Hub-and-Node API)",
                "api_root": API_ROOT,
                "exchange_rates": {str(y): r for y, r in EXCHANGE_RATES.items()},
                "rate_source": RATE_SOURCE,
                "provisional_years": sorted(PROVISIONAL_YEARS),
                "datasets": {}}
    for name in DATASETS:
        print(f"  fetching {name} ...", flush=True)
        snapshot["datasets"][name] = fetch_dataset(name)
    return snapshot


def refresh(path=SNAPSHOT_PATH):
    snapshot = build_snapshot()
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8") as fh:
        json.dump(snapshot, fh, separators=(",", ":"))
    return snapshot


# ---------------------------------------------------------------------------
# snapshot access
# ---------------------------------------------------------------------------
_cache = {}


def clear_cache():
    """Forget loaded snapshots, so a build after an install reads the new one."""
    _cache.clear()


def load(path=SNAPSHOT_PATH):
    if path not in _cache:
        if not os.path.exists(path):
            raise KnbsError(
                f"KNBS reference snapshot is missing ({path}). "
                "Run: python knbs_reference.py refresh")
        with open(path, encoding="utf-8") as fh:
            _cache[path] = json.load(fh)
    return _cache[path]


def series(dataset_id, path=SNAPSHOT_PATH, **coords):
    """{year: value} for rows matching the given dimensions.

    ``series('principal_domestic_exports', COMMODITY='OW_CT_FLWR',
    UNIT_MEASURE='KES')`` -> ``{2020: 6.09e10, ...}``.  Values are in base
    units (KSh, tonnes, headcount); missing observations are absent.
    """
    snap = load(path)
    entry = snap["datasets"].get(dataset_id)
    if entry is None:
        raise KnbsError(f"KNBS snapshot has no dataset {dataset_id!r}")
    out = {}
    for key, value in entry["rows"].items():
        parts = dict(p.split("=", 1) for p in key.split("|"))
        if all(parts.get(k) == v for k, v in coords.items()):
            out[int(parts["TIME_PERIOD"])] = value
    return out


def rate(year, path=SNAPSHOT_PATH):
    snap = load(path)
    value = snap["exchange_rates"].get(str(int(year)))
    if value is None:
        raise KnbsError(f"No KNBS annual average KSh/USD rate for {year}")
    return value


def to_usd_million(value_kes, year):
    return value_kes / rate(year) / 1e6


def is_provisional(year):
    return int(year) in PROVISIONAL_YEARS


def main(argv):
    if not argv or argv[0] != "refresh":
        print(__doc__)
        print("usage: python knbs_reference.py refresh")
        return 0
    snap = refresh()
    total = sum(len(d["rows"]) for d in snap["datasets"].values())
    print(f"wrote {SNAPSHOT_PATH}: {len(snap['datasets'])} datasets, "
          f"{total} observations")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
