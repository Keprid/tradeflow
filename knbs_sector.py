"""KNBS's account of the industry a product profile sells into.

ITC describes the market; KNBS describes Kenya.  A coffee profile reports
exports in US dollars, while the Economic Survey reports the hectares under
coffee, the tonnes picked by co-operatives as against estates, and the price
farmers were paid at auction.  The second is the context that makes the first
readable, and it comes from a source the rest of the profile does not touch.

Rows are declared as a dataset plus its dimensions and read from the vendored
snapshot in ``knbs_reference``, so nothing here reaches the portal at report
time.  Each series was checked against the printed Survey table it cites
before it was registered, and the quirks that checking found - a column that
mixes crop years with calendar years, a mineral priced in kilograms under a
tonne label - are recorded beside the rows they affect rather than silently
worked around.
"""

import knbs_reference as K
from knbs_reference import KnbsError

# The Survey's own window: five years, the latest it has.
COLUMNS = 5

MIN_YEAR = 2021          # the first year Economic Survey 2026 prints in full


def _row(dataset, label, unit, digits=None, divide=1.0, min_year=None,
         **coords):
    """One table row.  ``digits`` is the report's decimal places, one by
    default as in the Survey itself; a series published in whole hundreds
    declares zero.  ``min_year`` drops years known to be mis-filed."""
    return {"dataset": dataset, "label": label, "unit": unit,
            "digits": digits, "divide": divide, "min_year": min_year,
            "coords": coords}


# The Survey prices farm produce in the unit the farmer is paid in, and the
# portal drops that column, so each price row carries its own unit text.
_PRICE = "KSh per 100 Kg"

VIEWS = {
    # HS 0901.  Table 7.10 counts the October-September crop year and labels it
    # by the year it ends in, so the 2025 column is the 2024/25 season.
    "0901": {
        "title": "coffee",
        "rows": [
            _row("coffee_growers", "Area under coffee, all growers",
                 "Hectares", digits=0, min_year=MIN_YEAR, AGR_INDIC="AREA",
                 ITEM="CP012201", GROWER_TYPE="_T", UNIT_MEASURE="HA"),
            _row("coffee_growers", "Coffee produced, all growers",
                 "Tonnes", digits=0, min_year=MIN_YEAR,
                 AGR_INDIC="PRODUCTION", ITEM="CP012201", GROWER_TYPE="_T",
                 UNIT_MEASURE="TONNES"),
            _row("coffee_growers", "Produced by co-operatives",
                 "Tonnes", digits=0, min_year=MIN_YEAR,
                 AGR_INDIC="PRODUCTION", ITEM="CP012201",
                 GROWER_TYPE="COOPERATIVES", UNIT_MEASURE="TONNES"),
            _row("coffee_growers", "Produced by estates",
                 "Tonnes", digits=0, min_year=MIN_YEAR,
                 AGR_INDIC="PRODUCTION", ITEM="CP012201",
                 GROWER_TYPE="ESTATES", UNIT_MEASURE="TONNES"),
            _row("coffee_growers", "Average yield, co-operatives",
                 "Kg per hectare", min_year=MIN_YEAR,
                 AGR_INDIC="AVG_YIELD", ITEM="CP012201",
                 GROWER_TYPE="COOPERATIVES", UNIT_MEASURE="KG_HA"),
            _row("coffee_growers", "Average yield, estates",
                 "Kg per hectare", min_year=MIN_YEAR,
                 AGR_INDIC="AVG_YIELD", ITEM="CP012201",
                 GROWER_TYPE="ESTATES", UNIT_MEASURE="KG_HA"),
            _row("farm_gate_prices", "Average auction price received by farmers",
                 _PRICE, min_year=MIN_YEAR, ITEM="CP012201",
                 UNIT_MEASURE="KES"),
        ],
        "notes": ["KNBS counts coffee in crop years running October to "
                  "September, and labels each by the year it ends in, so the "
                  "2025 column is the 2024/25 season.  Yield divides current "
                  "production by the acreage three years earlier."],
    },
    # HS 0902.  Table 7.11 is calendar years; the portal's green-leaf rows
    # before 2020 carry made-tea weights, so nothing older than the Survey's
    # own window is read here.
    "0902": {
        "title": "tea",
        "rows": [
            _row("tea_growers", "Area under tea, all growers", "Hectares",
                 digits=0, min_year=MIN_YEAR, AGR_INDIC="AREA", ITEM="TEA",
                 GROWER_TYPE="_T", UNIT_MEASURE="HA"),
            _row("tea_growers", "Green leaf produced, all growers",
                 "Tonnes", digits=0, min_year=MIN_YEAR,
                 AGR_INDIC="PRODUCTION", ITEM="TEA01", GROWER_TYPE="_T",
                 UNIT_MEASURE="TONNES"),
            _row("tea_growers", "Made tea produced, all growers",
                 "Tonnes", digits=0, min_year=MIN_YEAR,
                 AGR_INDIC="PRODUCTION", ITEM="TEA09", GROWER_TYPE="_T",
                 UNIT_MEASURE="TONNES"),
            _row("tea_growers", "Made tea produced by smallholders",
                 "Tonnes", digits=0, min_year=MIN_YEAR,
                 AGR_INDIC="PRODUCTION", ITEM="TEA09",
                 GROWER_TYPE="SMALL_HOLDERS", UNIT_MEASURE="TONNES"),
            _row("tea_growers", "Made tea produced by estates",
                 "Tonnes", digits=0, min_year=MIN_YEAR,
                 AGR_INDIC="PRODUCTION", ITEM="TEA09", GROWER_TYPE="ESTATES",
                 UNIT_MEASURE="TONNES"),
            _row("tea_growers", "Average yield of made tea, smallholders",
                 "Kg per hectare", min_year=MIN_YEAR, AGR_INDIC="AVG_YIELD",
                 ITEM="TEA09", GROWER_TYPE="SMALL_HOLDERS",
                 UNIT_MEASURE="KG_HA"),
            _row("farm_gate_prices", "Average auction price received by farmers",
                 _PRICE, min_year=MIN_YEAR, ITEM="TEA", UNIT_MEASURE="KES"),
        ],
        "notes": ["Yield divides current production by the area four years "
                  "earlier."],
    },
    # HS 0603, inside Table 7.14.  The portal's cut-flower earnings before
    # 2022 hold cells of the published 2024 line, so that row starts at 2022
    # and the blank years are marked rather than filled.
    "0603": {
        "title": "cut flowers",
        "rows": [
            _row("fresh_horticulture", "Cut flowers exported", "Tonnes",
                 digits=0, AGR_INDIC="VOLUME", ITEM="CP0931",
                 UNIT_MEASURE="TONNES"),
            _row("fresh_horticulture", "Cut flowers exported, value",
                 "KSh billion", divide=1e9, min_year=2022,
                 AGR_INDIC="VALUE", ITEM="CP0931", UNIT_MEASURE="KES"),
            _row("fresh_horticulture",
                 "All fresh horticulture produce exported", "Tonnes",
                 digits=0, AGR_INDIC="VOLUME", ITEM="_T",
                 UNIT_MEASURE="TONNES"),
            _row("fresh_horticulture",
                 "All fresh horticulture produce exported, value",
                 "KSh billion", divide=1e9, AGR_INDIC="VALUE", ITEM="_T",
                 UNIT_MEASURE="KES"),
        ],
        "notes": ["KNBS excludes processed horticultural produce and nuts.",
                  "The portal's cut-flower earnings before 2022 are mis-filed, "
                  "so those years are left blank."],
    },
    # HS 06, 07 and 08 together - the same table, read across its three lines.
    "06": {
        "title": "horticulture",
        "rows": [
            _row("fresh_horticulture", "Cut flowers exported", "Tonnes",
                 digits=0, AGR_INDIC="VOLUME", ITEM="CP0931",
                 UNIT_MEASURE="TONNES"),
            _row("fresh_horticulture", "Fruits exported", "Tonnes",
                 digits=0, AGR_INDIC="VOLUME", ITEM="FRUITS",
                 UNIT_MEASURE="TONNES"),
            _row("fresh_horticulture", "All fresh horticulture produce",
                 "Tonnes", digits=0, AGR_INDIC="VOLUME", ITEM="_T",
                 UNIT_MEASURE="TONNES"),
            _row("fresh_horticulture",
                 "All fresh horticulture produce, value", "KSh billion",
                 divide=1e9, AGR_INDIC="VALUE", ITEM="_T", UNIT_MEASURE="KES"),
        ],
        "notes": ["KNBS excludes processed horticultural produce and nuts, and "
                  "publishes the value of the whole basket: it does not split "
                  "earnings between flowers, fruits and vegetables."],
    },
    # HS 25 and 26.  Table 8.9 prices minerals in KSh million, and labels the
    # three rows that are not weighed in tonnes - gold in kilograms, the two
    # gemstone lines in thousand carats.  Those rows are not read.
    "25": {
        "title": "mining",
        "rows": [
            _row("mining_output", "Value of all minerals produced",
                 "KSh million", divide=1e6, MINERAL="_T_VAL_MINERAL",
                 ENV_INDC="MINING", UNIT_MEASURE="KES"),
            _row("mining_output", "Soda ash produced", "Tonnes",
                 MINERAL="SODA_ASH", ENV_INDC="MINING", UNIT_MEASURE="T"),
            _row("mining_output", "Soda ash, value", "KSh million",
                 divide=1e6, MINERAL="SODA_ASH", ENV_INDC="MINING",
                 UNIT_MEASURE="KES"),
            _row("mining_output", "Crushed refined soda produced", "Tonnes",
                 MINERAL="CRUS_REFI_SODA", ENV_INDC="MINING",
                 UNIT_MEASURE="T"),
            _row("mining_output", "Titanium ore minerals produced", "Tonnes",
                 MINERAL="TITA_ORE_MINERALS", ENV_INDC="MINING",
                 UNIT_MEASURE="T"),
            _row("mining_output", "Titanium ore minerals, value",
                 "KSh million", divide=1e6, MINERAL="TITA_ORE_MINERALS",
                 ENV_INDC="MINING", UNIT_MEASURE="KES"),
        ],
        "notes": ["KNBS measures what Kenya's mines produce, which is not what "
                  "Kenya exports: most of the value above leaves as ore for "
                  "processing elsewhere."],
    },
    # HS 02 and 16.  The portal has no livestock numbers, so the only sectoral
    # line the Survey prints for meat is what an animal makes at the farm gate.
    "02": {
        "title": "meat",
        "rows": [
            _row("farm_gate_prices", "Beef (third grade), average farm-gate price",
                 _PRICE, ITEM="CP011221", UNIT_MEASURE="KES"),
            _row("farm_gate_prices", "Pig meat, average farm-gate price",
                 _PRICE, ITEM="CP011222", UNIT_MEASURE="KES"),
        ],
        "notes": ["Farm-gate prices are what the seller receives, not the "
                  "export price, and KNBS publishes no herd numbers or "
                  "slaughter weights in this series."],
    },
}

VIEWS["26"] = VIEWS["25"]
VIEWS["16"] = VIEWS["02"]
VIEWS["07"] = VIEWS["06"]
VIEWS["08"] = VIEWS["06"]


def _prefixes(code):
    """The four- and two-digit headings one config code stands for.

    Configs write codes as '0901', '06', '4420.10' or '6214.10-6214.90'; the
    leading digits of the first interval are what KNBS's chapters and headings
    are keyed on.
    """
    first = str(code).strip().split("-")[0].split(".")[0]
    if not first.isdigit():
        return []
    return ([first[:4], first[:2]] if len(first) >= 4 else [first[:2]])


def view_for(codes):
    """The registered view for a config's ``include_codes``, or None.

    Headings are matched before chapters so that a profile of HS 0603 gets the
    cut-flower table rather than the whole of horticulture.
    """
    codes = list(codes or [])
    for width in (4, 2):
        for code in codes:
            for prefix in _prefixes(code):
                if len(prefix) == width and prefix in VIEWS:
                    return prefix
    return None


def _source(dataset):
    entry = K.load()["datasets"].get(dataset) or {}
    return entry.get("survey_table") or ""


def context(codes, columns=COLUMNS):
    """The KNBS industry behind a product, ready for one table.

    Returns ``None`` when nothing is registered for the codes, when the
    snapshot has no such dataset, or when every candidate row is empty, so a
    profile of dairy or crafts prints no sectoral block at all.  Values are
    already in the unit the row is labelled with.
    """
    key = view_for(codes)
    if key is None:
        return None
    view = VIEWS[key]
    loaded = []
    for row in view["rows"]:
        try:
            values = K.series(row["dataset"], **row["coords"])
        except KnbsError:
            return None
        values = {y: v / row["divide"] for y, v in values.items()
                  if v and y >= max(row["min_year"] or 0, MIN_YEAR)}
        if values:
            loaded.append((row, values))
    if not loaded:
        return None
    latest = max(max(values) for _, values in loaded)
    start = latest - columns + 1
    rows = [{"label": row["label"], "unit": row["unit"], "digits": row["digits"],
             "values": {y: v for y, v in values.items() if y >= start}}
            for row, values in loaded]
    # A column is only worth printing if some row can fill it.
    years = [y for y in range(start, latest + 1)
             if any(y in r["values"] for r in rows)]
    rows = [dict(r, values={y: r["values"][y] for y in years
                            if y in r["values"]}) for r in rows]
    rows = [r for r in rows if r["values"]]
    if len(rows) < 2:
        return None
    return {"title": view["title"], "view": key, "rows": rows, "years": years,
            "provisional": [y for y in years if K.is_provisional(y)],
            "sources": list(dict.fromkeys(_source(r["dataset"])
                                          for r, _ in loaded)),
            "notes": list(view.get("notes") or [])}
