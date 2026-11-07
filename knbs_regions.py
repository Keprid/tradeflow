"""KNBS's own regional split of Kenya's trade in goods.

The databrowser ``TRADE_PARTNER`` dimension mixes countries with the
continents, sub-regions and blocs KNBS groups those countries into, so a code
is only meaningful together with its parent.  Economic Survey tables 6.14
(imports by origin) and 6.13 (exports by destination) print that tree; this
module describes it and turns one year of a cube into report rows.

The groupings are not guesses.  In the cube they add up to the published
total, and ``checks`` re-proves it for every year reported, so a change in
KNBS's taxonomy shows up as a failed identity rather than a silently wrong
share:

    Total  = Africa + America + Asia + Europe + Oceania + the unallocated tail
    Asia   = Middle East + Far East
    Europe = Western Europe + Eastern Europe

Africa is the one continent KNBS does not subdivide in the cube: the East
African Community sits inside it as a bloc, which is why blocs are reported as
"of which" lines instead of as components of the total.
"""

import knbs_reference as K

# (code, display name, indent level), in the order Economic Survey tables 6.13
# and 6.14 print them: total, then each continent with its sub-regions beneath.
TREE = [
    ("_T",  "Total",                0),
    ("F1",  "Africa",               1),
    ("S1",  "Asia",                 1),
    ("F97", "Middle East",          2),
    ("F_E", "Far East",             2),
    ("E1",  "Europe",               1),
    ("E5",  "Western Europe",       2),
    ("E2",  "Eastern Europe",       2),
    ("A1",  "America",              1),
    ("O3A", "Australia & Oceania",  1),
]

CONTINENTS = ["F1", "S1", "E1", "A1", "O3A"]

# The unallocated tail.  Exports carry a "Total of all other countries" line
# that splits into aircraft and ship stores plus the unspecified remainder;
# imports publish the unspecified remainder alone.  Whichever codes the cube
# holds is what closes the total.
OTHER_ROWS = [
    ("_TOTHER",    "All other destinations",         1),
    ("ACRT_SHIP",  "Aircraft and ships stores",      2),
    ("ALL_OT_CNT", "Other countries not specified",  2),
]
OTHER_TAIL = ("_TOTHER", "ALL_OT_CNT")

# Blocs that cut across the region tree, shown as "of which" lines beneath the
# region they sit in.
BLOCS = [
    ("B6",  "European Union",         "E5"),
    ("R19", "East African Community", "F1"),
]

# KNBS carries a second TRADE_PARTNER line labelled the same as another
# partner (GB1/United Kingdom, CD1/Democratic Republic of Congo, SO1/Somalia).
# Both lines of a pair can hold a value in the same year and Survey does not
# say how they differ, so only the canonical code is ever read.
PARTNER_VARIANTS = {"GB1", "CD1", "SO1"}

# Every TRADE_PARTNER code that is a grouping rather than a country.
GROUP_CODES = ({code for code, _, _ in TREE}
               | {code for code, _, _ in OTHER_ROWS}
               | {code for code, _, _ in BLOCS}
               | PARTNER_VARIANTS)

DATASETS = {
    "exports_by_destination": "Exports",
    "imports_by_origin": "Imports",
}

# A year whose identities miss the parent by more than this share is reported
# with reconciles=False: KNBS's own rounding, not a taxonomy change.
RECONCILE_TOLERANCE_PCT = 0.5


def _entry(dataset_id):
    snap = K.load()
    entry = snap["datasets"].get(dataset_id)
    if entry is None:
        raise K.KnbsError(f"KNBS snapshot has no dataset {dataset_id!r}")
    return entry


def _year_values(dataset_id, year):
    """{partner code: USD million} for one year of one dataset."""
    rates = {int(y) for y in K.load()["exchange_rates"]}
    if int(year) not in rates:
        raise K.KnbsError(
            f"no published KSh/USD annual average rate for {year}")
    out = {}
    for key, value in _entry(dataset_id)["rows"].items():
        parts = dict(p.split("=", 1) for p in key.split("|"))
        if parts["UNIT_MEASURE"] != "KES":
            continue
        if int(parts["TIME_PERIOD"]) != int(year):
            continue
        code = parts["TRADE_PARTNER"]
        if code in PARTNER_VARIANTS:
            continue          # see PARTNER_VARIANTS: unexplained second lines
        out[code] = K.to_usd_million(value, year)
    return out


def years(dataset_id="exports_by_destination"):
    """Years the cube carries a total for *and* a published rate to convert."""
    out = set()
    for key in _entry(dataset_id)["rows"]:
        parts = dict(p.split("=", 1) for p in key.split("|"))
        if parts["TRADE_PARTNER"] == "_T" and parts["UNIT_MEASURE"] == "KES":
            out.add(int(parts["TIME_PERIOD"]))
    rates = {int(y) for y in K.load()["exchange_rates"]}
    return sorted(out & rates)


def partners(dataset_id="exports_by_destination"):
    """{code: label} for the individual countries one cube names.

    KNBS publishes only its main partners separately; the rest hide inside
    the continent lines, so a partner absent here simply has no bilateral
    series to check against.
    """
    labels = _entry(dataset_id)["labels"].get("TRADE_PARTNER", {})
    return {code: label for code, label in labels.items()
            if code not in GROUP_CODES and not code.startswith("OTHER_")}


def _check(parent, children, values):
    """How far the children fall short of the parent, and in what share."""
    if parent not in values or not children:
        return None
    total = sum(values[c] for c in children if c in values)
    if not values[parent]:
        return None
    residual = values[parent] - total
    return {"parent": parent, "children": list(children),
            "residual": round(residual, 1),
            "residual_pct": round(100.0 * residual / values[parent], 2)}


def _row(code, name, level, values, total, **extra):
    value = values.get(code)
    return dict(extra, code=code, name=name, level=level, value=value,
                share=(value / total) if value is not None else None)


def _tail_rows(values, total):
    """The lines that close the total, in whatever shape the cube uses."""
    names = {code: name for code, name, _ in OTHER_ROWS}
    rows = []
    if "_TOTHER" in values:
        rows.append(_row("_TOTHER", names["_TOTHER"], 1, values, total))
        for code in ("ACRT_SHIP", "ALL_OT_CNT"):
            if code in values:
                rows.append(_row(code, names[code], 2, values, total))
    else:
        for code in ("ALL_OT_CNT", "ACRT_SHIP"):
            if code in values:
                rows.append(_row(code, names[code], 1, values, total))
    return rows


def decomposition(dataset_id="exports_by_destination", year=None, blocs=True):
    """One year of a KNBS trade dataset split the way KNBS publishes it.

    Returns the report rows (``level`` is the indent, ``bloc`` marks the
    "of which" lines), the total they are shares of, and the additive
    identities re-proved against the same figures.
    """
    if dataset_id not in DATASETS:
        raise K.KnbsError(f"no KNBS region tree for {dataset_id!r}")
    if year is None:
        available = years(dataset_id)
        if not available:
            raise K.KnbsError(f"KNBS {dataset_id} snapshot has no usable year")
        year = available[-1]
    values = _year_values(dataset_id, year)
    total = values.get("_T")
    if not total:
        raise K.KnbsError(f"KNBS {dataset_id} has no total for {year}")

    rows = [_row(code, name, level, values, total) for code, name, level in TREE]
    rows += _tail_rows(values, total)
    if blocs:
        rows = _insert_blocs(rows, [
            _row(code, name, 0, values, total, bloc=True, under=under)
            for code, name, under in BLOCS if code in values])

    tail = next((code for code in OTHER_TAIL if code in values), None)
    checks = [c for c in (
        _check("_T", CONTINENTS + ([tail] if tail else []), values),
        _check("S1", ["F97", "F_E"], values),
        _check("E1", ["E5", "E2"], values),
        _check("_TOTHER", ["ACRT_SHIP", "ALL_OT_CNT"], values)) if c]
    worst = max((abs(c["residual_pct"]) for c in checks), default=0.0)
    entry = _entry(dataset_id)
    return {
        "dataset": dataset_id,
        "flow": DATASETS[dataset_id],
        "year": int(year),
        "provisional": K.is_provisional(year),
        "rows": rows,
        "total": total,
        "checks": checks,
        "reconciles": worst <= RECONCILE_TOLERANCE_PCT,
        "max_residual_pct": round(worst, 2),
        "source": (f"KNBS Open Data Portal, {entry['label']} "
                   f"({entry['survey_table']})"),
        "rate_source": K.RATE_SOURCE,
    }


def _insert_blocs(rows, blocs):
    """Place each bloc as an 'of which' line under the region it sits in."""
    out = []
    for row in rows:
        out.append(row)
        for bloc in blocs:
            if bloc["under"] == row["code"]:
                bloc = dict(bloc)
                bloc["level"] = row["level"] + 1
                out.append(bloc)
    return out
