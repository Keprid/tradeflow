"""Cross-check our USD figures against KNBS's published KSh series.

KNBS values Kenya's trade in Kenyan shillings; ITC Trade Map and UNCTAD give
US dollars.  Because both ultimately describe the same customs flow, the KNBS
series divided by the Economic Survey's annual average exchange rate is an
independent check on the numbers a report prints.

Every check returns the same shape so report builders can render it without
knowing which series produced it:

    {
      "label":  "Cut flowers (HS 0603) against KNBS 'O/W Cut Flowers'",
      "rows":   [{"year": 2024, "ours": 722.9, "knbs": 719.3,
                  "diff_pct": -0.49, "provisional": True}, ...],
      "max_abs_diff_pct": 0.49,
      "status": "consistent" | "review" | "unavailable",
      "source": "KNBS Open Data Portal ... (ES 2026 Table 6.7)",
      "rate_source": "Economic Survey ... Table 6.19 ...",
    }

``ours`` is always USD million.  A year we have no KNBS rate for, or no KNBS
observation for, is dropped from the rows rather than guessed; if nothing
survives the status is ``unavailable`` and the report should say so.
"""

import re

import knbs_reference as K
import knbs_regions as R

# Above this absolute percentage difference the pair is flagged for review.
# Customs valuation, rounding in the printed Survey table and re-exports
# routinely move a series by a few tenths of a percent.
REVIEW_THRESHOLD_PCT = 2.0

# HS prefix -> KNBS "Values of Principal Domestic Exports" commodity code.
# Longest matching prefix wins; a profile can override the lookup by setting
# ``knbs_commodity`` in its config.
HS_TO_KNBS_COMMODITY = {
    "0603": "OW_CT_FLWR",
    "0901": "COFF_URTD",
    "0902": "TEA",
    "03": "FISH",
    "06": "HRTCTRE",
    "07": "VEG_FRT",
    "08": "VEG_FRT",
    "3301": "ESS_OIL",
}

# Our downloads name partners the way ITC does; KNBS names some of them
# differently.  Keys and values are both reduced by :func:`_partner_key`.
PARTNER_ALIASES = {
    "united states america": "usa",
    "united states": "usa",
    "iran islamic republic": "iran",
    "yemen": "yemen arab republic",
    "russia": "russian federation",
    "china taiwan province": "taiwan",
    "united republic tanzania": "tanzania",
    "czechia": "czech republic",
    "swaziland": "eswatini",
    "south korea": "korea republic",
}

# Words that carry no meaning when matching one partner name to another.
_PARTNER_STOP = {"the", "of", "and", "in"}

# ITC abbreviates what KNBS spells out.
_PARTNER_ABBREVIATIONS = {"rep": "republic", "dem": "democratic"}


def _partner_key(name):
    """A partner name reduced to what both sources agree on."""
    text = re.sub(r"\(.*?\)", " ", str(name or "").lower())
    text = re.sub(r"[^a-z0-9 ]", " ", text)
    words = (_PARTNER_ABBREVIATIONS.get(w, w) for w in text.split())
    return " ".join(w for w in words if w not in _PARTNER_STOP)


def partner_code(name, dataset_id="exports_by_destination"):
    """KNBS's TRADE_PARTNER code for a partner we name, or ``None``.

    KNBS names only its main partners separately, so a partner outside that
    list has no bilateral line to check against rather than a missing value.
    """
    labels = R.partners(dataset_id)
    by_key = {_partner_key(label): code for code, label in labels.items()}
    key = _partner_key(name)
    code = by_key.get(key)
    if code is None and key in PARTNER_ALIASES:
        code = by_key.get(_partner_key(PARTNER_ALIASES[key]))
    return code


def _finish(label, ours, knbs, source):
    rows = []
    for year in sorted(set(ours) & set(knbs)):
        our_value, knbs_value = ours[year], knbs[year]
        if not our_value or not knbs_value:
            continue
        diff = 100.0 * (knbs_value - our_value) / our_value
        rows.append({"year": year, "ours": our_value, "knbs": knbs_value,
                     "diff_pct": round(diff, 2),
                     "provisional": K.is_provisional(year)})
    if not rows:
        return {"label": label, "rows": [], "max_abs_diff_pct": None,
                "status": "unavailable", "source": source,
                "review_threshold_pct": REVIEW_THRESHOLD_PCT,
                "rate_source": K.RATE_SOURCE}
    worst = max(abs(r["diff_pct"]) for r in rows)
    return {"label": label, "rows": rows, "max_abs_diff_pct": round(worst, 2),
            "status": "consistent" if worst <= REVIEW_THRESHOLD_PCT else "review",
            "source": source, "review_threshold_pct": REVIEW_THRESHOLD_PCT,
            "rate_source": K.RATE_SOURCE}


def _usd_millions(dataset, **coords):
    """KNBS base-unit observations for one combination, as USD million.

    Years without a published annual average rate are dropped.
    """
    raw = K.series(dataset, **coords)
    return {y: K.to_usd_million(v, y) for y, v in raw.items()
            if y in K.EXCHANGE_RATES}


def _describe(dataset):
    snap = K.load()
    entry = snap["datasets"].get(dataset) or {}
    return (f"KNBS Open Data Portal, {entry.get('label', dataset)} "
            f"({entry.get('survey_table', '')})")


def _tidy(text):
    """KNBS's own labels: collapsed spacing, no trailing punctuation."""
    return " ".join(str(text or "").split()).strip(" ,;")


def check_product(ours_by_year, hs_code=None, commodity=None):
    """Product profile Table 1 'Kenya' row against the KNBS commodity line."""
    code = commodity or HS_TO_KNBS_COMMODITY.get(
        str(hs_code or "")[:4]) or HS_TO_KNBS_COMMODITY.get(
        str(hs_code or "")[:2])
    if code is None:
        return {"label": "KNBS cross-check", "rows": [], "status": "unavailable",
                "max_abs_diff_pct": None, "source": "no KNBS commodity mapping",
                "review_threshold_pct": REVIEW_THRESHOLD_PCT,
                "rate_source": K.RATE_SOURCE}
    snap = K.load()
    knbs_label = _tidy(snap["datasets"]["principal_domestic_exports"]["labels"]
                       ["COMMODITY"].get(code, code))
    ours = {int(y): v for y, v in ours_by_year.items()}
    knbs = _usd_millions("principal_domestic_exports",
                         COMMODITY=code, UNIT_MEASURE="KES")
    return _finish(f"KNBS '{knbs_label}' vs our USD figures",
                   ours, knbs, _describe("principal_domestic_exports"))


def check_services_exports(ours_by_year):
    ours = {int(y): v for y, v in ours_by_year.items()}
    knbs = _usd_millions("bop_annual", BOP_INDIC="S", ACC_ENTRY="C",
                         UNIT_MEASURE="KES", UNIT_MULT="6")
    return _finish("KNBS services credit vs UNCTAD exports",
                   ours, knbs, _describe("bop_annual"))


def check_services_imports(ours_by_year):
    ours = {int(y): v for y, v in ours_by_year.items()}
    knbs = _usd_millions("bop_annual", BOP_INDIC="S", ACC_ENTRY="D",
                         UNIT_MEASURE="KES", UNIT_MULT="6")
    return _finish("KNBS services debit vs UNCTAD imports",
                   ours, knbs, _describe("bop_annual"))


def check_goods_exports(ours_by_year):
    """Kenya's total merchandise exports, against Survey table 6.1."""
    ours = {int(y): v for y, v in ours_by_year.items()}
    knbs = _usd_millions("merchandise_trade", TRADE_FL="EXPORT",
                         INTD_INDIC="BOMT", UNIT_MEASURE="KES")
    return _finish("KNBS merchandise exports vs our USD figures",
                   ours, knbs, _describe("merchandise_trade"))


def check_partner(partner_name, ours_by_year, direction="exports"):
    """One leg of a bilateral flow against KNBS's own customs line for it.

    ``direction`` is "exports" (Kenya selling to the partner) or "imports".
    Returns ``None`` when KNBS does not break the partner out at all, so the
    caller prints no table rather than an empty one.
    """
    dataset = ("exports_by_destination" if direction == "exports"
               else "imports_by_origin")
    code = partner_code(partner_name, dataset)
    if code is None:
        return None
    label = _tidy(R.partners(dataset)[code])
    ours = {int(y): v for y, v in (ours_by_year or {}).items() if v}
    knbs = _usd_millions(dataset, TRADE_PARTNER=code, UNIT_MEASURE="KES")
    return _finish(f"KNBS {direction} ({label}) vs our USD figures",
                   ours, knbs, _describe(dataset))


def partner_checks(partner_name, exports_by_year=None, imports_by_year=None):
    """The bilateral cross-checks a partner's brief can be held to.

    A partner KNBS buries in its "Other Countries in ..." line yields an empty
    list: there is nothing to reconcile against, which is not the same as a
    check that failed.
    """
    out = []
    for direction, ours in (("exports", exports_by_year),
                            ("imports", imports_by_year)):
        if not ours:
            continue
        check = check_partner(partner_name, ours, direction)
        if check is not None:
            out.append(check)
    return out


def services_2025():
    """The Balance of Payments services line, in USD million, by year.

    Returns ``{"exports": {...}, "imports": {...}, "balance": {...}}`` for
    every year the snapshot carries, which currently runs one year beyond
    UNCTAD's EBOPS series.
    """
    exports = _usd_millions("bop_annual",
                            BOP_INDIC="S", ACC_ENTRY="C",
                            UNIT_MEASURE="KES", UNIT_MULT="6")
    imports = _usd_millions("bop_annual",
                            BOP_INDIC="S", ACC_ENTRY="D",
                            UNIT_MEASURE="KES", UNIT_MULT="6")
    balance = {y: exports[y] - imports[y] for y in sorted(set(exports) & set(imports))}
    return {"exports": exports, "imports": imports, "balance": balance}


def services_source():
    """Citation for the Balance of Payments services line."""
    return _describe("bop_annual")


def services_extension(upto_year):
    """The years KNBS publishes beyond ``upto_year``, as USD million.

    UNCTAD's EBOPS series trails KNBS's Balance of Payments by a year, so this
    is the tail that carries a services report forward.  Returns lists aligned
    to ``years`` plus a per-year ``provisional`` flag, or ``None`` when the
    snapshot holds nothing past ``upto_year``.
    """
    s = services_2025()
    years = [y for y in sorted(s["balance"]) if y > upto_year]
    if not years:
        return None
    return {"years": years,
            "exports": [s["exports"][y] for y in years],
            "imports": [s["imports"][y] for y in years],
            "balance": [s["balance"][y] for y in years],
            "provisional": [K.is_provisional(y) for y in years],
            "source": services_source(),
            "rate_source": K.RATE_SOURCE}
