"""Trade-agreement view of Kenya's partner-level trade.

The optional ''Table 8 Kenya Exports/Imports by Partner'' workbooks (built by
make_tables from the ''kenyas-exports-to-world-by-importer'' and
''kenyas-imports-from-world-by-exporter'' downloads) list every partner
country Kenya trades with.  This module groups that partner-level data into
the market blocs Kenya belongs to under its trade agreements (EAC, COMESA,
AfCFTA, AGOA, EU-EPA/UK, GSP markets and bilateral partners) so the report
can feature how Kenya performs against the markets covered by those
agreements.

Blocs overlap by design -- EAC is a subset of COMESA and both are subsets of
AfCFTA -- so the same trade is shown under each market-access umbrella, and
the narrative notes the nesting.
"""

import json
import os

from country_names import display_name, is_africa

_CONFIG_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                            "config", "trade_agreements.json")

_UNIT = "USD billion"


def _key(name):
    """Normalised matching key: lowercase alphanumerics only."""
    return "".join(ch for ch in str(name or "").lower() if ch.isalnum())


def load_agreements(path=None):
    """Load and normalise config/trade_agreements.json.

    Returns a list of agreement dicts; each carries ``members`` as a list of
    names (strings) or ``{"name": ..., "aliases": [...]}`` objects, or the
    flag ``africa: true`` meaning every African market.
    """
    with open(path or _CONFIG_PATH, "r", encoding="utf-8") as f:
        raw = json.load(f)
    agreements = []
    for ag in raw.get("agreements", []):
        members = ag.get("members") or []
        member_objs = []
        for m in members:
            if isinstance(m, str):
                member_objs.append({"name": m, "aliases": []})
            else:
                member_objs.append({
                    "name": m.get("name", ""),
                    "aliases": m.get("aliases") or [],
                })
        ag = dict(ag)
        ag["members"] = member_objs
        ag["_keys"] = set()
        for m in member_objs:
            ag["_keys"].update(_member_keys(m["name"]))
            for a in m["aliases"]:
                ag["_keys"].update(_member_keys(a))
        agreements.append(ag)
    return agreements


def _member_keys(name):
    out = {_key(name)}
    disp = display_name(name)
    if disp:
        out.add(_key(disp))
    return out


def _item_keys(label):
    out = {_key(label)}
    disp = display_name(label)
    if disp and disp != label:
        out.add(_key(disp))
    return out


def _matches_agreement(ag, label):
    """True when ``label`` is a partner country covered by agreement ``ag``.

    For the AfCFTA (``africa``) bloc every African market counts; otherwise
    the label must match one of the agreement's member names/aliases.
    """
    if ag.get("africa"):
        return is_africa(label)
    keys = _item_keys(label)
    return bool(keys & ag["_keys"])


def _series_for_items(items, years):
    """Aligned-to-``years`` list of per-year sums for a filtered item list."""
    per_year = {y: [] for y in years}
    for it in (items or []):
        for k, y in enumerate(years):
            v = it["years"][k] if k < len(it["years"]) else None
            per_year[y].append(v)
    out = []
    for y in years:
        vals = [v for v in per_year[y] if v is not None]
        out.append(sum(vals) if vals else None)
    return out


def _sum_series(series, year, years):
    if year not in years:
        return None
    return series[years.index(year)]


def _total_by_year(parsed):
    total = parsed.get("total")
    years = parsed.get("years") or []
    if not total or not total.get("years"):
        return {year: None for year in years}
    out = {}
    for k, year in enumerate(years):
        out[year] = total["years"][k] if k < len(total["years"]) else None
    return out


def _top_partners(items, year, years):
    out = []
    for it in (items or []):
        k = years.index(year) if year in years else -1
        v = it["years"][k] if (k >= 0 and k < len(it["years"])) else None
        if v is None:
            continue
        out.append({"name": display_name(it["name"]), "value": v})
    out.sort(key=lambda d: d["value"], reverse=True)
    return out[:3]


def analyze(exp, imp, agreements, year):
    """Aggregate parsed Table 8 exports/imports by agreement bloc.

    ``exp`` / ``imp`` : outputs of generate_report.parse_rank_table on the
    two Table 8 workbooks.  Values are in USD billions.

    Returns a dict ready for the narrative and chart sections, or None when
    the data cannot be read.
    """
    if exp is None or imp is None:
        return None
    years = list(exp.get("years") or [])
    if not years:
        return None
    if year not in years:
        year = years[-1]

    te = _total_by_year(exp)
    ti = _total_by_year(imp)

    # Markets covered by at least one agreement (deduplicated union), used for
    # the overall coverage share so overlapping blocs (EAC<COMESA<AfCFTA) are
    # not double-counted.
    covered_export_names = set()
    covered_import_names = set()
    exp_items = exp.get("items") or []
    imp_items = imp.get("items") or []
    for ag in agreements:
        for it in exp_items:
            if it["name"] != "Kenya" and _matches_agreement(ag, it["name"]):
                covered_export_names.add(it["name"])
        for it in imp_items:
            if it["name"] != "Kenya" and _matches_agreement(ag, it["name"]):
                covered_import_names.add(it["name"])
    covered_exports = _sum_series(_series_for_items(
        [it for it in exp_items if it["name"] in covered_export_names], years), year, years)
    covered_imports = _sum_series(_series_for_items(
        [it for it in imp_items if it["name"] in covered_import_names], years), year, years)

    rows = []
    for ag in agreements:
        exp_members = [it for it in exp_items
                       if _matches_agreement(ag, it["name"])
                       and _key(it["name"]) != _key("Kenya")]
        imp_members = [it for it in imp_items
                       if _matches_agreement(ag, it["name"])
                       and _key(it["name"]) != _key("Kenya")]
        matched_names = {it["name"] for it in exp_members + imp_members}
        if matched_names:
            e_series = _series_for_items(exp_members, years)
            i_series = _series_for_items(imp_members, years)
        else:
            e_series = [None] * len(years)
            i_series = [None] * len(years)

        exports = e_series[years.index(year)] if year in years else None
        imports = i_series[years.index(year)] if year in years else None
        exports_share = (exports / te.get(year)) if (exports is not None
                                                     and te.get(year)) else None
        imports_share = (imports / ti.get(year)) if (imports is not None
                                                     and ti.get(year)) else None

        missing = []
        for m in ag.get("members") or []:
            mkeys = set(_member_keys(m["name"]))
            for a in m.get("aliases") or []:
                mkeys.update(_member_keys(a))
            if not any(mkeys & _item_keys(n) for n in matched_names):
                missing.append(display_name(m["name"]))
        rows.append({
            "id": ag["id"],
            "name": ag["name"],
            "short": ag["short"],
            "kind": ag.get("kind"),
            "blurb": ag.get("blurb", ""),
            "africa": bool(ag.get("africa")),
            "year": year,
            "years": list(years),
            "exports_series": e_series,
            "imports_series": i_series,
            "exports": exports,
            "imports": imports,
            "exports_share": exports_share,
            "imports_share": imports_share,
            "balance": (exports - imports) if (exports is not None
                                               and imports is not None) else None,
            "partners": len(matched_names),
            "top_exports": _top_partners(exp_members, year, years),
            "top_imports": _top_partners(imp_members, year, years),
            "missing": missing,
        })

    def _tx_total(total_map, year):
        v = total_map.get(year)
        return v if v is not None else None

    exports_total = _tx_total(te, year)
    imports_total = _tx_total(ti, year)

    # Partners carrying value yet not covered by any listed agreement.
    uncovered = [it["name"] for it in exp_items + imp_items
                 if it["name"] not in covered_export_names
                 and it["name"] not in covered_import_names
                 and _key(it["name"]) != _key("Kenya")]
    seen = set()
    unmatched = []
    for n in uncovered:
        key = _key(n)
        if key not in seen:
            seen.add(key)
            unmatched.append(display_name(n))

    return {
        "year": year,
        "years": years,
        "unit": _UNIT,
        "rows": rows,
        "exports_total": exports_total,
        "imports_total": imports_total,
        "exports_matched": covered_exports,
        "imports_matched": covered_imports,
        "exports_share": (covered_exports / exports_total)
        if exports_total else None,
        "imports_share": (covered_imports / imports_total)
        if imports_total else None,
        "unmatched": unmatched,
    }


def build(exp, imp, agreements=None, year=None):
    """Convenience entry point used by generate_report."""
    return analyze(exp, imp, agreements or load_agreements(), year)