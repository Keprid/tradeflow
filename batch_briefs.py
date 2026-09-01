#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Trade Brief / Desktop Research Batch Generator
==============================================

General, market-agnostic runner for the KEPROBA RID desktop-research
deliverables (FY 2026/27 workplan).  It is *data-driven*: every deliverable
type is a reusable template applied to one or more markets, so the same code
works for any present or future market without editing the module.

Deliverable types supported (the 10 desktop-research tasks in the plan):
  * trade_brief        - full Kenya-<Market> trade-flow report (.docx),
                          includes the Africa & Kenya destination insight
                          (Task 2 / 3: 12 trade briefs + 4 regional briefs)
  * regional_brief     - one combined brief for a group of markets
                          (regional focus briefs)
  * secondary_research - compact market fact-sheet for quick-turnaround
                          desk research (e.g. Egypt / Nigeria)
  * performance_brief  - quarterly/annual export-performance summary
                          (quarterly + annual export & nation-brand briefs)
  * market_intelligence - value-chain analysis for a prioritized value chain
  * tariff_ntb         - tariff / non-tariff barrier documentation for a
                          value chain (desk scan)
  * prefeasibility     - market-gravity ranking across markets to justify
                          warehouse / facility siting

The module REUSES the existing single-country pipeline from
``generate_report.py`` and only adds iteration + lightweight aggregation.
It makes NO change to the single-country behaviour -- running
``generate_report.py`` alone still works exactly as before (status quo
preserved).

CLI
---
    python3 batch_briefs.py --manifest manifest.json --out-dir output/briefs

Manifest JSON::

    {
      "report": {"year": 2025, "month_year": "AUGUST 2026"},
      "hs_level": 4,
      "markets": [                        # one entry per market (Trade Map
        {"name": "Saudi Arabia",          #  folder of 7 ITC tables)
         "excel_dir": "sample_data",
         "config": "config/saudi_arabia.json",
         "output": "KENYA-Saudi Arabia TRADE FLOW.docx"}
      ],
      "deliverables": [
        {"type": "trade_brief", "markets": ["Saudi Arabia"]},
        {"type": "regional_brief", "region": "EAC",
         "markets": ["Saudi Arabia"]},
        {"type": "secondary_research", "markets": ["Egypt", "Nigeria"]},
        {"type": "performance_brief", "markets": ["Saudi Arabia"]},
        {"type": "market_intelligence", "name": "Horticulture",
         "hs_regex": "^(06|07|08)", "markets": ["Saudi Arabia"]},
        {"type": "tariff_ntb", "name": "Horticulture",
         "hs_regex": "^(06|07|08)"},
        {"type": "prefeasibility", "markets": ["Egypt", "Nigeria", "Serbia"]}
      ]
    }

``markets`` is the source-of-truth list of available markets; every
deliverable references markets by name.  ``config`` per market is optional:
a minimal config is auto-generated from the name when absent.

``hs_level`` (default 4): all automatic ITC Trade Map pulls MUST request
4-digit (HS4) Harmonized System detail, not the default coarser level.
``generate_report`` emits a warning if an HS2-level table is fed in.
"""

import argparse
import copy
import json
import os
import re
import sys

from country_names import display_name, title_partner

# Imported lazily so this helper can be inspected/imported without building
# the (heavy) matplotlib/docx stack unless a run is actually requested.
_GR = None


def _gr():
    global _GR
    if _GR is None:
        import generate_report as gr
        _GR = gr
    return _GR


# ---------------------------------------------------------------------------
# Small shared helpers
# ---------------------------------------------------------------------------
def _load_json(path):
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def _slug(name):
    return title_partner(name).lower().replace(" ", "_").replace("-", "_")


def _base_config_for(name, report_block, base=None):
    """Auto-build a minimal config dict for a market not in the manifest."""
    cfg = copy.deepcopy(base or {})
    cfg.setdefault("country", {})
    cfg["country"].setdefault("name", name)
    cfg["country"].setdefault("short", title_partner(name).replace(" ", "").upper())
    cfg["country"].setdefault("title", title_partner(name).upper())
    cfg["country"].setdefault("adjective", name)
    cfg["country"].setdefault("possessive", f"{name}'s")
    cfg["country"].setdefault("capital", "[RESEARCH NEEDED: capital]")
    cfg.setdefault("report", {}).update(copy.deepcopy(report_block or {}))
    cfg.setdefault("quick_facts", [])
    cfg.setdefault("references", ["Source: International Trade Centre (ITC) Database; "
                                  "Compiled by KEPROBA"])
    return cfg


def _resolve_excel_dir(excel_dir):
    gr = _gr()
    p = os.path.abspath(excel_dir)
    if os.path.isdir(p):
        return p
    base = os.path.abspath(os.path.join(gr.BASE_DIR, excel_dir))
    return base if os.path.isdir(base) else p


def _market_cfg(market, default_report, out_dir):
    """Return (cfg, cfg_abs) for a market entry, loading a stored config or
    auto-generating one.  Keeps the caller's report block authoritative."""
    gr = _gr()
    name = market["name"]
    cfg_path = market.get("config")
    if cfg_path:
        cfg_abs = os.path.abspath(cfg_path)
        cfg = gr.load_config(cfg_abs)
    else:
        base_cfg = os.path.join(gr.BASE_DIR, "config", _slug(name) + ".json")
        cfg = _base_config_for(name, default_report,
                               base=_load_json(base_cfg) if os.path.exists(base_cfg) else None)
        cfg_abs = os.path.join(out_dir, f"_auto_{_slug(name)}.json")
    if default_report:
        cfg.setdefault("report", {})
        cfg["report"].update(copy.deepcopy(default_report))
    return cfg, cfg_abs


def _find_market(manifest_markets, name):
    for m in manifest_markets:
        if title_partner(m["name"]).lower() == title_partner(name).lower():
            return m
    return {"name": name, "excel_dir": "sample_data"}


def _load_analysis(market, default_report, out_dir):
    gr = _gr()
    cfg, _ = _market_cfg(market, default_report, out_dir)
    excel_dir = _resolve_excel_dir(market.get("excel_dir", "sample_data"))
    return gr.Analysis(cfg).load(excel_dir), cfg


def _top_import_insight(a, n=20):
    """Text lines for the top-N imported products of a market, each showing
    how much the market imports of that product worldwide, from Africa (where
    the per-product-partner detail is provided) and from Kenya."""
    total = a.imports_2025() or 0
    lines = [f"- Top {n} imported products ({a.year}) -- worldwide vs Africa vs Kenya:"]
    for r in a.top_import_products(n):
        ww = r["worldwide_val_bin"]
        ksh = r["kenya_share"]
        ash = r.get("africa_share")
        seg = f"{_usd(ww) if ww is not None else 'n/a'} imported worldwide"
        if ash is not None:
            seg += f"; {ash*100:.1f}% from Africa"
        if ksh is not None:
            seg += f"; {ksh*100:.3f}% from Kenya (incl. in the Africa share)"
        lines.append(f"  - {display_name(r['name'])} ({r['code']}): {seg}")
    return lines


# ---------------------------------------------------------------------------
# Deliverable generators (each generic across any market)
# ---------------------------------------------------------------------------
def deliverable_trade_brief(market, out_dir, default_report, hs_level):
    """Task 2/3 -- full Kenya-<Market> trade-flow report (.docx)."""
    gr = _gr()
    name = market["name"]
    cfg, cfg_abs = _market_cfg(market, default_report, out_dir)
    excel_dir = _resolve_excel_dir(market.get("excel_dir", "sample_data"))
    out = market.get("output") or f"KENYA-{title_partner(name)} TRADE FLOW.docx"
    out_path = os.path.abspath(os.path.join(out_dir, out))
    gr.cfg_path = cfg_abs
    gr.build_report(cfg, excel_dir, out_path, os.path.join(out_dir, ".tmp"))
    return out_path


def deliverable_regional_brief(market, out_dir, default_report, hs_level, region=None):
    """Regional focus brief -- a Kenya-<Market> trade-flow report framed for a
    regional grouping (EAC / COMESA / AfCFTA).  Reuses the full pipeline; the
    region label is carried into the output name."""
    out_path = deliverable_trade_brief(market, out_dir, default_report, hs_level)
    if region:
        name = market["name"]
        path = os.path.join(os.path.dirname(out_path),
                            f"REGIONAL-{region}-{title_partner(name)} TRADE FLOW.docx")
        if os.path.exists(out_path):
            os.replace(out_path, path)
            return path
    return out_path


def deliverable_secondary_research(market, out_dir, default_report, hs_level):
    """Task 1 -- compact market fact-sheet (desk research, e.g. Egypt/Nigeria)."""
    a, cfg = _load_analysis(market, default_report, out_dir)
    name = market["name"]
    path = os.path.join(out_dir, f"SECONDARY_RESEARCH_{_slug(name)}.md")
    afr = a.africa_export_share()
    kenya_dest = a.kenya_destination_share()
    lines = [
        f"# {name} -- Market Fact-Sheet (desk research)",
        "",
        f"- Data year: {a.year}  |  HS level: {hs_level}",
        f"- Total exports: {_usd(a.exports_2025())}  |  Total imports: {_usd(a.imports_2025())}",
    ]
    if afr is not None:
        lines.append(f"- Share of {name}'s exports destined to Africa: {afr*100:.1f}%")
    if kenya_dest is not None:
        lines.append(f"- Share of {name}'s exports destined to Kenya: {kenya_dest*100:.1f}%")
    lines += [
        "",
        "## Top markets",
        *[f"- {d['name']} ({d['share']*100:.1f}% of total)" for d in a.top(a.table1, 5)
          if d.get("share") is not None],
        "",
        "## Top export products",
        *[f"- {display_name(d['label'])} ({d['share']*100:.1f}%)" for d in a.top(a.table4, 5)
          if d.get("share") is not None],
        "",
        "## Top imported products (worldwide vs Africa vs Kenya)",
        *_top_import_insight(a),
        "",
        "_Auto-generated; edit narrative before release._",
    ]
    _write_text(path, lines)
    return path


def deliverable_performance_brief(market, out_dir, default_report, hs_level):
    """Task 4/5 -- quarterly / annual export-performance summary."""
    a, cfg = _load_analysis(market, default_report, out_dir)
    name = market["name"]
    path = os.path.join(out_dir, f"PERFORMANCE_{_slug(name)}.md")
    ex = a.kenya_exports_years()
    im = a.kenya_imports_years()
    lines = [
        f"# Export Performance Brief -- Kenya to {name}",
        "",
        f"- Data year: {a.year} | HS4",
        f"- Kenya's exports to {name}: {_usd(ex[a.iy])} "
        f"(series {[f'{v:.1f}' if v is not None else 'n/a' for v in ex]})",
        f"- Kenya's imports from {name}: {_usd(im[a.iy])}",
    ]
    g = a.exports_growth_2024_25()
    cagr = a.exports_cagr() if hasattr(a, "exports_cagr") else None
    if g is not None:
        lines.append(f"- Year-on-year growth: {g*100:+.1f}%")
    if cagr is not None:
        lines.append(f"- CAGR over series: {cagr*100:+.1f}%")
    if afr := a.africa_export_share():
        lines.append(f"- {name}'s exports to Africa as a whole: {afr*100:.1f}%")
    if kd := a.kenya_destination_share():
        lines.append(f"- Share of {name}'s exports to Kenya: {kd*100:.1f}%")
    lines += ["", "## Top imported products (worldwide vs Africa vs Kenya)", *_top_import_insight(a)]
    lines += ["", "_Auto-generated; edit narrative before release._"]
    _write_text(path, lines)
    return path


def deliverable_market_intelligence(dl, out_dir, default_report, hs_level):
    """Task 7 -- market-intelligence report on a prioritized value chain."""
    market = _find_market(_CURRENT_MARKETS, dl["markets"][0]) if dl.get("markets") else {"name": ""}
    a, cfg = _load_analysis(market, default_report, out_dir)
    vc = dl.get("name", "Value Chain")
    pat = re.compile(dl.get("hs_regex", ".*"))
    rel = [d for d in a.table5["items"]
           if d.get("code") and pat.match(str(d["code"]).strip("'"))
           or d.get("label") and pat.match(d["label"])]
    path = os.path.join(out_dir, f"INTELLIGENCE_{_slug(vc)}.md")
    lines = [
        f"# Market Intelligence -- {vc} (source market: {market['name']})",
        "",
        f"- Data year: {a.year} | HS4",
        "",
        "## Kenya's exports to the market (value chain products)",
        "",
    ]
    for d in rel[:10]:
        lines.append(f"- {display_name(d['label'])} "
                     f"share {d['share']*100:.1f}%" if d.get("share") is not None
                     else f"- {display_name(d['label'])}")
    lines += [
        "",
        f"- {market['name']}'s total imports: {_usd(a.imports_2025())}",
        f"- Exports to Africa: {(a.africa_export_share() or 0)*100:.1f}%",
        f"- Exports to Kenya: {(a.kenya_destination_share() or 0)*100:.1f}%",
        "",
        "_Competition, tariff and market-access detail to be completed from "
        "ITC Trade Map tariff/export-potential modules (HS4)._",
    ]
    _write_text(path, lines)
    return path


def deliverable_tariff_ntb(dl, out_dir, default_report, hs_level):
    """Task 8/14 -- tariff / non-tariff barrier documentation sheet (desk)."""
    vc = dl.get("name", "Value Chain")
    path = os.path.join(out_dir, f"BARRIERS_{_slug(vc)}.md")
    lines = [
        f"# Tariff & Non-Tariff Barrier Documentation -- {vc}",
        "",
        f"- HS codes (HS4 filter): {dl.get('hs_regex', 'n/a')}",
        "",
        "## Desk scan (ITC Trade Map: Tariffs & Market Access module)",
        "- [ ] MFN applied tariff, Kenya origin",
        "- [ ] Preferential / regional rate (EAC, COMESA, AfCFTA, AGOA)",
        "- [ ] Non-tariff measures (SPS, TBT, licensing, quotas)",
        "- [ ] Documentary / conformity requirements",
        "",
        "_Template -- populate rows per market from the ITC Tariff module (HS4). "
        "This is a reusable scaffold for any value chain._",
    ]
    _write_text(path, lines)
    return path


def deliverable_prefeasibility(dl, out_dir, default_report, hs_level):
    """Task 10 -- warehouse pre-feasibility: rank candidate markets by a
    simple gravity score (import demand x growth x Kenya supply share)."""
    market_names = dl.get("markets", [])
    scoring = []
    for mn in market_names:
        m = _find_market(_CURRENT_MARKETS, mn)
        try:
            a, cfg = _load_analysis(m, default_report, out_dir)
            imports = a.imports_2025()
            growth = a.imports_growth_2024_25()
            kenya_ex = a.kenya_exports_years()[a.iy]
            score = None
            if imports:
                score = (imports or 0) * ((growth or 0) + 1.0)
                if kenya_ex:
                    score += kenya_ex
            scoring.append({"market": mn, "imports_usd": imports,
                            "growth": growth, "kenya_exports_usd": kenya_ex,
                            "score": round(score, 2) if score else None})
        except Exception as exc:
            scoring.append({"market": mn, "error": str(exc)})
    scoring.sort(key=lambda r: r.get("score") if r.get("score") is not None else -1e18,
                 reverse=True)
    path = os.path.join(out_dir, "PREFEASIBILITY_warehouses.md")
    lines = [
        "# Warehouse Pre-Feasibility -- Market Gravity Ranking",
        "",
        "| Rank | Market | Imports (USD B) | Import growth | Kenya exports (USD M) | Gravity score |",
        "|---|---|---|---|---|---|",
    ]
    for i, r in enumerate(scoring, 1):
        if "error" in r:
            lines.append(f"| {i} | {r['market']} | n/a | n/a | n/a | error: {r['error']} |")
            continue
        g = f"{r['growth']*100:+.1f}%" if r.get("growth") is not None else "n/a"
        lines.append(f"| {i} | {r['market']} | {r['imports_usd']:.2f} | {g} | "
                     f"{r['kenya_exports_usd']:.0f} | {r['score']} |")
    lines += [
        "",
        "_Gravity score = market import value x (1 + growth) + Kenya export value; "
        "scores are indicative, not a financial appraisal._",
    ]
    _write_text(path, lines)
    return path


# Deliverable type -> (handler, needs_a_market entry)
DELIVERABLES = {
    "trade_brief": (deliverable_trade_brief, True),
    "regional_brief": (deliverable_regional_brief, True),
    "secondary_research": (deliverable_secondary_research, True),
    "performance_brief": (deliverable_performance_brief, True),
    "market_intelligence": (deliverable_market_intelligence, False),
    "tariff_ntb": (deliverable_tariff_ntb, False),
    "prefeasibility": (deliverable_prefeasibility, False),
}

_CURRENT_MARKETS = []


def _usd(v, unit="billion"):
    if v is None:
        return "n/a"
    if v >= 1000 and unit in {"billion", "million", "thousand"}:
        steps = {"thousand": "million", "million": "billion", "billion": "trillion"}
        return f"{v/1000:,.1f} {steps[unit]}"
    return f"{v:,.1f} {unit}"


def _write_text(path, lines):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))


def run(manifest, out_dir, only=None):
    os.makedirs(out_dir, exist_ok=True)
    data = _load_json(manifest)
    default_report = data.get("report") or {}
    hs_level = int(data.get("hs_level", 4))
    markets = data.get("markets", [])
    global _CURRENT_MARKETS
    _CURRENT_MARKETS = markets
    results = []

    def _matches(mname):
        if not only:
            return True
        return any(t.lstrip("* ").lower() in title_partner(mname).lower()
                   for t in only)

    # 1) deliverables explicitly requested
    for dl in data.get("deliverables", []):
        dtype = dl.get("type")
        if dtype not in DELIVERABLES:
            results.append(("FAIL", dtype, "unknown deliverable type"))
            print(f"[FAIL] {dtype}: unknown deliverable type")
            continue
        handler, needs_market = DELIVERABLES[dtype]
        if needs_market:
            # run once per referenced market, passing the market entry to the
            # generic (market-agnostic) handler
            dl_market_names = dl.get("markets") or [m["name"] for m in markets]
            for mn in dl_market_names:
                market = _find_market(markets, mn)
                if not _matches(mn):
                    continue
                try:
                    kwargs = {}
                    if dtype == "regional_brief":
                        kwargs["region"] = dl.get("region", "Regional")
                    out = handler(market, out_dir, default_report, hs_level, **kwargs)
                    results.append(("ok", f"{dtype}:{mn}", out))
                    print(f"[OK]   {dtype}:{mn:18s} -> {out}")
                except Exception as exc:
                    results.append(("FAIL", f"{dtype}:{mn}", str(exc)))
                    print(f"[FAIL] {dtype}:{mn:18s} {exc}")
        else:
            try:
                out = handler(dl, out_dir, default_report, hs_level)
                results.append(("ok", dl.get("type"), out))
                print(f"[OK]   {dl.get('type'):22s} -> {out}")
            except Exception as exc:
                results.append(("FAIL", dl.get("type"), str(exc)))
                print(f"[FAIL] {dl.get('type'):22s} {exc}")

    # 2) backward-compatible: top-level "markets" -> a trade brief each
    for market in markets:
        if not _matches(market["name"]):
            continue
        try:
            out = deliverable_trade_brief(market, out_dir, default_report, hs_level)
            results.append(("ok-brief", market["name"], out))
            print(f"[OK]   trade brief {market['name']:18s} -> {out}")
        except Exception as exc:
            results.append(("FAIL", market["name"], str(exc)))
            print(f"[FAIL] trade brief {market['name']:18s} {exc}")

    # 3) backward-compatible: top-level "products" -> value-chain summaries
    for prod in data.get("products", []):
        try:
            name = f"{prod['name']} value chain"
            dl = {"type": "market_intelligence", "name": prod["name"],
                  "hs_regex": prod.get("hs_regex", ".*"),
                  "markets": [prod.get("source_market", markets[0]["name"])]
                  if markets else []}
            out = deliverable_market_intelligence(dl, out_dir, default_report, hs_level)
            results.append(("ok-chain", name, out))
            print(f"[OK]   value chain {name:18s} -> {out}")
        except Exception as exc:
            results.append(("FAIL", prod.get("name", "?"), str(exc)))
            print(f"[FAIL] value chain {prod.get('name', '?'):18s} {exc}")

    n_ok = sum(1 for r in results if r[0].startswith("ok"))
    n_fail = sum(1 for r in results if r[0] == "FAIL")
    print(f"\nBatch complete: {n_ok} succeeded, {n_fail} failed of {len(results)}.")
    return results


def main():
    ap = argparse.ArgumentParser(
        description="Batch-generate desktop-research deliverables for multiple markets.")
    ap.add_argument("--manifest", required=True, help="Path to the manifest JSON")
    ap.add_argument("--out-dir", default="output/briefs",
                    help="Output directory for generated deliverables")
    ap.add_argument("--only", default=None,
                    help="Comma-separated partial market names to run (filter)")
    args = ap.parse_args()

    only = {s.strip().lower() for s in args.only.split(",")} if args.only else None
    run(args.manifest, args.out_dir, only=only)


if __name__ == "__main__":
    main()
