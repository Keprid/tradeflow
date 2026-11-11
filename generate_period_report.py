#!/usr/bin/env python3
"""Generate the KEPROBA period-performance Word report (quarterly,
semi-annual or annual) from the workbook built by ``make_period_tables.py``.

Structure follows the approved sample report (headings are unnumbered, the
way the sample's Heading styles are):

  title page, list of tables, list of figures,
  trade performance in <period> (intro + comparative perspectives)
     - Table: top destinations (sub-periods of both years, change, %)
     - Figure: share of exports by markets
  composition of export commodities
     - Table: top export products + growth/decline exhibit
     - Figure: share of top exports by products
  deduction + implication for the country
  Annex 1: imports by supplying countries
     Annex 2: imports by products
  references

Usage:
    python generate_period_report.py \\
        --tables "output/hy2026/PERFORMANCE TABLES JAN JUN 2026.xlsx" \\
        --output "output/hy2026/KENYA EXPORT PERFORMANCE JANUARY-JUNE 2026.docx"
"""

import argparse
import os
import re
import sys
from datetime import date

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import openpyxl
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt, RGBColor

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE_DIR)

from charts import consolidate, draw_share_pie, slice_callouts, new_fig, finish
from country_names import display_name
from generate_report import ReportBuilder, chart_font  # noqa: E402

SRC_KRA = "Source: KRA Database; Compiled by KEPROBA"
SRC_KNBS = "Source: KNBS; Compiled by KEPROBA"
RED = RGBColor(0xC0, 0x00, 0x00)


# ---------------------------------------------------------------------------
# workbook reading
# ---------------------------------------------------------------------------
def read_meta(wb):
    meta = {}
    if "META" not in wb.sheetnames:
        return meta
    for row in wb["META"].iter_rows(min_row=1, values_only=True):
        if row and row[0]:
            meta[str(row[0])] = row[1]
    return meta


SUB_NAMES = {"Q1", "Q2", "Q3", "Q4",
             "January", "February", "March", "April", "May", "June", "July",
             "August", "September", "October", "November", "December"}


def _is_sub(text):
    return str(text or "").strip() in SUB_NAMES


def read_rank_sheet(wb, name):
    """One summary sheet -> {items, all_other, grand, subs, chg_col}.

    Cell layout (1-based columns): labels | yr-1 sub-periods | yr-1 total |
    yr-2 sub-periods | yr-2 total | change | %.  All value cells are read
    from the cached formula results.
    """
    ws = wb[name]
    hdr = {c: ws.cell(1, c).value for c in range(1, ws.max_column + 1)}
    nlab = next((c for c in sorted(hdr) if _is_sub(hdr[c])), None)
    if nlab is None:
        raise ValueError("no sub-period columns found in sheet %r" % name)
    subs = []
    c = nlab
    while c in hdr and _is_sub(hdr[c]):
        subs.append(str(hdr[c]).strip())
        c += 1
    n = len(subs)
    t0, t1 = nlab + n, nlab + 2 * n + 1
    chg = nlab + 2 * n + 2

    def grab(r):
        label = str(ws.cell(r, nlab - 1).value or "").strip()
        months, totals = [], []
        for base in (nlab, nlab + n + 1):
            sub = []
            for k in range(n):
                v = ws.cell(r, base + k).value
                sub.append(float(v) if isinstance(v, (int, float)) else None)
            months.append(sub)
            v = ws.cell(r, base + n).value
            totals.append(float(v) if isinstance(v, (int, float)) else None)
        change = ws.cell(r, chg).value
        pct = ws.cell(r, chg + 1).value
        # a zero (or missing) previous-year base leaves the % change
        # undefined; the workbook's own formula returns blank there, so
        # honour that even when an older build cached a 0.0 in the cell
        if isinstance(pct, (int, float)) and (totals[0] or 0):
            pct = float(pct)
        else:
            pct = None
        return {"name": label, "months": months, "totals": totals,
                "change": float(change) if isinstance(change, (int, float))
                          else None,
                "pct": pct}

    items, all_other, grand = [], None, None
    for r in range(2, ws.max_row + 1):
        name_v = str(ws.cell(r, nlab - 1).value or "").strip()
        if not name_v:
            continue
        rec = grab(r)
        rank = ws.cell(r, nlab - 2).value if nlab >= 3 else None
        rec["rank"] = rank if isinstance(rank, (int, float)) else None
        low = name_v.lower()
        if low.startswith("all other"):
            all_other = rec
        elif low in ("grand total", "total"):
            grand = rec
        else:
            items.append(rec)
    return {"items": items, "all_other": all_other, "grand": grand,
            "subs": subs}


def read_bot(wb):
    ws = wb["BOT"]
    out = {}
    for r in range(1, ws.max_row + 1):
        label = str(ws.cell(r, 1).value or "").strip()
        vals = [ws.cell(r, c).value for c in range(2, ws.max_column + 1)]
        if label:
            out[label] = [float(v) if isinstance(v, (int, float)) else None
                          for v in vals]
    return out


# ---------------------------------------------------------------------------
# report tables
# ---------------------------------------------------------------------------
def _fill_arrow(cell, value, kind="chg"):
    if value is None:
        if kind == "pct":
            # a percentage off a zero (or missing) base is undefined, not 0.0%
            cell.text = ""
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run("n/a")
            run.font.size = Pt(8)
        else:
            cell.text = ""
        return
    if kind == "chg":
        text = ("▲ {:,.1f}".format(value) if value > 0
                else ("▼ {:,.1f}".format(abs(value)) if value < 0
                      else "— 0.0"))
    else:
        text = ("▲ {:.1%}".format(value) if value > 0
                else ("▼ {:.1%}".format(abs(value)) if value < 0
                      else "— 0.0%"))
    cell.text = ""
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(text)
    run.font.size = Pt(8)
    run.font.color.rgb = (RGBColor(0x00, 0x7A, 0x33) if value > 0
                          else (RED if value < 0 else RGBColor(0, 0, 0)))


def add_period_table(b, parsed, label_header, span=""):
    """Rank | label | sub-periods(y-1) | total | sub-periods(y) | total |
    change | % - mirroring the approved sample."""
    subs = parsed["subs"]
    n = len(subs)
    rows = list(parsed["items"])
    if parsed["all_other"]:
        rows.append(parsed["all_other"])
    if parsed["grand"]:
        rows.append(parsed["grand"])
    ncols = 2 + 2 * n + 4
    table = b.doc.add_table(rows=2 + len(rows), cols=ncols)
    table.style = "Table Grid"
    widths = [900, 2600] + [820] * n + [1000] + [820] * n + [1000, 1150, 750]
    b._set_table_widths(table, widths)

    # Capture the layout grid once for the header writes (table.cell()
    # rebuilds it on every access). All merges run afterwards via live cells.
    hgrid = table._cells

    def hcell(r, c, text):
        b._cell_text(hgrid[r * ncols + c], text, bold=True,
                     align=WD_ALIGN_PARAGRAPH.CENTER)

    y0, y1 = parsed.get("years", ("", ""))
    c0 = 2
    c1 = c0 + n + 1
    t0, t1 = c0 + n, c1 + n
    chg = t1 + 1
    hcell(0, 0, "Rank\nin\n%s" % y1)
    hcell(0, 1, label_header)
    for base, yr in ((c0, y0), (c1, y1)):
        for k, s in enumerate(subs):
            hcell(0, base + k, s)
        hcell(0, t0 if base == c0 else t1,
              ("Total %s %s" % (span, yr)) if span else ("Total %s" % yr))
        b._cell_text(hgrid[ncols + (t0 if base == c0 else t1)],
                     "Value in Kshs Billion", bold=True,
                     align=WD_ALIGN_PARAGRAPH.CENTER)
        b._cell_text(hgrid[ncols + base], "Value in Kshs Billion", bold=True,
                     align=WD_ALIGN_PARAGRAPH.CENTER)
    hcell(0, chg, "Change in %s-%s" % (y0, y1))
    b._cell_text(hgrid[ncols + chg], "Value in Kshs Billion", bold=True,
                 align=WD_ALIGN_PARAGRAPH.CENTER)
    hcell(1, chg + 1, "%")
    table.cell(0, 0).merge(table.cell(1, 0))
    table.cell(0, 1).merge(table.cell(1, 1))
    for base in (c0, c1):
        table.cell(1, base).merge(table.cell(1, base + n - 1))
    table.cell(0, chg).merge(table.cell(0, chg + 1))

    # table.cell() rebuilds the whole layout grid on every access (it is an
    # uncached property), so filling row-by-row is O((rows*cols)^2). Capture
    # the grid once now that the header merges are done and index it flat.
    grid = table._cells

    def cell_at(r, c):
        return grid[r * ncols + c]

    ri = 2
    for d in rows:
        bold = d.get("name", "").lower() in ("grand total", "total")
        rank = d.get("rank")
        b._cell_text(cell_at(ri, 0), str(rank) if rank not in (None, "") else "",
                     bold=bold, align=WD_ALIGN_PARAGRAPH.CENTER)
        name_txt = display_name(d["name"])
        if d["name"].isupper():
            # ranked tables list markets in capitals; keep mapped short
            # names ("Russia") in that style rather than mixing cases
            name_txt = name_txt.upper()
        b._cell_text(cell_at(ri, 1), name_txt, bold=bold, wrap=True)
        for side, base in ((0, c0), (1, c1)):
            for k, v in enumerate(d["months"][side]):
                b._cell_text(cell_at(ri, base + k),
                             "{:,.1f}".format(v) if v is not None else "",
                             bold=bold, align=WD_ALIGN_PARAGRAPH.CENTER)
            tot = d["totals"][side]
            b._cell_text(cell_at(ri, base + n),
                         "{:,.1f}".format(tot) if tot is not None else "",
                         bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
        _fill_arrow(cell_at(ri, chg), d.get("change"), "chg")
        _fill_arrow(cell_at(ri, chg + 1), d.get("pct"), "pct")
        ri += 1
    b._fit_table_on_page(table)
    return table


def add_split_table(b, left_title, right_title, left_rows, right_rows):
    """Two-column exhibit of growers vs decliners (sample's compact table)."""
    n = max(len(left_rows), len(right_rows), 1) + 1
    t = b.doc.add_table(rows=n, cols=2)
    t.style = "Table Grid"
    b._set_table_widths(t, [4800, 4800])
    b._cell_text(t.cell(0, 0), left_title, bold=True,
                 align=WD_ALIGN_PARAGRAPH.CENTER)
    b._cell_text(t.cell(0, 1), right_title, bold=True,
                 align=WD_ALIGN_PARAGRAPH.CENTER)

    def fill(col, rows):
        if not rows:
            b._cell_text(t.cell(1, col), "None", italic=True)
            return
        for i, r in enumerate(rows):
            b._cell_text(t.cell(1 + i, col),
                         "%s (%+.1f)" % (display_name(r["name"]), r["change"]),
                         wrap=True)

    fill(0, left_rows)
    fill(1, right_rows)
    b._fit_table_on_page(t)
    return t


PIE_PALETTE = ["#4472C4", "#ED7D31", "#A5A5A5", "#FFC000", "#5B9BD5",
               "#70AD47", "#C00000", "#7030A0"]


def share_pie(out_path, labels, values, title):
    if not any(v for v in values if v):
        # A period the uploads do not cover has no shares to draw, and
        # matplotlib refuses an all-zero pie - state the gap plainly instead
        # of failing the whole report.
        fig, ax = new_fig(width=7.9, height=4.9)
        ax.set_title(title, fontsize=13, weight="bold", fontfamily=chart_font())
        ax.text(0.5, 0.5, "No data recorded for this period", ha="center",
                va="center", fontsize=14, color="#666666",
                transform=ax.transAxes, style="italic")
        ax.set_axis_off()
        finish(fig, out_path, ax=ax)
        return out_path
    labels, values = consolidate(labels, values, max_slices=7, min_pct=2.5,
                                other_label="All other")
    fig, ax = new_fig(width=7.9, height=4.9)
    colors = [PIE_PALETTE[i % len(PIE_PALETTE)] for i in range(len(labels))]
    _labs, _vals, wedges = draw_share_pie(
        ax, labels, values, colors, style="pie",
        other_label="All other", max_slices=7, min_pct=2.5, pct_inside=False)
    slice_callouts(fig, wedges, _labs, _vals)
    ax.set_title(title, fontsize=13, weight="bold", fontfamily=chart_font())
    finish(fig, out_path, ax=ax)
    return out_path


def _pie_inputs(parsed, map_name):
    """Pie series over the ranked rows plus the workbook's "All other" row,
    so shares are taken against the grand total, not just the listed items."""
    labels = [map_name(d["name"]) for d in parsed["items"]]
    values = [d["totals"][1] or 0 for d in parsed["items"]]
    if parsed["all_other"]:
        labels.append(map_name(parsed["all_other"]["name"]))
        values.append(parsed["all_other"]["totals"][1] or 0)
    return labels, values


# ---------------------------------------------------------------------------
# narrative helpers
# ---------------------------------------------------------------------------
_KEEP_CAPS = {"USA", "UAE", "UK", "US", "EU", "DRC", "USSR", "GCC"}

# Words left lowercase when a capitalised market name runs through the
# narrative ("Democratic Rep of Congo", "Isle of Man").
_MKT_MINOR = {"of", "and", "the", "for", "in", "da", "de"}

# Markets that take a definite article in "to/from X" phrasings.
_THE_MARKETS = {"usa", "uae", "uk", "us", "netherlands"}


def _mkt(name):
    """Narration-friendly market name ('JAPAN' -> 'Japan', 'USA' kept)."""
    disp = display_name(name)
    if disp and disp.isupper():
        words = [w if w.strip(".,") in _KEEP_CAPS else
                 "-".join(p.capitalize() for p in w.split("-"))
                 for w in disp.split()]
        disp = " ".join(words)
    fixed = []
    for i, w in enumerate(disp.split()):
        fixed.append(w.lower() if i and w.lower() in _MKT_MINOR else w)
    return " ".join(fixed)


def _mkt_the(name):
    """Market name with its definite article where one is required
    ('growth to the USA', 'exports to the Netherlands')."""
    disp = _mkt(name)
    return ("the " + disp) if disp.lower() in _THE_MARKETS else disp


def _strip_dangling(text):
    """Drop trailing function words left behind by a label trim."""
    prev = None
    while prev != text:
        prev = text
        text = re.sub(
            r"\s+(and|or|of|the|with|for|in|to|on|at|by|from)$", "",
            text).rstrip(" ,;:")
    return text


def _short(label, maxlen=60):
    """Product label for running text.

    Keeps the label's semicolon-separated clauses intact while they fit;
    only a label that still runs long is trimmed at a comma or word
    boundary.  No ellipsis is ever appended, so sentences stay
    grammatical.
    """
    text = str(label or "").strip()
    if len(text) <= maxlen:
        return text
    parts = text.split(";")
    kept = parts[0].strip()
    for part in parts[1:]:
        cand = kept + "; " + part.strip()
        if len(cand) > maxlen:
            break
        kept = cand
    if len(kept) > maxlen:
        cand = kept[:maxlen]
        cut = None
        if "," in cand:
            comma_cut = cand.rsplit(",", 1)[0].rstrip(" ,;:")
            if len(comma_cut) >= max(20, maxlen // 3):
                cut = comma_cut
        if cut is None:
            cut = cand.rsplit(" ", 1)[0].rstrip(" ,;:")
        kept = _strip_dangling(cut)
    return kept


def _join_and(parts):
    """Serial list: 'A and B' / 'A, B, and C' (empty entries dropped)."""
    parts = [p for p in parts if p]
    if not parts:
        return ""
    if len(parts) == 1:
        return parts[0]
    if len(parts) == 2:
        return " and ".join(parts)
    return ", ".join(parts[:-1]) + ", and " + parts[-1]


def _chg_phrase(prev, cur, y_prev, Y):
    """'an increase of Ksh 2.2 billion (5.0%), from Ksh 44.0 billion in
    2025 to Ksh 46.3 billion in 2026' -- reference-report wording."""
    chg = cur - prev
    word = "an increase" if chg >= 0 else "a decline"
    pct = (" (%+.1f%%)" % (chg / prev * 100)) if prev else ""
    return ("%s of %s%s, from %s in %d to %s in %d"
            % (word, _fmt_b(abs(chg)), pct, _fmt_b(prev), y_prev,
               _fmt_b(cur), Y))


def _driver_clause(drivers, y_prev, Y, maxn=2, flow="exports"):
    """'an increase of Ksh A in exports of X, from Ksh B in <yp> to Ksh C
    in <Y>' clauses.  Each clause carries commas, so three or more are
    joined with semicolons."""
    parts = []
    for prod, prev, cur, chg in drivers[:maxn]:
        word = "an increase" if chg > 0 else "a decline"
        parts.append("%s of %s in %s of %s, from %s in %d to %s in %d"
                     % (word, _fmt_b(abs(chg)), flow, _short(prod).lower(),
                        _fmt_b(prev), y_prev, _fmt_b(cur), Y))
    if len(parts) <= 1:
        return "".join(parts)
    return "; ".join(parts[:-1]) + "; and " + parts[-1]


def _up_list(items):
    """'Name (up 54.7%, from Ksh A to Ksh B)' list for the overview."""
    return _join_and(
        "%s (up %.1f%%, from %s to %s)"
        % (_short(d["name"]), (d["pct"] or 0) * 100,
           _fmt_b(d["totals"][0]), _fmt_b(d["totals"][1]))
        for d in items)


def _interventions():
    """The standing raft of proposed interventions (reference report)."""
    return [
        "**Market Intelligence** \u2013 Set up a trade intelligence unit "
        "and real-time dashboard; commission diagnostic studies on the "
        "declines in strategic markets; strengthen embassy trade "
        "attach\u00e9s.",
        "**Diplomatic Engagement** \u2013 Reactivate trade dialogues with "
        "markets registering declines; align with the EU on compliance "
        "requirements; use AfCFTA and trade missions to reinforce ties "
        "with underperforming markets.",
        "**Diversification** \u2013 Expand into new markets, building on "
        "the gains recorded by the fastest-growing destinations, and "
        "promote higher-value product lines in traditional sectors.",
        "**Trade Facilitation** \u2013 Streamline customs and logistics; "
        "expand export credit/insurance and trade finance to manage "
        "currency risk.",
    ]


def _shares(parsed):
    """[(name, current total, share of grand total)] for one rank sheet."""
    grand = parsed.get("grand") or {}
    gtot = (grand.get("totals") or [None, None])[1]
    out = []
    for d in parsed["items"]:
        v = d["totals"][1] or 0.0
        out.append((d["name"], v, (v / gtot) if gtot else 0.0))
    return out


def _cum_share(parsed, n):
    grand = parsed.get("grand") or {}
    gtot = (grand.get("totals") or [None, None])[1]
    if not gtot:
        return 0.0
    return sum((d["totals"][1] or 0.0) for d in parsed["items"][:n]) / gtot


def _movers(parsed, topn=4, min_abs=1.0):
    rows = [d for d in parsed["items"]
            if d.get("change") is not None and abs(d["change"]) >= min_abs]
    rows.sort(key=lambda d: -abs(d["change"]))
    return rows[:topn]


def _growth_decline(parsed, threshold=1.0):
    growers = [d for d in parsed["items"]
               if d.get("change") is not None and d["change"] > threshold]
    decliners = [d for d in parsed["items"]
                 if d.get("change") is not None and d["change"] < -threshold]
    growers.sort(key=lambda d: -d["change"])
    decliners.sort(key=lambda d: d["change"])
    return growers, decliners


def read_xtab(wb, name, year):
    """Partner x product sheet -> {(partner, product): [prev, cur]}."""
    if name not in wb.sheetnames:
        return {}
    ws = wb[name]
    hdr = {str(ws.cell(1, c).value or "").strip().lower(): c
           for c in range(1, ws.max_column + 1)}
    if not {"year", "partner", "product", "value"} <= set(hdr):
        return {}
    cy, cp, cg, cv = (hdr["year"], hdr["partner"], hdr["product"],
                      hdr["value"])
    out = {}
    for r in range(2, ws.max_row + 1):
        y = ws.cell(r, cy).value
        partner = ws.cell(r, cp).value
        product = ws.cell(r, cg).value
        value = ws.cell(r, cv).value
        if y is None or partner is None or product is None:
            continue
        try:
            yi = int(y)
            val = float(value or 0.0)
        except (TypeError, ValueError):
            continue
        key = (str(partner).strip().upper(), str(product).strip())
        rec = out.setdefault(key, [0.0, 0.0])
        rec[1 if yi == year else 0] += val
    return out


def _drivers(xtab, key, kind, year, y_prev, topn=2, want="decline",
             min_abs=0.05):
    """Biggest product (or market) movements behind one partner (or product).

    kind="partner" -> key is a partner, drivers are products (and vice
    versa).  Returns [(other, prev, cur, change)] sorted by the movement;
    movements below ``min_abs`` billion are dropped so a nil change is
    never narrated as a driver.
    """
    rows = []
    key_u = str(key or "").strip().upper()
    key_raw = str(key or "").strip()
    for (partner, product), (pv, cv) in xtab.items():
        if kind == "partner":
            if partner != key_u:
                continue
            other = product
        else:
            if product != key_raw:
                continue
            other = partner
        rows.append((other, pv, cv, cv - pv))
    if want == "decline":
        rows = [r for r in rows if r[3] < 0]
        rows.sort(key=lambda r: r[3])
    else:
        rows = [r for r in rows if r[3] > 0]
        rows.sort(key=lambda r: -r[3])
    rows = [r for r in rows if abs(r[3]) >= min_abs]
    return rows[:topn]


_ORDINALS = {1: "first", 2: "second", 3: "third", 4: "fourth"}


def _period_ordinal(period, months):
    if not months:
        return "review"
    if period == "quarter":
        return _ORDINALS.get((months[0] - 1) // 3 + 1, "review")
    if period == "semi":
        return "first" if months[0] <= 6 else "second"
    return ""


# ---------------------------------------------------------------------------
# report assembly
# ---------------------------------------------------------------------------
def _fmt_b(v):
    return "Kshs {:,.1f} billion".format(v) if v is not None else "n/a"


def _pct(v):
    return "{:.1f}%".format(v * 100) if v is not None else "n/a"


def build_report(tables_path, out_path, month_year=None, top_n=None):
    wb = openpyxl.load_workbook(tables_path, data_only=True)
    meta = read_meta(wb)
    exp_des = read_rank_sheet(wb, "EXP DES")
    exp_pdt = read_rank_sheet(wb, "EXP PDT")
    imp_sup = read_rank_sheet(wb, "IMP SUP")
    imp_pdt = read_rank_sheet(wb, "IMP PDT")
    bot = read_bot(wb)
    year = int(meta.get("year", 2026))
    xtab_exp = read_xtab(wb, "XTAB EXP", year)
    xtab_imp = read_xtab(wb, "XTAB IMP", year)
    wb.close()
    if top_n:  # optional row cap; by default every workbook row is shown
        for sheet in (exp_des, exp_pdt, imp_sup, imp_pdt):
            sheet["items"] = sheet["items"][:top_n]

    prev = int(meta.get("prev_year", year - 1))
    period = meta.get("period", "semi")
    span = meta.get("span", "")
    span_short = meta.get("span_short", span)
    months = [int(m) for m in str(meta.get("months", "")).split("|")
              if str(m).strip()]
    for sheet in (exp_des, exp_pdt, imp_sup, imp_pdt):
        sheet["years"] = (prev, year)

    def botv(label, idx):
        row = bot.get(label) or []
        return row[idx] if idx < len(row) else None

    # bot columns: Q1-25 Q2-25 HY-25 Q1-26 Q2-26 HY-26 (6 for semi)
    n_sub = len(exp_des["subs"])
    i_prev_tot, i_cur_tot = n_sub, 2 * n_sub + 1
    exp_p, exp_c = botv("Export", i_prev_tot), botv("Export", i_cur_tot)
    imp_p, imp_c = botv("Import", i_prev_tot), botv("Import", i_cur_tot)
    bal_p, bal_c = botv("Bal of Trade", i_prev_tot), botv("Bal of Trade", i_cur_tot)

    comp = bool(exp_p) and bool(exp_c)
    exp_growth = ((exp_c - exp_p) / exp_p) if comp else None
    imp_growth = ((imp_c - imp_p) / imp_p) if (comp and imp_p) else None

    period_word = {"quarter": "Quarter", "semi": "Half-Year",
                   "year": "Year"}.get(period, "Half-Year")
    ordinal = _period_ordinal(period, months)
    period_phrase = {"quarter": "the %s quarter" % ordinal,
                     "semi": "the %s half-year" % ordinal,
                     "year": "the year"}.get(period, "the review period")
    period_cap = period_phrase[0].upper() + period_phrase[1:]

    mkts = _shares(exp_des)
    prods = _shares(exp_pdt)
    imp_mkts = _shares(imp_sup)
    imp_prods = _shares(imp_pdt)
    g_mkt, d_mkt = _growth_decline(exp_des) if comp else ([], [])
    g_prd, d_prd = _growth_decline(exp_pdt) if comp else ([], [])

    # ---- section 1: overview -------------------------------------------
    overview = []
    if comp:
        chg_e = exp_c - exp_p
        overview.append(
            "Kenya\u2019s trade sector recorded %s in %s of %d (%s), with "
            "total merchandise exports valued at %s compared to %s in the "
            "same period of %d, %s of %s (%+.1f%%)."
            % ("notable growth" if exp_growth > 0 else
               ("a contraction" if exp_growth < 0 else "broadly flat trade"),
               period_phrase, year, span, _fmt_b(exp_c), _fmt_b(exp_p),
               prev, "an increase" if chg_e >= 0 else "a decline",
               _fmt_b(abs(chg_e)), exp_growth * 100))
        drive = []
        if g_mkt:
            drive.append("stronger demand from destination countries "
                         "including %s"
                         % _join_and(_mkt(d["name"]) for d in g_mkt[:4]))
        elif mkts:
            drive.append("stronger demand from destination countries "
                         "including %s"
                         % _join_and(_mkt(x[0]) for x in mkts[:4]))
        growers_p = sorted(
            [d for d in exp_pdt["items"]
             if d.get("change") is not None and d["change"] > 0.3
             and d.get("pct") is not None],
            key=lambda d: -d["change"])
        if growers_p:
            drive.append("significant contributions from key sectors such as "
                         "%s, all of which recorded export growth in %d "
                         "compared to %d"
                         % (_up_list(growers_p[:5]), year, prev))
        p2 = ("Export growth during this period was driven by %s. "
              % ", and ".join(drive)) if drive else ""
        if mkts:
            p2 += ("The leading buyers of Kenyan goods were %s (%s), %s "
                   "(%s) and %s (%s) of total exports. "
                   % (_mkt(mkts[0][0]), _pct(mkts[0][2]),
                      _mkt(mkts[1][0]), _pct(mkts[1][2]),
                      _mkt(mkts[2][0]), _pct(mkts[2][2])))
        if prods:
            p2 += ("%s was the top export at %s, accounting for %s of "
                    "total exports. "
                    % (_short(prods[0][0]), _fmt_b(prods[0][1]),
                       _pct(prods[0][2])))
        if imp_growth is not None and exp_growth is not None \
                and imp_growth >= exp_growth:
            p2 += ("However, the pace of import growth outpaced that of "
                   "exports, resulting in a widening trade deficit \u2013 a "
                   "key indicator of the country\u2019s persistent reliance "
                   "on foreign goods.")
        overview.append(p2)
        overview.append(
            "This report provides a comprehensive comparative analysis of "
            "Kenya\u2019s trade performance between %s of %d and %d. It "
            "examines critical trends in export destinations, the "
            "composition of export commodities, and the implications of the "
            "reported performance for the country\u2019s trade balance."
            % (period_phrase, year, prev))
    else:
        overview = [
            "Kenya\u2019s trade sector recorded total merchandise exports of "
            "%s in %s of %d (%s)." % (_fmt_b(exp_c), period_phrase, year,
                                      span),
            ("The leading buyers of Kenyan goods were %s, with %s alone "
             "absorbing %s of total exports."
             % (", ".join(_mkt(x[0]) for x in mkts[:4]),
                _mkt(mkts[0][0]), _pct(mkts[0][2]))) if mkts else "",
            "This report examines critical trends in export destinations and "
            "the composition of export commodities in the review period.",
        ]

    # ---- section 2: comparative perspectives (plain paragraphs) ---------
    if comp:
        chg_e = exp_c - exp_p
        chg_i = imp_c - imp_p
        s2 = [
            "In %s %d, total exports were valued at %s compared to %s in "
            "the same period in %d, an %s of %s (%+.1f%%)."
            % (span_short, year, _fmt_b(exp_c), _fmt_b(exp_p), prev,
               "increase" if chg_e >= 0 else "decline",
               _fmt_b(abs(chg_e)), exp_growth * 100),
            "On the other hand, total imports into Kenya were valued at %s "
            "in %s %d, an %s of %s (%+.1f%%) from %s in the same period in "
            "%d."
            % (_fmt_b(imp_c), span_short, year,
               "increase" if chg_i >= 0 else "decline",
               _fmt_b(abs(chg_i)),
               imp_growth * 100, _fmt_b(imp_p), prev),
        ]
        if bal_c is not None and bal_p is not None:
            wider = abs(bal_c) > abs(bal_p)
            s2.append(
                "The merchandise trade %s %s by %s (%.1f%%), from %s in %d "
                "to %s in %d."
                % ("deficit" if bal_c < 0 else "surplus",
                   "widened" if wider else "narrowed",
                   _fmt_b(abs(abs(bal_c) - abs(bal_p))),
                   (abs(bal_c) / abs(bal_p) - 1) * 100 if bal_p else 0,
                   _fmt_b(abs(bal_p)), prev, _fmt_b(abs(bal_c)), year))
    else:
        s2 = [
            "In the period %s %d, total exports were valued at %s."
            % (span_short, year, _fmt_b(exp_c)),
            "On the other hand, total imports were valued at %s in the same "
            "period." % _fmt_b(imp_c),
            "The balance of trade for merchandise trade for the period "
            "recorded %s."
            % (("a deficit of %s" % _fmt_b(abs(bal_c)))
               if (bal_c is not None and bal_c < 0)
               else (("a surplus of %s" % _fmt_b(bal_c))
                     if bal_c is not None else "n/a")),
        ]

    # ---- section 3: markets narrative + driver bullets -------------------
    markets_note = None
    market_bullets = []
    if comp:
        markets_note = (
            "The last two columns in the above table represent the change "
            "in value and the percentage change, respectively, between the "
            "two years. In %s %d, compared with the same period in %d, the "
            "Kenyan export destination markets that showed significant "
            "change were as follows."
            % (span_short, year, prev))
        for d in d_mkt[:3]:
            drv = _drivers(xtab_exp, d["name"], "partner", year, prev,
                           topn=2, want="decline")
            if drv:
                market_bullets.append(
                    "The decline in exports to %s for the reviewed period "
                    "resulted from %s."
                    % (_mkt_the(d["name"]), _driver_clause(drv, prev, year)))
            else:
                market_bullets.append(
                    "Exports to %s declined by %s (%+.1f%%), from %s in %d "
                    "to %s in %d."
                    % (_mkt_the(d["name"]), _fmt_b(abs(d["change"])),
                       (d.get("pct") or 0) * 100, _fmt_b(d["totals"][0]),
                       prev, _fmt_b(d["totals"][1]), year))
        for d in g_mkt[:2]:
            drv = _drivers(xtab_exp, d["name"], "partner", year, prev,
                           topn=2, want="growth")
            if drv:
                market_bullets.append(
                    "Export growth to %s was driven by %s."
                    % (_mkt_the(d["name"]),
                       _driver_clause(drv, prev, year)))
    else:
        if mkts:
            market_bullets = [
                "%s was the leading destination market, absorbing %s of "
                "exports (%s of the total)."
                % (_mkt(mkts[0][0]), _fmt_b(mkts[0][1]), _pct(mkts[0][2])),
                "The top ten destination markets accounted for %s of total "
                "exports." % _pct(_cum_share(exp_des, 10)),
            ]

    # ---- section 4: products narrative -----------------------------------
    products_note = None
    product_paras = []
    if comp:
        products_note = (
            "Comparing %s %d with the same period in %d, the top export "
            "products that experienced significant change were as follows."
            % (span_short, year, prev))
        for d in d_prd[:2]:
            drv = _drivers(xtab_exp, d["name"], "product", year, prev,
                           topn=3, want="decline")
            if drv:
                product_paras.append(
                    "A %s decline in %s was mainly driven by reduced "
                    "exports to %s."
                    % (_fmt_b(abs(d["change"])), _short(d["name"]),
                       _join_and("%s (down %s, from %s to %s)"
                                 % (_mkt_the(x[0]), _fmt_b(abs(x[3])),
                                    _fmt_b(x[1]), _fmt_b(x[2]))
                                 for x in drv[:3])))
        if g_prd:
            product_paras.append(
                "On the growth side, %s registered the largest gains in the "
                "review period." % _up_list(g_prd[:4]))
    else:
        product_paras = [
            "%s was the leading export product in the review period."
            % _short(prods[0][0]) if prods else "",
            "The top ten export products accounted for %s of total exports."
            % _pct(_cum_share(exp_pdt, 10)),
        ]

    # ---- section 5: deduction --------------------------------------------
    deduction = []
    if comp:
        deduction.append(
            "%s of %d marked %s for Kenya\u2019s merchandise trade. "
            "Total exports %s by %s (%+.1f%%), from %s in %d to %s in %d, "
            "signalling %s demand for Kenyan products in key international "
            "markets."
            % (period_cap, year,
               "a period of robust expansion" if exp_growth > 0
               else "a challenging period",
               "rose" if exp_growth >= 0 else "fell",
               _fmt_b(abs(exp_c - exp_p)), exp_growth * 100,
               _fmt_b(exp_p), prev, _fmt_b(exp_c), year,
               "improved" if exp_growth >= 0 else "weakened"))
        top_cnt = len(exp_des["items"])
        if g_mkt:
            d2 = ("This growth was broad, with %d of the top %d export "
                  "destinations registering increases exceeding Kshs 1 "
                  "billion, led by %s (%s)"
                  % (len(g_mkt), top_cnt,
                     _mkt(g_mkt[0]["name"]), _fmt_b(g_mkt[0]["change"])))
            if len(g_mkt) > 1:
                d2 += (", followed by %s"
                       % _join_and("%s (%s)" % (_mkt(x["name"]),
                                                _fmt_b(x["change"]))
                                   for x in g_mkt[1:4]))
            d2 += "."
        else:
            d2 = ("Export growth over the period was concentrated, with "
                  "none of the top %d export destinations registering "
                  "increases exceeding Kshs 1 billion." % top_cnt)
        if d_mkt:
            d2 += (" However, this positive export performance was partially "
                   "offset by declines in shipments to %s, reflecting "
                   "shifting global demand patterns and competitive "
                   "pressures in traditional markets."
                   % ", ".join("%s (%s)" % (_mkt(x["name"]),
                                            _fmt_b(abs(x["change"])))
                               for x in d_mkt[:3]))
        deduction.append(d2)
        if g_prd or d_prd:
            d3 = ""
            if g_prd:
                lead = g_prd[0]
                d3 = ("On the product performance front, Kenya\u2019s export "
                      "basket demonstrated significant dynamism. The "
                      "standout performer was %s, which recorded an increase "
                      "of %s (%+.1f%%), from %s in %d to %s in %d. "
                      % (_short(lead["name"]), _fmt_b(lead["change"]),
                         (lead.get("pct") or 0) * 100,
                         _fmt_b(lead["totals"][0]), prev,
                         _fmt_b(lead["totals"][1]), year))
                others = g_prd[1:5]
                if others:
                    d3 += ("Traditional mainstays such as %s continued to "
                           "register steady growth. "
                           % _join_and("%s (%s)" % (_short(x["name"]),
                                                    _fmt_b(x["change"]))
                                       for x in others))
            if d_prd:
                d3 += ("The decline in %s was the notable contraction, "
                       "suggesting that Kenya\u2019s export growth was "
                       "largely resilient in the review period."
                       % _join_and("%s (%s)" % (_short(x["name"]),
                                                _fmt_b(abs(x["change"])))
                                   for x in d_prd[:2]))
            if d3:
                deduction.append(d3)
        if bal_c is not None and bal_p is not None:
            wider = abs(bal_c) > abs(bal_p)
            stance = "deficit" if bal_c < 0 else "surplus"
            deduction.append(
                "The combined effect of import growth outpacing export "
                "expansion resulted in a %s trade %s. The merchandise "
                "trade %s %s by %s (%.1f%%), from %s in %d to %s in %d. "
                "This imbalance underscores Kenya\u2019s continued reliance "
                "on imported capital goods, machinery, vehicles and energy "
                "products to support its economic activities, even as its "
                "export sector shows encouraging growth and diversification."
                % ("widened" if wider else "narrowed", stance,
                   stance, "widened" if wider else "narrowed",
                   _fmt_b(abs(abs(bal_c) - abs(bal_p))),
                   (abs(bal_c) / abs(bal_p) - 1) * 100 if bal_p else 0,
                   _fmt_b(abs(bal_p)), prev, _fmt_b(abs(bal_c)), year))
    else:
        deduction = [
            "Kenya\u2019s merchandise trade in %s of %d recorded total "
            "exports of %s against an import bill of %s, yielding a %s of "
            "%s. Trade remains concentrated in a narrow set of markets and "
            "products: the top ten destinations absorbed %s of exports and "
            "the top ten products earned %s of the total."
            % (period_phrase, year, _fmt_b(exp_c), _fmt_b(imp_c),
               "deficit" if (bal_c or 0) < 0 else "surplus",
               _fmt_b(abs(bal_c or 0)), _pct(_cum_share(exp_des, 10)),
               _pct(_cum_share(exp_pdt, 10))),
            "The leading destination markets were %s. On the product front, "
            "%s led export earnings. This concentration in both products and "
            "markets exposes earnings to commodity price volatility, weather "
            "shocks and external demand fluctuations."
            % (", ".join(_mkt(x[0]) for x in mkts[:5]),
                _short(prods[0][0]) if prods else "-"),
        ]

    # ---- implication for the country -------------------------------------
    implication = []
    if comp:
        implication.append(
            "%s of %d demonstrated that Kenya\u2019s export sector "
            "remains resilient and increasingly diversified. The growth in "
            "exports to %s suggests that there are diverse opportunities "
            "in these markets that require targeted market-development "
            "interventions. %s remain among Kenya\u2019s strongest export "
            "performers, confirming the need for ongoing productivity "
            "improvements, compliance with international standards, "
            "climate-resilience interventions and support for value "
            "addition."
            % (period_cap, year,
               _join_and(_mkt_the(x["name"]) for x in g_mkt[:5]) if g_mkt
               else _join_and(_mkt(x[0]) for x in mkts[:5]),
               _join_and(_short(x["name"]) for x in g_prd[:4]) if g_prd
               else _join_and(_short(x[0]) for x in prods[:4])))
        if bal_c is not None and bal_p is not None:
            implication.append(
                "The trade balance moved from %s in %d to %s in %d, with "
                "total imports %s by %.1f%%. This raises concerns about "
                "long-term trade balance sustainability and reinforces the "
                "need for policies that promote domestic production, export "
                "competitiveness and value addition. The country should "
                "prioritise support for emerging export sectors while "
                "sustaining competitiveness in traditional sectors such as "
                "tea, coffee and horticulture."
                % (_fmt_b(abs(bal_p)), prev, _fmt_b(abs(bal_c)), year,
                   "growing" if imp_growth >= 0 else "declining",
                   abs(imp_growth) * 100 if imp_growth is not None else 0))
        if g_prd:
            implication.append(
                "The growth in %s presents a potential opportunity for "
                "export diversification beyond traditional commodities. "
                "Sustaining these gains will require investment in "
                "productive capacity, quality upgrading and market access "
                "for higher-value products."
                % _join_and(_short(x["name"]) for x in g_prd[:3]))
        if d_mkt:
            implication.append(
                "The reported decline in exports to strategic markets "
                "including %s suggests that greater diplomatic and trade "
                "engagement may be required to address market-specific "
                "challenges, including compliance requirements, tariffs and "
                "competition from other suppliers."
                % _join_and(_mkt(x["name"]) for x in d_mkt[:3]))
    else:
        implication = [
            "%s of %d demonstrated that Kenya\u2019s export sector "
            "remains anchored on traditional products and markets. %s and "
            "the wider East African region remain among Kenya\u2019s "
            "strongest export performers, confirming the need for ongoing "
            "productivity improvements, compliance with international "
            "standards and support for value addition."
            % (period_cap, year,
               _mkt(mkts[0][0]) if mkts else "Kenya\u2019s traditional "
               "markets"),
            "The trade balance recorded a %s of %s in the review period, "
            "with total imports of %s against exports of %s. This "
            "reinforces the need for policies that promote domestic "
            "production, export competitiveness and value addition."
            % ("deficit" if (bal_c or 0) < 0 else "surplus",
               _fmt_b(abs(bal_c or 0)), _fmt_b(imp_c), _fmt_b(exp_c)),
            "The concentration of export growth in a limited number of "
            "markets highlights the need for continued efforts to diversify "
            "Kenya\u2019s export destinations and products.",
        ]
    interventions_lead = (
        "In response to the reported declines in exports to strategic "
        "markets%s, and to reinforce overall export competitiveness, a "
        "coordinated raft of interventions is proposed across the following "
        "areas."
        % ((" such as " + ", ".join(_mkt(x["name"]) for x in d_mkt[:2]))
           if d_mkt else ""))

    # ---- annex 1: imports by supplying countries -------------------------
    annex1 = []
    if comp:
        lead = imp_sup["items"][0] if imp_sup["items"] else None
        second = imp_sup["items"][1] if len(imp_sup["items"]) > 1 else None
        for pos, d in enumerate(imp_sup["items"][:2]):
            drv = _drivers(xtab_imp, d["name"], "partner", year, prev,
                           topn=3,
                           want="decline" if d["change"] < 0 else "growth")
            if pos == 0:
                base = ("During the reviewed period, %s remained the "
                        "leading source of imports into Kenya, recording %s"
                        % (_mkt(d["name"]),
                           _chg_phrase(d["totals"][0], d["totals"][1],
                                       prev, year)))
            else:
                base = ("Imports from %s recorded %s"
                        % (_mkt(d["name"]),
                           _chg_phrase(d["totals"][0], d["totals"][1],
                                       prev, year)))
            if drv:
                base += (", due to %s."
                         % _driver_clause(drv, prev, year, maxn=3,
                                          flow="imports"))
            else:
                base += "."
            annex1.append(base)
        annex1.append(
            "Total imports into Kenya %s by %.1f%% (from %s to %s) over the "
            "review period."
            % ("grew" if imp_growth >= 0 else "contracted",
               abs(imp_growth) * 100 if imp_growth is not None else 0,
               _fmt_b(imp_p), _fmt_b(imp_c)))
        rest = [d for d in _movers(imp_sup, topn=12, min_abs=0.5)
                if d is not lead and d is not second]
        for d in rest[:7]:
            drv = _drivers(xtab_imp, d["name"], "partner", year, prev,
                           topn=3,
                           want="decline" if d["change"] < 0 else "growth")
            base = ("Imports from %s recorded %s"
                    % (_mkt(d["name"]),
                       _chg_phrase(d["totals"][0], d["totals"][1],
                                   prev, year)))
            if drv:
                base += (", driven by %s."
                         % _driver_clause(drv, prev, year, maxn=3,
                                          flow="imports"))
            else:
                base += "."
            annex1.append(base)
    else:
        annex1 = [
            "Kenya imported goods worth %s in %s of %d. The leading sources "
            "were %s."
            % (_fmt_b(imp_c), period_phrase, year,
               ", ".join(_mkt(x[0]) for x in imp_mkts[:5])),
            "The top three sources supplied %s of the total import bill."
            % _pct(sum(x[2] for x in imp_mkts[:3])),
        ]

    # ---- annex 2: imports by products ------------------------------------
    annex2 = []
    if comp:
        if imp_pdt["items"]:
            d = imp_pdt["items"][0]
            drv = _drivers(xtab_imp, d["name"], "product", year, prev,
                           topn=3,
                           want="decline" if d["change"] < 0 else "growth")
            txt = ("During the reviewed period, %s remained the top import, "
                   "with imports recording %s"
                   % (_short(d["name"]),
                      _chg_phrase(d["totals"][0], d["totals"][1],
                                  prev, year)))
            if drv:
                txt += (", driven by %s in imports from %s."
                        % ("an increase" if d["change"] >= 0
                           else "a decline",
                           _join_and("%s (%s)" % (_mkt_the(x[0]),
                                                  _fmt_b(abs(x[3])))
                                     for x in drv[:3])))
            else:
                txt += "."
            annex2.append(txt)
        annex2.append(
            "Over the review period, total imports into Kenya %s by %s "
            "(%.1f%%)."
            % ("grew" if (imp_c or 0) >= (imp_p or 0) else "declined",
               _fmt_b(abs((imp_c or 0) - (imp_p or 0))),
               abs(imp_growth) * 100 if imp_growth is not None else 0))
    else:
        annex2 = [
            "%s remained the top import item in the review period."
            % (_short(imp_prods[0][0]) if imp_prods else "-"),
            "Total imports were valued at %s." % _fmt_b(imp_c),
        ]

    out_dir = os.path.dirname(os.path.abspath(out_path))
    os.makedirs(out_dir, exist_ok=True)
    b = ReportBuilder({"country": {"name": "Kenya"}}, {})
    skey = str(meta.get("source", "KNBS"))
    if "KNBS" in skey and "KRA" in skey:
        src = "Source: KNBS and KRA Database; Compiled by KEPROBA"
    elif "KRA" in skey:
        src = SRC_KRA
    else:
        src = SRC_KNBS

    # ---- title page ----
    b.add_letterhead()
    b.add_para("KENYA", size=22, bold=True,
               align=WD_ALIGN_PARAGRAPH.CENTER, space_after=6)
    b.add_para("EXPORT PERFORMANCE IN %s %d-%d" % (span.upper(), prev, year),
               size=20, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER,
               space_after=18)
    b.add_para("REPORT", size=24, bold=True,
               align=WD_ALIGN_PARAGRAPH.CENTER, space_after=30)
    for _ in range(3):
        b.add_para("")
    b.add_para("Prepared by", size=16, bold=True,
               align=WD_ALIGN_PARAGRAPH.CENTER, space_after=14)
    b.add_para("KENYA EXPORT PROMOTION", size=20, bold=True, italic=True,
               align=WD_ALIGN_PARAGRAPH.CENTER, space_after=6)
    b.add_para("& BRANDING AGENCY", size=20, bold=True, italic=True,
               align=WD_ALIGN_PARAGRAPH.CENTER, space_after=28)
    b.add_para(month_year or date.today().strftime("%B %Y"), size=16,
               bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    b.page_break()

    # ---- lists on one page (no table of contents) ----
    cap_t1 = ("Table 1: Top Destination Markets in %s (%d-%d)"
              % (span_short, prev, year))
    cap_t2 = ("Table 2: Top Export Products in %s (%d-%d)"
              % (span_short, prev, year))
    cap_f1 = ("Figure 1: Share of Exports by Market, %s (%d-%d)"
              % (span_short, prev, year))
    cap_f2 = ("Figure 2: Share of Top Exports by Product, %s (%d-%d)"
              % (span_short, prev, year))
    b.add_list_placeholder("List of Tables", "table", [cap_t1, cap_t2])
    b.add_para("")
    b.add_list_placeholder("List of Figures", "figure", [cap_f1, cap_f2])
    b.start_body_section()

    # ---- 1. performance overview ----
    p = b.doc.add_paragraph()
    r = p.add_run("KENYA\u2019S TRADE PERFORMANCE IN %s (%d-%d)"
                  % (span_short.upper(), prev, year))
    b._style_run(r, bold=True)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for para in overview:
        if para:
            b.add_para(para)
    b.page_break()

    # ---- 2. comparative perspectives ----
    b.add_heading("Comparative Perspectives of Kenya\u2019s Trade in "
                  "%s (%d-%d)" % (span_short, prev, year))
    b.add_source(src)
    for line in s2:
        b.add_para(line)
    b.page_break()

    # ---- 3. lead destination markets ----
    b.add_heading("Lead Destination Markets %s (%d-%d)"
                  % (span_short, prev, year))
    b.add_table_caption(cap_t1)
    add_period_table(b, exp_des, "Destination", span_short)
    b.add_source(src)
    if comp and (g_mkt or d_mkt):
        add_split_table(
            b, "Export markets that registered growth (by more than Kshs 1 "
               "billion)",
            "Economies whose imports from Kenya declined (by more than "
            "Kshs 1 billion)",
            g_mkt[:6], d_mkt[:6])
        b.add_source(src)
    if markets_note:
        b.add_note(markets_note)
    for bullet in market_bullets:
        b.add_bullet(bullet)

    b.add_heading("Comparative Perspective on Shares of Lead Export Markets "
                  "%s (%d-%d)" % (span_short, prev, year), level=2)
    fig1 = os.path.join(out_dir, "fig_markets_%s.png" % year)
    labels, values = _pie_inputs(exp_des, display_name)
    share_pie(fig1, labels, values,
              "Share of Exports by Market, %s (%d-%d)"
              % (span_short, prev, year))
    b.add_table_caption(cap_f1)
    b.add_figure(fig1)
    b.add_source(src)
    b.page_break()

    # ---- 4. export product composition ----
    b.add_heading("Composition of Kenya\u2019s Export Commodities")
    b.add_table_caption(cap_t2)
    add_period_table(b, exp_pdt, "Product Label", span_short)
    b.add_source(src)
    if comp and (g_prd or d_prd):
        add_split_table(
            b, "Higher Performing Products (Growth by over Kshs 1 billion)",
            "Weaker Performing Products (Decline by over Kshs 1 billion)",
            g_prd[:6], d_prd[:6])
        b.add_source(src)
    if products_note:
        b.add_note(products_note)
    for para in product_paras:
        if para:
            b.add_para(para)

    b.add_heading("Share of Lead Export Products in %s (%d-%d)"
                  % (span_short, prev, year), level=2)
    fig2 = os.path.join(out_dir, "fig_products_%s.png" % year)
    labels, values = _pie_inputs(exp_pdt, lambda name: name)
    share_pie(fig2, labels, values,
              "Share of Top Exports by Product, %s (%d-%d)"
              % (span_short, prev, year))
    b.add_table_caption(cap_f2)
    b.add_figure(fig2)
    b.add_source(src)
    b.page_break()

    # ---- 5. deduction / implication ----
    b.add_heading("Deductions from the %s\u2019s performance in %s %d."
                  % (period_word.lower(), span_short, year))
    for para in deduction:
        if para:
            b.add_para(para)
    b.add_heading("Implications for the country", level=2)
    for para in implication:
        if para:
            b.add_para(para)
    b.add_para(interventions_lead)
    for bullet in _interventions():
        b.add_bullet(bullet)
    b.page_break()

    # ---- 6. annexes ----
    b.add_heading("Annex")
    b.add_heading("Annex 1: Kenya\u2019s imports by source country.",
                  level=1)
    add_period_table(b, imp_sup, "Source", span_short)
    b.add_source(src)
    for para in annex1:
        b.add_bullet(para)
    b.page_break()

    b.add_heading("Annex 2: Kenya\u2019s imports by product.", level=1)
    add_period_table(b, imp_pdt, "Product Label", span_short)
    b.add_source(src)
    for para in annex2:
        b.add_bullet(para)

    b.add_footer()
    b.doc.save(out_path)
    print("Saved:", out_path)
    return out_path


def main():
    ap = argparse.ArgumentParser(
        description="Generate the quarterly/semi-annual/annual performance "
                    "Word report from make_period_tables.py output.")
    ap.add_argument("--tables", required=True)
    ap.add_argument("--output", required=True)
    ap.add_argument("--month-year", default=None)
    ap.add_argument("--top", type=int, default=None,
                    help="show at most N rows per ranked table "
                         "(default: every row the workbook carries)")
    a = ap.parse_args()
    build_report(a.tables, a.output, month_year=a.month_year, top_n=a.top)


if __name__ == "__main__":
    main()
