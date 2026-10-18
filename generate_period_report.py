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
        return {"name": label, "months": months, "totals": totals,
                "change": float(change) if isinstance(change, (int, float))
                          else None,
                "pct": float(pct) if isinstance(pct, (int, float)) else None}

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

    def hcell(r, c, text):
        b._cell_text(table.cell(r, c), text, bold=True,
                     align=WD_ALIGN_PARAGRAPH.CENTER)

    y0, y1 = parsed.get("years", ("", ""))
    hcell(0, 0, "Rank\nin\n%s" % y1)
    table.cell(0, 0).merge(table.cell(1, 0))
    hcell(0, 1, label_header)
    table.cell(0, 1).merge(table.cell(1, 1))
    c0 = 2
    c1 = c0 + n + 1
    t0, t1 = c0 + n, c1 + n
    chg = t1 + 1
    for base, yr in ((c0, y0), (c1, y1)):
        for k, s in enumerate(subs):
            hcell(0, base + k, s)
        hcell(0, t0 if base == c0 else t1,
              ("Total %s %s" % (span, yr)) if span else ("Total %s" % yr))
        b._cell_text(table.cell(1, t0 if base == c0 else t1),
                     "Value in Ksh. Billion", bold=True,
                     align=WD_ALIGN_PARAGRAPH.CENTER)
        u = table.cell(1, base)
        b._cell_text(u, "Value in Ksh. Billion", bold=True,
                     align=WD_ALIGN_PARAGRAPH.CENTER)
        u.merge(table.cell(1, base + n - 1))
    hcell(0, chg, "Change in %s-%s" % (y0, y1))
    table.cell(0, chg).merge(table.cell(0, chg + 1))
    b._cell_text(table.cell(1, chg), "Value in Ksh. Billion", bold=True,
                 align=WD_ALIGN_PARAGRAPH.CENTER)
    hcell(1, chg + 1, "%")

    ri = 2
    for d in rows:
        bold = d.get("name", "").lower() in ("grand total", "total")
        rank = d.get("rank")
        b._cell_text(table.cell(ri, 0), str(rank) if rank not in (None, "") else "",
                     bold=bold, align=WD_ALIGN_PARAGRAPH.CENTER)
        b._cell_text(table.cell(ri, 1), display_name(d["name"]), bold=bold,
                     wrap=True)
        for side, base in ((0, c0), (1, c1)):
            for k, v in enumerate(d["months"][side]):
                b._cell_text(table.cell(ri, base + k),
                             "{:,.1f}".format(v) if v is not None else "",
                             bold=bold, align=WD_ALIGN_PARAGRAPH.CENTER)
            tot = d["totals"][side]
            b._cell_text(table.cell(ri, base + n),
                         "{:,.1f}".format(tot) if tot is not None else "",
                         bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
        _fill_arrow(table.cell(ri, chg), d.get("change"), "chg")
        _fill_arrow(table.cell(ri, chg + 1), d.get("pct"), "pct")
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
    wb.close()
    if top_n:  # optional row cap; by default every workbook row is shown
        for sheet in (exp_des, exp_pdt, imp_sup, imp_pdt):
            sheet["items"] = sheet["items"][:top_n]

    year = int(meta.get("year", 2026))
    prev = int(meta.get("prev_year", year - 1))
    period = meta.get("period", "semi")
    span = meta.get("span", "")
    span_short = meta.get("span_short", span)
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

    period_word = {"quarter": "Quarter", "semi": "Half-Year",
                   "year": "Year"}.get(period, "Half-Year")
    title = "KENYA EXPORT PERFORMANCE IN {} {} AND {}".format(
        span.upper(), prev, year)

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
    b.add_para(title, size=22, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER,
               space_after=24)
    b.add_para("REPORT", size=24, bold=True,
               align=WD_ALIGN_PARAGRAPH.CENTER, space_after=30)
    for _ in range(3):
        b.add_para("")
    b.add_para("Prepared", size=16, bold=True,
               align=WD_ALIGN_PARAGRAPH.CENTER, space_after=6)
    b.add_para("By", size=16, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER,
               space_after=14)
    b.add_para("KENYA EXPORT PROMOTION", size=20, bold=True, italic=True,
               align=WD_ALIGN_PARAGRAPH.CENTER, space_after=6)
    b.add_para("& BRANDING AGENCY", size=20, bold=True, italic=True,
               align=WD_ALIGN_PARAGRAPH.CENTER, space_after=28)
    b.add_para(month_year or date.today().strftime("%B %Y"), size=16,
               bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    b.page_break()

    # ---- TOC / lists ----
    b.add_toc()
    b.page_break()
    b.add_list_placeholder("List of Tables", "table")
    b.page_break()
    b.add_list_placeholder("List of Figures", "figure")
    b.start_body_section()

    tno = [0]

    def tnext():
        tno[0] += 1
        return tno[0]

    # ---- 1. performance overview ----
    p = b.doc.add_paragraph()
    r = p.add_run("KENYA\u2019S TRADE PERFORMANCE IN {} ({} AND {})".format(
        span_short.upper(), prev, year))
    r.bold = True
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    exp_growth = ((exp_c - exp_p) / exp_p
                  if (exp_c is not None and exp_p) else None)
    imp_growth = ((imp_c - imp_p) / imp_p
                  if (imp_c is not None and imp_p) else None)
    bal_widen = (bal_c - bal_p) if (bal_c is not None and bal_p is not None) \
        else None

    def change_word(growth):
        if growth is None:
            return ""
        if growth > 0:
            return "an increase of {}".format(_pct(growth))
        if growth < 0:
            return "a decline of {}".format(_pct(abs(growth)))
        return "broadly unchanged"

    exp_suffix = change_word(exp_growth)
    if exp_suffix:
        b.add_para(
            "In {} {}, total exports were valued at {} compared to {} in the "
            "same period of {}, {}.".format(
                span, year, _fmt_b(exp_c), _fmt_b(exp_p), prev, exp_suffix))
    else:  # no comparable base: report the level without a change clause
        b.add_para(
            "In {} {}, total exports were valued at {}{}.".format(
                span, year, _fmt_b(exp_c),
                (", against {} in the same period of {}".format(
                    _fmt_b(exp_p), prev)) if exp_p is not None else ""))
    if imp_c is not None:
        if imp_growth is None:
            tail = ("" if imp_p is None else
                    ", against {} in the same period of {}".format(
                        _fmt_b(imp_p), prev))
        elif imp_growth == 0:
            tail = ", broadly unchanged from the same period of {}".format(prev)
        else:
            tail = ", {} on the same period of {}".format(
                change_word(imp_growth), prev)
        b.add_para(
            "Total imports were valued at {} in {} {}{}.".format(
                _fmt_b(imp_c), span, year, tail))
    if bal_c is not None:
        if bal_c < 0:
            bal_phrase = "a trade deficit of {}".format(_fmt_b(abs(bal_c)))
        else:
            bal_phrase = "a trade surplus of {}".format(_fmt_b(bal_c))
        # for a deficit, a *smaller* balance number means a wider gap
        if bal_widen is None or bal_widen == 0:
            move = "broadly unchanged from"
        else:
            wider = (bal_widen < 0) if bal_c < 0 else (bal_widen > 0)
            move = "{} by {} from".format("widening" if wider else "narrowing",
                                          _fmt_b(abs(bal_widen)))
        b.add_para(
            "The merchandise trade balance for the period {} {} recorded {}, "
            "{} the same period of {}.".format(
                span, year, bal_phrase, move, prev))
    b.add_source(src)

    # ---- lead destination markets ----
    b.add_heading("Lead Destination Markets {} ({}-{})".format(
        span_short, prev, year))
    b.add_table_caption("Table {}: Top Destination Markets in {} ({}-{})".format(
        tnext(), span_short, prev, year))
    add_period_table(b, exp_des, "Destination", span_short)
    b.add_source(src)
    decl = [d for d in exp_des["items"]
            if d.get("change") is not None and d["change"] < 0]
    decl.sort(key=lambda d: d["change"])
    if decl:
        names = ", ".join(display_name(d["name"]) for d in decl[:3])
        b.add_para("Exports to {} registered the largest declines over the "
                   "period under review.".format(names))

    b.add_heading("Comparative Perspective on Shares of Lead Export Markets "
                  "{} ({} & {})".format(span_short, prev, year), level=2)
    fig1 = os.path.join(out_dir, "fig_markets_%s.png" % year)
    labels, values = _pie_inputs(exp_des, display_name)
    share_pie(fig1, labels, values,
              "Share of Exports by Markets, {} ({}-{})".format(
                  span_short, prev, year))
    b.add_table_caption("Figure 1: Share of Exports by Markets, {} ({}-{})".format(
        span_short, prev, year))
    b.add_figure(fig1)
    b.add_source(src)

    # ---- export product composition ----
    b.add_heading("Composition of Kenya\u2019s Export Commodities")
    b.add_table_caption("Table {}: Top Export Products in {} {} & {}".format(
        tnext(), span_short, prev, year))
    add_period_table(b, exp_pdt, "Product Label", span_short)
    b.add_source(src)
    growers = [d for d in exp_pdt["items"]
               if d.get("change") is not None and d["change"] > 1.0]
    growers.sort(key=lambda d: -d["change"])
    decliners = [d for d in exp_pdt["items"]
                 if d.get("change") is not None and d["change"] < -1.0]
    decliners.sort(key=lambda d: d["change"])
    if growers or decliners:
        b.add_table_caption(
            "Table {}: Export Products Registering Growth / Decline of Over "
            "Ksh. 1 Billion in {} ({}-{})".format(tnext(), span_short, prev, year))
        add_split_table(
            b, "Higher Performing Products (growth over Ksh. 1 billion)",
            "Weaker Performing Products (decline over Ksh. 1 billion)",
            growers[:6], decliners[:6])
        b.add_source(src)

    b.add_heading("Share of Lead Export Products in {} ({} & {})".format(
        span_short, prev, year), level=2)
    fig2 = os.path.join(out_dir, "fig_products_%s.png" % year)
    labels, values = _pie_inputs(exp_pdt, lambda name: name)
    share_pie(fig2, labels, values,
              "Share of Top Exports by Products, {} ({}-{})".format(
                  span_short, prev, year))
    b.add_table_caption("Figure 2: Share of Top Exports by Products, {} "
                        "({}-{})".format(span_short, prev, year))
    b.add_figure(fig2)
    b.add_source(src)

    # ---- deduction / implication ----
    b.add_heading("Deduction from the {} performance {} {}.".format(
        period_word, span_short, year))
    if exp_growth is None:  # no comparable base: state the level only
        b.add_para(
            "Kenya\u2019s export performance over {} {} had exports valued at "
            "{}{}.".format(
                span, year, _fmt_b(exp_c),
                (" against {} in the corresponding period of {}".format(
                    _fmt_b(exp_p), prev)) if exp_p is not None else ""))
    else:
        direction = ("improved" if exp_growth > 0
                     else ("weakened" if exp_growth < 0 else "held steady"))
        b.add_para(
            "Kenya\u2019s export performance over {} {} {} with exports valued at "
            "{} against {} in the corresponding period of {}.".format(
                span, year, direction, _fmt_b(exp_c), _fmt_b(exp_p), prev))
    if decl:
        dest_gr = sorted((d for d in exp_des["items"]
                          if d.get("change") is not None and d["change"] > 0),
                         key=lambda d: -d["change"])
        decl_names = ", ".join(display_name(d["name"]) for d in decl[:3])
        if dest_gr:  # both sides must be destinations, not products
            b.add_para(
                "The decline was most pronounced in {} while {} drove the "
                "positive side.".format(
                    decl_names,
                    ", ".join(display_name(d["name"]) for d in dest_gr[:3])))
        else:
            b.add_para(
                "The decline was most pronounced in {}.".format(decl_names))
    b.add_heading("Implication for the country", level=2)
    if bal_c is None:
        bal_txt = "trade imbalance"
    elif bal_c < 0 and (bal_widen is None or bal_widen < 0):
        bal_txt = "widening trade deficit"
    elif bal_c < 0:
        bal_txt = "trade deficit"
    else:
        bal_txt = "trade surplus"
    b.add_para(
        "The {} points to the need for sustained export "
        "diversification, deeper penetration of the lead destination markets "
        "and value addition in the leading export commodities identified "
        "above.".format(bal_txt))

    # ---- annexes ----
    b.add_heading("Annex")
    b.add_heading("Annex 1: Kenya\u2019s imports by supplying countries.")
    b.add_table_caption("Table {}: Kenya\u2019s Imports by Supplying "
                        "Countries, {} ({}-{})".format(tnext(), span_short,
                                                       prev, year))
    add_period_table(b, imp_sup, "Source", span_short)
    b.add_source(src)

    b.add_heading("Annex 2: Kenya\u2019s imports by products.")
    b.add_table_caption("Table {}: Kenya\u2019s Imports by Products, {} "
                        "({}-{})".format(tnext(), span_short, prev, year))
    add_period_table(b, imp_pdt, "Product Label", span_short)
    b.add_source(src)

    b.add_heading("References")
    b.add_para(src.split("Source: ", 1)[-1])

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
