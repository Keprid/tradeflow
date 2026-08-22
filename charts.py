#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
charts.py
=========

Shared pie / doughnut chart engine used by every report generator.

Features (mirrors professional chart-tool capabilities):

* Styles: "pie", "exploded_pie", "donut", "exploded_donut", "3d",
  "3d_exploded". Exploded styles move all slices away from the centre;
  on top of that the largest slice is called out with a bigger offset.
  The "3d" styles add a drop shadow for a pseudo-3D look.
* Small-slice consolidation: slices below ``min_pct`` beyond
  ``max_slices`` are collected into a single "Other" slice, so the pie
  never carries more points than stay readable.
* Overlap-safe labels: percentage callouts are drawn only on slices big
  enough to hold them; the full name + share of every slice lives in the
  legend rendered by :func:`share_legend` below the chart.

Only matplotlib is required.
"""

import math

import matplotlib.pyplot as plt

# Style constants -----------------------------------------------------------
WEDGE_WIDTH = 0.42        # ring thickness for donut styles
EDGE_COLOR = "white"
EDGE_WIDTH = 1.2
BASE_EXPLODE = 0.045      # offset applied to every slice when exploded
CALLOUT_EXTRA = 0.055     # extra offset for the called-out (largest) slice
PCT_LABEL_MIN = 3.5       # hide % text on slices smaller than this (%)
LEGEND_FONTSIZE = 8.5

# 3D rendering (Office-style extruded pie)
ASPECT_3D = 0.55          # vertical squash of the pie ellipse (rotX feel)
DEPTH = 0.18              # total side depth as fraction of the radius
DEPTH_STEPS = 12          # layered copies used to fake the extrusion


def _darken(hex_color, factor):
    """Multiply an '#rrggbb' colour by ``factor`` (0-1) -> darker shade."""
    hex_color = hex_color.lstrip("#")
    r, g, b = (int(hex_color[i:i + 2], 16) for i in (0, 2, 4))
    return "#%02x%02x%02x" % (int(r * factor), int(g * factor), int(b * factor))


def consolidate(labels, values, max_slices=7, min_pct=2.5,
                other_label="Other"):
    """Collect small slices into one ``other_label`` slice.

    Keeps the largest ``max_slices`` slices whose share is at least
    ``min_pct`` percent; everything else is summed into a single slice.
    Returns (labels, values) lists ready to plot.
    """
    total = sum(v for v in values if v is not None)
    if total <= 0:
        return list(labels), [v or 0 for v in values]
    keep, other_v = [], 0.0
    for label, v in zip(labels, values):
        v = v or 0.0
        if len(keep) < max_slices and v / total * 100 >= min_pct and v > 0:
            keep.append((label, v))
        else:
            other_v += v
    if other_v > 0:
        keep.append((other_label, other_v))
    return [l for l, _ in keep], [v for _, v in keep]


def series_shares(items):
    """Percent shares for charting from parsed rank-table ``items``.

    Uses the parsed 'share' column when it carries values; otherwise derives
    the shares from the latest year whose values are present (some datasets
    ship an empty latest-year column, which leaves every share cell blank).
    """
    parsed = [d.get("share") for d in items]
    if any(s is not None for s in parsed):
        return [(s or 0.0) * 100 for s in parsed]
    rows = [d.get("years") or [] for d in items]
    n = max((len(r) for r in rows), default=0)
    for k in range(n - 1, -1, -1):
        vals = [r[k] if k < len(r) and r[k] else 0.0 for r in rows]
        total = sum(vals)
        if total > 0:
            return [v / total * 100 for v in vals]
    return [0.0] * len(items)


def _draw_extruded(ax, values, colors, offsets, donut=False, pct_inside=True):
    """Office-style 3D pie: layered darkened copies build the side wall,
    the unsquashed top face carries the colours and labels. Callers must
    apply the vertical squash via ``ax.set_aspect(ASPECT_3D)`` afterwards.
    """
    n = len(values)
    # bottom-up: deepest layer first so nearer layers paint over it
    for step in range(DEPTH_STEPS, 0, -1):
        f = 0.45 + (0.40 * (DEPTH_STEPS - step) / DEPTH_STEPS)
        layer_colors = [_darken(c, f) for c in colors]
        ax.pie(values, explode=offsets, center=(0, -DEPTH * step / DEPTH_STEPS),
               radius=1.0, startangle=90, counterclock=False,
               colors=layer_colors[:n],
               wedgeprops=dict(width=None, edgecolor=_darken("#FFFFFF", f),
                               linewidth=0.6))
    result = ax.pie(
        values, labels=None, autopct="%1.1f%%" if pct_inside else None,
        startangle=90, counterclock=False, colors=list(colors)[:n],
        explode=offsets,
        wedgeprops=dict(edgecolor=EDGE_COLOR, linewidth=EDGE_WIDTH),
        textprops={"fontsize": 8})
    contents = list(result)
    wedges = contents[0]
    autotexts = contents[2] if len(contents) > 2 else []
    for at in autotexts:
        at.set_fontsize(8)
        at.set_color("white")
        at.set_fontweight("bold")
    return wedges


def draw_share_pie(ax, labels, values, colors, *, style="donut",
                   other_label="Other", max_slices=7, min_pct=2.5,
                   pct_inside=True):
    """Draw a styled pie/doughnut on ``ax``; returns (labels, values, wedges).

    style : "pie" | "exploded_pie" | "donut" | "exploded_donut" |
            "3d" | "3d_exploded"
    pct_inside : draw the white percentage callouts inside the slices; turn
                 off when the shares are carried by outside slice labels
                 (see :func:`slice_callouts`) to avoid printing every share
                 twice.
    """
    labels, values = consolidate(labels, values, max_slices=max_slices,
                                 min_pct=min_pct, other_label=other_label)
    n = len(labels)
    donut = style in ("donut", "exploded_donut")
    three_d = style in ("3d", "3d_exploded")
    shadow = style in ("3d", "3d_exploded")
    explode_all = style in ("exploded_pie", "exploded_donut", "3d_exploded")

    offsets = [0.0] * n
    if n:
        biggest = max(range(n), key=lambda i: values[i])
        for i in range(n):
            if explode_all:
                offsets[i] += BASE_EXPLODE
            # the largest slice is always called out on exploded styles
            if i == biggest and explode_all:
                offsets[i] += CALLOUT_EXTRA

    def pct_fmt(pct):
        return f"{pct:.1f}%" if pct >= PCT_LABEL_MIN else ""

    if three_d:
        palette = list(colors)[:n]
        while len(palette) < n:
            palette.append(palette[len(palette) % max(1, len(colors))] if colors else "#888888")
        wedges = _draw_extruded(ax, values, palette, offsets, pct_inside=pct_inside)
        ax.set_aspect(ASPECT_3D)
        lim = 1.0 + max(offsets or [0]) + DEPTH + 0.05
        ax.set_xlim(-lim, lim)
        ax.set_ylim(-lim * ASPECT_3D * 2, lim)
        return labels, values, wedges

    result = ax.pie(
        values, labels=None,
        autopct=(pct_fmt if pct_inside else None),
        startangle=90, counterclock=False, colors=list(colors)[:n],
        explode=offsets,
        pctdistance=(0.80 if donut else 0.68),
        wedgeprops=dict(width=(WEDGE_WIDTH if donut else None),
                        edgecolor=EDGE_COLOR, linewidth=EDGE_WIDTH),
        textprops={"fontsize": 8},
        shadow=shadow,
    )
    contents = list(result)
    wedges = contents[0]
    autotexts = contents[2] if len(contents) > 2 else []
    for at in autotexts:
        at.set_fontsize(8)
        at.set_color("white")
        at.set_fontweight("bold")
    ax.set_aspect("equal")
    return labels, values, wedges


def share_legend(fig, wedges, labels, values, ncol=2, fontsize=None):
    """Legend below the chart with 'name – share %' entries."""
    fontsize = fontsize or LEGEND_FONTSIZE
    total = sum(values) or 1.0
    entries = [f"{l} – {v / total * 100:.1f}%" for l, v in zip(labels, values)]
    fig.legend(wedges, entries, loc="lower center",
               bbox_to_anchor=(0.5, -0.02), ncol=ncol,
               frameon=False, fontsize=fontsize)


def side_legend(fig, wedges, labels, values, fontsize=None):
    """Legend to the right of the chart (template style: category names).

    Anchored beyond the axes' right edge (axes-relative transform) so the
    legend can never mingle with the slices or their data labels.
    """
    fontsize = fontsize or LEGEND_FONTSIZE
    ax = fig.axes[0] if fig.axes else None
    kwargs = dict(loc="center left", frameon=False, fontsize=fontsize,
                  handlelength=1.2, labelspacing=0.7)
    if ax is not None:
        kwargs.update(bbox_to_anchor=(1.04, 0.5), bbox_transform=ax.transAxes)
    else:
        kwargs.update(bbox_to_anchor=(0.98, 0.5))
    fig.legend(wedges, list(labels), **kwargs)


def slice_callouts(fig, wedges, labels, values, fontsize=None):
    """Place each 'name – share %' label right next to its own slice.

    Replaces the one-sided legend: every slice gets a short radial stub plus
    a straight connector to its label (Excel call-out style), so even slices
    with near-identical colours are unambiguous. The actual drawing happens
    in :func:`_render_slice_callouts` from finish(), once the final layout -
    and therefore the anti-overlap pass - can use settled geometry.
    """
    total = sum(values) or 1.0
    entries = [f"{l} – {v / total * 100:.1f}%" for l, v in zip(labels, values)]
    fig._slice_callouts = list(zip(wedges, entries))


def _render_slice_callouts(fig):
    """Draw the stored slice call-outs with de-overlapped label positions."""
    callouts = getattr(fig, "_slice_callouts", None)
    ax = fig.axes[0] if fig.axes else None
    if not callouts or ax is None:
        return
    fontsize = LEGEND_FONTSIZE
    inv_axes = ax.transAxes.inverted()
    gap = 0.062          # min vertical spacing between labels (axes fraction)
    items = []
    for w, text in callouts:
        mid = math.radians((w.theta1 + w.theta2) / 2.0)
        ux, uy = math.cos(mid), math.sin(mid)
        cx, cy = w.center
        arc = (cx + w.r * ux, cy + w.r * uy)                 # on the rim
        stub = (cx + (w.r + 0.14) * ux, cy + (w.r + 0.14) * uy)
        fx, fy = inv_axes.transform(ax.transData.transform(stub))
        right = ux >= 0
        items.append({"w": w, "text": text, "arc": arc, "stub": stub,
                      "fx": fx, "fy": fy, "right": right})
    # anti-overlap: nudge labels apart within each side, top-anchored order
    for side in (True, False):
        group = sorted((it for it in items if it["right"] == side),
                       key=lambda it: it["fy"])
        prev = None
        for it in group:
            y = it["fy"]
            if prev is not None and y - prev < gap:
                y = prev + gap
            y = min(max(y, 0.03), 1.35)
            it["fy"] = y
            prev = y
    for it in items:
        tx = min(it["fx"] + 0.02, 1.30) if it["right"] else max(it["fx"] - 0.02, -0.55)
        ax.annotate(
            it["text"],
            xy=it["stub"], xycoords="data",
            xytext=(tx, it["fy"]), textcoords="axes fraction",
            ha=("left" if it["right"] else "right"), va="center",
            fontsize=fontsize, color="#1a1a1a", annotation_clip=False,
            arrowprops=dict(arrowstyle="-", color=it["w"].get_facecolor(),
                            lw=0.9, shrinkA=0, shrinkB=1))


def new_fig(width=7.9, height=4.9, dpi=160):
    """Figure sized to give outside labels/legend breathing room."""
    fig, ax = plt.subplots(figsize=(width, height), dpi=dpi)
    return fig, ax


def _shrink_axes_for_legends(fig):
    """Pull the axes in so figure-level legends sit clear of the chart.

    tight_layout() ignores figure-level legends, so without this the pie
    would keep its full width and slide under a side legend.
    """
    try:
        fig.canvas.draw()
    except Exception:
        return
    renderer = fig.canvas.get_renderer()
    fw, fh = fig.bbox.width, fig.bbox.height
    for ax in fig.axes:
        pos = ax.get_position()
        x0, y0, w, h = pos.x0, pos.y0, pos.width, pos.height
        for leg in fig.legends:
            bb = leg.get_window_extent(renderer)
            lx0, lx1 = bb.x0 / fw, bb.x1 / fw
            ly0, ly1 = bb.y0 / fh, bb.y1 / fh
            beside = (ly0 < y0 + h) and (ly1 > y0)
            overlaps_x = (lx0 < x0 + w) and (lx1 > x0)
            if beside and overlaps_x:
                if lx0 >= x0 + w / 2:      # legend sits on the right half
                    w = max(0.28, lx0 - 0.02 - x0)
                else:                      # legend sits on the left half
                    new_x0 = max(0.0, min(x0 + w - 0.28, lx1 + 0.02))
                    w = pos.x0 + pos.width - new_x0
                    x0 = new_x0
        ax.set_position([x0, y0, w, h])


def finish(fig, out_path, ax=None):
    """Tight layout, save, close."""
    if getattr(fig, "_slice_callouts", None):
        fig.canvas.draw()          # settle final positions before labelling
        _render_slice_callouts(fig)
        try:
            fig.tight_layout()
        except Exception:
            pass
    elif fig.legends:
        _shrink_axes_for_legends(fig)
    else:
        fig.tight_layout()
    fig.savefig(out_path, bbox_inches="tight")
    plt.close(fig)
