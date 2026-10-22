"""
input_validation.py
===================

Pre-flight validation of the raw ITC goods downloads (webapp mode "raw").

``validate_goods_inputs`` runs before ``make_tables.generate_tables`` and
turns silent data problems -- mismatched partner countries, year ranges
that do not overlap, missing or duplicated total rows, non-numeric value
cells -- into actionable job-log lines plus a list of hard problems::

    checks, problems = validate_goods_inputs(excel_dir)

``checks`` is a list of ``(level, message)`` pairs where level is
``"ok"``, ``"warn"`` or ``"error"``; ``log_line`` renders one for the
job log.  A non-empty ``problems`` list means the job must stop: the
webapp raises PipelineError with those messages so /api/status shows a
clean, actionable error instead of a mid-pipeline crash.

Scope: the six goods source files behind Tables 1-6 only.  Ready-made
inputs, services, quarterly and product modes are untouched.
"""

import os
import re

from country_names import display_name
from make_tables import _pick_years, find_source_files, parse_source
from pipeline_errors import PipelineError
from xlsx_compat import convert_to_xlsx
from xlsx_tools import to_float

_ROLES = ("table1", "table2", "table3", "table4", "table5", "table6")
_WORLD_ROLES = ("table1", "table2", "table3", "table4")
_KENYA_ROLES = ("table5", "table6")
_BETA_HEADER = ("reporterCd", "reporterLabel", "partnerCd", "partnerLabel",
                "productCd", "productLabel")
_BETA_YEAR_RE = re.compile(r"^\s*((?:19|20)\d{2})(?:\s*\(([^)]*)\))?\s*$")
_QUANTITY_RE = re.compile(r"\b(quantity|tonnes?|tons?|units?|mixed)\b",
                          re.IGNORECASE)
_PLACEHOLDERS = {"", "...", "..", "-", "\u2013", "\u2014", "n/a"}
_SYMBOLS = {"ok": "\u2713", "warn": "\u26a0", "error": "\u2717"}


def log_line(level, message):
    """Render one check as a job-log line (ok / warning / failure glyph)."""
    flat = " ".join(str(message).split())
    return f"{_SYMBOLS.get(level, '?')} {flat}"


def validate_goods_inputs(excel_dir):
    """Validate the six raw ITC downloads in *excel_dir*.

    Returns ``(checks, problems)`` -- see the module docstring.  Never
    raises: file-discovery failures come back as error entries too.
    """
    checks = []
    problems = []

    def ok(msg):
        checks.append(("ok", msg))

    def warn(msg):
        checks.append(("warn", msg))

    def fail(msg):
        checks.append(("error", msg))
        problems.append(msg)

    try:
        files = find_source_files(excel_dir)
    except PipelineError as exc:
        text = str(exc)
        checks.append(("error", text.splitlines()[0]))
        problems.append(text)
        return checks, problems
    ok("Source files located: " + ", ".join(
        f"{r}={os.path.basename(files[r])}" for r in _ROLES))

    insp = {}
    for role in _ROLES:
        try:
            insp[role] = _inspect(files[role])
        except Exception as exc:                      # noqa: BLE001
            fail(f"{role}: cannot read {os.path.basename(files[role])}: {exc}")

    complete = len(insp) == len(_ROLES)
    _check_columns(insp, fail, ok, complete)
    _check_units(insp, warn, ok)
    _check_values(insp, fail, ok, complete)
    _check_totals(insp, fail, ok, complete)

    if complete:
        _check_labels(insp, fail, ok)
        _check_years(insp, fail, warn, ok)
    else:
        warn("Cross-file checks (partner country, year overlap) skipped: "
             "not all six source files could be read")

    return checks, problems


# ---------------------------------------------------------------------------
# Per-file inspection
# ---------------------------------------------------------------------------
def _inspect(path):
    """Read one source workbook: header/unit facts plus the parsed rows."""
    wb, _used = convert_to_xlsx(path)
    ws = wb[wb.sheetnames[0]]
    grid = [list(row) for row in ws.iter_rows(values_only=True)]
    wb.close()

    header = [str(c).strip() if c is not None else ""
              for c in (grid[0] if grid else [])]
    info = {
        "file": os.path.basename(str(path)),
        "beta": bool(header) and header[0] == "reporterCd",
        "header": header,
        "units": _header_units(header),
        "head_text": " ".join(str(c) for row in grid[:6] for c in row
                              if c is not None),
    }
    rows, ycols, years, labels = parse_source(path, _grid=grid)
    info["rows"] = rows
    info["ycols"] = ycols
    info["years"] = years
    info["labels"] = labels
    info["garbage"] = _garbage_cells(rows, ycols)
    return info


def _header_units(header):
    """Unit hints from beta year headers ('2021 (USD Thousand)')."""
    units = []
    for cell in header:
        m = _BETA_YEAR_RE.match(cell)
        if m and m.group(2):
            units.append(m.group(2).strip())
    return units


def _garbage_cells(rows, ycols):
    """Year cells that are neither numbers nor recognised Trade Map blanks."""
    found = []
    for row in rows:
        for i, year in ycols:
            raw = row[i] if i < len(row) else None
            if raw is None:
                continue
            text = str(raw).strip()
            if not text:
                continue
            if to_float(text.lstrip("'")) is not None:
                continue
            if text.casefold() in _PLACEHOLDERS:
                continue
            label = str(row[5] or row[3] or "").strip() or "?"
            found.append((label, year, text))
    return found


def _count_totals(role, rows):
    """Grand-total rows as the extractors see them for *role*."""
    n = 0
    for row in rows:
        code = str(row[4] or "").strip().upper()
        if code != "TOTAL":
            continue
        if role in _KENYA_ROLES or str(row[2] or "").strip() == "000":
            n += 1
    return n


# ---------------------------------------------------------------------------
# Checks
# ---------------------------------------------------------------------------
def _check_columns(insp, fail, ok, complete):
    bad = [r for r, info in insp.items()
           if info["beta"] and tuple(info["header"][:6]) != _BETA_HEADER]
    if bad:
        fail("Unexpected beta header columns in " + ", ".join(sorted(bad))
             + " -- expected reporterCd, reporterLabel, partnerCd, "
             "partnerLabel, productCd, productLabel")
    elif complete and all(i["beta"] for i in insp.values()):
        ok("Header layout: Trade Map beta columns present in all six files")


def _check_units(insp, warn, ok):
    issues = []
    units_seen = set()
    for role in _ROLES:
        info = insp.get(role)
        if info is None:
            continue
        if info["beta"]:
            units = info["units"]
            if not units:
                issues.append(f"{role} (no unit on year headers)")
                continue
            units_seen.update(units)
            if not any("usd" in u.lower() for u in units):
                issues.append(f"{role} ({', '.join(sorted(set(units)))})")
        elif _QUANTITY_RE.search(info["head_text"]):
            issues.append(f"{role} (looks like a quantity download)")
    if issues:
        warn("Year columns are not USD values: " + "; ".join(issues)
             + " -- download the 'value (USD)' copies")
    elif units_seen:
        ok("Units: " + ", ".join(sorted(units_seen)))
    else:
        ok("Units: value columns (classic Trade Map layout)")


def _check_values(insp, fail, ok, complete):
    hit = False
    for role in _ROLES:
        bad = insp.get(role, {}).get("garbage") or []
        if not bad:
            continue
        hit = True
        label, year, text = bad[0]
        extra = f" (+{len(bad) - 1} more)" if len(bad) > 1 else ""
        fail(f"{role}: {len(bad)} non-numeric value cell(s), e.g. row "
             f"{label!r} {year} = {text!r}{extra}")
    if not hit and complete:
        ok("All year cells are numeric or recognised blanks")


def _check_totals(insp, fail, ok, complete):
    hit = False
    for role in _ROLES:
        info = insp.get(role)
        if info is None:
            continue
        n = _count_totals(role, info["rows"])
        if n == 0:
            hit = True
            fail(f"{role}: no grand-total row found (expected one "
                 "'All Products' / World row)")
        elif n > 1:
            hit = True
            fail(f"{role}: {n} grand-total rows found -- expected exactly "
                 "one (duplicated or merged download rows?)")
    if not hit and complete:
        ok("Exactly one grand-total row in each of the six files")


def _check_labels(insp, fail, ok):
    def norm(s):
        return display_name(s or "").casefold()

    reps = {r: (insp[r]["labels"] or {}).get("reporter", "")
            for r in _WORLD_ROLES}
    k_reps = {r: (insp[r]["labels"] or {}).get("reporter", "")
              for r in _KENYA_ROLES}
    k_parts = {r: (insp[r]["labels"] or {}).get("partner", "")
               for r in _KENYA_ROLES}

    for role, rep in k_reps.items():
        if norm(rep) != "kenya":
            fail(f"{role}: reporter reads {rep!r} -- expected 'Kenya' "
                 "(wrong download in this slot?)")

    world = {norm(v) for v in reps.values()}
    if len(world) > 1:
        fail("Tables 1-4 disagree on the partner country: "
             + ", ".join(f"{r}={reps[r]!r}" for r in _WORLD_ROLES))
        return

    partners = {norm(v) for v in k_parts.values()}
    if len(partners) > 1:
        fail("Tables 5-6 disagree on the partner: "
             + ", ".join(f"{r}={k_parts[r]!r}" for r in _KENYA_ROLES))
        return

    if partners and world and partners != world:
        fail(f"Tables 5-6 use partner {k_parts['table5']!r} but tables 1-4 "
             f"report {reps['table1']!r} -- mixed country downloads")
        return

    if norm(k_reps["table5"]) == "kenya":
        ok(f"Partner consistent: {reps['table1']} <-> Kenya (tables 1-6)")


def _check_years(insp, fail, warn, ok):
    spans = {r: list(insp[r]["years"]) for r in _ROLES}
    common = set.intersection(*(set(spans[r]) for r in _ROLES))
    if not common:
        fail("Files share no common report years: "
             + ", ".join(f"{r}={_span(spans[r])}" for r in _ROLES))
        return
    full = _pick_years(sorted(common))
    differing = [r for r in _ROLES if set(spans[r]) != set(full)]
    if differing:
        warn("Year ranges differ ("
             + ", ".join(f"{r}={_span(spans[r])}" for r in differing)
             + f"); report window shrinks to {_span(full)}")
    elif len(full) < 4:
        warn(f"Report window has only {len(full)} year(s) ({_span(full)}); "
             "tables read best with at least 4")
    else:
        ok(f"Report window: {_span(full)} ({len(full)} years)")


def _span(years):
    if not years:
        return "none"
    if len(years) == 1:
        return str(years[0])
    return f"{years[0]}-{years[-1]}"
