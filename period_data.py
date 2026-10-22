"""Monthly trade-data loading and period aggregation for the
quarter / semi-annual / annual performance tables.

Two raw shapes are accepted (files in one folder are merged):

* KNBS monthly data - Year, Month, Partner, Partner.1, HS_C, Value, Flow
  where Flow is "Domestic exports" / "Re-exports" / "Imports".
  **Total exports = Domestic exports + Re-exports** at every level
  (month, quarter, half year, year).  Rows whose Flow is blank or
  unrecognised are dropped and reported (see WARNINGS) -- never
  silently counted as exports.
* KRA customs extracts - "Exports by HS Destination" / "Imports by HS
  Origin" layouts (Year, Month, Destination/Origin Name, HS_C, CIF Value).

Everything is normalised to: Year, Month, Side, Partner, Group, GroupLabel,
Value (Ksh. Billion).
"""

import hashlib
import os
import re
import sys

from pipeline_errors import PipelineError

import openpyxl
import pandas as pd

from hs_chapters import CHAPTER_NAMES

MONTH_NAMES = {1: "January", 2: "February", 3: "March", 4: "April",
               5: "May", 6: "June", 7: "July", 8: "August", 9: "September",
               10: "October", 11: "November", 12: "December"}

# Data-quality warnings raised while parsing the current folder of raw
# extracts (blank/unrecognised Flow rows, ...).  load_monthly() clears and
# refills this list on every call (cache hits restore the saved copy), and
# the webapp copies it into the job log.
WARNINGS = []


def _warn(msg):
    WARNINGS.append(msg)
    print("  [warn] %s" % msg)


def _flow_side(value):
    """Classify one raw Flow cell -> "Exports" / "Imports" / None.

    KNBS Flow values are "Domestic exports" / "Re-exports" / "Imports";
    case, hyphenation and stray whitespace vary between extracts.  Only
    recognised labels count: blank or garbage cells return None so the
    caller drops and reports the row instead of silently inflating exports.
    """
    if value is None:
        return None
    text = str(value).replace("\u00a0", " ").strip().lower()
    if not text or text in ("nan", "none", "n/a", "na", "null",
                            "-", "..", "..."):
        return None
    if text.startswith("import"):
        return "Imports"
    if text.startswith(("domestic export", "re-export", "reexport",
                        "re export", "export")):
        return "Exports"
    return None

def normalise_hs(value):
    """HS code padded to the next standard HS length (2 / 4 / 6 / 8 digits),
    or '' when not parseable.

    Padding stops at the first standard width so a chapter code stays a
    chapter: "01" -> "01" (not "00000001") and "0101" -> "0101".
    """
    if value is None:
        return ""
    s = str(value).strip()
    if s.endswith(".0"):  # float artifacts: 1012100.0
        s = s[:-2]
    if not s or s.lower() in ("nan", "none", "#n/a", "(blank)", ""):
        return ""
    s = re.sub(r"\D", "", s)
    if not s:
        return ""
    for width in (2, 4, 6, 8):
        if len(s) <= width:
            return s.zfill(width)
    return s


def _chapter(hs):
    return hs[:2]


def export_group(hs):
    """(code, label) for an export HS code.

    KEPROBA export product groups (verified against the completed half-year
    workbook): the named groups pull their HS codes out of the plain chapter
    rows, everything else is grouped by HS chapter.
    """
    hs = normalise_hs(hs)
    if not hs:
        return ("", "Other (unclassified)")
    ch = _chapter(hs)
    if hs[:4] in ("0901", "0902"):
        return (hs[:4], "Coffee" if hs[:4] == "0901" else "Tea")
    if ch in ("06", "07", "08"):
        return ("06-08", "Horticulture")
    if ch in ("60", "61", "62", "63"):
        return ("60-63", "Textile and Apparel Accessories")
    if ch == "09":
        return ("0903-0910", "Mate, spices and other")
    return (ch, CHAPTER_NAMES.get(ch, "Chapter %s" % ch))


def import_group(hs):
    """(code, label) for an import HS code (plain HS chapters)."""
    hs = normalise_hs(hs)
    ch = _chapter(hs) if hs else ""
    if not ch or ch not in CHAPTER_NAMES:
        return ("", "Other (unclassified)")
    return (ch, CHAPTER_NAMES[ch])


def _candidate_sheets(path):
    if path.lower().endswith(".csv"):
        return [None]
    try:
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        names = wb.sheetnames
        wb.close()
        return names or [None]
    except Exception:
        pass
    # Legacy binary .xls (or anything openpyxl cannot open): let pandas
    # enumerate the sheets so each one is still read individually.
    try:
        with pd.ExcelFile(path) as xf:
            return list(xf.sheet_names) or [None]
    except Exception:
        return [None]


def _pick(cols, *names):
    for n in names:
        for c in cols:
            if c.lower() == n.lower():
                return c
    return None


def _looks_like_header(row):
    """True when a row names Year and Month: that row is the table header.

    Sheets are deciphered individually, so a title / preamble row above the
    real header does not make the sheet unusable.
    """
    if row is None:
        return False
    low = {str(c).strip().lower() for c in row if c is not None}
    return "year" in low and "month" in low


def _read_xlsx(path, sheet):
    """Fast path: stream a workbook sheet with openpyxl instead of
    ``pd.read_excel``, which is an order of magnitude slower on large
    KNBS extracts (300k+ rows)."""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb[sheet] if sheet in wb.sheetnames else wb[wb.sheetnames[0]]
    it = ws.iter_rows(values_only=True)
    # Find the header row among the first 10 rows (row 1 for clean raw
    # extracts; a later row when the sheet opens with a title).
    header = None
    for _ in range(10):
        try:
            row = next(it)
        except StopIteration:
            break
        if _looks_like_header(row):
            header = row
            break
    if header is None:
        wb.close()
        return None
    cols = [str(c).strip() if c is not None else "" for c in header]
    # Disambiguate duplicate headers the way pandas would ("Partner",
    # "Partner.1", ...) so the rest of the loader sees the same names.
    seen = {}
    for i, c in enumerate(cols):
        if c in seen:
            seen[c] += 1
            cols[i] = "%s.%d" % (c, seen[c])
        else:
            seen[c] = 0
    lower = [c.lower() for c in cols]
    need = [i for i, c in enumerate(lower)
            if c in ("year", "month", "value", "cif value", "fob value",
                     "net value", "flow", "hs_c", "hs code",
                     "commodity code", "destination name", "destination",
                     "destin", "origin name", "origin", "partner",
                     "partner.1", "partner name")]
    if not need or "year" not in lower or "month" not in lower:
        # without Year/Month nothing can be attributed to a period; bail out
        # instead of raising (the caller would fall back to the slow parser).
        wb.close()
        return None
    rows = []
    yi, mi = lower.index("year"), lower.index("month")
    for r in it:
        if r is None or len(r) <= max(need):
            continue
        # A row is only useful if it carries a year and a month.
        if r[yi] is None or r[mi] is None:
            continue
        rows.append([r[i] if i < len(r) else None for i in need])
    wb.close()
    return pd.DataFrame(rows, columns=[cols[i] for i in need])


def _read_one(path, sheet):
    """One raw file/sheet -> normalised DataFrame, or None if unusable."""
    try:
        if path.lower().endswith(".csv"):
            df = pd.read_csv(path)
        else:
            try:
                df = _read_xlsx(path, sheet)
            except Exception:
                df = pd.read_excel(path, sheet_name=0 if sheet is None else sheet)
    except Exception:
        return None
    if isinstance(df, dict):
        # sheet_name=None returns every sheet; keep the first usable one.
        df = next((v for v in df.values() if v is not None and len(v)), None)
    if df is None or not len(df):
        return None
    df.columns = [str(c).strip() for c in df.columns]
    cols = df.columns
    col_year = _pick(cols, "Year")
    col_month = _pick(cols, "Month")
    col_val = _pick(cols, "Value", "CIF Value", "FOB Value", "NET VALUE")
    if col_year is None or col_month is None or col_val is None:
        return None

    col_flow = _pick(cols, "Flow")
    col_hs = _pick(cols, "HS_C", "HS Code", "Commodity Code")
    col_dest = _pick(cols, "Destination Name", "Destination", "Destin")
    col_orig = _pick(cols, "Origin Name", "Origin")
    col_partner = _pick(cols, "Partner.1", "Partner Name", "Partner")

    # Classify (and drop unrecognisable) Flow rows before the output frame
    # is built, so every column below is derived from the same row set.
    sides = None
    if col_flow is not None:
        sides = df[col_flow].map(_flow_side)
        unknown = sides.isna()
        if unknown.any():
            # fillna BEFORE astype: recent pandas keeps NaN as NaN through
            # astype(str), and value_counts() would then hide the blanks.
            sample = (df[col_flow][unknown].fillna("(blank)")
                      .astype(str).str.strip().replace("", "(blank)")
                      .value_counts().head(3))
            detail = ", ".join("%s x%d" % (repr(v), n)
                               for v, n in sample.items())
            _warn("%s%s: dropped %d row(s) with blank/unrecognised Flow "
                  "(%s)" % (os.path.basename(path),
                            "" if sheet is None else "/" + str(sheet),
                            int(unknown.sum()), detail))
            df = df[~unknown]
            sides = sides[~unknown]
        if not len(df):
            return None

    out = pd.DataFrame()
    out["Year"] = pd.to_numeric(df[col_year], errors="coerce")
    out["Month"] = pd.to_numeric(df[col_month], errors="coerce")

    if sides is not None:
        # Total exports = domestic exports + re-exports.
        out["Side"] = sides
    elif col_dest is not None:
        out["Side"] = "Exports"
    elif col_orig is not None:
        out["Side"] = "Imports"
    else:
        return None
    # Provenance for the report's source line: files carrying a Flow column
    # are KNBS monthly extracts, the Destination/Origin layout is KRA.
    out["SourceKind"] = "KNBS" if col_flow is not None else "KRA"

    pcol = col_partner or col_dest or col_orig
    out["Partner"] = (df[pcol].astype(str).str.strip().str.upper()
                      if pcol is not None else "")

    hs = df[col_hs] if col_hs is not None else pd.Series([""] * len(df))
    groups = [export_group(h) if s == "Exports" else import_group(h)
              for h, s in zip(hs, out["Side"])]
    out["Group"] = [g[0] for g in groups]
    out["GroupLabel"] = [g[1] for g in groups]

    out["Value"] = pd.to_numeric(df[col_val], errors="coerce").fillna(0.0) / 1e9
    out = out.dropna(subset=["Year", "Month"])
    out["Year"] = out["Year"].astype(int)
    out["Month"] = out["Month"].astype(int)
    out["SourceFile"] = os.path.basename(path)
    # SourceUnit = file + sheet: dedup treats one sheet as one unit of
    # record (same key on two sheets of one workbook = copied rows).
    out["SourceUnit"] = (os.path.basename(path) if sheet is None
                         else "%s/%s" % (os.path.basename(path), sheet))
    return out[["Year", "Month", "Side", "Partner", "Group", "GroupLabel",
                "Value", "SourceFile", "SourceUnit", "SourceKind"]]


def _dir_fingerprint(data_dir):
    """Cheap, content-addressed key for a folder of raw trade extracts.

    Filenames + sizes + mtimes are enough: raw KNBS/KRA extracts are
    read-only inputs, so a byte-identical file always has the same size and
    mtime.  Hashing the contents (237 MB for a full KNBS workbook) would cost
    as much as the parse we are trying to skip.
    """
    parts = []
    for fname in sorted(os.listdir(data_dir)):
        if fname.startswith(("~$", ".")):  # lock files, dot-directories
            continue
        if not fname.lower().endswith((".xls", ".xlsx", ".xlsm", ".csv")):
            continue
        try:
            st = os.stat(os.path.join(data_dir, fname))
        except OSError:
            continue  # file removed between listdir and stat
        parts.append("%s|%d|%d" % (fname, st.st_size, int(st.st_mtime)))
    return hashlib.sha256("\n".join(parts).encode("utf-8")).hexdigest()


def load_monthly(data_dir):
    """Merge every raw file in ``data_dir`` into one monthly DataFrame.

    The normalised result is cached on disk, keyed on the folder fingerprint.
    Re-running for a different period/quarter reuses the cache instead of
    re-parsing the ~240 MB KNBS workbook (which takes ~80 s).

    Rows identical to a row from a *different* source unit (file/sheet:
    the same extract uploaded twice, or one workbook's raw sheets overlapping
    each other) are counted once; rows repeated inside one sheet are kept,
    since two declarations of equal value are two real transactions.

    Data-quality warnings raised during the parse (blank/unrecognised Flow
    rows dropped) are stashed in ``WARNINGS`` on every call -- cache hits
    restore the copy saved with the cached frame.
    """
    frames = []
    del WARNINGS[:]              # fresh list for this load (cache refills it)
    cache_dir = os.path.join(data_dir, ".cache")
    # v4: caches carry the parse-time WARNINGS too; monthly3- and older
    # pickles (built by the lenient Flow mapper that counted garbage rows
    # as exports) are never read again.
    cache_path = os.path.join(
        cache_dir, "monthly4-%s.pkl" % _dir_fingerprint(data_dir))
    if os.path.exists(cache_path):
        try:
            payload = pd.read_pickle(cache_path)
            out, warns = payload["df"], payload["warnings"]
            WARNINGS.extend(warns)
            for w in warns:
                print("  [warn] %s" % w)
            print("  cached  %-58s %7d rows"
                  % (os.path.basename(data_dir), len(out)))
            return out
        except Exception:
            pass  # corrupt cache: fall through and rebuild
    for fname in sorted(os.listdir(data_dir)):
        if fname.startswith(("~$", ".")):  # lock files, dot-directories
            continue
        if not fname.lower().endswith((".xls", ".xlsx", ".xlsm", ".csv")):
            continue
        path = os.path.join(data_dir, fname)
        for sheet in _candidate_sheets(path):
            df = _read_one(path, sheet)
            if df is not None and len(df):
                frames.append(df)
                print("  loaded %-58s %7d rows"
                      % (os.path.basename(path) + "/" + str(sheet), len(df)))
    if not frames:
        raise PipelineError(
            "[ERROR] No usable monthly trade data found in '%s'.\n"
            "Expected the KNBS monthly workbook and/or the raw KRA extracts "
            "(...Exports by HS Destination... / ...Imports by HS Origin...). "
            "Every sheet of every file is inspected individually; summary "
            "and pivot sheets without Year/Month columns are ignored."
            % data_dir)
    out = pd.concat(frames, ignore_index=True)
    # The same extract uploaded twice - or raw sheets copied over each other
    # inside one workbook - must not be counted twice: when one
    # (year, month, side, partner, group, value) key appears under two
    # different source units, keep the first unit's rows only.  Repeats
    # inside one sheet are real transactions and are kept.
    key = ["Year", "Month", "Side", "Partner", "Group", "Value"]
    # Map each key to its first (alphabetically) source unit.  A row survives
    # only when it came from that winning unit.
    out = out.sort_values(["SourceUnit", "SourceFile"])
    first = out.drop_duplicates(subset=key, keep="first")
    first = first[key + ["SourceUnit"]]
    before = len(out)
    out = out.merge(first, on=key, how="left")
    out = out[out["SourceUnit_x"] == out["SourceUnit_y"]].drop(
        columns=["SourceUnit_y"]).rename(columns={"SourceUnit_x": "SourceUnit"})
    dropped = before - len(out)
    if dropped:
        print("  [info] %d rows repeated across files/sheets counted once"
              % dropped)
    try:
        os.makedirs(cache_dir, exist_ok=True)
        pd.to_pickle({"df": out, "warnings": list(WARNINGS)}, cache_path)
    except Exception:
        pass  # read-only volume: skip caching, still return the result
    return out


def period_spec(period, year, half=1, quarter=1, prev_year=None):
    """Sub-period breakdown + display labels for the review period.

    quarter -> the three months;  semi -> Q1+Q2 or Q3+Q4;  year -> Q1..Q4.

    The comparison is always the *same* sub-period one year earlier (the way
    the sample report compares Apr-Jun 2025 with Apr-Jun 2026).  Pass
    ``prev_year`` to override the default ``year - 1``.
    """
    year = int(year)
    if period == "quarter":
        if not 1 <= int(quarter) <= 4:
            raise PipelineError("[ERROR] --quarter must be 1..4")
        q = int(quarter)
        months = list(range(3 * q - 2, 3 * q + 1))
        subs = [(MONTH_NAMES[m], (m,)) for m in months]
    elif period == "semi":
        if int(half) not in (1, 2):
            raise PipelineError("[ERROR] --half must be 1 or 2")
        h = int(half)
        months = list(range(1, 7)) if h == 1 else list(range(7, 13))
        subs = [("Q%d" % (2 * h - 1), tuple(range(1, 4)) if h == 1 else (7, 8, 9)),
                ("Q%d" % (2 * h), tuple(range(4, 7)) if h == 1 else (10, 11, 12))]
    elif period == "year":
        months = list(range(1, 13))
        subs = [("Q1", (1, 2, 3)), ("Q2", (4, 5, 6)),
                ("Q3", (7, 8, 9)), ("Q4", (10, 11, 12))]
    else:
        raise PipelineError("[ERROR] --period must be quarter, semi or year")
    span = "%s-%s" % (MONTH_NAMES[months[0]], MONTH_NAMES[months[-1]])
    span_short = "%s-%s" % (MONTH_NAMES[months[0]][:3],
                            MONTH_NAMES[months[-1]][:3])
    return {"period": period, "year": year,
            "prev_year": int(prev_year) if prev_year else year - 1,
            "months": months, "subs": subs, "span": span,
            "span_upper": span.upper(), "span_short": span_short}


def data_sheets(df):
    """Compact monthly aggregates the summary cells are calculated from.

    Returns {"DATA PARTNER": DataFrame, "DATA GROUP": DataFrame}.
    """
    p = (df.groupby(["Year", "Month", "Side", "Partner"], dropna=False)["Value"]
           .sum().reset_index())
    g = (df.groupby(["Year", "Month", "Side", "Group", "GroupLabel"],
                    dropna=False)["Value"].sum().reset_index())
    return {"DATA PARTNER": p, "DATA GROUP": g}
