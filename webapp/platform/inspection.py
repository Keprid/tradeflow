"""Read-only previews and data-quality summaries for completed jobs."""

import os
import re
from pathlib import Path

import openpyxl

_YEAR_RE = re.compile(r"^\s*((?:19|20)\d{2})\b")


def _cell(value):
    if value is None or isinstance(value, (str, int, float, bool)):
        return value
    if hasattr(value, "isoformat"):
        try:
            return value.isoformat()
        except (TypeError, ValueError):
            return str(value)
    return str(value)


def _year_columns(header):
    out = []
    for index, value in enumerate(header):
        match = _YEAR_RE.match(str(value or ""))
        if match:
            out.append((index, int(match.group(1)), str(value)))
    return out


def preview_workbooks(directory, limit_files=8, max_rows=200):
    directory = Path(directory)
    files = []
    if not directory.is_dir():
        return {"files": files}
    paths = sorted(p for p in directory.glob("*.xlsx")
                   if not p.name.startswith("~$"))[:limit_files]
    for path in paths:
        entry = _preview_file(path, max_rows)
        if entry:
            files.append(entry)
    return {"files": files}


def _preview_file(path, max_rows):
    try:
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    except Exception:  # noqa: BLE001 - preview is best-effort
        return {"name": path.name, "error": "Could not open workbook."}
    try:
        ws = wb[wb.sheetnames[0]]
        rows = []
        for row in ws.iter_rows(values_only=True):
            values = [_cell(v) for v in row]
            if any(v not in (None, "") for v in values):
                rows.append(values)
            if len(rows) >= max_rows:
                break
        header = rows[0] if rows else []
        body = rows[1:] if len(rows) > 1 else []
        year_cols = _year_columns(header)
        series = []
        if year_cols:
            totals = {year: 0.0 for _, year, _ in year_cols}
            counts = {year: 0 for _, year, _ in year_cols}
            for row in body:
                for index, year, _ in year_cols:
                    if index >= len(row):
                        continue
                    value = row[index]
                    if isinstance(value, (int, float)):
                        totals[year] += float(value)
                        counts[year] += 1
            series.append({
                "label": ws.title,
                "years": [year for _, year, _ in year_cols],
                "values": [round(totals[year]) for _, year, _ in year_cols],
                "counts": [counts[year] for _, year, _ in year_cols],
            })
        return {
            "name": path.name,
            "sheets": list(wb.sheetnames),
            "active": ws.title,
            "columns": [str(c) for c in header],
            "rows": body,
            "row_count": len(body),
            "truncated": len(body) >= max_rows,
            "years": [year for _, year, _ in year_cols],
            "series": series,
        }
    finally:
        wb.close()


def quality_report(job_dir, manifest):
    job_dir = Path(job_dir)
    source_dir = job_dir / "uploads"
    report = {
        "generated_at": manifest.get("finished_at"),
        "mode": manifest.get("mode"),
        "report_type": manifest.get("report_type"),
        "provenance": manifest.get("provenance") or {},
        "checks": [],
        "problems": [],
        "files": [],
    }
    provenance = report["provenance"]
    for item in provenance.get("files", []) if isinstance(provenance, dict) \
            else []:
        report["files"].append({
            "name": item.get("name"),
            "bytes": item.get("bytes"),
            "sha256": item.get("sha256"),
            "source": item.get("source"),
        })
    if manifest.get("mode") == "raw" and source_dir.is_dir():
        try:
            import input_validation
            checks, problems = input_validation.validate_goods_inputs(
                str(source_dir))
            report["checks"] = [{"level": level, "message": message}
                                for level, message in checks]
            report["problems"] = list(problems)
        except Exception as exc:  # noqa: BLE001
            report["checks"] = [{"level": "warn",
                                 "message": f"Deep validation skipped: {exc}"}]
    else:
        report["checks"] = [{
            "level": "ok",
            "message": "Input validation is only defined for raw goods "
                       "uploads; this mode is checked by its own pipeline."}]
    report["summary"] = {
        "checks": len(report["checks"]),
        "warnings": sum(1 for c in report["checks"]
                        if c["level"] == "warn"),
        "errors": sum(1 for c in report["checks"]
                      if c["level"] == "error"),
        "files": len(report["files"]),
    }
    return report
