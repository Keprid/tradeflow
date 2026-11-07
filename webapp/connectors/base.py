"""Shared connector plumbing: parameter metadata, HTTP and file helpers."""

import gzip
import io
import json
import os
import time
import urllib.error
import urllib.parse
import urllib.request
from dataclasses import dataclass, field
from pathlib import Path

from webapp.platform import settings

USER_AGENT = "tradeflow-open-data/1.0 (+https://github.com/anomalyco/opencode)"
DEFAULT_TIMEOUT = 60


class ConnectorError(RuntimeError):
    """A user-facing connector failure (bad input or upstream problem)."""

    def __init__(self, message, status_code=400):
        super().__init__(message)
        self.status_code = status_code


@dataclass
class Param:
    id: str
    label: str
    type: str = "text"
    required: bool = False
    default: object = ""
    help: str = ""
    choices: list = field(default_factory=list)

    def as_dict(self):
        return {
            "id": self.id, "label": self.label, "type": self.type,
            "required": self.required, "default": self.default,
            "help": self.help, "choices": list(self.choices),
        }


@dataclass
class FetchResult:
    """Files produced by a connector, ready to stage as job uploads."""

    files: dict          # filename -> bytes
    provenance: dict      # source, urls, retrieved_at, notes, fingerprints
    notes: list = field(default_factory=list)

    def file_list(self):
        return [{"name": name, "bytes": len(data)}
                for name, data in sorted(self.files.items())]


class Connector:
    id = ""
    label = ""
    description = ""
    report_types = ("goods",)
    requires_credentials = False
    experimental = False
    params = ()
    # Reference sources build no report of their own, so they name the reports
    # and tables that read them instead.
    used_by = ()

    def validate(self, values):
        values = {p.id: values.get(p.id, p.default) for p in self.params}
        missing = [p.label for p in self.params
                   if p.required and not str(values.get(p.id) or "").strip()]
        if missing:
            raise ConnectorError("Please provide: " + ", ".join(missing))
        return values

    def fetch(self, values):  # pragma: no cover - abstract
        raise NotImplementedError


def _decode_body(raw, headers):
    encoding = (headers.get("Content-Encoding") or "").lower()
    if "gzip" in encoding:
        try:
            raw = gzip.decompress(raw)
        except OSError:
            pass
    return raw


def http_request(url, method="GET", data=None, headers=None, timeout=None,
                 retries=2, retry_status=(429, 500, 502, 503, 504)):
    """Fetch ``url`` with small retries; returns ``(body, response_headers)``."""
    headers = dict(headers or {})
    headers.setdefault("User-Agent", USER_AGENT)
    headers.setdefault("Accept", "application/json, text/csv, */*")
    if isinstance(data, dict):
        data = urllib.parse.urlencode(data).encode("utf-8")
        headers.setdefault("Content-Type",
                           "application/x-www-form-urlencoded")
    last_error = None
    for attempt in range(retries + 1):
        request = urllib.request.Request(url, data=data, headers=headers,
                                         method=method)
        try:
            with urllib.request.urlopen(
                    request, timeout=timeout or DEFAULT_TIMEOUT) as response:
                return _decode_body(response.read(), response.headers), \
                    dict(response.headers)
        except urllib.error.HTTPError as exc:
            body = exc.read()[:800]
            last_error = ConnectorError(
                f"{url.split('?')[0]} returned HTTP {exc.code}: "
                f"{body.decode('utf-8', 'replace')}", status_code=502)
            if exc.code not in retry_status or attempt >= retries:
                raise last_error from exc
        except (urllib.error.URLError, TimeoutError, OSError) as exc:
            last_error = ConnectorError(
                f"Could not reach {url.split('?')[0]}: {exc}", status_code=502)
            if attempt >= retries:
                raise last_error from exc
        time.sleep(0.8 * (attempt + 1))
    raise last_error


def http_json(url, **kwargs):
    body, _ = http_request(url, **kwargs)
    try:
        return json.loads(body.decode("utf-8"))
    except (UnicodeDecodeError, ValueError) as exc:
        raise ConnectorError(f"{url.split('?')[0]} did not return JSON.",
                             status_code=502) from exc


def http_text(url, **kwargs):
    body, _ = http_request(url, **kwargs)
    return body.decode("utf-8-sig", errors="replace")


def cached_json(url, cache_name, ttl_seconds, **kwargs):
    path = Path(settings.CACHE_DIR) / cache_name
    try:
        if path.exists() and time.time() - path.stat().st_mtime < ttl_seconds:
            return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        pass
    data = http_json(url, **kwargs)
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(json.dumps(data), encoding="utf-8")
    except OSError:
        pass
    return data


def slugify(text, fallback="source"):
    out = "".join(ch if ch.isalnum() else "-" for ch in (text or "").lower())
    out = "-".join(part for part in out.split("-") if part)
    return out or fallback


def as_int(value, default=0):
    try:
        return int(float(str(value).replace(",", "").strip()))
    except (TypeError, ValueError):
        return default


def as_float(value, default=0.0):
    try:
        return float(str(value).replace(",", "").strip())
    except (TypeError, ValueError):
        return default


def zip_files(files):
    buf = io.BytesIO()
    import zipfile
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, data in sorted(files.items()):
            archive.writestr(name, data)
    return buf.getvalue()
