"""Cached UN Comtrade reference lists used by the connectors."""

import re

from .base import ConnectorError, cached_json

PARTNER_URL = ("https://comtradeapi.un.org/files/v1/app/reference/"
               "partnerAreas.json")
HS_URL = "https://comtradeapi.un.org/files/v1/app/reference/HS.json"
REF_TTL = 7 * 24 * 3600


def _rows(payload):
    if isinstance(payload, dict):
        return payload.get("results") or payload.get("data") or []
    return payload or []


def partners():
    payload = cached_json(PARTNER_URL, "comtrade_partner_areas.json", REF_TTL)
    rows = []
    for row in _rows(payload):
        code = row.get("PartnerCode", row.get("id"))
        name = (row.get("PartnerDesc") or row.get("text") or "").strip()
        iso3 = (row.get("PartnerCodeIsoAlpha3") or "").strip().upper()
        if code is None or not name:
            continue
        rows.append({
            "code": str(int(code)) if str(code).isdigit() else str(code),
            "iso3": iso3,
            "name": name,
            "is_group": bool(row.get("isGroup")),
        })
    if not rows:
        raise ConnectorError("UN Comtrade's country reference list was empty.",
                             status_code=502)
    return rows


def _norm(text):
    return re.sub(r"[^a-z0-9]+", " ", (text or "").lower()).strip()


def find_countries(query):
    q = _norm(query)
    if not q:
        return []
    rows = partners()
    exact = [r for r in rows
             if q in (_norm(r["name"]), _norm(r["iso3"]), r["code"])]
    if exact:
        return exact
    prefix = [r for r in rows if _norm(r["name"]).startswith(q)
              or r["iso3"].lower().startswith(q)]
    if prefix:
        return prefix
    return [r for r in rows if q in _norm(r["name"])]


def resolve_country(query, purpose="partner", allow_groups=False):
    query = str(query or "").strip()
    if not query:
        raise ConnectorError(f"Please choose a {purpose} country.")
    matches = find_countries(query)
    if not allow_groups:
        countries = [m for m in matches if not m["is_group"]]
        if countries:
            matches = countries
    if not matches:
        raise ConnectorError(f"Unknown {purpose} country: {query!r}.")
    if len(matches) > 1:
        names = ", ".join(m["name"] for m in matches[:8])
        raise ConnectorError(
            f"{query!r} is ambiguous ({names}); use the full country name "
            f"or its UN M49 code.")
    return matches[0]


def country_by_code(code):
    code = str(code or "").strip()
    for row in partners():
        if row["code"] == code:
            return row
    return None


def country_name(code, default="World"):
    if str(code) in ("0", "000", ""):
        return default
    row = country_by_code(code)
    return row["name"] if row else f"Code {code}"


def hs_labels():
    payload = cached_json(HS_URL, "comtrade_hs.json", REF_TTL,
                          timeout=120)
    labels = {}
    for row in _rows(payload):
        code = str(row.get("id") or "").strip()
        text = str(row.get("text") or "").strip()
        if not code:
            continue
        text = re.sub(r"^%s\s*-\s*" % re.escape(code), "", text)
        labels[code] = text or code
    if not labels:
        raise ConnectorError("UN Comtrade's HS reference list was empty.",
                             status_code=502)
    return labels


def hs_label(code, labels=None):
    code = str(code or "").strip()
    if code.upper() == "TOTAL":
        return "All Products"
    labels = labels if labels is not None else hs_labels()
    return labels.get(code) or labels.get(code.zfill(4)) or f"HS {code}"
