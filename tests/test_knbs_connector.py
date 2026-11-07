"""The KNBS reference connector and the snapshot install step."""

import csv
import io
import json
import urllib.error

import pytest
from fastapi.testclient import TestClient

import knbs_reference as K
from webapp.connectors import connector_catalog
from webapp.connectors.base import ConnectorError
from webapp.connectors.knbs import (OBSERVATIONS_FILE, SNAPSHOT_FILE, USED_BY,
                                    KnbsConnector)
from webapp.main import app
from webapp.platform import registry, settings

FETCH_ID = "a" * 32


def _snapshot(rows=None):
    """A two-dataset snapshot in the shape knbs_reference.build_snapshot makes."""
    return {
        "generated_at": "2026-11-07T10:00:00",
        "source": "KNBS Open Data Portal (Hub-and-Node API)",
        "api_root": K.API_ROOT,
        "exchange_rates": {"2024": 134.82, "2025": 129.30},
        "rate_source": K.RATE_SOURCE,
        "provisional_years": [2025],
        "datasets": {
            "principal_domestic_exports": {
                "flow": "KE1,DF_INTD_VALUE_PDOM_EXPRT,1.0",
                "label": "Values of Principal Domestic Exports",
                "survey_table": "ES 2026 Table 6.7",
                "units": {"KES": 1e6},
                "labels": {},
                "rows": {} if rows is None else rows,
            },
            "coffee_growers": {
                "flow": "KE1,DF_AGR_PRO_ARE_AVE_YIE_COF_TYP_GRO,1.0",
                "label": "Production, Area and Average Yield of Coffee",
                "survey_table": "ES 2026 Table 7.10",
                "units": {"TONNES": 1e3},
                "labels": {},
                "rows": {
                    "AGR_INDIC=PRODUCTION|GROWER_TYPE=_T|ITEM=CP012201|"
                    "TIME_PERIOD=2025|UNIT_MEASURE=TONNES": 51400.0,
                },
            },
        },
    }


ROWS = {"COMMODITY=OW_CT_FLWR|TIME_PERIOD=2025|UNIT_MEASURE=KES": 1.032e11}


@pytest.fixture
def env(tmp_path, monkeypatch):
    """Point the registry, the fetch store and the snapshot at a temp dir."""
    monkeypatch.setattr(settings, "DB_PATH", tmp_path / "platform.db")
    monkeypatch.setattr(settings, "FETCH_DIR", tmp_path / "fetches")
    monkeypatch.setattr(registry, "_INITIALISED", False)
    registry.init_db()
    destination = tmp_path / "data" / "knbs_reference.json"
    monkeypatch.setattr(K, "SNAPSHOT_PATH", str(destination))
    K.clear_cache()
    yield tmp_path
    K.clear_cache()


@pytest.fixture
def client():
    return TestClient(app)


@pytest.fixture
def portal(monkeypatch):
    """Stand in for the live portal with a known snapshot."""
    snapshot = _snapshot(ROWS)
    monkeypatch.setattr(K, "build_snapshot", lambda: json.loads(
        json.dumps(snapshot)))
    return snapshot


# --------------------------------------------------------------------------
# the connector
# --------------------------------------------------------------------------
def test_connector_is_registered_as_a_reference_source():
    entry = next(s for s in connector_catalog() if s["id"] == "knbs")
    assert entry["report_types"] == ["reference"]
    assert entry["experimental"] is False
    assert entry["requires_credentials"] is False
    assert entry["params"] == []          # nothing to fill in: one button
    # A reader has to be able to find the installed series without fetching
    # first, so the catalog names the tables each report prints.
    assert entry["used_by"] == list(USED_BY)
    assert any("Table 13a" in line for line in entry["used_by"])


def test_fetch_returns_the_snapshot_and_a_readable_csv(portal):
    result = KnbsConnector().fetch({})
    assert set(result.files) == {SNAPSHOT_FILE, OBSERVATIONS_FILE}
    assert json.loads(result.files[SNAPSHOT_FILE]) == portal

    rows = list(csv.DictReader(io.StringIO(
        result.files[OBSERVATIONS_FILE].decode("utf-8"))))
    assert len(rows) == 2                 # one per observation, no zero fill
    flower = next(r for r in rows if r["dataset"] == "principal_domestic_exports")
    assert flower["TIME_PERIOD"] == "2025"
    assert flower["COMMODITY"] == "OW_CT_FLWR"
    assert flower["survey_table"] == "ES 2026 Table 6.7"
    assert float(flower["value"]) == pytest.approx(1.032e11)
    coffee = next(r for r in rows if r["dataset"] == "coffee_growers")
    assert coffee["GROWER_TYPE"] == "_T"
    assert float(coffee["value"]) == 51400.0


def test_fetch_provenance_counts_what_it_built(portal):
    result = KnbsConnector().fetch({})
    assert result.provenance["observations"] == 2
    assert result.provenance["datasets"] == ["coffee_growers",
                                             "principal_domestic_exports"]
    assert result.provenance["install_file"] == SNAPSHOT_FILE
    assert result.provenance["url"] == K.API_ROOT
    # A snapshot on its own builds nothing, so the notes have to say where the
    # installed series will show up instead.
    assert any("not a report input" in note for note in result.notes)
    assert [note for note in result.notes if note in USED_BY] == list(USED_BY)


@pytest.mark.parametrize("failure", [
    K.KnbsError("no unit scale for 'CARATS'"),
    urllib.error.URLError("portal unreachable"),
    ValueError("not JSON"),
])
def test_a_portal_failure_is_a_connector_error(monkeypatch, failure):
    def boom():
        raise failure
    monkeypatch.setattr(K, "build_snapshot", boom)
    with pytest.raises(ConnectorError) as exc:
        KnbsConnector().fetch({})
    assert exc.value.status_code == 502


def test_an_empty_portal_is_refused(monkeypatch):
    empty = _snapshot(ROWS)
    for entry in empty["datasets"].values():
        entry["rows"] = {}
    monkeypatch.setattr(K, "build_snapshot", lambda: empty)
    with pytest.raises(ConnectorError) as exc:
        KnbsConnector().fetch({})
    assert exc.value.status_code == 502
    assert "no observations" in str(exc.value)


# --------------------------------------------------------------------------
# fetch -> install, through the API
# --------------------------------------------------------------------------
def _fetch(client):
    response = client.post("/api/sources/knbs/fetch", json={})
    assert response.status_code == 200, response.text
    return response.json()


def test_fetch_route_offers_an_install_url(env, client, portal):
    body = _fetch(client)
    assert body["install_url"] == f"/api/fetches/{body['fetch_id']}/install"
    staged = env / "fetches" / body["fetch_id"]
    assert (staged / SNAPSHOT_FILE).is_file()
    assert (staged / OBSERVATIONS_FILE).is_file()


def test_other_sources_offer_no_install_url(env, client, monkeypatch):
    from webapp.connectors.base import FetchResult
    monkeypatch.setattr("webapp.connectors.unctad.UnctadConnector.fetch",
                        lambda self, values: FetchResult(
                            files={"services.xlsx": b"x"}, provenance={}))
    body = client.post("/api/sources/unctad/fetch", json={}).json()
    assert body["install_url"] is None


def test_install_replaces_the_snapshot_and_clears_the_cache(env, client,
                                                            portal):
    body = _fetch(client)
    destination = env / "data" / "knbs_reference.json"
    destination.parent.mkdir(parents=True, exist_ok=True)
    destination.write_text(json.dumps({"datasets": {"stale": {"rows": {}}}}),
                           encoding="utf-8")
    K.load(path=str(destination))          # prime the cache with the old file
    assert K._cache

    response = client.post(body["install_url"])
    assert response.status_code == 200, response.text
    installed = response.json()
    assert installed["datasets"] == 2
    assert installed["observations"] == 2
    assert installed["installed"] == str(destination)
    assert not K._cache                    # the next build re-reads the file
    assert json.loads(destination.read_text(encoding="utf-8")) == portal
    assert K.series("coffee_growers", path=str(destination),
                    AGR_INDIC="PRODUCTION", GROWER_TYPE="_T",
                    ITEM="CP012201", UNIT_MEASURE="TONNES") == {2025: 51400.0}
    assert not destination.with_name(destination.name + ".installing").exists()


def test_install_refuses_a_snapshot_without_observations(env, client,
                                                         monkeypatch):
    empty = _snapshot(ROWS)
    for entry in empty["datasets"].values():
        entry["rows"] = {}
    monkeypatch.setattr(K, "build_snapshot", lambda: empty)
    # The connector refuses to build it, so stage the file the way a corrupt
    # download would arrive: a fetch record pointing at an empty snapshot.
    directory = env / "fetches" / FETCH_ID
    directory.mkdir(parents=True)
    (directory / SNAPSHOT_FILE).write_text(json.dumps(empty), encoding="utf-8")
    registry.create_fetch(FETCH_ID, "knbs", {}, {}, [], directory, 3600)

    destination = env / "data" / "knbs_reference.json"
    destination.parent.mkdir(parents=True, exist_ok=True)
    destination.write_text('{"keep": true}', encoding="utf-8")
    response = client.post(f"/api/fetches/{FETCH_ID}/install")
    assert response.status_code == 400
    assert "no observations" in response.json()["detail"]
    assert destination.read_text(encoding="utf-8") == '{"keep": true}'


def test_install_refuses_unreadable_json(env, client):
    directory = env / "fetches" / FETCH_ID
    directory.mkdir(parents=True)
    (directory / SNAPSHOT_FILE).write_text("{not json", encoding="utf-8")
    registry.create_fetch(FETCH_ID, "knbs", {}, {}, [], directory, 3600)
    response = client.post(f"/api/fetches/{FETCH_ID}/install")
    assert response.status_code == 400
    assert "not readable JSON" in response.json()["detail"]


def test_install_is_only_for_the_knbs_source(env, client):
    directory = env / "fetches" / FETCH_ID
    directory.mkdir(parents=True)
    (directory / SNAPSHOT_FILE).write_text("{}", encoding="utf-8")
    registry.create_fetch(FETCH_ID, "unctad", {}, {}, [], directory, 3600)
    response = client.post(f"/api/fetches/{FETCH_ID}/install")
    assert response.status_code == 400
    assert "Only the KNBS reference snapshot" in response.json()["detail"]


def test_install_of_an_unknown_fetch_is_404(env, client):
    assert client.post(f"/api/fetches/{FETCH_ID}/install").status_code == 404


def test_install_needs_the_api_key_when_one_is_set(env, client, portal,
                                                   monkeypatch):
    body = _fetch(client)
    monkeypatch.setattr(settings, "API_KEY", "secret")
    assert client.post(body["install_url"]).status_code == 401
    assert client.post(body["install_url"],
                       headers={"x-api-key": "wrong"}).status_code == 401
    ok = client.post(body["install_url"], headers={"x-api-key": "secret"})
    assert ok.status_code == 200


def test_a_reference_fetch_still_cannot_build_a_report(env, client, portal):
    body = _fetch(client)
    for payload in ({"report_type": "goods"}, {"report_type": "services"}, {}):
        response = client.post(body["run_url"], json=payload)
        assert response.status_code == 400, payload
        detail = response.json()["detail"]
        assert "reference dataset" in detail
        assert "Install the snapshot" in detail