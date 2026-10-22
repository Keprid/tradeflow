"""API-level tests for the FastAPI webapp (validation & hardening)."""

import pytest
from fastapi.testclient import TestClient

from webapp.main import app

c = TestClient(app)

JOB0 = "0" * 32


def test_index_serves_split_assets():
    r = c.get("/")
    assert r.status_code == 200
    assert b"/static/style.css" in r.content
    assert b"/static/app.js" in r.content
    assert b"role=\"alert\"" in r.content          # error box is announced
    assert b"<noscript>" in r.content
    assert b'integrity="sha384-' in r.content     # SRI on the CDN stylesheet


def test_static_assets():
    assert c.get("/static/style.css").status_code == 200
    assert c.get("/static/app.js").status_code == 200
    assert c.get("/static/keproba_letterhead.png").status_code == 200


def test_configs_include_auto():
    r = c.get("/api/configs")
    assert r.status_code == 200
    ids = [x["id"] for x in r.json()]
    assert ids[0] == "__auto__"


def test_status_unknown_job_404():
    assert c.get(f"/api/status/{JOB0}").status_code == 404


def test_status_invalid_job_id_404():
    assert c.get("/api/status/not-a-job").status_code == 404


def test_download_traversal_rejected():
    assert c.get("/api/download/..%2F..%2Fetc%2Fpasswd").status_code == 404
    assert c.get(f"/api/download/{JOB0}").status_code == 404


def test_promotion_missing_manifest_is_404_not_500():
    assert c.get(f"/api/promotion/{JOB0}").status_code == 404


def test_tables_and_product_tables_404():
    assert c.get(f"/api/tables/{JOB0}").status_code == 404
    assert c.get(f"/api/product-tables/{JOB0}").status_code == 404


def test_competitor_file_404():
    assert c.get(f"/api/competitor-file/{JOB0}/png").status_code == 404


def test_run_requires_files():
    r = c.post("/api/run", data={"top": "20"})
    assert r.status_code == 400
    assert "select files" in r.json()["detail"].lower()


def test_run_top_bounds():
    files = {"files": ("t.csv", b"a,b\n1,2\n", "text/csv")}
    r = c.post("/api/run", data={"top": "0"}, files=files)
    assert r.status_code == 400
    r = c.post("/api/run", data={"top": "101"}, files=files)
    assert r.status_code == 400
    r = c.post("/api/run", data={"top": "50", "config": "__auto__"},
               files=files)
    # passes validation; fails later at upload-set detection (400, not 500)
    assert r.status_code == 400


def test_run_rejects_config_traversal():
    files = {"files": ("t.csv", b"a,b\n1,2\n", "text/csv")}
    r = c.post("/api/run",
               data={"top": "20", "config": "../../etc/passwd"},
               files=files)
    assert r.status_code == 400
    assert "Invalid config id" in r.json()["detail"]


def test_run_rejects_bad_report_type():
    files = {"files": ("t.csv", b"a,b\n1,2\n", "text/csv")}
    r = c.post("/api/run",
               data={"top": "20", "report_type": "nope"}, files=files)
    assert r.status_code == 400


def test_run_quarterly_requires_period():
    files = {"files": ("t.csv", b"a,b\n1,2\n", "text/csv")}
    r = c.post("/api/run",
               data={"top": "20", "report_type": "quarterly"}, files=files)
    assert r.status_code == 400
    assert "review period" in r.json()["detail"]


def test_run_rejects_unsupported_extension():
    files = {"files": ("malware.exe", b"MZ", "application/octet-stream")}
    r = c.post("/api/run", data={"top": "20"}, files=files)
    assert r.status_code == 400
    assert "Unsupported file type" in r.json()["detail"]
