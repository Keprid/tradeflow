"""SQLite registry for jobs, durable queue rows, provenance and fetches."""

import json
import sqlite3
import threading
import time
from contextlib import contextmanager

from . import settings

_LOCK = threading.Lock()
_INITIALISED = False

_SCHEMA = """
CREATE TABLE IF NOT EXISTS jobs (
    id              TEXT PRIMARY KEY,
    created         REAL NOT NULL,
    updated         REAL NOT NULL,
    status          TEXT NOT NULL,
    report_type     TEXT,
    mode            TEXT,
    config          TEXT,
    period          TEXT,
    owner_hash      TEXT,
    attempts        INTEGER NOT NULL DEFAULT 0,
    max_attempts    INTEGER NOT NULL DEFAULT 2,
    spec_json       TEXT,
    provenance_json TEXT,
    manifest_json   TEXT,
    quality_json    TEXT,
    error           TEXT,
    report_version  TEXT,
    started_at      REAL,
    finished_at     REAL
);
CREATE INDEX IF NOT EXISTS idx_jobs_status ON jobs(status, created);
CREATE INDEX IF NOT EXISTS idx_jobs_owner ON jobs(owner_hash, created);

CREATE TABLE IF NOT EXISTS fetches (
    id              TEXT PRIMARY KEY,
    created         REAL NOT NULL,
    source          TEXT NOT NULL,
    params_json     TEXT,
    provenance_json TEXT,
    files_json      TEXT,
    dir             TEXT NOT NULL,
    expires_at      REAL
);

CREATE TABLE IF NOT EXISTS comments (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    job_id      TEXT NOT NULL,
    author      TEXT,
    body        TEXT NOT NULL,
    created     REAL NOT NULL
);
CREATE INDEX IF NOT EXISTS idx_comments_job ON comments(job_id, created);
"""

_JOB_FIELDS = {
    "status", "report_type", "mode", "config", "period", "owner_hash",
    "attempts", "max_attempts", "spec_json", "provenance_json",
    "manifest_json", "quality_json", "error", "report_version",
    "started_at", "finished_at",
}


def _connect():
    conn = sqlite3.connect(str(settings.DB_PATH), timeout=30)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA busy_timeout=30000")
    return conn


def init_db():
    global _INITIALISED
    with _LOCK:
        if _INITIALISED:
            return
        conn = _connect()
        try:
            conn.executescript(_SCHEMA)
            conn.commit()
        finally:
            conn.close()
        _INITIALISED = True


@contextmanager
def _transaction():
    init_db()
    conn = _connect()
    try:
        with conn:
            yield conn
    finally:
        conn.close()


def _decode(row):
    if row is None:
        return None
    data = dict(row)
    for key in ("spec_json", "provenance_json", "manifest_json",
                "quality_json", "params_json", "files_json"):
        if key in data:
            raw = data.pop(key)
            base = key[:-5]
            try:
                data[base] = json.loads(raw) if raw else None
            except (TypeError, ValueError):
                data[base] = None
    return data


def create_job(job_id, report_type, mode="", config="", period="", spec=None,
               owner_hash=None, provenance=None, max_attempts=None,
               report_version=None):
    now = time.time()
    spec = spec or {}
    with _transaction() as conn:
        conn.execute(
            """INSERT INTO jobs
               (id, created, updated, status, report_type, mode, config,
                period, owner_hash, attempts, max_attempts, spec_json,
                provenance_json, report_version)
               VALUES (?,?,?,?,?,?,?,?,?,0,?,?,?,?)""",
            (job_id, now, now, "queued", report_type, mode, config, period,
             owner_hash, max_attempts or settings.MAX_ATTEMPTS,
             json.dumps(spec), json.dumps(provenance or {}),
             report_version or settings.REPORT_VERSION),
        )
    return get_job(job_id)


def get_job(job_id):
    init_db()
    conn = _connect()
    try:
        row = conn.execute("SELECT * FROM jobs WHERE id=?", (job_id,)).fetchone()
        return _decode(row)
    finally:
        conn.close()


def update_job(job_id, **fields):
    clean = {k: v for k, v in fields.items() if k in _JOB_FIELDS}
    if not clean:
        return
    if isinstance(clean.get("spec"), (dict, list)):
        clean["spec_json"] = json.dumps(clean.pop("spec"))
    if isinstance(clean.get("provenance"), (dict, list)):
        clean["provenance_json"] = json.dumps(clean.pop("provenance"))
    if isinstance(clean.get("manifest"), (dict, list)):
        clean["manifest_json"] = json.dumps(clean.pop("manifest"))
    if isinstance(clean.get("quality"), (dict, list)):
        clean["quality_json"] = json.dumps(clean.pop("quality"))
    for key in ("spec_json", "provenance_json", "manifest_json",
                "quality_json"):
        if key in clean and not isinstance(clean[key], str):
            clean[key] = json.dumps(clean[key])
    clean["updated"] = time.time()
    assignments = ", ".join(f"{k}=?" for k in clean)
    values = list(clean.values()) + [job_id]
    with _transaction() as conn:
        conn.execute(f"UPDATE jobs SET {assignments} WHERE id=?", values)


def set_status(job_id, status, **fields):
    fields["status"] = status
    if status in ("done", "error"):
        fields.setdefault("finished_at", time.time())
    update_job(job_id, **fields)


def list_jobs(owner_hash=None, limit=50):
    init_db()
    conn = _connect()
    try:
        if owner_hash:
            rows = conn.execute(
                "SELECT * FROM jobs WHERE owner_hash=? "
                "ORDER BY created DESC LIMIT ?",
                (owner_hash, limit)).fetchall()
        else:
            rows = conn.execute(
                "SELECT * FROM jobs ORDER BY created DESC LIMIT ?",
                (limit,)).fetchall()
        return [_decode(r) for r in rows]
    finally:
        conn.close()


def claim_next():
    """Atomically move the oldest queued job to ``running``."""
    init_db()
    conn = _connect()
    try:
        while True:
            row = conn.execute(
                "SELECT id FROM jobs WHERE status='queued' "
                "ORDER BY created LIMIT 1").fetchone()
            if row is None:
                return None
            job_id = row["id"]
            now = time.time()
            with conn:
                changed = conn.execute(
                    "UPDATE jobs SET status='running', started_at=?, "
                    "updated=?, attempts=attempts+1 "
                    "WHERE id=? AND status='queued'",
                    (now, now, job_id)).rowcount
            if changed:
                return get_job(job_id)
    finally:
        conn.close()


def requeue(job_id, error=None):
    update_job(job_id, status="queued", error=error)


def reset_running():
    """Requeue jobs a previous process left running (crash/restart)."""
    init_db()
    conn = _connect()
    try:
        with conn:
            rows = conn.execute(
                "SELECT id, attempts, max_attempts FROM jobs "
                "WHERE status='running'").fetchall()
            for row in rows:
                if row["attempts"] < row["max_attempts"]:
                    conn.execute(
                        "UPDATE jobs SET status='queued', updated=? "
                        "WHERE id=?", (time.time(), row["id"]))
                else:
                    conn.execute(
                        "UPDATE jobs SET status='error', error=?, updated=? "
                        "WHERE id=?",
                        ("Interrupted by a server restart and the retry "
                         "budget is exhausted.", time.time(), row["id"]))
    finally:
        conn.close()


def queue_depth():
    init_db()
    conn = _connect()
    try:
        row = conn.execute(
            "SELECT COUNT(*) AS n FROM jobs WHERE status='queued'").fetchone()
        return int(row["n"])
    finally:
        conn.close()


def status_counts():
    init_db()
    conn = _connect()
    try:
        return {r["status"]: int(r["n"]) for r in conn.execute(
            "SELECT status, COUNT(*) AS n FROM jobs GROUP BY status")}
    finally:
        conn.close()


def delete_job(job_id):
    init_db()
    conn = _connect()
    try:
        with conn:
            conn.execute("DELETE FROM comments WHERE job_id=?", (job_id,))
            conn.execute("DELETE FROM jobs WHERE id=?", (job_id,))
    finally:
        conn.close()


def delete_jobs_before(cutoff):
    init_db()
    conn = _connect()
    try:
        with conn:
            rows = conn.execute(
                "SELECT id FROM jobs WHERE updated < ?", (cutoff,)).fetchall()
            for row in rows:
                conn.execute("DELETE FROM comments WHERE job_id=?", (row["id"],))
                conn.execute("DELETE FROM jobs WHERE id=?", (row["id"],))
            return [r["id"] for r in rows]
    finally:
        conn.close()


# ---------------------------------------------------------------------------
# Fetched connector outputs
# ---------------------------------------------------------------------------
def create_fetch(fetch_id, source, params, provenance, files, directory,
                 ttl_seconds):
    now = time.time()
    with _transaction() as conn:
        conn.execute(
            """INSERT INTO fetches
               (id, created, source, params_json, provenance_json,
                files_json, dir, expires_at)
               VALUES (?,?,?,?,?,?,?,?)""",
            (fetch_id, now, source, json.dumps(params or {}),
             json.dumps(provenance or {}), json.dumps(files or []),
             str(directory), now + ttl_seconds),
        )
    return get_fetch(fetch_id)


def get_fetch(fetch_id):
    init_db()
    conn = _connect()
    try:
        row = conn.execute("SELECT * FROM fetches WHERE id=?",
                           (fetch_id,)).fetchone()
        return _decode(row)
    finally:
        conn.close()


def expired_fetches(now=None):
    init_db()
    now = now or time.time()
    conn = _connect()
    try:
        rows = conn.execute(
            "SELECT * FROM fetches WHERE expires_at IS NOT NULL "
            "AND expires_at < ?", (now,)).fetchall()
        return [_decode(r) for r in rows]
    finally:
        conn.close()


def delete_fetch(fetch_id):
    init_db()
    conn = _connect()
    try:
        with conn:
            conn.execute("DELETE FROM fetches WHERE id=?", (fetch_id,))
    finally:
        conn.close()


# ---------------------------------------------------------------------------
# Comments / annotations
# ---------------------------------------------------------------------------
def add_comment(job_id, body, author=""):
    now = time.time()
    with _transaction() as conn:
        cur = conn.execute(
            "INSERT INTO comments (job_id, author, body, created) "
            "VALUES (?,?,?,?)", (job_id, author or "", body, now))
        return {"id": cur.lastrowid, "job_id": job_id, "author": author or "",
                "body": body, "created": now}


def list_comments(job_id):
    init_db()
    conn = _connect()
    try:
        rows = conn.execute(
            "SELECT * FROM comments WHERE job_id=? ORDER BY created",
            (job_id,)).fetchall()
        return [dict(r) for r in rows]
    finally:
        conn.close()
