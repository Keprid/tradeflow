"""Environment-driven platform settings.

Defaults are chosen so local development behaves exactly as before: no API
key, one renderer, and job ownership enforced only for jobs created by this
version of the app (legacy job folders without a registry row stay readable).
"""

import os
import secrets
from pathlib import Path

WEBAPP_DIR = Path(__file__).resolve().parent.parent
BASE_DIR = WEBAPP_DIR.parent


def _env(name, default=""):
    value = os.environ.get(name)
    return default if value is None else value.strip()


def _int_env(name, default):
    raw = _env(name, "")
    if not raw:
        return default
    try:
        return int(raw)
    except ValueError:
        return default


DATA_DIR = Path(_env("TRADEFLOW_DATA_DIR", str(WEBAPP_DIR))).resolve()
DATA_DIR.mkdir(parents=True, exist_ok=True)

DB_PATH = Path(_env("TRADEFLOW_DB", str(DATA_DIR / "platform.db"))).resolve()

# Jobs, uploads and fetches stay in the existing locations unless redirected.
JOBS_DIR = Path(_env("TRADEFLOW_JOBS_DIR", str(WEBAPP_DIR / "jobs"))).resolve()
FETCH_DIR = Path(_env("TRADEFLOW_FETCH_DIR",
                      str(DATA_DIR / "fetches"))).resolve()
CACHE_DIR = Path(_env("TRADEFLOW_CACHE_DIR",
                      str(DATA_DIR / "cache"))).resolve()
for _p in (JOBS_DIR, FETCH_DIR, CACHE_DIR):
    _p.mkdir(parents=True, exist_ok=True)

REPORT_VERSION = _env("TRADEFLOW_REPORT_VERSION", "2026.1")
JOB_RETENTION_SECONDS = _int_env("TRADEFLOW_JOB_TTL", 24 * 3600)
FETCH_RETENTION_SECONDS = _int_env("TRADEFLOW_FETCH_TTL", 6 * 3600)
WORKER_COUNT = max(1, _int_env("TRADEFLOW_WORKERS", 1))
MAX_ATTEMPTS = max(1, _int_env("TRADEFLOW_JOB_ATTEMPTS", 2))
SHARE_TTL_SECONDS = _int_env("TRADEFLOW_SHARE_TTL", 7 * 24 * 3600)

# Optional coarse API gate for private deployments.  When unset the app stays
# open, matching the current single-user workflow.
API_KEY = _env("TRADEFLOW_API_KEY", "")

# Job ownership is enforced for registered jobs.  Set to 0 only for a trusted
# single-user deployment that needs bookmarkable bare job URLs.
ENFORCE_JOB_OWNERSHIP = _env("TRADEFLOW_ENFORCE_JOB_OWNERSHIP", "1").lower() \
    not in ("0", "false", "no", "off")

# Optional completion webhook (POST JSON) for downstream automation.
WEBHOOK_URL = _env("TRADEFLOW_WEBHOOK_URL", "")
WEBHOOK_TIMEOUT = _int_env("TRADEFLOW_WEBHOOK_TIMEOUT", 5)

# Connector credentials are read from the environment, never from the browser.
UNCTAD_CLIENT_ID = _env("UNCTAD_CLIENT_ID", "")
UNCTAD_API_KEY = _env("UNCTAD_API_KEY", "")
COMTRADE_API_KEY = _env("COMTRADE_API_KEY", "")

SECRET_FILE = DATA_DIR / ".platform_secret"


def _load_secret():
    env = _env("TRADEFLOW_SECRET", "")
    if env:
        return env.encode("utf-8")
    try:
        if SECRET_FILE.exists():
            value = SECRET_FILE.read_text(encoding="utf-8").strip()
            if value:
                return value.encode("utf-8")
        value = secrets.token_urlsafe(48)
        SECRET_FILE.write_text(value, encoding="utf-8")
        try:
            SECRET_FILE.chmod(0o600)
        except OSError:
            pass
        return value.encode("utf-8")
    except OSError:
        # Read-only disk: fall back to an ephemeral secret.  Owner/share
        # tokens then stop verifying after a restart, which is safe.
        return secrets.token_bytes(48)


SECRET = _load_secret()
