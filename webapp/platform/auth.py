"""Job ownership tokens, signed share links and the optional API gate."""

import base64
import hashlib
import hmac
import secrets
import time

from . import settings

_OWNER_HEADER = "x-job-token"
_SHARE_PARAM = "share"
_OWNER_PARAM = "token"


def new_owner_token():
    return secrets.token_urlsafe(32)


def token_hash(token):
    return hashlib.sha256((token or "").encode("utf-8")).hexdigest()


def _b64(raw):
    return base64.urlsafe_b64encode(raw).decode("ascii").rstrip("=")


def _unb64(text):
    pad = "=" * (-len(text) % 4)
    return base64.urlsafe_b64decode(text + pad)


def _sign(message):
    return _b64(hmac.new(settings.SECRET, message.encode("utf-8"),
                         hashlib.sha256).digest())


def make_share_token(job_id, ttl_seconds=None):
    ttl = ttl_seconds or settings.SHARE_TTL_SECONDS
    expires = int(time.time() + max(60, ttl))
    return f"{expires}.{_sign(f'share:{job_id}:{expires}')}", expires


def verify_share_token(job_id, token):
    if not token or "." not in token:
        return False
    expires_raw, signature = token.split(".", 1)
    try:
        expires = int(expires_raw)
    except ValueError:
        return False
    if expires < time.time():
        return False
    expected = _sign(f"share:{job_id}:{expires}")
    return hmac.compare_digest(expected, signature)


def token_from_request(request):
    token = request.headers.get(_OWNER_HEADER) or ""
    if not token:
        token = request.query_params.get(_OWNER_PARAM) or ""
    return token


def share_from_request(request):
    return request.query_params.get(_SHARE_PARAM) or ""


def is_authorized(request, job, allow_share=True):
    """True when the caller may read/act on *job*.

    Legacy job folders with no registry row (``job`` is None) and jobs with
    no owner hash stay accessible, so the upgrade never strands old links.
    """
    if not settings.ENFORCE_JOB_OWNERSHIP:
        return True
    if job is None:
        return True
    owner_hash = job.get("owner_hash")
    if not owner_hash:
        return True
    if allow_share and verify_share_token(
            job.get("id"), share_from_request(request)):
        return True
    supplied = token_from_request(request)
    return bool(supplied) and hmac.compare_digest(
        token_hash(supplied), owner_hash)


def api_key_ok(request):
    if not settings.API_KEY:
        return True
    supplied = request.headers.get("x-api-key") or \
        request.query_params.get("api_key") or ""
    return hmac.compare_digest(supplied, settings.API_KEY)
