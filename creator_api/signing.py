from __future__ import annotations

import hashlib
import hmac
import time
from typing import TYPE_CHECKING
from urllib.parse import urlencode

if TYPE_CHECKING:
    from .config import Settings


def _signature(secret: str, job_id: str, user_id: str, expires: int) -> str:
    payload = f"{job_id}:{user_id}:{expires}".encode()
    return hmac.new(secret.encode(), payload, hashlib.sha256).hexdigest()


def signed_download_url(settings: Settings, job_id: str, user_id: str) -> str:
    expires = int(time.time()) + settings.signed_url_ttl_seconds
    query = urlencode(
        {
            "user": user_id,
            "expires": expires,
            "signature": _signature(settings.download_signing_secret, job_id, user_id, expires),
        }
    )
    return f"{settings.public_api_url}/v1/video/jobs/{job_id}/download?{query}"


def valid_download_signature(
    settings: Settings,
    job_id: str,
    user_id: str,
    expires: int,
    signature: str,
) -> bool:
    if expires < int(time.time()):
        return False
    expected = _signature(settings.download_signing_secret, job_id, user_id, expires)
    return hmac.compare_digest(expected, signature)
