"""Short-lived HMAC tokens for public one-time workflows."""

import hashlib
import hmac
import secrets
import time


def create_expiring_token(secret: str, purpose: str, lifetime_seconds: int = 86400) -> str:
    issued = int(time.time())
    nonce = secrets.token_urlsafe(18)
    body = f"v1.{issued}.{int(lifetime_seconds)}.{nonce}"
    signature = hmac.new(str(secret).encode(), f"{purpose}:{body}".encode(), hashlib.sha256).hexdigest()
    return f"{body}.{signature}"


def verify_expiring_token(token: str, secret: str, purpose: str, max_lifetime_seconds: int = 604800) -> bool:
    try:
        version, issued_text, lifetime_text, nonce, supplied = str(token).split(".", 4)
        issued, lifetime = int(issued_text), int(lifetime_text)
        if version != "v1" or not nonce or lifetime <= 0 or lifetime > max_lifetime_seconds:
            return False
        body = f"{version}.{issued}.{lifetime}.{nonce}"
        expected = hmac.new(str(secret).encode(), f"{purpose}:{body}".encode(), hashlib.sha256).hexdigest()
        now = int(time.time())
        return hmac.compare_digest(expected, supplied) and issued - 60 <= now <= issued + lifetime
    except Exception:
        return False
