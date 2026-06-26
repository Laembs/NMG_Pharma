"""Cockpit-SSO (Kasse-Seite): Einmal-Token vom NMGone-Cockpit pruefen.

Spiegelbild der Pruef-Logik aus ``pennone/sso.py`` bzw. ``cockpit_sso.py`` –
bewusst nur Standardbibliothek, damit zwischen Cockpit, ONE und Kasse keine
Zusatz-Abhaengigkeit synchron gehalten werden muss. Das Cockpit signiert Firma +
Login mit dem gemeinsamen Geheimnis (Env ``PENNONE_SSO_SECRET``); die Kasse prueft
Signatur + Ablauf und verbraucht das Token einmalig.
"""
from __future__ import annotations

import base64
import hashlib
import hmac
import json
import os
import time


def secret() -> str:
    """Gemeinsames SSO-Geheimnis (muss mit Cockpit/ONE uebereinstimmen)."""
    return os.environ.get("PENNONE_SSO_SECRET", "").strip()


def _b64d(s: str) -> bytes:
    return base64.urlsafe_b64decode(s + "=" * (-len(s) % 4))


def _b64e(b: bytes) -> str:
    return base64.urlsafe_b64encode(b).rstrip(b"=").decode("ascii")


def verify_token(token: str, sec: str | None = None) -> dict | None:
    """Token pruefen (Signatur + Ablauf). Gibt Payload-dict oder None."""
    sec = sec if sec is not None else secret()
    if not sec:
        return None
    try:
        body, sig = (token or "").split(".", 1)
    except ValueError:
        return None
    expect = _b64e(hmac.new(sec.encode("utf-8"), body.encode("ascii"),
                            hashlib.sha256).digest())
    if not hmac.compare_digest(sig, expect):
        return None
    try:
        payload = json.loads(_b64d(body))
    except (ValueError, json.JSONDecodeError):
        return None
    if int(payload.get("exp", 0)) < int(time.time()):
        return None
    return payload


# ── Einmal-Verwendung (Replay-Schutz, In-Memory; Token leben nur ~30s) ───────
_seen: dict[str, int] = {}  # jti -> exp


def consume_jti(payload: dict) -> bool:
    """True, wenn das Token noch nicht benutzt wurde; markiert es danach."""
    now = int(time.time())
    for jti, exp in list(_seen.items()):
        if exp < now:
            _seen.pop(jti, None)
    jti = payload.get("jti")
    if not jti or jti in _seen:
        return False
    _seen[jti] = int(payload.get("exp", now))
    return True
