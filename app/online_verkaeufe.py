"""Online-Verkäufe der Web-Kasse (nmgkasse.pennone.de) für die Desktop-Kasse abrufen.

Die per Handy/Browser erfassten Verkäufe liegen in der zentralen Online-Datenbank.
Diese kleine Hilfe holt sie über die Lese-API der Web-Kasse, damit die Desktop-Kasse
sie in „Verkäufe" mit dem Zusatz „(online)" anzeigen kann.

Authentifiziert per kurzlebigem HMAC-Token mit dem gemeinsamen Geheimnis (dieselbe
Datei wie das Cockpit-SSO). Reiner Lesezugriff – es wird nichts geändert. Nur
Standardbibliothek, damit keine Zusatz-Abhängigkeit nötig ist.

Konfiguration (optional über Umgebungsvariablen):
  NMGKASSE_URL    – Basis-URL (Standard https://nmgkasse.pennone.de)
  NMGKASSE_FIRMA  – Firmen-Slug (Standard nmgpharma)
  PENNONE_SSO_SECRET / NMGKASSE_SSO_SECRET_FILE – Geheimnis bzw. Pfad zur Datei
"""
from __future__ import annotations

import base64
import hashlib
import hmac
import json
import os
import secrets
import time
import urllib.error
import urllib.request

DEFAULT_BASE_URL = "https://nmgkasse.pennone.de"
DEFAULT_FIRMA = "nmgpharma"
DEFAULT_SECRET_FILE = r"C:\pennone_one\data\cockpit_sso_secret.txt"


def _secret() -> str:
    env = os.environ.get("PENNONE_SSO_SECRET")
    if env:
        return env.strip()
    path = os.environ.get("NMGKASSE_SSO_SECRET_FILE", DEFAULT_SECRET_FILE)
    try:
        with open(path, encoding="utf-8") as fh:
            return fh.read().strip()
    except OSError:
        return ""


def _b64e(b: bytes) -> str:
    return base64.urlsafe_b64encode(b).rstrip(b"=").decode("ascii")


def _make_token(secret: str, firma: str, login: str = "desktop", ttl: int = 30) -> str:
    payload = {"f": firma, "l": login,
               "exp": int(time.time()) + ttl, "jti": secrets.token_hex(8)}
    body = _b64e(json.dumps(payload, separators=(",", ":"), sort_keys=True).encode("utf-8"))
    sig = _b64e(hmac.new(secret.encode("utf-8"), body.encode("ascii"),
                         hashlib.sha256).digest())
    return f"{body}.{sig}"


def basis_url() -> str:
    return (os.environ.get("NMGKASSE_URL") or DEFAULT_BASE_URL).rstrip("/")


def fetch(base_url: str | None = None, firma: str | None = None, timeout: float = 6.0):
    """Holt die Online-Verkäufe. Gibt (liste, fehler) zurück – fehler=None bei Erfolg.

    Jeder Eintrag: {id, datum, kunde_name, status, erfasst_von, positionen, stueck}.
    """
    base = (base_url or basis_url()).rstrip("/")
    fa = firma or os.environ.get("NMGKASSE_FIRMA") or DEFAULT_FIRMA
    sec = _secret()
    if not sec:
        return [], "Kein SSO-Geheimnis gefunden (cockpit_sso_secret.txt fehlt)."
    url = f"{base}/api/verkaeufe?token={_make_token(sec, fa)}"
    try:
        with urllib.request.urlopen(url, timeout=timeout) as resp:
            data = json.loads(resp.read().decode("utf-8"))
        return list(data.get("verkaeufe", [])), None
    except urllib.error.HTTPError as exc:
        return [], f"Server antwortet {exc.code}"
    except urllib.error.URLError as exc:
        return [], f"nicht erreichbar ({exc.reason})"
    except Exception as exc:  # pragma: no cover - defensiv
        return [], f"Fehler: {exc.__class__.__name__}"
