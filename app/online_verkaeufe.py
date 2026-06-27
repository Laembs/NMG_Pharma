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


def fetch_vorbestellungen(base_url: str | None = None, firma: str | None = None,
                          timeout: float = 6.0):
    """Offene Online-Vorbestellungen abrufen. Gibt (liste, fehler) zurück.

    Jeder Eintrag: {auftrag, datum, kunde_name, liefertermin, pzn, bezeichnung, menge}.
    """
    base = (base_url or basis_url()).rstrip("/")
    fa = firma or os.environ.get("NMGKASSE_FIRMA") or DEFAULT_FIRMA
    sec = _secret()
    if not sec:
        return [], "Kein SSO-Geheimnis gefunden."
    url = f"{base}/api/vorbestellungen?token={_make_token(sec, fa)}"
    try:
        with urllib.request.urlopen(url, timeout=timeout) as resp:
            data = json.loads(resp.read().decode("utf-8"))
        return list(data.get("vorbestellungen", [])), None
    except urllib.error.HTTPError as exc:
        return [], f"Server antwortet {exc.code}"
    except urllib.error.URLError as exc:
        return [], f"nicht erreichbar ({exc.reason})"
    except Exception as exc:  # pragma: no cover - defensiv
        return [], f"Fehler: {exc.__class__.__name__}"


def push_lager(artikel, bestand, base_url: str | None = None, firma: str | None = None,
               timeout: float = 10.0):
    """Schiebt Artikelstamm + Bestand der PC-Kasse zum Server (PC ist die Quelle).
    Gibt (ok, fehler) zurück. Der Server zeigt dem Handy dann verfügbar = Bestand
    − offene Online-Bestellungen."""
    base = (base_url or basis_url()).rstrip("/")
    fa = firma or os.environ.get("NMGKASSE_FIRMA") or DEFAULT_FIRMA
    sec = _secret()
    if not sec:
        return False, "Kein SSO-Geheimnis gefunden."
    import urllib.parse
    url = f"{base}/api/lager/push?" + urllib.parse.urlencode({"token": _make_token(sec, fa)})
    data = json.dumps({"artikel": artikel, "bestand": bestand}).encode("utf-8")
    try:
        req = urllib.request.Request(
            url, data=data, method="POST", headers={"Content-Type": "application/json"})
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            return True, json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as exc:
        return False, f"Server antwortet {exc.code}"
    except urllib.error.URLError as exc:
        return False, f"nicht erreichbar ({exc.reason})"
    except Exception as exc:  # pragma: no cover - defensiv
        return False, f"Fehler: {exc.__class__.__name__}"


def mark_uebernommen(auftrag_id, von: str = "", base_url: str | None = None,
                     firma: str | None = None, timeout: float = 6.0):
    """Markiert einen Online-Auftrag am Server als in die PC-Kasse übernommen.
    Gibt (ok, fehler) zurück."""
    base = (base_url or basis_url()).rstrip("/")
    fa = firma or os.environ.get("NMGKASSE_FIRMA") or DEFAULT_FIRMA
    sec = _secret()
    if not sec:
        return False, "Kein SSO-Geheimnis gefunden."
    import urllib.parse
    q = urllib.parse.urlencode({"token": _make_token(sec, fa), "von": von or "PC-Kasse"})
    url = f"{base}/api/verkauf/{int(auftrag_id)}/uebernommen?{q}"
    try:
        req = urllib.request.Request(url, data=b"", method="POST")
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            json.loads(resp.read().decode("utf-8"))
        return True, None
    except urllib.error.HTTPError as exc:
        return False, f"Server antwortet {exc.code}"
    except urllib.error.URLError as exc:
        return False, f"nicht erreichbar ({exc.reason})"
    except Exception as exc:  # pragma: no cover - defensiv
        return False, f"Fehler: {exc.__class__.__name__}"


def fetch_detail(auftrag_id, base_url: str | None = None, firma: str | None = None,
                 timeout: float = 6.0):
    """Holt einen einzelnen Online-Auftrag (Kopf + Positionen). Gibt (daten, fehler)."""
    base = (base_url or basis_url()).rstrip("/")
    fa = firma or os.environ.get("NMGKASSE_FIRMA") or DEFAULT_FIRMA
    sec = _secret()
    if not sec:
        return None, "Kein SSO-Geheimnis gefunden."
    url = f"{base}/api/verkauf/{int(auftrag_id)}?token={_make_token(sec, fa)}"
    try:
        with urllib.request.urlopen(url, timeout=timeout) as resp:
            return json.loads(resp.read().decode("utf-8")), None
    except urllib.error.HTTPError as exc:
        return None, f"Server antwortet {exc.code}"
    except urllib.error.URLError as exc:
        return None, f"nicht erreichbar ({exc.reason})"
    except Exception as exc:  # pragma: no cover - defensiv
        return None, f"Fehler: {exc.__class__.__name__}"
