"""Zentrale Cockpit-Meldungen/Aufgaben über den Hetzner-Dienst (nmgkasse.pennone.de).

Spiegelbild zu :mod:`app.online_verkaeufe`: damit ein an einem Arbeitsplatz
gemeldeter Warenausgang an ALLEN Arbeitsplätzen erscheint, liegen die Meldungen
zentral beim Web-Dienst statt in einer lokalen Datei. Diese kleine Hilfe spricht
die `/api/meldungen`-Endpunkte an (lesen/anlegen/abhaken/löschen).

Authentifiziert per kurzlebigem HMAC-Token mit dem gemeinsamen Geheimnis – die
Token-/Secret-/URL-Logik wird aus :mod:`app.online_verkaeufe` wiederverwendet,
damit es genau EINE Quelle dafür gibt. Nur Standardbibliothek (urllib).

Konfiguration (wie bei den Verkäufen):
  NMGKASSE_URL    – Basis-URL (Standard https://nmgkasse.pennone.de)
  NMGKASSE_FIRMA  – Firmen-Slug (Standard nmgpharma)
  PENNONE_SSO_SECRET / NMGKASSE_SSO_SECRET_FILE – Geheimnis bzw. Pfad zur Datei
"""
from __future__ import annotations

import json
import os
import urllib.error
import urllib.parse
import urllib.request

from app import online_verkaeufe as _ov


def is_configured() -> bool:
    """True, wenn ein SSO-Geheimnis vorliegt – nur dann ist der Online-Weg möglich."""
    return bool(_ov._secret())


def _firma(firma: str | None) -> str:
    return firma or os.environ.get("NMGKASSE_FIRMA") or _ov.DEFAULT_FIRMA


def _token(firma: str) -> str | None:
    sec = _ov._secret()
    if not sec:
        return None
    return _ov._make_token(sec, firma)


def fetch(base_url: str | None = None, firma: str | None = None, timeout: float = 6.0):
    """Holt die zentrale Meldungsliste. Gibt (liste, fehler) zurück – fehler=None bei Erfolg.

    Jeder Eintrag: {id, ts, text, kind, erstellt_von, done}.
    """
    base = (base_url or _ov.basis_url()).rstrip("/")
    tok = _token(_firma(firma))
    if not tok:
        return [], "Kein SSO-Geheimnis gefunden."
    url = f"{base}/api/meldungen?token={tok}"
    try:
        with urllib.request.urlopen(url, timeout=timeout) as resp:
            data = json.loads(resp.read().decode("utf-8"))
        return data.get("meldungen", []), None
    except urllib.error.HTTPError as exc:
        return [], f"Server antwortet {exc.code}"
    except urllib.error.URLError as exc:
        return [], f"nicht erreichbar ({exc.reason})"
    except Exception as exc:  # pragma: no cover - defensiv
        return [], f"Fehler: {exc.__class__.__name__}"


def add(text, kind="manuell", by="", key=None,
        base_url: str | None = None, firma: str | None = None, timeout: float = 6.0):
    """Legt eine Meldung/Aufgabe zentral an. Gibt (res, fehler) zurück.
    res z.B. {ok, id, dedupe}. key dedupt offene Einträge (1 Meldung pro Vorgang)."""
    base = (base_url or _ov.basis_url()).rstrip("/")
    tok = _token(_firma(firma))
    if not tok:
        return None, "Kein SSO-Geheimnis gefunden."
    url = f"{base}/api/meldungen?token={tok}"
    payload = {"text": text, "kind": kind, "by": by, "key": key}
    data = json.dumps(payload).encode("utf-8")
    try:
        req = urllib.request.Request(
            url, data=data, method="POST", headers={"Content-Type": "application/json"})
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            return json.loads(resp.read().decode("utf-8")), None
    except urllib.error.HTTPError as exc:
        return None, f"Server antwortet {exc.code}"
    except urllib.error.URLError as exc:
        return None, f"nicht erreichbar ({exc.reason})"
    except Exception as exc:  # pragma: no cover - defensiv
        return None, f"Fehler: {exc.__class__.__name__}"


def _post(path: str, params: dict, firma: str | None, base_url: str | None, timeout: float):
    base = (base_url or _ov.basis_url()).rstrip("/")
    tok = _token(_firma(firma))
    if not tok:
        return None, "Kein SSO-Geheimnis gefunden."
    q = dict(params)
    q["token"] = tok
    url = f"{base}{path}?" + urllib.parse.urlencode(q)
    try:
        req = urllib.request.Request(url, data=b"", method="POST")
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            return json.loads(resp.read().decode("utf-8")), None
    except urllib.error.HTTPError as exc:
        return None, f"Server antwortet {exc.code}"
    except urllib.error.URLError as exc:
        return None, f"nicht erreichbar ({exc.reason})"
    except Exception as exc:  # pragma: no cover - defensiv
        return None, f"Fehler: {exc.__class__.__name__}"


def set_done(meldung_id, done=True, von: str = "",
             base_url: str | None = None, firma: str | None = None, timeout: float = 6.0):
    return _post(f"/api/meldungen/{int(meldung_id)}/done",
                 {"done": 1 if done else 0, "von": von or ""}, firma, base_url, timeout)


def delete(meldung_id, base_url: str | None = None, firma: str | None = None, timeout: float = 6.0):
    return _post(f"/api/meldungen/{int(meldung_id)}/delete", {}, firma, base_url, timeout)
