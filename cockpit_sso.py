"""Cockpit-Seite des ONE-SSO: Einmal-Token bauen + ONE-URL zusammensetzen.

Der Cockpit-Login hat den Benutzer bereits gegen Pennone One geprueft
(cockpit_auth). Klickt der Mitarbeiter danach auf eine Web-Kachel (z. B.
Personal), erzeugt das Cockpit hier ein kurzlebiges, signiertes Token mit
Firma + Login und baut daraus die ONE-Anmelde-URL ``/sso/cockpit?token=...``.
ONE legt damit OHNE erneute Passworteingabe die Web-Session an.

Die Token-Logik ist identisch zu ``pennone/sso.py`` (nur Standardbibliothek),
damit zwischen beiden Programmen keine Zusatz-Abhaengigkeit synchron gehalten
werden muss. Das gemeinsame Geheimnis liegt in einer Datei im Pennone-
Datenverzeichnis, die beide Seiten lesen (lokaler Hybrid-Betrieb auf einer
Maschine); per Env ``PENNONE_SSO_SECRET`` ueberschreibbar (spaeter zentral).
"""
from __future__ import annotations

import base64
import hashlib
import hmac
import json
import os
import secrets
import time
from urllib.parse import urlencode

TTL = 30
DEFAULT_PENNONE_DIR = r"C:\pennone_one"
DEFAULT_ONE_BASE_URL = "http://127.0.0.1:8765"


# ── Konfiguration (cockpit_config.json, von cockpit_auth mitgenutzt) ──────────
def _config() -> dict:
    try:
        import cockpit_auth
        return cockpit_auth._config()  # gleiche cockpit_config.json
    except Exception:
        return {}


def pennone_dir() -> str:
    return _config().get("pennone_dir") or DEFAULT_PENNONE_DIR


def one_base_url() -> str:
    return (_config().get("one_base_url") or DEFAULT_ONE_BASE_URL).rstrip("/")


def sso_secret() -> str:
    """Gemeinsames Geheimnis (Env > Datei im Pennone-Datenverzeichnis).

    Existiert die Datei noch nicht, wird sie erzeugt – egal ob zuerst das Cockpit
    oder zuerst Pennone laeuft, beide landen auf derselben Datei/demselben Wert.
    """
    env = os.environ.get("PENNONE_SSO_SECRET")
    if env:
        return env.strip()
    path = os.path.join(pennone_dir(), "data", "cockpit_sso_secret.txt")
    try:
        if os.path.exists(path):
            with open(path, encoding="utf-8") as fh:
                val = fh.read().strip()
                if val:
                    return val
        os.makedirs(os.path.dirname(path), exist_ok=True)
        val = secrets.token_urlsafe(32)
        with open(path, "w", encoding="utf-8") as fh:
            fh.write(val)
        return val
    except OSError as exc:
        raise RuntimeError(f"SSO-Geheimnis nicht lesbar/schreibbar ({path}): {exc}")


# ── Token (identisch zu pennone/sso.py) ──────────────────────────────────────
def _b64e(b: bytes) -> str:
    return base64.urlsafe_b64encode(b).rstrip(b"=").decode("ascii")


def make_token(firma: str, login: str, ttl: int = TTL) -> str:
    payload = {"f": firma, "l": login,
               "exp": int(time.time()) + ttl, "jti": secrets.token_hex(8)}
    body = _b64e(json.dumps(payload, separators=(",", ":"), sort_keys=True).encode("utf-8"))
    sig = _b64e(hmac.new(sso_secret().encode("utf-8"), body.encode("ascii"),
                         hashlib.sha256).digest())
    return f"{body}.{sig}"


def sso_url(firma: str, login: str, next_path: str | None = None,
            base_url: str | None = None) -> str:
    """Vollstaendige SSO-Anmelde-URL fuer diesen Benutzer.

    ``next_path`` (optionaler lokaler Pfad) lenkt nach der Anmeldung gezielt auf
    einen Bereich, z. B. ``/personal`` fuer die Personal-Kachel.
    ``base_url`` waehlt das Ziel-Programm (Default: ONE). Fuer eine andere
    Pennone-Web-App – z. B. die Kasse unter ``https://nmgkasse.pennone.de`` –
    wird die passende Basis-URL uebergeben; Token/Geheimnis bleiben dieselben,
    sodass jede App mit ``/sso/cockpit``-Endpunkt denselben Login akzeptiert.
    """
    base = (base_url or one_base_url()).rstrip("/")
    params = {"token": make_token(firma, login)}
    if next_path:
        params["next"] = next_path
    return f"{base}/sso/cockpit?{urlencode(params)}"
