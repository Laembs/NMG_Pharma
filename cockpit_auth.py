"""Cockpit-Login mit lokalem Cache + woechentlicher Pruefung gegen Pennone One.

Quelle der Wahrheit = Pennone One (Plattform-DB, Tabellen firma+benutzer, bcrypt).
Ablauf wie vom User gewuenscht:
  - Beim Login werden die Daten gegen Pennone abgeglichen.
  - Das Ergebnis wird lokal hinterlegt (verschluesselter bcrypt-Hash + Identitaet).
  - Danach wird nur ~1x pro Woche erneut gegen Pennone geprueft; dazwischen laeuft
    der Login lokal (offline-faehig, schnell).
  - Ist Pennone faellig aber nicht erreichbar, greift der lokale Cache (kein
    Aussperren), mit Hinweis.

Pennone One laeuft derzeit lokal; wir lesen die Plattform-DB schreibgeschuetzt
(file:...?mode=ro), damit der laufende uvicorn-Prozess nicht gestoert wird.
Spaeter (zentraler Server) wird hier nur der Zugriffsweg getauscht.
"""
from __future__ import annotations

import os
import json
import sqlite3
from datetime import date

import bcrypt

RECHECK_DAYS = 7
DEFAULT_PENNONE_DB = r"C:\pennone_one\data\platform.sqlite"

# NMGone laeuft immer fuer EINE Firma (NMG-Pharma). Im Login wird daher keine
# Firma abgefragt; der Mandanten-Slug ist fest (per cockpit_config.json
# ueberschreibbar, falls der Slug in Pennone anders heisst).
DEFAULT_FIRMA_SLUG = "nmgpharma"
DEFAULT_FIRMA_NAME = "NMG-Pharma"


class ServerUnreachable(Exception):
    """Pennone-Plattform-DB nicht erreichbar/lesbar (offline / Datei fehlt)."""


# ── Pfade ────────────────────────────────────────────────────────────────────
def _base() -> str:
    try:
        from app.config import USERDATA_ROOT, BASE_DIR
        return str(USERDATA_ROOT) if USERDATA_ROOT else str(BASE_DIR)
    except Exception:
        return os.path.dirname(os.path.abspath(__file__))


def _config_path() -> str:
    return os.path.join(_base(), "cockpit_config.json")


def _cache_path() -> str:
    return os.path.join(_base(), "cockpit_auth_cache.json")


def _config() -> dict:
    try:
        with open(_config_path(), encoding="utf-8") as fh:
            return json.load(fh) or {}
    except Exception:
        return {}


def platform_db_path() -> str:
    """Pfad zur Pennone-Plattform-DB (ueber cockpit_config.json ueberschreibbar)."""
    return _config().get("pennone_platform_db") or DEFAULT_PENNONE_DB


def configured_firma() -> dict:
    """Feste NMGone-Firma (Slug + Anzeigename). Kein Firmen-Abfrage im Login."""
    cfg = _config()
    return {"slug": (cfg.get("firma_slug") or DEFAULT_FIRMA_SLUG).strip().lower(),
            "name": cfg.get("firma_name") or DEFAULT_FIRMA_NAME}


# ── Cache ────────────────────────────────────────────────────────────────────
def _load_cache() -> dict:
    try:
        with open(_cache_path(), encoding="utf-8") as fh:
            return json.load(fh)
    except Exception:
        return {"sessions": {}, "last": None}


def _save_cache(data: dict) -> None:
    try:
        os.makedirs(os.path.dirname(_cache_path()), exist_ok=True)
        with open(_cache_path(), "w", encoding="utf-8") as fh:
            json.dump(data, fh, ensure_ascii=False)
    except Exception:
        pass


def _hash(password: str) -> str:
    return bcrypt.hashpw(password.encode("utf-8")[:72], bcrypt.gensalt()).decode("utf-8")


def _verify(password: str, hashed: str) -> bool:
    try:
        return bcrypt.checkpw(password.encode("utf-8")[:72], hashed.encode("utf-8"))
    except (ValueError, TypeError):
        return False


# ── Server (Pennone) ─────────────────────────────────────────────────────────
def _connect_ro() -> sqlite3.Connection:
    path = platform_db_path()
    if not os.path.exists(path):
        raise ServerUnreachable(f"Pennone-DB nicht gefunden: {path}")
    try:
        con = sqlite3.connect(f"file:{path}?mode=ro", uri=True, timeout=2.0)
        con.row_factory = sqlite3.Row
        return con
    except sqlite3.Error as exc:
        raise ServerUnreachable(str(exc))


def server_authenticate(firma: str, login: str, password: str) -> dict | None:
    """Gegen Pennone pruefen. dict bei Erfolg, None bei falschen Daten.
    ServerUnreachable, wenn Pennone nicht erreichbar ist."""
    con = _connect_ro()
    try:
        row = con.execute(
            """SELECT b.*, f.slug AS firma_slug, f.name AS firma_name
               FROM benutzer b JOIN firma f ON f.id = b.firma_id
               WHERE f.slug = ? AND b.login = ? AND b.aktiv = 1
                 AND f.geloescht_am IS NULL""",
            (firma.strip().lower(), login.strip().lower()),
        ).fetchone()
    except sqlite3.Error as exc:
        raise ServerUnreachable(str(exc))
    finally:
        con.close()
    if row and _verify(password, row["pw_hash"]):
        return {
            "firma": row["firma_slug"], "firma_name": row["firma_name"],
            "login": row["login"], "name": row["name"] or row["login"],
            "rolle": row["rolle"], "superadmin": bool(row["superadmin"] or 0),
            "must_change_pw": bool(row["must_change_pw"] or 0),
        }
    return None


def list_firmen() -> list[dict]:
    """Firmen (slug+name) fuer die Login-Auswahl. Leer, wenn Pennone offline."""
    try:
        con = _connect_ro()
    except ServerUnreachable:
        return []
    try:
        rows = con.execute(
            "SELECT slug, name FROM firma WHERE geloescht_am IS NULL ORDER BY name"
        ).fetchall()
        return [dict(r) for r in rows]
    except sqlite3.Error:
        return []
    finally:
        con.close()


# ── Login (Cache + woechentliche Pruefung) ───────────────────────────────────
def _due(cache_entry: dict | None) -> bool:
    if not cache_entry or not cache_entry.get("last_check"):
        return True
    try:
        last = date.fromisoformat(cache_entry["last_check"])
    except ValueError:
        return True
    return (date.today() - last).days >= RECHECK_DAYS


def login(firma: str, login_name: str, password: str) -> dict:
    """Cockpit-Login. Liefert {ok, source, name, rolle, ..., msg}.

    NMGone hat eine feste Firma: wird keine uebergeben, gilt configured_firma().
    """
    firma = (firma or "").strip().lower() or configured_firma()["slug"]
    login_name = (login_name or "").strip().lower()
    key = f"{firma}/{login_name}"
    cache = _load_cache()
    entry = cache.get("sessions", {}).get(key)

    def _ok(source, ident, msg=""):
        cache["last"] = {"firma": firma, "login": login_name}
        _save_cache(cache)
        return {"ok": True, "source": source, "msg": msg,
                "name": ident.get("name"), "rolle": ident.get("rolle"),
                "superadmin": ident.get("superadmin", False),
                "firma": firma, "login": login_name}

    # 1) Nicht faellig + lokal stimmig -> ohne Server (offline/schnell)
    if not _due(entry) and entry and _verify(password, entry.get("pw_hash", "")):
        return _ok("local", entry)

    # 2) Server pruefen (faellig, neues Geraet, oder lokaler Hash passt nicht)
    try:
        u = server_authenticate(firma, login_name, password)
    except ServerUnreachable as exc:
        # Pennone nicht erreichbar -> lokaler Cache als Notnagel
        if entry and _verify(password, entry.get("pw_hash", "")):
            return _ok("local-offline", entry,
                       msg="Pennone One nicht erreichbar – lokal angemeldet.")
        return {"ok": False, "msg": f"Pennone One nicht erreichbar ({exc}) und keine "
                                    f"lokale Anmeldung vorhanden."}

    if u is None:
        # Server lehnt ab -> evtl. veralteter Cache entfernen
        cache.get("sessions", {}).pop(key, None)
        _save_cache(cache)
        return {"ok": False, "msg": "Firma, Benutzer oder Passwort stimmen nicht."}

    if u.get("must_change_pw"):
        return {"ok": False, "msg": "Bitte zuerst das Passwort in Pennone One ändern, "
                                    "dann hier anmelden."}

    # Erfolg -> Cache auffrischen (frischer lokaler Hash + Identitaet + Datum)
    cache.setdefault("sessions", {})[key] = {
        "pw_hash": _hash(password), "name": u["name"], "rolle": u["rolle"],
        "superadmin": u["superadmin"], "firma_name": u.get("firma_name", ""),
        "last_check": date.today().isoformat(),
    }
    return _ok("server", cache["sessions"][key])


def last_login() -> dict | None:
    """Zuletzt benutzte Firma/Login (zum Vorbefuellen des Formulars)."""
    return _load_cache().get("last")
