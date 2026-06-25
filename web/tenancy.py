"""Mandantenfähigkeit: eine Datenbank pro Firma.

Idee (siehe Plan): Statt einer globalen ``DB_PATH`` (wie in ``app/config.py``)
bekommt jede Lizenz-Firma ihre eigene SQLite-Datei unter ``web/tenants/<slug>/firma.sqlite``.
Eine kleine zentrale **Plattform-DB** (``web/data/platform.sqlite``) hält nur die
Verwaltungsdaten: Firmen, Benutzer (mit Passwort-Hash) und welche Module pro
Firma lizenziert sind.

Die Fachdaten-Tabellen (Mitarbeiter, Abwesenheiten …) sind schema-gleich zur
Desktop-App – das Schema ist aus ``app/personal_app.py:init_db`` abgeleitet, damit
bestehende SQL-Abfragen unverändert weiterlaufen.
"""
from __future__ import annotations

import re
import sqlite3
from pathlib import Path

# ── Pfade ────────────────────────────────────────────────────────────────────
WEB_DIR = Path(__file__).resolve().parent
DATA_DIR = WEB_DIR / "data"
TENANTS_DIR = WEB_DIR / "tenants"
PLATFORM_DB = DATA_DIR / "platform.sqlite"

DATA_DIR.mkdir(parents=True, exist_ok=True)
TENANTS_DIR.mkdir(parents=True, exist_ok=True)


def _slug(name: str) -> str:
    """Macht aus einem Firmennamen einen dateisystem-tauglichen Ordnernamen."""
    s = re.sub(r"[^a-z0-9]+", "-", (name or "").lower()).strip("-")
    return s or "firma"


# ── Plattform-DB (Verwaltung) ────────────────────────────────────────────────
def platform_con() -> sqlite3.Connection:
    con = sqlite3.connect(PLATFORM_DB)
    con.row_factory = sqlite3.Row
    con.execute("PRAGMA foreign_keys = ON")
    return con


def init_platform_db() -> None:
    """Legt die Verwaltungstabellen an (idempotent)."""
    with platform_con() as con:
        con.execute("""CREATE TABLE IF NOT EXISTS firma (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            slug TEXT NOT NULL UNIQUE,
            erstellt_am TEXT DEFAULT (datetime('now')))""")
        con.execute("""CREATE TABLE IF NOT EXISTS benutzer (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            firma_id INTEGER NOT NULL REFERENCES firma(id),
            login TEXT NOT NULL,
            passwort_hash TEXT NOT NULL,
            anzeigename TEXT,
            ist_admin INTEGER DEFAULT 0,
            UNIQUE(firma_id, login))""")
        con.execute("""CREATE TABLE IF NOT EXISTS firma_modul (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            firma_id INTEGER NOT NULL REFERENCES firma(id),
            modul_key TEXT NOT NULL,
            aktiv INTEGER DEFAULT 1,
            gueltig_bis TEXT,
            UNIQUE(firma_id, modul_key))""")
        con.commit()


# ── Firmen-DB (Fachdaten, eine pro Firma) ────────────────────────────────────
def tenant_db_path(slug: str) -> Path:
    p = TENANTS_DIR / slug / "firma.sqlite"
    p.parent.mkdir(parents=True, exist_ok=True)
    return p


def tenant_dir(slug: str) -> Path:
    """Ordner einer Firma – für Datei-Ablage (z. B. Personalakte-Dokumente)."""
    p = TENANTS_DIR / slug
    p.mkdir(parents=True, exist_ok=True)
    return p


def tenant_con(slug: str) -> sqlite3.Connection:
    """Öffnet die Fachdaten-DB einer Firma und stellt das Schema sicher."""
    con = sqlite3.connect(tenant_db_path(slug))
    con.row_factory = sqlite3.Row
    con.execute("PRAGMA foreign_keys = ON")
    _ensure_tenant_schema(con)
    return con


def _add_col(con: sqlite3.Connection, table: str, col: str, ddl: str) -> None:
    """Fügt eine Spalte hinzu, falls sie noch fehlt (einfache Migration)."""
    cols = {r["name"] for r in con.execute(f"PRAGMA table_info({table})").fetchall()}
    if col not in cols:
        con.execute(f"ALTER TABLE {table} ADD COLUMN {ddl}")


def _ensure_tenant_schema(con: sqlite3.Connection) -> None:
    """Personal-Schema – abgeleitet aus app/personal_app.py:init_db, erweitert um
    Genehmigungs-Workflow (Abwesenheiten), Zeiterfassung und Personalakte.

    Bewusst schema-kompatibel gehalten, damit Logik/Queries zwischen Desktop und
    Web austauschbar bleiben. Spaltenerweiterungen laufen über _add_col (idempotent).
    """
    con.execute("""CREATE TABLE IF NOT EXISTS tbl_mitarbeiter (
        id INTEGER PRIMARY KEY AUTOINCREMENT, vorname TEXT, name TEXT,
        abteilung TEXT, position TEXT, board_x INTEGER DEFAULT 60, board_y INTEGER DEFAULT 60,
        urlaubsanspruch INTEGER DEFAULT 30, personalverantwortlich INTEGER DEFAULT 0)""")
    con.execute("""CREATE TABLE IF NOT EXISTS tbl_mitarbeiter_vorgesetzter (
        id INTEGER PRIMARY KEY AUTOINCREMENT, mitarbeiter_id INTEGER NOT NULL,
        vorgesetzter_id INTEGER NOT NULL, art TEXT DEFAULT 'disziplinarisch',
        ist_primaer INTEGER DEFAULT 0)""")
    con.execute("""CREATE TABLE IF NOT EXISTS tbl_abwesenheit (
        id INTEGER PRIMARY KEY AUTOINCREMENT, mitarbeiter_id INTEGER NOT NULL,
        art TEXT DEFAULT 'Urlaub', von TEXT, bis TEXT, notiz TEXT, unterart TEXT)""")

    # Zeiterfassung (Kommt/Geht je Tag, Pause in Minuten)
    con.execute("""CREATE TABLE IF NOT EXISTS tbl_zeiterfassung (
        id INTEGER PRIMARY KEY AUTOINCREMENT, mitarbeiter_id INTEGER NOT NULL,
        datum TEXT NOT NULL, kommt TEXT, geht TEXT,
        pause_minuten INTEGER DEFAULT 0, notiz TEXT)""")

    # Digitale Personalakte (Metadaten; Dateien liegen im Tenant-Ordner)
    con.execute("""CREATE TABLE IF NOT EXISTS tbl_dokument (
        id INTEGER PRIMARY KEY AUTOINCREMENT, mitarbeiter_id INTEGER NOT NULL,
        kategorie TEXT DEFAULT 'Sonstiges', titel TEXT, dateiname TEXT,
        ablage TEXT, groesse INTEGER DEFAULT 0,
        hochgeladen_am TEXT DEFAULT (datetime('now')))""")

    # ── Migrationen für Bestands-DBs ─────────────────────────────────────────
    _add_col(con, "tbl_abwesenheit", "status", "status TEXT DEFAULT 'beantragt'")
    _add_col(con, "tbl_abwesenheit", "entschieden_am", "entschieden_am TEXT")
    _add_col(con, "tbl_abwesenheit", "erstellt_am", "erstellt_am TEXT")
    _add_col(con, "tbl_mitarbeiter", "sollstunden_tag", "sollstunden_tag REAL DEFAULT 8")
    con.commit()


def create_firma(name: str, modules: list[str] | None = None) -> tuple[int, str]:
    """Legt eine Firma + ihre (leere) Fachdaten-DB an und schaltet Module frei.

    Gibt (firma_id, slug) zurück. Idempotent über den Slug.
    """
    slug = _slug(name)
    with platform_con() as con:
        cur = con.execute(
            "INSERT OR IGNORE INTO firma(name, slug) VALUES(?, ?)", (name, slug))
        if cur.lastrowid:
            firma_id = cur.lastrowid
        else:
            firma_id = con.execute(
                "SELECT id FROM firma WHERE slug=?", (slug,)).fetchone()["id"]
        for key in (modules or []):
            con.execute(
                "INSERT OR IGNORE INTO firma_modul(firma_id, modul_key, aktiv) VALUES(?, ?, 1)",
                (firma_id, key))
        con.commit()
    # Fachdaten-DB + Schema sofort anlegen
    tenant_con(slug).close()
    return firma_id, slug
