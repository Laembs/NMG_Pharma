from __future__ import annotations
import sqlite3
from datetime import datetime
from pathlib import Path
from .config import DB_PATH

DB_SCHEMA_VERSION = "1.1"


def _columns(con: sqlite3.Connection, table: str) -> set[str]:
    try:
        return {row[1] for row in con.execute(f"PRAGMA table_info({table})").fetchall()}
    except sqlite3.Error:
        return set()


def _table_exists(con: sqlite3.Connection, table: str) -> bool:
    row = con.execute("SELECT name FROM sqlite_master WHERE type='table' AND name=?", (table,)).fetchone()
    return row is not None


def run_migrations(db_path: Path = DB_PATH) -> list[str]:
    """Führt sichere Datenbank-Migrationen aus, ohne bestehende Daten zu löschen."""
    actions: list[str] = []
    if not Path(db_path).exists():
        return actions

    con = sqlite3.connect(db_path)
    try:
        con.execute("PRAGMA foreign_keys = ON")
        con.execute("CREATE TABLE IF NOT EXISTS meta(key TEXT PRIMARY KEY, value TEXT NOT NULL)")

        if _table_exists(con, "tbl_auswertungen"):
            cols = _columns(con, "tbl_auswertungen")
            if "datenquelle" not in cols:
                con.execute("ALTER TABLE tbl_auswertungen ADD COLUMN datenquelle TEXT DEFAULT 'NMG'")
                actions.append("tbl_auswertungen.datenquelle ergänzt")
            if "programm_version" not in cols:
                con.execute("ALTER TABLE tbl_auswertungen ADD COLUMN programm_version TEXT")
                actions.append("tbl_auswertungen.programm_version ergänzt")
            con.execute("UPDATE tbl_auswertungen SET datenquelle='NMG' WHERE datenquelle IS NULL OR datenquelle='' ")

        if _table_exists(con, "tbl_auswertungspositionen"):
            cols = _columns(con, "tbl_auswertungspositionen")
            if "datenquelle" not in cols:
                con.execute("ALTER TABLE tbl_auswertungspositionen ADD COLUMN datenquelle TEXT DEFAULT 'NMG'")
                actions.append("tbl_auswertungspositionen.datenquelle ergänzt")
            con.execute("UPDATE tbl_auswertungspositionen SET datenquelle='NMG' WHERE datenquelle IS NULL OR datenquelle='' ")

        con.execute(
            """CREATE TABLE IF NOT EXISTS tbl_update_log(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                zeitpunkt TEXT DEFAULT CURRENT_TIMESTAMP,
                von_version TEXT,
                nach_version TEXT,
                paket TEXT,
                status TEXT,
                meldung TEXT
            )"""
        )
        con.execute(
            """CREATE TABLE IF NOT EXISTS tbl_system_log(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                zeitpunkt TEXT DEFAULT CURRENT_TIMESTAMP,
                bereich TEXT,
                meldung TEXT
            )"""
        )
        con.execute("INSERT OR REPLACE INTO meta(key,value) VALUES('db_schema_version', ?)", (DB_SCHEMA_VERSION,))
        con.execute("INSERT OR REPLACE INTO meta(key,value) VALUES('last_migration_at', ?)", (datetime.now().isoformat(timespec='seconds'),))
        con.commit()
        return actions
    finally:
        con.close()
