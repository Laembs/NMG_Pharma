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
        # V1.1 SP19: NMG-Rabatte-Historie. Bei jedem Rabatte-Import wird vorher
        # ein Snapshot des aktuellen Stands abgelegt, damit Diff zum letzten
        # Stand und Verlauf pro PZN moeglich sind.
        con.execute(
            """CREATE TABLE IF NOT EXISTS tbl_nmg_rabatte_snapshots(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                erstellt_am TEXT DEFAULT CURRENT_TIMESTAMP,
                quelle TEXT,
                anzahl_eintraege INTEGER DEFAULT 0,
                bemerkung TEXT
            )"""
        )
        con.execute(
            """CREATE TABLE IF NOT EXISTS tbl_nmg_rabatte_historie(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                snapshot_id INTEGER NOT NULL,
                nmg_pzn TEXT NOT NULL,
                artikel TEXT,
                rabatt REAL,
                quelle TEXT,
                letzte_aktualisierung TEXT,
                FOREIGN KEY(snapshot_id) REFERENCES tbl_nmg_rabatte_snapshots(id)
            )"""
        )
        con.execute("CREATE INDEX IF NOT EXISTS idx_nmg_rabatte_historie_pzn ON tbl_nmg_rabatte_historie(nmg_pzn)")
        con.execute("CREATE INDEX IF NOT EXISTS idx_nmg_rabatte_historie_snapshot ON tbl_nmg_rabatte_historie(snapshot_id)")
        actions.append("tbl_nmg_rabatte_snapshots + tbl_nmg_rabatte_historie sichergestellt")

        # Bestell-App: Bestellungen als Kopf + Positionen. Kopf = ein Kunde mit
        # Liefertermin/Bestellart, Positionen = mehrere NMG-Artikelzeilen mit
        # Menge + aufgeloestem Rabatt (PK-Kondition -> NMG -> manuell).
        con.execute(
            """CREATE TABLE IF NOT EXISTS tbl_bestellungen(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                datum TEXT,
                kundennummer TEXT,
                apotheke TEXT,
                bestellart TEXT DEFAULT 'Bestellung',
                lieferzeit TEXT,
                liefertermin TEXT,
                status TEXT DEFAULT 'offen',
                notizen TEXT,
                bearbeiter TEXT,
                erstellt_am TEXT DEFAULT CURRENT_TIMESTAMP,
                geaendert_am TEXT,
                msk_erfasst INTEGER DEFAULT 0,
                msk_von TEXT,
                msk_am TEXT
            )"""
        )
        # MSK-Erfassungsstatus pro Verkauf (muss in MSK eingegeben werden).
        if _table_exists(con, "tbl_bestellungen"):
            bcols = _columns(con, "tbl_bestellungen")
            for col, ddl in (("msk_erfasst", "msk_erfasst INTEGER DEFAULT 0"),
                             ("msk_von", "msk_von TEXT"), ("msk_am", "msk_am TEXT")):
                if col not in bcols:
                    con.execute(f"ALTER TABLE tbl_bestellungen ADD COLUMN {ddl}")
        con.execute(
            """CREATE TABLE IF NOT EXISTS tbl_bestellpositionen(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                bestell_id INTEGER NOT NULL,
                pzn TEXT,
                artikelname TEXT,
                df TEXT,
                pck TEXT,
                apu REAL,
                menge INTEGER DEFAULT 1,
                rabatt_prozent REAL,
                rabatt_quelle TEXT,
                charge TEXT,
                verfall TEXT,
                bestellart TEXT DEFAULT 'Bestellung',
                lieferzeit TEXT,
                liefertermin TEXT,
                FOREIGN KEY(bestell_id) REFERENCES tbl_bestellungen(id) ON DELETE CASCADE
            )"""
        )
        # Charge/Verfall + Liefervorgabe PRO Position: jede Position kann eine
        # eigene Bestellart (Bestellung/Vorbestellung/abgesagt) + Lieferzeit/-termin
        # haben. ALTER-Guard fuer DBs ohne diese Spalten.
        if _table_exists(con, "tbl_bestellpositionen"):
            pcols = _columns(con, "tbl_bestellpositionen")
            for col, ddl in (("charge", "charge TEXT"), ("verfall", "verfall TEXT"),
                             ("bestellart", "bestellart TEXT DEFAULT 'Bestellung'"),
                             ("lieferzeit", "lieferzeit TEXT"), ("liefertermin", "liefertermin TEXT")):
                if col not in pcols:
                    con.execute(f"ALTER TABLE tbl_bestellpositionen ADD COLUMN {ddl}")
        con.execute("CREATE INDEX IF NOT EXISTS idx_bestellungen_datum ON tbl_bestellungen(datum)")
        con.execute("CREATE INDEX IF NOT EXISTS idx_bestellungen_kundennummer ON tbl_bestellungen(kundennummer)")
        con.execute("CREATE INDEX IF NOT EXISTS idx_bestellpositionen_bestell ON tbl_bestellpositionen(bestell_id)")
        con.execute("CREATE INDEX IF NOT EXISTS idx_bestellpositionen_pzn ON tbl_bestellpositionen(pzn)")
        con.execute("CREATE INDEX IF NOT EXISTS idx_bestellpositionen_bestellart ON tbl_bestellpositionen(bestellart)")
        actions.append("tbl_bestellungen + tbl_bestellpositionen sichergestellt")

        # Kasse-App: Wareneingang -> Lagerbestand. Wareneingang bucht NMG-Artikel
        # mit Charge/Verfall/Menge ins Lager; der Verkauf (tbl_bestellungen) zieht
        # spaeter daraus ab. Lagerbestand: eine Zeile pro PZN x Charge x Verfall.
        con.execute(
            """CREATE TABLE IF NOT EXISTS tbl_wareneingang(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                datum TEXT, lieferant TEXT, lieferschein TEXT,
                bearbeiter TEXT, notizen TEXT,
                erstellt_am TEXT DEFAULT CURRENT_TIMESTAMP
            )"""
        )
        con.execute(
            """CREATE TABLE IF NOT EXISTS tbl_wareneingang_positionen(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                we_id INTEGER NOT NULL,
                pzn TEXT, artikelname TEXT, charge TEXT, verfall TEXT,
                menge INTEGER DEFAULT 0, ek REAL,
                FOREIGN KEY(we_id) REFERENCES tbl_wareneingang(id) ON DELETE CASCADE
            )"""
        )
        con.execute(
            """CREATE TABLE IF NOT EXISTS tbl_lagerbestand(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                pzn TEXT, artikelname TEXT, charge TEXT, verfall TEXT,
                menge INTEGER DEFAULT 0, aktualisiert_am TEXT,
                UNIQUE(pzn, charge, verfall)
            )"""
        )
        con.execute("CREATE INDEX IF NOT EXISTS idx_wareneingang_pos_we ON tbl_wareneingang_positionen(we_id)")
        con.execute("CREATE INDEX IF NOT EXISTS idx_lagerbestand_pzn ON tbl_lagerbestand(pzn)")
        actions.append("tbl_wareneingang(_positionen) + tbl_lagerbestand sichergestellt")

        con.execute("INSERT OR REPLACE INTO meta(key,value) VALUES('db_schema_version', ?)", (DB_SCHEMA_VERSION,))
        con.execute("INSERT OR REPLACE INTO meta(key,value) VALUES('last_migration_at', ?)", (datetime.now().isoformat(timespec='seconds'),))
        con.commit()
        return actions
    finally:
        con.close()
