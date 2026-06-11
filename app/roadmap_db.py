from __future__ import annotations

import sqlite3
from pathlib import Path
from .config import DB_PATH


VALID_STATUS = {"Idee", "Offen", "Begonnen", "Erledigt"}


def ensure_roadmap_table():
    """Legt die Roadmap-Tabelle an, ohne bestehende Daten zu löschen."""
    with sqlite3.connect(DB_PATH) as con:
        con.execute("""
            CREATE TABLE IF NOT EXISTS tbl_roadmap (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                bereich TEXT NOT NULL DEFAULT 'Sonstiges',
                titel TEXT NOT NULL,
                beschreibung TEXT,
                status TEXT NOT NULL DEFAULT 'Idee',
                prioritaet TEXT NOT NULL DEFAULT 'Normal',
                erstellt_am TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
                geaendert_am TEXT
            )
        """)
        con.execute("""
            CREATE INDEX IF NOT EXISTS idx_roadmap_status
            ON tbl_roadmap(status)
        """)
        con.execute("""
            CREATE INDEX IF NOT EXISTS idx_roadmap_bereich
            ON tbl_roadmap(bereich)
        """)
        con.commit()


def add_roadmap_item(bereich: str, titel: str, beschreibung: str = "", status: str = "Idee", prioritaet: str = "Normal") -> int:
    ensure_roadmap_table()

    bereich = (bereich or "Sonstiges").strip()
    titel = (titel or "").strip()
    beschreibung = (beschreibung or "").strip()
    status = (status or "Idee").strip()
    prioritaet = (prioritaet or "Normal").strip()

    if not titel:
        raise ValueError("Bitte einen Titel für den Wunsch eingeben.")

    if status not in VALID_STATUS:
        status = "Idee"

    with sqlite3.connect(DB_PATH) as con:
        cur = con.execute("""
            INSERT INTO tbl_roadmap (
                bereich,
                titel,
                beschreibung,
                status,
                prioritaet,
                erstellt_am
            )
            VALUES (?, ?, ?, ?, ?, CURRENT_TIMESTAMP)
        """, (bereich, titel, beschreibung, status, prioritaet))
        con.commit()
        return int(cur.lastrowid)


def list_roadmap_items(status: str | None = None):
    ensure_roadmap_table()

    with sqlite3.connect(DB_PATH) as con:
        con.row_factory = sqlite3.Row

        if status:
            return con.execute("""
                SELECT id, bereich, titel, beschreibung, status, prioritaet, erstellt_am, geaendert_am
                FROM tbl_roadmap
                WHERE status = ?
                ORDER BY datetime(COALESCE(geaendert_am, erstellt_am)) DESC, id DESC
            """, (status,)).fetchall()

        return con.execute("""
            SELECT id, bereich, titel, beschreibung, status, prioritaet, erstellt_am, geaendert_am
            FROM tbl_roadmap
            ORDER BY
                CASE status
                    WHEN 'Begonnen' THEN 1
                    WHEN 'Offen' THEN 2
                    WHEN 'Idee' THEN 3
                    WHEN 'Erledigt' THEN 4
                    ELSE 9
                END,
                datetime(COALESCE(geaendert_am, erstellt_am)) DESC,
                id DESC
        """).fetchall()


def update_roadmap_status(item_id: int, status: str):
    ensure_roadmap_table()
    status = (status or "").strip()
    if status not in VALID_STATUS:
        raise ValueError("Ungültiger Roadmap-Status.")

    with sqlite3.connect(DB_PATH) as con:
        con.execute("""
            UPDATE tbl_roadmap
            SET status = ?,
                geaendert_am = CURRENT_TIMESTAMP
            WHERE id = ?
        """, (status, int(item_id)))
        con.commit()


def seed_default_roadmap_items():
    """Legt Startpunkte nur an, wenn die Tabelle noch leer ist."""
    ensure_roadmap_table()
    with sqlite3.connect(DB_PATH) as con:
        count = con.execute("SELECT COUNT(*) FROM tbl_roadmap").fetchone()[0]
        if count:
            return 0

    defaults = [
        ("Import", "Universeller Dateiimport testen", "xlsx, xlsm, csv und txt für Neue Auswertung und alle Datenimporte sauber testen.", "Begonnen", "Hoch"),
        ("Datenbank", "Austauschdatenbank vollständig integrieren", "Import, Anzeige, Anzahl, Suche und spätere Analyse-Anbindung.", "Offen", "Hoch"),
        ("Analyse", "Analyse auf Artikelstamm + Austauschdatenbank umstellen", "Reihenfolge: Artikelstamm prüfen, Austauschdatenbank prüfen, nur unbekannte Fälle an Schulbank.", "Offen", "Hoch"),
        ("Schulbank", "Schulbank produktiv machen", "Lernen, Nicht lernen und Rückgängig mit dauerhafter Historie.", "Offen", "Hoch"),
        ("Kundenverwaltung", "Kundendatenbank prüfen", "Kundenstammdaten, Eingabemaske, Excel-Synchronisierung und spätere Massenmail-Funktion.", "Idee", "Normal"),
        ("E-Mail", "Massenmail-Funktion prüfen", "Serienmails an Kundengruppen mit Versandhistorie und Datenschutzprüfung.", "Idee", "Normal"),
        ("Datenbank", "Datenbankübersicht eingebaut", "Übersicht über Tabellen, Datensätze, Spalten und Zweck.", "Erledigt", "Normal"),
        ("Datenbank", "Artikelstamm eingebaut", "tbl_artikelstamm mit Import und 161.811 Datensätzen.", "Erledigt", "Hoch"),
    ]

    for bereich, titel, beschreibung, status, prioritaet in defaults:
        add_roadmap_item(bereich, titel, beschreibung, status, prioritaet)

    return len(defaults)
