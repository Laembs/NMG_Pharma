"""Kasse-Modul · Fachlogik (Web-Pilot).

Schema-kompatibel zur künftigen zentralen DB (siehe docs/Plan_Kasse_Web.pdf):
Verkauf/Aufträge, Wareneingang, Lagerbestand (Kern der GDP-Rückverfolgung),
Tagesabschluss und Protokoll. P0 legt nur das Tenant-Schema an und liefert die
Übersichts-Kennzahlen; der Verkaufs-Flow (P2) baut darauf auf.

Die Tabellen tragen die Kassen-Namen aus der Desktop-Kasse (app/kasse_app.py),
damit bestehende SQL-Logik später möglichst unverändert weiterläuft.
"""
from __future__ import annotations

import sqlite3
from datetime import date


def ensure_schema(con: sqlite3.Connection) -> None:
    """Legt die Kassen-Tabellen in der Firmen-DB an (idempotent).

    Wird zu Beginn jeder Kassen-Route aufgerufen, damit das Modul ohne globalen
    Migrationslauf selbsttragend ist (analog zum Personal-Schema in tenancy.py).
    """
    con.execute("""CREATE TABLE IF NOT EXISTS tbl_bestellungen (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        kunde_id INTEGER, kunde_name TEXT,
        datum TEXT, liefertermin TEXT,
        summe REAL DEFAULT 0, status TEXT DEFAULT 'offen',
        erstellt_am TEXT DEFAULT (datetime('now')))""")
    con.execute("""CREATE TABLE IF NOT EXISTS tbl_bestellpositionen (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        bestellung_id INTEGER NOT NULL,
        pzn TEXT, bezeichnung TEXT, menge INTEGER DEFAULT 0,
        einzelpreis REAL DEFAULT 0, rabatt_prozent REAL DEFAULT 0,
        bestellart TEXT)""")
    con.execute("""CREATE TABLE IF NOT EXISTS tbl_wareneingang (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        lieferant TEXT, datum TEXT, beleg TEXT,
        erstellt_am TEXT DEFAULT (datetime('now')))""")
    con.execute("""CREATE TABLE IF NOT EXISTS tbl_wareneingang_positionen (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        wareneingang_id INTEGER NOT NULL,
        pzn TEXT, menge INTEGER DEFAULT 0, charge TEXT, verfall TEXT)""")
    con.execute("""CREATE TABLE IF NOT EXISTS tbl_lagerbestand (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        pzn TEXT, bezeichnung TEXT, charge TEXT, verfall TEXT,
        menge INTEGER DEFAULT 0)""")
    con.execute("""CREATE TABLE IF NOT EXISTS tbl_kasse_tagesabschluss (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        datum TEXT, summe REAL DEFAULT 0, anzahl INTEGER DEFAULT 0,
        abgeschlossen_am TEXT DEFAULT (datetime('now')))""")
    con.execute("""CREATE TABLE IF NOT EXISTS tbl_kasse_log (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        zeit TEXT DEFAULT (datetime('now')),
        benutzer TEXT, aktion TEXT, detail TEXT)""")
    con.commit()


def uebersicht(con: sqlite3.Connection) -> dict:
    """Kennzahlen + letzte Vorgänge für die Kassen-Startseite."""
    heute = date.today().isoformat()

    def _scalar(sql: str, args: tuple = ()) -> float:
        row = con.execute(sql, args).fetchone()
        return row[0] if row and row[0] is not None else 0

    auftraege_heute = _scalar(
        "SELECT COUNT(*) FROM tbl_bestellungen WHERE datum = ?", (heute,))
    offene = _scalar(
        "SELECT COUNT(*) FROM tbl_bestellungen WHERE status = 'offen'")
    umsatz_heute = _scalar(
        "SELECT COALESCE(SUM(summe), 0) FROM tbl_bestellungen WHERE datum = ?",
        (heute,))
    lagerpositionen = _scalar("SELECT COUNT(*) FROM tbl_lagerbestand")

    letzte = [dict(r) for r in con.execute(
        """SELECT id, kunde_name, datum, summe, status
             FROM tbl_bestellungen
            ORDER BY id DESC LIMIT 8""").fetchall()]

    return {
        "auftraege_heute": int(auftraege_heute),
        "offene": int(offene),
        "umsatz_heute": float(umsatz_heute),
        "lagerpositionen": int(lagerpositionen),
        "letzte": letzte,
    }


def lagerbestand(con: sqlite3.Connection) -> list[dict]:
    """Bestand je PZN/Charge/Verfall (Kern der GDP-Rückverfolgung)."""
    return [dict(r) for r in con.execute(
        """SELECT pzn, bezeichnung, charge, verfall, menge
             FROM tbl_lagerbestand
            ORDER BY bezeichnung, verfall""").fetchall()]
