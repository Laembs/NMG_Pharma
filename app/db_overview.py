from __future__ import annotations

import sqlite3
from pathlib import Path
from datetime import datetime

from .config import DB_PATH


TABLE_DESCRIPTIONS = {
    "tbl_artikelstamm": {
        "name": "Artikelstamm",
        "zweck": "Zentrale PZN-Stammdatenbank für alle eingelesenen Artikel.",
        "inhalt": "PZN, Artikel, DF, PCK, Herst, Quelle, Zeitstempel",
    },
    "tbl_austauschdatenbank": {
        "name": "Austauschdatenbank",
        "zweck": "Hauptwissensquelle für bekannte Austauschartikel.",
        "inhalt": "PZN alt, PZN NMG, Austauschtext, Quelle, Status",
    },
    "tbl_lernvorschlaege": {
        "name": "Schulbank Lernvorschläge",
        "zweck": "Neue, übernommene und abgelehnte Lernfälle.",
        "inhalt": "Produkt alt, Produkt neu, Status, erstellt/bearbeitet",
    },
    "tbl_lernhistorie": {
        "name": "Lernhistorie",
        "zweck": "Nachvollziehbare Historie gelernter oder bearbeiteter Fälle.",
        "inhalt": "Zeitpunkt, Quelle, PZN, Austausch, Aktion, Status",
    },
    "tbl_pzn_basisdaten": {
        "name": "PZN-Basisdaten",
        "zweck": "Gelernte Stammdaten je PZN aus importierten Rohdaten.",
        "inhalt": "PZN, Artikelname, Hersteller, DF, PCK, Treffer, Quelle",
    },
    "tbl_pzn_basis_stimmen": {
        "name": "PZN-Basis Stimmen",
        "zweck": "Mehrheits-/Stimmenlogik für gelernte PZN-Felder.",
        "inhalt": "PZN, Feld, Wert, Anzahl, Quelle",
    },
    "tbl_hersteller_lern": {
        "name": "Hersteller-Lernstand",
        "zweck": "Gelernter Hersteller je PZN.",
        "inhalt": "PZN, Herstellerkürzel, Treffer, Quelle",
    },
    "tbl_hersteller_stimmen": {
        "name": "Hersteller-Stimmen",
        "zweck": "Mehrheitslogik für Hersteller je PZN.",
        "inhalt": "PZN, Herstellerkürzel, Anzahl, Quelle",
    },
    "tbl_pzn_ek_rohdaten": {
        "name": "EK-Rohdatenhistorie",
        "zweck": "Historie echter EK-Werte aus Rohdatenimporten.",
        "inhalt": "PZN, EK, Quelldatei, EK-Spalte, Importdatum",
    },
    "tbl_nmg_stamm": {
        "name": "NMG-Stammdaten",
        "zweck": "Offizieller NMG-/PK-Artikelstamm.",
        "inhalt": "PZN, Artikelname, Hersteller, APU, Taxe, Menge, Wirkstoffe",
    },
    "nmg_rabatte": {
        "name": "NMG-Rabatte",
        "zweck": "Rabattinformationen zu NMG-Artikeln.",
        "inhalt": "NMG-PZN, Artikel, Rabatt, Quelle",
    },
    "tbl_austauschartikel": {
        "name": "Alte Austauschartikel",
        "zweck": "Bisherige Austauschlogik vor tbl_austauschdatenbank.",
        "inhalt": "Original-PZN, NMG-PZN, Austauschtext, Quelle, Treffer",
    },
    "tbl_lieferfaehigkeit": {
        "name": "Lieferfähigkeit",
        "zweck": "Lieferstatus und Bevorratung zu NMG-Artikeln.",
        "inhalt": "NMG-PZN, lieferbar, Bevorratung, Liefervorschlag",
    },
    "tbl_referenz_h_o": {
        "name": "Referenz H-O",
        "zweck": "Geprüfte Referenzzuordnungen für Auswertungen.",
        "inhalt": "Original-PZN, Sortiment, NMG-PZN, APU, Rabatt, Austausch",
    },
    "tbl_auswertungen": {
        "name": "Gespeicherte Auswertungen",
        "zweck": "Kopfdaten jeder erstellten/importierten Analyse.",
        "inhalt": "Apotheke, Quelldatei, Ausgabedatei, Positionen, Treffer",
    },
    "tbl_auswertungspositionen": {
        "name": "Auswertungspositionen",
        "zweck": "Einzelpositionen aller gespeicherten Analysen.",
        "inhalt": "PZN, Artikel, EK, Absatz, NMG-Zuordnung, Austausch, Umsatz",
    },
    "tbl_rohdaten_mapping": {
        "name": "Rohdaten-Mapping",
        "zweck": "Merkt erkannte Spalten je importierter Datei.",
        "inhalt": "Dateiname, PZN-Spalte, Hersteller-Spalte, EK-Spalte, Absatz-Spalte",
    },
    "tbl_import_log": {
        "name": "Import-Protokoll",
        "zweck": "Protokolliert Datenimporte.",
        "inhalt": "Datei, Typ, Datensätze, Meldung, Datum",
    },
    "tbl_update_log": {
        "name": "Update-Protokoll",
        "zweck": "Protokolliert installierte Updates.",
        "inhalt": "Versionen, Paket, Status, Meldung",
    },
    "tbl_system_log": {
        "name": "System-Protokoll",
        "zweck": "Allgemeine Systemmeldungen.",
        "inhalt": "Bereich, Meldung, Zeitpunkt",
    },
    "meta": {
        "name": "Meta",
        "zweck": "Technische Versions- und Statusinformationen.",
        "inhalt": "Schlüssel/Wert-Paare",
    },
}


SYSTEM_TABLE_PREFIXES = ("sqlite_",)


def _table_exists(con: sqlite3.Connection, table: str) -> bool:
    row = con.execute(
        "SELECT name FROM sqlite_master WHERE type='table' AND name=?",
        (table,),
    ).fetchone()
    return row is not None


def _count_rows(con: sqlite3.Connection, table: str) -> int:
    try:
        return int(con.execute(f'SELECT COUNT(*) FROM "{table}"').fetchone()[0])
    except Exception:
        return 0


def _columns(con: sqlite3.Connection, table: str) -> list[str]:
    try:
        return [str(row[1]) for row in con.execute(f'PRAGMA table_info("{table}")').fetchall()]
    except Exception:
        return []


def ensure_known_overview_tables() -> None:
    """Legt keine Fachdaten neu an, sorgt aber dafür, dass optionale Tabellen sichtbar werden."""
    if not Path(DB_PATH).exists():
        return
    with sqlite3.connect(DB_PATH) as con:
        con.execute(
            """
            CREATE TABLE IF NOT EXISTS tbl_artikelstamm (
                pzn TEXT PRIMARY KEY,
                artikel TEXT,
                df TEXT,
                pck TEXT,
                herst TEXT,
                quelle TEXT,
                treffer INTEGER DEFAULT 1,
                erstellt_am TEXT DEFAULT CURRENT_TIMESTAMP,
                aktualisiert_am TEXT
            )
            """
        )
        con.execute(
            """
            CREATE TABLE IF NOT EXISTS tbl_austauschdatenbank (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                pzn_alt TEXT,
                artikel_alt TEXT,
                pzn_nmg TEXT,
                artikel_nmg TEXT,
                freitext_austausch TEXT NOT NULL DEFAULT '',
                quelle TEXT NOT NULL DEFAULT 'Manuell',
                status TEXT NOT NULL DEFAULT 'aktiv',
                erstellt_am TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
                aktualisiert_am TEXT
            )
            """
        )
        con.execute(
            """
            CREATE TABLE IF NOT EXISTS tbl_lernvorschlaege (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                produkt_alt TEXT NOT NULL,
                produkt_neu TEXT NOT NULL,
                status TEXT NOT NULL DEFAULT 'neu',
                erstellt_am TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
                bearbeitet_am TEXT
            )
            """
        )
        con.commit()


def get_database_overview() -> dict:
    """Liefert eine strukturierte Übersicht aller SQLite-Tabellen."""
    ensure_known_overview_tables()
    db_path = Path(DB_PATH)
    result = {
        "db_path": str(db_path),
        "db_exists": db_path.exists(),
        "db_size_bytes": db_path.stat().st_size if db_path.exists() else 0,
        "created_at": datetime.now().isoformat(timespec="seconds"),
        "tables": [],
        "total_tables": 0,
        "total_rows": 0,
    }
    if not db_path.exists():
        return result

    with sqlite3.connect(db_path) as con:
        rows = con.execute(
            "SELECT name FROM sqlite_master WHERE type='table' ORDER BY name"
        ).fetchall()
        table_names = [r[0] for r in rows if not str(r[0]).startswith(SYSTEM_TABLE_PREFIXES)]
        for table in table_names:
            desc = TABLE_DESCRIPTIONS.get(table, {})
            count = _count_rows(con, table)
            cols = _columns(con, table)
            item = {
                "table": table,
                "display_name": desc.get("name", table),
                "rows": count,
                "purpose": desc.get("zweck", "Noch nicht beschrieben."),
                "content": desc.get("inhalt", ", ".join(cols)),
                "columns": cols,
                "known": table in TABLE_DESCRIPTIONS,
            }
            result["tables"].append(item)
            result["total_rows"] += count
        result["total_tables"] = len(result["tables"])
    return result


def format_size(num_bytes: int) -> str:
    size = float(num_bytes or 0)
    for unit in ["B", "KB", "MB", "GB"]:
        if size < 1024 or unit == "GB":
            return f"{size:.1f} {unit}" if unit != "B" else f"{int(size)} B"
        size /= 1024
    return f"{size:.1f} GB"
