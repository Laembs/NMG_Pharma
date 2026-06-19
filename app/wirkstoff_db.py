"""V1.1 SP2: Wirkstoff- und Staerken-Datenbank.

Die Excel `Artikel_mit_Wirkstoff_und_Staerke_ergaenzt.xlsx` enthaelt pro PZN
einen Wirkstoff (z.B. 'Ramipril') und eine Staerke (z.B. '5MG'). Beides liegt
weder im Artikelstamm noch in den NMG-/Austausch-Tabellen.

Eigene Tabelle `tbl_wirkstoff_staerke` damit:
- bestehende Stammtabellen unangetastet bleiben,
- der Import unabhaengig wiederholbar ist (volle Re-Importe ohne Datenverlust
  woanders),
- Joins per `pzn = pzn` ueber den Primary-Key-Index laufen (keine UDFs in
  WHERE-Klauseln, vgl. V1.1 SP1).
"""

from __future__ import annotations

import re
import sqlite3
from datetime import datetime
from pathlib import Path

from .config import DB_PATH
from .file_loader import load_worksheet


def _pzn(value) -> str:
    """Normalisiert PZN auf 8 Stellen, nur Ziffern. /N-Suffix wird vorher
    abgeschnitten, '.0'-Suffix entfernt. Identisch zur Logik in
    artikel_db._pzn / vergleichssuche._pzn_norm.
    """
    if value is None:
        return ""
    text = str(value).strip()
    if not text:
        return ""
    text = re.sub(r"\s*/\s*N\s*$", "", text, flags=re.IGNORECASE)
    if text.endswith(".0") and text[:-2].isdigit():
        text = text[:-2]
    digits = "".join(ch for ch in text if ch.isdigit())
    if not digits:
        return ""
    return digits.zfill(8)


def _clean(value) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    text = str(value).strip()
    return " ".join(text.split())


def _norm_header(value) -> str:
    text = _clean(value).lower()
    text = (text.replace("ä", "ae").replace("ö", "oe")
                .replace("ü", "ue").replace("ß", "ss"))
    return re.sub(r"[^a-z0-9]+", "", text)


def ensure_wirkstoff_table() -> None:
    """Legt tbl_wirkstoff_staerke samt Indizes an. Idempotent."""
    with sqlite3.connect(DB_PATH) as con:
        con.execute("""
            CREATE TABLE IF NOT EXISTS tbl_wirkstoff_staerke (
                pzn TEXT PRIMARY KEY,
                wirkstoff TEXT,
                staerke TEXT,
                quelle TEXT NOT NULL DEFAULT 'Import',
                importdatum TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
            )
        """)
        con.execute("""
            CREATE INDEX IF NOT EXISTS idx_wirkstoff_staerke_wirkstoff
            ON tbl_wirkstoff_staerke (wirkstoff)
        """)
        con.execute("""
            CREATE INDEX IF NOT EXISTS idx_wirkstoff_staerke_staerke
            ON tbl_wirkstoff_staerke (staerke)
        """)
        con.commit()


_HEADER_ALIASES = {
    "pzn":       {"pzn"},
    "wirkstoff": {"wirkstoff", "wirkstoffergaenzt", "wirkstoffergnzt"},
    "staerke":   {"staerke", "strke", "starke", "dosis", "dosierung"},
}


def _detect_columns(ws) -> dict[str, int]:
    """Findet PZN-, Wirkstoff- und Staerke-Spalte in den ersten Zeilen.
    Liefert {feldname: col_index, 'header_row': N}.
    """
    best_score = -1
    best_map: dict[str, int] = {}
    best_row = 1
    for row in range(1, min(ws.max_row, 15) + 1):
        header_map: dict[str, int] = {}
        for col in range(1, ws.max_column + 1):
            key = _norm_header(ws.cell(row, col).value)
            if key:
                header_map[key] = col

        score = 0
        mapping: dict[str, int] = {}
        for field, aliases in _HEADER_ALIASES.items():
            for alias in aliases:
                if alias in header_map:
                    mapping[field] = header_map[alias]
                    score += 1
                    break

        if score > best_score:
            best_score = score
            best_map = mapping
            best_row = row

    if "pzn" not in best_map:
        raise ValueError("PZN-Spalte nicht erkannt.")
    if "wirkstoff" not in best_map and "staerke" not in best_map:
        raise ValueError("Weder Wirkstoff- noch Staerken-Spalte erkannt.")
    best_map["header_row"] = best_row
    return best_map


def import_wirkstoff_excel(path: str | Path, quelle: str = "Import") -> dict:
    """Importiert die Wirkstoff-/Staerken-Excel in tbl_wirkstoff_staerke.

    Liefert dict mit Statistik:
        gelesen, importiert (eingefuegt), aktualisiert, uebersprungen, fehler.

    UPSERT pro PZN: bestehende Eintraege werden ueberschrieben wenn neue
    Werte vorhanden sind, leere neue Werte ueberschreiben aber NICHT
    bestehende Werte (damit ein Teil-Import bestehende Felder nicht loescht).
    """
    path = Path(path)
    ensure_wirkstoff_table()

    ws, _source_type = load_worksheet(path)
    cols = _detect_columns(ws)
    header_row = cols.pop("header_row")
    pzn_col = cols["pzn"]
    wirkstoff_col = cols.get("wirkstoff")
    staerke_col = cols.get("staerke")

    stats = {"gelesen": 0, "importiert": 0, "aktualisiert": 0,
             "uebersprungen": 0, "fehler": 0}

    importdatum = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    batch: list[tuple[str, str, str, str, str]] = []
    seen_pzns: set[str] = set()

    for row_idx in range(header_row + 1, ws.max_row + 1):
        stats["gelesen"] += 1
        try:
            pzn = _pzn(ws.cell(row_idx, pzn_col).value)
            if not pzn:
                stats["uebersprungen"] += 1
                continue
            wirkstoff = _clean(ws.cell(row_idx, wirkstoff_col).value) if wirkstoff_col else ""
            staerke = _clean(ws.cell(row_idx, staerke_col).value) if staerke_col else ""
            if not wirkstoff and not staerke:
                stats["uebersprungen"] += 1
                continue

            if pzn in seen_pzns:
                stats["uebersprungen"] += 1
                continue
            seen_pzns.add(pzn)

            batch.append((pzn, wirkstoff, staerke, quelle, importdatum))
        except Exception:
            stats["fehler"] += 1

    with sqlite3.connect(DB_PATH) as con:
        vorher = int(con.execute("SELECT COUNT(*) FROM tbl_wirkstoff_staerke").fetchone()[0])
        con.execute("BEGIN")
        con.executemany("""
            INSERT INTO tbl_wirkstoff_staerke (pzn, wirkstoff, staerke, quelle, importdatum)
            VALUES (?, ?, ?, ?, ?)
            ON CONFLICT(pzn) DO UPDATE SET
                wirkstoff = CASE WHEN excluded.wirkstoff <> '' THEN excluded.wirkstoff ELSE tbl_wirkstoff_staerke.wirkstoff END,
                staerke   = CASE WHEN excluded.staerke   <> '' THEN excluded.staerke   ELSE tbl_wirkstoff_staerke.staerke END,
                quelle    = excluded.quelle,
                importdatum = excluded.importdatum
        """, batch)
        con.commit()
        nachher = int(con.execute("SELECT COUNT(*) FROM tbl_wirkstoff_staerke").fetchone()[0])

    stats["importiert"] = max(0, nachher - vorher)
    stats["aktualisiert"] = max(0, len(batch) - stats["importiert"])
    return stats


def wirkstoff_count() -> int:
    """Anzahl Zeilen in tbl_wirkstoff_staerke (fuer Status-Anzeige)."""
    ensure_wirkstoff_table()
    with sqlite3.connect(DB_PATH) as con:
        return int(con.execute("SELECT COUNT(*) FROM tbl_wirkstoff_staerke").fetchone()[0])


def lookup_wirkstoff(pzn: str) -> dict:
    """Schneller PK-Lookup. Liefert {'wirkstoff': str, 'staerke': str}."""
    pzn_n = _pzn(pzn)
    if not pzn_n:
        return {"wirkstoff": "", "staerke": ""}
    ensure_wirkstoff_table()
    with sqlite3.connect(DB_PATH) as con:
        r = con.execute(
            "SELECT wirkstoff, staerke FROM tbl_wirkstoff_staerke WHERE pzn = ? LIMIT 1",
            (pzn_n,),
        ).fetchone()
    if not r:
        return {"wirkstoff": "", "staerke": ""}
    return {"wirkstoff": r[0] or "", "staerke": r[1] or ""}
