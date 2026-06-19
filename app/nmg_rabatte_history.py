"""V1.1 SP19: Snapshot/Diff/Verlauf fuer nmg_rabatte.

Bei jedem Rabatte-Import wird der aktuelle Stand von nmg_rabatte vorher 1:1
in tbl_nmg_rabatte_historie kopiert (mit gemeinsamer snapshot_id). Damit
lassen sich Diff (neueste vs. vorletzter Snapshot) und Verlauf pro PZN
abfragen.
"""
from __future__ import annotations

import sqlite3
from typing import Any

from .config import DB_PATH


def take_snapshot(con: sqlite3.Connection, quelle: str = "", bemerkung: str = "") -> int:
    """Legt einen Snapshot des aktuellen nmg_rabatte-Stands an.

    Liefert die snapshot_id. Wenn nmg_rabatte leer ist, wird trotzdem ein
    Snapshot-Kopf erstellt (anzahl_eintraege=0).
    """
    cur = con.execute(
        "INSERT INTO tbl_nmg_rabatte_snapshots(quelle, bemerkung) VALUES(?, ?)",
        (quelle or "", bemerkung or ""),
    )
    snapshot_id = int(cur.lastrowid)
    try:
        rows = con.execute(
            "SELECT nmg_pzn, artikel, rabatt, quelle, letzte_aktualisierung FROM nmg_rabatte"
        ).fetchall()
    except sqlite3.Error:
        rows = []
    if rows:
        con.executemany(
            """INSERT INTO tbl_nmg_rabatte_historie
               (snapshot_id, nmg_pzn, artikel, rabatt, quelle, letzte_aktualisierung)
               VALUES (?, ?, ?, ?, ?, ?)""",
            [(snapshot_id, r[0], r[1], r[2], r[3], r[4]) for r in rows],
        )
    con.execute(
        "UPDATE tbl_nmg_rabatte_snapshots SET anzahl_eintraege=? WHERE id=?",
        (len(rows), snapshot_id),
    )
    con.commit()
    return snapshot_id


def list_snapshots(con: sqlite3.Connection, limit: int = 50) -> list[dict]:
    rows = con.execute(
        """SELECT id, erstellt_am, quelle, anzahl_eintraege, bemerkung
           FROM tbl_nmg_rabatte_snapshots
           ORDER BY id DESC LIMIT ?""",
        (limit,),
    ).fetchall()
    return [
        {"id": r[0], "erstellt_am": r[1], "quelle": r[2], "anzahl_eintraege": r[3], "bemerkung": r[4]}
        for r in rows
    ]


def latest_snapshot_id(con: sqlite3.Connection) -> int | None:
    row = con.execute("SELECT id FROM tbl_nmg_rabatte_snapshots ORDER BY id DESC LIMIT 1").fetchone()
    return int(row[0]) if row else None


def current_stats(con: sqlite3.Connection) -> dict:
    """Statistik ueber den aktuellen nmg_rabatte-Stand."""
    try:
        row = con.execute(
            """SELECT COUNT(*), MIN(rabatt), MAX(rabatt), AVG(rabatt),
                      MAX(letzte_aktualisierung)
               FROM nmg_rabatte
               WHERE rabatt IS NOT NULL"""
        ).fetchone()
    except sqlite3.Error:
        return {"anzahl": 0, "min": None, "max": None, "avg": None, "letzte_aktualisierung": None}
    return {
        "anzahl": int(row[0] or 0),
        "min": row[1],
        "max": row[2],
        "avg": row[3],
        "letzte_aktualisierung": row[4],
    }


def diff_against_snapshot(con: sqlite3.Connection, snapshot_id: int | None = None) -> dict:
    """Diff: aktueller nmg_rabatte gegen einen Snapshot.

    snapshot_id=None nimmt den letzten verfuegbaren Snapshot.
    Liefert dict mit Listen: neu, geaendert, entfernt, unveraendert_count.
    """
    if snapshot_id is None:
        snapshot_id = latest_snapshot_id(con)
    if snapshot_id is None:
        return {"snapshot_id": None, "neu": [], "geaendert": [], "entfernt": [], "unveraendert_count": 0}

    current = {
        r[0]: {"nmg_pzn": r[0], "artikel": r[1], "rabatt": r[2], "quelle": r[3], "letzte_aktualisierung": r[4]}
        for r in con.execute(
            "SELECT nmg_pzn, artikel, rabatt, quelle, letzte_aktualisierung FROM nmg_rabatte"
        ).fetchall()
    }
    old = {
        r[0]: {"nmg_pzn": r[0], "artikel": r[1], "rabatt": r[2], "quelle": r[3], "letzte_aktualisierung": r[4]}
        for r in con.execute(
            """SELECT nmg_pzn, artikel, rabatt, quelle, letzte_aktualisierung
               FROM tbl_nmg_rabatte_historie WHERE snapshot_id=?""",
            (snapshot_id,),
        ).fetchall()
    }

    neu = [current[pzn] for pzn in current if pzn not in old]
    entfernt = [old[pzn] for pzn in old if pzn not in current]
    geaendert = []
    unveraendert_count = 0
    for pzn in current:
        if pzn not in old:
            continue
        a, b = old[pzn], current[pzn]
        if _diff_rabatt(a["rabatt"], b["rabatt"]):
            geaendert.append({
                "nmg_pzn": pzn,
                "artikel": b["artikel"] or a["artikel"],
                "rabatt_alt": a["rabatt"],
                "rabatt_neu": b["rabatt"],
                "quelle": b["quelle"],
            })
        else:
            unveraendert_count += 1
    return {
        "snapshot_id": snapshot_id,
        "neu": neu,
        "geaendert": geaendert,
        "entfernt": entfernt,
        "unveraendert_count": unveraendert_count,
    }


def _diff_rabatt(a: Any, b: Any) -> bool:
    try:
        return abs(float(a or 0) - float(b or 0)) > 1e-9
    except (TypeError, ValueError):
        return str(a or "") != str(b or "")


def history_for_pzn(con: sqlite3.Connection, nmg_pzn: str) -> list[dict]:
    """Verlauf eines Rabatts ueber alle Snapshots + aktueller Stand am Ende."""
    rows = con.execute(
        """SELECT s.id, s.erstellt_am, s.quelle, h.artikel, h.rabatt, h.letzte_aktualisierung
           FROM tbl_nmg_rabatte_historie h
           JOIN tbl_nmg_rabatte_snapshots s ON s.id = h.snapshot_id
           WHERE h.nmg_pzn=?
           ORDER BY s.id ASC""",
        (str(nmg_pzn),),
    ).fetchall()
    verlauf = [
        {
            "snapshot_id": r[0],
            "erstellt_am": r[1],
            "snapshot_quelle": r[2],
            "artikel": r[3],
            "rabatt": r[4],
            "letzte_aktualisierung": r[5],
        }
        for r in rows
    ]
    # Aktueller Stand als letzte Zeile (snapshot_id=None signalisiert "live").
    try:
        row = con.execute(
            "SELECT artikel, rabatt, quelle, letzte_aktualisierung FROM nmg_rabatte WHERE nmg_pzn=?",
            (str(nmg_pzn),),
        ).fetchone()
    except sqlite3.Error:
        row = None
    if row:
        verlauf.append({
            "snapshot_id": None,
            "erstellt_am": "(aktuell)",
            "snapshot_quelle": row[2],
            "artikel": row[0],
            "rabatt": row[1],
            "letzte_aktualisierung": row[3],
        })
    return verlauf
