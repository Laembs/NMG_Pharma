"""Personal-Modul · reine Logik (Mitarbeiter + Abwesenheiten).

Portiert die CRUD-/Abfrage-Logik aus ``app/personal_app.py`` in tkinter-freie
Funktionen. Jede Funktion bekommt eine offene sqlite3.Connection (row_factory =
sqlite3.Row) und gibt Dicts/Listen zurück.
"""
from __future__ import annotations

import sqlite3
from datetime import date

ABW_ARTEN = ("Urlaub", "Sonderurlaub", "Krankheit", "Fortbildung", "Sonstiges")


# ── Mitarbeiter ──────────────────────────────────────────────────────────────
def list_mitarbeiter(con: sqlite3.Connection) -> list[dict]:
    rows = con.execute(
        """SELECT id, vorname, name, abteilung, position,
                  urlaubsanspruch, personalverantwortlich
             FROM tbl_mitarbeiter
            ORDER BY name, vorname""").fetchall()
    return [dict(r) for r in rows]


def get_mitarbeiter(con: sqlite3.Connection, mid: int) -> dict | None:
    row = con.execute(
        "SELECT * FROM tbl_mitarbeiter WHERE id=?", (mid,)).fetchone()
    return dict(row) if row else None


def create_mitarbeiter(con: sqlite3.Connection, *, vorname: str, name: str,
                       abteilung: str = "", position: str = "",
                       urlaubsanspruch: int = 30,
                       personalverantwortlich: bool = False) -> int:
    cur = con.execute(
        """INSERT INTO tbl_mitarbeiter
               (vorname, name, abteilung, position, urlaubsanspruch, personalverantwortlich)
           VALUES (?,?,?,?,?,?)""",
        (vorname.strip(), name.strip(), abteilung.strip(), position.strip(),
         int(urlaubsanspruch or 0), 1 if personalverantwortlich else 0))
    con.commit()
    return cur.lastrowid


def update_mitarbeiter(con: sqlite3.Connection, mid: int, *, vorname: str, name: str,
                       abteilung: str = "", position: str = "",
                       urlaubsanspruch: int = 30,
                       personalverantwortlich: bool = False) -> None:
    con.execute(
        """UPDATE tbl_mitarbeiter
              SET vorname=?, name=?, abteilung=?, position=?,
                  urlaubsanspruch=?, personalverantwortlich=?
            WHERE id=?""",
        (vorname.strip(), name.strip(), abteilung.strip(), position.strip(),
         int(urlaubsanspruch or 0), 1 if personalverantwortlich else 0, mid))
    con.commit()


def delete_mitarbeiter(con: sqlite3.Connection, mid: int) -> None:
    con.execute("DELETE FROM tbl_abwesenheit WHERE mitarbeiter_id=?", (mid,))
    con.execute("DELETE FROM tbl_mitarbeiter_vorgesetzter WHERE mitarbeiter_id=? OR vorgesetzter_id=?", (mid, mid))
    con.execute("DELETE FROM tbl_mitarbeiter WHERE id=?", (mid,))
    con.commit()


# ── Abwesenheiten ────────────────────────────────────────────────────────────
def list_abwesenheiten(con: sqlite3.Connection, mid: int | None = None) -> list[dict]:
    sql = """SELECT a.id, a.mitarbeiter_id, a.art, a.von, a.bis, a.notiz, a.unterart,
                    m.vorname, m.name
               FROM tbl_abwesenheit a
               JOIN tbl_mitarbeiter m ON m.id = a.mitarbeiter_id"""
    params: tuple = ()
    if mid is not None:
        sql += " WHERE a.mitarbeiter_id=?"
        params = (mid,)
    sql += " ORDER BY a.von DESC, a.id DESC"
    return [dict(r) for r in con.execute(sql, params).fetchall()]


def create_abwesenheit(con: sqlite3.Connection, *, mitarbeiter_id: int, art: str,
                       von: str, bis: str, notiz: str = "", unterart: str = "") -> int:
    if art not in ABW_ARTEN:
        art = "Sonstiges"
    cur = con.execute(
        """INSERT INTO tbl_abwesenheit(mitarbeiter_id, art, von, bis, notiz, unterart)
           VALUES (?,?,?,?,?,?)""",
        (mitarbeiter_id, art, von, bis, notiz.strip(), unterart.strip()))
    con.commit()
    return cur.lastrowid


def delete_abwesenheit(con: sqlite3.Connection, aid: int) -> None:
    con.execute("DELETE FROM tbl_abwesenheit WHERE id=?", (aid,))
    con.commit()


def _tage_zwischen(von: str, bis: str) -> int:
    """Kalendertage inkl. Start- und Endtag (einfache Zählung, ohne Feiertage)."""
    try:
        v = date.fromisoformat(von)
        b = date.fromisoformat(bis)
    except (ValueError, TypeError):
        return 0
    return max(0, (b - v).days + 1)


def urlaubsuebersicht(con: sqlite3.Connection, jahr: int | None = None) -> list[dict]:
    """Pro Mitarbeiter: Anspruch, genommene Urlaubstage, Rest – für ein Jahr."""
    jahr = jahr or date.today().year
    out: list[dict] = []
    for m in list_mitarbeiter(con):
        genommen = 0
        for a in list_abwesenheiten(con, m["id"]):
            if a["art"] not in ("Urlaub", "Sonderurlaub"):
                continue
            if a["von"] and a["von"][:4] == str(jahr):
                genommen += _tage_zwischen(a["von"], a["bis"])
        anspruch = m["urlaubsanspruch"] or 0
        out.append({
            "id": m["id"],
            "vorname": m["vorname"], "name": m["name"],
            "anspruch": anspruch, "genommen": genommen,
            "rest": anspruch - genommen,
        })
    return out
