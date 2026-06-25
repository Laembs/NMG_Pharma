"""Personal-Modul · reine Logik (Mitarbeiter, Abwesenheiten, Zeit, Akte).

Portiert die CRUD-/Abfrage-Logik aus ``app/personal_app.py`` in tkinter-freie
Funktionen und erweitert sie um Genehmigungs-Workflow, Kalender, Zeiterfassung
und digitale Personalakte. Jede Funktion bekommt eine offene sqlite3.Connection
(row_factory = sqlite3.Row) und gibt Dicts/Listen zurück.
"""
from __future__ import annotations

import calendar as _cal
import sqlite3
from datetime import date, datetime, timedelta

ABW_ARTEN = ("Urlaub", "Sonderurlaub", "Krankheit", "Fortbildung", "Sonstiges")
ABW_STATUS = ("beantragt", "genehmigt", "abgelehnt")
DOK_KATEGORIEN = ("Vertrag", "Zeugnis", "Bescheinigung", "Lohnabrechnung", "Sonstiges")


# ── Mitarbeiter ──────────────────────────────────────────────────────────────
def list_mitarbeiter(con: sqlite3.Connection) -> list[dict]:
    rows = con.execute(
        """SELECT id, vorname, name, abteilung, position,
                  urlaubsanspruch, personalverantwortlich, sollstunden_tag
             FROM tbl_mitarbeiter
            ORDER BY name, vorname""").fetchall()
    return [dict(r) for r in rows]


def get_mitarbeiter(con: sqlite3.Connection, mid: int) -> dict | None:
    row = con.execute(
        "SELECT * FROM tbl_mitarbeiter WHERE id=?", (mid,)).fetchone()
    return dict(row) if row else None


def create_mitarbeiter(con: sqlite3.Connection, *, vorname: str, name: str,
                       abteilung: str = "", position: str = "",
                       urlaubsanspruch: int = 30, sollstunden_tag: float = 8.0,
                       personalverantwortlich: bool = False) -> int:
    cur = con.execute(
        """INSERT INTO tbl_mitarbeiter
               (vorname, name, abteilung, position, urlaubsanspruch,
                sollstunden_tag, personalverantwortlich)
           VALUES (?,?,?,?,?,?,?)""",
        (vorname.strip(), name.strip(), abteilung.strip(), position.strip(),
         int(urlaubsanspruch or 0), float(sollstunden_tag or 0),
         1 if personalverantwortlich else 0))
    con.commit()
    return cur.lastrowid


def update_mitarbeiter(con: sqlite3.Connection, mid: int, *, vorname: str, name: str,
                       abteilung: str = "", position: str = "",
                       urlaubsanspruch: int = 30, sollstunden_tag: float = 8.0,
                       personalverantwortlich: bool = False) -> None:
    con.execute(
        """UPDATE tbl_mitarbeiter
              SET vorname=?, name=?, abteilung=?, position=?,
                  urlaubsanspruch=?, sollstunden_tag=?, personalverantwortlich=?
            WHERE id=?""",
        (vorname.strip(), name.strip(), abteilung.strip(), position.strip(),
         int(urlaubsanspruch or 0), float(sollstunden_tag or 0),
         1 if personalverantwortlich else 0, mid))
    con.commit()


def delete_mitarbeiter(con: sqlite3.Connection, mid: int) -> None:
    con.execute("DELETE FROM tbl_abwesenheit WHERE mitarbeiter_id=?", (mid,))
    con.execute("DELETE FROM tbl_zeiterfassung WHERE mitarbeiter_id=?", (mid,))
    con.execute("DELETE FROM tbl_dokument WHERE mitarbeiter_id=?", (mid,))
    con.execute("DELETE FROM tbl_mitarbeiter_vorgesetzter WHERE mitarbeiter_id=? OR vorgesetzter_id=?", (mid, mid))
    con.execute("DELETE FROM tbl_mitarbeiter WHERE id=?", (mid,))
    con.commit()


# ── Abwesenheiten ────────────────────────────────────────────────────────────
def list_abwesenheiten(con: sqlite3.Connection, mid: int | None = None,
                       status: str | None = None) -> list[dict]:
    sql = """SELECT a.id, a.mitarbeiter_id, a.art, a.von, a.bis, a.notiz, a.unterart,
                    a.status, a.erstellt_am, a.entschieden_am, m.vorname, m.name
               FROM tbl_abwesenheit a
               JOIN tbl_mitarbeiter m ON m.id = a.mitarbeiter_id"""
    where, params = [], []
    if mid is not None:
        where.append("a.mitarbeiter_id=?")
        params.append(mid)
    if status is not None:
        where.append("a.status=?")
        params.append(status)
    if where:
        sql += " WHERE " + " AND ".join(where)
    sql += " ORDER BY a.von DESC, a.id DESC"
    return [dict(r) for r in con.execute(sql, tuple(params)).fetchall()]


def get_abwesenheit(con: sqlite3.Connection, aid: int) -> dict | None:
    row = con.execute("SELECT * FROM tbl_abwesenheit WHERE id=?", (aid,)).fetchone()
    return dict(row) if row else None


def create_abwesenheit(con: sqlite3.Connection, *, mitarbeiter_id: int, art: str,
                       von: str, bis: str, notiz: str = "", unterart: str = "",
                       status: str = "beantragt") -> int:
    if art not in ABW_ARTEN:
        art = "Sonstiges"
    if status not in ABW_STATUS:
        status = "beantragt"
    cur = con.execute(
        """INSERT INTO tbl_abwesenheit
               (mitarbeiter_id, art, von, bis, notiz, unterart, status, erstellt_am)
           VALUES (?,?,?,?,?,?,?, datetime('now'))""",
        (mitarbeiter_id, art, von, bis, notiz.strip(), unterart.strip(), status))
    con.commit()
    return cur.lastrowid


def set_abwesenheit_status(con: sqlite3.Connection, aid: int, status: str) -> None:
    """Antrag genehmigen oder ablehnen (Genehmigungs-Workflow)."""
    if status not in ABW_STATUS:
        return
    con.execute(
        "UPDATE tbl_abwesenheit SET status=?, entschieden_am=datetime('now') WHERE id=?",
        (status, aid))
    con.commit()


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
    """Pro Mitarbeiter: Anspruch, genehmigte und beantragte Urlaubstage, Rest.

    „genommen" zählt nur GENEHMIGTE Urlaubs-/Sonderurlaubstage; „offen" zählt die
    noch nicht entschiedenen Anträge (Status beantragt) für das Jahr.
    """
    jahr = jahr or date.today().year
    out: list[dict] = []
    for m in list_mitarbeiter(con):
        genommen = offen = 0
        for a in list_abwesenheiten(con, m["id"]):
            if a["art"] not in ("Urlaub", "Sonderurlaub"):
                continue
            if not (a["von"] and a["von"][:4] == str(jahr)):
                continue
            tage = _tage_zwischen(a["von"], a["bis"])
            if a["status"] == "genehmigt":
                genommen += tage
            elif a["status"] == "beantragt":
                offen += tage
        anspruch = m["urlaubsanspruch"] or 0
        out.append({
            "id": m["id"],
            "vorname": m["vorname"], "name": m["name"],
            "anspruch": anspruch, "genommen": genommen, "offen": offen,
            "rest": anspruch - genommen,
        })
    return out


# ── Kalender ─────────────────────────────────────────────────────────────────
_MONATE_DE = ("", "Januar", "Februar", "März", "April", "Mai", "Juni", "Juli",
              "August", "September", "Oktober", "November", "Dezember")


def _overlaps(tag: date, von: str, bis: str) -> bool:
    try:
        v = date.fromisoformat(von)
        b = date.fromisoformat(bis)
    except (ValueError, TypeError):
        return False
    return v <= tag <= b


def kalender_monat(con: sqlite3.Connection, jahr: int, monat: int) -> dict:
    """Monats-Kalender: Wochen mit Tagen, je Tag die laufenden Abwesenheiten.

    Nur genehmigte und beantragte Einträge erscheinen (abgelehnte nicht).
    """
    abw = [a for a in list_abwesenheiten(con) if a["status"] in ("genehmigt", "beantragt")]
    wochen: list[list[dict]] = []
    heute = date.today()
    for woche in _cal.Calendar(firstweekday=0).monthdatescalendar(jahr, monat):
        zeile = []
        for tag in woche:
            eintraege = []
            if tag.month == monat:
                for a in abw:
                    if _overlaps(tag, a["von"], a["bis"]):
                        eintraege.append({
                            "name": f'{a["name"]}, {a["vorname"]}',
                            "art": a["art"], "status": a["status"]})
            zeile.append({
                "tag": tag.day, "im_monat": tag.month == monat,
                "ist_heute": tag == heute, "eintraege": eintraege})
        wochen.append(zeile)
    return {
        "jahr": jahr, "monat": monat, "monat_name": _MONATE_DE[monat],
        "wochen": wochen,
        "prev": (jahr - 1, 12) if monat == 1 else (jahr, monat - 1),
        "next": (jahr + 1, 1) if monat == 12 else (jahr, monat + 1),
        "wochentage": ("Mo", "Di", "Mi", "Do", "Fr", "Sa", "So"),
    }


# ── Zeiterfassung ────────────────────────────────────────────────────────────
def _minuten(hhmm: str | None) -> int | None:
    """'08:30' -> 510 Minuten seit Mitternacht; None bei leer/ungültig."""
    if not hhmm:
        return None
    try:
        h, m = hhmm.split(":")[:2]
        return int(h) * 60 + int(m)
    except (ValueError, AttributeError):
        return None


def _ist_minuten(row: dict) -> int:
    """Gearbeitete Minuten eines Eintrags = (geht - kommt) - pause."""
    k, g = _minuten(row.get("kommt")), _minuten(row.get("geht"))
    if k is None or g is None or g < k:
        return 0
    return max(0, g - k - int(row.get("pause_minuten") or 0))


def _hhmm(minuten: int) -> str:
    vz = "-" if minuten < 0 else ""
    minuten = abs(int(minuten))
    return f"{vz}{minuten // 60}:{minuten % 60:02d}"


def list_zeiten(con: sqlite3.Connection, mid: int | None = None,
                jahr: int | None = None, monat: int | None = None) -> list[dict]:
    sql = """SELECT z.id, z.mitarbeiter_id, z.datum, z.kommt, z.geht,
                    z.pause_minuten, z.notiz, m.vorname, m.name, m.sollstunden_tag
               FROM tbl_zeiterfassung z
               JOIN tbl_mitarbeiter m ON m.id = z.mitarbeiter_id"""
    where, params = [], []
    if mid is not None:
        where.append("z.mitarbeiter_id=?")
        params.append(mid)
    if jahr is not None and monat is not None:
        where.append("substr(z.datum,1,7)=?")
        params.append(f"{jahr:04d}-{monat:02d}")
    if where:
        sql += " WHERE " + " AND ".join(where)
    sql += " ORDER BY z.datum DESC, z.id DESC"
    rows = []
    for r in con.execute(sql, tuple(params)).fetchall():
        d = dict(r)
        ist = _ist_minuten(d)
        soll = int(round((d.get("sollstunden_tag") or 0) * 60)) if (d.get("kommt") and d.get("geht")) else 0
        d["ist_text"] = _hhmm(ist) if (d.get("kommt") and d.get("geht")) else "—"
        d["ueber_min"] = ist - soll
        d["ueber_text"] = _hhmm(ist - soll) if (d.get("kommt") and d.get("geht")) else "—"
        rows.append(d)
    return rows


def zeit_monatssumme(con: sqlite3.Connection, jahr: int, monat: int) -> list[dict]:
    """Pro Mitarbeiter: Ist-, Soll- und Überstunden für den Monat."""
    out: list[dict] = []
    for m in list_mitarbeiter(con):
        ist = soll = 0
        for z in list_zeiten(con, m["id"], jahr, monat):
            if z["kommt"] and z["geht"]:
                ist += _ist_minuten(z)
                soll += int(round((m.get("sollstunden_tag") or 0) * 60))
        out.append({
            "id": m["id"], "vorname": m["vorname"], "name": m["name"],
            "ist_text": _hhmm(ist), "soll_text": _hhmm(soll),
            "ueber_min": ist - soll, "ueber_text": _hhmm(ist - soll)})
    return out


def create_zeit(con: sqlite3.Connection, *, mitarbeiter_id: int, datum: str,
                kommt: str = "", geht: str = "", pause_minuten: int = 0,
                notiz: str = "") -> int:
    cur = con.execute(
        """INSERT INTO tbl_zeiterfassung(mitarbeiter_id, datum, kommt, geht, pause_minuten, notiz)
           VALUES (?,?,?,?,?,?)""",
        (mitarbeiter_id, datum, kommt or None, geht or None,
         int(pause_minuten or 0), notiz.strip()))
    con.commit()
    return cur.lastrowid


def stempeln(con: sqlite3.Connection, mitarbeiter_id: int) -> str:
    """Stempeluhr: offener Tag ohne 'geht' -> Gehen setzen; sonst neues Kommen.

    Gibt 'kommen' oder 'gehen' zurück (welche Aktion erfolgte).
    """
    jetzt = datetime.now()
    heute = jetzt.date().isoformat()
    zeit = jetzt.strftime("%H:%M")
    offen = con.execute(
        """SELECT id FROM tbl_zeiterfassung
            WHERE mitarbeiter_id=? AND datum=? AND kommt IS NOT NULL AND geht IS NULL
            ORDER BY id DESC LIMIT 1""", (mitarbeiter_id, heute)).fetchone()
    if offen:
        con.execute("UPDATE tbl_zeiterfassung SET geht=? WHERE id=?", (zeit, offen["id"]))
        con.commit()
        return "gehen"
    con.execute(
        "INSERT INTO tbl_zeiterfassung(mitarbeiter_id, datum, kommt) VALUES (?,?,?)",
        (mitarbeiter_id, heute, zeit))
    con.commit()
    return "kommen"


def delete_zeit(con: sqlite3.Connection, zid: int) -> None:
    con.execute("DELETE FROM tbl_zeiterfassung WHERE id=?", (zid,))
    con.commit()


# ── Digitale Personalakte (Dokument-Metadaten) ───────────────────────────────
def list_dokumente(con: sqlite3.Connection, mid: int) -> list[dict]:
    rows = con.execute(
        """SELECT id, mitarbeiter_id, kategorie, titel, dateiname, ablage,
                  groesse, hochgeladen_am
             FROM tbl_dokument WHERE mitarbeiter_id=?
            ORDER BY hochgeladen_am DESC, id DESC""", (mid,)).fetchall()
    return [dict(r) for r in rows]


def get_dokument(con: sqlite3.Connection, did: int) -> dict | None:
    row = con.execute("SELECT * FROM tbl_dokument WHERE id=?", (did,)).fetchone()
    return dict(row) if row else None


def create_dokument(con: sqlite3.Connection, *, mitarbeiter_id: int, kategorie: str,
                    titel: str, dateiname: str, ablage: str, groesse: int) -> int:
    if kategorie not in DOK_KATEGORIEN:
        kategorie = "Sonstiges"
    cur = con.execute(
        """INSERT INTO tbl_dokument(mitarbeiter_id, kategorie, titel, dateiname, ablage, groesse)
           VALUES (?,?,?,?,?,?)""",
        (mitarbeiter_id, kategorie, titel.strip(), dateiname, ablage, int(groesse or 0)))
    con.commit()
    return cur.lastrowid


def delete_dokument(con: sqlite3.Connection, did: int) -> dict | None:
    """Löscht den Metadatensatz und gibt ihn zurück (Aufrufer löscht die Datei)."""
    dok = get_dokument(con, did)
    if dok:
        con.execute("DELETE FROM tbl_dokument WHERE id=?", (did,))
        con.commit()
    return dok
