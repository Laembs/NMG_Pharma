"""Geteilte Helfer fuer den Nutzerdaten-Pfad und die geteilte Cockpit-Aufgaben-
und Meldungsliste.

Die Aufgaben/Meldungen liegen in der gemeinsamen SQLite-DB
(``tbl_cockpit_meldungen``) – NICHT mehr in einer lokalen JSON-Datei. Damit
sind sie an allen Arbeitsplaetzen sichtbar, sobald die DB zentral/geteilt liegt,
und gleichzeitige Schreiber (mehrere PCs) kollidieren nicht (WAL +
busy_timeout). Fach-Apps haengen ueber :func:`push_cockpit_todo` System-
Meldungen an (z.B. "Warenausgang gemeldet"); das Cockpit zeigt manuelle und
System-Eintraege in derselben Liste.

WICHTIG (Multiuser): "sichtbar an allen Arbeitsplaetzen" setzt voraus, dass
alle Arbeitsplaetze auf DIESELBE DB zeigen (data_root / NMGONE_DATA_ROOT auf
einen gemeinsamen Pfad bzw. spaeter ein zentraler DB-Server). Liegt je PC eine
eigene lokale DB, bleibt jede Liste lokal – unabhaengig von diesem Code.

Bewusst ohne GUI-/tk-Abhaengigkeit, damit Cockpit und Fach-Apps ohne
Import-Zyklus darauf zugreifen koennen.
"""
import os
import sqlite3
from datetime import datetime


def userdata_base():
    """Basisverzeichnis fuer Nutzerdaten (gleiche Quelle wie das Cockpit)."""
    try:
        from app.config import USERDATA_ROOT, BASE_DIR
        return str(USERDATA_ROOT) if USERDATA_ROOT else str(BASE_DIR)
    except Exception:
        return os.path.abspath(".")


def cockpit_todos_path():
    """Pfad der alten JSON-Aufgabenliste (nur noch fuer die Einmal-Migration)."""
    return os.path.join(userdata_base(), "cockpit_todos.json")


def _db_path():
    from app.config import DB_PATH
    return str(DB_PATH)


def _conn():
    con = sqlite3.connect(_db_path(), timeout=30)
    con.execute("PRAGMA busy_timeout=30000")
    try:
        con.execute("PRAGMA journal_mode=WAL")
    except Exception:
        pass
    return con


def ensure_meldungen_table(con):
    con.execute(
        """CREATE TABLE IF NOT EXISTS tbl_cockpit_meldungen(
               id           INTEGER PRIMARY KEY AUTOINCREMENT,
               ts           TEXT,
               text         TEXT NOT NULL,
               kind         TEXT DEFAULT 'manuell',
               erstellt_von TEXT,
               done         INTEGER DEFAULT 0,
               done_von     TEXT,
               done_am      TEXT,
               dedupe_key   TEXT
           )""")
    _migrate_json_if_needed(con)


def _retire_json(path):
    """Alte JSON-Datei nach erfolgreicher Migration wegraeumen (nicht loeschen)."""
    try:
        os.replace(path, path + ".migrated")
    except Exception:
        pass


def _migrate_json_if_needed(con):
    """Einmalige Uebernahme alter JSON-Aufgaben in die DB (nichts verlieren)."""
    p = cockpit_todos_path()
    if not os.path.exists(p):
        return
    try:
        n = con.execute("SELECT COUNT(*) FROM tbl_cockpit_meldungen").fetchone()[0]
        if n:
            # Tabelle bereits befuellt -> nicht doppelt importieren, nur wegraeumen.
            _retire_json(p)
            return
        import json
        with open(p, encoding="utf-8") as fh:
            data = json.load(fh)
        if isinstance(data, list):
            for t in data:
                ts = t.get("ts") or datetime.now().strftime("%d.%m. %H:%M")
                con.execute(
                    "INSERT INTO tbl_cockpit_meldungen(ts,text,kind,erstellt_von,done,dedupe_key) "
                    "VALUES(?,?,?,?,?,?)",
                    (ts, str(t.get("text", ""))[:200], t.get("kind", "manuell"),
                     t.get("by", ""), 1 if t.get("done") else 0, t.get("key")))
            con.commit()
        _retire_json(p)
    except Exception:
        # Best-Effort: schlaegt die Migration fehl, bleibt die JSON liegen.
        pass


def _fmt_ts(ts):
    if not ts:
        return ""
    try:
        return datetime.fromisoformat(ts).strftime("%d.%m. %H:%M")
    except Exception:
        return str(ts)


def _row_to_dict(row):
    cid, ts, text, kind, by, done = row
    return {"id": cid, "text": text, "kind": kind or "manuell",
            "by": by or "", "ts": _fmt_ts(ts), "done": bool(done)}


def _local_list_todos():
    """Alle Aufgaben/Meldungen fuer das Cockpit (offene zuerst, dann neueste)."""
    try:
        with _conn() as con:
            ensure_meldungen_table(con)
            rows = con.execute(
                "SELECT id, ts, text, kind, erstellt_von, done "
                "FROM tbl_cockpit_meldungen ORDER BY done ASC, id DESC").fetchall()
        return [_row_to_dict(r) for r in rows]
    except Exception:
        return []


def _local_add_todo(text, by=""):
    """Manuelle Aufgabe (vom Cockpit-Eingabefeld) anlegen."""
    text = (text or "").strip()
    if not text:
        return False
    try:
        with _conn() as con:
            ensure_meldungen_table(con)
            con.execute(
                "INSERT INTO tbl_cockpit_meldungen(ts,text,kind,erstellt_von,done) "
                "VALUES(?,?, 'manuell', ?, 0)",
                (datetime.now().isoformat(timespec="seconds"), text[:200], by))
            con.commit()
        return True
    except Exception:
        return False


def _local_toggle_todo(todo_id, by=""):
    try:
        with _conn() as con:
            ensure_meldungen_table(con)
            row = con.execute("SELECT done FROM tbl_cockpit_meldungen WHERE id=?",
                              (todo_id,)).fetchone()
            if row:
                neu = 0 if row[0] else 1
                con.execute(
                    "UPDATE tbl_cockpit_meldungen SET done=?, done_von=?, done_am=? WHERE id=?",
                    (neu, by if neu else None,
                     datetime.now().isoformat(timespec="seconds") if neu else None, todo_id))
                con.commit()
    except Exception:
        pass


def _local_delete_todo(todo_id):
    try:
        with _conn() as con:
            ensure_meldungen_table(con)
            con.execute("DELETE FROM tbl_cockpit_meldungen WHERE id=?", (todo_id,))
            con.commit()
    except Exception:
        pass


def _local_push_cockpit_todo(text, *, kind="system", by="System", key=None):
    """Haengt eine System-Meldung an die LOKALE Cockpit-Tabelle an (Fallback).

    ``key`` (optional) verhindert Doppelmeldungen: existiert bereits ein noch
    offener Eintrag mit demselben ``key``, wird nichts hinzugefuegt. So entsteht
    pro Vorgang (z.B. pro Warenausgang) genau EINE Meldung, auch wenn die
    ausloesende Aktion mehrere Positionen umfasst.

    Gibt True zurueck, wenn ein neuer Eintrag angelegt wurde.
    """
    text = (text or "").strip()
    if not text:
        return False
    try:
        with _conn() as con:
            ensure_meldungen_table(con)
            if key is not None:
                r = con.execute(
                    "SELECT 1 FROM tbl_cockpit_meldungen WHERE dedupe_key=? AND done=0 LIMIT 1",
                    (key,)).fetchone()
                if r:
                    return False
            con.execute(
                "INSERT INTO tbl_cockpit_meldungen(ts,text,kind,erstellt_von,done,dedupe_key) "
                "VALUES(?,?,?,?,0,?)",
                (datetime.now().isoformat(timespec="seconds"), text[:200], kind, by, key))
            con.commit()
        return True
    except Exception:
        return False


# ── Dispatcher: zentral (Hetzner) wenn konfiguriert, sonst lokale DB ──────────
# Ist der Online-Dienst konfiguriert (SSO-Geheimnis vorhanden), laeuft die
# geteilte Liste ueber nmgkasse.pennone.de – damit an ALLEN Arbeitsplaetzen
# sichtbar (wie die Web-Kasse-Verkaeufe). Faellt der Dienst aus, wird der lokale
# DB-Stand als Fallback benutzt. Per env NMGONE_MELDUNGEN_LOCAL=1 erzwingbar lokal.
#
# Bekannte Grenze (Phase 1): schlaegt ein Schreibvorgang online fehl, landet er
# lokal und wird (noch) nicht automatisch nachsynchronisiert -> Outbox/Resync ist
# der dokumentierte Folgeschritt.
def _online():
    if os.environ.get("NMGONE_MELDUNGEN_LOCAL", "").strip() in ("1", "true", "True", "ja"):
        return None
    try:
        from app import online_meldungen as om
        if om.is_configured():
            return om
    except Exception:
        pass
    return None


def _online_row_to_dict(r):
    return {"id": r.get("id"), "text": r.get("text", ""),
            "kind": r.get("kind") or "manuell",
            "by": r.get("erstellt_von") or "",
            "ts": _fmt_ts(r.get("ts")), "done": bool(r.get("done"))}


def list_todos():
    om = _online()
    if om is not None:
        rows, err = om.fetch()
        if err is None:
            return [_online_row_to_dict(r) for r in rows]
    return _local_list_todos()


def add_todo(text, by=""):
    text = (text or "").strip()
    if not text:
        return False
    om = _online()
    if om is not None:
        res, err = om.add(text, kind="manuell", by=by, key=None)
        if err is None and res and res.get("ok"):
            return True
    return _local_add_todo(text, by)


def toggle_todo(todo_id, by=""):
    om = _online()
    if om is not None:
        rows, err = om.fetch()
        if err is None:
            cur = next((r for r in rows if r.get("id") == todo_id), None)
            target = (not bool(cur.get("done"))) if cur else True
            _res, e2 = om.set_done(todo_id, done=target, von=by)
            if e2 is None:
                return
    _local_toggle_todo(todo_id, by)


def delete_todo(todo_id):
    om = _online()
    if om is not None:
        _res, err = om.delete(todo_id)
        if err is None:
            return
    _local_delete_todo(todo_id)


def push_cockpit_todo(text, *, kind="system", by="System", key=None):
    """Haengt eine System-Meldung an die geteilte Cockpit-Liste an (zentral wenn
    moeglich, sonst lokal). ``key`` dedupt offene Eintraege -> genau 1 Meldung pro
    Vorgang. Gibt True zurueck, wenn ein neuer Eintrag angelegt wurde."""
    text = (text or "").strip()
    if not text:
        return False
    om = _online()
    if om is not None:
        res, err = om.add(text, kind=kind, by=by, key=key)
        if err is None and res and res.get("ok"):
            return not res.get("dedupe", False)
    return _local_push_cockpit_todo(text, kind=kind, by=by, key=key)
