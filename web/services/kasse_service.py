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
    con.execute("""CREATE TABLE IF NOT EXISTS tbl_kunden (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT, inhaber TEXT, plz TEXT, ort TEXT)""")
    con.execute("""CREATE TABLE IF NOT EXISTS tbl_artikel (
        pzn TEXT PRIMARY KEY, bezeichnung TEXT)""")
    # Migration: woher kam der Auftrag + wer hat ihn erfasst (für die spätere
    # Zusammenführung mit der Desktop-Kasse: "User (online)").
    _add_col(con, "tbl_bestellungen", "erfasst_von", "erfasst_von TEXT")
    _add_col(con, "tbl_bestellungen", "quelle", "quelle TEXT DEFAULT 'online'")
    # In die PC-Kasse übernommen (= dort als Verkauf abgeschlossen) -> aus der
    # Online-To-do-Liste nehmen, aber in der Web-Historie behalten.
    _add_col(con, "tbl_bestellungen", "uebernommen", "uebernommen INTEGER DEFAULT 0")
    _add_col(con, "tbl_bestellungen", "uebernommen_am", "uebernommen_am TEXT")
    _add_col(con, "tbl_bestellungen", "uebernommen_von", "uebernommen_von TEXT")
    con.commit()


def _add_col(con: sqlite3.Connection, table: str, col: str, ddl: str) -> None:
    """Fügt eine Spalte hinzu, falls sie noch fehlt (einfache Migration)."""
    cols = {r[1] for r in con.execute(f"PRAGMA table_info({table})").fetchall()}
    if col not in cols:
        con.execute(f"ALTER TABLE {table} ADD COLUMN {ddl}")


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
        """SELECT id, kunde_name, datum, summe, status, erfasst_von, quelle
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


# ── Schnellverkauf (Handy: Kunde + Artikel -> bestellen/vorbestellen) ─────────
def artikel_vorschlag(con: sqlite3.Connection) -> list[dict]:
    """PZN + Bezeichnung fuer die Artikel-Auswahl (aus dem Lagerbestand).

    Vorbestellungen duerfen auch nicht gelagerte Artikel enthalten – die Liste
    dient nur als Tipphilfe (datalist), nicht als Zwang.
    """
    return [dict(r) for r in con.execute(
        """SELECT pzn, bezeichnung FROM tbl_lagerbestand
            WHERE pzn IS NOT NULL AND pzn <> ''
            GROUP BY pzn ORDER BY bezeichnung""").fetchall()]


def kunden_vorschlag(con: sqlite3.Connection) -> list[str]:
    """Bisher erfasste Kundennamen (Tipphilfe fuer das Kundenfeld)."""
    return [r[0] for r in con.execute(
        """SELECT kunde_name FROM tbl_bestellungen
            WHERE kunde_name IS NOT NULL AND kunde_name <> ''
            GROUP BY kunde_name ORDER BY kunde_name""").fetchall()]


def schnellverkauf_speichern(con: sqlite3.Connection, kunde_name: str, art: str,
                             liefertermin: str, positionen, benutzer: str) -> int:
    """Legt eine Bestellung bzw. Vorbestellung mit Positionen an.

    ``art`` = "Bestellung" oder "Vorbestellung" (steuert den Status). ``positionen``
    ist eine Liste von (pzn, bezeichnung, menge, einzelpreis) – leere/0-Mengen-Zeilen
    werden uebersprungen. Gibt die neue Bestell-ID zurueck (0, wenn keine gueltige
    Position dabei war).
    """
    ist_vor = (art or "").strip().lower().startswith("vor")
    status = "vorbestellt" if ist_vor else "offen"
    heute = date.today().isoformat()

    rows, summe = [], 0.0
    for pzn, bez, menge, preis in positionen:
        pzn = (pzn or "").strip()
        bez = (bez or "").strip()
        if not pzn and not bez:
            continue
        try:
            m = int(float(str(menge or "0").replace(",", ".")))
        except ValueError:
            m = 0
        try:
            p = float(str(preis or "0").replace(",", "."))
        except ValueError:
            p = 0.0
        if m <= 0:
            continue
        rows.append((pzn, bez, m, p))
        summe += m * p
    if not rows:
        return 0

    cur = con.execute(
        """INSERT INTO tbl_bestellungen(kunde_name, datum, liefertermin, summe, status)
           VALUES(?,?,?,?,?)""",
        (kunde_name.strip() or "—", heute, (liefertermin or "").strip() or None,
         summe, status))
    best_id = cur.lastrowid
    con.executemany(
        """INSERT INTO tbl_bestellpositionen
               (bestellung_id, pzn, bezeichnung, menge, einzelpreis, bestellart)
           VALUES(?,?,?,?,?,?)""",
        [(best_id, pzn, bez, m, p, ("Vorbestellung" if ist_vor else "Bestellung"))
         for (pzn, bez, m, p) in rows])
    con.execute(
        "INSERT INTO tbl_kasse_log(benutzer, aktion, detail) VALUES(?,?,?)",
        (benutzer, "schnellverkauf",
         f"{'Vorbestellung' if ist_vor else 'Bestellung'}: {kunde_name} "
         f"({len(rows)} Pos., {summe:.2f} EUR)"))
    con.commit()
    return best_id


# ── Mobiler Verkaufs-Flow v2 (Warenkorb + Bestands-Split + Apothekensuche) ────
def _reservierungen(con: sqlite3.Connection) -> dict:
    """Je PZN die Menge, die durch offene (noch nicht in die PC-Kasse übernommene)
    Online-Bestellungen reserviert ist. Vorbestellungen reservieren nichts (haben
    ohnehin keinen Bestand). So blockiert jeder offene Handyverkauf seine Ware."""
    rows = con.execute(
        """SELECT p.pzn, COALESCE(SUM(p.menge), 0) AS res
             FROM tbl_bestellpositionen p
             JOIN tbl_bestellungen b ON b.id = p.bestellung_id
            WHERE COALESCE(b.quelle, 'online') = 'online'
              AND COALESCE(b.uebernommen, 0) = 0
              AND p.bestellart = 'Bestellung'
              AND p.pzn IS NOT NULL AND p.pzn <> ''
            GROUP BY p.pzn""").fetchall()
    return {r["pzn"]: int(r["res"]) for r in rows}


def artikel_mit_bestand(con: sqlite3.Connection) -> list[dict]:
    """Alle bekannten Artikel (Artikelstamm + Lager) je PZN mit **verfügbarem**
    Bestand = Lagerbestand − offene Online-Reservierungen. Artikel ohne Bestand
    erscheinen mit 0 (auswählbar -> Vorbestellung). Verfügbar wird nie negativ.
    """
    roh = con.execute(
        """SELECT pzn, MAX(bezeichnung) AS bezeichnung,
                  COALESCE(SUM(bestand), 0) AS bestand
             FROM (
                 SELECT pzn, bezeichnung, 0 AS bestand FROM tbl_artikel
                 UNION ALL
                 SELECT pzn, bezeichnung, menge AS bestand FROM tbl_lagerbestand
             )
            WHERE pzn IS NOT NULL AND pzn <> ''
            GROUP BY pzn ORDER BY bezeichnung""").fetchall()
    res = _reservierungen(con)
    return [{"pzn": r["pzn"], "bezeichnung": r["bezeichnung"],
             "bestand": max(0, int(r["bestand"]) - res.get(r["pzn"], 0))}
            for r in roh]


def lager_ersetzen(con: sqlite3.Connection, artikel, bestand) -> tuple[int, int]:
    """Ersetzt Artikelstamm + Lagerbestand komplett mit den Daten der PC-Kasse
    (PC ist die Bestands-Quelle). Gibt (#artikel, #bestandszeilen) zurück."""
    con.execute("DELETE FROM tbl_artikel")
    con.executemany(
        "INSERT OR REPLACE INTO tbl_artikel(pzn, bezeichnung) VALUES(?,?)",
        [(a.get("pzn"), a.get("bezeichnung")) for a in artikel if a.get("pzn")])
    con.execute("DELETE FROM tbl_lagerbestand")
    con.executemany(
        """INSERT INTO tbl_lagerbestand(pzn, bezeichnung, charge, verfall, menge)
           VALUES(?,?,?,?,?)""",
        [(b.get("pzn"), b.get("bezeichnung"), b.get("charge"), b.get("verfall"),
          int(b.get("menge") or 0)) for b in bestand if b.get("pzn")])
    con.commit()
    n_a = con.execute("SELECT COUNT(*) FROM tbl_artikel").fetchone()[0]
    n_b = con.execute("SELECT COUNT(*) FROM tbl_lagerbestand").fetchone()[0]
    return n_a, n_b


def kunden_alle(con: sqlite3.Connection) -> list[dict]:
    """Apotheken für die Suche (Name/Inhaber/PLZ) – clientseitig gefiltert."""
    return [dict(r) for r in con.execute(
        "SELECT id, name, inhaber, plz, ort FROM tbl_kunden ORDER BY name").fetchall()]


def verkauf_speichern_v2(con: sqlite3.Connection, kunde_id, kunde_name: str,
                         positionen, benutzer: str) -> dict | None:
    """Ein Auftrag mit markierten Positionen + automatischer Bestands-Split.

    ``positionen`` = Liste von (pzn, bezeichnung, menge). Je Position wird gegen
    den Lagerbestand geprüft: der lieferbare Teil bekommt bestellart 'Bestellung',
    der Rest (und alles ohne Bestand) bestellart 'Vorbestellung'. Gibt eine
    Zusammenfassung zurück (oder None, wenn keine gültige Position dabei war).
    """
    bestand = {r["pzn"]: int(r["bestand"]) for r in artikel_mit_bestand(con)}
    pos_rows, vorbestellt_stueck = [], 0
    for pzn, bez, menge in positionen:
        pzn = (pzn or "").strip()
        bez = (bez or "").strip()
        try:
            m = int(float(str(menge or "0").replace(",", ".")))
        except ValueError:
            m = 0
        if (not pzn and not bez) or m <= 0:
            continue
        verf = bestand.get(pzn, 0) if pzn else 0
        lieferbar = max(0, min(m, verf))
        vor = m - lieferbar
        if lieferbar > 0:
            pos_rows.append((pzn, bez, lieferbar, "Bestellung"))
        if vor > 0:
            pos_rows.append((pzn, bez, vor, "Vorbestellung"))
            vorbestellt_stueck += vor
    if not pos_rows:
        return None

    status = "offen" if any(p[3] == "Bestellung" for p in pos_rows) else "vorbestellt"
    try:
        kid = int(kunde_id) if str(kunde_id or "").strip() else None
    except ValueError:
        kid = None
    cur = con.execute(
        """INSERT INTO tbl_bestellungen(kunde_id, kunde_name, datum, status,
                                        erfasst_von, quelle)
           VALUES(?,?,?,?,?,?)""",
        (kid, kunde_name.strip() or "—", date.today().isoformat(), status,
         benutzer, "online"))
    best_id = cur.lastrowid
    con.executemany(
        """INSERT INTO tbl_bestellpositionen
               (bestellung_id, pzn, bezeichnung, menge, bestellart)
           VALUES(?,?,?,?,?)""",
        [(best_id, *p) for p in pos_rows])
    vorbestellt_pos = sum(1 for p in pos_rows if p[3] == "Vorbestellung")
    con.execute(
        "INSERT INTO tbl_kasse_log(benutzer, aktion, detail) VALUES(?,?,?)",
        (benutzer, "verkauf",
         f"{kunde_name}: {len(pos_rows)} Pos. ({vorbestellt_stueck} Stk Vorbestellung)"))
    con.commit()
    return {
        "best_id": best_id,
        "positionen": len(pos_rows),
        "vorbestellt_pos": vorbestellt_pos,
        "vorbestellt_stueck": vorbestellt_stueck,
        "status": status,
    }
