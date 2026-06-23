"""Biosimilar-Wissensbasis aus der Gelben Liste.

Eigene Datenschicht, die die *klinische Wahrheit* abbildet:

    Wirkstoff (z.B. Adalimumab)
      -> Referenzprodukt / Original (z.B. Humira)
      -> n Biosimilars (Amgevita, Hyrimoz, Idacio, ...)

Diese Schicht liegt bewusst NEBEN tbl_austauschdatenbank (PZN->PZN-Austausch)
und wird nicht damit vermischt. Sie verzweigt nach unten zu PZNs, indem die
Markennamen aus der Liste gegen den Artikelstamm aufgeloest werden. Zur
Absicherung dient die Doppel-Bruecke: ein PZN-Kandidat zaehlt nur dann als
sicherer Auto-Treffer, wenn sein Wirkstoff (tbl_wirkstoff_staerke) zum
Wirkstoff der Gruppe passt. Alles Unsichere landet in einer Pruef-Warteschlange.

Tabellen:
    tbl_biosimilar_gruppe   (1 Zeile je Wirkstoff)
    tbl_biosimilar_produkt  (1 Zeile je Arzneimittel, rolle=original|biosimilar)
    tbl_biosimilar_pzn      (Verzweigung Name->PZN, match_status=auto|manuell|offen)
    tbl_biosimilar_changelog(Verschiebungs-Log je Import-Stand)

Quelle: https://www.gelbe-liste.de/biosimilars/zugelassene-biosimilars
Excel-Layout (3 Spalten, Forward-Fill): Wirkstoff | Arzneimittel | Referenzprodukt
"""

from __future__ import annotations

import re
import sqlite3
from datetime import datetime
from pathlib import Path

from .config import DB_PATH
from .file_loader import load_worksheet


# ── Normalisierung ───────────────────────────────────────────────────────────
def _clean(value) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    text = str(value).replace("\xa0", " ").strip()
    return " ".join(text.split())


def _norm_name(value) -> str:
    """Vergleichsschluessel fuer Wirkstoff-/Markennamen: klein, ohne Sonder-
    zeichen. 'Insulin glargin' -> 'insulinglargin', 'Epoetin alfa' -> ..."""
    text = _clean(value).lower()
    text = (text.replace("ä", "ae").replace("ö", "oe")
                .replace("ü", "ue").replace("ß", "ss"))
    return re.sub(r"[^a-z0-9]+", "", text)


def _pzn(value) -> str:
    """8-stellig, nur Ziffern. /N-Lagervarianten-Suffix wird abgeschnitten
    (vgl. wirkstoff_db._pzn / SP29)."""
    if value is None:
        return ""
    text = str(value).strip()
    if not text:
        return ""
    text = re.sub(r"\s*/\s*\w+\s*$", "", text)
    if text.endswith(".0") and text[:-2].isdigit():
        text = text[:-2]
    digits = "".join(ch for ch in text if ch.isdigit())
    return digits.zfill(8) if digits else ""


# ── Schema ───────────────────────────────────────────────────────────────────
def ensure_biosimilar_tables() -> None:
    """Legt alle Biosimilar-Tabellen samt Indizes an. Idempotent."""
    with sqlite3.connect(DB_PATH) as con:
        con.execute("""
            CREATE TABLE IF NOT EXISTS tbl_biosimilar_gruppe (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                wirkstoff TEXT NOT NULL,
                wirkstoff_norm TEXT NOT NULL UNIQUE,
                referenzprodukt TEXT,
                quelle TEXT NOT NULL DEFAULT 'Gelbe Liste',
                stand TEXT,
                erstellt_am TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
                aktualisiert_am TEXT
            )
        """)
        con.execute("""
            CREATE TABLE IF NOT EXISTS tbl_biosimilar_produkt (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                gruppe_id INTEGER NOT NULL,
                name TEXT NOT NULL,
                name_norm TEXT NOT NULL,
                rolle TEXT NOT NULL DEFAULT 'biosimilar',
                status TEXT NOT NULL DEFAULT 'aktiv',
                quelle TEXT NOT NULL DEFAULT 'Gelbe Liste',
                stand TEXT,
                erstellt_am TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
                aktualisiert_am TEXT,
                UNIQUE(gruppe_id, name_norm)
            )
        """)
        con.execute("""
            CREATE TABLE IF NOT EXISTS tbl_biosimilar_pzn (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                produkt_id INTEGER NOT NULL,
                pzn TEXT NOT NULL,
                artikelname TEXT,
                match_status TEXT NOT NULL DEFAULT 'offen',
                erstellt_am TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
                UNIQUE(produkt_id, pzn)
            )
        """)
        con.execute("""
            CREATE TABLE IF NOT EXISTS tbl_biosimilar_changelog (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                stand TEXT,
                typ TEXT NOT NULL,
                wirkstoff TEXT,
                produkt TEXT,
                detail TEXT,
                erstellt_am TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
            )
        """)
        con.execute("CREATE INDEX IF NOT EXISTS idx_biosim_produkt_gruppe ON tbl_biosimilar_produkt (gruppe_id)")
        con.execute("CREATE INDEX IF NOT EXISTS idx_biosim_produkt_norm ON tbl_biosimilar_produkt (name_norm)")
        con.execute("CREATE INDEX IF NOT EXISTS idx_biosim_pzn_produkt ON tbl_biosimilar_pzn (produkt_id)")
        con.execute("CREATE INDEX IF NOT EXISTS idx_biosim_pzn_pzn ON tbl_biosimilar_pzn (pzn)")
        con.commit()


# ── Parser (Gelbe Liste Excel, Forward-Fill) ─────────────────────────────────
def parse_gelbe_liste(path: str | Path) -> list[dict]:
    """Liest die Gelbe-Liste-Excel und liefert je Wirkstoff ein dict:
        {"wirkstoff": str, "referenzprodukt": str, "biosimilars": [str, ...]}

    Layout: 3 Spalten (Wirkstoff | Arzneimittel | Referenzprodukt). Wirkstoff
    und Referenzprodukt stehen nur in der ersten Zeile der Gruppe (Forward-Fill).
    Das Referenzprodukt ist das Original und wird mit rolle='original' gefuehrt.
    """
    ws, _ = load_worksheet(Path(path))
    # Spalten anhand der Kopfzeile finden (tolerant gegen \xa0 / Reihenfolge).
    header_row = 1
    col_w = col_a = col_r = None
    for col in range(1, ws.max_column + 1):
        key = _norm_name(ws.cell(header_row, col).value)
        if key.startswith("wirkstoff"):
            col_w = col
        elif key.startswith("arzneimittel") or key.startswith("praeparat") or key.startswith("biosimilar"):
            col_a = col
        elif key.startswith("referenz"):
            col_r = col
    if col_w is None or col_a is None:
        # Fallback auf feste Reihenfolge der bekannten Vorlage.
        col_w, col_a, col_r = 1, 2, 3

    groups: dict[str, dict] = {}
    order: list[str] = []
    cur_key = None
    for row in range(header_row + 1, ws.max_row + 1):
        w = _clean(ws.cell(row, col_w).value)
        a = _clean(ws.cell(row, col_a).value)
        r = _clean(ws.cell(row, col_r).value) if col_r else ""
        if w:
            cur_key = _norm_name(w)
            if cur_key and cur_key not in groups:
                groups[cur_key] = {"wirkstoff": w, "referenzprodukt": r, "biosimilars": []}
                order.append(cur_key)
            elif cur_key and r and not groups[cur_key]["referenzprodukt"]:
                groups[cur_key]["referenzprodukt"] = r
        if a and cur_key and cur_key in groups:
            if a not in groups[cur_key]["biosimilars"]:
                groups[cur_key]["biosimilars"].append(a)
    return [groups[k] for k in order]


# ── Import + Diff (Verschiebungs-Erkennung) ──────────────────────────────────
def import_gelbe_liste_excel(path: str | Path, quelle: str = "Gelbe Liste",
                             stand: str | None = None) -> dict:
    """Importiert die Gelbe-Liste-Excel, erkennt Verschiebungen gegen den
    bisherigen Stand und schreibt sie ins Changelog.

    Ablauf:
      1. Schnappschuss der aktuell aktiven Produkte je Gruppe (vorher).
      2. Gruppen + Produkte upserten (Original aus Referenzprodukt).
      3. Diff: neue Produkte -> Changelog 'produkt_neu'; vorher aktive, jetzt
         fehlende Produkte -> status='entfernt' + Changelog 'produkt_entfernt'.
      4. Namen -> PZN aufloesen (Auto/Offen), manuelle Treffer bleiben erhalten.

    Liefert Statistik inkl. Changelog-Eintraegen.
    """
    ensure_biosimilar_tables()
    stand = stand or datetime.now().strftime("%Y-%m-%d")
    parsed = parse_gelbe_liste(path)

    stats = {
        "stand": stand,
        "gruppen": len(parsed),
        "produkte": 0,
        "gruppen_neu": 0,
        "produkte_neu": 0,
        "produkte_entfernt": 0,
        "changelog": [],
    }

    with sqlite3.connect(DB_PATH) as con:
        con.row_factory = sqlite3.Row

        def log(typ, wirkstoff, produkt, detail):
            con.execute(
                "INSERT INTO tbl_biosimilar_changelog (stand, typ, wirkstoff, produkt, detail) "
                "VALUES (?,?,?,?,?)", (stand, typ, wirkstoff, produkt, detail))
            stats["changelog"].append(
                {"typ": typ, "wirkstoff": wirkstoff, "produkt": produkt, "detail": detail})

        for g in parsed:
            wirkstoff = g["wirkstoff"]
            wnorm = _norm_name(wirkstoff)
            referenz = g["referenzprodukt"]
            if not wnorm:
                continue

            grp = con.execute(
                "SELECT id, referenzprodukt FROM tbl_biosimilar_gruppe WHERE wirkstoff_norm=?",
                (wnorm,)).fetchone()
            if grp is None:
                cur = con.execute(
                    "INSERT INTO tbl_biosimilar_gruppe (wirkstoff, wirkstoff_norm, referenzprodukt, quelle, stand) "
                    "VALUES (?,?,?,?,?)", (wirkstoff, wnorm, referenz, quelle, stand))
                gruppe_id = cur.lastrowid
                stats["gruppen_neu"] += 1
                log("gruppe_neu", wirkstoff, referenz, f"Neuer Wirkstoff mit Original '{referenz}'")
            else:
                gruppe_id = grp["id"]
                con.execute(
                    "UPDATE tbl_biosimilar_gruppe SET referenzprodukt=COALESCE(NULLIF(?,''),referenzprodukt), "
                    "stand=?, aktualisiert_am=CURRENT_TIMESTAMP WHERE id=?",
                    (referenz, stand, gruppe_id))
                if referenz and grp["referenzprodukt"] and _norm_name(referenz) != _norm_name(grp["referenzprodukt"]):
                    log("referenz_geaendert", wirkstoff, referenz,
                        f"Original '{grp['referenzprodukt']}' -> '{referenz}'")

            # Schnappschuss vorher (aktive Produkte dieser Gruppe).
            vorher = {
                r["name_norm"]: r["name"]
                for r in con.execute(
                    "SELECT name_norm, name FROM tbl_biosimilar_produkt "
                    "WHERE gruppe_id=? AND status='aktiv'", (gruppe_id,)).fetchall()
            }

            # Soll-Produkte: Original (rolle=original) + Biosimilars.
            soll: list[tuple[str, str]] = []
            if referenz:
                soll.append((referenz, "original"))
            for b in g["biosimilars"]:
                soll.append((b, "biosimilar"))

            soll_norms = set()
            for name, rolle in soll:
                nnorm = _norm_name(name)
                if not nnorm:
                    continue
                soll_norms.add(nnorm)
                stats["produkte"] += 1
                exists = con.execute(
                    "SELECT id, status FROM tbl_biosimilar_produkt WHERE gruppe_id=? AND name_norm=?",
                    (gruppe_id, nnorm)).fetchone()
                if exists is None:
                    con.execute(
                        "INSERT INTO tbl_biosimilar_produkt (gruppe_id, name, name_norm, rolle, status, quelle, stand) "
                        "VALUES (?,?,?,?, 'aktiv', ?, ?)", (gruppe_id, name, nnorm, rolle, quelle, stand))
                    stats["produkte_neu"] += 1
                    log("produkt_neu", wirkstoff, name, f"Neues {rolle} fuer {wirkstoff}")
                else:
                    con.execute(
                        "UPDATE tbl_biosimilar_produkt SET name=?, rolle=?, status='aktiv', "
                        "stand=?, aktualisiert_am=CURRENT_TIMESTAMP WHERE id=?",
                        (name, rolle, stand, exists["id"]))

            # Verschiebung: vorher aktiv, jetzt nicht mehr in der Liste -> entfernt.
            for nnorm, name in vorher.items():
                if nnorm not in soll_norms:
                    con.execute(
                        "UPDATE tbl_biosimilar_produkt SET status='entfernt', "
                        "stand=?, aktualisiert_am=CURRENT_TIMESTAMP WHERE gruppe_id=? AND name_norm=?",
                        (stand, gruppe_id, nnorm))
                    stats["produkte_entfernt"] += 1
                    log("produkt_entfernt", wirkstoff, name, f"Nicht mehr in Liste (Stand {stand})")

        con.commit()

    # PZN-Aufloesung nachziehen.
    resolve = resolve_pzn_matches()
    stats["pzn_auto"] = resolve["auto"]
    stats["pzn_offen"] = resolve["offen"]
    stats["produkte_ohne_pzn"] = resolve["ohne_kandidat"]
    return stats


def _wirkstoff_passt(kandidat_wirkstoff: str, gruppe_wirkstoff_norm: str) -> bool:
    """Wirkstoff-Bruecke fuer den sicheren Auto-Treffer.

    Sicher (auto), wenn der Wirkstoff des PZN-Kandidaten zum Wirkstoff der
    Gruppe passt:
      - exakte Gleichheit, ODER
      - der kuerzere Name ist Praefix des laengeren UND mindestens 8 Zeichen
        lang. Das faengt Salz-/Form-Suffixe ('Enoxaparin' -> 'Enoxaparin-
        Natrium') ein, verhindert aber das Kreuzen generischer Staemme
        ('Insulin' -> glargin/aspart, 'Epoetin' -> alfa/zeta), deren Stamm
        unter 8 Zeichen liegt.
    Synonyme mit voellig anderem Namen (Erythropoietin <-> Epoetin) bleiben
    bewusst 'offen' und landen in der Pruef-Warteschlange.
    """
    cw = _norm_name(kandidat_wirkstoff)
    gw = gruppe_wirkstoff_norm
    if not cw or not gw:
        return False
    if cw == gw:
        return True
    kurz, lang = (cw, gw) if len(cw) <= len(gw) else (gw, cw)
    return len(kurz) >= 8 and lang.startswith(kurz)


# ── Namen -> PZN Aufloesung (Doppel-Bruecke) ─────────────────────────────────
def resolve_pzn_matches() -> dict:
    """Loest aktive Produktnamen gegen den Artikelstamm auf.

    Pro Produkt:
      - Kandidaten = tbl_artikelstamm.artikel LIKE 'NAME%'
      - Wirkstoff des Kandidaten (tbl_wirkstoff_staerke) == Wirkstoff der Gruppe
        -> match_status='auto' (sichere Verzweigung)
      - Markentreffer, aber Wirkstoff weicht ab/fehlt -> 'offen' (Pruef-Queue)
      - kein Markentreffer -> Produkt bleibt ohne PZN (Pruef-Queue)

    Manuell bestaetigte Treffer (match_status='manuell') bleiben unangetastet.
    Re-runnable: 'auto'/'offen'-Zeilen werden je Produkt neu berechnet.
    """
    ensure_biosimilar_tables()
    result = {"auto": 0, "offen": 0, "ohne_kandidat": 0, "produkte": 0}

    with sqlite3.connect(DB_PATH) as con:
        con.row_factory = sqlite3.Row
        produkte = con.execute("""
            SELECT p.id, p.name, p.name_norm, g.wirkstoff_norm
            FROM tbl_biosimilar_produkt p
            JOIN tbl_biosimilar_gruppe g ON g.id = p.gruppe_id
            WHERE p.status='aktiv'
        """).fetchall()

        for p in produkte:
            result["produkte"] += 1
            # Manuelle Entscheidungen schuetzen.
            manuell = {
                r["pzn"] for r in con.execute(
                    "SELECT pzn FROM tbl_biosimilar_pzn WHERE produkt_id=? AND match_status='manuell'",
                    (p["id"],)).fetchall()
            }
            con.execute(
                "DELETE FROM tbl_biosimilar_pzn WHERE produkt_id=? AND match_status IN ('auto','offen')",
                (p["id"],))

            like = p["name"].upper().strip() + "%"
            kandidaten = con.execute("""
                SELECT a.pzn, a.artikel, w.wirkstoff
                FROM tbl_artikelstamm a
                LEFT JOIN tbl_wirkstoff_staerke w ON w.pzn = a.pzn
                WHERE a.artikel LIKE ?
            """, (like,)).fetchall()

            if not kandidaten:
                result["ohne_kandidat"] += 1
                continue

            for k in kandidaten:
                if k["pzn"] in manuell:
                    continue
                passt = _wirkstoff_passt(k["wirkstoff"], p["wirkstoff_norm"])
                status = "auto" if passt else "offen"
                con.execute(
                    "INSERT OR IGNORE INTO tbl_biosimilar_pzn (produkt_id, pzn, artikelname, match_status) "
                    "VALUES (?,?,?,?)", (p["id"], k["pzn"], k["artikel"], status))
                result[status] += 1

        con.commit()
    return result


# ── Abfragen ─────────────────────────────────────────────────────────────────
def biosimilar_overview() -> list[dict]:
    """Gruppen mit Zaehlern fuer eine Uebersicht."""
    ensure_biosimilar_tables()
    with sqlite3.connect(DB_PATH) as con:
        con.row_factory = sqlite3.Row
        rows = con.execute("""
            SELECT g.wirkstoff, g.referenzprodukt, g.stand,
                   SUM(CASE WHEN p.rolle='biosimilar' AND p.status='aktiv' THEN 1 ELSE 0 END) AS anzahl_biosimilar,
                   SUM(CASE WHEN p.status='aktiv' THEN 1 ELSE 0 END) AS anzahl_produkte
            FROM tbl_biosimilar_gruppe g
            LEFT JOIN tbl_biosimilar_produkt p ON p.gruppe_id = g.id
            GROUP BY g.id
            ORDER BY g.wirkstoff
        """).fetchall()
    return [dict(r) for r in rows]


def get_pruef_queue() -> list[dict]:
    """Produkte, die NICHT sicher aufgeloest sind: kein auto/manuell-PZN.
    Zeigt die unsicheren 'offen'-Kandidaten zur manuellen Entscheidung.
    """
    ensure_biosimilar_tables()
    with sqlite3.connect(DB_PATH) as con:
        con.row_factory = sqlite3.Row
        rows = con.execute("""
            SELECT p.id AS produkt_id, p.name, p.rolle, g.wirkstoff, g.referenzprodukt,
                   (SELECT COUNT(*) FROM tbl_biosimilar_pzn x
                      WHERE x.produkt_id=p.id AND x.match_status='offen') AS offene_kandidaten
            FROM tbl_biosimilar_produkt p
            JOIN tbl_biosimilar_gruppe g ON g.id = p.gruppe_id
            WHERE p.status='aktiv'
              AND NOT EXISTS (
                  SELECT 1 FROM tbl_biosimilar_pzn y
                  WHERE y.produkt_id=p.id AND y.match_status IN ('auto','manuell')
              )
            ORDER BY g.wirkstoff, p.rolle DESC, p.name
        """).fetchall()
    return [dict(r) for r in rows]


def biosimilar_counts() -> dict:
    """Kennzahlen fuer Status-/Kachelanzeige."""
    ensure_biosimilar_tables()
    with sqlite3.connect(DB_PATH) as con:
        def one(sql):
            return int(con.execute(sql).fetchone()[0])
        return {
            "gruppen": one("SELECT COUNT(*) FROM tbl_biosimilar_gruppe"),
            "produkte_aktiv": one("SELECT COUNT(*) FROM tbl_biosimilar_produkt WHERE status='aktiv'"),
            "pzn_auto": one("SELECT COUNT(*) FROM tbl_biosimilar_pzn WHERE match_status='auto'"),
            "pzn_manuell": one("SELECT COUNT(*) FROM tbl_biosimilar_pzn WHERE match_status='manuell'"),
            "pzn_offen": one("SELECT COUNT(*) FROM tbl_biosimilar_pzn WHERE match_status='offen'"),
        }
