"""Import von Kundendaten und Wareneingang aus TXT / CSV / Excel.

Nutzt file_loader.load_table (Format-/Delimiter-/Encoding-Erkennung) und
file_loader.find_column (Spalten-Zuordnung per Synonymen). Beide Importe sind
hier gekapselt, damit sie aus der Kasse-App (auch standalone) aufrufbar sind.

PDF wird (noch) nicht unterstuetzt - keine PDF-Bibliothek installiert; eine
zuverlaessige Tabellen-Extraktion aus PDF braucht z.B. pdfplumber und bleibt
fehleranfaellig. Siehe Hinweis in der GUI.
"""
from __future__ import annotations

import getpass
import sqlite3
from collections import OrderedDict
from datetime import datetime, date, timedelta

from .file_loader import load_table, find_column

SUPPORTED_PATTERNS = ("*.xlsx", "*.xlsm", "*.csv", "*.txt")


def _cell(row, idx):
    if idx is None or idx >= len(row):
        return ""
    v = row[idx]
    return "" if v is None else str(v).strip()


def _to_int(text):
    t = str(text or "").strip().replace(",", ".")
    if not t:
        return None
    try:
        return int(float(t))
    except ValueError:
        return None


def _to_float(text):
    t = str(text or "").strip().replace(",", ".")
    if not t:
        return None
    try:
        return float(t)
    except ValueError:
        return None


def _ensure_kunden_center(con):
    """Spiegelt das tbl_kunden_center-Schema von NMGone (gui.py), damit beide
    Programme dieselbe Tabelle verwenden."""
    con.execute(
        """CREATE TABLE IF NOT EXISTS tbl_kunden_center(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            kundennummer TEXT, kundenname TEXT, kundentyp TEXT,
            ansprechpartner TEXT, telefon TEXT, email TEXT,
            status TEXT DEFAULT 'aktiv', notizen TEXT,
            erstellt_am TEXT DEFAULT CURRENT_TIMESTAMP, geaendert_am TEXT, bearbeiter TEXT
        )"""
    )
    have = {r[1] for r in con.execute("PRAGMA table_info(tbl_kunden_center)")}
    for col in ("plz", "ort", "strasse", "inhaber", "ansprechpartner2"):
        if col not in have:
            con.execute(f"ALTER TABLE tbl_kunden_center ADD COLUMN {col} TEXT")


def import_kunden(db_path, path) -> dict:
    """Importiert Kundendaten in tbl_kunden_center. Upsert ueber kundennummer."""
    table = load_table(path)
    headers = list(table.headers)
    cols = {
        "kundennummer": find_column(headers, ["Kundennummer", "Kunden-Nr", "KdNr", "Nr", "Kundennr"],
                                    ["kundennr", "kdnr", "kundennummer"]),
        "kundenname": find_column(headers, ["Apothekenname", "Apotheke", "Kundenname", "Name", "Kunde"],
                                  ["apotheke", "kundenname", "name"]),
        "plz": find_column(headers, ["PLZ", "Postleitzahl"], ["plz", "postleit"]),
        "ort": find_column(headers, ["Ort", "Stadt"], ["ort", "stadt"]),
        "strasse": find_column(headers, ["Straße", "Strasse", "Adresse"], ["strasse", "adresse"]),
        "inhaber": find_column(headers, ["Inhaber", "Inhaberin", "Besitzer"], ["inhaber", "besitzer"]),
        "telefon": find_column(headers, ["Telefon", "Tel", "Telefonnummer"], ["telefon", "tel"]),
        "email": find_column(headers, ["E-Mail", "Email", "Mail"], ["mail"]),
    }
    if cols["kundennummer"] is None and cols["kundenname"] is None:
        raise ValueError("Keine erkennbare Kundennummer- oder Namens-Spalte gefunden.")

    erkannt = [k for k, v in cols.items() if v is not None]
    neu = aktualisiert = uebersprungen = 0
    bearbeiter = getpass.getuser()
    jetzt = datetime.now().isoformat(timespec="seconds")

    with sqlite3.connect(db_path) as con:
        _ensure_kunden_center(con)
        for row in table.rows:
            werte = {k: _cell(row, idx) for k, idx in cols.items()}
            if not werte["kundennummer"] and not werte["kundenname"]:
                uebersprungen += 1
                continue
            knr = werte["kundennummer"]
            existing = None
            if knr:
                existing = con.execute(
                    "SELECT id FROM tbl_kunden_center WHERE kundennummer=? LIMIT 1", (knr,)
                ).fetchone()
            if existing:
                con.execute(
                    "UPDATE tbl_kunden_center SET kundenname=?, plz=?, ort=?, strasse=?, "
                    "inhaber=?, telefon=?, email=?, geaendert_am=?, bearbeiter=? WHERE id=?",
                    (werte["kundenname"], werte["plz"], werte["ort"], werte["strasse"],
                     werte["inhaber"], werte["telefon"], werte["email"], jetzt, bearbeiter, existing[0]),
                )
                aktualisiert += 1
            else:
                con.execute(
                    "INSERT INTO tbl_kunden_center(kundennummer, kundenname, plz, ort, strasse, "
                    "inhaber, telefon, email, status, bearbeiter) VALUES(?,?,?,?,?,?,?,?,?,?)",
                    (knr, werte["kundenname"], werte["plz"], werte["ort"], werte["strasse"],
                     werte["inhaber"], werte["telefon"], werte["email"], "aktiv", bearbeiter),
                )
                neu += 1
        con.commit()

    return {"gelesen": len(table.rows), "neu": neu, "aktualisiert": aktualisiert,
            "uebersprungen": uebersprungen, "spalten": erkannt, "quelle": table.source_type}


def import_wareneingang(db_path, path) -> dict:
    """Importiert Wareneingaenge -> tbl_lagerbestand (+ Historie). Nur NMG-PZNs."""
    table = load_table(path)
    headers = list(table.headers)
    c_pzn = find_column(headers, ["PZN", "PZN-Code", "Artikel-PZN"], ["pzn"])
    c_charge = find_column(headers, ["Charge", "Chargennummer", "Lot", "Los"], ["charge", "lot"])
    c_verfall = find_column(headers, ["Verfall", "Verfalldatum", "Haltbarkeit", "MHD", "Verwendbar bis"],
                            ["verfall", "mhd", "haltbar"])
    c_menge = find_column(headers, ["Menge", "Anzahl", "Stück", "Stueck", "Bestand"],
                          ["menge", "anzahl", "stueck", "stck"])
    c_ek = find_column(headers, ["EK", "Einkaufspreis", "EK-Preis", "Einkauf"], ["ek", "einkauf"])
    if c_pzn is None:
        raise ValueError("Keine PZN-Spalte gefunden.")
    if c_menge is None:
        raise ValueError("Keine Mengen-Spalte gefunden.")

    neu = erhoeht = kein_nmg = uebersprungen = 0
    jetzt = datetime.now().isoformat(timespec="seconds")
    from pathlib import Path as _P
    beleg = _P(path).name

    with sqlite3.connect(db_path) as con:
        cur = con.execute(
            "INSERT INTO tbl_wareneingang(datum, lieferant, lieferschein, bearbeiter) VALUES(?,?,?,?)",
            (jetzt, "Import", beleg, getpass.getuser()))
        we_id = cur.lastrowid
        for row in table.rows:
            pzn = _cell(row, c_pzn)
            menge = _to_int(_cell(row, c_menge))
            if not pzn or not menge or menge <= 0:
                uebersprungen += 1
                continue
            art = con.execute("SELECT artikelname FROM tbl_nmg_stamm WHERE pzn=? LIMIT 1", (pzn,)).fetchone()
            if not art:
                kein_nmg += 1
                continue
            artikelname = art[0]
            charge = _cell(row, c_charge)
            verfall = _cell(row, c_verfall)
            from .kasse_app import _normalize_verfall
            norm = _normalize_verfall(verfall)
            if norm is not None:  # gueltig -> aufgefuellt; ungueltig -> Rohwert behalten
                verfall = norm
            ek = _to_float(_cell(row, c_ek))
            con.execute(
                "INSERT INTO tbl_wareneingang_positionen(we_id,pzn,artikelname,charge,verfall,menge,ek) "
                "VALUES(?,?,?,?,?,?,?)", (we_id, pzn, artikelname, charge, verfall, menge, ek))
            upd = con.execute(
                "UPDATE tbl_lagerbestand SET menge=menge+?, aktualisiert_am=? "
                "WHERE pzn=? AND COALESCE(charge,'')=? AND COALESCE(verfall,'')=?",
                (menge, jetzt, pzn, charge, verfall))
            if upd.rowcount == 0:
                con.execute(
                    "INSERT INTO tbl_lagerbestand(pzn,artikelname,charge,verfall,menge,aktualisiert_am) "
                    "VALUES(?,?,?,?,?,?)", (pzn, artikelname, charge, verfall, menge, jetzt))
                neu += 1
            else:
                erhoeht += 1
        con.commit()

    return {"gelesen": len(table.rows), "neu_chargen": neu, "erhoehte_chargen": erhoeht,
            "kein_nmg": kein_nmg, "uebersprungen": uebersprungen, "quelle": table.source_type}


def _parse_datum(t):
    """Datum -> ISO 'YYYY-MM-DD'. Akzeptiert deutsche Formate (TT.MM.JJJJ,
    TT.MM.JJ, TT/MM/JJJJ), ISO und Excel-Seriennummern. Leer/ungueltig -> None."""
    t = str(t or "").strip()
    if not t:
        return None
    for fmt in ("%Y-%m-%d", "%d.%m.%Y", "%d.%m.%y", "%d/%m/%Y", "%d/%m/%y", "%Y/%m/%d"):
        try:
            return datetime.strptime(t, fmt).strftime("%Y-%m-%d")
        except ValueError:
            pass
    # Excel-Seriennummer (Tage seit 1899-12-30)
    try:
        serial = int(float(t.replace(",", ".")))
        if 1 <= serial <= 60000:
            return (date(1899, 12, 30) + timedelta(days=serial)).isoformat()
    except ValueError:
        pass
    return None


def import_verkaeufe(db_path, path, als_vorbestellung=False) -> dict:
    """Importiert historische Bestellungen/Verkaeufe (oder Vorbestellungen) aus
    TXT/CSV/Excel. Spalten: Kundennummer, PZN, Menge, Preis, Rabatt,
    Bestell-/Lieferdatum. Je Zeile eine Position; Zeilen mit gleichem Kunden +
    Datum werden zu EINER Bestellung zusammengefasst. Bucht KEINEN Lagerbestand ab
    (Historie)."""
    table = load_table(path)
    headers = list(table.headers)
    c_knr = find_column(headers, ["Kundennummer", "Kunden-Nr", "KdNr", "Nr", "Kundennr"],
                        ["kundennr", "kdnr", "kundennummer"])
    c_pzn = find_column(headers, ["PZN", "PZN-Code", "Artikel-PZN"], ["pzn"])
    c_menge = find_column(headers, ["Menge", "Anzahl", "Stück", "Stueck"], ["menge", "anzahl", "stueck"])
    c_preis = find_column(headers, ["Preis", "APU", "Einzelpreis", "EK", "VK"], ["preis", "apu", "ek", "vk"])
    c_rabatt = find_column(headers, ["Rabatt", "Rabatt %", "Rabatt%"], ["rabatt"])
    c_datum = find_column(headers, ["Bestelldatum", "Lieferdatum", "Datum", "Bestell-/Lieferdatum"],
                          ["bestelldatum", "lieferdatum", "datum"])
    if c_knr is None or c_pzn is None or c_menge is None:
        raise ValueError("Spalten Kundennummer, PZN und Menge werden benoetigt.")

    bestellart = "Vorbestellung" if als_vorbestellung else "Bestellung"
    status = "offen" if als_vorbestellung else "erledigt"
    erkannt = [n for n, c in (("Preis", c_preis), ("Rabatt", c_rabatt), ("Datum", c_datum))
               if c is not None]
    heute = date.today().isoformat()

    gruppen = OrderedDict()
    uebersprungen = datum_unklar = 0
    for row in table.rows:
        knr = _cell(row, c_knr)
        pzn = _cell(row, c_pzn)
        menge = _to_int(_cell(row, c_menge))
        if not knr or not pzn or not menge or menge <= 0:
            uebersprungen += 1
            continue
        datum = _parse_datum(_cell(row, c_datum)) if c_datum is not None else None
        if c_datum is not None and _cell(row, c_datum) and datum is None:
            datum_unklar += 1
        if datum is None:
            datum = heute
        preis = _to_float(_cell(row, c_preis)) if c_preis is not None else None
        rabatt = _to_float(_cell(row, c_rabatt)) if c_rabatt is not None else None
        gruppen.setdefault((knr, datum), []).append((pzn, menge, preis, rabatt))

    bearbeiter = getpass.getuser()
    bestellungen = positionen = 0
    with sqlite3.connect(db_path) as con:
        hat_kc = con.execute(
            "SELECT 1 FROM sqlite_master WHERE type='table' AND name='tbl_kunden_center'").fetchone()
        for (knr, datum), items in gruppen.items():
            apotheke = ""
            if hat_kc:
                r = con.execute("SELECT kundenname FROM tbl_kunden_center WHERE kundennummer=? LIMIT 1",
                                (knr,)).fetchone()
                apotheke = r[0] if r else ""
            cur = con.execute(
                "INSERT INTO tbl_bestellungen(datum,kundennummer,apotheke,status,bestellart,bearbeiter) "
                "VALUES(?,?,?,?,?,?)", (datum, knr, apotheke, status, bestellart, bearbeiter))
            bid = cur.lastrowid
            bestellungen += 1
            for pzn, menge, preis, rabatt in items:
                nmg = con.execute("SELECT artikelname, apu FROM tbl_nmg_stamm WHERE pzn=? LIMIT 1",
                                  (pzn,)).fetchone()
                name = nmg[0] if nmg else pzn
                apu = preis if preis is not None else (nmg[1] if nmg else None)
                con.execute(
                    "INSERT INTO tbl_bestellpositionen(bestell_id,pzn,artikelname,apu,menge,"
                    "rabatt_prozent,rabatt_quelle,bestellart,liefertermin) VALUES(?,?,?,?,?,?,?,?,?)",
                    (bid, pzn, name, apu, menge, rabatt or 0, "Import", bestellart, datum))
                positionen += 1
        con.commit()

    return {"gelesen": len(table.rows), "bestellungen": bestellungen, "positionen": positionen,
            "uebersprungen": uebersprungen, "datum_unklar": datum_unklar, "spalten": erkannt,
            "quelle": table.source_type, "als_vorbestellung": als_vorbestellung}
