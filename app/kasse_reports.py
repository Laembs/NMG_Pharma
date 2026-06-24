"""Auswertungen fuer die Kasse-App: Verfall-Uebersicht, Inventur/Lagerbestand
und Umsatz/Tagesabschluss.

Reine Lese-/Auswerte-Logik + PDF-Erzeugung (fpdf2, wie app/kurzbericht.py). Die
GUI (Reiter „Auswertung“) liegt in app/kasse_app.py und ruft diese Funktionen.

Umsaetze sind grundsaetzlich NETTO: apu * menge * (1 - rabatt_prozent/100).
Es zaehlen nur echte Bestell-Positionen (bestellart='Bestellung') aus nicht
stornierten Verkaeufen - identisch zur Logik in app/report_app.py.
"""
from __future__ import annotations

import sqlite3
from datetime import date, datetime
from pathlib import Path

from .config import OUTPUT_DIR


# ── Formatierung ─────────────────────────────────────────────────────────────
def _eur(v, suffix=" €") -> str:
    """Deutsche Geldformatierung: 1.234,56 €."""
    try:
        s = f"{float(v):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except (TypeError, ValueError):
        s = "0,00"
    return s + suffix


def _pdf_text(s) -> str:
    """Macht Text fuer die fpdf2-Kernschrift (latin-1) sicher: Euro-Zeichen und
    Sonderstriche ersetzen, Rest notfalls verlustfrei ersetzen."""
    s = str(s)
    for a, b in (("€", "EUR"), ("–", "-"), ("—", "-"), ("’", "'"), ("→", "->"),
                 ("✓", "OK"), ("·", "-")):
        s = s.replace(a, b)
    return s.encode("latin-1", "replace").decode("latin-1")


# ── Verfall ──────────────────────────────────────────────────────────────────
def verfall_to_date(verfall) -> date | None:
    """'MM/JJ' bzw. 'MM/JJJJ' -> letzter Tag des Verfallmonats (Ware ist bis Ende
    des aufgedruckten Monats verwendbar). Ungueltig -> None."""
    s = str(verfall or "").strip()
    if "/" not in s:
        return None
    mon, _, jahr = s.partition("/")
    if not (mon.isdigit() and jahr.isdigit()):
        return None
    m = int(mon)
    j = int(jahr)
    if len(jahr) == 2:
        j += 2000
    if not (1 <= m <= 12):
        return None
    # erster Tag des Folgemonats minus 1 Tag = letzter Tag des Verfallmonats.
    from calendar import monthrange
    return date(j, m, monthrange(j, m)[1])


# Horizont-Auswahl im Verfall-Reiter -> Warn-/Anzeige-Fenster in Tagen.
# "Alle" = kein Horizont (alles anzeigen). Standard 3 Monate = 90 Tage.
VERFALL_MONATE = {"3 Monate": 90, "6 Monate": 180, "9 Monate": 270,
                  "12 Monate": 365, "Alle": None}


def verfall_rows(db_path, heute: date | None = None, warn_tage: int = 90) -> list[dict]:
    """Lagerbestand mit Verfall-Bewertung, sortiert nach Verfalldatum (frueheste
    zuerst, Zeilen ohne Verfalldatum ganz hinten). warn_tage = ab wann 'bald'.
    Jede Zeile enthaelt auch apu und wert (apu*menge) fuer die Summen."""
    heute = heute or date.today()
    with sqlite3.connect(db_path) as con:
        rows = con.execute(
            "SELECT l.pzn, COALESCE(l.artikelname,''), COALESCE(l.charge,''), "
            "COALESCE(l.verfall,''), l.menge, "
            "(SELECT apu FROM tbl_nmg_stamm n WHERE n.pzn=l.pzn LIMIT 1) AS apu, l.ek "
            "FROM tbl_lagerbestand l WHERE l.menge > 0"
        ).fetchall()
    out = []
    for pzn, art, charge, verfall, menge, apu, ek in rows:
        d = verfall_to_date(verfall)
        tage = (d - heute).days if d else None
        if tage is None:
            status = "ohne"
        elif tage < 0:
            status = "abgelaufen"
        elif tage <= warn_tage:
            status = "bald"
        else:
            status = "ok"
        out.append({"pzn": pzn, "artikelname": art, "charge": charge, "verfall": verfall,
                    "menge": menge, "apu": apu, "wert": (apu or 0) * (menge or 0),
                    "ek": ek, "lagerwert": (ek or 0) * (menge or 0),
                    "datum": d, "tage": tage, "status": status})
    out.sort(key=lambda r: (r["datum"] is None, r["datum"] or date.max))
    return out


def inventur_rows(db_path) -> list[dict]:
    """Kompletter Lagerbestand (menge<>0) mit apu/wert (Verkaufswert) und ek/lagerwert
    (Einkaufswert), fuer Inventur-Liste + Summen."""
    with sqlite3.connect(db_path) as con:
        rows = con.execute(
            "SELECT l.pzn, COALESCE(l.artikelname,''), COALESCE(l.charge,''), "
            "COALESCE(l.verfall,''), l.menge, "
            "(SELECT apu FROM tbl_nmg_stamm n WHERE n.pzn=l.pzn LIMIT 1) AS apu, l.ek "
            "FROM tbl_lagerbestand l WHERE l.menge <> 0 ORDER BY l.artikelname, l.verfall"
        ).fetchall()
    return [{"pzn": p, "artikelname": a, "charge": c, "verfall": v, "menge": m,
             "apu": apu, "wert": (apu or 0) * (m or 0),
             "ek": ek, "lagerwert": (ek or 0) * (m or 0)}
            for p, a, c, v, m, apu, ek in rows]


# ── Umsatz / Tagesabschluss ──────────────────────────────────────────────────
_UMSATZ_BASE = (
    "SELECT {periode} AS periode, COUNT(DISTINCT b.id) AS anzahl, "
    "COALESCE(SUM(p.apu*p.menge),0) AS brutto, "
    "COALESCE(SUM(p.apu*p.menge*COALESCE(p.rabatt_prozent,0)/100.0),0) AS rabatt, "
    "COALESCE(SUM(p.apu*p.menge*(1-COALESCE(p.rabatt_prozent,0)/100.0)),0) AS netto, "
    "COALESCE(SUM(p.menge),0) AS pakete "
    "FROM tbl_bestellpositionen p JOIN tbl_bestellungen b ON b.id=p.bestell_id "
    "WHERE COALESCE(p.bestellart,'Bestellung')='Bestellung' "
    "AND COALESCE(b.status,'offen')<>'storniert' AND p.apu IS NOT NULL "
)

_PERIODE_SQL = {
    "Tag": "substr(b.datum,1,10)",
    "Monat": "substr(b.datum,1,7)",
    "Jahr": "substr(b.datum,1,4)",
}


def umsatz_rows(db_path, granularitaet="Tag", von=None, bis=None) -> list[dict]:
    """Umsatz gruppiert nach Tag/Monat/Jahr. von/bis als 'YYYY-MM-DD' (optional).
    Liefert je Periode: anzahl Verkaeufe, Brutto (APU x Menge), Rabatt gegeben,
    Netto-Umsatz."""
    periode = _PERIODE_SQL.get(granularitaet, _PERIODE_SQL["Tag"])
    sql = _UMSATZ_BASE.format(periode=periode)
    params: list = []
    if von:
        sql += "AND b.datum >= ? "
        params.append(von)
    if bis:
        sql += "AND b.datum <= ? "
        params.append(bis)
    sql += "GROUP BY periode ORDER BY periode DESC"
    with sqlite3.connect(db_path) as con:
        rows = con.execute(sql, params).fetchall()
    return [{"periode": r[0] or "", "anzahl": r[1], "brutto": r[2],
             "rabatt": r[3], "netto": r[4], "pakete": int(r[5] or 0)} for r in rows]


def tagesabschluss_data(db_path, tag: str) -> dict:
    """Kennzahlen + Positionen fuer einen einzelnen Tag ('YYYY-MM-DD')."""
    with sqlite3.connect(db_path) as con:
        head = con.execute(
            _UMSATZ_BASE.format(periode="substr(b.datum,1,10)") +
            "AND substr(b.datum,1,10)=? GROUP BY periode", (tag,)).fetchone()
        positionen = con.execute(
            "SELECT p.pzn, COALESCE(p.artikelname,''), SUM(p.menge), "
            "SUM(p.apu*p.menge), SUM(p.apu*p.menge*COALESCE(p.rabatt_prozent,0)/100.0), "
            "SUM(p.apu*p.menge*(1-COALESCE(p.rabatt_prozent,0)/100.0)) "
            "FROM tbl_bestellpositionen p JOIN tbl_bestellungen b ON b.id=p.bestell_id "
            "WHERE COALESCE(p.bestellart,'Bestellung')='Bestellung' "
            "AND COALESCE(b.status,'offen')<>'storniert' AND p.apu IS NOT NULL "
            "AND substr(b.datum,1,10)=? GROUP BY p.pzn, p.artikelname "
            "ORDER BY SUM(p.apu*p.menge*(1-COALESCE(p.rabatt_prozent,0)/100.0)) DESC",
            (tag,)).fetchall()
    return {
        "tag": tag,
        "anzahl": head[1] if head else 0,
        "pakete": int(sum((p[2] or 0) for p in positionen)),
        "brutto": head[2] if head else 0.0,
        "rabatt": head[3] if head else 0.0,
        "netto": head[4] if head else 0.0,
        "positionen": positionen,
    }


def _ensure_ta_table(con):
    con.execute(
        "CREATE TABLE IF NOT EXISTS tbl_kasse_tagesabschluss("
        "datum TEXT PRIMARY KEY, nr INTEGER, erzeugt_am TEXT, erzeugt_von TEXT, "
        "anzahl INTEGER, pakete INTEGER, brutto REAL, rabatt REAL, netto REAL)")


def tagesabschluss_nr(db_path, tag, assign=False, data=None):
    """Laufende Tagesabschluss-Nr fuer einen Tag. assign=False -> vorhandene Nr
    oder None. assign=True -> bei Bedarf neue Nr (max+1) vergeben + Kennzahlen
    festschreiben, dann zurueckgeben."""
    import getpass
    with sqlite3.connect(db_path) as con:
        _ensure_ta_table(con)
        row = con.execute("SELECT nr FROM tbl_kasse_tagesabschluss WHERE datum=?", (tag,)).fetchone()
        if row:
            return row[0]
        if not assign:
            return None
        nr = (con.execute("SELECT COALESCE(MAX(nr),0) FROM tbl_kasse_tagesabschluss")
              .fetchone()[0]) + 1
        d = data or tagesabschluss_data(db_path, tag)
        con.execute(
            "INSERT INTO tbl_kasse_tagesabschluss(datum, nr, erzeugt_am, erzeugt_von, "
            "anzahl, pakete, brutto, rabatt, netto) VALUES(?,?,?,?,?,?,?,?,?)",
            (tag, nr, datetime.now().isoformat(timespec="seconds"), getpass.getuser(),
             d["anzahl"], d["pakete"], d["brutto"], d["rabatt"], d["netto"]))
        con.commit()
        return nr


def tage_mit_umsatz(db_path) -> list[str]:
    """Alle Tage ('YYYY-MM-DD'), an denen es nicht stornierte Bestell-Umsaetze
    gab (aufsteigend) - fuer den automatischen Tagesabschluss-Nachlauf."""
    with sqlite3.connect(db_path) as con:
        rows = con.execute(
            "SELECT DISTINCT substr(b.datum,1,10) "
            "FROM tbl_bestellpositionen p JOIN tbl_bestellungen b ON b.id=p.bestell_id "
            "WHERE COALESCE(p.bestellart,'Bestellung')='Bestellung' "
            "AND COALESCE(b.status,'offen')<>'storniert' AND p.apu IS NOT NULL "
            "ORDER BY 1").fetchall()
    return [r[0] for r in rows if r[0]]


# ── PDF-Bausteine ────────────────────────────────────────────────────────────
def _new_pdf(orientation="P"):
    from fpdf import FPDF
    pdf = FPDF(orientation=orientation, unit="mm", format="A4")
    pdf.set_auto_page_break(auto=True, margin=14)
    pdf.add_page()
    return pdf


def _titel(pdf, titel, untertitel=""):
    from fpdf.enums import XPos, YPos
    pdf.set_font("Helvetica", "B", 18)
    pdf.set_text_color(11, 74, 134)
    pdf.cell(0, 10, _pdf_text(titel), new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    if untertitel:
        pdf.set_font("Helvetica", "", 10)
        pdf.set_text_color(120, 120, 120)
        pdf.cell(0, 6, _pdf_text(untertitel), new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.ln(2)


def _tabelle(pdf, spalten, zeilen, kopf_fill=(11, 74, 134)):
    """spalten = [(titel, breite_mm, align)], zeilen = [(zell-tuple, fill_rgb|None)]."""
    from fpdf.enums import XPos, YPos
    pdf.set_font("Helvetica", "B", 9)
    pdf.set_fill_color(*kopf_fill)
    pdf.set_text_color(255, 255, 255)
    for titel, breite, _align in spalten:
        pdf.cell(breite, 7, _pdf_text(titel), border=0, align="C", fill=True)
    pdf.ln(7)
    pdf.set_font("Helvetica", "", 9)
    for zellen, fill in zeilen:
        if fill:
            pdf.set_fill_color(*fill)
        pdf.set_text_color(40, 40, 40)
        for (titel, breite, align), wert in zip(spalten, zellen):
            pdf.cell(breite, 6, _pdf_text(wert), border="B", align=align, fill=bool(fill))
        pdf.ln(6)


def _summenzeile(pdf, text):
    from fpdf.enums import XPos, YPos
    pdf.ln(2)
    pdf.set_font("Helvetica", "B", 10)
    pdf.set_text_color(11, 74, 134)
    pdf.cell(0, 7, _pdf_text(text), new_x=XPos.LMARGIN, new_y=YPos.NEXT)


def verfall_pdf(db_path, out_path=None, heute: date | None = None,
                monate: str = "Alle") -> Path:
    """Verfall-Report: abgelaufene und bald ablaufende Chargen (rot/gelb). monate
    waehlt den Horizont (siehe VERFALL_MONATE); 'Alle' zeigt den ganzen Bestand."""
    heute = heute or date.today()
    warn_tage = VERFALL_MONATE.get(monate, 90)
    rows = verfall_rows(db_path, heute, warn_tage if warn_tage is not None else 90)
    if warn_tage is not None:
        # Nur abgelaufene + im Horizont ablaufende anzeigen.
        rows = [r for r in rows if r["status"] in ("abgelaufen", "bald")]
    out_path = Path(out_path) if out_path else (
        OUTPUT_DIR / "auswertungen" / f"Verfall_{heute:%Y-%m-%d}.pdf")
    out_path.parent.mkdir(parents=True, exist_ok=True)

    abgelaufen = sum(1 for r in rows if r["status"] == "abgelaufen")
    bald = sum(1 for r in rows if r["status"] == "bald")
    horizont = "ganzer Bestand" if warn_tage is None else f"Horizont {monate}"
    pdf = _new_pdf("P")
    _titel(pdf, "Verfall-Übersicht (Lager)",
           f"Stand {heute:%d.%m.%Y}  -  {horizont}  -  {abgelaufen} abgelaufen, {bald} bald")
    spalten = [("PZN", 22, "L"), ("Artikel", 66, "L"), ("Charge", 24, "L"),
               ("Verfall", 18, "C"), ("Bestand", 16, "R"), ("APU", 18, "R"),
               ("Wert", 22, "R"), ("Status", 24, "L")]
    zeilen = []
    for r in rows:
        fill = {"abgelaufen": (250, 219, 216), "bald": (252, 243, 207)}.get(r["status"])
        stat = {"abgelaufen": "abgelaufen", "ok": "ok", "ohne": "-"}.get(
            r["status"], f"in {r['tage']} T.")
        zeilen.append(((r["pzn"], r["artikelname"], r["charge"], r["verfall"] or "-",
                        str(r["menge"]), _eur(r["apu"], " EUR") if r["apu"] is not None else "-",
                        _eur(r["wert"], " EUR"), stat), fill))
    _tabelle(pdf, spalten, zeilen)
    bestand = sum(r["menge"] or 0 for r in rows)
    wert = sum(r["wert"] for r in rows)
    lagerwert = sum(r["lagerwert"] for r in rows)
    _summenzeile(pdf, f"Summe Bestand: {bestand}    Verkaufswert (APU x Bestand): "
                 f"{_eur(wert, ' EUR')}    Lagerwert (EK x Bestand): {_eur(lagerwert, ' EUR')}")
    pdf.output(str(out_path))
    return out_path


def inventur_pdf(db_path, out_path=None, heute: date | None = None) -> Path:
    """Inventur-/Lagerbestandsliste zum Abhaken (Soll-Bestand, EK/Verkaufswerte,
    Zaehl-Spalte). Querformat fuer die zusaetzlichen Wert-Spalten."""
    heute = heute or date.today()
    rows = inventur_rows(db_path)
    out_path = Path(out_path) if out_path else (
        OUTPUT_DIR / "auswertungen" / f"Inventur_{heute:%Y-%m-%d}.pdf")
    out_path.parent.mkdir(parents=True, exist_ok=True)

    pdf = _new_pdf("L")
    gesamt = sum(r["menge"] or 0 for r in rows)
    wert = sum(r["wert"] for r in rows)
    lagerwert = sum(r["lagerwert"] for r in rows)
    _titel(pdf, "Inventur / Lagerbestand",
           f"Stand {heute:%d.%m.%Y}  -  {len(rows)} Positionen, Gesamtmenge {gesamt}")
    spalten = [("PZN", 24, "L"), ("Artikel", 62, "L"), ("Charge", 26, "L"),
               ("Verfall", 18, "C"), ("Soll", 14, "R"), ("APU", 22, "R"),
               ("Verkaufswert", 28, "R"), ("EK", 22, "R"), ("Lagerwert", 28, "R"),
               ("Gezählt", 24, "C")]
    zeilen = [((r["pzn"], r["artikelname"], r["charge"], r["verfall"] or "-", str(r["menge"]),
                _eur(r["apu"], " EUR") if r["apu"] is not None else "-",
                _eur(r["wert"], " EUR"),
                _eur(r["ek"], " EUR") if r["ek"] is not None else "-",
                _eur(r["lagerwert"], " EUR"), ""),
               (245, 249, 253) if i % 2 else None) for i, r in enumerate(rows)]
    _tabelle(pdf, spalten, zeilen)
    _summenzeile(pdf, f"Summe Bestand: {gesamt}    Verkaufswert (APU x Bestand): "
                 f"{_eur(wert, ' EUR')}    Lagerwert (EK x Bestand): {_eur(lagerwert, ' EUR')}")
    pdf.output(str(out_path))
    return out_path


def umsatz_pdf(db_path, granularitaet="Monat", von=None, bis=None, out_path=None) -> Path:
    rows = umsatz_rows(db_path, granularitaet, von, bis)
    heute = date.today()
    out_path = Path(out_path) if out_path else (
        OUTPUT_DIR / "auswertungen" / f"Umsatz_{granularitaet}_{heute:%Y-%m-%d}.pdf")
    out_path.parent.mkdir(parents=True, exist_ok=True)

    pdf = _new_pdf("P")
    zr = f"{von or '...'} bis {bis or '...'}" if (von or bis) else "gesamter Zeitraum"
    _titel(pdf, f"Umsatzübersicht je {granularitaet}", f"Stand {heute:%d.%m.%Y}  -  {zr}")
    spalten = [("Zeitraum", 30, "L"), ("Anzahl Verkäufe", 28, "R"),
               ("Anzahl Packungen", 30, "R"), ("APU Brutto", 34, "R"),
               ("Rabatt (Netto)", 32, "R"), ("APU Netto", 34, "R")]
    zeilen = [((r["periode"], str(r["anzahl"]), str(r["pakete"]), _eur(r["brutto"], " EUR"),
                _eur(r["rabatt"], " EUR"), _eur(r["netto"], " EUR")),
               (245, 249, 253) if i % 2 else None) for i, r in enumerate(rows)]
    # Summenzeile.
    if rows:
        zeilen.append(((
            "Summe", str(sum(r["anzahl"] for r in rows)),
            str(sum(r["pakete"] for r in rows)),
            _eur(sum(r["brutto"] for r in rows), " EUR"),
            _eur(sum(r["rabatt"] for r in rows), " EUR"),
            _eur(sum(r["netto"] for r in rows), " EUR")), (216, 234, 247)))
    _tabelle(pdf, spalten, zeilen)
    pdf.output(str(out_path))
    return out_path


# ── Auftrags-Historie / Protokoll (wer hat was wann gemacht) ─────────────────
def auftrag_historie(db_path, bestell_id):
    """Kopfdaten eines Verkaufs + alle Protokoll-Eintraege dazu (chronologisch).
    Liefert (kopf|None, [(zeitpunkt, bearbeiter, aktion, details), ...])."""
    with sqlite3.connect(db_path) as con:
        kopf = con.execute(
            "SELECT id, datum, COALESCE(apotheke,''), COALESCE(status,'offen'), "
            "COALESCE(bearbeiter,''), COALESCE(msk_erfasst,0), COALESCE(msk_von,''), "
            "COALESCE(msk_am,'') FROM tbl_bestellungen WHERE id=?", (bestell_id,)).fetchone()
        eintraege = con.execute(
            "SELECT zeitpunkt, COALESCE(bearbeiter,''), COALESCE(aktion,''), COALESCE(details,'') "
            "FROM tbl_kasse_log WHERE bestell_id=? ORDER BY id", (bestell_id,)).fetchall()
    return kopf, eintraege


def auftrag_historie_pdf(db_path, bestell_id, out_path=None) -> Path:
    """Auftrags-Verlauf als PDF: wer hat was wann mit diesem Auftrag gemacht."""
    kopf, eintraege = auftrag_historie(db_path, bestell_id)
    out_path = Path(out_path) if out_path else (
        OUTPUT_DIR / "auswertungen" / f"Auftrag_{bestell_id}_Verlauf.pdf")
    out_path.parent.mkdir(parents=True, exist_ok=True)

    apotheke = kopf[2] if kopf else ""
    status = kopf[3] if kopf else ""
    msk = ("erfasst von " + (kopf[6] or "?") + " am " + (kopf[7] or "")[:16].replace("T", " ")) \
        if (kopf and kopf[5]) else "nicht erfasst"
    pdf = _new_pdf("P")
    _titel(pdf, f"Auftrag #{bestell_id} - Verlauf",
           f"{apotheke}  -  Status {status}  -  MSK: {msk}")
    spalten = [("Zeitpunkt", 34, "L"), ("Mitarbeiter", 34, "L"),
               ("Aktion", 44, "L"), ("Details", 70, "L")]
    zeilen = [((str(z).replace("T", " ")[:16], b, a, d), (245, 249, 253) if i % 2 else None)
              for i, (z, b, a, d) in enumerate(eintraege)]
    if not zeilen:
        zeilen = [(("-", "-", "keine Protokoll-Einträge", ""), None)]
    _tabelle(pdf, spalten, zeilen)
    pdf.output(str(out_path))
    return out_path


def protokoll_pdf(zeilen, out_path=None, titel="Änderungs-Protokoll", untertitel="") -> Path:
    """Protokoll-Auswertung als PDF aus bereits gefilterten Zeilen.
    zeilen = [(zeit, bearbeiter, aktion, auftrag, kunde, details), ...]."""
    out_path = Path(out_path) if out_path else (
        OUTPUT_DIR / "auswertungen" / f"Protokoll_{date.today():%Y-%m-%d}.pdf")
    out_path.parent.mkdir(parents=True, exist_ok=True)
    pdf = _new_pdf("L")  # Querformat - viele Spalten
    _titel(pdf, titel, untertitel or f"Stand {date.today():%d.%m.%Y}  -  {len(zeilen)} Einträge")
    spalten = [("Zeitpunkt", 32, "L"), ("Mitarbeiter", 34, "L"), ("Aktion", 40, "L"),
               ("Auftrag", 18, "L"), ("Kunde", 50, "L"), ("Details", 105, "L")]
    pdf_zeilen = [((z[0], z[1], z[2], z[3], z[4], z[5]), (245, 249, 253) if i % 2 else None)
                  for i, z in enumerate(zeilen)]
    if not pdf_zeilen:
        pdf_zeilen = [(("-", "-", "keine Einträge", "", "", ""), None)]
    _tabelle(pdf, spalten, pdf_zeilen)
    pdf.output(str(out_path))
    return out_path


def tagesabschluss_pdf_path(tag: str) -> Path:
    """Zielpfad des Tagesabschluss-PDFs fuer einen Tag (auch zur Existenzpruefung
    durch den automatischen Lauf)."""
    return OUTPUT_DIR / "tagesabschluss" / f"Tagesabschluss_{tag}.pdf"


def tagesabschluss_pdf(db_path, tag: str, out_path=None) -> Path:
    """Tagesabschluss: Kennzahlen des Tages + Artikel-Aufstellung."""
    from fpdf.enums import XPos, YPos
    data = tagesabschluss_data(db_path, tag)
    nr = tagesabschluss_nr(db_path, tag, assign=True, data=data)
    out_path = Path(out_path) if out_path else tagesabschluss_pdf_path(tag)
    out_path.parent.mkdir(parents=True, exist_ok=True)

    try:
        tag_anzeige = datetime.strptime(tag, "%Y-%m-%d").strftime("%d.%m.%Y")
    except ValueError:
        tag_anzeige = tag
    pdf = _new_pdf("P")
    _titel(pdf, f"Tagesabschluss Nr. {nr}",
           f"{tag_anzeige}  -  {data['anzahl']} Verkäufe, {data['pakete']} Packungen")

    # Kennzahl-Boxen.
    y0 = pdf.get_y()
    boxen = [("APU-Summe", data["brutto"]), ("Rabatt-Summe", data["rabatt"]),
             ("Umsatz-Summe (netto)", data["netto"])]
    gap = 5
    bw = (pdf.epw - 2 * gap) / 3
    for i, (label, wert) in enumerate(boxen):
        x = pdf.l_margin + i * (bw + gap)
        pdf.set_fill_color(232, 245, 238)
        pdf.rect(x, y0, bw, 22, style="F")
        pdf.set_xy(x + 3, y0 + 3)
        pdf.set_font("Helvetica", "", 9)
        pdf.set_text_color(90, 90, 90)
        pdf.cell(bw - 6, 5, _pdf_text(label), new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        pdf.set_xy(x + 3, y0 + 10)
        pdf.set_font("Helvetica", "B", 15)
        pdf.set_text_color(30, 142, 90)
        pdf.cell(bw - 6, 9, _pdf_text(_eur(wert, " EUR")))
    pdf.set_y(y0 + 22 + 5)

    spalten = [("PZN", 26, "L"), ("Artikel", 86, "L"), ("Menge", 18, "R"),
               ("APU brutto", 28, "R"), ("Umsatz netto", 32, "R")]
    zeilen = [((p[0], p[1], str(int(p[2] or 0)), _eur(p[3], " EUR"), _eur(p[5], " EUR")),
               (245, 249, 253) if i % 2 else None)
              for i, p in enumerate(data["positionen"])]
    _tabelle(pdf, spalten, zeilen)
    pdf.output(str(out_path))
    return out_path
