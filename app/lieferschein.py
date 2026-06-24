"""Lieferschein aus einem gespeicherten Verkauf erzeugen (HTML), drucken/oeffnen
und per E-Mail (Outlook) versenden.

Bewusst getrennt von der Auftragsbestaetigung (app/auftrag.py): der Lieferschein
begleitet die Ware und zeigt KEINE Preise, dafuer Charge/Verfall je Position.
Wird erzeugt, sobald ein Verkauf in MSK erfasst und damit zur Lieferung
freigegeben wurde (kasse_app._msk_markieren).

Die Vorlage ist - wie bei der Auftragsbestaetigung - eine austauschbare
HTML-Datei mit {{platzhaltern}} im nutzer-beschreibbaren Ordner 'vorlagen/'.
"""
from __future__ import annotations

import html as _html
import shutil
import sqlite3
import sys
from datetime import datetime
from pathlib import Path

from .config import ASSETS_DIR, OUTPUT_DIR, USERDATA_ROOT, BASE_DIR, DB_PATH
from . import einstellungen

_ROOT = USERDATA_ROOT or BASE_DIR
VORLAGEN_DIR = _ROOT / "vorlagen"
TEMPLATE_NAME = "lieferschein.html"


def template_path() -> Path:
    """Pfad zur (nutzer-editierbaren) Lieferschein-Vorlage; legt sie beim ersten
    Mal aus der mitgelieferten Standardvorlage an."""
    VORLAGEN_DIR.mkdir(parents=True, exist_ok=True)
    p = VORLAGEN_DIR / TEMPLATE_NAME
    if not p.exists():
        default = ASSETS_DIR / "lieferschein_vorlage.html"
        if default.exists():
            shutil.copy2(default, p)
    return p


def _load(db_path, bestell_id):
    """Kopf, Positionen und Kundendaten eines Verkaufs laden. Spiegelt
    auftrag._load, holt aber zusaetzlich den MSK-Status fuer den Lieferschein-Fuss."""
    with sqlite3.connect(db_path) as con:
        h = con.execute(
            "SELECT id, datum, kundennummer, apotheke, bestellart, lieferzeit, liefertermin, "
            "COALESCE(msk_von,''), COALESCE(msk_am,'') "
            "FROM tbl_bestellungen WHERE id=?", (bestell_id,)
        ).fetchone()
        if not h:
            raise ValueError(f"Verkauf #{bestell_id} nicht gefunden.")
        keys = ("id", "datum", "kundennummer", "apotheke", "bestellart", "lieferzeit",
                "liefertermin", "msk_von", "msk_am")
        header = dict(zip(keys, h))
        positions = [
            dict(zip(("pzn", "artikelname", "df", "pck", "menge", "charge", "verfall", "bestellart"), r))
            for r in con.execute(
                "SELECT pzn, artikelname, df, pck, menge, COALESCE(charge,''), COALESCE(verfall,''), "
                "COALESCE(bestellart,'Bestellung') "
                "FROM tbl_bestellpositionen WHERE bestell_id=? ORDER BY id", (bestell_id,))
        ]
        kunde = {}
        knr = header.get("kundennummer")
        exists = con.execute(
            "SELECT 1 FROM sqlite_master WHERE type='table' AND name='tbl_kunden_center'").fetchone()
        if knr and exists:
            have = {r[1] for r in con.execute("PRAGMA table_info(tbl_kunden_center)")}
            want = [c for c in ("kundenname", "plz", "ort", "strasse", "inhaber", "email") if c in have]
            if want:
                row = con.execute(
                    f"SELECT {','.join(want)} FROM tbl_kunden_center WHERE kundennummer=? LIMIT 1",
                    (knr,)).fetchone()
                if row:
                    kunde = dict(zip(want, row))
    return header, positions, kunde


def render(db_path=DB_PATH, bestell_id=None) -> Path:
    """Erzeugt den Lieferschein als HTML-Datei und gibt den Pfad zurueck.
    Nur tatsaechlich gelieferte Positionen (Bestellung) - Vorbestellungen/abgesagte
    gehoeren nicht auf den Lieferschein."""
    header, positions, kunde = _load(db_path, bestell_id)
    positions = [p for p in positions if (p.get("bestellart") or "Bestellung") == "Bestellung"]
    tpl = template_path()
    text = tpl.read_text(encoding="utf-8")

    rows = []
    gesamt_menge = 0
    for i, p in enumerate(positions, 1):
        gesamt_menge += p.get("menge") or 0
        rows.append(
            "<tr><td>{pos}</td><td>{pzn}</td><td>{art}</td><td>{pck}</td>"
            "<td>{charge}</td><td>{verfall}</td><td class='r'>{menge}</td></tr>".format(
                pos=i, pzn=_html.escape(str(p["pzn"] or "")),
                art=_html.escape(str(p["artikelname"] or "")),
                pck=_html.escape(str(p.get("pck") or "")),
                charge=_html.escape(str(p.get("charge") or "—")),
                verfall=_html.escape(str(p.get("verfall") or "—")),
                menge=p["menge"] or 0))

    apotheke = kunde.get("kundenname") or header.get("apotheke") or ""
    msk_am = (header.get("msk_am") or "")[:16].replace("T", " ")
    repl = {
        "lieferscheinnr": str(header["id"]),
        "auftragsnr": str(header["id"]),
        "datum": datetime.now().strftime("%d.%m.%Y"),
        "kundennummer": header.get("kundennummer") or "",
        "apotheke": apotheke,
        "inhaber": kunde.get("inhaber") or "",
        "strasse": kunde.get("strasse") or "",
        "plz": kunde.get("plz") or "",
        "ort": kunde.get("ort") or "",
        "bestellart": header.get("bestellart") or "",
        "lieferzeit": header.get("lieferzeit") or "",
        "liefertermin": header.get("liefertermin") or "",
        "gesamt_menge": str(gesamt_menge),
        "msk_von": header.get("msk_von") or "",
        "msk_am": msk_am,
        "firma": einstellungen.get(db_path, "firma_name"),
        "absender_kontakt": einstellungen.get(db_path, "firma_kontakt"),
    }
    for k, v in repl.items():
        text = text.replace("{{%s}}" % k, _html.escape(str(v)))
    text = text.replace("{{absender_adresse}}", einstellungen.absender_adresse_html(db_path))
    text = text.replace("{{positionen}}", "".join(rows))

    out_dir = OUTPUT_DIR / "lieferscheine"
    out_dir.mkdir(parents=True, exist_ok=True)
    safe = "".join(c for c in apotheke if c.isalnum() or c in " _-").strip().replace(" ", "_")[:40]
    path = out_dir / f"Lieferschein_{header['id']}_{safe or 'kunde'}.html"
    path.write_text(text, encoding="utf-8")
    return path


def kunde_email(db_path, bestell_id) -> str:
    _h, _p, kunde = _load(db_path, bestell_id)
    return (kunde.get("email") or "").strip()


def send_via_outlook(to, subject, html_body, attachment=None):
    """Oeffnet eine Outlook-Mail (wie auftrag.send_via_outlook). Raised bei Fehler."""
    if not sys.platform.startswith("win"):
        raise RuntimeError("Outlook-Versand ist nur unter Windows verfügbar.")
    import win32com.client  # pywin32, wie in gui.py
    outlook = win32com.client.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)
    if to:
        mail.To = to
    mail.Subject = subject
    mail.HTMLBody = html_body
    if attachment:
        mail.Attachments.Add(str(attachment))
    mail.Display(True)
