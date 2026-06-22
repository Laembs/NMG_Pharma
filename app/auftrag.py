"""Auftragsbestaetigung aus einem gespeicherten Verkauf erzeugen (HTML) und
drucken/per E-Mail (Outlook) versenden.

Die Vorlage ist austauschbar: eine HTML-Datei mit {{platzhaltern}}. Beim ersten
Aufruf wird die mitgelieferte Standardvorlage (assets/) in einen nutzer-
beschreibbaren Ordner 'vorlagen/' kopiert; dort kann sie frei angepasst werden.
"""
from __future__ import annotations

import html as _html
import shutil
import sqlite3
import sys
from datetime import datetime
from pathlib import Path

from .config import ASSETS_DIR, OUTPUT_DIR, USERDATA_ROOT, BASE_DIR, DB_PATH

_ROOT = USERDATA_ROOT or BASE_DIR
VORLAGEN_DIR = _ROOT / "vorlagen"
TEMPLATE_NAME = "auftragsbestaetigung.html"


def template_path() -> Path:
    """Pfad zur (nutzer-editierbaren) Vorlage; legt sie beim ersten Mal aus der
    mitgelieferten Standardvorlage an."""
    VORLAGEN_DIR.mkdir(parents=True, exist_ok=True)
    p = VORLAGEN_DIR / TEMPLATE_NAME
    if not p.exists():
        default = ASSETS_DIR / "auftragsbestaetigung_vorlage.html"
        if default.exists():
            shutil.copy2(default, p)
    return p


def _eur(v):
    if v is None:
        return ""
    return f"{v:,.2f} €".replace(",", "X").replace(".", ",").replace("X", ".")


def _load(db_path, bestell_id):
    with sqlite3.connect(db_path) as con:
        h = con.execute(
            "SELECT id, datum, kundennummer, apotheke, bestellart, lieferzeit, liefertermin "
            "FROM tbl_bestellungen WHERE id=?", (bestell_id,)
        ).fetchone()
        if not h:
            raise ValueError(f"Verkauf #{bestell_id} nicht gefunden.")
        keys = ("id", "datum", "kundennummer", "apotheke", "bestellart", "lieferzeit", "liefertermin")
        header = dict(zip(keys, h))
        positions = [
            dict(zip(("pzn", "artikelname", "df", "pck", "apu", "menge", "rabatt", "charge",
                      "verfall", "bestellart", "lieferzeit", "liefertermin"), r))
            for r in con.execute(
                "SELECT pzn, artikelname, df, pck, apu, menge, rabatt_prozent, charge, verfall, "
                "COALESCE(bestellart,'Bestellung'), COALESCE(lieferzeit,''), COALESCE(liefertermin,'') "
                "FROM tbl_bestellpositionen WHERE bestell_id=?", (bestell_id,))
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
    """Erzeugt die Auftragsbestaetigung als HTML-Datei und gibt den Pfad zurueck."""
    header, positions, kunde = _load(db_path, bestell_id)
    # Vorbestellungen (und abgesagte) gehoeren NICHT auf die Auftragsbestaetigung -
    # nur tatsaechlich gelieferte Bestellungen.
    positions = [p for p in positions if (p.get("bestellart") or "Bestellung") == "Bestellung"]
    tpl = template_path()
    text = tpl.read_text(encoding="utf-8")

    rows = []
    gesamt = 0.0
    for i, p in enumerate(positions, 1):
        apu = p["apu"]
        rab = p["rabatt"] or 0
        summe = None
        if apu is not None:
            summe = apu * (p["menge"] or 0) * (1 - rab / 100.0)
            gesamt += summe
        liefer = " · ".join(x for x in [p.get("bestellart") or "", p.get("lieferzeit") or "",
                                        p.get("liefertermin") or ""] if x)
        rows.append(
            "<tr><td>{pos}</td><td>{pzn}</td><td>{art}</td><td>{liefer}</td>"
            "<td class='r'>{menge}</td><td class='r'>{apu}</td><td class='r'>{rab}</td>"
            "<td class='r'>{summe}</td></tr>".format(
                pos=i, pzn=_html.escape(str(p["pzn"] or "")),
                art=_html.escape(str(p["artikelname"] or "")),
                liefer=_html.escape(liefer),
                menge=p["menge"] or 0,
                apu=_eur(apu), rab=f"{rab:.0f}", summe=_eur(summe)))

    apotheke = kunde.get("kundenname") or header.get("apotheke") or ""
    repl = {
        "auftragsnr": str(header["id"]),
        "datum": str(header.get("datum") or datetime.now().strftime("%Y-%m-%d")),
        "kundennummer": header.get("kundennummer") or "",
        "apotheke": apotheke,
        "inhaber": kunde.get("inhaber") or "",
        "strasse": kunde.get("strasse") or "",
        "plz": kunde.get("plz") or "",
        "ort": kunde.get("ort") or "",
        "bestellart": header.get("bestellart") or "",
        "lieferzeit": header.get("lieferzeit") or "",
        "liefertermin": header.get("liefertermin") or "",
        "gesamt": _eur(gesamt),
    }
    for k, v in repl.items():
        text = text.replace("{{%s}}" % k, _html.escape(str(v)))
    text = text.replace("{{positionen}}", "".join(rows))

    out_dir = OUTPUT_DIR / "auftragsbestaetigungen"
    out_dir.mkdir(parents=True, exist_ok=True)
    safe = "".join(c for c in apotheke if c.isalnum() or c in " _-").strip().replace(" ", "_")[:40]
    path = out_dir / f"Auftrag_{header['id']}_{safe or 'kunde'}.html"
    path.write_text(text, encoding="utf-8")
    return path


def kunde_email(db_path, bestell_id) -> str:
    _h, _p, kunde = _load(db_path, bestell_id)
    return (kunde.get("email") or "").strip()


def send_via_outlook(to, subject, html_body, attachment=None):
    """Oeffnet eine Outlook-Mail (wie NMGones bestehender Versand). Raised bei
    Fehler -> die GUI zeigt die Meldung."""
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
