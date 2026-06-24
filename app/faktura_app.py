"""NMG Faktura-App (eigenstaendig) - rechtskonforme Rechnungen & Gutschriften.

Prototyp. Eigenes Fenster / Taskleisten-Icon (run_standalone), teilt sich die
Datenbank mit NMGone (app/config.py). Kerngedanken:

  * APU wird je Position als Wert EINGEFROREN (Snapshot zum Beleg, kein Live-Join)
  * Festgeschriebene Belege sind unveraenderbar (GoBD); Korrektur nur via Storno
  * Lueckenlose Nummernkreise je Belegart/Jahr
  * Firmenstammdaten (Steuernr/USt-IdNr/Logo/Versenderadresse) frei pflegbar
  * Mitarbeiter (Name + E-Mail) wird automatisch dem Beleg zugeordnet

Datenmodell siehe migrations.py (tbl_faktura_*).
"""
from __future__ import annotations

import base64
import getpass
import json
import os
import sqlite3
import subprocess
from datetime import datetime, date
from pathlib import Path
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog

from .config import DB_PATH, ASSETS_DIR, OUTPUT_DIR
from . import theme
from . import tour

# Palette aus dem zentralen Theme (gemeinsamer Look mit NMGone/Kasse).
BG = theme.CARD
SHELL_BG = theme.BG
ACCENT = theme.PRIMARY
ACCENT_LIGHT = theme.SELECT_BG
BORDER = theme.BORDER
TEXT = theme.INK
MUTED = theme.MUTED
OK_GREEN = theme.SUCCESS

# Firmenstammdaten: (schluessel, label, mehrzeilig?) - steuert die Einstellungsmaske.
STAMM_FELDER = [
    ("firma_name", "Firmenname", False),
    ("firma_strasse", "Strasse / Hausnr.", False),
    ("firma_plz", "PLZ", False),
    ("firma_ort", "Ort", False),
    ("firma_ustid", "USt-IdNr.", False),
    ("firma_steuernr", "Steuernummer", False),
    ("firma_telefon", "Telefon", False),
    ("firma_email", "E-Mail", False),
    ("firma_web", "Webseite", False),
    ("firma_bank", "Bank", False),
    ("firma_iban", "IBAN", False),
    ("firma_bic", "BIC", False),
    ("zahlungsziel_tage", "Zahlungsziel (Tage)", False),
    ("ust_satz_standard", "USt-Satz Standard (%)", False),
    ("rechnung_fusstext", "Fusstext der Rechnung", True),
]

STAMM_DEFAULTS = {"zahlungsziel_tage": "14", "ust_satz_standard": "19",
                  "rechnung_fusstext": "Zahlbar innerhalb des Zahlungsziels ohne Abzug. "
                                       "Aufbewahrungspflichtig 10 Jahre (GoBD)."}

# Frei konfigurierbare Belegnummern-Formate. Platzhalter:
#   {JJJJ}=Jahr 4-stellig, {JJ}=Jahr 2-stellig, {MM}=Monat, {NR}=Zähler, {NR:5}=Zähler 5-stellig
STD_NR_FORMAT = {
    "rechnung": "RE-{JJJJ}-{NR:5}",
    "gutschrift": "GU-{JJJJ}-{NR:5}",
    "storno": "ST-{JJJJ}-{NR:5}",
    "quartalsverguetung": "QV-{JJJJ}-{NR:5}",
}

# Frei gestaltbare Rechnungs-/Beleg-Vorlage (gilt für Rechnung, Gutschrift, Quartalsvergütung).
TPL_DEFAULTS = {
    "tpl_akzentfarbe": "#185fa5",
    "tpl_logo_pos": "links",
    "tpl_spalte_apu": "1",
    "tpl_spalte_ust": "1",
    "tpl_titel_rechnung": "Rechnung",
    "tpl_titel_gutschrift": "Gutschrift",
    "tpl_titel_storno": "Storno-Rechnung",
    "tpl_titel_quartalsverguetung": "Quartalsvergütung",
    "tpl_kopftext_rechnung": "",
    "tpl_kopftext_gutschrift": "Wir schreiben Ihnen folgenden Betrag gut:",
    "tpl_kopftext_quartalsverguetung": "Vergütung für Ihren Quartalsumsatz:",
    "tpl_fusstext": "Zahlbar innerhalb des Zahlungsziels ohne Abzug. "
                    "Aufbewahrungspflichtig 10 Jahre (GoBD).",
}
STAMM_DEFAULTS.update(TPL_DEFAULTS)


# ── Helfer ───────────────────────────────────────────────────────────────────
def _con() -> sqlite3.Connection:
    con = sqlite3.connect(DB_PATH)
    con.execute("PRAGMA foreign_keys = ON")
    return con


def _table_exists(con, name) -> bool:
    return con.execute("SELECT 1 FROM sqlite_master WHERE type='table' AND name=?",
                       (name,)).fetchone() is not None


def _spalte_existiert(con, table, col) -> bool:
    try:
        return col in {r[1] for r in con.execute(f"PRAGMA table_info({table})")}
    except sqlite3.Error:
        return False


def _parse_datum(s, default_iso: str | None = None) -> str:
    """'2026-06-30' / '30.06.2026' / '' -> ISO-Datum. Leer/ungültig -> default (heute)."""
    s = str(s or "").strip()
    fallback = default_iso or date.today().isoformat()
    if not s:
        return fallback
    for fmt in ("%Y-%m-%d", "%d.%m.%Y", "%d.%m.%y"):
        try:
            return datetime.strptime(s, fmt).date().isoformat()
        except ValueError:
            continue
    return fallback


def _eur(v) -> str:
    if v is None:
        return "—"
    return f"{v:,.2f} €".replace(",", "X").replace(".", ",").replace("X", ".")


def _parse_num(s) -> float:
    """'1.234,56 €' / '1234.56' / '' -> float. Robust gegen DE/EN-Format."""
    s = str(s or "").replace("€", "").replace("%", "").replace(" ", "").strip()
    if not s:
        return 0.0
    if "," in s:
        s = s.replace(".", "").replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return 0.0


def _now() -> str:
    return datetime.now().isoformat(timespec="seconds")


def get_setting(schluessel: str, default: str = "") -> str:
    try:
        with _con() as con:
            row = con.execute("SELECT wert FROM tbl_faktura_einstellungen WHERE schluessel=?",
                              (schluessel,)).fetchone()
        if row and row[0] is not None:
            return row[0]
    except sqlite3.Error:
        pass
    return STAMM_DEFAULTS.get(schluessel, default)


def set_setting(schluessel: str, wert: str) -> None:
    with _con() as con:
        con.execute("INSERT INTO tbl_faktura_einstellungen(schluessel, wert) VALUES(?,?) "
                    "ON CONFLICT(schluessel) DO UPDATE SET wert=excluded.wert",
                    (schluessel, wert))
        con.commit()


def aktueller_mitarbeiter() -> dict:
    """Liefert den aktuell angemeldeten Mitarbeiter (Auto-Zuordnung ueber den
    Windows-Benutzernamen). Faellt auf den Benutzernamen zurueck, wenn noch
    kein Eintrag gepflegt ist."""
    benutzer = getpass.getuser()
    try:
        with _con() as con:
            row = con.execute(
                "SELECT name, email FROM tbl_faktura_mitarbeiter "
                "WHERE benutzer=? AND aktiv=1 ORDER BY id DESC LIMIT 1",
                (benutzer,)).fetchone()
        if row:
            return {"benutzer": benutzer, "name": row[0] or benutzer, "email": row[1] or ""}
    except sqlite3.Error:
        pass
    return {"benutzer": benutzer, "name": benutzer, "email": ""}


def _log(aktion: str, beleg_id=None, details: str = "") -> None:
    try:
        with _con() as con:
            con.execute("INSERT INTO tbl_faktura_log(bearbeiter, aktion, beleg_id, details) "
                        "VALUES(?,?,?,?)", (getpass.getuser(), aktion, beleg_id, details))
            con.commit()
    except sqlite3.Error:
        pass


def _format_nummer(vorlage: str, jahr: int, monat: int, zaehler: int) -> str:
    """Wendet das konfigurierte Format mit Platzhaltern an. Ist kein {NR} enthalten,
    wird der Zähler 5-stellig angehängt (sonst wären Nummern nicht eindeutig)."""
    import re
    s = (vorlage or "").strip() or "{JJJJ}-{NR:5}"
    s = s.replace("{JJJJ}", f"{jahr:04d}").replace("{JJ}", f"{jahr % 100:02d}")
    s = s.replace("{MM}", f"{monat:02d}")
    if "{NR" not in s:
        s += "-{NR:5}"
    s = re.sub(r"\{NR(?::(\d+))?\}",
               lambda m: f"{zaehler:0{int(m.group(1))}d}" if m.group(1) else str(zaehler), s)
    return s


def naechste_nummer(belegart: str) -> str:
    """Zieht eine lueckenlose Nummer fuer die Belegart im aktuellen Jahr.
    Das Format ist in den Einstellungen frei konfigurierbar (nr_format_<belegart>)."""
    heute = date.today()
    jahr = heute.year
    vorlage = get_setting(f"nr_format_{belegart}", STD_NR_FORMAT.get(belegart, "BE-{JJJJ}-{NR:5}"))
    with _con() as con:
        con.execute("INSERT OR IGNORE INTO tbl_faktura_nummernkreis(belegart, jahr, praefix, letzter_zaehler) "
                    "VALUES(?,?,?,0)", (belegart, jahr, belegart[:2].upper()))
        con.execute("UPDATE tbl_faktura_nummernkreis SET letzter_zaehler = letzter_zaehler + 1 "
                    "WHERE belegart=? AND jahr=?", (belegart, jahr))
        zaehler = con.execute("SELECT letzter_zaehler FROM tbl_faktura_nummernkreis "
                              "WHERE belegart=? AND jahr=?", (belegart, jahr)).fetchone()[0]
        con.commit()
    return _format_nummer(vorlage, jahr, heute.month, zaehler)


# ── PDF-Erzeugung ────────────────────────────────────────────────────────────
def _find_browser():
    for p in (r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
              r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
              r"C:\Program Files\Google\Chrome\Application\chrome.exe",
              r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"):
        if os.path.exists(p):
            return p
    return None


def _logo_data_uri() -> str:
    pfad = get_setting("firma_logo_pfad", "")
    if pfad and os.path.exists(pfad):
        try:
            ext = Path(pfad).suffix.lower().lstrip(".") or "png"
            mime = "jpeg" if ext in ("jpg", "jpeg") else ext
            raw = base64.b64encode(Path(pfad).read_bytes()).decode("ascii")
            return f"data:image/{mime};base64,{raw}"
        except Exception:
            return ""
    return ""


def _hex_tint(hexcol: str, frac: float = 0.12) -> str:
    """Hellt eine Hex-Farbe in Richtung Weiß auf (für dezente Hintergrundflächen)."""
    try:
        h = (hexcol or "").lstrip("#")
        r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
        r = int(255 + (r - 255) * frac); g = int(255 + (g - 255) * frac); b = int(255 + (b - 255) * frac)
        return f"#{r:02x}{g:02x}{b:02x}"
    except Exception:
        return "#eef3f8"


# Freies Layout: Standard-Positionen je Block in Millimeter (A4 = 210 × 297 mm).
LAYOUT_DEFAULT = {
    "logo": {"x": 15, "y": 15}, "firma": {"x": 15, "y": 34}, "beleginfo": {"x": 135, "y": 15},
    "kunde": {"x": 15, "y": 52}, "kopftext": {"x": 15, "y": 80}, "positionen": {"x": 15, "y": 92},
    "summe": {"x": 120, "y": 178}, "sachbearbeiter": {"x": 15, "y": 235},
    "fuss": {"x": 15, "y": 255}, "bank": {"x": 15, "y": 266},
}
BLOCK_BREITE_MM = {"logo": 60, "firma": 90, "beleginfo": 65, "kunde": 90, "kopftext": 180,
                   "positionen": 180, "summe": 80, "sachbearbeiter": 120, "fuss": 180, "bank": 180}
BLOCK_HOEHE_MM = {"logo": 18, "firma": 16, "beleginfo": 24, "kunde": 24, "kopftext": 10,
                  "positionen": 55, "summe": 26, "sachbearbeiter": 10, "fuss": 12, "bank": 8}
BLOCK_LABEL = {"logo": "Logo / Absenderkopf", "firma": "Absender-Anschrift",
               "beleginfo": "Belegtitel + Nummer/Datum", "kunde": "Empfänger",
               "kopftext": "Kopftext", "positionen": "Positionstabelle", "summe": "Summen",
               "sachbearbeiter": "Sachbearbeiter", "fuss": "Fußtext", "bank": "Bankverbindung"}
# Beispiel-Inhalte für den Layout-Editor (damit man sieht, was im Block steht).
BLOCK_SAMPLE = {
    "logo": "[ Logo / Firmenname ]",
    "firma": "Musterstr. 1 · 12345 Musterstadt\nUSt-IdNr. DE123456789",
    "beleginfo": "RECHNUNG\nNr. RE-2026-00001\nDatum: 24.06.2026",
    "kunde": "Beleg an\nMuster-Apotheke\nMarktplatz 7 · 54321 Beispielheim",
    "kopftext": "Kopftext (frei) …",
    "positionen": "PZN   Artikel        Menge  APU      Netto\n04812… Adalimumab 40mg  10  412,50  4.125,00\n05123… Etanercept 50mg   6  298,00  1.788,00",
    "summe": "Netto      5.913,00\nUSt 19 %   1.123,47\nGesamt     7.036,47 €",
    "sachbearbeiter": "Sachbearbeiter: Max Muster · max@nmg.de",
    "fuss": "Zahlbar innerhalb 14 Tagen ohne Abzug. (GoBD)",
    "bank": "Musterbank · IBAN DE.. · BIC ..",
}


# Wählbare Datenbank-/Beleg-Platzhalter für eigene Felder im Layout-Editor.
FELD_QUELLEN = [
    ("Kunde – Name", "{kunde_name}"), ("Kunde – Nummer", "{kunde_nr}"),
    ("Kunde – Anschrift", "{kunde_adresse}"), ("Kunde – USt-IdNr.", "{kunde_ustid}"),
    ("Beleg – Titel", "{titel}"), ("Beleg – Nummer", "{beleg_nr}"),
    ("Beleg – Datum", "{beleg_datum}"), ("Beleg – Leistungsdatum", "{leistungsdatum}"),
    ("Beleg – Netto", "{netto}"), ("Beleg – USt", "{ust}"), ("Beleg – Brutto", "{brutto}"),
    ("Firma – Name", "{firma_name}"), ("Firma – Anschrift", "{firma_anschrift}"),
    ("Firma – USt-IdNr.", "{firma_ustid}"), ("Firma – Steuernummer", "{firma_steuernr}"),
    ("Firma – IBAN", "{firma_iban}"), ("Firma – Telefon", "{firma_telefon}"),
    ("Firma – E-Mail", "{firma_email}"),
    ("Sachbearbeiter", "{sachbearbeiter}"), ("Sachbearbeiter – E-Mail", "{sachbearbeiter_email}"),
]


def _feld_tokens(beleg: dict) -> dict:
    """Werte für die Platzhalter eigener Felder (zum Druckzeitpunkt aufgelöst)."""
    netto = beleg.get("netto")
    ust = beleg.get("ust_betrag")
    return {
        "{kunde_name}": beleg.get("kunde_name") or "", "{kunde_nr}": beleg.get("kunde_nr") or "",
        "{kunde_adresse}": beleg.get("kunde_adresse") or "", "{kunde_ustid}": beleg.get("kunde_ustid") or "",
        "{titel}": get_setting(f"tpl_titel_{beleg.get('belegart')}", "") or "Beleg",
        "{beleg_nr}": beleg.get("beleg_nr") or "", "{beleg_datum}": beleg.get("beleg_datum") or "",
        "{leistungsdatum}": beleg.get("leistungsdatum") or beleg.get("beleg_datum") or "",
        "{netto}": _eur(netto) if netto is not None else "", "{ust}": _eur(ust) if ust is not None else "",
        "{brutto}": _eur(beleg.get("brutto")),
        "{firma_name}": get_setting("firma_name"),
        "{firma_anschrift}": " · ".join(x for x in (get_setting("firma_strasse"),
            f"{get_setting('firma_plz')} {get_setting('firma_ort')}".strip()) if x),
        "{firma_ustid}": get_setting("firma_ustid"), "{firma_steuernr}": get_setting("firma_steuernr"),
        "{firma_iban}": get_setting("firma_iban"), "{firma_telefon}": get_setting("firma_telefon"),
        "{firma_email}": get_setting("firma_email"),
        "{sachbearbeiter}": beleg.get("mitarbeiter") or "",
        "{sachbearbeiter_email}": beleg.get("mitarbeiter_email") or "",
    }


def _feld_aufloesen(text: str, beleg: dict) -> str:
    out = text or ""
    for token, wert in _feld_tokens(beleg).items():
        out = out.replace(token, str(wert))
    return out


def _zusatz_laden() -> list:
    roh = get_setting("tpl_zusatzfelder", "")
    if roh:
        try:
            d = json.loads(roh)
            if isinstance(d, list):
                return d
        except Exception:
            pass
    return []


def _layout_laden() -> dict:
    roh = get_setting("tpl_layout", "")
    data = {}
    if roh:
        try:
            data = json.loads(roh)
        except Exception:
            data = {}
    if not isinstance(data, dict):
        data = {}
    out = {}
    for k in LAYOUT_DEFAULT:
        d = data.get(k, {}) if isinstance(data.get(k), dict) else {}
        out[k] = {"x": float(d.get("x", LAYOUT_DEFAULT[k]["x"])),
                  "y": float(d.get("y", LAYOUT_DEFAULT[k]["y"])),
                  "w": float(d.get("w", BLOCK_BREITE_MM.get(k, 90))),
                  "h": float(d.get("h", BLOCK_HOEHE_MM.get(k, 12)))}
    return out


def _beleg_bloecke(beleg: dict, positionen: list[dict], accent: str) -> dict:
    """Erzeugt das HTML je Inhaltsblock (für Fließ- und Freies Layout wiederverwendbar)."""
    art = beleg["belegart"]
    spalte_apu = get_setting("tpl_spalte_apu", "1") == "1"
    spalte_ust = get_setting("tpl_spalte_ust", "1") == "1"
    firma_adr = " · ".join(x for x in (get_setting("firma_strasse"),
                f"{get_setting('firma_plz')} {get_setting('firma_ort')}".strip()) if x)
    steuerzeile = " · ".join(x for x in (
        f"USt-IdNr. {get_setting('firma_ustid')}" if get_setting("firma_ustid") else "",
        f"St.-Nr. {get_setting('firma_steuernr')}" if get_setting("firma_steuernr") else "") if x)
    bankzeile = " · ".join(x for x in (get_setting("firma_bank"),
                f"IBAN {get_setting('firma_iban')}" if get_setting("firma_iban") else "",
                f"BIC {get_setting('firma_bic')}" if get_setting("firma_bic") else "") if x)
    logo = _logo_data_uri()
    logo_html = f'<img src="{logo}" style="max-height:70px; max-width:220px;">' if logo else \
                f'<div style="font-size:20px; font-weight:700;">{get_setting("firma_name") or "Firma"}</div>'
    titel = get_setting(f"tpl_titel_{art}") or {"rechnung": "Rechnung", "gutschrift": "Gutschrift",
            "storno": "Storno-Rechnung", "quartalsverguetung": "Quartalsvergütung"}.get(art, "Beleg")
    kopftext = get_setting(f"tpl_kopftext_{art}", "")
    fuss = get_setting("tpl_fusstext") or get_setting("rechnung_fusstext")

    nach_satz: dict[float, list] = {}
    for p in positionen:
        nach_satz.setdefault(p["ust_satz"], [0.0, 0.0])
        nach_satz[p["ust_satz"]][0] += p["netto_zeile"]
        nach_satz[p["ust_satz"]][1] += p["ust_zeile"]
    ths = '<th>PZN</th><th>Artikel</th><th class="r">Menge</th>'
    ths += '<th class="r">APU</th>' if spalte_apu else ''
    ths += '<th class="r">USt</th>' if spalte_ust else ''
    ths += '<th class="r">Netto</th>'
    zeilen = ""
    for p in positionen:
        tds = (f'<td class="mono">{p["pzn"] or ""}</td><td>{p["bezeichnung"] or ""}</td>'
               f'<td class="r">{p["menge"]:g}</td>')
        tds += f'<td class="r">{_eur(p["apu_einzel"])}</td>' if spalte_apu else ''
        tds += f'<td class="r">{p["ust_satz"]:g} %</td>' if spalte_ust else ''
        tds += f'<td class="r">{_eur(p["netto_zeile"])}</td>'
        zeilen += f"<tr>{tds}</tr>"
    satz_zeilen = ""
    for satz, (netto, ust) in sorted(nach_satz.items()):
        satz_zeilen += (f'<tr><td class="sub">Netto {satz:g} %</td><td class="r">{_eur(netto)}</td></tr>'
                        f'<tr><td class="sub">zzgl. {satz:g} % USt</td><td class="r">{_eur(ust)}</td></tr>')
    ma = beleg.get("mitarbeiter") or ""
    ma_zeile = (f"Sachbearbeiter: {ma}" + (f" · {beleg.get('mitarbeiter_email')}"
                if beleg.get("mitarbeiter_email") else "")) if ma else ""
    ustid = (' · USt-IdNr. ' + beleg['kunde_ustid']) if beleg.get("kunde_ustid") else ''

    return {
        "logo": logo_html,
        "firma": f'<div class="sub">{firma_adr}</div><div class="sub">{steuerzeile}</div>',
        "beleginfo": (f'<div class="titel">{titel}</div>'
                      f'<div class="sub">Nr. <span class="mono">{beleg.get("beleg_nr") or "(Entwurf)"}</span></div>'
                      f'<div class="sub">Datum: {beleg.get("beleg_datum") or ""}</div>'
                      f'<div class="sub">Leistung: {beleg.get("leistungsdatum") or beleg.get("beleg_datum") or ""}</div>'),
        "kunde": (f'<span class="sub">Beleg an</span><br><b>{beleg.get("kunde_name") or ""}</b><br>'
                  f'<span class="sub">{beleg.get("kunde_adresse") or ""}</span><br>'
                  f'<span class="sub">Kunden-Nr. {beleg.get("kunde_nr") or ""}{ustid}</span>'),
        "kopftext": f'<div class="kopftext">{kopftext}</div>' if kopftext else '',
        "positionen": f'<table><tr>{ths}</tr>{zeilen}</table>',
        "summe": (f'<table class="tot">{satz_zeilen}<tr>'
                  f'<td style="font-weight:700;border-top:1px solid #999;">Gesamt</td>'
                  f'<td class="r" style="font-weight:700;border-top:1px solid #999;">{_eur(beleg.get("brutto"))}</td></tr></table>'),
        "sachbearbeiter": f'<div class="note">{ma_zeile}</div>' if ma_zeile else '',
        "fuss": f'<div class="sub">{fuss}</div>',
        "bank": f'<div class="sub">{bankzeile}</div>',
    }


def _html_rahmen(accent: str, inhalt: str) -> str:
    accent_light = _hex_tint(accent, 0.12)
    return f"""<!DOCTYPE html><html lang="de"><head><meta charset="utf-8"><style>
*{{box-sizing:border-box;-webkit-print-color-adjust:exact;print-color-adjust:exact;}}
body{{font-family:'Segoe UI',Arial,sans-serif;color:#2c2c2a;font-size:13px;margin:0;}}
.row{{display:flex;justify-content:space-between;align-items:flex-start;}}
.titel{{font-size:22px;font-weight:700;color:{accent};}}
.sub{{color:#5f5e5a;font-size:12px;}}
.mono{{font-family:Consolas,monospace;}}
table{{width:100%;border-collapse:collapse;font-size:12px;}}
th,td{{text-align:left;padding:6px 8px;border-bottom:1px solid #e3e2dc;}}
th{{color:{accent};}} td.r,th.r{{text-align:right;}}
.kopftext{{font-size:13px;}}
.tot td{{border:none;padding:3px 8px;}}
.note{{background:{accent_light};color:{accent};border-radius:6px;padding:9px 12px;font-size:12px;}}
.seite{{position:relative;width:210mm;height:297mm;}}
.blk{{position:absolute;}}
@media print{{@page{{size:A4;margin:0;}}}}
</style></head><body>{inhalt}</body></html>"""


def beleg_als_html(beleg: dict, positionen: list[dict]) -> str:
    """Baut die Beleg-HTML. Bei aktivem freien Layout (tpl_layout_aktiv) werden die
    Blöcke an den im Layout-Editor gespeicherten Positionen absolut platziert, sonst
    klassisches Fließlayout."""
    accent = (get_setting("tpl_akzentfarbe") or "#185fa5").strip() or "#185fa5"
    bloecke = _beleg_bloecke(beleg, positionen, accent)

    if get_setting("tpl_layout_aktiv", "0") == "1":
        layout = _layout_laden()
        teile = []
        for key, pos in layout.items():
            inhalt = bloecke.get(key, "")
            if not inhalt:
                continue
            breite = pos.get("w") or BLOCK_BREITE_MM.get(key, 90)
            teile.append(f'<div class="blk" style="left:{pos["x"]:.1f}mm;top:{pos["y"]:.1f}mm;'
                         f'width:{breite:.1f}mm;">{inhalt}</div>')
        # eigene Felder (Freitext / Datenfelder)
        for z in _zusatz_laden():
            txt = _feld_aufloesen(z.get("text", ""), beleg)
            if not txt:
                continue
            teile.append(f'<div class="blk" style="left:{z.get("x", 20):.1f}mm;top:{z.get("y", 20):.1f}mm;'
                         f'width:{z.get("w", 60):.1f}mm;font-size:{z.get("size", 11)}px;'
                         f'white-space:pre-wrap;">{txt}</div>')
        return _html_rahmen(accent, f'<div class="seite">{"".join(teile)}</div>')

    # Fließlayout (Standard)
    kopf_l = f'<div style="width:50%;">{bloecke["logo"]}<div style="margin-top:6px;">{bloecke["firma"]}</div></div>'
    kopf_r = f'<div style="text-align:right;width:50%;">{bloecke["beleginfo"]}</div>'
    kopf = (kopf_r + kopf_l) if get_setting("tpl_logo_pos", "links") == "rechts" else (kopf_l + kopf_r)
    body = (f'<div style="padding:18mm;">'
            f'<div class="row">{kopf}</div>'
            f'<div style="margin-top:18px;">{bloecke["kunde"]}</div>'
            f'<div style="margin-top:14px;">{bloecke["kopftext"]}</div>'
            f'<div style="margin-top:14px;">{bloecke["positionen"]}</div>'
            f'<div style="width:46%;margin-left:auto;margin-top:10px;">{bloecke["summe"]}</div>'
            f'<div style="margin-top:14px;">{bloecke["sachbearbeiter"]}</div>'
            f'<div style="margin-top:16px;">{bloecke["fuss"]}</div>'
            f'<div style="margin-top:6px;">{bloecke["bank"]}</div></div>')
    return _html_rahmen(accent, body)


def _sanitize_dateiname(s: str) -> str:
    """Entfernt unter Windows unzulaessige Zeichen. ';' bleibt als Trennzeichen."""
    for ch in '\\/:*?"<>|\n\r\t':
        s = (s or "").replace(ch, "-")
    return s.strip().strip(".")[:180]


def _quartal_von_datum(iso: str) -> int:
    try:
        return (int(iso[5:7]) - 1) // 3 + 1
    except Exception:
        return (date.today().month - 1) // 3 + 1


def beleg_ablage(beleg: dict) -> tuple[Path, str]:
    """Liefert (Verzeichnis, Dateiname-ohne-Endung) nach Ordnerschema:
        Rechnungen/<Jahr>/<Monat>, Gutschriften/<Jahr>/<Monat>,
        Quartalsverguetung/<Jahr>/Q<n>.
    Dateiname: Kundennummer;Rechnungsnummer;Apotheke;Datum.
    Basis ist OUTPUT_DIR (künftig per SharePoint/OneDrive umlenkbar über config.py)."""
    art = beleg.get("belegart") or "rechnung"
    datum = beleg.get("beleg_datum") or date.today().isoformat()
    jahr, monat = datum[:4], datum[5:7]
    base = Path(OUTPUT_DIR)
    if art == "quartalsverguetung":
        bezug = beleg.get("zeitraum_von") or datum
        verzeichnis = base / "Quartalsverguetung" / bezug[:4] / f"Q{_quartal_von_datum(bezug)}"
    elif art == "gutschrift":
        verzeichnis = base / "Gutschriften" / jahr / monat
    else:  # rechnung + storno
        verzeichnis = base / "Rechnungen" / jahr / monat
    verzeichnis.mkdir(parents=True, exist_ok=True)
    name = _sanitize_dateiname(";".join((
        str(beleg.get("kunde_nr") or "ohneNr"),
        str(beleg.get("beleg_nr") or "Entwurf"),
        str(beleg.get("kunde_name") or "Kunde"),
        str(datum))))
    return verzeichnis, name


def erzeuge_pdf(beleg: dict, positionen: list[dict]) -> str | None:
    """Schreibt HTML + konvertiert via Edge/Chrome headless zu PDF.
    Rueckgabe: PDF-Pfad oder None. Ohne Browser bleibt das HTML als Fallback."""
    ziel, name = beleg_ablage(beleg)
    html_pfad = ziel / f"{name}.html"
    pdf_pfad = ziel / f"{name}.pdf"
    html_pfad.write_text(beleg_als_html(beleg, positionen), encoding="utf-8")

    browser = _find_browser()
    if not browser:
        return str(html_pfad)
    try:
        uri = html_pfad.resolve().as_uri()
        if pdf_pfad.exists():
            pdf_pfad.unlink()
        subprocess.run([browser, "--headless", "--disable-gpu", "--no-pdf-header-footer",
                        f"--print-to-pdf={pdf_pfad}", uri],
                       timeout=40, capture_output=True)
        if pdf_pfad.exists():
            return str(pdf_pfad)
    except Exception:
        pass
    return str(html_pfad)


def vorschau_pdf(belegart: str = "rechnung") -> str | None:
    """Erzeugt eine Beispiel-PDF mit der aktuell gespeicherten Vorlage (kein echter
    Beleg, kein Nummernkreis-Verbrauch). Rückgabe: Pfad oder None."""
    ma = aktueller_mitarbeiter()
    heute = date.today()
    musternr = _format_nummer(get_setting(f"nr_format_{belegart}", STD_NR_FORMAT.get(belegart, "")),
                              heute.year, heute.month, 123)
    beleg = {"belegart": belegart, "beleg_nr": musternr, "kunde_name": "Muster-Apotheke",
             "kunde_adresse": "Marktplatz 7 · 54321 Beispielheim", "kunde_nr": "10042",
             "kunde_ustid": "DE987654321", "beleg_datum": heute.isoformat(),
             "leistungsdatum": heute.isoformat(), "brutto": 5913.0 * 1.19,
             "mitarbeiter": ma["name"], "mitarbeiter_email": ma["email"]}
    pos = [{"pzn": "04812345", "bezeichnung": "Adalimumab 40 mg", "menge": 10, "apu_einzel": 412.5,
            "ust_satz": 19, "netto_zeile": 4125.0, "ust_zeile": 783.75},
           {"pzn": "05123987", "bezeichnung": "Etanercept 50 mg", "menge": 6, "apu_einzel": 298.0,
            "ust_satz": 19, "netto_zeile": 1788.0, "ust_zeile": 339.72}]
    ziel = Path(OUTPUT_DIR) / "faktura"
    ziel.mkdir(parents=True, exist_ok=True)
    html_pfad = ziel / "_vorschau.html"
    pdf_pfad = ziel / "_vorschau.pdf"
    html_pfad.write_text(beleg_als_html(beleg, pos), encoding="utf-8")
    browser = _find_browser()
    if not browser:
        return str(html_pfad)
    try:
        if pdf_pfad.exists():
            pdf_pfad.unlink()
        subprocess.run([browser, "--headless", "--disable-gpu", "--no-pdf-header-footer",
                        f"--print-to-pdf={pdf_pfad}", html_pfad.resolve().as_uri()],
                       timeout=40, capture_output=True)
        if pdf_pfad.exists():
            return str(pdf_pfad)
    except Exception:
        pass
    return str(html_pfad)


# ── UI ───────────────────────────────────────────────────────────────────────
def _card(parent, title=None):
    outer = tk.Frame(parent, bg=BORDER)
    inner = tk.Frame(outer, bg=BG)
    inner.pack(fill="both", expand=True, padx=1, pady=1)
    if title:
        head = tk.Frame(inner, bg=BG)
        head.pack(fill="x", padx=14, pady=(10, 0))
        tk.Label(head, text=title, bg=BG, fg=ACCENT, font=(theme.FONT, 11, "bold")).pack(side="left")
        tk.Frame(inner, bg=ACCENT_LIGHT, height=2).pack(fill="x", padx=14, pady=(7, 0))
    body = tk.Frame(inner, bg=BG)
    body.pack(fill="both", expand=True, padx=14, pady=(10, 12))
    return outer, body


class FakturaPanel(tk.Frame):
    """Hauptpanel der Faktura-App: dunkle Sidebar + wechselnde Seiten."""

    def __init__(self, master, on_close=None):
        super().__init__(master, bg=SHELL_BG)
        self.on_close = on_close
        self.positionen: list[dict] = []   # Positionen der gerade bearbeiteten Rechnung

        self.sidebar = theme.Sidebar(self, title="Faktura", subtitle="Rechnungen & Gutschriften")
        self.sidebar.pack(side="left", fill="y")
        self.sidebar.add_section("Übersicht")
        self.sidebar.add_item("start", "🏠", "Startseite", lambda: self.show("start"), active=True)
        self.sidebar.add_section("Aufträge")
        self.sidebar.add_item("kunden", "📋", "Aufträge", lambda: self.show("kunden"))
        self.sidebar.add_section("Rechnungen")
        self.sidebar.add_item("rechnungen", "🧾", "Rechnungen", lambda: self.show("rechnungen"))
        self.sidebar.add_item("neu", "➕", "Neue Rechnung", lambda: self.show("neu"))
        self.sidebar.add_section("Vergütung")
        self.sidebar.add_item("verguetung", "💶", "Quartalsvergütung", lambda: self.show("verguetung"))
        self.sidebar.add_section("Konfiguration")
        self.sidebar.add_item("staffel", "📈", "Staffel", lambda: self.show("staffel"))
        self.sidebar.add_item("einstellungen", "⚙", "Einstellungen", lambda: self.show("einst_firma"))
        # Aufklappbare Unterpunkte der Einstellungen
        self._einst_subs = [
            self.sidebar.add_subitem("einst_firma", "Firmendaten", lambda: self.show("einst_firma")),
            self.sidebar.add_subitem("einst_nummern", "Belegnummern", lambda: self.show("einst_nummern")),
            self.sidebar.add_subitem("einst_layout", "Layouts", lambda: self.show("einst_layout")),
        ]
        self.sidebar.add_footer_note("NMG Faktura · Prototyp")

        self.content = tk.Frame(self, bg=SHELL_BG)
        self.content.pack(side="left", fill="both", expand=True)
        self.show("start")

    def show(self, key):
        # Einstellungs-Unterbaum auf-/zuklappen
        self._einst_aufklappen(key.startswith("einst_"))
        for w in self.content.winfo_children():
            w.destroy()
        self.sidebar.set_active(key)
        {"start": self._page_start, "kunden": self._page_kunden,
         "rechnungen": self._page_rechnungen, "neu": self._page_neu,
         "verguetung": self._page_verguetung, "staffel": self._page_staffel,
         "einst_firma": self._page_einst_firma, "einst_nummern": self._page_einst_nummern,
         "einst_layout": self._page_einst_layout}[key]()

    def _einst_aufklappen(self, auf: bool):
        for w in self._einst_subs:
            if auf:
                w.pack(fill="x", padx=10, pady=0)
            else:
                w.pack_forget()

    def _scrollseite(self):
        """Liefert ein scrollbares Wrap-Frame (Mausrad aktiv) für lange Seiten."""
        canvas = tk.Canvas(self.content, bg=SHELL_BG, highlightthickness=0)
        scroll = ttk.Scrollbar(self.content, orient="vertical", command=canvas.yview)
        wrap = tk.Frame(canvas, bg=SHELL_BG)
        wrap.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=wrap, anchor="nw", width=900)
        canvas.configure(yscrollcommand=scroll.set)
        canvas.pack(side="left", fill="both", expand=True, padx=(24, 0))
        scroll.pack(side="right", fill="y")
        def _wheel(e):
            canvas.yview_scroll(int(-1 * (e.delta / 120)), "units")
        canvas.bind("<Enter>", lambda e: canvas.bind_all("<MouseWheel>", _wheel))
        canvas.bind("<Leave>", lambda e: canvas.unbind_all("<MouseWheel>"))
        return wrap

    def _metric_cards(self, parent, items):
        """Reihe von Kennzahl-Karten im Tagesabschluss-Stil. items = Liste aus
        (key, label) oder (key, label, fg). Liefert (frame, {key: wert-Label})."""
        row = tk.Frame(parent, bg=SHELL_BG)
        row.pack(fill="x", pady=(12, 4))
        out = {}
        for item in items:
            key, label = item[0], item[1]
            fg = item[2] if len(item) > 2 else ACCENT
            c = tk.Frame(row, bg="#f2f6fb", highlightbackground="#dde7f1", highlightthickness=1)
            c.pack(side="left", fill="both", expand=True, padx=(0, 8))
            tk.Label(c, text=label, bg="#f2f6fb", fg="#666",
                     font=(theme.FONT, 9)).pack(anchor="w", padx=12, pady=(8, 0))
            val = tk.Label(c, text="–", bg="#f2f6fb", fg=fg, font=(theme.FONT, 17, "bold"))
            val.pack(anchor="w", padx=12, pady=(0, 9))
            out[key] = val
        return row, out

    # ── Seite: Startseite / Dashboard ────────────────────────────────────────
    def _page_start(self):
        ma = aktueller_mitarbeiter()
        theme.page_header(self.content, "Faktura",
                          f"Angemeldet: {ma['name']}" + (f" · {ma['email']}" if ma['email'] else ""),
                          bg=SHELL_BG).pack(fill="x", padx=24, pady=(20, 12))
        kpi = self._kennzahlen()
        grid = tk.Frame(self.content, bg=SHELL_BG)
        grid.pack(fill="x", padx=24)
        kacheln = [
            ("Rechnungen gesamt", str(kpi["anzahl_rechnungen"]), ACCENT),
            (f"Umsatz {kpi['monat_label']} (netto)", _eur(kpi["monat_netto"]), OK_GREEN),
            ("Offene Entwürfe", str(kpi["entwuerfe"]), theme.WARNING),
            ("Gutschriften", str(kpi["gutschriften"]), theme.PURPLE),
        ]
        for i, (label, wert, farbe) in enumerate(kacheln):
            grid.columnconfigure(i, weight=1)
            card = tk.Frame(grid, bg=BG, highlightbackground=BORDER, highlightthickness=1)
            card.grid(row=0, column=i, sticky="ew", padx=(0 if i == 0 else 8, 0))
            tk.Label(card, text=label, bg=BG, fg=MUTED, font=theme.SMALL).pack(anchor="w", padx=14, pady=(12, 0))
            tk.Label(card, text=wert, bg=BG, fg=farbe, font=(theme.FONT, 22, "bold")).pack(anchor="w", padx=14, pady=(0, 12))

        bar = tk.Frame(self.content, bg=SHELL_BG)
        bar.pack(fill="x", padx=24, pady=14)
        theme.PillButton(bar, "➕  Neue Rechnung", lambda: self.show("neu"),
                         kind="primary", font_size=11, padx=16, pady=9).pack(side="left")
        theme.PillButton(bar, "💶  Quartalsvergütung berechnen", lambda: self.show("verguetung"),
                         kind="success", font_size=11, padx=16, pady=9).pack(side="left", padx=(8, 0))

        outer, body = _card(self.content, "Zuletzt erstellt")
        outer.pack(fill="both", expand=True, padx=24, pady=(0, 20))
        cols = ("nr", "art", "datum", "kunde", "brutto", "status")
        tree = ttk.Treeview(body, columns=cols, show="headings", height=10)
        for c, t, w in (("nr", "Beleg-Nr.", 120), ("art", "Art", 90), ("datum", "Datum", 90),
                        ("kunde", "Kunde", 220), ("brutto", "Brutto", 110), ("status", "Status", 120)):
            tree.heading(c, text=t)
            tree.column(c, width=w, anchor="e" if c == "brutto" else "w")
        tree.pack(fill="both", expand=True)
        theme.style_treeview(tree)
        with _con() as con:
            rows = con.execute(
                "SELECT beleg_nr, belegart, beleg_datum, kunde_name, brutto, status "
                "FROM tbl_faktura_belege ORDER BY id DESC LIMIT 12").fetchall()
        for r in rows:
            art = {"rechnung": "Rechnung", "gutschrift": "Gutschrift", "storno": "Storno",
                   "quartalsverguetung": "Quartalsverg."}.get(r[1], r[1])
            tree.insert("", "end", values=(r[0] or "(Entwurf)", art, r[2] or "", r[3] or "",
                                           _eur(r[4]), (r[5] or "").capitalize()))

    def _kennzahlen(self) -> dict:
        monat = date.today().strftime("%Y-%m")
        with _con() as con:
            anzahl = con.execute("SELECT COUNT(*) FROM tbl_faktura_belege WHERE belegart='rechnung'").fetchone()[0]
            entwuerfe = con.execute("SELECT COUNT(*) FROM tbl_faktura_belege "
                                    "WHERE status='entwurf'").fetchone()[0]
            gut = con.execute("SELECT COUNT(*) FROM tbl_faktura_belege WHERE belegart='gutschrift'").fetchone()[0]
            monat_netto = con.execute(
                "SELECT COALESCE(SUM(netto),0) FROM tbl_faktura_belege "
                "WHERE belegart='rechnung' AND status='festgeschrieben' AND substr(beleg_datum,1,7)=?",
                (monat,)).fetchone()[0]
        return {"anzahl_rechnungen": anzahl, "entwuerfe": entwuerfe, "gutschriften": gut,
                "monat_netto": monat_netto, "monat_label": date.today().strftime("%m/%Y")}

    # ── Seite: Kunden-Umsätze ────────────────────────────────────────────────
    def _page_kunden(self):
        theme.page_header(self.content, "Aufträge",
                          "Offene Aufträge je Apotheke (Kasse-Verkäufe, noch nicht abgerechnet)",
                          bg=SHELL_BG).pack(fill="x", padx=24, pady=(20, 8))
        bar = tk.Frame(self.content, bg=SHELL_BG)
        bar.pack(fill="x", padx=24, pady=(0, 8))
        tk.Label(bar, text="Zeitraum (JJJJ oder JJJJ-MM, leer = alle):", bg=SHELL_BG, fg=MUTED,
                 font=theme.SMALL).pack(side="left")
        self.kunden_zeitraum = tk.StringVar()
        ent = ttk.Entry(bar, textvariable=self.kunden_zeitraum, width=10)
        ent.pack(side="left", padx=6)
        ent.bind("<Return>", lambda e: self._lade_kunden())
        theme.PillButton(bar, "Aktualisieren", self._lade_kunden, kind="accent",
                         font_size=10, padx=12, pady=6).pack(side="left", padx=(6, 0))
        self.kunden_info = tk.StringVar()
        tk.Label(bar, textvariable=self.kunden_info, bg=SHELL_BG, fg=theme.FAINT,
                 font=theme.SMALL).pack(side="left", padx=(14, 0))

        bar2 = tk.Frame(self.content, bg=SHELL_BG)
        bar2.pack(fill="x", padx=24, pady=(0, 8))
        theme.PillButton(bar2, "🧾  Rechnung erstellen (Auswahl)", self._sammelrechnung,
                         kind="primary", font_size=10, padx=14, pady=7).pack(side="left")
        theme.PillButton(bar2, "✖  Für jetzt entfernen", self._kunde_ausblenden,
                         kind="ghost", font_size=10, padx=14, pady=7).pack(side="left", padx=(8, 0))
        tk.Label(bar2, text="Rechnungsdatum:", bg=SHELL_BG, fg=MUTED,
                 font=theme.SMALL).pack(side="left", padx=(16, 2))
        self.kunden_rechnungsdatum = tk.StringVar(value=date.today().isoformat())
        ttk.Entry(bar2, textvariable=self.kunden_rechnungsdatum, width=12).pack(side="left")
        tk.Label(bar2, text="(leer = heute · „Aktualisieren\" holt ausgeblendete zurück)",
                 bg=SHELL_BG, fg=theme.FAINT, font=theme.SMALL).pack(side="left", padx=(10, 0))

        outer, body = _card(self.content, "Übersicht je Kunde")
        outer.pack(fill="both", expand=True, padx=24, pady=(0, 20))
        cols = ("nr", "name", "auftraege", "artikel", "umsatz", "umsatz_rab")
        tree = ttk.Treeview(body, columns=cols, show="headings", height=15)
        spalten = (("nr", "Kundennr.", 110), ("name", "Apotheke", 220),
                   ("auftraege", "Aufträge", 90), ("artikel", "Artikel", 90),
                   ("umsatz", "Umsatz", 130), ("umsatz_rab", "Umsatz n. Rabatt", 140))
        for c, t, w in spalten:
            tree.heading(c, text=t, command=lambda cc=c: self._kunden_sort(cc))
            tree.column(c, width=w, anchor="e" if c in ("auftraege", "artikel", "umsatz", "umsatz_rab") else "w")
        tree.pack(fill="both", expand=True)
        theme.style_treeview(tree)
        self._kunden_tree = tree
        self._kunden_rows = []
        self._kunden_sortspalte = ("umsatz_rab", True)
        self._lade_kunden()

    def _lade_kunden(self):
        self._kunden_hidden = set()  # „Aktualisieren" holt ausgeblendete Kunden zurück
        zeitraum = (self.kunden_zeitraum.get() or "").strip()
        wo = "b.status <> 'storniert' AND COALESCE(p.bestellart,'') <> 'abgesagt'"
        with _con() as con:
            offen_filter = " AND b.faktura_beleg_id IS NULL" if \
                _spalte_existiert(con, "tbl_bestellungen", "faktura_beleg_id") else ""
        wo += offen_filter
        params: tuple = ()
        if zeitraum:
            wo += " AND b.datum LIKE ?"
            params = (zeitraum + "%",)
        sql = f"""
            SELECT b.kundennummer,
                   COALESCE(k.kundenname, b.apotheke, '—') AS name,
                   COUNT(DISTINCT b.id) AS auftraege,
                   COALESCE(SUM(p.menge),0) AS artikel,
                   COALESCE(SUM(p.apu*p.menge),0) AS umsatz,
                   COALESCE(SUM(p.apu*p.menge*(1-COALESCE(p.rabatt_prozent,0)/100.0)),0) AS umsatz_rab
            FROM tbl_bestellungen b
            JOIN tbl_bestellpositionen p ON p.bestell_id = b.id
            LEFT JOIN tbl_kunden_center k ON k.kundennummer = b.kundennummer
            WHERE {wo}
            GROUP BY b.kundennummer, name
        """
        try:
            with _con() as con:
                if not _table_exists(con, "tbl_bestellungen"):
                    self._kunden_rows = []
                else:
                    self._kunden_rows = con.execute(sql, params).fetchall()
        except sqlite3.Error as e:
            messagebox.showerror("Fehler", f"Konnte Umsätze nicht laden:\n{e}")
            self._kunden_rows = []
        self._kunden_render()

    def _kunden_sort(self, spalte):
        akt, desc = self._kunden_sortspalte
        desc = not desc if akt == spalte else True
        self._kunden_sortspalte = (spalte, desc)
        self._kunden_render()

    def _kunden_render(self):
        idx = {"nr": 0, "name": 1, "auftraege": 2, "artikel": 3, "umsatz": 4, "umsatz_rab": 5}
        spalte, desc = self._kunden_sortspalte
        hidden = getattr(self, "_kunden_hidden", set())
        sichtbar = [r for r in self._kunden_rows if r[0] not in hidden]
        rows = sorted(sichtbar, key=lambda r: r[idx[spalte]], reverse=desc)
        tree = self._kunden_tree
        tree.delete(*tree.get_children())
        s_auf = s_art = s_ums = s_rab = 0
        for r in rows:
            s_auf += r[2]; s_art += r[3]; s_ums += r[4]; s_rab += r[5]
            tree.insert("", "end", values=(r[0] or "—", r[1], r[2], f"{r[3]:g}",
                        _eur(r[4]), _eur(r[5])))
        if rows:
            tree.insert("", "end", values=("", "Σ Gesamt", s_auf, f"{s_art:g}",
                        _eur(s_ums), _eur(s_rab)), tags=("summe",))
            tree.tag_configure("summe", background="#EEF3F8", font=(theme.FONT, 10, "bold"))
        # Sortpfeil im Spaltenkopf
        for c in ("nr", "name", "auftraege", "artikel", "umsatz", "umsatz_rab"):
            txt = tree.heading(c, "text").rstrip(" ▲▼")
            tree.heading(c, text=txt + (" ▼" if desc else " ▲") if c == spalte else txt)
        ausgeblendet = len(self._kunden_rows) - len(rows)
        self.kunden_info.set(f"{len(rows)} Kunden mit Verkäufen"
                             + (f" · {ausgeblendet} ausgeblendet" if ausgeblendet else ""))

    def _kunde_auswahl(self):
        sel = self._kunden_tree.selection()
        if not sel:
            return None
        vals = self._kunden_tree.item(sel[0], "values")
        if not vals or not vals[0]:   # Σ-Summenzeile hat keine Kundennr.
            return None
        return {"nr": vals[0], "name": vals[1]}

    def _kunden_auswahl_alle(self):
        """Alle markierten Kunden (Mehrfachauswahl), ohne die Σ-Summenzeile."""
        treffer = []
        for iid in self._kunden_tree.selection():
            vals = self._kunden_tree.item(iid, "values")
            if vals and vals[0]:
                treffer.append({"nr": vals[0], "name": vals[1]})
        return treffer

    def _kunde_ausblenden(self):
        kunden = self._kunden_auswahl_alle()
        if not kunden:
            messagebox.showinfo("Ausblenden", "Bitte einen oder mehrere Kunden auswählen.")
            return
        for k in kunden:
            self._kunden_hidden.add(k["nr"])
        self._kunden_render()

    def _sammelrechnung(self):
        k = self._kunde_auswahl()
        if not k:
            messagebox.showinfo("Rechnung", "Bitte einen Kunden auswählen.")
            return
        zeitraum = (self.kunden_zeitraum.get() or "").strip()
        rechnungsdatum = _parse_datum(self.kunden_rechnungsdatum.get())
        # nur offene (noch nicht abgerechnete) Auftraege heranziehen
        wo = ("b.kundennummer=? AND b.status<>'storniert' AND COALESCE(p.bestellart,'')<>'abgesagt' "
              "AND b.faktura_beleg_id IS NULL")
        params = [k["nr"]]
        if zeitraum:
            wo += " AND b.datum LIKE ?"
            params.append(zeitraum + "%")
        satz = _parse_num(get_setting("ust_satz_standard")) or 19.0
        with _con() as con:
            schon = con.execute(
                "SELECT beleg_nr FROM tbl_faktura_belege WHERE belegart='rechnung' AND kunde_nr=? "
                "AND notizen LIKE ?", (k["nr"], f"Sammelrechnung {zeitraum or 'alle'}%")).fetchone()
            rows = con.execute(
                f"SELECT p.pzn, MAX(p.artikelname), p.apu, COALESCE(p.rabatt_prozent,0) AS rab, "
                f"SUM(p.menge) AS menge FROM tbl_bestellungen b "
                f"JOIN tbl_bestellpositionen p ON p.bestell_id=b.id WHERE {wo} "
                f"GROUP BY p.pzn, p.apu, rab ORDER BY MAX(p.artikelname)", params).fetchall()
            kdaten = con.execute(
                "SELECT COALESCE(strasse,''), COALESCE(plz,''), COALESCE(ort,'') "
                "FROM tbl_kunden_center WHERE kundennummer=? LIMIT 1", (k["nr"],)).fetchone()
        if not rows:
            messagebox.showinfo("Rechnung", "Keine abrechenbaren Positionen im gewählten Zeitraum.")
            return
        if schon and not messagebox.askyesno("Bereits abgerechnet",
                f"Für diesen Zeitraum existiert schon Rechnung {schon[0]}. Trotzdem neue erstellen?"):
            return
        adresse = " · ".join(x for x in ((kdaten[0] if kdaten else ""),
                  f"{kdaten[1]} {kdaten[2]}".strip() if kdaten else "") if x)
        # Bei „Auftragsrabatt"-Gutschrift wird zum vollen APU fakturiert (der Rabatt
        # wandert in die Gutschrift); sonst wird der Rabatt auf der Rechnung abgezogen.
        voll_apu = (get_setting("auto_gutschrift", "0") == "1"
                    and get_setting("auto_gutschrift_typ", "") == "auftragsrabatt")
        positionen = []
        for pzn, name, apu, rab, menge in rows:
            apu = apu or 0.0
            rab_eff = 0 if voll_apu else (rab or 0)
            netto = round(apu * menge * (1 - rab_eff / 100.0), 2)
            ust = round(netto * satz / 100.0, 2)
            positionen.append({"pzn": pzn, "bezeichnung": name, "menge": menge, "apu_einzel": apu,
                               "rabatt": rab or 0, "ust_satz": satz, "netto_zeile": netto,
                               "ust_zeile": ust, "brutto_zeile": round(netto + ust, 2)})
        netto_s = round(sum(p["netto_zeile"] for p in positionen), 2)
        ust_s = round(sum(p["ust_zeile"] for p in positionen), 2)
        ma = aktueller_mitarbeiter()
        nr = naechste_nummer("rechnung")
        with _con() as con:
            cur = con.execute(
                "INSERT INTO tbl_faktura_belege(belegart, beleg_nr, kunde_nr, kunde_name, kunde_adresse, "
                "beleg_datum, leistungsdatum, zeitraum_von, zeitraum_bis, netto, ust_betrag, brutto, "
                "status, mitarbeiter, mitarbeiter_email, festgeschrieben_am, notizen) "
                "VALUES('rechnung',?,?,?,?,?,?,?,?,?,?,?, 'festgeschrieben',?,?,?,?)",
                (nr, k["nr"], k["name"], adresse, rechnungsdatum, rechnungsdatum,
                 f"{zeitraum}" if zeitraum else None, None, netto_s, ust_s, round(netto_s + ust_s, 2),
                 ma["name"], ma["email"], _now(),
                 f"Sammelrechnung {zeitraum or 'alle'} ({len(positionen)} Positionen)"))
            bid = cur.lastrowid
            for i, p in enumerate(positionen, 1):
                con.execute(
                    "INSERT INTO tbl_faktura_positionen(beleg_id, pos_nr, pzn, bezeichnung, menge, "
                    "apu_einzel, rabatt, ust_satz, netto_zeile, ust_zeile, brutto_zeile) "
                    "VALUES(?,?,?,?,?,?,?,?,?,?,?)",
                    (bid, i, p["pzn"], p["bezeichnung"], p["menge"], p["apu_einzel"], p["rabatt"],
                     p["ust_satz"], p["netto_zeile"], p["ust_zeile"], p["brutto_zeile"]))
            # Abgerechnete Auftraege mit der Rechnung verknuepfen (verschwinden aus
            # der Auftragsliste; ein Storno gibt sie wieder frei).
            upd_wo = ("kundennummer=? AND status<>'storniert' AND faktura_beleg_id IS NULL "
                      "AND EXISTS(SELECT 1 FROM tbl_bestellpositionen p WHERE p.bestell_id="
                      "tbl_bestellungen.id AND COALESCE(p.bestellart,'')<>'abgesagt')")
            upd_params = [bid, _now(), k["nr"]]
            if zeitraum:
                upd_wo += " AND datum LIKE ?"
                upd_params.append(zeitraum + "%")
            con.execute(f"UPDATE tbl_bestellungen SET faktura_beleg_id=?, abgerechnet_am=? "
                        f"WHERE {upd_wo}", upd_params)
            con.commit()
        beleg = self._beleg_dict(bid)
        pdf = erzeuge_pdf(beleg, self._positionen_dict(bid))
        with _con() as con:
            con.execute("UPDATE tbl_faktura_belege SET pdf_pfad=? WHERE id=?", (pdf, bid))
            con.commit()
        _log("sammelrechnung", bid, f"{nr} {k['name']} {zeitraum or 'alle'}: {_eur(netto_s + ust_s)}")
        if pdf and os.path.exists(pdf):
            try:
                os.startfile(pdf)  # type: ignore[attr-defined]
            except Exception:
                pass
        gut = self._auto_gutschrift(bid)
        msg = (f"Rechnung {nr} über {_eur(netto_s + ust_s)} erstellt.\n"
               f"Gespeichert unter:\n{pdf}")
        if gut:
            msg += f"\n\nAutomatische Gutschrift {gut['nr']} über {_eur(gut['brutto'])} erstellt."
        messagebox.showinfo("Rechnung erstellt", msg)
        self._kunden_hidden.add(k["nr"])
        self._kunden_render()

    # ── Seite: Rechnungsliste ────────────────────────────────────────────────
    def _page_rechnungen(self):
        theme.page_header(self.content, "Rechnungen", "Alle Belege · Doppelklick öffnet das PDF",
                          bg=SHELL_BG).pack(fill="x", padx=24, pady=(20, 8))
        bar = tk.Frame(self.content, bg=SHELL_BG)
        bar.pack(fill="x", padx=24, pady=(0, 8))
        theme.PillButton(bar, "➕  Neue Rechnung", lambda: self.show("neu"),
                         kind="primary", font_size=10, padx=14, pady=7).pack(side="left")
        theme.PillButton(bar, "↩  Gutschrift zur Auswahl", self._gutschrift_zur_auswahl,
                         kind="success", font_size=10, padx=14, pady=7).pack(side="left", padx=(8, 0))
        theme.PillButton(bar, "⛔  Storno zur Auswahl", self._storno_der_auswahl,
                         kind="danger", font_size=10, padx=14, pady=7).pack(side="left", padx=(8, 0))

        bar2 = tk.Frame(self.content, bg=SHELL_BG)
        bar2.pack(fill="x", padx=24, pady=(0, 8))
        tk.Label(bar2, text="Ordner öffnen:", bg=SHELL_BG, fg=MUTED,
                 font=theme.SMALL).pack(side="left", padx=(0, 4))
        theme.PillButton(bar2, "📂 Rechnungen", lambda: self._ordner_oeffnen("Rechnungen"),
                         kind="neutral", font_size=10, padx=12, pady=6).pack(side="left")
        theme.PillButton(bar2, "📂 Gutschriften", lambda: self._ordner_oeffnen("Gutschriften"),
                         kind="neutral", font_size=10, padx=12, pady=6).pack(side="left", padx=(6, 0))
        theme.PillButton(bar2, "📂 Quartalsvergütung", lambda: self._ordner_oeffnen("Quartalsverguetung"),
                         kind="neutral", font_size=10, padx=12, pady=6).pack(side="left", padx=(6, 0))
        theme.PillButton(bar2, "📁 Ordner der Auswahl", self._ordner_der_auswahl,
                         kind="ghost", font_size=10, padx=12, pady=6).pack(side="left", padx=(12, 0))

        outer, body = _card(self.content, "Belegübersicht")
        outer.pack(fill="both", expand=True, padx=24, pady=(0, 20))
        cols = ("nr", "art", "datum", "kunde", "brutto", "status", "bearbeiter")
        tree = ttk.Treeview(body, columns=cols, show="headings", height=16)
        for c, t, w in (("nr", "Beleg-Nr.", 120), ("art", "Art", 90), ("datum", "Datum", 90),
                        ("kunde", "Kunde", 200), ("brutto", "Brutto", 100),
                        ("status", "Status", 110), ("bearbeiter", "Sachbearbeiter", 140)):
            tree.heading(c, text=t)
            tree.column(c, width=w, anchor="e" if c == "brutto" else "w")
        tree.pack(fill="both", expand=True)
        theme.style_treeview(tree)
        self._rech_tree = tree
        tree.bind("<Double-1>", lambda e: self._oeffne_pdf_der_auswahl())
        self._lade_rechnungen()

    def _lade_rechnungen(self):
        tree = self._rech_tree
        tree.delete(*tree.get_children())
        with _con() as con:
            if not _table_exists(con, "tbl_faktura_belege"):
                return
            rows = con.execute(
                "SELECT id, beleg_nr, belegart, beleg_datum, kunde_name, brutto, status, mitarbeiter "
                "FROM tbl_faktura_belege ORDER BY id DESC").fetchall()
        for r in rows:
            art = {"rechnung": "Rechnung", "gutschrift": "Gutschrift", "storno": "Storno",
                   "quartalsverguetung": "Quartalsverg."}.get(r[2], r[2])
            tree.insert("", "end", iid=str(r[0]),
                        values=(r[1] or "(Entwurf)", art, r[3] or "", r[4] or "",
                                _eur(r[5]), (r[6] or "").capitalize(), r[7] or ""))

    def _auswahl_id(self):
        sel = self._rech_tree.selection()
        return int(sel[0]) if sel else None

    def _oeffne_pdf_der_auswahl(self):
        bid = self._auswahl_id()
        if not bid:
            return
        with _con() as con:
            row = con.execute("SELECT pdf_pfad FROM tbl_faktura_belege WHERE id=?", (bid,)).fetchone()
        if row and row[0] and os.path.exists(row[0]):
            try:
                os.startfile(row[0])  # type: ignore[attr-defined]
            except Exception:
                messagebox.showinfo("PDF", f"Datei liegt unter:\n{row[0]}")
        else:
            messagebox.showinfo("Kein PDF", "Für diesen Beleg wurde noch kein PDF erzeugt "
                                            "(nur festgeschriebene Belege haben ein PDF).")

    def _ordner_oeffnen(self, unter=""):
        """Öffnet den Ablageordner (Rechnungen/Gutschriften/Quartalsverguetung) im Explorer."""
        ziel = Path(OUTPUT_DIR) / unter if unter else Path(OUTPUT_DIR)
        ziel.mkdir(parents=True, exist_ok=True)
        try:
            os.startfile(str(ziel))  # type: ignore[attr-defined]
        except Exception:
            messagebox.showinfo("Ordner", f"Ordner:\n{ziel}")

    def _ordner_der_auswahl(self):
        """Öffnet den Ordner, in dem das PDF des ausgewählten Belegs liegt."""
        bid = self._auswahl_id()
        if not bid:
            messagebox.showinfo("Ordner", "Bitte zuerst einen Beleg auswählen.")
            return
        with _con() as con:
            row = con.execute("SELECT pdf_pfad FROM tbl_faktura_belege WHERE id=?", (bid,)).fetchone()
        if row and row[0] and os.path.exists(row[0]):
            try:
                os.startfile(os.path.dirname(row[0]))  # type: ignore[attr-defined]
            except Exception:
                messagebox.showinfo("Ordner", f"Ordner:\n{os.path.dirname(row[0])}")
        else:
            messagebox.showinfo("Ordner", "Für diesen Beleg gibt es noch kein gespeichertes PDF.")

    def _gutschrift_zur_auswahl(self):
        bid = self._auswahl_id()
        if not bid:
            messagebox.showinfo("Gutschrift", "Bitte zuerst eine Rechnung auswählen.")
            return
        with _con() as con:
            row = con.execute("SELECT beleg_nr, belegart, status, kunde_nr, kunde_name, kunde_adresse, "
                              "kunde_ustid FROM tbl_faktura_belege WHERE id=?", (bid,)).fetchone()
        if not row or row[1] != "rechnung" or row[2] != "festgeschrieben":
            messagebox.showinfo("Gutschrift", "Gutschriften nur zu festgeschriebenen Rechnungen.")
            return
        betrag = self._frage_betrag(f"Gutschrift zu {row[0]}", "Gutschrift-Betrag (brutto, €):")
        if betrag is None or betrag <= 0:
            return
        ma = aktueller_mitarbeiter()
        satz = _parse_num(get_setting("ust_satz_standard")) or 19.0
        netto = round(betrag / (1 + satz / 100.0), 2)
        ust = round(betrag - netto, 2)
        nr = naechste_nummer("gutschrift")
        heute = date.today().isoformat()
        with _con() as con:
            cur = con.execute(
                "INSERT INTO tbl_faktura_belege(belegart, beleg_nr, kunde_nr, kunde_name, kunde_adresse, "
                "kunde_ustid, beleg_datum, leistungsdatum, bezug_beleg_id, netto, ust_betrag, brutto, "
                "status, mitarbeiter, mitarbeiter_email, festgeschrieben_am, notizen) "
                "VALUES('gutschrift',?,?,?,?,?,?,?,?,?,?,?, 'festgeschrieben',?,?,?,?)",
                (nr, row[3], row[4], row[5], row[6], heute, heute, bid, netto, ust, betrag,
                 ma["name"], ma["email"], _now(), f"Bonus-Gutschrift zu Rechnung {row[0]} (§17 UStG)"))
            gid = cur.lastrowid
            con.execute("INSERT INTO tbl_faktura_positionen(beleg_id, pos_nr, pzn, bezeichnung, menge, "
                        "apu_einzel, rabatt, ust_satz, netto_zeile, ust_zeile, brutto_zeile) "
                        "VALUES(?,1,'','Bonus-Gutschrift', 1, ?, 0, ?, ?, ?, ?)",
                        (gid, netto, satz, netto, ust, betrag))
            con.commit()
        beleg = self._beleg_dict(gid)
        pdf = erzeuge_pdf(beleg, self._positionen_dict(gid))
        with _con() as con:
            con.execute("UPDATE tbl_faktura_belege SET pdf_pfad=? WHERE id=?", (pdf, gid))
            con.commit()
        _log("gutschrift_erstellt", gid, f"{nr} zu Rechnung {row[0]}: {_eur(betrag)}")
        messagebox.showinfo("Gutschrift", f"Gutschrift {nr} über {_eur(betrag)} erstellt.")
        self._lade_rechnungen()

    def _storno_der_auswahl(self):
        bid = self._auswahl_id()
        if not bid:
            messagebox.showinfo("Storno", "Bitte zuerst eine Rechnung auswählen.")
            return
        with _con() as con:
            row = con.execute("SELECT beleg_nr, belegart, status FROM tbl_faktura_belege "
                              "WHERE id=?", (bid,)).fetchone()
        if not row or row[1] != "rechnung":
            messagebox.showinfo("Storno", "Nur Rechnungen können storniert werden.")
            return
        if row[2] != "festgeschrieben":
            messagebox.showinfo("Storno", "Nur festgeschriebene Rechnungen können storniert werden.")
            return
        if not messagebox.askyesno("Storno",
                f"Rechnung {row[0]} stornieren? Es wird eine Storno-Rechnung mit negativen "
                f"Beträgen erzeugt; die Originalrechnung bleibt erhalten (GoBD)."):
            return
        orig = self._beleg_dict(bid)
        ma = aktueller_mitarbeiter()
        nr = naechste_nummer("storno")
        heute = date.today().isoformat()
        with _con() as con:
            cur = con.execute(
                "INSERT INTO tbl_faktura_belege(belegart, beleg_nr, kunde_nr, kunde_name, kunde_adresse, "
                "kunde_ustid, beleg_datum, leistungsdatum, bezug_beleg_id, netto, ust_betrag, brutto, "
                "status, mitarbeiter, mitarbeiter_email, festgeschrieben_am, notizen) "
                "VALUES('storno',?,?,?,?,?,?,?,?,?,?,?, 'festgeschrieben',?,?,?,?)",
                (nr, orig.get("kunde_nr"), orig.get("kunde_name"), orig.get("kunde_adresse"),
                 orig.get("kunde_ustid"), heute, heute, bid, -orig["netto"], -orig["ust_betrag"],
                 -orig["brutto"], ma["name"], ma["email"], _now(),
                 f"Storno der Rechnung {row[0]}"))
            sid = cur.lastrowid
            for i, p in enumerate(self._positionen_dict(bid), 1):
                con.execute(
                    "INSERT INTO tbl_faktura_positionen(beleg_id, pos_nr, pzn, bezeichnung, menge, "
                    "apu_einzel, rabatt, ust_satz, netto_zeile, ust_zeile, brutto_zeile) "
                    "VALUES(?,?,?,?,?,?,?,?,?,?,?)",
                    (sid, i, p["pzn"], p["bezeichnung"], -p["menge"], p["apu_einzel"], p["rabatt"],
                     p["ust_satz"], -p["netto_zeile"], -p["ust_zeile"], -p["brutto_zeile"]))
            con.execute("UPDATE tbl_faktura_belege SET status='storniert' WHERE id=?", (bid,))
            # Verknuepfte Auftraege wieder freigeben -> tauchen erneut in der
            # Auftragsliste auf und sind in der Kasse weiter bearbeitbar.
            if _spalte_existiert(con, "tbl_bestellungen", "faktura_beleg_id"):
                con.execute("UPDATE tbl_bestellungen SET faktura_beleg_id=NULL, abgerechnet_am=NULL "
                            "WHERE faktura_beleg_id=?", (bid,))
            con.commit()
        beleg = self._beleg_dict(sid)
        pdf = erzeuge_pdf(beleg, self._positionen_dict(sid))
        with _con() as con:
            con.execute("UPDATE tbl_faktura_belege SET pdf_pfad=? WHERE id=?", (pdf, sid))
            con.commit()
        _log("storniert", sid, f"{nr} storniert {row[0]}")
        verk = self._storno_verkettete_gutschriften(bid)
        msg = f"Storno-Rechnung {nr} erstellt. Original {row[0]} ist storniert."
        if verk:
            msg += f"\nVerknüpfte Gutschrift(en) ebenfalls storniert: {', '.join(verk)}."
        messagebox.showinfo("Storno", msg)
        self._lade_rechnungen()

    def _storno_verkettete_gutschriften(self, rechnung_bid):
        """Storniert ALLE an einer stornierten Rechnung hängenden Gutschriften (automatisch
        UND manuell), falls in den Einstellungen aktiviert. Rückgabe: Liste der GU-Nummern."""
        if get_setting("auto_gutschrift_storno", "1") != "1":
            return []
        with _con() as con:
            ids = [r[0] for r in con.execute(
                "SELECT id FROM tbl_faktura_belege WHERE belegart='gutschrift' AND bezug_beleg_id=? "
                "AND status='festgeschrieben'",
                (rechnung_bid,)).fetchall()]
        erzeugt = []
        for gid in ids:
            g = self._beleg_dict(gid)
            ma = aktueller_mitarbeiter()
            nr = naechste_nummer("storno")
            heute = date.today().isoformat()
            with _con() as con:
                cur = con.execute(
                    "INSERT INTO tbl_faktura_belege(belegart, beleg_nr, kunde_nr, kunde_name, kunde_adresse, "
                    "kunde_ustid, beleg_datum, leistungsdatum, bezug_beleg_id, netto, ust_betrag, brutto, "
                    "status, mitarbeiter, mitarbeiter_email, festgeschrieben_am, notizen) "
                    "VALUES('storno',?,?,?,?,?,?,?,?,?,?,?, 'festgeschrieben',?,?,?,?)",
                    (nr, g.get("kunde_nr"), g.get("kunde_name"), g.get("kunde_adresse"), g.get("kunde_ustid"),
                     heute, heute, gid, -(g.get("netto") or 0), -(g.get("ust_betrag") or 0),
                     -(g.get("brutto") or 0), ma["name"], ma["email"], _now(),
                     f"Storno der Gutschrift {g.get('beleg_nr')}"))
                sid = cur.lastrowid
                for i, p in enumerate(self._positionen_dict(gid), 1):
                    con.execute(
                        "INSERT INTO tbl_faktura_positionen(beleg_id, pos_nr, pzn, bezeichnung, menge, "
                        "apu_einzel, rabatt, ust_satz, netto_zeile, ust_zeile, brutto_zeile) "
                        "VALUES(?,?,?,?,?,?,?,?,?,?,?)",
                        (sid, i, p["pzn"], p["bezeichnung"], -p["menge"], p["apu_einzel"], p["rabatt"],
                         p["ust_satz"], -p["netto_zeile"], -p["ust_zeile"], -p["brutto_zeile"]))
                con.execute("UPDATE tbl_faktura_belege SET status='storniert' WHERE id=?", (gid,))
                con.commit()
            beleg = self._beleg_dict(sid)
            pdf = erzeuge_pdf(beleg, self._positionen_dict(sid))
            with _con() as con:
                con.execute("UPDATE tbl_faktura_belege SET pdf_pfad=? WHERE id=?", (pdf, sid))
                con.commit()
            _log("auto_gutschrift_storniert", sid, f"{nr} storniert {g.get('beleg_nr')}")
            erzeugt.append(g.get("beleg_nr"))
        return erzeugt

    def _frage_betrag(self, titel, prompt):
        dlg = tk.Toplevel(self)
        dlg.title(titel)
        dlg.configure(bg=BG)
        dlg.transient(self.winfo_toplevel())
        dlg.grab_set()
        tk.Label(dlg, text=prompt, bg=BG, fg=TEXT, font=theme.BODY).pack(padx=20, pady=(18, 6))
        var = tk.StringVar()
        ent = ttk.Entry(dlg, textvariable=var, width=20)
        ent.pack(padx=20)
        ent.focus_set()
        res = {"v": None}

        def ok():
            res["v"] = _parse_num(var.get())
            dlg.destroy()
        theme.PillButton(dlg, "Erstellen", ok, kind="success", font_size=10).pack(pady=14)
        ent.bind("<Return>", lambda e: ok())
        self.wait_window(dlg)
        return res["v"]

    # ── Seite: Neue Rechnung ─────────────────────────────────────────────────
    def _page_neu(self):
        self.positionen = []
        theme.page_header(self.content, "Neue Rechnung",
                          "APU wird beim Hinzufügen eingefroren · Sachbearbeiter automatisch",
                          bg=SHELL_BG).pack(fill="x", padx=24, pady=(20, 8))
        wrap = tk.Frame(self.content, bg=SHELL_BG)
        wrap.pack(fill="both", expand=True, padx=24, pady=(0, 16))

        # Rechnungsdatum
        drow = tk.Frame(wrap, bg=SHELL_BG)
        drow.pack(fill="x", pady=(0, 8))
        tk.Label(drow, text="Rechnungsdatum:", bg=SHELL_BG, fg=MUTED,
                 font=theme.SMALL).pack(side="left")
        self.neu_datum = tk.StringVar(value=date.today().isoformat())
        ttk.Entry(drow, textvariable=self.neu_datum, width=12).pack(side="left", padx=6)
        tk.Label(drow, text="(JJJJ-MM-TT oder TT.MM.JJJJ · leer = heute)", bg=SHELL_BG,
                 fg=theme.FAINT, font=theme.SMALL).pack(side="left")

        # Kunde
        k_outer, k_body = _card(wrap, "Kunde")
        k_outer.pack(fill="x")
        self.kunde_var = tk.StringVar()
        srow = tk.Frame(k_body, bg=BG)
        srow.pack(fill="x")
        tk.Label(srow, text="Suche (Nr./Name):", bg=BG, fg=MUTED, font=theme.SMALL).pack(side="left")
        self.kunde_such = ttk.Entry(srow, width=30)
        self.kunde_such.pack(side="left", padx=6)
        self.kunde_such.bind("<KeyRelease>", lambda e: self._kunde_suche())
        self.kunde_lbl = tk.Label(srow, textvariable=self.kunde_var, bg=BG, fg=ACCENT,
                                  font=theme.BODY_BOLD)
        self.kunde_lbl.pack(side="left", padx=(16, 0))
        self.kunde_list = tk.Listbox(k_body, height=4, font=theme.SMALL)
        self.kunde_list.pack(fill="x", pady=(6, 0))
        self.kunde_list.bind("<<ListboxSelect>>", lambda e: self._kunde_waehlen())
        self._kunde = {}

        # Position erfassen
        p_outer, p_body = _card(wrap, "Position hinzufügen")
        p_outer.pack(fill="x", pady=(12, 0))
        prow = tk.Frame(p_body, bg=BG)
        prow.pack(fill="x")
        tk.Label(prow, text="Artikel (PZN/Name):", bg=BG, fg=MUTED, font=theme.SMALL).pack(side="left")
        self.art_such = ttk.Entry(prow, width=26)
        self.art_such.pack(side="left", padx=6)
        self.art_such.bind("<KeyRelease>", lambda e: self._art_suche())
        tk.Label(prow, text="Menge:", bg=BG, fg=MUTED, font=theme.SMALL).pack(side="left", padx=(12, 2))
        self.menge_var = tk.StringVar(value="1")
        ttk.Entry(prow, textvariable=self.menge_var, width=6).pack(side="left")
        tk.Label(prow, text="USt %:", bg=BG, fg=MUTED, font=theme.SMALL).pack(side="left", padx=(12, 2))
        self.ust_var = tk.StringVar(value=get_setting("ust_satz_standard"))
        ttk.Entry(prow, textvariable=self.ust_var, width=5).pack(side="left")
        theme.PillButton(prow, "Hinzufügen", self._position_add, kind="accent",
                         font_size=10, padx=12, pady=6).pack(side="left", padx=(12, 0))
        self.art_list = tk.Listbox(p_body, height=4, font=theme.SMALL)
        self.art_list.pack(fill="x", pady=(6, 0))
        self.art_list.bind("<<ListboxSelect>>", lambda e: self._art_waehlen())
        self._artikel = {}

        # Positionsliste
        l_outer, l_body = _card(wrap, "Positionen")
        l_outer.pack(fill="both", expand=True, pady=(12, 0))
        cols = ("pzn", "bez", "menge", "apu", "ust", "netto")
        tree = ttk.Treeview(l_body, columns=cols, show="headings", height=6)
        for c, t, w in (("pzn", "PZN", 90), ("bez", "Artikel", 240), ("menge", "Menge", 70),
                        ("apu", "APU", 100), ("ust", "USt", 60), ("netto", "Netto", 110)):
            tree.heading(c, text=t)
            tree.column(c, width=w, anchor="e" if c in ("apu", "netto", "menge") else "w")
        tree.pack(fill="both", expand=True)
        theme.style_treeview(tree)
        self._pos_tree = tree
        tree.bind("<Delete>", lambda e: self._position_entfernen())

        # Summen als Kennzahl-Karten (Tagesabschluss-Stil)
        _, self._summe_labels = self._metric_cards(wrap, [
            ("netto", "Netto"), ("ust", "zzgl. USt"), ("brutto", "Gesamt (brutto)", OK_GREEN)])
        # Aktionen
        foot = tk.Frame(wrap, bg=SHELL_BG)
        foot.pack(fill="x", pady=(8, 0))
        theme.PillButton(foot, "✔  Festschreiben + PDF", self._festschreiben,
                         kind="success", font_size=10, padx=14, pady=8).pack(side="right")
        theme.PillButton(foot, "Als Entwurf speichern", self._entwurf_speichern,
                         kind="neutral", font_size=10, padx=14, pady=8).pack(side="right", padx=(0, 8))
        theme.PillButton(foot, "Position löschen", self._position_entfernen,
                         kind="ghost", font_size=10, padx=12, pady=8).pack(side="right", padx=(0, 8))
        self._pos_refresh()

    def _kunde_suche(self):
        q = self.kunde_such.get().strip()
        self.kunde_list.delete(0, "end")
        self._kunde_treffer = []
        if len(q) < 2:
            return
        like = f"%{q}%"
        with _con() as con:
            if not _table_exists(con, "tbl_kunden_center"):
                return
            rows = con.execute(
                "SELECT kundennummer, kundenname, COALESCE(strasse,''), COALESCE(plz,''), "
                "COALESCE(ort,'') FROM tbl_kunden_center "
                "WHERE kundenname LIKE ? OR kundennummer LIKE ? ORDER BY kundenname LIMIT 20",
                (like, like)).fetchall()
        self._kunde_treffer = rows
        for r in rows:
            self.kunde_list.insert("end", f"{r[0]}  ·  {r[1]}  ·  {r[3]} {r[4]}")

    def _kunde_waehlen(self):
        sel = self.kunde_list.curselection()
        if not sel:
            return
        r = self._kunde_treffer[sel[0]]
        adresse = " · ".join(x for x in (r[2], f"{r[3]} {r[4]}".strip()) if x)
        self._kunde = {"nr": r[0], "name": r[1], "adresse": adresse}
        self.kunde_var.set(f"✓ {r[1]}")

    def _art_suche(self):
        q = self.art_such.get().strip()
        self.art_list.delete(0, "end")
        self._art_treffer = []
        if len(q) < 2:
            return
        like = f"%{q}%"
        with _con() as con:
            if not _table_exists(con, "tbl_nmg_stamm"):
                return
            rows = con.execute(
                "SELECT pzn, artikelname, apu FROM tbl_nmg_stamm "
                "WHERE pzn LIKE ? OR artikelname LIKE ? ORDER BY artikelname LIMIT 20",
                (like, like)).fetchall()
        self._art_treffer = rows
        for r in rows:
            self.art_list.insert("end", f"{r[0]}  ·  {r[1]}  ·  {_eur(r[2])}")

    def _art_waehlen(self):
        sel = self.art_list.curselection()
        if not sel:
            return
        r = self._art_treffer[sel[0]]
        self._artikel = {"pzn": r[0], "bezeichnung": r[1], "apu": r[2] or 0.0}
        self.art_such.delete(0, "end")
        self.art_such.insert(0, f"{r[1]}")

    def _position_add(self):
        if not self._artikel:
            messagebox.showinfo("Position", "Bitte zuerst einen Artikel auswählen.")
            return
        menge = _parse_num(self.menge_var.get()) or 1
        satz = _parse_num(self.ust_var.get()) or 19
        apu = self._artikel["apu"]
        netto = round(apu * menge, 2)
        ust = round(netto * satz / 100.0, 2)
        self.positionen.append({
            "pzn": self._artikel["pzn"], "bezeichnung": self._artikel["bezeichnung"],
            "menge": menge, "apu_einzel": apu, "rabatt": 0, "ust_satz": satz,
            "netto_zeile": netto, "ust_zeile": ust, "brutto_zeile": round(netto + ust, 2)})
        self._artikel = {}
        self._pos_refresh()

    def _position_entfernen(self):
        sel = self._pos_tree.selection()
        if not sel:
            return
        idx = self._pos_tree.index(sel[0])
        if 0 <= idx < len(self.positionen):
            del self.positionen[idx]
            self._pos_refresh()

    def _pos_refresh(self):
        tree = self._pos_tree
        tree.delete(*tree.get_children())
        netto = ust = 0.0
        for p in self.positionen:
            netto += p["netto_zeile"]
            ust += p["ust_zeile"]
            tree.insert("", "end", values=(p["pzn"], p["bezeichnung"], f'{p["menge"]:g}',
                        _eur(p["apu_einzel"]), f'{p["ust_satz"]:g} %', _eur(p["netto_zeile"])))
        self._summe_labels["netto"].config(text=_eur(netto))
        self._summe_labels["ust"].config(text=_eur(ust))
        self._summe_labels["brutto"].config(text=_eur(netto + ust))

    def _validieren(self):
        if not self._kunde:
            messagebox.showinfo("Rechnung", "Bitte einen Kunden auswählen.")
            return False
        if not self.positionen:
            messagebox.showinfo("Rechnung", "Mindestens eine Position erforderlich.")
            return False
        return True

    def _summen(self):
        netto = round(sum(p["netto_zeile"] for p in self.positionen), 2)
        ust = round(sum(p["ust_zeile"] for p in self.positionen), 2)
        return netto, ust, round(netto + ust, 2)

    def _beleg_speichern(self, status, beleg_nr=None):
        netto, ust, brutto = self._summen()
        ma = aktueller_mitarbeiter()
        datum = _parse_datum(self.neu_datum.get() if hasattr(self, "neu_datum") else "")
        with _con() as con:
            cur = con.execute(
                "INSERT INTO tbl_faktura_belege(belegart, beleg_nr, kunde_nr, kunde_name, kunde_adresse, "
                "beleg_datum, leistungsdatum, netto, ust_betrag, brutto, status, mitarbeiter, "
                "mitarbeiter_email, festgeschrieben_am) "
                "VALUES('rechnung',?,?,?,?,?,?,?,?,?,?,?,?,?)",
                (beleg_nr, self._kunde["nr"], self._kunde["name"], self._kunde["adresse"],
                 datum, datum, netto, ust, brutto, status, ma["name"], ma["email"],
                 _now() if status == "festgeschrieben" else None))
            bid = cur.lastrowid
            for i, p in enumerate(self.positionen, 1):
                con.execute(
                    "INSERT INTO tbl_faktura_positionen(beleg_id, pos_nr, pzn, bezeichnung, menge, "
                    "apu_einzel, rabatt, ust_satz, netto_zeile, ust_zeile, brutto_zeile) "
                    "VALUES(?,?,?,?,?,?,?,?,?,?,?)",
                    (bid, i, p["pzn"], p["bezeichnung"], p["menge"], p["apu_einzel"], p["rabatt"],
                     p["ust_satz"], p["netto_zeile"], p["ust_zeile"], p["brutto_zeile"]))
            con.commit()
        return bid

    def _entwurf_speichern(self):
        if not self._validieren():
            return
        bid = self._beleg_speichern("entwurf")
        _log("entwurf_gespeichert", bid)
        messagebox.showinfo("Entwurf", "Rechnung als Entwurf gespeichert.")
        self.show("rechnungen")

    def _festschreiben(self):
        if not self._validieren():
            return
        if not get_setting("firma_name"):
            if not messagebox.askyesno("Stammdaten fehlen",
                    "Es sind noch keine Firmenstammdaten hinterlegt. Trotzdem festschreiben?"):
                return
        nr = naechste_nummer("rechnung")
        bid = self._beleg_speichern("festgeschrieben", beleg_nr=nr)
        beleg = self._beleg_dict(bid)
        pdf = erzeuge_pdf(beleg, self._positionen_dict(bid))
        with _con() as con:
            con.execute("UPDATE tbl_faktura_belege SET pdf_pfad=? WHERE id=?", (pdf, bid))
            con.commit()
        _log("festgeschrieben", bid, nr)
        if pdf and os.path.exists(pdf):
            try:
                os.startfile(pdf)  # type: ignore[attr-defined]
            except Exception:
                pass
        gut = self._auto_gutschrift(bid)
        msg = f"Rechnung {nr} festgeschrieben.\nPDF: {pdf or '—'}"
        if gut:
            msg += f"\n\nAutomatische Gutschrift {gut['nr']} über {_eur(gut['brutto'])} erstellt."
        messagebox.showinfo("Festgeschrieben", msg)
        self.show("rechnungen")

    def _auto_gutschrift(self, rechnung_bid):
        """Erzeugt – falls aktiviert – automatisch eine Gutschrift zur Rechnung.
        Berechnung per Einstellung:
          * prozent_netto / prozent_brutto / fix -> ein Sammelbetrag
          * auftragsrabatt -> je Position der Rabatt aus dem Auftrag (APU × Menge × Rabatt%),
            als detaillierte Gutschrift. Rückgabe: {'nr','brutto'} oder None."""
        if get_setting("auto_gutschrift", "0") != "1":
            return None
        typ = get_setting("auto_gutschrift_typ", "prozent_netto")
        r = self._beleg_dict(rechnung_bid)
        satz = _parse_num(get_setting("ust_satz_standard")) or 19.0

        positionen = []
        if typ == "auftragsrabatt":
            # Rabatt je Position aus dem Auftrag in Gutschrift-Zeilen umwandeln.
            for p in self._positionen_dict(rechnung_bid):
                rab = p.get("rabatt") or 0
                apu = p.get("apu_einzel") or 0
                menge = p.get("menge") or 0
                if rab <= 0 or apu <= 0 or menge == 0:
                    continue
                pos_satz = p.get("ust_satz") or satz
                netto = round(apu * menge * rab / 100.0, 2)
                if netto <= 0:
                    continue
                ust = round(netto * pos_satz / 100.0, 2)
                positionen.append({"pzn": p.get("pzn") or "",
                    "bezeichnung": f"{p.get('bezeichnung') or ''} – {rab:g}% Rabatt",
                    "menge": menge, "apu_einzel": apu, "rabatt": rab, "ust_satz": pos_satz,
                    "netto_zeile": netto, "ust_zeile": ust, "brutto_zeile": round(netto + ust, 2)})
            if not positionen:
                return None
            notiz = f"Automatische Gutschrift (Auftragsrabatt) zu Rechnung {r.get('beleg_nr')} (§17 UStG)"
        else:
            wert = _parse_num(get_setting("auto_gutschrift_wert", "0"))
            if wert <= 0:
                return None
            if typ == "prozent_netto":
                netto = round((r.get("netto") or 0) * wert / 100.0, 2)
                ust = round(netto * satz / 100.0, 2)
                brutto = round(netto + ust, 2)
            else:  # prozent_brutto oder fix -> Bruttobetrag, dann §17-Split
                brutto = round((r.get("brutto") or 0) * wert / 100.0, 2) if typ == "prozent_brutto" \
                    else round(wert, 2)
                netto = round(brutto / (1 + satz / 100.0), 2)
                ust = round(brutto - netto, 2)
            if brutto <= 0:
                return None
            positionen = [{"pzn": "", "bezeichnung": "Gutschrift zur Rechnung", "menge": 1,
                           "apu_einzel": netto, "rabatt": 0, "ust_satz": satz,
                           "netto_zeile": netto, "ust_zeile": ust, "brutto_zeile": brutto}]
            notiz = f"Automatische Gutschrift zu Rechnung {r.get('beleg_nr')} (§17 UStG)"

        netto_s = round(sum(p["netto_zeile"] for p in positionen), 2)
        ust_s = round(sum(p["ust_zeile"] for p in positionen), 2)
        brutto_s = round(netto_s + ust_s, 2)
        ma = aktueller_mitarbeiter()
        heute = date.today().isoformat()
        nr = naechste_nummer("gutschrift")
        with _con() as con:
            cur = con.execute(
                "INSERT INTO tbl_faktura_belege(belegart, beleg_nr, kunde_nr, kunde_name, kunde_adresse, "
                "kunde_ustid, beleg_datum, leistungsdatum, bezug_beleg_id, netto, ust_betrag, brutto, "
                "status, mitarbeiter, mitarbeiter_email, festgeschrieben_am, notizen) "
                "VALUES('gutschrift',?,?,?,?,?,?,?,?,?,?,?, 'festgeschrieben',?,?,?,?)",
                (nr, r.get("kunde_nr"), r.get("kunde_name"), r.get("kunde_adresse"), r.get("kunde_ustid"),
                 heute, heute, rechnung_bid, netto_s, ust_s, brutto_s, ma["name"], ma["email"], _now(),
                 notiz))
            gid = cur.lastrowid
            for i, p in enumerate(positionen, 1):
                con.execute("INSERT INTO tbl_faktura_positionen(beleg_id, pos_nr, pzn, bezeichnung, menge, "
                            "apu_einzel, rabatt, ust_satz, netto_zeile, ust_zeile, brutto_zeile) "
                            "VALUES(?,?,?,?,?,?,?,?,?,?,?)",
                            (gid, i, p["pzn"], p["bezeichnung"], p["menge"], p["apu_einzel"], p["rabatt"],
                             p["ust_satz"], p["netto_zeile"], p["ust_zeile"], p["brutto_zeile"]))
            con.commit()
        beleg = self._beleg_dict(gid)
        pdf = erzeuge_pdf(beleg, self._positionen_dict(gid))
        with _con() as con:
            con.execute("UPDATE tbl_faktura_belege SET pdf_pfad=? WHERE id=?", (pdf, gid))
            con.commit()
        _log("auto_gutschrift", gid, f"{nr} zu {r.get('beleg_nr')}: {_eur(brutto_s)}")
        return {"nr": nr, "brutto": brutto_s}

    def _beleg_dict(self, bid) -> dict:
        with _con() as con:
            con.row_factory = sqlite3.Row
            r = con.execute("SELECT * FROM tbl_faktura_belege WHERE id=?", (bid,)).fetchone()
        return dict(r) if r else {}

    def _positionen_dict(self, bid) -> list[dict]:
        with _con() as con:
            con.row_factory = sqlite3.Row
            rows = con.execute("SELECT * FROM tbl_faktura_positionen WHERE beleg_id=? ORDER BY pos_nr",
                               (bid,)).fetchall()
        return [dict(r) for r in rows]

    # ── Seite: Quartalsvergütung ─────────────────────────────────────────────
    def _page_verguetung(self):
        basis = get_setting("bonus_basis", "netto")
        theme.page_header(self.content, "Quartalsvergütung",
                          f"Feste Euro-Vergütung je Kunde ab Quartals-Umsatzschwelle "
                          f"({basis.capitalize()} aus festgeschriebenen Rechnungen, ohne Storno)",
                          bg=SHELL_BG).pack(fill="x", padx=24, pady=(20, 8))
        bar = tk.Frame(self.content, bg=SHELL_BG)
        bar.pack(fill="x", padx=24, pady=(0, 8))
        heute = date.today()
        akt_q = (heute.month - 1) // 3 + 1
        vor_q = akt_q - 1 or 4
        vor_jahr = heute.year if akt_q > 1 else heute.year - 1
        tk.Label(bar, text="Jahr:", bg=SHELL_BG, fg=MUTED, font=theme.SMALL).pack(side="left")
        self.verg_jahr = tk.StringVar(value=str(vor_jahr))
        ttk.Entry(bar, textvariable=self.verg_jahr, width=6).pack(side="left", padx=(4, 10))
        tk.Label(bar, text="Quartal:", bg=SHELL_BG, fg=MUTED, font=theme.SMALL).pack(side="left")
        self.verg_quartal = tk.StringVar(value=str(vor_q))
        ttk.Combobox(bar, textvariable=self.verg_quartal, width=4, state="readonly",
                     values=["1", "2", "3", "4"], style="NMG.TCombobox").pack(side="left", padx=(4, 10))
        theme.PillButton(bar, "Berechnen", self._verg_berechnen, kind="accent",
                         font_size=10, padx=14, pady=6).pack(side="left", padx=(6, 0))
        theme.PillButton(bar, "💶  Vergütungen erzeugen", self._verg_erzeugen,
                         kind="success", font_size=10, padx=14, pady=6).pack(side="left", padx=(8, 0))

        outer, body = _card(self.content, "Vorschau")
        outer.pack(fill="both", expand=True, padx=24, pady=(0, 20))
        cols = ("kunde", "umsatz", "stufe", "bonus", "status")
        tree = ttk.Treeview(body, columns=cols, show="headings", height=12)
        for c, t, w in (("kunde", "Kunde", 240), ("umsatz", "Quartalsumsatz", 130),
                        ("stufe", "Stufe", 160), ("bonus", "Vergütung", 110), ("status", "Status", 140)):
            tree.heading(c, text=t)
            tree.column(c, width=w, anchor="e" if c in ("umsatz", "bonus") else "w")
        tree.pack(fill="both", expand=True)
        theme.style_treeview(tree)
        self._verg_tree = tree
        self._verg_preview = []
        if self._staffel_leer():
            tk.Label(body, text="Hinweis: Es sind noch keine Staffel-Stufen hinterlegt "
                                "(siehe Staffel).", bg=BG, fg=theme.WARNING,
                     font=theme.SMALL).pack(anchor="w", pady=(6, 0))

    def _staffel_leer(self) -> bool:
        with _con() as con:
            return con.execute("SELECT COUNT(*) FROM tbl_faktura_bonus_staffel").fetchone()[0] == 0

    def _staffel_bonus(self, umsatz: float):
        """Findet die passende Stufe für einen Umsatz. Rückgabe (bonus, bezeichnung)
        oder (0, None), wenn keine Stufe greift."""
        with _con() as con:
            rows = con.execute("SELECT schwelle_von, schwelle_bis, bonus_betrag, bezeichnung "
                               "FROM tbl_faktura_bonus_staffel ORDER BY schwelle_von").fetchall()
        treffer = None
        for von, bis, bonus, bez in rows:
            if umsatz >= (von or 0) and (bis is None or umsatz <= bis):
                treffer = (bonus or 0.0, bez)  # höchste passende Stufe gewinnt
        return treffer or (0.0, None)

    def _quartal_grenzen(self, jahr: str, quartal: int):
        m1 = (quartal - 1) * 3 + 1
        m3 = m1 + 2
        letzter = {1: 31, 2: 30, 3: 30, 4: 31}[quartal]
        return m1, m3, f"{jahr}-{m1:02d}-01", f"{jahr}-{m3:02d}-{letzter}"

    def _verg_berechnen(self):
        jahr = (self.verg_jahr.get() or "").strip()
        if len(jahr) != 4 or not jahr.isdigit():
            messagebox.showinfo("Quartalsvergütung", "Bitte ein gültiges Jahr (JJJJ) eingeben.")
            return
        quartal = int(self.verg_quartal.get() or "1")
        m1, m3, von, bis = self._quartal_grenzen(jahr, quartal)
        marker = f"Quartalsvergütung Q{quartal}/{jahr}"
        # Basis: festgeschriebene Rechnungen (status='festgeschrieben' schließt stornierte
        # aus, da ein Storno das Original auf 'storniert' setzt). Netto/Brutto per Einstellung.
        feld = "brutto" if get_setting("bonus_basis", "netto") == "brutto" else "netto"
        with _con() as con:
            rows = con.execute(
                f"SELECT kunde_nr, kunde_name, COALESCE(SUM({feld}),0) AS umsatz "
                "FROM tbl_faktura_belege WHERE belegart='rechnung' AND status='festgeschrieben' "
                "AND substr(beleg_datum,1,4)=? AND CAST(substr(beleg_datum,6,2) AS INTEGER) BETWEEN ? AND ? "
                "GROUP BY kunde_nr, kunde_name ORDER BY umsatz DESC", (jahr, m1, m3)).fetchall()
            schon = {r[0] for r in con.execute(
                "SELECT kunde_nr FROM tbl_faktura_belege WHERE belegart='quartalsverguetung' "
                "AND notizen LIKE ?", (marker + "%",)).fetchall()}
        tree = self._verg_tree
        tree.delete(*tree.get_children())
        self._verg_preview = []
        for kunde_nr, name, umsatz in rows:
            bonus, bez = self._staffel_bonus(umsatz)
            if bonus <= 0:
                status = "keine Stufe"
            elif kunde_nr in schon:
                status = "bereits erzeugt"
            else:
                status = "wird erzeugt"
                self._verg_preview.append({"jahr": jahr, "quartal": quartal, "von": von, "bis": bis,
                                           "kunde_nr": kunde_nr, "kunde_name": name,
                                           "umsatz": umsatz, "bonus": bonus})
            tree.insert("", "end", values=(name or kunde_nr or "—", _eur(umsatz),
                        bez or "—", _eur(bonus) if bonus else "—", status))
        if not rows:
            messagebox.showinfo("Quartalsvergütung",
                                f"Keine festgeschriebenen Rechnungen in Q{quartal}/{jahr}.")

    def _verg_erzeugen(self):
        if not self._verg_preview:
            messagebox.showinfo("Quartalsvergütung", "Bitte zuerst 'Berechnen' – es sind keine "
                                "neuen Vergütungen offen.")
            return
        if not messagebox.askyesno("Quartalsvergütung",
                f"{len(self._verg_preview)} Vergütung(en) erzeugen?"):
            return
        satz = _parse_num(get_setting("ust_satz_standard")) or 19.0
        ma = aktueller_mitarbeiter()
        heute = date.today().isoformat()
        anzahl = 0
        for e in self._verg_preview:
            netto = round(e["bonus"] / (1 + satz / 100.0), 2)
            ust = round(e["bonus"] - netto, 2)
            nr = naechste_nummer("quartalsverguetung")
            marker = f"Quartalsvergütung Q{e['quartal']}/{e['jahr']}"
            with _con() as con:
                cur = con.execute(
                    "INSERT INTO tbl_faktura_belege(belegart, beleg_nr, kunde_nr, kunde_name, "
                    "beleg_datum, leistungsdatum, zeitraum_von, zeitraum_bis, netto, ust_betrag, "
                    "brutto, status, mitarbeiter, mitarbeiter_email, festgeschrieben_am, notizen) "
                    "VALUES('quartalsverguetung',?,?,?,?,?,?,?,?,?,?, 'festgeschrieben',?,?,?,?)",
                    (nr, e["kunde_nr"], e["kunde_name"], heute, heute, e["von"], e["bis"],
                     netto, ust, e["bonus"], ma["name"], ma["email"], _now(),
                     f"{marker} (Umsatz {_eur(e['umsatz'])}, §17 UStG)"))
                gid = cur.lastrowid
                con.execute("INSERT INTO tbl_faktura_positionen(beleg_id, pos_nr, pzn, bezeichnung, "
                            "menge, apu_einzel, rabatt, ust_satz, netto_zeile, ust_zeile, brutto_zeile) "
                            "VALUES(?,1,'',?,1,?,0,?,?,?,?)",
                            (gid, marker, netto, satz, netto, ust, e["bonus"]))
                con.commit()
            beleg = self._beleg_dict(gid)
            pdf = erzeuge_pdf(beleg, self._positionen_dict(gid))
            with _con() as con:
                con.execute("UPDATE tbl_faktura_belege SET pdf_pfad=? WHERE id=?", (pdf, gid))
                con.commit()
            _log("quartalsverguetung_erstellt", gid,
                 f"{nr} {e['kunde_name']} Q{e['quartal']}/{e['jahr']}: {_eur(e['bonus'])}")
            anzahl += 1
        messagebox.showinfo("Quartalsvergütung", f"{anzahl} Vergütung(en) erstellt.")
        self._verg_berechnen()

    # ── Seite: Bonus-Staffel ─────────────────────────────────────────────────
    def _page_staffel(self):
        theme.page_header(self.content, "Monats-Bonus · Staffelung",
                          "Feste Euro-Beträge je Umsatzstufe",
                          bg=SHELL_BG).pack(fill="x", padx=24, pady=(20, 8))
        basis_bar = tk.Frame(self.content, bg=SHELL_BG)
        basis_bar.pack(fill="x", padx=24, pady=(0, 8))
        tk.Label(basis_bar, text="Schwelle messen am:", bg=SHELL_BG, fg=MUTED,
                 font=theme.SMALL).pack(side="left")
        self.bonus_basis_var = tk.StringVar(value=get_setting("bonus_basis", "netto"))
        cb = ttk.Combobox(basis_bar, textvariable=self.bonus_basis_var, width=10, state="readonly",
                          values=["netto", "brutto"], style="NMG.TCombobox")
        cb.pack(side="left", padx=6)
        cb.bind("<<ComboboxSelected>>",
                lambda e: set_setting("bonus_basis", self.bonus_basis_var.get()))
        tk.Label(basis_bar, text="(Monatsumsatz je Kunde)", bg=SHELL_BG, fg=theme.FAINT,
                 font=theme.SMALL).pack(side="left", padx=(4, 0))

        outer, body = _card(self.content, "Stufen")
        outer.pack(fill="both", expand=True, padx=24, pady=(0, 20))
        cols = ("von", "bis", "bonus", "bez")
        tree = ttk.Treeview(body, columns=cols, show="headings", height=8)
        for c, t, w in (("von", "Umsatz ab (€)", 130), ("bis", "Umsatz bis (€)", 130),
                        ("bonus", "Bonus (€)", 120), ("bez", "Bezeichnung", 220)):
            tree.heading(c, text=t)
            tree.column(c, width=w, anchor="e" if c in ("von", "bis", "bonus") else "w")
        tree.pack(fill="both", expand=True)
        theme.style_treeview(tree)
        self._staffel_tree = tree

        form = tk.Frame(body, bg=BG)
        form.pack(fill="x", pady=(10, 0))
        self.st_von = tk.StringVar(); self.st_bis = tk.StringVar()
        self.st_bonus = tk.StringVar(); self.st_bez = tk.StringVar()
        for lbl, var, w in (("ab €", self.st_von, 10), ("bis €", self.st_bis, 10),
                            ("Bonus €", self.st_bonus, 10), ("Bezeichnung", self.st_bez, 20)):
            tk.Label(form, text=lbl, bg=BG, fg=MUTED, font=theme.SMALL).pack(side="left", padx=(0, 2))
            ttk.Entry(form, textvariable=var, width=w).pack(side="left", padx=(0, 10))
        theme.PillButton(form, "Stufe hinzufügen", self._staffel_add, kind="accent",
                         font_size=10, padx=12, pady=6).pack(side="left")
        theme.PillButton(form, "Auswahl löschen", self._staffel_del, kind="ghost",
                         font_size=10, padx=12, pady=6).pack(side="left", padx=(8, 0))
        self._lade_staffel()

    def _lade_staffel(self):
        tree = self._staffel_tree
        tree.delete(*tree.get_children())
        with _con() as con:
            rows = con.execute("SELECT id, schwelle_von, schwelle_bis, bonus_betrag, bezeichnung "
                               "FROM tbl_faktura_bonus_staffel ORDER BY schwelle_von").fetchall()
        for r in rows:
            tree.insert("", "end", iid=str(r[0]),
                        values=(_eur(r[1]), _eur(r[2]) if r[2] is not None else "∞",
                                _eur(r[3]), r[4] or ""))

    def _staffel_add(self):
        von = _parse_num(self.st_von.get())
        bis = _parse_num(self.st_bis.get()) or None
        bonus = _parse_num(self.st_bonus.get())
        with _con() as con:
            con.execute("INSERT INTO tbl_faktura_bonus_staffel(gueltig_ab, schwelle_von, schwelle_bis, "
                        "bonus_betrag, bezeichnung) VALUES(?,?,?,?,?)",
                        (date.today().isoformat(), von, bis, bonus, self.st_bez.get().strip()))
            con.commit()
        for v in (self.st_von, self.st_bis, self.st_bonus, self.st_bez):
            v.set("")
        self._lade_staffel()

    def _staffel_del(self):
        sel = self._staffel_tree.selection()
        if not sel:
            return
        with _con() as con:
            con.execute("DELETE FROM tbl_faktura_bonus_staffel WHERE id=?", (int(sel[0]),))
            con.commit()
        self._lade_staffel()

    # ── Einstellungen: Firmendaten ───────────────────────────────────────────
    def _page_einst_firma(self):
        theme.page_header(self.content, "Einstellungen · Firmendaten",
                          "Versender-Stammdaten, Logo und Mitarbeiter",
                          bg=SHELL_BG).pack(fill="x", padx=24, pady=(20, 8))
        wrap = self._scrollseite()

        # Firmenstammdaten
        outer, body = _card(wrap, "Firmenstammdaten (Versender)")
        outer.pack(fill="x", pady=(0, 12))
        self._stamm_vars = {}
        grid = tk.Frame(body, bg=BG)
        grid.pack(fill="x")
        for i, (key, label, mehrzeilig) in enumerate(STAMM_FELDER):
            r, c = divmod(i, 2)
            cell = tk.Frame(grid, bg=BG)
            cell.grid(row=r, column=c, sticky="ew", padx=8, pady=5)
            grid.columnconfigure(c, weight=1)
            tk.Label(cell, text=label, bg=BG, fg=MUTED, font=theme.SMALL).pack(anchor="w")
            if mehrzeilig:
                txt = tk.Text(cell, height=3, width=40, font=theme.BODY,
                              highlightbackground=BORDER, highlightthickness=1, relief="flat")
                txt.insert("1.0", get_setting(key))
                txt.pack(fill="x")
                self._stamm_vars[key] = txt
            else:
                var = tk.StringVar(value=get_setting(key))
                ttk.Entry(cell, textvariable=var, width=40).pack(fill="x")
                self._stamm_vars[key] = var

        # Logo
        logo_row = tk.Frame(body, bg=BG)
        logo_row.pack(fill="x", pady=(8, 0))
        tk.Label(logo_row, text="Logo:", bg=BG, fg=MUTED, font=theme.SMALL).pack(side="left")
        self.logo_var = tk.StringVar(value=get_setting("firma_logo_pfad") or "— kein Logo —")
        tk.Label(logo_row, textvariable=self.logo_var, bg=BG, fg=TEXT,
                 font=theme.SMALL).pack(side="left", padx=8)
        theme.PillButton(logo_row, "Logo wählen…", self._logo_waehlen, kind="neutral",
                         font_size=10, padx=12, pady=5).pack(side="left", padx=(8, 0))
        theme.PillButton(body, "💾  Stammdaten speichern", self._stamm_speichern, kind="success",
                         font_size=10, padx=14, pady=8).pack(anchor="e", pady=(10, 0))

        # Automatische Gutschrift je Rechnung
        ag_outer, ag_body = _card(wrap, "Automatische Gutschrift je Rechnung")
        ag_outer.pack(fill="x", pady=(0, 12))
        tk.Label(ag_body, text="Beim Festschreiben einer Rechnung wird automatisch eine verknüpfte "
                               "Gutschrift erzeugt (mit §17-USt-Split).", bg=BG, fg=MUTED,
                 font=theme.SMALL, wraplength=830, justify="left").pack(anchor="w")
        self.ag_aktiv = tk.IntVar(value=1 if get_setting("auto_gutschrift", "0") == "1" else 0)
        tk.Checkbutton(ag_body, text="Automatische Gutschriften erzeugen", variable=self.ag_aktiv,
                       bg=BG, fg=TEXT, selectcolor=theme.CARD_ALT, activebackground=BG,
                       font=theme.BODY).pack(anchor="w", pady=(8, 0))
        self.ag_storno = tk.IntVar(value=1 if get_setting("auto_gutschrift_storno", "1") == "1" else 0)
        tk.Checkbutton(ag_body, text="Beim Stornieren einer Rechnung die zugehörigen Gutschriften "
                                     "(automatisch UND manuell) mit stornieren", variable=self.ag_storno,
                       bg=BG, fg=TEXT, selectcolor=theme.CARD_ALT, activebackground=BG,
                       font=theme.SMALL).pack(anchor="w")
        agrow = tk.Frame(ag_body, bg=BG)
        agrow.pack(fill="x", pady=(8, 0))
        tk.Label(agrow, text="Berechnung:", bg=BG, fg=MUTED, font=theme.SMALL).pack(side="left", padx=(0, 2))
        self._ag_typ_labels = {"prozent_netto": "% vom Netto", "prozent_brutto": "% vom Brutto",
                               "fix": "Fixbetrag (€)", "auftragsrabatt": "Auftragsrabatt (aus Positionen)"}
        self.ag_typ = tk.StringVar(value=self._ag_typ_labels.get(
            get_setting("auto_gutschrift_typ", "prozent_netto"), "% vom Netto"))
        ttk.Combobox(agrow, textvariable=self.ag_typ, width=26, state="readonly",
                     values=list(self._ag_typ_labels.values()), style="NMG.TCombobox").pack(side="left")
        tk.Label(agrow, text="Wert:", bg=BG, fg=MUTED, font=theme.SMALL).pack(side="left", padx=(14, 2))
        self.ag_wert = tk.StringVar(value=get_setting("auto_gutschrift_wert", "0"))
        ttk.Entry(agrow, textvariable=self.ag_wert, width=10).pack(side="left")
        tk.Label(ag_body, text="Bei Auftragsrabatt (aus Positionen) wird die Rechnung zum vollen APU "
                               "erstellt; die Gutschrift bildet je Artikel den Auftragsrabatt ab "
                               "(z. B. 60 %). Das Feld Wert wird dann nicht benötigt.",
                 bg=BG, fg=MUTED, font=theme.SMALL, wraplength=830, justify="left").pack(anchor="w", pady=(6, 0))
        theme.PillButton(ag_body, "💾  Speichern", self._autogut_speichern, kind="success",
                         font_size=10, padx=14, pady=8).pack(anchor="e", pady=(10, 0))

        # Mitarbeiter
        m_outer, m_body = _card(wrap, "Mitarbeiter (Sachbearbeiter)")
        m_outer.pack(fill="x", pady=(0, 24))
        akt = aktueller_mitarbeiter()
        tk.Label(m_body, text=f"Angemeldet als Windows-Benutzer: {akt['benutzer']} — "
                              f"wird Rechnungen automatisch zugeordnet.",
                 bg=BG, fg=MUTED, font=theme.SMALL, wraplength=820, justify="left").pack(anchor="w")
        cols = ("benutzer", "name", "email", "telefon")
        tree = ttk.Treeview(m_body, columns=cols, show="headings", height=5)
        for c, t, w in (("benutzer", "Windows-Benutzer", 160), ("name", "Name", 200),
                        ("email", "E-Mail", 240), ("telefon", "Telefon", 120)):
            tree.heading(c, text=t)
            tree.column(c, width=w)
        tree.pack(fill="x", pady=(8, 0))
        theme.style_treeview(tree)
        self._ma_tree = tree
        form = tk.Frame(m_body, bg=BG)
        form.pack(fill="x", pady=(8, 0))
        self.ma_benutzer = tk.StringVar(value=akt["benutzer"])
        self.ma_name = tk.StringVar(value=akt["name"] if akt["name"] != akt["benutzer"] else "")
        self.ma_email = tk.StringVar(value=akt["email"])
        self.ma_tel = tk.StringVar()
        for lbl, var, w in (("Benutzer", self.ma_benutzer, 16), ("Name", self.ma_name, 20),
                            ("E-Mail", self.ma_email, 24), ("Telefon", self.ma_tel, 12)):
            tk.Label(form, text=lbl, bg=BG, fg=MUTED, font=theme.SMALL).pack(side="left", padx=(0, 2))
            ttk.Entry(form, textvariable=var, width=w).pack(side="left", padx=(0, 8))
        theme.PillButton(form, "Speichern", self._ma_speichern, kind="accent",
                         font_size=10, padx=12, pady=6).pack(side="left")
        theme.PillButton(form, "Auswahl löschen", self._ma_del, kind="ghost",
                         font_size=10, padx=12, pady=6).pack(side="left", padx=(8, 0))
        self._lade_mitarbeiter()

    # ── Einstellungen: Belegnummern ──────────────────────────────────────────
    def _page_einst_nummern(self):
        theme.page_header(self.content, "Einstellungen · Belegnummern",
                          "Frei konfigurierbares Nummernformat je Belegart",
                          bg=SHELL_BG).pack(fill="x", padx=24, pady=(20, 8))
        wrap = self._scrollseite()

        # Belegnummern (frei konfigurierbar)
        nr_outer, nr_body = _card(wrap, "Belegnummern (frei konfigurierbar)")
        nr_outer.pack(fill="x", pady=(0, 12))
        tk.Label(nr_body, text="Platzhalter:  {JJJJ} = Jahr · {JJ} = Jahr 2-stellig · {MM} = Monat · "
                               "{NR} = Zähler · {NR:5} = Zähler 5-stellig (z. B. RE-{JJJJ}-{NR:5})",
                 bg=BG, fg=MUTED, font=theme.SMALL, wraplength=830, justify="left").pack(anchor="w")
        self._nr_vars = {}
        nrgrid = tk.Frame(nr_body, bg=BG)
        nrgrid.pack(fill="x", pady=(6, 0))
        for i, (key, label) in enumerate([("nr_format_rechnung", "Rechnung"),
                ("nr_format_gutschrift", "Gutschrift"), ("nr_format_storno", "Storno"),
                ("nr_format_quartalsverguetung", "Quartalsvergütung")]):
            r, c = divmod(i, 2)
            cell = tk.Frame(nrgrid, bg=BG)
            cell.grid(row=r, column=c, sticky="ew", padx=8, pady=5)
            nrgrid.columnconfigure(c, weight=1)
            tk.Label(cell, text=label, bg=BG, fg=MUTED, font=theme.SMALL).pack(anchor="w")
            var = tk.StringVar(value=get_setting(key, STD_NR_FORMAT.get(
                key.replace("nr_format_", ""), "")))
            ttk.Entry(cell, textvariable=var, width=30).pack(fill="x")
            self._nr_vars[key] = var
        theme.PillButton(nr_body, "💾  Nummern speichern", self._nummern_speichern, kind="success",
                         font_size=10, padx=14, pady=8).pack(anchor="e", pady=(10, 0))

    # ── Einstellungen: Layouts ───────────────────────────────────────────────
    def _page_einst_layout(self):
        theme.page_header(self.content, "Einstellungen · Layouts",
                          "Vorlage gestalten – gilt für Rechnung, Gutschrift und Quartalsvergütung",
                          bg=SHELL_BG).pack(fill="x", padx=24, pady=(20, 8))
        wrap = self._scrollseite()

        # Vorlage / Layout
        tpl_outer, tpl_body = _card(wrap, "Vorlage / Layout (gilt für Rechnung · Gutschrift · Quartalsvergütung)")
        tpl_outer.pack(fill="x", pady=(0, 12))
        self._tpl_vars = {}
        r1 = tk.Frame(tpl_body, bg=BG)
        r1.pack(fill="x")
        tk.Label(r1, text="Akzentfarbe (Hex):", bg=BG, fg=MUTED, font=theme.SMALL).pack(side="left")
        self._tpl_vars["tpl_akzentfarbe"] = tk.StringVar(value=get_setting("tpl_akzentfarbe"))
        ttk.Entry(r1, textvariable=self._tpl_vars["tpl_akzentfarbe"], width=10).pack(side="left", padx=(4, 16))
        tk.Label(r1, text="Logo-Position:", bg=BG, fg=MUTED, font=theme.SMALL).pack(side="left")
        self._tpl_vars["tpl_logo_pos"] = tk.StringVar(value=get_setting("tpl_logo_pos"))
        ttk.Combobox(r1, textvariable=self._tpl_vars["tpl_logo_pos"], width=8, state="readonly",
                     values=["links", "rechts"], style="NMG.TCombobox").pack(side="left", padx=(4, 16))
        self._tpl_apu = tk.IntVar(value=1 if get_setting("tpl_spalte_apu", "1") == "1" else 0)
        self._tpl_ust = tk.IntVar(value=1 if get_setting("tpl_spalte_ust", "1") == "1" else 0)
        tk.Checkbutton(r1, text="APU-Spalte", variable=self._tpl_apu, bg=BG, fg=TEXT,
                       selectcolor=theme.CARD_ALT, activebackground=BG, font=theme.SMALL).pack(side="left")
        tk.Checkbutton(r1, text="USt-Spalte", variable=self._tpl_ust, bg=BG, fg=TEXT,
                       selectcolor=theme.CARD_ALT, activebackground=BG, font=theme.SMALL).pack(side="left", padx=(8, 0))

        tgrid = tk.Frame(tpl_body, bg=BG)
        tgrid.pack(fill="x", pady=(8, 0))
        for i, (key, label) in enumerate([("tpl_titel_rechnung", "Titel Rechnung"),
                ("tpl_titel_gutschrift", "Titel Gutschrift"),
                ("tpl_titel_quartalsverguetung", "Titel Quartalsvergütung"),
                ("tpl_titel_storno", "Titel Storno")]):
            r, c = divmod(i, 2)
            cell = tk.Frame(tgrid, bg=BG)
            cell.grid(row=r, column=c, sticky="ew", padx=8, pady=5)
            tgrid.columnconfigure(c, weight=1)
            tk.Label(cell, text=label, bg=BG, fg=MUTED, font=theme.SMALL).pack(anchor="w")
            v = tk.StringVar(value=get_setting(key))
            ttk.Entry(cell, textvariable=v, width=30).pack(fill="x")
            self._tpl_vars[key] = v
        for key, label in [("tpl_kopftext_rechnung", "Kopftext Rechnung (optional)"),
                           ("tpl_kopftext_gutschrift", "Kopftext Gutschrift"),
                           ("tpl_kopftext_quartalsverguetung", "Kopftext Quartalsvergütung"),
                           ("tpl_fusstext", "Fußtext (alle Belege)")]:
            tk.Label(tpl_body, text=label, bg=BG, fg=MUTED, font=theme.SMALL).pack(anchor="w", pady=(6, 0))
            txt = tk.Text(tpl_body, height=2, font=theme.BODY, highlightbackground=BORDER,
                          highlightthickness=1, relief="flat")
            txt.insert("1.0", get_setting(key))
            txt.pack(fill="x")
            self._tpl_vars[key] = txt
        btnrow = tk.Frame(tpl_body, bg=BG)
        btnrow.pack(fill="x", pady=(10, 0))
        theme.PillButton(btnrow, "💾  Vorlage speichern", self._vorlage_speichern, kind="success",
                         font_size=10, padx=14, pady=8).pack(side="left")
        theme.PillButton(btnrow, "👁  Vorschau (PDF)", self._vorschau_oeffnen, kind="accent",
                         font_size=10, padx=14, pady=8).pack(side="left", padx=(8, 0))
        theme.PillButton(btnrow, "🎨  Layout frei gestalten…", self._layout_editor_oeffnen,
                         kind="primary", font_size=10, padx=14, pady=8).pack(side="left", padx=(8, 0))
        self._layout_aktiv = tk.IntVar(value=1 if get_setting("tpl_layout_aktiv", "0") == "1" else 0)
        tk.Checkbutton(btnrow, text="Freies Layout verwenden", variable=self._layout_aktiv,
                       bg=BG, fg=TEXT, selectcolor=theme.CARD_ALT, activebackground=BG, font=theme.SMALL,
                       command=lambda: set_setting("tpl_layout_aktiv", str(self._layout_aktiv.get()))
                       ).pack(side="left", padx=(12, 0))
        tk.Label(tpl_body, text="Tipp: Der Button Layout frei gestalten öffnet den Drag-&-Drop-Editor "
                                "(Vollbild) mit Beispiel-Inhalten und größenveränderbaren Feldern.",
                 bg=BG, fg=MUTED, font=theme.SMALL, wraplength=830, justify="left").pack(anchor="w", pady=(10, 0))

    def _logo_waehlen(self):
        pfad = filedialog.askopenfilename(
            title="Logo wählen", filetypes=[("Bilder", "*.png *.jpg *.jpeg *.gif"), ("Alle", "*.*")])
        if pfad:
            self.logo_var.set(pfad)
            set_setting("firma_logo_pfad", pfad)

    def _stamm_speichern(self):
        for key, widget in self._stamm_vars.items():
            wert = widget.get("1.0", "end").strip() if isinstance(widget, tk.Text) else widget.get().strip()
            set_setting(key, wert)
        _log("stammdaten_gespeichert")
        messagebox.showinfo("Gespeichert", "Firmenstammdaten gespeichert.")

    def _autogut_speichern(self):
        typ_rueck = {v: k for k, v in self._ag_typ_labels.items()}
        set_setting("auto_gutschrift", str(self.ag_aktiv.get()))
        set_setting("auto_gutschrift_storno", str(self.ag_storno.get()))
        set_setting("auto_gutschrift_typ", typ_rueck.get(self.ag_typ.get(), "prozent_netto"))
        set_setting("auto_gutschrift_wert", self.ag_wert.get().strip())
        _log("auto_gutschrift_einstellung")
        zustand = "aktiv" if self.ag_aktiv.get() else "aus"
        messagebox.showinfo("Gespeichert", f"Automatische Gutschrift: {zustand} "
                            f"({self.ag_typ.get()} {self.ag_wert.get()}).")

    def _nummern_speichern(self):
        for key, var in self._nr_vars.items():
            set_setting(key, var.get().strip())
        _log("nummernformat_gespeichert")
        beispiele = "\n".join(
            f"{k.replace('nr_format_', '').capitalize()}: "
            f"{_format_nummer(v.get().strip(), date.today().year, date.today().month, 1)}"
            for k, v in self._nr_vars.items())
        messagebox.showinfo("Gespeichert", "Belegnummern-Formate gespeichert.\n\nBeispiel:\n" + beispiele)

    def _vorlage_speichern(self, silent=False):
        set_setting("tpl_spalte_apu", "1" if self._tpl_apu.get() else "0")
        set_setting("tpl_spalte_ust", "1" if self._tpl_ust.get() else "0")
        for key, widget in self._tpl_vars.items():
            wert = widget.get("1.0", "end").strip() if isinstance(widget, tk.Text) else widget.get().strip()
            set_setting(key, wert)
        _log("vorlage_gespeichert")
        if not silent:
            messagebox.showinfo("Gespeichert", "Vorlage gespeichert. Gilt für Rechnung, "
                                "Gutschrift und Quartalsvergütung.")

    def _vorschau_oeffnen(self):
        self._vorlage_speichern(silent=True)  # aktuellen Stand übernehmen
        pfad = vorschau_pdf("rechnung")
        if pfad and os.path.exists(pfad):
            try:
                os.startfile(pfad)  # type: ignore[attr-defined]
            except Exception:
                messagebox.showinfo("Vorschau", f"Vorschau liegt unter:\n{pfad}")
        else:
            messagebox.showinfo("Vorschau", "Konnte keine Vorschau erzeugen (Browser fehlt?).")

    def _layout_editor_oeffnen(self):
        self._vorlage_speichern(silent=True)  # Texte/Farben für die Editor-Vorschau übernehmen
        LayoutEditor(self.winfo_toplevel(), on_save=lambda: self._layout_aktiv.set(1))

    def _lade_mitarbeiter(self):
        tree = self._ma_tree
        tree.delete(*tree.get_children())
        with _con() as con:
            rows = con.execute("SELECT id, benutzer, name, email, telefon FROM tbl_faktura_mitarbeiter "
                               "WHERE aktiv=1 ORDER BY name").fetchall()
        for r in rows:
            tree.insert("", "end", iid=str(r[0]), values=(r[1] or "", r[2] or "", r[3] or "", r[4] or ""))

    def _ma_speichern(self):
        benutzer = self.ma_benutzer.get().strip()
        if not benutzer:
            messagebox.showinfo("Mitarbeiter", "Windows-Benutzer darf nicht leer sein.")
            return
        with _con() as con:
            # Upsert ueber benutzer: bestehenden aktiven Eintrag aktualisieren.
            row = con.execute("SELECT id FROM tbl_faktura_mitarbeiter WHERE benutzer=? AND aktiv=1",
                              (benutzer,)).fetchone()
            if row:
                con.execute("UPDATE tbl_faktura_mitarbeiter SET name=?, email=?, telefon=? WHERE id=?",
                            (self.ma_name.get().strip(), self.ma_email.get().strip(),
                             self.ma_tel.get().strip(), row[0]))
            else:
                con.execute("INSERT INTO tbl_faktura_mitarbeiter(benutzer, name, email, telefon) "
                            "VALUES(?,?,?,?)", (benutzer, self.ma_name.get().strip(),
                            self.ma_email.get().strip(), self.ma_tel.get().strip()))
            con.commit()
        _log("mitarbeiter_gespeichert", details=benutzer)
        self._lade_mitarbeiter()

    def _ma_del(self):
        sel = self._ma_tree.selection()
        if not sel:
            return
        with _con() as con:
            con.execute("UPDATE tbl_faktura_mitarbeiter SET aktiv=0 WHERE id=?", (int(sel[0]),))
            con.commit()
        self._lade_mitarbeiter()


class LayoutEditor(tk.Toplevel):
    """Freier Layout-Editor: alle Beleg-Blöcke per Maus auf einer A4-Fläche verschieben
    und in der Größe ziehen (Griff unten rechts). Eigene Felder (Freitext / Datenfelder)
    lassen sich hinzufügen. Speichert Position+Größe (mm) als JSON (tpl_layout) und die
    eigenen Felder (tpl_zusatzfelder)."""
    MIN_MM = 12  # kleinste Block-Kantenlänge

    def __init__(self, master, on_save=None):
        super().__init__(master)
        self.title("Layout frei gestalten")
        self.configure(bg=SHELL_BG)
        self.on_save = on_save
        self.layout = _layout_laden()
        self.zusatz = _zusatz_laden()
        self._sel = None
        self._mode = "move"
        self._px = self._py = 0
        # Skalierung so wählen, dass das GANZE Fenster in die Arbeitsfläche passt
        # (sonst öffnet es hinter der Taskleiste / zu groß).
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        self.scale = max(1.8, min(3.0, (sh - 280) / 297, (sw - 140) / 210))

        kopf = tk.Frame(self, bg=SHELL_BG)
        kopf.pack(fill="x", padx=16, pady=(12, 4))
        tk.Label(kopf, text="Ziehen = verschieben · Griff unten rechts = Größe · A4 (210 × 297 mm)",
                 bg=SHELL_BG, fg=TEXT, font=(theme.FONT, 12, "bold")).pack(side="left")

        bar = tk.Frame(self, bg=SHELL_BG)
        bar.pack(fill="x", padx=16, pady=(0, 6))
        theme.PillButton(bar, "➕  Freitext", self._add_freitext, kind="neutral",
                         font_size=10, padx=12, pady=6).pack(side="left")
        theme.PillButton(bar, "➕  Datenfeld", self._add_datenfeld, kind="neutral",
                         font_size=10, padx=12, pady=6).pack(side="left", padx=(8, 0))
        theme.PillButton(bar, "🗑  Feld löschen", self._feld_loeschen, kind="ghost",
                         font_size=10, padx=12, pady=6).pack(side="left", padx=(8, 0))
        tk.Label(bar, text="(eigenes Feld anklicken, dann löschen)", bg=SHELL_BG, fg=theme.FAINT,
                 font=theme.SMALL).pack(side="left", padx=(10, 0))

        mitte = tk.Frame(self, bg=SHELL_BG)
        mitte.pack(fill="both", expand=True, padx=16)
        w, h = int(210 * self.scale), int(297 * self.scale)
        self.canvas = tk.Canvas(mitte, width=w, height=h, bg="white",
                                highlightbackground=BORDER, highlightthickness=1)
        self.canvas.pack()
        self.boxes = {}
        self._zeichne()

        fuss = tk.Frame(self, bg=SHELL_BG)
        fuss.pack(fill="x", padx=16, pady=10)
        theme.PillButton(fuss, "💾  Speichern & aktivieren", self._speichern, kind="success",
                         font_size=10, padx=14, pady=8).pack(side="left")
        theme.PillButton(fuss, "👁  Vorschau (PDF)", self._vorschau, kind="accent",
                         font_size=10, padx=14, pady=8).pack(side="left", padx=(8, 0))
        theme.PillButton(fuss, "Standard wiederherstellen", self._standard, kind="neutral",
                         font_size=10, padx=14, pady=8).pack(side="left", padx=(8, 0))
        theme.PillButton(fuss, "Schließen", self.destroy, kind="ghost",
                         font_size=10, padx=14, pady=8).pack(side="right")

        # Fenster passend dimensionieren und oben links positionieren.
        win_w, win_h = w + 60, h + 190
        self.geometry(f"{win_w}x{min(win_h, sh - 70)}+30+15")
        self.minsize(min(win_w, sw - 40), 400)
        self.transient(master)
        self.grab_set()

    def _box(self, key, label, sample, data, custom=False):
        sc = self.scale
        x, y = data["x"] * sc, data["y"] * sc
        w = data.get("w", 60) * sc
        h = data.get("h", 12) * sc
        tag, htag = f"b_{key}", f"h_{key}"
        fill = "#FBEFD9" if custom else ACCENT_LIGHT
        outl = theme.WARNING if custom else ACCENT
        rect = self.canvas.create_rectangle(x, y, x + w, y + h, fill=fill, outline=outl,
                                            width=1, tags=(tag,))
        lbl = self.canvas.create_text(x + 5, y + 3, anchor="nw", text=label, fill=outl,
                                      font=(theme.FONT, 7, "bold"), tags=(tag,))
        txt = self.canvas.create_text(x + 5, y + 16, anchor="nw", text=sample, fill=theme.INK,
                                      font=(theme.MONO, 7), width=max(20, w - 10), tags=(tag,))
        handle = self.canvas.create_rectangle(x + w - 9, y + h - 9, x + w, y + h,
                                              fill=outl, outline=outl, tags=(htag,))
        self.canvas.tag_bind(tag, "<Button-1>", lambda e, k=key: self._press(e, k, "move"))
        self.canvas.tag_bind(tag, "<B1-Motion>", lambda e, k=key: self._drag(e, k))
        self.canvas.tag_bind(htag, "<Button-1>", lambda e, k=key: self._press(e, k, "resize"))
        self.canvas.tag_bind(htag, "<B1-Motion>", lambda e, k=key: self._drag(e, k))
        self.boxes[key] = {"rect": rect, "label": lbl, "text": txt, "handle": handle,
                           "x": x, "y": y, "w": w, "h": h, "data": data, "custom": custom}

    def _zeichne(self):
        self.canvas.delete("all")
        self.boxes = {}
        for key in LAYOUT_DEFAULT:
            self._box(key, BLOCK_LABEL.get(key, key), BLOCK_SAMPLE.get(key, ""),
                      self.layout[key], custom=False)
        for z in self.zusatz:
            label = "Datenfeld" if z.get("typ") == "feld" else "Freitext"
            self._box(f"z_{z['id']}", label, z.get("text", ""), z, custom=True)

    def _render_box(self, key):
        b = self.boxes[key]
        self.canvas.coords(b["rect"], b["x"], b["y"], b["x"] + b["w"], b["y"] + b["h"])
        self.canvas.coords(b["label"], b["x"] + 5, b["y"] + 3)
        self.canvas.coords(b["text"], b["x"] + 5, b["y"] + 16)
        self.canvas.itemconfig(b["text"], width=max(20, b["w"] - 10))
        self.canvas.coords(b["handle"], b["x"] + b["w"] - 9, b["y"] + b["h"] - 9,
                           b["x"] + b["w"], b["y"] + b["h"])

    def _press(self, e, key, mode):
        self._sel, self._mode = key, mode
        self._px, self._py = e.x, e.y
        self.canvas.tag_raise(f"b_{key}")
        self.canvas.tag_raise(f"h_{key}")

    def _drag(self, e, key):
        b = self.boxes[key]
        d = b["data"]
        sc = self.scale
        pw, ph = 210 * sc, 297 * sc
        dx, dy = e.x - self._px, e.y - self._py
        mn = self.MIN_MM * sc
        if self._mode == "resize":
            b["w"] = min(max(b["w"] + dx, mn), pw - b["x"])
            b["h"] = min(max(b["h"] + dy, mn), ph - b["y"])
            d["w"] = round(b["w"] / sc, 1)
            d["h"] = round(b["h"] / sc, 1)
        else:
            b["x"] = min(max(b["x"] + dx, 0), pw - b["w"])
            b["y"] = min(max(b["y"] + dy, 0), ph - b["h"])
            d["x"] = round(b["x"] / sc, 1)
            d["y"] = round(b["y"] / sc, 1)
        self._render_box(key)
        self._px, self._py = e.x, e.y

    def _add_feld(self, text, typ):
        nid = max([z["id"] for z in self.zusatz], default=0) + 1
        self.zusatz.append({"id": nid, "typ": typ, "text": text,
                            "x": 20, "y": 20, "w": 70, "h": 12, "size": 11})
        self._zeichne()

    def _add_freitext(self):
        txt = simpledialog.askstring("Freitext hinzufügen",
                                     "Text (Platzhalter wie {kunde_name} sind erlaubt):", parent=self)
        if txt:
            self._add_feld(txt, "freitext")

    def _add_datenfeld(self):
        dlg = tk.Toplevel(self)
        dlg.title("Datenfeld wählen")
        dlg.configure(bg=BG)
        dlg.transient(self)
        dlg.grab_set()
        tk.Label(dlg, text="Feld aus der Datenbank/Beleg:", bg=BG, fg=TEXT,
                 font=theme.BODY).pack(padx=20, pady=(16, 6))
        var = tk.StringVar(value=FELD_QUELLEN[0][0])
        ttk.Combobox(dlg, textvariable=var, width=36, state="readonly",
                     values=[l for l, _ in FELD_QUELLEN], style="NMG.TCombobox").pack(padx=20)
        res = {"t": None}

        def ok():
            res["t"] = dict(FELD_QUELLEN).get(var.get())
            dlg.destroy()
        theme.PillButton(dlg, "Einfügen", ok, kind="success", font_size=10).pack(pady=14)
        self.wait_window(dlg)
        if res["t"]:
            self._add_feld(res["t"], "feld")

    def _feld_loeschen(self):
        if not self._sel or not self._sel.startswith("z_"):
            messagebox.showinfo("Löschen", "Bitte zuerst ein EIGENES Feld anklicken. "
                                "Standardblöcke lassen sich nicht löschen.", parent=self)
            return
        zid = int(self._sel.split("_")[1])
        self.zusatz = [z for z in self.zusatz if z["id"] != zid]
        self._sel = None
        self._zeichne()

    def _speichern(self):
        set_setting("tpl_layout", json.dumps(self.layout))
        set_setting("tpl_zusatzfelder", json.dumps(self.zusatz))
        set_setting("tpl_layout_aktiv", "1")
        _log("layout_gespeichert")
        if self.on_save:
            self.on_save()
        messagebox.showinfo("Layout", "Layout gespeichert und aktiviert. Rechnung, Gutschrift "
                            "und Quartalsvergütung nutzen es ab jetzt.", parent=self)

    def _standard(self):
        self.layout = {
            k: {"x": LAYOUT_DEFAULT[k]["x"], "y": LAYOUT_DEFAULT[k]["y"],
                "w": BLOCK_BREITE_MM.get(k, 90), "h": BLOCK_HOEHE_MM.get(k, 12)} for k in LAYOUT_DEFAULT}
        self._zeichne()

    def _vorschau(self):
        set_setting("tpl_layout", json.dumps(self.layout))
        set_setting("tpl_zusatzfelder", json.dumps(self.zusatz))
        set_setting("tpl_layout_aktiv", "1")
        pfad = vorschau_pdf("rechnung")
        if pfad and os.path.exists(pfad):
            try:
                os.startfile(pfad)  # type: ignore[attr-defined]
            except Exception:
                messagebox.showinfo("Vorschau", f"Datei: {pfad}", parent=self)


def run_standalone():
    """Startet Faktura als eigenständiges Fenster (eigenes Taskleisten-Icon).
    Wird von start_faktura.py genutzt."""
    if os.name == "nt":
        try:
            import ctypes
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("NMG.Faktura")
        except Exception:
            pass
    try:
        from .migrations import run_migrations
        run_migrations()
    except Exception:
        pass
    root = tk.Tk()
    root.title("NMG Faktura")
    root.geometry("1100x720")
    root.minsize(960, 620)
    root.configure(bg=SHELL_BG)
    theme.apply_theme(root)
    theme.apply_widget_defaults(root)
    try:
        root.iconbitmap(str(ASSETS_DIR / "Faktura.ico"))
    except Exception:
        pass
    FakturaPanel(root, on_close=root.destroy).pack(fill="both", expand=True)
    tour.maybe_show(root, "faktura", tour.faktura_steps())
    root.mainloop()


if __name__ == "__main__":
    run_standalone()
