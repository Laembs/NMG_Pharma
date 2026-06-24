"""Defektmeldung / Nichtverfuegbarkeitsbescheinigung erzeugen (HTML).

Wenn ein Artikel nicht vorraetig/lieferbar ist, kann der Apotheke ein Dokument
ausgestellt werden, das sie zur Dokumentation der Nichtverfuegbarkeit (z.B. zur
Vorlage beim Kostentraeger) nutzen kann.

WICHTIG: Die konkrete rechtliche Formulierung / Paragraphen-Verweise sind in der
Vorlage bewusst als anpassbarer Platzhalter gehalten - sie muessen fachlich/
rechtlich geprueft und in vorlagen/defektmeldung.html eingetragen werden. Das
Modul fuellt nur die Datenfelder (Apotheke, Artikelliste, Datum, Grund).

Anders als Auftrag/Lieferschein liest dieses Modul KEINEN gespeicherten Verkauf,
sondern erhaelt die Positionen direkt aus der GUI (ad-hoc-Dokument).
"""
from __future__ import annotations

import html as _html
import shutil
import sqlite3
from datetime import datetime
from pathlib import Path

from .config import ASSETS_DIR, OUTPUT_DIR, USERDATA_ROOT, BASE_DIR, DB_PATH
from . import einstellungen

_ROOT = USERDATA_ROOT or BASE_DIR
VORLAGEN_DIR = _ROOT / "vorlagen"
TEMPLATE_NAME = "defektmeldung.html"


def template_path() -> Path:
    """Pfad zur (nutzer-editierbaren) Defektmeldung-Vorlage; legt sie beim ersten
    Mal aus der mitgelieferten Standardvorlage an."""
    VORLAGEN_DIR.mkdir(parents=True, exist_ok=True)
    p = VORLAGEN_DIR / TEMPLATE_NAME
    if not p.exists():
        default = ASSETS_DIR / "defektmeldung_vorlage.html"
        if default.exists():
            shutil.copy2(default, p)
    return p


def _kunde_laden(db_path, knr) -> dict:
    if not knr:
        return {}
    with sqlite3.connect(db_path) as con:
        exists = con.execute(
            "SELECT 1 FROM sqlite_master WHERE type='table' AND name='tbl_kunden_center'").fetchone()
        if not exists:
            return {}
        have = {r[1] for r in con.execute("PRAGMA table_info(tbl_kunden_center)")}
        want = [c for c in ("kundenname", "plz", "ort", "strasse", "inhaber", "email") if c in have]
        if not want:
            return {}
        row = con.execute(
            f"SELECT {','.join(want)} FROM tbl_kunden_center WHERE kundennummer=? LIMIT 1",
            (knr,)).fetchone()
        return dict(zip(want, row)) if row else {}


def render(db_path=DB_PATH, knr=None, kunde=None, positionen=None, grund="",
           datum=None) -> Path:
    """Erzeugt die Defektmeldung als HTML-Datei und gibt den Pfad zurueck.
    positionen = [{pzn, artikelname, menge}, ...]."""
    positionen = positionen or []
    if kunde is None:
        kunde = _kunde_laden(db_path, knr)
    tpl = template_path()
    text = tpl.read_text(encoding="utf-8")

    rows = []
    for i, p in enumerate(positionen, 1):
        rows.append(
            "<tr><td>{pos}</td><td>{pzn}</td><td>{art}</td><td class='r'>{menge}</td></tr>".format(
                pos=i, pzn=_html.escape(str(p.get("pzn") or "")),
                art=_html.escape(str(p.get("artikelname") or "")),
                menge=p.get("menge") or ""))

    apotheke = kunde.get("kundenname") or ""
    firma = einstellungen.firma_felder(db_path)
    repl = {
        "datum": (datum or datetime.now().strftime("%d.%m.%Y")),
        "kundennummer": knr or "",
        "apotheke": apotheke,
        "inhaber": kunde.get("inhaber") or "",
        "strasse": kunde.get("strasse") or "",
        "plz": kunde.get("plz") or "",
        "ort": kunde.get("ort") or "",
        "grund": grund or "nicht lieferbar",
        "firma": firma["firma_name"],
        "absender_kontakt": firma["firma_kontakt"],
    }
    for k, v in repl.items():
        text = text.replace("{{%s}}" % k, _html.escape(str(v)))
    # Mehrzeilige/HTML-Felder roh ersetzen (nach dem Escape-Durchlauf).
    text = text.replace("{{absender_adresse}}", einstellungen.absender_adresse_html(db_path))
    text = text.replace("{{rechtstext}}", einstellungen.rechtstext_html(db_path))
    text = text.replace("{{positionen}}", "".join(rows))

    out_dir = OUTPUT_DIR / "defektmeldungen"
    out_dir.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    safe = "".join(c for c in apotheke if c.isalnum() or c in " _-").strip().replace(" ", "_")[:40]
    path = out_dir / f"Defektmeldung_{safe or 'kunde'}_{stamp}.html"
    path.write_text(text, encoding="utf-8")
    return path
