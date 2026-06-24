"""Zentrale, in der App editierbare Einstellungen der Kasse (Texte, Firmendaten,
Parameter). Speicherung als einfache Schluessel/Wert-Tabelle in der gemeinsamen
Datenbank, damit NMGone und die Kasse-.exe dieselben Werte sehen.

Genutzt von den Dokument-Vorlagen (Auftrag/Lieferschein/Defektmeldung) und vom
automatischen Tagesabschluss. Fehlt ein Wert (oder ist leer), greift der Default.
"""
from __future__ import annotations

import getpass
import html as _html
import sqlite3
from datetime import datetime

from .config import DB_PATH

# Schluessel + Standardwerte. Die Defaults entsprechen den bisher fest in den
# Vorlagen hinterlegten Texten, damit sich ohne Pflege nichts aendert.
DEFAULTS = {
    "firma_name": "Ihre Firma GmbH",
    "firma_adresse": "Musterstraße 1\n12345 Musterstadt",
    "firma_kontakt": "Tel. 01234 / 56789 · info@ihre-firma.de",
    "defekt_rechtstext": (
        "Diese Bescheinigung dient der Apotheke als Nachweis der Nichtverfügbarkeit "
        "der oben genannten Arzneimittel und kann zur Dokumentation gegenüber dem "
        "Kostenträger (Krankenkasse) verwendet werden.\n\n"
        "[Hier die rechtlich geprüfte Formulierung und die einschlägigen gesetzlichen "
        "Grundlagen eintragen.]"
    ),
    "tagesabschluss_stunde": "18",
}


def _ensure(con):
    con.execute(
        "CREATE TABLE IF NOT EXISTS tbl_kasse_einstellungen("
        "schluessel TEXT PRIMARY KEY, wert TEXT, geaendert_am TEXT, bearbeiter TEXT)")


def get(db_path=DB_PATH, key="", default=None) -> str:
    """Liefert den gespeicherten Wert. Ist nichts gespeichert ODER der Wert leer,
    greift der uebergebene default bzw. der DEFAULTS-Eintrag."""
    fallback = default if default is not None else DEFAULTS.get(key, "")
    try:
        with sqlite3.connect(db_path) as con:
            _ensure(con)
            row = con.execute(
                "SELECT wert FROM tbl_kasse_einstellungen WHERE schluessel=?", (key,)).fetchone()
    except sqlite3.Error:
        return fallback
    if row and row[0] is not None and str(row[0]).strip() != "":
        return str(row[0])
    return fallback


def get_int(db_path=DB_PATH, key="", default=0) -> int:
    try:
        return int(str(get(db_path, key, str(default))).strip())
    except (ValueError, TypeError):
        return default


def set_(db_path, key, value):
    """Speichert einen Wert (Upsert)."""
    with sqlite3.connect(db_path) as con:
        _ensure(con)
        con.execute(
            "INSERT INTO tbl_kasse_einstellungen(schluessel, wert, geaendert_am, bearbeiter) "
            "VALUES(?,?,?,?) ON CONFLICT(schluessel) DO UPDATE SET "
            "wert=excluded.wert, geaendert_am=excluded.geaendert_am, bearbeiter=excluded.bearbeiter",
            (key, value, datetime.now().isoformat(timespec="seconds"), getpass.getuser()))
        con.commit()


def set_many(db_path, werte: dict):
    for k, v in werte.items():
        set_(db_path, k, v)


def firma_felder(db_path=DB_PATH) -> dict:
    """Firmendaten fuer den Dokumentkopf (Auftrag/Lieferschein/Defektmeldung)."""
    return {
        "firma_name": get(db_path, "firma_name"),
        "firma_adresse": get(db_path, "firma_adresse"),
        "firma_kontakt": get(db_path, "firma_kontakt"),
    }


def _mehrzeilig_html(text: str) -> str:
    """Mehrzeiligen Text HTML-sicher machen (escapen) und Zeilenumbrueche -> <br>."""
    return _html.escape(str(text or "")).replace("\n", "<br>")


def absender_adresse_html(db_path=DB_PATH) -> str:
    return _mehrzeilig_html(get(db_path, "firma_adresse"))


def rechtstext_html(db_path=DB_PATH) -> str:
    return _mehrzeilig_html(get(db_path, "defekt_rechtstext"))
