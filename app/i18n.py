"""SP11: Einfache Dict-basierte Mehrsprachigkeit fuer NMGone.

Warum nicht gettext?
- gettext braucht .po/.mo-Dateien + Build-Schritt + Locale-Setup
- Wir haben bisher ~9.000 Zeilen GUI-Code, fast alles deutsche Strings
- Dict-basiert ist 10x schneller eingebaut und reicht voellig fuer den
  ersten Aufschlag. Wenn das Volumen waechst, koennen wir spaeter zu
  gettext migrieren.

Wie es funktioniert:
- TRANSLATIONS-Dict mappt deutschen Quelltext -> {sprache: uebersetzung}
- T("Neue Auswertung") liefert die Uebersetzung in der aktiv gesetzten
  Sprache zurueck. Fehlt eine Uebersetzung: Fallback auf DE, dann auf
  den Key selbst.
- Sprache wird in USERDATA_ROOT/language.json gespeichert -> persistent
  ueber Programmstarts hinweg, vor dem Splash bereits lesbar (kein DB-
  Zugriff noetig).
- Aenderung der Sprache wirkt erst beim naechsten Start (kein Live-
  Switch im UI). Reicht fuer SP11.
"""
from __future__ import annotations
import json
from pathlib import Path

from .config import USERDATA_ROOT, BASE_DIR

LANGUAGES = {
    "DE": "Deutsch",
    "EN": "English",
    "SK": "Slovenčina",
    "CZ": "Čeština",
}
DEFAULT_LANGUAGE = "DE"

_current_language = DEFAULT_LANGUAGE


def _settings_path() -> Path:
    base = USERDATA_ROOT if USERDATA_ROOT else BASE_DIR
    return base / "language.json"


def load_language() -> str:
    global _current_language
    try:
        p = _settings_path()
        if p.exists():
            data = json.loads(p.read_text(encoding="utf-8"))
            lang = str(data.get("language", DEFAULT_LANGUAGE)).strip().upper()
            if lang in LANGUAGES:
                _current_language = lang
    except Exception:
        pass
    return _current_language


def save_language(lang: str) -> None:
    global _current_language
    lang = str(lang).strip().upper()
    if lang not in LANGUAGES:
        return
    _current_language = lang
    try:
        p = _settings_path()
        p.parent.mkdir(parents=True, exist_ok=True)
        p.write_text(json.dumps({"language": lang}, ensure_ascii=False), encoding="utf-8")
    except Exception:
        pass


def get_language() -> str:
    return _current_language


# ── Uebersetzungen ─────────────────────────────────────────────────────────
# SP11 deckt die Top-Strings ab: Navigation, Splash, Hauptbuttons,
# Update-Dialoge. Der Rest bleibt erstmal deutsch und wird Schritt fuer
# Schritt nachgezogen, wenn echte SK/EN-Nutzer Stellen melden.
TRANSLATIONS: dict[str, dict[str, str]] = {
    "Sprache auswählen": {
        "DE": "Sprache auswählen",
        "EN": "Select language",
        "SK": "Vyberte jazyk",
        "CZ": "Vyberte jazyk",
    },
    "Starten": {
        "DE": "Starten",
        "EN": "Start",
        "SK": "Spustiť",
        "CZ": "Spustit",
    },
    "Willkommen bei NMGone": {
        "DE": "Willkommen bei NMGone",
        "EN": "Welcome to NMGone",
        "SK": "Vitajte v NMGone",
        "CZ": "Vítejte v NMGone",
    },
    "Startseite": {
        "DE": "Startseite",
        "EN": "Home",
        "SK": "Domov",
        "CZ": "Domů",
    },
    "Neue Auswertung": {
        "DE": "Bedarfsanalyse",
        "EN": "Demand analysis",
        "SK": "Analýza dopytu",
        "CZ": "Analýza poptávky",
    },
    "Apps": {
        "DE": "Apps",
        "EN": "Apps",
        "SK": "Aplikácie",
        "CZ": "Aplikace",
    },
    "Analysen": {
        "DE": "Analysen",
        "EN": "Analyses",
        "SK": "Analýzy",
        "CZ": "Analýzy",
    },
    "Schulbank": {
        "DE": "Schulbank",
        "EN": "Learning",
        "SK": "Učenie",
        "CZ": "Učení",
    },
    "Daten aktualisieren": {
        "DE": "Daten aktualisieren",
        "EN": "Update data",
        "SK": "Aktualizovať údaje",
        "CZ": "Aktualizovat data",
    },
    "Update / Backup": {
        "DE": "Update / Backup",
        "EN": "Update / Backup",
        "SK": "Aktualizácia / Záloha",
        "CZ": "Aktualizace / Záloha",
    },
    "Datenbankübersicht": {
        "DE": "Datenbankübersicht",
        "EN": "Database overview",
        "SK": "Prehľad databázy",
        "CZ": "Přehled databáze",
    },
    "Report": {
        "DE": "Report",
        "EN": "Report",
        "SK": "Správa",
        "CZ": "Zpráva",
    },
    "Roadmap": {
        "DE": "Roadmap",
        "EN": "Roadmap",
        "SK": "Plán",
        "CZ": "Plán",
    },
    "Cloud / DB-Pfad": {
        "DE": "Cloud / DB-Pfad",
        "EN": "Cloud / DB path",
        "SK": "Cloud / cesta DB",
        "CZ": "Cloud / cesta DB",
    },
    "Hilfe": {
        "DE": "Hilfe",
        "EN": "Help",
        "SK": "Pomoc",
        "CZ": "Nápověda",
    },
    "Kunden": {
        "DE": "Kunden",
        "EN": "Customers",
        "SK": "Zákazníci",
        "CZ": "Zákazníci",
    },
    "Bestellung": {
        "DE": "Bestellung",
        "EN": "Orders",
        "SK": "Objednávky",
        "CZ": "Objednávky",
    },
    "ToDo": {
        "DE": "ToDo",
        "EN": "To-do",
        "SK": "Úlohy",
        "CZ": "Úkoly",
    },
    "Mitarbeiter": {
        "DE": "Mitarbeiter",
        "EN": "Staff",
        "SK": "Zamestnanci",
        "CZ": "Zaměstnanci",
    },
    "Produktanalyse": {
        "DE": "Produktanalyse",
        "EN": "Product analysis",
        "SK": "Analýza produktov",
        "CZ": "Analýza produktů",
    },
    "Abweichungsanalyse": {
        "DE": "Abweichungsanalyse",
        "EN": "Variance analysis",
        "SK": "Analýza odchýlok",
        "CZ": "Analýza odchylek",
    },
    "Speichern": {
        "DE": "Speichern",
        "EN": "Save",
        "SK": "Uložiť",
        "CZ": "Uložit",
    },
    "Abbrechen": {
        "DE": "Abbrechen",
        "EN": "Cancel",
        "SK": "Zrušiť",
        "CZ": "Zrušit",
    },
    "Importieren": {
        "DE": "Importieren",
        "EN": "Import",
        "SK": "Importovať",
        "CZ": "Importovat",
    },
    "Öffnen": {
        "DE": "Öffnen",
        "EN": "Open",
        "SK": "Otvoriť",
        "CZ": "Otevřít",
    },
    "Schließen": {
        "DE": "Schließen",
        "EN": "Close",
        "SK": "Zavrieť",
        "CZ": "Zavřít",
    },
    "Aktualisieren": {
        "DE": "Aktualisieren",
        "EN": "Refresh",
        "SK": "Obnoviť",
        "CZ": "Obnovit",
    },
    "Update suchen": {
        "DE": "Update suchen",
        "EN": "Check for update",
        "SK": "Skontrolovať aktualizáciu",
        "CZ": "Zkontrolovat aktualizaci",
    },
    "Update installieren": {
        "DE": "Update installieren",
        "EN": "Install update",
        "SK": "Inštalovať aktualizáciu",
        "CZ": "Instalovat aktualizaci",
    },
}


def T(key: str, **kwargs) -> str:
    """Translate. Fallback: aktive Sprache -> DE -> Key selbst.
    Optional kwargs werden in .format(...) gefuettert.
    """
    translations = TRANSLATIONS.get(key, {})
    text = translations.get(_current_language) or translations.get("DE") or key
    if kwargs:
        try:
            return text.format(**kwargs)
        except Exception:
            return text
    return text
