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


# ── Laufzeit-Woerterbuch (grosse Volluebersetzung) ──────────────────────────
# Waehrend TRANSLATIONS oben die handgepflegten Spezialfaelle (z.B. Navigation,
# bei denen der DE-Anzeigetext vom Key abweicht) haelt, sammelt _RUNTIME die
# automatisch extrahierten UI-Strings je Sprache. Quelle: app/translations_*.py.
# Schluessel ist immer der DEUTSCHE Anzeigetext (= das, was im Code-Literal
# steht und zur Laufzeit im Widget landet), damit die Auto-Uebersetzung greift.
_RUNTIME: dict[str, dict[str, str]] = {}


def register_translations(mapping: dict, lang: str) -> None:
    """Registriert {deutscher_text: uebersetzung} fuer eine Sprache."""
    lang = str(lang).strip().upper()
    for de, tr in mapping.items():
        if de and tr:
            _RUNTIME.setdefault(de, {})[lang] = tr


def translate(text):
    """Zentrale Uebersetzungsfunktion fuer die Auto-Patch-Schicht.

    Bei DE (Standard) wird der Text unveraendert zurueckgegeben -> kein
    Verhalten/Performance-Einfluss fuer bestehende Nutzer. Sonst: Lookup in
    _RUNTIME, dann im handgepflegten TRANSLATIONS, sonst Fallback DE-Text.
    """
    if _current_language == "DE":
        return text
    if not isinstance(text, str) or not text:
        return text
    entry = _RUNTIME.get(text)
    if entry:
        return entry.get(_current_language) or text
    entry = TRANSLATIONS.get(text)
    if entry:
        return entry.get(_current_language) or entry.get("DE") or text
    return text


def T(key: str, **kwargs) -> str:
    """Translate. Fallback: aktive Sprache -> DE -> Key selbst.
    Optional kwargs werden in .format(...) gefuettert.
    """
    translations = TRANSLATIONS.get(key)
    if translations:
        text = translations.get(_current_language) or translations.get("DE") or key
    else:
        text = translate(key)
    if kwargs:
        try:
            return text.format(**kwargs)
        except Exception:
            return text
    return text


# ── Auto-Uebersetzung: Tkinter zentral patchen ──────────────────────────────
# Statt ~1.500 Aufrufstellen einzeln in T(...) zu wrappen, haengen wir uns an
# die wenigen Tkinter-Einstiegspunkte (text=/title/Menue/Tab/Heading/messagebox)
# und schicken jeden String durch translate(). Fuer DE ist das ein No-Op, daher
# fuer den Bestand risikolos; fehlende Eintraege fallen auf Deutsch zurueck.
_auto_installed = False


def install_auto_translation() -> None:
    global _auto_installed
    if _auto_installed:
        return
    try:
        import tkinter as tk
        from tkinter import ttk
    except Exception:
        return
    _auto_installed = True

    def _patch_text_widget(cls):
        orig_init = cls.__init__

        def __init__(self, *a, **kw):
            if "text" in kw:
                kw["text"] = translate(kw["text"])
            return orig_init(self, *a, **kw)

        cls.__init__ = __init__
        orig_cfg = cls.configure

        def configure(self, cnf=None, **kw):
            if isinstance(cnf, dict) and "text" in cnf:
                cnf = dict(cnf)
                cnf["text"] = translate(cnf["text"])
            if "text" in kw:
                kw["text"] = translate(kw["text"])
            return orig_cfg(self, cnf, **kw)

        cls.configure = configure
        cls.config = configure

    for cls in (tk.Label, tk.Button, tk.Checkbutton, tk.Radiobutton,
                tk.Menubutton, tk.LabelFrame,
                ttk.Label, ttk.Button, ttk.Checkbutton, ttk.Radiobutton,
                ttk.Menubutton, ttk.Labelframe):
        try:
            _patch_text_widget(cls)
        except Exception:
            pass

    # Fenstertitel
    try:
        orig_title = tk.Wm.wm_title

        def wm_title(self, string=None):
            if isinstance(string, str):
                string = translate(string)
            return orig_title(self, string)

        tk.Wm.wm_title = wm_title
        tk.Wm.title = wm_title
    except Exception:
        pass

    # Menue-Eintraege
    try:
        orig_add = tk.Menu.add

        def menu_add(self, itemType, cnf={}, **kw):
            if "label" in kw:
                kw["label"] = translate(kw["label"])
            if isinstance(cnf, dict) and "label" in cnf:
                cnf = dict(cnf)
                cnf["label"] = translate(cnf["label"])
            return orig_add(self, itemType, cnf, **kw)

        tk.Menu.add = menu_add
    except Exception:
        pass

    # Notebook-Tabs
    try:
        orig_nb_add = ttk.Notebook.add

        def nb_add(self, child, **kw):
            if "text" in kw:
                kw["text"] = translate(kw["text"])
            return orig_nb_add(self, child, **kw)

        ttk.Notebook.add = nb_add
        orig_nb_ins = ttk.Notebook.insert

        def nb_ins(self, pos, child, **kw):
            if "text" in kw:
                kw["text"] = translate(kw["text"])
            return orig_nb_ins(self, pos, child, **kw)

        ttk.Notebook.insert = nb_ins
        orig_nb_tab = ttk.Notebook.tab

        def nb_tab(self, tab_id, option=None, **kw):
            if "text" in kw:
                kw["text"] = translate(kw["text"])
            return orig_nb_tab(self, tab_id, option, **kw)

        ttk.Notebook.tab = nb_tab
    except Exception:
        pass

    # Treeview-Spaltenkoepfe
    try:
        orig_head = ttk.Treeview.heading

        def tv_head(self, column, option=None, **kw):
            if "text" in kw:
                kw["text"] = translate(kw["text"])
            return orig_head(self, column, option, **kw)

        ttk.Treeview.heading = tv_head
    except Exception:
        pass

    # messagebox / simpledialog
    try:
        from tkinter import messagebox as _mb

        def _wrap_mb(orig):
            def f(title=None, message=None, **kw):
                if isinstance(title, str):
                    title = translate(title)
                if isinstance(message, str):
                    message = translate(message)
                return orig(title, message, **kw)
            return f

        for _fn in ("showinfo", "showwarning", "showerror", "askquestion",
                    "askokcancel", "askyesno", "askyesnocancel", "askretrycancel"):
            _o = getattr(_mb, _fn, None)
            if _o:
                setattr(_mb, _fn, _wrap_mb(_o))
    except Exception:
        pass

    try:
        from tkinter import simpledialog as _sd

        def _wrap_sd(orig):
            def f(title, prompt, **kw):
                if isinstance(title, str):
                    title = translate(title)
                if isinstance(prompt, str):
                    prompt = translate(prompt)
                return orig(title, prompt, **kw)
            return f

        for _fn in ("askstring", "askinteger", "askfloat"):
            _o = getattr(_sd, _fn, None)
            if _o:
                setattr(_sd, _fn, _wrap_sd(_o))
    except Exception:
        pass


def _load_external_translations() -> None:
    """Laedt die grossen, automatisch gepflegten Woerterbuecher (falls vorhanden)."""
    for mod_name, lang in (("translations_sk", "SK"),
                           ("translations_en", "EN"),
                           ("translations_cz", "CZ")):
        try:
            mod = __import__(f"{__package__}.{mod_name}", fromlist=["DATA"])
            data = getattr(mod, "DATA", None) or getattr(mod, mod_name.split("_")[1].upper(), None)
            if isinstance(data, dict):
                register_translations(data, lang)
        except Exception:
            pass


_load_external_translations()
