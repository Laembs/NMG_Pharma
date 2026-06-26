__version__ = "0.9.0"

# Mehrsprachigkeit zentral aktivieren: Sprache laden + Tkinter-Auto-Uebersetzung
# installieren, bevor irgendeine App ihre Widgets baut. Bei DE (Standard) ist
# das ein No-Op, daher fuer bestehende Nutzer ohne Effekt. Bewusst tolerant:
# in Nicht-GUI-Kontexten (Skripte ohne Tkinter) faellt es still durch.
try:  # pragma: no cover - reine Verdrahtung
    from . import i18n as _i18n
    _i18n.load_language()
    _i18n.install_auto_translation()
except Exception:
    pass
