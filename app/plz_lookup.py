"""Offline-PLZ-Pruefung und Ort-Autofill.

Liest die gebuendelte Tabelle ``assets/plz_orte.csv`` (PLZ;Ort, UTF-8),
ein Eintrag je PLZ (bei Mehrfachorten die haeufigste Stadt). Quelle:
zauberware/postal-codes-json-xml-csv (Geonames-Stand). Funktioniert ohne
Internet; im gefrorenen Build liegt die CSV im gebuendelten assets-Ordner.
"""
from __future__ import annotations

import csv
import sys
from functools import lru_cache
from pathlib import Path


def _csv_path() -> Path:
    from .config import ASSETS_DIR
    meipass = Path(getattr(sys, "_MEIPASS", ASSETS_DIR.parent))
    for cand in (ASSETS_DIR / "plz_orte.csv", meipass / "assets" / "plz_orte.csv"):
        if cand.exists():
            return cand
    return ASSETS_DIR / "plz_orte.csv"


@lru_cache(maxsize=1)
def _table() -> dict[str, str]:
    table: dict[str, str] = {}
    try:
        with open(_csv_path(), encoding="utf-8", newline="") as f:
            reader = csv.reader(f, delimiter=";")
            next(reader, None)  # Kopfzeile
            for row in reader:
                if len(row) >= 2 and row[0].strip():
                    table.setdefault(row[0].strip(), row[1].strip())
    except OSError:
        pass
    return table


def has_data() -> bool:
    """True, wenn die PLZ-Tabelle geladen werden konnte."""
    return bool(_table())


def is_valid_plz(plz: str | None) -> bool:
    """True, wenn die PLZ 5-stellig numerisch ist UND in der Tabelle existiert."""
    p = (plz or "").strip()
    return len(p) == 5 and p.isdigit() and p in _table()


def lookup_ort(plz: str | None) -> str | None:
    """Liefert den Ort zur PLZ oder None, wenn unbekannt."""
    return _table().get((plz or "").strip())
