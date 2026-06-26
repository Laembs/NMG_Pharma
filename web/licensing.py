"""Modul-/Lizenz-Freischaltung.

Welche Module eine Firma sehen darf, steht in der Plattform-Tabelle
``firma_modul``. ``MODULE_CATALOG`` ist der zentrale Katalog aller Module, die
das Produkt grundsätzlich anbietet (egal ob lizenziert). Beim Pilot ist nur
``personal`` tatsächlich umgesetzt; die übrigen sind als „kommt noch" gelistet,
damit das Dashboard/Lizenzmodell schon vollständig aussieht.
"""
from __future__ import annotations

from fastapi import HTTPException, Request

from .tenancy import platform_con

# key, Anzeige-Titel, Kurzbeschreibung, ist im Pilot lauffähig?
MODULE_CATALOG: list[dict] = [
    {"key": "personal",  "titel": "Personal",  "desc": "Mitarbeiter, Urlaub, Zeit & Akte",   "ready": True},
    {"key": "faktura",   "titel": "Faktura",   "desc": "Rechnungen & Gutschriften",          "ready": False},
    {"key": "kasse",     "titel": "Kasse",     "desc": "Verkauf & Wareneingang",             "ready": True},
    {"key": "einkauf",   "titel": "Einkauf",   "desc": "Beschaffung EU-Ausland",             "ready": False},
    {"key": "gdp",       "titel": "GDP",       "desc": "Wareneingang & Retouren",            "ready": False},
    {"key": "meldungen", "titel": "Meldungen", "desc": "GDP-Meldewesen & Kühlkette",         "ready": False},
]
_CATALOG_BY_KEY = {m["key"]: m for m in MODULE_CATALOG}


def lizenzierte_module(firma_id: int) -> set[str]:
    """Liefert die Keys der aktuell aktiv lizenzierten Module einer Firma."""
    with platform_con() as con:
        rows = con.execute(
            """SELECT modul_key FROM firma_modul
               WHERE firma_id=? AND aktiv=1
                 AND (gueltig_bis IS NULL OR gueltig_bis >= date('now'))""",
            (firma_id,)).fetchall()
    return {r["modul_key"] for r in rows}


def module_fuer_dashboard(firma_id: int) -> list[dict]:
    """Katalog-Einträge der lizenzierten Module (für die Dashboard-Kacheln)."""
    aktiv = lizenzierte_module(firma_id)
    return [_CATALOG_BY_KEY[k] for k in (m["key"] for m in MODULE_CATALOG) if k in aktiv]


def require_module(modul_key: str):
    """FastAPI-Dependency: sperrt eine Route, wenn das Modul nicht lizenziert ist.

    Verwendung:  @router.get(..., dependencies=[Depends(require_module("personal"))])
    """
    def _dep(request: Request) -> None:
        user = getattr(request.state, "user", None)
        if not user:
            raise HTTPException(status_code=401, detail="Nicht angemeldet")
        if modul_key not in lizenzierte_module(user["firma_id"]):
            raise HTTPException(
                status_code=403,
                detail=f"Modul '{modul_key}' ist fuer Ihre Firma nicht freigeschaltet.")
    return _dep
