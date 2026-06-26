"""Kasse-Modul · Web-Routen (durch Lizenz-Gate geschützt).

P0 (Gerüst): Übersicht + Lagerbestand-Ansicht + Verkaufs-Platzhalter. Der
eigentliche Verkaufs-Flow (Kunde → Artikel → Charge → Rabatt → Position →
Abschluss) folgt in P2. Jede Route stellt zuerst das Tenant-Schema sicher,
damit das Modul ohne separaten Migrationslauf läuft.
"""
from __future__ import annotations

from fastapi import APIRouter, Depends, Request

from ..licensing import require_module
from ..services import kasse_service as svc
from ..templating import page, tenant_for

# Jede Route verlangt das lizenzierte Modul "kasse".
router = APIRouter(
    prefix="/kasse",
    dependencies=[Depends(require_module("kasse"))],
)


@router.get("")
def uebersicht(request: Request):
    con = tenant_for(request)
    try:
        svc.ensure_schema(con)
        daten = svc.uebersicht(con)
    finally:
        con.close()
    return page(request, "kasse/uebersicht.html", daten=daten, titel="Kasse")


@router.get("/lager")
def lager(request: Request):
    con = tenant_for(request)
    try:
        svc.ensure_schema(con)
        bestand = svc.lagerbestand(con)
    finally:
        con.close()
    return page(request, "kasse/lager.html", bestand=bestand, titel="Lagerbestand")


@router.get("/verkauf")
def verkauf(request: Request):
    # Platzhalter: der Verkaufs-Flow ist die nächste Ausbaustufe (P2).
    con = tenant_for(request)
    try:
        svc.ensure_schema(con)
    finally:
        con.close()
    return page(request, "kasse/verkauf.html", titel="Verkauf")
