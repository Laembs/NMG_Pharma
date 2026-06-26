"""Kasse-Modul · Web-Routen (durch Lizenz-Gate geschützt).

P0 (Gerüst): Übersicht + Lagerbestand-Ansicht + Verkaufs-Platzhalter. Der
eigentliche Verkaufs-Flow (Kunde → Artikel → Charge → Rabatt → Position →
Abschluss) folgt in P2. Jede Route stellt zuerst das Tenant-Schema sicher,
damit das Modul ohne separaten Migrationslauf läuft.
"""
from __future__ import annotations

from fastapi import APIRouter, Depends, Form, Request
from fastapi.responses import RedirectResponse

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
    """Mobiler Verkaufs-Flow: Artikel suchen -> Warenkorb -> Apotheke -> speichern."""
    con = tenant_for(request)
    try:
        svc.ensure_schema(con)
        artikel = svc.artikel_mit_bestand(con)
        kunden = svc.kunden_alle(con)
    finally:
        con.close()
    return page(request, "kasse/verkauf.html", titel="Schnellverkauf",
                artikel=artikel, kunden=kunden)


@router.post("/verkauf")
def verkauf_speichern(request: Request,
                      kunde_id: str = Form(""),
                      kunde_name: str = Form(""),
                      pzn: list[str] = Form(default=[]),
                      bezeichnung: list[str] = Form(default=[]),
                      menge: list[str] = Form(default=[])):
    """Speichert den Auftrag (mit Bestands-Split) und kehrt zur Übersicht zurück."""
    user = request.session["user"]
    positionen = list(zip(pzn, bezeichnung, menge))
    con = tenant_for(request)
    try:
        svc.ensure_schema(con)
        res = svc.verkauf_speichern_v2(
            con, kunde_id, kunde_name, positionen,
            user.get("anzeigename") or user.get("login") or "?")
    finally:
        con.close()
    if not res:
        return RedirectResponse("/kasse/verkauf?fehler=1", status_code=303)
    return RedirectResponse(
        f"/kasse?ok={res['best_id']}&vor={res['vorbestellt_stueck']}", status_code=303)
