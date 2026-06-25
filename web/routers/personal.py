"""Personal-Modul · Web-Routen (durch Lizenz-Gate geschützt)."""
from __future__ import annotations

from fastapi import APIRouter, Depends, Form, Request
from fastapi.responses import RedirectResponse

from ..licensing import require_module
from ..services import personal_service as svc
from ..templating import page, tenant_for

# Jede Route in diesem Router verlangt das lizenzierte Modul "personal".
router = APIRouter(
    prefix="/personal",
    dependencies=[Depends(require_module("personal"))],
)


@router.get("")
def liste(request: Request):
    con = tenant_for(request)
    try:
        mitarbeiter = svc.list_mitarbeiter(con)
        urlaub = svc.urlaubsuebersicht(con)
    finally:
        con.close()
    return page(request, "personal/list.html",
                mitarbeiter=mitarbeiter, urlaub=urlaub, titel="Personal")


@router.get("/neu")
def neu_form(request: Request):
    return page(request, "personal/form.html", ma=None, titel="Neuer Mitarbeiter")


@router.post("/neu")
def neu_speichern(request: Request,
                  vorname: str = Form(...), name: str = Form(...),
                  abteilung: str = Form(""), position: str = Form(""),
                  urlaubsanspruch: int = Form(30),
                  personalverantwortlich: bool = Form(False)):
    con = tenant_for(request)
    try:
        svc.create_mitarbeiter(
            con, vorname=vorname, name=name, abteilung=abteilung, position=position,
            urlaubsanspruch=urlaubsanspruch, personalverantwortlich=personalverantwortlich)
    finally:
        con.close()
    return RedirectResponse("/personal", status_code=303)


@router.get("/{mid}/bearbeiten")
def bearbeiten_form(request: Request, mid: int):
    con = tenant_for(request)
    try:
        ma = svc.get_mitarbeiter(con, mid)
    finally:
        con.close()
    return page(request, "personal/form.html", ma=ma, titel="Mitarbeiter bearbeiten")


@router.post("/{mid}/bearbeiten")
def bearbeiten_speichern(request: Request, mid: int,
                         vorname: str = Form(...), name: str = Form(...),
                         abteilung: str = Form(""), position: str = Form(""),
                         urlaubsanspruch: int = Form(30),
                         personalverantwortlich: bool = Form(False)):
    con = tenant_for(request)
    try:
        svc.update_mitarbeiter(
            con, mid, vorname=vorname, name=name, abteilung=abteilung, position=position,
            urlaubsanspruch=urlaubsanspruch, personalverantwortlich=personalverantwortlich)
    finally:
        con.close()
    return RedirectResponse("/personal", status_code=303)


@router.post("/{mid}/loeschen")
def loeschen(request: Request, mid: int):
    con = tenant_for(request)
    try:
        svc.delete_mitarbeiter(con, mid)
    finally:
        con.close()
    return RedirectResponse("/personal", status_code=303)


# ── Abwesenheiten ────────────────────────────────────────────────────────────
@router.get("/abwesenheiten")
def abwesenheiten(request: Request):
    con = tenant_for(request)
    try:
        eintraege = svc.list_abwesenheiten(con)
        mitarbeiter = svc.list_mitarbeiter(con)
    finally:
        con.close()
    return page(request, "personal/abwesenheiten.html",
                eintraege=eintraege, mitarbeiter=mitarbeiter,
                arten=svc.ABW_ARTEN, titel="Abwesenheiten")


@router.post("/abwesenheiten")
def abwesenheit_speichern(request: Request,
                          mitarbeiter_id: int = Form(...), art: str = Form(...),
                          von: str = Form(...), bis: str = Form(...),
                          notiz: str = Form("")):
    con = tenant_for(request)
    try:
        svc.create_abwesenheit(
            con, mitarbeiter_id=mitarbeiter_id, art=art, von=von, bis=bis, notiz=notiz)
    finally:
        con.close()
    return RedirectResponse("/personal/abwesenheiten", status_code=303)


@router.post("/abwesenheiten/{aid}/loeschen")
def abwesenheit_loeschen(request: Request, aid: int):
    con = tenant_for(request)
    try:
        svc.delete_abwesenheit(con, aid)
    finally:
        con.close()
    return RedirectResponse("/personal/abwesenheiten", status_code=303)
