"""Personal-Modul · Web-Routen (durch Lizenz-Gate geschützt)."""
from __future__ import annotations

import re
import unicodedata
from datetime import date
from pathlib import Path

from fastapi import APIRouter, Depends, File, Form, HTTPException, Request, UploadFile
from fastapi.responses import FileResponse, RedirectResponse

from ..licensing import require_module
from ..services import personal_service as svc
from ..templating import page, tenant_for
from ..tenancy import tenant_dir

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
                  urlaubsanspruch: int = Form(30), sollstunden_tag: float = Form(8.0),
                  personalverantwortlich: bool = Form(False)):
    con = tenant_for(request)
    try:
        svc.create_mitarbeiter(
            con, vorname=vorname, name=name, abteilung=abteilung, position=position,
            urlaubsanspruch=urlaubsanspruch, sollstunden_tag=sollstunden_tag,
            personalverantwortlich=personalverantwortlich)
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
                         urlaubsanspruch: int = Form(30), sollstunden_tag: float = Form(8.0),
                         personalverantwortlich: bool = Form(False)):
    con = tenant_for(request)
    try:
        svc.update_mitarbeiter(
            con, mid, vorname=vorname, name=name, abteilung=abteilung, position=position,
            urlaubsanspruch=urlaubsanspruch, sollstunden_tag=sollstunden_tag,
            personalverantwortlich=personalverantwortlich)
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


# ── Abwesenheiten (Antrag → Genehmigung) ─────────────────────────────────────
@router.get("/abwesenheiten")
def abwesenheiten(request: Request):
    con = tenant_for(request)
    try:
        eintraege = svc.list_abwesenheiten(con)
        mitarbeiter = svc.list_mitarbeiter(con)
        offen = [e for e in eintraege if e["status"] == "beantragt"]
    finally:
        con.close()
    return page(request, "personal/abwesenheiten.html",
                eintraege=eintraege, offen=offen, mitarbeiter=mitarbeiter,
                arten=svc.ABW_ARTEN, titel="Abwesenheiten")


@router.post("/abwesenheiten")
def abwesenheit_speichern(request: Request,
                          mitarbeiter_id: int = Form(...), art: str = Form(...),
                          von: str = Form(...), bis: str = Form(...),
                          notiz: str = Form("")):
    con = tenant_for(request)
    try:
        # Krankheit/Sonstiges sind keine genehmigungspflichtigen Anträge.
        status = "genehmigt" if art in ("Krankheit", "Sonstiges") else "beantragt"
        svc.create_abwesenheit(
            con, mitarbeiter_id=mitarbeiter_id, art=art, von=von, bis=bis,
            notiz=notiz, status=status)
    finally:
        con.close()
    return RedirectResponse("/personal/abwesenheiten", status_code=303)


@router.post("/abwesenheiten/{aid}/genehmigen")
def abwesenheit_genehmigen(request: Request, aid: int):
    con = tenant_for(request)
    try:
        svc.set_abwesenheit_status(con, aid, "genehmigt")
    finally:
        con.close()
    return RedirectResponse("/personal/abwesenheiten", status_code=303)


@router.post("/abwesenheiten/{aid}/ablehnen")
def abwesenheit_ablehnen(request: Request, aid: int):
    con = tenant_for(request)
    try:
        svc.set_abwesenheit_status(con, aid, "abgelehnt")
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


# ── Kalender ─────────────────────────────────────────────────────────────────
@router.get("/kalender")
def kalender(request: Request, jahr: int | None = None, monat: int | None = None):
    heute = date.today()
    jahr = jahr or heute.year
    monat = monat if monat in range(1, 13) else heute.month
    con = tenant_for(request)
    try:
        kal = svc.kalender_monat(con, jahr, monat)
    finally:
        con.close()
    return page(request, "personal/kalender.html", kal=kal, titel="Kalender")


# ── Zeiterfassung (Kommt/Geht, Überstunden) ──────────────────────────────────
@router.get("/zeiterfassung")
def zeiterfassung(request: Request, jahr: int | None = None, monat: int | None = None):
    heute = date.today()
    jahr = jahr or heute.year
    monat = monat if monat in range(1, 13) else heute.month
    con = tenant_for(request)
    try:
        mitarbeiter = svc.list_mitarbeiter(con)
        zeiten = svc.list_zeiten(con, jahr=jahr, monat=monat)
        summe = svc.zeit_monatssumme(con, jahr, monat)
    finally:
        con.close()
    return page(request, "personal/zeiterfassung.html",
                mitarbeiter=mitarbeiter, zeiten=zeiten, summe=summe,
                jahr=jahr, monat=monat, heute=heute.isoformat(), titel="Zeiterfassung")


@router.post("/zeiterfassung")
def zeit_speichern(request: Request,
                   mitarbeiter_id: int = Form(...), datum: str = Form(...),
                   kommt: str = Form(""), geht: str = Form(""),
                   pause_minuten: int = Form(0), notiz: str = Form("")):
    con = tenant_for(request)
    try:
        svc.create_zeit(con, mitarbeiter_id=mitarbeiter_id, datum=datum,
                        kommt=kommt, geht=geht, pause_minuten=pause_minuten, notiz=notiz)
    finally:
        con.close()
    return RedirectResponse("/personal/zeiterfassung", status_code=303)


@router.post("/zeiterfassung/stempeln")
def zeit_stempeln(request: Request, mitarbeiter_id: int = Form(...)):
    con = tenant_for(request)
    try:
        svc.stempeln(con, mitarbeiter_id)
    finally:
        con.close()
    return RedirectResponse("/personal/zeiterfassung", status_code=303)


@router.post("/zeiterfassung/{zid}/loeschen")
def zeit_loeschen(request: Request, zid: int):
    con = tenant_for(request)
    try:
        svc.delete_zeit(con, zid)
    finally:
        con.close()
    return RedirectResponse("/personal/zeiterfassung", status_code=303)


# ── Digitale Personalakte (Verträge, Dokumente) ──────────────────────────────
_ERLAUBTE_ENDUNGEN = {".pdf", ".png", ".jpg", ".jpeg", ".docx", ".doc", ".txt", ".odt"}


def _sicherer_dateiname(name: str) -> str:
    """ASCII-only, ohne Pfadanteile – verhindert Path-Traversal."""
    name = unicodedata.normalize("NFKD", name or "").encode("ascii", "ignore").decode()
    name = name.replace("\\", "/").split("/")[-1]
    name = re.sub(r"[^A-Za-z0-9._-]+", "_", name).strip("._") or "datei"
    return name[:120]


def _akte_dir(request: Request, mid: int) -> Path:
    slug = request.session["user"]["firma_slug"]
    p = tenant_dir(slug) / "dokumente" / str(mid)
    p.mkdir(parents=True, exist_ok=True)
    return p


@router.get("/{mid}/akte")
def akte(request: Request, mid: int):
    con = tenant_for(request)
    try:
        ma = svc.get_mitarbeiter(con, mid)
        if not ma:
            raise HTTPException(status_code=404, detail="Mitarbeiter nicht gefunden")
        dokumente = svc.list_dokumente(con, mid)
    finally:
        con.close()
    return page(request, "personal/akte.html", ma=ma, dokumente=dokumente,
                kategorien=svc.DOK_KATEGORIEN, titel="Personalakte")


@router.post("/{mid}/akte/upload")
async def akte_upload(request: Request, mid: int,
                      kategorie: str = Form("Sonstiges"), titel: str = Form(""),
                      datei: UploadFile = File(...)):
    endung = Path(datei.filename or "").suffix.lower()
    if endung not in _ERLAUBTE_ENDUNGEN:
        raise HTTPException(status_code=400, detail=f"Dateityp {endung} nicht erlaubt")
    ablage_name = f"{int(date.today().strftime('%Y%m%d'))}_{_sicherer_dateiname(datei.filename)}"
    ziel = _akte_dir(request, mid) / ablage_name
    inhalt = await datei.read()
    ziel.write_bytes(inhalt)
    con = tenant_for(request)
    try:
        svc.create_dokument(
            con, mitarbeiter_id=mid, kategorie=kategorie,
            titel=titel or datei.filename, dateiname=datei.filename or ablage_name,
            ablage=ablage_name, groesse=len(inhalt))
    finally:
        con.close()
    return RedirectResponse(f"/personal/{mid}/akte", status_code=303)


@router.get("/dokument/{did}")
def dokument_download(request: Request, did: int):
    con = tenant_for(request)
    try:
        dok = svc.get_dokument(con, did)
    finally:
        con.close()
    if not dok:
        raise HTTPException(status_code=404, detail="Dokument nicht gefunden")
    pfad = _akte_dir(request, dok["mitarbeiter_id"]) / dok["ablage"]
    if not pfad.exists():
        raise HTTPException(status_code=404, detail="Datei fehlt")
    return FileResponse(str(pfad), filename=dok["dateiname"] or dok["ablage"])


@router.post("/dokument/{did}/loeschen")
def dokument_loeschen(request: Request, did: int):
    con = tenant_for(request)
    try:
        dok = svc.delete_dokument(con, did)
    finally:
        con.close()
    if dok:
        pfad = _akte_dir(request, dok["mitarbeiter_id"]) / dok["ablage"]
        try:
            pfad.unlink(missing_ok=True)
        except OSError:
            pass
        return RedirectResponse(f"/personal/{dok['mitarbeiter_id']}/akte", status_code=303)
    return RedirectResponse("/personal", status_code=303)
