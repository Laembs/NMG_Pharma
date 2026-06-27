"""NMGone-Web · FastAPI-Einstieg (Pilot).

Start (Dev):  uvicorn web.app:app --reload
Dann:         http://localhost:8000
"""
from __future__ import annotations

import os
from pathlib import Path

from fastapi import FastAPI, Form, Request
from fastapi.responses import FileResponse, JSONResponse, RedirectResponse
from fastapi.staticfiles import StaticFiles
from starlette.middleware.sessions import SessionMiddleware

from . import auth, sso
from .licensing import landing_path, module_fuer_dashboard
from .routers import kasse, personal
from .services import kasse_service
from .templating import page, templates
from .tenancy import init_platform_db, platform_con, tenant_con

WEB_DIR = Path(__file__).resolve().parent

app = FastAPI(title="NMGone-Web (Pilot)")

app.mount("/static", StaticFiles(directory=str(WEB_DIR / "static")), name="static")
app.include_router(personal.router)
app.include_router(kasse.router)

# Pfade, die ohne Login erreichbar sind (inkl. PWA-Dateien im Root-Scope).
# /api/* prüft selbst per signiertem Token (kein Session-Login).
_OEFFENTLICH = ("/login", "/logout", "/static", "/healthz",
                "/sw.js", "/manifest.webmanifest", "/sso/cockpit", "/api/")


# ── PWA: Service Worker + Manifest im Root-Scope ─────────────────────────────
# Der Service Worker muss von "/" ausgeliefert werden, damit sein Scope die
# ganze App umfasst (aus /static/ wäre der Scope zu eng).
@app.get("/sw.js")
def service_worker():
    return FileResponse(str(WEB_DIR / "static" / "sw.js"), media_type="text/javascript")


@app.get("/manifest.webmanifest")
def manifest():
    return FileResponse(str(WEB_DIR / "static" / "manifest.webmanifest"),
                        media_type="application/manifest+json")


@app.on_event("startup")
def _startup() -> None:
    init_platform_db()


@app.middleware("http")
async def auth_guard(request: Request, call_next):
    """Stellt request.state.user bereit und schützt nicht-öffentliche Seiten."""
    request.state.user = request.session.get("user")
    path = request.url.path
    ist_oeffentlich = path == "/" or any(path.startswith(p) for p in _OEFFENTLICH)
    if not ist_oeffentlich and not request.state.user:
        return RedirectResponse("/login", status_code=303)
    return await call_next(request)


@app.get("/healthz")
def healthz():
    return {"status": "ok"}


@app.get("/")
def root(request: Request):
    user = request.session.get("user")
    ziel = landing_path(user["firma_id"]) if user else "/login"
    return RedirectResponse(ziel, status_code=303)


# ── Login / Logout ───────────────────────────────────────────────────────────
@app.get("/login")
def login_form(request: Request, fehler: str | None = None):
    with platform_con() as con:
        firmen = [dict(r) for r in con.execute(
            "SELECT slug, name FROM firma ORDER BY name").fetchall()]
    return templates.TemplateResponse(request, "login.html", {
        "firmen": firmen, "fehler": fehler, "user": None, "nav_module": [],
    })


@app.post("/login")
def login_submit(request: Request,
                 firma: str = Form(...), login: str = Form(...),
                 passwort: str = Form(...)):
    user = auth.authenticate(firma, login, passwort)
    if not user:
        return RedirectResponse(
            "/login?fehler=Anmeldung+fehlgeschlagen", status_code=303)
    auth.login_session(request, user)
    return RedirectResponse(landing_path(user["firma_id"]), status_code=303)


@app.get("/logout")
def logout(request: Request):
    auth.logout_session(request)
    return RedirectResponse("/login", status_code=303)


# ── SSO aus dem NMGone-Cockpit (Einmal-Token) ────────────────────────────────
@app.get("/sso/cockpit")
def sso_cockpit(request: Request, token: str = "", next: str = "/kasse"):
    """Anmeldung per signiertem Einmal-Token vom Cockpit – ohne Passwort.

    Das Cockpit hat den Benutzer bereits geprueft und signiert Firma+Login mit dem
    gemeinsamen Geheimnis (PENNONE_SSO_SECRET). Wir pruefen Signatur/Ablauf,
    verbrauchen das Token einmalig und legen die Session an.
    """
    if not sso.secret():
        return RedirectResponse("/login?fehler=SSO+nicht+konfiguriert", status_code=303)
    payload = sso.verify_token(token)
    if not payload or not sso.consume_jti(payload):
        return RedirectResponse("/login?fehler=SSO-Link+ungueltig+oder+abgelaufen",
                                status_code=303)
    user = auth.user_by_slug_login(payload.get("f", ""), payload.get("l", ""))
    if not user:
        return RedirectResponse("/login?fehler=Benutzer+in+der+Kasse+unbekannt",
                                status_code=303)
    auth.login_session(request, user)
    # Open-Redirect-Schutz: nur lokale Pfade zulassen.
    ziel = next if next.startswith("/") and not next.startswith("//") else "/kasse"
    return RedirectResponse(ziel, status_code=303)


# ── Lese-API für die Desktop-Kasse (Online-Verkäufe abrufen) ─────────────────
@app.get("/api/verkaeufe")
def api_verkaeufe(token: str = ""):
    """Liefert die online erfassten Verkäufe als JSON – für die Desktop-Kasse.

    Geschützt per signiertem Token (gleiches Geheimnis wie das Cockpit-SSO); es
    wird KEIN jti verbraucht, damit die Desktop-Kasse beliebig oft abrufen kann.
    Reiner Lesezugriff, keine Änderung.
    """
    payload = sso.verify_token(token)
    if not payload:
        return JSONResponse({"error": "ungueltiges oder abgelaufenes Token"}, status_code=401)
    slug = payload.get("f", "")
    with platform_con() as con:
        firma = con.execute("SELECT slug FROM firma WHERE slug=?", (slug,)).fetchone()
    if not firma:
        return JSONResponse({"error": "Firma unbekannt"}, status_code=404)
    con = tenant_con(slug)
    try:
        kasse_service.ensure_schema(con)
        rows = con.execute(
            """SELECT b.id, b.datum, b.kunde_name, b.status,
                      COALESCE(b.erfasst_von, '') AS erfasst_von,
                      COUNT(p.id) AS positionen,
                      COALESCE(SUM(p.menge), 0) AS stueck
                 FROM tbl_bestellungen b
                 LEFT JOIN tbl_bestellpositionen p ON p.bestellung_id = b.id
                WHERE COALESCE(b.quelle, 'online') = 'online'
                  AND COALESCE(b.uebernommen, 0) = 0
                GROUP BY b.id ORDER BY b.id DESC LIMIT 200""").fetchall()
    finally:
        con.close()
    return JSONResponse({"verkaeufe": [dict(r) for r in rows]})


@app.get("/api/verkauf/{auftrag_id}")
def api_verkauf_detail(auftrag_id: int, token: str = ""):
    """Detail eines Online-Auftrags (Kopf + Positionen mit Bestellung/Vorbestellung)
    für die Detailansicht der Desktop-Kasse. Token-geschützt, reiner Lesezugriff."""
    payload = sso.verify_token(token)
    if not payload:
        return JSONResponse({"error": "ungueltiges oder abgelaufenes Token"}, status_code=401)
    slug = payload.get("f", "")
    with platform_con() as con:
        firma = con.execute("SELECT slug FROM firma WHERE slug=?", (slug,)).fetchone()
    if not firma:
        return JSONResponse({"error": "Firma unbekannt"}, status_code=404)
    con = tenant_con(slug)
    try:
        kasse_service.ensure_schema(con)
        kopf = con.execute(
            """SELECT id, datum, kunde_name, status, liefertermin,
                      COALESCE(erfasst_von, '') AS erfasst_von
                 FROM tbl_bestellungen WHERE id=?""", (auftrag_id,)).fetchone()
        if not kopf:
            return JSONResponse({"error": "Auftrag nicht gefunden"}, status_code=404)
        pos = con.execute(
            """SELECT pzn, bezeichnung, menge,
                      COALESCE(bestellart, 'Bestellung') AS bestellart
                 FROM tbl_bestellpositionen WHERE bestellung_id=? ORDER BY id""",
            (auftrag_id,)).fetchall()
    finally:
        con.close()
    return JSONResponse({"kopf": dict(kopf), "positionen": [dict(r) for r in pos]})


@app.get("/api/vorbestellungen")
def api_vorbestellungen(token: str = ""):
    """Offene Vorbestell-Positionen der online erfassten Aufträge (für den
    Vorbestellungen-Reiter der Desktop-Kasse). Token-geschützt, nur Lesen."""
    payload = sso.verify_token(token)
    if not payload:
        return JSONResponse({"error": "ungueltiges oder abgelaufenes Token"}, status_code=401)
    slug = payload.get("f", "")
    with platform_con() as con:
        firma = con.execute("SELECT slug FROM firma WHERE slug=?", (slug,)).fetchone()
    if not firma:
        return JSONResponse({"error": "Firma unbekannt"}, status_code=404)
    con = tenant_con(slug)
    try:
        kasse_service.ensure_schema(con)
        rows = con.execute(
            """SELECT b.id AS auftrag, b.datum, b.kunde_name, b.liefertermin,
                      p.pzn, p.bezeichnung, p.menge
                 FROM tbl_bestellpositionen p
                 JOIN tbl_bestellungen b ON b.id = p.bestellung_id
                WHERE COALESCE(b.quelle, 'online') = 'online'
                  AND COALESCE(b.uebernommen, 0) = 0
                  AND p.bestellart = 'Vorbestellung'
                ORDER BY b.id DESC, p.id""").fetchall()
    finally:
        con.close()
    return JSONResponse({"vorbestellungen": [dict(r) for r in rows]})


@app.post("/api/verkauf/{auftrag_id}/uebernommen")
def api_verkauf_uebernommen(auftrag_id: int, token: str = "", von: str = ""):
    """Markiert einen Online-Auftrag als in die PC-Kasse übernommen (Schreibzugriff).

    Token-geschützt (gleiches Geheimnis). Idempotent: nur noch nicht übernommene
    Aufträge werden gesetzt. Der Auftrag bleibt in der Web-Historie, fällt aber aus
    den To-do-Listen (Verkäufe/Vorbestellungen) heraus.
    """
    payload = sso.verify_token(token)
    if not payload:
        return JSONResponse({"error": "ungueltiges oder abgelaufenes Token"}, status_code=401)
    slug = payload.get("f", "")
    with platform_con() as con:
        firma = con.execute("SELECT slug FROM firma WHERE slug=?", (slug,)).fetchone()
    if not firma:
        return JSONResponse({"error": "Firma unbekannt"}, status_code=404)
    con = tenant_con(slug)
    try:
        kasse_service.ensure_schema(con)
        cur = con.execute(
            """UPDATE tbl_bestellungen
                  SET uebernommen=1, uebernommen_am=datetime('now'), uebernommen_von=?
                WHERE id=? AND COALESCE(uebernommen,0)=0""",
            (von or payload.get("l", ""), auftrag_id))
        con.commit()
    finally:
        con.close()
    return JSONResponse({"ok": True, "geaendert": cur.rowcount})


@app.post("/api/lager/push")
async def api_lager_push(request: Request, token: str = ""):
    """Nimmt den aktuellen Bestand + Artikelstamm der PC-Kasse entgegen (PC ist die
    Quelle) und ersetzt damit den Online-Bestand. Token-geschützt. Body (JSON):
    {"artikel": [{pzn, bezeichnung}], "bestand": [{pzn, bezeichnung, charge, verfall, menge}]}."""
    payload = sso.verify_token(token)
    if not payload:
        return JSONResponse({"error": "ungueltiges oder abgelaufenes Token"}, status_code=401)
    slug = payload.get("f", "")
    with platform_con() as con:
        firma = con.execute("SELECT slug FROM firma WHERE slug=?", (slug,)).fetchone()
    if not firma:
        return JSONResponse({"error": "Firma unbekannt"}, status_code=404)
    try:
        body = await request.json()
    except Exception:
        return JSONResponse({"error": "ungueltiger Body"}, status_code=400)
    con = tenant_con(slug)
    try:
        kasse_service.ensure_schema(con)
        n_a, n_b = kasse_service.lager_ersetzen(
            con, body.get("artikel", []), body.get("bestand", []))
    finally:
        con.close()
    return JSONResponse({"ok": True, "artikel": n_a, "bestand": n_b})


# ── Dashboard ────────────────────────────────────────────────────────────────
@app.get("/dashboard")
def dashboard(request: Request):
    user = request.session["user"]
    kacheln = module_fuer_dashboard(user["firma_id"])
    return page(request, "dashboard.html", kacheln=kacheln, titel="Übersicht")


# WICHTIG: zuletzt hinzugefügte Middleware ist die ÄUSSERSTE. Die
# SessionMiddleware muss vor auth_guard laufen (das request.session liest),
# daher wird sie hier ganz am Ende registriert.
app.add_middleware(
    SessionMiddleware,
    secret_key=os.environ.get("WEB_SESSION_SECRET", "dev-only-unsicher-bitte-aendern"),
    same_site="lax",
)
