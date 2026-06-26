"""NMGone-Web · FastAPI-Einstieg (Pilot).

Start (Dev):  uvicorn web.app:app --reload
Dann:         http://localhost:8000
"""
from __future__ import annotations

import os
from pathlib import Path

from fastapi import FastAPI, Form, Request
from fastapi.responses import FileResponse, RedirectResponse
from fastapi.staticfiles import StaticFiles
from starlette.middleware.sessions import SessionMiddleware

from . import auth, sso
from .licensing import module_fuer_dashboard
from .routers import kasse, personal
from .templating import page, templates
from .tenancy import init_platform_db, platform_con

WEB_DIR = Path(__file__).resolve().parent

app = FastAPI(title="NMGone-Web (Pilot)")

app.mount("/static", StaticFiles(directory=str(WEB_DIR / "static")), name="static")
app.include_router(personal.router)
app.include_router(kasse.router)

# Pfade, die ohne Login erreichbar sind (inkl. PWA-Dateien im Root-Scope).
_OEFFENTLICH = ("/login", "/logout", "/static", "/healthz",
                "/sw.js", "/manifest.webmanifest", "/sso/cockpit")


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
    ziel = "/dashboard" if request.session.get("user") else "/login"
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
    return RedirectResponse("/dashboard", status_code=303)


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
