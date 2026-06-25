"""Gemeinsame Jinja2-Umgebung + kleine Request-Helfer (vermeidet Zirkel-Importe)."""
from __future__ import annotations

import sqlite3
from pathlib import Path

from fastapi import Request
from fastapi.templating import Jinja2Templates

from .tenancy import tenant_con

TEMPLATES_DIR = Path(__file__).resolve().parent / "templates"
templates = Jinja2Templates(directory=str(TEMPLATES_DIR))


def tenant_for(request: Request) -> sqlite3.Connection:
    """Öffnet die Fachdaten-DB der eingeloggten Firma."""
    user = request.session["user"]
    return tenant_con(user["firma_slug"])


def page(request: Request, name: str, **ctx):
    """Rendert ein Template mit Standard-Kontext (eingeloggter Benutzer + Module)."""
    from .licensing import module_fuer_dashboard
    user = request.session.get("user")
    base = {
        "request": request,
        "user": user,
        "nav_module": module_fuer_dashboard(user["firma_id"]) if user else [],
    }
    base.update(ctx)
    return templates.TemplateResponse(request, name, base)
