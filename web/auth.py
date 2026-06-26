"""Login, Session und Passwort-Hashing.

Die Anmeldung erfolgt mit *Firma + Login + Passwort*. Nach erfolgreichem Login
liegt die Benutzer-Info in einem signierten Session-Cookie (Starlette
SessionMiddleware). ``current_user`` ist eine FastAPI-Dependency, die geschützte
Routen verwenden.
"""
from __future__ import annotations

import bcrypt
from fastapi import Request

from .tenancy import platform_con


def _to_bytes(plain: str) -> bytes:
    # bcrypt verarbeitet max. 72 Bytes – längere Passwörter werden abgeschnitten.
    return plain.encode("utf-8")[:72]


def hash_password(plain: str) -> str:
    return bcrypt.hashpw(_to_bytes(plain), bcrypt.gensalt()).decode("ascii")


def verify_password(plain: str, hashed: str) -> bool:
    try:
        return bcrypt.checkpw(_to_bytes(plain), hashed.encode("ascii"))
    except (ValueError, TypeError):
        return False


def create_user(firma_id: int, login: str, passwort: str,
                anzeigename: str | None = None, ist_admin: bool = False) -> None:
    """Legt einen Benutzer an bzw. setzt sein Passwort neu (idempotent)."""
    with platform_con() as con:
        con.execute(
            """INSERT INTO benutzer(firma_id, login, passwort_hash, anzeigename, ist_admin)
               VALUES(?,?,?,?,?)
               ON CONFLICT(firma_id, login) DO UPDATE SET
                   passwort_hash=excluded.passwort_hash,
                   anzeigename=excluded.anzeigename,
                   ist_admin=excluded.ist_admin""",
            (firma_id, login, hash_password(passwort), anzeigename or login,
             1 if ist_admin else 0))
        con.commit()


def authenticate(firma_slug: str, login: str, passwort: str) -> dict | None:
    """Prüft die Anmeldedaten. Gibt bei Erfolg das Session-User-Dict zurück."""
    with platform_con() as con:
        row = con.execute(
            """SELECT b.id AS user_id, b.firma_id, b.login, b.anzeigename,
                      b.passwort_hash, f.slug AS firma_slug, f.name AS firma_name
                 FROM benutzer b JOIN firma f ON f.id = b.firma_id
                WHERE f.slug=? AND LOWER(b.login)=LOWER(?)""",
            (firma_slug, login)).fetchone()
    if not row or not verify_password(passwort, row["passwort_hash"]):
        return None
    return {
        "user_id": row["user_id"],
        "firma_id": row["firma_id"],
        "firma_slug": row["firma_slug"],
        "firma_name": row["firma_name"],
        "login": row["login"],
        "anzeigename": row["anzeigename"],
    }


def user_by_slug_login(firma_slug: str, login: str) -> dict | None:
    """Benutzer per Firma-Slug + Login holen – OHNE Passwortpruefung.

    Fuer den SSO-Login (das Cockpit hat den Benutzer bereits geprueft). Liefert
    dasselbe Session-User-Dict wie ``authenticate``.
    """
    with platform_con() as con:
        row = con.execute(
            """SELECT b.id AS user_id, b.firma_id, b.login, b.anzeigename,
                      f.slug AS firma_slug, f.name AS firma_name
                 FROM benutzer b JOIN firma f ON f.id = b.firma_id
                WHERE f.slug=? AND LOWER(b.login)=LOWER(?)""",
            (firma_slug, login)).fetchone()
    if not row:
        return None
    return {
        "user_id": row["user_id"],
        "firma_id": row["firma_id"],
        "firma_slug": row["firma_slug"],
        "firma_name": row["firma_name"],
        "login": row["login"],
        "anzeigename": row["anzeigename"],
    }


def login_session(request: Request, user: dict) -> None:
    request.session["user"] = user


def logout_session(request: Request) -> None:
    request.session.pop("user", None)


def current_user(request: Request) -> dict | None:
    """Liest den eingeloggten Benutzer aus der Session (oder None)."""
    return request.session.get("user")
