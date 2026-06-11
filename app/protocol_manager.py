from __future__ import annotations

import json
import os
import platform
import shutil
import subprocess
import traceback
import urllib.parse
import zipfile
from datetime import datetime
from pathlib import Path

from .config import BASE_DIR, LOG_DIR, DB_PATH
from .backup import APP_VERSION, DB_SCHEMA_VERSION, versionsinfo

PROTOCOL_ROOT = LOG_DIR / "protokolle"
SUPPORT_DIR = LOG_DIR / "supportpakete"
DEFAULT_RECIPIENT = "Laemb@hotmail.de"

CATEGORIES = {
    "programm": "Programm",
    "fehler": "Fehler",
    "neue_auswertung": "Neue_Auswertung",
    "schulbank": "Schulbank",
    "datenaktualisierung": "Datenaktualisierung",
    "produktanalyse": "Produktanalyse",
    "marktanalyse": "Marktanalyse",
    "abweichungsanalyse": "Abweichungsanalyse",
    "update_backup": "Update_Backup",
    "admin": "Admin",
    "protokolle": "Protokolle",
}


def _now() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def _stamp() -> str:
    return datetime.now().strftime("%Y%m%d_%H%M%S")


def _safe_category(category: str) -> str:
    key = str(category or "programm").strip().lower()
    return CATEGORIES.get(key, CATEGORIES["programm"])


def ensure_protocol_dirs() -> Path:
    PROTOCOL_ROOT.mkdir(parents=True, exist_ok=True)
    SUPPORT_DIR.mkdir(parents=True, exist_ok=True)
    for folder in CATEGORIES.values():
        (PROTOCOL_ROOT / folder).mkdir(parents=True, exist_ok=True)
    return PROTOCOL_ROOT


def _daily_file(category: str) -> Path:
    folder = PROTOCOL_ROOT / _safe_category(category)
    folder.mkdir(parents=True, exist_ok=True)
    return folder / f"{datetime.now():%Y-%m-%d}.log"


def log_event(category: str, action: str, details: str = "", level: str = "INFO", user: str = "") -> Path:
    """Schreibt ein einfaches Tagesprotokoll. Bestehende Dateien werden nur ergänzt."""
    ensure_protocol_dirs()
    path = _daily_file(category)
    user_part = f" | Bearbeiter: {user}" if user else ""
    text = (
        f"[{_now()}] [{str(level or 'INFO').upper()}] {str(action or '').strip()}{user_part}\n"
        f"{str(details or '').strip()}\n"
        "-" * 90 + "\n"
    )
    with path.open("a", encoding="utf-8") as f:
        f.write(text)
    return path


def log_exception(category: str, action: str, exc: BaseException | None = None, user: str = "") -> Path:
    tb = traceback.format_exc()
    if exc is not None and tb.strip() == "NoneType: None":
        tb = "".join(traceback.format_exception(type(exc), exc, exc.__traceback__))
    details = f"Fehler bei: {action}\n\n{tb}"
    return log_event("fehler", action, details, "ERROR", user=user)


def list_protocol_files(category: str | None = None) -> list[dict]:
    ensure_protocol_dirs()
    result = []
    folders = [PROTOCOL_ROOT / _safe_category(category)] if category else [PROTOCOL_ROOT / v for v in CATEGORIES.values()]
    for folder in folders:
        if not folder.exists():
            continue
        for path in sorted(folder.glob("*.log"), key=lambda p: p.stat().st_mtime, reverse=True):
            try:
                result.append({
                    "path": path,
                    "name": path.name,
                    "category": folder.name,
                    "size": path.stat().st_size,
                    "mtime": datetime.fromtimestamp(path.stat().st_mtime).strftime("%d.%m.%Y %H:%M:%S"),
                })
            except OSError:
                continue
    return result


def read_protocol_file(path: str | Path, max_chars: int = 20000) -> str:
    path = Path(path)
    text = path.read_text(encoding="utf-8", errors="replace")
    if len(text) > max_chars:
        return text[-max_chars:]
    return text


def delete_protocol_file(path: str | Path) -> None:
    path = Path(path)
    root = PROTOCOL_ROOT.resolve()
    if root not in path.resolve().parents:
        raise ValueError("Datei liegt nicht im Protokollverzeichnis.")
    path.unlink(missing_ok=True)


def _db_overview_payload() -> dict:
    payload = {"status": "nicht verfügbar"}
    try:
        import sqlite3
        payload = {"db_path": str(DB_PATH), "tables": []}
        if DB_PATH.exists():
            with sqlite3.connect(DB_PATH) as con:
                rows = con.execute("SELECT name FROM sqlite_master WHERE type='table' ORDER BY name").fetchall()
                for (table,) in rows:
                    if str(table).startswith("sqlite_"):
                        continue
                    try:
                        count = con.execute(f'SELECT COUNT(*) FROM "{table}"').fetchone()[0]
                    except Exception:
                        count = None
                    payload["tables"].append({"table": table, "rows": count})
    except Exception as exc:
        payload = {"status": f"Fehler: {exc}"}
    return payload


def create_support_package() -> Path:
    ensure_protocol_dirs()
    SUPPORT_DIR.mkdir(parents=True, exist_ok=True)
    out = SUPPORT_DIR / f"NMG_Supportpaket_{_stamp()}.zip"
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("versionsinfo.txt", versionsinfo())
        zf.writestr("systeminfo.json", json.dumps({
            "created_at": datetime.now().isoformat(timespec="seconds"),
            "app_version": APP_VERSION,
            "db_schema_version": DB_SCHEMA_VERSION,
            "platform": platform.platform(),
            "python": platform.python_version(),
            "db_path": str(DB_PATH),
            "db_exists": DB_PATH.exists(),
            "db_size_bytes": DB_PATH.stat().st_size if DB_PATH.exists() else 0,
        }, ensure_ascii=False, indent=2))
        zf.writestr("datenbankuebersicht.json", json.dumps(_db_overview_payload(), ensure_ascii=False, indent=2))
        if PROTOCOL_ROOT.exists():
            for file in PROTOCOL_ROOT.rglob("*.log"):
                zf.write(file, f"protokolle/{file.relative_to(PROTOCOL_ROOT)}")
        update_log = BASE_DIR / "logs" / "update_log.txt"
        if update_log.exists():
            zf.write(update_log, "update_log.txt")
    log_event("protokolle", "Supportpaket erstellt", str(out))
    return out


def open_mail_with_attachment(path: str | Path, recipient: str = DEFAULT_RECIPIENT, subject: str = "NMG Analyse Protokoll") -> bool:
    """Öffnet nach Möglichkeit Outlook mit Anhang. Fallback: Standard-Mailprogramm ohne echten Anhang."""
    path = Path(path).resolve()
    body = f"Hallo,\n\nanbei das NMG Analyse Protokoll.\n\nDatei: {path}\n"

    # Outlook per COM über PowerShell. Funktioniert nur unter Windows mit installiertem Outlook.
    if os.name == "nt":
        ps = f"""
        $ErrorActionPreference = 'Stop'
        $outlook = New-Object -ComObject Outlook.Application
        $mail = $outlook.CreateItem(0)
        $mail.To = '{recipient}'
        $mail.Subject = '{subject.replace("'", "''")}'
        $mail.Body = @'
{body}
'@
        $mail.Attachments.Add('{str(path).replace("'", "''")}') | Out-Null
        $mail.Display()
        """
        try:
            subprocess.run(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", ps], check=True, timeout=20)
            log_event("protokolle", "Mailentwurf mit Anhang geöffnet", str(path))
            return True
        except Exception as exc:
            log_event("protokolle", "Outlook-Anhang fehlgeschlagen, Fallback mailto", str(exc), "WARN")

    query = urllib.parse.urlencode({"subject": subject, "body": body})
    mailto = f"mailto:{recipient}?{query}"
    try:
        if os.name == "nt":
            os.startfile(mailto)  # type: ignore[attr-defined]
        else:
            subprocess.Popen(["xdg-open", mailto])
        return False
    except Exception:
        raise
