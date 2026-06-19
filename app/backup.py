from __future__ import annotations

import json
import shutil
import sqlite3
import zipfile
from datetime import datetime
from pathlib import Path

from .config import BASE_DIR, DATA_DIR, OUTPUT_DIR, DB_PATH, BACKUP_DIR as _CONFIG_BACKUP_DIR

# SP10 Critical Fix: vorher stand hier "BACKUP_DIR = BASE_DIR / 'backups'",
# was im installierten Programm auf C:\Program Files\NMGone\_internal\backups
# zeigte (read-only!). Beim Programmstart wollte backup_auto_taeglich() dort
# mkdir aufrufen und crashte mit PermissionError [WinError 5].
# Fix: BACKUP_DIR aus config nehmen (USERDATA_ROOT/backups, also
# C:\ProgramData\NMGone\backups im installierten Programm, writable).
BACKUP_DIR = _CONFIG_BACKUP_DIR
APP_VERSION = "1.1.3"
APP_VERSION_DISPLAY = "V1.1 SP3"
DB_SCHEMA_VERSION = "1.1"


def _ts() -> str:
    return datetime.now().strftime("%Y%m%d_%H%M%S")


def _db_counts() -> dict:
    counts = {}
    if not DB_PATH.exists():
        return counts
    tables = [
        "tbl_pzn_basisdaten",
        "tbl_nmg_stamm",
        "tbl_austauschartikel",
        "tbl_lieferfaehigkeit",
        "tbl_lernhistorie",
        "tbl_auswertungen",
        "tbl_auswertungspositionen",
    ]
    try:
        con = sqlite3.connect(DB_PATH)
        cur = con.cursor()
        cur.execute("SELECT name FROM sqlite_master WHERE type='table'")
        existing = {r[0] for r in cur.fetchall()}
        for table in tables:
            if table in existing:
                cur.execute(f"SELECT COUNT(*) FROM {table}")
                counts[table] = cur.fetchone()[0]
        con.close()
    except Exception as exc:
        counts["status"] = f"Zählung nicht möglich: {exc}"
    return counts


def backup_erstellen() -> Path:
    """Erstellt ein ZIP-Backup der aktuellen Datenbank und wichtiger Nutzdaten."""
    BACKUP_DIR.mkdir(parents=True, exist_ok=True)
    if not DB_PATH.exists():
        raise FileNotFoundError(f"Datenbank nicht gefunden: {DB_PATH}")

    backup_path = BACKUP_DIR / f"NMGone_Backup_{APP_VERSION}_{_ts()}.zip"
    manifest = {
        "app_version": APP_VERSION,
        "db_schema_version": DB_SCHEMA_VERSION,
        "created_at": datetime.now().isoformat(timespec="seconds"),
        "database": "data/nmg_startdatenbank.sqlite",
        "counts": _db_counts(),
    }

    with zipfile.ZipFile(backup_path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.write(DB_PATH, "data/nmg_startdatenbank.sqlite")
        # Kleine Protokoll-/Konfigurationsdateien mitsichern, aber keine großen Rohdatenordner verdoppeln.
        for rel in ["data/lernimport_bericht_1_4.csv", "data/lernimport_zusammenfassung_1_4.txt"]:
            p = BASE_DIR / rel
            if p.exists():
                zf.write(p, rel)
        zf.writestr("manifest.json", json.dumps(manifest, ensure_ascii=False, indent=2))

    return backup_path


def backup_pruefen(backup_file: str | Path) -> dict:
    """Liest das Manifest und prüft, ob eine Datenbank im Backup enthalten ist."""
    backup_file = Path(backup_file)
    if not backup_file.exists():
        raise FileNotFoundError(str(backup_file))
    with zipfile.ZipFile(backup_file, "r") as zf:
        names = set(zf.namelist())
        has_db = "data/nmg_startdatenbank.sqlite" in names or "nmg_startdatenbank.sqlite" in names
        manifest = {}
        if "manifest.json" in names:
            manifest = json.loads(zf.read("manifest.json").decode("utf-8"))
        manifest["has_database"] = has_db
        manifest["file"] = str(backup_file)
        return manifest


def backup_wiederherstellen(backup_file: str | Path) -> Path:
    """Stellt die Datenbank aus einem Backup wieder her. Vorher wird die aktuelle DB gesichert."""
    backup_file = Path(backup_file)
    info = backup_pruefen(backup_file)
    if not info.get("has_database"):
        raise ValueError("Das Backup enthält keine nmg_startdatenbank.sqlite.")

    DATA_DIR.mkdir(parents=True, exist_ok=True)
    BACKUP_DIR.mkdir(parents=True, exist_ok=True)

    if DB_PATH.exists():
        safety_copy = BACKUP_DIR / f"Vor_Restore_DB_{_ts()}.sqlite"
        shutil.copy2(DB_PATH, safety_copy)

    # SP10: vorher BASE_DIR / ".restore_tmp" - im installierten Programm
    # read-only. Jetzt OS-Tempdir.
    import tempfile as _tempfile
    tmp_dir = Path(_tempfile.gettempdir()) / "nmgone_restore_tmp"
    if tmp_dir.exists():
        shutil.rmtree(tmp_dir)
    tmp_dir.mkdir(parents=True, exist_ok=True)

    try:
        with zipfile.ZipFile(backup_file, "r") as zf:
            names = set(zf.namelist())
            member = "data/nmg_startdatenbank.sqlite" if "data/nmg_startdatenbank.sqlite" in names else "nmg_startdatenbank.sqlite"
            zf.extract(member, tmp_dir)
            extracted = tmp_dir / member
            shutil.copy2(extracted, DB_PATH)
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)

    return DB_PATH



def backup_auto_taeglich(aufbewahrung_tage: int = 7) -> dict:
    """Erstellt beim Programmstart höchstens ein automatisches Backup pro Tag.
    Es werden nur Auto-Backups rotiert; manuelle Backups bleiben erhalten.
    """
    BACKUP_DIR.mkdir(parents=True, exist_ok=True)
    today = datetime.now().strftime("%Y%m%d")
    existing_today = sorted(BACKUP_DIR.glob(f"NMGone_AutoBackup_{today}_*.zip"))
    if existing_today:
        return {"created": False, "file": existing_today[-1], "message": "Auto-Backup für heute bereits vorhanden."}

    if not DB_PATH.exists():
        return {"created": False, "file": None, "message": "Keine Datenbank für Auto-Backup gefunden."}

    backup_path = BACKUP_DIR / f"NMGone_AutoBackup_{today}_{datetime.now().strftime('%H%M%S')}.zip"
    manifest = {
        "app_version": APP_VERSION,
        "db_schema_version": DB_SCHEMA_VERSION,
        "created_at": datetime.now().isoformat(timespec="seconds"),
        "type": "auto_daily",
        "retention_days": aufbewahrung_tage,
        "database": "data/nmg_startdatenbank.sqlite",
        "counts": _db_counts(),
    }
    with zipfile.ZipFile(backup_path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.write(DB_PATH, "data/nmg_startdatenbank.sqlite")
        zf.writestr("manifest.json", json.dumps(manifest, ensure_ascii=False, indent=2))

    # Rotation: nur die letzten N Auto-Backups behalten, älteste zuerst löschen.
    auto_backups = sorted(BACKUP_DIR.glob("NMGone_AutoBackup_*.zip"), key=lambda p: p.stat().st_mtime)
    deleted = []
    while len(auto_backups) > aufbewahrung_tage:
        old = auto_backups.pop(0)
        try:
            old.unlink()
            deleted.append(str(old))
        except OSError:
            pass
    return {"created": True, "file": backup_path, "deleted": deleted, "message": "Auto-Backup erstellt."}


def versionsinfo() -> str:
    return f"NMGone {APP_VERSION_DISPLAY} ({APP_VERSION}) | Datenbank-Schema {DB_SCHEMA_VERSION} | Datenbank: {DB_PATH}"
