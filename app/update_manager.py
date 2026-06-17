from __future__ import annotations
import json
import os
import shutil
import subprocess
import sys
import tempfile
import zipfile
from datetime import datetime
from pathlib import Path
from .config import BASE_DIR, UPDATE_DIR, BACKUP_DIR, DB_PATH, LOG_DIR
from .backup import backup_erstellen, APP_VERSION, DB_SCHEMA_VERSION
from .migrations import run_migrations

PACKAGE_EXT = ".nmgupdate"
PROTECTED_DIRS = {"data", "backups", "gespeicherte_analysen", "ausgaben", "updates", "logs"}
PROTECTED_FILES = {"nmg_startdatenbank.sqlite"}

ROLLBACK_PREFIX = "Rollback_Vor_Update_"


def create_rollback_snapshot(target_version: str = "") -> Path:
    """Sichert Programmdateien + Datenbank vor einem Update als Rücksprungpunkt."""
    BACKUP_DIR.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    safe_target = str(target_version or "unbekannt").replace(".", "_")
    out = BACKUP_DIR / f"{ROLLBACK_PREFIX}{APP_VERSION}_nach_{safe_target}_{stamp}.zip"
    manifest = {
        "type": "rollback_snapshot",
        "from_version": APP_VERSION,
        "target_version": target_version,
        "created_at": datetime.now().isoformat(timespec="seconds"),
    }
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("manifest.json", json.dumps(manifest, ensure_ascii=False, indent=2))
        # Programmdateien sichern, Nutzerdaten nicht komplett duplizieren.
        for rel_root in ["app", "assets"]:
            root = BASE_DIR / rel_root
            if root.exists():
                for f in root.rglob("*"):
                    if f.is_file() and "__pycache__" not in f.parts:
                        zf.write(f, str(f.relative_to(BASE_DIR)))
        for rel in ["start.py", "start.bat", "requirements.txt", "version.json", "README.md"]:
            f = BASE_DIR / rel
            if f.exists():
                zf.write(f, rel)
        if DB_PATH.exists():
            zf.write(DB_PATH, "data/nmg_startdatenbank.sqlite")
    return out


def list_rollback_snapshots() -> list[Path]:
    BACKUP_DIR.mkdir(parents=True, exist_ok=True)
    return sorted(BACKUP_DIR.glob(f"{ROLLBACK_PREFIX}*.zip"), key=lambda x: x.stat().st_mtime, reverse=True)


def restore_rollback_snapshot(snapshot: str | Path) -> dict:
    """Stellt einen vor einem Update erzeugten Rücksprungpunkt wieder her."""
    snapshot = Path(snapshot)
    if not snapshot.exists():
        raise FileNotFoundError(str(snapshot))
    # Sicherheitsbackup des aktuellen Zustands, bevor zurückgespielt wird.
    safety = create_rollback_snapshot("vor_rollback")
    restored = []
    with zipfile.ZipFile(snapshot, "r") as zf:
        names = zf.namelist()
        if "manifest.json" not in names:
            raise ValueError("Rollback-Paket enthält kein manifest.json.")
        manifest = json.loads(zf.read("manifest.json").decode("utf-8"))
        if manifest.get("type") != "rollback_snapshot":
            raise ValueError("Dies ist kein gültiger Rücksprungpunkt.")
        tmp = Path(tempfile.mkdtemp(prefix="nmg_rollback_"))
        try:
            zf.extractall(tmp)
            for rel in names:
                if rel.endswith("/") or rel == "manifest.json":
                    continue
                src = tmp / rel
                dst = BASE_DIR / rel
                dst.parent.mkdir(parents=True, exist_ok=True)
                shutil.copy2(src, dst)
                restored.append(rel)
        finally:
            shutil.rmtree(tmp, ignore_errors=True)
    _write_install_log(snapshot.name, APP_VERSION, str(manifest.get("from_version", "unbekannt")), "ROLLBACK", f"{len(restored)} Dateien wiederhergestellt. Sicherheitskopie: {safety}")
    return {"status": "OK", "snapshot": str(snapshot), "restored": restored, "safety": str(safety), "restart_required": True}


def _version_tuple(v: str) -> tuple:
    clean = v.replace("SP", "").replace("_", ".").replace("-", ".")
    nums = []
    for part in clean.split('.'):
        num = ''.join(ch for ch in part if ch.isdigit())
        nums.append(int(num or 0))
    return tuple(nums + [0] * (4 - len(nums)))


def read_update_manifest(package: str | Path) -> dict:
    package = Path(package)
    if not package.exists():
        raise FileNotFoundError(str(package))
    with zipfile.ZipFile(package, "r") as zf:
        if "manifest.json" not in zf.namelist():
            raise ValueError("Updatepaket enthält kein manifest.json.")
        return json.loads(zf.read("manifest.json").decode("utf-8"))


def validate_update_package(package: str | Path) -> dict:
    package = Path(package)
    if package.suffix.lower() not in (PACKAGE_EXT, ".zip"):
        raise ValueError("Bitte ein .nmgupdate-Paket auswählen.")
    manifest = read_update_manifest(package)
    target = str(manifest.get("target_version", "")).strip()
    if not target:
        raise ValueError("Updatepaket enthält keine Zielversion.")
    return manifest


def install_update_package(package: str | Path) -> dict:
    """Installiert ein Updatepaket.

    Format:
    - manifest.json
    - files/<Programmdateien>

    Datenordner, Backups und gespeicherte Analysen werden bewusst nicht überschrieben.
    """
    package = Path(package)
    manifest = validate_update_package(package)
    target_version = str(manifest.get("target_version"))

    UPDATE_DIR.mkdir(parents=True, exist_ok=True)
    BACKUP_DIR.mkdir(parents=True, exist_ok=True)

    backup_path = None
    rollback_path = create_rollback_snapshot(target_version)
    if DB_PATH.exists():
        backup_path = backup_erstellen()

    tmp = Path(tempfile.mkdtemp(prefix="nmg_update_"))
    copied = []
    skipped = []
    try:
        with zipfile.ZipFile(package, "r") as zf:
            zf.extractall(tmp)
        files_root = tmp / "files"
        if not files_root.exists():
            raise ValueError("Updatepaket enthält keinen files/-Ordner.")

        # Sicherheitskopie des Pakets behalten
        shutil.copy2(package, UPDATE_DIR / package.name)

        for src in files_root.rglob("*"):
            if src.is_dir():
                continue
            rel = src.relative_to(files_root)
            parts = set(rel.parts)
            if parts & PROTECTED_DIRS or rel.name in PROTECTED_FILES:
                skipped.append(str(rel))
                continue
            dst = BASE_DIR / rel
            dst.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(src, dst)
            copied.append(str(rel))

        migration_actions = run_migrations(DB_PATH)
        _write_install_log(package.name, APP_VERSION, target_version, "OK", f"{len(copied)} Dateien kopiert. Migrationen: {migration_actions}")
        return {
            "status": "OK",
            "from_version": APP_VERSION,
            "target_version": target_version,
            "copied": copied,
            "skipped": skipped,
            "backup": str(backup_path) if backup_path else None,
            "rollback": str(rollback_path),
            "migrations": migration_actions,
            "restart_required": True,
        }
    except Exception as exc:
        _write_install_log(package.name, APP_VERSION, target_version, "FEHLER", str(exc))
        raise
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def _write_install_log(package: str, from_v: str, to_v: str, status: str, msg: str) -> None:
    try:
        import sqlite3
        if DB_PATH.exists():
            con = sqlite3.connect(DB_PATH)
            con.execute("CREATE TABLE IF NOT EXISTS tbl_update_log(id INTEGER PRIMARY KEY AUTOINCREMENT, zeitpunkt TEXT DEFAULT CURRENT_TIMESTAMP, von_version TEXT, nach_version TEXT, paket TEXT, status TEXT, meldung TEXT)")
            con.execute("INSERT INTO tbl_update_log(von_version,nach_version,paket,status,meldung) VALUES(?,?,?,?,?)", (from_v, to_v, package, status, msg))
            con.commit(); con.close()
    except Exception:
        pass
    # SP10: vorher BASE_DIR / "logs" - im installierten Programm read-only.
    LOG_DIR.mkdir(parents=True, exist_ok=True)
    with (LOG_DIR / "update_log.txt").open("a", encoding="utf-8") as f:
        f.write(f"{datetime.now().isoformat(timespec='seconds')} | {status} | {package} | {from_v} -> {to_v} | {msg}\n")


def write_version_file() -> Path:
    # SP10: vorher BASE_DIR / "version.json" - im installierten Programm
    # read-only -> Schreibversuch warf PermissionError, wurde aber von
    # gui.py __init__ als except: pass geschluckt (Version-Tracking lief
    # einfach still in Leere). Jetzt LOG_DIR (in USERDATA_ROOT, writable).
    LOG_DIR.mkdir(parents=True, exist_ok=True)
    p = LOG_DIR / "version.json"
    payload = {
        "app": "NMGone",
        "version": APP_VERSION,
        "db_schema_version": DB_SCHEMA_VERSION,
        "build_date": datetime.now().strftime("%Y-%m-%d"),
    }
    p.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    return p



def find_update_packages() -> list[dict]:
    """Sucht lokale Updatepakete im updates/-Ordner und sortiert neue Versionen nach Zielversion."""
    UPDATE_DIR.mkdir(parents=True, exist_ok=True)
    packages = []
    for pkg in sorted(UPDATE_DIR.glob(f"*{PACKAGE_EXT}")):
        try:
            manifest = validate_update_package(pkg)
            target = str(manifest.get("target_version", ""))
            packages.append({
                "path": pkg,
                "name": pkg.name,
                "target_version": target,
                "is_newer": _version_tuple(target) > _version_tuple(APP_VERSION),
                "manifest": manifest,
            })
        except Exception:
            continue
    packages.sort(key=lambda x: _version_tuple(str(x.get("target_version", "0"))), reverse=True)
    return packages


def find_newest_update() -> dict | None:
    """Gibt das neueste Updatepaket zurück, wenn es neuer als die aktuelle Version ist."""
    for item in find_update_packages():
        if item.get("is_newer"):
            return item
    return None


def restart_application() -> None:
    """Startet die Anwendung nach einem Update neu und beendet den aktuellen Prozess."""
    try:
        if getattr(sys, "frozen", False):
            subprocess.Popen([sys.executable], cwd=str(BASE_DIR))
        else:
            start_py = BASE_DIR / "start.py"
            if start_py.exists():
                subprocess.Popen([sys.executable, str(start_py)], cwd=str(BASE_DIR))
            else:
                subprocess.Popen([sys.executable] + sys.argv, cwd=str(BASE_DIR))
    except Exception as exc:
        raise RuntimeError(f"Automatischer Neustart fehlgeschlagen: {exc}")

def open_updates_folder() -> None:
    UPDATE_DIR.mkdir(parents=True, exist_ok=True)
    if sys.platform.startswith("win"):
        os.startfile(str(UPDATE_DIR))  # type: ignore[attr-defined]
    elif sys.platform == "darwin":
        os.system(f'open "{UPDATE_DIR}"')
    else:
        os.system(f'xdg-open "{UPDATE_DIR}"')
