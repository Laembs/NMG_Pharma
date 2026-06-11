from __future__ import annotations
import json
import zipfile
from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parent
VERSION = "3.0.0"
OUT = ROOT / "updates" / f"NMG_Update_{VERSION.replace('.', '_')}.nmgupdate"
INCLUDE = ["app", "assets", "installer", "start.py", "requirements.txt", "README.md", "build_exe.bat", "build_installer.bat", "setup_dev.bat", "INSTALLATION_HINWEISE.txt", "README_3_0.txt"]
EXCLUDE_DIRS = {"__pycache__", ".git", "data", "backups", "gespeicherte_analysen", "ausgaben", "updates", "logs", "dist", "build"}

manifest = {
    "app": "NMG Analyse",
    "target_version": VERSION,
    "db_schema_version": "1.1",
    "created_at": datetime.now().isoformat(timespec="seconds"),
    "notes": "Version 3.0.0: Installationsroutine, Programm-/Daten-Trennung, Vollversions-Update und Setup-Vorlage.",
}

OUT.parent.mkdir(parents=True, exist_ok=True)
with zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED) as zf:
    zf.writestr("manifest.json", json.dumps(manifest, ensure_ascii=False, indent=2))
    for item in INCLUDE:
        path = ROOT / item
        if not path.exists():
            continue
        if path.is_file():
            zf.write(path, f"files/{path.name}")
        else:
            for p in path.rglob("*"):
                if p.is_dir() or any(part in EXCLUDE_DIRS for part in p.relative_to(ROOT).parts):
                    continue
                zf.write(p, f"files/{p.relative_to(ROOT).as_posix()}")
print(OUT)
