#!/usr/bin/env python3
"""
build_release.py – NMG Analyse Build & Release Skript
======================================================
Erstellt eine .exe via PyInstaller und bereitet ein GitHub Release vor.

Voraussetzungen:
    pip install pyinstaller requests

Verwendung:
    python build_release.py                  # Build + Release vorbereiten
    python build_release.py --version 2.0.0  # explizite Version
    python build_release.py --build-only     # nur .exe, kein GitHub
    python build_release.py --upload         # nach Build auch auf GitHub hochladen

GitHub-Token:
    Umgebungsvariable GITHUB_TOKEN setzen oder in .env-Datei:
    GITHUB_TOKEN=ghp_xxxxxxxxxxxx
    GITHUB_REPO=deinuser/nmg-analyse
"""

import os
import sys
import json
import shutil
import subprocess
import argparse
from pathlib import Path
from datetime import datetime

# ── Konfiguration ────────────────────────────────────────────────────────────
REPO        = os.getenv("GITHUB_REPO", "DEIN_GITHUB_USER/nmg-analyse")
APP_NAME    = "NMG_Analyse"
ENTRY_POINT = "run.py"          # Startskript (muss existieren)
ICON_FILE   = "assets/nmg_logo.ico"  # optional, leer lassen wenn kein Icon
DIST_DIR    = Path("dist")
BUILD_DIR   = Path("build")
RELEASE_DIR = Path("releases")

# ── Argumente ────────────────────────────────────────────────────────────────
parser = argparse.ArgumentParser(description="NMG Analyse Build & Release")
parser.add_argument("--version", help="Versionsnummer z.B. 1.5.0")
parser.add_argument("--build-only", action="store_true", help="Nur bauen, kein GitHub")
parser.add_argument("--upload", action="store_true", help="Direkt auf GitHub hochladen")
args = parser.parse_args()

# ── Version ermitteln ────────────────────────────────────────────────────────
def get_version():
    if args.version:
        return args.version
    # Aus backup.py lesen
    try:
        for f in Path(".").rglob("backup.py"):
            content = f.read_text(encoding="utf-8")
            for line in content.splitlines():
                if "APP_VERSION" in line and "=" in line:
                    return line.split("=")[1].strip().strip('"').strip("'")
    except Exception:
        pass
    return datetime.now().strftime("%Y.%m.%d")

VERSION = get_version()
print(f"\n🏗  NMG Analyse – Build v{VERSION}")
print("=" * 50)

# ── PyInstaller Build ─────────────────────────────────────────────────────────
def build_exe():
    print("\n[1/4] Bereinige alte Build-Artefakte...")
    for d in (DIST_DIR, BUILD_DIR):
        if d.exists():
            shutil.rmtree(d)

    print(f"[2/4] Baue {APP_NAME}.exe mit PyInstaller...")
    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--onefile",
        "--windowed",
        "--name", APP_NAME,
        "--clean",
    ]
    if ICON_FILE and Path(ICON_FILE).exists():
        cmd += ["--icon", ICON_FILE]
    # Alle .py-Dateien aus dem Paket einbinden
    for pyfile in Path("nmg_analyse").rglob("*.py"):
        cmd += ["--hidden-import", str(pyfile).replace("/","\\").replace("\\", ".").rstrip(".py")]
    cmd += [
        "--add-data", "assets;assets",
        "--add-data", "nmg_analyse;nmg_analyse",
        ENTRY_POINT,
    ]
    result = subprocess.run(cmd, capture_output=False)
    if result.returncode != 0:
        print("❌ PyInstaller fehlgeschlagen!")
        sys.exit(1)

    exe_path = DIST_DIR / f"{APP_NAME}.exe"
    if not exe_path.exists():
        print(f"❌ {exe_path} nicht gefunden!")
        sys.exit(1)
    print(f"✅ Build erfolgreich: {exe_path} ({exe_path.stat().st_size / 1024 / 1024:.1f} MB)")
    return exe_path

# ── Release-Paket erstellen ──────────────────────────────────────────────────
def create_release_package(exe_path):
    print(f"\n[3/4] Erstelle Release-Paket v{VERSION}...")
    RELEASE_DIR.mkdir(exist_ok=True)
    zip_name = RELEASE_DIR / f"{APP_NAME}_v{VERSION}"
    pkg_dir = RELEASE_DIR / f"{APP_NAME}_v{VERSION}"
    pkg_dir.mkdir(exist_ok=True)

    shutil.copy2(exe_path, pkg_dir / f"{APP_NAME}.exe")

    # release.json für Auto-Update
    release_info = {
        "version": VERSION,
        "name": f"NMG Analyse v{VERSION}",
        "date": datetime.now().strftime("%Y-%m-%d"),
        "download_url": f"https://github.com/{REPO}/releases/download/v{VERSION}/{APP_NAME}_v{VERSION}.zip",
        "changelog": f"Release v{VERSION}"
    }
    (pkg_dir / "release.json").write_text(json.dumps(release_info, indent=2, ensure_ascii=False), encoding="utf-8")

    # ZIP erstellen
    archive = shutil.make_archive(str(zip_name), "zip", RELEASE_DIR, f"{APP_NAME}_v{VERSION}")
    shutil.rmtree(pkg_dir)
    print(f"✅ Release-Paket: {archive}")
    return Path(archive), release_info

# ── GitHub Release ────────────────────────────────────────────────────────────
def upload_to_github(archive_path, release_info):
    try:
        import requests
    except ImportError:
        print("⚠  requests nicht installiert: pip install requests")
        return

    token = os.getenv("GITHUB_TOKEN", "")
    if not token:
        try:
            from pathlib import Path as P
            env = {l.split("=")[0].strip(): l.split("=")[1].strip()
                   for l in P(".env").read_text().splitlines() if "=" in l and not l.startswith("#")}
            token = env.get("GITHUB_TOKEN", "")
        except Exception:
            pass

    if not token:
        print("⚠  Kein GitHub-Token – manuell auf GitHub hochladen:")
        print(f"   {archive_path}")
        return

    headers = {"Authorization": f"token {token}", "Accept": "application/vnd.github+json"}
    tag = f"v{VERSION}"

    print(f"\n[4/4] Erstelle GitHub Release {tag}...")
    create_url = f"https://api.github.com/repos/{REPO}/releases"
    payload = {"tag_name": tag, "name": release_info["name"],
                "body": release_info["changelog"], "draft": False, "prerelease": False}
    resp = requests.post(create_url, json=payload, headers=headers)
    if resp.status_code not in (200, 201):
        print(f"❌ Release-Erstellung fehlgeschlagen: {resp.text}")
        return

    upload_url = resp.json()["upload_url"].split("{")[0]
    print(f"📤 Lade {archive_path.name} hoch...")
    with open(archive_path, "rb") as f:
        up_resp = requests.post(
            f"{upload_url}?name={archive_path.name}",
            data=f,
            headers={**headers, "Content-Type": "application/zip"}
        )
    if up_resp.status_code in (200, 201):
        print(f"✅ Hochgeladen: {up_resp.json().get('browser_download_url','')}")
    else:
        print(f"❌ Upload fehlgeschlagen: {up_resp.text}")

# ── Hauptprogramm ─────────────────────────────────────────────────────────────
if __name__ == "__main__":
    exe = build_exe()
    archive, info = create_release_package(exe)

    if not args.build_only:
        print("\n📋 Nächste Schritte:")
        print(f"   1. {archive} auf GitHub hochladen")
        print(f"   2. release.json im Repo aktualisieren (Version: {VERSION})")
        print(f"   3. Git-Tag erstellen: git tag v{VERSION} && git push --tags")

    if args.upload:
        upload_to_github(archive, info)

    print(f"\n✅ Build abgeschlossen: v{VERSION}")