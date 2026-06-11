from __future__ import annotations

import json
import os
import shutil
import sys
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent.parent


def _is_windows() -> bool:
    return os.name == "nt"


def _load_install_data_root() -> Path | None:
    """Ermittelt den Datenordner.

    Prioritaet:
    1. install_config.json im Programmordner
    2. Umgebungsvariable NMG_ANALYSE_DATA_ROOT
    3. Windows: C:/ProgramData/NMG Analyse
    4. portable Nutzung: Programmordner
    """
    cfg = BASE_DIR / "install_config.json"
    if cfg.exists():
        try:
            payload = json.loads(cfg.read_text(encoding="utf-8"))
            raw = str(payload.get("data_root", "")).strip()
            if raw:
                return Path(os.path.expandvars(raw)).expanduser()
        except Exception:
            pass

    env = os.environ.get("NMG_ANALYSE_DATA_ROOT", "").strip()
    if env:
        return Path(os.path.expandvars(env)).expanduser()

    if _is_windows() and not os.environ.get("NMG_ANALYSE_PORTABLE"):
        program_data = os.environ.get("ProgramData", r"C:/ProgramData")
        return Path(program_data) / "NMG Analyse"

    return None


USERDATA_ROOT = _load_install_data_root()
if USERDATA_ROOT:
    USERDATA_ROOT.mkdir(parents=True, exist_ok=True)
    DATA_DIR = USERDATA_ROOT / "data"
    OUTPUT_DIR = USERDATA_ROOT / "ausgaben"
    SAVED_ANALYSES_DIR = USERDATA_ROOT / "gespeicherte_analysen"
    BACKUP_DIR = USERDATA_ROOT / "backups"
    UPDATE_DIR = USERDATA_ROOT / "updates"
    LOG_DIR = USERDATA_ROOT / "logs"
else:
    DATA_DIR = BASE_DIR / "data"
    OUTPUT_DIR = BASE_DIR / "ausgaben"
    SAVED_ANALYSES_DIR = BASE_DIR / "gespeicherte_analysen"
    BACKUP_DIR = BASE_DIR / "backups"
    UPDATE_DIR = BASE_DIR / "updates"
    LOG_DIR = BASE_DIR / "logs"

ASSETS_DIR = BASE_DIR / "assets"
VERSION_FILE = BASE_DIR / "version.json"
DB_PATH = DATA_DIR / "nmg_startdatenbank.sqlite"

REFERENCE_XLSX = DATA_DIR / "NMG_Hochpreiser_Vollversion_1_3_NMG_STAMM_APU_TAXEK.xlsx"
LINDEN_REFERENCE_XLSX = DATA_DIR / "Linden_Apo_Auengrund_Referenz.xlsx"
ROSEN_REFERENCE_XLSX = DATA_DIR / "Rosen_apo_Forst_Referenz.xlsx"
SONNEN_REFERENCE_XLSX = DATA_DIR / "Sonnen_Apotheke_Erfurt_Referenz.xlsx"
HANDCHECKED_DIR = DATA_DIR / "lerndaten_hand"
HISTORICAL_ANALYSIS_DIR = DATA_DIR / "historische_auswertungen"


def _runtime_seed_dirs() -> list[Path]:
    dirs = []
    meipass = Path(getattr(sys, "_MEIPASS", BASE_DIR))
    for candidate in [BASE_DIR / "data", meipass / "data"]:
        if candidate.exists() and candidate not in dirs:
            dirs.append(candidate)
    return dirs


def _copy_seed_data_if_needed() -> None:
    """Kopiert die ausgelieferten Startdaten in den Nutzdatenordner.

    Bestehende Datenbanken und Nutzerdaten werden nicht ueberschrieben.
    Diese Recovery-Version liefert den Lernstand aus Prototyp 2.6 als Startbestand mit.
    """
    for folder in [DATA_DIR, OUTPUT_DIR, SAVED_ANALYSES_DIR, BACKUP_DIR, UPDATE_DIR, LOG_DIR]:
        folder.mkdir(parents=True, exist_ok=True)

    if DB_PATH.exists() and DB_PATH.stat().st_size > 1024:
        return

    for src in _runtime_seed_dirs():
        seed_db = src / "nmg_startdatenbank.sqlite"
        if not seed_db.exists():
            continue
        DATA_DIR.mkdir(parents=True, exist_ok=True)
        for item in src.iterdir():
            dest = DATA_DIR / item.name
            if item.is_dir():
                if not dest.exists():
                    shutil.copytree(item, dest)
            else:
                if not dest.exists() or (item.name == "nmg_startdatenbank.sqlite" and dest.stat().st_size < 1024):
                    shutil.copy2(item, dest)
        return


_copy_seed_data_if_needed()
