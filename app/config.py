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
    1. install_config.json im Programmordner (vom Setup geschrieben)
    2. Umgebungsvariable NMGONE_DATA_ROOT
    3. Installiertes .exe auf Windows: C:/ProgramData/NMGone
    4. Dev-Modus (python start.py): lokaler data/-Ordner im Repo

    SP7: Skelett fuer kuenftige SharePoint/OneDrive-Umleitung. Bewusst
    auskommentiert - wird in einem spaeteren SP aktiviert, wenn das
    Feature freigegeben wird. Dann liest die App den Pfad aus
    install_config.json -> "data_root_override" (oder aus der meta-Tabelle)
    und routet darueber. Bis dahin bleibt alles wie heute.
    """
    cfg = BASE_DIR / "install_config.json"
    if cfg.exists():
        try:
            payload = json.loads(cfg.read_text(encoding="utf-8"))
            # TODO SP-future: zuerst override pruefen, wenn freigegeben
            # override = str(payload.get("data_root_override", "")).strip()
            # if override and payload.get("feature_custom_data_root_enabled"):
            #     return Path(os.path.expandvars(override)).expanduser()
            raw = str(payload.get("data_root", "")).strip()
            if raw:
                return Path(os.path.expandvars(raw)).expanduser()
        except Exception:
            pass

    env = os.environ.get("NMGONE_DATA_ROOT", "").strip()
    if env:
        return Path(os.path.expandvars(env)).expanduser()

    # Nur die gebaute .exe nutzt ProgramData. Dev-Modus arbeitet lokal,
    # damit Code-Experimente nicht in produktive Mitarbeiter-Daten greifen.
    is_frozen = getattr(sys, "frozen", False)
    if _is_windows() and is_frozen and not os.environ.get("NMGONE_PORTABLE"):
        program_data = os.environ.get("ProgramData", r"C:/ProgramData")
        return Path(program_data) / "NMGone"

    return None


USERDATA_ROOT = _load_install_data_root()
if USERDATA_ROOT:
    USERDATA_ROOT.mkdir(parents=True, exist_ok=True)
    DATA_DIR = USERDATA_ROOT / "data"
    OUTPUT_DIR = USERDATA_ROOT / "ausgaben"
    SAVED_ANALYSES_DIR = USERDATA_ROOT / "gespeicherte_analysen"
    IMPORT_DIR = USERDATA_ROOT / "importierte_analysen"
    BACKUP_DIR = USERDATA_ROOT / "backups"
    UPDATE_DIR = USERDATA_ROOT / "updates"
    LOG_DIR = USERDATA_ROOT / "logs"
else:
    DATA_DIR = BASE_DIR / "data"
    OUTPUT_DIR = BASE_DIR / "ausgaben"
    SAVED_ANALYSES_DIR = BASE_DIR / "gespeicherte_analysen"
    IMPORT_DIR = BASE_DIR / "importierte_analysen"
    BACKUP_DIR = BASE_DIR / "backups"
    UPDATE_DIR = BASE_DIR / "updates"
    LOG_DIR = BASE_DIR / "logs"


def jahr_quartal_pfad(base, kategorie: str = "", dt=None):
    """V1.1 SP12: Liefert einen Ausgabe-Pfad nach Schema
        base/[kategorie]/<Jahr>/Q<n>/
    Erstellt das Verzeichnis falls noetig.

    Wenn kategorie leer ist, wird der Kategorie-Schritt uebersprungen
    (z.B. fuer base = importierte_analysen/PK ist die Kategorie bereits
    im base-Pfad enthalten).
    """
    from datetime import datetime as _dt
    if dt is None:
        dt = _dt.now()
    quartal = (dt.month - 1) // 3 + 1
    p = Path(base)
    if kategorie:
        p = p / kategorie
    p = p / str(dt.year) / f"Q{quartal}"
    p.mkdir(parents=True, exist_ok=True)
    return p

ASSETS_DIR = BASE_DIR / "assets"
VERSION_FILE = BASE_DIR / "version.json"


def _select_db_filename() -> str:
    """Waehlt die zu nutzende SQLite-Datei.

    Umschalten ohne Code-Aenderung ueber Umgebungsvariablen:
      * NMGONE_DEMO=1            -> data/nmg_demodatenbank.sqlite (Demo)
      * NMGONE_DB_FILE=<name>    -> data/<name> (beliebige Datei, hat Vorrang)
    Ohne beides: die normale Produktiv-Datenbank.
    """
    explicit = os.environ.get("NMGONE_DB_FILE", "").strip()
    if explicit:
        return explicit
    if os.environ.get("NMGONE_DEMO", "").strip() in ("1", "true", "True", "ja"):
        return "nmg_demodatenbank.sqlite"
    # Demo-Paket: liegt KEINE Produktiv-DB, aber eine Demo-DB vor, laeuft das
    # Programm automatisch im Demo-Modus (z.B. das ausgelieferte NMGone-Demo).
    if not (DATA_DIR / "nmg_startdatenbank.sqlite").exists() \
            and (DATA_DIR / "nmg_demodatenbank.sqlite").exists():
        return "nmg_demodatenbank.sqlite"
    return "nmg_startdatenbank.sqlite"


DB_FILENAME = _select_db_filename()
DB_PATH = DATA_DIR / DB_FILENAME
IS_DEMO_DB = DB_FILENAME != "nmg_startdatenbank.sqlite"

# Interner Programmname. Im Demo-Modus heisst das Programm ueberall sichtbar
# "NMGone Demo", damit Demo- und Echtbetrieb nie verwechselt werden.
APP_NAME = "NMGone Demo" if IS_DEMO_DB else "NMGone"
DEMO_SUFFIX = "  [DEMO]" if IS_DEMO_DB else ""


def fenstertitel(basis: str) -> str:
    """Haengt im Demo-Modus eine sichtbare DEMO-Markierung an einen Titel."""
    return f"{basis}{DEMO_SUFFIX}"

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
    for folder in [DATA_DIR, OUTPUT_DIR, SAVED_ANALYSES_DIR, IMPORT_DIR, BACKUP_DIR, UPDATE_DIR, LOG_DIR]:
        folder.mkdir(parents=True, exist_ok=True)

    # Demo-/Sondermodus: Datei wird NICHT automatisch erzeugt. Sie muss
    # vorhanden sein (z.B. via scripts/seed_demodaten.py). Klarer Hinweis
    # statt stillem Kopieren der Produktivdaten.
    if IS_DEMO_DB:
        if not (DB_PATH.exists() and DB_PATH.stat().st_size > 1024):
            raise FileNotFoundError(
                f"Demo-/Sonderdatenbank fehlt: {DB_PATH}\n"
                f"Bitte zuerst erzeugen, z.B.:  python scripts/seed_demodaten.py"
            )
        return

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


def _migrate_zf_to_zw_dirs() -> None:
    """ZF -> ZW: Vorhandene Unterordner 'ZF' in 'ZW' umbenennen (die Datenquelle
    Zukunftswerk wurde von 'ZF' auf 'ZW' vereinheitlicht). Non-destruktiv:
    existiert 'ZW' schon, werden die Inhalte hineinverschoben statt ueberschrieben.
    """
    for base in (SAVED_ANALYSES_DIR, IMPORT_DIR):
        old = base / "ZF"
        new = base / "ZW"
        try:
            if not old.exists():
                continue
            if not new.exists():
                old.rename(new)
            else:
                for item in old.iterdir():
                    target = new / item.name
                    if not target.exists():
                        shutil.move(str(item), str(target))
                try:
                    old.rmdir()  # nur entfernen, wenn restlos geleert
                except OSError:
                    pass
        except Exception:
            # Best-Effort: schlaegt die Migration fehl, bleibt der Alt-Ordner bestehen.
            pass


_copy_seed_data_if_needed()
_migrate_zf_to_zw_dirs()
