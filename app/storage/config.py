"""Konfiguration der Speicher-Schicht (Modus-Aufloesung).

Bewusst getrennt von der globalen app/config.py gehalten, damit die laufende
App in diesem Schritt unberuehrt bleibt. Liefert ausschliesslich den
gewuenschten Speicher-Modus; Default ist immer LOCAL.

Aufloesungs-Reihenfolge (erste Quelle gewinnt):
    1. Umgebungsvariable NMGONE_STORAGE_MODE  (local|cloud|sharepoint)
    2. install_config.json -> "storage_mode"  (vom Setup geschrieben)
    3. Default: local

Damit bleibt der lokale Betrieb der Standard. Cloud/SharePoint werden erst
aktiv, wenn sie explizit gesetzt UND implementiert sind.
"""
from __future__ import annotations

import json
import os
from pathlib import Path

from .storage_provider import StorageMode


def _install_config_path() -> Path | None:
    try:
        from ..config import BASE_DIR
        return Path(BASE_DIR) / "install_config.json"
    except Exception:
        return None


def get_storage_mode() -> StorageMode:
    env = os.environ.get("NMGONE_STORAGE_MODE", "").strip()
    if env:
        return StorageMode.from_string(env)

    cfg = _install_config_path()
    if cfg and cfg.exists():
        try:
            payload = json.loads(cfg.read_text(encoding="utf-8"))
            raw = str(payload.get("storage_mode", "")).strip()
            if raw:
                return StorageMode.from_string(raw)
        except Exception:
            pass

    return StorageMode.LOCAL
