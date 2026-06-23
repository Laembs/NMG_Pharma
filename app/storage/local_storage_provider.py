"""Lokaler Speicher-Provider (heute aktiv).

Bildet logische Namen auf das lokale Dateisystem unterhalb eines Wurzel-
Verzeichnisses ab. Die Wurzel kommt standardmaessig aus der bestehenden
app/config.py (USERDATA_ROOT bzw. BASE_DIR), damit sich am tatsaechlichen
Speicherort gegenueber heute NICHTS aendert.

Dieser Provider liefert ueber resolve() echte Path-Objekte zurueck. Dadurch
kann bestehender Code, der mit Pfaden arbeitet, schrittweise und ohne
Verhaltensaenderung auf die Storage-Schicht umgestellt werden.
"""
from __future__ import annotations

from pathlib import Path
from typing import Iterable

from .storage_provider import StorageMode, StorageProvider


def _default_root() -> Path:
    """Wurzelverzeichnis aus der bestehenden Konfiguration ableiten.

    Lazy import, damit dieses Modul auch isoliert (z.B. im Test) ohne die
    vollstaendige App-Initialisierung importierbar bleibt.
    """
    try:
        from ..config import USERDATA_ROOT, BASE_DIR
        return Path(USERDATA_ROOT or BASE_DIR)
    except Exception:
        return Path.cwd()


class LocalStorageProvider(StorageProvider):
    mode = StorageMode.LOCAL

    def __init__(self, root: Path | str | None = None):
        self.root = Path(root) if root is not None else _default_root()

    def _full(self, name: str) -> Path:
        # Logische Namen sind '/'-getrennt und immer relativ zur Wurzel.
        rel = Path(str(name).replace("\\", "/").lstrip("/"))
        if rel.is_absolute() or ".." in rel.parts:
            raise ValueError(f"Ungueltiger logischer Name: {name!r}")
        return self.root / rel

    def exists(self, name: str) -> bool:
        return self._full(name).exists()

    def read_bytes(self, name: str) -> bytes:
        return self._full(name).read_bytes()

    def write_bytes(self, name: str, data: bytes) -> None:
        target = self._full(name)
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_bytes(data)

    def delete(self, name: str) -> None:
        target = self._full(name)
        if target.exists():
            target.unlink()

    def list(self, prefix: str = "") -> Iterable[str]:
        base = self._full(prefix) if prefix else self.root
        if not base.exists():
            return []
        results = []
        for p in base.rglob("*"):
            if p.is_file():
                results.append(p.relative_to(self.root).as_posix())
        return results

    def resolve(self, name: str) -> Path:
        return self._full(name)

    def supports_local_path(self) -> bool:
        return True
