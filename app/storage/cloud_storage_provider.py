"""Cloud-Speicher-Provider (PLATZHALTER / Schnittstelle).

Dieser Provider ist bewusst noch nicht implementiert. Er definiert nur die
geplante Schnittstelle fuer einen spaeteren Cloud-/Objektspeicher-Betrieb
(z.B. eigenes Backend-REST-API, S3-kompatibler Objektspeicher o.ae.).

Designziele fuer die spaetere Umsetzung:
    - Byte-/Text-basiert (KEINE echten Dateipfade -> supports_local_path=False).
    - Authentifizierung/Token-Handling gekapselt im Provider.
    - Optionaler lokaler Cache, lokaler Fallback bleibt jederzeit moeglich.
    - Abhaengigkeiten (requests/SDK) NUR per lazy import, damit der lokale
      Betrieb keine zusaetzlichen Pakete benoetigt.
"""
from __future__ import annotations

from typing import Iterable

from .storage_provider import StorageMode, StorageProvider

_NOT_READY = (
    "CloudStorageProvider ist noch nicht implementiert. "
    "Speicher-Modus 'cloud' ist derzeit ein Platzhalter; "
    "bitte STORAGE_MODE auf 'local' belassen."
)


class CloudStorageProvider(StorageProvider):
    mode = StorageMode.CLOUD

    def __init__(self, **options):
        # Konfiguration (Endpoint, Bucket, Credentials-Quelle) wird spaeter
        # hier entgegengenommen. Vorerst nur gemerkt, kein Verbindungsaufbau.
        self.options = options

    def exists(self, name: str) -> bool:
        raise NotImplementedError(_NOT_READY)

    def read_bytes(self, name: str) -> bytes:
        raise NotImplementedError(_NOT_READY)

    def write_bytes(self, name: str, data: bytes) -> None:
        raise NotImplementedError(_NOT_READY)

    def delete(self, name: str) -> None:
        raise NotImplementedError(_NOT_READY)

    def list(self, prefix: str = "") -> Iterable[str]:
        raise NotImplementedError(_NOT_READY)
