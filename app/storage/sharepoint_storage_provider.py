"""SharePoint-/OneDrive-Speicher-Provider (PLATZHALTER / Schnittstelle).

Noch nicht implementiert. Definiert die geplante Schnittstelle fuer den
spaeteren Betrieb gegen SharePoint/OneDrive ueber Microsoft Graph.

Wichtige Vorgaben fuer die spaetere Umsetzung:
    - KEINE harte Microsoft-Abhaengigkeit: msal / Office365-REST / requests
      werden ausschliesslich per lazy import innerhalb der Methoden geladen.
      Lokale Nutzer ohne diese Pakete sind nicht betroffen.
    - Datei-Locking/Konflikte beachten: Schreibzugriffe ueber Graph muessen
      ETag-/Versions-Pruefung nutzen (If-Match), um Ueberschreiben bei
      gleichzeitigem Zugriff zu erkennen. Bei Konflikt: klar definierter
      Fehler statt stillem Datenverlust.
    - Robuste Synchronisation: Schreiben moeglichst atomar (Upload in temp +
      Umbenennen) und mit Wiederholung bei transienten Fehlern.
    - Lokaler Fallback bleibt erhalten: bei Nichterreichbarkeit kann der
      Aufrufer auf den LocalStorageProvider zurueckfallen.
    - Byte-/Text-basiert -> supports_local_path=False.
"""
from __future__ import annotations

from typing import Iterable

from .storage_provider import StorageMode, StorageProvider

_NOT_READY = (
    "SharePointStorageProvider ist noch nicht implementiert. "
    "Speicher-Modus 'sharepoint' ist derzeit ein Platzhalter; "
    "bitte STORAGE_MODE auf 'local' belassen."
)


class SharePointStorageProvider(StorageProvider):
    mode = StorageMode.SHAREPOINT

    def __init__(self, **options):
        # Spaeter: site_id / drive_id / Ordner, Auth-Quelle (msal). Hier noch
        # kein Verbindungsaufbau und kein Import von Microsoft-Bibliotheken.
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
