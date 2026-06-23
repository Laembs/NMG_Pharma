"""Speicher-Schicht fuer NMGone.

Oeffentliche API. Bestehender Code bleibt unberuehrt; wer die Schicht nutzen
moechte, holt sich ueber get_storage_provider() das passende Backend.

    from app.storage import get_storage_provider
    sp = get_storage_provider()            # heute: LocalStorageProvider
    sp.write_bytes("ausgaben/test.txt", b"hallo")

Der Modus kommt aus app/storage/config.py (Default: local). Cloud/SharePoint
sind derzeit dokumentierte Platzhalter und werfen NotImplementedError, bis sie
implementiert sind - der lokale Betrieb ist davon nicht betroffen.
"""
from __future__ import annotations

from .storage_provider import StorageMode, StorageProvider
from .local_storage_provider import LocalStorageProvider
from .cloud_storage_provider import CloudStorageProvider
from .sharepoint_storage_provider import SharePointStorageProvider
from .config import get_storage_mode

__all__ = [
    "StorageMode",
    "StorageProvider",
    "LocalStorageProvider",
    "CloudStorageProvider",
    "SharePointStorageProvider",
    "get_storage_mode",
    "get_storage_provider",
]

# Einfacher Prozess-Cache, damit nicht bei jedem Aufruf neu gebaut wird.
_PROVIDER: StorageProvider | None = None


def get_storage_provider(mode: StorageMode | str | None = None,
                         force_new: bool = False) -> StorageProvider:
    """Liefert den konfigurierten Speicher-Provider.

    mode      : optionaler Override; sonst aus get_storage_mode().
    force_new : Cache umgehen (z.B. fuer Tests).
    """
    global _PROVIDER

    if mode is None and not force_new and _PROVIDER is not None:
        return _PROVIDER

    selected = (
        mode if isinstance(mode, StorageMode)
        else StorageMode.from_string(mode) if mode is not None
        else get_storage_mode()
    )

    if selected is StorageMode.CLOUD:
        provider: StorageProvider = CloudStorageProvider()
    elif selected is StorageMode.SHAREPOINT:
        provider = SharePointStorageProvider()
    else:
        provider = LocalStorageProvider()

    if mode is None and not force_new:
        _PROVIDER = provider
    return provider
