"""Abstrakte Speicher-Schnittstelle fuer NMGone.

Ziel: Die Programmlogik soll Datei-Artefakte (Excel-Exporte, Importe,
Vorlagen, Backups) ueber *logische Namen* ansprechen statt ueber feste
Dateipfade. Dadurch kann derselbe Code spaeter unveraendert gegen lokalen
Speicher, einen Cloud-Objektspeicher oder SharePoint/OneDrive laufen.

WICHTIG (Schritt 1 der Umstellung):
    Diese Schicht ist rein additiv und wird zunaechst NICHT in bestehenden
    Code eingebunden. NMGone laeuft lokal unveraendert weiter. Die konkreten
    Provider werden schrittweise verdrahtet.

Abgrenzung:
    Diese Schicht kapselt DATEI-Artefakte. Die SQLite-Datenbank bleibt
    vorerst dateibasiert (Ort wird weiterhin ueber app/config.py bzw.
    install_config.json bestimmt). Echter Cloud-Mehrbenutzerbetrieb der DB
    ist ein spaeterer, eigener Schritt (Server-DB/REST) und laeuft NICHT
    ueber diese Datei-Schicht.
"""
from __future__ import annotations

from abc import ABC, abstractmethod
from enum import Enum
from pathlib import Path
from typing import Iterable


class StorageMode(str, Enum):
    """Betriebsart der Speicherschicht. Default ist immer LOCAL."""

    LOCAL = "local"
    CLOUD = "cloud"
    SHAREPOINT = "sharepoint"

    @classmethod
    def from_string(cls, value: str | None) -> "StorageMode":
        text = (value or "").strip().lower()
        for mode in cls:
            if mode.value == text:
                return mode
        return cls.LOCAL


class StorageProvider(ABC):
    """Gemeinsame Schnittstelle aller Speicher-Backends.

    Logische Namen sind relative, mit '/' getrennte Pfade unterhalb eines
    Bereichs, z.B. "ausgaben/2026/Q2/report.xlsx". Die Provider bilden diese
    Namen intern auf ihren jeweiligen Speicher ab.

    Hinweis zu resolve(): Nur lokal-basierte Provider koennen einen echten
    Dateisystempfad liefern. Cloud-/SharePoint-Provider duerfen hier
    NotImplementedError werfen - Aufrufer, die mit beliebigen Backends
    funktionieren sollen, muessen byte-/textbasiert arbeiten.
    """

    mode: StorageMode = StorageMode.LOCAL

    # --- Lesen -----------------------------------------------------------
    @abstractmethod
    def exists(self, name: str) -> bool:
        """True, wenn unter dem logischen Namen ein Objekt existiert."""

    @abstractmethod
    def read_bytes(self, name: str) -> bytes:
        """Roh-Bytes lesen. FileNotFoundError, wenn nicht vorhanden."""

    def read_text(self, name: str, encoding: str = "utf-8") -> str:
        return self.read_bytes(name).decode(encoding)

    # --- Schreiben -------------------------------------------------------
    @abstractmethod
    def write_bytes(self, name: str, data: bytes) -> None:
        """Roh-Bytes schreiben/ueberschreiben (Zielbereich anlegen)."""

    def write_text(self, name: str, text: str, encoding: str = "utf-8") -> None:
        self.write_bytes(name, text.encode(encoding))

    @abstractmethod
    def delete(self, name: str) -> None:
        """Objekt loeschen. Kein Fehler, wenn es nicht existiert."""

    @abstractmethod
    def list(self, prefix: str = "") -> Iterable[str]:
        """Logische Namen unterhalb eines Praefix auflisten."""

    # --- Lokaler Brueckenkopf -------------------------------------------
    def resolve(self, name: str) -> Path:
        """Echten Dateisystempfad zum logischen Namen liefern.

        Nur fuer lokal-basierte Provider sinnvoll. Erlaubt es, bestehenden
        Code (der mit echten Pfaden arbeitet) schrittweise zu migrieren,
        ohne ihn sofort byte-basiert umzuschreiben.
        """
        raise NotImplementedError(
            f"{type(self).__name__} liefert keinen lokalen Pfad (mode={self.mode})."
        )

    def supports_local_path(self) -> bool:
        """True, wenn resolve() einen echten Pfad zurueckgibt."""
        return False
