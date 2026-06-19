"""V1.1 SP8: Archivierung von gespeicherten Analysen.

User waehlt einen Datumsbereich, alle tbl_auswertungen-Eintraege darin
werden:
1. Als JSON (auswertungen.json + positionen.json) in ein ZIP geschrieben
2. Die zugehoerigen Excel-Dateien (ausgabedatei) wandern in den excel/-
   Unterordner im ZIP
3. Danach werden die DB-Zeilen geloescht; CASCADE auf
   tbl_auswertungspositionen ist via FOREIGN KEY definiert.

Das ZIP liegt unter BACKUP_DIR/analysen_archiv/ und hat einen
sprechenden Dateinamen mit Zeitraum.

Wiederherstellung: einzelne Excel aus dem ZIP entpacken und mit
_open_file oeffnen. Die DB-Zeilen kommen NICHT zurueck - das Archiv
ist eine "Cold Storage"-Loesung; archivierte Auswertungen tauchen
nicht mehr in Suche oder Produktanalyse auf, bis der User sie
manuell wieder importiert (zukuenftiges Feature).

Loeschen des ZIPs selbst ist Admin-only (Passwort wie beim
DB-Cleanup-Dialog).
"""

from __future__ import annotations

import io
import json
import re
import shutil
import sqlite3
import tempfile
import zipfile
from datetime import datetime
from pathlib import Path

from .config import BACKUP_DIR, DB_PATH


ARCHIV_DIR = BACKUP_DIR / "analysen_archiv"
_ARCHIV_VERSION = "1"  # Schema-Version, falls wir spaeter Format aendern.


def ensure_archiv_dir() -> Path:
    """Legt den Archiv-Ordner an, falls noch nicht da."""
    ARCHIV_DIR.mkdir(parents=True, exist_ok=True)
    return ARCHIV_DIR


def _safe_chunk(text: str, maxlen: int = 40) -> str:
    """Bereinigt einen String fuer Dateinamen."""
    cleaned = re.sub(r"[^A-Za-z0-9_-]+", "_", (text or "")).strip("_")
    return cleaned[:maxlen] if cleaned else "leer"


def zaehle_zeitraum(von: str, bis: str) -> int:
    """Wieviele Auswertungen liegen im Datumsbereich? Fuer Vorschau-Dialog."""
    von, bis = _normalize_zeitraum(von, bis)
    with sqlite3.connect(DB_PATH) as con:
        return int(con.execute(
            """SELECT COUNT(*) FROM tbl_auswertungen
               WHERE date(datum) >= date(?) AND date(datum) <= date(?)""",
            (von, bis),
        ).fetchone()[0])


def _normalize_zeitraum(von: str, bis: str) -> tuple[str, str]:
    """Akzeptiert YYYY-MM-DD oder DD.MM.YYYY und gibt YYYY-MM-DD zurueck."""
    def _to_iso(v: str) -> str:
        v = (v or "").strip()
        if not v:
            raise ValueError("Datum ist leer.")
        if re.match(r"^\d{4}-\d{2}-\d{2}$", v):
            return v
        if re.match(r"^\d{2}\.\d{2}\.\d{4}$", v):
            d, m, y = v.split(".")
            return f"{y}-{m}-{d}"
        raise ValueError(f"Datum '{v}' nicht erkannt (YYYY-MM-DD oder DD.MM.YYYY).")
    return _to_iso(von), _to_iso(bis)


def archiviere_zeitraum(von: str, bis: str) -> dict:
    """Archiviert alle Auswertungen im Datumsbereich.

    Schritte:
    1. Auswertungen + Positionen aus DB lesen
    2. Excel-Dateien einsammeln (falls ausgabedatei existiert)
    3. ZIP unter ARCHIV_DIR ablegen
    4. DB-Zeilen loeschen (Positionen via CASCADE oder explizit)

    Liefert dict mit Statistik:
        archiviert (int), excel_anzahl (int), zip_pfad (Path),
        groesse_bytes (int).
    """
    von, bis = _normalize_zeitraum(von, bis)
    ensure_archiv_dir()

    with sqlite3.connect(DB_PATH) as con:
        con.row_factory = sqlite3.Row
        auswertungen = [dict(r) for r in con.execute(
            """SELECT * FROM tbl_auswertungen
               WHERE date(datum) >= date(?) AND date(datum) <= date(?)
               ORDER BY datum""",
            (von, bis),
        ).fetchall()]
        if not auswertungen:
            return {"archiviert": 0, "excel_anzahl": 0, "zip_pfad": None,
                    "groesse_bytes": 0}

        ids = [r["id"] for r in auswertungen]
        placeholders = ",".join("?" * len(ids))
        positionen = [dict(r) for r in con.execute(
            f"SELECT * FROM tbl_auswertungspositionen WHERE auswertung_id IN ({placeholders})",
            ids,
        ).fetchall()]

    zeitstempel = datetime.now().strftime("%Y%m%d_%H%M%S")
    zip_name = f"Analysen_Archiv_{von}_bis_{bis}_{zeitstempel}.zip"
    zip_pfad = ARCHIV_DIR / zip_name

    excel_anzahl = 0
    metadaten = {
        "archiv_version": _ARCHIV_VERSION,
        "erstellt_am": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "zeitraum_von": von,
        "zeitraum_bis": bis,
        "auswertungen_count": len(auswertungen),
        "positionen_count": len(positionen),
    }

    with zipfile.ZipFile(zip_pfad, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("meta.json", json.dumps(metadaten, indent=2, ensure_ascii=False))
        zf.writestr("auswertungen.json", json.dumps(auswertungen, indent=2, ensure_ascii=False, default=str))
        zf.writestr("positionen.json", json.dumps(positionen, indent=2, ensure_ascii=False, default=str))

        for aw in auswertungen:
            datei = aw.get("ausgabedatei") or ""
            if not datei:
                continue
            quelle = Path(datei)
            if not quelle.exists():
                continue
            # Zielname im ZIP: excel/<id>_<apotheke>.xlsx
            zielname = f"excel/{aw['id']:06d}_{_safe_chunk(aw.get('apotheke') or 'auswertung')}.xlsx"
            try:
                zf.write(quelle, arcname=zielname)
                excel_anzahl += 1
            except OSError:
                # Datei gesperrt o.ae. - nicht fatal, der Archiv-Index
                # verzeichnet sie trotzdem; spaeter kann ein Restore-Pfad
                # den Verlust melden.
                pass

    # Erst nach erfolgreichem ZIP-Schreiben die DB-Zeilen entfernen.
    with sqlite3.connect(DB_PATH) as con:
        con.execute("PRAGMA foreign_keys = ON")
        con.execute(
            f"DELETE FROM tbl_auswertungspositionen WHERE auswertung_id IN ({placeholders})",
            ids,
        )
        con.execute(
            f"DELETE FROM tbl_auswertungen WHERE id IN ({placeholders})",
            ids,
        )
        con.commit()

    return {
        "archiviert": len(auswertungen),
        "excel_anzahl": excel_anzahl,
        "zip_pfad": zip_pfad,
        "groesse_bytes": zip_pfad.stat().st_size,
    }


def liste_archive() -> list[dict]:
    """Liste aller ZIPs im Archiv-Ordner mit Metadaten."""
    ensure_archiv_dir()
    out = []
    for zip_pfad in sorted(ARCHIV_DIR.glob("*.zip"), reverse=True):
        info = {
            "pfad": zip_pfad,
            "name": zip_pfad.name,
            "groesse_bytes": zip_pfad.stat().st_size,
            "erstellt_am": "",
            "zeitraum_von": "",
            "zeitraum_bis": "",
            "auswertungen_count": 0,
        }
        try:
            with zipfile.ZipFile(zip_pfad, "r") as zf:
                with zf.open("meta.json") as mf:
                    meta = json.loads(mf.read().decode("utf-8"))
                    info.update({
                        "erstellt_am": meta.get("erstellt_am", ""),
                        "zeitraum_von": meta.get("zeitraum_von", ""),
                        "zeitraum_bis": meta.get("zeitraum_bis", ""),
                        "auswertungen_count": int(meta.get("auswertungen_count", 0)),
                    })
        except (zipfile.BadZipFile, KeyError, json.JSONDecodeError):
            pass
        out.append(info)
    return out


def liste_analysen_im_archiv(zip_pfad: Path) -> list[dict]:
    """Liefert die Auswertungs-Header (id, datum, apotheke, etc.) aus dem ZIP."""
    with zipfile.ZipFile(zip_pfad, "r") as zf:
        with zf.open("auswertungen.json") as af:
            return json.loads(af.read().decode("utf-8"))


def excel_aus_archiv(zip_pfad: Path, auswertung_id: int) -> Path | None:
    """Extrahiert die Excel zu auswertung_id in einen Temp-Ordner und
    liefert den Pfad zurueck. None wenn keine Excel im ZIP fuer diese ID.
    """
    with zipfile.ZipFile(zip_pfad, "r") as zf:
        prefix = f"excel/{auswertung_id:06d}_"
        kandidaten = [n for n in zf.namelist() if n.startswith(prefix)]
        if not kandidaten:
            return None
        name = kandidaten[0]
        ziel_dir = Path(tempfile.gettempdir()) / "NMGone_Archiv"
        ziel_dir.mkdir(parents=True, exist_ok=True)
        ziel = ziel_dir / Path(name).name
        with zf.open(name) as src, ziel.open("wb") as dst:
            shutil.copyfileobj(src, dst)
        return ziel


def loesche_archiv(zip_pfad: Path) -> None:
    """Endgueltig loeschen. Aufrufer ist verantwortlich fuer die
    Admin-Passwort-Abfrage.
    """
    p = Path(zip_pfad)
    if p.exists() and p.is_file() and p.suffix.lower() == ".zip":
        p.unlink()
