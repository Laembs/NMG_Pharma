from __future__ import annotations

import sqlite3
from datetime import datetime
from pathlib import Path

from .config import DB_PATH
from .protocol_manager import PROTOCOL_ROOT, ensure_protocol_dirs, log_event


def _stamp() -> str:
    return datetime.now().strftime("%Y%m%d_%H%M%S")


def _count(con: sqlite3.Connection, sql: str, params: tuple = ()) -> int:
    try:
        row = con.execute(sql, params).fetchone()
        return int(row[0] or 0) if row else 0
    except Exception:
        return -1


def _table_exists(con: sqlite3.Connection, table: str) -> bool:
    try:
        return con.execute("SELECT name FROM sqlite_master WHERE type='table' AND name=?", (table,)).fetchone() is not None
    except Exception:
        return False


def run_database_health_check() -> dict:
    """Prüft die Datenbank ohne Daten zu verändern und schreibt ein Prüfprotokoll."""
    ensure_protocol_dirs()
    out_dir = PROTOCOL_ROOT / "Datenbankcheck"
    out_dir.mkdir(parents=True, exist_ok=True)
    out = out_dir / f"Datenbankcheck_{_stamp()}.log"

    result = {
        "status": "OK",
        "warnings": [],
        "errors": [],
        "report": out,
    }

    lines: list[str] = []
    lines.append("NMGone - Datenbank-Gesundheitscheck")
    lines.append("=" * 90)
    lines.append(f"Erstellt: {datetime.now().strftime('%d.%m.%Y %H:%M:%S')}")
    lines.append(f"Datenbank: {DB_PATH}")
    lines.append("")

    if not Path(DB_PATH).exists():
        result["status"] = "FEHLER"
        result["errors"].append("Datenbankdatei wurde nicht gefunden.")
        lines.append("FEHLER: Datenbankdatei wurde nicht gefunden.")
        out.write_text("\n".join(lines) + "\n", encoding="utf-8")
        return result

    try:
        lines.append(f"Dateigröße: {Path(DB_PATH).stat().st_size} Bytes")
    except Exception:
        pass

    required_tables = [
        "meta",
        "tbl_nmg_stamm",
        "tbl_austauschdatenbank",
        "tbl_lernvorschlaege",
        "tbl_lernhistorie",
        "tbl_pzn_basisdaten",
        "tbl_hersteller_lern",
        "tbl_auswertungen",
        "tbl_auswertungspositionen",
        "tbl_import_log",
    ]

    try:
        with sqlite3.connect(DB_PATH) as con:
            con.row_factory = sqlite3.Row
            lines.append("")
            lines.append("SQLite Integritätsprüfung")
            lines.append("-" * 90)
            try:
                integrity = con.execute("PRAGMA integrity_check").fetchone()[0]
            except Exception as exc:
                integrity = f"Fehler: {exc}"
            lines.append(f"PRAGMA integrity_check: {integrity}")
            if str(integrity).lower() != "ok":
                result["status"] = "FEHLER"
                result["errors"].append(f"SQLite integrity_check meldet: {integrity}")

            lines.append("")
            lines.append("Pflichttabellen")
            lines.append("-" * 90)
            for table in required_tables:
                exists = _table_exists(con, table)
                lines.append(f"{table}: {'OK' if exists else 'FEHLT'}")
                if not exists:
                    result["warnings"].append(f"Tabelle fehlt: {table}")

            lines.append("")
            lines.append("Datensatz-Zählung")
            lines.append("-" * 90)
            table_names = [r[0] for r in con.execute("SELECT name FROM sqlite_master WHERE type='table' ORDER BY name").fetchall() if not str(r[0]).startswith("sqlite_")]
            for table in table_names:
                c = _count(con, f'SELECT COUNT(*) FROM "{table}"')
                lines.append(f"{table}: {c}")

            checks = []
            if _table_exists(con, "tbl_auswertungspositionen"):
                checks.append(("Auswertungspositionen ohne PZN", _count(con, "SELECT COUNT(*) FROM tbl_auswertungspositionen WHERE pzn IS NULL OR TRIM(pzn)=''")))
                checks.append(("Auswertungspositionen ohne Kopf-Auswertung", _count(con, """
                    SELECT COUNT(*)
                    FROM tbl_auswertungspositionen p
                    LEFT JOIN tbl_auswertungen a ON a.id = p.auswertung_id
                    WHERE a.id IS NULL
                """)))
            if _table_exists(con, "tbl_austauschdatenbank"):
                checks.append(("Aktive Austauschdatensätze ohne Alt-PZN", _count(con, "SELECT COUNT(*) FROM tbl_austauschdatenbank WHERE COALESCE(status,'aktiv')='aktiv' AND (pzn_alt IS NULL OR TRIM(pzn_alt)='')")))
                checks.append(("Aktive Austauschdatensätze ohne Ziel/Freitext", _count(con, "SELECT COUNT(*) FROM tbl_austauschdatenbank WHERE COALESCE(status,'aktiv')='aktiv' AND COALESCE(pzn_nmg,'')='' AND COALESCE(freitext_austausch,'')=''")))
                duplicate_active = _count(con, """
                    SELECT COUNT(*) FROM (
                        SELECT pzn_alt
                        FROM tbl_austauschdatenbank
                        WHERE COALESCE(status,'aktiv')='aktiv' AND COALESCE(pzn_alt,'') <> ''
                        GROUP BY pzn_alt
                        HAVING COUNT(*) > 1
                    )
                """)
                checks.append(("Alt-PZN mit mehreren aktiven Austauschdatensätzen", duplicate_active))
            if _table_exists(con, "tbl_lernvorschlaege"):
                checks.append(("Lernvorschläge ohne Alt-PZN", _count(con, "SELECT COUNT(*) FROM tbl_lernvorschlaege WHERE pzn_alt IS NULL OR TRIM(pzn_alt)=''")))
                duplicate_suggestions = _count(con, """
                    SELECT COUNT(*) FROM (
                        SELECT COALESCE(pzn_alt,''), COALESCE(pzn_nmg,''), COALESCE(freitext_austausch,'')
                        FROM tbl_lernvorschlaege
                        WHERE status IN ('neu','uebernommen','abgelehnt')
                        GROUP BY COALESCE(pzn_alt,''), COALESCE(pzn_nmg,''), COALESCE(freitext_austausch,'')
                        HAVING COUNT(*) > 1
                    )
                """)
                checks.append(("Doppelte aktive Lernvorschläge", duplicate_suggestions))
            if _table_exists(con, "tbl_nmg_stamm"):
                checks.append(("NMG-Stammdaten ohne PZN", _count(con, "SELECT COUNT(*) FROM tbl_nmg_stamm WHERE pzn IS NULL OR TRIM(pzn)=''")))

            lines.append("")
            lines.append("Fachliche Auffälligkeiten")
            lines.append("-" * 90)
            for name, count in checks:
                if count < 0:
                    lines.append(f"{name}: Prüfung nicht möglich")
                    result["warnings"].append(f"Prüfung nicht möglich: {name}")
                elif count == 0:
                    lines.append(f"{name}: OK")
                else:
                    lines.append(f"{name}: {count} Auffälligkeit(en)")
                    result["warnings"].append(f"{name}: {count}")

    except Exception as exc:
        result["status"] = "FEHLER"
        result["errors"].append(str(exc))
        lines.append("")
        lines.append(f"FEHLER BEIM CHECK: {exc}")

    if result["errors"]:
        result["status"] = "FEHLER"
    elif result["warnings"]:
        result["status"] = "WARNUNG"

    lines.append("")
    lines.append("Ergebnis")
    lines.append("-" * 90)
    lines.append(f"Status: {result['status']}")
    if result["warnings"]:
        lines.append("Warnungen:")
        for item in result["warnings"]:
            lines.append(f"- {item}")
    if result["errors"]:
        lines.append("Fehler:")
        for item in result["errors"]:
            lines.append(f"- {item}")

    out.write_text("\n".join(lines) + "\n", encoding="utf-8")
    try:
        log_event("datenbankcheck", "Datenbank-Gesundheitscheck erstellt", f"Status: {result['status']} | Datei: {out}")
    except Exception:
        pass
    return result
