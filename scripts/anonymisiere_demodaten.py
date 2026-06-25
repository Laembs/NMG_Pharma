# -*- coding: utf-8 -*-
"""
Anonymisiert die Demo-Datenbank fuer eine GEFAHRLOSE oeffentliche Weitergabe.

Hintergrund: data/nmg_demodatenbank.sqlite ist eine Kopie der Produktiv-DB und
enthaelt damit reale Geschaeftsdaten (echte Partnerkonditionen/Rabatte, echte
EK-Einkaufspreise, reale Apotheken-/Kundennamen in Lern- und Analysetabellen).

Dieses Skript:
  1. LEERT die historischen Lern-/Analyse-/EK-/PK-Tabellen (fuer ein Demo nicht
     noetig und kommerziell sensibel).
  2. NEUTRALISIERT sensible Klartextfelder in den verbleibenden Tabellen
     (Rabattquellen, Austausch-Herkunft, Profil).
  3. VACUUM, um die Datei zu verkleinern.
  4. Prueft danach automatisch, dass keine realen Marker mehr enthalten sind.

Oeffentliche Stammdaten (PZN, Arzneimittelnamen, Taxe-/APU-Preise, Biosimilar-
Gruppen aus der Gelben Liste) bleiben erhalten - das sind keine Geschaefts-
geheimnisse, sondern oeffentliche Referenzdaten.

Idempotent. Auf die Produktiv-DB wird NIE zugegriffen.
"""
import os
import re
import sqlite3
import random

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DST = os.path.join(ROOT, "data", "nmg_demodatenbank.sqlite")

random.seed(7)

# Historische/sensible Tabellen, die fuers Demo komplett geleert werden ------
LEEREN = [
    # echte Lernstaende (Quelle = reale Apotheken-Dateien)
    "tbl_pzn_basis_stimmen", "tbl_pzn_basisdaten",
    "tbl_hersteller_stimmen", "tbl_hersteller_lern",
    "tbl_lernhistorie", "tbl_lernvorschlaege",
    # echte EK-Einkaufspreise
    "tbl_pzn_ek_rohdaten",
    # reale gespeicherte Auswertungen + Positionen (named Apotheken)
    "tbl_auswertungen", "tbl_auswertungspositionen",
    # Liefer-/Importspuren mit realen Dateinamen
    "tbl_lieferfaehigkeit", "tbl_import_log", "tbl_rohdaten_mapping",
    # reale Partnerkonditionen (PK)
    "tbl_pk_konditionen", "tbl_pk_kunden",
    "tbl_pk_konditionen_import", "tbl_pk_rabatte_import",
    "tbl_nmg_rabatte_historie", "tbl_nmg_rabatte_snapshots",
    "tbl_nmg_stamm_import", "tbl_apu_hap_import",
    # Altlasten-Logs mit realen Bezuegen
    "tbl_faktura_log",
]


def main():
    if not os.path.exists(DST):
        raise SystemExit(f"Demo-DB fehlt: {DST} (zuerst seed_demodaten.py laufen lassen)")
    con = sqlite3.connect(DST)
    cur = con.cursor()
    vorhandene = {r[0] for r in cur.execute(
        "SELECT name FROM sqlite_master WHERE type='table'")}

    # 1) Sensible Historien-Tabellen leeren -------------------------------
    geleert = 0
    for t in LEEREN:
        if t in vorhandene:
            n = cur.execute(f"SELECT count(*) FROM {t}").fetchone()[0]
            cur.execute(f"DELETE FROM {t}")
            if n:
                geleert += n
    print(f"Geleert: {geleert} Zeilen aus {len(LEEREN)} Historien-/PK-Tabellen")

    # 2) Klartextfelder neutralisieren ------------------------------------
    # nmg_rabatte: Quelle verschleiern + Rabatte auf synthetische Werte setzen
    rab_werte = [5, 7.5, 10, 12.5, 15, 20, 25, 30]
    for (pzn,) in cur.execute("SELECT nmg_pzn FROM nmg_rabatte").fetchall():
        cur.execute(
            "UPDATE nmg_rabatte SET rabatt=?, quelle='Demo-Kondition', "
            "letzte_aktualisierung='2026-01-01' WHERE nmg_pzn=?",
            (random.choice(rab_werte) / 100.0, pzn))

    # Austausch-Tabellen: reale Herkunft/Dateinamen entfernen
    cur.execute(
        "UPDATE tbl_austauschdatenbank "
        "SET quelle='Stammdaten-Demo', bemerkung=NULL "
        "WHERE quelle!='Demo'")
    cur.execute(
        "UPDATE tbl_austauschartikel "
        "SET quelle='Stammdaten-Demo', bemerkung=NULL "
        "WHERE quelle!='Demo'")

    # Artikelstamm: Importquelle generalisieren (kein echter Dateibezug)
    cur.execute("UPDATE tbl_artikelstamm SET quelle='Stammdaten' "
                "WHERE quelle IS NOT NULL")
    # NMG-Stamm: Quelle-Dateiname entfernen
    cur.execute("UPDATE tbl_nmg_stamm SET quelle='Stammdaten' "
                "WHERE quelle IS NOT NULL")

    # Belegzeilen: interne Konditionsquelle ('Partnerkonditionen ...') saeubern
    cur.execute("UPDATE tbl_bestellpositionen SET rabatt_quelle='NMG-Kondition' "
                "WHERE rabatt_quelle LIKE '%Partnerkonditionen%'")

    # Eigenes Mitarbeiterprofil / Meta neutralisieren
    if "tbl_mitarbeiterprofil" in vorhandene:
        cur.execute("DELETE FROM tbl_mitarbeiterprofil")
    cur.execute("DELETE FROM meta WHERE key LIKE 'pk_rabatte_%'")
    cur.execute("UPDATE meta SET value='NMGone Demo' "
                "WHERE key='basis'")

    con.commit()

    # 3) Datei verkleinern ------------------------------------------------
    cur.execute("VACUUM")
    con.commit()

    # 4) Verifizieren -----------------------------------------------------
    marker = re.compile(
        r"pauly|buttlar|walkhoff|greif|anklam|stelmach|hofwiesen|born_apo|"
        r"auengrund|rosen apo|sonnen apo|handpr|\.xlsx|apotheke am markt|"
        r"partnerkonditionen", re.I)
    tabs = [r[0] for r in cur.execute(
        "SELECT name FROM sqlite_master WHERE type='table'")]
    treffer = {}
    for t in tabs:
        cols = [c[1] for c in cur.execute(f'PRAGMA table_info("{t}")')]
        if not cols:
            continue
        try:
            rows = cur.execute(f'SELECT * FROM "{t}"').fetchall()
        except sqlite3.Error:
            continue
        n = 0
        for row in rows:
            for v in row:
                if isinstance(v, str) and marker.search(v):
                    n += 1
                    break
        if n:
            treffer[t] = n
    con.close()

    size_mb = os.path.getsize(DST) / (1024 * 1024)
    print(f"Dateigroesse nach VACUUM: {size_mb:.1f} MB")
    if treffer:
        print("!! VERBLEIBENDE REALE MARKER:")
        for t, n in treffer.items():
            print(f"   {t}: {n}")
        raise SystemExit("Anonymisierung unvollstaendig - bitte pruefen.")
    print("OK: keine realen Apotheken-/Datei-/Konditionsmarker mehr gefunden.")


if __name__ == "__main__":
    main()
