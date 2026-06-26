# -*- coding: utf-8 -*-
"""Zentrale Datenquelle der Revisions-/Aenderungsuebersicht.

Eine Quelle der Wahrheit fuer ZWEI Ausgaben:
  * app/revision_uebersicht.py  -> die Testoberflaeche (Fenster) "Was wurde veraendert"
  * scripts/build_handout_revision.py -> das PDF-Handout zum Nachvollziehen

Wenn sich an den Apps/Aenderungen etwas aendert, NUR hier pflegen – Oberflaeche
und PDF ziehen automatisch nach. Bewusst reiner Text/Listen, keine UI-Importe,
damit das Modul auch ohne Tkinter (z.B. im Build-Skript) ladbar bleibt.
"""
from __future__ import annotations

# Stand dieser Revision (Demo/Wareneingang/Meldungen/Einkauf-Welle) ------------
REVISION = {
    "titel": "Revision Sommer 2026 – Vom Tool zur Programm-Familie",
    "untertitel": "Wareneingang & Retouren · Meldungen · Einkauf · Demo-Modus",
    "datum": "2026-06-25",
    "version": "V2.1.0",
    "claim": (
        "NMGone ist nicht mehr ein Programm, sondern eine Familie eigenstaendiger "
        "Apps rund um eine gemeinsame Datenbank. Diese Revision legt das Fundament "
        "fuer den naechsten grossen Schritt – die zentrale Mehrbenutzer-Architektur."
    ),
    "kennzahlen": [
        ("3", "neue eigenstaendige Apps", "Wareneingang & Retouren, Meldungen, Einkauf"),
        ("8", "Apps mit eigenem Symbol", "jede per Taskleiste & Kachel startbar"),
        ("52", "Module in app/", "~42.000 Zeilen gemeinsamer Code"),
        ("1", "Demo-Modus", "gefahrlos vorfuehren – Produktiv-DB bleibt unberuehrt"),
    ],
}

# Die Programm-Familie ---------------------------------------------------------
# status: "neu" (in dieser Revision dazugekommen), "erweitert" (angefasst),
#         "stabil" (unveraendert mitgelaufen).
APPS = [
    {
        "key": "nmgone", "icon": "🏠", "name": "NMGone (Kern)", "status": "erweitert",
        "farbe": "#0B4A86", "start": "Doppelklick / start.py",
        "zweck": "Zentrale: Bedarfsanalyse, Produktanalyse, Kunden, Wissens-Datenbanken, "
                 "Dashboard mit Kacheln zu allen Apps.",
        "neu": "Drei neue Kacheln + Start-Logik (Wareneingang, Meldungen, Einkauf) "
               "ins Dashboard und in die App-Liste eingehaengt.",
    },
    {
        "key": "gdp", "icon": "📦", "name": "Wareneingang & Retouren (GDP)", "status": "neu",
        "farbe": "#0B6E6E", "start": "Kachel / NMGone.exe --gdp / start_gdp.py",
        "zweck": "Wareneingangspruefung, Chargen-Rueckverfolgung (Kunde↔Charge), "
                 "Kundenqualifizierung, Retouren/Reklamationen mit Gutschrift-Workflow, "
                 "revisionssicheres Protokoll – wie es die Grosshandelserlaubnis verlangt.",
        "neu": "Komplett neue App (~1.775 Zeilen), eigene GDP-Tabellen, eigenes Icon (GDP.ico).",
    },
    {
        "key": "meldungen", "icon": "🔔", "name": "Meldungen", "status": "neu",
        "farbe": "#B5391F", "start": "Kachel / NMGone.exe --meldungen / start_meldungen.py",
        "zweck": "GDP-Meldewesen: Abweichungs-/Qualitaetsmeldungen mit Workflow & CAPA, "
                 "Kuehlsachenkontrolle (Temperatur-Soll-Grenzen) und Selbstinspektion. "
                 "Buendelt die Pflichten, die vorher unuebersichtlich in der GDP-App lagen.",
        "neu": "Neue App (~975 Zeilen). Kuehlkette & Selbstinspektion aus der GDP-App "
               "hierher ausgelagert, Protokoll mit GDP geteilt.",
    },
    {
        "key": "einkauf", "icon": "🛒", "name": "Einkauf", "status": "neu",
        "farbe": "#0B4A86", "start": "Kachel / NMGone.exe --einkauf / start_einkauf.py",
        "zweck": "Beschaffung EU-Ausland: Lieferanten je Land mit Waehrung & Lieferzeit, "
                 "Einkaufsquellen (PZN×Lieferant), §129-Margenrechner (Import-EK → AVP → "
                 "garantierter Preisabstand + eigene Marge), Aufgaben/Wiedervorlagen als Meldungen.",
        "neu": "Neue App (~1.835 Zeilen), eigene Einkauf-Tabellen, Wechselkurs-Umrechnung.",
    },
    {
        "key": "kasse", "icon": "🛒", "name": "Kasse", "status": "stabil",
        "farbe": "#8B5A00", "start": "Kachel / NMGone.exe --kasse / start_kasse.py",
        "zweck": "Verkauf an Apotheken, Wareneingang ins Lager, Lieferschein, Defektmeldung, "
                 "EK/Lagerwert-Auswertungen.",
        "neu": "Liefert den Wareneingang, auf dem die neue GDP-App aufsetzt.",
    },
    {
        "key": "faktura", "icon": "🧾", "name": "Faktura", "status": "stabil",
        "farbe": "#0B4A86", "start": "Kachel / NMGone.exe --faktura / start_faktura.py",
        "zweck": "Rechnungen, Gutschriften, Quartalsverguetung – mit anpassbaren Vorlagen.",
        "neu": "Unveraendert mitgelaufen.",
    },
    {
        "key": "personal", "icon": "👤", "name": "Mitarbeiter / Personal", "status": "stabil",
        "farbe": "#6B4FB3", "start": "Kachel / NMGone.exe --personal / start_personal.py",
        "zweck": "Organigramm, Abwesenheiten, Arbeitsbereiche, Custom-Felder (EAV) "
                 "und Vorgesetzten-Matrix.",
        "neu": "Unveraendert mitgelaufen.",
    },
    {
        "key": "report", "icon": "📑", "name": "Auswertungen / Report", "status": "stabil",
        "farbe": "#0B6E6E", "start": "Kachel / NMGone.exe --report / start_report.py",
        "zweck": "Verkaeufe, Kunden, Artikel frei auswerten und exportieren.",
        "neu": "Unveraendert mitgelaufen.",
    },
    {
        "key": "hilfe", "icon": "❓", "name": "Hilfe / Handbuch", "status": "erweitert",
        "farbe": "#11823B", "start": "NMGone.exe --hilfe / start_hilfe.py",
        "zweck": "Bebildertes Handbuch zu allen Modulen mit Screenshot-Slots.",
        "neu": "Drei neue Kapitel (Wareneingang & Retouren, Meldungen, Einkauf) ergaenzt.",
    },
]

# Was wurde in dieser Revision veraendert? -------------------------------------
# kategorie -> Liste von (titel, detail)
AENDERUNGEN = [
    ("Neue Apps", [
        ("Wareneingang & Retouren (GDP)",
         "Eigenstaendige App fuer Wareneingangspruefung, Chargen-Rueckverfolgung, "
         "Kundenqualifizierung und Retouren mit Gutschrift-Workflow."),
        ("Meldungen",
         "Eigenstaendige App fuer GDP-Meldewesen, Kuehlsachenkontrolle und "
         "Selbstinspektion – aus der GDP-App ausgelagert und gebuendelt."),
        ("Einkauf",
         "Eigenstaendige App fuer Beschaffung EU-Ausland, §129-Margenrechner und "
         "Aufgaben/Wiedervorlagen, die als Meldungen auf dem Dashboard erscheinen."),
    ]),
    ("Integration in NMGone", [
        ("Dashboard-Kacheln",
         "Drei neue Kacheln (Wareneingang, Meldungen, Einkauf) in der Apps-Sektion "
         "UND in der grossen App-Liste – jede oeffnet die App als eigenen Prozess."),
        ("Start-Schalter",
         "start.py / NMGone.exe versteht jetzt --gdp, --meldungen, --einkauf, --personal; "
         "jede App bekommt ihr eigenes Taskleisten-Symbol (AUMID)."),
        ("Hilfe mitgezogen",
         "Drei neue, bebilderte Hilfe-Kapitel; Screenshot-Slots unter assets/hilfe/."),
    ]),
    ("Datenmodell (migrations.py)", [
        ("Einkauf-Tabellen",
         "tbl_einkauf_lieferanten, _quellen, _wechselkurse, _aufgaben, _einstellungen, _log "
         "mit Indizes auf PZN, Lieferant, Faelligkeit und Status."),
        ("GDP-Tabellen",
         "Zentral in gdp_app.ensure_gdp_tables erzeugt; migrations ruft sie beim Start "
         "ebenfalls auf, damit auch NMGone die Tabellen sichert (Single Source)."),
    ]),
    ("Demo-Modus zum Vorfuehren", [
        ("Eigene Demo-Datenbank",
         "scripts/seed_demodaten.py kopiert die Startdatenbank nach "
         "data/nmg_demodatenbank.sqlite und befuellt die operativen Tabellen mit "
         "mind. 12 Datensaetzen je Schritt – Produktiv-DB bleibt unberuehrt."),
        ("Start per Doppelklick",
         "Demo-Modus starten.bat, Wareneingang + Retouren (Demo) starten.bat und "
         "Meldungen (Demo) starten.bat setzen NMGONE_DEMO=1 und starten gefahrlos."),
    ]),
    ("Kunden-Landkarte (Vorbereitung)", [
        ("Offline-Geokodierung",
         "Neues Modul app/geo_de.py bildet PLZ-Leitregionen offline auf Koordinaten ab "
         "(keine Zusatzpakete, kein Internet) – Basis fuer die Deutschlandkarte der Kunden."),
    ]),
    ("Diese Testoberflaeche", [
        ("Revisions-Uebersicht",
         "Neues, isoliertes Fenster (app/revision_uebersicht.py), das genau zeigt, was "
         "veraendert wurde – plus dieses PDF-Handout zum Nachvollziehen."),
    ]),
]

# Aufraeumen: was erledigt, was empfohlen --------------------------------------
AUFGERAEUMT = [
    "Leere Stray-Datei 'python' (0 Bytes) entfernt.",
    "Alle __pycache__-Ordner im Projekt geleert (regenerieren sich von selbst).",
    "Geprueft: alle 52 app-Module + alle start-Skripte kompilieren fehlerfrei.",
    "Geprueft: neue Apps korrekt in Nav-Dispatch, beide Kachel-Listen und start.py verdrahtet.",
]

# (pfad, groesse, empfehlung) – NICHT automatisch geloescht (User entscheidet)
AUFRAEUM_EMPFEHLUNG = [
    ("app -vor Schulbankpatch/", "2,6 MB", "Alter Code-Backup-Ordner – nach Sichtung loeschbar (liegt in .gitignore)."),
    ("dist/  +  dist_setup/  +  build/", "≈ 296 MB", "Reine Build-Artefakte – jederzeit neu baubar, gefahrlos loeschbar."),
    ("app.zip", "164 KB", "Alter Code-Snapshot – durch git ersetzt, loeschbar."),
    ("updates_alt/", "404 KB", "Alte Update-Pakete (Binaer-Blobs), loeschbar."),
    ("organigramm_test.db", "44 KB", "Test-Datenbank im Wurzelverzeichnis – gehoert nach data/ oder weg."),
    ("README_2_6*.txt … README_3_*.txt", "7 Dateien", "Versions-Readmes alter Staende – in docs/ archivieren oder loeschen."),
    ("NMG_Analyse.spec", "1 Datei", "Alte PyInstaller-Spec; aktiv ist NMGone.spec – Altspec loeschbar."),
]

# Roadmap / "Naechste Revolution" ---------------------------------------------
# (titel, beschreibung, aufwand) aufwand: "klein"/"mittel"/"gross"
ROADMAP = [
    ("Zentrale Mehrbenutzer-Datenbank",
     "Weg von der einzelnen SQLite-Datei hin zu einem zentralen DB-Server, damit "
     "mehrere Arbeitsplaetze gleichzeitig arbeiten (NICHT ueber SharePoint).", "gross"),
    ("Echte Screenshots in die Hilfe",
     "Die Bild-Slots der neuen Kapitel (Wareneingang, Meldungen, Einkauf) mit echten "
     "Screenshots fuellen.", "klein"),
    ("Kunden-Deutschlandkarte fertigstellen",
     "Die Offline-Geokodierung (geo_de.py) in der Kunden-App zur Karte ausbauen "
     "(„wo sitzt welcher Kunde“).", "mittel"),
    ("Biosimilar-Wissensbasis verzahnen",
     "Die neue Gelbe-Liste-Wissensbasis mit Bedarfsanalyse & GUI verbinden "
     "(Wirkstoff-Biosimilar-Zuordnung).", "mittel"),
    ("Produktanalyse-Bugs fixen",
     "EK durchweg NULL (Umsatz=0); PZN-Normalisierung ohne '/N'-Strip korrigieren.", "klein"),
    ("App-Symbole & Setup finalisieren",
     "Eigene Icons fuer Meldungen/Einkauf, Verknuepfungen im Installer ergaenzen.", "klein"),
]
