"""V1.1 SP11: Zentraler Changelog fuer die Versionsinfo-Seite.

Eintraege absteigend nach Version sortiert. Neue Releases hier anhaengen.
Format pro Eintrag: (display_name, datum_iso, lines).
"""

CHANGELOG: list[tuple[str, str, list[str]]] = [
    ("V1.1 SP19", "2026-06-19", [
        "Neue Auswertung ist 11x schneller: trace_lookup_source nutzt Index-Cache statt pro Zeile die komplette Austauschdatenbank zu scannen. 414 Zeilen jetzt in ~16 s statt ~3 Min.",
        "Bugfix Neue Auswertung (SP18-Regression): Doppelklick-Schutz auf 'Auswertung starten' verhindert stille Mehrfachstarts (Log zeigte 90 Aufrufe in einer Sekunde).",
        "Bugfix Status-Box: erscheint jetzt zuverlaessig auch wenn die rechte Sidebar eingeklappt ist (Fallback am Hauptfenster oben rechts).",
        "Bugfix Background-Jobs: Exceptions in on_done/on_error werden ins Fehler-Log geschrieben + messagebox angezeigt statt geraeuschlos zu verschwinden.",
        "NMG-Rabatte-Uebersicht (neue Kachel in Apps + Dashboard): 4 Reiter - alle Rabatte mit Filter, Statistik, Diff zum letzten Stand, Verlauf pro PZN.",
        "Vor jedem PK-Rabatte-Import wird automatisch ein Snapshot des aktuellen Stands angelegt - damit Diff/Verlauf nachvollziehbar sind.",
        "Auto-Restart nach Setup: NMGone startet nach jedem Update automatisch wieder (kein Haekchen mehr am Wizard-Ende noetig).",
    ]),
    ("V1.1 SP18", "2026-06-19", [
        "Neue Auswertung laeuft jetzt im Hintergrund (analog manueller Import in SP15). UI bleibt klickbar.",
        "Background-Status-Box umgezogen: von rechts oben im Header in die rechte Sidebar ueber den Backup-Status.",
        "Wenn keine Jobs laufen, ist die Box unsichtbar.",
        "UnknownInputFormatError + Format-Assistent-Pfad jetzt in on_error des Background-Jobs (interaktiv im UI-Thread).",
    ]),
    ("V1.1 SP17", "2026-06-19", [
        "Manueller Import: nicht-erkannte Dateien werden uebersprungen statt pro-Datei modal abgefragt.",
        "TXT-Sammeldatei 'nicht_erkannte_dateien_TIMESTAMP.txt' in importierte_analysen/<typ>/<Jahr>/Q<n>/ mit Liste der Dateinamen + Fehlertext.",
        "Am Ende EINE einmalige Sammel-Frage 'Sollen die N nicht-erkannten Dateien manuell gemapped werden?' (statt Pro-Datei-Dialog).",
        "Aufraeum: doppelte Statistik-Messagebox am Import-Ende war drin, jetzt nur noch einmal.",
    ]),
    ("V1.1 SP16", "2026-06-19", [
        "Import-Performance fuer rohdaten-Modus (Excel ohne Standard-Header).",
        "Vorher: pro Zeile mehrere SQL-Queries mit pzn_norm()-UDF in WHERE - Full-Table-Scan ueber zigtausend Eintraege.",
        "Jetzt: Pre-Cache vor der Schleife. Alle PZNs einmal sammeln, dann Basisdaten/Austausch/NMG-Stamm/Rabatte/Lieferfaehigkeit in einem Rutsch laden (IN-Lookup mit Index).",
        "Erwartet: 50-200x schneller im rohdaten-Modus. Bei erkanntem Header weiterhin der bereits schnelle Pfad aus SP13.",
    ]),
    ("V1.1 SP15", "2026-06-19", [
        "Manuelle Imports laufen jetzt im HINTERGRUND - kein blockierender Modal-Dialog mehr.",
        "Status-Box oben rechts mit Titel + aktueller Datei + Fortschritts-Balken; verschwindet nach Abschluss.",
        "UI bleibt waehrend des Imports klickbar; mehrere Imports koennen parallel laufen.",
        "Statistik-Toast unten in der Statusleiste nach Abschluss (8 s).",
        "Neue Methode _run_background analog _run_busy, aber non-modal.",
    ]),
    ("V1.1 SP14", "2026-06-19", [
        "Manuelle Imports: Fortschritts-Anzeige im Busy-Dialog. Zeigt 'Datei 5 von 96: foo.xlsx' statt 'Bitte warten'.",
        "Neuer _run_busy(..., progress=True)-Modus: Worker bekommt ein progress(text)-Callable und kann den Subtitle waehrend der Verarbeitung aktualisieren.",
    ]),
    ("V1.1 SP13", "2026-06-19", [
        "Manueller Analysen Import deutlich schneller: pro Excel-Zeile war vorher 2-3 Einzel-Queries als implizite Transaktion. Bei 5000 Zeilen = 10.000-15.000 Transaktionen, dauerte minutenlang.",
        "Fix: Inserts werden gesammelt und als executemany in einer Transaktion geschrieben; Duplikat-Check liest bestehende Schluessel einmal als Set.",
        "Erwartet: 5000 Zeilen jetzt in unter 30 s (vorher mehrere Minuten).",
        "Stellen: _insert_learning_suggestions, _insert_learning_suggestions_from_auswertung, import_historical_market_file (Positionen-INSERT).",
    ]),
    ("V1.1 SP12", "2026-06-19", [
        "Gespeicherte Analysen: Refresh nach Admin-Loeschen funktioniert wieder.",
        "Gespeicherte Analysen: 'Item end not found'-Fehler bei Seitenwechsel weg (defensive Guards).",
        "Manuelle Analysen Import: Fortschritts-Dialog (Busy-Modal) statt eingefrorenes UI.",
        "Manuelle Imports: Original-Datei wird in importierte_analysen/<PK|ZF>/<Jahr>/Q<n>/ kopiert.",
        "Kalender-Popup schliesst beim Klick ausserhalb.",
        "Dashboard-Editor: Fenster groesser, Speicher-Button immer sichtbar.",
        "Dashboard: 'Alles abwaehlen' funktioniert jetzt wirklich (Sentinel _NONE_ vs _EMPTY_).",
        "Globale Suche auf Startseite: Tabelle nur sichtbar wenn Treffer da; Hoehe nur 3 Zeilen.",
        "Produktanalyse: nur Artikel mit Namen im Artikelstamm.",
        "Re-Import: 'schon importiert' greift nur, wenn die Auswertung noch lebt.",
        "Neue Ordner-Struktur: ausgaben/<Kategorie>/<Jahr>/Q<n>/ und importierte_analysen/<PK|ZF>/<Jahr>/Q<n>/.",
    ]),
    ("V1.1 SP11", "2026-06-19", [
        "Produktanalyse: Basis-Info-Zeile und Beschreibungs-Zeile in den Excel-Sheets entfernt (Zeile 2 + 3).",
        "Versionsinfo: neue Seite mit Liste aller Versionen und Changelog pro Version.",
        "Startseite: 'Datenbankstatus'-Box (Werte waren nicht real) und Datenbank-Dateiname in 'Aktiver Bearbeiter' entfernt.",
        "Dashboard-Editor: 'Alles abwaehlen + Speichern' funktioniert jetzt (vorher Block-Dialog).",
        "Button-Style auf Hauptseiten und Sidebars vereinheitlicht.",
    ]),
    ("V1.1 SP10", "2026-06-19", [
        "Bug: Produktanalyse zeigt nur noch Artikel, die im Artikelstamm einen Namen haben (keine leeren Zeilen).",
        "Bug: Re-Import einer Datei klappt, wenn die zugehoerige Auswertung zwischendurch geloescht wurde.",
        "Gespeicherte Analysen: neue Filter-Buttons Importiert / Programm / Produktanalyse.",
        "Sidebar-Button 'Zeitraum LOESCHEN' (rot, Admin-Passwort, voller DB-Backup vorher).",
        "Datums-Felder als Kalender-Picker (tkcalendar) in Suche und Zeitraum-Dialogen.",
    ]),
    ("V1.1 SP9", "2026-06-19", [
        "Globale Suche: zeigt initial 3 Treffer + 'Mehr anzeigen (N)'-Button.",
        "Globale Suche als eigenes Toplevel-Tool (analog Vergleichs-Suche).",
        "Kachel in der Apps-Seite + Dashboard-Editor-Eintrag.",
    ]),
    ("V1.1 SP8", "2026-06-19", [
        "Gespeicherte Analysen: getrennte Suchfelder (Apotheke, Kunden-Nr., Kunde, Datum von/bis).",
        "'Auswertung oeffnen' oeffnet jetzt direkt die Excel.",
        "Zeitraum-Archivierung: Auswertungen wandern als ZIP nach backups/analysen_archiv/, DB-Zeilen werden geloescht.",
        "Archiv-Verwaltung: Liste der ZIPs, Doppelklick extrahiert Excel und oeffnet. Loeschen nur per Admin-Passwort.",
    ]),
    ("V1.1 SP7", "2026-06-19", [
        "Startseite: alte Austauschsuche entfernt, durch globale Suche ersetzt.",
        "Globale Suche findet Kunden, gespeicherte Analysen und Artikel parallel.",
        "Doppelklick verzweigt zu Kunden-Center, Excel oder Vergleichs-Suche.",
    ]),
    ("V1.1 SP6", "2026-06-19", [
        "Produktanalyse: zwei neue Reiter pro Kundentyp - 'NMG-Sortiment' und 'Austausch vorhanden'.",
        "Bei PK+ZF: 9 Reiter total (3 Sichten x 3 Reiter).",
    ]),
    ("V1.1 SP5", "2026-06-19", [
        "Vergleichs-Suche: Hersteller-Filter mit Wert (#hexal, #ratio, ...).",
        "Internes SQL-LIMIT bei Filtern auf 2000 hochgesetzt, damit Filter nicht ins Leere greifen.",
        "Aufraeumen: abstraktes #herst aus SP4 entfernt.",
    ]),
    ("V1.1 SP4", "2026-06-19", [
        "Vergleichs-Suche: Filter-Syntax mit '#' eingefuehrt: #nmg, #austausch, #schulbank, #wirkstoff, #herst.",
        "Unbekannte #-Tags werden verworfen, UND-Logik bei mehreren Filtern.",
    ]),
    ("V1.1 SP3", "2026-06-19", [
        "Vergleichs-Suche: 'Quelle'-Spalte im Austauschartikel-Detail-Panel.",
        "Datenbankuebersicht: 'Datenbankinhalte leeren' direkt erreichbar (Ctrl+Alt+A-Toggle entfaellt).",
    ]),
    ("V1.1 SP2", "2026-06-19", [
        "Neue Wirkstoff/Staerke-Datenbank (tbl_wirkstoff_staerke).",
        "Import-Button in Datenbankuebersicht. 162k Zeilen in ~14 s importiert.",
        "Vergleichs-Suche kann nach Wirkstoff, Hersteller und Staerke suchen.",
        "Detail-Panel zeigt Wirkstoff + Staerke.",
    ]),
    ("V1.1 SP1", "2026-06-18", [
        "Produktanalyse-Performance: UDFs in JOIN-/EXISTS-Klauseln entfernt.",
        "Direkter =-Join nutzt die vorhandenen PZN-Indizes, PK+ZF in ~183 ms statt >3 Minuten.",
    ]),
    ("V1.1.0", "2026-06-18", [
        "Meilenstein-Release. Neue Vergleichs-Suche als eigenes Toplevel-Fenster.",
        "Live-Suche nach PZN oder Artikelname ab 3 Zeichen.",
        "Sucht parallel in Austauschdatenbank, NMG-Stamm, Artikelstamm, PZN-Basisdaten.",
        "Detail-Panel zeigt Stammdaten + Austauschartikel-Tabelle.",
        "Dashboard-Editor um 20 % verkleinert, nicht mehr maximiert.",
    ]),
    ("V1.0 SP30", "2026-06-18", [
        "Produktanalyse mit Vorkommen + Erste/Letzte Sichtung + PZN-Normalisierung in Austausch-Filtern.",
    ]),
    ("V1.0 SP29", "2026-06-18", [
        "PK-Rabatte: PZN-Suffix '/N' (SK-Lagervariante) wird korrekt abgeschnitten.",
        "Max-Rabatt pro PZN, Counter-Refresh, Lieferfaehigkeit-Anzeige entfernt.",
    ]),
    ("V1.0 SP28", "2026-06-18", [
        "Rabatt-Parser-Fix beim PK-Import. Counter zeigen ehrliche Werte.",
    ]),
    ("V1.0 SP24-SP27", "2026-06-17", [
        "pzn_norm()-Lookup, Statusliste + Start-Warnung, In-Memory-Caches fuer Lookups.",
    ]),
]


def get_changelog() -> list[tuple[str, str, list[str]]]:
    """Liefert alle Changelog-Eintraege (neueste zuerst)."""
    return list(CHANGELOG)
