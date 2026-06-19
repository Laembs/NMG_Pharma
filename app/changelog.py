"""V1.1 SP11: Zentraler Changelog fuer die Versionsinfo-Seite.

Eintraege absteigend nach Version sortiert. Neue Releases hier anhaengen.
Format pro Eintrag: (display_name, datum_iso, lines).
"""

CHANGELOG: list[tuple[str, str, list[str]]] = [
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
