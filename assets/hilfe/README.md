# Hilfe-Bilder

Hier liegen die Screenshots für die Hilfe-App (`app/hilfe_app.py`).

## So fügst du Bilder hinzu

1. Mach einen Screenshot der gewünschten Maske (Windows: `Win` + `Shift` + `S`,
   dann in z.B. Paint einfügen und als **PNG** speichern).
2. Lege die PNG-Datei in den passenden Themenordner unter `assets/hilfe/<thema>/`.
3. Benenne sie **exakt** wie unten angegeben.
4. Hilfe-App neu starten – das Bild erscheint automatisch anstelle des Platzhalters.

Solange ein Bild fehlt, zeigt die Hilfe einen Platzhalter mit genau diesem Pfad
an. Du musst **keinen Code ändern**, nur die Dateien ablegen.

Breite ist egal – die App skaliert automatisch auf max. 760 px Breite herunter.

## Erwartete Dateien

| Ordner            | Dateiname                  | Zeigt                                   |
|-------------------|----------------------------|-----------------------------------------|
| `start`           | `01_startbildschirm.png`   | Startbildschirm nach Programmstart      |
| `dashboard`       | `01_dashboard.png`         | Dashboard mit Kacheln und Info-Widgets  |
| `bedarfsanalyse`  | `01_bedarfsanalyse.png`    | Bedarfsanalyse mit Dateiauswahl/Ergebnis|
| `produktanalyse`  | `01_produktanalyse.png`    | Produktanalyse Chancen-Übersicht        |
| `kunden`          | `01_kunden.png`            | Kunden-Maske mit Reitern                |
| `kasse`           | `01_kasse.png`             | Kasse Verkaufsmaske                     |
| `auswertungen`    | `01_auswertungen.png`      | Auswertungen mit Ergebnistabelle        |
| `auswertungen`    | `02_auswertung_frei.png`   | Baukasten-Dialog (freie Auswertung)     |
| `schulbank`       | `01_schulbank.png`         | Schulbank mit Lernvorschlägen           |
| `daten`           | `01_daten.png`             | Seite „Daten aktualisieren"             |
| `backup`          | `01_backup.png`            | Backup- und Update-Funktionen           |

Weitere Bilder kannst du jederzeit ergänzen – dazu in `app/hilfe_app.py` beim
passenden Thema eine `("img", "dateiname.png", "Bildunterschrift")`-Zeile
einfügen.
