# NMGone — Übersetzungen

Hier liegt die Excel-Datei mit allen übersetzbaren Texten der App.

## Workflow

**Du (oder ein Muttersprachler):**
1. Datei `NMGone_Translations.xlsx` öffnen
2. Tab **"Translations"** wählen
3. Spalten **D (SK)** und **E (CZ)** mit Übersetzungen füllen
4. Spalten A (Key) und B (DE) **nicht** verändern — das ist die Referenz
5. Speichern → entweder
   - per OneDrive / E-Mail an Laemb schicken, oder
   - im GitHub-Web direkt hochladen (Drag-and-Drop auf die alte Datei → "Commit changes")

**Beim nächsten NMGone-Build** zieht ein kleines Skript die Übersetzungen automatisch in `app/i18n.py` und sie sind in der App sichtbar.

## Was ist drin

Aktuell ~28 Top-Strings: Navigation, Splash, Hauptbuttons. Bei jedem Service-Pack kommen neue Strings dazu — die Excel wird dann erweitert.

## Hinweise für Übersetzer

- **Kurz halten**: viele Texte sind Button-Beschriftungen — wenn sie zu lang werden, brechen sie das Layout.
- **Fachbegriffe**: PZN, APU, HAP, PK, ZF bleiben unverändert (deutsche Pharma-Standardabkürzungen).
- **Unsicher?** Spalte F gibt einen Hinweis, ob es noch ein Stub ist (= aktuell fällt die App auf Deutsch zurück).
