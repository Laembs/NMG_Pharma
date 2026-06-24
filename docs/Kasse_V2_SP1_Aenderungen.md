# Kasse-Ausbau V2.0 SP1 — Änderungsdokumentation

**Datum:** 2026-06-24
**Branch:** `feat/bestell-app`
**Status:** umgesetzt, getestet (py_compile, Headless-Build aller Views, reale PDF-Erzeugung,
sauberer App-Start). **Als V2.0 SP3 (2.0.3) versioniert** – noch nicht committet, noch nicht
gebaut (Build/Release durch dich, siehe Abschnitt 0).

---

## 0. Release: V2.0 SP3 (erledigt)
Die Kasse-Arbeit ist als **V2.0 SP3 (Version 2.0.3, 2026-06-24)** versioniert. Gebumpt wurde:
`version.json`, `app/backup.py` (APP_VERSION/APP_VERSION_DISPLAY), `app/changelog.py`
(neuer Eintrag „V2.0 SP3" ganz oben), `build_installer.bat` und neues
`installer/NMGone_Setup_2_0_3.iss`.

Hintergrund Nummerierung: SP1 = Mitarbeiter-App, SP2 = Bedarfsanalyse (beide parallele
Workstreams im selben Arbeitsverzeichnis), **SP3 = Kasse** (dieser Stand).

**Noch offen (deine Aufgabe):** `build_installer.bat` ausführen (PyInstaller + Inno Setup →
`dist_setup\NMGone_Setup_2_0_3.exe`), als **Draft-Release** auf GitHub stellen und nach
deinem Test publishen. Commit nur der Kasse-Dateien (kein `git add -A`, siehe unten).

---

## 1. Was umgesetzt wurde (6 Punkte)

### 1) Lieferschein nach MSK-Bestätigung
- Sobald ein Verkauf im Reiter **Verkäufe** als „✓ In MSK erfasst“ markiert wird, fragt
  die Kasse, ob ein **Lieferschein** erzeugt und geöffnet werden soll.
- Funktioniert auch für **mehrere** gleichzeitig markierte Verkäufe (Bulk).
- Der Lieferschein zeigt **PZN, Artikel, PCK, Charge, Verfall, Menge** sowie ein
  Unterschriftsfeld — **bewusst keine Preise** (Lieferschein, keine Rechnung).
- Zusätzlich neuer Button **„📋 Lieferschein“** in der Verkaufs-Detailansicht
  (Doppelklick auf einen Verkauf) zum jederzeitigen Nachdruck.
- Vorlage ist frei anpassbar (Firmendaten/Logo): `vorlagen/lieferschein.html`
  (wird beim ersten Mal automatisch aus `assets/lieferschein_vorlage.html` angelegt).
- Ausgabe landet in `ausgaben/lieferscheine/`.

### 2) Überbestand — Aufteilen in Bestellung + Vorbestellung
- Beim Hinzufügen einer Position im **Verkauf**: reicht der Lagerbestand nicht, kommt
  eine Abfrage mit **drei** Möglichkeiten:
  - **Ja** → verfügbare Menge sofort als **Bestellung** liefern, **Rest** automatisch als
    **Vorbestellung** aufnehmen (es entstehen zwei Positionszeilen).
  - **Nein** → nur den **vorhandenen Bestand** abverkaufen (Rest wird verworfen).
  - **Abbrechen** → nichts hinzufügen.
- Vorher wurde die *ganze* Position zur Vorbestellung — jetzt wird sauber gesplittet.

### 3) Neuer Reiter „Auswertung“
Drei Unter-Tabs:
- **Umsatz / Tagesabschluss** — Umsatz je **Tag / Monat / Jahr** mit Spalten
  *Anzahl Verkäufe · Anzahl Packungen · APU Brutto · Rabatt (Netto) · APU Netto*,
  optionaler Von/Bis-Zeitraum, Summenzeile. Es zählen nur echte Bestellungen aus
  nicht stornierten Verkäufen. (APU Brutto = APU×Menge ohne Rabatt; APU Netto =
  APU×Menge mit Rabatt; Rabatt (Netto) = gewährter Rabattbetrag.)
- **Verfall** — Lagerbestand nach Verfall: **abgelaufen = rot**, **läuft in ≤90 Tagen
  ab = gelb**, mit Warnzähler oben.
- **Inventur** — Zählliste mit Soll-Bestand je Charge.

### 4) Tagesabschluss (Umsatz pro Tag) — manuell + automatisch
- **Manuell:** Knopf „🧾 Tagesabschluss heute (PDF)“ oder **Doppelklick** auf einen Tag
  in der Umsatztabelle (nur bei Gruppierung „Tag“).
- Der Tagesabschluss-PDF enthält Kennzahlen (APU brutto, Rabatt gegeben, Umsatz netto)
  plus Artikel-Aufstellung des Tages.
- **Automatisch:** die laufende Kasse erzeugt **jeden Abend punktgenau um 18:00 Uhr**
  den Tagesabschluss des Tages und vergibt die laufende **Nr.** (einmal pro Tag). Beim
  Programmstart werden **verpasste Vortage** (Umsatz vorhanden, aber noch keine Nr.)
  nachgeholt und chronologisch nummeriert. „Erledigt" = es existiert eine Nr. in
  `tbl_kasse_tagesabschluss` (nicht die PDF-Datei) – so wird nie doppelt nummeriert.
- Zeitsteuerung: präziser `after()`-Timer auf 18:00 (kein 30-Min-Polling mehr);
  Uhrzeit über `KassePanel.AUTO_TAGESABSCHLUSS_STUNDE` einstellbar. Voraussetzung ist
  eine **laufende Kasse**; läuft sie um 18 Uhr nicht, wird der Tag beim nächsten Start
  nachgeholt.
- Ausgabe: `ausgaben/tagesabschluss/Tagesabschluss_JJJJ-MM-TT.pdf`.

### 5) PDF-Reports
- Verfall, Inventur und Umsatz/Tagesabschluss sind alle als **druckbares PDF**
  exportierbar (Bibliothek `fpdf2`, ist mit 2.8.7 installiert).
- Ausgabe: `ausgaben/auswertungen/` bzw. `ausgaben/tagesabschluss/`.

### 6) Datenbank-Robustheit (Mehrbenutzer-Vorbereitung)
- Jeder DB-Zugriff der Kasse nutzt jetzt `busy_timeout` (30 s). Damit bricht der
  Parallelbetrieb von NMGone und Kasse auf derselben Datenbank nicht mehr sofort mit
  „database is locked“ ab. (WAL war bereits aktiv.)

### 7) MSK-Abfrage beim manuellen Lieferschein (Nachtrag)
- Klickt man in der Verkaufs-Detailansicht auf **„📋 Lieferschein“** und der Auftrag ist
  **noch nicht** in MSK erfasst, kommt eine Abfrage: **Ja** = als „in MSK erfasst“
  markieren und Lieferschein erzeugen · **Nein** = Lieferschein trotzdem erzeugen ·
  **Abbrechen**. Ist der Auftrag bereits erfasst, wird direkt erzeugt.

### 8) Auswertung „Wer hat was mit dem Auftrag gemacht“ (Nachtrag)
- **Pro Auftrag:** neuer Knopf **„🕘 Verlauf“** in der Detailansicht zeigt alle
  Protokoll-Einträge zu diesem Auftrag (Mitarbeiter, Aktion, Zeitpunkt, Details) und
  exportiert sie als PDF.
- **Übergreifend:** im Reiter **Protokoll** neuer Knopf **„📄 Als PDF“** – exportiert die
  aktuell angezeigten/gefilterten Einträge (z. B. nach Auftrag-Nr oder Mitarbeiter
  gefiltert) als PDF.
- Grundlage ist die bestehende Protokoll-Tabelle `tbl_kasse_log`; protokolliert werden
  u. a. Verkauf gespeichert, MSK erfasst/offen, Lieferschein erzeugt, Storno,
  Bestandskorrektur, Report erzeugt.

### 9) Verfall: Zeitraum-Auswahl + Summen (Nachtrag)
- Im Reiter **Auswertung → Verfall** neues Dropdown **3 / 6 / 9 / 12 Monate / Alle**
  (Standard **3 Monate = 90 Tage**). Bei einer Monatsauswahl werden nur abgelaufene und
  innerhalb des Zeitraums ablaufende Chargen gezeigt; **Alle** zeigt den ganzen Bestand.
- Unter der Tabelle: **Summe Bestand** und **Verkaufswert (APU × Bestand)** (APU ist der
  Verkaufspreis an die Apotheke, daher Verkaufs- und nicht Lagerwert/Einkaufswert).
- Der PDF-Report übernimmt die gewählte Monatsauswahl und enthält APU-/Wert-Spalten
  plus die Summenzeile.

### 10) Inventur: Verkaufswert-Summe (Nachtrag)
- Inventur zeigt jetzt zusätzlich den **Verkaufswert (APU × Bestand)** – am Bildschirm und
  im PDF (mit Summenzeile).

### 11) Tagesabschluss: eigener Reiter mit Kalender, Nr und Kennzahlen (Nachtrag)
- **Auswertung → Tagesabschluss** ist ein **eigener Reiter**.
- **Kalender-Datumsauswahl** (`tkcalendar.DateEntry`, Fallback = Textfeld).
- **Laufende Tagesabschluss-Nr**: beim Erzeugen des PDFs wird eine fortlaufende Nummer
  vergeben und in der neuen Tabelle `tbl_kasse_tagesabschluss` gespeichert (eine Nr pro
  Tag, stabil bei erneutem Erzeugen).
- **Kennzahlen am Bildschirm**: Menge der Verkäufe, verkaufte Packungen, APU-Summe,
  Rabatt-Summe, Umsatz-Summe – plus Artikeltabelle des Tages.
- Hinweis: Die Tabelle `tbl_kasse_tagesabschluss` wird automatisch angelegt
  (`ensure_kasse_tables` + defensiv beim ersten Zugriff).

### 12) Vollbild-Start (Nachtrag)
- Die Kasse öffnet jetzt **maximiert** (`run_standalone` → `root.state("zoomed")`,
  Fallback für Nicht-Windows). Normale Fenstergröße bleibt als wiederhergestellter Zustand.

### 14) EK (Einkaufspreis) + Lagerwert (Nachtrag)
- **Neue Spalte `ek` in `tbl_lagerbestand`** (automatisch nachgerüstet). EK je Lagerzeile.
- **Wareneingang:** neues Feld **„EK €"** – bei Artikelwahl mit dem APU vorbelegt, frei
  überschreibbar. Wird in `tbl_lagerbestand` und `tbl_wareneingang_positionen` gespeichert
  (neuer EK überschreibt, sonst bleibt der bestehende erhalten).
- **Import:** eine EK-Spalte (Synonyme „EK/Einkaufspreis/Einkauf") wird – falls vorhanden –
  übernommen und in den Lagerbestand geschrieben.
- **Lagerwert (EK × Bestand)** wird jetzt **überall zusätzlich** zum Verkaufswert gezeigt:
  Lagerbestand-Tabelle (Wareneingang, Spalten EK + Lagerwert), Artikel-Summenleiste,
  Verfall-Summe, Inventur (Spalten EK + Lagerwert, Summe) und die PDFs (Verfall-Summe;
  Inventur-PDF jetzt **Querformat** mit APU/Verkaufswert/EK/Lagerwert).
- Solange kein EK erfasst ist, ist der Lagerwert 0 bzw. „—" (APU/Verkaufswert bleiben wie
  bisher). Ein echter EK ≠ APU lässt sich so jederzeit nachpflegen, ohne weitere Umbauten.
- Storno bucht eine geleerte Lagerzeile mit EK-Fallback = APU der Position zurück.

### 16) Einstellungen-Reiter (Texte in der App ändern) (Nachtrag)
- Neuer **Reiter „Einstellungen"** (Nav links, Symbol ⚙). Hier editierbar – **ohne Dateien
  anzufassen**:
  - **Firmendaten** (Firma, Adresse mehrzeilig, Kontakt) → Kopf von Auftragsbestätigung,
    Lieferschein **und** Defektmeldung.
  - **Defektmeldung-Rechtstext** (mehrzeilig) → genau der Platzhalter, den du füllen wolltest.
  - **Tagesabschluss-Uhrzeit** (0–23) → steuert den automatischen Lauf.
- Speicherung in neuer Tabelle **`tbl_kasse_einstellungen`** (Schlüssel/Wert) via neuem
  Modul `app/einstellungen.py`. Leere Felder fallen auf den Standardtext zurück.
- Die Dokument-Vorlagen nutzen jetzt Platzhalter (`{{firma}}`, `{{absender_adresse}}`,
  `{{absender_kontakt}}`, Defektmeldung `{{rechtstext}}`), die beim Erzeugen aus den
  Einstellungen gefüllt werden.
- **⚠ WICHTIG – bestehende Vorlagen:** Die neuen Platzhalter wirken nur in **frisch
  angelegten** Vorlagen. Wer schon eine `vorlagen/auftragsbestaetigung.html` (oder
  `lieferschein.html` / `defektmeldung.html`) hat – z. B. eine selbst angepasste –, muss
  diese Datei **löschen** (wird dann neu mit Platzhaltern erzeugt) **oder** die Platzhalter
  von Hand eintragen. Sonst greifen die Firmendaten/der Rechtstext aus den Einstellungen
  dort nicht. Auf einem neuen Rechner/Installation greift alles automatisch.
- Speichern plant den automatischen Tagesabschluss mit der (ggf. neuen) Uhrzeit neu.

### 15) Defektmeldung / Nichtverfügbarkeitsbescheinigung (Nachtrag)
- Neuer **Reiter „Defektmeldung"** (Nav links, Symbol ⚠). Aufbau wie ein Verkauf:
  **Apotheke suchen → Artikel hinzufügen** (mit Anzeige des aktuellen Bestands) →
  **Grund** wählen (nicht lieferbar / nicht vorrätig / Herstellerengpass / … + Freitext)
  → **„Defektmeldung erzeugen"**.
- Neues Modul `app/defektmeldung.py` + Vorlage `assets/defektmeldung_vorlage.html`
  (→ beim ersten Mal nach `vorlagen/defektmeldung.html` kopiert, frei editierbar).
  Ausgabe als HTML in `ausgaben/defektmeldungen/`, wird im Browser geöffnet (druckbar/„als
  PDF speichern"). Jede Erzeugung wird im Protokoll vermerkt („Defektmeldung erzeugt").
- **⚠ Rechtstext bewusst als Platzhalter:** Die konkrete gesetzliche Formulierung /
  Paragraphen-Verweise sind NICHT fest eingebaut, sondern als klar markierter Platzhalter
  in der Vorlage – muss fachlich/rechtlich geprüft und eingetragen werden.

### 13) Artikel-Übersicht: Summenleiste (Nachtrag)
- Im Reiter **Artikel** unten eine Leiste mit **Gesamtbestand** und **Verkaufswert
  (APU × Bestand)**. Sucht man nach PZN/Artikel, beziehen sich die Summen auf die
  **angezeigte Auswahl** (Beschriftung wechselt „gesamt“ ↔ „Auswahl“).

---

## 2. Geänderte / neue Dateien (NUR dieser Task)

**Neu angelegt**
- `app/lieferschein.py` — Lieferschein-Erzeugung (HTML aus austauschbarer Vorlage).
- `app/kasse_reports.py` — Datenabfragen + PDF-Erzeugung für Verfall/Inventur/Umsatz/Tagesabschluss.
- `assets/lieferschein_vorlage.html` — Standard-Lieferscheinvorlage.
- `docs/Kasse_V2_SP1_Aenderungen.md` — dieses Dokument.

**Geändert**
- `app/kasse_app.py` — neuer Reiter, Lieferschein-Hook, Überbestand-Split, `_conn`-Timeout,
  Auto-Tagesabschluss.
- `app/changelog.py` — Eintrag „V2.0 SP1“ (2026-06-24).
- `app/hilfe_app.py` — Kasse-Hilfeabschnitt um die neuen Funktionen erweitert.

---

## 3. Worauf du SPEZIELL achten solltest

### ⚠️ 3.1 Beim Committen: fremde Arbeit ist im Arbeitsverzeichnis!
Im Working Tree liegen **viele Änderungen, die NICHT zu diesem Task gehören** (anderer
laufender Stand). **Kein `git add -A`/`git commit -a`** — sonst landet halbfertige fremde
Arbeit im selben Commit. Nur diese Dateien gehören zu „Kasse V2.0 SP1“:

```
git add app/kasse_app.py app/changelog.py app/hilfe_app.py \
        app/lieferschein.py app/kasse_reports.py \
        assets/lieferschein_vorlage.html docs/Kasse_V2_SP1_Aenderungen.md
```

**Nicht von mir** (separat prüfen/committen oder stehen lassen): `app/faktura_app.py`,
`start_faktura.py`, `docs/Faktura_*`, `app/exporter.py`, `app/gui.py`, `app/migrations.py`,
`app/kurzbericht.py`, `docs/Mitarbeiter_Organigramm_Doku.md`, `start_organigramm_test.py`,
die `praesentation_assets/`- und `*.pptx/ppsx`-Sachen.

### 3.2 Noch offen (bewusst NICHT gemacht)
- **Versionsbump + Build:** `version.json`, `backup.py`, Installer-`.iss` wurden **nicht**
  angefasst. Nur der Changelog-Text steht. Du bündelst SPs und baust selbst.
- **Hilfe-Screenshots:** Die Hilfe-Texte sind ergänzt, aber es liegen **keine neuen
  Screenshots** vom Auswertung-Reiter unter `assets/hilfe/kasse/`. (Passt zur Regel
  „Hilfe inkl. Screenshot mitpflegen“ — Bilder noch ablegen.)
- **Echter Klicktest** in der GUI steht aus (bisher nur Headless + reale PDF-Tests).

### 3.3 Lieferschein-Vorlage
- Die Vorlage enthält **Platzhalter-Firmendaten** („Ihre Firma GmbH“, Musterstraße …).
  Vor echtem Einsatz in `vorlagen/lieferschein.html` die eigenen Daten/Logo eintragen.
- Wurde schon mal eine alte `vorlagen/lieferschein.html` angelegt, wird sie **nicht**
  überschrieben — ggf. löschen, dann wird die neue Vorlage frisch kopiert.

### 3.4 Verhalten Tagesabschluss
- Der automatische Lauf braucht eine **laufende Kasse** (Desktop-App). Läuft sie am
  Abend nicht, wird der Abschluss beim **nächsten Start** nachgeholt.
- Ab 18 Uhr wird der heutige Abschluss bei jedem 30-Min-Tick neu geschrieben, damit auch
  späte Online-Verkäufe enthalten sind (überschreibt die Datei des Tages).

### 3.5 Verfall-Logik
- „Bald ablaufend“ = Verfall innerhalb **90 Tagen**. Schwelle sitzt in
  `app/kasse_reports.py` (`tage <= 90`) — bei Bedarf dort anpassen.
- Verfall „MM/JJ“ wird als **Ende des angegebenen Monats** gewertet (Ware bis Monatsende
  brauchbar).

### 3.6 PDF-Zeichensatz
- Die PDFs nutzen die fpdf2-Kernschrift (latin-1). Das **€-Zeichen** wird in PDFs als
  „EUR“ geschrieben (technisch bedingt). In der Bildschirm-Tabelle steht weiterhin „€“.

---

## 4. Wie du es testen kannst
1. Kasse starten (`start_kasse.py` bzw. NMGone → Kasse).
2. **Verkäufe** → einen Verkauf „In MSK erfasst“ setzen → Lieferschein-Abfrage prüfen.
3. **Verkauf** → Artikel mit Menge größer als Bestand → Drei-Wege-Abfrage prüfen.
4. **Auswertung** → Umsatz (Tag/Monat/Jahr), Verfall-Farben, Inventur, je „📄 PDF“.
5. **Auswertung → Umsatz** → Doppelklick auf einen Tag → Tagesabschluss-PDF.

---

## 17. Changelog-Text für die Kasse (zum Einfügen in app/changelog.py)
Als eigener Eintrag (z. B. `("V2.0 SP2", "2026-06-24", [ ... ])`) – Datum/Versionslabel nach
deiner Wahl. Diese Zeilen wurden durch die parallele Mitarbeiter-Arbeit aus „V2.0 SP1"
verdrängt und hier gesichert:

- Kasse: Lieferschein – wird beim „In MSK erfasst"-Setzen angeboten und ist jederzeit in der Verkaufs-Detailansicht erneut erzeugbar (Charge/Verfall, ohne Preise). Beim manuellen Erzeugen wird abgefragt, ob der Auftrag in MSK erfasst wurde.
- Kasse-Verkauf: reicht der Bestand nicht, kann die verfügbare Menge sofort als Bestellung geliefert und der Rest als Vorbestellung aufgenommen werden (oder nur der vorhandene Bestand abverkauft werden).
- Kasse: neuer Reiter „Auswertung" – Umsatz je Tag/Monat/Jahr (Anzahl Verkäufe, Anzahl Packungen, APU Brutto, Rabatt (Netto), APU Netto), Verfall (Zeitraum 3/6/9/12 Monate oder Alle; abgelaufen rot, bald gelb) und Inventur. Mit Summen Verkaufswert (APU×Bestand) und Lagerwert (EK×Bestand). Alles als PDF.
- Kasse-Tagesabschluss: eigener Reiter mit Kalender-Datumsauswahl, fortlaufender Nummer und Kennzahlen (Verkäufe, Packungen, APU-, Rabatt-, Umsatz-Summe). Wird jeden Abend automatisch (Uhrzeit einstellbar, Standard 18 Uhr) erzeugt und nummeriert; verpasste Vortage werden beim Start nachgeholt.
- Kasse: Auftrags-Verlauf (wer/was/wann) je Auftrag einsehbar und als PDF; Protokoll gefiltert als PDF exportierbar.
- Kasse: EK (Einkaufspreis) im Wareneingang erfassbar (mit APU vorbelegt) und beim Import übernommen; daraus überall der Lagerwert (EK×Bestand) zusätzlich zum Verkaufswert.
- Kasse: Artikel-Reiter mit Summenleiste (Gesamtbestand, Verkaufswert, Lagerwert), passt sich der Suche an. Kasse startet maximiert (Vollbild).
- Kasse: neuer Reiter „Defektmeldung" – Nichtverfügbarkeits-Bescheinigung für die Apotheke (Apotheke + Artikel wählen, Grund angeben).
- Kasse: neuer Reiter „Einstellungen" – Firmendaten (Dokumentkopf), Defektmeldung-Rechtstext und Tagesabschluss-Uhrzeit direkt in der App pflegbar.
- Kasse: robusterer Datenbankzugriff (busy_timeout) für den Parallelbetrieb von NMGone und Kasse.
- Kasse-Auswertung: einheitliches Design – Kennzahlen in Umsatz, Verfall und Inventur als Karten (wie beim Tagesabschluss), Reiter (Umsatz/Tagesabschluss/Verfall/Inventur) im NMGone-Blau gestylt.
- Kasse-Verkauf: „Freie Position" – frei benannter Posten mit eigenem Preis/Menge/Rabatt, nicht bestandsgeführt, wird gespeichert und erscheint mit Preis auf der Auftragsbestätigung (z. B. Botendienst-Zuschlag).
