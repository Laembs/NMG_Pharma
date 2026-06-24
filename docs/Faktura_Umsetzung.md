# Faktura-App – Umsetzungs- & Übergabe-Doku

**Stand:** 2026-06-24 · **Status:** Prototyp (lauffähig, getestet) · **Autor-Notiz:** Diskussions-
und Arbeitsstand, noch nicht für den Produktivbetrieb freigegeben.

---

## 1. Was gebaut wurde (Überblick)

Eine **eigenständige App** „NMG Faktura" für rechtskonforme Rechnungen und Gutschriften
(Deutschland), nach dem Muster der Kasse-App. Eigenes Fenster, eigenes Taskleisten-Icon,
gemeinsame Datenbank mit NMGone.

**Abgedeckter Kreislauf:** Rechnung → (Storno) / (Sofort-Gutschrift) → Monats-Bonus aus Staffel.

---

## 2. Geänderte / neue Dateien

| Datei | Art | Inhalt |
|---|---|---|
| `app/faktura_app.py` | **neu** | Die komplette App (UI + Logik + PDF) |
| `start_faktura.py` | **neu** | Start-Einstieg (wie `start_kasse.py`) |
| `app/migrations.py` | geändert | Block „Faktura-Tabellen" (7 Tabellen) ergänzt |
| `assets/Faktura.ico` | neu | Fenster-/App-Icon |
| `docs/Konzept_Faktura.md` | neu | Fachkonzept (Recht, Datenmodell, Phasen) |
| `docs/Faktura_Handout.html/.pdf` | neu | Druckbares 3-Seiten-Handout mit Bildern |
| `docs/Faktura_Umsetzung.md` | neu | **dieses Dokument** |

**Es wurde nichts an bestehender NMGone-/Kasse-Logik geändert** – nur additiv ergänzt.

---

## 3. So startet man die App

```
python start_faktura.py
```

Beim Start laufen automatisch die Migrationen (legen fehlende Tabellen an). Es öffnet sich
ein eigenes Fenster „NMG Faktura". Die Daten landen in derselben SQLite-DB wie NMGone
(Dev: `data/nmg_startdatenbank.sqlite`, Produktiv: `C:/ProgramData/NMGone/...`).

---

## 4. Die Seiten der App

1. **Startseite** – 4 Kennzahlen-Kacheln (Rechnungen gesamt, Monatsumsatz netto, offene
   Entwürfe, Gutschriften), Schnellaktionen, „Zuletzt erstellt"-Liste.
1b. **Aufträge** (linke Navigation „Aufträge") – alle Apotheken, an die verkauft wurde
   (Quelle: Kasse-Verkäufe `tbl_bestellungen` + `tbl_bestellpositionen`): Kundennr.,
   Apotheke, Aufträge, Artikel, Umsatz, Umsatz nach Rabatt. Optionaler Zeitraum-Filter
   (JJJJ / JJJJ-MM), sortierbare Spalten, Summenzeile. **Stornierte Aufträge und abgesagte
   Positionen werden ausgeschlossen.** Aktionen:
   - **Rechnung erstellen (Auswahl):** Sammelrechnung über alle Verkäufe des gewählten
     Kunden im Zeitraum (Positionen je PZN/Preis/Rabatt zusammengefasst, APU eingefroren),
     festschreiben + PDF in die Ablage.
   - **Für jetzt entfernen:** blendet den Kunden temporär aus (z. B. „später drucken").
     „Aktualisieren" holt alle wieder zurück (rein im Arbeitsspeicher, nicht gespeichert).
2. **Rechnungen** – Belegliste; Doppelklick öffnet das PDF; Buttons für **Gutschrift**
   (fester €-Betrag) und **Storno** zur ausgewählten Rechnung. Ist die **automatische
   Gutschrift** aktiv (Einstellungen · Firmendaten), wird beim Festschreiben jeder Rechnung
   (auch Sammelrechnung) automatisch eine verknüpfte Gutschrift mit PDF erzeugt.
   Über die Buttons **📂 Rechnungen / Gutschriften / Quartalsvergütung** öffnet man die
   jeweiligen Ablageordner direkt im Explorer; **📁 Ordner der Auswahl** öffnet den Ordner
   des markierten Belegs.
3. **Neue Rechnung** – Kundensuche (aus `tbl_kunden_center`), Artikelsuche (aus
   `tbl_nmg_stamm`), **APU wird beim Hinzufügen eingefroren**, Live-Summen,
   „Als Entwurf speichern" und „Festschreiben + PDF".
4. **Quartalsvergütung** – Jahr + Quartal wählen → **Quartals**-Umsatz je Kunde aus den
   **festgeschriebenen Rechnungen** (stornierte zählen nicht, da das Original auf
   „storniert" steht; Netto/Brutto per Einstellung) → passende Staffel-Stufe → Vorschau →
   Vergütungen erzeugen (eigene Nummern `QV-…`, §17-USt-Split, Doppel-Schutz pro Quartal).
   *Hinweis: ersetzt den früheren „Monats-Bonus" – die Vergütung ist jetzt quartalsweise.*
5. **Staffel** – Euro-Stufen pflegen (ab/bis/Bonus) + Schalter **Netto/Brutto**.
6. **Einstellungen** – beim Klick klappt in der Navigation ein **Unterbaum** auf mit drei
   scrollbaren (Mausrad) Unterseiten:
   - **Firmendaten:** Stammdaten (Steuernr., USt-IdNr., Adresse, Bank, Logo, Zahlungsziel,
     USt-Satz), **Automatische Gutschrift je Rechnung** (Schalter „Automatische Gutschriften
     erzeugen" + Berechnung % vom Netto / % vom Brutto / Fixbetrag / **Auftragsrabatt (aus
     Positionen)** + Option „beim Stornieren der Rechnung die zugehörigen Gutschriften –
     automatisch UND manuell – mit stornieren") und **Mitarbeiter** (Name/E-Mail,
     Auto-Zuordnung über Windows-Benutzer).
   - **Belegnummern (frei konfigurierbar):** je Belegart ein Format mit Platzhaltern
     `{JJJJ}` `{JJ}` `{MM}` `{NR}` `{NR:5}` (z. B. `RE-{JJJJ}-{NR:5}`). Fehlt `{NR}`, wird der
     Zähler 5-stellig angehängt (Lückenlosigkeit bleibt gewahrt).
   - **Layouts:** Akzentfarbe, Logo-Position, Spalten an/aus (APU, USt), Titel und Kopftext
     **je Belegart**, gemeinsamer Fußtext. **Eine Vorlage – Rechnung, Gutschrift und
     Quartalsvergütung nutzen sie.** Button **Vorschau (PDF)** + **Layout frei gestalten**.

**Layout frei gestalten (Drag & Drop):** Editor mit A4-Fläche (Fenster wird passend zur
Arbeitsfläche skaliert, öffnet nicht hinter der Taskleiste), auf der alle Blöcke (Logo,
Absender, Belegtitel, Empfänger, Kopftext, Positionstabelle, Summen, Sachbearbeiter, Fußtext,
Bank) mit **Beispiel-Inhalten** angezeigt werden. Blöcke lassen sich **frei verschieben** und am
**Griff unten rechts in der Größe ziehen**. Speichern legt Position+Größe (mm) als JSON ab
(`tpl_layout`) und aktiviert „Freies Layout"; die Belege werden dann absolut platziert. Checkbox
**„Freies Layout verwenden"** schaltet zwischen freiem und klassischem Fließlayout um.

**Eigene Felder:** Im Editor lassen sich zusätzliche Felder anlegen:
- **Freitext** – beliebiger Text (darf Platzhalter wie `{kunde_name}` enthalten).
- **Datenfeld** – Auswahl aus Datenbank-/Beleg-Platzhaltern (Kunde, Beleg, Firma,
  Sachbearbeiter), die beim Druck mit echten Werten gefüllt werden.
Eigene Felder werden in `tpl_zusatzfelder` (JSON) gespeichert, sind verschieb-/größenbar und
über „Feld löschen" entfernbar (Standardblöcke bleiben erhalten).
   - **Mitarbeiter** (Name/E-Mail, Auto-Zuordnung über Windows-Benutzer).

---

## 4b. Ablage der PDFs (Ordnerschema + Dateiname)

Belege werden beim Festschreiben automatisch abgelegt unter `OUTPUT_DIR` (Dev: `ausgaben/`,
Produktiv: `C:/ProgramData/NMGone/ausgaben/`):

| Belegart | Ordner |
|---|---|
| Rechnung / Storno | `Rechnungen/<Jahr>/<Monat>/` |
| Gutschrift | `Gutschriften/<Jahr>/<Monat>/` |
| Quartalsvergütung | `Quartalsverguetung/<Jahr>/Q<n>/` |

**Dateiname:** `Kundennummer;Rechnungsnummer;Apothekenname;Datum.pdf`
(unzulässige Windows-Zeichen werden zu `-`, `;` bleibt als Trennzeichen).

**Künftig SharePoint/OneDrive:** Die Basis ist `OUTPUT_DIR` aus `config.py`. Sobald der
zentrale Datenpfad dort umgelenkt wird (vgl. `Konzept_Zentrale_Daten.md`), landen die PDFs
automatisch im selben Schema auf SharePoint/OneDrive – ohne Änderung an der Faktura-App.

---

## 5. Datenmodell (neue Tabellen)

| Tabelle | Zweck |
|---|---|
| `tbl_faktura_einstellungen` | Firmenstammdaten als Schlüssel/Wert |
| `tbl_faktura_mitarbeiter` | Sachbearbeiter (Name/E-Mail, `benutzer` = Windows-Login) |
| `tbl_faktura_belege` | Belegkopf (Rechnung/Storno/Gutschrift), Kunden-Snapshot, Summen, Status |
| `tbl_faktura_positionen` | Positionen mit **`apu_einzel` = eingefrorener APU** |
| `tbl_faktura_nummernkreis` | Lückenlose Zähler je Belegart + Jahr (RE/GU/ST) |
| `tbl_faktura_bonus_staffel` | Euro-Stufen für den Monats-Bonus |
| `tbl_faktura_log` | Audit-Trail (wer/wann/was) |

---

## 6. Umgesetzte rechtliche/technische Kernprinzipien

- **APU-Snapshot:** Der Preis wird je Position als Wert kopiert (kein Live-Join auf
  `tbl_nmg_stamm`). Spätere Preislisten-Importe verändern alte Belege nicht.
- **Unveränderbarkeit (GoBD):** Festgeschriebene Belege werden nicht editiert/gelöscht;
  Korrektur nur über **Storno + Neu**.
- **Lückenlose Nummern:** Die Belegnummer wird **erst beim Festschreiben** gezogen
  (verworfene Entwürfe erzeugen keine Lücken).
- **§17 UStG:** Gutschriften/Boni weisen Netto + USt getrennt aus (Entgeltminderung).
- **Auto-Sachbearbeiter:** Beim Festschreiben wird der angemeldete Mitarbeiter automatisch
  auf den Beleg geschrieben (Name + E-Mail).

---

## 7. ⚠️ Worauf du SPEZIELL achten solltest

### Rechtlich / steuerlich (vor Produktivstart klären)
1. **Begriff „Gutschrift":** Aktuell steht „Gutschrift" auf dem Beleg. In Deutschland
   doppeldeutig (Storno vs. Bonus). Eine als „Gutschrift" betitelte Korrektur kann als
   unberechtigter Steuerausweis (§14c UStG) gewertet werden. **Endgültige Benennung +
   Rechnungsvorlage mit dem Steuerberater abstimmen.**
2. **USt-Satz der Boni:** Der §17-Split rechnet mit dem **Standard-USt-Satz** aus den
   Einstellungen (Default 19 %). Bei gemischten Sätzen (7 %/steuerfrei) ist der Bonus-Split
   nur näherungsweise – dann manuell prüfen.
3. **Pflichtangaben:** Damit eine Rechnung §14-konform ist, **müssen die Firmenstammdaten
   (USt-IdNr. ODER Steuernummer, vollständige Adresse) in den Einstellungen ausgefüllt
   sein.** Festschreiben warnt zwar, lässt aber zu – nicht versehentlich ohne Stammdaten
   festschreiben.

### Technisch
4. **PDF braucht Edge oder Chrome:** Die PDF-Erzeugung nutzt Edge/Chrome im Headless-Modus.
   Ist keiner installiert, entsteht ersatzweise eine **HTML-Datei** (kein PDF). Auf den
   Zielrechnern muss ein Browser vorhanden sein.
5. **Nummernkreise nicht manuell anfassen:** `tbl_faktura_nummernkreis` niemals von Hand
   ändern – das bricht die Lückenlosigkeit. (Für die Demo stehen die Zähler aktuell auf 0,
   die erste echte Rechnung wird also `RE-2026-00001`.)
6. **Festgeschriebene Belege sind final:** Es gibt bewusst **kein Bearbeiten/Löschen**.
   Fehler werden über Storno korrigiert. Das ist gewollt (GoBD), nicht vergessen.
7. **Geteilte Datenbank:** Faktura nutzt dieselbe DB wie NMGone/Kasse. Kunden kommen aus
   `tbl_kunden_center`, Artikel/APU aus `tbl_nmg_stamm`. **Sind diese leer, findet die
   Suche nichts** – zuerst Kunden/Artikel in NMGone/Kasse pflegen.
8. **Backup:** Vor dem ersten echten Einsatz ein DB-Backup ziehen (NMGone-Backup nutzen).

### Abrechnungs-Status der Aufträge (neu)
- Die **Aufträge-Seite zeigt nur OFFENE** (noch nicht abgerechnete) Verkäufe. Beim Erstellen
  einer Sammelrechnung werden die enthaltenen Aufträge mit der Rechnung verknüpft
  (`tbl_bestellungen.faktura_beleg_id`) und verschwinden aus der Liste – so kann derselbe
  Verkauf nicht doppelt abgerechnet werden.
- Ein **Storno gibt die Aufträge wieder frei** (Verknüpfung zurückgesetzt) → sie tauchen
  erneut als offener Auftrag auf und sind in der Kasse über die Auftragsnummer weiter
  bearbeitbar. Danach kann eine neue/korrigierte Rechnung erstellt werden.
- Das **Rechnungsdatum** ist wählbar (Aufträge-Seite und „Neue Rechnung"); es bestimmt auch
  den Monatsordner der Ablage. Leer = heute.

### Prototyp-Grenzen (bewusst noch nicht gebaut)
9. **Vergütung = Bruttobetrag:** Der Staffel-Betrag wird als Brutto behandelt und intern in
    Netto + USt gesplittet. Falls der Bonus als Netto gemeint ist, muss die Logik angepasst
    werden (1 Zeile).
10. **Kein Positions-Rabatt im UI:** Das Rabatt-Feld je Position existiert im Datenmodell und
    wird bei der **Sammelrechnung** aus den Verkäufen übernommen, in der **manuellen** „Neue
    Rechnung" aber noch nicht eingegeben (dort immer 0 %).
11. **Nicht im NMGone-Dashboard:** Faktura ist (wie gewünscht) standalone und noch nicht als
    Kachel in NMGone verdrahtet.

---

## 8. Getestet (automatisierte Smoke-/E2E-Tests, danach wieder aufgeräumt)

- Migration legt alle 7 Tabellen an ✓
- Einstellungen speichern/laden, Mitarbeiter-Auto-Erkennung ✓
- Nummernkreis lückenlos (RE-2026-00001, 00002 …) ✓
- Alle 6 GUI-Seiten bauen fehlerfrei auf ✓
- PDF-Erzeugung end-to-end (echte .pdf) ✓
- E2E: 2 Rechnungen (13.600 € netto) → Monats-Bonus erkennt Stufe 2 (150 €) → Gutschrift ✓
- Storno: Original → „storniert", Storno-Beleg mit negativen Beträgen ✓

---

## 9. Sinnvolle nächste Schritte

- Echte **Sammelrechnung** (offene Verkäufe sammeln → eine Monatsrechnung je Kunde)
- **Positions-Rabatt** im UI anbinden
- **ZUGFeRD/XRechnung**-Export (B2B-Pflicht: Empfang seit 2025, Versand ab 2027/28)
- **Kachel-Anbindung** in NMGone (Nav-Dispatch / Apps-Tile / Dashboard)
- **Hilfe-App** um eine Faktura-Seite + Screenshot ergänzen
- Rechnungsvorlage final mit Steuerberater abstimmen
