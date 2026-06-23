# Konzept: Faktura-Modul für NMGone / Kasse (rechtskonform Deutschland)

**Stand:** 2026-06-23 · **Status:** Konzept – noch nichts gebaut · **Zweck:** Diskussionsgrundlage

> Ziel in einem Satz: NMGone soll künftig **rechtskonforme Rechnungen** schreiben können,
> dazu passende **Gutschriften** (Storno *und* Bonus), wahlweise **sofort** oder als
> **Monatsabrechnung**, plus eine **gestaffelte Monats-Bonus-Gutschrift** ab bestimmten
> Umsatzschwellen. Preise (APU) werden **zum Verkaufszeitpunkt eingefroren**.

---

## 1. Wichtige Begriffsklärung vorab (sonst redet man aneinander vorbei)

Das Wort **„Gutschrift"** hat in Deutschland **zwei völlig verschiedene Bedeutungen**.
Das muss im Programm sauber getrennt werden, sonst gibt es Ärger mit dem Finanzamt:

| Gemeint ist… | Korrekter Begriff im Programm | Wirkung |
|---|---|---|
| **Eine Rechnung korrigieren / stornieren** | **Storno-Rechnung / Korrekturrechnung** | Hebt eine fehlerhafte Rechnung (teilweise) auf |
| **Geld zurück / Rabatt nachträglich gewähren** (dein „X Betrag" / Staffel-Bonus) | **kaufmännische Gutschrift / Bonus-Gutschrift / Rückvergütung** | Mindert nachträglich das Entgelt → §17 UStG |
| **Self-Billing** (Kunde rechnet selbst ab) | **Gutschrift i.S.d. §14 Abs. 2 UStG** | Eher selten, hier vermutlich **nicht** gemeint |

➡️ **Empfehlung:** Im Programm nicht einfach „Gutschrift" auf das Dokument drucken,
sondern den jeweils zutreffenden Begriff. Eine als „Gutschrift" betitelte Storno-Rechnung
kann sonst als unberechtigter Steuerausweis (§14c UStG) gewertet werden.

**Was du beschreibst, sind eigentlich drei verschiedene Belege:**
1. **Rechnung** (Verkauf, mit aktuellem APU)
2. **Bonus-Gutschrift sofort** – nachträglicher Rabatt direkt zur Rechnung
3. **Monats-Bonus-Gutschrift** – gestaffelt ab Umsatzschwelle pro Monat

---

## 2. Was eine Rechnung in Deutschland zwingend enthalten muss (§14 UStG)

Pflichtangaben auf **jeder** Rechnung:

1. Vollständiger **Name + Anschrift** des leistenden Unternehmers (NMG) **und** des Kunden
2. **Steuernummer ODER USt-IdNr.** von NMG
3. **Ausstellungsdatum**
4. **Fortlaufende, einmalige Rechnungsnummer** (lückenlos, siehe §4)
5. **Menge + handelsübliche Bezeichnung** der Ware (PZN + Artikelname)
6. **Zeitpunkt der Lieferung/Leistung** (Liefer-/Leistungsdatum, ggf. = Rechnungsdatum)
7. **Entgelt, aufgeschlüsselt nach Steuersätzen** (Netto je Satz)
8. Anzuwendender **Steuersatz** + **Steuerbetrag** (bei Arzneimitteln i. d. R. **19 %**)
9. Hinweis bei **Steuerbefreiung** (falls zutreffend)
10. Ggf. Hinweis auf **Skonto / Zahlungsbedingungen**

**Kleinbetragsrechnung bis 250 € brutto (§33 UStDV):** vereinfacht – Name/Anschrift NMG,
Datum, Menge/Bezeichnung, Bruttobetrag, Steuersatz. (Trotzdem sauber durchnummerieren.)

**Storno/Korrektur** muss die **Original-Rechnungsnummer referenzieren**
(„storniert Rechnung Nr. …").

---

## 3. GoBD – die eigentliche technische Hürde

Rechnungen sind steuerrelevante Belege. Daraus folgt für die Software:

- **Unveränderbarkeit:** Eine **festgeschriebene** Rechnung darf **nie mehr geändert**
  werden. Korrektur **nur** über Storno + Neu. → Status-Modell „Entwurf → festgeschrieben".
- **Lückenlose Nummernkreise:** keine Lücken, keine Doppelvergabe, nachvollziehbar.
- **Aufbewahrung 10 Jahre** – inkl. originalgetreuer **PDF-Wiedergabe**.
- **Nachvollziehbarkeit / Audit-Trail:** wer hat wann was erstellt/storniert
  (ihr habt bereits ein Audit-Log-Muster – das wiederverwenden).
- **Verfahrensdokumentation:** kurze Beschreibung, wie Rechnungen entstehen (für Prüfer).

> Faustregel: Eine festgeschriebene Rechnung ist „in Stein gemeißelt".
> Es gibt **kein Löschen und kein Bearbeiten** – nur Storno und Neuanlage.

---

## 4. Nummernkreise (Detail, weil oft unterschätzt)

- Getrennte, jeweils **fortlaufende** Kreise je Belegart, z. B.
  - `RE-2026-00001` (Rechnung)
  - `ST-2026-00001` (Storno)
  - `GU-2026-00001` (Bonus-Gutschrift)
- Zähler **erst beim Festschreiben** ziehen (nicht beim Entwurf), damit keine Lücken
  durch verworfene Entwürfe entstehen.
- Jahreswechsel: Präfix mit Jahr, Zähler darf pro Jahr neu bei 1 starten – muss aber
  innerhalb des Jahres lückenlos sein.

---

## 5. APU zum Verkaufszeitpunkt „einfrieren" (zentrale Anforderung)

> „Rechnungen müssen immer den aktuellen APU beim Verkauf haben."

Heute liegt der APU in `tbl_nmg_stamm.apu` und **ändert sich**, wenn eine neue
APU/HAP-Liste importiert wird. Eine Rechnung muss aber **dauerhaft** den Preis zeigen,
der **am Verkaufstag** galt – auch wenn der Stammpreis morgen anders ist.

**Lösung – Preis-Snapshot:** Bei Rechnungserstellung wird der APU **als Wert in die
Rechnungsposition kopiert** (nicht per Join live nachgeschlagen). Die Rechnung verweist
also nicht auf den Stammdatenpreis, sondern trägt ihren eigenen, eingefrorenen Preis.

Das deckt zugleich die GoBD-Unveränderbarkeit ab und ist Standard bei jeder Faktura.

---

## 6. Belegfluss / Ablauf

### 6a. Sofort-Rechnung (Standardfall)
```
Verkauf an Kasse / Auftrag
        │  (APU zum Tag wird eingefroren)
        ▼
   Rechnung (Entwurf)  ──►  Festschreiben ──►  PDF + Nummer  ──►  Versand / Druck
        │
        └─(optional)─►  Bonus-Gutschrift sofort  (z. B. fixer/prozentualer Rabatt)
```

### 6b. Monatsabrechnung (Sammelrechnung)
```
Im Lauf des Monats: viele Verkäufe je Kunde werden gesammelt (offene Positionen)
        ▼
Monatslauf (z. B. 1. des Folgemonats):
   pro Kunde EINE Sammelrechnung über alle Positionen des Vormonats
        ▼
   danach automatisch: Monats-Bonus-Gutschrift gemäß Staffel (siehe §7)
```

➡️ Pro Kunde konfigurierbar: **Abrechnungsmodus = „sofort" oder „monatlich"**.

---

## 7. Bonus-Gutschrift mit Staffelung (dein „X Betrag", konfigurierbar)

Du kennst die genauen Stufen noch nicht – darum wird die **Staffel als Einstellung**
gebaut, die du jederzeit selbst pflegst. Vorschlag für die Konfiguration:

| Stufe | Monatsumsatz ab | Monatsumsatz bis | Bonus-Typ | Bonus-Wert |
|------:|----------------:|-----------------:|-----------|-----------:|
| 1 | 5.000 € | 9.999 € | Prozent | 1,0 % |
| 2 | 10.000 € | 19.999 € | Prozent | 2,0 % |
| 3 | 20.000 € | — | Prozent | 3,0 % |

- **Bonus-Typ** wahlweise **Prozent** vom Monatsumsatz **oder Fixbetrag**.
- 2–3 Stufen genügen anfangs, Tabelle ist beliebig erweiterbar.
- Staffel ist **datiert** (`gültig_ab`), damit historische Monate korrekt bleiben,
  wenn du die Stufen später änderst.
- Berechnungsbasis klären: **Netto- oder Bruttoumsatz?** (siehe Rückfragen §10)

**Steuerlich (wichtig):** Eine nachträgliche Bonus-Gutschrift mindert das Entgelt
→ **§17 UStG**: die **Umsatzsteuer wird anteilig mit korrigiert**. Die Gutschrift muss
also Netto + USt getrennt ausweisen, nicht nur einen Bruttobetrag.

---

## 8. Datenmodell (Vorschlag, schließt an bestehende Tabellen an)

Bestehend: `tbl_kunden_center` (Kunden), `tbl_nmg_stamm` (Artikel inkl. `apu`).

**Neu / zu ergänzen:**

```
tbl_kunden_center  (ERGÄNZEN um Faktura-Felder)
  + ust_id            -- USt-IdNr. des Kunden
  + steuernummer
  + rechnungsadresse  -- falls abweichend von plz/ort/strasse
  + abrechnungsmodus  -- 'sofort' | 'monatlich'
  + zahlungsziel_tage -- z. B. 14
  + bonus_aktiv       -- 0/1

fak_beleg              -- Belegkopf (Rechnung/Storno/Gutschrift)
  id, belegart ('rechnung'|'storno'|'bonus_gutschrift'),
  beleg_nr, nummernkreis, kunde_id,
  beleg_datum, leistungsdatum,
  zeitraum_von, zeitraum_bis,   -- nur bei Monatsrechnung
  bezug_beleg_id,               -- Storno/Gutschrift referenziert Rechnung
  netto_summe, ust_summe, brutto_summe,
  status ('entwurf'|'festgeschrieben'|'versendet'|'bezahlt'|'storniert'),
  pdf_pfad, erstellt_am, erstellt_von, festgeschrieben_am

fak_beleg_position     -- Belegzeilen mit eingefrorenem Preis
  id, beleg_id, pos_nr, pzn, bezeichnung, menge,
  apu_einzel,           -- << Snapshot zum Verkaufstag
  rabatt, netto_zeile, ust_satz, ust_zeile, brutto_zeile

fak_nummernkreis       -- lückenlose Zähler je Belegart/Jahr
  belegart, jahr, praefix, letzter_zaehler

fak_bonus_staffel      -- konfigurierbare Stufen (§7)
  id, gueltig_ab, schwelle_von, schwelle_bis,
  bonus_typ ('prozent'|'fix'), bonus_wert, bezeichnung

fak_audit              -- wer/wann/was (oder bestehendes Audit-Log nutzen)
```

---

## 9. E-Rechnung – Pflicht nicht verpassen (B2B)

Seit **01.01.2025** müssen Unternehmen **E-Rechnungen empfangen** können
(strukturiertes Format **XRechnung / ZUGFeRD**, nicht bloß PDF).
**Versand** wird stufenweise Pflicht:
- ab **01.01.2027** für Unternehmen mit Vorjahresumsatz > 800.000 €
- ab **01.01.2028** für **alle** (B2B)

➡️ Für ein neues Faktura-Modul heißt das: **von Anfang an ZUGFeRD-fähig planen**
(PDF mit eingebettetem XML). Kein Muss für die erste Version, aber das Datenmodell
sollte alle dafür nötigen Felder schon vorhalten (USt-IdNr., Einheiten, Steuersätze).

---

## 10. Offene Rückfragen an dich (für morgen)

1. **„Gutschrift nach jeder Rechnung"** – ist das ein **fester Rabatt** auf jede
   Rechnung (z. B. immer 2 %), oder etwas anderes? Begriff = Bonus oder Storno?
   → *Stand 2026-06-23: Arbeitsbegriff bleibt vorerst **„Gutschrift"** (genaue
   Benennung folgt). Achtung: auf dem gedruckten Beleg ggf. präziser betiteln,
   siehe §1 / §14c UStG.*
2. **Staffel-Basis:** Bonus auf **Netto- oder Bruttoumsatz** des Monats?
   Und: **prozentual** oder **Fixbetrag** je Stufe?
3. **Staffel-Schwellen:** sobald du die Zahlen hast (Programm bleibt konfigurierbar).
4. **Monatslauf:** Stichtag (z. B. 1. des Folgemonats)? Sammelrechnung pro Kunde?
5. **Steuersatz:** durchgehend 19 %, oder gibt es Positionen mit 7 % / steuerfrei?
6. **NMG-Stammdaten:** USt-IdNr./Steuernummer, Firmenanschrift, Logo, Bankverbindung
   für den Rechnungsfuß – liegen vor?
7. **Wo verkauft wird:** entsteht die Rechnung aus der **Kasse**, aus **Aufträgen**
   (`auftrag.py`) oder aus beidem?

---

## 11. Vorgeschlagene Umsetzung in Phasen (wenn es losgeht)

- **Phase 1 – Sofort-Rechnung:** Belegkopf/-positionen, APU-Snapshot, Nummernkreis,
  Festschreiben, PDF-Druck mit allen §14-Pflichtangaben, Storno.
- **Phase 2 – Gutschriften:** Bonus-Gutschrift sofort, §17-konforme USt-Korrektur.
- **Phase 3 – Monatsabrechnung:** Sammeln offener Positionen, Monatslauf,
  Sammelrechnung pro Kunde.
- **Phase 4 – Staffel-Bonus:** konfigurierbare Stufen, automatische Monats-Bonus-Gutschrift.
- **Phase 5 – E-Rechnung:** ZUGFeRD/XRechnung-Export.

> Diese Reihenfolge ist so gewählt, dass nach jeder Phase etwas **fertig Nutzbares**
> entsteht und der rechtskonforme Kern (Phase 1) zuerst steht.

---

*Hinweis: Das ist ein Fachkonzept, keine Steuerberatung. Die finale Rechnungsvorlage
und die §17-Behandlung der Boni sollten vor Produktivstart kurz mit dem Steuerberater
abgestimmt werden.*
