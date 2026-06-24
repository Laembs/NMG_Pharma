# Mitarbeiter · Organigramm · Abwesenheiten — Dokumentation

Stand: 24.06.2026 · Branch `feat/bestell-app`

Diese Datei hält fest, **was tatsächlich gebaut wurde**, wo es liegt und **worauf
du besonders achten solltest**. Sie ist absichtlich ehrlich getrennt nach
„läuft schon in der echten App" vs. „nur Testversion".

---

## 1 · Was in dieser Sitzung NEU entstanden ist (auf der Platte)

| Datei | Was | Git-Status |
|---|---|---|
| `start_organigramm_test.py` | **Eigenständige Testversion** (Tkinter) mit Organigramm-Board **und** Abwesenheits-Kalender | untracked (neu) |
| `organigramm_test.db` | Test-Datenbank mit Demo-Daten für das Test-Script | **ignoriert** (`*.db`) |
| `docs/organigramm_konzept_druck.html` | Druckbares Konzept (HTML) | untracked (neu) |
| `docs/organigramm_konzept.pdf` | Dasselbe als PDF (4 Seiten) | untracked (neu) |
| `docs/organigramm_prototyp.html` | Interaktiver HTML-Prototyp (UX-Spielwiese) | bereits getrackt |
| `docs/Mitarbeiter_Organigramm_Doku.md` | Diese Doku | untracked (neu) |

**Das Herzstück ist `start_organigramm_test.py`.** Es ist komplett eigenständig,
benutzt eine eigene Test-DB und **fasst die echte App nicht an**.

---

## 2 · Was die echte App (`app/gui.py`) heute schon kann

Die **Phase 1** (Mitarbeiter-Zusatzfelder + Vorgesetzten-Matrix) ist bereits
Teil des letzten Releases (V2.0-Commit) — also schon vor dieser Sitzung gebaut.
Konkret im `Mitarbeiter`-Center:

- `_ensure_mitarbeiter_phase1_tables()` legt 3 Tabellen an:
  - `tbl_mitarbeiter_feld` / `tbl_mitarbeiter_wert` — frei definierbare Zusatzfelder
  - `tbl_mitarbeiter_vorgesetzter` — mehrere Vorgesetzte je Person, mit Art + `ist_primaer`
- `_mitarbeiter_felder_dialog()` — Zusatzfelder verwalten
- `_mitarbeiter_detail_dialog()` — Reiter Stammdaten / Zusatzfelder / Vorgesetzte

> ⚠️ Diese Phase-1-Funktionen wurden **noch nie real in der GUI durchgeklickt**
> (nur Syntax- und SQL-Test). Ein echter Smoke-Test steht aus.

**Was die echte App NOCH NICHT hat:** das grafische Organigramm-Board und die
Abwesenheits-/Urlaubsplanung. Beides existiert bisher **nur** im Test-Script.

---

## 3 · Die Testversion im Detail (`start_organigramm_test.py`)

Start: `python start_organigramm_test.py` · Umschalter oben rechts im Fenster.

### Ansicht „🗺 Organigramm"
- Post-it-Karten je Mitarbeiter, **mit der Maus verschiebbar** (Position gespeichert)
- **Doppelklick** auf Karte → Name / Rolle / Abteilung bearbeiten
- **🔗 Verbinden** einschalten → Start-Karte, dann Ziel-Karte klicken; Art oben
  wählen (disziplinarisch / fachlich / Vertretung); Abfrage „als primär?"
- **× auf einer Linie** → Verbindung entfernen
- **Teilbereich-Dropdown** → Gesamt oder einzelne Abteilung
- **Auto-Layout** → ordnet sichtbare Karten als Baum aus den primären Kanten
- Karte zeigt unten den **primären Vorgesetzten** (Rückspiegelung)
- Linienstil = Art: durchgezogen+Pfeil / gestrichelt / gepunktet

### Ansicht „🗓 Abwesenheiten" (neu auf deinen Wunsch)
- **Team-Kalender** im Monatsraster: Zeilen = Mitarbeiter, Spalten = Tage
- Wochenenden grau, „heute" hervorgehoben
- **Farbige Balken** je Abwesenheit, feste Arten:
  - 🟩 Urlaub · 🟥 Krankheit · 🟧 Fortbildung · 🟪 Sonstiges
- **➕ Abwesenheit eintragen** → Mitarbeiter, Art, Von–Bis (`JJJJ-MM-TT`), Notiz
- **Balken anklicken** → Detail ansehen / löschen
- **◀ ▶ / Heute** → Monat wechseln, **Teilbereich-Filter**
- Pro Person Zeile **„Urlaub X T · Krank Y T (Jahr)"** — zählt nur Werktage
- Tabelle: `tbl_abwesenheit(mitarbeiter_id, art, von, bis, notiz)`

---

## 4 · WORAUF DU BESONDERS ACHTEN SOLLTEST

1. **Dieser Ordner ist eine „- Kopie".** Beobachtung: Änderungen an *getrackten*
   Dateien (z. B. `app/gui.py`) tauchen hier **nicht als Diff** auf — der Stand
   gleicht dem Commit. Es sieht so aus, als würde dieser Ordner mit dem
   Hauptrepo abgeglichen. **Echte Code-Änderungen also im Hauptarbeitsordner
   machen**, nicht hier. Neue, *untrackte* Dateien (Test-Script, Konzepte)
   bleiben dagegen erhalten.

2. **Organigramm-Board + Abwesenheiten sind NUR Testversion.** Sie sind noch
   **nicht** in der echten App. Übertragung nach `gui.py` ist der nächste Schritt.

3. **`organigramm_test.db` ist Wegwerf-Test.** Bereits durch `*.db` in
   `.gitignore` → wird nicht committet. **Keine echten Mitarbeiter / echten
   Krankheitsdaten** dort eintragen — es ist nur eine Spielwiese.

4. **Datenschutz (wichtig beim echten Einbau):** Mitarbeiterdaten und v. a.
   **Krankheitstage** sind sensibel (DSGVO). Empfehlung: nur „krank ja/nein"
   speichern, **keinen Krankheitsgrund**. Beim Online-/Mehrbenutzer-Schritt
   Zugriffsrechte klar regeln (wer darf wessen Abwesenheiten sehen?).

5. **Phase 1 in der echten App ist ungetestet in der GUI.** Vor dem Weiterbauen
   einmal real durchklicken (Felder anlegen, Vorgesetzten zuordnen, speichern).

6. **Zwei Datenquellen für Mitarbeiter.** In der echten App kommen Karten aus
   `tbl_mitarbeiterprofil` (login-basiert) **und** `tbl_mitarbeiter`
   (id-basiert). Zusatzfelder/Vorgesetzte/Abwesenheiten hängen an
   `tbl_mitarbeiter` (id). Beim Online-Schritt müssen die Quellen
   zusammengeführt werden.

7. **Beim Portieren nach `gui.py`:** `tbl_abwesenheit` in
   `_ensure_mitarbeiter_phase1_tables` ergänzen; Board und Kalender als
   *zusätzliche* Ansichten im Mitarbeiter-Center (Karten-Raster bleibt).

8. **Datumsformat** überall `JJJJ-MM-TT` (ISO), damit Sortierung/Vergleich passt.

---

## 5 · Nächste sinnvolle Schritte (Vorschlag)

1. Phase 1 in der echten App real durchklicken (Smoke-Test).
2. Optik/Bedienung der Testversion final abnehmen.
3. Organigramm-Board nach `gui.py` übertragen (zusätzliche Ansicht).
4. Abwesenheits-Kalender nach `gui.py` übertragen (+ `tbl_abwesenheit`).
5. Optional: Resturlaub-Anspruch, Feiertage, Überlappungs-Warnung, weitere Arten.
