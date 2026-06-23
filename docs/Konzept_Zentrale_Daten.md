# Konzept & Lösungsvorschlag: Zentrale Daten für NMGone + Kasse

**Stand:** 2026-06-22 · **Status:** Konzept (noch nichts gebaut)

---

## 1. Das Ziel in einem Satz

> Was an einem Ort eingepflegt wird (z. B. eine Rabatt-Liste oder ein neuer Kunde),
> sollen **alle Mitarbeiter und alle Kassen sofort sehen** – egal ob im Büro an
> NMGone oder unterwegs an der Kasse.

NMGone bleibt dabei ein lokal installiertes Programm. Die Kasse soll europa-/weltweit
funktionieren.

---

## 2. Grundprinzip: Programm lokal, Daten zentral

Der wichtigste Gedanke, der alles löst:

- **„Lokal" heißt: das *Programm* ist lokal installiert und läuft lokal.**
- **Die geteilten *Daten* liegen zentral an *einem* Ort.**

Denn: Daten, die nur auf einem PC liegen, kann per Definition niemand sonst sehen.
Sobald mehrere Mitarbeiter dieselbe Rabatt-Liste oder denselben Kunden sehen sollen,
muss es *einen gemeinsamen Datenort* geben. NMGone darf lokal bleiben – aber seine
geteilten Daten nicht.

---

## 3. Zwei Sorten Daten (der Schlüssel)

Nicht alle Daten müssen geteilt werden. Wir trennen bewusst:

| | **Sorte 1 – Gemeinsame Betriebsdaten** | **Sorte 2 – Schwere Analyse-Daten** |
|---|---|---|
| Beispiele | Artikel, Kunden, Rabatte/Konditionen, Verkäufe, Lager, Wareneingang | große Vergleichs-/Austauschtabellen, `pzn_norm`-Logik, Auswertungen |
| Liegt | **zentral** (Cloud) | **lokal** in NMGone |
| Zugriff | alle lesen **und** schreiben (NMGone + Kasse) | nur NMGone, schnell vor Ort |
| Warum | müssen überall aktuell sein | NMGone-intern, über Distanz zu langsam |

So sieht jeder neue Kunde / jede neue Rabatt-Liste sofort überall – ohne dass NMGones
schwere Analyse über das Internet kriechen muss.

---

## 4. Architektur

```
   NMGone (lokal, Büro)                      Kasse (Web/PWA, weltweit)
   ────────────────────                      ─────────────────────────
   • schwere Analyse  ── lokal               • Verkauf, Lager
   • Betriebsdaten  ─┐                        ─┐
                     │                          │
                     ▼                          ▼
            ┌───────────────────────────────────────────┐
            │   API-Backend (EU-Cloud, abgesichert)      │
            │   – einziger Zugang zu den zentralen Daten │
            └───────────────────────────────────────────┘
                                 │
                                 ▼
            ┌───────────────────────────────────────────┐
            │   Zentrale Datenbank (PostgreSQL, EU)       │
            │   = die EINE Quelle der Wahrheit            │
            └───────────────────────────────────────────┘
```

- **Zentrale DB** (PostgreSQL) in einer **EU-Region** (DSGVO – Pharma-/Kundendaten).
- **API-Backend** (FastAPI, bleibt Python) ist der einzige Zugang zur DB. Die DB selbst
  steht nie offen im Internet → sicher.
- **NMGone** liest/schreibt die Betriebsdaten über das API; die schwere Analyse bleibt
  lokal.
- **Kasse** läuft als Web-App/PWA (kein App-Store nötig) und nutzt dasselbe API.

---

## 5. Der entscheidende Ablauf: lokal importieren → alle sehen es

**Beispiel Rabatt-Liste:**
1. Du klickst in NMGone wie gewohnt auf „Importieren" und wählst deine Excel/CSV-Datei
   (die Datei liegt lokal – das bleibt so).
2. NMGone liest die Datei lokal ein und **schreibt das Ergebnis in die zentrale DB**
   (statt wie bisher in die lokale Datei).
3. Ab diesem Moment sehen **alle NMGone-Mitarbeiter und alle Kassen** die neuen Rabatte –
   automatisch, ohne Extra-Schritt, weil alle denselben zentralen Ort lesen.

**Beispiel neuer Kunde:**
- Anlegen in NMGone *oder* unterwegs in der Kasse → das API vergibt **zentral die
  Kundennummer** (und daraus die MSK-Nr 216+Nr) → der Kunde ist sofort überall sichtbar.
- Zentrale Nummernvergabe heißt: keine zwei Stellen vergeben dieselbe Nummer.

Kurz: **Der Import-Knopf bleibt in NMGone – nur sein Ziel wechselt** von „lokale Datei"
zu „zentrale DB".

---

## 6. Warum das sicher ist (Abgrenzung zur OneDrive-Idee)

Das ist **nicht** das frühere „SQLite-Datei auf SharePoint teilen" (das führt zu
Konflikt-Kopien und Datenverlust). Hier gilt:

- Es gibt **keine geteilte Datei**, sondern eine **echte Datenbank** mit Transaktionen
  und Sperren.
- **Neue Datensätze anlegen** macht praktisch nie Konflikte (man fügt hinzu, streitet
  nicht um eine Zeile).
- **Gleichzeitiges Bearbeiten** desselben Datensatzes fängt die Datenbank von Haus aus ab.

---

## 7. Wie wir das zusammen umsetzen (Phasen)

| Phase | Was passiert | Wo | Risiko | Wer |
|---|---|---|---|---|
| **A – Aufräumen** | Datenzugriff bündeln (heute 149 verstreute DB-Verbindungen → eine Schicht); Logik aus der Oberfläche lösen. **Verhalten bleibt identisch.** | lokal | sehr gering, nebenbei machbar | ich baue, du testest |
| **B – Backend** | API lokal gegen die bestehende Datenbank bauen | lokal | gering | ich baue |
| **C – Cloud** | Zentrale DB + EU-Hosting einrichten, Daten umziehen | Cloud | mittel | ich baue, du besorgst Zugänge |
| **D – Umstellen** | NMGone (Betriebsdaten) + Kasse auf das API umstellen; Kasse als Web/PWA | beides | mittel | ich baue |
| **E – Go-Live** | Test mit echten Daten, DSGVO-Check, Mitarbeiter-Zugänge, live | Cloud | – | gemeinsam |

**„Zusammen" konkret:** Ich übernehme Umbau und Programmierung. Du entscheidest die
Fachlogik (z. B. „dürfen Kunden auch unterwegs angelegt werden?"), testest mit echten
Abläufen und besorgst die Cloud-Zugänge. Phase A können wir **sofort und nebenbei**
starten, während du normal weiterprogrammierst.

---

## 8. Was angeschafft werden muss (erst ab Phase C)

| Posten | Wofür | Grobe Kosten |
|---|---|---|
| EU-Hosting (Server) | API laufen lassen | ~5–30 €/Monat |
| Managed PostgreSQL (EU) | zentrale DB, DSGVO | ~15–50 €/Monat |
| Domain + TLS | sichere https-Verbindung | ~10 €/Jahr, TLS meist gratis |
| Backup/Monitoring | Datensicherheit | oft inklusive |

Realistischer Start: **~30–80 €/Monat**. Kein App-Store-Account nötig (PWA). Das Teure
ist nicht die Anschaffung, sondern die Entwicklungszeit.

---

## 9. Erste Schritte (was jetzt zu tun ist)

1. **Phase A starten** – als risikoarmen Pilot die Kasse auf die gebündelte
   Datenschicht umstellen (ein sauberer, abgegrenzter Commit als Vorlage für alle
   weiteren Module). Stört die laufende Arbeit nicht.
2. **Eine Entscheidung von dir:** Sitzen die NMGone-Mitarbeiter **alle an einem
   Standort** (dann reicht der zentrale Ort evtl. als Server bei euch im Haus) oder
   **verteilt** (dann von Anfang an Cloud)? Davon hängt der genaue Aufbau in Phase C ab.

---

## 10. Offene Entscheidungen (sammeln, nicht jetzt klären)

- Standort der NMGone-Mitarbeiter (siehe oben).
- Dürfen Kunden auch **von der Kasse unterwegs** angelegt werden? (Ja ist möglich –
  zentrale Nummernvergabe regelt es.)
- Feinschnitt: welche Tabellen genau zentral vs. lokal (grobe Liste steht in Abschnitt 3,
  Feinschliff passiert in Phase A automatisch).
- Anmeldung der Mitarbeiter (Login/Tokens) – Detail für Phase B/C.

---

## 11. Zweistufiger Start: erst eigene Domain, dann Umzug zu NMG

Damit das Konzept **real gezeigt werden kann, bevor mit der Firma gesprochen wird**,
bauen wir in zwei Stufen:

### Stufe 1 – Aufbau & Test auf eigener Domain
- Läuft unter einer eigenen Subdomain, z. B. `kasse.jagdeal.de`
  (Domain `jagdeal.de` liegt bei **Wix**).
- **Nur mit Test-/Dummy-Daten** – ausdrücklich KEINE echten NMG-Kundendaten
  (das wäre sonst schon ein DSGVO-Thema, solange Firma + AVV + EU-Hosting nicht stehen).
- Vorteil: Man kann jederzeit allein loslegen, ohne die Firma vorher einzubinden, und
  hat am Ende etwas **Vorzeigbares** fürs Gespräch mit der Geschäftsführung.

### Stufe 2 – Umzug auf die NMG-Domain (nach GO der Firma)
- Aus `kasse.jagdeal.de` wird `kasse.nmg-pharma.de`.
- **Ein Umzug ist fast nur ein Adress-Wechsel, kein Neubau**: Programm, Datenbank und
  Cloud-Server bleiben dieselben; es ändert sich nur, welche Adresse auf den Server
  zeigt (DNS) + ein Konfigurationswert in der App.
- Erst hier kommen die **echten Stammdaten** hoch.

### Technische Vorsorge (mache ich)
- Die Adresse wird **nirgends fest einbetoniert**, sondern bleibt an *einer* Stelle
  einstellbar → Umziehen = einen Wert ändern.
- https/Zertifikat ist pro Domain, aber für die neue Adresse genauso automatisch/gratis.

### Was du wofür brauchst
| | Stufe 1 (eigene Domain) | Stufe 2 (NMG-Domain) |
|---|---|---|
| Firma einbinden? | nein | **ja** (Zugang zu DNS von nmg-pharma.de) |
| Echte Kundendaten? | nein (nur Dummy) | ja |
| Domain | `jagdeal.de` (Wix, hast du) | `nmg-pharma.de` (Firma) |
| Cloud-Kosten | fallen ab Teststart an (~30–80 €/Monat) | laufen weiter |

> **Wix-Hinweis:** Subdomain bei Wix per DNS-Eintrag (A-Record/CNAME) im Dashboard
> möglich. Falls Wix das je nach Tarif einschränkt, zur Not nur die DNS-Verwaltung
> umziehen – die Wix-Seite läuft davon unberührt weiter.
