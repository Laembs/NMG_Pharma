"""NMGone Hilfe-/Handbuch-Modul ("Hilfe").

Eigenstaendiges Programm wie die Kasse und die Auswertungen (eigenes Fenster /
Taskleisten-Icon, AppUserModelID NMG.Hilfe). Zeigt ein bebildertes Handbuch:
links die Themen (dunkle Sidebar, einheitlich mit NMGone/Kasse/Report), rechts
der scrollbare Inhalt aus Text + Bildern.

Die Bilder werden zur Laufzeit aus  assets/hilfe/<thema>/<datei>.png  geladen.
Fehlt ein Bild noch, erscheint ein sauberer Platzhalter, der genau sagt, welcher
Screenshot dort hingehoert und wo die Datei abzulegen ist. So kann der Inhalt
Schritt fuer Schritt mit echten Screenshots befuellt werden, ohne Code zu aendern.

Das Modul ist reine Anzeige - es liest und schreibt KEINE Programmdaten.
"""
from __future__ import annotations

import os
import sys
import webbrowser
from pathlib import Path
import tkinter as tk
from tkinter import ttk, messagebox

from .config import ASSETS_DIR, DEMO_SUFFIX
from . import theme

# Wurzel fuer die Hilfe-Bilder:  assets/hilfe/<thema-key>/<datei>
HELP_IMG_DIR = ASSETS_DIR / "hilfe"
# Druckbare Handouts (HTML -> im Browser als PDF druckbar):  assets/handouts/<datei>.html
HANDOUT_DIR = ASSETS_DIR / "handouts"

# Thema -> zugehoeriges druckbares Handout. Jedes Thema mit Eintrag bekommt im
# Kopf den Button "Handout drucken (PDF)". "faq" hat bewusst keins.
TOPIC_HANDOUTS = {
    "start":          "index.html",
    "dashboard":      "dashboard.html",
    "bedarfsanalyse": "bedarfsanalyse.html",
    "produktanalyse": "produktanalyse.html",
    "kunden":         "kunden.html",
    "auskunft":       "pzn_auskunft.html",
    "markt":          "markt_insights.html",
    "kasse":          "kasse.html",
    "gdp":            "gdp.html",
    "meldungen":      "meldungen.html",
    "einkauf":        "einkauf.html",
    "faktura":        "faktura.html",
    "personal":       "personal.html",
    "auswertungen":   "auswertungen.html",
    "schulbank":      "schulbank.html",
    "daten":          "daten.html",
    "backup":         "backup.html",
}


# ── Inhalt ───────────────────────────────────────────────────────────────────
# Jedes Thema:  key -> {"title", "subtitle", "blocks"}
# Block-Typen (erstes Tupel-Element):
#   ("h",   text)                 Zwischenueberschrift
#   ("p",   text)                 Absatz
#   ("ul",  [item, item, ...])    Aufzaehlung
#   ("step",[item, item, ...])    nummerierte Schritt-fuer-Schritt-Liste
#   ("tip", text)                 gruener Hinweis-Kasten
#   ("warn",text)                 gelber Achtung-Kasten
#   ("img", datei, bildunterschrift)   Bild aus assets/hilfe/<key>/<datei>
TOPICS = [
    ("start",         "🏁", "Erste Schritte"),
    ("dashboard",     "🏠", "Dashboard"),
    ("bedarfsanalyse","📊", "Bedarfsanalyse"),
    ("produktanalyse","📈", "Produktanalyse"),
    ("kunden",        "👥", "Kunden"),
    ("auskunft",      "🔎", "PZN-Schnellauskunft"),
    ("markt",         "📦", "Markt-Insights"),
    ("kasse",         "🛒", "Kasse"),
    ("gdp",           "📦", "Wareneingang & Retouren"),
    ("meldungen",     "🔔", "Meldungen"),
    ("einkauf",       "🛒", "Einkauf"),
    ("faktura",       "🧾", "Faktura"),
    ("personal",      "👤", "Mitarbeiter"),
    ("auswertungen",  "📑", "Auswertungen"),
    ("schulbank",     "🎓", "Schulbank"),
    ("daten",         "🗄", "Daten aktualisieren"),
    ("backup",        "💾", "Backup & Update"),
    ("faq",           "❓", "FAQ & Tipps"),
]

HELP_CONTENT = {
    "start": {
        "title": "Erste Schritte",
        "subtitle": "Was NMGone ist und wie du dich zurechtfindest.",
        "blocks": [
            ("p", "NMGone ist die zentrale Plattform für die NMG-Pharma-Arbeit: "
                  "Bedarfs- und Produktanalysen, Kundenstamm, Kasse, Auswertungen "
                  "und die Pflege der Stammdaten – alles aus einem Fenster heraus."),
            ("h", "Die wichtigsten Bereiche"),
            ("ul", [
                "Dashboard – dein Startbildschirm mit Kacheln und Info-Widgets.",
                "Bedarfsanalyse – die PK-/ZW-Auswertung für Apotheken.",
                "Kunden – der Kundenstamm mit Historie.",
                "Kasse – Wareneingang, Lager und Verkauf (eigenes Fenster).",
                "Auswertungen – Verkäufe, Kunden und Artikel frei auswerten (eigenes Fenster).",
            ]),
            ("img", "01_startbildschirm.png", "Der Startbildschirm direkt nach dem Programmstart."),
            ("h", "Bedienung"),
            ("p", "Links die dunkle Navigationsleiste, in der Mitte der Arbeitsbereich. "
                  "Über die Navigation springst du jederzeit zwischen den Bereichen. "
                  "Kasse, Auswertungen und diese Hilfe öffnen sich als eigene Fenster "
                  "mit eigenem Symbol in der Taskleiste."),
            ("tip", "Alle Bereiche teilen sich dieselbe Datenbank. Was du in der Kasse "
                    "verkaufst, taucht sofort in den Auswertungen auf."),
        ],
    },
    "dashboard": {
        "title": "Dashboard",
        "subtitle": "Dein Startbildschirm – anpassbar an deine Arbeitsweise.",
        "blocks": [
            ("p", "Das Dashboard zeigt oben die Schnellzugriff-Kacheln und darunter "
                  "Info-Widgets (Daten-Aktualität, offene ToDos, letzte Auswertungen)."),
            ("img", "01_dashboard.png", "Dashboard mit Schnellzugriff-Kacheln und Info-Widgets."),
            ("h", "Kacheln anpassen"),
            ("step", [
                "Auf das Zahnrad bzw. „Dashboard anpassen“ klicken.",
                "Gewünschte Kacheln an- oder abwählen.",
                "Auswahl speichern – das Dashboard merkt sich deine Auswahl.",
            ]),
            ("tip", "Du kannst auch alle Kacheln abwählen, wenn du ein ganz schlankes "
                    "Dashboard willst – die Bereiche bleiben über die Navigation erreichbar."),
        ],
    },
    "bedarfsanalyse": {
        "title": "Bedarfsanalyse",
        "subtitle": "PK-/ZW-Auswertung für Apotheken (früher „Neue Auswertung“).",
        "blocks": [
            ("p", "Die Bedarfsanalyse erzeugt aus einer eingelesenen Apotheken-Datei "
                  "eine Auswertung der Bedarfe – getrennt nach Partnerkonditions-Kunden "
                  "(PK) und Zukunftswerk (ZW)."),
            ("h", "Ablauf"),
            ("step", [
                "Bedarfsanalyse öffnen.",
                "Datei der Apotheke auswählen (Excel/.xls – auch das Apotheken-XML-Format).",
                "Kundentyp / Datenquelle prüfen (PK oder ZW).",
                "Analyse starten und Ergebnis prüfen.",
                "Bei Bedarf als Analyse speichern – sie erscheint dann unter „Ges. Analysen“.",
            ]),
            ("img", "01_bedarfsanalyse.png", "Die Bedarfsanalyse mit Dateiauswahl und Ergebnisliste."),
            ("warn", "Apotheken-Dateien sind oft als „.xls“ benannt, intern aber XML/HTML. "
                     "NMGone erkennt das automatisch – wenn eine Datei nicht lädt, hilft der "
                     "Rohdaten-Format-Assistent weiter."),
        ],
    },
    "produktanalyse": {
        "title": "Produktanalyse",
        "subtitle": "Produktchancen erkennen und aufbereiten.",
        "blocks": [
            ("p", "Die Produktanalyse zeigt, bei welchen Produkten Chancen bestehen – "
                  "etwa über Mengen, Umsätze und Vergleichswerte."),
            ("img", "01_produktanalyse.png", "Die Produktanalyse mit Chancen-Übersicht."),
            ("h", "So gehst du vor"),
            ("step", [
                "Produktanalyse öffnen.",
                "Zeitraum bzw. Datenbasis wählen.",
                "Ergebnis sichten und nach Bedarf exportieren.",
            ]),
        ],
    },
    "kunden": {
        "title": "Kunden",
        "subtitle": "Der Kundenstamm mit Adressen, Konditionen und Historie.",
        "blocks": [
            ("p", "Unter „Kunden“ pflegst du den Kundenstamm: Apotheke, Inhaber, Adresse, "
                  "PK/ZW-Zuordnung, Rechnungsart und mehr. Die Kundendaten gehören allein "
                  "diesem Bereich – Auswertungen lesen sie nur, ändern sie aber nie."),
            ("img", "01_kunden.png", "Die Kunden-Maske mit Reitern für die Stammdaten."),
            ("h", "Einen Kunden bearbeiten"),
            ("step", [
                "Kunde in der Liste auswählen.",
                "Reiter mit den Stammdaten öffnen.",
                "Felder anpassen und speichern.",
            ]),
            ("tip", "Über die Kundennummer hängen Bestellungen, Auswertungen und Kasse "
                    "zusammen – sie sollte eindeutig und gepflegt sein."),
        ],
    },
    "auskunft": {
        "title": "PZN-Schnellauskunft",
        "subtitle": "Alles zu einem Artikel auf einen Blick – ohne ganze Auswertung.",
        "blocks": [
            ("p", "Die PZN-Schnellauskunft beantwortet in Sekunden „Was wissen wir über diesen "
                  "Artikel?“. Sie sammelt zu einer PZN alle verfügbaren Informationen aus allen "
                  "Tabellen der Datenbank. Reine Anzeige – es werden keine Daten verändert."),
            ("h", "Artikel suchen"),
            ("step", [
                "Suchbegriff eingeben – eine PZN (auch unvollständig) oder Teil des Artikelnamens.",
                "In der Trefferliste den gewünschten Artikel anklicken.",
            ]),
            ("img", "01_auskunft.png", "Die PZN-Schnellauskunft mit Suchfeld und 360°-Artikelinfo."),
            ("h", "Was die Auskunft zeigt"),
            ("ul", [
                "NMG-Stamm: Artikelname, Hersteller, APU, Taxe-EK/-VK, Menge, Einheit, Wirkstoffe.",
                "Artikelstamm: Darreichungsform, Packung, Hersteller, Einkaufspreis (EK).",
                "Rabatt: hinterlegter NMG-Rabattsatz und Quelle.",
                "Lieferfähigkeit, bester effektiver EK, Biosimilar-Gruppe und Austausch zu NMG.",
            ]),
            ("tip", "Ruft eine Apotheke an und fragt nach Preis, Rabatt oder einer NMG-Alternative, "
                    "hast du die Antwort sofort – ohne in mehreren Listen zu blättern."),
        ],
    },
    "markt": {
        "title": "Markt-Insights",
        "subtitle": "Was fragen unsere Apotheken-Kunden tatsächlich nach?",
        "blocks": [
            ("p", "Die Markt-Insights aggregieren über alle bisher erstellten Auswertungen hinweg, "
                  "was die Apotheken-Kunden wirklich nachfragen – das Gesamtbild statt einer einzelnen "
                  "Apotheke."),
            ("h", "Zwei Sichtweisen"),
            ("ul", [
                "NMG-Sicht – sortiert nach Rabatt-Potenzial: wo steckt für uns der größte Hebel "
                "(nur NMG-Treffer).",
                "Bedarf-Sicht – sortiert nach Gesamt-Absatz: was wird am meisten nachgefragt, egal ob "
                "NMG oder nicht.",
            ]),
            ("img", "01_markt.png", "Markt-Insights mit umschaltbarer NMG- und Bedarf-Sicht."),
            ("tip", "Die Bedarf-Sicht ist ideal für Einkauf und Sortimentsplanung, die NMG-Sicht für "
                    "den Vertrieb. Je mehr Bedarfsanalysen erstellt wurden, desto aussagekräftiger."),
        ],
    },
    "kasse": {
        "title": "Kasse",
        "subtitle": "Lagerbestand führen und an Apotheken verkaufen.",
        "blocks": [
            ("p", "Die Kasse ist ein eigenes Programm mit eigenem Taskleisten-Symbol. "
                  "Sie führt den Lagerbestand und verkauft an Apotheken. Die "
                  "Warenannahme (Wareneingang mit Charge und Verfall) läuft über die "
                  "eigene App „Wareneingang & Retouren“ – beide teilen sich dieselbe "
                  "Datenbank, gebuchte Ware steht also sofort in der Kasse bereit."),
            ("h", "Übersicht – die Startseite"),
            ("p", "Beim Öffnen zeigt die Kasse jetzt zuerst den Reiter „Übersicht“: "
                  "Umsatz und Verkäufe von heute, offene Aufträge und Vorbestellungen, "
                  "Verkaufs- und Lagerwert sowie eine Warnliste mit bald ablaufenden und "
                  "abgelaufenen Chargen. Rechts stehen die letzten Verkäufe – ein "
                  "Doppelklick öffnet den jeweiligen Verkauf. Über „🔄 Aktualisieren“ "
                  "werden alle Kennzahlen neu berechnet."),
            ("tip", "Der Einkaufspreis (EK) wird bei der Warenannahme in der App "
                    "„Wareneingang & Retouren“ erfasst (dort mit dem APU vorbelegt und "
                    "überschreibbar). Die Kasse zeigt daraus überall den Lagerwert "
                    "(EK × Bestand) – zusätzlich zum Verkaufswert (APU × Bestand)."),
            ("img", "01_kasse.png", "Die Kasse mit Verkaufsmaske."),
            ("h", "Typischer Ablauf"),
            ("step", [
                "Ware annehmen – in der App „Wareneingang & Retouren“; der Lagerbestand erhöht sich.",
                "Verkauf an eine Apotheke anlegen (Kunde wählen).",
                "Artikel/Mengen erfassen, abschließen.",
                "Der Verkauf landet sofort in den Auswertungen.",
            ]),
            ("h", "Artikel-Übersicht & Lagerbestand"),
            ("p", "Im Reiter „Artikel“ siehst du den Artikelstamm und den aktuellen "
                  "Lagerbestand. Eine Leiste unten zeigt Gesamtbestand, Verkaufswert "
                  "(APU × Bestand) und Lagerwert (EK × Bestand); sobald du nach PZN oder "
                  "Artikel suchst, beziehen sich die Summen nur noch auf die angezeigte "
                  "Auswahl. Ein Doppelklick auf einen Artikel öffnet seine Chargen und Verfälle."),
            ("h", "Bestand korrigieren"),
            ("p", "Im Chargen-Fenster (Doppelklick auf einen Artikel) korrigierst du den "
                  "Bestand einer einzelnen Charge per Doppelklick – z. B. nach Inventur "
                  "oder Bruch. Ein Grund ist Pflicht und wird revisionssicher im Protokoll "
                  "festgehalten. Das Einbuchen neu eingetroffener Ware geschieht dagegen in "
                  "der App „Wareneingang & Retouren“."),
            ("img", "02_artikel.png", "Reiter „Artikel“: Artikelstamm, Lagerbestand und Bestandskorrektur per Doppelklick auf eine Charge."),
            ("h", "Freie Position"),
            ("p", "Über „➕ Freie Position“ (unter der Positionsliste) erfasst du einen "
                  "frei benannten Posten mit eigenem Preis, Menge und Rabatt – z. B. einen "
                  "Botendienst-Zuschlag. Er wird NICHT vom Lager abgebucht, aber gespeichert "
                  "und erscheint mit Preis auf der Auftragsbestätigung."),
            ("h", "Wenn der Bestand nicht reicht"),
            ("p", "Buchst du mehr als auf Lager ist, fragt die Kasse nach: Du kannst die "
                  "verfügbare Menge sofort als Bestellung liefern und den Rest automatisch "
                  "als Vorbestellung aufnehmen lassen – oder nur den vorhandenen Bestand "
                  "abverkaufen."),
            ("img", "03_vorbestellungen.png", "Reiter „Vorbestellungen“: offene Vorbestellungen disponieren und als Verkauf übernehmen."),
            ("h", "MSK & Lieferschein"),
            ("p", "Im Reiter „Verkäufe“ markierst du einen Verkauf als „In MSK erfasst“. "
                  "Direkt danach bietet die Kasse an, den Lieferschein zu erzeugen und zu "
                  "öffnen. Über die Detailansicht eines Verkaufs lässt er sich jederzeit "
                  "erneut drucken."),
            ("tip", "Lieferschein-Vorlage anpassen: vorlagen/lieferschein.html (eigene "
                    "Firmendaten/Logo). Der Lieferschein zeigt Charge/Verfall, aber keine Preise."),
            ("p", "Erzeugst du einen Lieferschein über den Knopf in der Detailansicht, fragt "
                  "die Kasse zuerst, ob der Auftrag in MSK erfasst wurde – auf „Ja“ wird er "
                  "gleich entsprechend markiert."),
            ("h", "Wer hat was mit einem Auftrag gemacht?"),
            ("p", "In der Detailansicht eines Verkaufs öffnet „🕘 Verlauf“ die komplette "
                  "Historie (angelegt, MSK erfasst, Lieferschein, Storno, …) mit Mitarbeiter "
                  "und Zeitpunkt – auch als PDF. Das vollständige Protokoll im Reiter "
                  "„Protokoll“ lässt sich nach Auftrag-Nr oder Mitarbeiter filtern und als "
                  "PDF exportieren."),
            ("h", "Reiter „Auswertung“"),
            ("ul", [
                "Umsatz – je Tag/Monat/Jahr mit Anzahl Verkäufe, Anzahl Packungen, "
                "APU Brutto, Rabatt (Netto) und APU Netto; Zeitraum frei einschränkbar.",
                "Tagesabschluss – Tag per Kalender wählen; zeigt die Tagesabschluss-Nr "
                "und die Kennzahlen (Menge der Verkäufe, verkaufte Packungen, APU-, "
                "Rabatt- und Umsatz-Summe).",
                "Verfall – Zeitraum wählbar (3/6/9/12 Monate oder Alle, Standard 3 Monate); "
                "abgelaufene (rot) und bald ablaufende (gelb) Chargen, mit Summe aus "
                "Bestand und Verkaufswert (APU × Bestand).",
                "Inventur – Zählliste mit Soll-Bestand und Verkaufswert zum Ausdrucken.",
            ]),
            ("p", "Den Tagesabschluss erzeugst du im Reiter „Tagesabschluss“ (oder per "
                  "Doppelklick auf einen Tag in der Umsatztabelle). Beim Erzeugen des PDFs "
                  "wird eine laufende Nummer vergeben. Er entsteht jeden Abend um 18 Uhr "
                  "zusätzlich automatisch (mit Nummer); verpasste Vortage werden beim "
                  "nächsten Start nachgeholt."),
            ("img", "04_auswertung.png", "Reiter „Auswertung“: Umsatz, Tagesabschluss, Verfall und Inventur."),
            ("h", "Defektmeldung (Nichtverfügbarkeit)"),
            ("p", "Ist ein Artikel nicht vorrätig oder nicht lieferbar, erzeugst du im "
                  "Reiter „Defektmeldung“ eine Bescheinigung für die Apotheke: Apotheke "
                  "suchen, Artikel hinzufügen (der aktuelle Bestand wird angezeigt), Grund "
                  "wählen und „Defektmeldung erzeugen“."),
            ("warn", "Der genaue Rechtstext der Defektmeldung muss fachlich/rechtlich geprüft "
                     "und eingetragen werden – das geht im Reiter „Einstellungen“."),
            ("img", "05_defektmeldung.png", "Reiter „Defektmeldung“: Nichtverfügbarkeit für die Apotheke bescheinigen."),
            ("h", "Einstellungen"),
            ("p", "Im Reiter „Einstellungen“ (Zahnrad) pflegst du die Firmendaten (Kopf von "
                  "Auftragsbestätigung, Lieferschein und Defektmeldung), den Rechtstext der "
                  "Defektmeldung und die Uhrzeit des automatischen Tagesabschlusses – alles "
                  "ohne Dateien bearbeiten zu müssen."),
            ("img", "06_einstellungen.png", "Reiter „Einstellungen“: Firmendaten, Dokument-Texte und Parameter."),
            ("warn", "Stornierte oder abgesagte Verkäufe werden in den Umsatz-Auswertungen "
                     "nicht mitgezählt."),
        ],
    },
    "einkauf": {
        "title": "Einkauf",
        "subtitle": "Beschaffung EU-Ausland, §129-Margen, Aufgaben & Meldungen.",
        "blocks": [
            ("p", "Die Einkauf-App ist ein eigenes Programm mit eigenem Taskleisten-Symbol – "
                  "das „Sorglos-Paket“ für den Einkäufer. Sie bündelt Aufgaben/Wiedervorlagen, "
                  "die Lieferanten und Beschaffungsquellen aus dem EU-Ausland sowie die "
                  "Margenkalkulation nach §129 SGB V. Alles teilt sich die NMGone-Datenbank."),
            ("img", "01_dashboard.png", "Das Einkauf-Dashboard mit Kennzahlen und Meldungen."),
            ("tip", "Zum Ausprobieren: unter „Einstellungen“ → „Beispieldaten anlegen“ klicken (oder "
                    "den Button auf dem leeren Dashboard). Das legt Beispiel-Lieferanten, -Quellen und "
                    "-Aufgaben an; vorhandene Daten bleiben erhalten."),
            ("h", "Aufgaben & Meldungen"),
            ("p", "Jede Wiedervorlage, Anfrage, Reklamation oder ein Rückruf wird als Aufgabe mit "
                  "Fälligkeit und Priorität erfasst – optional einem Lieferanten zugeordnet. "
                  "Fällige und überfällige Aufgaben erscheinen automatisch als Meldungen oben auf "
                  "dem Dashboard („melde dich bei …“) und lassen sich dort direkt abhaken."),
            ("step", [
                "Links „Aufgaben & Meldungen“ öffnen, „Neue Aufgabe“ wählen.",
                "Titel, Kategorie, Priorität und Fälligkeit eintragen (Schnell-Buttons: Heute, +1 Woche …).",
                "Optional einen Lieferanten zuordnen.",
                "Erledigte Aufgaben verschwinden aus den Meldungen.",
            ]),
            ("h", "Lieferanten & Beschaffungsquellen"),
            ("p", "Unter „Lieferanten“ pflegst du Bezugsquellen je Land mit Währung, Lieferzeit, "
                  "Mindestbestellwert, Zahlungsziel und GDP-Zertifizierung. Unter "
                  "„Beschaffungsquellen“ hinterlegst du pro Artikel (PZN) den Einkaufspreis in "
                  "Fremdwährung und die Mindestabnahme – die Basis für Marge und Bestellvorschlag."),
            ("img", "02_quellen.png", "Beschaffungsquellen mit EK in Fremdwährung, EUR-Umrechnung und Marge."),
            ("h", "Margenrechner §129"),
            ("p", "Der Margenrechner führt Import-EK (Fremdwährung) → deutscher AVP/Festbetrag → "
                  "garantierter Preisabstand zusammen. Du gibst eine PZN ein (oder lädst sie aus den "
                  "NMG-Stammdaten), der Rechner prüft den nach §129 erforderlichen Preisabstand, zeigt "
                  "den maximal zulässigen Import-AVP, die Ersparnis für die Kasse und deine eigene Marge."),
            ("img", "03_marge.png", "Margenrechner §129 mit Status-Banner, Kennzahlen und Marge."),
            ("h", "Beschaffungsvorschlag"),
            ("p", "Der „Beschaffungsvorschlag“ wertet alle aktiven Quellen aus und zeigt je Artikel "
                  "automatisch die beste Quelle (höchste Marge) – mit §129-Status, EK in EUR, Marge "
                  "und dem Absatz aus den Kasse-Verkäufen. Über die Filter blendest du nur "
                  "§129-konforme Artikel oder nur Artikel mit Absatz ein. Per „Nachbestellen“ legst "
                  "du direkt eine Aufgabe beim passenden Lieferanten an."),
            ("h", "Export"),
            ("p", "Lieferanten, Beschaffungsquellen und der Beschaffungsvorschlag lassen sich jeweils "
                  "per „⤓ Excel“ als formatierte Tabelle exportieren (Ablage im Ausgaben-Ordner)."),
            ("tip", "Die §129-Staffel (z. B. AVP bis 100 € → 15 %) ist in den Einstellungen anpassbar. "
                    "Die Wechselkurse pflegst du unter „Wechselkurse“ als Tageskurs (EUR je 1 Einheit). "
                    "Alle Tabellen lassen sich per Klick auf die Spaltenüberschrift sortieren."),
        ],
    },
    "gdp": {
        "title": "Wareneingang & Retouren (GDP)",
        "subtitle": "Chargen-Rückverfolgung, Kundenqualifizierung & Retouren.",
        "blocks": [
            ("p", "Diese App ist das Großhandelsmodul. Hier nimmst du Ware an "
                  "(Wareneingang mit Charge und Verfall) und wickelst sie GDP-konform "
                  "ab: Chargen-Rückverfolgung, Kundenqualifizierung und Retouren. Eigenes "
                  "Fenster mit eigenem Taskleisten-Symbol, gemeinsame Datenbank mit NMGone – "
                  "eingebuchte Ware steht sofort im Lager der Kasse bereit."),
            ("tip", "Kühlsachenkontrolle, Selbstinspektion und das Meldewesen liegen "
                    "jetzt in der eigenen App „Meldungen“ – so bleibt der Wareneingang "
                    "übersichtlich."),
            ("img", "01_uebersicht.png", "Die Übersicht mit Kennzahlen und offenen GDP-Pflichten."),
            ("h", "Übersicht (Ampel)"),
            ("p", "Die Startseite zeigt offene Retouren, ungeprüfte Wareneingänge, "
                  "abgelaufene Lizenzen und Chargen mit baldigem Verfall. Jede Meldung "
                  "springt direkt in den passenden Bereich."),
            ("h", "Wareneingang annehmen & GDP-Prüfung"),
            ("p", "Oben buchst du eingetroffene NMG-Ware ins Lager: Artikel suchen, "
                  "Charge, Verfall, Menge und EK erfassen, Lieferant/Lieferschein angeben "
                  "und „Einbuchen“. Die Ware steht damit sofort im Lager der Kasse. "
                  "Alternativ importierst du eine Wareneingangs-Liste (Excel/CSV/TXT)."),
            ("step", [
                "Ware einbuchen (oder Liste importieren) – sie erscheint in der Liste darunter.",
                "Eingang wählen und „GDP-Prüfung erfassen“: Transport-Temperatur, Unversehrtheit und Dokumente prüfen.",
                "Ergebnis wird als GDP-konform / nicht konform festgeschrieben.",
            ]),
            ("h", "Chargen-Rückverfolgung & Rückruf"),
            ("p", "Charge oder PZN eingeben: links siehst du den Eingang, rechts jede "
                  "Apotheke, die diese Charge erhalten hat. Im Rückruffall „Rückruf "
                  "auslösen“ und über „Verteiler exportieren“ die Liste der betroffenen "
                  "Apotheken (mit E-Mail) als CSV sichern."),
            ("img", "02_rueckverfolgung.png", "Chargen-Rückverfolgung Kunde ↔ Charge mit Rückruf-Verteiler."),
            ("h", "Kundenqualifizierung"),
            ("p", "Nur lizenzierte Apotheken dürfen beliefert werden. Hier hinterlegst du "
                  "Lizenznummer, -typ und Gültigkeit. Abgelaufene oder nicht freigegebene "
                  "Apotheken werden rot markiert."),
            ("h", "Retouren & Reklamationen"),
            ("p", "Meldet eine Apotheke telefonisch oder per E-Mail eine Rücksendung, "
                  "erfasst ein Mitarbeiter sie hier. Danach ist sie für alle sichtbar "
                  "(das Dashboard zeigt die Zahl offener Retouren)."),
            ("step", [
                "„Neue Retoure / Reklamation“: Apotheke, Artikel/Charge, Menge und Grund erfassen.",
                "„Prüfen“ nimmt den Vorgang in Bearbeitung.",
                "„Gutschrift“ erzeugt einen Gutschrift-Beleg (Anbindung an die Faktura).",
                "„In Retourenbestand“ nimmt die Ware in Quarantäne, „Vernichten“ dokumentiert die Entsorgung.",
            ]),
            ("img", "03_retouren.png", "Retouren-Workflow von der Erfassung bis zur Gutschrift."),
            ("h", "Retourenbestand: freigeben oder abschreiben"),
            ("p", "Zurückgenommene Ware geht NICHT direkt in den Verkauf, sondern in "
                  "einen gesperrten Retourenbestand (Quarantäne). Im Bereich "
                  "„Retourenbestand“ steht der Bestand als „Normalbestand (Retourenbestand)“, "
                  "z. B. 25 (5) – also 25 verkaufbar, 5 in Quarantäne."),
            ("step", [
                "Zeile im Retourenbestand wählen.",
                "„In Bestand freigeben“ – einwandfreie Ware wird wieder verkaufbar (geht in den Normalbestand).",
                "oder „Abschreiben“ – mit Grund; die Menge wird ausgebucht und der EK-Wert dokumentiert.",
            ]),
            ("img", "04_retourenbestand.png", "Retourenbestand mit Freigabe und Abschreibung."),
            ("tip", "Jeder Schritt landet im revisionssicheren Protokoll – inklusive "
                    "Wer/Was/Wann. Protokoll und Abschreibungen lassen sich als CSV exportieren."),
        ],
    },
    "meldungen": {
        "title": "Meldungen",
        "subtitle": "GDP-Meldewesen, Kühlsachenkontrolle & Selbstinspektion.",
        "blocks": [
            ("p", "Die Meldungen-App bündelt die Qualitäts- und Überwachungsaufgaben an "
                  "einem Ort – früher steckten sie verstreut im Wareneingang. Eigenes "
                  "Fenster mit eigenem Taskleisten-Symbol, gemeinsame Datenbank mit NMGone."),
            ("img", "01_uebersicht.png", "Die Übersicht mit offenen Meldungen und Überwachungspflichten."),
            ("h", "Übersicht (Ampel)"),
            ("p", "Die Startseite zeigt offene und kritische Meldungen, überfällige "
                  "Vorgänge, offene Temperatur-Abweichungen und fällige Selbstinspektionen. "
                  "Jede Meldung springt direkt in den passenden Bereich."),
            ("h", "Meldungen erfassen & bearbeiten"),
            ("step", [
                "„Neue Meldung“: Art (Abweichung, Qualitätsmangel, Reklamation …), Titel, "
                "Priorität, Betrifft, Verantwortlichen und Fälligkeit erfassen.",
                "Über „In Bearbeitung“ und „Erledigt“ den Status führen.",
                "„Maßnahme / CAPA“ dokumentiert die Korrektur- und Vorbeugemaßnahme.",
                "Doppelklick öffnet die Detailansicht mit dem kompletten Verlauf.",
            ]),
            ("img", "02_meldung_neu.png", "Eine neue Meldung mit Typ, Priorität und Verantwortlichem."),
            ("tip", "Über den Status-Filter siehst du wahlweise nur offene oder alle "
                    "Meldungen; „Export CSV“ sichert die Liste."),
            ("h", "Kühlsachenkontrolle / Temperatur"),
            ("p", "Messpunkte mit Soll-Grenzen (z. B. Kühlschrank 2–8 °C) verwalten und "
                  "Messungen erfassen. Liegt ein Wert außerhalb der Grenze, wird er als "
                  "Abweichung markiert – dazu eine Maßnahme (CAPA) dokumentieren und auf "
                  "„behoben“ setzen."),
            ("img", "03_kuehlkette.png", "Messpunkte mit Soll-Grenzen und erfasste Temperaturmessungen."),
            ("h", "Selbstinspektion"),
            ("step", [
                "„Neue Inspektion (Vorlage)“ erzeugt eine GDP-Checkliste.",
                "Jeden Prüfpunkt per Doppelklick auf ok / Abweichung / n.z. setzen.",
                "Zu Abweichungen eine Maßnahme erfassen, am Ende abschließen.",
            ]),
            ("img", "04_inspektion.png", "Selbstinspektion mit Checkliste und Maßnahmen."),
            ("tip", "Alle Vorgänge (Meldungen, Kühlkette, Inspektion) landen im "
                    "revisionssicheren Protokoll – es ist dasselbe Protokoll wie im "
                    "Wareneingang und lässt sich als CSV exportieren."),
        ],
    },
    "faktura": {
        "title": "Faktura",
        "subtitle": "Rechnungen, Gutschriften & Quartalsvergütung.",
        "blocks": [
            ("p", "Die Faktura ist ein eigenes Programm mit eigenem Taskleisten-Symbol. "
                  "Sie erstellt Rechnungen, Gutschriften und Stornos sowie die "
                  "Quartalsvergütung – jeweils als fertiges PDF mit deinem Layout."),
            ("img", "01_start.png", "Die Faktura-Startseite mit der Übersicht."),
            ("h", "Eine neue Rechnung erstellen"),
            ("step", [
                "Links „Neue Rechnung“ öffnen (oder über „Aufträge“ einen Auftrag wählen).",
                "Kunde wählen und Positionen erfassen (PZN, Menge, APU).",
                "Belegdaten prüfen (Belegnummer, Datum, Leistungsdatum).",
                "Als PDF erzeugen – der Beleg wird abgelegt und kann geöffnet werden.",
            ]),
            ("img", "02_rechnung_neu.png", "Die Maske „Neue Rechnung“ mit Positionen."),
            ("h", "Belegnummern & Layout"),
            ("p", "Unter „Einstellungen“ pflegst du Firmendaten, die Belegnummern und das "
                  "Layout. Die Nummernkreise sind frei konfigurierbar über Platzhalter: "
                  "{JJJJ} = Jahr, {MM} = Monat, {NR} = Zähler (z. B. {NR:5} = 5-stellig). "
                  "Im Layout legst du Akzentfarbe und Logo der Belege fest."),
            ("img", "03_einstellungen.png", "Einstellungen: Firmendaten und frei konfigurierbare Belegnummern."),
            ("h", "Quartalsvergütung"),
            ("p", "Im Bereich „Quartalsvergütung“ erzeugst du die Vergütungsbelege eines "
                  "Quartals; die zugehörige „Staffel“ legst du unter Konfiguration fest."),
            ("tip", "Die PDFs entstehen über Edge/Chrome im Hintergrund. Ist kein Browser "
                    "vorhanden, bleibt eine HTML-Fassung des Belegs erhalten."),
        ],
    },
    "personal": {
        "title": "Mitarbeiter & Personal",
        "subtitle": "Organigramm, Abwesenheiten & Arbeitsbereiche.",
        "blocks": [
            ("p", "Das Mitarbeiter-Board ist ein eigenes Programm mit drei Ansichten, die "
                  "du oben umschaltest: Organigramm, Arbeitsbereiche und Abwesenheiten. "
                  "Alle teilen sich die NMGone-Datenbank."),
            ("img", "01_organigramm.png", "Organigramm: Mitarbeiterkarten mit Vorgesetzten-Verbindungen."),
            ("h", "Organigramm"),
            ("p", "Die Mitarbeiterkarten lassen sich frei anordnen. Über den Verbinden-Modus "
                  "ziehst du Linien vom Vorgesetzten zum Mitarbeiter; Farben kennzeichnen die "
                  "Abteilungen."),
            ("h", "Arbeitsbereiche"),
            ("p", "Hier ordnest du Mitarbeitern ihre Arbeitsbereiche und Kategorien zu – "
                  "umschaltbar nach Mitarbeiter oder nach Bereich."),
            ("img", "02_arbeitsbereiche.png", "Arbeitsbereiche: Zuordnung von Mitarbeitern und Bereichen."),
            ("h", "Abwesenheiten"),
            ("p", "Urlaub und andere Abwesenheiten pflegst du im Kalender (Monats- oder "
                  "Jahresansicht). Der Urlaubsverfall wird mitgeführt."),
            ("img", "03_abwesenheiten.png", "Abwesenheiten: Urlaubskalender mit Monats- und Jahresansicht."),
            ("h", "Anträge genehmigen oder ablehnen"),
            ("p", "Jede neu eingetragene Abwesenheit ist zunächst ein Antrag (Status "
                  "„Beantragt“) und wird im Kalender schraffiert mit ⏳ dargestellt. In der "
                  "Ansicht „Anträge“ entscheiden die als personalverantwortlich markierten "
                  "Mitarbeiter über jeden Antrag: Genehmigen oder Ablehnen (mit Grund). Erst "
                  "ein genehmigter Antrag zählt auf das Urlaubskonto."),
            ("step", [
                "Mitarbeiter trägt seine Abwesenheit ein → Status „Beantragt“.",
                "In „Anträge“ den Eintrag wählen und „Genehmigen“ oder „Ablehnen“ klicken.",
                "Die Entscheidung (wer/wann, ggf. Ablehnungsgrund) wird mitgespeichert.",
            ]),
            ("tip", "Zweite Stufe: In der Parameter-App (Einstellungen) lässt sich aktivieren, "
                    "dass die Geschäftsführung jeden vom Personalverantwortlichen genehmigten "
                    "Antrag zusätzlich freigeben muss."),
        ],
    },
    "auswertungen": {
        "title": "Auswertungen",
        "subtitle": "Verkäufe, Kunden und Artikel frei auswerten und exportieren.",
        "blocks": [
            ("p", "Die Auswertungen sind ein eigenes Fenster mit vier Perspektiven: "
                  "Verkäufe, Kunden, Artikel und „Frei“ (selbst gebauter Bericht). "
                  "Alle Umsätze sind netto."),
            ("img", "01_auswertungen.png", "Auswertungen mit Perspektiven-Navigation und Ergebnistabelle."),
            ("h", "Eine Auswertung erstellen"),
            ("step", [
                "Perspektive links wählen (Verkäufe / Kunden / Artikel / Frei).",
                "Zeitraum einstellen (akt. Monat, 3/6/12 Monate oder Von–Bis).",
                "Optional über „Spalten…“ die Ausgabespalten wählen und sortieren.",
                "„Auswerten“ klicken.",
            ]),
            ("h", "Freie Auswertung (Baukasten)"),
            ("p", "In der Perspektive „Frei“ baust du dir über „Bauen…“ eine eigene "
                  "Auswertung: Gruppieren nach Dimensionen (z.B. Ort, Monat), Kennzahlen "
                  "(Umsatz, Menge …) und Filter. Vorlagen lassen sich speichern und laden."),
            ("img", "02_auswertung_frei.png", "Der Baukasten-Dialog für die freie Auswertung."),
            ("tip", "Export als Excel ist immer möglich. Word und PDF erscheinen, wenn die "
                    "zugehörigen Bibliotheken installiert sind."),
        ],
    },
    "schulbank": {
        "title": "Schulbank",
        "subtitle": "Lernvorschläge prüfen und bestätigen.",
        "blocks": [
            ("p", "Die Schulbank sammelt Lernvorschläge – z.B. Zuordnungen, die das "
                  "Programm vorschlägt und die du bestätigen oder korrigieren kannst. "
                  "Bestätigte Einträge verbessern künftige Analysen."),
            ("img", "01_schulbank.png", "Die Schulbank mit offenen Lernvorschlägen."),
        ],
    },
    "daten": {
        "title": "Daten aktualisieren",
        "subtitle": "Stammdaten und Importe pflegen.",
        "blocks": [
            ("p", "Hier hältst du die Datenbasis aktuell: APU/HAP, NMG- und PK-Rabatte, "
                  "Artikelstamm sowie manuelle Analyse-Importe."),
            ("img", "01_daten.png", "Die Seite „Daten aktualisieren“ mit den Import-Kacheln."),
            ("tip", "Das Info-Widget „Daten-Aktualität“ auf dem Dashboard zeigt per Ampel, "
                    "welche Datenquelle frisch ist und welche eine Aktualisierung braucht."),
        ],
    },
    "backup": {
        "title": "Backup & Update",
        "subtitle": "Daten sichern und das Programm aktuell halten.",
        "blocks": [
            ("p", "NMGone kann die Datenbank sichern und wiederherstellen sowie Updates "
                  "einspielen. Sichere regelmäßig – besonders vor größeren Importen."),
            ("img", "01_backup.png", "Die Backup- und Update-Funktionen."),
            ("step", [
                "Backup erstellen – schreibt eine Sicherung der Datenbank.",
                "Backup wiederherstellen – spielt eine Sicherung zurück.",
                "Update installieren – bringt das Programm auf den neuen Stand.",
            ]),
            ("warn", "Beim Wiederherstellen werden die aktuellen Daten durch das Backup "
                     "ersetzt. Im Zweifel vorher noch ein frisches Backup ziehen."),
        ],
    },
    "faq": {
        "title": "FAQ & Tipps",
        "subtitle": "Häufige Fragen und kleine Kniffe.",
        "blocks": [
            ("h", "Eine Datei lädt nicht – was tun?"),
            ("p", "Apotheken-„.xls“ sind oft XML/HTML. NMGone erkennt das meist automatisch. "
                  "Klappt es nicht, öffnet sich der Rohdaten-Format-Assistent und führt dich "
                  "durch die Zuordnung der Spalten."),
            ("h", "Wo liegen meine Daten?"),
            ("p", "Die installierte Version speichert unter C:\\ProgramData\\NMGone. "
                  "Datenbank, Ausgaben, gespeicherte Analysen und Backups liegen dort."),
            ("h", "Kasse/Auswertungen lassen sich nicht doppelt öffnen"),
            ("p", "Das ist gewollt: Läuft das Fenster bereits, holt NMGone es nicht erneut, "
                  "sondern weist darauf hin. Hol das Fenster über die Taskleiste nach vorn."),
            ("tip", "Diese Hilfe wächst mit: Lege eigene Screenshots als PNG in den passenden "
                    "Ordner unter assets/hilfe/ ab (der Platzhalter nennt dir den genauen Pfad)."),
        ],
    },
}


class HilfePanel(tk.Frame):
    """Hilfe-Oberflaeche. Laeuft als NMGone-Toplevel und als eigene .exe."""

    def __init__(self, master, on_close=None, nmgone_action=None):
        super().__init__(master, bg=theme.BG)
        self._on_close = on_close
        self._nmgone_action = nmgone_action
        self._topic = TOPICS[0][0]
        self._nav_buttons: dict[str, tk.Button] = {}
        self._img_refs: list = []   # PhotoImage-Referenzen festhalten (sonst GC)
        self._build()
        self._select_topic(self._topic)

    # ── Aufbau ───────────────────────────────────────────────────────────────
    def _build(self):
        self.columnconfigure(1, weight=1)
        self.rowconfigure(0, weight=1)

        # Linke Sidebar (dunkel, einheitlich mit NMGone/Kasse/Report)
        nav = tk.Frame(self, bg=theme.SIDEBAR, width=244)
        nav.grid(row=0, column=0, sticky="ns")
        nav.grid_propagate(False)

        self._app_icon = theme.load_icon(ASSETS_DIR / "Hilfe.ico", 56)
        if self._app_icon:
            tk.Label(nav, image=self._app_icon, bg=theme.SIDEBAR).pack(anchor="w", padx=20, pady=(18, 2))
        tk.Label(nav, text="Hilfe", bg=theme.SIDEBAR, fg="white",
                 font=(theme.FONT, 16, "bold")).pack(anchor="w", padx=20, pady=(2, 2))
        tk.Label(nav, text="Handbuch · Schritt für Schritt", bg=theme.SIDEBAR,
                 fg=theme.SIDEBAR_MUTED, font=(theme.FONT, 9)).pack(anchor="w", padx=20, pady=(0, 14))

        for key, icon, label in TOPICS:
            b = tk.Button(nav, text=f"   {icon}   {label}", anchor="w", relief="flat",
                          bg=theme.SIDEBAR, fg=theme.SIDEBAR_TEXT, activebackground=theme.SIDEBAR_ACTIVE,
                          activeforeground="white", bd=0, font=(theme.FONT, 11),
                          cursor="hand2", command=lambda k=key: self._select_topic(k))
            b.pack(fill="x", padx=10, pady=1, ipady=8)
            self._nav_buttons[key] = b

        tk.Frame(nav, bg=theme.SIDEBAR).pack(fill="both", expand=True)
        nmb = tk.Button(nav, text="↩  NMGone öffnen", relief="flat", bg=theme.SIDEBAR_ACTIVE,
                        fg="white", activebackground="#1B5085", activeforeground="white", bd=0,
                        font=(theme.FONT, 10), cursor="hand2", command=self._open_nmgone)
        nmb.pack(fill="x", padx=10, pady=(2, 18), ipady=8)

        # Rechter Inhalt (scrollbar)
        content = tk.Frame(self, bg=theme.BG)
        content.grid(row=0, column=1, sticky="nsew")
        content.columnconfigure(0, weight=1)
        content.rowconfigure(1, weight=1)

        self.header = tk.Frame(content, bg=theme.BG)
        self.header.grid(row=0, column=0, sticky="ew", padx=28, pady=(22, 8))
        # Rechts: Handout als druckbares PDF oeffnen (nur Themen mit Handout)
        self._handout_btn = tk.Button(
            self.header, text="📄  Handout drucken (PDF)", relief="flat",
            bg=theme.PRIMARY, fg="white", activebackground=theme.SIDEBAR_ACTIVE,
            activeforeground="white", bd=0, font=(theme.FONT, 10, "bold"),
            cursor="hand2", padx=14, pady=8, command=self._open_handout)
        # Links: Titel + Untertitel
        head_text = tk.Frame(self.header, bg=theme.BG)
        head_text.pack(side="left", fill="x", expand=True)
        self._title_lbl = tk.Label(head_text, text="", bg=theme.BG, fg=theme.INK,
                                    font=(theme.FONT, 22, "bold"))
        self._title_lbl.pack(anchor="w")
        self._subtitle_lbl = tk.Label(head_text, text="", bg=theme.BG, fg=theme.MUTED,
                                       font=(theme.FONT, 12))
        self._subtitle_lbl.pack(anchor="w", pady=(2, 0))

        # Scrollbarer Bereich: Canvas + inneres Frame
        wrap = tk.Frame(content, bg=theme.BG)
        wrap.grid(row=1, column=0, sticky="nsew", padx=(20, 8), pady=(0, 16))
        wrap.rowconfigure(0, weight=1)
        wrap.columnconfigure(0, weight=1)
        self._canvas = tk.Canvas(wrap, bg=theme.BG, highlightthickness=0, bd=0)
        self._canvas.grid(row=0, column=0, sticky="nsew")
        vsb = ttk.Scrollbar(wrap, orient="vertical", command=self._canvas.yview)
        vsb.grid(row=0, column=1, sticky="ns")
        self._canvas.configure(yscrollcommand=vsb.set)
        self._inner = tk.Frame(self._canvas, bg=theme.BG)
        self._inner_id = self._canvas.create_window((0, 0), window=self._inner, anchor="nw")
        self._inner.bind("<Configure>",
                         lambda e: self._canvas.configure(scrollregion=self._canvas.bbox("all")))
        self._canvas.bind("<Configure>",
                          lambda e: self._canvas.itemconfigure(self._inner_id, width=e.width))
        # Mausrad
        self._canvas.bind_all("<MouseWheel>", self._on_wheel)

    def _on_wheel(self, ev):
        try:
            self._canvas.yview_scroll(int(-1 * (ev.delta / 120)), "units")
        except Exception:
            pass

    # ── Themenwechsel ────────────────────────────────────────────────────────
    def _select_topic(self, key):
        self._topic = key
        for k, b in self._nav_buttons.items():
            b.configure(bg=theme.SIDEBAR_ACTIVE if k == key else theme.SIDEBAR,
                        fg="white" if k == key else theme.SIDEBAR_TEXT)
        data = HELP_CONTENT.get(key, {})
        self._title_lbl.configure(text=data.get("title", ""))
        self._subtitle_lbl.configure(text=data.get("subtitle", ""))
        # Handout-Button passend zum Thema ein-/ausblenden
        self._current_handout = TOPIC_HANDOUTS.get(key)
        if self._current_handout:
            if not self._handout_btn.winfo_ismapped():
                self._handout_btn.pack(side="right", anchor="n", padx=(12, 0))
        else:
            self._handout_btn.pack_forget()
        self._render_blocks(key, data.get("blocks", []))
        self._canvas.yview_moveto(0.0)

    # ── Inhalt rendern ───────────────────────────────────────────────────────
    def _render_blocks(self, topic_key, blocks):
        for w in self._inner.winfo_children():
            w.destroy()
        self._img_refs.clear()
        col = tk.Frame(self._inner, bg=theme.BG)
        col.pack(fill="x", anchor="w", padx=8, pady=(2, 24))
        wrap_px = 760

        for block in blocks:
            kind = block[0]
            if kind == "h":
                tk.Label(col, text=block[1], bg=theme.BG, fg=theme.PRIMARY,
                         font=(theme.FONT, 14, "bold"), justify="left",
                         wraplength=wrap_px).pack(anchor="w", pady=(16, 4))
            elif kind == "p":
                tk.Label(col, text=block[1], bg=theme.BG, fg=theme.INK,
                         font=(theme.FONT, 11), justify="left",
                         wraplength=wrap_px).pack(anchor="w", pady=(2, 6))
            elif kind == "ul":
                for item in block[1]:
                    row = tk.Frame(col, bg=theme.BG)
                    row.pack(anchor="w", fill="x", pady=1)
                    tk.Label(row, text="•", bg=theme.BG, fg=theme.ACCENT,
                             font=(theme.FONT, 11, "bold")).pack(side="left", anchor="n", padx=(6, 8))
                    tk.Label(row, text=item, bg=theme.BG, fg=theme.INK, font=(theme.FONT, 11),
                             justify="left", wraplength=wrap_px - 30).pack(side="left", anchor="w")
            elif kind == "step":
                for i, item in enumerate(block[1], start=1):
                    row = tk.Frame(col, bg=theme.BG)
                    row.pack(anchor="w", fill="x", pady=2)
                    badge = tk.Label(row, text=str(i), bg=theme.PRIMARY, fg="white",
                                     font=(theme.FONT, 10, "bold"), width=2)
                    badge.pack(side="left", anchor="n", padx=(6, 10))
                    tk.Label(row, text=item, bg=theme.BG, fg=theme.INK, font=(theme.FONT, 11),
                             justify="left", wraplength=wrap_px - 40).pack(side="left", anchor="w")
            elif kind == "tip":
                self._callout(col, "💡  Tipp", block[1], theme.SUCCESS, "#EAF6EE")
            elif kind == "warn":
                self._callout(col, "⚠  Achtung", block[1], theme.WARNING, "#FBF3E2")
            elif kind == "img":
                self._image_block(col, topic_key, block[1], block[2], wrap_px)

    def _callout(self, parent, title, text, accent, bg):
        box = tk.Frame(parent, bg=bg, highlightbackground=accent, highlightthickness=1)
        box.pack(fill="x", anchor="w", pady=(8, 8))
        inner = tk.Frame(box, bg=bg)
        inner.pack(fill="x", padx=12, pady=10)
        tk.Label(inner, text=title, bg=bg, fg=accent,
                 font=(theme.FONT, 10, "bold")).pack(anchor="w")
        tk.Label(inner, text=text, bg=bg, fg=theme.INK, font=(theme.FONT, 11),
                 justify="left", wraplength=720).pack(anchor="w", pady=(2, 0))

    def _image_block(self, parent, topic_key, filename, caption, wrap_px):
        path = HELP_IMG_DIR / topic_key / filename
        img = self._load_image(path, max_w=wrap_px)
        frame = tk.Frame(parent, bg=theme.BG)
        frame.pack(anchor="w", pady=(10, 6))
        if img is not None:
            self._img_refs.append(img)
            holder = tk.Label(frame, image=img, bg=theme.CARD, bd=0,
                              highlightbackground=theme.BORDER, highlightthickness=1)
            holder.pack(anchor="w")
        else:
            # Platzhalter: sagt genau, welcher Screenshot wohin gehoert.
            ph = tk.Frame(frame, bg=theme.CARD_ALT, highlightbackground=theme.BORDER,
                          highlightthickness=1, width=wrap_px, height=210)
            ph.pack(anchor="w")
            ph.pack_propagate(False)
            tk.Label(ph, text="🖼", bg=theme.CARD_ALT, fg=theme.FAINT,
                     font=(theme.FONT, 34)).pack(pady=(34, 4))
            tk.Label(ph, text="Screenshot hier ablegen", bg=theme.CARD_ALT, fg=theme.MUTED,
                     font=(theme.FONT, 11, "bold")).pack()
            rel = os.path.join("assets", "hilfe", topic_key, filename)
            tk.Label(ph, text=rel, bg=theme.CARD_ALT, fg=theme.FAINT,
                     font=(theme.MONO, 9)).pack(pady=(4, 0))
        tk.Label(frame, text=caption, bg=theme.BG, fg=theme.MUTED,
                 font=(theme.FONT, 9, "italic"), justify="left",
                 wraplength=wrap_px).pack(anchor="w", pady=(4, 0))

    @staticmethod
    def _load_image(path: Path, max_w: int = 760):
        """Laedt ein Bild und skaliert es bei Bedarf auf max_w Breite (Seitenverhaeltnis
        bleibt). PIL bevorzugt; ohne PIL Fallback auf tk.PhotoImage (PNG/GIF)."""
        if not path.exists():
            return None
        try:
            from PIL import Image, ImageTk
            im = Image.open(path)
            im.load()
            if im.mode not in ("RGB", "RGBA"):
                im = im.convert("RGBA")
            if im.width > max_w:
                h = int(im.height * (max_w / im.width))
                im = im.resize((max_w, h), Image.LANCZOS)
            return ImageTk.PhotoImage(im)
        except Exception:
            try:
                img = tk.PhotoImage(file=str(path))
                if img.width() > max_w:
                    factor = max(1, int(img.width() / max_w + 0.999))
                    img = img.subsample(factor, factor)
                return img
            except Exception:
                return None

    # ── Handout drucken ──────────────────────────────────────────────────────
    def _open_handout(self):
        """Oeffnet das Handout zum aktuellen Thema im Standardbrowser. Dort laesst
        es sich ueber Strg+P als PDF speichern/ausdrucken."""
        fn = getattr(self, "_current_handout", None)
        if not fn:
            return
        path = HANDOUT_DIR / fn
        if not path.exists():
            messagebox.showinfo(
                "Handout",
                f"Das Handout wurde nicht gefunden:\n{path}\n\n"
                "Es sollte unter assets/handouts/ liegen.")
            return
        try:
            webbrowser.open(path.as_uri())
        except Exception as exc:
            messagebox.showwarning("Handout", f"Das Handout konnte nicht geöffnet werden:\n{exc}")

    # ── NMGone-Verbindung ────────────────────────────────────────────────────
    def _open_nmgone(self):
        if self._nmgone_action:
            try:
                self._nmgone_action()
                return
            except Exception:
                pass
        try:
            import subprocess
            if getattr(sys, "frozen", False):
                subprocess.Popen([sys.executable])
            else:
                start_py = Path(__file__).resolve().parent.parent / "start.py"
                subprocess.Popen([sys.executable, str(start_py)])
        except Exception:
            pass


def run_standalone():
    """Startet die Hilfe als eigenes Fenster (eigenes Taskleisten-Icon).
    Genutzt von start_hilfe.py und von NMGone.exe --hilfe."""
    if os.name == "nt":
        try:
            import ctypes
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("NMG.Hilfe")
        except Exception:
            pass
    root = tk.Tk()
    root.title(f"NMGone{DEMO_SUFFIX} Hilfe")
    root.geometry("1180x800")
    root.minsize(980, 640)
    root.configure(bg=theme.BG)
    theme.apply_theme(root)
    theme.apply_widget_defaults(root)
    for ico in ("Hilfe.ico", "NMGone.ico"):
        try:
            root.iconbitmap(str(ASSETS_DIR / ico))
            break
        except Exception:
            continue
    HilfePanel(root, on_close=root.destroy).pack(fill="both", expand=True)
    root.mainloop()


if __name__ == "__main__":
    sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
    run_standalone()
