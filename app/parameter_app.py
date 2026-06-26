#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Parameter & Berechtigungen · NMGone (eigenständige App)
=======================================================
Zentrale Stelle, an der festgelegt und nachgesehen wird, *welcher Mitarbeiter
was genau tun darf*. Pro Modul (NMGone, Kasse, Faktura, Einkauf, GDP,
Meldungen, Personal …) ist jede Berechtigung ein einzelner Punkt im Katalog.

Zwei Betriebsarten (oben rechts umschaltbar über das Schloss):
  • Ansehen (Standard)  – jeder darf alles *nachsehen*, nichts ändern.
  • Admin-Modus (PIN)   – freischalten, sperren, Rollen und Katalog pflegen.

Ansichten (Kopfleiste):
  • Übersicht / Matrix  – alle Mitarbeiter × alle Berechtigungen auf einen Blick
  • Mitarbeiter         – pro Person: Rollen zuweisen, einzelne Punkte
                          freischalten (✓) oder sperren (✗)
  • Rollen              – Bündel von Berechtigungen (Vertrieb, Lager, Labor …)
  • Katalog             – Stammliste aller Berechtigungs-Punkte je Modul
  • Protokoll           – wer hat wann was geändert (Revisionssicherheit)

Effektives Recht = Rollen des Mitarbeiters  ∪  Admin-Rolle
                   danach übersteuert eine persönliche Ausnahme (Override).

Start:  python start_parameter.py
Teilt sich die Datenbank mit NMGone (app/config.py · DB_PATH).
"""
from __future__ import annotations
import os
import sys
import getpass
import hashlib
import sqlite3
from datetime import datetime
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog

from . import theme  # zentrales Design-System (Cockpit-Optik, eine Quelle)

# ── Palette: zentral aus app/theme.py (gleiche Quelle wie Cockpit & übrige Apps)
PRIMARY      = theme.PRIMARY
PRIMARY_DARK = theme.PRIMARY_DARK
ACCENT       = theme.ACCENT
SUCCESS      = theme.SUCCESS
DANGER       = theme.DANGER
WARNING      = theme.WARNING
PURPLE       = theme.PURPLE
SIDEBAR      = theme.SIDEBAR
SIDEBAR_TEXT = theme.SIDEBAR_TEXT
BG           = theme.BG
CARD         = theme.CARD
CARD_ALT     = theme.CARD_ALT
INK          = theme.INK
MUTED        = theme.MUTED
FAINT        = theme.FAINT
BORDER       = theme.BORDER
DIVIDER      = theme.DIVIDER
SELECT_BG    = theme.SELECT_BG
FONT         = theme.FONT

from app.config import DB_PATH, fenstertitel, ASSETS_DIR  # gemeinsame NMGone-Datenbank

# Zustände eines Rechts für einen Mitarbeiter
ST_ALLOW, ST_DENY, ST_NONE = "allow", "deny", "none"

# ──────────────────────────────────────────────────────────────────────────────
#  Start-Katalog: alle Berechtigungs-Punkte, gruppiert nach Modul.
#  (modul, schluessel, titel, beschreibung)
#  Wird nur beim allerersten Start angelegt; danach im Katalog pflegbar.
# ──────────────────────────────────────────────────────────────────────────────
SEED_BERECHTIGUNGEN = [
    # ── Allgemein (modulübergreifend) ──────────────────────────────────────
    ("Allgemein", "app.start",            "Programm starten",          "Darf NMGone und die Apps überhaupt öffnen."),
    ("Allgemein", "daten.export",          "Daten exportieren",         "Listen/Tabellen als Excel, CSV oder PDF ausgeben."),
    ("Allgemein", "drucken",               "Drucken / Vorschau",        "Belege und Listen drucken oder als Vorschau öffnen."),
    ("Allgemein", "email.senden",          "Per E-Mail versenden",      "Belege/Listen per E-Mail verschicken."),
    ("Allgemein", "einstellungen.ansehen", "Einstellungen ansehen",     "Konfiguration einsehen (ohne Änderung)."),
    ("Allgemein", "einstellungen.aendern", "Einstellungen ändern",      "Programm-Einstellungen anpassen."),

    # ── NMGone · Analyse ───────────────────────────────────────────────────
    ("NMGone · Analyse", "bedarf.starten",   "Bedarfsanalyse starten",    "Neue Bedarfsanalyse (PK/ZW) durchführen."),
    ("NMGone · Analyse", "bedarf.export",    "Bedarfsanalyse exportieren","Kurzbericht/Excel der Analyse ausgeben."),
    ("NMGone · Analyse", "analysen.ansehen", "Gespeicherte Analysen öffnen","Vorhandene Analysen ansehen."),
    ("NMGone · Analyse", "analysen.loeschen","Gespeicherte Analysen löschen","Gespeicherte Analysen entfernen."),
    ("NMGone · Analyse", "produktanalyse",   "Produktanalyse",            "Produktchancen erstellen/auswerten."),
    ("NMGone · Analyse", "abweichungsanalyse","Abweichungsanalyse",       "Manuelle vs. Programm-Auswertung vergleichen."),
    ("NMGone · Analyse", "schulbank.ansehen","Schulbank ansehen",         "Lernvorschläge (Schulbank) einsehen."),
    ("NMGone · Analyse", "schulbank.entscheiden","Lernvorschläge entscheiden","Vorschläge übernehmen oder ablehnen."),
    ("NMGone · Analyse", "schulbank.manuelle_pruefung","Manuelle Prüfung","Manuelle Prüfung der Lernvorschläge."),
    ("NMGone · Analyse", "rabatte.ansehen",  "NMG-Rabatte ansehen",       "Rabatt-Übersicht, Statistik, Verlauf einsehen."),
    ("NMGone · Analyse", "rabatte.bearbeiten","NMG-Rabatte bearbeiten",    "Rabatte/Gültigkeiten ändern (mit Audit-Log)."),
    ("NMGone · Analyse", "austausch.ansehen","Austauschdatenbank ansehen","Biosimilar/Austausch-Einträge einsehen."),
    ("NMGone · Analyse", "austausch.bearbeiten","Austauschdatenbank pflegen","Austausch-/Biosimilar-Daten bearbeiten."),
    ("NMGone · Analyse", "kunden.ansehen",   "Kundenstamm ansehen",       "Kunden, Steckbrief und Historie einsehen."),
    ("NMGone · Analyse", "kunden.bearbeiten","Kundenstamm bearbeiten",    "Kunden anlegen/ändern, Listen importieren."),
    ("NMGone · Analyse", "kunden.email",     "Kunden-E-Mail-Versand",     "Serien-/Einzel-E-Mails an Kunden senden."),
    ("NMGone · Analyse", "suche.global",     "Globale / Vergleichs-Suche","Kunde, Analyse, PZN oder Artikel suchen."),
    ("NMGone · Analyse", "todo.ansehen",     "ToDo ansehen",              "Aufgaben und offene Punkte einsehen."),
    ("NMGone · Analyse", "todo.bearbeiten",  "ToDo bearbeiten",           "Aufgaben anlegen, ändern, erledigen."),

    # ── NMGone · Daten & Import ────────────────────────────────────────────
    ("NMGone · Daten", "import.zw",          "Manuelle Analysen / ZW importieren","Zukunftswerk-/manuelle Analysen einlesen."),
    ("NMGone · Daten", "import.partnerkonditionen","Partnerkonditionen importieren","PK-Datensätze einlesen."),
    ("NMGone · Daten", "import.apu_hap",     "APU/HAP-Daten importieren", "Apothekenpreise/Herstellerabgabepreise einlesen."),
    ("NMGone · Daten", "import.nmg_artikel", "NMG-Artikel importieren",   "NMG-Artikelstamm einlesen."),
    ("NMGone · Daten", "import.pk_rabatte",  "PK-Rabatte importieren",    "Partnerkonditionen-Rabatte einlesen."),
    ("NMGone · Daten", "artikelstamm.ansehen","Artikelstamm ansehen",     "Artikelstamm einsehen."),
    ("NMGone · Daten", "artikelstamm.bearbeiten","Artikelstamm pflegen",  "Artikelstamm-Einträge ändern."),
    ("NMGone · Daten", "auswertungsvorlage.bearbeiten","Auswertungsvorlage pflegen","Vorlage für Auswertungen anpassen."),

    # ── NMGone · System (Update / Backup / Datenbank) ──────────────────────
    ("NMGone · System", "system.update.suchen",     "Nach Updates suchen",      "Online nach neuer Version suchen."),
    ("NMGone · System", "system.update.installieren","Update installieren",     "Update-Paket einspielen (Neustart)."),
    ("NMGone · System", "backup.erstellen",         "Datensicherung erstellen", "Backup der Datenbank anstoßen."),
    ("NMGone · System", "backup.wiederherstellen",  "Sicherung wiederherstellen","Datenbank aus einem Backup zurückspielen."),
    ("NMGone · System", "db.uebersicht",            "Datenbankübersicht ansehen","Tabellen/Stände der Datenbank einsehen."),
    ("NMGone · System", "db.pfad.aendern",          "Cloud / DB-Pfad ändern",   "Speicherort der Datenbank umstellen."),
    ("NMGone · System", "roadmap.ansehen",          "Roadmap ansehen",          "Offene/erledigte Punkte einsehen."),
    ("NMGone · System", "roadmap.bearbeiten",       "Roadmap bearbeiten",       "Roadmap-Punkte anlegen/ändern."),

    # ── Kasse ──────────────────────────────────────────────────────────────
    ("Kasse", "kasse.uebersicht",        "Kasse-Übersicht ansehen",   "Kennzahlen und Status der Kasse sehen."),
    ("Kasse", "kasse.verkauf",           "Verkauf erfassen",          "Verkäufe an Apotheken erfassen."),
    ("Kasse", "kasse.vorbestellung",     "Vorbestellungen verarbeiten","Vorbestellungen importieren/übernehmen."),
    ("Kasse", "kasse.verkaeufe.import",  "Verkäufe importieren",      "Externe Verkaufsdaten einlesen."),
    ("Kasse", "kasse.wareneingang",      "Wareneingang buchen",       "Ware annehmen und ins Lager buchen."),
    ("Kasse", "kasse.lager.ansehen",     "Lagerbestand ansehen",      "Bestände und Lagerwert einsehen."),
    ("Kasse", "kasse.lager.bearbeiten",  "Lagerbestand / Inventur",   "Bestände manuell anpassen/inventieren."),
    ("Kasse", "kasse.preise.bearbeiten", "Preise bearbeiten",         "Verkaufs-/EK-Preise ändern."),
    ("Kasse", "kasse.lieferschein",      "Lieferscheine erstellen",   "Lieferscheine drucken/erzeugen."),
    ("Kasse", "kasse.auftragsbestaetigung","Auftragsbestätigung",     "Auftragsbestätigungen erstellen."),
    ("Kasse", "kasse.defektmeldung",     "Defektmeldung erfassen",    "Defekte Ware melden."),
    ("Kasse", "kasse.auswertung",        "Kassen-Auswertung",         "Umsatz, Tagesabschluss, Verfall, Inventur."),
    ("Kasse", "kasse.protokoll",         "Kassen-Protokoll ansehen",  "Protokoll der Kassen-Vorgänge einsehen."),
    ("Kasse", "kasse.einstellungen",     "Kassen-Einstellungen",      "Einstellungen der Kasse-App pflegen."),

    # ── Faktura ────────────────────────────────────────────────────────────
    ("Faktura", "faktura.auftrag",       "Aufträge bearbeiten",       "Aufträge anlegen und pflegen."),
    ("Faktura", "faktura.rechnung",      "Rechnungen erstellen",      "Ausgangsrechnungen schreiben."),
    ("Faktura", "faktura.gutschrift",    "Gutschriften erstellen",    "Gutschriften ausstellen."),
    ("Faktura", "faktura.quartal",       "Quartalsvergütung",         "Quartalsabrechnung/-vergütung erstellen."),
    ("Faktura", "faktura.staffel",       "Staffel / Konditionen",     "Staffeln und Konditionen pflegen."),
    ("Faktura", "faktura.einstellungen", "Faktura-Einstellungen",     "Vorlagen, Nummernkreise, Stammdaten pflegen."),

    # ── Einkauf ────────────────────────────────────────────────────────────
    ("Einkauf", "einkauf.dashboard",     "Einkauf-Dashboard",         "Einkaufs-Übersicht/Dashboard ansehen."),
    ("Einkauf", "einkauf.aufgaben",      "Aufgaben & Meldungen",      "Aufgaben/Meldungen im Einkauf bearbeiten."),
    ("Einkauf", "einkauf.lieferanten",   "Lieferanten pflegen",       "Lieferantenstamm anlegen/ändern."),
    ("Einkauf", "einkauf.quellen",       "Beschaffungsquellen",       "Beschaffungsquellen pflegen."),
    ("Einkauf", "einkauf.vorschlag",     "Beschaffungsvorschlag",     "Beschaffungsvorschläge erzeugen/prüfen."),
    ("Einkauf", "einkauf.beschaffung",   "Beschaffung EU-Ausland",    "Einkauf/Beschaffung anlegen und führen."),
    ("Einkauf", "einkauf.importkandidaten","Importkandidaten",        "Hersteller-Pipeline/Importkandidaten pflegen."),
    ("Einkauf", "einkauf.statistik",     "Statistik & Erfolg",        "Lieferanten-/Einkaufserfolg auswerten."),
    ("Einkauf", "einkauf.margenrechner", "Margenrechner §129",        "Margen-/Preiskalkulation nutzen."),
    ("Einkauf", "einkauf.kurse",         "Wechselkurse pflegen",      "Wechselkurse erfassen/aktualisieren."),
    ("Einkauf", "einkauf.einstellungen", "Einkauf-Einstellungen",     "Einstellungen der Einkauf-App pflegen."),

    # ── GDP (Wareneingang & Retouren) ──────────────────────────────────────
    ("GDP", "gdp.uebersicht",            "GDP-Übersicht",             "Offene GDP-Pflichten/Status ansehen."),
    ("GDP", "gdp.wareneingang",          "GDP-Wareneingang",          "Wareneingang mit Chargenerfassung."),
    ("GDP", "gdp.produktionsbestand",    "Produktionsbestand",        "Produktionsbestand verwalten."),
    ("GDP", "gdp.warenausgang",          "Warenausgang / Avis",       "Warenausgang und Avis bearbeiten."),
    ("GDP", "gdp.charge.ansehen",        "Chargen-Rückverfolgung",    "Chargen rückverfolgen/einsehen."),
    ("GDP", "gdp.bewegungen",            "Warenbewegungen ansehen",   "Alle Warenbewegungen einsehen."),
    ("GDP", "gdp.bestandsdiff",          "Bestandsdifferenzen",       "Bestandsdifferenzen prüfen/buchen."),
    ("GDP", "gdp.retoure",               "Retouren / Reklamation",    "Retouren annehmen und abwickeln."),
    ("GDP", "gdp.retourenbestand",       "Retourenbestand",           "Retourenbestand verwalten."),
    ("GDP", "gdp.gutschrift",            "Retouren-Gutschrift",       "Gutschrift für Retoure auslösen."),
    ("GDP", "gdp.kunde.qualifizieren",   "Kundenqualifizierung",      "Kundenqualifizierung (Erlaubnis) pflegen."),
    ("GDP", "gdp.protokoll",             "GDP-Protokoll ansehen",     "Protokoll der GDP-Vorgänge einsehen."),
    ("GDP", "gdp.einstellungen",         "GDP-Einstellungen",         "Einstellungen der GDP-App pflegen."),

    # ── Meldungen / Qualität ───────────────────────────────────────────────
    ("Meldungen", "meldungen.uebersicht","Meldungen-Übersicht",       "Offene Meldungen/Pflichten ansehen."),
    ("Meldungen", "meldungen.gdp",       "GDP-Meldungen erfassen",    "Abweichungen/Meldungen erfassen."),
    ("Meldungen", "meldungen.kuehlkette","Kühlsachenkontrolle",       "Kühlketten-/Temperaturkontrolle führen."),
    ("Meldungen", "meldungen.selbstinspektion","Selbstinspektion",    "Selbstinspektionen anlegen/auswerten."),
    ("Meldungen", "meldungen.protokoll", "Meldungen-Protokoll",       "Protokoll der Meldungen einsehen."),

    # ── Buchhaltung (Vorerfassung / DATEV) ─────────────────────────────────
    ("Buchhaltung", "buchhaltung.uebersicht","Buchhaltung-Übersicht", "Status/Übersicht der Buchhaltung ansehen."),
    ("Buchhaltung", "buchhaltung.belege","Belege erfassen",           "Belege vorerfassen."),
    ("Buchhaltung", "buchhaltung.kontenrahmen","Kontenrahmen pflegen","Konten/Kontenrahmen verwalten."),
    ("Buchhaltung", "buchhaltung.kontierung","Kontierung",            "Belege kontieren."),
    ("Buchhaltung", "buchhaltung.datev_export","DATEV-Export",        "Export ans Steuerbüro erzeugen."),
    ("Buchhaltung", "buchhaltung.erechnung","eRechnung empfangen",    "Eingehende eRechnungen verarbeiten."),
    ("Buchhaltung", "buchhaltung.einstellungen","Buchhaltung-Einstellungen","Berater-/Mandanten-/Wirtschaftsjahr pflegen."),

    # ── Personal ───────────────────────────────────────────────────────────
    ("Personal", "personal.organigramm.ansehen", "Organigramm ansehen", "Organigramm/Struktur einsehen."),
    ("Personal", "personal.organigramm.bearbeiten","Organigramm bearbeiten","Karten/Beziehungen ändern."),
    ("Personal", "personal.arbeitsbereiche","Arbeitsbereiche pflegen","Arbeitsbereiche und Vertretungen verwalten."),
    ("Personal", "personal.abwesenheit.ansehen","Abwesenheiten ansehen","Urlaub/Krankheit im Kalender sehen."),
    ("Personal", "personal.abwesenheit.bearbeiten","Abwesenheiten eintragen","Urlaub/Krankheit erfassen/ändern."),
    ("Personal", "personal.urlaub.entscheiden","Urlaub genehmigen",   "Über Urlaub/Verfall entscheiden."),
    ("Personal", "personal.stammdaten.bearbeiten","Mitarbeiter-Stammdaten","Mitarbeiter anlegen/ändern."),

    # ── Auswertungen (Report) ──────────────────────────────────────────────
    ("Auswertungen", "report.ansehen",   "Auswertungen öffnen",       "Verkäufe, Kunden, Artikel frei auswerten."),
    ("Auswertungen", "report.export",    "Auswertungen exportieren",  "Auswertungsergebnisse ausgeben."),

    # ── Verwaltung (diese Parameter-App) ───────────────────────────────────
    ("Verwaltung", "param.ansehen",      "Berechtigungen ansehen",    "Diese Parameter-App nur lesend nutzen."),
    ("Verwaltung", "param.admin",        "Berechtigungen verwalten",  "Rechte freischalten/sperren, Rollen & Katalog pflegen."),
]

# Start-Rollen: (name, beschreibung, ist_admin, [schluessel, ...])
# '*' = alle Berechtigungen (wird bei Admin automatisch über ist_admin abgedeckt).
SEED_ROLLEN = [
    ("Administrator", "Vollzugriff auf alles inkl. dieser Verwaltung.", 1, []),
    ("Geschäftsführung", "Sieht alles, steuert Personal & Einkauf.", 0, [
        "app.start", "daten.export", "drucken", "email.senden", "einstellungen.ansehen",
        "bedarf.starten", "bedarf.export", "analysen.ansehen", "produktanalyse",
        "abweichungsanalyse", "rabatte.ansehen", "kunden.ansehen", "suche.global",
        "todo.ansehen", "todo.bearbeiten", "db.uebersicht", "roadmap.ansehen", "roadmap.bearbeiten",
        "kasse.uebersicht", "kasse.auswertung", "faktura.quartal", "faktura.staffel",
        "einkauf.dashboard", "einkauf.statistik", "einkauf.margenrechner",
        "gdp.uebersicht", "meldungen.uebersicht", "buchhaltung.uebersicht",
        "personal.organigramm.ansehen", "personal.abwesenheit.ansehen", "personal.urlaub.entscheiden",
        "report.ansehen", "report.export", "param.ansehen",
    ]),
    ("Vertrieb / Außendienst", "Kunden, Verkauf, Analyse.", 0, [
        "app.start", "daten.export", "drucken", "email.senden",
        "bedarf.starten", "bedarf.export", "analysen.ansehen", "rabatte.ansehen",
        "austausch.ansehen", "kunden.ansehen", "kunden.bearbeiten", "kunden.email",
        "suche.global", "todo.ansehen", "todo.bearbeiten",
        "kasse.verkauf", "kasse.lager.ansehen", "kasse.lieferschein", "kasse.auftragsbestaetigung",
        "report.ansehen", "param.ansehen",
    ]),
    ("Innendienst", "Auftragsbearbeitung, Faktura-Vorbereitung.", 0, [
        "app.start", "daten.export", "drucken", "email.senden",
        "kunden.ansehen", "kunden.bearbeiten", "suche.global", "todo.ansehen", "todo.bearbeiten",
        "kasse.verkauf", "kasse.vorbestellung", "kasse.lager.ansehen",
        "kasse.lieferschein", "kasse.auftragsbestaetigung", "kasse.defektmeldung",
        "faktura.auftrag", "faktura.rechnung", "faktura.gutschrift",
        "report.ansehen", "param.ansehen",
    ]),
    ("Lager / Logistik", "Wareneingang, Bestände, Retouren.", 0, [
        "app.start", "drucken",
        "kasse.wareneingang", "kasse.lager.ansehen", "kasse.lager.bearbeiten", "kasse.defektmeldung",
        "gdp.uebersicht", "gdp.wareneingang", "gdp.produktionsbestand", "gdp.warenausgang",
        "gdp.charge.ansehen", "gdp.bewegungen", "gdp.bestandsdiff",
        "gdp.retoure", "gdp.retourenbestand",
        "param.ansehen",
    ]),
    ("Labor / QS", "Qualität, Meldungen, Selbstinspektion.", 0, [
        "app.start", "drucken",
        "gdp.charge.ansehen", "gdp.bewegungen",
        "meldungen.uebersicht", "meldungen.gdp", "meldungen.kuehlkette",
        "meldungen.selbstinspektion", "meldungen.protokoll",
        "param.ansehen",
    ]),
    ("Buchhaltung / Faktura", "Rechnungen, Gutschriften, Buchhaltung.", 0, [
        "app.start", "daten.export", "drucken", "email.senden",
        "faktura.auftrag", "faktura.rechnung", "faktura.gutschrift", "faktura.quartal",
        "faktura.staffel", "faktura.einstellungen",
        "buchhaltung.uebersicht", "buchhaltung.belege", "buchhaltung.kontenrahmen",
        "buchhaltung.kontierung", "buchhaltung.datev_export", "buchhaltung.erechnung",
        "gdp.gutschrift", "report.ansehen", "report.export", "param.ansehen",
    ]),
    ("Einkauf", "Beschaffung & Margen.", 0, [
        "app.start", "daten.export", "drucken",
        "einkauf.dashboard", "einkauf.aufgaben", "einkauf.lieferanten", "einkauf.quellen",
        "einkauf.vorschlag", "einkauf.beschaffung", "einkauf.importkandidaten",
        "einkauf.statistik", "einkauf.margenrechner", "einkauf.kurse",
        "rabatte.ansehen", "suche.global", "param.ansehen",
    ]),
    ("GDP-Beauftragter", "Wareneingang, Chargen, Kundenqualifizierung, Meldungen.", 0, [
        "app.start", "daten.export", "drucken",
        "gdp.uebersicht", "gdp.wareneingang", "gdp.produktionsbestand", "gdp.warenausgang",
        "gdp.charge.ansehen", "gdp.bewegungen", "gdp.bestandsdiff", "gdp.retoure",
        "gdp.retourenbestand", "gdp.gutschrift", "gdp.kunde.qualifizieren", "gdp.protokoll",
        "meldungen.uebersicht", "meldungen.gdp", "meldungen.kuehlkette",
        "meldungen.selbstinspektion", "meldungen.protokoll",
        "param.ansehen",
    ]),
    ("Gast / Nur-Lesen", "Darf nur ansehen, nichts ändern.", 0, [
        "app.start", "einstellungen.ansehen",
        "analysen.ansehen", "rabatte.ansehen", "austausch.ansehen", "kunden.ansehen",
        "suche.global", "todo.ansehen",
        "kasse.uebersicht", "kasse.lager.ansehen", "kasse.auswertung",
        "gdp.uebersicht", "meldungen.uebersicht",
        "personal.organigramm.ansehen", "personal.abwesenheit.ansehen",
        "report.ansehen", "param.ansehen",
    ]),
]


# Beim Erweitern von SEED_BERECHTIGUNGEN hochzaehlen: zwingt bestehende
# Datenbanken, fehlende Katalog-Punkte EINMAL nachzuziehen (siehe _sync_catalog).
SEED_CATALOG_VERSION = 2


def _hash_pin(pin: str) -> str:
    return hashlib.sha256(("nmgparam::" + (pin or "")).encode("utf-8")).hexdigest()


def _sync_catalog(con) -> int:
    """Bringt eine bereits bestueckte DB auf den aktuellen Katalogstand.

    Laeuft nur, wenn die gespeicherte catalog_version < SEED_CATALOG_VERSION ist
    (also einmal pro Versionssprung). Dadurch:
      * neue Berechtigungs-Punkte werden nachgetragen,
      * vom Admin geloeschte Built-ins kommen NICHT bei jedem Start zurueck,
      * Modul/Reihenfolge der Built-ins bleiben sauber gruppiert.
    Gibt die Zahl neu angelegter Punkte zurueck.
    """
    row = con.execute("SELECT wert FROM tbl_param_config WHERE schluessel='catalog_version'").fetchone()
    have_ver = int(row[0]) if row and str(row[0]).isdigit() else 0
    if have_ver >= SEED_CATALOG_VERSION:
        return 0

    seed_index = {key: i for i, (_, key, _, _) in enumerate(SEED_BERECHTIGUNGEN)}
    existing = {r[0] for r in con.execute("SELECT schluessel FROM tbl_berechtigung")}
    added = 0
    for modul, key, titel, besch in SEED_BERECHTIGUNGEN:
        if key in existing:
            # Built-in: Modul + Sortierung kanonisch halten (Titel/Beschreibung
            # NICHT ueberschreiben, damit Admin-Anpassungen erhalten bleiben).
            con.execute("UPDATE tbl_berechtigung SET modul=?, sort=? WHERE schluessel=?",
                        (modul, seed_index[key], key))
        else:
            con.execute("INSERT INTO tbl_berechtigung(modul,schluessel,titel,beschreibung,sort) "
                        "VALUES(?,?,?,?,?)", (modul, key, titel, besch, seed_index[key]))
            added += 1
    # selbst angelegte (Nicht-Built-in) Punkte ans Ende sortieren
    con.execute("UPDATE tbl_berechtigung SET sort=10000+id "
                "WHERE schluessel NOT IN (%s)" % ",".join("?" * len(seed_index)),
                tuple(seed_index.keys()))
    con.execute("INSERT INTO tbl_param_config(schluessel,wert) VALUES('catalog_version',?) "
                "ON CONFLICT(schluessel) DO UPDATE SET wert=excluded.wert",
                (str(SEED_CATALOG_VERSION),))
    return added


# ──────────────────────────────────────────────────────────────────────────────
#  Datenbank
# ──────────────────────────────────────────────────────────────────────────────
def init_db():
    con = sqlite3.connect(DB_PATH)
    # Mitarbeiter-Tabelle teilen wir mit der Personal-App; falls eine frische
    # DB ohne Personal-App genutzt wird, legen wir das kompatible Minimum an.
    con.execute("""CREATE TABLE IF NOT EXISTS tbl_mitarbeiter (
        id INTEGER PRIMARY KEY AUTOINCREMENT, vorname TEXT, name TEXT,
        abteilung TEXT, position TEXT, board_x INTEGER DEFAULT 60, board_y INTEGER DEFAULT 60,
        urlaubsanspruch INTEGER DEFAULT 30, personalverantwortlich INTEGER DEFAULT 0)""")

    con.execute("""CREATE TABLE IF NOT EXISTS tbl_berechtigung (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        modul TEXT NOT NULL DEFAULT 'Allgemein',
        schluessel TEXT NOT NULL UNIQUE,
        titel TEXT NOT NULL,
        beschreibung TEXT DEFAULT '',
        sort INTEGER DEFAULT 0,
        aktiv INTEGER DEFAULT 1)""")
    con.execute("""CREATE TABLE IF NOT EXISTS tbl_rolle (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT NOT NULL UNIQUE,
        beschreibung TEXT DEFAULT '',
        ist_admin INTEGER DEFAULT 0)""")
    con.execute("""CREATE TABLE IF NOT EXISTS tbl_rolle_recht (
        rolle_id INTEGER NOT NULL,
        berechtigung_id INTEGER NOT NULL,
        PRIMARY KEY (rolle_id, berechtigung_id))""")
    con.execute("""CREATE TABLE IF NOT EXISTS tbl_ma_rolle (
        mitarbeiter_id INTEGER NOT NULL,
        rolle_id INTEGER NOT NULL,
        PRIMARY KEY (mitarbeiter_id, rolle_id))""")
    # erlaubt: 1 = persönlich freigeschaltet, 0 = persönlich gesperrt (übersteuert Rolle)
    con.execute("""CREATE TABLE IF NOT EXISTS tbl_ma_recht_override (
        mitarbeiter_id INTEGER NOT NULL,
        berechtigung_id INTEGER NOT NULL,
        erlaubt INTEGER NOT NULL,
        PRIMARY KEY (mitarbeiter_id, berechtigung_id))""")
    con.execute("""CREATE TABLE IF NOT EXISTS tbl_param_config (
        schluessel TEXT PRIMARY KEY, wert TEXT)""")
    con.execute("""CREATE TABLE IF NOT EXISTS tbl_param_audit (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        zeitpunkt TEXT DEFAULT CURRENT_TIMESTAMP,
        benutzer TEXT, aktion TEXT, detail TEXT)""")

    # Start-Katalog nur einmal befüllen
    have = con.execute("SELECT COUNT(*) FROM tbl_berechtigung").fetchone()[0]
    if not have:
        for i, (modul, key, titel, besch) in enumerate(SEED_BERECHTIGUNGEN):
            con.execute("INSERT INTO tbl_berechtigung(modul,schluessel,titel,beschreibung,sort) "
                        "VALUES(?,?,?,?,?)", (modul, key, titel, besch, i))
        con.execute("INSERT INTO tbl_param_config(schluessel,wert) VALUES('catalog_version',?) "
                    "ON CONFLICT(schluessel) DO UPDATE SET wert=excluded.wert",
                    (str(SEED_CATALOG_VERSION),))
    else:
        # Bereits bestueckte DB: fehlende Katalog-Punkte einmalig nachziehen.
        _sync_catalog(con)
    # Start-Rollen nur einmal anlegen
    if not con.execute("SELECT COUNT(*) FROM tbl_rolle").fetchone()[0]:
        key_to_id = {r[0]: r[1] for r in con.execute("SELECT schluessel,id FROM tbl_berechtigung")}
        for name, besch, is_admin, keys in SEED_ROLLEN:
            cur = con.execute("INSERT INTO tbl_rolle(name,beschreibung,ist_admin) VALUES(?,?,?)",
                              (name, besch, is_admin))
            rid = cur.lastrowid
            for k in keys:
                if k in key_to_id:
                    con.execute("INSERT OR IGNORE INTO tbl_rolle_recht(rolle_id,berechtigung_id) "
                                "VALUES(?,?)", (rid, key_to_id[k]))
    con.commit()
    con.close()


def _audit(aktion: str, detail: str = ""):
    try:
        con = sqlite3.connect(DB_PATH)
        con.execute("INSERT INTO tbl_param_audit(benutzer,aktion,detail) VALUES(?,?,?)",
                    (getpass.getuser(), aktion, detail))
        con.commit()
        con.close()
    except Exception:
        pass


# ──────────────────────────────────────────────────────────────────────────────
#  App
# ──────────────────────────────────────────────────────────────────────────────
class App:
    def __init__(self, root):
        self.root = root
        root.title(fenstertitel("NMGone · Parameter & Berechtigungen"))
        root.geometry("1240x800")
        # Im Vollbild (maximiert) starten. 'zoomed' = Windows; sonst -zoomed/Bildschirmgroesse.
        try:
            root.state("zoomed")
        except tk.TclError:
            try:
                root.attributes("-zoomed", True)
            except tk.TclError:
                root.geometry(f"{root.winfo_screenwidth()}x{root.winfo_screenheight()}+0+0")
        root.configure(bg=BG)
        self.style = theme.apply_theme(root)   # zentrale ttk-Optik (Cockpit-Look)
        theme.apply_widget_defaults(root)
        self._style_tree()

        self.admin = False
        self.view = "matrix"
        self.sel_ma = None        # gewählter Mitarbeiter (Mitarbeiter-Ansicht)
        self.sel_rolle = None     # gewählte Rolle (Rollen-Ansicht)
        self.matrix_modul = "Alle"

        # Cockpit-Layout: links dunkle Sidebar, rechts Arbeitsbereich.
        shell = tk.Frame(root, bg=BG)
        shell.pack(fill="both", expand=True)
        self._build_sidebar(shell)

        right = tk.Frame(shell, bg=BG)
        right.pack(side="left", fill="both", expand=True)
        self._build_topbar(right)
        self.main = tk.Frame(right, bg=BG)
        self.main.pack(side="top", fill="both", expand=True)
        self._build_status(right)

        self.load()
        self.show_matrix()

    # ---- ttk-Styles -------------------------------------------------------
    def _style_tree(self):
        s = self.style
        s.configure("P.Treeview", background=CARD, fieldbackground=CARD, foreground=INK,
                    borderwidth=0, rowheight=28, font=(FONT, 10))
        s.map("P.Treeview", background=[("selected", SELECT_BG)], foreground=[("selected", PRIMARY)])
        s.configure("P.Treeview.Heading", background="#EEF3F8", foreground=PRIMARY,
                    relief="flat", font=(FONT, 10, "bold"), padding=6)
        s.configure("TCombobox", fieldbackground=CARD, background=CARD, foreground=INK,
                    arrowcolor=PRIMARY, padding=5)

    # ---- Linke Navigation (Cockpit-Sidebar) -------------------------------
    def _build_sidebar(self, parent):
        self.sidebar = theme.Sidebar(parent, width=250, title="Parameter",
                                     subtitle="Berechtigungen")
        self.sidebar.pack(side="left", fill="y")
        # NMGone-Logo oben in der Sidebar (wie in den übrigen Apps).
        try:
            _logo_path = ASSETS_DIR / "NMGone.png"
            if _logo_path.exists():
                _raw = tk.PhotoImage(file=str(_logo_path))
                _factor = max(1, round(max(_raw.width() / 200, _raw.height() / 120)))
                self._sidebar_logo = _raw.subsample(_factor, _factor)
                self.sidebar.set_logo(self._sidebar_logo)
        except Exception:
            pass
        self.sidebar.add_section("Ansichten")
        for key, icon, label in (("matrix", "📋", "Übersicht"),
                                 ("ma", "👤", "Mitarbeiter"),
                                 ("rollen", "🧩", "Rollen")):
            self.sidebar.add_item(key, icon, label, lambda k=key: self._nav(k),
                                  active=(key == "matrix"))
        self.sidebar.add_section("Verwaltung")
        for key, icon, label in (("katalog", "🗂", "Katalog"),
                                 ("einst", "⚙", "Einstellungen"),
                                 ("audit", "📜", "Protokoll")):
            self.sidebar.add_item(key, icon, label, lambda k=key: self._nav(k))
        self.sidebar.add_footer_note("Wer darf was —\nzentral festlegen.")

    def _build_topbar(self, parent):
        """Schmale Kopfleiste rechts oben: Admin-/Ansehen-Umschalter (Schloss)."""
        h = tk.Frame(parent, bg=BG, height=54)
        h.pack(side="top", fill="x")
        h.pack_propagate(False)
        tk.Label(h, text="🔐  Parameter & Berechtigungen", bg=BG, fg=INK,
                 font=(FONT, 15, "bold")).pack(side="left", padx=(22, 6), pady=12)
        self.lock_btn = tk.Button(h, text="🔒  Ansehen", command=self.toggle_admin,
                                  bg=SIDEBAR, fg="white", relief="flat",
                                  font=(FONT, 10, "bold"), padx=14, pady=7, cursor="hand2",
                                  activebackground=PRIMARY_DARK, activeforeground="white")
        self.lock_btn.pack(side="right", padx=(0, 22))

    def _nav(self, key):
        {"matrix": self.show_matrix, "ma": self.show_mitarbeiter,
         "rollen": self.show_rollen, "katalog": self.show_katalog,
         "einst": self.show_einstellungen, "audit": self.show_audit}[key]()

    def _nav_state(self):
        # Aktiven Eintrag in der Sidebar hervorheben.
        self.sidebar.set_active(self.view)

    def _build_status(self, parent):
        self.status = tk.StringVar(value="Bereit.")
        s = tk.Frame(parent, bg="#E3EAF1", height=26)
        s.pack(side="bottom", fill="x")
        tk.Label(s, textvariable=self.status, bg="#E3EAF1", fg=MUTED,
                 font=(FONT, 9), anchor="w").pack(side="left", padx=14)
        self.mode_lbl = tk.Label(s, text="", bg="#E3EAF1", fg=MUTED, font=(FONT, 9, "bold"))
        self.mode_lbl.pack(side="right", padx=14)
        self._refresh_mode_label()

    def _refresh_mode_label(self):
        if self.admin:
            self.mode_lbl.configure(text="● Admin-Modus aktiv — Änderungen möglich", fg=SUCCESS)
        else:
            self.mode_lbl.configure(text="● Ansehen — schreibgeschützt", fg=MUTED)

    def _clear_main(self):
        for w in self.main.winfo_children():
            w.destroy()

    # ---- Admin-Modus ------------------------------------------------------
    def _get_pin_hash(self):
        con = sqlite3.connect(DB_PATH)
        row = con.execute("SELECT wert FROM tbl_param_config WHERE schluessel='admin_pin'").fetchone()
        con.close()
        return row[0] if row else None

    def _set_pin_hash(self, h):
        con = sqlite3.connect(DB_PATH)
        con.execute("INSERT INTO tbl_param_config(schluessel,wert) VALUES('admin_pin',?) "
                    "ON CONFLICT(schluessel) DO UPDATE SET wert=excluded.wert", (h,))
        con.commit()
        con.close()

    def toggle_admin(self):
        if self.admin:
            self.admin = False
            self._after_mode_change("Admin-Modus verlassen.")
            return
        pin_hash = self._get_pin_hash()
        if not pin_hash:
            # Erste Einrichtung: PIN festlegen
            if not messagebox.askyesno(
                "Admin-PIN einrichten",
                "Es ist noch keine Admin-PIN vergeben.\n\n"
                "Jetzt eine PIN festlegen? Danach lässt sich der Admin-Modus nur "
                "noch mit dieser PIN öffnen.", parent=self.root):
                return
            p1 = simpledialog.askstring("Neue Admin-PIN", "Neue PIN eingeben (mind. 4 Zeichen):",
                                        show="•", parent=self.root)
            if not p1 or len(p1) < 4:
                messagebox.showinfo("Admin-PIN", "Abgebrochen — PIN muss mindestens 4 Zeichen haben.", parent=self.root)
                return
            p2 = simpledialog.askstring("Bestätigen", "PIN zur Bestätigung erneut eingeben:",
                                        show="•", parent=self.root)
            if p1 != p2:
                messagebox.showwarning("Admin-PIN", "Die Eingaben stimmen nicht überein.", parent=self.root)
                return
            self._set_pin_hash(_hash_pin(p1))
            _audit("admin_pin_gesetzt", "Erste Admin-PIN eingerichtet.")
            self.admin = True
            self._after_mode_change("Admin-PIN eingerichtet — Admin-Modus aktiv.")
            return
        # PIN abfragen
        p = simpledialog.askstring("Admin-Modus", "Admin-PIN eingeben:", show="•", parent=self.root)
        if p is None:
            return
        if _hash_pin(p) == pin_hash:
            self.admin = True
            _audit("admin_login", "Admin-Modus geöffnet.")
            self._after_mode_change("Admin-Modus aktiv — Änderungen möglich.")
        else:
            _audit("admin_login_fehler", "Falsche PIN eingegeben.")
            messagebox.showwarning("Admin-Modus", "Falsche PIN.", parent=self.root)

    def _after_mode_change(self, msg):
        self.lock_btn.configure(
            text=("🔓  Admin" if self.admin else "🔒  Ansehen"),
            bg=(SUCCESS if self.admin else SIDEBAR),
            activebackground=(theme.SUCCESS_DARK if self.admin else PRIMARY_DARK))
        self._refresh_mode_label()
        self.status.set(msg)
        self._nav(self.view)  # aktuelle Ansicht mit neuen Rechten neu zeichnen

    def _require_admin(self) -> bool:
        if not self.admin:
            messagebox.showinfo(
                "Nur Ansehen",
                "Dafür ist der Admin-Modus nötig.\n\n"
                "Oben rechts auf das Schloss klicken und die Admin-PIN eingeben.",
                parent=self.root)
            return False
        return True

    def change_pin(self):
        if not self._require_admin():
            return
        p1 = simpledialog.askstring("PIN ändern", "Neue PIN (mind. 4 Zeichen):", show="•", parent=self.root)
        if not p1 or len(p1) < 4:
            return
        p2 = simpledialog.askstring("Bestätigen", "Neue PIN erneut eingeben:", show="•", parent=self.root)
        if p1 != p2:
            messagebox.showwarning("PIN ändern", "Die Eingaben stimmen nicht überein.", parent=self.root)
            return
        self._set_pin_hash(_hash_pin(p1))
        _audit("admin_pin_geaendert", "Admin-PIN geändert.")
        messagebox.showinfo("PIN ändern", "Admin-PIN wurde geändert.", parent=self.root)

    # ---- Daten laden ------------------------------------------------------
    def load(self):
        con = sqlite3.connect(DB_PATH)
        con.row_factory = sqlite3.Row
        self.emps = [dict(r) for r in con.execute(
            "SELECT id,vorname,name,abteilung,position FROM tbl_mitarbeiter "
            "ORDER BY name, vorname").fetchall()]
        self.berecht = [dict(r) for r in con.execute(
            "SELECT * FROM tbl_berechtigung WHERE aktiv=1 ORDER BY sort, id").fetchall()]
        self.rollen = [dict(r) for r in con.execute(
            "SELECT * FROM tbl_rolle ORDER BY ist_admin DESC, name").fetchall()]
        rr = con.execute("SELECT rolle_id, berechtigung_id FROM tbl_rolle_recht").fetchall()
        mr = con.execute("SELECT mitarbeiter_id, rolle_id FROM tbl_ma_rolle").fetchall()
        ov = con.execute("SELECT mitarbeiter_id, berechtigung_id, erlaubt FROM tbl_ma_recht_override").fetchall()
        con.close()

        self.b_by_id = {b["id"]: b for b in self.berecht}
        self.role_by_id = {r["id"]: r for r in self.rollen}
        self.role_rights = {}
        for rid, bid in rr:
            self.role_rights.setdefault(rid, set()).add(bid)
        self.ma_roles = {}
        for mid, rid in mr:
            self.ma_roles.setdefault(mid, set()).add(rid)
        self.overrides = {(mid, bid): erl for mid, bid, erl in ov}

        # Module in stabiler Reihenfolge (nach erstem Auftreten)
        self.module = []
        for b in self.berecht:
            if b["modul"] not in self.module:
                self.module.append(b["modul"])

    def _name(self, mid):
        e = next((x for x in self.emps if x["id"] == mid), None)
        return f"{e['vorname']} {e['name']}".strip() if e else f"#{mid}"

    def _initials(self, e):
        return ((e.get("vorname") or " ")[:1] + (e.get("name") or " ")[:1]).upper().strip() or "?"

    def effective(self, mid):
        """schluessel/berechtigung_id -> (state, quelle). Override schlägt Rolle."""
        roles = self.ma_roles.get(mid, set())
        is_admin = any(self.role_by_id.get(r, {}).get("ist_admin") for r in roles)
        res = {}
        for b in self.berecht:
            bid = b["id"]
            ov = self.overrides.get((mid, bid))
            if ov is not None:
                res[bid] = (ST_ALLOW if ov == 1 else ST_DENY, "override")
            elif is_admin:
                res[bid] = (ST_ALLOW, "admin")
            elif any(bid in self.role_rights.get(r, set()) for r in roles):
                res[bid] = (ST_ALLOW, "rolle")
            else:
                res[bid] = (ST_NONE, "")
        return res

    # =========================================================================
    #  ANSICHT 1 · ÜBERSICHT / MATRIX
    # =========================================================================
    def show_matrix(self):
        self.view = "matrix"
        self._nav_state()
        self._clear_main()

        bar = tk.Frame(self.main, bg=CARD, height=50)
        bar.pack(side="top", fill="x")
        bar.configure(highlightbackground=BORDER, highlightthickness=1)
        tk.Label(bar, text="Modul:", bg=CARD, fg=MUTED, font=(FONT, 10, "bold")).pack(side="left", padx=(16, 6), pady=10)
        self.modul_cb = ttk.Combobox(bar, state="readonly", width=18,
                                     values=["Alle"] + self.module)
        self.modul_cb.set(self.matrix_modul if self.matrix_modul in (["Alle"] + self.module) else "Alle")
        self.modul_cb.pack(side="left", pady=10)
        self.modul_cb.bind("<<ComboboxSelected>>", lambda e: self._draw_matrix())
        tk.Label(bar, text="✓ erlaubt    ✗ gesperrt    ·  kein Zugriff",
                 bg=CARD, fg=FAINT, font=(FONT, 9)).pack(side="left", padx=18)
        tk.Label(bar, text="Doppelklick auf eine Zeile öffnet den Mitarbeiter.",
                 bg=CARD, fg=FAINT, font=(FONT, 9)).pack(side="right", padx=16)

        self.matrix_wrap = tk.Frame(self.main, bg=BG)
        self.matrix_wrap.pack(side="top", fill="both", expand=True)
        self._draw_matrix()

    def _draw_matrix(self):
        for w in self.matrix_wrap.winfo_children():
            w.destroy()
        self.matrix_modul = self.modul_cb.get()
        if not self.emps:
            self._empty_hint(self.matrix_wrap,
                             "Noch keine Mitarbeiter angelegt.",
                             "Mitarbeiter werden in der Mitarbeiter-App (Organigramm) gepflegt.\n"
                             "Diese App teilt sich die Mitarbeiterliste mit NMGone.")
            return

        berecht = [b for b in self.berecht
                   if self.matrix_modul == "Alle" or b["modul"] == self.matrix_modul]

        cols = [f"e{e['id']}" for e in self.emps]
        tree = ttk.Treeview(self.matrix_wrap, style="P.Treeview", columns=cols,
                            show="tree headings", selectmode="browse")
        tree.heading("#0", text="Berechtigung")
        tree.column("#0", width=300, anchor="w", stretch=False)
        for e in self.emps:
            tree.heading(f"e{e['id']}", text=self._initials(e))
            tree.column(f"e{e['id']}", width=46, anchor="center", stretch=False)

        tree.tag_configure("modul", background="#EEF3F8", foreground=PRIMARY, font=(FONT, 10, "bold"))

        eff_by_ma = {e["id"]: self.effective(e["id"]) for e in self.emps}
        cur_modul = None
        parent = ""
        for b in berecht:
            if b["modul"] != cur_modul:
                cur_modul = b["modul"]
                parent = tree.insert("", "end", text=f"  {cur_modul}",
                                     values=[""] * len(cols), open=True, tags=("modul",))
            vals = []
            for e in self.emps:
                state = eff_by_ma[e["id"]][b["id"]][0]
                vals.append({ST_ALLOW: "✓", ST_DENY: "✗", ST_NONE: "·"}[state])
            tree.insert(parent, "end", text=f"     {b['titel']}", values=vals,
                        tags=(f"b{b['id']}",))

        vsb = ttk.Scrollbar(self.matrix_wrap, orient="vertical", command=tree.yview)
        hsb = ttk.Scrollbar(self.matrix_wrap, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        self.matrix_wrap.rowconfigure(0, weight=1)
        self.matrix_wrap.columnconfigure(0, weight=1)

        def on_dbl(_ev):
            col = tree.identify_column(tree.winfo_pointerx() - tree.winfo_rootx())
            try:
                idx = int(col.replace("#", "")) - 1
            except ValueError:
                idx = -1
            if 0 <= idx < len(self.emps):
                self.sel_ma = self.emps[idx]["id"]
                self.show_mitarbeiter()
        tree.bind("<Double-Button-1>", on_dbl)

        self.status.set(f"Übersicht · {len(berecht)} Berechtigungen × {len(self.emps)} Mitarbeiter · "
                        f"Modul: {self.matrix_modul}")

    # =========================================================================
    #  ANSICHT 2 · MITARBEITER (Rollen + einzelne Punkte)
    # =========================================================================
    def show_mitarbeiter(self):
        self.view = "ma"
        self._nav_state()
        self._clear_main()
        if not self.emps:
            self._empty_hint(self.main, "Noch keine Mitarbeiter angelegt.",
                             "Mitarbeiter werden in der Mitarbeiter-App gepflegt.")
            return
        if self.sel_ma is None or self.sel_ma not in {e["id"] for e in self.emps}:
            self.sel_ma = self.emps[0]["id"]

        # Aktionsleiste: Sammel-Zuweisung + Rechte kopieren/übertragen
        tb = tk.Frame(self.main, bg=CARD, height=48)
        tb.pack(side="top", fill="x")
        tb.configure(highlightbackground=BORDER, highlightthickness=1)
        tk.Label(tb, text="Mitarbeiter-Berechtigungen", bg=CARD, fg=PRIMARY,
                 font=(FONT, 12, "bold")).pack(side="left", padx=16, pady=10)
        if self.admin:
            tk.Button(tb, text="👥  Rolle an mehrere zuweisen", command=self._bulk_role_dialog,
                      bg=PRIMARY, fg="white", relief="flat", font=(FONT, 10, "bold"),
                      padx=12, pady=6, cursor="hand2").pack(side="left", padx=(8, 4), pady=8)
            tk.Button(tb, text="📋  Rechte kopieren / übertragen",
                      command=lambda: self._copy_rights_dialog(self.sel_ma),
                      bg=ACCENT, fg="white", relief="flat", font=(FONT, 10, "bold"),
                      padx=12, pady=6, cursor="hand2").pack(side="left", padx=4, pady=8)
        else:
            tk.Label(tb, text="Zum Zuweisen/Kopieren Admin-Modus aktivieren (Schloss oben rechts).",
                     bg=CARD, fg=FAINT, font=(FONT, 9)).pack(side="left", padx=10)

        body = tk.Frame(self.main, bg=BG)
        body.pack(fill="both", expand=True)

        # linke Liste
        left = tk.Frame(body, bg=CARD, width=250)
        left.pack(side="left", fill="y")
        left.pack_propagate(False)
        left.configure(highlightbackground=BORDER, highlightthickness=1)
        tk.Label(left, text="Mitarbeiter", bg=CARD, fg=PRIMARY,
                 font=(FONT, 11, "bold")).pack(anchor="w", padx=14, pady=(12, 6))
        self.ma_list = tk.Listbox(left, bg=CARD, fg=INK, font=(FONT, 10), bd=0,
                                  highlightthickness=0, activestyle="none",
                                  selectbackground=SELECT_BG, selectforeground=PRIMARY)
        self.ma_list.pack(fill="both", expand=True, padx=8, pady=(0, 10))
        for e in self.emps:
            self.ma_list.insert("end", f"  {e['vorname']} {e['name']}".rstrip())
        idx = next((i for i, e in enumerate(self.emps) if e["id"] == self.sel_ma), 0)
        self.ma_list.selection_set(idx)
        self.ma_list.see(idx)
        self.ma_list.bind("<<ListboxSelect>>", self._on_ma_select)

        # rechte Detailfläche
        self.ma_detail = tk.Frame(body, bg=BG)
        self.ma_detail.pack(side="left", fill="both", expand=True)
        self._draw_ma_detail()

    def _on_ma_select(self, _ev):
        sel = self.ma_list.curselection()
        if not sel:
            return
        self.sel_ma = self.emps[sel[0]]["id"]
        self._draw_ma_detail()

    def _draw_ma_detail(self):
        for w in self.ma_detail.winfo_children():
            w.destroy()
        e = next((x for x in self.emps if x["id"] == self.sel_ma), None)
        if not e:
            return

        head = tk.Frame(self.ma_detail, bg=BG)
        head.pack(fill="x", padx=18, pady=(14, 6))
        if self.admin:
            tk.Button(head, text="→ Rechte auf andere übertragen",
                      command=lambda src=self.sel_ma: self._copy_rights_dialog(src),
                      bg="#EDF1F6", fg=PRIMARY, relief="flat", font=(FONT, 9, "bold"),
                      padx=10, pady=5, cursor="hand2").pack(side="right", anchor="n")
        tk.Label(head, text=f"{e['vorname']} {e['name']}".strip(), bg=BG, fg=INK,
                 font=(FONT, 18, "bold")).pack(anchor="w")
        sub = " · ".join(x for x in [e.get("abteilung"), e.get("position")] if x)
        if sub:
            tk.Label(head, text=sub, bg=BG, fg=MUTED, font=(FONT, 11)).pack(anchor="w")

        # Rollen-Zuweisung
        rcard = tk.Frame(self.ma_detail, bg=CARD, highlightbackground=BORDER, highlightthickness=1)
        rcard.pack(fill="x", padx=18, pady=8)
        tk.Label(rcard, text="Rollen", bg=CARD, fg=PRIMARY, font=(FONT, 12, "bold")).pack(anchor="w", padx=14, pady=(12, 2))
        tk.Label(rcard, text="Rollen bündeln Berechtigungen. Mehrfachauswahl möglich.",
                 bg=CARD, fg=MUTED, font=(FONT, 9)).pack(anchor="w", padx=14)
        rwrap = tk.Frame(rcard, bg=CARD)
        rwrap.pack(fill="x", padx=10, pady=10)
        my_roles = self.ma_roles.get(self.sel_ma, set())
        self._role_vars = {}
        for i, r in enumerate(self.rollen):
            var = tk.BooleanVar(value=(r["id"] in my_roles))
            self._role_vars[r["id"]] = var
            txt = r["name"] + ("  ★" if r["ist_admin"] else "")
            cb = tk.Checkbutton(rwrap, text=txt, variable=var, bg=CARD, fg=INK,
                                font=(FONT, 10), anchor="w", selectcolor=CARD,
                                activebackground=CARD, cursor=("hand2" if self.admin else "arrow"),
                                state=("normal" if self.admin else "disabled"),
                                command=lambda rid=r["id"]: self._toggle_role(rid))
            cb.grid(row=i // 3, column=i % 3, sticky="w", padx=8, pady=3)

        # Berechtigungen pro Modul
        cap = tk.Frame(self.ma_detail, bg=BG)
        cap.pack(fill="x", padx=18, pady=(8, 2))
        tk.Label(cap, text="Berechtigungen", bg=BG, fg=PRIMARY, font=(FONT, 12, "bold")).pack(side="left")
        tk.Label(cap, text="(✓ erlaubt · ✗ gesperrt · grau = aus Rolle/Admin · gelb = persönliche Ausnahme)",
                 bg=BG, fg=FAINT, font=(FONT, 9)).pack(side="left", padx=8)

        # scrollbarer Bereich
        wrap = tk.Frame(self.ma_detail, bg=BG)
        wrap.pack(fill="both", expand=True, padx=18, pady=(2, 8))
        canvas = tk.Canvas(wrap, bg=BG, highlightthickness=0)
        vsb = ttk.Scrollbar(wrap, orient="vertical", command=canvas.yview)
        inner = tk.Frame(canvas, bg=BG)
        inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=inner, anchor="nw", width=860)
        canvas.configure(yscrollcommand=vsb.set)
        canvas.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")
        canvas.bind_all("<MouseWheel>", lambda ev: canvas.yview_scroll(int(-ev.delta / 120), "units"))

        eff = self.effective(self.sel_ma)
        cur_modul = None
        for b in self.berecht:
            if b["modul"] != cur_modul:
                cur_modul = b["modul"]
                tk.Label(inner, text=cur_modul.upper(), bg=BG, fg=ACCENT,
                         font=(FONT, 9, "bold")).pack(anchor="w", padx=4, pady=(12, 2))
            self._perm_row(inner, b, eff[b["id"]])

        self.status.set(f"{e['vorname']} {e['name']} · {len(my_roles)} Rolle(n) · "
                        + ("Admin-Modus: Punkte freischalten/sperren möglich." if self.admin
                           else "Ansehen — zum Ändern Admin-Modus aktivieren."))

    def _perm_row(self, parent, b, eff):
        state, quelle = eff
        row = tk.Frame(parent, bg=CARD, highlightbackground=BORDER, highlightthickness=1)
        row.pack(fill="x", padx=4, pady=2)
        # Status-Punkt
        dot = {ST_ALLOW: "✓", ST_DENY: "✗", ST_NONE: "·"}[state]
        dotcol = {ST_ALLOW: SUCCESS, ST_DENY: DANGER, ST_NONE: FAINT}[state]
        tk.Label(row, text=dot, bg=CARD, fg=dotcol, font=(FONT, 13, "bold"), width=2).pack(side="left", padx=(10, 4), pady=6)
        info = tk.Frame(row, bg=CARD)
        info.pack(side="left", fill="x", expand=True)
        tk.Label(info, text=b["titel"], bg=CARD, fg=INK, font=(FONT, 10, "bold"), anchor="w").pack(anchor="w")
        tk.Label(info, text=b["beschreibung"] or b["schluessel"], bg=CARD, fg=MUTED,
                 font=(FONT, 9), anchor="w").pack(anchor="w")
        # Quelle-Badge
        badge_txt = {"admin": "Admin-Rolle", "rolle": "aus Rolle",
                     "override": "persönliche Ausnahme", "": "—"}[quelle]
        badge_col = {"admin": PURPLE, "rolle": ACCENT, "override": WARNING, "": FAINT}[quelle]
        tk.Label(row, text=badge_txt, bg=CARD, fg=badge_col, font=(FONT, 9, "bold")).pack(side="left", padx=10)

        if self.admin:
            btns = tk.Frame(row, bg=CARD)
            btns.pack(side="right", padx=8)
            self._mini_btn(btns, "✓ frei", SUCCESS, lambda: self._set_override(b["id"], 1))
            self._mini_btn(btns, "✗ sperr", DANGER, lambda: self._set_override(b["id"], 0))
            if quelle == "override":
                self._mini_btn(btns, "↺ Standard", "#6B7886", lambda: self._set_override(b["id"], None))

    def _mini_btn(self, parent, text, color, cmd):
        tk.Button(parent, text=text, command=cmd, bg=color, fg="white", relief="flat",
                  font=(FONT, 9, "bold"), padx=8, pady=3, cursor="hand2",
                  activebackground=color, activeforeground="white").pack(side="left", padx=2)

    def _toggle_role(self, rid):
        if not self.admin:
            return
        on = self._role_vars[rid].get()
        con = sqlite3.connect(DB_PATH)
        if on:
            con.execute("INSERT OR IGNORE INTO tbl_ma_rolle(mitarbeiter_id,rolle_id) VALUES(?,?)",
                        (self.sel_ma, rid))
        else:
            con.execute("DELETE FROM tbl_ma_rolle WHERE mitarbeiter_id=? AND rolle_id=?",
                        (self.sel_ma, rid))
        con.commit()
        con.close()
        _audit("rolle_zuordnung", f"{self._name(self.sel_ma)} · Rolle '{self.role_by_id[rid]['name']}' "
                                  f"{'zugewiesen' if on else 'entfernt'}")
        self.load()
        self._draw_ma_detail()

    def _set_override(self, bid, erlaubt):
        if not self._require_admin():
            return
        con = sqlite3.connect(DB_PATH)
        if erlaubt is None:
            con.execute("DELETE FROM tbl_ma_recht_override WHERE mitarbeiter_id=? AND berechtigung_id=?",
                        (self.sel_ma, bid))
            akt = "auf Standard zurückgesetzt"
        else:
            con.execute("INSERT INTO tbl_ma_recht_override(mitarbeiter_id,berechtigung_id,erlaubt) "
                        "VALUES(?,?,?) ON CONFLICT(mitarbeiter_id,berechtigung_id) "
                        "DO UPDATE SET erlaubt=excluded.erlaubt", (self.sel_ma, bid, erlaubt))
            akt = "freigeschaltet" if erlaubt == 1 else "gesperrt"
        con.commit()
        con.close()
        _audit("recht_override", f"{self._name(self.sel_ma)} · '{self.b_by_id[bid]['titel']}' {akt}")
        self.load()
        self._draw_ma_detail()

    # ---- Sammel-Zuweisung & Kopieren -------------------------------------
    def _emp_multipicker(self, parent, exclude=None, preselect=None, height=12):
        """Scrollbare Mehrfachauswahl der Mitarbeiter (Listbox 'extended') mit
        Alle/Keine und Abteilungsfilter. Liefert eine Funktion -> Liste von IDs."""
        exclude = exclude or set()
        emps = [e for e in self.emps if e["id"] not in exclude]
        box = tk.Frame(parent, bg=BG)
        top = tk.Frame(box, bg=BG)
        top.pack(fill="x", pady=(0, 4))
        tk.Label(top, text="Abteilung:", bg=BG, fg=MUTED, font=(FONT, 9)).pack(side="left")
        depts = ["Alle"] + sorted({(e.get("abteilung") or "—") for e in emps})
        dept_cb = ttk.Combobox(top, state="readonly", width=18, values=depts)
        dept_cb.set("Alle")
        dept_cb.pack(side="left", padx=6)
        lb = tk.Listbox(box, selectmode="extended", height=height, bg=CARD, fg=INK,
                        font=(FONT, 10), bd=0, highlightthickness=1,
                        highlightbackground=BORDER, activestyle="none",
                        selectbackground=SELECT_BG, selectforeground=PRIMARY,
                        exportselection=False)
        lb.pack(fill="both", expand=True, pady=2)
        shown = []  # parallele Liste der angezeigten emp-ids

        def refill():
            d = dept_cb.get()
            lb.delete(0, "end")
            shown.clear()
            for e in emps:
                if d != "Alle" and (e.get("abteilung") or "—") != d:
                    continue
                shown.append(e["id"])
                lb.insert("end", f"  {e['vorname']} {e['name']}".rstrip()
                          + (f"   ({e['abteilung']})" if e.get("abteilung") else ""))
            if preselect:
                for i, mid in enumerate(shown):
                    if mid in preselect:
                        lb.selection_set(i)
        refill()
        dept_cb.bind("<<ComboboxSelected>>", lambda _e: refill())
        btnrow = tk.Frame(box, bg=BG)
        btnrow.pack(fill="x", pady=(4, 0))
        tk.Button(btnrow, text="Alle", command=lambda: lb.selection_set(0, "end"),
                  padx=10, pady=3, cursor="hand2").pack(side="left")
        tk.Button(btnrow, text="Keine", command=lambda: lb.selection_clear(0, "end"),
                  padx=10, pady=3, cursor="hand2").pack(side="left", padx=6)
        return box, (lambda: [shown[i] for i in lb.curselection()])

    def _bulk_role_dialog(self):
        if not self._require_admin():
            return
        if not self.rollen:
            return
        win = tk.Toplevel(self.root)
        win.title("Rolle an mehrere Mitarbeiter")
        win.configure(bg=BG)
        win.transient(self.root)
        win.grab_set()
        win.geometry("520x560")
        tk.Label(win, text="Rolle an mehrere Mitarbeiter zuweisen",
                 bg=BG, fg=INK, font=(FONT, 14, "bold")).pack(anchor="w", padx=16, pady=(14, 2))
        tk.Label(win, text="Wähle eine Rolle, die Aktion und die betroffenen Mitarbeiter.",
                 bg=BG, fg=MUTED, font=(FONT, 10)).pack(anchor="w", padx=16)

        row = tk.Frame(win, bg=BG)
        row.pack(fill="x", padx=16, pady=(12, 4))
        tk.Label(row, text="Rolle:", bg=BG, fg=PRIMARY, font=(FONT, 10, "bold")).pack(side="left")
        role_cb = ttk.Combobox(row, state="readonly", width=30,
                               values=[r["name"] + ("  ★" if r["ist_admin"] else "") for r in self.rollen])
        role_cb.current(0)
        role_cb.pack(side="left", padx=8)

        mode = tk.StringVar(value="zuweisen")
        mrow = tk.Frame(win, bg=BG)
        mrow.pack(fill="x", padx=16, pady=(2, 8))
        tk.Radiobutton(mrow, text="Zuweisen (hinzufügen)", variable=mode, value="zuweisen",
                       bg=BG, fg=INK, font=(FONT, 10), selectcolor=CARD, activebackground=BG).pack(side="left")
        tk.Radiobutton(mrow, text="Entfernen", variable=mode, value="entfernen",
                       bg=BG, fg=INK, font=(FONT, 10), selectcolor=CARD, activebackground=BG).pack(side="left", padx=12)

        tk.Label(win, text="Mitarbeiter:", bg=BG, fg=PRIMARY, font=(FONT, 10, "bold")).pack(anchor="w", padx=16)
        picker, get_ids = self._emp_multipicker(win, height=12)
        picker.pack(fill="both", expand=True, padx=16, pady=(2, 8))

        def apply():
            targets = get_ids()
            if not targets:
                messagebox.showinfo("Hinweis", "Bitte mindestens einen Mitarbeiter wählen.", parent=win)
                return
            r = self.rollen[role_cb.current()]
            add = (mode.get() == "zuweisen")
            con = sqlite3.connect(DB_PATH)
            for mid in targets:
                if add:
                    con.execute("INSERT OR IGNORE INTO tbl_ma_rolle(mitarbeiter_id,rolle_id) VALUES(?,?)", (mid, r["id"]))
                else:
                    con.execute("DELETE FROM tbl_ma_rolle WHERE mitarbeiter_id=? AND rolle_id=?", (mid, r["id"]))
            con.commit()
            con.close()
            _audit("rolle_sammel", f"Rolle '{r['name']}' {'zugewiesen an' if add else 'entfernt von'} "
                                   f"{len(targets)} Mitarbeiter: " + ", ".join(self._name(m) for m in targets))
            win.destroy()
            self.load()
            self.show_mitarbeiter()
            self.status.set(f"Rolle '{r['name']}' bei {len(targets)} Mitarbeiter(n) "
                            f"{'zugewiesen' if add else 'entfernt'}.")

        bar = tk.Frame(win, bg=BG)
        bar.pack(fill="x", padx=16, pady=12)
        tk.Button(bar, text="Abbrechen", command=win.destroy, padx=14, pady=7).pack(side="right", padx=(8, 0))
        tk.Button(bar, text="Anwenden", command=apply, bg=PRIMARY, fg="white", relief="flat",
                  padx=18, pady=8, cursor="hand2").pack(side="right")

    def _copy_rights_dialog(self, source_id=None):
        if not self._require_admin():
            return
        if len(self.emps) < 2:
            messagebox.showinfo("Rechte kopieren",
                                "Dafür werden mindestens zwei Mitarbeiter benötigt.", parent=self.root)
            return
        win = tk.Toplevel(self.root)
        win.title("Rechte kopieren / übertragen")
        win.configure(bg=BG)
        win.transient(self.root)
        win.grab_set()
        win.geometry("560x640")
        tk.Label(win, text="Rechte von einem Mitarbeiter übertragen",
                 bg=BG, fg=INK, font=(FONT, 14, "bold")).pack(anchor="w", padx=16, pady=(14, 2))
        tk.Label(win, text="Kopiert die Rollen (und optional die persönlichen Ausnahmen) eines\n"
                           "Vorlage-Mitarbeiters auf einen oder mehrere Ziel-Mitarbeiter.",
                 bg=BG, fg=MUTED, font=(FONT, 10), justify="left").pack(anchor="w", padx=16)

        row = tk.Frame(win, bg=BG)
        row.pack(fill="x", padx=16, pady=(12, 6))
        tk.Label(row, text="Vorlage (Quelle):", bg=BG, fg=PRIMARY, font=(FONT, 10, "bold")).pack(side="left")
        src_cb = ttk.Combobox(row, state="readonly", width=34,
                              values=[f"{e['vorname']} {e['name']}".strip()
                                      + (f"  ({e['abteilung']})" if e.get("abteilung") else "")
                                      for e in self.emps])
        src_idx = next((i for i, e in enumerate(self.emps) if e["id"] == source_id), 0)
        src_cb.current(src_idx)
        src_cb.pack(side="left", padx=8)

        opt = tk.Frame(win, bg=BG)
        opt.pack(fill="x", padx=16, pady=(2, 4))
        copy_roles = tk.BooleanVar(value=True)
        copy_overrides = tk.BooleanVar(value=False)
        tk.Checkbutton(opt, text="Rollen kopieren", variable=copy_roles, bg=BG, fg=INK,
                       font=(FONT, 10), selectcolor=CARD, activebackground=BG).pack(anchor="w")
        tk.Checkbutton(opt, text="Persönliche Ausnahmen (Freischaltungen/Sperren) mitkopieren",
                       variable=copy_overrides, bg=BG, fg=INK, font=(FONT, 10),
                       selectcolor=CARD, activebackground=BG).pack(anchor="w")

        mode = tk.StringVar(value="ergaenzen")
        mrow = tk.Frame(win, bg=BG)
        mrow.pack(fill="x", padx=16, pady=(2, 8))
        tk.Radiobutton(mrow, text="Ergänzen (zu vorhandenen hinzufügen)", variable=mode, value="ergaenzen",
                       bg=BG, fg=INK, font=(FONT, 10), selectcolor=CARD, activebackground=BG).pack(anchor="w")
        tk.Radiobutton(mrow, text="Ersetzen (Ziel zuerst leeren, dann übernehmen)", variable=mode, value="ersetzen",
                       bg=BG, fg=INK, font=(FONT, 10), selectcolor=CARD, activebackground=BG).pack(anchor="w")

        tk.Label(win, text="Ziel-Mitarbeiter:", bg=BG, fg=PRIMARY, font=(FONT, 10, "bold")).pack(anchor="w", padx=16)
        # Quelle aus Zielauswahl ausschließen; bei Wechsel der Quelle neu aufbauen
        picker_holder = tk.Frame(win, bg=BG)
        picker_holder.pack(fill="both", expand=True, padx=16, pady=(2, 8))
        state = {"get_ids": None}

        def rebuild_picker(*_a):
            for w in picker_holder.winfo_children():
                w.destroy()
            src = self.emps[src_cb.current()]["id"]
            pk, getter = self._emp_multipicker(picker_holder, exclude={src}, height=11)
            pk.pack(fill="both", expand=True)
            state["get_ids"] = getter
        rebuild_picker()
        src_cb.bind("<<ComboboxSelected>>", rebuild_picker)

        def apply():
            src = self.emps[src_cb.current()]["id"]
            targets = state["get_ids"]() if state["get_ids"] else []
            if not targets:
                messagebox.showinfo("Hinweis", "Bitte mindestens einen Ziel-Mitarbeiter wählen.", parent=win)
                return
            if not copy_roles.get() and not copy_overrides.get():
                messagebox.showinfo("Hinweis", "Bitte mindestens 'Rollen' oder 'Ausnahmen' wählen.", parent=win)
                return
            ersetzen = (mode.get() == "ersetzen")
            src_roles = sorted(self.ma_roles.get(src, set()))
            src_ovr = [(bid, erl) for (mid, bid), erl in self.overrides.items() if mid == src]
            con = sqlite3.connect(DB_PATH)
            for mid in targets:
                if copy_roles.get():
                    if ersetzen:
                        con.execute("DELETE FROM tbl_ma_rolle WHERE mitarbeiter_id=?", (mid,))
                    for rid in src_roles:
                        con.execute("INSERT OR IGNORE INTO tbl_ma_rolle(mitarbeiter_id,rolle_id) VALUES(?,?)", (mid, rid))
                if copy_overrides.get():
                    if ersetzen:
                        con.execute("DELETE FROM tbl_ma_recht_override WHERE mitarbeiter_id=?", (mid,))
                    for bid, erl in src_ovr:
                        con.execute("INSERT INTO tbl_ma_recht_override(mitarbeiter_id,berechtigung_id,erlaubt) "
                                    "VALUES(?,?,?) ON CONFLICT(mitarbeiter_id,berechtigung_id) "
                                    "DO UPDATE SET erlaubt=excluded.erlaubt", (mid, bid, erl))
            con.commit()
            con.close()
            was = []
            if copy_roles.get():
                was.append(f"{len(src_roles)} Rolle(n)")
            if copy_overrides.get():
                was.append(f"{len(src_ovr)} Ausnahme(n)")
            _audit("rechte_kopiert",
                   f"{' + '.join(was)} von {self._name(src)} ({mode.get()}) auf "
                   f"{len(targets)} Mitarbeiter: " + ", ".join(self._name(m) for m in targets))
            win.destroy()
            self.load()
            self.show_mitarbeiter()
            self.status.set(f"{' + '.join(was)} von {self._name(src)} auf "
                            f"{len(targets)} Mitarbeiter übertragen ({mode.get()}).")

        bar = tk.Frame(win, bg=BG)
        bar.pack(fill="x", padx=16, pady=12)
        tk.Button(bar, text="Abbrechen", command=win.destroy, padx=14, pady=7).pack(side="right", padx=(8, 0))
        tk.Button(bar, text="Übertragen", command=apply, bg=PRIMARY, fg="white", relief="flat",
                  padx=18, pady=8, cursor="hand2").pack(side="right")

    # =========================================================================
    #  ANSICHT 3 · ROLLEN
    # =========================================================================
    def show_rollen(self):
        self.view = "rollen"
        self._nav_state()
        self._clear_main()

        body = tk.Frame(self.main, bg=BG)
        body.pack(fill="both", expand=True)

        left = tk.Frame(body, bg=CARD, width=270)
        left.pack(side="left", fill="y")
        left.pack_propagate(False)
        left.configure(highlightbackground=BORDER, highlightthickness=1)
        head = tk.Frame(left, bg=CARD)
        head.pack(fill="x", padx=14, pady=(12, 6))
        tk.Label(head, text="Rollen", bg=CARD, fg=PRIMARY, font=(FONT, 11, "bold")).pack(side="left")
        if self.admin:
            tk.Button(head, text="➕ Neu", command=self._rolle_neu, bg=PRIMARY, fg="white",
                      relief="flat", font=(FONT, 9, "bold"), padx=8, pady=3, cursor="hand2").pack(side="right")
        self.rolle_list = tk.Listbox(left, bg=CARD, fg=INK, font=(FONT, 10), bd=0,
                                     highlightthickness=0, activestyle="none",
                                     selectbackground=SELECT_BG, selectforeground=PRIMARY)
        self.rolle_list.pack(fill="both", expand=True, padx=8, pady=(0, 10))
        for r in self.rollen:
            n = sum(1 for _ in self.role_rights.get(r["id"], set()))
            star = "★ " if r["ist_admin"] else ""
            self.rolle_list.insert("end", f"  {star}{r['name']}  ({'alle' if r['ist_admin'] else n})")
        if self.sel_rolle is None or self.sel_rolle not in {r["id"] for r in self.rollen}:
            self.sel_rolle = self.rollen[0]["id"] if self.rollen else None
        if self.sel_rolle is not None:
            idx = next((i for i, r in enumerate(self.rollen) if r["id"] == self.sel_rolle), 0)
            self.rolle_list.selection_set(idx)
        self.rolle_list.bind("<<ListboxSelect>>", self._on_rolle_select)

        self.rolle_detail = tk.Frame(body, bg=BG)
        self.rolle_detail.pack(side="left", fill="both", expand=True)
        self._draw_rolle_detail()

    def _on_rolle_select(self, _ev):
        sel = self.rolle_list.curselection()
        if not sel:
            return
        self.sel_rolle = self.rollen[sel[0]]["id"]
        self._draw_rolle_detail()

    def _draw_rolle_detail(self):
        for w in self.rolle_detail.winfo_children():
            w.destroy()
        r = self.role_by_id.get(self.sel_rolle)
        if not r:
            self._empty_hint(self.rolle_detail, "Keine Rolle gewählt.", "")
            return

        head = tk.Frame(self.rolle_detail, bg=BG)
        head.pack(fill="x", padx=18, pady=(14, 4))
        tk.Label(head, text=r["name"], bg=BG, fg=INK, font=(FONT, 18, "bold")).pack(side="left")
        if self.admin and not r["ist_admin"]:
            tk.Button(head, text="🗑 Löschen", command=lambda: self._rolle_loeschen(r),
                      bg="#EDF1F6", fg=DANGER, relief="flat", font=(FONT, 9, "bold"),
                      padx=10, pady=4, cursor="hand2").pack(side="right")
            tk.Button(head, text="✏ Umbenennen", command=lambda: self._rolle_umbenennen(r),
                      bg="#EDF1F6", fg=PRIMARY, relief="flat", font=(FONT, 9, "bold"),
                      padx=10, pady=4, cursor="hand2").pack(side="right", padx=6)
        tk.Label(self.rolle_detail, text=r["beschreibung"] or "", bg=BG, fg=MUTED,
                 font=(FONT, 11)).pack(anchor="w", padx=18)

        if r["ist_admin"]:
            box = tk.Frame(self.rolle_detail, bg="#F3EEFA", highlightbackground=PURPLE, highlightthickness=1)
            box.pack(fill="x", padx=18, pady=14)
            tk.Label(box, text="★  Diese Rolle hat automatisch ALLE Berechtigungen.",
                     bg="#F3EEFA", fg=PURPLE, font=(FONT, 11, "bold")).pack(anchor="w", padx=14, pady=12)
            n_ma = sum(1 for v in self.ma_roles.values() if r["id"] in v)
            tk.Label(self.rolle_detail, text=f"Zugewiesen an {n_ma} Mitarbeiter.",
                     bg=BG, fg=MUTED, font=(FONT, 10)).pack(anchor="w", padx=18)
            return

        tk.Label(self.rolle_detail, text="Enthaltene Berechtigungen"
                 + ("  ·  Häkchen setzen/entfernen" if self.admin else "  ·  (nur Ansehen)"),
                 bg=BG, fg=PRIMARY, font=(FONT, 12, "bold")).pack(anchor="w", padx=18, pady=(10, 2))

        wrap = tk.Frame(self.rolle_detail, bg=BG)
        wrap.pack(fill="both", expand=True, padx=18, pady=(2, 10))
        canvas = tk.Canvas(wrap, bg=BG, highlightthickness=0)
        vsb = ttk.Scrollbar(wrap, orient="vertical", command=canvas.yview)
        inner = tk.Frame(canvas, bg=BG)
        inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=inner, anchor="nw", width=860)
        canvas.configure(yscrollcommand=vsb.set)
        canvas.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")
        canvas.bind_all("<MouseWheel>", lambda ev: canvas.yview_scroll(int(-ev.delta / 120), "units"))

        have = self.role_rights.get(r["id"], set())
        self._rolle_vars = {}
        cur_modul = None
        for b in self.berecht:
            if b["modul"] != cur_modul:
                cur_modul = b["modul"]
                tk.Label(inner, text=cur_modul.upper(), bg=BG, fg=ACCENT,
                         font=(FONT, 9, "bold")).pack(anchor="w", padx=4, pady=(12, 2))
            var = tk.BooleanVar(value=(b["id"] in have))
            self._rolle_vars[b["id"]] = var
            cb = tk.Checkbutton(inner, text=f"  {b['titel']}   —   {b['beschreibung']}",
                                variable=var, bg=BG, fg=INK, font=(FONT, 10), anchor="w",
                                selectcolor=CARD, activebackground=BG,
                                cursor=("hand2" if self.admin else "arrow"),
                                state=("normal" if self.admin else "disabled"),
                                command=lambda bid=b["id"]: self._toggle_rolle_recht(r["id"], bid))
            cb.pack(anchor="w", padx=6)

        n_ma = sum(1 for v in self.ma_roles.values() if r["id"] in v)
        self.status.set(f"Rolle '{r['name']}' · {len(have)} Berechtigungen · an {n_ma} Mitarbeiter zugewiesen")

    def _toggle_rolle_recht(self, rid, bid):
        if not self.admin:
            return
        on = self._rolle_vars[bid].get()
        con = sqlite3.connect(DB_PATH)
        if on:
            con.execute("INSERT OR IGNORE INTO tbl_rolle_recht(rolle_id,berechtigung_id) VALUES(?,?)", (rid, bid))
        else:
            con.execute("DELETE FROM tbl_rolle_recht WHERE rolle_id=? AND berechtigung_id=?", (rid, bid))
        con.commit()
        con.close()
        _audit("rolle_recht", f"Rolle '{self.role_by_id[rid]['name']}' · '{self.b_by_id[bid]['titel']}' "
                              f"{'hinzugefügt' if on else 'entfernt'}")
        self.load()

    def _rolle_neu(self):
        if not self._require_admin():
            return
        name = simpledialog.askstring("Neue Rolle", "Name der Rolle:", parent=self.root)
        if not name:
            return
        besch = simpledialog.askstring("Neue Rolle", "Kurzbeschreibung (optional):", parent=self.root) or ""
        con = sqlite3.connect(DB_PATH)
        try:
            cur = con.execute("INSERT INTO tbl_rolle(name,beschreibung,ist_admin) VALUES(?,?,0)", (name, besch))
            new_id = cur.lastrowid
            con.commit()
        except sqlite3.IntegrityError:
            messagebox.showwarning("Neue Rolle", "Eine Rolle mit diesem Namen existiert bereits.", parent=self.root)
            con.close()
            return
        con.close()
        _audit("rolle_neu", f"Rolle '{name}' angelegt")
        self.sel_rolle = new_id
        self.load()
        self.show_rollen()

    def _rolle_umbenennen(self, r):
        if not self._require_admin():
            return
        name = simpledialog.askstring("Umbenennen", "Neuer Name:", initialvalue=r["name"], parent=self.root)
        if not name or name == r["name"]:
            return
        con = sqlite3.connect(DB_PATH)
        try:
            con.execute("UPDATE tbl_rolle SET name=? WHERE id=?", (name, r["id"]))
            con.commit()
        except sqlite3.IntegrityError:
            messagebox.showwarning("Umbenennen", "Name bereits vergeben.", parent=self.root)
            con.close()
            return
        con.close()
        _audit("rolle_umbenannt", f"'{r['name']}' → '{name}'")
        self.load()
        self.show_rollen()

    def _rolle_loeschen(self, r):
        if not self._require_admin():
            return
        n_ma = sum(1 for v in self.ma_roles.values() if r["id"] in v)
        if not messagebox.askyesno("Rolle löschen",
                f"Rolle '{r['name']}' wirklich löschen?\n\n"
                f"Sie ist aktuell {n_ma} Mitarbeiter(n) zugewiesen — diese Zuordnungen "
                f"werden mit entfernt.", parent=self.root):
            return
        con = sqlite3.connect(DB_PATH)
        con.execute("DELETE FROM tbl_rolle WHERE id=?", (r["id"],))
        con.execute("DELETE FROM tbl_rolle_recht WHERE rolle_id=?", (r["id"],))
        con.execute("DELETE FROM tbl_ma_rolle WHERE rolle_id=?", (r["id"],))
        con.commit()
        con.close()
        _audit("rolle_geloescht", f"Rolle '{r['name']}' gelöscht")
        self.sel_rolle = None
        self.load()
        self.show_rollen()

    # =========================================================================
    #  ANSICHT 4 · KATALOG
    # =========================================================================
    def show_katalog(self):
        self.view = "katalog"
        self._nav_state()
        self._clear_main()

        bar = tk.Frame(self.main, bg=CARD, height=50)
        bar.pack(side="top", fill="x")
        bar.configure(highlightbackground=BORDER, highlightthickness=1)
        tk.Label(bar, text="Katalog aller Berechtigungs-Punkte", bg=CARD, fg=PRIMARY,
                 font=(FONT, 12, "bold")).pack(side="left", padx=16, pady=10)
        if self.admin:
            tk.Button(bar, text="➕ Neuer Punkt", command=lambda: self._katalog_form(None),
                      bg=PRIMARY, fg="white", relief="flat", font=(FONT, 10, "bold"),
                      padx=12, pady=6, cursor="hand2").pack(side="right", padx=16, pady=8)
            tk.Button(bar, text="🔑 PIN ändern", command=self.change_pin,
                      bg="#EDF1F6", fg=PRIMARY, relief="flat", font=(FONT, 10, "bold"),
                      padx=12, pady=6, cursor="hand2").pack(side="right", pady=8)
        else:
            tk.Label(bar, text="Zum Bearbeiten Admin-Modus aktivieren.",
                     bg=CARD, fg=FAINT, font=(FONT, 9)).pack(side="right", padx=16)

        wrap = tk.Frame(self.main, bg=BG)
        wrap.pack(fill="both", expand=True)
        cols = ("modul", "titel", "schluessel", "beschreibung")
        tree = ttk.Treeview(wrap, style="P.Treeview", columns=cols, show="headings", selectmode="browse")
        for c, t, w in (("modul", "Modul", 130), ("titel", "Titel", 220),
                        ("schluessel", "Schlüssel", 200), ("beschreibung", "Beschreibung", 380)):
            tree.heading(c, text=t)
            tree.column(c, width=w, anchor="w")
        for b in self.berecht:
            tree.insert("", "end", iid=str(b["id"]),
                        values=(b["modul"], b["titel"], b["schluessel"], b["beschreibung"]))
        vsb = ttk.Scrollbar(wrap, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=vsb.set)
        tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        wrap.rowconfigure(0, weight=1)
        wrap.columnconfigure(0, weight=1)
        self.katalog_tree = tree

        def on_dbl(_ev):
            sel = tree.selection()
            if sel and self.admin:
                self._katalog_form(self.b_by_id.get(int(sel[0])))
            elif sel:
                self._require_admin()
        tree.bind("<Double-Button-1>", on_dbl)
        self.status.set(f"Katalog · {len(self.berecht)} Berechtigungs-Punkte in {len(self.module)} Modulen"
                        + ("  ·  Doppelklick zum Bearbeiten" if self.admin else ""))

    def _katalog_form(self, b):
        if not self._require_admin():
            return
        win = tk.Toplevel(self.root)
        win.title("Berechtigung bearbeiten" if b else "Neue Berechtigung")
        win.configure(bg=BG)
        win.transient(self.root)
        win.grab_set()
        win.geometry("480x360")
        fields = [("modul", "Modul"), ("titel", "Titel"),
                  ("schluessel", "Schlüssel (technisch, eindeutig)"), ("beschreibung", "Beschreibung")]
        vars_ = {}
        for i, (k, lab) in enumerate(fields):
            tk.Label(win, text=lab, bg=BG, fg=PRIMARY, font=(FONT, 10, "bold")).grid(
                row=i, column=0, sticky="w", padx=16, pady=9)
            v = tk.StringVar(value=(b.get(k, "") if b else ""))
            vars_[k] = v
            tk.Entry(win, textvariable=v, width=34).grid(row=i, column=1, sticky="ew", padx=16, pady=9)
        win.columnconfigure(1, weight=1)
        tk.Label(win, text="Der Schlüssel verbindet die Berechtigung mit der Programm-Logik\n"
                           "(z. B. kasse.verkauf). Nur ändern, wenn du weißt was du tust.",
                 bg=BG, fg=MUTED, font=(FONT, 9), justify="left").grid(
            row=len(fields), column=0, columnspan=2, sticky="w", padx=16, pady=(2, 8))

        def save():
            data = {k: v.get().strip() for k, v in vars_.items()}
            if not data["titel"] or not data["schluessel"]:
                messagebox.showwarning("Eingabe", "Titel und Schlüssel sind Pflicht.", parent=win)
                return
            con = sqlite3.connect(DB_PATH)
            try:
                if b:
                    con.execute("UPDATE tbl_berechtigung SET modul=?,titel=?,schluessel=?,beschreibung=? WHERE id=?",
                                (data["modul"] or "Allgemein", data["titel"], data["schluessel"],
                                 data["beschreibung"], b["id"]))
                    act = f"'{data['titel']}' geändert"
                else:
                    mx = con.execute("SELECT COALESCE(MAX(sort),0)+1 FROM tbl_berechtigung").fetchone()[0]
                    con.execute("INSERT INTO tbl_berechtigung(modul,titel,schluessel,beschreibung,sort) "
                                "VALUES(?,?,?,?,?)", (data["modul"] or "Allgemein", data["titel"],
                                                      data["schluessel"], data["beschreibung"], mx))
                    act = f"'{data['titel']}' angelegt"
                con.commit()
            except sqlite3.IntegrityError:
                messagebox.showwarning("Speichern", "Dieser Schlüssel ist bereits vergeben.", parent=win)
                con.close()
                return
            con.close()
            _audit("katalog", act)
            win.destroy()
            self.load()
            self.show_katalog()

        bar = tk.Frame(win, bg=BG)
        bar.grid(row=len(fields) + 1, column=0, columnspan=2, sticky="ew", padx=16, pady=14)
        if b:
            tk.Button(bar, text="🗑 Löschen", command=lambda: self._katalog_loeschen(b, win),
                      bg="#EDF1F6", fg=DANGER, relief="flat", padx=12, pady=7, cursor="hand2").pack(side="left")
        tk.Button(bar, text="Abbrechen", command=win.destroy, padx=14, pady=7).pack(side="right", padx=(8, 0))
        tk.Button(bar, text="Speichern", command=save, bg=PRIMARY, fg="white", relief="flat",
                  padx=18, pady=8, cursor="hand2").pack(side="right")

    def _katalog_loeschen(self, b, win):
        if not messagebox.askyesno("Löschen",
                f"Berechtigung '{b['titel']}' löschen?\n\n"
                "Alle Rollen-Zuordnungen und persönlichen Ausnahmen dazu werden mit entfernt.",
                parent=win):
            return
        con = sqlite3.connect(DB_PATH)
        con.execute("DELETE FROM tbl_berechtigung WHERE id=?", (b["id"],))
        con.execute("DELETE FROM tbl_rolle_recht WHERE berechtigung_id=?", (b["id"],))
        con.execute("DELETE FROM tbl_ma_recht_override WHERE berechtigung_id=?", (b["id"],))
        con.commit()
        con.close()
        _audit("katalog_geloescht", f"'{b['titel']}' gelöscht")
        win.destroy()
        self.load()
        self.show_katalog()

    # =========================================================================
    #  ANSICHT · EINSTELLUNGEN (App-übergreifende Schalter)
    # =========================================================================
    def _get_config(self, key, default=""):
        con = sqlite3.connect(DB_PATH)
        row = con.execute("SELECT wert FROM tbl_param_config WHERE schluessel=?", (key,)).fetchone()
        con.close()
        return row[0] if row else default

    def _set_config(self, key, wert):
        con = sqlite3.connect(DB_PATH)
        con.execute("INSERT INTO tbl_param_config(schluessel,wert) VALUES(?,?) "
                    "ON CONFLICT(schluessel) DO UPDATE SET wert=excluded.wert", (key, str(wert)))
        con.commit()
        con.close()

    def show_einstellungen(self):
        self.view = "einst"
        self._nav_state()
        self._clear_main()
        tk.Label(self.main, text="Einstellungen", bg=BG, fg=PRIMARY,
                 font=(FONT, 13, "bold")).pack(anchor="w", padx=18, pady=(14, 2))
        tk.Label(self.main, text="App-übergreifende Schalter. Änderungen brauchen den Admin-Modus.",
                 bg=BG, fg=MUTED, font=(FONT, 10)).pack(anchor="w", padx=18)

        card = tk.Frame(self.main, bg=CARD, highlightbackground=BORDER, highlightthickness=1)
        card.pack(fill="x", padx=18, pady=14)
        tk.Label(card, text="🏖  Urlaub / Abwesenheiten – Genehmigung", bg=CARD, fg=INK,
                 font=(FONT, 11, "bold")).pack(anchor="w", padx=16, pady=(14, 2))
        tk.Label(card, text="Anträge werden zunächst von den Personalverantwortlichen genehmigt oder "
                            "abgelehnt (1. Stufe). Ist die folgende Option aktiv, muss ein genehmigter "
                            "Antrag zusätzlich von der Geschäftsführung freigegeben werden (2. Stufe).",
                 bg=CARD, fg=MUTED, font=(FONT, 9), wraplength=620, justify="left").pack(anchor="w", padx=16, pady=(0, 8))

        self.gf_var = tk.IntVar(value=1 if str(self._get_config("urlaub_gf_freigabe", "0")) == "1" else 0)

        def toggle():
            if not self._require_admin():
                self.gf_var.set(1 - self.gf_var.get())  # zurücksetzen
                return
            self._set_config("urlaub_gf_freigabe", self.gf_var.get())
            _audit("urlaub_gf_freigabe", "AN" if self.gf_var.get() else "AUS")
            self.status.set("Geschäftsführungs-Freigabe " + ("aktiviert." if self.gf_var.get() else "deaktiviert."))

        tk.Checkbutton(card, text="Geschäftsführung muss genehmigte Urlaubsanträge zusätzlich freigeben",
                       variable=self.gf_var, command=toggle, bg=CARD, fg=INK, font=(FONT, 10),
                       activebackground=CARD, selectcolor="white", anchor="w",
                       cursor="hand2").pack(anchor="w", padx=16, pady=(0, 14))

        self.status.set("Einstellungen")

    # =========================================================================
    #  ANSICHT 5 · PROTOKOLL
    # =========================================================================
    def show_audit(self):
        self.view = "audit"
        self._nav_state()
        self._clear_main()
        tk.Label(self.main, text="Änderungsprotokoll", bg=BG, fg=PRIMARY,
                 font=(FONT, 13, "bold")).pack(anchor="w", padx=18, pady=(14, 2))
        tk.Label(self.main, text="Jede Freischaltung, Sperrung und Rollen-Änderung wird hier "
                                 "revisionssicher festgehalten.", bg=BG, fg=MUTED,
                 font=(FONT, 10)).pack(anchor="w", padx=18)

        wrap = tk.Frame(self.main, bg=BG)
        wrap.pack(fill="both", expand=True, padx=18, pady=10)
        cols = ("zeit", "benutzer", "aktion", "detail")
        tree = ttk.Treeview(wrap, style="P.Treeview", columns=cols, show="headings")
        for c, t, w in (("zeit", "Zeitpunkt", 150), ("benutzer", "Benutzer", 130),
                        ("aktion", "Aktion", 150), ("detail", "Detail", 520)):
            tree.heading(c, text=t)
            tree.column(c, width=w, anchor="w")
        con = sqlite3.connect(DB_PATH)
        rows = con.execute("SELECT zeitpunkt,benutzer,aktion,detail FROM tbl_param_audit "
                           "ORDER BY id DESC LIMIT 500").fetchall()
        con.close()
        for z, u, a, d in rows:
            try:
                z = datetime.fromisoformat(z).strftime("%d.%m.%Y %H:%M")
            except Exception:
                pass
            tree.insert("", "end", values=(z, u, a, d))
        vsb = ttk.Scrollbar(wrap, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=vsb.set)
        tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        wrap.rowconfigure(0, weight=1)
        wrap.columnconfigure(0, weight=1)
        self.status.set(f"Protokoll · {len(rows)} jüngste Einträge")

    # ---- gemeinsame Helfer ------------------------------------------------
    def _empty_hint(self, parent, title, text):
        box = tk.Frame(parent, bg=BG)
        box.place(relx=0.5, rely=0.4, anchor="center")
        tk.Label(box, text=title, bg=BG, fg=INK, font=(FONT, 14, "bold")).pack(pady=(0, 6))
        if text:
            tk.Label(box, text=text, bg=BG, fg=MUTED, font=(FONT, 10), justify="center").pack()


def run_standalone():
    init_db()
    root = tk.Tk()
    App(root)
    root.mainloop()


# ──────────────────────────────────────────────────────────────────────────────
#  TESTUMGEBUNG – isolierte DB, Demo-Mitarbeiter, Admin sofort offen
# ──────────────────────────────────────────────────────────────────────────────
DEMO_MITARBEITER = [
    # (vorname, name, abteilung, position, rollen-name, [extra-allow], [extra-deny])
    ("Anna",  "Maier",  "Geschäftsführung", "Geschäftsführung", "Geschäftsführung", [], []),
    ("Ben",   "Krause", "Vertrieb",         "Außendienst",      "Vertrieb / Außendienst",
        ["einkauf.margenrechner"], ["rabatte.ansehen"]),
    ("Carla", "Sommer", "Vertrieb",         "Innendienst",      "Innendienst", ["gdp.retoure"], []),
    ("David", "Reuter", "Lager",            "Logistik",         "Lager / Logistik", [], []),
    ("Eva",   "Lang",   "Labor",            "QS / Analytik",    "Labor / QS", [], []),
    ("Felix", "Thiel",  "Buchhaltung",      "Buchhalter",       "Buchhaltung / Faktura", [], []),
    ("Gina",  "Wolf",   "Einkauf",          "Einkäuferin",      "Einkauf", [], []),
    ("Hans",  "Berg",   "Qualität",         "GDP-Beauftragter", "GDP-Beauftragter", [], []),
]


def _seed_testdaten():
    """Legt Demo-Mitarbeiter inkl. Rollen-Zuordnungen und ein paar persönlichen
    Ausnahmen an – nur, wenn die Test-DB noch keine Mitarbeiter hat (damit eigene
    Experimente über Neustarts erhalten bleiben)."""
    con = sqlite3.connect(DB_PATH)
    if con.execute("SELECT COUNT(*) FROM tbl_mitarbeiter").fetchone()[0]:
        con.close()
        return
    role_id = {r[0]: r[1] for r in con.execute("SELECT name, id FROM tbl_rolle")}
    berecht_id = {r[0]: r[1] for r in con.execute("SELECT schluessel, id FROM tbl_berechtigung")}
    for vn, nn, ab, po, rolle, allow, deny in DEMO_MITARBEITER:
        cur = con.execute("INSERT INTO tbl_mitarbeiter(vorname,name,abteilung,position) "
                          "VALUES(?,?,?,?)", (vn, nn, ab, po))
        mid = cur.lastrowid
        if rolle in role_id:
            con.execute("INSERT OR IGNORE INTO tbl_ma_rolle(mitarbeiter_id,rolle_id) VALUES(?,?)",
                        (mid, role_id[rolle]))
        for key in allow:
            if key in berecht_id:
                con.execute("INSERT OR IGNORE INTO tbl_ma_recht_override(mitarbeiter_id,berechtigung_id,erlaubt) "
                            "VALUES(?,?,1)", (mid, berecht_id[key]))
        for key in deny:
            if key in berecht_id:
                con.execute("INSERT OR IGNORE INTO tbl_ma_recht_override(mitarbeiter_id,berechtigung_id,erlaubt) "
                            "VALUES(?,?,0)", (mid, berecht_id[key]))
    con.commit()
    con.close()


def run_testmode():
    """Startet die App in einer eigenen Test-Datenbank mit Demo-Daten und sofort
    geöffnetem Admin-Modus (volle Rechte, keine PIN). Die echte NMGone-Datenbank
    wird dabei NICHT berührt."""
    global DB_PATH
    from app.config import DATA_DIR
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    DB_PATH = DATA_DIR / "nmg_parameter_testumgebung.sqlite"  # isolierte Test-DB
    init_db()
    _seed_testdaten()

    root = tk.Tk()
    app = App(root)
    root.title("NMGone · Parameter & Berechtigungen — TESTUMGEBUNG (volle Rechte)")
    # Admin ohne PIN sofort aktiv schalten
    app.admin = True
    app._after_mode_change("TESTUMGEBUNG – volle Rechte aktiv. Eigene Test-Datenbank, "
                           "echte Daten bleiben unberührt.")
    root.mainloop()


def main():
    if "--test" in sys.argv or "--testmode" in sys.argv:
        run_testmode()
    else:
        run_standalone()


if __name__ == "__main__":
    sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
    main()
