# -*- coding: utf-8 -*-
"""Erzeugt die Planung "Kasse als Web/PWA" fuer NMGone als PDF.

Eine Quelle, ein PDF: Wie aus der bestehenden Desktop-Kasse eine browser- und
handytaugliche Web-Anwendung (PWA) auf dem zentralen Server wird – Architektur,
Datenmodell, Funktionsumfang v1, die Mobile-Besonderheiten (Kamera-Scan),
eine Umsetzungs-Roadmap mit Zeitschaetzung und die offenen Entscheidungen.

Scope: Kassen-/Verkaufs-Modul der NMGone-Programmfamilie als Web-App gegen die
zentrale PostgreSQL-DB (Hetzner). Pharma-Grosshandel B2B (Verkauf an Apotheken),
KEINE Bareinnahmen -> kein KassenSichV/TSE/DSFinV-K-Thema.

Aufruf:  python scripts/build_plan_kasse_web.py
"""
from __future__ import annotations

import os
from datetime import date

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.lib.enums import TA_RIGHT, TA_CENTER
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak,
    HRFlowable,
)

# --- Farben / Markenwelt ------------------------------------------------------
BLAU = colors.HexColor("#0B4A86")
TUERKIS = colors.HexColor("#0B6E6E")
GRAU = colors.HexColor("#5A6675")
HELLGRAU = colors.HexColor("#EEF1F5")
ZEILE = colors.HexColor("#F6F8FA")
GRUEN = colors.HexColor("#11823B")
ROT = colors.HexColor("#B5391F")
GELB = colors.HexColor("#C88200")
WARNBG = colors.HexColor("#FCF3E6")

OUT = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                   "docs", "Plan_Kasse_Web.pdf")

# --- Schriften (Arial fuer saubere Umlaute + Euro/Paragraf) -------------------
try:
    pdfmetrics.registerFont(TTFont("AppFont", "C:/Windows/Fonts/arial.ttf"))
    pdfmetrics.registerFont(TTFont("AppFont-Bold", "C:/Windows/Fonts/arialbd.ttf"))
    BASE, BOLD = "AppFont", "AppFont-Bold"
except Exception:
    BASE, BOLD = "Helvetica", "Helvetica-Bold"

styles = getSampleStyleSheet()


def S(name, **kw):
    base = kw.pop("parent", styles["Normal"])
    return ParagraphStyle(name, parent=base, **kw)


st_title = S("t", parent=styles["Title"], fontName=BOLD, fontSize=23, textColor=BLAU, leading=27, spaceAfter=2)
st_sub = S("s", fontName=BASE, fontSize=11.5, textColor=GRAU, spaceAfter=2)
st_h1 = S("h1", fontName=BOLD, fontSize=15, textColor=BLAU, spaceBefore=14, spaceAfter=6, leading=18)
st_h2 = S("h2", fontName=BOLD, fontSize=11.5, textColor=TUERKIS, spaceBefore=8, spaceAfter=3)
st_body = S("b", fontName=BASE, fontSize=9.5, leading=13, textColor=colors.HexColor("#23303D"))
st_small = S("sm", fontName=BASE, fontSize=8, leading=10.5, textColor=GRAU)
st_cell = S("c", fontName=BASE, fontSize=8.5, leading=11)
st_cell_b = S("cb", fontName=BOLD, fontSize=8.5, leading=11)
st_cell_w = S("cw", fontName=BOLD, fontSize=8.5, leading=11, textColor=colors.white)
st_mono = S("m", fontName=BASE, fontSize=8.5, leading=11, textColor=BLAU)
st_kpi_num = S("kn", fontName=BOLD, fontSize=18, textColor=BLAU, alignment=TA_CENTER, leading=20)
st_kpi_lbl = S("kl", fontName=BASE, fontSize=7.5, textColor=GRAU, alignment=TA_CENTER, leading=9)


def header_footer(canvas, doc):
    canvas.saveState()
    w, h = A4
    canvas.setFillColor(BLAU)
    canvas.rect(0, h - 12 * mm, w, 12 * mm, stroke=0, fill=1)
    canvas.setFillColor(colors.white)
    canvas.setFont(BOLD, 9)
    canvas.drawString(18 * mm, h - 8 * mm, "NMGone  ·  Planung Kasse als Web/PWA")
    canvas.setFont(BASE, 8)
    canvas.drawRightString(w - 18 * mm, h - 8 * mm, date.today().strftime("Stand %d.%m.%Y"))
    canvas.setStrokeColor(HELLGRAU)
    canvas.setLineWidth(0.6)
    canvas.line(18 * mm, 12 * mm, w - 18 * mm, 12 * mm)
    canvas.setFillColor(GRAU)
    canvas.setFont(BASE, 7.5)
    canvas.drawString(18 * mm, 8 * mm,
                      "Interne Projektplanung. Zeitschaetzungen sind Richtwerte, keine Festpreise.")
    canvas.drawRightString(w - 18 * mm, 8 * mm, "Seite %d" % doc.page)
    canvas.restoreState()


def kpi_row(items):
    cells = []
    for num, lbl in items:
        inner = Table([[Paragraph(num, st_kpi_num)], [Paragraph(lbl, st_kpi_lbl)]],
                      colWidths=[40 * mm])
        inner.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, -1), HELLGRAU),
            ("TOPPADDING", (0, 0), (-1, 0), 7),
            ("BOTTOMPADDING", (0, -1), (-1, -1), 7),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ]))
        cells.append(inner)
    t = Table([cells], colWidths=[44 * mm] * len(cells))
    t.setStyle(TableStyle([("LEFTPADDING", (0, 0), (-1, -1), 2),
                           ("RIGHTPADDING", (0, 0), (-1, -1), 2)]))
    return t


def table(rows, header, col_w, head_bg=BLAU):
    data = [[Paragraph(hh, st_cell_w) for hh in header]]
    for r in rows:
        line = []
        for i, c in enumerate(r):
            sty = st_mono if (i == 0 and len(header) >= 3) else st_cell
            line.append(Paragraph(str(c), sty))
        data.append(line)
    t = Table(data, colWidths=col_w, repeatRows=1)
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), head_bg),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, ZEILE]),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#D7DEE6")),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ("LEFTPADDING", (0, 0), (-1, -1), 5),
        ("RIGHTPADDING", (0, 0), (-1, -1), 5),
    ]))
    return t


def note_box(title, body_html, bg=WARNBG, edge=GELB):
    inner = [Paragraph(f"<b>{title}</b>", st_body), Spacer(1, 2),
             Paragraph(body_html, st_body)]
    t = Table([[inner]], colWidths=[174 * mm])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), bg),
        ("BOX", (0, 0), (-1, -1), 0.6, edge),
        ("LEFTPADDING", (0, 0), (-1, -1), 8),
        ("RIGHTPADDING", (0, 0), (-1, -1), 8),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
    ]))
    return t


story = []

# ============================================================ TITELSEITE
story.append(Spacer(1, 4 * mm))
story.append(Paragraph("Planung: Kasse als Web/PWA", st_title))
story.append(Paragraph("Die NMGone-Kasse im Browser &amp; auf dem Handy – zentral, mehrbenutzerfähig", st_sub))
story.append(Paragraph("Architektur · Datenmodell · Funktionsumfang · Mobile-Scan · Roadmap · offene Punkte", st_sub))
story.append(Spacer(1, 4 * mm))
story.append(HRFlowable(width="100%", thickness=1.2, color=BLAU, spaceAfter=8))

story.append(kpi_row([
    ("PWA", "Browser + Handy, 1 Code"),
    ("PostgreSQL", "zentrale DB (steht)"),
    ("FastAPI", "Stack wie Pennone"),
    ("Kamera", "PZN-/Barcode-Scan"),
]))
story.append(Spacer(1, 5 * mm))

story.append(Paragraph("Worum es geht", st_h2))
story.append(Paragraph(
    "Die heutige Kasse ist eine ausgereifte <b>Windows-Desktop-Anwendung</b> (Verkauf an Apotheken, "
    "Chargen-/Verfallverfolgung, kundenspezifische Rabatte, Vorbestellungen, Tagesabschluss). Sie soll "
    "künftig <b>im Browser und auf dem Handy</b> laufen. Der saubere Weg dafür ist <b>nicht</b>, die "
    "Desktop-Version zu portieren, sondern die Kasse als <b>Web-App</b> auf dem bereits aufgesetzten "
    "Hetzner-Server neu zu bauen – gegen die <b>zentrale PostgreSQL-Datenbank</b>, die dort schon läuft. "
    "Damit ist sie von Anfang an zentral und mehrbenutzerfähig.", st_body))
story.append(Spacer(1, 3 * mm))

story.append(note_box(
    "„Browser oder Handy-App“ ist in Wirklichkeit EINE Lösung",
    "Es braucht <b>keine getrennten iOS-/Android-Apps</b>. Eine <b>PWA</b> (Progressive Web App) läuft im "
    "Browser auf PC, Tablet und Handy und lässt sich auf dem Handy „zum Startbildschirm hinzufügen“ – "
    "eigenes Icon, Vollbild, fühlt sich an wie eine echte App. <b>Ein Code, alle Geräte</b>, kein App-Store. "
    "Pennone One nutzt dieses Muster bereits; die Kasse setzt darauf auf.",
    bg=colors.HexColor("#EAF2FB"), edge=BLAU))
story.append(Spacer(1, 4 * mm))

story.append(note_box(
    "Aktualisierung: getroffene Entscheidungen &amp; erstes Geruest (P0) steht",
    "<b>1. Zentral &amp; mehrbenutzerfaehig</b> – bestaetigt. "
    "<b>2. Lokal-zuerst + Sync von Anfang an</b>: die Kasse arbeitet gegen eine lokale DB und "
    "gleicht im Hintergrund mit der zentralen DB ab – bei Netzausfall wird weiterverkauft "
    "(bewusst NICHT online-only). <b>3. Eigenes Desktop-Fenster</b> (pywebview, kein Browser) "
    "<b>UND Handy/Browser von Anfang an</b> – dieselbe FastAPI-App, am PC als Fenster, am Handy als PWA. "
    "<b>P0 ist gebaut und getestet</b>: Login, Desktop-Fenster, PWA-Manifest/Service-Worker, "
    "Lizenz-Gate und die Kassen-Startseite laufen (Rauchtest gruen).",
    bg=colors.HexColor("#EAF7EE"), edge=GRUEN))
story.append(Spacer(1, 4 * mm))

story.append(Paragraph("Inhalt dieser Planung", st_h2))
story.append(Paragraph(
    "1. Architektur &amp; Einordnung in die Server-Landschaft<br/>"
    "2. Datenmodell – was aus der Desktop-Kasse übernommen wird<br/>"
    "3. Funktionsumfang: v1 (MVP) und später<br/>"
    "4. Mobile-Besonderheiten – Kamera-Scan, PWA, Belege<br/>"
    "5. Umsetzungs-Roadmap mit Zeitschätzung<br/>"
    "6. Offene Entscheidungen", st_body))

story.append(PageBreak())

# ============================================================ 1. ARCHITEKTUR
story.append(Paragraph("1 · Architektur &amp; Einordnung", st_h1))
story.append(Paragraph(
    "Die Kasse wird ein <b>eigener Web-Dienst</b> auf demselben Hetzner-Server – gleiches Muster wie "
    "Pennone One: ein <b>FastAPI</b>-Prozess hinter dem <b>Caddy</b>-Reverse-Proxy (automatisches HTTPS), "
    "als <b>systemd</b>-Dienst, der die zentrale <b>PostgreSQL</b>-DB nutzt.", st_body))
story.append(Spacer(1, 3 * mm))
arch = [
    ("Frontend/UI", "Server-gerendert mit Jinja2 + HTMX, touch-optimiert; als PWA installierbar"),
    ("Backend", "FastAPI (Python), wie Pennone – ein Stack, geteilte Infrastruktur"),
    ("Datenbank", "Zentrale PostgreSQL 18 auf dem Server (läuft bereits)"),
    ("Betrieb", "Eigener systemd-Dienst + eigener Caddy-Block (Auto-HTTPS)"),
    ("Adresse", "z.B. kasse.pennone.de (Subdomain) – als App aufs Handy installierbar"),
    ("Sicherheit", "Login + Rollen; DB nur server-intern, kein direkter Client-DB-Zugriff"),
]
story.append(table(arch, ["Baustein", "Umsetzung"], [40 * mm, 134 * mm]))
story.append(Spacer(1, 3 * mm))

story.append(Paragraph("Wie sich die Kasse in die Gesamtlandschaft einfügt", st_h2))
land = [
    ("Pennone One", "Server", "Web-App / PWA – läuft (one.pennone.de)"),
    ("Kasse", "Server", "Web-App / PWA – dieses Vorhaben, mobil-tauglich"),
    ("PostgreSQL", "Server", "Zentrale DB für alles – steht bereits"),
    ("NMGone (Analyse-Kern)", "bleibt Desktop", "nutzt später dieselbe zentrale DB"),
]
story.append(table(land, ["Komponente", "Wo", "Form / Stand"], [50 * mm, 32 * mm, 92 * mm], head_bg=TUERKIS))
story.append(Spacer(1, 3 * mm))
story.append(note_box(
    "Warum das auch das NMGone-Dilemma löst",
    "Der NMGone-Analyse-Kern steckt an <b>576 Stellen in 43 Dateien</b> direkt auf SQLite. Statt alles "
    "umzubauen, werden nur die Teile, die <b>mobil/Browser</b> sein müssen (Kasse, später ein Cockpit), "
    "web-first neu gebaut – der schwere Desktop-Kern bleibt und teilt sich nur die zentrale Datenbank.",
    bg=colors.HexColor("#EAF2FB"), edge=BLAU))

story.append(PageBreak())

# ============================================================ 2. DATENMODELL
story.append(Paragraph("2 · Datenmodell", st_h1))
story.append(Paragraph(
    "Das bestehende Kassen-Schema wird <b>1:1 nach PostgreSQL portiert</b> (gleiche Logik, andere "
    "Engine). Stammdaten (Artikel, Kunden, Rabatte) kommen aus der <b>zentralen DB</b>, die mit NMGone "
    "geteilt wird – keine Doppelpflege.", st_body))
story.append(Spacer(1, 3 * mm))
tabellen = [
    ("tbl_bestellungen", "Verkaufs-/Auftragsköpfe (Kunde, Datum, Liefertermin, Summe, Status)"),
    ("tbl_bestellpositionen", "Positionen je Auftrag (PZN, Menge, Preis, Rabatt, Bestellart)"),
    ("tbl_wareneingang", "Wareneingangs-Köpfe (Lieferant, Datum, Beleg)"),
    ("tbl_wareneingang_positionen", "Eingangspositionen (PZN, Menge, Charge, Verfall)"),
    ("tbl_lagerbestand", "Bestand je PZN/Charge/Verfall – Kern der GDP-Rückverfolgung"),
    ("tbl_kasse_log", "Aktionsprotokoll (wer/wann/was) – Nachvollziehbarkeit"),
    ("tbl_kasse_tagesabschluss", "Tagesabschlüsse / Z-Werte"),
    ("tbl_kasse_einstellungen", "Kassen-Einstellungen (pro Standort/Firma)"),
]
story.append(table(tabellen, ["Tabelle", "Inhalt"], [58 * mm, 116 * mm]))
story.append(Spacer(1, 3 * mm))
story.append(Paragraph(
    "<b>Geteilte Stammdaten</b> (zentral, read-only für die Kasse): Artikel/PZN, Kundenstamm, "
    "kundenspezifische Rabatte. <b>Mehrmandanten/Standort:</b> alle Kassentabellen bekommen eine "
    "Firmen-/Standort-Dimension, damit mehrere Nutzer und ggf. mehrere Standorte sauber getrennt sind.",
    st_small))

story.append(PageBreak())

# ============================================================ 3. FUNKTIONSUMFANG
story.append(Paragraph("3 · Funktionsumfang", st_h1))
story.append(Paragraph(
    "v1 ist die schlanke, voll verkaufsfähige Kasse. Alles Weitere wird modular nachgezogen, ohne "
    "den Kern umzubauen.", st_body))
story.append(Spacer(1, 3 * mm))

story.append(Paragraph("3a · v1 (MVP) – das, was eine Kasse zum Verkaufen braucht", st_h2))
v1 = [
    ("Verkauf", "Kunde suchen → Artikel (PZN) suchen → Charge/Verfall wählen → Rabatt automatisch → Position → Abschluss"),
    ("Freie Position", "Position ohne Stammartikel (manuell) erfassen"),
    ("Lagerbestand", "Bestand ansehen, Bestand je Charge, Verfügbarkeit prüfen"),
    ("Beleg / Auftrag", "Als PDF erzeugen (zum Teilen/Drucken) statt Drucker-Treiber"),
    ("Übersicht", "Dashboard mit Kennzahlen + letzte Vorgänge"),
    ("Tagesabschluss", "Tageswerte abschließen und festschreiben"),
]
story.append(table(v1, ["Funktion", "Inhalt"], [40 * mm, 134 * mm]))
story.append(Spacer(1, 4 * mm))

story.append(Paragraph("3b · Später (nach v1)", st_h2))
spaeter = [
    ("Wareneingang mobil", "Eingang per Handy erfassen, mit Kamera-Scan – füllt das Lager"),
    ("Vorbestellungen", "Offene Vorbestellungen je Kunde, Auftragsumwandlung"),
    ("Kunden-Insights", "Top-Artikel je Kunde, Verkaufshistorie"),
    ("MSV3-Import", "Bestellungen aus dem MSV3-Postfach automatisch in den Verkauf"),
    ("Statistiken", "Auswertungen, Export"),
    ("Offline-Betrieb", "Weiterverkaufen ohne Internet + späterer Sync (siehe Abschnitt 6)"),
]
story.append(table(spaeter, ["Funktion", "Inhalt"], [40 * mm, 134 * mm], head_bg=TUERKIS))
story.append(Spacer(1, 3 * mm))
story.append(note_box(
    "Kein Bargeld – kein TSE/KassenSichV",
    "Der Verkauf läuft <b>B2B unbar</b> (an Apotheken). Es gibt keine Bareinnahmen, daher entfällt das "
    "gesamte Thema <b>KassenSichV / TSE / DSFinV-K</b>. Das hält die Kasse technisch und rechtlich "
    "deutlich schlanker als eine Ladenkasse.",
    bg=colors.HexColor("#EAF7EE"), edge=GRUEN))

story.append(PageBreak())

# ============================================================ 4. MOBILE
story.append(Paragraph("4 · Mobile-Besonderheiten", st_h1))
story.append(Paragraph(
    "Der größte Gewinn gegenüber der Desktop-Kasse ist die <b>Handy-Nutzung</b> – vor allem das "
    "Scannen mit der Kamera.", st_body))
story.append(Spacer(1, 3 * mm))
mobile = [
    ("PZN-/Barcode-Scan", "Mit der Handykamera Artikel scannen (Web-BarcodeDetector bzw. JS-Scanner) statt tippen – schnell im Lager &amp; am Verkauf"),
    ("PWA-Installation", "„Zum Startbildschirm hinzufügen“: eigenes Icon, Vollbild, App-Gefühl ohne App-Store"),
    ("Touch-UI", "Große Schaltflächen, wenige Klicks, fingerfreundliche Listen"),
    ("Beleg teilen", "PDF erzeugen und per Handy teilen/drucken – kein Treiber nötig"),
    ("Responsiv", "Gleiche App passt sich Handy, Tablet und PC an"),
]
story.append(table(mobile, ["Feature", "Nutzen"], [44 * mm, 130 * mm]))
story.append(Spacer(1, 3 * mm))
story.append(note_box(
    "Offline – entschieden: lokal-zuerst + Sync (von Anfang an)",
    "Eine Kasse darf bei Netzausfall <b>nie stehenbleiben</b>. Daher arbeitet sie am PC gegen eine "
    "<b>lokale Datenbank</b> und gleicht im Hintergrund mit der zentralen DB ab; Handys laufen online "
    "gegen zentral. Das ist <b>mehr Aufwand</b> (Sync- und Konfliktlogik), wurde aber bewusst <b>nach vorn "
    "geholt</b> – nicht als spätere Ausbaustufe. Am PC zeigt ein <b>pywebview-Fenster</b> die lokale "
    "App (kein Browser), am Handy dieselbe App als PWA mit Kamera-Scan.",
    bg=colors.HexColor("#EAF7EE"), edge=GRUEN))

story.append(PageBreak())

# ============================================================ 5. ROADMAP
story.append(Paragraph("5 · Umsetzungs-Roadmap mit Zeitschätzung", st_h1))
story.append(Paragraph(
    "Zeitschätzung in <b>Personentagen (PT)</b> inkl. Test. Reihenfolge = empfohlene Bearbeitung. "
    "Das <b>Herzstück ist der Verkaufs-Flow</b> (P2).", st_body))
story.append(Spacer(1, 3 * mm))
roadmap = [
    ("P0 ✓", "Gerüst (ERLEDIGT)", "FastAPI-Skelett, Login, PWA-Manifest/Service-Worker, Lizenz-Gate, pywebview-Fenster. Gebaut + getestet.", "fertig"),
    ("P1", "Postgres-Schema &amp; Stammdaten", "8 Kassentabellen nach Postgres; Artikel-/Kunden-/Rabatt-Zugriff aus zentraler DB.", "3–5"),
    ("P2", "Verkaufs-Flow (mobil)", "Kunde → Artikel → Charge/Verfall → Rabatt → Position → Abschluss. Das Herzstück.", "6–10"),
    ("P2", "Kamera-Scan", "PZN/Barcode per Handykamera in Suche &amp; Position.", "2–4"),
    ("P3", "Beleg-PDF &amp; Tagesabschluss", "Auftrags-/Beleg-PDF; Tagesabschluss festschreiben.", "3–5"),
    ("P3", "Lagerbestand-Ansicht", "Bestand je PZN/Charge, Verfügbarkeit.", "2–3"),
    ("P3", "Übersicht / Dashboard", "Kennzahlen + letzte Vorgänge.", "1–2"),
    ("P4", "Deploy &amp; Mehrbenutzer", "systemd-Dienst + Caddy + Subdomain; Rollen/Rechte, Standort-Trennung.", "2–4"),
    ("P4", "Hilfe &amp; Handout", "Kurzanleitung + Screenshots für die App-Familie.", "1–2"),
    ("–", "Später: Wareneingang mobil", "Eingang per Handy + Scan erfassen.", "3–5"),
    ("–", "Später: Vorbestellungen/MSV3", "Vorbestellungen, MSV3-Import, Statistiken.", "4–7"),
    ("–", "Später: Offline-Sync", "Service-Worker + Warteschlange + Abgleich.", "5–9"),
]
story.append(table(roadmap, ["Phase", "Vorhaben", "Inhalt", "PT"],
                   [14 * mm, 46 * mm, 98 * mm, 16 * mm]))
story.append(Spacer(1, 3 * mm))
story.append(Paragraph(
    "<b>Kern bis lauffähige, deployte Kasse (P0–P4): rund 22–38 Personentage.</b> Größter Block ist "
    "der Verkaufs-Flow (6–10 PT). Die „Später“-Stufen kommen modular oben drauf, ohne den Kern "
    "umzubauen.", st_body))
story.append(Spacer(1, 3 * mm))
story.append(Paragraph(
    "Lesehilfe: <b>klein</b> ≈ 1–2 PT · <b>mittel</b> ≈ 3–8 PT · <b>gross</b> ≈ 10+ PT. "
    "Ein Personentag = ein voller Arbeitstag konzentrierter Entwicklung inkl. Test.", st_small))

story.append(PageBreak())

# ============================================================ 6. OFFENE PUNKTE
story.append(Paragraph("6 · Offene Entscheidungen", st_h1))
story.append(Paragraph(
    "Diese Punkte bestimmen Architektur und Aufwand und sollten vor dem Bau des Gerüsts (P0) "
    "festgelegt werden.", st_body))
story.append(Spacer(1, 3 * mm))
offen = [
    ("Offline", "<b>ENTSCHIEDEN:</b> lokal-zuerst + Sync von Anfang an (PC lokal, Handy online)."),
    ("Geräte", "<b>ENTSCHIEDEN:</b> Desktop-Fenster (pywebview) UND Handy/Browser-PWA von Anfang an."),
    ("Login / Nutzer", "Eigener Kassen-Login – oder am zentralen Cockpit-/NMGone-Konto hängen?"),
    ("Drucken", "Reicht Beleg-PDF zum Teilen/Drucken – oder echter Bondrucker (Bluetooth/Netzwerk)?"),
    ("Adresse", "Subdomain kasse.pennone.de – oder eigene Firmendomain?"),
    ("Stammdaten-Quelle", "Artikel/Kunden/Rabatte aus der zentralen DB – wann werden sie dorthin migriert?"),
    ("Mehrbenutzer/Rollen", "Wer darf verkaufen, stornieren, Tagesabschluss machen? (über die Rechte-App)"),
    ("Standorte", "Ein Standort oder mehrere (getrennte Bestände/Kassen)?"),
]
story.append(table(offen, ["Thema", "Frage / Stand"], [44 * mm, 130 * mm], head_bg=ROT))
story.append(Spacer(1, 5 * mm))

story.append(HRFlowable(width="100%", thickness=0.8, color=HELLGRAU, spaceAfter=4))
story.append(Paragraph(
    "<b>Hinweis:</b> Diese Planung ist eine interne Projektunterlage zur Abstimmung. Zeitschätzungen "
    "sind Richtwerte (kein Festpreis) und hängen von den Entscheidungen in Abschnitt 6 ab. "
    "Stand: " + date.today().strftime("%d.%m.%Y") + ".",
    st_small))

# --- Build --------------------------------------------------------------------
os.makedirs(os.path.dirname(OUT), exist_ok=True)
doc = SimpleDocTemplate(
    OUT, pagesize=A4,
    leftMargin=18 * mm, rightMargin=18 * mm,
    topMargin=18 * mm, bottomMargin=16 * mm,
    title="Planung Kasse als Web/PWA NMGone", author="NMGone")
doc.build(story, onFirstPage=header_footer, onLaterPages=header_footer)
print("OK ->", OUT)
