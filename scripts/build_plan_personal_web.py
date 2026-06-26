# -*- coding: utf-8 -*-
"""Erzeugt die Planung "Personal-Modul als Web-Funktion" als PDF.

Eine Quelle, ein PDF: Was das Web-Personalmodul ist, was noch fehlt, die
kompletten Hosting-Kosten (Domain bis Webhoster, realistische DE-Preise),
ein Kundenpreis-Modell und die Deckungsbeitrags-Rechnung.

Architektur-Bezug: FastAPI + HTMX, eine SQLite PRO Firma (Mandant), Lizenz-Gate
ueber platform.sqlite. Hosting bewusst leichtgewichtig (ein VPS traegt
zweistellige Mandantenzahlen).

Aufruf:  python scripts/build_plan_personal_web.py
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
                   "docs", "Plan_Personal_Web.pdf")

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
st_cell_w = S("cw", fontName=BOLD, fontSize=8.5, leading=11, textColor=colors.white)
st_mono = S("m", fontName=BASE, fontSize=8.5, leading=11, textColor=BLAU)
st_kpi_num = S("kn", fontName=BOLD, fontSize=19, textColor=BLAU, alignment=TA_CENTER, leading=21)
st_kpi_lbl = S("kl", fontName=BASE, fontSize=7.5, textColor=GRAU, alignment=TA_CENTER, leading=9)
st_right = S("r", fontName=BASE, fontSize=8.5, leading=11, alignment=TA_RIGHT)
st_right_b = S("rb", fontName=BOLD, fontSize=8.5, leading=11, alignment=TA_RIGHT, textColor=BLAU)


def header_footer(canvas, doc):
    canvas.saveState()
    w, h = A4
    canvas.setFillColor(BLAU)
    canvas.rect(0, h - 12 * mm, w, 12 * mm, stroke=0, fill=1)
    canvas.setFillColor(colors.white)
    canvas.setFont(BOLD, 9)
    canvas.drawString(18 * mm, h - 8 * mm, "NMGone  ·  Planung Personal-Modul (Web)")
    canvas.setFont(BASE, 8)
    canvas.drawRightString(w - 18 * mm, h - 8 * mm, date.today().strftime("Stand %d.%m.%Y"))
    canvas.setStrokeColor(HELLGRAU)
    canvas.setLineWidth(0.6)
    canvas.line(18 * mm, 12 * mm, w - 18 * mm, 12 * mm)
    canvas.setFillColor(GRAU)
    canvas.setFont(BASE, 7.5)
    canvas.drawString(18 * mm, 8 * mm,
                      "Interne Projekt- und Kostenplanung. Preise sind Marktstaende (DE) und koennen sich aendern – vor Abschluss pruefen.")
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


def table(rows, header, col_w, head_bg=BLAU, right_cols=()):
    data = [[Paragraph(hh, st_cell_w) for hh in header]]
    for r in rows:
        line = []
        for i, c in enumerate(r):
            if i in right_cols:
                sty = st_right_b if i == len(header) - 1 else st_right
            elif i == 0 and len(header) >= 3:
                sty = st_mono
            else:
                sty = st_cell
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


def total_row_style(t, n_rows, n_cols):
    """Hebt die letzte Tabellenzeile als Summenzeile hervor."""
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, n_rows), (-1, n_rows), colors.HexColor("#EAF2FB")),
        ("LINEABOVE", (0, n_rows), (-1, n_rows), 0.8, BLAU),
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
story.append(Paragraph("Planung: Personal-Modul als Web-Funktion", st_title))
story.append(Paragraph("Mitarbeiterverwaltung im Browser – mandantenfaehig, DSGVO-konform gehostet", st_sub))
story.append(Paragraph("Was steht · Hosting-Kosten (Domain bis Webhoster) · Kundenpreise · Deckungsbeitrag", st_sub))
story.append(Spacer(1, 4 * mm))
story.append(HRFlowable(width="100%", thickness=1.2, color=BLAU, spaceAfter=8))

story.append(kpi_row([
    ("FastAPI", "Web-Stack + HTMX"),
    ("1 DB / Firma", "Mandantentrennung"),
    ("< 10 €", "Hosting/Monat (Start)"),
    ("19–79 €", "Kundenpreis/Monat"),
]))
story.append(Spacer(1, 5 * mm))

story.append(Paragraph("Worum es geht", st_h2))
story.append(Paragraph(
    "Das <b>Personal-Modul</b> aus der NMGone-Familie soll als <b>Web-Funktion</b> bereitstehen: "
    "Mitarbeiterstammdaten, Custom-Felder (EAV), Vorgesetzten-Matrix und Organigramm – erreichbar "
    "im Browser, ohne lokale Installation. Technisch laeuft das ueber eine <b>FastAPI</b>-Anwendung "
    "mit <b>HTMX</b>-Oberflaeche; jede Kundenfirma erhaelt eine <b>eigene SQLite-Datenbank</b> "
    "(saubere Mandantentrennung), der Zugang ist ueber ein <b>Lizenz-Gate</b> geschuetzt. Der Pilot "
    "ist lauffaehig und getestet.", st_body))
story.append(Spacer(1, 3 * mm))

story.append(note_box(
    "Warum das Modell wirtschaftlich attraktiv ist",
    "Die Architektur ist bewusst <b>leichtgewichtig</b>: eine App-Instanz, eine kleine SQLite-Datei "
    "je Firma. Damit traegt <b>ein einzelner kleiner Server</b> problemlos zweistellige Mandantenzahlen. "
    "Die Hosting-<b>Grenzkosten pro Kunde liegen unter 1 €/Monat</b> – der Server ist der einzige "
    "nennenswerte Fixposten. Der eigentliche Wert (und Preis) steckt in Support, DSGVO-Verantwortung "
    "und Weiterentwicklung, nicht in der Infrastruktur.",
    bg=colors.HexColor("#EAF2FB"), edge=BLAU))
story.append(Spacer(1, 4 * mm))

story.append(Paragraph("Inhalt dieser Planung", st_h2))
story.append(Paragraph(
    "1. Was steht – und was bis zum ersten zahlenden Kunden noch fehlt<br/>"
    "2. Hosting-Kosten Stufe 1 (Start) – Domain bis Webhoster, alle Posten<br/>"
    "3. Hosting-Kosten Stufe 2 (Wachstum) und die Managed-Alternative<br/>"
    "4. Kundenpreise – zwei Tarifmodelle<br/>"
    "5. Deckungsbeitrag – was am Ende uebrig bleibt<br/>"
    "6. Empfehlung &amp; naechste Schritte", st_body))

story.append(PageBreak())

# ============================================================ 1. STAND
story.append(Paragraph("1 · Was steht – und was noch fehlt", st_h1))
story.append(Paragraph(
    "Der Web-Pilot des Personal-Moduls laeuft mehrbenutzerfaehig mit Mandantentrennung und "
    "Lizenz-Gate. Funktional ist volle Paritaet zur Desktop-App nachgewiesen. Fuer den Schritt "
    "vom Pilot zum <b>verkaufsfertigen Produkt</b> fehlen noch ein paar Bausteine – vor allem "
    "das DSGVO-Paket, weil es sich um <b>Personaldaten</b> handelt.", st_body))
story.append(Spacer(1, 3 * mm))
stand = [
    ("Personal-Kernfunktionen", "Stammdaten, Custom-Felder, Vorgesetzten-Matrix, Organigramm", "steht (Pilot)", "–"),
    ("Mandantentrennung", "Eine SQLite je Firma, Lizenz-Gate aktiv", "steht", "–"),
    ("Self-Service-Onboarding", "Firma selbst anlegen, ohne manuelles Zutun", "offen", "mittel"),
    ("Passwort-Reset / E-Mail", "Einladung &amp; Reset per Transaktions-Mail", "offen", "klein"),
    ("DSGVO-Paket", "AV-Vertrag, Loeschkonzept, Datenexport – Pflicht bei Personaldaten", "offen", "mittel"),
    ("Abrechnung / Zahlung", "Stripe-Anbindung, Tarif-Verwaltung", "offen", "mittel"),
    ("Backups + Restore-Test", "Automatisiert je Mandant, Wiederherstellung getestet", "offen", "klein"),
    ("Rollen/Rechte im Web", "Analog Parameter-App: wer darf was", "teils", "mittel"),
]
story.append(table(stand, ["Baustein", "Inhalt", "Status", "Aufwand"],
                   [42 * mm, 86 * mm, 28 * mm, 18 * mm]))
story.append(Spacer(1, 3 * mm))
story.append(note_box(
    "DSGVO ist kein Nice-to-have",
    "Sobald du fremde <b>Personaldaten</b> verarbeitest, bist du <b>Auftragsverarbeiter</b>. "
    "Vor dem ersten zahlenden Kunden zwingend: <b>AV-Vertrag</b> (Vorlage genuegt), ein "
    "<b>Loeschkonzept</b> und ein <b>Datenexport</b> je Mandant. Deutsches Hosting (siehe Abschnitt 2) "
    "ist dabei ein echtes Verkaufsargument."))

story.append(PageBreak())

# ============================================================ 2. HOSTING STUFE 1
story.append(Paragraph("2 · Hosting-Kosten Stufe 1 – Start (1–10 Firmen)", st_h1))
story.append(Paragraph(
    "Alle Posten von der Domain bis zum Webhoster. Preise sind realistische deutsche "
    "Marktstaende (2026, Brutto-Orientierung). Anbieter sind Beispiele – austauschbar.", st_body))
story.append(Spacer(1, 3 * mm))
hosting1 = [
    ("Domain .de", "Netcup / INWX", "0,60 €", "7 €"),
    ("Domain .com (optional)", "INWX", "1,00 €", "12 €"),
    ("VPS (2 vCPU, 4 GB RAM)", "Hetzner CX22", "4,50 €", "54 €"),
    ("Automatische Backups (+20 %)", "Hetzner", "0,90 €", "11 €"),
    ("SSL-Zertifikat", "Let's Encrypt", "0 €", "0 €"),
    ("Transaktions-Mail (Reset/Einladung)", "Brevo Free (300/Tag)", "0 €", "0 €"),
    ("Uptime-Monitoring", "UptimeRobot Free", "0 €", "0 €"),
    ("<b>Summe (mit beiden Domains)</b>", "", "<b>~7 €</b>", "<b>~84 €</b>"),
]
t = table(hosting1, ["Posten", "Anbieter (Beispiel)", "pro Monat", "pro Jahr"],
          [62 * mm, 50 * mm, 30 * mm, 32 * mm], right_cols=(2, 3))
total_row_style(t, len(hosting1), 4)
story.append(t)
story.append(Spacer(1, 3 * mm))
story.append(note_box(
    "Unterm Strich: unter 10 €/Monat",
    "Realistisch alles drin bleibst du im Start <b>unter 10 €/Monat</b>. <b>Hetzner</b> betreibt "
    "deutsche Rechenzentren – das ist bei Personaldaten DSGVO-seitig stark und ein gutes Argument "
    "gegenueber US-Cloud. SSL ist via Let's Encrypt kostenlos, Mail- und Monitoring-Freitiers "
    "reichen fuer den Anfang locker.",
    bg=colors.HexColor("#EAF6EE"), edge=GRUEN))

story.append(PageBreak())

# ============================================================ 3. HOSTING STUFE 2
story.append(Paragraph("3 · Hosting-Kosten Stufe 2 – Wachstum (bis ~30 Firmen)", st_h1))
story.append(Paragraph(
    "Wenn mehr Mandanten und mehr Mail-Volumen dazukommen, waechst nur der Server eine Stufe "
    "und es kommt externer Backup-Speicher dazu. Immer noch <b>ein</b> Server.", st_body))
story.append(Spacer(1, 3 * mm))
hosting2 = [
    ("VPS (4 vCPU, 8 GB RAM)", "Hetzner CX32", "7,00 €", "84 €"),
    ("Automatische Backups", "Hetzner", "1,40 €", "17 €"),
    ("Externer Backup-Speicher (1 TB)", "Hetzner Storage Box", "3,90 €", "47 €"),
    ("Transaktions-Mail (mehr Volumen)", "Brevo Starter", "9,00 €", "108 €"),
    ("Domain(s)", "Netcup / INWX", "1,60 €", "19 €"),
    ("Monitoring (optional Pro)", "Better Stack", "0–7 €", "0–84 €"),
    ("<b>Summe</b>", "", "<b>~23–30 €</b>", "<b>~275–360 €</b>"),
]
t = table(hosting2, ["Posten", "Anbieter (Beispiel)", "pro Monat", "pro Jahr"],
          [62 * mm, 50 * mm, 30 * mm, 32 * mm], head_bg=TUERKIS, right_cols=(2, 3))
total_row_style(t, len(hosting2), 4)
story.append(t)
story.append(Spacer(1, 4 * mm))

story.append(Paragraph("3a · Alternative: „Ich will mich um nichts kuemmern“", st_h2))
story.append(Paragraph(
    "Statt einen VPS selbst zu pflegen (Updates, Patches) geht auch ein <b>Managed-PaaS</b> wie "
    "<b>Render</b>, <b>Railway</b> oder <b>Fly.io</b>: ~20–30 €/Monat ab Tag 1, dafuer kein "
    "Server-Administrieren. <b>Aber:</b> fuer das Modell „eine SQLite-Datei pro Firma“ ist ein "
    "<b>VPS mit persistenter Platte die bessere Wahl</b> – volle Kontrolle ueber die Mandanten-"
    "Dateien und das Backup. Empfehlung daher: VPS.", st_body))
story.append(Spacer(1, 2 * mm))
story.append(note_box(
    "Skalierungs-Hinweis",
    "SQLite-pro-Firma skaliert auf einem Server erstaunlich weit, weil die Mandanten sich keine "
    "Datenbank teilen. Erst bei sehr vielen gleichzeitigen Nutzern pro Firma oder dem Wunsch nach "
    "echter Ausfallsicherheit (zweiter Server) lohnt der Wechsel zu einer zentralen DB – das ist "
    "eine spaetere Entscheidung, keine, die den Start blockiert.",
    bg=colors.HexColor("#EAF2FB"), edge=BLAU))

story.append(PageBreak())

# ============================================================ 4. KUNDENPREISE
story.append(Paragraph("4 · Was du Kunden berechnen kannst", st_h1))
story.append(Paragraph(
    "Personaldaten haben hohen wahrgenommenen Wert und du nimmst dem Kunden die DSGVO-Verantwortung "
    "fuers Hosting ab. Zwei gaengige Modelle – fuer den Start ist Modell A (Flat pro Firma) am "
    "einfachsten zu kommunizieren.", st_body))
story.append(Spacer(1, 3 * mm))

story.append(Paragraph("4a · Modell A: Flat pro Firma (empfohlen)", st_h2))
tarife = [
    ("Starter", "bis 10 Mitarbeiter", "Stammdaten, Custom-Felder", "19 €"),
    ("Standard", "bis 30 Mitarbeiter", "+ Vorgesetzten-Matrix, Rollen", "39 €"),
    ("Pro", "unbegrenzt", "+ Organigramm, Prioritaets-Support", "79 €"),
]
t = table(tarife, ["Tarif", "Umfang", "Inhalt", "pro Monat"],
          [26 * mm, 40 * mm, 76 * mm, 32 * mm], right_cols=(3,))
story.append(t)
story.append(Spacer(1, 4 * mm))

story.append(Paragraph("4b · Modell B: Pro Mitarbeiter", st_h2))
story.append(Paragraph(
    "Skaliert mit der Kundengroesse: <b>2,50–3 € pro Mitarbeiter und Monat</b>, Mindestabnahme "
    "19 €. Fairer fuer sehr kleine Kunden, rechnet sich bei groesseren staerker – aber etwas "
    "erklaerungsbeduerftiger in der Rechnung.", st_body))
story.append(Spacer(1, 3 * mm))
story.append(note_box(
    "Markt-Einordnung",
    "Etablierte HR-Tools (Personio &amp; Co.) liegen bei <b>5–12 € pro Mitarbeiter</b> und Monat – "
    "mit grossem Funktionsumfang, aber auch hoher Komplexitaet. Als <b>schlanke Loesung fuer "
    "Grosshandel und Apotheken</b> positionierst du dich bewusst deutlich darunter: weniger Ballast, "
    "DSGVO-konform aus Deutschland, fairer Preis."))

story.append(PageBreak())

# ============================================================ 5. DECKUNGSBEITRAG
story.append(Paragraph("5 · Deckungsbeitrag – was uebrig bleibt", st_h1))
story.append(Paragraph(
    "Gegenueberstellung von Einnahmen (Modell A) und Hosting-Kosten. Der Server ist nahezu der "
    "einzige Fixposten und traegt alle Szenarien.", st_body))
story.append(Spacer(1, 3 * mm))
db = [
    ("5 Firmen × 39 €", "195 €", "~10 €", "~185 €"),
    ("15 Firmen × 39 €", "585 €", "~25 €", "~560 €"),
    ("30 Firmen × Ø 45 € (Mix)", "1.350 €", "~30 €", "~1.320 €"),
]
t = table(db, ["Szenario", "Einnahmen/Monat", "Hosting/Monat", "Marge/Monat"],
          [54 * mm, 40 * mm, 38 * mm, 42 * mm], head_bg=GRUEN, right_cols=(1, 2, 3))
story.append(t)
story.append(Spacer(1, 3 * mm))
story.append(Paragraph(
    "Die Hosting-<b>Grenzkosten pro Kunde liegen unter 1 €/Monat</b>. Dein realer Aufwand ist "
    "deine <b>Arbeitszeit</b> (Support, DSGVO, Weiterentwicklung) – nicht die Technik. Genau "
    "deshalb traegt sich das Modell schon ab wenigen Kunden.", st_body))
story.append(Spacer(1, 3 * mm))
story.append(note_box(
    "Einmalige &amp; weitere laufende Posten (nicht im Hosting oben)",
    "• <b>Zahlungsgebuehren:</b> Stripe ~1,5 % + 0,25 € je SEPA-Lastschrift.<br/>"
    "• <b>Rechtliches einmalig:</b> AGB/AV-Vertrag/Datenschutz pruefen lassen – einmalig einige "
    "hundert Euro, optional.<br/>"
    "• <b>Deine Entwicklungszeit</b> fuer die offenen Bausteine aus Abschnitt 1 – der groesste "
    "„Preis“, aber Eigenleistung.",
    bg=WARNBG, edge=GELB))

story.append(PageBreak())

# ============================================================ 6. EMPFEHLUNG
story.append(Paragraph("6 · Empfehlung &amp; naechste Schritte", st_h1))
story.append(Spacer(1, 1 * mm))
schritte = [
    ("1. Hosting aufsetzen", "Hetzner CX22 + Backups + .de-Domain → unter 10 €/Monat. Let's Encrypt fuer SSL."),
    ("2. DSGVO zuerst", "AV-Vertrag, Loeschkonzept, Datenexport je Mandant – vor dem ersten zahlenden Kunden."),
    ("3. Onboarding &amp; Reset", "Self-Service Firma anlegen + Passwort-Reset/Einladung per Brevo-Mail."),
    ("4. Backups testen", "Automatisch je Mandant sichern UND eine Wiederherstellung einmal echt durchspielen."),
    ("5. Abrechnung", "Stripe anbinden, Tarife 19/39/79 € hinterlegen."),
    ("6. Preis kommunizieren", "Flat-Tarife – einfach, kalkulierbar, planbar fuer den Kunden."),
]
story.append(table(schritte, ["Schritt", "Inhalt"], [44 * mm, 130 * mm]))
story.append(Spacer(1, 4 * mm))

story.append(note_box(
    "Kernaussage",
    "Die <b>Technik kostet fast nichts</b> (unter 10 €/Monat im Start, unter 1 € je Kunde). Schon "
    "<b>5 Kunden</b> bringen ~185 € Marge im Monat. Der Engpass ist nicht das Hosting, sondern die "
    "<b>Fertigstellung der offenen Bausteine</b> (v. a. DSGVO) und der laufende Support. Wenn die "
    "stehen, ist das ein wirtschaftlich sehr tragfaehiges Zusatzprodukt.",
    bg=colors.HexColor("#EAF2FB"), edge=BLAU))
story.append(Spacer(1, 5 * mm))

story.append(HRFlowable(width="100%", thickness=0.8, color=HELLGRAU, spaceAfter=4))
story.append(Paragraph(
    "<b>Hinweis &amp; Vorbehalt:</b> Diese Planung ist eine interne Projekt- und Kostenunterlage. "
    "Alle Preise sind <b>Marktstaende (Deutschland, 2026)</b> zur Orientierung und koennen sich "
    "aendern – vor Vertragsabschluss bei den jeweiligen Anbietern pruefen. Die rechtlichen "
    "Aussagen (DSGVO, AV-Vertrag) ersetzen keine Rechtsberatung. Stand: " +
    date.today().strftime("%d.%m.%Y") + ".", st_small))

# --- Build --------------------------------------------------------------------
os.makedirs(os.path.dirname(OUT), exist_ok=True)
doc = SimpleDocTemplate(
    OUT, pagesize=A4,
    leftMargin=18 * mm, rightMargin=18 * mm,
    topMargin=18 * mm, bottomMargin=16 * mm,
    title="Planung Personal-Modul Web NMGone", author="NMGone")
doc.build(story, onFirstPage=header_footer, onLaterPages=header_footer)
print("OK ->", OUT)
