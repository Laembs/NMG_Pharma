# -*- coding: utf-8 -*-
"""Erzeugt den Projekt- und Aufwandsbericht NMGone als PDF.

Eine Quelle, ein PDF: Umgesetztes, Geplantes (mit Zeitschaetzung) und eine
Aufwands-/Kostenrechnung inkl. Verhandlungsspanne fuer NMG Pharma.

Scope: NUR das NMGone-Desktop-Programm (Programm-Familie). Der separate
Web-SaaS-Pilot ist bewusst NICHT enthalten.
"""
from __future__ import annotations

import os
from datetime import date

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.lib.enums import TA_LEFT, TA_RIGHT, TA_CENTER
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak,
    HRFlowable, KeepTogether,
)

# --- Farben / Markenwelt ------------------------------------------------------
BLAU = colors.HexColor("#0B4A86")
TUERKIS = colors.HexColor("#0B6E6E")
GRAU = colors.HexColor("#5A6675")
HELLGRAU = colors.HexColor("#EEF1F5")
ZEILE = colors.HexColor("#F6F8FA")
GRUEN = colors.HexColor("#11823B")
ROT = colors.HexColor("#B5391F")

OUT = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                   "docs", "Projektbericht_NMGone.pdf")

# --- Stile --------------------------------------------------------------------
styles = getSampleStyleSheet()


def S(name, **kw):
    base = kw.pop("parent", styles["Normal"])
    return ParagraphStyle(name, parent=base, **kw)


st_title = S("t", parent=styles["Title"], fontSize=24, textColor=BLAU, leading=28, spaceAfter=2)
st_sub = S("s", fontSize=11.5, textColor=GRAU, spaceAfter=2)
st_h1 = S("h1", fontSize=15, textColor=BLAU, spaceBefore=14, spaceAfter=6, leading=18,
          fontName="Helvetica-Bold")
st_h2 = S("h2", fontSize=11.5, textColor=TUERKIS, spaceBefore=8, spaceAfter=3,
          fontName="Helvetica-Bold")
st_body = S("b", fontSize=9.5, leading=13, textColor=colors.HexColor("#23303D"))
st_small = S("sm", fontSize=8, leading=10, textColor=GRAU)
st_cell = S("c", fontSize=8.5, leading=11)
st_cell_b = S("cb", fontSize=8.5, leading=11, fontName="Helvetica-Bold")
st_cell_w = S("cw", fontSize=8.5, leading=11, textColor=colors.white, fontName="Helvetica-Bold")
st_kpi_num = S("kn", fontSize=20, textColor=BLAU, fontName="Helvetica-Bold", alignment=TA_CENTER, leading=22)
st_kpi_lbl = S("kl", fontSize=7.5, textColor=GRAU, alignment=TA_CENTER, leading=9)
st_right = S("r", fontSize=8.5, leading=11, alignment=TA_RIGHT)
st_right_b = S("rb", fontSize=8.5, leading=11, alignment=TA_RIGHT, fontName="Helvetica-Bold")


def header_footer(canvas, doc):
    canvas.saveState()
    w, h = A4
    # Kopfbalken
    canvas.setFillColor(BLAU)
    canvas.rect(0, h - 12 * mm, w, 12 * mm, stroke=0, fill=1)
    canvas.setFillColor(colors.white)
    canvas.setFont("Helvetica-Bold", 9)
    canvas.drawString(18 * mm, h - 8 * mm, "NMGone  ·  Projekt- & Aufwandsbericht")
    canvas.setFont("Helvetica", 8)
    canvas.drawRightString(w - 18 * mm, h - 8 * mm, date.today().strftime("Stand %d.%m.%Y"))
    # Fusszeile
    canvas.setStrokeColor(HELLGRAU)
    canvas.setLineWidth(0.6)
    canvas.line(18 * mm, 12 * mm, w - 18 * mm, 12 * mm)
    canvas.setFillColor(GRAU)
    canvas.setFont("Helvetica", 7.5)
    canvas.drawString(18 * mm, 8 * mm,
                      "Vertraulich – interne Projektunterlage. Schaetzwerte, keine Rechts- oder Steuerberatung.")
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
            ("BOX", (0, 0), (-1, -1), 0, HELLGRAU),
        ]))
        cells.append(inner)
    t = Table([cells], colWidths=[44 * mm] * len(cells))
    t.setStyle(TableStyle([("LEFTPADDING", (0, 0), (-1, -1), 2),
                           ("RIGHTPADDING", (0, 0), (-1, -1), 2)]))
    return t


def status_table(rows, header, col_w, status_col=None):
    data = [[Paragraph(h, st_cell_w) for h in header]]
    for r in rows:
        data.append([Paragraph(str(c), st_cell) for c in r])
    t = Table(data, colWidths=col_w, repeatRows=1)
    style = [
        ("BACKGROUND", (0, 0), (-1, 0), BLAU),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, ZEILE]),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#D7DEE6")),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ("LEFTPADDING", (0, 0), (-1, -1), 5),
        ("RIGHTPADDING", (0, 0), (-1, -1), 5),
    ]
    t.setStyle(TableStyle(style))
    return t


def money(v):
    return "{:,.0f} €".format(v).replace(",", ".")


story = []

# ============================================================ TITELSEITE
story.append(Spacer(1, 6 * mm))
story.append(Paragraph("Projektbericht NMGone", st_title))
story.append(Paragraph("Von der Einzelauswertung zur Pharma-Großhandels-Programmfamilie", st_sub))
story.append(Paragraph("Stand der Umsetzung · offene Roadmap mit Zeitschätzung · Aufwand &amp; Wert", st_sub))
story.append(Spacer(1, 5 * mm))
story.append(HRFlowable(width="100%", thickness=1.2, color=BLAU, spaceAfter=8))

story.append(kpi_row([
    ("8", "eigenständige Apps"),
    ("52", "Module in app/"),
    ("~46.000", "Zeilen Code (app/)"),
    ("19+", "Funktionsbereiche"),
]))
story.append(Spacer(1, 5 * mm))

story.append(Paragraph("Worum es geht", st_h2))
story.append(Paragraph(
    "NMGone ist über die Projektlaufzeit von einem einzelnen Analyse-Werkzeug zu einer "
    "<b>Familie eigenständiger Programme rund um eine gemeinsame Datenbank</b> gewachsen. "
    "Jede App ist per eigenem Symbol startbar und bildet einen abgeschlossenen Arbeitsbereich "
    "eines Pharma-Parallelimporteurs / Großhandels ab: Beschaffung, Wareneingang nach GDP, "
    "Verkauf, Faktura, Meldewesen, Personal und Auswertung. Dieser Bericht hält fest, was "
    "umgesetzt ist, was noch geplant ist (mit Zeitschätzung) und welchen Aufwand bzw. Wert "
    "das Gesamtwerk darstellt.", st_body))
story.append(Spacer(1, 3 * mm))
story.append(Paragraph(
    "<b>Hinweis zum Umfang:</b> Dieser Bericht betrifft ausschließlich das NMGone-"
    "Desktop-Programm. Der separat geführte Web-/SaaS-Pilot ist bewusst <b>nicht</b> Teil "
    "dieser Aufstellung.", st_small))
story.append(Spacer(1, 4 * mm))

story.append(Paragraph("Inhalt", st_h2))
story.append(Paragraph(
    "1. Was ist umgesetzt (Module &amp; Funktionen)<br/>"
    "2. Was ist noch geplant – Roadmap mit Zeitschätzung<br/>"
    "3. Aufwand, Kosten &amp; Verhandlungsspanne (letzte Seite)", st_body))

story.append(PageBreak())

# ============================================================ 1. UMGESETZT
story.append(Paragraph("1 · Umgesetzt", st_h1))
story.append(Paragraph(
    "Acht startbare Apps plus gemeinsame Infrastruktur. „Status“ bezieht sich auf die "
    "Reife: <b>produktiv</b> = im Einsatz nutzbar, <b>nutzbar</b> = funktionsfähig, Feinschliff offen.",
    st_body))
story.append(Spacer(1, 3 * mm))

apps = [
    ("NMGone (Kern)", "Dashboard mit App-Kacheln, Bedarfsanalyse, Produktanalyse, Kunden-CRM "
        "(Steckbrief, ABC-Analyse, Offline-Deutschlandkarte), Wissens-Datenbanken.", "produktiv"),
    ("Kasse", "Wareneingang → Lager → Verkauf an Apotheken, Lieferschein, Defektmeldung, "
        "EK-/Lagerwert-Auswertungen.", "produktiv"),
    ("Faktura", "Rechnungen, Gutschriften, Quartalsvergütung mit anpassbaren Vorlagen.", "produktiv"),
    ("Wareneingang &amp; Retouren (GDP)", "Wareneingangsprüfung, Chargen-Rückverfolgung "
        "(Kunde↔Charge), Kundenqualifizierung, Retouren/Reklamation mit Gutschrift-Workflow, "
        "revisionssicheres Protokoll.", "nutzbar"),
    ("Meldungen", "GDP-Meldewesen (Abweichungen/CAPA), Kühlsachenkontrolle, Selbstinspektion.", "nutzbar"),
    ("Einkauf", "Beschaffung EU-Ausland, Lieferanten je Land (Währung/Lieferzeit), "
        "§129-Margenrechner, Aufgaben/Wiedervorlagen.", "nutzbar"),
    ("Mitarbeiter / Personal", "Organigramm-Datenbasis, Abwesenheiten, Arbeitsbereiche, "
        "Custom-Felder (EAV), Vorgesetzten-Matrix.", "nutzbar"),
    ("Auswertungen / Report &amp; Kurzbericht", "Freie Auswertungen + Export; Kurzbericht als "
        "Excel/PDF mit Diagrammen.", "nutzbar"),
    ("Hilfe / Handbuch", "Bebildertes Handbuch zu allen Modulen, druckbare Handouts.", "nutzbar"),
    ("Parameter / Berechtigungen", "Rollen, Overrides, Admin-PIN – „wer darf was“.", "nutzbar"),
]
story.append(status_table(
    [(n, b, s) for n, b, s in apps],
    ["App / Modul", "Funktionsumfang", "Status"],
    [46 * mm, 110 * mm, 18 * mm]))

story.append(Spacer(1, 5 * mm))
story.append(Paragraph("Querschnitt &amp; Infrastruktur", st_h2))
infra = [
    ("Wissens-Datenbanken", "Austausch-/Biosimilar- (Gelbe Liste), Wirkstoff-, Hersteller-, "
        "Artikel-Stammdaten, NMG-Rabatte inkl. Import &amp; PZN-Normalisierung."),
    ("Daten &amp; Betrieb", "SQLite-Datenmodell mit Migrationen, Backup, Update-Manager, "
        "Demo-Modus (eigene Demo-DB, Produktiv-DB unberührt)."),
    ("Oberfläche", "Zentrales Design-System (theme.py), moderne Testoberfläche, "
        "Revisions-Übersicht „Was wurde verändert“, Filter-Persistenz."),
    ("Auslieferung", "PyInstaller-Build, Windows-Installer (Inno Setup), eigene App-Symbole, "
        "Start per Taskleiste/Kachel, mehrsprachige Basis (i18n)."),
]
story.append(status_table(infra, ["Bereich", "Inhalt"], [40 * mm, 134 * mm]))

story.append(PageBreak())

# ============================================================ 2. GEPLANT
story.append(Paragraph("2 · Geplant – Roadmap mit Zeitschätzung", st_h1))
story.append(Paragraph(
    "Zeitschätzung in <b>Personentagen (PT)</b> für die Restumsetzung inkl. Test &amp; "
    "Einbau in die bestehende Programm-Familie. Priorität: H = hoch, M = mittel, N = niedrig.",
    st_body))
story.append(Spacer(1, 3 * mm))

roadmap = [
    ("Zentrale Mehrbenutzer-Datenbank", "Weg von der einzelnen SQLite-Datei hin zu einem "
        "zentralen DB-Server, damit mehrere Arbeitsplätze gleichzeitig arbeiten.", "H", "20–35"),
    ("Produktanalyse-Bugs fixen", "EK durchweg NULL (Umsatz=0); PZN-Normalisierung ohne "
        "'/N'-Strip korrigieren.", "H", "1–2"),
    ("Echte Screenshots in die Hilfe", "Bild-Slots der neuen Kapitel (Wareneingang, "
        "Meldungen, Einkauf) mit echten Screenshots füllen.", "M", "1–2"),
    ("App-Symbole &amp; Setup finalisieren", "Eigene Icons für Meldungen/Einkauf, "
        "Verknüpfungen im Installer ergänzen.", "M", "1–2"),
    ("Kunden-Deutschlandkarte fertigstellen", "Offline-Geokodierung (geo_de.py) in der "
        "Kunden-App zur Karte ausbauen.", "M", "3–5"),
    ("Biosimilar-Wissensbasis verzahnen", "Gelbe-Liste-Wissensbasis mit Bedarfsanalyse &amp; "
        "GUI verbinden (Wirkstoff-Biosimilar-Zuordnung).", "M", "4–6"),
    ("Mitarbeiter-Organigramm Phase 2", "Aus der Datenbasis die grafische Organigramm-"
        "Darstellung erzeugen.", "M", "4–6"),
    ("Berechtigungen echt durchsetzen", "Rollen/Overrides nicht nur verwalten, sondern in "
        "allen Apps tatsächlich erzwingen.", "M", "4–7"),
    ("MSV3-Bestell-Import (Kasse)", "MSV3-Bestellungen aus Outlook (sales@) in den Verkauf "
        "importieren – Mail-Format ist der Blocker.", "M", "5–8"),
    ("Modulübergreifendes Report-Programm", "Alle NMGone-Logs/Protokolle in einem Report "
        "bündeln.", "N", "5–8"),
    ("XLS-Pseudo-Format überall", "Apotheken-.xls (XML/HTML) auch in Manueller Import &amp; "
        "Testoberfläche sauber laden.", "N", "1–2"),
    ("Filter-Persistenz ausrollen", "Gehaltene Filter über alle Apps hinweg (bisher nur "
        "Einkauf).", "N", "1–2"),
]
story.append(status_table(roadmap,
    ["Vorhaben", "Inhalt", "Prio", "PT"],
    [44 * mm, 100 * mm, 12 * mm, 18 * mm]))

story.append(Spacer(1, 4 * mm))
sum_min = 50
sum_max = 85
story.append(Paragraph(
    f"<b>Summe Restaufwand: rund {sum_min}–{sum_max} Personentage</b> – getrieben vor allem "
    "vom großen Brocken „zentrale Mehrbenutzer-Datenbank“. Ohne diesen liegt der Rest bei "
    "rund 30–50 PT.", st_body))

story.append(Spacer(1, 4 * mm))
story.append(Paragraph("Lesehilfe Aufwandsklassen", st_h2))
story.append(Paragraph(
    "<b>klein</b> ≈ 1–2 PT · <b>mittel</b> ≈ 3–8 PT · <b>groß</b> ≈ 15+ PT. "
    "Ein Personentag = ein voller Arbeitstag konzentrierter Entwicklung inkl. Test.", st_small))

story.append(PageBreak())

# ============================================================ 3. AUFWAND & KOSTEN
story.append(Paragraph("3 · Aufwand, Kosten &amp; Verhandlungsspanne", st_h1))
story.append(Paragraph(
    "Die folgende Rechnung schätzt den <b>Wiederbeschaffungswert</b>: Was würde es kosten, "
    "NMGone in heutigem Umfang professionell neu erstellen zu lassen? Das ist der belastbare "
    "Anker für eine Verhandlung – unabhängig davon, wie schnell die Umsetzung tatsächlich war.",
    st_body))
story.append(Spacer(1, 3 * mm))

# 3a Aufwand je Baustein
story.append(Paragraph("3a · Geschätzter Erstellungsaufwand (Agentur-Äquivalent)", st_h2))
aufwand = [
    ("Kern: Dashboard, Bedarfs-/Produktanalyse, Kunden-CRM", "40"),
    ("Wissens-Datenbanken inkl. Import &amp; Normalisierung", "25"),
    ("Kasse (Wareneingang/Lager/Verkauf)", "25"),
    ("Faktura (Rechnungen/Gutschriften/Vorlagen)", "20"),
    ("GDP – Wareneingang &amp; Retouren", "18"),
    ("Einkauf (EU-Beschaffung, §129-Rechner)", "18"),
    ("Meldungen (Meldewesen/Kühlkette/Inspektion)", "10"),
    ("Personal / Organigramm", "15"),
    ("Report + Kurzbericht", "12"),
    ("Hilfe / Handbuch", "8"),
    ("Parameter / Berechtigungen", "10"),
    ("Oberfläche, Theme, Testoberfläche, Revision", "15"),
    ("Infrastruktur: DB, Migrationen, Backup, Update, Installer, Demo", "25"),
]
total_pt = sum(int(x) for _, x in aufwand)
rows = [[Paragraph(n, st_cell), Paragraph(p, st_right)] for n, p in aufwand]
rows.append([Paragraph("<b>Summe Erstellungsaufwand</b>", st_cell_b),
             Paragraph(f"<b>{total_pt}</b>", st_right_b)])
t = Table([[Paragraph("Baustein", st_cell_w), Paragraph("PT", st_cell_w)]] + rows,
          colWidths=[150 * mm, 24 * mm])
t.setStyle(TableStyle([
    ("BACKGROUND", (0, 0), (-1, 0), BLAU),
    ("ROWBACKGROUNDS", (0, 1), (-1, -2), [colors.white, ZEILE]),
    ("BACKGROUND", (0, -1), (-1, -1), HELLGRAU),
    ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#D7DEE6")),
    ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ("TOPPADDING", (0, 0), (-1, -1), 3.5),
    ("BOTTOMPADDING", (0, 0), (-1, -1), 3.5),
    ("LEFTPADDING", (0, 0), (-1, -1), 5),
    ("RIGHTPADDING", (0, 0), (-1, -1), 5),
]))
story.append(t)
story.append(Paragraph(
    f"≈ {total_pt} PT entsprechen grob <b>{total_pt/21:.0f}–{total_pt/19:.0f} Personenmonaten</b> "
    "Vollzeit-Entwicklung eines erfahrenen Entwicklers (inkl. Konzept, Test, Feinschliff).", st_small))
story.append(Spacer(1, 4 * mm))

# 3b Kosten
story.append(Paragraph("3b · Wiederbeschaffungswert nach Tagessatz", st_h2))
saetze = [("konservativ", 650), ("markt­üblich (Senior-Freelance)", 850), ("Agentur/Spezialanbieter", 1150)]
rows = [[Paragraph(lbl, st_cell), Paragraph(money(rate) + " / PT", st_right),
         Paragraph(money(rate * total_pt), st_right_b)] for lbl, rate in saetze]
t = Table([[Paragraph("Szenario", st_cell_w), Paragraph("Tagessatz", st_cell_w),
            Paragraph(f"Wert bei {total_pt} PT", st_cell_w)]] + rows,
          colWidths=[78 * mm, 48 * mm, 48 * mm])
t.setStyle(TableStyle([
    ("BACKGROUND", (0, 0), (-1, 0), TUERKIS),
    ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, ZEILE]),
    ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#D7DEE6")),
    ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ("TOPPADDING", (0, 0), (-1, -1), 5),
    ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
    ("LEFTPADDING", (0, 0), (-1, -1), 5),
    ("RIGHTPADDING", (0, 0), (-1, -1), 5),
]))
story.append(t)
story.append(Spacer(1, 2 * mm))
lo = saetze[0][1] * total_pt
hi = saetze[2][1] * total_pt
story.append(Paragraph(
    f"<b>Wiederbeschaffungswert NMGone: rund {money(lo)} bis {money(hi)}.</b> "
    "Realistischer Verhandlungsanker für ein fertiges, im Betrieb erprobtes, fachlich "
    "spezialisiertes System: die Mitte, also Größenordnung <b>" + money(saetze[1][1]*total_pt) + "</b>.",
    st_body))
story.append(Spacer(1, 4 * mm))

# 3c Verhandlung
story.append(Paragraph("3c · Was man bei NMG Pharma ansetzen kann", st_h2))
modelle = [
    ("Einmal-Kauf (Buy-out)", "Quellcode + Nutzungsrechte gehen an NMG Pharma.",
     f"{money(160000)} – {money(300000)}"),
    ("Lizenz + Wartung/Jahr", "Nutzungsrecht; jährliche Pflege/Weiterentwicklung 15–20 % p.a.",
     f"{money(60000)}–{money(120000)} + {money(15000)}–{money(25000)}/J."),
    ("Pro-Arbeitsplatz-Lizenz", "Pro Seat/Standort, planbar skalierend.",
     "≈ " + money(1500) + "–" + money(3500) + " / Platz"),
    ("Dienstleister-Modell", "Weiterbetrieb/-entwicklung als Auftrag, nach Tagessatz.",
     money(700) + "–" + money(950) + " / PT"),
]
rows = [[Paragraph(m, st_cell_b), Paragraph(b, st_cell), Paragraph(v, st_right)] for m, b, v in modelle]
t = Table([[Paragraph("Modell", st_cell_w), Paragraph("Prinzip", st_cell_w),
            Paragraph("Größenordnung", st_cell_w)]] + rows,
          colWidths=[42 * mm, 80 * mm, 52 * mm])
t.setStyle(TableStyle([
    ("BACKGROUND", (0, 0), (-1, 0), BLAU),
    ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, ZEILE]),
    ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#D7DEE6")),
    ("VALIGN", (0, 0), (-1, -1), "TOP"),
    ("TOPPADDING", (0, 0), (-1, -1), 4),
    ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ("LEFTPADDING", (0, 0), (-1, -1), 5),
    ("RIGHTPADDING", (0, 0), (-1, -1), 5),
]))
story.append(t)
story.append(Spacer(1, 3 * mm))

story.append(Paragraph("Argumente, die den Preis stützen", st_h2))
story.append(Paragraph(
    "• <b>Fachtiefe:</b> bildet GDP-Pflichten, §129-Marge und Parallelimport-Spezifika ab – "
    "Standardsoftware kann das nicht von der Stange.<br/>"
    "• <b>Ersparnis statt Kosten:</b> ersetzt mehrere Insel-/Excel-Lösungen und manuelle "
    "Schritte; spart laufend Arbeitszeit und reduziert Fehler/Audit-Risiken.<br/>"
    "• <b>Betriebsreif &amp; ausgeliefert:</b> Installer, Demo-Modus, Handbuch, Update-Mechanik – "
    "kein Prototyp, sondern eingesetztes System.<br/>"
    "• <b>Abhängigkeit/Klumpenrisiko:</b> Wissen liegt heute bei einer Person – ein Buy-out "
    "oder Wartungsvertrag sichert NMG Pharma genau dagegen ab (rechtfertigt Aufpreis).",
    st_body))
story.append(Spacer(1, 4 * mm))

story.append(HRFlowable(width="100%", thickness=0.8, color=HELLGRAU, spaceAfter=4))
story.append(Paragraph(
    "<b>Methodik &amp; Vorbehalt:</b> Aufwand bottom-up je Modul geschätzt und gegen den "
    f"Code-Umfang (~46.000 Zeilen in app/, 52 Module) plausibilisiert. Tagessätze sind "
    "marktübliche Spannen für spezialisierte B2B-/Pharma-Software (DE, Stand 2026). Alle "
    "Beträge sind <b>Schätzwerte zur Verhandlungsvorbereitung</b>, netto, und ersetzen keine "
    "Rechts-, Steuer- oder Sachverständigenberatung. Tatsächliche Preise hängen von Rechte-"
    "umfang, Wartungszusagen und Verhandlung ab.", st_small))

# --- Build --------------------------------------------------------------------
os.makedirs(os.path.dirname(OUT), exist_ok=True)
doc = SimpleDocTemplate(
    OUT, pagesize=A4,
    leftMargin=18 * mm, rightMargin=18 * mm,
    topMargin=18 * mm, bottomMargin=16 * mm,
    title="Projektbericht NMGone", author="NMGone")
doc.build(story, onFirstPage=header_footer, onLaterPages=header_footer)
print("OK ->", OUT)
print("Summe Erstellungsaufwand (PT):", total_pt)
