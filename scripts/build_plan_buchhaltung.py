# -*- coding: utf-8 -*-
"""Erzeugt die Planung "Buchhaltungs-App" fuer NMGone als PDF.

Eine Quelle, ein PDF: Was die App leisten soll, was wir konkret reinbringen,
Standard-Kontenlisten (SKR03 + SKR04), USt-Behandlung, eRechnung-Empfang,
GoBD/DATEV, eine Umsetzungs-Roadmap mit Zeitschaetzung und – klar markiert –
die Punkte, die mit dem Steuerbuero abgestimmt werden muessen.

Scope: Buchhaltungs-/Vorerfassungs-App innerhalb der NMGone-Programmfamilie,
KEINE vollwertige Finanzbuchhaltung mit Bilanz. Keine Bareinnahmen -> kein
KassenSichV/TSE/DSFinV-K-Thema.

Aufruf:  python scripts/build_plan_buchhaltung.py
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
GELB = colors.HexColor("#C88200")
WARNBG = colors.HexColor("#FCF3E6")

OUT = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                   "docs", "Plan_Buchhaltung_App.pdf")

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
st_kpi_num = S("kn", fontName=BOLD, fontSize=19, textColor=BLAU, alignment=TA_CENTER, leading=21)
st_kpi_lbl = S("kl", fontName=BASE, fontSize=7.5, textColor=GRAU, alignment=TA_CENTER, leading=9)
st_right = S("r", fontName=BASE, fontSize=8.5, leading=11, alignment=TA_RIGHT)
st_right_b = S("rb", fontName=BOLD, fontSize=8.5, leading=11, alignment=TA_RIGHT)


def header_footer(canvas, doc):
    canvas.saveState()
    w, h = A4
    canvas.setFillColor(BLAU)
    canvas.rect(0, h - 12 * mm, w, 12 * mm, stroke=0, fill=1)
    canvas.setFillColor(colors.white)
    canvas.setFont(BOLD, 9)
    canvas.drawString(18 * mm, h - 8 * mm, "NMGone  ·  Planung Buchhaltungs-App")
    canvas.setFont(BASE, 8)
    canvas.drawRightString(w - 18 * mm, h - 8 * mm, date.today().strftime("Stand %d.%m.%Y"))
    canvas.setStrokeColor(HELLGRAU)
    canvas.setLineWidth(0.6)
    canvas.line(18 * mm, 12 * mm, w - 18 * mm, 12 * mm)
    canvas.setFillColor(GRAU)
    canvas.setFont(BASE, 7.5)
    canvas.drawString(18 * mm, 8 * mm,
                      "Interne Projektplanung. Konten-/USt-Angaben sind Vorschlaege, keine Steuerberatung – final mit dem Steuerbuero abstimmen.")
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
story.append(Paragraph("Planung: Buchhaltungs-App", st_title))
story.append(Paragraph("Gepruefte Daten aus NMGone an Buchhaltung &amp; Steuerbuero – nach deutschen Vorgaben", st_sub))
story.append(Paragraph("Was wir reinbringen · Umsetzungs-Roadmap · offene Punkte fuer das Steuerbuero", st_sub))
story.append(Spacer(1, 4 * mm))
story.append(HRFlowable(width="100%", thickness=1.2, color=BLAU, spaceAfter=8))

story.append(kpi_row([
    ("4", "Quellen / Vorgaenge"),
    ("SKR03+04", "Kontenrahmen"),
    ("DATEV", "Export-Format"),
    ("eRechnung", "Empfang + Versand"),
]))
story.append(Spacer(1, 5 * mm))

story.append(Paragraph("Worum es geht", st_h2))
story.append(Paragraph(
    "NMGone ist gross genug geworden, dass die laufenden Bewegungsdaten an die Buchhaltung "
    "bzw. das Steuerbuero uebergeben werden sollen: <b>Abverkauf</b> (Kasse/Verkauf an Apotheken), "
    "<b>Einkauf/Wareneingang</b> (Beschaffung, auch EU-Ausland) und <b>Ausgangsrechnungen</b> "
    "(Faktura). Die neue App sammelt diese Daten, ordnet sie Konten und Steuerschluesseln zu, "
    "<b>prueft</b> sie auf Plausibilitaet/Vollstaendigkeit und exportiert nur freigegebene, "
    "gepruefte Buchungen im <b>DATEV-Format</b> an das Steuerbuero.", st_body))
story.append(Spacer(1, 3 * mm))

story.append(note_box(
    "Klare Scope-Abgrenzung",
    "Die App ist eine <b>Vorerfassungs- und Export-App</b>, keine vollwertige Finanzbuchhaltung "
    "mit Bilanz/GuV. Das eigentliche Buchen, der Jahresabschluss und die Steuererklaerung bleiben "
    "beim <b>Steuerbuero</b> – das reduziert Haftung und Aufwand erheblich.<br/>"
    "<b>Kein Bargeld:</b> Es gibt keine Bareinnahmen, daher entfaellt das gesamte Thema "
    "KassenSichV / TSE / DSFinV-K.", bg=colors.HexColor("#EAF2FB"), edge=BLAU))
story.append(Spacer(1, 4 * mm))

story.append(Paragraph("Inhalt dieser Planung", st_h2))
story.append(Paragraph(
    "1. Datenquellen &amp; wie sie zu Buchungen werden<br/>"
    "2. Standard-Kontenlisten SKR03 und SKR04 (Vorschlag)<br/>"
    "3. Umsatzsteuer &amp; Steuerschluessel<br/>"
    "4. eRechnung – Empfang (Pflicht) &amp; Versand aus der Faktura<br/>"
    "5. Pruefung, Freigabe &amp; GoBD/DATEV-Export<br/>"
    "6. Umsetzungs-Roadmap mit Zeitschaetzung<br/>"
    "7. Offene Punkte – mit dem Steuerbuero abzustimmen", st_body))

story.append(PageBreak())

# ============================================================ 1. DATENQUELLEN
story.append(Paragraph("1 · Datenquellen &amp; Buchungslogik", st_h1))
story.append(Paragraph(
    "Jede Quelle wird verdichtet (keine Einzel-Buchung je Artikel) und ueber ein "
    "konfigurierbares Mapping „Geschaeftsvorfall → Konto + Steuerschluessel“ in "
    "Buchungssaetze (Soll/Haben) uebersetzt.", st_body))
story.append(Spacer(1, 3 * mm))
quellen = [
    ("Faktura (Ausgangsrechnungen)", "Rechnungen &amp; Gutschriften an Apotheken",
     "Forderung an Erloese + USt; Gutschrift entsprechend negativ"),
    ("Kasse / Verkauf", "Verkauf an Apotheken (unbar, B2B)",
     "Verdichtete Erloese je Steuersatz + USt (gegen Debitor/Forderung)"),
    ("Wareneingang / Einkauf", "Beschaffung Inland &amp; EU-Ausland",
     "Wareneingang + Vorsteuer gegen Verbindlichkeit; EU: i.g. Erwerb / §13b"),
    ("Umbuchung Produktion → Vertrieb", "Interner Warenuebergang Produktion an Verkauf",
     "Bestands-/Wertumbuchung zwischen Bereichen – KEINE USt (interner Vorgang)"),
    ("Sonstiges (spaeter)", "Reisekosten, Buerokosten u. ä.",
     "Optionaler Erfassungsdialog – erst bei Bedarf ausbauen"),
]
story.append(table(quellen, ["Quelle / Vorgang", "Inhalt", "Wird zu (Buchung)"],
                   [46 * mm, 56 * mm, 72 * mm]))
story.append(Spacer(1, 3 * mm))
story.append(note_box(
    "Interne Umbuchung Produktion → Vertrieb",
    "Der Warenuebergang von der <b>Produktion</b> (Parallelimport/Konfektionierung) in den "
    "<b>Vertrieb</b> ist <b>kein B2B-Verkauf</b>, sondern eine interne Umbuchung. Wie sie verbucht "
    "wird, haengt an der Firmenstruktur:<br/>"
    "• <b>Eine Firma, zwei Bereiche:</b> reine Bestands-/Wertumbuchung zwischen "
    "Kostenstellen/Lagern – ohne USt.<br/>"
    "• <b>Zwei Betriebe/Gesellschaften:</b> konzerninterne Lieferung mit interner Verrechnung.<br/>"
    "Die App fuehrt <b>Bereich/Betrieb</b> als Dimension mit, damit beide Varianten ohne "
    "Code-Aenderung funktionieren. Welche zutrifft, ist noch offen (siehe Abschnitt 7).",
    bg=colors.HexColor("#EAF2FB"), edge=BLAU))
story.append(Spacer(1, 3 * mm))
story.append(Paragraph(
    "<b>Parallelimport-Besonderheit:</b> Einkauf aus dem EU-Ausland ist kein normaler "
    "19%-Wareneingang, sondern <b>innergemeinschaftlicher Erwerb</b> bzw. faellt unter "
    "<b>Reverse-Charge (§13b UStG)</b>. Dafuer gibt es eigene Steuerschluessel (siehe Abschnitt 3). "
    "Welche Faelle bei euch konkret vorkommen, gehoert auf die Klaerliste fuers Steuerbuero.", st_small))

story.append(PageBreak())

# ============================================================ 2. KONTENLISTEN
story.append(Paragraph("2 · Standard-Kontenlisten (Vorschlag)", st_h1))
story.append(Paragraph(
    "Beide gaengigen Kontenrahmen werden in der App hinterlegt; das Steuerbuero waehlt einen "
    "(meist SKR04 bei GmbH/Grosshandel). Die Konten sind <b>konfigurierbar</b>, nicht fest "
    "verdrahtet. Unten die fuer NMGone relevanten Standardkonten.", st_body))
story.append(Spacer(1, 3 * mm))

story.append(Paragraph("2a · SKR03 (prozessgliederungsorientiert)", st_h2))
skr03 = [
    ("1200", "Bank", "Aktiv / Geld"),
    ("1360", "Geldtransit", "Aktiv / Geld"),
    ("1400", "Forderungen aus Lieferungen &amp; Leistungen (Debitoren)", "Aktiv"),
    ("1576", "Abziehbare Vorsteuer 19 %", "USt"),
    ("1571", "Abziehbare Vorsteuer 7 %", "USt"),
    ("1577", "Abziehbare Vorsteuer i.g. Erwerb 19 %", "USt"),
    ("1600", "Verbindlichkeiten aus Lieferungen &amp; Leistungen (Kreditoren)", "Passiv"),
    ("1776", "Umsatzsteuer 19 %", "USt"),
    ("1771", "Umsatzsteuer 7 %", "USt"),
    ("3200", "Wareneingang 19 % Vorsteuer", "Aufwand"),
    ("3300", "Wareneingang 7 % Vorsteuer", "Aufwand"),
    ("3425", "Innergem. Erwerb 19 % Vorst. u. USt", "Aufwand"),
    ("8400", "Erloese 19 % USt", "Ertrag"),
    ("8300", "Erloese 7 % USt", "Ertrag"),
    ("8125", "Steuerfreie innergem. Lieferung (§4 Nr.1b)", "Ertrag"),
    ("8120", "Steuerfreie Umsaetze §4 Nr.1a (Ausfuhr Drittland)", "Ertrag"),
    ("4660", "Reisekosten Arbeitnehmer", "Aufwand"),
    ("4670", "Reisekosten Unternehmer", "Aufwand"),
    ("4930", "Buerobedarf", "Aufwand"),
]
story.append(table(skr03, ["Konto", "Bezeichnung", "Art"], [16 * mm, 130 * mm, 28 * mm]))
story.append(Spacer(1, 4 * mm))

story.append(Paragraph("2b · SKR04 (abschlussgliederungsorientiert)", st_h2))
skr04 = [
    ("1800", "Bank", "Aktiv / Geld"),
    ("1460", "Geldtransit", "Aktiv / Geld"),
    ("1200", "Forderungen aus Lieferungen &amp; Leistungen (Debitoren)", "Aktiv"),
    ("1406", "Abziehbare Vorsteuer 19 %", "USt"),
    ("1401", "Abziehbare Vorsteuer 7 %", "USt"),
    ("1407", "Abziehbare Vorsteuer i.g. Erwerb 19 %", "USt"),
    ("3300", "Verbindlichkeiten aus Lieferungen &amp; Leistungen (Kreditoren)", "Passiv"),
    ("3806", "Umsatzsteuer 19 %", "USt"),
    ("3801", "Umsatzsteuer 7 %", "USt"),
    ("5200", "Wareneingang 19 % Vorsteuer", "Aufwand"),
    ("5300", "Wareneingang 7 % Vorsteuer", "Aufwand"),
    ("5425", "Innergem. Erwerb 19 % Vorst. u. USt", "Aufwand"),
    ("4400", "Erloese 19 % USt", "Ertrag"),
    ("4300", "Erloese 7 % USt", "Ertrag"),
    ("4125", "Steuerfreie innergem. Lieferung (§4 Nr.1b)", "Ertrag"),
    ("4120", "Steuerfreie Umsaetze §4 Nr.1a (Ausfuhr Drittland)", "Ertrag"),
    ("6650", "Reisekosten Arbeitnehmer", "Aufwand"),
    ("6670", "Reisekosten Unternehmer", "Aufwand"),
    ("6815", "Buerobedarf", "Aufwand"),
]
story.append(table(skr04, ["Konto", "Bezeichnung", "Art"], [16 * mm, 130 * mm, 28 * mm], head_bg=TUERKIS))
story.append(Spacer(1, 2 * mm))
story.append(Paragraph(
    "Kontonummern sind die DATEV-Standardkonten der jeweiligen Rahmen. <b>Verbindlich ist die "
    "Liste, die euer Steuerbuero fuehrt</b> – vor Produktivstart 1:1 abgleichen.", st_small))

story.append(PageBreak())

# ============================================================ 3. USt
story.append(Paragraph("3 · Umsatzsteuer &amp; Steuerschluessel", st_h1))
story.append(Paragraph(
    "Heute deckt NMGone im Einkauf <b>19 %</b> und ggf. <b>7 %</b> ab. Die App haelt die "
    "Steuersaetze als Katalog vor (erweiterbar), damit spaetere Faelle wie Reisekosten, "
    "EU-Erwerb oder §13b ohne Code-Aenderung dazukommen. Im DATEV-Export wird je Buchung "
    "ein <b>Steuerschluessel (BU-Schluessel)</b> mitgegeben.", st_body))
story.append(Spacer(1, 3 * mm))
ust = [
    ("Wareneinkauf Inland 19 %", "19 % Vorsteuer", "9"),
    ("Wareneinkauf Inland 7 %", "7 % Vorsteuer", "8"),
    ("Verkauf / Erloes 19 %", "19 % Umsatzsteuer", "3"),
    ("Verkauf / Erloes 7 %", "7 % Umsatzsteuer", "2"),
    ("Innergem. Erwerb EU 19 %", "Erwerbsteuer + Vorsteuer", "i.g.E. (z.B. 91)"),
    ("Leistung §13b (Reverse-Charge)", "Steuerschuld Leistungsempfaenger", "§13b-Schluessel"),
    ("Steuerfreie innergem. Lieferung", "0 % – mit Nachweis (USt-ID)", "1 / steuerfrei"),
    ("Reisekosten u.ä. (spaeter)", "19 % / 7 % / teils ohne Vorsteuer", "offen"),
]
story.append(table(ust, ["Geschaeftsvorfall", "USt-Behandlung", "DATEV-Schl. (typ.)"],
                   [62 * mm, 70 * mm, 42 * mm]))
story.append(Spacer(1, 3 * mm))
story.append(note_box(
    "Achtung Steuerschluessel",
    "Die angegebenen DATEV-Schluessel (2/3/8/9 …) sind die <b>klassischen Automatikkonten-"
    "Schluessel</b> und dienen nur der Orientierung. Welche Schluessel in eurem Mandanten "
    "tatsaechlich gelten – besonders fuer <b>i.g. Erwerb</b> und <b>§13b</b> – legt das "
    "Steuerbuero fest. Diese Schluessel werden in der App pro Geschaeftsvorfall hinterlegt, "
    "nicht geraten."))

story.append(PageBreak())

# ============================================================ 4. eRECHNUNG
story.append(Paragraph("4 · eRechnung – Empfang &amp; Versand", st_h1))
story.append(Paragraph(
    "Seit <b>01.01.2025</b> muss jedes Unternehmen im inlaendischen B2B-Geschaeft "
    "<b>elektronische Rechnungen empfangen</b> koennen. Eine eRechnung ist ein <b>strukturierter "
    "Datensatz</b> (kein PDF-Bild) nach EN 16931 – in Deutschland vor allem <b>XRechnung</b> (reines "
    "XML) und <b>ZUGFeRD</b> (PDF/A-3 mit eingebettetem XML). Empfang gehoert in die Buchhaltungs-App, "
    "Versand in die <b>Faktura</b>.", st_body))
story.append(Spacer(1, 3 * mm))

story.append(Paragraph("4a · Empfang (Pflicht heute) – in der Buchhaltungs-App", st_h2))
erech = [
    ("Empfangsweg", "Zentrales E-Mail-Postfach (z.B. rechnung@) einlesen / Datei-Import per Drag&amp;Drop"),
    ("Formate lesen", "XRechnung (UBL/CII-XML) und ZUGFeRD (PDF/A-3 mit XML-Anhang) parsen"),
    ("Pruefung", "Pflichtfelder &amp; EN-16931-Validierung; USt-ID, Betraege, Steuer pruefen"),
    ("Vorschau", "Strukturierte Daten lesbar anzeigen (nicht nur XML) + Original archivieren"),
    ("Uebernahme", "Geprueft → als Eingangsrechnung in die Buchung (Kreditor + Vorsteuer)"),
    ("Archiv (GoBD)", "Original-XML/PDF unveraenderbar 10 Jahre aufbewahren, mit Buchung verknuepft"),
]
story.append(table(erech, ["Baustein", "Inhalt"], [40 * mm, 134 * mm]))
story.append(Spacer(1, 4 * mm))

story.append(Paragraph("4b · Versand aus der Faktura (vorbereiten zur Umsetzung)", st_h2))
story.append(Paragraph(
    "Die Faktura erzeugt heute PDF-Rechnungen. Fuer den eRechnung-Versand wird daraus zusaetzlich "
    "ein <b>strukturierter Datensatz</b> erzeugt. Der elegante Weg fuer B2B ist <b>ZUGFeRD</b>: das "
    "gewohnte PDF bleibt sichtbar, das XML steckt eingebettet darin (ein Beleg, beides drin).", st_body))
story.append(Spacer(1, 2 * mm))
esend = [
    ("Format B2B", "ZUGFeRD (PDF/A-3 mit eingebettetem factur-x.xml), Profil EN 16931"),
    ("Format Behoerde", "XRechnung (reines XML) inkl. Leitweg-ID – nur falls an oeffentliche Stellen"),
    ("XML-Erzeugung", "Aus vorhandenen Faktura-Daten EN-16931-XML (CII) erzeugen; Bibliothek z.B. factur-x / drafthorse"),
    ("PDF/A-3", "Bestehendes Rechnungs-PDF PDF/A-3-konform machen + XML einbetten"),
    ("Pflichtfelder", "USt-ID Verkaeufer+Kaeufer, Steueraufschluesselung, Zahlungsbedingungen, Bankverbindung, Leistungsdatum"),
    ("Validierung", "Vor Versand gegen EN-16931-Schematron (KoSIT-Validator) pruefen"),
    ("Kundenkennzeichen", "Je Apotheke: eRechnung ja/nein + Format + Empfangsadresse; Fallback PDF (mit Zustimmung)"),
    ("Versand &amp; Archiv", "Per E-Mail; ausgehende eRechnung revisionssicher 10 Jahre archivieren"),
]
story.append(table(esend, ["Baustein", "Inhalt"], [40 * mm, 134 * mm], head_bg=TUERKIS))
story.append(Spacer(1, 3 * mm))
story.append(note_box(
    "Warum jetzt vorbereiten",
    "Der <b>Empfang</b> ist seit 2025 Pflicht. Beim <b>Versand</b> gelten Uebergangsfristen: ab "
    "<b>01.01.2027</b> muessen Unternehmen mit Vorjahresumsatz &gt; 800.000 € eRechnungen senden, ab "
    "<b>01.01.2028</b> alle. Die Faktura jetzt ZUGFeRD-faehig zu machen, verteilt den Aufwand und "
    "macht NMG-Pharma fristgerecht – statt kurz vor Stichtag.",
    bg=colors.HexColor("#EAF2FB"), edge=BLAU))

story.append(PageBreak())

# ============================================================ 5. PRUEFUNG / GoBD
story.append(Paragraph("5 · Pruefung, Freigabe &amp; GoBD/DATEV", st_h1))
story.append(Paragraph(
    "„Gepruefte Daten“ heisst: vor dem Export laeuft eine Pruefkette, und nur freigegebene "
    "Buchungen verlassen das Haus.", st_body))
story.append(Spacer(1, 3 * mm))

story.append(Paragraph("5a · Pruef- und Freigabe-Workflow", st_h2))
pruef = [
    ("Plausibilitaet", "Soll = Haben, Konto gesetzt, Steuerschluessel gesetzt, Datum im Wirtschaftsjahr"),
    ("Vollstaendigkeit", "Belegnummern lueckenlos &amp; ohne Dubletten; Erloessumme Kasse ↔ verbucht"),
    ("Status-Kette", "Entwurf → geprueft → freigegeben → exportiert (nur freigegebene gehen raus)"),
    ("4-Augen (optional)", "Freigabe durch berechtigte Rolle – ueber die Parameter-/Rechte-App"),
    ("Fehlerliste", "Bildschirm + Druck/Export aller offenen Beanstandungen vor dem Lauf"),
]
story.append(table(pruef, ["Schritt", "Inhalt"], [40 * mm, 134 * mm]))
story.append(Spacer(1, 4 * mm))

story.append(Paragraph("5b · GoBD-Konformitaet &amp; Export", st_h2))
gobd = [
    ("Unveraenderbarkeit", "Freigegebene/exportierte Buchungen read-only; Korrektur nur per Storno"),
    ("Nachvollziehbarkeit", "Audit-Log (wer/wann/was) je Buchung und je Export-Lauf"),
    ("Export-Journal", "Jeder DATEV-Lauf protokolliert: Zeitraum, Satzanzahl, Pruefsumme, Datei-Hash"),
    ("Beleg-Verknuepfung", "Buchung ↔ Original-Beleg (PDF/XML), GoBD-gerecht 10 Jahre archiviert"),
    ("DATEV-Export", "EXTF-CSV (Buchungsstapel), Header mit Berater-/Mandanten-Nr. + Wirtschaftsjahr, CP1252"),
    ("Verfahrensdoku", "GoBD-Verfahrensdokumentation als druckbares PDF (Pflicht bei Pruefung)"),
]
story.append(table(gobd, ["Anforderung", "Umsetzung"], [40 * mm, 134 * mm], head_bg=TUERKIS))

story.append(PageBreak())

# ============================================================ 6. ROADMAP
story.append(Paragraph("6 · Umsetzungs-Roadmap mit Zeitschaetzung", st_h1))
story.append(Paragraph(
    "Zeitschaetzung in <b>Personentagen (PT)</b> inkl. Test und Einbau in die bestehende "
    "Programmfamilie. Reihenfolge = empfohlene Bearbeitung.", st_body))
story.append(Spacer(1, 3 * mm))
roadmap = [
    ("P0", "Fundament &amp; Entscheidungen", "Scope fixieren, Kontenrahmen-Wahl, Berater-/Mandanten-Nr., GoBD-Leitplanken.", "1–2"),
    ("P1", "Datenmodell &amp; Quellen", "Buchungssatz-Modell; Anbindung Faktura, Kasse, Wareneingang; Mapping-Tabelle.", "5–8"),
    ("P1", "Kontenrahmen SKR03/04", "Beide Listen hinterlegen, konfigurierbar, Konto-Auswahl in der UI.", "2–3"),
    ("P1", "USt-/Steuerschluessel-Katalog", "Saetze 19/7 % + i.g.Erwerb/§13b/steuerfrei als pflegbare Liste.", "2–3"),
    ("P2", "Pruef- &amp; Freigabe-Workflow", "Plausibilitaet, Vollstaendigkeit, Status-Kette, Fehlerliste.", "4–6"),
    ("P3", "DATEV-Export (EXTF)", "Buchungsstapel-CSV mit Header, gegen DATEV-Doku validiert.", "4–7"),
    ("P3", "GoBD: Sperre, Audit, Journal", "Unveraenderbarkeit, Audit-Log, Export-Journal, Belegablage.", "4–6"),
    ("P4", "eRechnung-Empfang (Buchhaltung)", "XRechnung/ZUGFeRD lesen, validieren, archivieren, uebernehmen.", "6–10"),
    ("P4", "eRechnung-Versand (Faktura)", "ZUGFeRD/XRechnung aus Faktura erzeugen, KoSIT-validieren, versenden + archivieren.", "6–9"),
    ("P5", "Einbau in App-Familie", "Nav/Kacheln/Dashboard, Demo-Modus, Rechte, Theme, Start-.bat.", "3–5"),
    ("P5", "Hilfe, Handout &amp; Verfahrensdoku", "Hilfe-Kapitel + Screenshots, GoBD-Verfahrensdokumentation-PDF.", "2–4"),
    ("–", "Spaeter: Sonstige Belege", "Reisekosten/Buerokosten-Erfassung, weitere Konten.", "2–4"),
]
story.append(table(roadmap, ["Phase", "Vorhaben", "Inhalt", "PT"],
                   [14 * mm, 46 * mm, 98 * mm, 16 * mm]))
story.append(Spacer(1, 3 * mm))
story.append(Paragraph(
    "<b>Kern (P0–P5, ohne „Spaeter“): rund 39–63 Personentage.</b> Groesster Block ist die "
    "eRechnung (Empfang + Versand, zusammen 12–19 PT). Reine Pflicht-Minimalstrecke (Quellen → "
    "Pruefung → DATEV-Export + GoBD, ohne eRechnung-Versand) liegt bei rund <b>26–40 PT</b>.", st_body))
story.append(Spacer(1, 3 * mm))
story.append(Paragraph(
    "Lesehilfe: <b>klein</b> ≈ 1–2 PT · <b>mittel</b> ≈ 3–8 PT · <b>gross</b> ≈ 15+ PT. "
    "Ein Personentag = ein voller Arbeitstag konzentrierter Entwicklung inkl. Test.", st_small))

story.append(PageBreak())

# ============================================================ 7. OFFENE PUNKTE
story.append(Paragraph("7 · Offene Punkte – mit dem Steuerbuero klaeren", st_h1))
story.append(Paragraph(
    "Diese Punkte sind haftungs- und korrektheitsrelevant und sollten <b>nicht</b> aus dem "
    "Bauch heraus entschieden werden. Sie betreffen direkt Architektur und Konten-Mapping.", st_body))
story.append(Spacer(1, 3 * mm))
offen = [
    ("Firmenstruktur", "Eine Firma (Bereiche Produktion/Vertrieb) oder zwei Betriebe? Entscheidet, ob die Umbuchung USt-relevant ist."),
    ("Kontenrahmen", "SKR03 oder SKR04? (Welche Konten-Liste fuehrt ihr fuer uns?)"),
    ("Stammdaten DATEV", "Berater-Nummer, Mandanten-Nummer, Wirtschaftsjahr / Kontenlaenge."),
    ("Parallelimport-USt", "i.g. Erwerb vs. §13b: welche Faelle, welche Steuerschluessel?"),
    ("Erloes-Verdichtung", "Tageserloese je Steuersatz verdichten – oder Einzelrechnungen je Debitor?"),
    ("Debitoren/Kreditoren", "Personenkonten je Kunde/Lieferant gewuenscht? Nummernkreise?"),
    ("Lieferweg", "DATEV Unternehmen online, Datei-Upload oder E-Mail an die Kanzlei?"),
    ("Buchungstakt", "Monatlich, quartalsweise? Passend zur Umsatzsteuer-Voranmeldung?"),
    ("eRechnung-Archiv", "Reicht euer/unser Archiv, oder schreibt die Kanzlei ein System vor?"),
    ("Reisekosten &amp; Co.", "Sollen weitere Aufwandsarten rein – und mit welchen Konten?"),
]
story.append(table(offen, ["Thema", "Frage an die Kanzlei"], [44 * mm, 130 * mm], head_bg=ROT))
story.append(Spacer(1, 5 * mm))

story.append(HRFlowable(width="100%", thickness=0.8, color=HELLGRAU, spaceAfter=4))
story.append(Paragraph(
    "<b>Hinweis &amp; Vorbehalt:</b> Diese Planung ist eine interne Projektunterlage zur "
    "Abstimmung. Konten, Steuerschluessel und Verfahren sind <b>Vorschlaege auf Basis der "
    "DATEV-Standardkontenrahmen</b> und ersetzen keine Steuer- oder Rechtsberatung. Verbindlich "
    "ist stets die Vorgabe des Steuerbueros. Stand: " + date.today().strftime("%d.%m.%Y") + ".",
    st_small))

# --- Build --------------------------------------------------------------------
os.makedirs(os.path.dirname(OUT), exist_ok=True)
doc = SimpleDocTemplate(
    OUT, pagesize=A4,
    leftMargin=18 * mm, rightMargin=18 * mm,
    topMargin=18 * mm, bottomMargin=16 * mm,
    title="Planung Buchhaltungs-App NMGone", author="NMGone")
doc.build(story, onFirstPage=header_footer, onLaterPages=header_footer)
print("OK ->", OUT)
