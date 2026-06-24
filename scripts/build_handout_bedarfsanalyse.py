# -*- coding: utf-8 -*-
"""Erzeugt das Anwender-Handout zur Bedarfsanalyse als PDF (reportlab)."""

from pathlib import Path

from reportlab.lib import colors
from reportlab.lib.enums import TA_LEFT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (
    BaseDocTemplate, Frame, PageTemplate, Paragraph, Spacer, Table, TableStyle,
    ListFlowable, ListItem, HRFlowable, KeepTogether,
)

OUT = Path(r"C:/Users/USER/Downloads/Handout_Bedarfsanalyse.pdf")

NAVY = colors.HexColor("#0B4A86")
LIGHT = colors.HexColor("#D9EAF7")
GREEN = colors.HexColor("#1E8E5A")
GREYTX = colors.HexColor("#444444")
GREYLT = colors.HexColor("#777777")
RULE = colors.HexColor("#C7D6E6")
TIPBG = colors.HexColor("#EAF5EF")
WARNBG = colors.HexColor("#FCEFEA")

# Schriften (Arial fuer saubere Umlaute + Euro-Zeichen)
try:
    pdfmetrics.registerFont(TTFont("AppFont", "C:/Windows/Fonts/arial.ttf"))
    pdfmetrics.registerFont(TTFont("AppFont-Bold", "C:/Windows/Fonts/arialbd.ttf"))
    pdfmetrics.registerFont(TTFont("AppFont-It", "C:/Windows/Fonts/ariali.ttf"))
    BASE, BOLD, ITAL = "AppFont", "AppFont-Bold", "AppFont-It"
except Exception:
    BASE, BOLD, ITAL = "Helvetica", "Helvetica-Bold", "Helvetica-Oblique"

styles = getSampleStyleSheet()

def S(name, **kw):
    kw.setdefault("fontName", BASE)
    return ParagraphStyle(name, parent=styles["Normal"], **kw)

body = S("body", fontSize=10, leading=15, textColor=GREYTX, spaceAfter=5)
body_b = S("body_b", fontSize=10, leading=15, textColor=GREYTX, fontName=BOLD)
h1 = S("h1", fontSize=15, leading=19, textColor=NAVY, fontName=BOLD, spaceBefore=14, spaceAfter=3)
h2 = S("h2", fontSize=11.5, leading=15, textColor=NAVY, fontName=BOLD, spaceBefore=9, spaceAfter=3)
small = S("small", fontSize=8.5, leading=12, textColor=GREYLT)
cell = S("cell", fontSize=9, leading=12.5, textColor=GREYTX)
cell_b = S("cell_b", fontSize=9, leading=12.5, textColor=colors.white, fontName=BOLD)
tip_t = S("tip_t", fontSize=9.5, leading=14, textColor=GREYTX)
title_st = S("title_st", fontSize=26, leading=30, textColor=NAVY, fontName=BOLD)
subtitle_st = S("subtitle_st", fontSize=13, leading=17, textColor=GREEN, fontName=BOLD)


def bullets(items, gap=3):
    li = [ListItem(Paragraph(t, body), leftIndent=6, value="•") for t in items]
    return ListFlowable(li, bulletType="bullet", bulletColor=NAVY, bulletFontSize=9,
                        leftIndent=12, spaceBefore=1, spaceAfter=gap)


def heading(text):
    return KeepTogether([Paragraph(text, h1),
                         HRFlowable(width="100%", thickness=1.2, color=RULE,
                                    spaceBefore=2, spaceAfter=6)])


def callout(title, text, bg=TIPBG, bar=GREEN):
    inner = [Paragraph(f'<b>{title}</b>', S("ct", fontSize=9.5, leading=14, textColor=bar, fontName=BOLD)),
             Paragraph(text, tip_t)]
    t = Table([[inner]], colWidths=[165 * mm])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), bg),
        ("LINEBEFORE", (0, 0), (0, -1), 3, bar),
        ("LEFTPADDING", (0, 0), (-1, -1), 9),
        ("RIGHTPADDING", (0, 0), (-1, -1), 9),
        ("TOPPADDING", (0, 0), (-1, -1), 7),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
    ]))
    return t


def datatable(header, rows, colw):
    data = [[Paragraph(h, cell_b) for h in header]]
    for r in rows:
        data.append([Paragraph(str(c), cell) for c in r])
    t = Table(data, colWidths=colw, repeatRows=1)
    st = [
        ("BACKGROUND", (0, 0), (-1, 0), NAVY),
        ("GRID", (0, 0), (-1, -1), 0.5, RULE),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 5),
        ("RIGHTPADDING", (0, 0), (-1, -1), 5),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F4F8FC")]),
    ]
    t.setStyle(TableStyle(st))
    return t


def steps(items):
    data = []
    for i, (t, d) in enumerate(items, 1):
        num = Paragraph(str(i), S("num", fontSize=11, leading=20, textColor=colors.white, fontName=BOLD, alignment=1))
        txt = [Paragraph(f"<b>{t}</b>", body_b), Paragraph(d, body)]
        data.append([num, txt])
    t = Table(data, colWidths=[9 * mm, 156 * mm])
    style = [("VALIGN", (0, 0), (-1, -1), "TOP"),
             ("TOPPADDING", (0, 0), (-1, -1), 4), ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
             ("LEFTPADDING", (0, 0), (0, -1), 0), ("RIGHTPADDING", (0, 0), (0, -1), 6)]
    for r in range(len(items)):
        style.append(("BACKGROUND", (0, r), (0, r), NAVY))
    t.setStyle(TableStyle(style))
    return t


# ---------------------------------------------------------------- Seitenrahmen
def on_page(canvas, doc):
    canvas.saveState()
    # Kopfband
    canvas.setFillColor(NAVY)
    canvas.rect(0, A4[1] - 12 * mm, A4[0], 12 * mm, fill=1, stroke=0)
    canvas.setFillColor(colors.white)
    canvas.setFont(BOLD, 9)
    canvas.drawString(18 * mm, A4[1] - 8 * mm, "NMGone  ·  Bedarfsanalyse")
    canvas.setFont(BASE, 8)
    canvas.drawRightString(A4[0] - 18 * mm, A4[1] - 8 * mm, "Anwender-Handout")
    # Fusszeile
    canvas.setStrokeColor(RULE)
    canvas.line(18 * mm, 14 * mm, A4[0] - 18 * mm, 14 * mm)
    canvas.setFillColor(GREYLT)
    canvas.setFont(BASE, 8)
    canvas.drawString(18 * mm, 10 * mm, "NMG Pharma · Bedarfsanalyse – Anwender-Handout")
    canvas.drawRightString(A4[0] - 18 * mm, 10 * mm, f"Seite {doc.page}")
    canvas.restoreState()


doc = BaseDocTemplate(
    str(OUT), pagesize=A4,
    leftMargin=18 * mm, rightMargin=18 * mm, topMargin=18 * mm, bottomMargin=18 * mm,
    title="Bedarfsanalyse – Anwender-Handout", author="NMG Pharma",
)
frame = Frame(doc.leftMargin, doc.bottomMargin, doc.width, doc.height - 4 * mm, id="main")
doc.addPageTemplates([PageTemplate(id="all", frames=[frame], onPage=on_page)])

story = []
A = story.append

# ----------------------------------------------------------------- Titelblock
A(Spacer(1, 6 * mm))
A(Paragraph("Bedarfsanalyse", title_st))
A(Paragraph("Anwender-Handout – Auswertung, Wirkstoff-Austausch, Kurzbericht", subtitle_st))
A(Spacer(1, 3 * mm))
A(Paragraph(
    "Dieses Handout beschreibt, wie eine Bedarfsanalyse für eine Apotheke erstellt wird, "
    "welche Dateien eingelesen werden können, wie die Ergebnisse aufgebaut sind "
    "(Detail-Auswertung, Wirkstoff-Blatt, Kurzbericht) und wie die Biosimilar-Wissensbasis "
    "(Gelbe Liste) gepflegt wird.", body))
A(HRFlowable(width="100%", thickness=1.2, color=RULE, spaceBefore=6, spaceAfter=2))

# 1 -------------------------------------------------------------------------
A(heading("1. Wozu dient die Bedarfsanalyse?"))
A(Paragraph(
    "Die Bedarfsanalyse liest die Abverkaufs- bzw. Verbrauchsdaten einer Apotheke ein und "
    "gleicht jede PZN mit dem NMG-Sortiment ab. Sie zeigt, welche Artikel auf NMG-Produkte "
    "umgestellt werden können, welches Einsparpotenzial daraus entsteht und welche "
    "wirkstoffgleichen NMG-Alternativen es gibt.", body))
A(Paragraph("Aus einer Auswertung entstehen drei Bausteine:", body))
A(bullets([
    "<b>Detail-Auswertung</b> (Excel) – jede PZN mit NMG-Treffer, Rabatt und Umsatz.",
    "<b>Tabellenblatt „Wirkstoff &amp; NMG-Austausch“</b> – Wirkstoff je Artikel und passende NMG-Produkte.",
    "<b>Kurzbericht</b> (Excel + PDF) – verdichtete Management-Übersicht mit Diagrammen für das Gespräch.",
]))

# 2 -------------------------------------------------------------------------
A(heading("2. Schritt für Schritt: Auswertung erstellen"))
A(steps([
    ("Bedarfsanalyse öffnen", "In der linken Navigation auf „Bedarfsanalyse“ klicken."),
    ("Kundentyp wählen", "Partnerkondition (PK) oder Zukunftswerk (ZW). Die Auswahl wird gemerkt."),
    ("Kundennummer (optional)", "Kann leer bleiben – die Zuordnung ist auch nach der Auswertung möglich."),
    ("Kundenname / Apotheke", "z. B. „Rosen Apotheke Forst“."),
    ("Auswertungsname", "Name für den Ausgabeordner und die Analyse."),
    ("Rohdaten auswählen", "Die Abverkaufs-/Verbrauchsdatei der Apotheke laden (Formate siehe Punkt 3)."),
    ("Auswertungsvorlage (optional)", "Optisches Layout der Ausgabe; kann auch leer bleiben."),
    ("Auswertung starten", "Die Analyse läuft im Hintergrund; am Ende öffnet sich der Ergebnis-Ordner."),
]))

# 3 -------------------------------------------------------------------------
A(heading("3. Welche Dateien können eingelesen werden?"))
A(datatable(
    ["Format", "Hinweis"],
    [
        [".xlsx / .xlsm", "Modernes Excel."],
        [".xls", "Auch „unechte“ .xls aus Apothekensoftware (intern XML- oder HTML-Tabellen) werden gelesen."],
        [".csv / .txt", "Trennzeichen (; , Tab |) und Umlaut-Kodierung werden automatisch erkannt."],
    ],
    [38 * mm, 127 * mm]))
A(Spacer(1, 3))
A(Paragraph("<b>Pflichtspalten:</b> PZN und eine Mengen-/Absatzspalte. Erkannt werden u. a. "
            "„Packungen“, „Abverkäufe 6 Monate“, „Verkaufsmenge“, „Verbrauch“, „Menge“, „Anzahl“. "
            "Optional: Artikelname, DF, Pck, Hersteller, EK.", body))
A(callout("Format nicht erkannt?",
          "Bei einem nicht unterstützten Dateiformat kommt sofort eine Meldung. Sind nur die "
          "Spalten unklar, öffnet sich der Format-Assistent, in dem PZN- und Mengenspalte "
          "manuell zugeordnet werden können.", WARNBG, colors.HexColor("#C0392B")))

# 4 -------------------------------------------------------------------------
A(heading("4. Mehrfache PZN – Mengen summieren"))
A(Paragraph(
    "Verkaufs- bzw. Transaktionslisten enthalten dieselbe PZN oft mehrfach (jede Zeile ist ein "
    "einzelner Verkauf), teilweise mit negativen Mengen (Retouren). Beim Start erkennt das "
    "Programm das und fragt nach:", body))
A(bullets([
    "<b>Ja (empfohlen bei Verkaufslisten):</b> Je PZN entsteht eine Zeile, die Mengen werden "
    "addiert (Retouren werden gegengerechnet). Das ist schneller und liefert den echten Bedarf.",
    "<b>Nein:</b> Jede Zeile bleibt einzeln stehen (wie in der Rohdatei).",
]))
A(callout("Tipp bei großen Listen",
          "Eine Liste mit z. B. 25.000 Verkaufszeilen wird durch das Summieren auf wenige tausend "
          "eindeutige PZN verdichtet – die Auswertung läuft dadurch deutlich schneller (rund 30 "
          "Sekunden statt mehreren Minuten)."))

# 5 -------------------------------------------------------------------------
A(heading("5. Wo liegen die Ergebnisse?"))
A(Paragraph(
    "Für jede Auswertung wird ein eigener Kundenordner angelegt (getrennt nach PK / ZW). "
    "Er öffnet sich nach Abschluss automatisch und enthält:", body))
A(bullets([
    "die <b>Detail-Auswertung</b> (Excel, inkl. Tabellenblatt „Wirkstoff &amp; NMG-Austausch“),",
    "eine <b>Kopie der Rohdaten</b>,",
    "den <b>Kurzbericht</b> als Excel und als PDF.",
]))
A(Paragraph("<i>Der Kurzbericht wird unmittelbar nach der Auswertung im Hintergrund erstellt "
            "und liegt wenige Sekunden später im selben Ordner.</i>",
            S("it", fontSize=9, leading=13, textColor=GREYLT, fontName=ITAL)))

# 6 -------------------------------------------------------------------------
A(heading("6. Die Detail-Auswertung – Spalten"))
A(datatable(
    ["Spalte", "Bedeutung"],
    [
        ["PZN / Artikelname / DF / Pck / Herst / EK", "Stammdaten des abgegebenen Artikels (aus Rohdaten, ergänzt aus dem Stamm)."],
        ["Abverkäufe (6/12 Monate)", "Absatz-/Verbrauchsmenge im Zeitraum."],
        ["im Sortiment", "„X“ = NMG-Treffer bzw. Austausch möglich."],
        ["PZN NMG / APU NMG", "NMG-Artikel, auf den umgestellt werden kann, samt Apothekeneinkaufspreis."],
        ["NMG Rabatt", "Rabattsatz auf den NMG-Artikel."],
        ["lieferbar / Bevorratung / Liefervorschlag", "Lieferfähigkeit und Empfehlung."],
        ["austauschbar gegen", "Alternativen – bleibt leer, wenn bereits eine „PZN NMG“ vorhanden ist (Details im Wirkstoff-Blatt)."],
        ["NMG Rabatt in Euro", "APU × Rabattsatz (je Packung)."],
        ["NMG Rabatt Gesamt nach Absatz", "Rabatt in Euro × Menge = Einsparpotenzial."],
        ["Umsatz", "EK × Menge."],
    ],
    [55 * mm, 110 * mm]))

# 7 -------------------------------------------------------------------------
A(heading("7. Tabellenblatt „Wirkstoff & NMG-Austausch“"))
A(Paragraph(
    "Das zweite Tabellenblatt listet je abgegebenem Artikel den Wirkstoff und die passenden "
    "austauschbaren NMG-Artikel auf. So ist auf einen Blick erkennbar, welche NMG-Alternative "
    "für einen Wirkstoff in Frage kommt.", body))
A(datatable(
    ["Spalte", "Inhalt"],
    [
        ["PZN / Artikel (abgegeben)", "Der ausgewertete Artikel aus den Rohdaten."],
        ["Wirkstoff / Stärke", "Wirkstoff des abgegebenen Artikels."],
        ["Austauschbare NMG-Artikel", "NMG-Produkte mit demselben Wirkstoff (PZN – Name), ggf. mehrere."],
    ],
    [55 * mm, 110 * mm]))
A(Spacer(1, 2))
A(Paragraph("<b>Beispiele:</b> Adalimumab → YUFLYMA · Infliximab → REMSIMA · Ustekinumab → STELARA · "
            "Pegfilgrastim → PELGRAZ.", small))

# 8 -------------------------------------------------------------------------
A(heading("8. Der Kurzbericht (Excel + PDF)"))
A(Paragraph("Verdichtete Übersicht für den Apotheken-Inhaber bzw. das Vertriebsgespräch. Enthält:", body))
A(bullets([
    "<b>Einsparpotenzial – 6 und 12 Monate</b> nebeneinander (aus dem Datenzeitraum hochgerechnet).",
    "<b>Top-10 Hebel-Artikel</b> mit der größten Einsparung (Balkendiagramm).",
    "<b>Umstellungsgrad</b> (Donut) – Anteil umstellbarer Positionen.",
    "<b>Kennzahlen:</b> Positionen gesamt, davon umstellbar, Gesamtumsatz.",
]))

# 9 -------------------------------------------------------------------------
A(heading("9. Berechnungslogik (Kurzüberblick)"))
A(datatable(
    ["Kennzahl", "Formel"],
    [
        ["NMG Rabatt in Euro", "APU × Rabattsatz"],
        ["Einsparpotenzial (Rabatt gesamt)", "NMG Rabatt in Euro × Menge"],
        ["Umsatz", "EK × Menge"],
        ["Jahres-Hochrechnung", "Einsparung pro Monat × 12"],
    ],
    [70 * mm, 95 * mm]))

# 10 ------------------------------------------------------------------------
A(heading("10. Gelbe Liste (Biosimilars) importieren"))
A(Paragraph(
    "Über <b>Daten aktualisieren → Kachel „Gelbe Liste (Biosimilars)“ → Import</b> wird die "
    "Biosimilar-Wissensbasis eingelesen. Die Excel hat drei Spalten:", body))
A(datatable(
    ["Wirkstoff", "Arzneimittel", "Referenzprodukt"],
    [["Adalimumab", "Amgevita / Hyrimoz / Idacio …", "Humira"],
     ["Etanercept", "Benepali / Erelzi …", "Enbrel"]],
    [45 * mm, 75 * mm, 45 * mm]))
A(Spacer(1, 3))
A(Paragraph(
    "Nach dem Import zeigt eine Meldung, wie viele Wirkstoff-Gruppen und Produkte erkannt und wie "
    "viele PZN automatisch zugeordnet wurden. Die Datei kann jederzeit erneut hochgeladen werden, "
    "um zu aktualisieren.", body))
A(callout("Gut zu wissen",
          "Das Wirkstoff-Blatt der Bedarfsanalyse funktioniert bereits ohne die Gelbe Liste. Der "
          "Import erweitert künftig den Biosimilar-Abgleich (Original ↔ Biosimilar-Gruppen)."))

# 11 ------------------------------------------------------------------------
A(heading("11. Hinweise & häufige Fragen"))
A(bullets([
    "<b>Oberfläche ruckelt während der Auswertung?</b> Bei sehr großen Dateien rechnet das "
    "Programm im Hintergrund – kurzes Ruckeln ist normal, ein hartes Einfrieren nicht mehr.",
    "<b>Umsatz ist 0?</b> Dann fehlt der EK in den Rohdaten; er wird – sofern vorhanden – aus den "
    "Stammdaten ergänzt.",
    "<b>„austauschbar gegen“ ist leer?</b> Korrekt, wenn bereits eine „PZN NMG“ steht – die "
    "Austausch-Details stehen im Wirkstoff-Blatt.",
    "<b>Datei wird abgelehnt?</b> Nur .xlsx, .xlsm, .xls, .csv, .txt sind erlaubt; bei unklaren "
    "Spalten hilft der Format-Assistent.",
]))

doc.build(story)
print("OK ->", OUT, OUT.stat().st_size, "bytes")
