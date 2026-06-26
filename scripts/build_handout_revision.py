# -*- coding: utf-8 -*-
"""Erzeugt das Handout zur Revision (Was wurde veraendert) als PDF (reportlab).

Quelle ist app/revision_data.py – dieselbe Quelle wie die Testoberflaeche
app/revision_uebersicht.py. Ausgabe:
    docs/Handout_Revision_Uebersicht.pdf   (dort sucht die Oberflaeche sie)
und – falls vorhanden – zusaetzlich ins Downloads-Verzeichnis.

Aufruf:  python scripts/build_handout_revision.py
"""
from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from app import revision_data as RD  # noqa: E402

from reportlab.lib import colors  # noqa: E402
from reportlab.lib.pagesizes import A4  # noqa: E402
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet  # noqa: E402
from reportlab.lib.units import mm  # noqa: E402
from reportlab.pdfbase import pdfmetrics  # noqa: E402
from reportlab.pdfbase.ttfonts import TTFont  # noqa: E402
from reportlab.platypus import (  # noqa: E402
    BaseDocTemplate, Frame, PageTemplate, Paragraph, Spacer, Table, TableStyle,
    HRFlowable, KeepTogether,
)

OUT = ROOT / "docs" / "Handout_Revision_Uebersicht.pdf"

NAVY = colors.HexColor("#0B4A86")
TEAL = colors.HexColor("#0B6E6E")
RED = colors.HexColor("#B5391F")
GREEN = colors.HexColor("#1E8E5A")
ACCENT = colors.HexColor("#208ACD")
WARN = colors.HexColor("#C88200")
LIGHT = colors.HexColor("#D9EAF7")
CARDBG = colors.HexColor("#F4F6F9")
TIPBG = colors.HexColor("#EAF5EF")
WARNBG = colors.HexColor("#FCEFEA")
GREYTX = colors.HexColor("#444444")
GREYLT = colors.HexColor("#777777")
RULE = colors.HexColor("#C7D6E6")

# Schriften (Arial fuer saubere Umlaute + Euro/Paragraf)
try:
    pdfmetrics.registerFont(TTFont("AppFont", "C:/Windows/Fonts/arial.ttf"))
    pdfmetrics.registerFont(TTFont("AppFont-Bold", "C:/Windows/Fonts/arialbd.ttf"))
    pdfmetrics.registerFont(TTFont("AppFont-It", "C:/Windows/Fonts/ariali.ttf"))
    BASE, BOLD, ITAL = "AppFont", "AppFont-Bold", "AppFont-It"
except Exception:
    BASE, BOLD, ITAL = "Helvetica", "Helvetica-Bold", "Helvetica-Oblique"

styles = getSampleStyleSheet()


def S(name, **kw):
    return ParagraphStyle(name, parent=styles["Normal"], **kw)


H1 = S("H1", fontName=BOLD, fontSize=19, textColor=NAVY, leading=23, spaceAfter=2)
SUB = S("SUB", fontName=BASE, fontSize=10.5, textColor=GREYLT, leading=14, spaceAfter=6)
H2 = S("H2", fontName=BOLD, fontSize=13.5, textColor=NAVY, leading=17,
       spaceBefore=12, spaceAfter=4)
BODY = S("BODY", fontName=BASE, fontSize=10, textColor=GREYTX, leading=14.5, spaceAfter=3)
BODYW = S("BODYW", fontName=BASE, fontSize=10, textColor=colors.white, leading=14.5)
ITEMT = S("ITEMT", fontName=BOLD, fontSize=10, textColor=colors.HexColor("#1F2933"), leading=13)
ITEMD = S("ITEMD", fontName=BASE, fontSize=9.5, textColor=GREYTX, leading=13)
SMALL = S("SMALL", fontName=BASE, fontSize=8.5, textColor=GREYLT, leading=11)
MONO = S("MONO", fontName=BASE, fontSize=8.5, textColor=NAVY, leading=11)
KZAHL = S("KZAHL", fontName=BOLD, fontSize=22, textColor=NAVY, leading=24)
BADGE = S("BADGE", fontName=BOLD, fontSize=8, textColor=colors.white, leading=10,
          alignment=1)

story = []


def hr(space_before=4, space_after=6):
    story.append(Spacer(1, space_before))
    story.append(HRFlowable(width="100%", thickness=0.7, color=RULE))
    story.append(Spacer(1, space_after))


def bullet_item(titel, detail, color=NAVY):
    """Ein Punkt: farbiges Quadrat + fetter Titel + Beschreibung darunter."""
    inner = [Paragraph(titel, ITEMT)]
    if detail:
        inner.append(Paragraph(detail, ITEMD))
    dot = Table([[""]], colWidths=[3.2 * mm], rowHeights=[3.2 * mm])
    dot.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), color),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
    ]))
    t = Table([[dot, inner]], colWidths=[7 * mm, 158 * mm])
    t.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ("LEFTPADDING", (0, 0), (0, 0), 2),
    ]))
    return t


def badge(text, color):
    b = Table([[Paragraph(text, BADGE)]], colWidths=[26 * mm])
    b.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), color),
        ("TOPPADDING", (0, 0), (-1, -1), 2),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
    ]))
    return b


# ── Kopf ─────────────────────────────────────────────────────────────────────
story.append(Paragraph(RD.REVISION["titel"], H1))
story.append(Paragraph(
    f'{RD.REVISION["untertitel"]} &nbsp;·&nbsp; Stand {RD.REVISION["datum"]} '
    f'&nbsp;·&nbsp; {RD.REVISION["version"]}', SUB))
story.append(Paragraph(RD.REVISION["claim"], BODY))
hr()

# ── Kennzahlen ───────────────────────────────────────────────────────────────
kz_cells = []
for zahl, titel, sub in RD.REVISION["kennzahlen"]:
    cell = [Paragraph(zahl, KZAHL),
            Paragraph(f"<b>{titel}</b>", ITEMT),
            Paragraph(sub, SMALL)]
    kz_cells.append(cell)
kz = Table([kz_cells], colWidths=[42 * mm] * 4)
kz.setStyle(TableStyle([
    ("BACKGROUND", (0, 0), (-1, -1), CARDBG),
    ("BOX", (0, 0), (-1, -1), 0.5, RULE),
    ("INNERGRID", (0, 0), (-1, -1), 0.5, colors.white),
    ("VALIGN", (0, 0), (-1, -1), "TOP"),
    ("LEFTPADDING", (0, 0), (-1, -1), 8),
    ("RIGHTPADDING", (0, 0), (-1, -1), 8),
    ("TOPPADDING", (0, 0), (-1, -1), 8),
    ("BOTTOMPADDING", (0, 0), (-1, -1), 10),
]))
story.append(kz)

# ── Programm-Familie ─────────────────────────────────────────────────────────
story.append(Paragraph("Die Programm-Familie", H2))
story.append(Paragraph(
    "Eigenstaendige Apps – jede mit eigenem Taskleisten-Symbol – rund um eine "
    "gemeinsame Datenbank. Status: <b>NEU</b> = in dieser Revision dazugekommen, "
    "<b>ERWEITERT</b> = angefasst, <b>STABIL</b> = unveraendert mitgelaufen.", BODY))
story.append(Spacer(1, 4))

STATUS_COLOR = {"neu": GREEN, "erweitert": ACCENT, "stabil": GREYLT}
rows = [[Paragraph("<b>App</b>", ITEMD), Paragraph("<b>Status</b>", ITEMD),
         Paragraph("<b>Zweck &amp; Neuerung</b>", ITEMD)]]
for app in RD.APPS:
    st = app["status"]
    rows.append([
        Paragraph(f'{app["icon"]} <b>{app["name"]}</b><br/>'
                  f'<font size=7 color="#777777">{app["start"]}</font>', ITEMD),
        badge(st.upper(), STATUS_COLOR.get(st, GREYLT)),
        Paragraph(f'{app["zweck"]}<br/><font color="#0B4A86"><b>Neu:</b></font> '
                  f'{app["neu"]}', ITEMD),
    ])
tbl = Table(rows, colWidths=[44 * mm, 28 * mm, 96 * mm], repeatRows=1)
tbl.setStyle(TableStyle([
    ("BACKGROUND", (0, 0), (-1, 0), LIGHT),
    ("GRID", (0, 0), (-1, -1), 0.5, RULE),
    ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ("VALIGN", (2, 1), (2, -1), "TOP"),
    ("LEFTPADDING", (0, 0), (-1, -1), 6),
    ("RIGHTPADDING", (0, 0), (-1, -1), 6),
    ("TOPPADDING", (0, 0), (-1, -1), 5),
    ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
    ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, CARDBG]),
]))
story.append(tbl)

# ── Was wurde veraendert ─────────────────────────────────────────────────────
story.append(Paragraph("Was wurde veraendert", H2))
for kategorie, eintraege in RD.AENDERUNGEN:
    block = [Paragraph(kategorie, S("KAT", fontName=BOLD, fontSize=11,
                                    textColor=TEAL, leading=14, spaceBefore=6, spaceAfter=2))]
    for titel, detail in eintraege:
        block.append(bullet_item(titel, detail, color=TEAL))
    story.append(KeepTogether(block))

# ── Aufraeumen ───────────────────────────────────────────────────────────────
story.append(Paragraph("Aufgeraeumt &amp; geprueft", H2))
done_rows = [[Paragraph("✓ " + t, S("DONE", fontName=BASE, fontSize=9.5,
                                    textColor=colors.HexColor("#0D6630"), leading=13))]
             for t in RD.AUFGERAEUMT]
done = Table(done_rows, colWidths=[168 * mm])
done.setStyle(TableStyle([
    ("BACKGROUND", (0, 0), (-1, -1), TIPBG),
    ("BOX", (0, 0), (-1, -1), 0.5, GREEN),
    ("LEFTPADDING", (0, 0), (-1, -1), 8),
    ("TOPPADDING", (0, 0), (-1, -1), 3),
    ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
]))
story.append(done)
story.append(Spacer(1, 6))
story.append(Paragraph(
    "<b>Empfohlen – aber bewusst NICHT automatisch geloescht.</b> Diese Dinge liegen "
    "nur lokal herum (alle in .gitignore). Du entscheidest, wenn du zurueck bist:", BODY))
rec_rows = [[Paragraph("<b>Pfad</b>", ITEMD), Paragraph("<b>Groesse</b>", ITEMD),
             Paragraph("<b>Empfehlung</b>", ITEMD)]]
for pfad, groesse, empf in RD.AUFRAEUM_EMPFEHLUNG:
    rec_rows.append([Paragraph(pfad, MONO), Paragraph(groesse, ITEMD),
                     Paragraph(empf, ITEMD)])
rec = Table(rec_rows, colWidths=[46 * mm, 22 * mm, 100 * mm], repeatRows=1)
rec.setStyle(TableStyle([
    ("BACKGROUND", (0, 0), (-1, 0), WARNBG),
    ("GRID", (0, 0), (-1, -1), 0.5, RULE),
    ("VALIGN", (0, 0), (-1, -1), "TOP"),
    ("LEFTPADDING", (0, 0), (-1, -1), 6),
    ("TOPPADDING", (0, 0), (-1, -1), 4),
    ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
]))
story.append(rec)

# ── Roadmap ──────────────────────────────────────────────────────────────────
story.append(Paragraph("Naechste Revolution – der Fahrplan", H2))
AUF_COLOR = {"klein": GREEN, "mittel": WARN, "gross": RED}
AUF_TXT = {"klein": "klein", "mittel": "mittel", "gross": "gross"}
for i, (titel, beschr, aufwand) in enumerate(RD.ROADMAP, start=1):
    head = Table([[
        Paragraph(f'<b>{i}. {titel}</b>', S("RT", fontName=BOLD, fontSize=10.5,
                                            textColor=colors.HexColor("#1F2933"), leading=13)),
        badge(AUF_TXT.get(aufwand, "mittel"), AUF_COLOR.get(aufwand, WARN)),
    ]], colWidths=[140 * mm, 28 * mm])
    head.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                              ("LEFTPADDING", (0, 0), (-1, -1), 0)]))
    story.append(KeepTogether([head, Paragraph(beschr, ITEMD), Spacer(1, 5)]))

hr()
story.append(Paragraph(
    "Diese Uebersicht gibt es auch als Fenster: <b>„Revisions-Uebersicht starten.bat“</b> "
    "(oder <font name='%s'>python start_revision.py</font>). Inhalt und PDF stammen aus "
    "derselben Quelle (app/revision_data.py) – einmal pflegen, beides aktuell." % BASE, SMALL))


# ── Seitenrahmen ─────────────────────────────────────────────────────────────
def _decorate(canvas, doc):
    canvas.saveState()
    canvas.setFillColor(NAVY)
    canvas.rect(0, A4[1] - 8 * mm, A4[0], 8 * mm, fill=1, stroke=0)
    canvas.setFillColor(GREYLT)
    canvas.setFont(BASE, 8)
    canvas.drawString(18 * mm, 10 * mm, "NMGone · Revisions-Uebersicht")
    canvas.drawRightString(A4[0] - 18 * mm, 10 * mm, f"Seite {doc.page}")
    canvas.restoreState()


def build():
    OUT.parent.mkdir(parents=True, exist_ok=True)
    doc = BaseDocTemplate(
        str(OUT), pagesize=A4,
        leftMargin=18 * mm, rightMargin=18 * mm,
        topMargin=16 * mm, bottomMargin=16 * mm,
        title="NMGone – Revisions-Uebersicht", author="NMGone")
    frame = Frame(doc.leftMargin, doc.bottomMargin,
                  doc.width, doc.height, id="main")
    doc.addPageTemplates([PageTemplate(id="t", frames=[frame], onPage=_decorate)])
    doc.build(story)
    print("PDF erzeugt:", OUT)

    # Komfort: zusaetzlich nach Downloads kopieren, falls vorhanden
    dl = Path.home() / "Downloads"
    if dl.is_dir():
        import shutil
        try:
            shutil.copy2(OUT, dl / OUT.name)
            print("Kopie:", dl / OUT.name)
        except Exception as exc:
            print("Downloads-Kopie uebersprungen:", exc)


if __name__ == "__main__":
    build()
