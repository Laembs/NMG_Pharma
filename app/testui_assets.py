"""Erzeugt die Hilfe-Bilder fuer die Testoberflaeche programmatisch (PIL).

Bewusst zur Laufzeit generiert statt als Binaerdatei im Repo: bleibt
versionierbar, immer konsistent zur Palette und ohne Asset-Pflege.
Aufruf: ensure_help_images(zielordner) -> dict[name, pfad].
"""
from __future__ import annotations

from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

# ── Palette (an die bestehende NMGone-Optik angelehnt) ───────────────────────
PRIMARY = (11, 74, 134)      # #0B4A86
ACCENT = (32, 138, 205)      # #208ACD
SUCCESS = (17, 130, 59)      # #11823B
WARNING = (200, 130, 0)      # #C88200
INK = (31, 41, 51)           # #1F2933
MUTED = (98, 110, 125)       # #626E7D
BG = (244, 246, 249)         # #F4F6F9
CARD = (255, 255, 255)
BORDER = (220, 226, 232)


def _font(size: int, bold: bool = False) -> ImageFont.FreeTypeFont:
    candidates = (
        ["segoeuib.ttf", "arialbd.ttf", "DejaVuSans-Bold.ttf"] if bold
        else ["segoeui.ttf", "arial.ttf", "DejaVuSans.ttf"]
    )
    for name in candidates:
        try:
            return ImageFont.truetype(name, size)
        except Exception:
            continue
    return ImageFont.load_default()


def _round_rect(draw, box, radius, fill=None, outline=None, width=1):
    draw.rounded_rectangle(box, radius=radius, fill=fill, outline=outline, width=width)


def _text_center(draw, cx, cy, text, font, fill):
    l, t, r, b = draw.textbbox((0, 0), text, font=font)
    draw.text((cx - (r - l) / 2, cy - (b - t) / 2), text, font=font, fill=fill)


def _draw_check(draw, x, y, size, color, width=3):
    """Zeichnet einen Haken als Vektor (zuverlaessiger als ein Glyph)."""
    draw.line([(x, y + size * 0.55), (x + size * 0.38, y + size)], fill=color, width=width)
    draw.line([(x + size * 0.38, y + size), (x + size, y)], fill=color, width=width)


def _shadow_card(img, box, radius=18):
    """Weiche Schlagschatten-Karte fuer einen modernen, professionellen Look."""
    x0, y0, x1, y1 = box
    shadow = Image.new("RGBA", img.size, (0, 0, 0, 0))
    sd = ImageDraw.Draw(shadow)
    sd.rounded_rectangle((x0 + 3, y0 + 6, x1 + 3, y1 + 6), radius=radius, fill=(15, 30, 50, 40))
    try:
        from PIL import ImageFilter
        shadow = shadow.filter(ImageFilter.GaussianBlur(7))
    except Exception:
        pass
    img.alpha_composite(shadow)


def _new_canvas(w, h):
    img = Image.new("RGBA", (w, h), BG + (255,))
    return img, ImageDraw.Draw(img)


def _save(img, path: Path):
    img.convert("RGB").save(path, "PNG")


# ── Bild 1: Prozess-Flow der Neuen Auswertung ────────────────────────────────
def _draw_prozess_flow(path: Path):
    W, H = 1100, 360
    img, d = _new_canvas(W, H)
    _text_center(d, W // 2, 44, "So entsteht eine Auswertung", _font(34, True), PRIMARY)
    _text_center(d, W // 2, 84, "Fünf Schritte – vollautomatisch im Hintergrund", _font(19), MUTED)

    steps = [
        ("1", "Datei wählen", "xlsx · xls · csv", ACCENT),
        ("2", "Einlesen", "Spalten erkennen", ACCENT),
        ("3", "Abgleich", "PZN · Biosimilar", PRIMARY),
        ("4", "Berechnen", "Absatz · Umsatz", PRIMARY),
        ("5", "Fertig", "Excel + Ordner", SUCCESS),
    ]
    n = len(steps)
    card_w, card_h, gap = 170, 150, 30
    total = n * card_w + (n - 1) * gap
    x = (W - total) // 2
    y = 150
    for i, (num, title, sub, color) in enumerate(steps):
        box = (x, y, x + card_w, y + card_h)
        _shadow_card(img, box)
        _round_rect(d, box, 16, fill=CARD + (255,), outline=BORDER + (255,), width=1)
        d.ellipse((x + card_w / 2 - 22, y + 18, x + card_w / 2 + 22, y + 62), fill=color + (255,))
        _text_center(d, x + card_w / 2, y + 40, num, _font(26, True), (255, 255, 255))
        _text_center(d, x + card_w / 2, y + 90, title, _font(20, True), INK)
        _text_center(d, x + card_w / 2, y + 118, sub, _font(15), MUTED)
        if i < n - 1:
            ax = x + card_w + gap / 2
            d.line((ax - 10, y + card_h / 2, ax + 10, y + card_h / 2), fill=ACCENT + (255,), width=4)
            d.polygon([(ax + 10, y + card_h / 2 - 7), (ax + 10, y + card_h / 2 + 7), (ax + 20, y + card_h / 2)], fill=ACCENT + (255,))
        x += card_w + gap
    _save(img, path)


# ── Bild 2: Unterstützte Formate ─────────────────────────────────────────────
def _draw_formate(path: Path):
    W, H = 1100, 300
    img, d = _new_canvas(W, H)
    _text_center(d, W // 2, 44, "Welche Dateien kann ich laden?", _font(32, True), PRIMARY)

    formats = [
        ("XLSX", "Excel (neu)", SUCCESS, True),
        ("XLS", "Excel (alt)", SUCCESS, True),
        ("CSV", "Trennzeichen", ACCENT, True),
        ("TXT", "Texttabelle", ACCENT, True),
    ]
    cw, ch, gap = 220, 130, 30
    total = len(formats) * cw + (len(formats) - 1) * gap
    x = (W - total) // 2
    y = 110
    for ext, label, color, ok in formats:
        box = (x, y, x + cw, y + ch)
        _shadow_card(img, box)
        _round_rect(d, box, 16, fill=CARD + (255,), outline=BORDER + (255,), width=1)
        _round_rect(d, (x + 22, y + 28, x + 92, y + 102), 12, fill=color + (255,))
        _text_center(d, x + 57, y + 65, ext, _font(20, True), (255, 255, 255))
        d.text((x + 110, y + 38), label, font=_font(20, True), fill=INK)
        _draw_check(d, x + 110, y + 72, 14, SUCCESS, width=3)
        d.text((x + 132, y + 70), "wird gelesen", font=_font(16), fill=SUCCESS)
        x += cw + gap
    _text_center(d, W // 2, 270, "Auch alte .xls-Dateien werden jetzt automatisch verarbeitet.", _font(18), MUTED)
    _save(img, path)


# ── Bild 3: Fortschritts-Anzeige erklärt ─────────────────────────────────────
def _draw_fortschritt(path: Path):
    W, H = 1100, 330
    img, d = _new_canvas(W, H)
    _text_center(d, W // 2, 44, "Du siehst jederzeit, dass gearbeitet wird", _font(30, True), PRIMARY)

    box = (140, 100, W - 140, 270)
    _shadow_card(img, box)
    _round_rect(d, box, 18, fill=CARD + (255,), outline=BORDER + (255,), width=1)

    d.text((175, 128), "Auswertung läuft …", font=_font(22, True), fill=INK)
    d.text((box[2] - 175, 130), "00:07", font=_font(20, True), fill=ACCENT, anchor="ra")

    # Fortschrittsbalken
    bx0, bx1, by = 175, box[2] - 175, 185
    _round_rect(d, (bx0, by, bx1, by + 22), 11, fill=(232, 237, 242, 255))
    fill_x = bx0 + int((bx1 - bx0) * 0.62)
    _round_rect(d, (bx0, by, fill_x, by + 22), 11, fill=ACCENT + (255,))

    # Stufen-Checkliste
    stages = [("check", "Datei eingelesen", SUCCESS), ("check", "Stammdaten geladen", SUCCESS),
              ("dot", "Positionen werden abgeglichen", ACCENT), ("ring", "Excel schreiben", MUTED)]
    sx = 175
    for mark, label, color in stages:
        if mark == "check":
            _draw_check(d, sx, 230, 16, color, width=3)
        elif mark == "dot":
            d.ellipse((sx, 232, sx + 15, 247), fill=color + (255,))
        else:
            d.ellipse((sx, 232, sx + 15, 247), outline=color + (255,), width=2)
        d.text((sx + 26, 230), label, font=_font(16), fill=INK if color != MUTED else MUTED)
        sx += d.textbbox((0, 0), label, font=_font(16))[2] + 60
    _save(img, path)


_IMAGES = {
    "prozess_flow.png": _draw_prozess_flow,
    "formate.png": _draw_formate,
    "fortschritt.png": _draw_fortschritt,
}


def ensure_help_images(target_dir: str | Path, force: bool = False) -> dict[str, Path]:
    target = Path(target_dir)
    target.mkdir(parents=True, exist_ok=True)
    out: dict[str, Path] = {}
    for name, fn in _IMAGES.items():
        p = target / name
        if force or not p.exists():
            try:
                fn(p)
            except Exception:
                continue
        out[name] = p
    return out


if __name__ == "__main__":
    here = Path(__file__).resolve().parent.parent / "testui_assets"
    res = ensure_help_images(here, force=True)
    for k, v in res.items():
        print("ok", k, v)
