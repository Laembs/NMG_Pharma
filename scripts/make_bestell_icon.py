"""Erzeugt assets/bestell.ico (Einkaufswagen auf NMG-Blau) ohne Fremdbibliotheken.

Reines Python: zeichnet je ein BGRA-Bitmap pro Groesse und packt sie als
32-bpp-ICO. Wiederverwendbar - bei Designaenderung einfach neu ausfuehren:
    python scripts/make_bestell_icon.py
"""
from __future__ import annotations

import struct
from pathlib import Path

OUT = Path(__file__).resolve().parent.parent / "assets" / "bestell.ico"

BLUE = (0x86, 0x4a, 0x0b, 255)   # #0b4a86 als BGRA
WHITE = (255, 255, 255, 255)
SIZES = [16, 32, 48]


def _blend(dst, src):
    a = src[3] / 255.0
    return (
        round(src[0] * a + dst[0] * (1 - a)),
        round(src[1] * a + dst[1] * (1 - a)),
        round(src[2] * a + dst[2] * (1 - a)),
        max(dst[3], src[3]),
    )


class Canvas:
    def __init__(self, size):
        self.s = size
        self.px = [[(0, 0, 0, 0) for _ in range(size)] for _ in range(size)]

    def set(self, x, y, color):
        if 0 <= x < self.s and 0 <= y < self.s:
            self.px[y][x] = _blend(self.px[y][x], color)

    def rounded_bg(self, color, radius):
        s, r = self.s, radius
        for y in range(s):
            for x in range(s):
                inside = True
                for cx, cy in ((r, r), (s - 1 - r, r), (r, s - 1 - r), (s - 1 - r, s - 1 - r)):
                    if ((x < r and y < r) or (x > s - 1 - r and y < r)
                            or (x < r and y > s - 1 - r) or (x > s - 1 - r and y > s - 1 - r)):
                        if (x - cx) ** 2 + (y - cy) ** 2 > r * r and \
                           abs(x - cx) <= r and abs(y - cy) <= r:
                            inside = False
                            break
                if inside:
                    self.set(x, y, color)

    def disc(self, cx, cy, r, color):
        for y in range(int(cy - r) - 1, int(cy + r) + 2):
            for x in range(int(cx - r) - 1, int(cx + r) + 2):
                if (x - cx) ** 2 + (y - cy) ** 2 <= r * r:
                    self.set(x, y, color)

    def fill_poly(self, pts, color):
        ys = [p[1] for p in pts]
        for y in range(int(min(ys)), int(max(ys)) + 1):
            xs = []
            n = len(pts)
            for i in range(n):
                x1, y1 = pts[i]
                x2, y2 = pts[(i + 1) % n]
                if (y1 <= y < y2) or (y2 <= y < y1):
                    xs.append(x1 + (y - y1) * (x2 - x1) / (y2 - y1))
            xs.sort()
            for i in range(0, len(xs) - 1, 2):
                for x in range(int(round(xs[i])), int(round(xs[i + 1])) + 1):
                    self.set(x, y, color)

    def thick_line(self, x1, y1, x2, y2, w, color):
        steps = int(max(abs(x2 - x1), abs(y2 - y1)) * 2) + 1
        for i in range(steps + 1):
            t = i / steps
            self.disc(x1 + (x2 - x1) * t, y1 + (y2 - y1) * t, w / 2.0, color)


def draw_cart(size):
    c = Canvas(size)
    u = size / 32.0  # Einheit relativ zu 32er-Raster

    def U(v):
        return v * u

    c.rounded_bg(BLUE, max(2, round(size * 0.18)))

    # Griff oben links + Verbindung zum Korb
    c.thick_line(U(4), U(8), U(9), U(8), max(1.5, U(2)), WHITE)
    c.thick_line(U(9), U(8), U(11), U(13), max(1.5, U(2)), WHITE)

    # Korb (Trapez)
    c.fill_poly([(U(10), U(13)), (U(26), U(13)), (U(23), U(22)), (U(13), U(22))], WHITE)

    # kleine blaue Streben im Korb fuer "Gitter"-Anmutung (nur ab 32px sichtbar)
    if size >= 32:
        for gx in (15, 18, 21):
            c.thick_line(U(gx), U(14), U(gx) - U(0.5), U(21), max(1, U(1)), BLUE)

    # Raeder
    c.disc(U(15), U(26), U(2.3), WHITE)
    c.disc(U(22), U(26), U(2.3), WHITE)
    return c


def ico_image(c: Canvas) -> bytes:
    s = c.s
    # BITMAPINFOHEADER: Hoehe = 2*s (Farbe + Maske)
    hdr = struct.pack("<IiiHHIIiiII", 40, s, s * 2, 1, 32, 0, 0, 0, 0, 0, 0)
    xor = bytearray()
    for y in range(s - 1, -1, -1):       # bottom-up
        for x in range(s):
            b, g, r, a = c.px[y][x]
            xor += bytes((b, g, r, a))
    # AND-Maske (1bpp, alle 0 -> Alpha steuert Transparenz), Zeilen auf 32bit
    row_bytes = ((s + 31) // 32) * 4
    andmask = bytes(row_bytes * s)
    return hdr + bytes(xor) + andmask


def build():
    images = [ico_image(draw_cart(s)) for s in SIZES]
    out = bytearray()
    out += struct.pack("<HHH", 0, 1, len(images))
    offset = 6 + 16 * len(images)
    for s, img in zip(SIZES, images):
        w = 0 if s >= 256 else s
        out += struct.pack("<BBBBHHII", w, w, 0, 0, 1, 32, len(img), offset)
        offset += len(img)
    for img in images:
        out += img
    OUT.write_bytes(out)
    print(f"geschrieben: {OUT}  ({len(out)} bytes, Groessen {SIZES})")


if __name__ == "__main__":
    build()
