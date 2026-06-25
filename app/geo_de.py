"""Offline-Geokodierung Deutschland fuer die Kunden-Landkarte.

Kein Internet, keine Zusatzpakete: Eine PLZ wird ueber ihre zweistellige
Leitregion auf einen ungefaehren Mittelpunkt (lat/lon) abgebildet. Das reicht,
um Apotheken auf einer Deutschlandkarte ihrer Region zuzuordnen ("wo sitzt
welcher Kunde"). Zusaetzlich ein stark vereinfachter Grenzumriss zum Zeichnen.

Genauigkeit: Stadt-/Regionsebene (~20-40 km), bewusst grob. Fuer eine
strassengenaue Karte braeuchte es einen echten Geocoder/Tiles.
"""
from __future__ import annotations

# ── Zweistellige PLZ-Leitregion -> ungefaehrer Mittelpunkt (lat, lon) ─────────
PLZ2_LATLON: dict[str, tuple[float, float]] = {
    "01": (51.05, 13.74), "02": (51.18, 14.43), "03": (51.76, 14.33), "04": (51.34, 12.37),
    "06": (51.48, 11.97), "07": (50.87, 11.60), "08": (50.72, 12.49), "09": (50.83, 12.92),
    "10": (52.52, 13.40), "12": (52.46, 13.43), "13": (52.56, 13.36), "14": (52.40, 13.06),
    "15": (52.34, 14.55), "16": (52.83, 13.82), "17": (53.56, 13.26), "18": (54.09, 12.14),
    "19": (53.63, 11.41),
    "20": (53.55, 9.99), "21": (53.25, 10.41), "22": (53.60, 9.90), "23": (53.87, 10.69),
    "24": (54.32, 10.14), "25": (53.93, 9.52), "26": (53.14, 8.21), "27": (53.55, 8.58),
    "28": (53.08, 8.81), "29": (52.63, 10.08),
    "30": (52.37, 9.73), "31": (52.15, 9.95), "32": (52.11, 8.67), "33": (51.72, 8.75),
    "34": (51.31, 9.49), "35": (50.58, 8.68), "36": (50.55, 9.68), "37": (51.54, 9.92),
    "38": (52.27, 10.52), "39": (52.13, 11.62),
    "40": (51.23, 6.78), "41": (51.19, 6.44), "42": (51.26, 7.18), "44": (51.51, 7.47),
    "45": (51.46, 7.01), "46": (51.50, 6.85), "47": (51.43, 6.76), "48": (51.96, 7.63),
    "49": (52.28, 8.05),
    "50": (50.94, 6.96), "51": (51.03, 7.00), "52": (50.78, 6.08), "53": (50.74, 7.10),
    "54": (49.76, 6.64), "55": (50.00, 8.27), "56": (50.36, 7.59), "57": (50.88, 8.02),
    "58": (51.36, 7.47), "59": (51.68, 7.82),
    "60": (50.11, 8.68), "61": (50.23, 8.62), "63": (49.97, 9.00), "64": (49.87, 8.65),
    "65": (50.08, 8.24), "66": (49.24, 6.99), "67": (49.44, 7.77), "68": (49.49, 8.47),
    "69": (49.40, 8.69),
    "70": (48.78, 9.18), "71": (48.90, 9.19), "72": (48.52, 9.06), "73": (48.70, 9.65),
    "74": (49.14, 9.22), "75": (48.89, 8.70), "76": (49.01, 8.40), "77": (48.47, 7.94),
    "78": (48.06, 8.46), "79": (47.99, 7.85),
    "80": (48.14, 11.58), "81": (48.11, 11.60), "82": (47.99, 11.34), "83": (47.86, 12.12),
    "84": (48.54, 12.15), "85": (48.77, 11.42), "86": (48.37, 10.90), "87": (47.73, 10.32),
    "88": (47.78, 9.61), "89": (48.40, 9.99),
    "90": (49.45, 11.08), "91": (49.45, 10.90), "92": (49.44, 11.86), "93": (49.01, 12.10),
    "94": (48.57, 13.46), "95": (50.32, 11.92), "96": (49.89, 10.90), "97": (49.79, 9.95),
    "98": (50.61, 10.69), "99": (50.98, 11.03),
}

# Geografische Grenzen Deutschlands (fuer die Projektion)
LAT_MIN, LAT_MAX = 47.2, 55.1
LON_MIN, LON_MAX = 5.8, 15.1

# ── Stark vereinfachter Grenzumriss (lat, lon), im Uhrzeigersinn ──────────────
GERMANY_OUTLINE: list[tuple[float, float]] = [
    (54.90, 8.30), (54.80, 9.40), (54.38, 10.18), (54.40, 11.20), (54.18, 12.10),
    (54.40, 13.40), (53.90, 14.20), (52.85, 14.14), (52.34, 14.55), (51.50, 14.75),
    (51.15, 15.04), (50.92, 14.82), (50.70, 14.00), (50.20, 12.90), (49.70, 12.50),
    (48.95, 13.30), (48.57, 13.46), (48.10, 12.90), (47.70, 13.00), (47.58, 13.02),
    (47.50, 11.40), (47.42, 10.98), (47.27, 10.28), (47.55, 9.18), (47.62, 8.60),
    (47.59, 7.60), (48.00, 7.60), (48.60, 8.00), (49.00, 8.20), (49.23, 6.99),
    (49.80, 6.40), (50.32, 6.36), (50.78, 6.08), (51.05, 6.00), (51.40, 6.10),
    (51.83, 6.10), (52.20, 7.00), (52.62, 7.05), (53.05, 7.20), (53.35, 7.20),
    (53.60, 7.20), (53.70, 8.00), (53.88, 8.65), (54.30, 8.60), (54.50, 8.55),
    (54.90, 8.30),
]


def plz_to_latlon(plz: str | None) -> tuple[float, float] | None:
    """Ungefaehre Koordinate (lat, lon) zur PLZ ueber ihre 2-stellige Region."""
    p = (plz or "").strip()
    if len(p) < 2 or not p[:2].isdigit():
        return None
    return PLZ2_LATLON.get(p[:2])


def project(lat: float, lon: float, width: int, height: int,
            pad: int = 14) -> tuple[float, float]:
    """Aequirektangulaere Projektion auf eine Canvas-Flaeche (width x height).

    Laengen werden mit cos(mittlere Breite) gestaucht, damit Deutschland nicht
    in die Breite gezogen wirkt. Liefert (x, y) in Pixeln.
    """
    import math
    lat_mid = (LAT_MIN + LAT_MAX) / 2.0
    cos_mid = math.cos(math.radians(lat_mid)) or 1.0
    lon_span = (LON_MAX - LON_MIN) * cos_mid
    lat_span = (LAT_MAX - LAT_MIN)
    # gemeinsamer Massstab, damit Seitenverhaeltnis stimmt
    sx = (width - 2 * pad) / lon_span
    sy = (height - 2 * pad) / lat_span
    s = min(sx, sy)
    # zentrieren
    used_w = lon_span * s
    used_h = lat_span * s
    off_x = (width - used_w) / 2.0
    off_y = (height - used_h) / 2.0
    x = off_x + ((lon - LON_MIN) * cos_mid) * s
    y = off_y + (LAT_MAX - lat) * s  # lat invertiert (Norden oben)
    return x, y
