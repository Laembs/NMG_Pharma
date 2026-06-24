"""Erzeugt die Screenshots fuer die Hilfe-App (app/hilfe_app.py).

Baut die echten Fenster/Seiten auf und greift mit PIL ImageGrab exakt den
Fensterinhalt ab - kein manuelles Zuschneiden. Die Bilder landen direkt in
  assets/hilfe/<thema>/<datei>.png
also genau dort, wo die Hilfe sie zur Laufzeit erwartet.

Aufruf (aus dem Repo-Root):
    python scripts/hilfe_screenshots.py            # alles
    python scripts/hilfe_screenshots.py nmgone     # nur NMGone-Seiten
    python scripts/hilfe_screenshots.py kasse
    python scripts/hilfe_screenshots.py report

Modale Dialoge (messagebox/simpledialog) werden vorab neutralisiert, damit der
Aufbau nicht blockiert. Das Skript veraendert keine Programmlogik - es ist ein
reines Werkzeug zum Befuellen der Hilfe-Bilder und kann nach jeder UI-Aenderung
erneut laufen.
"""
from __future__ import annotations

import os
import sys
import time
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO))

# ── Modale Dialoge neutralisieren, BEVOR app-Module importiert werden ─────────
import tkinter as tk
import tkinter.messagebox as mb
import tkinter.simpledialog as sd

for _n in ("showinfo", "showwarning", "showerror"):
    setattr(mb, _n, lambda *a, **k: None)
mb.askyesno = lambda *a, **k: False
mb.askokcancel = lambda *a, **k: False
mb.askquestion = lambda *a, **k: "no"
mb.askretrycancel = lambda *a, **k: False
sd.askstring = lambda *a, **k: None
sd.askinteger = lambda *a, **k: None
sd.askfloat = lambda *a, **k: None
# Sicherheitsnetz: kein wait_window blockiert den Screenshot-Lauf.
tk.Misc.wait_window = lambda self, window=None: None

from app.config import ASSETS_DIR  # noqa: E402
from app import theme              # noqa: E402

HILFE = ASSETS_DIR / "hilfe"
GEO = "1360x860+20+20"


def _settle(win):
    win.update_idletasks()
    for _ in range(3):
        win.update(); time.sleep(0.15)
    try:
        win.deiconify(); win.lift(); win.focus_force()
    except Exception:
        pass
    win.update(); time.sleep(0.35); win.update()


def _capture_hwnd_printwindow(win):
    """Captured ein bestimmtes Fenster via PrintWindow (ctypes). Funktioniert
    auch, wenn das Fenster verdeckt ist oder nicht im Vordergrund liegt - das
    Fenster zeichnet sich selbst in einen Speicher-DC. Liefert ein PIL-Image
    (ganzes Fenster inkl. Titelleiste) oder None bei Fehler."""
    import ctypes
    from ctypes import wintypes
    from PIL import Image

    user32 = ctypes.windll.user32
    gdi32 = ctypes.windll.gdi32

    hwnd = win.winfo_id()
    root = user32.GetAncestor(hwnd, 2)  # GA_ROOT -> Top-Level inkl. Rahmen
    hwnd = root or hwnd

    rect = wintypes.RECT()
    user32.GetWindowRect(hwnd, ctypes.byref(rect))
    w, h = rect.right - rect.left, rect.bottom - rect.top
    if w <= 0 or h <= 0:
        return None

    hwndDC = user32.GetWindowDC(hwnd)
    mfcDC = gdi32.CreateCompatibleDC(hwndDC)
    bmp = gdi32.CreateCompatibleBitmap(hwndDC, w, h)
    gdi32.SelectObject(mfcDC, bmp)
    # PW_RENDERFULLCONTENT = 2 -> rendert auch modernen Inhalt korrekt
    user32.PrintWindow(hwnd, mfcDC, 2)

    class BMIH(ctypes.Structure):
        _fields_ = [("biSize", wintypes.DWORD), ("biWidth", wintypes.LONG),
                    ("biHeight", wintypes.LONG), ("biPlanes", wintypes.WORD),
                    ("biBitCount", wintypes.WORD), ("biCompression", wintypes.DWORD),
                    ("biSizeImage", wintypes.DWORD), ("biXPelsPerMeter", wintypes.LONG),
                    ("biYPelsPerMeter", wintypes.LONG), ("biClrUsed", wintypes.DWORD),
                    ("biClrImportant", wintypes.DWORD)]

    bmi = BMIH()
    bmi.biSize = ctypes.sizeof(BMIH)
    bmi.biWidth, bmi.biHeight = w, -h  # negativ = top-down
    bmi.biPlanes, bmi.biBitCount, bmi.biCompression = 1, 32, 0
    buf = ctypes.create_string_buffer(w * h * 4)
    got = gdi32.GetDIBits(mfcDC, bmp, 0, h, buf, ctypes.byref(bmi), 0)
    gdi32.DeleteObject(bmp)
    gdi32.DeleteDC(mfcDC)
    user32.ReleaseDC(hwnd, hwndDC)
    if not got:
        return None
    return Image.frombuffer("RGB", (w, h), buf, "raw", "BGRX", 0, 1)


def grab(win, relpath: str):
    """Greift den Inhalt von 'win' ab und speichert ihn nach relpath. Bevorzugt
    PrintWindow (robust gegen Verdeckung), faellt sonst auf ImageGrab zurueck."""
    _settle(win)
    img = None
    try:
        img = _capture_hwnd_printwindow(win)
    except Exception as e:
        print(f"  printwindow-fehler: {e!r}")
    if img is None:
        from PIL import ImageGrab
        x, y = win.winfo_rootx(), win.winfo_rooty()
        w, h = win.winfo_width(), win.winfo_height()
        img = ImageGrab.grab(bbox=(x, y, x + w, y + h))
    out = HILFE / relpath
    out.parent.mkdir(parents=True, exist_ok=True)
    img.save(out)
    bright = img.convert("L").resize((1, 1)).getpixel((0, 0))
    print(f"SAVED {relpath}  {img.width}x{img.height}  brightness={bright}")


def cap_nmgone():
    from app.gui import NMGApp
    for m in ("_check_mitarbeiterprofil_pflicht", "_check_daten_benachrichtigungen",
              "_check_wissensbasis_leer", "_check_apotheken_analyse_faellig",
              "_startup_update_check"):
        if hasattr(NMGApp, m):
            setattr(NMGApp, m, lambda self, *a, **k: None)
    app = NMGApp()
    try:
        app.state("normal")
    except Exception:
        pass
    app.geometry(GEO)
    app.update()

    pages = [
        ("show_startseite",               ["start/01_startbildschirm.png",
                                           "dashboard/01_dashboard.png"]),
        ("show_neue_auswertung_page",     ["bedarfsanalyse/01_bedarfsanalyse.png"]),
        ("show_kunden_center",            ["kunden/01_kunden.png"]),
        ("show_schulbank_page",           ["schulbank/01_schulbank.png"]),
        ("show_daten_aktualisieren_page", ["daten/01_daten.png"]),
        ("show_update_backup_page",       ["backup/01_backup.png"]),
    ]
    for method, outs in pages:
        try:
            getattr(app, method)()
            app.update_idletasks(); app.update(); time.sleep(0.3)
            for rel in outs:
                grab(app, rel)
        except Exception as e:
            print(f"FAIL {method}: {e!r}")
    app.destroy()


def _set_icon(root, *names):
    for n in names:
        try:
            root.iconbitmap(str(ASSETS_DIR / n))
            return
        except Exception:
            continue


def _all_toplevels(root):
    """Alle Toplevel-Fenster unterhalb von root (rekursiv) - Dialoge haengen oft
    am Panel-Frame, nicht direkt an root."""
    found = []

    def walk(w):
        for c in w.winfo_children():
            if isinstance(c, tk.Toplevel):
                found.append(c)
            walk(c)

    walk(root)
    return found


def cap_kasse():
    from app.kasse_app import KassePanel, SHELL_BG
    root = tk.Tk()
    root.title("NMGone Kasse")
    root.geometry(GEO)
    root.configure(bg=SHELL_BG)
    _set_icon(root, "kasse.ico", "NMGone.ico")
    theme.apply_theme(root)
    theme.apply_widget_defaults(root)
    try:
        p = KassePanel(root)
        p.pack(fill="both", expand=True)
        root.update_idletasks(); root.update(); time.sleep(0.3)
        grab(root, "kasse/01_kasse.png")
        # Weitere Reiter durchschalten und je abgreifen.
        for key, rel in (("wareneingang",   "kasse/02_wareneingang.png"),
                         ("vorbestellungen", "kasse/03_vorbestellungen.png"),
                         ("auswertung",      "kasse/04_auswertung.png"),
                         ("defektmeldung",   "kasse/05_defektmeldung.png"),
                         ("einstellungen",   "kasse/06_einstellungen.png")):
            try:
                p._show_view(key)
                root.update_idletasks(); root.update(); time.sleep(0.3)
                grab(root, rel)
            except Exception as e:
                print(f"FAIL kasse/{key}: {e!r}")
    except Exception as e:
        print(f"FAIL kasse: {e!r}")
    root.destroy()


def cap_faktura():
    from app.faktura_app import FakturaPanel, SHELL_BG
    try:
        from app.migrations import run_migrations
        run_migrations()
    except Exception:
        pass
    root = tk.Tk()
    root.title("NMG Faktura")
    root.geometry(GEO)
    root.configure(bg=SHELL_BG)
    _set_icon(root, "Faktura.ico", "NMGone.ico")
    theme.apply_theme(root)
    theme.apply_widget_defaults(root)
    try:
        p = FakturaPanel(root, on_close=root.destroy)
        p.pack(fill="both", expand=True)
        root.update_idletasks(); root.update(); time.sleep(0.3)
        grab(root, "faktura/01_start.png")
        for key, rel in (("neu",         "faktura/02_rechnung_neu.png"),
                         ("einst_firma", "faktura/03_einstellungen.png")):
            try:
                p.show(key)
                root.update_idletasks(); root.update(); time.sleep(0.3)
                grab(root, rel)
            except Exception as e:
                print(f"FAIL faktura/{key}: {e!r}")
    except Exception as e:
        print(f"FAIL faktura: {e!r}")
    root.destroy()


def _seed_personal(dbpath, pa):
    """Schreibt Demo-Mitarbeiter/-Struktur in die TEMP-DB, damit die Screenshots
    nicht leer sind. Nutzt die in personal_app vorhandenen Demo-Konstanten."""
    import sqlite3
    import calendar
    from datetime import date
    con = sqlite3.connect(dbpath)
    for i, v, n, ab, po, x, y in pa.DEMO_EMPS:
        con.execute("INSERT INTO tbl_mitarbeiter(id,vorname,name,abteilung,position,board_x,board_y) "
                    "VALUES(?,?,?,?,?,?,?)", (i, v, n, ab, po, x, y))
    con.execute("UPDATE tbl_mitarbeiter SET personalverantwortlich=1 WHERE id=8")
    for c, p, a, pr in pa.DEMO_LINKS:
        con.execute("INSERT INTO tbl_mitarbeiter_vorgesetzter(mitarbeiter_id,vorgesetzter_id,art,ist_primaer) "
                    "VALUES(?,?,?,?)", (c, p, a, pr))
    t = date.today(); y, m = t.year, t.month

    def dd(day):
        return date(y, m, min(day, calendar.monthrange(y, m)[1])).isoformat()

    for mid, art, a, b in [(4, "Urlaub", 3, 9), (5, "Krankheit", 10, 12),
                           (2, "Fortbildung", 17, 18), (6, "Urlaub", 20, 27),
                           (7, "Sonstiges", 5, 5), (8, "Urlaub", 24, 28)]:
        con.execute("INSERT INTO tbl_abwesenheit(mitarbeiter_id,art,von,bis) VALUES(?,?,?,?)",
                    (mid, art, dd(a), dd(b)))
    for k in ("Vertrieb", "Labor"):
        con.execute("INSERT OR IGNORE INTO tbl_kategorie(name) VALUES(?)", (k,))
    bid = {}
    for name, kat in [("Außendienst Nord", "Vertrieb"), ("Innendienst", "Vertrieb"),
                      ("Analytik", "Labor"), ("QS", "Labor")]:
        cur = con.execute("INSERT INTO tbl_arbeitsbereich(name,kategorie) VALUES(?,?)", (name, kat))
        bid[name] = cur.lastrowid
    for mid, bname in [(4, "Außendienst Nord"), (5, "Innendienst"), (6, "Analytik"),
                       (7, "QS"), (3, "Analytik")]:
        con.execute("INSERT OR IGNORE INTO tbl_mitarbeiter_arbeitsbereich(mitarbeiter_id,bereich_id) "
                    "VALUES(?,?)", (mid, bid[bname]))
    con.commit(); con.close()


def cap_personal():
    """Personal-Board mit Demo-Daten in einer TEMP-DB (echte NMGone-DB bleibt
    unberührt - sie enthält bewusst keine Mitarbeiter-Demodaten)."""
    import tempfile
    from app import personal_app as pa
    tmp = Path(tempfile.gettempdir()) / "nmg_hilfe_personal_demo.sqlite"
    if tmp.exists():
        try:
            tmp.unlink()
        except Exception:
            pass
    orig_db = pa.DB_PATH
    pa.DB_PATH = tmp
    root = None
    try:
        pa.init_db(reset=True)
        _seed_personal(tmp, pa)
        root = tk.Tk()
        app = pa.App(root)
        root.geometry(GEO)  # ueberschreibt die App-eigene Geometrie
        root.update_idletasks(); root.update(); time.sleep(0.3)
        grab(root, "personal/01_organigramm.png")
        for method, rel in (("show_arbeitsbereiche", "personal/02_arbeitsbereiche.png"),
                            ("show_absence",          "personal/03_abwesenheiten.png")):
            try:
                getattr(app, method)()
                root.update_idletasks(); root.update(); time.sleep(0.3)
                grab(root, rel)
            except Exception as e:
                print(f"FAIL personal/{method}: {e!r}")
    except Exception as e:
        print(f"FAIL personal: {e!r}")
    finally:
        if root is not None:
            try:
                root.destroy()
            except Exception:
                pass
        pa.DB_PATH = orig_db
        try:
            tmp.unlink()
        except Exception:
            pass


def cap_report():
    from app.report_app import BerichtPanel
    root = tk.Tk()
    root.title("NMG Auswertungen")
    root.geometry(GEO)
    root.configure(bg=theme.BG)
    _set_icon(root, "Report.ico", "NMGone.ico")
    theme.apply_theme(root)
    theme.apply_widget_defaults(root)
    try:
        p = BerichtPanel(root)
        p.pack(fill="both", expand=True)
        root.update_idletasks(); root.update(); time.sleep(0.3)
        grab(root, "auswertungen/01_auswertungen.png")
        # Freie Auswertung: Baukasten-Dialog oeffnen und abgreifen.
        p._select_perspektive("frei")
        root.update(); time.sleep(0.2)
        p._open_builder_dialog()
        root.update_idletasks(); root.update(); time.sleep(0.4)
        tops = _all_toplevels(root)
        if tops:
            grab(tops[-1], "auswertungen/02_auswertung_frei.png")
        else:
            print("  (kein Baukasten-Dialog gefunden - nehme Hauptfenster)")
            grab(root, "auswertungen/02_auswertung_frei.png")
    except Exception as e:
        print(f"FAIL report: {e!r}")
    root.destroy()


def main():
    target = sys.argv[1] if len(sys.argv) > 1 else "all"
    if target in ("nmgone", "all"):
        cap_nmgone()
    if target in ("kasse", "all"):
        cap_kasse()
    if target in ("report", "all"):
        cap_report()
    if target in ("faktura", "all"):
        cap_faktura()
    if target in ("personal", "all"):
        cap_personal()
    print("FERTIG.")


if __name__ == "__main__":
    main()
