#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Mitarbeiter-Board · TESTVERSION (eigenständig)
==============================================
Probiert das geplante Phase-2 aus, OHNE die echte App zu berühren.
Eigene Test-Datenbank (organigramm_test.db) mit Demo-Daten.

Zwei Ansichten (oben rechts umschaltbar):
  • Organigramm   – Post-it-Karten ziehen, verbinden, lösen, Teilbereiche, Auto-Layout
  • Abwesenheiten – Team-Kalender: Urlaub / Krankheit / Fortbildung eintragen,
                    farbige Balken pro Tag, Urlaubstage-Zähler

Start:  python start_organigramm_test.py
"""
from __future__ import annotations
import os
import calendar
import getpass
import sqlite3
from datetime import date, timedelta, datetime
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog

# ── Palette (aus app/theme.py) ───────────────────────────────────────────────
PRIMARY      = "#0B4A86"
SIDEBAR      = "#0B2C4A"
SIDEBAR_TEXT = "#D7E6F4"
BG           = "#EEF2F7"
CARD         = "#FFFFFF"
INK          = "#1F2933"
MUTED        = "#626E7D"
FAINT        = "#9AA5B1"
BORDER       = "#DCE2E8"
GRID         = "#E4EAF1"
SHADOW       = "#C9D2DC"
WEEKEND      = "#F0F3F7"
TODAY        = "#FFF6E0"
SELECT_BG    = "#E8F1FB"
FONT         = "Segoe UI"

STD_URLAUB   = 30  # Standard-Urlaubsanspruch (Tage/Jahr), pro Person änderbar

CARD_W, CARD_H = 196, 96

ART_STYLE = {
    "disziplinarisch": dict(color="#185FA5", width=3, dash=None,   arrow=True),
    "fachlich":        dict(color="#BA7517", width=2, dash=(6, 4), arrow=False),
    "Vertretung":      dict(color="#5F5E5A", width=2, dash=(2, 5), arrow=False),
}
ARTEN = tuple(ART_STYLE.keys())

# Abwesenheits-Arten + Farben
ABW_STYLE = {
    "Urlaub":      "#11823B",
    "Sonderurlaub": "#0E7C86",
    "Krankheit":   "#C0392B",
    "Fortbildung": "#C88200",
    "Sonstiges":   "#6B4FB3",
}
ABW_ARTEN = tuple(ABW_STYLE.keys())

# Genehmigungs-Status für Abwesenheits-Anträge
ST_BEANTRAGT = "Beantragt"      # neu eingereicht, wartet auf Personalverantwortliche
ST_GF        = "GF-Freigabe"     # Stufe 1 genehmigt, wartet auf Geschäftsführung
ST_GENEHMIGT = "Genehmigt"       # endgültig genehmigt, zählt auf das Urlaubskonto
ST_ABGELEHNT = "Abgelehnt"       # abgelehnt, wird nicht im Kalender geführt
ST_OFFEN     = (ST_BEANTRAGT, ST_GF)   # noch zu entscheiden

ST_STYLE = {
    ST_BEANTRAGT: ("⏳ Wartet auf Genehmigung", "#C88200"),
    ST_GF:        ("⏳ Wartet auf GF-Freigabe", "#0E7C86"),
    ST_GENEHMIGT: ("✓ Genehmigt", "#11823B"),
    ST_ABGELEHNT: ("✗ Abgelehnt", "#C0392B"),
}

# Vertretungsstufen je Arbeitsbereich: 0 = Verantwortlich, 1..3 = Vertretungskette
STUFEN = ["Verantwortlich", "1. Vertretung", "2. Vertretung", "3. Vertretung"]
STUFE_SHORT = ["V", "1.V", "2.V", "3.V"]
STUFE_COLOR = ["#11823B", "#0E7C86", "#6B4FB3", "#C88200"]

# Gründe für Sonderurlaub (Auswahl, wenn Art = Sonderurlaub)
SONDERURLAUB_GRUENDE = [
    "Eigene Hochzeit",
    "Geburt eines Kindes",
    "Todesfall (naher Angehöriger)",
    "Umzug (betrieblich veranlasst)",
    "Schwere Erkrankung im Haushalt",
    "Dienstjubiläum",
    "Pflege Angehöriger",
    "Sonstiger Anlass",
]

DEPT_PALETTE = ["#0B4A86", "#11823B", "#6B4FB3", "#C88200", "#0E7C86", "#A8324A"]
WD = ["Mo", "Di", "Mi", "Do", "Fr", "Sa", "So"]
MONATE = ["", "Januar", "Februar", "März", "April", "Mai", "Juni",
          "Juli", "August", "September", "Oktober", "November", "Dezember"]

from app.config import DB_PATH, DEMO_SUFFIX  # gemeinsame NMGone-Datenbank
from app import tour

DEMO_EMPS = [
    (1, "Anna",  "Maier",  "Leitung",    "Geschäftsführung", 470, 40),
    (2, "Ben",   "Krause", "Vertrieb",   "Leitung Vertrieb", 250, 250),
    (3, "Carla", "Sommer", "Labor",      "Leitung Labor",    760, 250),
    (4, "David", "Reuter", "Vertrieb",   "Außendienst",      110, 470),
    (5, "Eva",   "Lang",   "Vertrieb",   "Innendienst",      370, 470),
    (6, "Felix", "Thiel",  "Labor",      "Laborant",         660, 470),
    (7, "Gina",  "Wolf",   "Labor",      "QS / Analytik",    900, 470),
    (8, "Hans",  "Berg",   "Verwaltung", "Buchhaltung",      470, 250),
]
DEMO_LINKS = [
    (2, 1, "disziplinarisch", 1), (3, 1, "disziplinarisch", 1), (8, 1, "disziplinarisch", 1),
    (4, 2, "disziplinarisch", 1), (5, 2, "disziplinarisch", 1),
    (6, 3, "disziplinarisch", 1), (7, 3, "disziplinarisch", 1),
    (6, 7, "fachlich", 0), (4, 5, "Vertretung", 0),
]


def _clamp_day(y, m, d):
    return min(d, calendar.monthrange(y, m)[1])


def init_db(reset=False):
    if reset and os.path.exists(DB_PATH):
        os.remove(DB_PATH)
    con = sqlite3.connect(DB_PATH)
    con.execute("""CREATE TABLE IF NOT EXISTS tbl_mitarbeiter (
        id INTEGER PRIMARY KEY AUTOINCREMENT, vorname TEXT, name TEXT,
        abteilung TEXT, position TEXT, board_x INTEGER DEFAULT 60, board_y INTEGER DEFAULT 60,
        urlaubsanspruch INTEGER DEFAULT 30, personalverantwortlich INTEGER DEFAULT 0)""")
    con.execute("""CREATE TABLE IF NOT EXISTS tbl_mitarbeiter_vorgesetzter (
        id INTEGER PRIMARY KEY AUTOINCREMENT, mitarbeiter_id INTEGER NOT NULL,
        vorgesetzter_id INTEGER NOT NULL, art TEXT DEFAULT 'disziplinarisch',
        ist_primaer INTEGER DEFAULT 0)""")
    con.execute("""CREATE TABLE IF NOT EXISTS tbl_abwesenheit (
        id INTEGER PRIMARY KEY AUTOINCREMENT, mitarbeiter_id INTEGER NOT NULL,
        art TEXT DEFAULT 'Urlaub', von TEXT, bis TEXT, notiz TEXT, unterart TEXT,
        status TEXT DEFAULT 'Beantragt', beantragt_am TEXT,
        entscheider TEXT, entschieden_am TEXT, ablehnung_grund TEXT,
        gf_entscheider TEXT, gf_am TEXT)""")
    # Geteilt mit der Parameter-App: Schalter 'urlaub_gf_freigabe' (2. Stufe)
    con.execute("""CREATE TABLE IF NOT EXISTS tbl_param_config (
        schluessel TEXT PRIMARY KEY, wert TEXT)""")
    con.execute("""CREATE TABLE IF NOT EXISTS tbl_urlaub_verfall (
        id INTEGER PRIMARY KEY AUTOINCREMENT, mitarbeiter_id INTEGER NOT NULL,
        jahr INTEGER, tage INTEGER, datum TEXT)""")
    con.execute("""CREATE TABLE IF NOT EXISTS tbl_arbeitsbereich (
        id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT, kategorie TEXT)""")
    con.execute("""CREATE TABLE IF NOT EXISTS tbl_kategorie (
        id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT UNIQUE)""")
    con.execute("""CREATE TABLE IF NOT EXISTS tbl_mitarbeiter_arbeitsbereich (
        mitarbeiter_id INTEGER NOT NULL, bereich_id INTEGER NOT NULL,
        stufe INTEGER DEFAULT 0,
        PRIMARY KEY (mitarbeiter_id, bereich_id))""")
    # Migration: neue Spalten ergänzen, falls Test-DB älter ist
    mcols = {r[1] for r in con.execute("PRAGMA table_info(tbl_mitarbeiter)").fetchall()}
    if "urlaubsanspruch" not in mcols:
        con.execute("ALTER TABLE tbl_mitarbeiter ADD COLUMN urlaubsanspruch INTEGER DEFAULT 30")
    if "personalverantwortlich" not in mcols:
        con.execute("ALTER TABLE tbl_mitarbeiter ADD COLUMN personalverantwortlich INTEGER DEFAULT 0")
    acols = {r[1] for r in con.execute("PRAGMA table_info(tbl_abwesenheit)").fetchall()}
    if "unterart" not in acols:
        con.execute("ALTER TABLE tbl_abwesenheit ADD COLUMN unterart TEXT")
    # Genehmigungs-Workflow: neue Spalten nachrüsten
    for col, ddl in [
        ("status", "status TEXT"), ("beantragt_am", "beantragt_am TEXT"),
        ("entscheider", "entscheider TEXT"), ("entschieden_am", "entschieden_am TEXT"),
        ("ablehnung_grund", "ablehnung_grund TEXT"),
        ("gf_entscheider", "gf_entscheider TEXT"), ("gf_am", "gf_am TEXT"),
    ]:
        if col not in acols:
            con.execute(f"ALTER TABLE tbl_abwesenheit ADD COLUMN {ddl}")
    # Altbestand (vor Einführung des Workflows) gilt als bereits genehmigt
    con.execute("UPDATE tbl_abwesenheit SET status='Genehmigt' WHERE status IS NULL OR status=''")
    try:
        zcols = {r[1] for r in con.execute("PRAGMA table_info(tbl_mitarbeiter_arbeitsbereich)").fetchall()}
        if zcols and "stufe" not in zcols:
            con.execute("ALTER TABLE tbl_mitarbeiter_arbeitsbereich ADD COLUMN stufe INTEGER DEFAULT 0")
    except Exception:
        pass
    if False:  # keine Demo-Daten in der echten NMGone-Datenbank
        for i, v, n, ab, po, x, y in DEMO_EMPS:
            con.execute("INSERT INTO tbl_mitarbeiter(id,vorname,name,abteilung,position,board_x,board_y) "
                        "VALUES(?,?,?,?,?,?,?)", (i, v, n, ab, po, x, y))
        # Demo: Hans Berg (Verwaltung) ist als personalverantwortlich eingeteilt
        con.execute("UPDATE tbl_mitarbeiter SET personalverantwortlich=1 WHERE id=8")
        for c, p, a, pr in DEMO_LINKS:
            con.execute("INSERT INTO tbl_mitarbeiter_vorgesetzter(mitarbeiter_id,vorgesetzter_id,art,ist_primaer) "
                        "VALUES(?,?,?,?)", (c, p, a, pr))
        # Demo-Abwesenheiten im aktuellen Monat
        t = date.today()
        y, m = t.year, t.month
        def dd(day): return date(y, m, _clamp_day(y, m, day)).isoformat()
        for mid, art, a, b in [
            (4, "Urlaub", 3, 9), (5, "Krankheit", 10, 12), (2, "Fortbildung", 17, 18),
            (6, "Urlaub", 20, 27), (7, "Sonstiges", 5, 5), (8, "Urlaub", 24, 28),
        ]:
            con.execute("INSERT INTO tbl_abwesenheit(mitarbeiter_id,art,von,bis) VALUES(?,?,?,?)",
                        (mid, art, dd(a), dd(b)))
        con.execute("INSERT INTO tbl_abwesenheit(mitarbeiter_id,art,von,bis,unterart) VALUES(?,?,?,?,?)",
                    (3, "Sonderurlaub", dd(14), dd(15), "Umzug (betrieblich veranlasst)"))
        # Demo-Arbeitsbereiche (id, name, kategorie)
        bereiche = [
            (1, "Apothekenbetreuung Nord", "Vertrieb"), (2, "Apothekenbetreuung Süd", "Vertrieb"),
            (3, "Key Accounts", "Vertrieb"), (4, "Außendienst-Koordination", "Vertrieb"),
            (5, "Messen & Events", "Marketing"), (6, "Social Media", "Marketing"),
            (7, "Newsletter", "Marketing"), (8, "Preislisten-Pflege", "Innendienst"),
            (9, "Retouren-Abwicklung", "Innendienst"), (10, "Musterversand", "Innendienst"),
            (11, "Qualitätskontrolle", "Labor"), (12, "Stabilitätsprüfung", "Labor"),
            (13, "Hygienemanagement", "Labor"), (14, "Gefahrstoffe", "Labor"),
            (15, "Lieferantenbewertung", "Einkauf"), (16, "Buchhaltung", "Verwaltung"),
            (17, "Datenschutz", "Verwaltung"), (18, "IT-Support", "Verwaltung"),
        ]
        for bid, bname, kat in bereiche:
            con.execute("INSERT INTO tbl_arbeitsbereich(id,name,kategorie) VALUES(?,?,?)", (bid, bname, kat))
        # Zuordnungen (mitarbeiter, bereich, stufe): 0=Verantwortlich, 1..3=Vertretung
        zuord = [(2, b, 0) for b in (1, 2, 3, 4, 5, 6, 7, 8, 9, 10)] + \
                [(3, b, 0) for b in (11, 12, 13, 14)] + \
                [(8, b, 0) for b in (16, 17, 18)] + [(1, 15, 0)] + [
            # Vertretungskette Bereich 1 (Apothekenbetreuung Nord): Ben -> David -> Eva -> Felix
            (4, 1, 1), (5, 1, 2), (6, 1, 3),
            # Bereich 3 (Key Accounts): Ben -> Eva -> David
            (5, 3, 1), (4, 3, 2),
            # Bereich 5 (Messen & Events): Ben -> Anna
            (1, 5, 1),
            # Bereich 11 (Qualitätskontrolle): Carla -> Felix -> Gina
            (6, 11, 1), (7, 11, 2),
            # Bereich 16 (Buchhaltung): Hans -> Anna
            (1, 16, 1),
        ]
        for mid, bid, stufe in zuord:
            con.execute("INSERT OR IGNORE INTO tbl_mitarbeiter_arbeitsbereich(mitarbeiter_id,bereich_id,stufe) VALUES(?,?,?)",
                        (mid, bid, stufe))
        con.commit()
    # Kategorie-Stammliste füllen (frisch ODER migriert): aus vorhandenen Bereichen ableiten
    if not con.execute("SELECT COUNT(*) FROM tbl_kategorie").fetchone()[0]:
        for (kn,) in con.execute("SELECT DISTINCT kategorie FROM tbl_arbeitsbereich "
                                 "WHERE kategorie IS NOT NULL AND kategorie<>''").fetchall():
            con.execute("INSERT OR IGNORE INTO tbl_kategorie(name) VALUES(?)", (kn,))
        con.commit()
    con.close()


class App:
    def __init__(self, root):
        self.root = root
        root.title(f"NMGone{DEMO_SUFFIX} · Mitarbeiter & Personal")
        root.geometry("1200x780")
        # Im Vollbild (maximiert) starten. 'zoomed' = Windows; sonst -zoomed/Bildschirmgroesse.
        try:
            root.state("zoomed")
        except tk.TclError:
            try:
                root.attributes("-zoomed", True)
            except tk.TclError:
                root.geometry(f"{root.winfo_screenwidth()}x{root.winfo_screenheight()}+0+0")
        root.configure(bg=BG)
        self.style = ttk.Style(root)
        try:
            self.style.theme_use("clam")
        except Exception:
            pass
        self.style.configure("TCombobox", fieldbackground=CARD, background=CARD,
                             foreground=INK, arrowcolor=PRIMARY, padding=5)

        self.view = "orga"
        self.connect_mode = False
        self.connect_source = None
        self.drag = None
        self.dept_color = {}
        t = date.today()
        self.cal_year, self.cal_month = t.year, t.month
        self.cal_mode = "monat"   # "monat" (Grundeinstellung) oder "jahr"
        self.ab_mode = "ma"       # Arbeitsbereiche: "ma" (nach Mitarbeiter) oder "bereich"
        # In der echten App: der per Login angemeldete Mitarbeiter. Hier ersatzweise
        # der Windows-Benutzer als Vorbelegung für "Neue Karte".
        try:
            self.current_user = getpass.getuser()
        except Exception:
            self.current_user = ""

        self._build_header()
        self.main = tk.Frame(root, bg=BG)
        self.main.pack(side="top", fill="both", expand=True)
        self._build_status()

        self.load()
        self.show_orga()

    # ===== Header mit Ansichts-Umschalter =====
    def _build_header(self):
        h = tk.Frame(self.root, bg=SIDEBAR, height=64)
        h.pack(side="top", fill="x")
        h.pack_propagate(False)
        tk.Label(h, text="👥  Mitarbeiter-Board", bg=SIDEBAR, fg="white",
                 font=(FONT, 18, "bold")).pack(side="left", padx=22)
        tk.Label(h, text="Organigramm · Abwesenheiten · Arbeitsbereiche",
                 bg=SIDEBAR, fg=SIDEBAR_TEXT, font=(FONT, 10)).pack(side="left", padx=4)
        tk.Button(h, text="⚙", command=self._settings_dialog, bg=SIDEBAR, fg="white",
                  relief="flat", font=(FONT, 17), padx=10, cursor="hand2",
                  activebackground=SIDEBAR, activeforeground="white").pack(side="right", padx=(0, 18))
        sw = tk.Frame(h, bg=SIDEBAR)
        sw.pack(side="right", padx=4)
        self.nav_orga = tk.Button(sw, text="🗺  Organigramm", command=self.show_orga,
                                  relief="flat", font=(FONT, 10, "bold"), padx=14, pady=7, cursor="hand2")
        self.nav_ab = tk.Button(sw, text="🧩  Arbeitsbereiche", command=self.show_arbeitsbereiche,
                                relief="flat", font=(FONT, 10, "bold"), padx=14, pady=7, cursor="hand2")
        self.nav_abw = tk.Button(sw, text="🗓  Abwesenheiten", command=self.show_absence,
                                 relief="flat", font=(FONT, 10, "bold"), padx=14, pady=7, cursor="hand2")
        self.nav_antr = tk.Button(sw, text="✅  Anträge", command=self.show_antraege,
                                  relief="flat", font=(FONT, 10, "bold"), padx=14, pady=7, cursor="hand2")
        self.nav_orga.pack(side="left", padx=4)
        self.nav_ab.pack(side="left", padx=4)
        self.nav_abw.pack(side="left", padx=4)
        self.nav_antr.pack(side="left", padx=4)

    def _update_antraege_badge(self):
        if not hasattr(self, "nav_antr"):
            return
        n = len(self._open_antraege())
        self.nav_antr.configure(text=f"✅  Anträge ({n})" if n else "✅  Anträge")

    def _nav_state(self):
        for btn, name in ((self.nav_orga, "orga"), (self.nav_ab, "ab"),
                          (self.nav_abw, "abw"), (self.nav_antr, "antr")):
            if self.view == name:
                btn.configure(bg="white", fg=PRIMARY)
            else:
                btn.configure(bg=SIDEBAR, fg=SIDEBAR_TEXT, activebackground=SIDEBAR)

    def _build_status(self):
        self.status = tk.StringVar(value="Bereit.")
        s = tk.Frame(self.root, bg="#E3EAF1", height=26)
        s.pack(side="bottom", fill="x")
        tk.Label(s, textvariable=self.status, bg="#E3EAF1", fg=MUTED,
                 font=(FONT, 9), anchor="w").pack(side="left", padx=14)

    def _clear_main(self):
        for w in self.main.winfo_children():
            w.destroy()

    # ===== Daten =====
    def load(self):
        con = sqlite3.connect(DB_PATH)
        con.row_factory = sqlite3.Row
        self.emps = [dict(r) for r in con.execute(
            "SELECT * FROM tbl_mitarbeiter ORDER BY name, vorname").fetchall()]
        self.links = [dict(r) for r in con.execute(
            "SELECT id, mitarbeiter_id AS child, vorgesetzter_id AS parent, art, ist_primaer "
            "FROM tbl_mitarbeiter_vorgesetzter").fetchall()]
        self.abw = [dict(r) for r in con.execute(
            "SELECT * FROM tbl_abwesenheit").fetchall()]
        self.verfall = [dict(r) for r in con.execute(
            "SELECT * FROM tbl_urlaub_verfall").fetchall()]
        self.bereiche = [dict(r) for r in con.execute(
            "SELECT * FROM tbl_arbeitsbereich ORDER BY kategorie, name").fetchall()]
        self.ab_zuord = [dict(r) for r in con.execute(
            "SELECT * FROM tbl_mitarbeiter_arbeitsbereich").fetchall()]
        self.kategorien = [r["name"] for r in con.execute(
            "SELECT name FROM tbl_kategorie ORDER BY name").fetchall()]
        con.close()
        for e in self.emps:
            e["x"] = e["board_x"] or 60
            e["y"] = e["board_y"] or 60
        self.emp_by_id = {e["id"]: e for e in self.emps}
        depts = sorted({e["abteilung"] or "—" for e in self.emps})
        self.dept_color = {d: DEPT_PALETTE[i % len(DEPT_PALETTE)] for i, d in enumerate(depts)}
        self._update_antraege_badge()

    def _name(self, eid):
        e = self.emp_by_id.get(eid, {})
        return f"{e.get('vorname','')} {e.get('name','')}".strip() or f"#{eid}"

    def _depts(self):
        return sorted({e["abteilung"] or "—" for e in self.emps})

    # =========================================================================
    #  ANSICHT 1 · ORGANIGRAMM
    # =========================================================================
    def show_orga(self):
        self.view = "orga"
        self._nav_state()
        self._clear_main()
        self.connect_mode = False
        self.connect_source = None

        t = tk.Frame(self.main, bg=CARD, height=52)
        t.pack(side="top", fill="x")
        t.configure(highlightbackground=BORDER, highlightthickness=1)
        tk.Label(t, text="Teilbereich:", bg=CARD, fg=MUTED, font=(FONT, 10, "bold")).pack(side="left", padx=(16, 6), pady=10)
        self.bereich = ttk.Combobox(t, state="readonly", width=22,
                                    values=["Gesamtunternehmen"] + self._depts())
        self.bereich.set("Gesamtunternehmen")
        self.bereich.pack(side="left", pady=10)
        self.bereich.bind("<<ComboboxSelected>>", lambda e: self.redraw_orga())
        self.connect_btn = tk.Button(t, text="🔗  Verbinden", command=self.toggle_connect,
                                     bg="#EDF1F6", fg=PRIMARY, relief="flat", font=(FONT, 10, "bold"),
                                     padx=14, pady=7, cursor="hand2")
        self.connect_btn.pack(side="left", padx=(16, 6), pady=8)
        tk.Label(t, text="Art:", bg=CARD, fg=MUTED, font=(FONT, 10, "bold")).pack(side="left", padx=(8, 4))
        self.art = ttk.Combobox(t, state="readonly", width=16, values=list(ARTEN))
        self.art.set(ARTEN[0])
        self.art.pack(side="left", pady=10)
        self._tbtn(t, "✦  Auto-Layout", self.auto_layout)
        self._tbtn(t, "➕  Neue Karte", lambda: self._card_form(None))
        leg = tk.Frame(t, bg=CARD)
        leg.pack(side="right", padx=16)
        for art in ARTEN:
            self._legend_line(leg, ART_STYLE[art]["color"], ART_STYLE[art]["dash"],
                              ART_STYLE[art]["width"], art)

        wrap = tk.Frame(self.main, bg=BG)
        wrap.pack(side="top", fill="both", expand=True)
        self.canvas = tk.Canvas(wrap, bg=BG, highlightthickness=0)
        hsb = ttk.Scrollbar(wrap, orient="horizontal", command=self.canvas.xview)
        vsb = ttk.Scrollbar(wrap, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(xscrollcommand=hsb.set, yscrollcommand=vsb.set,
                              scrollregion=(0, 0, 1600, 1100))
        self.canvas.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        wrap.rowconfigure(0, weight=1)
        wrap.columnconfigure(0, weight=1)
        self.redraw_orga()

    def _tbtn(self, parent, text, cmd):
        tk.Button(parent, text=text, command=cmd, bg=CARD, fg=INK, relief="flat",
                  font=(FONT, 10), padx=10, pady=7, cursor="hand2",
                  activebackground="#EDF1F6").pack(side="left", padx=4, pady=8)

    def _legend_line(self, parent, color, dash, width, label):
        box = tk.Frame(parent, bg=CARD)
        box.pack(side="left", padx=8)
        cv = tk.Canvas(box, width=30, height=12, bg=CARD, highlightthickness=0)
        cv.create_line(2, 6, 28, 6, fill=color, width=width, dash=dash)
        cv.pack(side="left")
        tk.Label(box, text=label, bg=CARD, fg=MUTED, font=(FONT, 9)).pack(side="left", padx=(4, 0))

    def visible_ids(self):
        b = self.bereich.get() if hasattr(self, "bereich") else "Gesamtunternehmen"
        if b == "Gesamtunternehmen":
            return {e["id"] for e in self.emps}
        return {e["id"] for e in self.emps if (e["abteilung"] or "—") == b}

    def primary_boss(self, eid):
        for lk in self.links:
            if lk["child"] == eid and lk["ist_primaer"]:
                p = self.emp_by_id.get(lk["parent"])
                if p:
                    return f"{p['vorname']} {p['name']}".strip(), lk["art"]
        return None, None

    def redraw_orga(self):
        self.canvas.delete("all")
        self._draw_grid()
        for e in self.emps:
            if e["id"] in self.visible_ids():
                self.draw_card(e)
        self.draw_connections()
        vis = len(self.visible_ids())
        self.status.set(f"{vis} Karten sichtbar · {self.bereich.get()}"
                        + ("   ·   Verbinden aktiv – Start- und Ziel-Karte anklicken" if self.connect_mode else ""))

    def _draw_grid(self):
        for x in range(0, 1600, 28):
            self.canvas.create_line(x, 0, x, 1100, fill=GRID, tags=("grid",))
        for y in range(0, 1100, 28):
            self.canvas.create_line(0, y, 1600, y, fill=GRID, tags=("grid",))
        self.canvas.tag_lower("grid")

    def _round_rect(self, x1, y1, x2, y2, r, **kw):
        pts = [x1 + r, y1, x2 - r, y1, x2, y1, x2, y1 + r, x2, y2 - r, x2, y2,
               x2 - r, y2, x1 + r, y2, x1, y2, x1, y2 - r, x1, y1 + r, x1, y1]
        return self.canvas.create_polygon(pts, smooth=True, **kw)

    def draw_card(self, e):
        x, y = e["x"], e["y"]
        tag = f"c{e['id']}"
        col = self.dept_color.get(e["abteilung"] or "—", PRIMARY)
        sel = (e["id"] == self.connect_source)
        self._round_rect(x + 3, y + 4, x + CARD_W + 3, y + CARD_H + 4, 14, fill=SHADOW, outline="", tags=(tag, "card"))
        self._round_rect(x, y, x + CARD_W, y + CARD_H, 14, fill=CARD,
                         outline=(col if sel else BORDER), width=(3 if sel else 1), tags=(tag, "card"))
        self.canvas.create_rectangle(x + 1, y + 14, x + 7, y + CARD_H - 14, fill=col, outline="", tags=(tag, "card"))
        ini = (e["vorname"][:1] + e["name"][:1]).upper()
        self.canvas.create_oval(x + 18, y + 18, x + 52, y + 52, fill=col, outline="", tags=(tag, "card"))
        self.canvas.create_text(x + 35, y + 35, text=ini, fill="white", font=(FONT, 13, "bold"), tags=(tag, "card"))
        name = f"{e['vorname']} {e['name']}".strip()
        self.canvas.create_text(x + 62, y + 24, text=name, anchor="w", fill=INK, font=(FONT, 12, "bold"), tags=(tag, "card"))
        self.canvas.create_text(x + 62, y + 44, text=e["position"] or "", anchor="w", fill=MUTED, font=(FONT, 10), tags=(tag, "card"))
        self.canvas.create_text(x + 62, y + 60, text=(e["abteilung"] or "").upper(), anchor="w", fill=col, font=(FONT, 8, "bold"), tags=(tag, "card"))
        boss, art = self.primary_boss(e["id"])
        if boss:
            self.canvas.create_text(x + 14, y + CARD_H - 13, text=f"▸ Vorgesetzt: {boss}", anchor="w",
                                    fill=FAINT, font=(FONT, 8), tags=(tag, "card"))
        if e.get("personalverantwortlich"):
            self.canvas.create_text(x + CARD_W - 12, y + 14, text="🔑", anchor="e",
                                    font=(FONT, 11), tags=(tag, "card"))
        self.canvas.tag_bind(tag, "<ButtonPress-1>", lambda ev, i=e["id"]: self.on_press(ev, i))
        self.canvas.tag_bind(tag, "<B1-Motion>", lambda ev, i=e["id"]: self.on_motion(ev, i))
        self.canvas.tag_bind(tag, "<ButtonRelease-1>", lambda ev, i=e["id"]: self.on_release(ev, i))
        self.canvas.tag_bind(tag, "<Double-Button-1>", lambda ev, i=e["id"]: self._card_form(self.emp_by_id[i]))

    def draw_connections(self):
        self.canvas.delete("conn")
        vis = self.visible_ids()
        for lk in self.links:
            if lk["child"] not in vis or lk["parent"] not in vis:
                continue
            k = self.emp_by_id[lk["child"]]
            p = self.emp_by_id[lk["parent"]]
            st = ART_STYLE.get(lk["art"], ART_STYLE["disziplinarisch"])
            x1, y1 = k["x"] + CARD_W / 2, k["y"]
            x2, y2 = p["x"] + CARD_W / 2, p["y"] + CARD_H
            self.canvas.create_line(x1, y1, x2, y2, fill=st["color"], width=st["width"],
                                    dash=st["dash"], smooth=False,
                                    arrow=("last" if st["arrow"] else "none"),
                                    arrowshape=(12, 14, 5), tags=("conn",))
            mx, my = (x1 + x2) / 2, (y1 + y2) / 2
            kn = f"kn{lk['id']}"
            self.canvas.create_oval(mx - 9, my - 9, mx + 9, my + 9, fill="white",
                                    outline=st["color"], width=1.5, tags=("conn", kn))
            self.canvas.create_text(mx, my, text="×", fill=st["color"], font=(FONT, 12, "bold"), tags=("conn", kn))
            self.canvas.tag_bind(kn, "<Button-1>", lambda ev, l=lk: self.remove_link(l))
        if self.canvas.find_withtag("card"):
            self.canvas.tag_raise("card")
            if self.canvas.find_withtag("conn"):
                self.canvas.tag_raise("conn", "card")

    def on_press(self, ev, eid):
        self.drag = {"id": eid, "x": self.canvas.canvasx(ev.x), "y": self.canvas.canvasy(ev.y), "moved": False}

    def on_motion(self, ev, eid):
        if not self.drag or self.drag["id"] != eid or self.connect_mode:
            return
        cx, cy = self.canvas.canvasx(ev.x), self.canvas.canvasy(ev.y)
        dx, dy = cx - self.drag["x"], cy - self.drag["y"]
        if abs(dx) > 2 or abs(dy) > 2:
            self.drag["moved"] = True
        self.canvas.move(f"c{eid}", dx, dy)
        e = self.emp_by_id[eid]
        e["x"] += dx
        e["y"] += dy
        self.drag["x"], self.drag["y"] = cx, cy
        self.draw_connections()

    def on_release(self, ev, eid):
        if not self.drag:
            return
        moved = self.drag["moved"]
        self.drag = None
        if moved:
            e = self.emp_by_id[eid]
            con = sqlite3.connect(DB_PATH)
            con.execute("UPDATE tbl_mitarbeiter SET board_x=?, board_y=? WHERE id=?",
                        (int(e["x"]), int(e["y"]), eid))
            con.commit()
            con.close()
            self.status.set("Position gespeichert.")
        elif self.connect_mode:
            self.pick_for_connection(eid)

    def toggle_connect(self):
        self.connect_mode = not self.connect_mode
        self.connect_source = None
        if self.connect_mode:
            self.connect_btn.configure(bg=PRIMARY, fg="white", text="✓  Verbinden aktiv")
        else:
            self.connect_btn.configure(bg="#EDF1F6", fg=PRIMARY, text="🔗  Verbinden")
        self.redraw_orga()

    def pick_for_connection(self, eid):
        if self.connect_source is None:
            self.connect_source = eid
            self.redraw_orga()
            self.status.set(f"Start: {self._name(eid)} – jetzt Ziel-(Vorgesetzten-)Karte anklicken.")
            return
        if eid == self.connect_source:
            return
        child, parent = self.connect_source, eid
        art = self.art.get()
        primaer = 1 if messagebox.askyesno("Verbindung",
            f"„{self._name(child)} → {self._name(parent)}“ ({art})\n\n"
            "Als primäre (organigramm-bildende) Beziehung markieren?") else 0
        con = sqlite3.connect(DB_PATH)
        if primaer:
            con.execute("UPDATE tbl_mitarbeiter_vorgesetzter SET ist_primaer=0 WHERE mitarbeiter_id=?", (child,))
        cur = con.execute("INSERT INTO tbl_mitarbeiter_vorgesetzter(mitarbeiter_id,vorgesetzter_id,art,ist_primaer) "
                          "VALUES(?,?,?,?)", (child, parent, art, primaer))
        con.commit()
        con.close()
        self.links.append({"id": cur.lastrowid, "child": child, "parent": parent, "art": art, "ist_primaer": primaer})
        self.connect_source = None
        self.connect_mode = False
        self.connect_btn.configure(bg="#EDF1F6", fg=PRIMARY, text="🔗  Verbinden")
        self.redraw_orga()
        self.status.set(f"Verbindung angelegt: {self._name(child)} → {self._name(parent)} ({art}).")

    def remove_link(self, lk):
        if not messagebox.askyesno("Verbindung entfernen",
            f"Verbindung „{self._name(lk['child'])} → {self._name(lk['parent'])}“ ({lk['art']}) entfernen?"):
            return
        con = sqlite3.connect(DB_PATH)
        con.execute("DELETE FROM tbl_mitarbeiter_vorgesetzter WHERE id=?", (lk["id"],))
        con.commit()
        con.close()
        self.links = [x for x in self.links if x["id"] != lk["id"]]
        self.redraw_orga()
        self.status.set("Verbindung entfernt.")

    def auto_layout(self):
        vis = self.visible_ids()
        parent_of = {l["child"]: l["parent"] for l in self.links if l["ist_primaer"]}
        kids = {i: [] for i in vis}
        for i in vis:
            p = parent_of.get(i)
            if p in vis:
                kids[p].append(i)
        roots = [i for i in vis if parent_of.get(i) not in vis]
        level, order = {}, []

        def dfs(i, d):
            level[i] = d
            order.append(i)
            for k in sorted(kids[i], key=lambda j: self.emp_by_id[j]["name"]):
                dfs(k, d + 1)
        for r in sorted(roots, key=lambda j: self.emp_by_id[j]["name"]):
            dfs(r, 0)
        per = {}
        con = sqlite3.connect(DB_PATH)
        for i in order:
            d = level[i]
            per[d] = per.get(d, 0)
            e = self.emp_by_id[i]
            e["x"] = 60 + per[d] * (CARD_W + 40)
            e["y"] = 40 + d * (CARD_H + 70)
            per[d] += 1
            con.execute("UPDATE tbl_mitarbeiter SET board_x=?, board_y=? WHERE id=?", (int(e["x"]), int(e["y"]), i))
        con.commit()
        con.close()
        self.redraw_orga()
        self.status.set("Auto-Layout angewendet.")

    def _card_form(self, e):
        win = tk.Toplevel(self.root)
        win.title("Mitarbeiter bearbeiten" if e else "Neue Karte")
        win.configure(bg=BG)
        win.transient(self.root)
        win.grab_set()
        win.geometry("470x450")
        # Vorbelegung bei NEUER Karte mit dem angemeldeten Benutzer.
        # In der echten App käme hier das volle Login-Profil (Abteilung, Position …).
        src = dict(e) if e else {}
        if not e and self.current_user:
            u = self.current_user.replace("_", ".").replace("-", ".")
            if "." in u:
                vn, _, nn = u.partition(".")
                src = {"vorname": vn.capitalize(), "name": nn.capitalize()}
            else:
                src = {"vorname": self.current_user}
        fields = [("vorname", "Vorname"), ("name", "Nachname"), ("abteilung", "Abteilung"),
                  ("position", "Position"), ("urlaubsanspruch", "Urlaubsanspruch (Tage/Jahr)")]
        if not e:
            tk.Label(win, text=f"Vorbelegt mit angemeldetem Benutzer: {self.current_user or '—'} · bitte prüfen.",
                     bg=BG, fg=MUTED, font=(FONT, 9), wraplength=420, justify="left").grid(
                row=0, column=0, columnspan=2, sticky="w", padx=16, pady=(12, 2))
        base = 1 if not e else 0
        vars_ = {}
        for i, (k, lab) in enumerate(fields):
            r = base + i
            tk.Label(win, text=lab, bg=BG, fg=PRIMARY, font=(FONT, 10, "bold")).grid(row=r, column=0, sticky="w", padx=16, pady=9)
            default = STD_URLAUB if (k == "urlaubsanspruch" and not e) else src.get(k, "")
            v = tk.StringVar(value=str(default if default not in (None, "") else ("" if k != "urlaubsanspruch" else STD_URLAUB)))
            vars_[k] = v
            tk.Entry(win, textvariable=v, width=28).grid(row=r, column=1, sticky="ew", padx=16, pady=9)
        win.columnconfigure(1, weight=1)
        pv_var = tk.BooleanVar(value=bool(src.get("personalverantwortlich")))
        pv_row = base + len(fields)
        tk.Checkbutton(win, text="Personalverantwortlich (entscheidet über Urlaub / Verfall)",
                       variable=pv_var, bg=BG, fg=INK, font=(FONT, 10), anchor="w").grid(
            row=pv_row, column=0, columnspan=2, sticky="w", padx=14, pady=(4, 2))
        bar_row = pv_row + 1

        def save():
            data = {k: v.get().strip() for k, v in vars_.items()}
            try:
                anspruch = int(data.get("urlaubsanspruch") or STD_URLAUB)
            except ValueError:
                anspruch = STD_URLAUB
            pv = 1 if pv_var.get() else 0
            con = sqlite3.connect(DB_PATH)
            if e:
                con.execute("UPDATE tbl_mitarbeiter SET vorname=?,name=?,abteilung=?,position=?,urlaubsanspruch=?,personalverantwortlich=? WHERE id=?",
                            (data["vorname"], data["name"], data["abteilung"], data["position"], anspruch, pv, e["id"]))
            else:
                con.execute("INSERT INTO tbl_mitarbeiter(vorname,name,abteilung,position,board_x,board_y,urlaubsanspruch,personalverantwortlich) "
                            "VALUES(?,?,?,?,?,?,?,?)",
                            (data["vorname"], data["name"], data["abteilung"], data["position"], 80, 80, anspruch, pv))
            con.commit()
            con.close()
            win.destroy()
            self.load()
            self.show_orga()

        bar = tk.Frame(win, bg=BG)
        bar.grid(row=bar_row, column=0, columnspan=2, sticky="ew", padx=16, pady=16)
        tk.Button(bar, text="Abbrechen", command=win.destroy, padx=14, pady=7).pack(side="right", padx=(8, 0))
        tk.Button(bar, text="Speichern", command=save, bg=PRIMARY, fg="white", relief="flat", padx=18, pady=8).pack(side="right")

    # =========================================================================
    #  ANSICHT 2 · ABWESENHEITEN (Team-Kalender)
    # =========================================================================
    NAME_W, DAY_W, ROW_H, HEAD_H = 210, 32, 40, 46

    def show_absence(self):
        self.view = "abw"
        self._nav_state()
        self._clear_main()

        t = tk.Frame(self.main, bg=CARD, height=52)
        t.pack(side="top", fill="x")
        t.configure(highlightbackground=BORDER, highlightthickness=1)
        tk.Button(t, text="◀", command=lambda: self._month_step(-1), bg=CARD, fg=PRIMARY,
                  relief="flat", font=(FONT, 12, "bold"), padx=10, pady=4, cursor="hand2").pack(side="left", padx=(14, 2), pady=8)
        self.month_lbl = tk.Label(t, text="", bg=CARD, fg=INK, font=(FONT, 12, "bold"), width=16)
        self.month_lbl.pack(side="left")
        tk.Button(t, text="▶", command=lambda: self._month_step(1), bg=CARD, fg=PRIMARY,
                  relief="flat", font=(FONT, 12, "bold"), padx=10, pady=4, cursor="hand2").pack(side="left", padx=(2, 14), pady=8)
        tk.Button(t, text="Heute", command=self._month_today, bg=CARD, fg=MUTED,
                  relief="flat", font=(FONT, 10), padx=8, pady=6, cursor="hand2").pack(side="left")

        # Monat / Jahr umschalten (Grundeinstellung: Monat)
        self.mode_monat = tk.Button(t, text="Monat", command=lambda: self._set_cal_mode("monat"),
                                    relief="flat", font=(FONT, 10, "bold"), padx=12, pady=6, cursor="hand2")
        self.mode_jahr = tk.Button(t, text="Jahr", command=lambda: self._set_cal_mode("jahr"),
                                   relief="flat", font=(FONT, 10, "bold"), padx=12, pady=6, cursor="hand2")
        self.mode_monat.pack(side="left", padx=(16, 2))
        self.mode_jahr.pack(side="left", padx=(0, 4))
        self.mode_monat.configure(bg=(PRIMARY if self.cal_mode == "monat" else "#EDF1F6"),
                                  fg=("white" if self.cal_mode == "monat" else PRIMARY))
        self.mode_jahr.configure(bg=(PRIMARY if self.cal_mode == "jahr" else "#EDF1F6"),
                                 fg=("white" if self.cal_mode == "jahr" else PRIMARY))

        tk.Label(t, text="Teilbereich:", bg=CARD, fg=MUTED, font=(FONT, 10, "bold")).pack(side="left", padx=(18, 6))
        self.abw_bereich = ttk.Combobox(t, state="readonly", width=20,
                                        values=["Alle Bereiche"] + self._depts())
        self.abw_bereich.set("Alle Bereiche")
        self.abw_bereich.pack(side="left", pady=10)
        self.abw_bereich.bind("<<ComboboxSelected>>", lambda e: self.draw_calendar())

        tk.Button(t, text="➕  Abwesenheit eintragen", command=self._abw_form, bg=PRIMARY, fg="white",
                  relief="flat", font=(FONT, 10, "bold"), padx=14, pady=7, cursor="hand2").pack(side="left", padx=(16, 6), pady=8)
        tk.Button(t, text="🏖  Resturlaub", command=self._resturlaub_overview, bg="#EDF1F6", fg=PRIMARY,
                  relief="flat", font=(FONT, 10, "bold"), padx=12, pady=7, cursor="hand2").pack(side="left", pady=8)

        leg = tk.Frame(t, bg=CARD)
        leg.pack(side="right", padx=16)
        for art, col in ABW_STYLE.items():
            box = tk.Frame(leg, bg=CARD)
            box.pack(side="left", padx=7)
            tk.Label(box, text="  ", bg=col).pack(side="left")
            tk.Label(box, text=art, bg=CARD, fg=MUTED, font=(FONT, 9)).pack(side="left", padx=(4, 0))

        wrap = tk.Frame(self.main, bg=BG)
        wrap.pack(side="top", fill="both", expand=True)
        self.cal = tk.Canvas(wrap, bg=CARD, highlightthickness=0)
        hsb = ttk.Scrollbar(wrap, orient="horizontal", command=self.cal.xview)
        vsb = ttk.Scrollbar(wrap, orient="vertical", command=self.cal.yview)
        self.cal.configure(xscrollcommand=hsb.set, yscrollcommand=vsb.set)
        self.cal.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        wrap.rowconfigure(0, weight=1)
        wrap.columnconfigure(0, weight=1)
        self.draw_calendar()

    def _month_step(self, delta):
        if self.cal_mode == "jahr":
            self.cal_year += delta
            self.draw_calendar()
            return
        m = self.cal_month + delta
        y = self.cal_year
        if m < 1:
            m, y = 12, y - 1
        elif m > 12:
            m, y = 1, y + 1
        self.cal_month, self.cal_year = m, y
        self.draw_calendar()

    def _month_today(self):
        t = date.today()
        self.cal_year, self.cal_month = t.year, t.month
        self._set_cal_mode("monat")

    def _abw_visible_emps(self):
        b = self.abw_bereich.get() if hasattr(self, "abw_bereich") else "Alle Bereiche"
        if b == "Alle Bereiche":
            return self.emps
        return [e for e in self.emps if (e["abteilung"] or "—") == b]

    def _year_count(self, eid, art, stati=(ST_GENEHMIGT,)):
        days = 0
        for a in self.abw:
            if a["mitarbeiter_id"] != eid or a["art"] != art:
                continue
            if (a.get("status") or ST_GENEHMIGT) not in stati:
                continue
            try:
                v = date.fromisoformat(a["von"])
                b = date.fromisoformat(a["bis"])
            except Exception:
                continue
            d = v
            while d <= b:
                if d.year == self.cal_year and d.weekday() < 5:
                    days += 1
                d = date.fromordinal(d.toordinal() + 1)
        return days

    def _verfall_count(self, eid, jahr):
        return sum(int(v["tage"] or 0) for v in self.verfall
                   if v["mitarbeiter_id"] == eid and v["jahr"] == jahr)

    def _rest_urlaub(self, e, jahr):
        anspruch = e.get("urlaubsanspruch") or STD_URLAUB
        return anspruch - self._year_count(e["id"], "Urlaub") - self._verfall_count(e["id"], jahr)

    def _top_manager(self, eid):
        """Oberste primäre Vorgesetzte über eid (nur noch informativ)."""
        parent_of = {l["child"]: l["parent"] for l in self.links if l["ist_primaer"]}
        seen, cur = set(), eid
        while cur in parent_of and parent_of[cur] not in seen:
            seen.add(cur)
            cur = parent_of[cur]
        return cur if cur != eid else None

    def _personalverantwortliche(self):
        """Namen der als personalverantwortlich eingeteilten Mitarbeiter (entscheiden über Urlaub/Verfall)."""
        return [f"{e['vorname']} {e['name']}".strip() for e in self.emps if e.get("personalverantwortlich")]

    def _geschaeftsfuehrung(self):
        """Namen der Geschäftsführung (für die 2. Freigabe-Stufe)."""
        gf = [f"{e['vorname']} {e['name']}".strip() for e in self.emps
              if "geschäftsführung" in ((e.get("position") or "") + " " + (e.get("abteilung") or "")).lower()]
        return gf or self._personalverantwortliche()

    def _gf_freigabe_aktiv(self):
        """2. Stufe (Geschäftsführung segnet ab) – Schalter aus der Parameter-App."""
        con = sqlite3.connect(DB_PATH)
        try:
            row = con.execute(
                "SELECT wert FROM tbl_param_config WHERE schluessel='urlaub_gf_freigabe'").fetchone()
        except Exception:
            row = None
        con.close()
        return bool(row and str(row[0]) == "1")

    def _open_antraege(self):
        """Alle noch zu entscheidenden Anträge (Beantragt + GF-Freigabe ausstehend)."""
        return [a for a in self.abw if (a.get("status") or ST_GENEHMIGT) in ST_OFFEN]

    def _set_cal_mode(self, mode):
        self.cal_mode = mode
        self.mode_monat.configure(bg=(PRIMARY if mode == "monat" else "#EDF1F6"),
                                  fg=("white" if mode == "monat" else PRIMARY))
        self.mode_jahr.configure(bg=(PRIMARY if mode == "jahr" else "#EDF1F6"),
                                 fg=("white" if mode == "jahr" else PRIMARY))
        self.draw_calendar()

    def draw_calendar(self):
        self.cal.delete("all")
        y = self.cal_year
        mode = self.cal_mode
        emps = self._abw_visible_emps()
        NAME_W, ROW_H, HEAD_H = self.NAME_W, self.ROW_H, self.HEAD_H
        today = date.today()

        if mode == "jahr":
            start, end = date(y, 1, 1), date(y, 12, 31)
            DAY_W = 6
            self.month_lbl.configure(text=f"Jahr {y}")
        else:
            m = self.cal_month
            ndays = calendar.monthrange(y, m)[1]
            start, end = date(y, m, 1), date(y, m, ndays)
            DAY_W = self.DAY_W
            self.month_lbl.configure(text=f"{MONATE[m]} {y}")

        n = (end - start).days + 1

        def dx(d):
            return NAME_W + (d - start).days * DAY_W
        total_w = NAME_W + n * DAY_W
        total_h = HEAD_H + len(emps) * ROW_H + 10
        self.cal.configure(scrollregion=(0, 0, total_w, total_h))

        # Hintergründe
        if mode == "monat":
            for i in range(n):
                d = start + timedelta(i)
                x = dx(d)
                if d.weekday() >= 5:
                    self.cal.create_rectangle(x, 0, x + DAY_W, total_h, fill=WEEKEND, outline="")
                if d == today:
                    self.cal.create_rectangle(x, 0, x + DAY_W, total_h, fill=TODAY, outline="")
        else:
            for mm in range(1, 13):
                if mm % 2 == 0:
                    ms, me = date(y, mm, 1), date(y, mm, calendar.monthrange(y, mm)[1])
                    self.cal.create_rectangle(dx(ms), 0, dx(me) + DAY_W, total_h, fill="#F6F8FB", outline="")
            if start <= today <= end:
                tx = dx(today) + DAY_W / 2
                self.cal.create_line(tx, 0, tx, total_h, fill="#E0A100", width=2)

        # Kopfzeile
        self.cal.create_rectangle(0, 0, total_w, HEAD_H, fill="#EEF3F8", outline="")
        self.cal.create_text(14, HEAD_H / 2, text="Mitarbeiter", anchor="w", fill=PRIMARY, font=(FONT, 11, "bold"))
        if mode == "monat":
            for i in range(n):
                d = start + timedelta(i)
                x = dx(d) + DAY_W / 2
                fg = MUTED if d.weekday() < 5 else "#A0461F"
                self.cal.create_text(x, 14, text=WD[d.weekday()], fill=fg, font=(FONT, 8))
                self.cal.create_text(x, 31, text=str(d.day), fill=(PRIMARY if d == today else INK), font=(FONT, 10, "bold"))
        else:
            for mm in range(1, 13):
                ms, me = date(y, mm, 1), date(y, mm, calendar.monthrange(y, mm)[1])
                cx = (dx(ms) + dx(me) + DAY_W) / 2
                self.cal.create_text(cx, HEAD_H / 2, text=MONATE[mm][:3], fill=INK, font=(FONT, 9, "bold"))

        # vertikale Trennlinien
        if mode == "monat":
            for col in range(n + 1):
                x = NAME_W + col * DAY_W
                self.cal.create_line(x, 0, x, total_h, fill=GRID)
        else:
            for mm in range(1, 13):
                x = dx(date(y, mm, 1))
                self.cal.create_line(x, 0, x, total_h, fill=GRID)
            self.cal.create_line(total_w, 0, total_w, total_h, fill=GRID)
        self.cal.create_line(NAME_W, 0, NAME_W, total_h, fill=BORDER, width=1)

        # Zeilen + Namen
        for idx, e in enumerate(emps):
            ry = HEAD_H + idx * ROW_H
            if idx % 2:
                self.cal.create_rectangle(0, ry, NAME_W, ry + ROW_H, fill="#F8FAFC", outline="")
            self.cal.create_line(0, ry, total_w, ry, fill=GRID)
            col = self.dept_color.get(e["abteilung"] or "—", PRIMARY)
            self.cal.create_rectangle(0, ry + 6, 5, ry + ROW_H - 6, fill=col, outline="")
            self.cal.create_text(14, ry + 15, text=f"{e['vorname']} {e['name']}".strip(),
                                 anchor="w", fill=INK, font=(FONT, 10, "bold"))
            gen = self._year_count(e["id"], "Urlaub")
            anspruch = e.get("urlaubsanspruch") or STD_URLAUB
            rest = self._rest_urlaub(e, y)
            kcount = self._year_count(e["id"], "Krankheit")
            offen = self._year_count(e["id"], "Urlaub", stati=ST_OFFEN)
            rest_col = "#C0392B" if rest < 0 else MUTED
            offen_txt = f" · {offen} offen" if offen else ""
            self.cal.create_text(14, ry + 31, anchor="w", fill=rest_col, font=(FONT, 8),
                                 text=f"Urlaub {gen}/{anspruch} · Rest {rest}{offen_txt} · Krank {kcount} ({y})")

        # Abwesenheits-Balken
        row_of = {e["id"]: i for i, e in enumerate(emps)}
        for a in self.abw:
            if a["mitarbeiter_id"] not in row_of:
                continue
            try:
                av = date.fromisoformat(a["von"])
                ab = date.fromisoformat(a["bis"])
            except Exception:
                continue
            if ab < start or av > end:
                continue
            status = a.get("status") or ST_GENEHMIGT
            if status == ST_ABGELEHNT:
                continue  # Abgelehnte sind keine Abwesenheit – nicht im Kalender führen
            s = max(av, start)
            ee = min(ab, end)
            ry = HEAD_H + row_of[a["mitarbeiter_id"]] * ROW_H
            x1 = dx(s) + 1
            x2 = dx(ee) + DAY_W - 1
            color = ABW_STYLE.get(a["art"], "#777")
            tagid = f"abw{a['id']}"
            if status == ST_GENEHMIGT:
                self.cal.create_rectangle(x1, ry + 7, x2, ry + ROW_H - 7, fill=color, outline="", tags=(tagid,))
            else:  # Beantragt / GF-Freigabe ausstehend → schraffiert + Rahmen
                self.cal.create_rectangle(x1, ry + 7, x2, ry + ROW_H - 7, fill=color, outline=color,
                                          width=1, stipple="gray50", tags=(tagid,))
            if x2 - x1 > 46:
                base = a["art"] if mode == "monat" else a["art"][:3]
                txt = base if status == ST_GENEHMIGT else "⏳ " + base
                self.cal.create_text((x1 + x2) / 2, ry + ROW_H / 2, text=txt,
                                     fill="white", font=(FONT, 8, "bold"), tags=(tagid,))
            self.cal.tag_bind(tagid, "<Button-1>", lambda ev, ab_=a: self._abw_click(ab_))
            self.cal.tag_bind(tagid, "<Enter>", lambda ev: self.cal.config(cursor="hand2"))
            self.cal.tag_bind(tagid, "<Leave>", lambda ev: self.cal.config(cursor=""))

        scope = f"Jahr {y}" if mode == "jahr" else f"{MONATE[self.cal_month]} {y}"
        self.status.set(f"Abwesenheiten · {scope} · {len(emps)} Mitarbeiter · "
                        f"Klick auf einen Balken zum Bearbeiten/Löschen.")

    # =========================================================================
    #  ANSICHT 4 · ANTRÄGE (genehmigen / ablehnen)
    # =========================================================================
    def show_antraege(self):
        self.view = "antr"
        self._nav_state()
        self._clear_main()

        gf = self._gf_freigabe_aktiv()
        t = tk.Frame(self.main, bg=CARD, height=52)
        t.pack(side="top", fill="x")
        t.configure(highlightbackground=BORDER, highlightthickness=1)
        tk.Label(t, text="✅  Urlaubs-/Abwesenheitsanträge", bg=CARD, fg=PRIMARY,
                 font=(FONT, 12, "bold")).pack(side="left", padx=16, pady=10)
        stufe_txt = ("Zweistufig: Personalverantwortliche → Geschäftsführung"
                     if gf else "Einstufig: Personalverantwortliche entscheiden")
        tk.Label(t, text=stufe_txt + "  ·  GF-Freigabe in der Parameter-App schaltbar",
                 bg=CARD, fg=FAINT, font=(FONT, 9)).pack(side="left", padx=8)

        wrap = tk.Frame(self.main, bg=BG)
        wrap.pack(side="top", fill="both", expand=True, padx=0, pady=0)

        cols = ("ma", "art", "zeit", "tage", "ein", "status")
        tree = ttk.Treeview(wrap, columns=cols, show="headings", selectmode="browse")
        for c, txt, w, anc in (("ma", "Mitarbeiter", 170, "w"), ("art", "Art", 120, "w"),
                               ("zeit", "Zeitraum", 180, "w"), ("tage", "Tage", 50, "center"),
                               ("ein", "Eingereicht", 120, "w"), ("status", "Status", 170, "w")):
            tree.heading(c, text=txt)
            tree.column(c, width=w, anchor=anc, stretch=(c == "status"))
        tree.tag_configure("beantragt", foreground="#9A6500")
        tree.tag_configure("gf", foreground="#0E7C86")

        offen = sorted(self._open_antraege(), key=lambda a: (a.get("beantragt_am") or "", a["id"]))
        self._antr_by_iid = {}
        for a in offen:
            try:
                tage = sum(1 for d in self._daterange(a["von"], a["bis"]) if d.weekday() < 5)
            except Exception:
                tage = "?"
            art_txt = a["art"] + (f" ({a['unterart']})" if a.get("unterart") else "")
            st_lbl = ST_STYLE.get(a.get("status"), (a.get("status"), MUTED))[0]
            tag = "gf" if a.get("status") == ST_GF else "beantragt"
            iid = tree.insert("", "end", values=(
                self._name(a["mitarbeiter_id"]), art_txt,
                f"{a['von']} – {a['bis']}", tage, a.get("beantragt_am") or "—", st_lbl), tags=(tag,))
            self._antr_by_iid[iid] = a

        vsb = ttk.Scrollbar(wrap, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=vsb.set)
        tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        wrap.rowconfigure(0, weight=1)
        wrap.columnconfigure(0, weight=1)
        self._antr_tree = tree

        if not offen:
            self._empty_overlay(wrap, "Keine offenen Anträge.",
                                "Neue Urlaubs-/Abwesenheitsanträge erscheinen hier zur Entscheidung.")

        def sel_antrag():
            s = tree.selection()
            if not s:
                messagebox.showinfo("Anträge", "Bitte zuerst einen Antrag in der Liste auswählen.", parent=self.root)
                return None
            return self._antr_by_iid.get(s[0])

        tree.bind("<Double-Button-1>", lambda e: (sel_antrag() and self._abw_click(sel_antrag())))

        bar = tk.Frame(self.main, bg=BG)
        bar.pack(side="bottom", fill="x", padx=16, pady=12)
        tk.Button(bar, text="Details …", command=lambda: (sel_antrag() and self._abw_click(sel_antrag())),
                  bg="#EDF1F6", fg=PRIMARY, relief="flat", font=(FONT, 10, "bold"),
                  padx=14, pady=8, cursor="hand2").pack(side="left")

        def do(genehmigen):
            a = sel_antrag()
            if a:
                self._entscheide(a, genehmigen, None)

        tk.Button(bar, text="✓ Genehmigen", command=lambda: do(True), bg="#11823B", fg="white",
                  relief="flat", font=(FONT, 10, "bold"), padx=16, pady=8, cursor="hand2").pack(side="right")
        tk.Button(bar, text="✗ Ablehnen", command=lambda: do(False), bg="#C0392B", fg="white",
                  relief="flat", font=(FONT, 10, "bold"), padx=16, pady=8, cursor="hand2").pack(side="right", padx=(0, 8))

        self.status.set(f"{len(offen)} offene Anträge · {stufe_txt}.")

    def _daterange(self, von, bis):
        v = date.fromisoformat(von)
        b = date.fromisoformat(bis)
        d = v
        while d <= b:
            yield d
            d = date.fromordinal(d.toordinal() + 1)

    def _empty_overlay(self, parent, title, sub):
        ov = tk.Frame(parent, bg=BG)
        ov.place(relx=0.5, rely=0.42, anchor="center")
        tk.Label(ov, text=title, bg=BG, fg=MUTED, font=(FONT, 13, "bold")).pack()
        tk.Label(ov, text=sub, bg=BG, fg=FAINT, font=(FONT, 10), justify="center").pack(pady=(4, 0))

    def _abw_click(self, a):
        win = tk.Toplevel(self.root)
        win.title("Abwesenheit")
        win.configure(bg=BG)
        win.transient(self.root)
        win.grab_set()
        win.geometry("400x300")
        art_txt = a["art"] + (f" – {a['unterart']}" if a.get("unterart") else "")
        info = (f"{self._name(a['mitarbeiter_id'])}\n{art_txt}\n"
                f"{a['von']}  bis  {a['bis']}" + (f"\n{a['notiz']}" if a.get("notiz") else ""))
        tk.Label(win, text=info, bg=BG, fg=INK, font=(FONT, 11), justify="left").pack(padx=20, pady=(18, 6), anchor="w")

        status = a.get("status") or ST_GENEHMIGT
        st_lbl, st_col = ST_STYLE.get(status, (status, MUTED))
        tk.Label(win, text=st_lbl, bg=BG, fg=st_col, font=(FONT, 11, "bold")).pack(padx=20, anchor="w")
        # Entscheidungs-Historie
        hist = []
        if a.get("entscheider"):
            hist.append(f"Stufe 1: {a['entscheider']} ({a.get('entschieden_am') or ''})")
        if a.get("gf_entscheider"):
            hist.append(f"GF-Freigabe: {a['gf_entscheider']} ({a.get('gf_am') or ''})")
        if status == ST_ABGELEHNT and a.get("ablehnung_grund"):
            hist.append(f"Grund: {a['ablehnung_grund']}")
        if hist:
            tk.Label(win, text="\n".join(hist), bg=BG, fg=MUTED, font=(FONT, 9),
                     justify="left").pack(padx=20, pady=(4, 0), anchor="w")

        bar = tk.Frame(win, bg=BG)
        bar.pack(side="bottom", fill="x", padx=16, pady=14)

        def loeschen():
            if not messagebox.askyesno("Löschen", "Diesen Eintrag wirklich löschen?", parent=win):
                return
            con = sqlite3.connect(DB_PATH)
            con.execute("DELETE FROM tbl_abwesenheit WHERE id=?", (a["id"],))
            con.commit()
            con.close()
            self.abw = [x for x in self.abw if x["id"] != a["id"]]
            win.destroy()
            self._refresh_after_decision()
            self.status.set("Abwesenheit gelöscht.")
        tk.Button(bar, text="Schließen", command=win.destroy, padx=14, pady=7).pack(side="right", padx=(8, 0))
        tk.Button(bar, text="Löschen", command=loeschen, bg="#C0392B", fg="white", relief="flat", padx=16, pady=7).pack(side="right")

        # Genehmigen / Ablehnen, je nach Status
        if status in ST_OFFEN:
            tk.Button(bar, text="✗ Ablehnen", command=lambda: self._entscheide(a, False, win),
                      bg="#EDF1F6", fg="#C0392B", relief="flat", font=(FONT, 10, "bold"),
                      padx=14, pady=7, cursor="hand2").pack(side="left")
            ok_lbl = "✓ GF-Freigabe" if status == ST_GF else "✓ Genehmigen"
            tk.Button(bar, text=ok_lbl, command=lambda: self._entscheide(a, True, win),
                      bg="#11823B", fg="white", relief="flat", font=(FONT, 10, "bold"),
                      padx=14, pady=7, cursor="hand2").pack(side="left", padx=(8, 0))

    def _refresh_after_decision(self):
        """Kalender + Antrags-Ansicht + Badge nach einer Änderung neu aufbauen."""
        if getattr(self, "view", None) == "antr":
            self.show_antraege()
        elif getattr(self, "view", None) == "abw":
            self.draw_calendar()
        self._update_antraege_badge()

    def _entscheide(self, a, genehmigen, win=None):
        """Stufe-1- oder GF-Entscheidung über einen Antrag treffen."""
        status = a.get("status") or ST_GENEHMIGT
        if status not in ST_OFFEN:
            return
        zweite_stufe = (status == ST_GF)
        if genehmigen:
            kand = self._geschaeftsfuehrung() if zweite_stufe else self._personalverantwortliche()
            titel = "GF-Freigabe" if zweite_stufe else "Genehmigen"
            person = self._pick_decider(titel, kand, win or self.root)
            if person is None:
                return
            now = datetime.now().strftime("%d.%m.%Y %H:%M")
            con = sqlite3.connect(DB_PATH)
            if zweite_stufe:
                con.execute("UPDATE tbl_abwesenheit SET status=?, gf_entscheider=?, gf_am=? WHERE id=?",
                            (ST_GENEHMIGT, person, now, a["id"]))
                a["status"], a["gf_entscheider"], a["gf_am"] = ST_GENEHMIGT, person, now
                msg = f"GF-Freigabe erteilt: {self._name(a['mitarbeiter_id'])} ({a['art']})."
            elif self._gf_freigabe_aktiv():
                con.execute("UPDATE tbl_abwesenheit SET status=?, entscheider=?, entschieden_am=? WHERE id=?",
                            (ST_GF, person, now, a["id"]))
                a["status"], a["entscheider"], a["entschieden_am"] = ST_GF, person, now
                msg = f"Genehmigt von {person} – wartet jetzt auf GF-Freigabe."
            else:
                con.execute("UPDATE tbl_abwesenheit SET status=?, entscheider=?, entschieden_am=? WHERE id=?",
                            (ST_GENEHMIGT, person, now, a["id"]))
                a["status"], a["entscheider"], a["entschieden_am"] = ST_GENEHMIGT, person, now
                msg = f"Genehmigt: {self._name(a['mitarbeiter_id'])} ({a['art']})."
            con.commit()
            con.close()
        else:
            kand = self._geschaeftsfuehrung() if zweite_stufe else self._personalverantwortliche()
            person = self._pick_decider("Ablehnen", kand, win or self.root)
            if person is None:
                return
            grund = simpledialog.askstring("Ablehnen", "Grund der Ablehnung (optional):",
                                           parent=win or self.root) or ""
            now = datetime.now().strftime("%d.%m.%Y %H:%M")
            con = sqlite3.connect(DB_PATH)
            con.execute("UPDATE tbl_abwesenheit SET status=?, entscheider=?, entschieden_am=?, "
                        "ablehnung_grund=? WHERE id=?",
                        (ST_ABGELEHNT, person, now, grund, a["id"]))
            con.commit()
            con.close()
            a["status"], a["entscheider"], a["entschieden_am"], a["ablehnung_grund"] = \
                ST_ABGELEHNT, person, now, grund
            msg = f"Abgelehnt: {self._name(a['mitarbeiter_id'])} ({a['art']})."
        if win is not None:
            win.destroy()
        self._refresh_after_decision()
        self.status.set(msg)

    def _pick_decider(self, titel, kandidaten, parent):
        """Kleiner Dialog: Wer trifft die Entscheidung? Gibt Namen oder None (Abbruch)."""
        kandidaten = [k for k in kandidaten if k] or ["—"]
        dlg = tk.Toplevel(parent)
        dlg.title(titel)
        dlg.configure(bg=BG)
        dlg.transient(parent)
        dlg.grab_set()
        dlg.resizable(False, False)
        tk.Label(dlg, text="Entscheidung durch:", bg=BG, fg=PRIMARY,
                 font=(FONT, 11, "bold")).pack(anchor="w", padx=18, pady=(16, 4))
        cb = ttk.Combobox(dlg, state="readonly", values=kandidaten, width=32)
        # Vorbelegung: angemeldeter Benutzer, falls in der Liste
        pre = next((k for k in kandidaten if self.current_user and self.current_user.lower() in k.lower()), kandidaten[0])
        cb.set(pre)
        cb.pack(padx=18, pady=(0, 6))
        if not self._personalverantwortliche():
            tk.Label(dlg, text="Hinweis: Noch keine Personalverantwortlichen markiert\n"
                              "(Organigramm-Karte bearbeiten).", bg=BG, fg=MUTED,
                     font=(FONT, 8), justify="left").pack(padx=18, anchor="w")
        res = {"v": None}

        def ok():
            res["v"] = cb.get()
            dlg.destroy()

        bar = tk.Frame(dlg, bg=BG)
        bar.pack(fill="x", padx=14, pady=14)
        tk.Button(bar, text="Abbrechen", command=dlg.destroy, padx=12, pady=6).pack(side="right", padx=(8, 0))
        tk.Button(bar, text="OK", command=ok, bg=PRIMARY, fg="white", relief="flat",
                  padx=18, pady=6).pack(side="right")
        dlg.wait_window()
        return res["v"]

    def _abw_form(self):
        win = tk.Toplevel(self.root)
        win.title("Abwesenheit eintragen")
        win.configure(bg=BG)
        win.transient(self.root)
        win.grab_set()
        win.geometry("470x380")
        win.columnconfigure(1, weight=1)
        names = {f"{e['vorname']} {e['name']}".strip(): e["id"] for e in self.emps}
        first = date(self.cal_year, self.cal_month, 1).isoformat()

        def lab(text, r):
            tk.Label(win, text=text, bg=BG, fg=PRIMARY, font=(FONT, 10, "bold")).grid(
                row=r, column=0, sticky="w", padx=16, pady=9)

        lab("Mitarbeiter", 0)
        mit = ttk.Combobox(win, state="readonly", values=list(names.keys()))
        if names:
            mit.current(0)
        mit.grid(row=0, column=1, sticky="ew", padx=16, pady=9)

        lab("Art", 1)
        art = ttk.Combobox(win, state="readonly", values=list(ABW_ARTEN))
        art.current(0)
        art.grid(row=1, column=1, sticky="ew", padx=16, pady=9)

        lab("Grund (Sonderurlaub)", 2)
        grund = ttk.Combobox(win, state="disabled", values=list(SONDERURLAUB_GRUENDE))
        grund.grid(row=2, column=1, sticky="ew", padx=16, pady=9)

        def on_art(_=None):
            if art.get() == "Sonderurlaub":
                grund.configure(state="readonly")
                if not grund.get():
                    grund.current(0)
            else:
                grund.set("")
                grund.configure(state="disabled")
        art.bind("<<ComboboxSelected>>", on_art)

        def date_row(text, r, default):
            lab(text, r)
            box = tk.Frame(win, bg=BG)
            box.grid(row=r, column=1, sticky="ew", padx=16, pady=9)
            box.columnconfigure(0, weight=1)
            ent = tk.Entry(box)
            ent.insert(0, default)
            ent.grid(row=0, column=0, sticky="ew")
            tk.Button(box, text="📅", command=lambda: self._pick_date(ent), relief="flat",
                      bg="#EDF1F6", fg=PRIMARY, padx=8, cursor="hand2").grid(row=0, column=1, padx=(6, 0))
            return ent
        von = date_row("Von", 3, first)
        bis = date_row("Bis", 4, first)

        lab("Notiz", 5)
        notiz = tk.Entry(win)
        notiz.grid(row=5, column=1, sticky="ew", padx=16, pady=9)

        def save():
            mname = mit.get()
            if mname not in names:
                messagebox.showinfo("Abwesenheit", "Bitte einen Mitarbeiter wählen.")
                return
            try:
                v = date.fromisoformat(von.get().strip())
                b = date.fromisoformat(bis.get().strip())
            except ValueError:
                messagebox.showinfo("Abwesenheit", "Bitte ein gültiges Datum wählen (Kalender 📅).")
                return
            if b < v:
                messagebox.showinfo("Abwesenheit", "„Bis“ darf nicht vor „Von“ liegen.")
                return
            a_art = art.get()
            if a_art == "Sonderurlaub" and not grund.get():
                messagebox.showinfo("Abwesenheit", "Bitte einen Grund für den Sonderurlaub wählen.")
                return
            unter = grund.get() if a_art == "Sonderurlaub" else ""
            now = datetime.now().strftime("%d.%m.%Y %H:%M")
            con = sqlite3.connect(DB_PATH)
            cur = con.execute(
                "INSERT INTO tbl_abwesenheit(mitarbeiter_id,art,von,bis,notiz,unterart,status,beantragt_am) "
                "VALUES(?,?,?,?,?,?,?,?)",
                (names[mname], a_art, v.isoformat(), b.isoformat(), notiz.get().strip(),
                 unter, ST_BEANTRAGT, now))
            con.commit()
            con.close()
            self.abw.append({"id": cur.lastrowid, "mitarbeiter_id": names[mname], "art": a_art,
                             "von": v.isoformat(), "bis": b.isoformat(), "notiz": notiz.get().strip(),
                             "unterart": unter, "status": ST_BEANTRAGT, "beantragt_am": now})
            win.destroy()
            self._refresh_after_decision()
            self.status.set(f"Antrag eingereicht: {mname} ({a_art}) – wartet auf Genehmigung.")
            # 31.03.-Hinweis, wenn Urlaub ins neue Jahr übertragen wird
            if a_art == "Urlaub" and b.year > v.year:
                messagebox.showwarning(
                    "Resturlaub / gesetzliche Frist",
                    f"Der Urlaub reicht ins Jahr {b.year}.\n\n"
                    f"Gesetzlicher Hinweis (§ 7 BUrlG): Resturlaub aus {v.year} muss "
                    f"grundsätzlich bis zum 31.03.{v.year + 1} genommen werden – sonst verfällt er.")

        bar = tk.Frame(win, bg=BG)
        bar.grid(row=6, column=0, columnspan=2, sticky="ew", padx=16, pady=16)
        tk.Button(bar, text="Abbrechen", command=win.destroy, padx=14, pady=7).pack(side="right", padx=(8, 0))
        tk.Button(bar, text="Speichern", command=save, bg=PRIMARY, fg="white", relief="flat", padx=18, pady=8).pack(side="right")

    # ----- Kalender-Datumsauswahl (eigenständig, ohne Zusatzpaket) -----
    def _pick_date(self, entry):
        try:
            cur = date.fromisoformat(entry.get().strip())
        except Exception:
            cur = date.today()
        top = tk.Toplevel(self.root)
        top.title("Datum wählen")
        top.configure(bg=CARD)
        top.transient(self.root)
        top.grab_set()
        top.resizable(False, False)
        st = {"y": cur.year, "m": cur.month}
        head = tk.Frame(top, bg=CARD)
        head.pack(fill="x", padx=8, pady=(8, 4))
        lbl = tk.Label(head, bg=CARD, fg=PRIMARY, font=(FONT, 11, "bold"), width=16)
        body = tk.Frame(top, bg=CARD)
        body.pack(padx=8, pady=(0, 8))

        def choose(d):
            entry.delete(0, "end")
            entry.insert(0, d.isoformat())
            top.destroy()

        def render():
            for w in body.winfo_children():
                w.destroy()
            lbl.configure(text=f"{MONATE[st['m']]} {st['y']}")
            for i, wd in enumerate(WD):
                tk.Label(body, text=wd, bg=CARD, fg=MUTED, font=(FONT, 8, "bold"), width=4).grid(row=0, column=i, padx=1, pady=2)
            cal = calendar.Calendar(firstweekday=0)
            for r, week in enumerate(cal.monthdayscalendar(st["y"], st["m"]), start=1):
                for c, day in enumerate(week):
                    if day == 0:
                        continue
                    d = date(st["y"], st["m"], day)
                    is_today = (d == date.today())
                    is_sel = (d == cur)
                    bg = PRIMARY if is_sel else (SELECT_BG if is_today else CARD)
                    fg = "white" if is_sel else INK
                    tk.Button(body, text=str(day), width=4, relief="flat", bg=bg, fg=fg,
                              font=(FONT, 9), cursor="hand2", activebackground=SELECT_BG,
                              command=lambda dd=d: choose(dd)).grid(row=r, column=c, padx=1, pady=1)

        def step(delta):
            m, y = st["m"] + delta, st["y"]
            if m < 1:
                m, y = 12, y - 1
            elif m > 12:
                m, y = 1, y + 1
            st["m"], st["y"] = m, y
            render()

        tk.Button(head, text="◀", command=lambda: step(-1), bg=CARD, fg=PRIMARY, relief="flat",
                  font=(FONT, 11, "bold"), cursor="hand2").pack(side="left")
        lbl.pack(side="left", expand=True)
        tk.Button(head, text="▶", command=lambda: step(1), bg=CARD, fg=PRIMARY, relief="flat",
                  font=(FONT, 11, "bold"), cursor="hand2").pack(side="left")
        render()

    # ----- Resturlaub-Übersicht + Verfall-Workflow -----
    def _resturlaub_overview(self):
        y = self.cal_year
        win = tk.Toplevel(self.root)
        win.title(f"Resturlaub {y}")
        win.configure(bg=BG)
        win.transient(self.root)
        win.grab_set()
        win.geometry("680x480")
        tk.Label(win, text=f"Resturlaub {y}", bg=BG, fg=PRIMARY, font=(FONT, 14, "bold")).pack(anchor="w", padx=16, pady=(14, 2))
        pv_now = self._personalverantwortliche()
        pv_txt = ("Personalverantwortlich: " + ", ".join(pv_now)) if pv_now else \
                 "Hinweis: Noch niemand als personalverantwortlich eingeteilt (im Organigramm bei einer Karte setzen)."
        tk.Label(win, text="Mitarbeiter auswählen und Verfall melden – die personalverantwortliche Person "
                           "entscheidet, ob der offene Resturlaub verfällt. Genommener Urlaub bleibt unangetastet.\n" + pv_txt,
                 bg=BG, fg=MUTED, font=(FONT, 9), wraplength=640, justify="left").pack(anchor="w", padx=16, pady=(0, 6))
        cols = ("ma", "anspruch", "gen", "verf", "rest", "hinweis")
        heads = {"ma": "Mitarbeiter", "anspruch": "Anspruch", "gen": "Genommen",
                 "verf": "Verfallen", "rest": "Rest offen", "hinweis": "Hinweis"}
        tree = ttk.Treeview(win, columns=cols, show="headings", height=12)
        for c in cols:
            tree.heading(c, text=heads[c])
            tree.column(c, width=(160 if c == "ma" else (210 if c == "hinweis" else 70)), anchor="w")
        tree.tag_configure("warn", foreground="#C0392B")
        tree.pack(fill="both", expand=True, padx=16, pady=6)
        id_of = {}

        def populate():
            for iid in tree.get_children(""):
                tree.delete(iid)
            id_of.clear()
            for e in self.emps:
                anspruch = e.get("urlaubsanspruch") or STD_URLAUB
                gen = self._year_count(e["id"], "Urlaub")
                verf = self._verfall_count(e["id"], y)
                rest = anspruch - gen - verf
                if rest > 0:
                    hinweis, tag = f"⚠ bis 31.03.{y + 1} nehmen, sonst verfällt", "warn"
                elif rest < 0:
                    hinweis, tag = "über Anspruch hinaus", "warn"
                else:
                    hinweis, tag = ("verfallen" if verf else "vollständig genommen"), ""
                iid = tree.insert("", "end",
                                  values=(f"{e['vorname']} {e['name']}".strip(), anspruch, gen, verf, rest, hinweis),
                                  tags=(tag,))
                id_of[iid] = e["id"]

        def verfall_melden():
            sel = tree.selection()
            if not sel:
                messagebox.showinfo("Resturlaub", "Bitte zuerst einen Mitarbeiter in der Liste auswählen.")
                return
            eid = id_of[sel[0]]
            e = self.emp_by_id[eid]
            rest = self._rest_urlaub(e, y)
            if rest <= 0:
                messagebox.showinfo("Resturlaub", f"{self._name(eid)} hat keinen offenen Resturlaub.")
                return
            pv = self._personalverantwortliche()
            if not pv:
                messagebox.showinfo(
                    "Keine personalverantwortliche Person",
                    "Es ist niemand als personalverantwortlich eingeteilt.\n\n"
                    "Bitte zuerst im Organigramm bei einem Mitarbeiter „Personalverantwortlich“ setzen "
                    "(Karte bearbeiten). Diese Person entscheidet über den Urlaubsverfall.")
                return
            empf = ", ".join(pv)
            ja = messagebox.askyesno(
                "Nachricht an die personalverantwortliche Person",
                f"An: {empf}\n\n"
                f"Der Resturlaub von {self._name(eid)} ({rest} Tage aus {y}) wurde nicht genommen "
                f"und verfällt gesetzlich zum 31.03.{y + 1}.\n\n"
                f"Soll der noch offene Urlaub jetzt verfallen (löschen)?\n\n"
                f"Ja = verfallen lassen     Nein = nichts tun\n"
                f"(Der bereits genommene Urlaub dieses Jahres bleibt erhalten.)")
            if not ja:
                self.status.set("Verfall abgelehnt – nichts geändert.")
                return
            con = sqlite3.connect(DB_PATH)
            con.execute("INSERT INTO tbl_urlaub_verfall(mitarbeiter_id,jahr,tage,datum) VALUES(?,?,?,?)",
                        (eid, y, rest, date.today().isoformat()))
            con.commit()
            con.close()
            self.verfall.append({"mitarbeiter_id": eid, "jahr": y, "tage": rest, "datum": date.today().isoformat()})
            populate()
            self.draw_calendar()
            self.status.set(f"Resturlaub von {self._name(eid)} verfallen: {rest} Tage ({y}).")

        bar = tk.Frame(win, bg=BG)
        bar.pack(fill="x", padx=16, pady=(2, 6))
        tk.Button(bar, text="✉  Verfall melden", command=verfall_melden, bg=PRIMARY, fg="white",
                  relief="flat", font=(FONT, 10, "bold"), padx=14, pady=7, cursor="hand2").pack(side="left")
        tk.Label(win, text="§ 7 BUrlG: Resturlaub verfällt grundsätzlich am 31.03. des Folgejahres. "
                           "Nur Werktage (Mo–Fr) werden gezählt.",
                 bg=BG, fg=MUTED, font=(FONT, 9), wraplength=640, justify="left").pack(anchor="w", padx=16, pady=(2, 6))
        tk.Button(win, text="Schließen", command=win.destroy, padx=16, pady=7).pack(side="right", padx=16, pady=(0, 14))
        populate()

    # =========================================================================
    #  ANSICHT 3 · ARBEITSBEREICHE (was macht eine Person alles)
    # =========================================================================
    def _ab_for(self, eid):
        bids = {z["bereich_id"] for z in self.ab_zuord if z["mitarbeiter_id"] == eid}
        return [b for b in self.bereiche if b["id"] in bids]

    def _ab_chain(self, bid):
        """Zuordnungen eines Bereichs nach Stufe sortiert: Verantwortlich → Vertretungen."""
        rows = [z for z in self.ab_zuord if z["bereich_id"] == bid]
        return sorted(rows, key=lambda z: z.get("stufe", 0))

    def _ab_stufe(self, eid, bid):
        for z in self.ab_zuord:
            if z["mitarbeiter_id"] == eid and z["bereich_id"] == bid:
                return z.get("stufe", 0)
        return 0

    def _short_name(self, eid):
        e = self.emp_by_id.get(eid, {})
        nn = (e.get("name") or "")
        return f"{e.get('vorname','')} {nn[:1]}.".strip() if nn else self._name(eid)

    def _ab_selected_id(self):
        nm = self.ab_emp.get()
        for e in self.emps:
            if f"{e['vorname']} {e['name']}".strip() == nm:
                return e["id"]
        return None

    def show_arbeitsbereiche(self):
        self.view = "ab"
        self._nav_state()
        self._clear_main()

        t = tk.Frame(self.main, bg=CARD, height=52)
        t.pack(side="top", fill="x")
        t.configure(highlightbackground=BORDER, highlightthickness=1)
        # Ansicht: nach Mitarbeiter ODER nach Bereich (zeigt Vertretungsketten)
        self.ab_mode_ma = tk.Button(t, text="nach Mitarbeiter", command=lambda: self._set_ab_mode("ma"),
                                    relief="flat", font=(FONT, 10, "bold"), padx=12, pady=6, cursor="hand2")
        self.ab_mode_be = tk.Button(t, text="nach Bereich", command=lambda: self._set_ab_mode("bereich"),
                                    relief="flat", font=(FONT, 10, "bold"), padx=12, pady=6, cursor="hand2")
        self.ab_mode_ma.pack(side="left", padx=(14, 2), pady=8)
        self.ab_mode_be.pack(side="left", padx=(0, 6), pady=8)

        tk.Label(t, text="Mitarbeiter:", bg=CARD, fg=MUTED, font=(FONT, 10, "bold")).pack(side="left", padx=(10, 6), pady=10)
        names = [f"{e['vorname']} {e['name']}".strip() for e in self.emps]
        self.ab_emp = ttk.Combobox(t, state="readonly", width=24, values=names)
        if self.emps:
            best = max(self.emps, key=lambda e: len(self._ab_for(e["id"])))
            self.ab_emp.set(f"{best['vorname']} {best['name']}".strip())
        self.ab_emp.pack(side="left", pady=10)
        self.ab_emp.bind("<<ComboboxSelected>>", lambda e: self._render_ab())
        tk.Button(t, text="➕  zuweisen", command=self._ab_assign, bg=PRIMARY, fg="white",
                  relief="flat", font=(FONT, 10, "bold"), padx=14, pady=7, cursor="hand2").pack(side="left", padx=12, pady=8)
        self._set_ab_mode(self.ab_mode, render=False)

        wrap = tk.Frame(self.main, bg=BG)
        wrap.pack(fill="both", expand=True)
        self.ab_canvas = tk.Canvas(wrap, bg=BG, highlightthickness=0)
        vsb = ttk.Scrollbar(wrap, orient="vertical", command=self.ab_canvas.yview)
        self.ab_canvas.configure(yscrollcommand=vsb.set)
        self.ab_canvas.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        wrap.rowconfigure(0, weight=1)
        wrap.columnconfigure(0, weight=1)
        self.ab_inner = tk.Frame(self.ab_canvas, bg=BG)
        self._ab_win = self.ab_canvas.create_window((0, 0), window=self.ab_inner, anchor="nw")
        self.ab_inner.bind("<Configure>", lambda e: self.ab_canvas.configure(scrollregion=self.ab_canvas.bbox("all")))
        self.ab_canvas.bind("<Configure>", lambda e: self.ab_canvas.itemconfig(self._ab_win, width=e.width))
        self._render_ab()

    def _set_ab_mode(self, mode, render=True):
        self.ab_mode = mode
        self.ab_mode_ma.configure(bg=(PRIMARY if mode == "ma" else "#EDF1F6"),
                                  fg=("white" if mode == "ma" else PRIMARY))
        self.ab_mode_be.configure(bg=(PRIMARY if mode == "bereich" else "#EDF1F6"),
                                  fg=("white" if mode == "bereich" else PRIMARY))
        if hasattr(self, "ab_emp"):
            self.ab_emp.configure(state=("readonly" if mode == "ma" else "disabled"))
        if render:
            self._render_ab()

    def _render_ab(self):
        for w in self.ab_inner.winfo_children():
            w.destroy()
        if self.ab_mode == "bereich":
            self._render_ab_bereich()
        else:
            self._render_ab_ma()

    def _chain_text(self, bid):
        return "   →   ".join(f"{STUFE_SHORT[z.get('stufe', 0)]} {self._short_name(z['mitarbeiter_id'])}"
                              for z in self._ab_chain(bid))

    # --- Ansicht nach Mitarbeiter: was macht diese Person, mit Rolle + Kette ---
    def _render_ab_ma(self):
        eid = self._ab_selected_id()
        if not eid:
            return
        e = self.emp_by_id[eid]
        areas = self._ab_for(eid)
        col_e = self.dept_color.get(e["abteilung"] or "—", PRIMARY)

        head = tk.Frame(self.ab_inner, bg=BG)
        head.pack(fill="x", padx=20, pady=(16, 8))
        ini = (e["vorname"][:1] + e["name"][:1]).upper()
        tk.Label(head, text=ini, bg=col_e, fg="white", font=(FONT, 14, "bold"), width=3).pack(side="left", padx=(0, 10), ipady=4)
        box = tk.Frame(head, bg=BG)
        box.pack(side="left")
        tk.Label(box, text=f"{e['vorname']} {e['name']}".strip(), bg=BG, fg=INK, font=(FONT, 15, "bold")).pack(anchor="w")
        tk.Label(box, text=" · ".join(filter(None, [e["position"] or "", e["abteilung"] or ""])),
                 bg=BG, fg=MUTED, font=(FONT, 10)).pack(anchor="w")
        tk.Label(head, text=f"{len(areas)} Arbeitsbereiche", bg=SELECT_BG, fg=PRIMARY,
                 font=(FONT, 10, "bold"), padx=12, pady=4).pack(side="right")

        if not areas:
            tk.Label(self.ab_inner, text="Diesem Mitarbeiter sind noch keine Arbeitsbereiche zugewiesen.\n"
                     "Mit „➕ zuweisen“ hinzufügen.", bg=BG, fg=MUTED, font=(FONT, 11), justify="left").pack(anchor="w", padx=24, pady=24)
            self.status.set(f"{e['vorname']} {e['name']}: keine Arbeitsbereiche.")
            return

        cats = {}
        for b in areas:
            cats.setdefault(b["kategorie"] or "Allgemein", []).append(b)
        cat_color = {k: DEPT_PALETTE[i % len(DEPT_PALETTE)] for i, k in enumerate(sorted(cats))}
        COLS = 2
        for kat in sorted(cats):
            color = cat_color[kat]
            sec = tk.Frame(self.ab_inner, bg=BG)
            sec.pack(fill="x", padx=20, pady=(10, 2))
            tk.Label(sec, text=f"{kat.upper()}", bg=BG, fg=color, font=(FONT, 9, "bold")).pack(side="left")
            tk.Label(sec, text=f"  ({len(cats[kat])})", bg=BG, fg=FAINT, font=(FONT, 9)).pack(side="left")
            grid = tk.Frame(self.ab_inner, bg=BG)
            grid.pack(fill="x", padx=20, pady=(2, 4))
            for i in range(COLS):
                grid.columnconfigure(i, weight=1, uniform="abcol")
            for idx, b in enumerate(sorted(cats[kat], key=lambda x: x["name"])):
                r, c = divmod(idx, COLS)
                chip = tk.Frame(grid, bg=CARD, highlightbackground=BORDER, highlightthickness=1)
                chip.grid(row=r, column=c, sticky="ew", padx=5, pady=5)
                tk.Frame(chip, bg=color, width=5).pack(side="left", fill="y")
                body = tk.Frame(chip, bg=CARD)
                body.pack(side="left", fill="x", expand=True, padx=8, pady=8)
                top = tk.Frame(body, bg=CARD)
                top.pack(fill="x")
                tk.Label(top, text=b["name"], bg=CARD, fg=INK, font=(FONT, 10, "bold"),
                         anchor="w", wraplength=210).pack(side="left", fill="x", expand=True)
                tk.Button(top, text="×", command=lambda bid=b["id"], i=eid: self._ab_remove(i, bid),
                          bg=CARD, fg=MUTED, relief="flat", font=(FONT, 12, "bold"), cursor="hand2",
                          activeforeground="#C0392B", activebackground=CARD).pack(side="right")
                my = self._ab_stufe(eid, b["id"])
                tk.Label(top, text=f" {STUFEN[my]} ", bg=STUFE_COLOR[my], fg="white",
                         font=(FONT, 8, "bold")).pack(side="right", padx=4)
                tk.Label(body, text="Kette:  " + self._chain_text(b["id"]), bg=CARD, fg=MUTED,
                         font=(FONT, 8), anchor="w", justify="left", wraplength=300).pack(anchor="w", pady=(4, 0))
        self.status.set(f"{e['vorname']} {e['name']}: {len(areas)} Arbeitsbereiche · Badge = Rolle, Kette = Vertretungsreihenfolge.")

    # --- Ansicht nach Bereich: Verantwortlich + Vertretungskette als Leiter ---
    def _render_ab_bereich(self):
        tk.Label(self.ab_inner, text="Arbeitsbereiche – Verantwortliche und Vertretungsketten",
                 bg=BG, fg=INK, font=(FONT, 14, "bold")).pack(anchor="w", padx=20, pady=(16, 4))
        cats = {}
        for b in self.bereiche:
            cats.setdefault(b["kategorie"] or "Allgemein", []).append(b)
        cat_color = {k: DEPT_PALETTE[i % len(DEPT_PALETTE)] for i, k in enumerate(sorted(cats))}
        for kat in sorted(cats):
            color = cat_color[kat]
            tk.Label(self.ab_inner, text=kat.upper(), bg=BG, fg=color, font=(FONT, 9, "bold")).pack(anchor="w", padx=22, pady=(12, 2))
            for b in sorted(cats[kat], key=lambda x: x["name"]):
                card = tk.Frame(self.ab_inner, bg=CARD, highlightbackground=BORDER, highlightthickness=1)
                card.pack(fill="x", padx=22, pady=5)
                bar = tk.Frame(card, bg=CARD)
                bar.pack(fill="x", padx=12, pady=(8, 4))
                tk.Frame(bar, bg=color, width=5, height=18).pack(side="left", padx=(0, 8))
                tk.Label(bar, text=b["name"], bg=CARD, fg=INK, font=(FONT, 11, "bold")).pack(side="left")
                tk.Button(bar, text="➕ Person", command=lambda bid=b["id"]: self._ab_add_to_bereich(bid),
                          bg="#EDF1F6", fg=PRIMARY, relief="flat", font=(FONT, 9, "bold"),
                          padx=8, pady=3, cursor="hand2").pack(side="right")
                chain = self._ab_chain(b["id"])
                if not chain:
                    tk.Label(card, text="— noch niemand zugeordnet —", bg=CARD, fg=FAINT,
                             font=(FONT, 9), anchor="w").pack(anchor="w", padx=26, pady=(0, 8))
                    continue
                for z in chain:
                    st = z.get("stufe", 0)
                    row = tk.Frame(card, bg=CARD)
                    row.pack(fill="x", padx=26, pady=1)
                    tk.Label(row, text=("● " if st == 0 else "↳ ") + STUFEN[st], bg=CARD, fg=STUFE_COLOR[st],
                             font=(FONT, 9, "bold"), width=18, anchor="w").pack(side="left")
                    tk.Label(row, text=self._name(z["mitarbeiter_id"]), bg=CARD, fg=INK, font=(FONT, 10)).pack(side="left")
                    tk.Button(row, text="×", command=lambda mid=z["mitarbeiter_id"], bid=b["id"]: self._ab_remove(mid, bid),
                              bg=CARD, fg=MUTED, relief="flat", font=(FONT, 10, "bold"), cursor="hand2",
                              activeforeground="#C0392B", activebackground=CARD).pack(side="right")
                tk.Frame(card, bg=CARD, height=4).pack()
        self.status.set(f"{len(self.bereiche)} Arbeitsbereiche · Kette: ● Verantwortlich, ↳ Vertretungen.")

    def _ab_add_to_bereich(self, bid):
        belegt = {z["mitarbeiter_id"] for z in self._ab_chain(bid)}
        frei = {f"{e['vorname']} {e['name']}".strip(): e["id"] for e in self.emps if e["id"] not in belegt}
        if not frei:
            messagebox.showinfo("Arbeitsbereich", "Alle Mitarbeiter sind diesem Bereich bereits zugeordnet.")
            return
        win = tk.Toplevel(self.root)
        win.title("Person zu Arbeitsbereich")
        win.configure(bg=BG)
        win.transient(self.root)
        win.grab_set()
        win.geometry("420x230")
        win.columnconfigure(1, weight=1)
        bn = next((b["name"] for b in self.bereiche if b["id"] == bid), "")
        tk.Label(win, text=bn, bg=BG, fg=PRIMARY, font=(FONT, 12, "bold")).grid(row=0, column=0, columnspan=2, sticky="w", padx=16, pady=(14, 8))
        tk.Label(win, text="Mitarbeiter:", bg=BG, fg=INK, font=(FONT, 10, "bold")).grid(row=1, column=0, sticky="w", padx=16, pady=9)
        mc = ttk.Combobox(win, state="readonly", values=list(frei.keys()))
        mc.current(0)
        mc.grid(row=1, column=1, sticky="ew", padx=16, pady=9)
        tk.Label(win, text="Rolle:", bg=BG, fg=INK, font=(FONT, 10, "bold")).grid(row=2, column=0, sticky="w", padx=16, pady=9)
        rc = ttk.Combobox(win, state="readonly", values=STUFEN)
        used = {z.get("stufe", 0) for z in self._ab_chain(bid)}
        rc.current(next((i for i in range(4) if i not in used), 0))
        rc.grid(row=2, column=1, sticky="ew", padx=16, pady=9)

        def save():
            mid = frei[mc.get()]
            stufe = STUFEN.index(rc.get())
            con = sqlite3.connect(DB_PATH)
            con.execute("INSERT OR REPLACE INTO tbl_mitarbeiter_arbeitsbereich(mitarbeiter_id,bereich_id,stufe) VALUES(?,?,?)",
                        (mid, bid, stufe))
            con.commit()
            con.close()
            win.destroy()
            self.load()
            self._render_ab()
        bar = tk.Frame(win, bg=BG)
        bar.grid(row=3, column=0, columnspan=2, sticky="ew", padx=16, pady=16)
        tk.Button(bar, text="Abbrechen", command=win.destroy, padx=14, pady=7).pack(side="right", padx=(8, 0))
        tk.Button(bar, text="Speichern", command=save, bg=PRIMARY, fg="white", relief="flat", padx=18, pady=8).pack(side="right")

    def _ab_remove(self, eid, bid):
        con = sqlite3.connect(DB_PATH)
        con.execute("DELETE FROM tbl_mitarbeiter_arbeitsbereich WHERE mitarbeiter_id=? AND bereich_id=?", (eid, bid))
        con.commit()
        con.close()
        self.ab_zuord = [z for z in self.ab_zuord if not (z["mitarbeiter_id"] == eid and z["bereich_id"] == bid)]
        self._render_ab()
        self.status.set("Arbeitsbereich entfernt.")

    def _ab_assign(self):
        eid = self._ab_selected_id()
        if not eid:
            messagebox.showinfo("Arbeitsbereiche", "Bitte zuerst einen Mitarbeiter wählen.")
            return
        win = tk.Toplevel(self.root)
        win.title("Arbeitsbereich zuweisen")
        win.configure(bg=BG)
        win.transient(self.root)
        win.grab_set()
        win.geometry("470x390")
        win.columnconfigure(1, weight=1)
        tk.Label(win, text=f"Für: {self._name(eid)}", bg=BG, fg=PRIMARY, font=(FONT, 12, "bold")).grid(
            row=0, column=0, columnspan=2, sticky="w", padx=16, pady=(14, 8))

        assigned = {b["id"] for b in self._ab_for(eid)}
        avail = {f"{b['name']}  ·  {b['kategorie']}": b["id"] for b in self.bereiche if b["id"] not in assigned}
        tk.Label(win, text="Vorhandenen wählen:", bg=BG, fg=INK, font=(FONT, 10, "bold")).grid(row=1, column=0, sticky="w", padx=16, pady=8)
        ex = ttk.Combobox(win, state="readonly", values=list(avail.keys()))
        ex.grid(row=1, column=1, sticky="ew", padx=16, pady=8)

        tk.Label(win, text="Rolle:", bg=BG, fg=INK, font=(FONT, 10, "bold")).grid(row=2, column=0, sticky="w", padx=16, pady=8)
        rc = ttk.Combobox(win, state="readonly", values=STUFEN)
        rc.current(0)
        rc.grid(row=2, column=1, sticky="ew", padx=16, pady=8)

        tk.Label(win, text="— oder neuen Bereich anlegen —", bg=BG, fg=FAINT, font=(FONT, 9)).grid(row=3, column=0, columnspan=2, pady=(6, 2))
        tk.Label(win, text="Name:", bg=BG, fg=INK, font=(FONT, 10, "bold")).grid(row=4, column=0, sticky="w", padx=16, pady=8)
        nm = tk.Entry(win)
        nm.grid(row=4, column=1, sticky="ew", padx=16, pady=8)
        tk.Label(win, text="Kategorie:", bg=BG, fg=INK, font=(FONT, 10, "bold")).grid(row=5, column=0, sticky="w", padx=16, pady=8)
        kat = ttk.Combobox(win, values=self.kategorien)
        kat.grid(row=5, column=1, sticky="ew", padx=16, pady=8)

        def save():
            stufe = STUFEN.index(rc.get()) if rc.get() in STUFEN else 0
            con = sqlite3.connect(DB_PATH)
            if nm.get().strip():
                cur = con.execute("INSERT INTO tbl_arbeitsbereich(name,kategorie) VALUES(?,?)",
                                  (nm.get().strip(), kat.get().strip() or "Allgemein"))
                bid = cur.lastrowid
            elif ex.get() in avail:
                bid = avail[ex.get()]
            else:
                con.close()
                messagebox.showinfo("Arbeitsbereich", "Bitte einen vorhandenen Bereich wählen oder einen Namen eingeben.")
                return
            con.execute("INSERT OR REPLACE INTO tbl_mitarbeiter_arbeitsbereich(mitarbeiter_id,bereich_id,stufe) VALUES(?,?,?)",
                        (eid, bid, stufe))
            con.commit()
            con.close()
            win.destroy()
            self.load()
            self.show_arbeitsbereiche()
            self.ab_emp.set(self._name(eid))
            self._render_ab()

        bar = tk.Frame(win, bg=BG)
        bar.grid(row=6, column=0, columnspan=2, sticky="ew", padx=16, pady=16)
        tk.Button(bar, text="Abbrechen", command=win.destroy, padx=14, pady=7).pack(side="right", padx=(8, 0))
        tk.Button(bar, text="Zuweisen", command=save, bg=PRIMARY, fg="white", relief="flat", padx=18, pady=8).pack(side="right")

    # =========================================================================
    #  EINSTELLUNGEN (⚙) – Kategorien frei verwalten
    # =========================================================================
    def _settings_dialog(self):
        win = tk.Toplevel(self.root)
        win.title("Einstellungen")
        win.configure(bg=BG)
        win.transient(self.root)
        win.grab_set()
        win.geometry("460x480")
        tk.Label(win, text="⚙  Einstellungen", bg=BG, fg=PRIMARY, font=(FONT, 15, "bold")).pack(anchor="w", padx=18, pady=(16, 2))
        tk.Label(win, text="Kategorien für Arbeitsbereiche", bg=BG, fg=INK, font=(FONT, 11, "bold")).pack(anchor="w", padx=18, pady=(8, 2))
        tk.Label(win, text="Frei anlegen, umbenennen oder löschen. Umbenennen zieht alle zugeordneten "
                           "Arbeitsbereiche automatisch mit.",
                 bg=BG, fg=MUTED, font=(FONT, 9), wraplength=410, justify="left").pack(anchor="w", padx=18, pady=(0, 8))

        listframe = tk.Frame(win, bg=BG)
        listframe.pack(fill="both", expand=True, padx=18)
        lb = tk.Listbox(listframe, font=(FONT, 11), activestyle="none", highlightthickness=1,
                        highlightbackground=BORDER, relief="flat", bg=CARD, fg=INK, selectbackground=SELECT_BG,
                        selectforeground=PRIMARY)
        sb = ttk.Scrollbar(listframe, orient="vertical", command=lb.yview)
        lb.configure(yscrollcommand=sb.set)
        lb.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        def used_count(name):
            return sum(1 for b in self.bereiche if (b["kategorie"] or "") == name)

        def refresh():
            lb.delete(0, "end")
            for k in self.kategorien:
                lb.insert("end", f"{k}    ({used_count(k)} Bereiche)")

        def selected_name():
            sel = lb.curselection()
            if not sel:
                return None
            return self.kategorien[sel[0]]

        def reload_and_refresh():
            self.load()
            refresh()
            if self.view == "ab":
                self._render_ab()

        def add_kat():
            name = simpledialog.askstring("Neue Kategorie", "Name der neuen Kategorie:", parent=win)
            if not name or not name.strip():
                return
            con = sqlite3.connect(DB_PATH)
            con.execute("INSERT OR IGNORE INTO tbl_kategorie(name) VALUES(?)", (name.strip(),))
            con.commit()
            con.close()
            reload_and_refresh()

        def rename_kat():
            old = selected_name()
            if not old:
                messagebox.showinfo("Kategorie", "Bitte zuerst eine Kategorie wählen.")
                return
            new = simpledialog.askstring("Kategorie umbenennen", f"Neuer Name für „{old}“:", initialvalue=old, parent=win)
            if not new or not new.strip() or new.strip() == old:
                return
            new = new.strip()
            con = sqlite3.connect(DB_PATH)
            con.execute("UPDATE tbl_kategorie SET name=? WHERE name=?", (new, old))
            con.execute("UPDATE tbl_arbeitsbereich SET kategorie=? WHERE kategorie=?", (new, old))
            con.commit()
            con.close()
            reload_and_refresh()

        def delete_kat():
            old = selected_name()
            if not old:
                messagebox.showinfo("Kategorie", "Bitte zuerst eine Kategorie wählen.")
                return
            n = used_count(old)
            if n > 0:
                if not messagebox.askyesno("Kategorie löschen",
                        f"„{old}“ wird von {n} Arbeitsbereich(en) genutzt.\n\n"
                        f"Diese auf „Allgemein“ setzen und Kategorie löschen?"):
                    return
            con = sqlite3.connect(DB_PATH)
            con.execute("UPDATE tbl_arbeitsbereich SET kategorie='Allgemein' WHERE kategorie=?", (old,))
            con.execute("INSERT OR IGNORE INTO tbl_kategorie(name) VALUES('Allgemein')")
            con.execute("DELETE FROM tbl_kategorie WHERE name=?", (old,))
            con.commit()
            con.close()
            reload_and_refresh()

        bar = tk.Frame(win, bg=BG)
        bar.pack(fill="x", padx=18, pady=12)
        tk.Button(bar, text="➕ Neu", command=add_kat, bg=PRIMARY, fg="white", relief="flat",
                  font=(FONT, 10, "bold"), padx=14, pady=7, cursor="hand2").pack(side="left", padx=(0, 8))
        tk.Button(bar, text="Umbenennen", command=rename_kat, padx=14, pady=7, cursor="hand2").pack(side="left", padx=4)
        tk.Button(bar, text="Löschen", command=delete_kat, padx=14, pady=7, cursor="hand2").pack(side="left", padx=4)
        tk.Button(bar, text="Schließen", command=win.destroy, padx=14, pady=7).pack(side="right")
        refresh()


def run_standalone():
    init_db()
    root = tk.Tk()
    App(root)
    tour.maybe_show(root, "personal", tour.personal_steps())
    root.mainloop()


def main():
    run_standalone()


if __name__ == "__main__":
    main()
