# -*- coding: utf-8 -*-
"""NMG Buchhaltung - Vorerfassung & Export an die Buchhaltung / das Steuerbuero.

Eigenstaendige App (eigenes Fenster / Taskleisten-Icon, AUMID NMG.Buchhaltung),
teilt sich die NMGone-Datenbank (app/config.py). Sie ist bewusst eine
VORERFASSUNGS- und EXPORT-App, KEINE vollwertige Finanzbuchhaltung:

  * sammelt Ein-/Ausgangsrechnungen (u.a. eingelesene eRechnungen),
  * ordnet sie Konten (SKR03/SKR04) und Steuerschluesseln zu,
  * prueft sie und exportiert geprueftе Buchungen (spaeter: DATEV).

Datenmodell (tbl_buha_*) wird von ensure_buha_tables() idempotent angelegt;
die Standard-Kontenrahmen SKR03/SKR04 werden einmalig geseedet.

run_standalone() wird von start_buchhaltung.py und NMGone.exe --buchhaltung genutzt.
Planung: docs/Plan_Buchhaltung_App.pdf
"""
from __future__ import annotations

import os
import sqlite3
from datetime import datetime, date
from pathlib import Path
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

from .config import DB_PATH, ASSETS_DIR
from . import theme
from . import erechnung as erech

# Wiederverwendbare Bausteine + Palette aus der GDP-App (gemeinsamer Look).
from .gdp_app import (
    _card, FormDialog, ME,
    BG, SHELL_BG, ACCENT, ACCENT_LIGHT, BORDER,
    TEXT, MUTED, OK_GREEN, WARN, DANGER,
    SIDEBAR, SIDEBAR_ACTIVE, SIDEBAR_TEXT, SIDEBAR_MUTED,
)

try:
    from . import tour
except Exception:  # tour ist optional
    tour = None


# ── Standard-Kontenrahmen (Vorschlag, mit Steuerbuero final abstimmen) ────────
# (konto, bezeichnung, art).  art in {Geld, Aktiv, Passiv, USt, Aufwand, Ertrag}
KONTENRAHMEN: dict[str, list[tuple[str, str, str]]] = {
    "SKR03": [
        ("1200", "Bank", "Geld"),
        ("1360", "Geldtransit", "Geld"),
        ("1400", "Forderungen aus Lieferungen & Leistungen", "Aktiv"),
        ("1576", "Abziehbare Vorsteuer 19 %", "USt"),
        ("1571", "Abziehbare Vorsteuer 7 %", "USt"),
        ("1577", "Abziehbare Vorsteuer i.g. Erwerb 19 %", "USt"),
        ("1600", "Verbindlichkeiten aus Lieferungen & Leistungen", "Passiv"),
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
    ],
    "SKR04": [
        ("1800", "Bank", "Geld"),
        ("1460", "Geldtransit", "Geld"),
        ("1200", "Forderungen aus Lieferungen & Leistungen", "Aktiv"),
        ("1406", "Abziehbare Vorsteuer 19 %", "USt"),
        ("1401", "Abziehbare Vorsteuer 7 %", "USt"),
        ("1407", "Abziehbare Vorsteuer i.g. Erwerb 19 %", "USt"),
        ("3300", "Verbindlichkeiten aus Lieferungen & Leistungen", "Passiv"),
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
    ],
}

STATUS_KETTE = ["entwurf", "geprueft", "freigegeben", "exportiert"]
STATUS_TAG = {"entwurf": "warn", "geprueft": "ok", "freigegeben": "ok", "exportiert": "muted"}

# Konto-Arten fuer die Kontenpflege (Anlegen/Bearbeiten).
KONTO_ARTEN = ["Geld", "Aktiv", "Passiv", "USt", "Aufwand", "Ertrag"]

KONTIERUNG_RICHTUNG = ["eingang", "ausgang", "intern"]

# Standard-Kontierungsregeln (Buchungsvorlagen) je Kontenrahmen:
# (name, richtung, konto_soll, konto_haben, steuerschluessel, ust_satz, bemerkung)
# Vorschlag, mit dem Steuerbuero abstimmen (Steuerschluessel besonders).
KONTIERUNG_DEFAULTS: dict[str, list[tuple]] = {
    "SKR03": [
        ("Wareneinkauf 19 %", "eingang", "3200", "1600", "9", 19, "Vorsteuer 19 %"),
        ("Wareneinkauf 7 %", "eingang", "3300", "1600", "8", 7, "Vorsteuer 7 %"),
        ("Innergem. Erwerb 19 %", "eingang", "3425", "1600", "i.g.E.", 19, "Erwerbsteuer + Vorsteuer"),
        ("Verkauf / Erlöse 19 %", "ausgang", "1400", "8400", "3", 19, "Umsatzsteuer 19 %"),
        ("Verkauf / Erlöse 7 %", "ausgang", "1400", "8300", "2", 7, "Umsatzsteuer 7 %"),
        ("Steuerfreie i.g. Lieferung", "ausgang", "1400", "8125", "steuerfrei", 0, "§4 Nr.1b (USt-ID)"),
    ],
    "SKR04": [
        ("Wareneinkauf 19 %", "eingang", "5200", "3300", "9", 19, "Vorsteuer 19 %"),
        ("Wareneinkauf 7 %", "eingang", "5300", "3300", "8", 7, "Vorsteuer 7 %"),
        ("Innergem. Erwerb 19 %", "eingang", "5425", "3300", "i.g.E.", 19, "Erwerbsteuer + Vorsteuer"),
        ("Verkauf / Erlöse 19 %", "ausgang", "1200", "4400", "3", 19, "Umsatzsteuer 19 %"),
        ("Verkauf / Erlöse 7 %", "ausgang", "1200", "4300", "2", 7, "Umsatzsteuer 7 %"),
        ("Steuerfreie i.g. Lieferung", "ausgang", "1200", "4125", "steuerfrei", 0, "§4 Nr.1b (USt-ID)"),
    ],
}


def _norm_mmjjjj(s) -> str | None:
    """Normalisiert eine Monats-/Jahresangabe auf 'MM.JJJJ'.
    Akzeptiert '7 2026', '07.2026', '7/2026', '07-2026'. Ungueltig -> None."""
    roh = str(s or "").strip()
    for sep in (" ", "/", "-"):
        roh = roh.replace(sep, ".")
    teile = [t for t in roh.split(".") if t]
    if len(teile) != 2 or not all(t.isdigit() for t in teile):
        return None
    mm, jjjj = teile
    monat = int(mm)
    if not (1 <= monat <= 12) or len(jjjj) != 4:
        return None
    return f"{monat:02d}.{jjjj}"


def ensure_buha_tables(db_path=DB_PATH):
    """Legt die Buchhaltungs-Tabellen idempotent an und seedet SKR03/04 einmalig."""
    with sqlite3.connect(db_path) as con:
        con.executescript(
            """
            CREATE TABLE IF NOT EXISTS tbl_buha_einstellungen(
                schluessel TEXT PRIMARY KEY, wert TEXT
            );
            CREATE TABLE IF NOT EXISTS tbl_buha_konten(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                kontenrahmen TEXT, konto TEXT, bezeichnung TEXT, art TEXT,
                UNIQUE(kontenrahmen, konto)
            );
            CREATE TABLE IF NOT EXISTS tbl_buha_belege(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                richtung TEXT DEFAULT 'eingang',          -- eingang | ausgang
                beleg_nr TEXT, partner_name TEXT, partner_ustid TEXT,
                datum TEXT, leistungsdatum TEXT,
                netto REAL DEFAULT 0, ust REAL DEFAULT 0, brutto REAL DEFAULT 0,
                waehrung TEXT DEFAULT 'EUR',
                quelle TEXT,                               -- erechnung | faktura | manuell
                status TEXT DEFAULT 'entwurf',
                konto_soll TEXT, konto_haben TEXT, steuerschluessel TEXT,
                xml_pfad TEXT, notiz TEXT,
                erstellt_am TEXT DEFAULT CURRENT_TIMESTAMP, erstellt_von TEXT
            );
            CREATE TABLE IF NOT EXISTS tbl_buha_positionen(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                beleg_id INTEGER NOT NULL,
                pos_nr INTEGER, pzn TEXT, bezeichnung TEXT,
                menge REAL DEFAULT 1, einzelpreis REAL DEFAULT 0,
                ust_satz REAL DEFAULT 19, netto REAL DEFAULT 0,
                FOREIGN KEY(beleg_id) REFERENCES tbl_buha_belege(id) ON DELETE CASCADE
            );
            CREATE TABLE IF NOT EXISTS tbl_buha_kontierung(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                kontenrahmen TEXT, name TEXT, richtung TEXT,
                konto_soll TEXT, konto_haben TEXT, steuerschluessel TEXT,
                ust_satz REAL DEFAULT 0, bemerkung TEXT, aktiv INTEGER DEFAULT 1
            );
            CREATE TABLE IF NOT EXISTS tbl_buha_log(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                zeitpunkt TEXT DEFAULT CURRENT_TIMESTAMP,
                bearbeiter TEXT, aktion TEXT, beleg_id INTEGER, details TEXT
            );
            CREATE INDEX IF NOT EXISTS idx_buha_pos_beleg ON tbl_buha_positionen(beleg_id);
            CREATE INDEX IF NOT EXISTS idx_buha_belege_datum ON tbl_buha_belege(datum);
            """
        )
        # Kontenrahmen einmalig seeden (nur fehlende Konten ergaenzen).
        for rahmen, konten in KONTENRAHMEN.items():
            for konto, bez, art in konten:
                con.execute(
                    "INSERT OR IGNORE INTO tbl_buha_konten(kontenrahmen, konto, bezeichnung, art) "
                    "VALUES(?,?,?,?)", (rahmen, konto, bez, art))
        # Kontierungsregeln je Kontenrahmen nur seeden, wenn noch keine vorhanden
        # (User-Aenderungen werden NICHT ueberschrieben).
        for rahmen, regeln in KONTIERUNG_DEFAULTS.items():
            vorhanden = con.execute("SELECT COUNT(*) FROM tbl_buha_kontierung WHERE kontenrahmen=?",
                                    (rahmen,)).fetchone()[0]
            if not vorhanden:
                for name, richtung, soll, haben, schl, satz, bem in regeln:
                    con.execute(
                        "INSERT INTO tbl_buha_kontierung(kontenrahmen,name,richtung,konto_soll,"
                        "konto_haben,steuerschluessel,ust_satz,bemerkung) VALUES(?,?,?,?,?,?,?,?)",
                        (rahmen, name, richtung, soll, haben, schl, satz, bem))
        # Standard-Einstellungen
        con.execute("INSERT OR IGNORE INTO tbl_buha_einstellungen(schluessel,wert) "
                    "VALUES('kontenrahmen','SKR04')")
        jahr = str(date.today().year)
        con.execute("INSERT OR IGNORE INTO tbl_buha_einstellungen(schluessel,wert) "
                    "VALUES('wj_beginn',?)", (f"01.{jahr}",))
        con.execute("INSERT OR IGNORE INTO tbl_buha_einstellungen(schluessel,wert) "
                    "VALUES('wj_ende',?)", (f"12.{jahr}",))
        con.commit()


def _eur(v) -> str:
    try:
        return f"{float(v):,.2f} €".replace(",", "X").replace(".", ",").replace("X", ".")
    except (TypeError, ValueError):
        return "—"


# ════════════════════════════════════════════════════════════════════════════
class BuchhaltungPanel(tk.Frame):
    """Hauptoberflaeche der Buchhaltungs-App (Sidebar + Seiten)."""

    NAV = (
        ("uebersicht",   "Übersicht",     "\U0001F4CA"),
        ("belege",       "Belege",        "\U0001F4C4"),
        ("kontenrahmen", "Kontenrahmen",  "\U0001F4D2"),
        ("kontierung",   "Kontierung",    "\U0001F501"),
        ("einstellungen", "Einstellungen", "⚙"),
    )
    TITEL = {k: t for k, t, _ in NAV}
    UNTERTITEL = {
        "uebersicht": "Ein-/Ausgangsrechnungen und offene Prüfungen auf einen Blick.",
        "belege": "Belege erfassen, eRechnungen einlesen, prüfen und freigeben.",
        "kontenrahmen": "Standard-Konten SKR03 / SKR04 (Vorschlag, mit Steuerbüro abstimmen).",
        "kontierung": "Buchungsvorlagen: Geschäftsvorfall → Soll-/Haben-Konto (+ Steuerschlüssel).",
        "einstellungen": "Kontenrahmen, Berater-/Mandanten-Nr., Wirtschaftsjahr.",
    }

    def __init__(self, master, db_path=DB_PATH, on_close=None, nmgone_action=None):
        super().__init__(master, bg=SHELL_BG)
        self.db_path = db_path
        self._on_close = on_close or (lambda: self.winfo_toplevel().destroy())
        self._nmgone_action = nmgone_action
        ensure_buha_tables(db_path)
        self._build()

    # ---------------------------------------------------------------- Datenbank
    def _conn(self):
        con = sqlite3.connect(self.db_path, timeout=30)
        con.execute("PRAGMA busy_timeout=30000")
        con.execute("PRAGMA foreign_keys=ON")
        return con

    def _log(self, aktion, beleg_id=None, details=""):
        try:
            with self._conn() as con:
                con.execute("INSERT INTO tbl_buha_log(bearbeiter,aktion,beleg_id,details) "
                            "VALUES(?,?,?,?)", (ME(), aktion, beleg_id, details))
                con.commit()
        except Exception:
            pass

    def _get_setting(self, key, default=""):
        with self._conn() as con:
            r = con.execute("SELECT wert FROM tbl_buha_einstellungen WHERE schluessel=?",
                            (key,)).fetchone()
        return r[0] if r and r[0] is not None else default

    def _set_setting(self, key, val):
        with self._conn() as con:
            con.execute("INSERT INTO tbl_buha_einstellungen(schluessel,wert) VALUES(?,?) "
                        "ON CONFLICT(schluessel) DO UPDATE SET wert=excluded.wert", (key, str(val)))
            con.commit()

    # ------------------------------------------------------------------- Aufbau
    def _build(self):
        self.columnconfigure(1, weight=1)
        self.rowconfigure(0, weight=1)

        self.sidebar = theme.Sidebar(self, width=256, title="Buchhaltung",
                                     subtitle="Vorerfassung & Export")
        self.sidebar.grid(row=0, column=0, sticky="ns")
        self._app_icon = theme.load_icon(ASSETS_DIR / "Buchhaltung.ico", 56)
        if self._app_icon:
            self.sidebar.set_logo(self._app_icon)
        for key, text, icon in self.NAV:
            self.sidebar.add_item(key, icon, text, lambda k=key: self._show_view(k))

        foot = self.sidebar.footer()
        tk.Button(foot, text="\U0001F3E0  NMGone öffnen", command=self._open_nmgone,
                  bg=SIDEBAR_ACTIVE, fg="#FFFFFF", relief="flat",
                  font=(theme.FONT, 10, "bold"), activebackground="#1B5085",
                  activeforeground="#FFFFFF", padx=10, pady=7, cursor="hand2").pack(fill="x", padx=10, pady=(0, 6))
        tk.Button(foot, text="Schließen", command=self._on_close, relief="flat",
                  bg="#0E3454", fg=SIDEBAR_TEXT, activebackground="#15466E",
                  activeforeground="#FFFFFF", padx=10, pady=5, cursor="hand2").pack(fill="x", padx=10)
        self.sidebar.add_footer_note(f"Datenbank:\n{Path(self.db_path).name}")

        main = tk.Frame(self, bg=SHELL_BG)
        main.grid(row=0, column=1, sticky="nsew")
        main.columnconfigure(0, weight=1)
        main.rowconfigure(1, weight=1)

        head = tk.Frame(main, bg=SHELL_BG)
        head.grid(row=0, column=0, sticky="ew", padx=22, pady=(18, 0))
        self._view_title = tk.Label(head, text="", font=(theme.FONT, 18, "bold"),
                                    fg=ACCENT, bg=SHELL_BG)
        self._view_title.pack(anchor="w")
        self._view_subtitle = tk.Label(head, text="", font=(theme.FONT, 10),
                                       fg=MUTED, bg=SHELL_BG)
        self._view_subtitle.pack(anchor="w", pady=(1, 0))
        tk.Frame(head, bg=ACCENT_LIGHT, height=3).pack(fill="x", pady=(10, 0))

        page = tk.Frame(main, bg=SHELL_BG)
        page.grid(row=1, column=0, sticky="nsew", padx=22, pady=12)
        page.columnconfigure(0, weight=1)
        page.rowconfigure(0, weight=1)

        self._views = {}
        self._builders = {
            "uebersicht": self._build_uebersicht,
            "belege": self._build_belege,
            "kontenrahmen": self._build_kontenrahmen,
            "kontierung": self._build_kontierung,
            "einstellungen": self._build_einstellungen,
        }
        self._refreshers = {}
        for key, builder in self._builders.items():
            frame = tk.Frame(page, bg=SHELL_BG)
            frame.grid(row=0, column=0, sticky="nsew")
            builder(frame)
            self._views[key] = frame

        self._current = None
        self._show_view("uebersicht")

    def _show_view(self, key):
        self._current = key
        self._views[key].tkraise()
        self._view_title.config(text=self.TITEL[key])
        self._view_subtitle.config(text=self.UNTERTITEL.get(key, ""))
        self.sidebar.set_active(key)
        ref = self._refreshers.get(key)
        if ref:
            ref()

    def _open_nmgone(self):
        if callable(self._nmgone_action):
            self._nmgone_action()
            return
        try:
            import subprocess, sys
            flags = getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0)
            if getattr(sys, "frozen", False):
                subprocess.Popen([sys.executable], close_fds=True, creationflags=flags)
            else:
                start_py = Path(__file__).resolve().parent.parent / "start.py"
                subprocess.Popen([sys.executable, str(start_py)], close_fds=True, creationflags=flags)
        except Exception as exc:
            messagebox.showerror("NMGone", f"NMGone konnte nicht gestartet werden:\n{exc}", parent=self)

    # =============================================================== ÜBERSICHT
    def _build_uebersicht(self, parent):
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(1, weight=1)
        self._kpi_row = tk.Frame(parent, bg=SHELL_BG)
        self._kpi_row.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        body = tk.Frame(parent, bg=SHELL_BG)
        body.grid(row=1, column=0, sticky="nsew")
        body.columnconfigure(0, weight=1)
        body.columnconfigure(1, weight=1)
        body.rowconfigure(0, weight=1)
        outer_a, self._hinweis_body = _card(body, "Nächste Schritte", "Was als nächstes ansteht")
        outer_a.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        outer_b, b2 = _card(body, "Schnellzugriff")
        outer_b.grid(row=0, column=1, sticky="nsew", padx=(8, 0))
        for key, label, icon in self.NAV[1:]:
            theme.PillButton(b2, f"{icon}  {label}", lambda k=key: self._show_view(k),
                             kind="ghost", font_size=10, padx=12, pady=8).pack(fill="x", pady=3)
        self._refreshers["uebersicht"] = self._refresh_uebersicht

    def _kpi_tile(self, parent, value, label, color):
        f = tk.Frame(parent, bg=BG, highlightbackground=BORDER, highlightthickness=1)
        tk.Label(f, text=value, bg=BG, fg=color, font=(theme.FONT, 19, "bold")).pack(anchor="w", padx=12, pady=(10, 0))
        tk.Label(f, text=label, bg=BG, fg=MUTED, font=(theme.FONT, 9)).pack(anchor="w", padx=12, pady=(0, 10))
        return f

    def _refresh_uebersicht(self):
        for w in self._kpi_row.winfo_children():
            w.destroy()
        with self._conn() as con:
            anz = con.execute("SELECT COUNT(*) FROM tbl_buha_belege").fetchone()[0]
            offen = con.execute("SELECT COUNT(*) FROM tbl_buha_belege WHERE status='entwurf'").fetchone()[0]
            vorsteuer = con.execute("SELECT COALESCE(SUM(ust),0) FROM tbl_buha_belege WHERE richtung='eingang'").fetchone()[0]
            ust = con.execute("SELECT COALESCE(SUM(ust),0) FROM tbl_buha_belege WHERE richtung='ausgang'").fetchone()[0]
        kpis = [(str(anz), "Belege gesamt", ACCENT), (str(offen), "noch im Entwurf", WARN if offen else OK_GREEN),
                (_eur(vorsteuer), "Vorsteuer (Eingang)", ACCENT), (_eur(ust), "Umsatzsteuer (Ausgang)", ACCENT)]
        for i, (val, lbl, col) in enumerate(kpis):
            self._kpi_row.columnconfigure(i, weight=1)
            self._kpi_tile(self._kpi_row, val, lbl, col).grid(row=0, column=i, sticky="ew", padx=(0 if i == 0 else 8, 0))
        for w in self._hinweis_body.winfo_children():
            w.destroy()
        rahmen = self._get_setting("kontenrahmen", "SKR04")
        berater = self._get_setting("berater_nr", "")
        zeilen = [
            ("✓" if berater else "○", f"Kontenrahmen: {rahmen}"),
            ("✓" if berater else "○", "Berater-/Mandanten-Nr. hinterlegt"
                if berater else "Berater-/Mandanten-Nr. noch offen (Einstellungen)"),
            ("○", f"{offen} Beleg(e) im Entwurf – prüfen & freigeben (Belege)"),
            ("○", "DATEV-Export: in Planung (Plan_Buchhaltung_App.pdf)"),
        ]
        for mark, txt in zeilen:
            tk.Label(self._hinweis_body, text=f"  {mark}  {txt}", bg=BG, fg=TEXT,
                     anchor="w", font=(theme.FONT, 10)).pack(fill="x", pady=2)

    # ================================================================== BELEGE
    def _build_belege(self, parent):
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(2, weight=1)
        bar = tk.Frame(parent, bg=SHELL_BG)
        bar.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        theme.PillButton(bar, "🧾  eRechnung einlesen", self._beleg_erechnung_import,
                         kind="primary", font_size=10).pack(side="left")
        theme.PillButton(bar, "↻  Aktualisieren", self._refresh_belege,
                         kind="neutral", font_size=10).pack(side="left", padx=8)
        theme.PillButton(bar, "✓  Status weiter", self._beleg_status_weiter,
                         kind="success", font_size=10).pack(side="left", padx=(0, 8))
        theme.PillButton(bar, "🗑  Entwurf löschen", self._beleg_loeschen,
                         kind="danger", font_size=10).pack(side="left")
        self._belege_count = tk.Label(bar, text="", bg=SHELL_BG, fg=MUTED, font=(theme.FONT, 9))
        self._belege_count.pack(side="left", padx=12)

        outer, body = _card(parent, "Ein- und Ausgangsbelege")
        outer.grid(row=2, column=0, sticky="nsew")
        cols = ("richtung", "nr", "partner", "datum", "netto", "ust", "brutto", "status", "quelle")
        tree = ttk.Treeview(body, columns=cols, show="headings", height=16)
        for c, t, w, anc in (("richtung", "Art", 70, "w"), ("nr", "Beleg-Nr.", 130, "w"),
                             ("partner", "Partner", 200, "w"), ("datum", "Datum", 90, "w"),
                             ("netto", "Netto", 90, "e"), ("ust", "USt", 80, "e"),
                             ("brutto", "Brutto", 95, "e"), ("status", "Status", 95, "w"),
                             ("quelle", "Quelle", 110, "w")):
            tree.heading(c, text=t)
            tree.column(c, width=w, anchor=anc)
        tree.pack(fill="both", expand=True)
        theme.style_treeview(tree)
        self._belege_tree = tree
        self._refreshers["belege"] = self._refresh_belege

    def _refresh_belege(self):
        tree = self._belege_tree
        tree.delete(*tree.get_children())
        with self._conn() as con:
            rows = con.execute(
                "SELECT id, richtung, beleg_nr, partner_name, datum, netto, ust, brutto, status, quelle "
                "FROM tbl_buha_belege ORDER BY id DESC").fetchall()
        for r in rows:
            tree.insert("", "end", iid=str(r[0]), values=(
                "Eingang" if r[1] == "eingang" else "Ausgang", r[2] or "—", r[3] or "",
                r[4] or "", _eur(r[5]), _eur(r[6]), _eur(r[7]),
                (r[8] or "").capitalize(), r[9] or ""))
        self._belege_count.config(text=f"{len(rows)} Beleg(e)")

    def _beleg_auswahl_id(self):
        sel = self._belege_tree.selection()
        return int(sel[0]) if sel else None

    def _beleg_erechnung_import(self):
        """Liest eine eRechnung (ZUGFeRD/Factur-X oder XRechnung) ein und legt sie
        als Beleg (Entwurf) an. Richtung wird abgefragt (Eingang/Ausgang)."""
        pfad = filedialog.askopenfilename(
            title="eRechnung einlesen (XML oder ZUGFeRD-PDF)",
            filetypes=[("eRechnung", "*.xml *.pdf"), ("XML", "*.xml"),
                       ("PDF (ZUGFeRD)", "*.pdf"), ("Alle Dateien", "*.*")])
        if not pfad:
            return
        try:
            daten = erech.lies_erechnung(pfad)
        except Exception as exc:
            messagebox.showerror("eRechnung", f"Konnte nicht gelesen werden:\n{exc}", parent=self)
            return
        richtung = "eingang" if messagebox.askyesno(
            "Richtung", "Eingangsrechnung (vom Lieferanten)?\n\n"
            "Ja = Eingang (Vorsteuer) · Nein = Ausgang (Umsatzsteuer)", parent=self) else "ausgang"
        partner = daten.get("verkaeufer", {}) if richtung == "eingang" else daten.get("kaeufer", {})
        s = daten.get("summen", {})
        jetzt = datetime.now().isoformat(timespec="seconds")
        with self._conn() as con:
            cur = con.execute(
                "INSERT INTO tbl_buha_belege(richtung, beleg_nr, partner_name, partner_ustid, datum, "
                "leistungsdatum, netto, ust, brutto, waehrung, quelle, status, xml_pfad, erstellt_am, erstellt_von) "
                "VALUES(?,?,?,?,?,?,?,?,?,?,?,'entwurf',?,?,?)",
                (richtung, daten.get("rechnungsnr", ""), partner.get("name", ""),
                 partner.get("ustid", ""), daten.get("datum", ""), daten.get("leistungsdatum"),
                 s.get("netto", 0), s.get("ust", 0), s.get("brutto", 0),
                 daten.get("waehrung", "EUR"), "erechnung", pfad, jetzt, ME()))
            bid = cur.lastrowid
            for p in daten.get("positionen", []):
                con.execute(
                    "INSERT INTO tbl_buha_positionen(beleg_id,pos_nr,pzn,bezeichnung,menge,einzelpreis,ust_satz,netto) "
                    "VALUES(?,?,?,?,?,?,?,?)",
                    (bid, p.get("nr"), p.get("pzn", ""), p.get("bezeichnung", ""),
                     p.get("menge", 0), p.get("einzelpreis", 0), p.get("ust_satz", 0), p.get("netto", 0)))
            con.commit()
        self._log("erechnung_import", bid, f"{richtung} · {daten.get('rechnungsnr', '')} · {pfad}")
        fehler = erech.pruefe_en16931(daten)
        self._refresh_belege()
        hinweis = ("\n\n⚠ Pflichtfeld-Hinweise:\n" + "\n".join("• " + x for x in fehler)) if fehler else ""
        messagebox.showinfo("Beleg angelegt",
                            f"{'Eingangs' if richtung == 'eingang' else 'Ausgangs'}rechnung "
                            f"{daten.get('rechnungsnr', '')} als Entwurf übernommen "
                            f"({_eur(s.get('brutto', 0))}).{hinweis}", parent=self)

    def _beleg_status_weiter(self):
        bid = self._beleg_auswahl_id()
        if not bid:
            messagebox.showinfo("Status", "Bitte zuerst einen Beleg auswählen.", parent=self)
            return
        with self._conn() as con:
            r = con.execute("SELECT status FROM tbl_buha_belege WHERE id=?", (bid,)).fetchone()
            if not r:
                return
            idx = STATUS_KETTE.index(r[0]) if r[0] in STATUS_KETTE else 0
            if idx >= len(STATUS_KETTE) - 1:
                messagebox.showinfo("Status", "Beleg ist bereits 'exportiert'.", parent=self)
                return
            neu = STATUS_KETTE[idx + 1]
            con.execute("UPDATE tbl_buha_belege SET status=? WHERE id=?", (neu, bid))
            con.commit()
        self._log("status", bid, f"-> {neu}")
        self._refresh_belege()

    def _beleg_loeschen(self):
        bid = self._beleg_auswahl_id()
        if not bid:
            messagebox.showinfo("Löschen", "Bitte zuerst einen Beleg auswählen.", parent=self)
            return
        with self._conn() as con:
            r = con.execute("SELECT status, beleg_nr FROM tbl_buha_belege WHERE id=?", (bid,)).fetchone()
            if not r:
                return
            if r[0] != "entwurf":
                messagebox.showwarning("Löschen", "Nur Belege im Status 'Entwurf' können gelöscht werden "
                                       "(GoBD-Unveränderbarkeit).", parent=self)
                return
            if not messagebox.askyesno("Löschen", f"Beleg {r[1] or '(Entwurf)'} wirklich löschen?", parent=self):
                return
            con.execute("DELETE FROM tbl_buha_belege WHERE id=?", (bid,))
            con.commit()
        self._log("loeschen", bid, r[1] or "")
        self._refresh_belege()

    # ============================================================ KONTENRAHMEN
    def _build_kontenrahmen(self, parent):
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(2, weight=1)
        bar = tk.Frame(parent, bg=SHELL_BG)
        bar.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        tk.Label(bar, text="Kontenrahmen anzeigen:", bg=SHELL_BG, fg=MUTED,
                 font=(theme.FONT, 10)).pack(side="left", padx=(0, 6))
        self._kr_var = tk.StringVar(value=self._get_setting("kontenrahmen", "SKR04"))
        cb = ttk.Combobox(bar, textvariable=self._kr_var, values=list(KONTENRAHMEN),
                          state="readonly", width=10, style="NMG.TCombobox")
        cb.pack(side="left")
        cb.bind("<<ComboboxSelected>>", lambda _e: self._refresh_kontenrahmen())
        theme.PillButton(bar, "➕ Konto", self._konto_add, kind="primary",
                         font_size=10, padx=12, pady=6).pack(side="left", padx=(14, 0))
        theme.PillButton(bar, "✏ Bearbeiten", self._konto_edit, kind="neutral",
                         font_size=10, padx=12, pady=6).pack(side="left", padx=(6, 0))
        theme.PillButton(bar, "🗑 Löschen", self._konto_del, kind="danger",
                         font_size=10, padx=12, pady=6).pack(side="left", padx=(6, 0))
        tk.Label(parent, text="Konten frei editierbar (hinzufügen / bearbeiten / löschen). Vorschlag "
                 "auf Basis der DATEV-Standardkonten – verbindlich ist die Liste des Steuerbüros.",
                 bg=SHELL_BG, fg=MUTED, font=(theme.FONT, 9)).grid(row=1, column=0, sticky="w", pady=(0, 6))

        outer, body = _card(parent, "Konten")
        outer.grid(row=2, column=0, sticky="nsew")
        cols = ("konto", "bezeichnung", "art")
        tree = ttk.Treeview(body, columns=cols, show="headings", height=18)
        for c, t, w, anc in (("konto", "Konto", 90, "w"), ("bezeichnung", "Bezeichnung", 360, "w"),
                             ("art", "Art", 110, "w")):
            tree.heading(c, text=t)
            tree.column(c, width=w, anchor=anc)
        tree.pack(fill="both", expand=True)
        theme.style_treeview(tree)
        tree.bind("<Double-1>", lambda _e: self._konto_edit())
        self._kr_tree = tree
        self._refreshers["kontenrahmen"] = self._refresh_kontenrahmen

    def _refresh_kontenrahmen(self):
        tree = self._kr_tree
        tree.delete(*tree.get_children())
        rahmen = self._kr_var.get()
        with self._conn() as con:
            rows = con.execute(
                "SELECT konto, bezeichnung, art FROM tbl_buha_konten WHERE kontenrahmen=? ORDER BY konto",
                (rahmen,)).fetchall()
        for konto, bez, art in rows:
            tree.insert("", "end", iid=konto, values=(konto, bez, art))

    def _konto_add(self):
        rahmen = self._kr_var.get()

        def save(v):
            konto = str(v.get("konto", "")).strip()
            bez = str(v.get("bezeichnung", "")).strip()
            art = v.get("art", "") or KONTO_ARTEN[0]
            if not konto or not bez:
                raise ValueError("Konto-Nummer und Bezeichnung sind Pflichtfelder.")
            with self._conn() as con:
                if con.execute("SELECT 1 FROM tbl_buha_konten WHERE kontenrahmen=? AND konto=?",
                               (rahmen, konto)).fetchone():
                    raise ValueError(f"Konto {konto} existiert in {rahmen} bereits.")
                con.execute("INSERT INTO tbl_buha_konten(kontenrahmen,konto,bezeichnung,art) "
                            "VALUES(?,?,?,?)", (rahmen, konto, bez, art))
                con.commit()
            self._log("konto_add", None, f"{rahmen} {konto} {bez}")
            self._refresh_kontenrahmen()

        FormDialog(self, f"Konto hinzufügen ({rahmen})", [
            ("konto", "Konto-Nummer", "text", ""),
            ("bezeichnung", "Bezeichnung", "text", ""),
            ("art", "Art", "combo", KONTO_ARTEN, "Aufwand"),
        ], save)

    def _konto_edit(self):
        rahmen = self._kr_var.get()
        sel = self._kr_tree.selection()
        if not sel:
            messagebox.showinfo("Bearbeiten", "Bitte zuerst ein Konto auswählen.", parent=self)
            return
        konto_alt = sel[0]
        with self._conn() as con:
            r = con.execute("SELECT konto, bezeichnung, art FROM tbl_buha_konten "
                            "WHERE kontenrahmen=? AND konto=?", (rahmen, konto_alt)).fetchone()
        if not r:
            return

        def save(v):
            konto = str(v.get("konto", "")).strip()
            bez = str(v.get("bezeichnung", "")).strip()
            art = v.get("art", "") or KONTO_ARTEN[0]
            if not konto or not bez:
                raise ValueError("Konto-Nummer und Bezeichnung sind Pflichtfelder.")
            with self._conn() as con:
                if konto != konto_alt and con.execute(
                        "SELECT 1 FROM tbl_buha_konten WHERE kontenrahmen=? AND konto=?",
                        (rahmen, konto)).fetchone():
                    raise ValueError(f"Konto {konto} existiert in {rahmen} bereits.")
                con.execute("UPDATE tbl_buha_konten SET konto=?, bezeichnung=?, art=? "
                            "WHERE kontenrahmen=? AND konto=?", (konto, bez, art, rahmen, konto_alt))
                con.commit()
            self._log("konto_edit", None, f"{rahmen} {konto_alt} -> {konto} {bez}")
            self._refresh_kontenrahmen()

        FormDialog(self, f"Konto bearbeiten ({rahmen})", [
            ("konto", "Konto-Nummer", "text", r[0]),
            ("bezeichnung", "Bezeichnung", "text", r[1]),
            ("art", "Art", "combo", KONTO_ARTEN, r[2] or "Aufwand"),
        ], save)

    def _konto_del(self):
        rahmen = self._kr_var.get()
        sel = self._kr_tree.selection()
        if not sel:
            messagebox.showinfo("Löschen", "Bitte zuerst ein Konto auswählen.", parent=self)
            return
        konto = sel[0]
        if not messagebox.askyesno("Löschen", f"Konto {konto} aus {rahmen} löschen?", parent=self):
            return
        with self._conn() as con:
            con.execute("DELETE FROM tbl_buha_konten WHERE kontenrahmen=? AND konto=?", (rahmen, konto))
            con.commit()
        self._log("konto_del", None, f"{rahmen} {konto}")
        self._refresh_kontenrahmen()

    # ============================================================== KONTIERUNG
    def _konto_optionen(self, rahmen):
        """Konten des Rahmens als Auswahl-Strings 'KONTO — Bezeichnung'."""
        with self._conn() as con:
            rows = con.execute("SELECT konto, bezeichnung FROM tbl_buha_konten "
                               "WHERE kontenrahmen=? ORDER BY konto", (rahmen,)).fetchall()
        return [f"{k} — {b}" for k, b in rows]

    @staticmethod
    def _konto_num(option):
        return str(option or "").split(" — ")[0].strip()

    def _build_kontierung(self, parent):
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(2, weight=1)
        bar = tk.Frame(parent, bg=SHELL_BG)
        bar.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        tk.Label(bar, text="Kontenrahmen:", bg=SHELL_BG, fg=MUTED,
                 font=(theme.FONT, 10)).pack(side="left", padx=(0, 6))
        self._ko_var = tk.StringVar(value=self._get_setting("kontenrahmen", "SKR04"))
        cb = ttk.Combobox(bar, textvariable=self._ko_var, values=list(KONTENRAHMEN),
                          state="readonly", width=10, style="NMG.TCombobox")
        cb.pack(side="left")
        cb.bind("<<ComboboxSelected>>", lambda _e: self._refresh_kontierung())
        theme.PillButton(bar, "➕ Regel", self._kontierung_add, kind="primary",
                         font_size=10, padx=12, pady=6).pack(side="left", padx=(14, 0))
        theme.PillButton(bar, "✏ Bearbeiten", self._kontierung_edit, kind="neutral",
                         font_size=10, padx=12, pady=6).pack(side="left", padx=(6, 0))
        theme.PillButton(bar, "🗑 Löschen", self._kontierung_del, kind="danger",
                         font_size=10, padx=12, pady=6).pack(side="left", padx=(6, 0))
        tk.Label(parent, text="Buchungsvorlagen je Geschäftsvorfall: welches Konto im Soll, welches "
                 "im Haben (+ Steuerschlüssel). Belege können daraus ihre Buchung ableiten.",
                 bg=SHELL_BG, fg=MUTED, font=(theme.FONT, 9)).grid(row=1, column=0, sticky="w", pady=(0, 6))

        outer, body = _card(parent, "Kontierungsregeln (Soll an Haben)")
        outer.grid(row=2, column=0, sticky="nsew")
        cols = ("name", "richtung", "soll", "haben", "schluessel", "satz")
        tree = ttk.Treeview(body, columns=cols, show="headings", height=16)
        for c, t, w, anc in (("name", "Geschäftsvorfall", 220, "w"), ("richtung", "Richtung", 90, "w"),
                             ("soll", "Soll-Konto", 160, "w"), ("haben", "Haben-Konto", 160, "w"),
                             ("schluessel", "St.-Schl.", 90, "w"), ("satz", "USt %", 70, "e")):
            tree.heading(c, text=t)
            tree.column(c, width=w, anchor=anc)
        tree.pack(fill="both", expand=True)
        theme.style_treeview(tree)
        tree.bind("<Double-1>", lambda _e: self._kontierung_edit())
        self._ko_tree = tree
        self._refreshers["kontierung"] = self._refresh_kontierung

    def _refresh_kontierung(self):
        tree = self._ko_tree
        tree.delete(*tree.get_children())
        rahmen = self._ko_var.get()
        with self._conn() as con:
            namen = {k: b for k, b in con.execute(
                "SELECT konto, bezeichnung FROM tbl_buha_konten WHERE kontenrahmen=?", (rahmen,)).fetchall()}
            rows = con.execute(
                "SELECT id, name, richtung, konto_soll, konto_haben, steuerschluessel, ust_satz "
                "FROM tbl_buha_kontierung WHERE kontenrahmen=? ORDER BY richtung, name", (rahmen,)).fetchall()
        for rid, name, richtung, soll, haben, schl, satz in rows:
            soll_t = f"{soll}  {namen.get(soll, '')[:14]}".strip()
            haben_t = f"{haben}  {namen.get(haben, '')[:14]}".strip()
            tree.insert("", "end", iid=str(rid), values=(
                name, richtung, soll_t, haben_t, schl or "", f"{satz:g}" if satz else "0"))

    def _kontierung_add(self):
        rahmen = self._ko_var.get()
        opt = self._konto_optionen(rahmen)
        if not opt:
            messagebox.showinfo("Kontierung", f"Für {rahmen} sind keine Konten vorhanden – "
                                "bitte zuerst im Kontenrahmen anlegen.", parent=self)
            return

        def save(v):
            name = str(v.get("name", "")).strip()
            soll = self._konto_num(v.get("konto_soll", ""))
            haben = self._konto_num(v.get("konto_haben", ""))
            if not name or not soll or not haben:
                raise ValueError("Geschäftsvorfall, Soll- und Haben-Konto sind Pflichtfelder.")
            with self._conn() as con:
                con.execute("INSERT INTO tbl_buha_kontierung(kontenrahmen,name,richtung,konto_soll,"
                            "konto_haben,steuerschluessel,ust_satz,bemerkung) VALUES(?,?,?,?,?,?,?,?)",
                            (rahmen, name, v.get("richtung", "eingang"), soll, haben,
                             str(v.get("steuerschluessel", "")).strip(), v.get("ust_satz", 0),
                             str(v.get("bemerkung", "")).strip()))
                con.commit()
            self._log("kontierung_add", None, f"{rahmen} {name}: S {soll} / H {haben}")
            self._refresh_kontierung()

        FormDialog(self, f"Kontierungsregel ({rahmen})", [
            ("name", "Geschäftsvorfall", "text", ""),
            ("richtung", "Richtung", "combo", KONTIERUNG_RICHTUNG, "eingang"),
            ("konto_soll", "Soll-Konto", "combo", opt, opt[0]),
            ("konto_haben", "Haben-Konto", "combo", opt, opt[0]),
            ("steuerschluessel", "Steuerschlüssel", "text", ""),
            ("ust_satz", "USt-Satz %", "float", "19"),
            ("bemerkung", "Bemerkung", "text", ""),
        ], save, width=520)

    def _kontierung_edit(self):
        rahmen = self._ko_var.get()
        sel = self._ko_tree.selection()
        if not sel:
            messagebox.showinfo("Bearbeiten", "Bitte zuerst eine Regel auswählen.", parent=self)
            return
        rid = int(sel[0])
        with self._conn() as con:
            r = con.execute("SELECT name, richtung, konto_soll, konto_haben, steuerschluessel, "
                            "ust_satz, bemerkung FROM tbl_buha_kontierung WHERE id=?", (rid,)).fetchone()
        if not r:
            return
        opt = self._konto_optionen(rahmen)
        nummap = {self._konto_num(o): o for o in opt}
        leer = opt[0] if opt else ""

        def save(v):
            name = str(v.get("name", "")).strip()
            soll = self._konto_num(v.get("konto_soll", ""))
            haben = self._konto_num(v.get("konto_haben", ""))
            if not name or not soll or not haben:
                raise ValueError("Geschäftsvorfall, Soll- und Haben-Konto sind Pflichtfelder.")
            with self._conn() as con:
                con.execute("UPDATE tbl_buha_kontierung SET name=?, richtung=?, konto_soll=?, "
                            "konto_haben=?, steuerschluessel=?, ust_satz=?, bemerkung=? WHERE id=?",
                            (name, v.get("richtung", "eingang"), soll, haben,
                             str(v.get("steuerschluessel", "")).strip(), v.get("ust_satz", 0),
                             str(v.get("bemerkung", "")).strip(), rid))
                con.commit()
            self._log("kontierung_edit", rid, f"{name}: S {soll} / H {haben}")
            self._refresh_kontierung()

        FormDialog(self, f"Kontierungsregel bearbeiten ({rahmen})", [
            ("name", "Geschäftsvorfall", "text", r[0]),
            ("richtung", "Richtung", "combo", KONTIERUNG_RICHTUNG, r[1] or "eingang"),
            ("konto_soll", "Soll-Konto", "combo", opt, nummap.get(r[2], leer)),
            ("konto_haben", "Haben-Konto", "combo", opt, nummap.get(r[3], leer)),
            ("steuerschluessel", "Steuerschlüssel", "text", r[4] or ""),
            ("ust_satz", "USt-Satz %", "float", str(r[5] or 0)),
            ("bemerkung", "Bemerkung", "text", r[6] or ""),
        ], save, width=520)

    def _kontierung_del(self):
        sel = self._ko_tree.selection()
        if not sel:
            messagebox.showinfo("Löschen", "Bitte zuerst eine Regel auswählen.", parent=self)
            return
        rid = int(sel[0])
        if not messagebox.askyesno("Löschen", "Diese Kontierungsregel löschen?", parent=self):
            return
        with self._conn() as con:
            con.execute("DELETE FROM tbl_buha_kontierung WHERE id=?", (rid,))
            con.commit()
        self._log("kontierung_del", rid, "")
        self._refresh_kontierung()

    # ============================================================ EINSTELLUNGEN
    def _build_einstellungen(self, parent):
        parent.columnconfigure(0, weight=1)
        outer, body = _card(parent, "Stammdaten für den Export ans Steuerbüro",
                            "Mit dem Steuerbüro abstimmen (siehe Plan, Abschnitt 7).")
        outer.grid(row=0, column=0, sticky="ew")
        self._einst_vars = {}
        felder = [
            ("kontenrahmen", "Kontenrahmen", "combo", list(KONTENRAHMEN)),
            ("berater_nr", "Berater-Nummer (DATEV)", "text"),
            ("mandant_nr", "Mandanten-Nummer (DATEV)", "text"),
            ("wj_beginn", "Wirtschaftsjahr-Beginn (MM.JJJJ)", "text"),
            ("wj_ende", "Wirtschaftsjahr-Ende (MM.JJJJ)", "text"),
        ]
        for key, label, kind, *extra in felder:
            row = tk.Frame(body, bg=BG)
            row.pack(fill="x", pady=5)
            tk.Label(row, text=label, bg=BG, fg=TEXT, width=26, anchor="w",
                     font=(theme.FONT, 10)).pack(side="left")
            val = self._get_setting(key, "SKR04" if key == "kontenrahmen" else "")
            if kind == "combo":
                var = tk.StringVar(value=val or "SKR04")
                ttk.Combobox(row, textvariable=var, values=extra[0], state="readonly",
                             width=20, style="NMG.TCombobox").pack(side="left")
            else:
                var = tk.StringVar(value=val)
                tk.Entry(row, textvariable=var, font=(theme.FONT, 10), relief="solid",
                         bd=1, highlightthickness=0, width=30).pack(side="left")
            self._einst_vars[key] = var
        btn = tk.Frame(body, bg=BG)
        btn.pack(fill="x", pady=(10, 2))
        theme.PillButton(btn, "Speichern", self._einst_speichern, kind="success",
                         font_size=10, padx=16, pady=7).pack(side="left")

    def _einst_speichern(self):
        werte = {k: v.get().strip() for k, v in self._einst_vars.items()}
        for key in ("wj_beginn", "wj_ende"):
            roh = werte.get(key, "")
            if roh:
                norm = _norm_mmjjjj(roh)
                if not norm:
                    messagebox.showwarning("Wirtschaftsjahr",
                                           "Wirtschaftsjahr-Beginn und -Ende müssen im Format "
                                           "MM.JJJJ angegeben werden (z.B. 07.2026 für ein "
                                           "abweichendes Wirtschaftsjahr).", parent=self)
                    return
                werte[key] = norm
        for key, val in werte.items():
            self._set_setting(key, val)
        self._log("einstellungen", None, "Stammdaten gespeichert")
        messagebox.showinfo("Gespeichert", "Einstellungen gespeichert.", parent=self)
        self._refresh_uebersicht()


# ════════════════════════════════════════════════════════════════════════════
def run_standalone():
    """Startet die Buchhaltungs-App als eigenstaendiges Fenster (eigenes Icon).
    Wird von start_buchhaltung.py und von NMGone.exe --buchhaltung genutzt."""
    if os.name == "nt":
        try:
            import ctypes
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("NMG.Buchhaltung")
        except Exception:
            pass
    try:
        from .migrations import run_migrations
        run_migrations()
    except Exception:
        pass
    ensure_buha_tables()
    root = tk.Tk()
    root.title("NMG Buchhaltung")
    root.geometry("1180x720")
    root.minsize(960, 600)
    try:
        root.state("zoomed")
    except tk.TclError:
        try:
            root.attributes("-zoomed", True)
        except tk.TclError:
            root.geometry(f"{root.winfo_screenwidth()}x{root.winfo_screenheight()}+0+0")
    root.configure(bg=SHELL_BG)
    theme.apply_theme(root)
    theme.apply_widget_defaults(root)
    for ico in ("Buchhaltung.ico", "Faktura.ico", "GDP.ico"):
        try:
            root.iconbitmap(str(ASSETS_DIR / ico))
            break
        except Exception:
            continue
    BuchhaltungPanel(root, on_close=root.destroy).pack(fill="both", expand=True)
    root.mainloop()


if __name__ == "__main__":
    run_standalone()
