"""NMG GDP-App - Wareneingang, Chargen-Rueckverfolgung, Kuehlkette & Retouren.

Eigenstaendige App (eigenes Fenster / Taskleisten-Icon), teilt sich die
NMGone-Datenbank (app/config.py). Sie verzahnt den Wareneingang aus der Kasse
(tbl_wareneingang / tbl_lagerbestand) mit einer vollstaendigen GDP- und
Retouren-Abwicklung, wie sie fuer die Grosshandelserlaubnis verlangt wird.

Module (Sidebar):
  - Uebersicht            : Kennzahlen + offene Pflichten (Ampel)
  - Wareneingang          : NMG-Ware annehmen (Charge/Verfall ins Lager) + GDP-Pruefung
  - Chargen-Rueckverfolgung: Kunde <-> Charge, gezielter Rueckruf
  - Retouren/Reklamationen: Workflow Grund -> Gutschrift -> Wiedereinlagern/Vernichten
  - Kundenqualifizierung  : nur lizenzierte Apotheken duerfen beliefert werden
  - Protokoll             : revisionssicheres Aenderungs-Log

run_standalone() wird von start_gdp.py und NMGone.exe --gdp genutzt.
"""
from __future__ import annotations

import csv
import getpass
import os
import sqlite3
from datetime import datetime, date, timedelta
from pathlib import Path
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

from .config import DB_PATH, ASSETS_DIR
from . import theme

try:
    from . import tour
except Exception:  # tour ist optional
    tour = None

# ── Palette (zentrales Theme, gemeinsamer Look mit NMGone/Kasse) ──────────────
BG = theme.CARD
SHELL_BG = theme.BG
ACCENT = theme.PRIMARY
ACCENT_DARK = theme.PRIMARY_DARK
ACCENT_LIGHT = theme.SELECT_BG
BORDER = theme.BORDER
HEAD_BG = "#EEF3F8"
TEXT = theme.INK
MUTED = theme.MUTED
OK_GREEN = theme.SUCCESS
WARN = theme.WARNING
DANGER = theme.DANGER

SIDEBAR = theme.SIDEBAR
SIDEBAR_ACTIVE = theme.SIDEBAR_ACTIVE
SIDEBAR_TEXT = theme.SIDEBAR_TEXT
SIDEBAR_MUTED = theme.SIDEBAR_MUTED

HEUTE = date.today
ME = lambda: getpass.getuser()


# ── kleine Helfer ─────────────────────────────────────────────────────────────
def _table_exists(con, name) -> bool:
    return con.execute(
        "SELECT 1 FROM sqlite_master WHERE type='table' AND name=?", (name,)
    ).fetchone() is not None


def _parse_verfall(text) -> date | None:
    """'MM/YYYY' oder 'YYYY-MM-DD' -> date (Monatsende fuer MM/YYYY)."""
    if not text:
        return None
    s = str(text).strip()
    try:
        if "/" in s:
            mm, yy = s.split("/")
            mm, yy = int(mm), int(yy)
            if mm == 12:
                return date(yy, 12, 31)
            return date(yy, mm + 1, 1) - timedelta(days=1)
        return datetime.strptime(s[:10], "%Y-%m-%d").date()
    except Exception:
        return None


def _normalize_verfall(t):
    """Leer -> ''. Gueltiges Verfalldatum (MM/JJ, MM/JJJJ, auch '.'/' ' als Trenner,
    einstelliger Monat) -> 'MM/JJ' bzw. 'MM/JJJJ'. Auch ohne Trenner:
    '1226' -> '12/26', '122026' -> '12/2026'. Ungueltig (z.B. Monat 13) -> None."""
    t = str(t or "").strip().replace(".", "/").replace("-", "/").replace(" ", "")
    if not t:
        return ""
    if "/" not in t and t.isdigit():
        if len(t) in (3, 4):
            t = f"{t[:-2]}/{t[-2:]}"
        elif len(t) in (5, 6):
            t = f"{t[:-4]}/{t[-4:]}"
        else:
            return None
    teile = [x for x in t.split("/") if x != ""]
    if len(teile) != 2:
        return None
    mon, jahr = teile
    if not (mon.isdigit() and jahr.isdigit()) or len(jahr) not in (2, 4):
        return None
    m = int(mon)
    if not (1 <= m <= 12):
        return None
    return f"{m:02d}/{jahr}"


def _heute_iso() -> str:
    return HEUTE().isoformat()


def ensure_gdp_tables(db_path=DB_PATH):
    """Alle GDP-Tabellen idempotent anlegen. WAL fuer parallelen Zugriff."""
    with sqlite3.connect(db_path) as con:
        try:
            con.execute("PRAGMA journal_mode=WAL")
        except sqlite3.Error:
            pass
        con.executescript(
            """
            CREATE TABLE IF NOT EXISTS tbl_gdp_auslieferung(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                datum TEXT, kundennummer TEXT, kunde_name TEXT,
                pzn TEXT, artikelname TEXT, charge TEXT, verfall TEXT,
                menge INTEGER DEFAULT 0, beleg_nr TEXT, bearbeiter TEXT,
                erstellt_am TEXT DEFAULT CURRENT_TIMESTAMP
            );
            CREATE TABLE IF NOT EXISTS tbl_gdp_we_pruefung(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                we_id INTEGER, datum TEXT, lieferant TEXT,
                transport_temp_c REAL, temp_ok INTEGER DEFAULT 1,
                unversehrt INTEGER DEFAULT 1, dokumente_ok INTEGER DEFAULT 1,
                gdp_konform INTEGER DEFAULT 1, geprueft_von TEXT, bemerkung TEXT,
                erstellt_am TEXT DEFAULT CURRENT_TIMESTAMP
            );
            CREATE TABLE IF NOT EXISTS tbl_gdp_messpunkt(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT, typ TEXT, soll_min REAL, soll_max REAL,
                aktiv INTEGER DEFAULT 1
            );
            CREATE TABLE IF NOT EXISTS tbl_gdp_temperatur(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                messpunkt_id INTEGER, zeitpunkt TEXT, temp_c REAL,
                status TEXT, erfasst_von TEXT, notiz TEXT,
                massnahme TEXT, behoben INTEGER DEFAULT 0
            );
            CREATE TABLE IF NOT EXISTS tbl_gdp_inspektion(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                datum TEXT, titel TEXT, typ TEXT, status TEXT,
                durchgefuehrt_von TEXT, naechste_faellig TEXT, bemerkung TEXT,
                erstellt_am TEXT DEFAULT CURRENT_TIMESTAMP
            );
            CREATE TABLE IF NOT EXISTS tbl_gdp_inspektion_punkt(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                inspektion_id INTEGER, kategorie TEXT, frage TEXT,
                ergebnis TEXT, bemerkung TEXT, massnahme TEXT
            );
            CREATE TABLE IF NOT EXISTS tbl_gdp_kunde_quali(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                kundennummer TEXT UNIQUE, kunde_name TEXT,
                lizenznummer TEXT, lizenz_typ TEXT, lizenz_gueltig_bis TEXT,
                qualifiziert INTEGER DEFAULT 0, geprueft_am TEXT,
                geprueft_von TEXT, bemerkung TEXT
            );
            CREATE TABLE IF NOT EXISTS tbl_gdp_retoure(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                datum TEXT, typ TEXT, kundennummer TEXT, kunde_name TEXT,
                pzn TEXT, artikelname TEXT, charge TEXT, verfall TEXT,
                menge INTEGER DEFAULT 0, grund TEXT, temperaturbruch INTEGER DEFAULT 0,
                status TEXT, entscheidung TEXT, gutschrift_beleg TEXT,
                faktura_beleg_id INTEGER, bearbeiter TEXT,
                erstellt_am TEXT DEFAULT CURRENT_TIMESTAMP, abgeschlossen_am TEXT,
                notiz TEXT
            );
            CREATE TABLE IF NOT EXISTS tbl_gdp_rueckruf(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                datum TEXT, charge TEXT, pzn TEXT, artikelname TEXT,
                grund TEXT, betroffene_kunden INTEGER DEFAULT 0, betroffene_menge INTEGER DEFAULT 0,
                status TEXT, ausgeloest_von TEXT, bemerkung TEXT,
                erstellt_am TEXT DEFAULT CURRENT_TIMESTAMP
            );
            CREATE TABLE IF NOT EXISTS tbl_gdp_log(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                zeitpunkt TEXT DEFAULT CURRENT_TIMESTAMP, bearbeiter TEXT,
                modul TEXT, aktion TEXT, bezug_id INTEGER, details TEXT
            );
            CREATE TABLE IF NOT EXISTS tbl_gdp_abschreibung(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                datum TEXT, pzn TEXT, artikelname TEXT, charge TEXT, verfall TEXT,
                menge INTEGER DEFAULT 0, grund TEXT, wert_ek REAL,
                retoure_id INTEGER, bearbeiter TEXT, erstellt_am TEXT DEFAULT CURRENT_TIMESTAMP
            );
            """
        )
        # Retourenbestand (Quarantaene) lebt als zusaetzliche Spalte je Lagerzeile:
        # Normalbestand = menge (verkaufbar), Retourenbestand = menge_retoure
        # (gesperrt, erst nach Freigabe verkaufbar). Nachruesten fuer Bestands-DBs.
        if _table_exists(con, "tbl_lagerbestand"):
            lcols = {r[1] for r in con.execute("PRAGMA table_info(tbl_lagerbestand)")}
            if "menge_retoure" not in lcols:
                con.execute("ALTER TABLE tbl_lagerbestand ADD COLUMN menge_retoure INTEGER DEFAULT 0")
        # Standard-Messpunkte anlegen, falls noch keiner existiert.
        if con.execute("SELECT COUNT(*) FROM tbl_gdp_messpunkt").fetchone()[0] == 0:
            con.executemany(
                "INSERT INTO tbl_gdp_messpunkt(name,typ,soll_min,soll_max,aktiv) VALUES(?,?,?,?,1)",
                [("Kuehlschrank 1 (2-8 C)", "Kuehlschrank", 2.0, 8.0),
                 ("Kuehlschrank 2 (2-8 C)", "Kuehlschrank", 2.0, 8.0),
                 ("Lager Trockenbereich", "Lager", 15.0, 25.0),
                 ("Transportbox Kuehlkette", "Transport", 2.0, 8.0)])
        con.commit()


# ── wiederverwendbare UI-Bausteine ────────────────────────────────────────────
def _card(parent, title=None, subtitle=None):
    """Weisse Karte mit 1px-Rahmen und optionalem Akzent-Titel. -> (outer, body)."""
    outer = tk.Frame(parent, bg=BORDER)
    inner = tk.Frame(outer, bg=BG)
    inner.pack(fill="both", expand=True, padx=1, pady=1)
    if title:
        head = tk.Frame(inner, bg=BG)
        head.pack(fill="x", padx=14, pady=(10, 0))
        tk.Label(head, text=title, bg=BG, fg=ACCENT,
                 font=(theme.FONT, 11, "bold")).pack(side="left")
        if subtitle:
            tk.Label(head, text=subtitle, bg=BG, fg=MUTED,
                     font=(theme.FONT, 9)).pack(side="left", padx=(8, 0))
        tk.Frame(inner, bg=ACCENT_LIGHT, height=2).pack(fill="x", padx=14, pady=(7, 0))
    body = tk.Frame(inner, bg=BG)
    body.pack(fill="both", expand=True, padx=14, pady=12)
    return outer, body


def _make_tree(parent, columns, widths, anchors=None, height=12):
    """Treeview mit vertikalem Scrollbalken + Zebra-Tags. -> tree."""
    wrap = tk.Frame(parent, bg=BG)
    wrap.pack(fill="both", expand=True)
    tree = ttk.Treeview(wrap, columns=columns, show="headings", height=height,
                        style="NMG.Treeview")
    vs = ttk.Scrollbar(wrap, orient="vertical", command=tree.yview)
    tree.configure(yscrollcommand=vs.set)
    tree.pack(side="left", fill="both", expand=True)
    vs.pack(side="right", fill="y")
    anchors = anchors or {}
    for col, head, w in zip(columns, columns, widths):
        tree.heading(col, text=col)
        tree.column(col, width=w, anchor=anchors.get(col, "w"), stretch=True)
    tree.tag_configure("odd", background="#FFFFFF")
    tree.tag_configure("even", background="#F7FAFD")
    tree.tag_configure("alert", background="#FCEBEA", foreground=DANGER)
    tree.tag_configure("ok", foreground=OK_GREEN)
    tree.tag_configure("warn", background="#FEF6E7", foreground="#8A5A00")
    return tree


def _badge(parent, text, color):
    return tk.Label(parent, text=f" {text} ", bg=color, fg="#FFFFFF",
                    font=(theme.FONT, 9, "bold"), padx=6, pady=1)


class SearchBox(tk.Frame):
    """Eingabezeile mit Live-Vorschlagsliste. fetch(text)->[(label,payload)],
    on_select(payload) bei Auswahl."""

    def __init__(self, master, fetch, on_select, width=46, height=6):
        super().__init__(master, bg=BG)
        self._fetch = fetch
        self._on_select = on_select
        self.var = tk.StringVar()
        self.entry = tk.Entry(self, textvariable=self.var, width=width)
        self.entry.pack(fill="x")
        self.lb = tk.Listbox(self, height=height, activestyle="dotbox")
        self._payloads = []
        self.entry.bind("<KeyRelease>", self._on_key)
        self.lb.bind("<ButtonRelease-1>", self._on_pick)
        self.lb.bind("<Return>", self._on_pick)
        self.entry.bind("<Down>", self._focus_list)

    def _focus_list(self, _e):
        if self._payloads:
            self.lb.focus_set()
            self.lb.selection_clear(0, "end")
            self.lb.selection_set(0)

    def _on_key(self, e):
        if e.keysym in ("Up", "Down", "Return"):
            return
        text = self.var.get().strip()
        if not text:
            self._hide()
            return
        results = list(self._fetch(text))
        self.lb.delete(0, "end")
        self._payloads = []
        for label, payload in results:
            self.lb.insert("end", label)
            self._payloads.append(payload)
        if results:
            self.lb.pack(fill="x", pady=(2, 0))
        else:
            self._hide()

    def _on_pick(self, _e):
        sel = self.lb.curselection()
        if not sel:
            return
        i = sel[0]
        self.var.set(self.lb.get(i))
        payload = self._payloads[i]
        self._hide()
        self._on_select(payload)

    def _hide(self):
        self.lb.pack_forget()

    def clear(self):
        self.var.set("")
        self._hide()


# Status -> (Treeview-Tag) Zuordnung fuer Retouren
RET_STATUS_TAG = {
    "Neu": "warn", "In Pruefung": "warn", "Gutschrift": "ok",
    "Im Retourenbestand": "warn", "Wiedereingelagert": "ok",
    "Vernichtet": "alert", "Abgelehnt": "alert",
}


class FormDialog(tk.Toplevel):
    """Schlankes, themengerechtes Eingabe-Dialogfenster.

    fields: Liste von (key, label, kind, options/default). kind in
    {'text','int','float','combo','check','date','multiline'}. on_ok(values)
    bekommt ein dict key->wert (Strings/Zahlen)."""

    def __init__(self, master, title, fields, on_ok, width=440):
        super().__init__(master)
        self.title(title)
        self.configure(bg=SHELL_BG)
        self.resizable(False, False)
        self.transient(master.winfo_toplevel())
        self._on_ok = on_ok
        self._vars: dict[str, tk.Variable] = {}
        self._kinds: dict[str, str] = {}

        head = tk.Frame(self, bg=ACCENT)
        head.pack(fill="x")
        tk.Label(head, text=title, bg=ACCENT, fg="#FFFFFF",
                 font=(theme.FONT, 13, "bold"), padx=16, pady=10).pack(anchor="w")

        body = tk.Frame(self, bg=SHELL_BG)
        body.pack(fill="both", expand=True, padx=18, pady=14)
        for key, label, kind, *extra in fields:
            self._kinds[key] = kind
            row = tk.Frame(body, bg=SHELL_BG)
            row.pack(fill="x", pady=4)
            tk.Label(row, text=label, bg=SHELL_BG, fg=TEXT, width=18, anchor="w",
                     font=(theme.FONT, 10)).pack(side="left")
            default = extra[1] if len(extra) > 1 else (extra[0] if extra and kind in ("text", "int", "float", "date", "multiline") else "")
            if kind == "combo":
                var = tk.StringVar(value=(extra[1] if len(extra) > 1 else ""))
                cb = ttk.Combobox(row, textvariable=var, values=list(extra[0]),
                                  state="readonly", style="NMG.TCombobox", width=26)
                cb.pack(side="left", fill="x", expand=True)
            elif kind == "check":
                var = tk.IntVar(value=int(extra[0]) if extra else 0)
                tk.Checkbutton(row, variable=var, bg=SHELL_BG, activebackground=SHELL_BG).pack(side="left")
            elif kind == "multiline":
                var = tk.StringVar(value=default)
                txt = tk.Text(row, height=3, width=28, font=(theme.FONT, 10),
                              relief="solid", bd=1, highlightthickness=0)
                txt.insert("1.0", default)
                txt.pack(side="left", fill="x", expand=True)
                self._vars[key] = txt  # Text-Widget direkt
                continue
            else:
                var = tk.StringVar(value=str(default))
                tk.Entry(row, textvariable=var, font=(theme.FONT, 10), relief="solid",
                         bd=1, highlightthickness=0).pack(side="left", fill="x", expand=True)
            self._vars[key] = var

        btns = tk.Frame(self, bg=SHELL_BG)
        btns.pack(fill="x", padx=18, pady=(0, 16))
        theme.PillButton(btns, "Speichern", self._ok, kind="success",
                         font_size=10, padx=16, pady=7).pack(side="right")
        theme.PillButton(btns, "Abbrechen", self.destroy, kind="neutral",
                         font_size=10, padx=16, pady=7).pack(side="right", padx=(0, 8))
        self.bind("<Escape>", lambda _e: self.destroy())
        self.update_idletasks()
        try:
            x = master.winfo_rootx() + 80
            y = master.winfo_rooty() + 60
            self.geometry(f"+{x}+{y}")
        except Exception:
            pass
        self.grab_set()

    def _ok(self):
        out = {}
        for key, var in self._vars.items():
            kind = self._kinds.get(key)
            if isinstance(var, tk.Text):
                out[key] = var.get("1.0", "end").strip()
            elif kind == "int":
                try:
                    out[key] = int(float(var.get() or 0))
                except ValueError:
                    out[key] = 0
            elif kind == "float":
                try:
                    out[key] = float(str(var.get()).replace(",", ".") or 0)
                except ValueError:
                    out[key] = 0.0
            elif kind == "check":
                out[key] = int(var.get())
            else:
                out[key] = var.get().strip()
        try:
            self._on_ok(out)
        except Exception as exc:
            messagebox.showerror("Fehler", str(exc), parent=self)
            return
        self.destroy()


# ════════════════════════════════════════════════════════════════════════════
class GDPPanel(tk.Frame):
    """Hauptoberflaeche der GDP-App (Sidebar + Seiten)."""

    NAV = (
        ("uebersicht",      "Uebersicht",            "\U0001F4CA"),
        ("wareneingang",    "Wareneingang",          "\U0001F4E6"),
        ("rueckverfolgung", "Chargen-Rueckverfolgung", "\U0001F50E"),
        ("retouren",        "Retouren / Reklamation", "↩"),
        ("retourenbestand", "Retourenbestand",       "⚖"),
        ("qualifizierung",  "Kundenqualifizierung",  "\U0001F6E1"),
        ("protokoll",       "Protokoll",             "\U0001F4DD"),
    )
    TITEL = {k: t for k, t, _ in NAV}
    UNTERTITEL = {
        "uebersicht": "Kennzahlen und offene GDP-Pflichten auf einen Blick.",
        "wareneingang": "NMG-Ware annehmen (Charge/Verfall ins Lager) und GDP-Pruefung erfassen.",
        "rueckverfolgung": "Welche Apotheke hat welche Charge erhalten? Gezielter Rueckruf.",
        "retouren": "Rueckwaren und Reklamationen von der Erfassung bis zur Gutschrift.",
        "retourenbestand": "Zurueckgenommene Ware in Quarantaene: freigeben oder abschreiben.",
        "qualifizierung": "Nur lizenzierte Apotheken duerfen beliefert werden.",
        "protokoll": "Revisionssicheres Protokoll aller GDP-Vorgaenge.",
    }

    def __init__(self, master, db_path=DB_PATH, on_close=None, nmgone_action=None):
        super().__init__(master, bg=SHELL_BG)
        self.db_path = db_path
        self._on_close = on_close or (lambda: self.winfo_toplevel().destroy())
        self._nmgone_action = nmgone_action
        ensure_gdp_tables(db_path)
        self._build()

    # ---------------------------------------------------------------- Datenbank
    def _conn(self):
        con = sqlite3.connect(self.db_path, timeout=30)
        con.execute("PRAGMA busy_timeout=30000")
        con.execute("PRAGMA foreign_keys=ON")
        return con

    def _log(self, modul, aktion, bezug_id=None, details=""):
        try:
            with self._conn() as con:
                con.execute(
                    "INSERT INTO tbl_gdp_log(zeitpunkt,bearbeiter,modul,aktion,bezug_id,details) "
                    "VALUES(?,?,?,?,?,?)",
                    (datetime.now().isoformat(timespec="seconds"), ME(),
                     modul, aktion, bezug_id, details))
                con.commit()
        except Exception:
            pass
        if getattr(self, "_log_tree", None) is not None and self._current == "protokoll":
            self._refresh_protokoll()

    # ------------------------------------------------------------------- Aufbau
    def _build(self):
        self.columnconfigure(1, weight=1)
        self.rowconfigure(0, weight=1)

        # ---- Sidebar ----
        left = tk.Frame(self, bg=SIDEBAR, width=256)
        left.grid(row=0, column=0, sticky="ns")
        left.grid_propagate(False)
        left.rowconfigure(2, weight=1)

        logo_box = tk.Frame(left, bg=SIDEBAR)
        logo_box.grid(row=0, column=0, sticky="ew", padx=16, pady=(18, 0))
        self._app_icon = theme.load_icon(ASSETS_DIR / "GDP.ico", 56)
        if self._app_icon:
            tk.Label(logo_box, image=self._app_icon, bg=SIDEBAR).pack(anchor="w")
        else:
            tk.Label(logo_box, text="\U0001F4E6", font=(theme.FONT, 28),
                     bg=SIDEBAR, fg="#FFFFFF").pack(anchor="w")
        tk.Label(left, text="Wareneingang\n& Retouren", font=(theme.FONT, 15, "bold"),
                 fg="#FFFFFF", bg=SIDEBAR, justify="left").grid(
            row=1, column=0, sticky="w", padx=18, pady=(6, 10))

        nav = tk.Frame(left, bg=SIDEBAR)
        nav.grid(row=2, column=0, sticky="new", padx=8)
        self._nav_buttons = {}
        self._nav_bars = {}
        for key, text, icon in self.NAV:
            rowf = tk.Frame(nav, bg=SIDEBAR)
            rowf.pack(fill="x", pady=1)
            bar = tk.Frame(rowf, bg=SIDEBAR, width=4)
            bar.pack(side="left", fill="y")
            b = tk.Button(rowf, text=f"   {icon}   {text}", anchor="w", relief="flat",
                          bg=SIDEBAR, fg=SIDEBAR_TEXT, font=(theme.FONT, 11), bd=0,
                          activebackground=SIDEBAR_ACTIVE, activeforeground="#FFFFFF",
                          cursor="hand2", command=lambda k=key: self._show_view(k))
            b.pack(side="left", fill="x", expand=True, ipady=6)
            b.bind("<Enter>", lambda _e, k=key: self._nav_hover(k, True))
            b.bind("<Leave>", lambda _e, k=key: self._nav_hover(k, False))
            self._nav_buttons[key] = b
            self._nav_bars[key] = bar

        bottom = tk.Frame(left, bg=SIDEBAR)
        bottom.grid(row=3, column=0, sticky="ew", padx=10, pady=(0, 6))
        tk.Button(bottom, text="\U0001F3E0  NMGone oeffnen", command=self._open_nmgone,
                  bg=SIDEBAR_ACTIVE, fg="#FFFFFF", relief="flat",
                  font=(theme.FONT, 10, "bold"), activebackground="#1B5085",
                  activeforeground="#FFFFFF", padx=10, pady=7, cursor="hand2").pack(fill="x", pady=(0, 6))
        tk.Button(bottom, text="Schliessen", command=self._on_close, relief="flat",
                  bg="#0E3454", fg=SIDEBAR_TEXT, activebackground="#15466E",
                  activeforeground="#FFFFFF", padx=10, pady=5, cursor="hand2").pack(fill="x")
        tk.Label(left, text=f"Datenbank:\n{Path(self.db_path).name}", justify="left",
                 bg=SIDEBAR, fg=SIDEBAR_MUTED, font=(theme.FONT, 9)).grid(
            row=4, column=0, sticky="w", padx=16, pady=12)

        # ---- Hauptbereich ----
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
            "wareneingang": self._build_wareneingang,
            "rueckverfolgung": self._build_rueckverfolgung,
            "retouren": self._build_retouren,
            "retourenbestand": self._build_retourenbestand,
            "qualifizierung": self._build_qualifizierung,
            "protokoll": self._build_protokoll,
        }
        self._refreshers = {}
        for key, builder in self._builders.items():
            frame = tk.Frame(page, bg=SHELL_BG)
            frame.grid(row=0, column=0, sticky="nsew")
            builder(frame)
            self._views[key] = frame

        self._current = None
        self._show_view("uebersicht")

    def _nav_hover(self, key, on):
        if key == self._current:
            return
        self._nav_buttons[key].config(bg=SIDEBAR_ACTIVE if on else SIDEBAR)

    def _show_view(self, key):
        self._current = key
        self._views[key].tkraise()
        self._view_title.config(text=self.TITEL[key])
        self._view_subtitle.config(text=self.UNTERTITEL.get(key, ""))
        for k, b in self._nav_buttons.items():
            active = (k == key)
            b.config(bg=SIDEBAR_ACTIVE if active else SIDEBAR,
                     fg="#FFFFFF" if active else SIDEBAR_TEXT,
                     font=(theme.FONT, 11, "bold" if active else "normal"))
            self._nav_bars[k].config(bg=ACCENT if active else SIDEBAR)
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

        outer_a, self._alert_body = _card(body, "Offene Pflichten", "Was als naechstes ansteht")
        outer_a.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        outer_b, b2 = _card(body, "Schnellzugriff")
        outer_b.grid(row=0, column=1, sticky="nsew", padx=(8, 0))
        for key, label, icon in self.NAV[1:]:
            theme.PillButton(b2, f"{icon}  {label}", lambda k=key: self._show_view(k),
                             kind="ghost", font_size=10, padx=12, pady=8).pack(fill="x", pady=3)
        self._refreshers["uebersicht"] = self._refresh_uebersicht

    def _kpi_tile(self, parent, value, label, color):
        f = tk.Frame(parent, bg=BG, highlightbackground=BORDER, highlightthickness=1)
        tk.Frame(f, bg=color, height=4).pack(fill="x")
        tk.Label(f, text=str(value), bg=BG, fg=color,
                 font=(theme.FONT, 26, "bold")).pack(pady=(12, 0), padx=18)
        tk.Label(f, text=label, bg=BG, fg=MUTED, font=(theme.FONT, 9),
                 wraplength=150, justify="center").pack(pady=(0, 12), padx=12)
        return f

    def _refresh_uebersicht(self):
        with self._conn() as con:
            ret_offen = con.execute(
                "SELECT COUNT(*) FROM tbl_gdp_retoure WHERE status IN ('Neu','In Pruefung')").fetchone()[0]
            # Lizenzen
            liz_abgelaufen = con.execute(
                "SELECT COUNT(*) FROM tbl_gdp_kunde_quali WHERE lizenz_gueltig_bis IS NOT NULL "
                "AND lizenz_gueltig_bis < ?", (_heute_iso(),)).fetchone()[0]
            # Chargen die in <90 Tagen ablaufen (aus Lagerbestand)
            chargen = []
            if _table_exists(con, "tbl_lagerbestand"):
                chargen = con.execute(
                    "SELECT pzn, artikelname, charge, verfall, menge FROM tbl_lagerbestand "
                    "WHERE menge>0").fetchall()
            we_ungeprueft = 0
            if _table_exists(con, "tbl_wareneingang"):
                we_ungeprueft = con.execute(
                    "SELECT COUNT(*) FROM tbl_wareneingang w WHERE NOT EXISTS "
                    "(SELECT 1 FROM tbl_gdp_we_pruefung p WHERE p.we_id=w.id)").fetchone()[0]
            # Retourenbestand (Quarantaene), der noch entschieden werden muss
            rb_offen = rb_stueck = 0
            if _table_exists(con, "tbl_lagerbestand"):
                lcols = {r[1] for r in con.execute("PRAGMA table_info(tbl_lagerbestand)")}
                if "menge_retoure" in lcols:
                    rb_offen, rb_stueck = con.execute(
                        "SELECT COUNT(*), COALESCE(SUM(menge_retoure),0) FROM tbl_lagerbestand "
                        "WHERE COALESCE(menge_retoure,0) > 0").fetchone()
        grenze = HEUTE() + timedelta(days=90)
        bald_ab = sum(1 for *_ , verf, menge in chargen
                      if (_parse_verfall(verf) and _parse_verfall(verf) <= grenze))

        for w in self._kpi_row.winfo_children():
            w.destroy()
        tiles = [
            (ret_offen, "Offene Retouren", WARN if ret_offen else OK_GREEN),
            (rb_offen, "Retourenbestand offen", WARN if rb_offen else OK_GREEN),
            (we_ungeprueft, "Wareneingaenge ungeprueft", WARN if we_ungeprueft else OK_GREEN),
            (liz_abgelaufen, "Lizenzen abgelaufen", DANGER if liz_abgelaufen else OK_GREEN),
            (bald_ab, "Chargen Verfall < 90 Tage", WARN if bald_ab else OK_GREEN),
        ]
        for i, (v, lbl, col) in enumerate(tiles):
            self._kpi_row.columnconfigure(i, weight=1)
            self._kpi_tile(self._kpi_row, v, lbl, col).grid(row=0, column=i, sticky="nsew", padx=4)

        for w in self._alert_body.winfo_children():
            w.destroy()
        alerts = []
        if liz_abgelaufen:
            alerts.append((DANGER, f"{liz_abgelaufen} Apotheke(n) mit abgelaufener Lizenz - Belieferung sperren.", "qualifizierung"))
        if we_ungeprueft:
            alerts.append((WARN, f"{we_ungeprueft} Wareneingang/-gaenge ohne GDP-Pruefung.", "wareneingang"))
        if ret_offen:
            alerts.append((WARN, f"{ret_offen} Retoure(n)/Reklamation(en) in Bearbeitung.", "retouren"))
        if rb_offen:
            alerts.append((WARN, f"{rb_offen} Charge(n) im Retourenbestand ({rb_stueck} St) - "
                                 "freigeben oder abschreiben.", "retourenbestand"))
        if bald_ab:
            alerts.append((WARN, f"{bald_ab} Charge(n) laufen in unter 90 Tagen ab.", "rueckverfolgung"))
        if not alerts:
            tk.Label(self._alert_body, text="✔  Alles erledigt - keine offenen GDP-Pflichten.",
                     bg=BG, fg=OK_GREEN, font=(theme.FONT, 11, "bold")).pack(anchor="w", pady=8)
        for col, text, target in alerts:
            row = tk.Frame(self._alert_body, bg=BG)
            row.pack(fill="x", pady=3)
            tk.Frame(row, bg=col, width=5).pack(side="left", fill="y")
            tk.Label(row, text=text, bg=BG, fg=TEXT, font=(theme.FONT, 10),
                     anchor="w", justify="left", wraplength=420).pack(side="left", padx=8)
            theme.PillButton(row, "oeffnen", lambda t=target: self._show_view(t),
                             kind="neutral", font_size=9, padx=10, pady=4).pack(side="right")

    # ============================================================ WARENEINGANG
    def _search_nmg(self, text, limit=25):
        like = f"%{text}%"
        with self._conn() as con:
            if not _table_exists(con, "tbl_nmg_stamm"):
                return []
            return con.execute(
                "SELECT pzn, artikelname FROM tbl_nmg_stamm "
                "WHERE pzn LIKE ? OR artikelname LIKE ? ORDER BY artikelname LIMIT ?",
                (like, like, limit),
            ).fetchall()

    def _build_wareneingang(self, parent):
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(2, weight=1)
        bar = tk.Frame(parent, bg=SHELL_BG)
        bar.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        theme.PillButton(bar, "✅  GDP-Pruefung erfassen", self._we_pruefung_dialog,
                         kind="primary", font_size=10).pack(side="left")
        theme.PillButton(bar, "↻  Aktualisieren", self._refresh_wareneingang,
                         kind="neutral", font_size=10).pack(side="left", padx=8)
        tk.Label(bar, text="Doppelklick auf einen Eingang = Positionen anzeigen.",
                 bg=SHELL_BG, fg=MUTED, font=(theme.FONT, 9)).pack(side="left", padx=12)

        # ---- Buchungsmaske: NMG-Ware mit Charge/Verfall ins Lager buchen ----
        outer_b, formb = _card(parent, "Artikel ins Lager buchen", "NMG-Ware annehmen")
        outer_b.grid(row=1, column=0, sticky="ew", pady=(0, 8))
        whead = tk.Frame(formb, bg=BG)
        whead.pack(fill="x")
        tk.Label(whead, text="Artikel suchen (PZN / Name):", bg=BG,
                 font=(theme.FONT, 9, "bold")).pack(side="left")
        tk.Button(whead, text="📥 Wareneingangs-Liste importieren", command=self._we_import,
                  font=(theme.FONT, 8), padx=6, pady=1).pack(side="right")
        self._we_search = SearchBox(
            formb,
            fetch=lambda t: [(f"{pzn}  ·  {name}", (pzn, name)) for pzn, name in self._search_nmg(t)],
            on_select=self._we_pick_artikel, height=5,
        )
        self._we_search.pack(fill="x", pady=(2, 4))
        self._we_artikel_label = tk.Label(formb, text="Kein Artikel gewaehlt.", bg=BG,
                                          fg=MUTED, font=(theme.FONT, 9, "italic"))
        self._we_artikel_label.pack(anchor="w")
        self._we_pzn = None

        row1 = tk.Frame(formb, bg=BG)
        row1.pack(fill="x", pady=(6, 0))
        self._we_charge_var = tk.StringVar()
        self._we_verfall_var = tk.StringVar()
        self._we_menge_var = tk.StringVar()
        self._we_ek_var = tk.StringVar()
        for label, var, w in (("Charge", self._we_charge_var, 12), ("Verfall", self._we_verfall_var, 10),
                              ("Menge", self._we_menge_var, 6), ("EK €", self._we_ek_var, 8)):
            tk.Label(row1, text=label + ":", bg=BG, font=(theme.FONT, 9, "bold")).pack(side="left", padx=(0, 4))
            tk.Entry(row1, textvariable=var, width=w).pack(side="left", padx=(0, 12))

        row2 = tk.Frame(formb, bg=BG)
        row2.pack(fill="x", pady=(6, 0))
        self._we_lief_var = tk.StringVar(value="NMG")
        self._we_ls_var = tk.StringVar()
        for label, var, w in (("Lieferant", self._we_lief_var, 18), ("Lieferschein", self._we_ls_var, 14)):
            tk.Label(row2, text=label + ":", bg=BG, font=(theme.FONT, 9, "bold")).pack(side="left", padx=(0, 4))
            tk.Entry(row2, textvariable=var, width=w).pack(side="left", padx=(0, 12))
        theme.PillButton(row2, "Einbuchen", self._we_buchen, kind="primary",
                         font_size=10, padx=14, pady=4).pack(side="left")

        # ---- Liste der erfassten Wareneingaenge + GDP-Pruefstatus ----
        outer, body = _card(parent)
        outer.grid(row=2, column=0, sticky="nsew")
        cols = ("Datum", "Lieferant", "Lieferschein", "Positionen", "Stueck", "GDP-Pruefung")
        self._we_tree = _make_tree(body, cols, (90, 150, 130, 80, 70, 150),
                                   anchors={"Positionen": "center", "Stueck": "center"})
        self._we_tree.bind("<Double-1>", lambda _e: self._we_show_positions())
        self._refreshers["wareneingang"] = self._refresh_wareneingang

    def _we_pick_artikel(self, payload):
        pzn, name = payload
        self._we_pzn = pzn
        self._we_artikel_label.config(text=f"Gewaehlt: {pzn} · {name}", fg=ACCENT)
        # EK mit dem APU vorbelegen (Standard); kann ueberschrieben werden.
        with self._conn() as con:
            row = con.execute("SELECT apu FROM tbl_nmg_stamm WHERE pzn=? LIMIT 1", (pzn,)).fetchone()
        apu = row[0] if row else None
        self._we_ek_var.set(f"{apu:.2f}".replace(".", ",") if apu is not None else "")

    def _we_buchen(self):
        pzn = getattr(self, "_we_pzn", None)
        if not pzn:
            messagebox.showwarning("Wareneingang", "Bitte zuerst einen NMG-Artikel waehlen.", parent=self)
            return
        charge = self._we_charge_var.get().strip()
        verfall = _normalize_verfall(self._we_verfall_var.get())
        if verfall is None:
            messagebox.showwarning("Wareneingang", "Verfalldatum ungueltig. Bitte als MM/JJ oder "
                                   "MM/JJJJ eingeben (Monat 1-12), z. B. 01/2027.", parent=self)
            return
        try:
            menge = int(self._we_menge_var.get().strip())
            if menge <= 0:
                raise ValueError
        except ValueError:
            messagebox.showwarning("Wareneingang", "Menge muss eine positive Zahl sein.", parent=self)
            return
        lieferant = self._we_lief_var.get().strip() or "NMG"
        lieferschein = self._we_ls_var.get().strip()
        with self._conn() as con:
            row = con.execute("SELECT artikelname, apu FROM tbl_nmg_stamm WHERE pzn=? LIMIT 1",
                              (pzn,)).fetchone()
            if not row:
                messagebox.showwarning("Wareneingang", f"PZN {pzn} ist kein NMG-Artikel.", parent=self)
                return
            artikelname, apu = row
            ek_text = self._we_ek_var.get().strip().replace(",", ".")
            try:
                ek = float(ek_text) if ek_text else apu
            except ValueError:
                ek = apu
            jetzt = datetime.now().isoformat(timespec="seconds")
            cur = con.execute(
                "INSERT INTO tbl_wareneingang(datum, lieferant, lieferschein, bearbeiter) VALUES(?,?,?,?)",
                (jetzt, lieferant, lieferschein or None, ME()))
            we_id = cur.lastrowid
            con.execute(
                "INSERT INTO tbl_wareneingang_positionen(we_id,pzn,artikelname,charge,verfall,menge,ek) "
                "VALUES(?,?,?,?,?,?,?)", (we_id, pzn, artikelname, charge, verfall, menge, ek))
            upd = con.execute(
                "UPDATE tbl_lagerbestand SET menge = menge + ?, ek=COALESCE(?, ek), aktualisiert_am = ? "
                "WHERE pzn=? AND COALESCE(charge,'')=? AND COALESCE(verfall,'')=?",
                (menge, ek, jetzt, pzn, charge, verfall))
            if upd.rowcount == 0:
                con.execute(
                    "INSERT INTO tbl_lagerbestand(pzn,artikelname,charge,verfall,menge,ek,aktualisiert_am) "
                    "VALUES(?,?,?,?,?,?,?)", (pzn, artikelname, charge, verfall, menge, ek, jetzt))
            con.commit()
        self._log("Wareneingang", "Einbuchen", we_id,
                  f"{artikelname} (PZN {pzn}) +{menge} · Charge {charge or '–'} · Verf {verfall or '–'}")
        self._we_pzn = None
        self._we_search.clear()
        self._we_artikel_label.config(text="Kein Artikel gewaehlt.", fg=MUTED)
        for var in (self._we_charge_var, self._we_verfall_var, self._we_menge_var,
                    self._we_ek_var, self._we_ls_var):
            var.set("")
        self._refresh_wareneingang()
        messagebox.showinfo("Wareneingang", f"{menge} × {artikelname} eingebucht.", parent=self)

    def _we_import(self):
        from . import kasse_import
        path = filedialog.askopenfilename(
            title="Wareneingangs-Liste waehlen (Excel, CSV oder TXT)",
            filetypes=[("Tabellen", "*.xlsx *.xlsm *.csv *.txt"), ("Alle Dateien", "*.*")],
            parent=self)
        if not path:
            return
        try:
            r = kasse_import.import_wareneingang(self.db_path, path)
        except Exception as e:
            messagebox.showerror("Wareneingang-Import", f"Import fehlgeschlagen:\n{e}", parent=self)
            return
        self._refresh_wareneingang()
        self._log("Wareneingang", "Import", None,
                  f"{r.get('quelle','')}: {r.get('neu_chargen',0)} neue Chargen, "
                  f"{r.get('erhoehte_chargen',0)} erhoeht")
        messagebox.showinfo(
            "Wareneingang-Import",
            f"Quelle: {r.get('quelle','')} · {r.get('gelesen',0)} Zeilen gelesen.\n\n"
            f"Neue Chargen: {r.get('neu_chargen',0)} · Bestand erhoeht: {r.get('erhoehte_chargen',0)}\n"
            f"Kein NMG-Artikel: {r.get('kein_nmg',0)} · Uebersprungen: {r.get('uebersprungen',0)}",
            parent=self)

    def _refresh_wareneingang(self):
        t = self._we_tree
        t.delete(*t.get_children())
        with self._conn() as con:
            if not _table_exists(con, "tbl_wareneingang"):
                t.insert("", "end", values=("-", "Keine Kasse-Wareneingaenge vorhanden", "", "", "", ""))
                return
            rows = con.execute(
                """SELECT w.id, w.datum, w.lieferant, w.lieferschein,
                          (SELECT COUNT(*) FROM tbl_wareneingang_positionen p WHERE p.we_id=w.id),
                          (SELECT COALESCE(SUM(menge),0) FROM tbl_wareneingang_positionen p WHERE p.we_id=w.id),
                          (SELECT gdp_konform FROM tbl_gdp_we_pruefung g WHERE g.we_id=w.id ORDER BY g.id DESC LIMIT 1)
                   FROM tbl_wareneingang w ORDER BY w.datum DESC, w.id DESC""").fetchall()
        for i, (wid, datum, lief, ls, npos, stk, konform) in enumerate(rows):
            if konform is None:
                pruef, tag = "offen", "warn"
            elif konform:
                pruef, tag = "✔ GDP-konform", "ok"
            else:
                pruef, tag = "✖ nicht konform", "alert"
            base = "even" if i % 2 else "odd"
            t.insert("", "end", iid=str(wid),
                     values=(datum or "", lief or "", ls or "", npos, stk, pruef),
                     tags=(base if tag in ("ok",) else tag,))

    def _we_selected_id(self):
        sel = self._we_tree.selection()
        if not sel:
            return None
        try:
            return int(sel[0])
        except ValueError:
            return None

    def _we_show_positions(self):
        wid = self._we_selected_id()
        if not wid:
            return
        with self._conn() as con:
            head = con.execute("SELECT datum, lieferant, lieferschein FROM tbl_wareneingang WHERE id=?",
                               (wid,)).fetchone()
            pos = con.execute(
                "SELECT pzn, artikelname, charge, verfall, menge, ek FROM tbl_wareneingang_positionen "
                "WHERE we_id=? ORDER BY artikelname", (wid,)).fetchall()
        win = tk.Toplevel(self)
        win.title(f"Wareneingang #{wid}")
        win.configure(bg=SHELL_BG)
        win.geometry("680x420")
        tk.Label(win, bg=ACCENT, fg="#FFFFFF", font=(theme.FONT, 12, "bold"),
                 text=f"  {head[1] or ''}  ·  {head[0] or ''}  ·  LS {head[2] or ''}",
                 anchor="w").pack(fill="x", ipady=8)
        wrap = tk.Frame(win, bg=SHELL_BG)
        wrap.pack(fill="both", expand=True, padx=12, pady=12)
        tree = _make_tree(wrap, ("PZN", "Artikel", "Charge", "Verfall", "Menge", "EK"),
                          (80, 240, 90, 80, 60, 70),
                          anchors={"Menge": "center", "EK": "e"})
        for i, (pzn, name, charge, verf, menge, ek) in enumerate(pos):
            tree.insert("", "end", values=(pzn or "", name or "", charge or "", verf or "",
                                           menge, f"{ek:.2f}" if ek else ""),
                        tags=("even" if i % 2 else "odd",))

    def _we_pruefung_dialog(self):
        wid = self._we_selected_id()
        if not wid:
            messagebox.showinfo("GDP-Pruefung", "Bitte zuerst einen Wareneingang in der Liste waehlen.", parent=self)
            return
        with self._conn() as con:
            head = con.execute("SELECT datum, lieferant FROM tbl_wareneingang WHERE id=?", (wid,)).fetchone()
        lief = head[1] if head else ""

        def save(v):
            konform = 1 if (v["temp_ok"] and v["unversehrt"] and v["dokumente_ok"]) else 0
            with self._conn() as con:
                con.execute(
                    """INSERT INTO tbl_gdp_we_pruefung
                       (we_id,datum,lieferant,transport_temp_c,temp_ok,unversehrt,
                        dokumente_ok,gdp_konform,geprueft_von,bemerkung)
                       VALUES(?,?,?,?,?,?,?,?,?,?)""",
                    (wid, _heute_iso(), lief, v["transport_temp_c"], v["temp_ok"],
                     v["unversehrt"], v["dokumente_ok"], konform, ME(), v["bemerkung"]))
                con.commit()
            self._log("Wareneingang", "GDP-Pruefung erfasst", wid,
                      f"{lief}: {'konform' if konform else 'NICHT konform'}")
            self._refresh_wareneingang()

        FormDialog(self, f"GDP-Wareneingangspruefung #{wid}", [
            ("transport_temp_c", "Transport-Temp (C)", "float", "6.0"),
            ("temp_ok", "Temperatur ok", "check", 1),
            ("unversehrt", "Ware unversehrt", "check", 1),
            ("dokumente_ok", "Dokumente vollstaendig", "check", 1),
            ("bemerkung", "Bemerkung", "multiline", ""),
        ], save)

    # =================================================== CHARGEN-RÜCKVERFOLGUNG
    def _outbound_rows(self, con, term=None, charge=None, kundennummer=None):
        """Auslieferungen Kunde<->Charge aus ZWEI Quellen vereint:
          1) tbl_gdp_auslieferung  - explizite/importierte GDP-Auslieferungen
          2) tbl_bestellpositionen + tbl_bestellungen - echte Kasse-Verkaeufe
             (bestellart='Bestellung'); so funktioniert die Rueckverfolgung im
             Echtbetrieb ohne Doppelpflege.
        term   = LIKE ueber Charge ODER PZN; charge = exakte Charge;
        kundennummer = exakter Kunde. Liefert Tupel
        (datum,kundennummer,kunde_name,pzn,artikelname,charge,verfall,menge,beleg_nr,quelle)."""
        rows = []
        cond, params = [], []
        if term:
            cond.append("(charge LIKE ? OR pzn LIKE ?)")
            params += [f"%{term}%", f"%{term}%"]
        if charge is not None:
            cond.append("COALESCE(charge,'')=?")
            params.append(charge or "")
        if kundennummer:
            cond.append("kundennummer=?")
            params.append(kundennummer)
        wsql = (" WHERE " + " AND ".join(cond)) if cond else ""
        for r in con.execute(
                "SELECT datum,kundennummer,kunde_name,pzn,artikelname,charge,verfall,menge,"
                f"COALESCE(beleg_nr,'') FROM tbl_gdp_auslieferung{wsql} ORDER BY datum DESC", params):
            rows.append((*r, "GDP"))
        if _table_exists(con, "tbl_bestellpositionen") and _table_exists(con, "tbl_bestellungen"):
            cond2 = ["COALESCE(p.bestellart,'Bestellung')='Bestellung'"]
            params2 = []
            if term:
                cond2.append("(p.charge LIKE ? OR p.pzn LIKE ?)")
                params2 += [f"%{term}%", f"%{term}%"]
            if charge is not None:
                cond2.append("COALESCE(p.charge,'')=?")
                params2.append(charge or "")
            if kundennummer:
                cond2.append("b.kundennummer=?")
                params2.append(kundennummer)
            q = ("SELECT b.datum, b.kundennummer, COALESCE(b.apotheke,''), p.pzn, p.artikelname, "
                 "p.charge, p.verfall, p.menge, ('Best-'||b.id) "
                 "FROM tbl_bestellpositionen p JOIN tbl_bestellungen b ON b.id=p.bestell_id "
                 "WHERE " + " AND ".join(cond2) + " ORDER BY b.datum DESC")
            for r in con.execute(q, params2):
                rows.append((*r, "Kasse"))
        return rows

    def _build_rueckverfolgung(self, parent):
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(2, weight=1)
        bar = tk.Frame(parent, bg=SHELL_BG)
        bar.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        tk.Label(bar, text="Charge oder PZN:", bg=SHELL_BG, fg=TEXT,
                 font=(theme.FONT, 10)).pack(side="left")
        self._rv_var = tk.StringVar()
        e = tk.Entry(bar, textvariable=self._rv_var, font=(theme.FONT, 11), width=22,
                     relief="solid", bd=1)
        e.pack(side="left", padx=8)
        e.bind("<Return>", lambda _e: self._rv_search())
        theme.PillButton(bar, "\U0001F50E  Suchen", self._rv_search, kind="primary",
                         font_size=10).pack(side="left")
        theme.PillButton(bar, "\U0001F4E2  Rueckruf ausloesen", self._rv_rueckruf,
                         kind="danger", font_size=10).pack(side="left", padx=8)
        theme.PillButton(bar, "⬇  Verteiler exportieren", self._rv_export,
                         kind="neutral", font_size=10).pack(side="left")
        theme.PillButton(bar, "\U0001F4C4  Chargendossier", self._rv_dossier,
                         kind="neutral", font_size=10).pack(side="left", padx=8)

        info = tk.Frame(parent, bg=SHELL_BG)
        info.grid(row=1, column=0, sticky="ew")
        self._rv_info = tk.Label(info, text="Charge eingeben und suchen - oder leer lassen fuer alle Chargen mit Ablauf < 90 Tage.",
                                 bg=SHELL_BG, fg=MUTED, font=(theme.FONT, 9))
        self._rv_info.pack(anchor="w", pady=(0, 6))

        split = tk.Frame(parent, bg=SHELL_BG)
        split.grid(row=2, column=0, sticky="nsew")
        split.columnconfigure(0, weight=1)
        split.columnconfigure(1, weight=1)
        split.rowconfigure(0, weight=1)
        outer_in, in_body = _card(split, "Eingang", "Wo kam die Charge herein?")
        outer_in.grid(row=0, column=0, sticky="nsew", padx=(0, 6))
        self._rv_in_tree = _make_tree(in_body, ("Datum", "Lieferant", "Artikel", "Charge", "Verfall", "Menge"),
                                      (80, 120, 200, 90, 80, 60), anchors={"Menge": "center"})
        outer_out, out_body = _card(split, "Ausgang", "Welche Apotheke hat diese Charge erhalten?")
        outer_out.grid(row=0, column=1, sticky="nsew", padx=(6, 0))
        self._rv_out_tree = _make_tree(out_body, ("Datum", "Apotheke", "Artikel", "Charge", "Menge", "Beleg"),
                                       (80, 160, 170, 90, 55, 90), anchors={"Menge": "center"})
        self._rv_out_tree.bind("<Double-1>", lambda _e: self._rv_show_kontakt())
        self._refreshers["rueckverfolgung"] = self._rv_search

    def _rv_search(self):
        term = self._rv_var.get().strip()
        like = f"%{term}%"
        for t in (self._rv_in_tree, self._rv_out_tree):
            t.delete(*t.get_children())
        grenze = HEUTE() + timedelta(days=90)
        aus = []
        with self._conn() as con:
            if term:
                ein = con.execute(
                    """SELECT w.datum, w.lieferant, p.artikelname, p.charge, p.verfall, p.menge
                       FROM tbl_wareneingang_positionen p JOIN tbl_wareneingang w ON w.id=p.we_id
                       WHERE p.charge LIKE ? OR p.pzn LIKE ? ORDER BY w.datum DESC""",
                    (like, like)).fetchall() if _table_exists(con, "tbl_wareneingang") else []
                aus = self._outbound_rows(con, term=term)
            else:
                ein = []
                if _table_exists(con, "tbl_lagerbestand"):
                    for pzn, name, charge, verf, menge in con.execute(
                        "SELECT pzn, artikelname, charge, verfall, menge FROM tbl_lagerbestand WHERE menge>0"):
                        d = _parse_verfall(verf)
                        if d and d <= grenze:
                            ein.append(("Lager", "", name, charge, verf, menge))
        for i, r in enumerate(ein):
            tag = "alert" if (_parse_verfall(r[4]) and _parse_verfall(r[4]) <= grenze) else ("even" if i % 2 else "odd")
            self._rv_in_tree.insert("", "end", values=r, tags=(tag,))
        kunden_set = set()
        self._rv_out_map = {}
        for i, r in enumerate(aus):
            # r = (datum,knr,kunde,pzn,artikel,charge,verfall,menge,beleg,quelle)
            kunden_set.add(r[1])
            beleg = f"{r[8]} ({r[9]})" if r[8] else r[9]
            iid = self._rv_out_tree.insert("", "end", values=(r[0], r[2], r[4], r[5], r[7], beleg),
                                           tags=("even" if i % 2 else "odd",))
            self._rv_out_map[iid] = (r[1], r[2])
        if term:
            self._rv_info.config(
                text=f"Charge/PZN '{term}': {len(ein)} Eingangs- und {len(aus)} Ausgangsbewegung(en) "
                     f"aus Kasse-Verkaeufen + GDP-Auslieferungen. "
                     f"{len(kunden_set)} Apotheke(n) bei Rueckruf betroffen.")
        else:
            self._rv_info.config(text=f"{len(ein)} Charge(n) im Lager mit Verfall < 90 Tage. "
                                      "Charge/PZN eingeben fuer die Kunden-Rueckverfolgung.")

    def _rv_rueckruf(self):
        term = self._rv_var.get().strip()
        if not term:
            messagebox.showinfo("Rueckruf", "Bitte zuerst eine konkrete Charge/PZN suchen.", parent=self)
            return
        with self._conn() as con:
            rows = self._outbound_rows(con, term=term)
        if not rows:
            messagebox.showinfo("Rueckruf",
                                f"Zu '{term}' sind keine Auslieferungen/Verkaeufe erfasst - kein Verteiler.",
                                parent=self)
            return
        n_kunden = len({r[1] for r in rows})
        summe = sum(r[7] or 0 for r in rows)
        artikel = next((r[4] for r in rows if r[4]), "")
        pzn = next((r[3] for r in rows if r[3]), "")
        if not messagebox.askyesno(
                "Rueckruf ausloesen",
                f"Charge/PZN '{term}'\nArtikel: {artikel}\n\n"
                f"{n_kunden} Apotheke(n), {summe} Stueck betroffen.\n\n"
                "Rueckruf anlegen und Verteiler vormerken?", parent=self):
            return

        def save(v):
            with self._conn() as con:
                con.execute(
                    """INSERT INTO tbl_gdp_rueckruf
                       (datum,charge,pzn,artikelname,grund,betroffene_kunden,betroffene_menge,
                        status,ausgeloest_von,bemerkung)
                       VALUES(?,?,?,?,?,?,?,?,?,?)""",
                    (_heute_iso(), term, pzn, artikel, v["grund"], n_kunden, summe,
                     "ausgeloest", ME(), v["bemerkung"]))
                con.commit()
            self._log("Rueckruf", "Rueckruf ausgeloest", None,
                      f"Charge {term}: {n_kunden} Apotheken, {summe} Stueck")
            messagebox.showinfo("Rueckruf",
                                f"Rueckruf dokumentiert. {n_kunden} Apotheke(n) im Verteiler.\n"
                                "Verteiler ueber 'Verteiler exportieren' als Liste sichern.", parent=self)

        FormDialog(self, f"Rueckruf Charge {term}", [
            ("grund", "Rueckruf-Grund", "combo",
             ["Qualitaetsmangel", "Behoerdlicher Rueckruf", "Temperaturabweichung",
              "Verpackungsfehler", "Faelschungsverdacht", "Sonstiges"], "Qualitaetsmangel"),
            ("bemerkung", "Bemerkung", "multiline", ""),
        ], save)

    def _rv_export(self):
        term = self._rv_var.get().strip()
        if not term:
            messagebox.showinfo("Export", "Bitte zuerst eine Charge/PZN suchen.", parent=self)
            return
        with self._conn() as con:
            rows = self._outbound_rows(con, term=term)
            kontakt = self._kontakt_map(con)
        if not rows:
            messagebox.showinfo("Export", "Keine betroffenen Apotheken gefunden.", parent=self)
            return
        rows = sorted(rows, key=lambda r: (r[2] or "", r[0] or ""))
        path = filedialog.asksaveasfilename(
            parent=self, title="Verteiler speichern", defaultextension=".csv",
            initialfile=f"Rueckruf_Verteiler_{term}.csv",
            filetypes=[("CSV", "*.csv")])
        if not path:
            return
        with open(path, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f, delimiter=";")
            w.writerow(["Datum", "Kundennr", "Apotheke", "Artikel", "Charge", "Verfall",
                        "Menge", "Beleg", "Quelle", "E-Mail", "Telefon", "Ort"])
            for r in rows:
                email, tel, ort = kontakt.get(r[1], ("", "", ""))
                w.writerow([r[0], r[1], r[2], r[4], r[5], r[6], r[7],
                            r[8], r[9], email, tel, ort])
        self._log("Rueckruf", "Verteiler exportiert", None, f"Charge {term}: {len(rows)} Zeilen")
        messagebox.showinfo("Export", f"Verteiler mit {len(rows)} Zeile(n) gespeichert:\n{path}", parent=self)

    def _kontakt_map(self, con):
        """kundennummer -> (email, telefon, ort) aus dem Kundenstamm (Best-Effort)."""
        out = {}
        if not _table_exists(con, "tbl_kunden_center"):
            return out
        have = {r[1] for r in con.execute("PRAGMA table_info(tbl_kunden_center)")}
        email_c = "email" if "email" in have else "NULL"
        tel_c = "telefon" if "telefon" in have else ("telefon1" if "telefon1" in have else "NULL")
        ort_c = "ort" if "ort" in have else "NULL"
        for knr, email, tel, ort in con.execute(
                f"SELECT kundennummer, {email_c}, {tel_c}, {ort_c} FROM tbl_kunden_center"):
            out[knr] = (email or "", tel or "", ort or "")
        return out

    def _rv_show_kontakt(self):
        sel = self._rv_out_tree.selection()
        if not sel or not getattr(self, "_rv_out_map", None):
            return
        knr, name = self._rv_out_map.get(sel[0], ("", ""))
        with self._conn() as con:
            email, tel, ort = self._kontakt_map(con).get(knr, ("", "", ""))
            quali = con.execute(
                "SELECT qualifiziert, lizenz_gueltig_bis FROM tbl_gdp_kunde_quali WHERE kundennummer=?",
                (knr,)).fetchone() if _table_exists(con, "tbl_gdp_kunde_quali") else None
        lz = ""
        if quali is not None:
            ok = quali[0] and not (quali[1] and quali[1] < _heute_iso())
            lz = f"\nQualifizierung: {'freigegeben' if ok else 'PRUEFEN/gesperrt'}"
            if quali[1]:
                lz += f" (Lizenz gueltig bis {quali[1]})"
        messagebox.showinfo(
            f"Apotheke {name}",
            f"Kundennr: {knr}\nOrt: {ort or '-'}\nE-Mail: {email or '-'}\n"
            f"Telefon: {tel or '-'}{lz}", parent=self)

    def _rv_dossier(self):
        """Vollstaendiges Chargendossier (Lebenslauf einer Charge) als CSV:
        Eingang -> GDP-Pruefung -> Ausgaenge/Apotheken -> Retouren -> Rueckrufe.
        Ideal fuer Inspektionen und die Rueckruf-Dokumentation."""
        term = self._rv_var.get().strip()
        if not term:
            messagebox.showinfo("Chargendossier", "Bitte zuerst eine konkrete Charge eingeben.", parent=self)
            return
        with self._conn() as con:
            kontakt = self._kontakt_map(con)
            eingang = con.execute(
                """SELECT w.datum, w.lieferant, w.lieferschein, p.pzn, p.artikelname,
                          p.charge, p.verfall, p.menge, p.ek
                   FROM tbl_wareneingang_positionen p JOIN tbl_wareneingang w ON w.id=p.we_id
                   WHERE p.charge=? ORDER BY w.datum""", (term,)).fetchall() \
                if _table_exists(con, "tbl_wareneingang") else []
            we_ids = [r[0] for r in con.execute(
                """SELECT DISTINCT w.id FROM tbl_wareneingang w
                   JOIN tbl_wareneingang_positionen p ON p.we_id=w.id WHERE p.charge=?""",
                (term,))] if _table_exists(con, "tbl_wareneingang") else []
            pruef = con.execute(
                "SELECT datum,lieferant,transport_temp_c,gdp_konform,geprueft_von,bemerkung "
                "FROM tbl_gdp_we_pruefung WHERE we_id IN (%s)" %
                (",".join("?" * len(we_ids)) or "NULL"), we_ids).fetchall() if we_ids else []
            ausgang = self._outbound_rows(con, charge=term)
            retouren = con.execute(
                "SELECT datum,typ,kunde_name,menge,grund,status,entscheidung,gutschrift_beleg "
                "FROM tbl_gdp_retoure WHERE charge=? ORDER BY id", (term,)).fetchall()
            rueckrufe = con.execute(
                "SELECT datum,grund,betroffene_kunden,betroffene_menge,status,ausgeloest_von "
                "FROM tbl_gdp_rueckruf WHERE charge=? ORDER BY id", (term,)).fetchall()
        if not (eingang or ausgang or retouren):
            messagebox.showinfo("Chargendossier",
                                f"Zur Charge '{term}' sind keine Bewegungen erfasst.", parent=self)
            return
        path = filedialog.asksaveasfilename(
            parent=self, title="Chargendossier speichern", defaultextension=".csv",
            initialfile=f"Chargendossier_{term}.csv", filetypes=[("CSV", "*.csv")])
        if not path:
            return
        with open(path, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f, delimiter=";")
            w.writerow([f"CHARGENDOSSIER  Charge {term}",
                        f"erstellt {datetime.now().isoformat(timespec='minutes')} von {ME()}"])
            w.writerow([])
            w.writerow(["== WARENEINGANG =="])
            w.writerow(["Datum", "Lieferant", "Lieferschein", "PZN", "Artikel", "Charge", "Verfall", "Menge", "EK"])
            w.writerows(eingang or [["(keine)"]])
            w.writerow([])
            w.writerow(["== GDP-WARENEINGANGSPRUEFUNG =="])
            w.writerow(["Datum", "Lieferant", "Transport-Temp", "konform", "geprueft von", "Bemerkung"])
            w.writerows([[a, b, c, "ja" if dd else "NEIN", e, g] for (a, b, c, dd, e, g) in pruef] or [["(keine)"]])
            w.writerow([])
            w.writerow(["== AUSGANG / BELIEFERTE APOTHEKEN =="])
            w.writerow(["Datum", "Kundennr", "Apotheke", "Artikel", "Menge", "Beleg", "Quelle", "E-Mail", "Telefon", "Ort"])
            for r in sorted(ausgang, key=lambda x: (x[2] or "")):
                email, tel, ort = kontakt.get(r[1], ("", "", ""))
                w.writerow([r[0], r[1], r[2], r[4], r[7], r[8], r[9], email, tel, ort])
            if not ausgang:
                w.writerow(["(keine)"])
            w.writerow([])
            w.writerow(["== RETOUREN / REKLAMATIONEN =="])
            w.writerow(["Datum", "Typ", "Apotheke", "Menge", "Grund", "Status", "Entscheidung", "Gutschrift"])
            w.writerows(retouren or [["(keine)"]])
            w.writerow([])
            w.writerow(["== RUECKRUFE =="])
            w.writerow(["Datum", "Grund", "betroffene Apotheken", "betroffene Menge", "Status", "ausgeloest von"])
            w.writerows(rueckrufe or [["(keine)"]])
        self._log("Rueckruf", "Chargendossier exportiert", None,
                  f"Charge {term}: {len(eingang)} Eingang, {len(ausgang)} Ausgang, {len(retouren)} Retouren")
        messagebox.showinfo("Chargendossier",
                            f"Dossier zur Charge {term} gespeichert:\n{path}\n\n"
                            f"Eingang: {len(eingang)} · Ausgang: {len(ausgang)} Apotheke(n) · "
                            f"Retouren: {len(retouren)}", parent=self)

    # ===================================================== RETOUREN / REKLAMAT.
    def _build_retouren(self, parent):
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(2, weight=1)
        bar = tk.Frame(parent, bg=SHELL_BG)
        bar.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        theme.PillButton(bar, "➕  Neue Retoure / Reklamation", self._ret_new,
                         kind="primary", font_size=10).pack(side="left")
        theme.PillButton(bar, "\U0001F50D  Pruefen", lambda: self._ret_action("pruefen"),
                         kind="neutral", font_size=10).pack(side="left", padx=6)
        theme.PillButton(bar, "\U0001F9FE  Gutschrift", lambda: self._ret_action("gutschrift"),
                         kind="accent", font_size=10).pack(side="left", padx=6)
        theme.PillButton(bar, "\U0001F4E5  In Retourenbestand", lambda: self._ret_action("wieder"),
                         kind="success", font_size=10).pack(side="left", padx=6)
        theme.PillButton(bar, "\U0001F5D1  Vernichten", lambda: self._ret_action("vernichten"),
                         kind="danger", font_size=10).pack(side="left", padx=6)
        theme.PillButton(bar, "✖  Ablehnen", lambda: self._ret_action("ablehnen"),
                         kind="neutral", font_size=10).pack(side="left", padx=6)
        theme.PillButton(bar, "\U0001F5D1  Storno", lambda: self._ret_action("storno"),
                         kind="neutral", font_size=10).pack(side="left", padx=6)

        bar2 = tk.Frame(parent, bg=SHELL_BG)
        bar2.grid(row=1, column=0, sticky="ew", pady=(0, 6))
        tk.Label(bar2, text="Filter Status:", bg=SHELL_BG, fg=MUTED,
                 font=(theme.FONT, 9)).pack(side="left")
        self._ret_filter = tk.StringVar(value="alle")
        cb = ttk.Combobox(bar2, textvariable=self._ret_filter, state="readonly", width=18,
                          style="NMG.TCombobox",
                          values=["alle", "offen (Neu/In Pruefung)", "Neu", "In Pruefung",
                                  "Gutschrift", "Wiedereingelagert", "Vernichtet", "Abgelehnt"])
        cb.pack(side="left", padx=8)
        cb.bind("<<ComboboxSelected>>", lambda _e: self._refresh_retouren())
        self._ret_count = tk.Label(bar2, text="", bg=SHELL_BG, fg=MUTED, font=(theme.FONT, 9))
        self._ret_count.pack(side="left", padx=12)

        outer, body = _card(parent)
        outer.grid(row=2, column=0, sticky="nsew")
        cols = ("Nr", "Datum", "Typ", "Apotheke", "Artikel", "Charge", "Menge", "Grund", "Status", "Gutschrift")
        self._ret_tree = _make_tree(body, cols, (40, 80, 90, 140, 170, 80, 55, 130, 110, 100),
                                    anchors={"Nr": "center", "Menge": "center"})
        self._ret_tree.bind("<Double-1>", lambda _e: self._ret_detail())
        self._refreshers["retouren"] = self._refresh_retouren

    def _refresh_retouren(self):
        t = self._ret_tree
        t.delete(*t.get_children())
        flt = self._ret_filter.get() if hasattr(self, "_ret_filter") else "alle"
        if flt == "alle":
            where, params = "", ()
        elif flt.startswith("offen"):
            where, params = "WHERE status IN ('Neu','In Pruefung')", ()
        else:
            where, params = "WHERE status=?", (flt,)
        with self._conn() as con:
            rows = con.execute(
                f"""SELECT id, datum, typ, kunde_name, artikelname, charge, menge, grund,
                          status, gutschrift_beleg, temperaturbruch
                   FROM tbl_gdp_retoure {where} ORDER BY id DESC""", params).fetchall()
            offen = con.execute(
                "SELECT COUNT(*) FROM tbl_gdp_retoure WHERE status IN ('Neu','In Pruefung')").fetchone()[0]
            ges = con.execute("SELECT COUNT(*) FROM tbl_gdp_retoure").fetchone()[0]
        for r in rows:
            (rid, datum, typ, kunde, artikel, charge, menge, grund, status, gut, tbruch) = r
            tag = RET_STATUS_TAG.get(status, "odd")
            grund_disp = ("❄ " if tbruch else "") + (grund or "")
            t.insert("", "end", iid=str(rid),
                     values=(rid, datum or "", typ or "", kunde or "", artikel or "",
                             charge or "", menge, grund_disp, status or "", gut or ""),
                     tags=(tag,))
        if hasattr(self, "_ret_count"):
            self._ret_count.config(text=f"{len(rows)} angezeigt · {offen} offen · {ges} gesamt")

    def _ret_selected(self):
        sel = self._ret_tree.selection()
        if not sel:
            return None
        with self._conn() as con:
            return con.execute("SELECT * FROM tbl_gdp_retoure WHERE id=?", (int(sel[0]),)).fetchone()

    def _ret_cols(self):
        with self._conn() as con:
            return [d[1] for d in con.execute("PRAGMA table_info(tbl_gdp_retoure)")]

    def _ret_new(self):
        # Artikel-/Kundenlisten vorbereiten (aus Lager + Kundenstamm)
        with self._conn() as con:
            kunden = con.execute(
                "SELECT kundennummer, kundenname FROM tbl_kunden_center ORDER BY kundenname"
            ).fetchall() if _table_exists(con, "tbl_kunden_center") else []
            # Artikel/Chargen aus ZWEI Quellen: aktueller Lagerbestand UND
            # tatsaechlich gelieferte Chargen (auch was nicht mehr im Lager liegt).
            artikel = {}  # (pzn,charge,verfall) -> (pzn,name,charge,verfall)
            if _table_exists(con, "tbl_lagerbestand"):
                for pzn, name, ch, vf in con.execute(
                        "SELECT DISTINCT pzn, artikelname, charge, verfall FROM tbl_lagerbestand"):
                    artikel[(pzn, ch or "", vf or "")] = (pzn, name, ch, vf)
            for r in self._outbound_rows(con):
                key = (r[3], r[5] or "", r[6] or "")
                artikel.setdefault(key, (r[3], r[4], r[5], r[6]))
        kunden_labels = [f"{n} ({knr})" for knr, n in kunden]
        kunden_map = {f"{n} ({knr})": (knr, n) for knr, n in kunden}
        chargen = sorted(artikel.values(), key=lambda x: (x[1] or "", x[2] or ""))
        art_labels = [f"{name} | Ch {ch or '-'} | {pzn}" for pzn, name, ch, vf in chargen]
        art_map = {f"{name} | Ch {ch or '-'} | {pzn}": (pzn, name, ch, vf)
                   for pzn, name, ch, vf in chargen}

        def save(v):
            knr, kname = kunden_map.get(v["kunde"], ("", v["kunde"]))
            pzn, name, charge, verfall = art_map.get(v["artikel"], ("", v["artikel"], "", ""))
            with self._conn() as con:
                cur = con.execute(
                    """INSERT INTO tbl_gdp_retoure
                       (datum,typ,kundennummer,kunde_name,pzn,artikelname,charge,verfall,
                        menge,grund,temperaturbruch,status,bearbeiter,notiz)
                       VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                    (_heute_iso(), v["typ"], knr, kname, pzn, name, charge, verfall,
                     v["menge"], v["grund"], v["temperaturbruch"], "Neu", ME(), v["notiz"]))
                rid = cur.lastrowid
                con.commit()
            self._log("Retoure", "Retoure angelegt", rid, f"{kname}: {name} (Ch {charge})")
            self._refresh_retouren()

        FormDialog(self, "Neue Retoure / Reklamation", [
            ("typ", "Vorgang", "combo", ["Retoure", "Reklamation"], "Retoure"),
            ("kunde", "Apotheke", "combo", kunden_labels, kunden_labels[0] if kunden_labels else ""),
            ("artikel", "Artikel / Charge", "combo", art_labels, art_labels[0] if art_labels else ""),
            ("menge", "Menge", "int", "1"),
            ("grund", "Ruecksendegrund", "combo",
             ["Falschlieferung", "Ablauf/Verfall nahe", "Transportschaden", "Temperaturbruch",
              "Ueberbestellung", "Qualitaetsmangel", "Bestellfehler Apotheke", "Sonstiges"],
             "Falschlieferung"),
            ("temperaturbruch", "Temperaturbruch", "check", 0),
            ("notiz", "Notiz", "multiline", ""),
        ], save, width=480)

    def _ret_detail(self):
        row = self._ret_selected()
        if not row:
            return
        cols = self._ret_cols()
        d = dict(zip(cols, row))
        with self._conn() as con:
            log = con.execute(
                "SELECT zeitpunkt, bearbeiter, aktion, details FROM tbl_gdp_log "
                "WHERE modul='Retoure' AND bezug_id=? ORDER BY id", (d["id"],)).fetchall()
        win = tk.Toplevel(self)
        win.title(f"Retoure #{d['id']}")
        win.configure(bg=SHELL_BG)
        win.geometry("560x520")
        tk.Label(win, bg=ACCENT, fg="#FFFFFF", anchor="w", font=(theme.FONT, 12, "bold"),
                 text=f"  {d['typ']} #{d['id']}  ·  {d['status']}").pack(fill="x", ipady=8)
        info = tk.Frame(win, bg=SHELL_BG)
        info.pack(fill="x", padx=16, pady=10)
        pairs = [("Datum", d["datum"]), ("Apotheke", d["kunde_name"]),
                 ("Artikel", d["artikelname"]), ("PZN", d["pzn"]),
                 ("Charge", d["charge"]), ("Verfall", d["verfall"]),
                 ("Menge", d["menge"]), ("Grund", d["grund"]),
                 ("Temperaturbruch", "Ja" if d["temperaturbruch"] else "Nein"),
                 ("Entscheidung", d["entscheidung"] or "-"),
                 ("Gutschrift", d["gutschrift_beleg"] or "-"),
                 ("Notiz", d["notiz"] or "-")]
        for i, (k, val) in enumerate(pairs):
            r, c = divmod(i, 2)
            cell = tk.Frame(info, bg=SHELL_BG)
            cell.grid(row=r, column=c, sticky="w", padx=6, pady=2)
            tk.Label(cell, text=f"{k}: ", bg=SHELL_BG, fg=MUTED, font=(theme.FONT, 9, "bold")).pack(side="left")
            tk.Label(cell, text=str(val), bg=SHELL_BG, fg=TEXT, font=(theme.FONT, 9)).pack(side="left")
        tk.Label(win, text="Verlauf", bg=SHELL_BG, fg=ACCENT,
                 font=(theme.FONT, 11, "bold")).pack(anchor="w", padx=16, pady=(8, 2))
        lw = tk.Frame(win, bg=SHELL_BG)
        lw.pack(fill="both", expand=True, padx=16, pady=(0, 12))
        tree = _make_tree(lw, ("Zeit", "Wer", "Aktion", "Details"), (130, 90, 150, 150))
        for i, lr in enumerate(log):
            tree.insert("", "end", values=lr, tags=("even" if i % 2 else "odd",))

    def _ret_action(self, action):
        row = self._ret_selected()
        if not row:
            messagebox.showinfo("Retoure", "Bitte zuerst eine Retoure in der Liste waehlen.", parent=self)
            return
        cols = self._ret_cols()
        d = dict(zip(cols, row))
        rid, status = d["id"], d["status"]

        if action == "pruefen":
            if status != "Neu":
                messagebox.showinfo("Pruefen", "Nur neue Vorgaenge koennen in Pruefung gehen.", parent=self)
                return
            self._ret_set(rid, status="In Pruefung")
            self._log("Retoure", "In Pruefung genommen", rid, "")
        elif action == "gutschrift":
            if status not in ("Neu", "In Pruefung"):
                messagebox.showinfo("Gutschrift", "Fuer diesen Status nicht mehr moeglich.", parent=self)
                return
            self._ret_gutschrift(d)
            return
        elif action == "wieder":
            if status not in ("Neu", "In Pruefung", "Gutschrift"):
                messagebox.showinfo("Retourenbestand",
                    "Nur offene oder gutgeschriebene Retouren koennen ins "
                    "Retourenlager genommen werden.", parent=self)
                return
            if not messagebox.askyesno("In Retourenbestand nehmen",
                    f"Charge {d['charge']} ({d['menge']} St) in den Retourenbestand "
                    "(Quarantaene) buchen?\n\nDie Ware ist dort GESPERRT und noch nicht "
                    "verkaufbar. Die Freigabe oder Abschreibung erfolgt anschliessend "
                    "im Bereich 'Retourenbestand'.", parent=self):
                return
            self._ret_restock(d)
        elif action == "vernichten":
            if not messagebox.askyesno("Vernichten",
                    f"Charge {d['charge']} ({d['menge']} St) als vernichtet dokumentieren?", parent=self):
                return
            self._ret_set(rid, status="Vernichtet", entscheidung="Vernichtung",
                          abgeschlossen_am=_heute_iso())
            self._log("Retoure", "Vernichtung dokumentiert", rid, f"{d['menge']} St Charge {d['charge']}")
        elif action == "ablehnen":
            self._ret_set(rid, status="Abgelehnt", entscheidung="Abgelehnt",
                          abgeschlossen_am=_heute_iso())
            self._log("Retoure", "Retoure abgelehnt", rid, "")
        elif action == "storno":
            if status not in ("Neu", "In Pruefung"):
                messagebox.showinfo("Storno",
                    "Nur Vorgaenge im Status 'Neu' oder 'In Pruefung' koennen "
                    "geloescht werden (sonst Buchungen/Gutschrift vorhanden).", parent=self)
                return
            if not messagebox.askyesno("Storno",
                    f"Retoure #{rid} ({d['kunde_name']}, {d['artikelname']}) "
                    "endgueltig loeschen?", parent=self):
                return
            with self._conn() as con:
                con.execute("DELETE FROM tbl_gdp_retoure WHERE id=?", (rid,))
                con.commit()
            self._log("Retoure", "Retoure storniert/geloescht", rid,
                      f"{d['kunde_name']}: {d['artikelname']}")
        self._refresh_retouren()

    def _ret_set(self, rid, **fields):
        sets = ", ".join(f"{k}=?" for k in fields)
        with self._conn() as con:
            con.execute(f"UPDATE tbl_gdp_retoure SET {sets} WHERE id=?",
                        (*fields.values(), rid))
            con.commit()

    def _ret_restock(self, d):
        """Ruecknahme ins Retourenlager: Menge in den GESPERRTEN Retourenbestand
        (tbl_lagerbestand.menge_retoure) buchen - NICHT in den verkaufbaren
        Normalbestand. Freigabe/Abschreibung erfolgt im Bereich Retourenbestand."""
        with self._conn() as con:
            if _table_exists(con, "tbl_lagerbestand"):
                ex = con.execute(
                    "SELECT id FROM tbl_lagerbestand WHERE pzn=? AND COALESCE(charge,'')=? "
                    "AND COALESCE(verfall,'')=?", (d["pzn"], d["charge"] or "", d["verfall"] or "")).fetchone()
                if ex:
                    con.execute("UPDATE tbl_lagerbestand SET menge_retoure=COALESCE(menge_retoure,0)+?, "
                                "aktualisiert_am=? WHERE id=?", (d["menge"], _heute_iso(), ex[0]))
                else:
                    con.execute(
                        "INSERT INTO tbl_lagerbestand(pzn,artikelname,charge,verfall,menge,menge_retoure,aktualisiert_am) "
                        "VALUES(?,?,?,?,0,?,?)",
                        (d["pzn"], d["artikelname"], d["charge"], d["verfall"], d["menge"], _heute_iso()))
                con.commit()
        self._ret_set(d["id"], status="Im Retourenbestand", entscheidung="In Quarantaene",
                      abgeschlossen_am=_heute_iso())
        self._log("Retoure", "In Retourenbestand genommen", d["id"],
                  f"+{d['menge']} St Charge {d['charge']} in Quarantaene")

    def _ret_gutschrift(self, d):
        """Erzeugt eine Gutschrift. Wenn die Faktura-Tabellen existieren, wird ein
        echter Gutschrift-Beleg angelegt (Faktura-Anbindung); sonst nur eine
        interne Gutschrift-Nummer vergeben."""
        beleg_nr = None
        faktura_id = None
        try:
            with self._conn() as con:
                if _table_exists(con, "tbl_faktura_belege"):
                    jahr = HEUTE().year
                    # Nummernkreis nutzen, falls vorhanden
                    nr = None
                    if _table_exists(con, "tbl_faktura_nummernkreis"):
                        r = con.execute(
                            "SELECT letzter_zaehler FROM tbl_faktura_nummernkreis "
                            "WHERE belegart='gutschrift' AND jahr=?", (jahr,)).fetchone()
                        if r:
                            nr = (r[0] or 0) + 1
                            con.execute("UPDATE tbl_faktura_nummernkreis SET letzter_zaehler=? "
                                        "WHERE belegart='gutschrift' AND jahr=?", (nr, jahr))
                    if nr is None:
                        cnt = con.execute("SELECT COUNT(*) FROM tbl_faktura_belege "
                                          "WHERE belegart='gutschrift'").fetchone()[0]
                        nr = 1001 + cnt
                    beleg_nr = f"GU-{jahr}-{nr}"
                    # Betrag aus APU schaetzen
                    apu = 0.0
                    if _table_exists(con, "tbl_nmg_stamm"):
                        rr = con.execute("SELECT apu FROM tbl_nmg_stamm WHERE pzn=? LIMIT 1",
                                         (d["pzn"],)).fetchone()
                        apu = float(rr[0]) if rr and rr[0] else 0.0
                    netto = round(apu * (d["menge"] or 0), 2)
                    ust = round(netto * 0.19, 2)
                    brutto = round(netto + ust, 2)
                    have = {c[1] for c in con.execute("PRAGMA table_info(tbl_faktura_belege)")}
                    if {"belegart", "beleg_nr", "kunde_nr", "kunde_name"} <= have:
                        cur = con.execute(
                            """INSERT INTO tbl_faktura_belege
                               (belegart,beleg_nr,kunde_nr,kunde_name,beleg_datum,netto,ust_betrag,brutto,
                                status,mitarbeiter,erstellt_am)
                               VALUES('gutschrift',?,?,?,?,?,?,?, 'offen', ?, ?)""",
                            (beleg_nr, d["kundennummer"], d["kunde_name"], _heute_iso(),
                             -netto, -ust, -brutto, ME(), datetime.now().isoformat(timespec="seconds")))
                        faktura_id = cur.lastrowid
                        if _table_exists(con, "tbl_faktura_positionen"):
                            pcols = {c[1] for c in con.execute("PRAGMA table_info(tbl_faktura_positionen)")}
                            if {"beleg_id", "pzn", "menge"} <= pcols:
                                con.execute(
                                    """INSERT INTO tbl_faktura_positionen
                                       (beleg_id,pos_nr,pzn,bezeichnung,menge,apu_einzel,rabatt,
                                        ust_satz,netto_zeile,ust_zeile,brutto_zeile)
                                       VALUES(?,1,?,?,?,?,0,19,?,?,?)""",
                                    (faktura_id, d["pzn"], d["artikelname"], d["menge"], apu,
                                     -netto, -ust, -brutto))
                    con.commit()
        except Exception as exc:
            beleg_nr = beleg_nr or None
            messagebox.showwarning("Gutschrift",
                                   f"Faktura-Beleg konnte nicht angelegt werden:\n{exc}\n"
                                   "Es wird nur eine interne Gutschrift-Nummer vergeben.", parent=self)
        if not beleg_nr:
            beleg_nr = f"GU-RET-{d['id']}"
        self._ret_set(d["id"], status="Gutschrift", entscheidung="Gutschrift erteilt",
                      gutschrift_beleg=beleg_nr, faktura_beleg_id=faktura_id)
        self._log("Retoure", "Gutschrift erstellt", d["id"],
                  f"Beleg {beleg_nr}" + (" (Faktura)" if faktura_id else ""))
        self._refresh_retouren()
        messagebox.showinfo("Gutschrift",
                            f"Gutschrift {beleg_nr} erstellt." +
                            ("\nAls offener Beleg in der Faktura hinterlegt." if faktura_id else ""),
                            parent=self)

    # ===================================================== RETOURENBESTAND
    def _build_retourenbestand(self, parent):
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(2, weight=1)
        bar = tk.Frame(parent, bg=SHELL_BG)
        bar.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        theme.PillButton(bar, "✅  In Bestand freigeben", self._rb_freigeben,
                         kind="success", font_size=10).pack(side="left")
        theme.PillButton(bar, "\U0001F5D1  Abschreiben", self._rb_abschreiben,
                         kind="danger", font_size=10).pack(side="left", padx=6)
        theme.PillButton(bar, "↻  Aktualisieren", self._refresh_retourenbestand,
                         kind="neutral", font_size=10).pack(side="left", padx=6)
        theme.PillButton(bar, "⬇  Abschreibungen (CSV)", self._rb_export_abschreibungen,
                         kind="neutral", font_size=10).pack(side="left", padx=6)
        tk.Label(bar, text="Anzeige: Normalbestand (Retourenbestand). Retourenware ist "
                           "gesperrt, bis sie freigegeben wird.",
                 bg=SHELL_BG, fg=MUTED, font=(theme.FONT, 9)).pack(side="left", padx=12)

        self._rb_summary = tk.Label(parent, text="", bg=SHELL_BG, fg=TEXT,
                                    font=(theme.FONT, 10, "bold"))
        self._rb_summary.grid(row=1, column=0, sticky="w", pady=(0, 6))

        split = tk.Frame(parent, bg=SHELL_BG)
        split.grid(row=2, column=0, sticky="nsew")
        split.columnconfigure(0, weight=3)
        split.columnconfigure(1, weight=2)
        split.rowconfigure(0, weight=1)
        outer_q, q_body = _card(split, "Quarantaene / Retourenbestand",
                                "Zu entscheiden: freigeben oder abschreiben")
        outer_q.grid(row=0, column=0, sticky="nsew", padx=(0, 6))
        self._rb_tree = _make_tree(
            q_body, ("PZN", "Artikel", "Charge", "Verfall", "Bestand (Retoure)", "Retoure", "EK"),
            (80, 210, 90, 80, 130, 70, 70),
            anchors={"Bestand (Retoure)": "center", "Retoure": "center", "EK": "e"})
        outer_a, a_body = _card(split, "Abschreibungen", "Dokumentierte Vernichtungen/Write-offs")
        outer_a.grid(row=0, column=1, sticky="nsew", padx=(6, 0))
        self._rb_ab_tree = _make_tree(
            a_body, ("Datum", "Artikel", "Charge", "Menge", "Grund", "Wert EK"),
            (80, 170, 80, 55, 130, 80), anchors={"Menge": "center", "Wert EK": "e"})
        self._refreshers["retourenbestand"] = self._refresh_retourenbestand

    def _refresh_retourenbestand(self):
        t = self._rb_tree
        t.delete(*t.get_children())
        ges_ret = wert = 0
        with self._conn() as con:
            rows = []
            if _table_exists(con, "tbl_lagerbestand"):
                lcols = {r[1] for r in con.execute("PRAGMA table_info(tbl_lagerbestand)")}
                if "menge_retoure" in lcols:
                    rows = con.execute(
                        "SELECT id, pzn, artikelname, charge, verfall, menge, "
                        "COALESCE(menge_retoure,0), ek FROM tbl_lagerbestand "
                        "WHERE COALESCE(menge_retoure,0) > 0 ORDER BY artikelname").fetchall()
            ab = con.execute(
                "SELECT datum, artikelname, charge, menge, grund, wert_ek "
                "FROM tbl_gdp_abschreibung ORDER BY id DESC LIMIT 200").fetchall()
        for i, (lid, pzn, name, charge, verf, menge, mret, ek) in enumerate(rows):
            ges_ret += mret
            wert += (ek or 0) * mret
            bestand = f"{menge}  ({mret})"
            t.insert("", "end", iid=str(lid),
                     values=(pzn or "", name or "", charge or "", verf or "", bestand, mret,
                             f"{ek:.2f}" if ek else ""),
                     tags=("warn",))
        at = self._rb_ab_tree
        at.delete(*at.get_children())
        for i, (datum, name, charge, menge, grund, w) in enumerate(ab):
            at.insert("", "end", values=(datum or "", name or "", charge or "", menge,
                                         grund or "", f"{w:.2f}" if w else ""),
                      tags=("even" if i % 2 else "odd",))
        self._rb_summary.config(
            text=f"Im Retourenbestand: {len(rows)} Charge(n) · {ges_ret} Stueck · "
                 f"EK-Wert ~ {wert:.2f} EUR   |   Zeile waehlen und freigeben oder abschreiben")

    def _rb_selected(self):
        sel = self._rb_tree.selection()
        if not sel:
            return None
        with self._conn() as con:
            r = con.execute(
                "SELECT id, pzn, artikelname, charge, verfall, menge, COALESCE(menge_retoure,0), ek "
                "FROM tbl_lagerbestand WHERE id=?", (int(sel[0]),)).fetchone()
        if not r:
            return None
        keys = ("id", "pzn", "artikelname", "charge", "verfall", "menge", "menge_retoure", "ek")
        return dict(zip(keys, r))

    def _rb_freigeben(self):
        d = self._rb_selected()
        if not d:
            messagebox.showinfo("Freigeben", "Bitte eine Zeile im Retourenbestand waehlen.", parent=self)
            return
        maxq = d["menge_retoure"]

        def save(v):
            q = max(0, min(int(v["menge"]), maxq))
            if q <= 0:
                return
            with self._conn() as con:
                con.execute(
                    "UPDATE tbl_lagerbestand SET menge=menge+?, menge_retoure=menge_retoure-?, "
                    "aktualisiert_am=? WHERE id=?", (q, q, _heute_iso(), d["id"]))
                con.commit()
            self._log("Retourenbestand", "In Bestand freigegeben", d["id"],
                      f"{q} St Charge {d['charge']} ({d['artikelname']}) -> Normalbestand")
            self._refresh_retourenbestand()
            messagebox.showinfo("Freigegeben",
                                f"{q} St der Charge {d['charge']} sind wieder verkaufbar.", parent=self)

        FormDialog(self, f"Freigeben: {d['artikelname']} (Ch {d['charge']})", [
            ("menge", "Menge freigeben", "int", str(maxq)),
        ], save)

    def _rb_abschreiben(self):
        d = self._rb_selected()
        if not d:
            messagebox.showinfo("Abschreiben", "Bitte eine Zeile im Retourenbestand waehlen.", parent=self)
            return
        maxq = d["menge_retoure"]

        def save(v):
            q = max(0, min(int(v["menge"]), maxq))
            if q <= 0:
                return
            wert = round((d["ek"] or 0) * q, 2)
            with self._conn() as con:
                con.execute("UPDATE tbl_lagerbestand SET menge_retoure=menge_retoure-?, "
                            "aktualisiert_am=? WHERE id=?", (q, _heute_iso(), d["id"]))
                con.execute(
                    """INSERT INTO tbl_gdp_abschreibung
                       (datum,pzn,artikelname,charge,verfall,menge,grund,wert_ek,bearbeiter)
                       VALUES(?,?,?,?,?,?,?,?,?)""",
                    (_heute_iso(), d["pzn"], d["artikelname"], d["charge"], d["verfall"],
                     q, v["grund"], wert, ME()))
                con.commit()
            self._log("Retourenbestand", "Abgeschrieben/vernichtet", d["id"],
                      f"{q} St Charge {d['charge']} ({d['artikelname']}), EK ~ {wert} EUR: {v['grund']}")
            self._refresh_retourenbestand()
            messagebox.showinfo("Abgeschrieben",
                                f"{q} St der Charge {d['charge']} abgeschrieben "
                                f"(EK-Wert ~ {wert} EUR). Im Protokoll dokumentiert.", parent=self)

        FormDialog(self, f"Abschreiben: {d['artikelname']} (Ch {d['charge']})", [
            ("menge", "Menge abschreiben", "int", str(maxq)),
            ("grund", "Grund", "combo",
             ["Beschaedigt", "Verfall ueberschritten", "Temperaturbruch",
              "Nicht GDP-konform", "Rueckruf", "Sonstiges"], "Beschaedigt"),
        ], save)

    def _rb_export_abschreibungen(self):
        path = filedialog.asksaveasfilename(
            parent=self, title="Abschreibungen exportieren", defaultextension=".csv",
            initialfile=f"Abschreibungen_{_heute_iso()}.csv", filetypes=[("CSV", "*.csv")])
        if not path:
            return
        with self._conn() as con:
            rows = con.execute(
                "SELECT datum,pzn,artikelname,charge,verfall,menge,grund,wert_ek,bearbeiter "
                "FROM tbl_gdp_abschreibung ORDER BY id").fetchall()
        with open(path, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f, delimiter=";")
            w.writerow(["Datum", "PZN", "Artikel", "Charge", "Verfall", "Menge", "Grund",
                        "Wert EK", "Bearbeiter"])
            w.writerows(rows)
        messagebox.showinfo("Export", f"{len(rows)} Abschreibung(en) gespeichert.", parent=self)

    # ===================================================== KUNDENQUALIFIZIERUNG
    def _build_qualifizierung(self, parent):
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(1, weight=1)
        bar = tk.Frame(parent, bg=SHELL_BG)
        bar.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        theme.PillButton(bar, "✎  Lizenz / Qualifizierung bearbeiten", self._quali_edit,
                         kind="primary", font_size=10).pack(side="left")
        theme.PillButton(bar, "↻  Aktualisieren", self._refresh_qualifizierung,
                         kind="neutral", font_size=10).pack(side="left", padx=8)
        tk.Label(bar, text="Nur qualifizierte Apotheken mit gueltiger Lizenz duerfen beliefert werden.",
                 bg=SHELL_BG, fg=MUTED, font=(theme.FONT, 9)).pack(side="left", padx=12)

        outer, body = _card(parent)
        outer.grid(row=1, column=0, sticky="nsew")
        cols = ("Kundennr", "Apotheke", "Ort", "Lizenznr", "Lizenz-Typ", "Gueltig bis", "Qualifiziert", "Status")
        self._quali_tree = _make_tree(body, cols, (80, 170, 110, 110, 150, 90, 90, 130),
                                      anchors={"Qualifiziert": "center"})
        self._quali_tree.bind("<Double-1>", lambda _e: self._quali_edit())
        self._refreshers["qualifizierung"] = self._refresh_qualifizierung

    def _refresh_qualifizierung(self):
        t = self._quali_tree
        t.delete(*t.get_children())
        with self._conn() as con:
            if not _table_exists(con, "tbl_kunden_center"):
                t.insert("", "end", values=("-", "Kein Kundenstamm vorhanden", "", "", "", "", "", ""))
                return
            rows = con.execute(
                """SELECT k.kundennummer, k.kundenname, COALESCE(k.ort,''),
                          q.lizenznummer, q.lizenz_typ, q.lizenz_gueltig_bis, q.qualifiziert
                   FROM tbl_kunden_center k
                   LEFT JOIN tbl_gdp_kunde_quali q ON q.kundennummer=k.kundennummer
                   ORDER BY k.kundenname""").fetchall()
        heute = _heute_iso()
        for i, (knr, name, ort, liz, typ, gueltig, quali) in enumerate(rows):
            if quali is None:
                status, tag = "nicht geprueft", "warn"
            elif gueltig and gueltig < heute:
                status, tag = "Lizenz abgelaufen", "alert"
            elif quali:
                status, tag = "freigegeben", "ok"
            else:
                status, tag = "gesperrt", "alert"
            t.insert("", "end", iid=knr,
                     values=(knr, name, ort, liz or "", typ or "", gueltig or "",
                             "Ja" if quali else ("Nein" if quali is not None else "-"), status),
                     tags=(tag,))

    def _quali_edit(self):
        sel = self._quali_tree.selection()
        if not sel:
            messagebox.showinfo("Qualifizierung", "Bitte eine Apotheke waehlen.", parent=self)
            return
        knr = sel[0]
        with self._conn() as con:
            kname = con.execute("SELECT kundenname FROM tbl_kunden_center WHERE kundennummer=?",
                                (knr,)).fetchone()
            kname = kname[0] if kname else knr
            cur = con.execute(
                "SELECT lizenznummer, lizenz_typ, lizenz_gueltig_bis, qualifiziert, bemerkung "
                "FROM tbl_gdp_kunde_quali WHERE kundennummer=?", (knr,)).fetchone()
        liz, typ, gueltig, quali, bem = cur or ("", "Apothekenbetriebserlaubnis", "", 0, "")

        def save(v):
            with self._conn() as con:
                con.execute(
                    """INSERT INTO tbl_gdp_kunde_quali
                       (kundennummer,kunde_name,lizenznummer,lizenz_typ,lizenz_gueltig_bis,
                        qualifiziert,geprueft_am,geprueft_von,bemerkung)
                       VALUES(?,?,?,?,?,?,?,?,?)
                       ON CONFLICT(kundennummer) DO UPDATE SET
                        kunde_name=excluded.kunde_name, lizenznummer=excluded.lizenznummer,
                        lizenz_typ=excluded.lizenz_typ, lizenz_gueltig_bis=excluded.lizenz_gueltig_bis,
                        qualifiziert=excluded.qualifiziert, geprueft_am=excluded.geprueft_am,
                        geprueft_von=excluded.geprueft_von, bemerkung=excluded.bemerkung""",
                    (knr, kname, v["lizenznummer"], v["lizenz_typ"], v["lizenz_gueltig_bis"],
                     v["qualifiziert"], _heute_iso(), ME(), v["bemerkung"]))
                con.commit()
            self._log("Qualifizierung", "Lizenz/Qualifizierung aktualisiert", None,
                      f"{kname}: {'freigegeben' if v['qualifiziert'] else 'gesperrt'}")
            self._refresh_qualifizierung()

        FormDialog(self, f"Qualifizierung: {kname}", [
            ("lizenznummer", "Lizenz-/Erlaubnisnr", "text", liz or ""),
            ("lizenz_typ", "Lizenz-Typ", "combo",
             ["Apothekenbetriebserlaubnis", "Grosshandelserlaubnis (83 AMG)", "Krankenhausapotheke"],
             typ or "Apothekenbetriebserlaubnis"),
            ("lizenz_gueltig_bis", "Gueltig bis (JJJJ-MM-TT)", "text", gueltig or ""),
            ("qualifiziert", "Freigegeben (qualifiziert)", "check", quali or 0),
            ("bemerkung", "Bemerkung", "multiline", bem or ""),
        ], save, width=480)

    # ===================================================== PROTOKOLL
    def _build_protokoll(self, parent):
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(1, weight=1)
        bar = tk.Frame(parent, bg=SHELL_BG)
        bar.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        tk.Label(bar, text="Modul:", bg=SHELL_BG, fg=TEXT, font=(theme.FONT, 10)).pack(side="left")
        self._log_filter = tk.StringVar(value="alle")
        cb = ttk.Combobox(bar, textvariable=self._log_filter, state="readonly", width=20,
                          style="NMG.TCombobox",
                          values=["alle", "Wareneingang", "Rueckruf", "Retoure",
                                  "Qualifizierung"])
        cb.pack(side="left", padx=8)
        cb.bind("<<ComboboxSelected>>", lambda _e: self._refresh_protokoll())
        theme.PillButton(bar, "⬇  Export CSV", self._log_export, kind="neutral",
                         font_size=10).pack(side="left", padx=8)

        outer, body = _card(parent)
        outer.grid(row=1, column=0, sticky="nsew")
        self._log_tree = _make_tree(body, ("Zeitpunkt", "Bearbeiter", "Modul", "Aktion", "Details"),
                                    (150, 110, 120, 180, 260))
        self._refreshers["protokoll"] = self._refresh_protokoll

    def _refresh_protokoll(self):
        t = getattr(self, "_log_tree", None)
        if t is None:
            return
        t.delete(*t.get_children())
        modul = self._log_filter.get()
        with self._conn() as con:
            if modul == "alle":
                rows = con.execute(
                    "SELECT zeitpunkt,bearbeiter,modul,aktion,details FROM tbl_gdp_log "
                    "ORDER BY id DESC LIMIT 500").fetchall()
            else:
                rows = con.execute(
                    "SELECT zeitpunkt,bearbeiter,modul,aktion,details FROM tbl_gdp_log "
                    "WHERE modul=? ORDER BY id DESC LIMIT 500", (modul,)).fetchall()
        for i, r in enumerate(rows):
            t.insert("", "end", values=r, tags=("even" if i % 2 else "odd",))

    def _log_export(self):
        path = filedialog.asksaveasfilename(
            parent=self, title="Protokoll exportieren", defaultextension=".csv",
            initialfile=f"GDP_Protokoll_{_heute_iso()}.csv", filetypes=[("CSV", "*.csv")])
        if not path:
            return
        with self._conn() as con:
            rows = con.execute(
                "SELECT zeitpunkt,bearbeiter,modul,aktion,details FROM tbl_gdp_log ORDER BY id").fetchall()
        with open(path, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f, delimiter=";")
            w.writerow(["Zeitpunkt", "Bearbeiter", "Modul", "Aktion", "Details"])
            w.writerows(rows)
        messagebox.showinfo("Export", f"{len(rows)} Protokoll-Zeile(n) gespeichert.", parent=self)


# ════════════════════════════════════════════════════════════════════════════
def run_standalone():
    """Startet die GDP-App als eigenstaendiges Fenster (eigenes Taskleisten-Icon).
    Wird von start_gdp.py und von NMGone.exe --gdp genutzt."""
    if os.name == "nt":
        try:
            import ctypes
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("NMG.GDP")
        except Exception:
            pass
    try:
        from .migrations import run_migrations
        run_migrations()
    except Exception:
        pass
    root = tk.Tk()
    root.title("NMG Wareneingang & Retouren")
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
    try:
        root.iconbitmap(str(ASSETS_DIR / "GDP.ico"))
    except Exception:
        pass
    GDPPanel(root, on_close=root.destroy).pack(fill="both", expand=True)
    if tour is not None:
        try:
            tour.maybe_show(root, "gdp", _gdp_tour_steps())
        except Exception:
            pass
    root.mainloop()


def _gdp_tour_steps():
    return [
        ("Wareneingang & Retouren",
         "Diese App verbindet den Wareneingang aus der Kasse mit GDP-Qualitaet "
         "und der Retouren-Abwicklung."),
        ("Chargen-Rueckverfolgung",
         "Du siehst zu jeder Charge, welche Apotheke sie erhalten hat - und kannst "
         "im Rueckruffall gezielt genau diese anschreiben."),
        ("Kuehlkette, Meldungen & Inspektion ausgelagert",
         "Temperaturueberwachung, das Meldewesen und die GDP-Selbstinspektion findest "
         "du jetzt uebersichtlich in der eigenen App 'Meldungen'."),
        ("Retouren",
         "Vom Ruecksendegrund ueber die Gutschrift (Faktura) bis zur "
         "Wiedereinlagerung oder Vernichtung - alles mit Charge."),
    ]


if __name__ == "__main__":
    import sys
    sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
    run_standalone()
