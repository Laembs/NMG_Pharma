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
from . import erechnung as erech

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
CARD = theme.CARD
CARD_ALT = theme.CARD_ALT
INK = theme.INK
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
            CREATE TABLE IF NOT EXISTS tbl_gdp_einstellungen(
                schluessel TEXT PRIMARY KEY, wert TEXT
            );
            CREATE TABLE IF NOT EXISTS tbl_gdp_produktionsbestand(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                pzn TEXT, artikelname TEXT, charge TEXT, verfall TEXT,
                menge INTEGER DEFAULT 0, ek REAL, aktualisiert_am TEXT,
                UNIQUE(pzn, charge, verfall)
            );
            CREATE TABLE IF NOT EXISTS tbl_gdp_warenausgang(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                datum TEXT, nummer TEXT, quelle TEXT, ziel TEXT DEFAULT 'Verkaufsbestand',
                status TEXT DEFAULT 'avisiert', bemerkung TEXT,
                erstellt_von TEXT, erstellt_am TEXT DEFAULT CURRENT_TIMESTAMP,
                bestaetigt_am TEXT, bestaetigt_von TEXT, we_id INTEGER
            );
            CREATE TABLE IF NOT EXISTS tbl_gdp_warenausgang_pos(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                wa_id INTEGER, pzn TEXT, artikelname TEXT, charge TEXT, verfall TEXT,
                menge INTEGER DEFAULT 0, ek REAL
            );
            CREATE TABLE IF NOT EXISTS tbl_gdp_bestandsdiff(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                datum TEXT, bereich TEXT, pzn TEXT, artikelname TEXT, charge TEXT, verfall TEXT,
                menge_vorher INTEGER, menge_diff INTEGER, menge_nachher INTEGER,
                grund TEXT, bemerkung TEXT, bearbeiter TEXT, erstellt_am TEXT DEFAULT CURRENT_TIMESTAMP
            );
            """
        )
        # Wareneingangs-Art je Beleg: 'bestand' (Fertigware aus Produktion ->
        # Verkaufsbestand, Standard) oder 'produktion' (Einkauf -> Produktions-
        # bestand). Spalte am gemeinsamen tbl_wareneingang nachruesten.
        if _table_exists(con, "tbl_wareneingang"):
            wcols = {r[1] for r in con.execute("PRAGMA table_info(tbl_wareneingang)")}
            if "art" not in wcols:
                con.execute("ALTER TABLE tbl_wareneingang ADD COLUMN art TEXT DEFAULT 'bestand'")
        # Betriebsmodus dieses Arbeitsplatzes (Verkauf oder Produktion). Standard
        # Verkauf; in den Einstellungen umschaltbar.
        con.execute("INSERT OR IGNORE INTO tbl_gdp_einstellungen(schluessel,wert) VALUES('betriebsmodus','verkauf')")
        # Rechnungsnummer am Warenausgang (optional; per Einstellung als Pflicht setzbar).
        wacols = {r[1] for r in con.execute("PRAGMA table_info(tbl_gdp_warenausgang)")}
        if "rechnungsnummer" not in wacols:
            con.execute("ALTER TABLE tbl_gdp_warenausgang ADD COLUMN rechnungsnummer TEXT")
        con.execute("INSERT OR IGNORE INTO tbl_gdp_einstellungen(schluessel,wert) VALUES('rechnungsnummer_pflicht','0')")
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


class ErechnungPruefDialog(tk.Toplevel):
    """Prüf- und Übernahme-Maske für eine eingelesene eRechnung.

    Der Mitarbeiter gleicht alle Positionen (Menge, EK, Charge/Verfall) gegen
    die tatsächlich gelieferte Ware ab und bucht sie – nach ausdrücklicher
    Bestätigung – als Wareneingang (Produktion) mit Einkaufspreisen.
    Die Zuordnung Lieferanten-Artikelnummer -> deutsche PZN bleibt vorerst
    manuell (Art-Nr/PZN ist je Zeile editierbar)."""

    def __init__(self, panel, daten, fehler):
        super().__init__(panel)
        self.panel = panel
        self.daten = daten
        self.title("eRechnung prüfen & als Wareneingang (Produktion) übernehmen")
        self.configure(bg=SHELL_BG)
        self.transient(panel.winfo_toplevel())
        v = daten.get("verkaeufer", {})
        s = daten.get("summen", {})

        head = tk.Frame(self, bg=ACCENT)
        head.pack(fill="x")
        tk.Label(head, text="eRechnung prüfen & übernehmen", bg=ACCENT, fg="#FFFFFF",
                 font=(theme.FONT, 13, "bold"), padx=16, pady=10).pack(anchor="w")

        info = tk.Frame(self, bg=SHELL_BG)
        info.pack(fill="x", padx=18, pady=(12, 2))
        kopf = (f"Lieferant: {v.get('name', '')}    USt-IdNr.: {v.get('ustid', '') or '—'} "
                f"({v.get('land', '')})\n"
                f"Rechnung: {daten.get('rechnungsnr', '')} ({daten.get('typ', '')})    "
                f"Datum: {daten.get('datum', '')}    Brutto: {s.get('brutto', 0):.2f} €    "
                f"[{daten.get('_format', '')}]")
        tk.Label(info, text=kopf, bg=SHELL_BG, fg=TEXT, justify="left",
                 font=(theme.FONT, 10)).pack(anchor="w")
        if fehler:
            tk.Label(info, text="⚠ eRechnung-Hinweise: " + " | ".join(fehler), bg=SHELL_BG,
                     fg=WARN, justify="left", wraplength=760,
                     font=(theme.FONT, 9)).pack(anchor="w", pady=(4, 0))
        tk.Label(self, text="Mengen/EK gegen die gelieferte Ware prüfen, Charge und Verfall "
                 "erfassen. Art-Nr/PZN ggf. auf die echte PZN korrigieren (Zuordnung noch manuell).",
                 bg=SHELL_BG, fg=MUTED, justify="left", wraplength=770,
                 font=(theme.FONT, 9)).pack(fill="x", padx=18, pady=(2, 6))

        cols = [("Art-Nr/PZN", 12), ("Bezeichnung", 26), ("Menge", 7),
                ("EK €", 9), ("Charge", 12), ("Verfall", 10)]
        wrap = tk.Frame(self, bg=SHELL_BG)
        wrap.pack(fill="both", expand=True, padx=18, pady=4)
        header = tk.Frame(wrap, bg=SHELL_BG)
        header.pack(fill="x")
        for txt, w in cols:
            tk.Label(header, text=txt, bg=SHELL_BG, fg=MUTED, width=w, anchor="w",
                     font=(theme.FONT, 9, "bold")).pack(side="left", padx=2)
        canvas = tk.Canvas(wrap, bg=SHELL_BG, highlightthickness=0, height=260)
        sb = ttk.Scrollbar(wrap, orient="vertical", command=canvas.yview)
        inner = tk.Frame(canvas, bg=SHELL_BG)
        inner.bind("<Configure>", lambda _e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=inner, anchor="nw")
        canvas.configure(yscrollcommand=sb.set)
        canvas.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        self.rows = []
        for p in daten.get("positionen", []):
            r = tk.Frame(inner, bg=SHELL_BG)
            r.pack(fill="x", pady=1)

            def feld(width, val):
                var = tk.StringVar(value=str(val))
                tk.Entry(r, textvariable=var, width=width, font=(theme.FONT, 9),
                         relief="solid", bd=1, highlightthickness=0).pack(side="left", padx=2)
                return var
            vrs = {
                "pzn": feld(12, p.get("pzn", "") or ""),
                "bez": feld(26, p.get("bezeichnung", "") or ""),
                "menge": feld(7, f"{p.get('menge', 0):g}"),
                "ek": feld(9, f"{p.get('einzelpreis', 0):.2f}"),
                "charge": feld(12, ""),
                "verfall": feld(10, ""),
            }
            self.rows.append(vrs)

        foot = tk.Frame(self, bg=SHELL_BG)
        foot.pack(fill="x", padx=18, pady=(6, 16))
        self._gep = tk.IntVar(value=0)
        tk.Checkbutton(foot, text="Geprüft: Positionen mit der gelieferten Ware abgeglichen.",
                       variable=self._gep, bg=SHELL_BG, fg=TEXT, activebackground=SHELL_BG,
                       font=(theme.FONT, 9)).pack(side="left")
        theme.PillButton(foot, "✅ Geprüft – Wareneingang buchen", self._buchen,
                         kind="success", font_size=10, padx=14, pady=7).pack(side="right")
        theme.PillButton(foot, "Abbrechen", self.destroy, kind="neutral",
                         font_size=10, padx=14, pady=7).pack(side="right", padx=(0, 8))
        self.bind("<Escape>", lambda _e: self.destroy())
        self.update_idletasks()
        try:
            self.geometry(f"+{panel.winfo_rootx() + 60}+{panel.winfo_rooty() + 40}")
        except Exception:
            pass
        self.grab_set()

    def _buchen(self):
        if not self._gep.get():
            messagebox.showwarning("Prüfung", "Bitte zuerst die Prüfung bestätigen (Häkchen).",
                                   parent=self)
            return
        pos = []
        for r in self.rows:
            try:
                menge = float(str(r["menge"].get()).replace(",", ".") or 0)
            except ValueError:
                menge = 0.0
            if menge <= 0:
                continue
            try:
                ek = float(str(r["ek"].get()).replace(",", ".") or 0)
            except ValueError:
                ek = 0.0
            pos.append((r["pzn"].get().strip(), r["bez"].get().strip(),
                        r["charge"].get().strip(), r["verfall"].get().strip(), menge, ek))
        if not pos:
            messagebox.showwarning("Wareneingang", "Keine Position mit Menge > 0.", parent=self)
            return
        lieferant = (self.daten.get("verkaeufer") or {}).get("name") or "Lieferant"
        lieferschein = self.daten.get("rechnungsnr") or None
        jetzt = datetime.now().isoformat(timespec="seconds")
        try:
            with self.panel._conn() as con:
                cur = con.execute(
                    "INSERT INTO tbl_wareneingang(datum, lieferant, lieferschein, bearbeiter, art) "
                    "VALUES(?,?,?,?, 'produktion')", (jetzt, lieferant, lieferschein, ME()))
                we_id = cur.lastrowid
                for pzn, name, charge, verfall, menge, ek in pos:
                    con.execute(
                        "INSERT INTO tbl_wareneingang_positionen"
                        "(we_id,pzn,artikelname,charge,verfall,menge,ek) VALUES(?,?,?,?,?,?,?)",
                        (we_id, pzn, name, charge, verfall, menge, ek))
                    upd = con.execute(
                        "UPDATE tbl_gdp_produktionsbestand SET menge=menge+?, ek=COALESCE(?,ek), "
                        "aktualisiert_am=? WHERE pzn=? AND COALESCE(charge,'')=? "
                        "AND COALESCE(verfall,'')=?",
                        (menge, ek, jetzt, pzn, charge, verfall))
                    if upd.rowcount == 0:
                        con.execute(
                            "INSERT INTO tbl_gdp_produktionsbestand"
                            "(pzn,artikelname,charge,verfall,menge,ek,aktualisiert_am) "
                            "VALUES(?,?,?,?,?,?,?)", (pzn, name, charge, verfall, menge, ek, jetzt))
                con.commit()
        except Exception as exc:
            messagebox.showerror("Wareneingang", f"Buchung fehlgeschlagen:\n{exc}", parent=self)
            return
        self.panel._log("Wareneingang", "eRechnung übernommen (Produktion)", we_id,
                        f"{lieferant} · Rg {lieferschein or '–'} · {len(pos)} Pos.")
        try:
            self.panel._refresh_wareneingang()
        except Exception:
            pass
        self.destroy()
        messagebox.showinfo("Übernommen",
                            f"Wareneingang (Produktion) gebucht: {len(pos)} Position(en).\n"
                            "Die GDP-Prüfung dieses Wareneingangs steht noch aus.",
                            parent=self.panel)


# ════════════════════════════════════════════════════════════════════════════
class GDPPanel(tk.Frame):
    """Hauptoberflaeche der GDP-App (Sidebar + Seiten)."""

    NAV = (
        ("uebersicht",      "Uebersicht",            "\U0001F4CA"),
        ("wareneingang",    "Wareneingang",          "\U0001F4E6"),
        ("produktionsbestand", "Produktionsbestand", "\U0001F3ED"),
        ("warenausgang",    "Warenausgang / Avis",   "\U0001F4E4"),
        ("rueckverfolgung", "Chargen-Rueckverfolgung", "\U0001F50E"),
        ("bewegungen",      "Warenbewegungen",       "\U0001F501"),
        ("bestandsdiff",    "Bestandsdifferenzen",   "Δ"),
        ("retouren",        "Retouren / Reklamation", "↩"),
        ("retourenbestand", "Retourenbestand",       "⚖"),
        ("qualifizierung",  "Kundenqualifizierung",  "\U0001F6E1"),
        ("protokoll",       "Protokoll",             "\U0001F4DD"),
        ("einstellungen",   "Einstellungen",         "⚙"),
    )
    TITEL = {k: t for k, t, _ in NAV}
    UNTERTITEL = {
        "uebersicht": "Kennzahlen und offene GDP-Pflichten auf einen Blick.",
        "wareneingang": "Ware annehmen (Charge/Verfall) und GDP-Pruefung erfassen.",
        "produktionsbestand": "Eingekaufte Ware fuer die Produktion - getrennt vom Verkaufsbestand.",
        "warenausgang": "Produzierte Ware avisieren - erscheint zur Bestaetigung im Verkaufs-Wareneingang.",
        "rueckverfolgung": "Welche Apotheke hat welche Charge erhalten? Gezielter Rueckruf.",
        "bewegungen": "Alle Ein- und Ausgaenge aus Verkauf und Produktion - durchsuchbar.",
        "bestandsdiff": "Manuelle Bestandskorrekturen (Inventur, Bruch, Schwund) erfassen und nachvollziehen.",
        "retouren": "Rueckwaren und Reklamationen von der Erfassung bis zur Gutschrift.",
        "retourenbestand": "Zurueckgenommene Ware in Quarantaene: freigeben oder abschreiben.",
        "qualifizierung": "Nur lizenzierte Apotheken duerfen beliefert werden.",
        "protokoll": "Revisionssicheres Protokoll aller GDP-Vorgaenge.",
        "einstellungen": "Module ein-/ausschalten - z. B. die beiden Wareneingaenge.",
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

    # ----------------------------------------------------------- Einstellungen
    def _setting(self, key, default=None):
        with self._conn() as con:
            r = con.execute("SELECT wert FROM tbl_gdp_einstellungen WHERE schluessel=?",
                            (key,)).fetchone()
        return r[0] if r else default

    def _set_setting(self, key, val):
        with self._conn() as con:
            con.execute("INSERT INTO tbl_gdp_einstellungen(schluessel,wert) VALUES(?,?) "
                        "ON CONFLICT(schluessel) DO UPDATE SET wert=excluded.wert",
                        (key, str(val)))
            con.commit()

    def _modus(self):
        """Betriebsmodus dieses Arbeitsplatzes: 'verkauf' oder 'produktion'.
        Strikt getrennt - bestimmt sichtbare Bereiche und aktiven Wareneingang."""
        m = self._setting("betriebsmodus", "verkauf")
        return m if m in ("verkauf", "produktion") else "verkauf"

    def _we_bestand_aktiv(self):
        return self._modus() == "verkauf"

    def _we_produktion_aktiv(self):
        return self._modus() == "produktion"

    def _rechnung_pflicht(self):
        return self._setting("rechnungsnummer_pflicht", "0") == "1"

    def _next_nummer(self, con, prefix):
        jahr = HEUTE().year
        n = con.execute("SELECT COUNT(*) FROM tbl_gdp_warenausgang WHERE nummer LIKE ?",
                        (f"{prefix}-{jahr}-%",)).fetchone()[0]
        return f"{prefix}-{jahr}-{1001 + n}"

    # ------------------------------------------------------------------- Aufbau
    def _build(self):
        self.columnconfigure(1, weight=1)
        self.rowconfigure(0, weight=1)

        # ---- Sidebar (zentrale theme.Sidebar) ----
        self.sidebar = theme.Sidebar(self, width=256, title="Wareneingang",
                                     subtitle="& Retouren")
        self.sidebar.grid(row=0, column=0, sticky="ns")
        self._app_icon = theme.load_icon(ASSETS_DIR / "GDP.ico", 56)
        if self._app_icon:
            self.sidebar.set_logo(self._app_icon)
        # Modus-Badge oberhalb der Eintraege (in den Navigations-Container).
        self._modus_badge = tk.Label(self.sidebar.body(), text="", font=(theme.FONT, 9, "bold"),
                                     fg="#FFFFFF", bg=SIDEBAR_ACTIVE, padx=8, pady=3)
        self._modus_badge.pack(anchor="w", padx=22, pady=(4, 8))
        self._nav_rowf = {}
        for key, text, icon in self.NAV:
            self._nav_rowf[key] = self.sidebar.add_item(
                key, icon, text, lambda k=key: self._show_view(k))

        foot = self.sidebar.footer()
        tk.Button(foot, text="\U0001F3E0  NMGone oeffnen", command=self._open_nmgone,
                  bg=SIDEBAR_ACTIVE, fg="#FFFFFF", relief="flat",
                  font=(theme.FONT, 10, "bold"), activebackground="#1B5085",
                  activeforeground="#FFFFFF", padx=10, pady=7, cursor="hand2").pack(fill="x", padx=10, pady=(0, 6))
        tk.Button(foot, text="Schliessen", command=self._on_close, relief="flat",
                  bg="#0E3454", fg=SIDEBAR_TEXT, activebackground="#15466E",
                  activeforeground="#FFFFFF", padx=10, pady=5, cursor="hand2").pack(fill="x", padx=10)
        self.sidebar.add_footer_note(f"Datenbank:\n{Path(self.db_path).name}")

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
            "produktionsbestand": self._build_produktionsbestand,
            "warenausgang": self._build_warenausgang,
            "rueckverfolgung": self._build_rueckverfolgung,
            "bewegungen": self._build_bewegungen,
            "bestandsdiff": self._build_bestandsdiff,
            "retouren": self._build_retouren,
            "retourenbestand": self._build_retourenbestand,
            "qualifizierung": self._build_qualifizierung,
            "protokoll": self._build_protokoll,
            "einstellungen": self._build_einstellungen,
        }
        self._refreshers = {}
        for key, builder in self._builders.items():
            frame = tk.Frame(page, bg=SHELL_BG)
            frame.grid(row=0, column=0, sticky="nsew")
            builder(frame)
            self._views[key] = frame

        self._apply_nav_visibility()
        self._current = None
        self._show_view("uebersicht")

    # Bereiche, die nur in einem Betriebsmodus sichtbar sind
    NAV_PRODUKTION = ("produktionsbestand", "warenausgang")
    NAV_VERKAUF = ("retouren", "retourenbestand", "qualifizierung")

    def _visible_nav_keys(self):
        produktion = self._we_produktion_aktiv()
        verkauf = self._we_bestand_aktiv()
        keys = set()
        for key, _t, _i in self.NAV:
            if key in self.NAV_PRODUKTION:
                if produktion:
                    keys.add(key)
            elif key in self.NAV_VERKAUF:
                if verkauf:
                    keys.add(key)
            else:
                keys.add(key)
        return keys

    def _apply_nav_visibility(self):
        """Blendet Nav-Eintraege passend zum Betriebsmodus ein/aus. Re-packt in
        NAV-Reihenfolge, damit die Sortierung erhalten bleibt."""
        visible = self._visible_nav_keys()
        for key, _t, _i in self.NAV:
            self._nav_rowf[key].pack_forget()
        for key, _t, _i in self.NAV:
            if key in visible:
                self._nav_rowf[key].pack(fill="x", padx=10, pady=1)
        if getattr(self, "_modus_badge", None) is not None:
            m = self._modus()
            self._modus_badge.config(
                text=f"Modus: {self.MODUS_INFO[m][0].strip()}",
                bg=(OK_GREEN if m == "verkauf" else "#7A4E12"))
        cur = getattr(self, "_current", None)
        if cur is not None and cur not in visible:
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

        outer_a, self._alert_body = _card(body, "Offene Pflichten", "Was als naechstes ansteht")
        outer_a.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        outer_b, self._quick_body = _card(body, "Schnellzugriff")
        outer_b.grid(row=0, column=1, sticky="nsew", padx=(8, 0))
        self._refreshers["uebersicht"] = self._refresh_uebersicht

    def _refresh_quickaccess(self):
        for w in self._quick_body.winfo_children():
            w.destroy()
        visible = self._visible_nav_keys()
        for key, label, icon in self.NAV[1:]:
            if key in visible and key != "einstellungen":
                theme.PillButton(self._quick_body, f"{icon}  {label}",
                                 lambda k=key: self._show_view(k),
                                 kind="ghost", font_size=10, padx=12, pady=8).pack(fill="x", pady=3)

    def _kpi_tile(self, parent, value, label, color):
        f = tk.Frame(parent, bg=BG, highlightbackground=BORDER, highlightthickness=1)
        tk.Frame(f, bg=color, height=4).pack(fill="x")
        tk.Label(f, text=str(value), bg=BG, fg=color,
                 font=(theme.FONT, 26, "bold")).pack(pady=(12, 0), padx=18)
        tk.Label(f, text=label, bg=BG, fg=MUTED, font=(theme.FONT, 9),
                 wraplength=150, justify="center").pack(pady=(0, 12), padx=12)
        return f

    def _refresh_uebersicht(self):
        self._refresh_quickaccess()
        verkauf = self._we_bestand_aktiv()
        with self._conn() as con:
            ret_offen = con.execute(
                "SELECT COUNT(*) FROM tbl_gdp_retoure WHERE status IN ('Neu','In Pruefung')").fetchone()[0]
            liz_abgelaufen = con.execute(
                "SELECT COUNT(*) FROM tbl_gdp_kunde_quali WHERE lizenz_gueltig_bis IS NOT NULL "
                "AND lizenz_gueltig_bis < ?", (_heute_iso(),)).fetchone()[0]
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
            rb_offen = rb_stueck = 0
            if _table_exists(con, "tbl_lagerbestand"):
                lcols = {r[1] for r in con.execute("PRAGMA table_info(tbl_lagerbestand)")}
                if "menge_retoure" in lcols:
                    rb_offen, rb_stueck = con.execute(
                        "SELECT COUNT(*), COALESCE(SUM(menge_retoure),0) FROM tbl_lagerbestand "
                        "WHERE COALESCE(menge_retoure,0) > 0").fetchone()
            avis_offen = 0
            if _table_exists(con, "tbl_gdp_warenausgang"):
                avis_offen = con.execute(
                    "SELECT COUNT(*) FROM tbl_gdp_warenausgang WHERE status='avisiert'").fetchone()[0]
            pb_chargen = pb_stueck = 0
            if _table_exists(con, "tbl_gdp_produktionsbestand"):
                pb_chargen, pb_stueck = con.execute(
                    "SELECT COUNT(*), COALESCE(SUM(menge),0) FROM tbl_gdp_produktionsbestand "
                    "WHERE menge>0").fetchone()
        grenze = HEUTE() + timedelta(days=90)
        bald_ab = sum(1 for *_ , verf, menge in chargen
                      if (_parse_verfall(verf) and _parse_verfall(verf) <= grenze))

        for w in self._kpi_row.winfo_children():
            w.destroy()
        if verkauf:
            tiles = [
                (ret_offen, "Offene Retouren", WARN if ret_offen else OK_GREEN),
                (rb_offen, "Retourenbestand offen", WARN if rb_offen else OK_GREEN),
                (avis_offen, "Avise zu bestaetigen", WARN if avis_offen else OK_GREEN),
                (liz_abgelaufen, "Lizenzen abgelaufen", DANGER if liz_abgelaufen else OK_GREEN),
                (bald_ab, "Chargen Verfall < 90 Tage", WARN if bald_ab else OK_GREEN),
            ]
        else:
            tiles = [
                (we_ungeprueft, "Wareneingaenge ungeprueft", WARN if we_ungeprueft else OK_GREEN),
                (pb_chargen, "Produktionsbestand (Chargen)", ACCENT),
                (pb_stueck, "Produktionsbestand (Stueck)", ACCENT),
                (avis_offen, "Warenausgaenge avisiert", ACCENT if avis_offen else OK_GREEN),
                (bald_ab, "Chargen Verfall < 90 Tage", WARN if bald_ab else OK_GREEN),
            ]
        for i, (v, lbl, col) in enumerate(tiles):
            self._kpi_row.columnconfigure(i, weight=1)
            self._kpi_tile(self._kpi_row, v, lbl, col).grid(row=0, column=i, sticky="nsew", padx=4)

        for w in self._alert_body.winfo_children():
            w.destroy()
        alerts = []
        if we_ungeprueft:
            alerts.append((WARN, f"{we_ungeprueft} Wareneingang/-gaenge ohne GDP-Pruefung.", "wareneingang"))
        if verkauf:
            if liz_abgelaufen:
                alerts.append((DANGER, f"{liz_abgelaufen} Apotheke(n) mit abgelaufener Lizenz - Belieferung sperren.", "qualifizierung"))
            if avis_offen:
                alerts.append((WARN, f"{avis_offen} avisierte Lieferung(en) aus der Produktion - im "
                                     "Wareneingang bestaetigen.", "wareneingang"))
            if ret_offen:
                alerts.append((WARN, f"{ret_offen} Retoure(n)/Reklamation(en) in Bearbeitung.", "retouren"))
            if rb_offen:
                alerts.append((WARN, f"{rb_offen} Charge(n) im Retourenbestand ({rb_stueck} St) - "
                                     "freigeben oder abschreiben.", "retourenbestand"))
        else:
            if pb_chargen:
                alerts.append((ACCENT, f"{pb_chargen} Charge(n) im Produktionsbestand ({pb_stueck} St) - "
                                       "produzieren und avisieren.", "produktionsbestand"))
            if avis_offen:
                alerts.append((WARN, f"{avis_offen} Warenausgang/-gaenge avisiert (warten auf "
                                     "Bestaetigung im Verkauf).", "warenausgang"))
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

    def _erechnung_einlesen(self):
        """Liest eine eingehende eRechnung (ZUGFeRD/Factur-X oder XRechnung) ein
        und öffnet die Prüf-/Übernahme-Maske, aus der – nach Bestätigung durch
        den Mitarbeiter – ein Wareneingang (Produktion) mit Einkaufspreisen
        gebucht werden kann."""
        pfad = filedialog.askopenfilename(
            title="eRechnung einlesen (XML oder ZUGFeRD-PDF)",
            filetypes=[("eRechnung", "*.xml *.pdf"), ("XML", "*.xml"),
                       ("PDF (ZUGFeRD)", "*.pdf"), ("Alle Dateien", "*.*")])
        if not pfad:
            return
        try:
            daten = erech.lies_erechnung(pfad)
        except Exception as exc:
            messagebox.showerror("eRechnung", f"Konnte nicht gelesen werden:\n{exc}\n\n"
                                 "Hinweis: Aus komprimierten ZUGFeRD-PDFs gelingt das Auslesen "
                                 "nur mit installierter Bibliothek 'facturx'.")
            return
        if not daten.get("positionen"):
            messagebox.showwarning("eRechnung", "Die eRechnung enthält keine Positionen.")
            return
        fehler = erech.pruefe_en16931(daten)
        ErechnungPruefDialog(self, daten, fehler)

    def _build_wareneingang(self, parent):
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(3, weight=1)
        bar = tk.Frame(parent, bg=SHELL_BG)
        bar.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        theme.PillButton(bar, "✅  GDP-Pruefung erfassen", self._we_pruefung_dialog,
                         kind="primary", font_size=10).pack(side="left")
        theme.PillButton(bar, "↻  Aktualisieren", self._refresh_wareneingang,
                         kind="neutral", font_size=10).pack(side="left", padx=8)
        theme.PillButton(bar, "🧾  eRechnung einlesen", self._erechnung_einlesen,
                         kind="neutral", font_size=10).pack(side="left", padx=(0, 8))
        tk.Label(bar, text="History:", bg=SHELL_BG, fg=MUTED,
                 font=(theme.FONT, 9)).pack(side="left", padx=(8, 4))
        self._we_filter = tk.StringVar(value="Alle")
        cb = ttk.Combobox(bar, textvariable=self._we_filter, state="readonly", width=22,
                          style="NMG.TCombobox",
                          values=["Alle", "Offen (ungeprueft)", "Erledigt (GDP-geprueft)"])
        cb.pack(side="left")
        cb.bind("<<ComboboxSelected>>", lambda _e: self._refresh_wareneingang())
        tk.Label(bar, text="Typ:", bg=SHELL_BG, fg=MUTED,
                 font=(theme.FONT, 9)).pack(side="left", padx=(8, 4))
        self._we_typ_filter = tk.StringVar(value="Alle")
        cbt = ttk.Combobox(bar, textvariable=self._we_typ_filter, state="readonly", width=18,
                           style="NMG.TCombobox",
                           values=["Alle", "Verkaufsbestand", "Produktion"])
        cbt.pack(side="left")
        cbt.bind("<<ComboboxSelected>>", lambda _e: self._refresh_wareneingang())
        self._we_count = tk.Label(bar, text="", bg=SHELL_BG, fg=MUTED, font=(theme.FONT, 9))
        self._we_count.pack(side="left", padx=12)

        # ---- Buchungsmaske: Ware mit Charge/Verfall annehmen ----
        outer_b, formb = _card(parent, "Wareneingang erfassen", "Ware annehmen")
        outer_b.grid(row=1, column=0, sticky="ew", pady=(0, 8))
        # Art des Wareneingangs (nur sichtbar bei mehreren aktiven Arten)
        self._we_art = tk.StringVar(value="bestand")
        self._we_art_frame = tk.Frame(formb, bg=BG)
        self._we_art_frame.pack(fill="x", pady=(0, 4))
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

        # ---- Avisierte Lieferungen aus der Produktion (zu bestaetigen) ----
        # Erscheinen hier automatisch, sobald in der Produktion ein Warenausgang
        # erzeugt wurde. Werden bei Anlieferung bestaetigt (haendisch / per Liste).
        self._we_avis_outer = tk.Frame(parent, bg="#FBE9C7", highlightbackground=WARN,
                                       highlightthickness=1)
        self._we_avis_outer.grid(row=2, column=0, sticky="ew", pady=(0, 8))
        avhead = tk.Frame(self._we_avis_outer, bg="#FBE9C7")
        avhead.pack(fill="x", padx=12, pady=(8, 4))
        tk.Label(avhead, text="🕓  Avisierte Lieferungen (zu bestaetigen)", bg="#FBE9C7",
                 fg="#8A5A00", font=(theme.FONT, 11, "bold")).pack(side="left")
        theme.PillButton(avhead, "✅  Bestaetigen", self._we_avis_confirm,
                         kind="success", font_size=9, padx=12, pady=4).pack(side="right")
        theme.PillButton(avhead, "📥  Per Liste bestaetigen", self._we_avis_confirm_liste,
                         kind="neutral", font_size=9, padx=12, pady=4).pack(side="right", padx=6)
        avbody = tk.Frame(self._we_avis_outer, bg="#FBE9C7")
        avbody.pack(fill="x", padx=12, pady=(0, 10))
        self._we_avis_tree = _make_tree(avbody, ("Nummer", "Datum", "Quelle", "Rechnung", "Positionen", "Stueck"),
                                        (115, 85, 130, 120, 85, 65), height=4,
                                        anchors={"Positionen": "center", "Stueck": "center"})
        self._we_avis_tree.bind("<Double-1>",
                                lambda _e: self._wa_show_positions(self._we_avis_selected()))

        # ---- Liste der erfassten Wareneingaenge + GDP-Pruefstatus ----
        outer, body = _card(parent)
        outer.grid(row=3, column=0, sticky="nsew")
        cols = ("Datum", "Art", "Lieferant", "Lieferschein", "Positionen", "Stueck", "GDP-Pruefung")
        self._we_tree = _make_tree(body, cols, (90, 130, 140, 120, 75, 65, 140),
                                   anchors={"Positionen": "center", "Stueck": "center"})
        self._we_tree.bind("<Double-1>", lambda _e: self._we_show_positions())
        self._refreshers["wareneingang"] = self._refresh_wareneingang
        self._we_sync_art_selector()

    def _we_avis_selected(self):
        sel = self._we_avis_tree.selection()
        return int(sel[0]) if sel else None

    def _we_avis_confirm(self):
        wa_id = self._we_avis_selected()
        if not wa_id:
            messagebox.showinfo("Bestaetigen", "Bitte eine avisierte Lieferung waehlen.", parent=self)
            return
        self._wa_confirm(wa_id)

    def _we_avis_confirm_liste(self):
        wa_id = self._we_avis_selected()
        if not wa_id:
            messagebox.showinfo("Per Liste bestaetigen",
                                "Bitte die zugehoerige avisierte Lieferung waehlen.", parent=self)
            return
        path = filedialog.askopenfilename(
            parent=self, title="Anlieferliste waehlen (Excel, CSV oder TXT)",
            filetypes=[("Tabellen", "*.xlsx *.xlsm *.csv *.txt"), ("Alle Dateien", "*.*")])
        if not path:
            return
        try:
            positions, rechnung = self._parse_lieferliste(path)
        except Exception as e:
            messagebox.showerror("Import", f"Datei konnte nicht gelesen werden:\n{e}", parent=self)
            return
        if not positions:
            messagebox.showinfo("Per Liste bestaetigen",
                                "Keine gueltigen Positionen (PZN + Menge) gefunden.", parent=self)
            return
        # Soll/Ist-Abgleich: avisierte Positionen vs. tatsaechliche Anlieferliste
        with self._conn() as con:
            avis = con.execute(
                "SELECT pzn, artikelname, charge, verfall, menge, ek FROM tbl_gdp_warenausgang_pos "
                "WHERE wa_id=?", (wa_id,)).fetchall()
        diffs = self._avis_diff(avis, positions)
        if diffs:
            mehr = "" if len(diffs) <= 15 else f"\n… und {len(diffs) - 15} weitere"
            if not messagebox.askyesno(
                    "Abweichung Avis ↔ Anlieferung",
                    "Die Anlieferliste weicht vom avisierten Warenausgang ab:\n\n"
                    + "\n".join(f"• {d}" for d in diffs[:15]) + mehr
                    + "\n\nTatsaechlich gelieferte Menge buchen und den Avis trotzdem bestaetigen?",
                    icon="warning", parent=self):
                return
        else:
            if not messagebox.askyesno("Per Liste bestaetigen",
                    f"Anlieferliste stimmt mit dem Avis ueberein.\n\n{len(positions)} Position(en) "
                    "in den Verkaufsbestand buchen und den Avis bestaetigen?", parent=self):
                return
        # Rechnungsnummer aus der Anlieferliste am Avis nachtragen (falls vorhanden)
        if rechnung:
            with self._conn() as con:
                con.execute("UPDATE tbl_gdp_warenausgang SET rechnungsnummer=? WHERE id=? "
                            "AND COALESCE(rechnungsnummer,'')=''", (rechnung, wa_id))
                con.commit()
        if self._wa_confirm(wa_id, override_positions=positions) and diffs:
            self._log("Wareneingang", "Avis mit Abweichung bestaetigt", wa_id, " | ".join(diffs))

    @staticmethod
    def _avis_diff(avis_pos, liste_pos):
        """Vergleicht avisierte Positionen (Soll) mit der Anlieferliste (Ist).
        Schluessel = (PZN, Charge). Liefert Liste von Abweichungstexten."""
        def agg(rows):
            d = {}
            for pzn, name, charge, verfall, menge, ek in rows:
                key = ((pzn or "").strip(), (charge or "").strip())
                m, nm = d.get(key, (0, name))
                d[key] = (m + (menge or 0), name or nm)
            return d
        soll, ist = agg(avis_pos), agg(liste_pos)
        diffs = []
        for key in sorted(set(soll) | set(ist)):
            pzn, charge = key
            soll_m, name_s = soll.get(key, (0, ""))
            ist_m, name_i = ist.get(key, (0, ""))
            name = name_s or name_i or pzn
            ch = f" Ch {charge}" if charge else " (ohne Charge)"
            if soll_m and not ist_m:
                diffs.append(f"FEHLT: {name}{ch} – avisiert {soll_m}, nicht in Liste")
            elif ist_m and not soll_m:
                diffs.append(f"ZUSAETZLICH: {name}{ch} – {ist_m} geliefert, nicht avisiert")
            elif soll_m != ist_m:
                diffs.append(f"MENGE: {name}{ch} – avisiert {soll_m}, geliefert {ist_m}")
        return diffs

    WE_ART_LABEL = {"bestand": "→ Verkaufsbestand", "produktion": "→ Produktion"}

    def _we_sync_art_selector(self):
        """(Re)baut den Art-Umschalter passend zu den aktiven Einstellungen.
        Beide Arten aus -> Hinweis statt Maske; eine Art -> fester Hinweis;
        beide -> Auswahl per Radiobutton."""
        f = getattr(self, "_we_art_frame", None)
        if f is None:
            return
        for w in f.winfo_children():
            w.destroy()
        opts = []
        if self._we_bestand_aktiv():
            opts.append(("bestand", "Aus Produktion → Verkaufsbestand"))
        if self._we_produktion_aktiv():
            opts.append(("produktion", "Einkauf → für Produktion (Rohware)"))
        if not opts:
            self._we_art.set("")
            tk.Label(f, text="⚠  Beide Wareneingaenge sind in den Einstellungen deaktiviert.",
                     bg=BG, fg=DANGER, font=(theme.FONT, 9, "bold")).pack(anchor="w")
            return
        keys = [k for k, _ in opts]
        if self._we_art.get() not in keys:
            self._we_art.set(keys[0])
        if len(opts) == 1:
            tk.Label(f, text="Wareneingang: " + opts[0][1], bg=BG, fg=ACCENT,
                     font=(theme.FONT, 9, "bold")).pack(anchor="w")
        else:
            head = tk.Frame(f, bg=BG)
            head.pack(fill="x")
            tk.Label(head, text="Art des Wareneingangs:", bg=BG, fg=MUTED,
                     font=(theme.FONT, 9, "bold")).pack(side="left", padx=(0, 10))
            for k, lbl in opts:
                tk.Radiobutton(head, text=lbl, variable=self._we_art, value=k, bg=BG,
                               activebackground=BG, selectcolor=BG, font=(theme.FONT, 9),
                               cursor="hand2").pack(side="left", padx=(0, 12))

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
        art = self._we_art.get() if hasattr(self, "_we_art") else "bestand"
        if not art:
            messagebox.showwarning("Wareneingang",
                "Beide Wareneingaenge sind deaktiviert. Bitte in den Einstellungen "
                "mindestens einen aktivieren.", parent=self)
            return
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
                "INSERT INTO tbl_wareneingang(datum, lieferant, lieferschein, bearbeiter, art) "
                "VALUES(?,?,?,?,?)",
                (jetzt, lieferant, lieferschein or None, ME(), art))
            we_id = cur.lastrowid
            con.execute(
                "INSERT INTO tbl_wareneingang_positionen(we_id,pzn,artikelname,charge,verfall,menge,ek) "
                "VALUES(?,?,?,?,?,?,?)", (we_id, pzn, artikelname, charge, verfall, menge, ek))
            # Ziel-Bestand je nach Art: Produktions-Einkauf -> getrennter
            # Produktionsbestand (nicht verkaufbar); sonst Verkaufsbestand.
            ziel = "tbl_gdp_produktionsbestand" if art == "produktion" else "tbl_lagerbestand"
            upd = con.execute(
                f"UPDATE {ziel} SET menge = menge + ?, ek=COALESCE(?, ek), aktualisiert_am = ? "
                "WHERE pzn=? AND COALESCE(charge,'')=? AND COALESCE(verfall,'')=?",
                (menge, ek, jetzt, pzn, charge, verfall))
            if upd.rowcount == 0:
                con.execute(
                    f"INSERT INTO {ziel}(pzn,artikelname,charge,verfall,menge,ek,aktualisiert_am) "
                    "VALUES(?,?,?,?,?,?,?)", (pzn, artikelname, charge, verfall, menge, ek, jetzt))
            con.commit()
        ziel_lbl = "Produktionsbestand" if art == "produktion" else "Verkaufsbestand"
        self._log("Wareneingang", f"Einbuchen ({ziel_lbl})", we_id,
                  f"{artikelname} (PZN {pzn}) +{menge} · Charge {charge or '–'} · Verf {verfall or '–'}")
        self._we_pzn = None
        self._we_search.clear()
        self._we_artikel_label.config(text="Kein Artikel gewaehlt.", fg=MUTED)
        for var in (self._we_charge_var, self._we_verfall_var, self._we_menge_var,
                    self._we_ek_var, self._we_ls_var):
            var.set("")
        self._refresh_wareneingang()
        if art == "produktion" and self._current != "produktionsbestand":
            self._refreshers.get("produktionsbestand", lambda: None)()
        messagebox.showinfo("Wareneingang",
                            f"{menge} × {artikelname} in den {ziel_lbl} eingebucht.", parent=self)

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
        flt = self._we_filter.get() if hasattr(self, "_we_filter") else "Alle"
        with self._conn() as con:
            if not _table_exists(con, "tbl_wareneingang"):
                t.insert("", "end", values=("-", "Noch keine Wareneingaenge erfasst", "", "", "", ""))
                if hasattr(self, "_we_count"):
                    self._we_count.config(text="")
                return
            has_art = "art" in {r[1] for r in con.execute("PRAGMA table_info(tbl_wareneingang)")}
            art_sel = "w.art" if has_art else "'bestand'"
            rows = con.execute(
                f"""SELECT w.id, w.datum, COALESCE({art_sel},'bestand'), w.lieferant, w.lieferschein,
                          (SELECT COUNT(*) FROM tbl_wareneingang_positionen p WHERE p.we_id=w.id),
                          (SELECT COALESCE(SUM(menge),0) FROM tbl_wareneingang_positionen p WHERE p.we_id=w.id),
                          (SELECT gdp_konform FROM tbl_gdp_we_pruefung g WHERE g.we_id=w.id ORDER BY g.id DESC LIMIT 1)
                   FROM tbl_wareneingang w ORDER BY w.datum DESC, w.id DESC""").fetchall()
        typ_flt = self._we_typ_filter.get() if hasattr(self, "_we_typ_filter") else "Alle"
        gezeigt = offen_n = erledigt_n = 0
        for i, (wid, datum, art, lief, ls, npos, stk, konform) in enumerate(rows):
            if typ_flt == "Verkaufsbestand" and art != "bestand":
                continue
            if typ_flt == "Produktion" and art != "produktion":
                continue
            ist_erledigt = konform is not None
            if ist_erledigt:
                erledigt_n += 1
            else:
                offen_n += 1
            if flt.startswith("Offen") and ist_erledigt:
                continue
            if flt.startswith("Erledigt") and not ist_erledigt:
                continue
            if konform is None:
                pruef, tag = "offen", "warn"
            elif konform:
                pruef, tag = "✔ GDP-konform", "ok"
            else:
                pruef, tag = "✖ nicht konform", "alert"
            base = "even" if gezeigt % 2 else "odd"
            art_disp = self.WE_ART_LABEL.get(art, art)
            t.insert("", "end", iid=str(wid),
                     values=(datum or "", art_disp, lief or "", ls or "", npos, stk, pruef),
                     tags=(base if tag in ("ok",) else tag,))
            gezeigt += 1
        if hasattr(self, "_we_count"):
            self._we_count.config(
                text=f"{gezeigt} angezeigt · {offen_n} offen · {erledigt_n} erledigt")
        self._refresh_we_avis()

    def _refresh_we_avis(self):
        """Avisierte Warenausgaenge (Ziel Verkaufsbestand) im Wareneingang zeigen.
        Karte nur sichtbar, wenn es offene Avise gibt UND Produktion aktiv ist."""
        t = getattr(self, "_we_avis_tree", None)
        if t is None:
            return
        t.delete(*t.get_children())
        rows = []
        if self._we_bestand_aktiv():  # Avise werden im Verkauf-Modus bestaetigt
            with self._conn() as con:
                if _table_exists(con, "tbl_gdp_warenausgang"):
                    rows = con.execute(
                        """SELECT w.id, w.nummer, w.datum, w.quelle, COALESCE(w.rechnungsnummer,''),
                                  (SELECT COUNT(*) FROM tbl_gdp_warenausgang_pos p WHERE p.wa_id=w.id),
                                  (SELECT COALESCE(SUM(menge),0) FROM tbl_gdp_warenausgang_pos p WHERE p.wa_id=w.id)
                           FROM tbl_gdp_warenausgang w
                           WHERE w.status='avisiert' AND w.ziel='Verkaufsbestand'
                           ORDER BY w.id DESC""").fetchall()
        for wid, nummer, datum, quelle, rnr, npos, stk in rows:
            t.insert("", "end", iid=str(wid),
                     values=(nummer or "", (datum or "")[:10], quelle or "", rnr, npos, stk),
                     tags=("warn",))
        if rows:
            self._we_avis_outer.grid()
        else:
            self._we_avis_outer.grid_remove()

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

    # ===================================================== PRODUKTIONSBESTAND
    def _build_produktionsbestand(self, parent):
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(2, weight=1)
        bar = tk.Frame(parent, bg=SHELL_BG)
        bar.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        theme.PillButton(bar, "\U0001F3ED  Produzieren → Warenausgang", self._pb_produzieren,
                         kind="success", font_size=10).pack(side="left")
        theme.PillButton(bar, "\U0001F5D1  Abschreiben", self._pb_abschreiben,
                         kind="danger", font_size=10).pack(side="left", padx=6)
        theme.PillButton(bar, "↻  Aktualisieren", self._refresh_produktionsbestand,
                         kind="neutral", font_size=10).pack(side="left", padx=6)
        tk.Label(bar, text="Eingekaufte Ware fuer die Produktion (gesperrt). 'Produzieren' "
                           "erzeugt einen Warenausgang, der im Verkaufs-Wareneingang avisiert wird.",
                 bg=SHELL_BG, fg=MUTED, font=(theme.FONT, 9)).pack(side="left", padx=12)

        self._pb_summary = tk.Label(parent, text="", bg=SHELL_BG, fg=TEXT,
                                    font=(theme.FONT, 10, "bold"))
        self._pb_summary.grid(row=1, column=0, sticky="w", pady=(0, 6))

        outer, body = _card(parent, "Produktionsbestand", "Rohware / Einkauf fuer die Produktion")
        outer.grid(row=2, column=0, sticky="nsew")
        self._pb_tree = _make_tree(
            body, ("PZN", "Artikel", "Charge", "Verfall", "Menge", "EK"),
            (90, 250, 110, 90, 70, 80), anchors={"Menge": "center", "EK": "e"})
        self._refreshers["produktionsbestand"] = self._refresh_produktionsbestand

    def _refresh_produktionsbestand(self):
        t = getattr(self, "_pb_tree", None)
        if t is None:
            return
        t.delete(*t.get_children())
        ges = wert = 0
        with self._conn() as con:
            rows = con.execute(
                "SELECT id, pzn, artikelname, charge, verfall, menge, ek FROM tbl_gdp_produktionsbestand "
                "WHERE menge > 0 ORDER BY artikelname").fetchall()
        grenze = HEUTE() + timedelta(days=90)
        for i, (pid, pzn, name, charge, verf, menge, ek) in enumerate(rows):
            ges += menge
            wert += (ek or 0) * menge
            d = _parse_verfall(verf)
            tag = "alert" if (d and d <= grenze) else ("even" if i % 2 else "odd")
            t.insert("", "end", iid=str(pid),
                     values=(pzn or "", name or "", charge or "", verf or "", menge,
                             f"{ek:.2f}" if ek else ""), tags=(tag,))
        self._pb_summary.config(
            text=f"Produktionsbestand: {len(rows)} Charge(n) · {ges} Stueck · EK-Wert ~ {wert:.2f} EUR")

    def _pb_selected(self):
        sel = self._pb_tree.selection()
        if not sel:
            return None
        with self._conn() as con:
            r = con.execute(
                "SELECT id, pzn, artikelname, charge, verfall, menge, ek FROM tbl_gdp_produktionsbestand "
                "WHERE id=?", (int(sel[0]),)).fetchone()
        if not r:
            return None
        return dict(zip(("id", "pzn", "artikelname", "charge", "verfall", "menge", "ek"), r))

    def _pb_produzieren(self):
        """Produktion fertig: reduziert den Produktionsbestand und legt einen
        avisierten Warenausgang an (Ziel Verkaufsbestand). Der Warenausgang
        erscheint danach automatisch im Verkaufs-Wareneingang zur Bestaetigung."""
        d = self._pb_selected()
        if not d:
            messagebox.showinfo("Produzieren", "Bitte eine Zeile im Produktionsbestand waehlen.", parent=self)
            return
        maxq = d["menge"]

        def save(v):
            q = max(0, min(int(v["menge"]), maxq))
            if q <= 0:
                return
            rnr = v.get("rechnungsnummer", "").strip()
            if self._rechnung_pflicht() and not rnr:
                raise ValueError("Rechnungsnummer ist als Pflichtfeld gesetzt (Einstellungen).")
            charge = v["charge"].strip() or d["charge"]
            verfall = _normalize_verfall(v["verfall"]) if v["verfall"].strip() else d["verfall"]
            jetzt = datetime.now().isoformat(timespec="seconds")
            with self._conn() as con:
                con.execute("UPDATE tbl_gdp_produktionsbestand SET menge=menge-?, aktualisiert_am=? WHERE id=?",
                            (q, jetzt, d["id"]))
                nummer = self._next_nummer(con, "WA")
                cur = con.execute(
                    """INSERT INTO tbl_gdp_warenausgang
                       (datum,nummer,quelle,ziel,status,bemerkung,rechnungsnummer,erstellt_von)
                       VALUES(?,?, 'Produktion', 'Verkaufsbestand', 'avisiert', ?,?, ?)""",
                    (jetzt, nummer, v.get("bemerkung", ""), rnr, ME()))
                wa_id = cur.lastrowid
                con.execute(
                    "INSERT INTO tbl_gdp_warenausgang_pos(wa_id,pzn,artikelname,charge,verfall,menge,ek) "
                    "VALUES(?,?,?,?,?,?,?)", (wa_id, d["pzn"], d["artikelname"], charge, verfall, q, d["ek"]))
                con.commit()
            self._log("Warenausgang", "Produziert / avisiert", wa_id,
                      f"{nummer}: {q} St {d['artikelname']} Charge {charge} -> Verkaufs-Wareneingang")
            self._refresh_produktionsbestand()
            self._refreshers.get("warenausgang", lambda: None)()
            self._refresh_wareneingang()
            messagebox.showinfo("Warenausgang erstellt",
                                f"Warenausgang {nummer} ({q} St) angelegt und an den "
                                "Verkaufs-Wareneingang avisiert. Dort bei Anlieferung bestaetigen.",
                                parent=self)

        rnr_label = "Rechnungsnummer *" if self._rechnung_pflicht() else "Rechnungsnummer"
        FormDialog(self, f"Produzieren → Warenausgang: {d['artikelname']}", [
            ("menge", "Menge produziert", "int", str(maxq)),
            ("charge", "Charge Fertigware", "text", d["charge"] or ""),
            ("verfall", "Verfall (MM/JJJJ)", "text", d["verfall"] or ""),
            ("rechnungsnummer", rnr_label, "text", ""),
            ("bemerkung", "Bemerkung", "text", ""),
        ], save, width=460)

    def _pb_abschreiben(self):
        d = self._pb_selected()
        if not d:
            messagebox.showinfo("Abschreiben", "Bitte eine Zeile im Produktionsbestand waehlen.", parent=self)
            return
        maxq = d["menge"]

        def save(v):
            q = max(0, min(int(v["menge"]), maxq))
            if q <= 0:
                return
            wert = round((d["ek"] or 0) * q, 2)
            with self._conn() as con:
                con.execute("UPDATE tbl_gdp_produktionsbestand SET menge=menge-?, aktualisiert_am=? WHERE id=?",
                            (q, _heute_iso(), d["id"]))
                con.execute(
                    """INSERT INTO tbl_gdp_abschreibung
                       (datum,pzn,artikelname,charge,verfall,menge,grund,wert_ek,bearbeiter)
                       VALUES(?,?,?,?,?,?,?,?,?)""",
                    (_heute_iso(), d["pzn"], d["artikelname"], d["charge"], d["verfall"],
                     q, "Produktionsbestand: " + v["grund"], wert, ME()))
                con.commit()
            self._log("Produktionsbestand", "Abgeschrieben", d["id"],
                      f"{q} St {d['artikelname']} Charge {d['charge']}, EK ~ {wert} EUR: {v['grund']}")
            self._refresh_produktionsbestand()
            messagebox.showinfo("Abgeschrieben",
                                f"{q} St abgeschrieben (EK-Wert ~ {wert} EUR).", parent=self)

        FormDialog(self, f"Abschreiben: {d['artikelname']} (Ch {d['charge']})", [
            ("menge", "Menge abschreiben", "int", str(maxq)),
            ("grund", "Grund", "combo",
             ["Beschaedigt", "Verfall ueberschritten", "Produktionsausschuss",
              "Nicht GDP-konform", "Sonstiges"], "Beschaedigt"),
        ], save)

    # ===================================================== WARENAUSGANG / AVIS
    WA_STATUS_TAG = {"avisiert": "warn", "bestaetigt": "ok", "storniert": "alert"}

    def _build_warenausgang(self, parent):
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(2, weight=1)
        bar = tk.Frame(parent, bg=SHELL_BG)
        bar.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        theme.PillButton(bar, "➕  Neuer Warenausgang", self._wa_new,
                         kind="primary", font_size=10).pack(side="left")
        theme.PillButton(bar, "\U0001F4E5  Liste importieren", self._wa_import,
                         kind="neutral", font_size=10).pack(side="left", padx=6)
        theme.PillButton(bar, "✖  Stornieren", self._wa_storno,
                         kind="neutral", font_size=10).pack(side="left", padx=6)
        theme.PillButton(bar, "↻  Aktualisieren", self._refresh_warenausgang,
                         kind="neutral", font_size=10).pack(side="left", padx=6)

        tk.Label(parent, text="Hier wird der Warenausgang nur angelegt/avisiert. Die Bestaetigung "
                              "erfolgt ausschliesslich im VERKAUF (Wareneingang), wenn die Ware "
                              "im Verteilerzentrum eintrifft. Doppelklick = Positionen.",
                 bg=SHELL_BG, fg=MUTED, font=(theme.FONT, 9), justify="left",
                 wraplength=900).grid(row=1, column=0, sticky="w", pady=(0, 6))

        outer, body = _card(parent)
        outer.grid(row=2, column=0, sticky="nsew")
        cols = ("Nummer", "Datum", "Quelle", "Rechnung", "Ziel", "Positionen", "Stueck", "Status")
        self._wa_tree = _make_tree(body, cols, (110, 85, 130, 120, 120, 75, 60, 100),
                                   anchors={"Positionen": "center", "Stueck": "center"})
        self._wa_tree.bind("<Double-1>", lambda _e: self._wa_show_positions())
        self._refreshers["warenausgang"] = self._refresh_warenausgang

    def _refresh_warenausgang(self):
        t = getattr(self, "_wa_tree", None)
        if t is None:
            return
        t.delete(*t.get_children())
        with self._conn() as con:
            rows = con.execute(
                """SELECT w.id, w.nummer, w.datum, w.quelle, COALESCE(w.rechnungsnummer,''), w.ziel, w.status,
                          (SELECT COUNT(*) FROM tbl_gdp_warenausgang_pos p WHERE p.wa_id=w.id),
                          (SELECT COALESCE(SUM(menge),0) FROM tbl_gdp_warenausgang_pos p WHERE p.wa_id=w.id)
                   FROM tbl_gdp_warenausgang w ORDER BY w.id DESC""").fetchall()
        for wid, nummer, datum, quelle, rnr, ziel, status, npos, stk in rows:
            t.insert("", "end", iid=str(wid),
                     values=(nummer or "", (datum or "")[:10], quelle or "", rnr, ziel or "",
                             npos, stk, status or ""),
                     tags=(self.WA_STATUS_TAG.get(status, "odd"),))

    def _wa_selected_id(self):
        sel = self._wa_tree.selection()
        return int(sel[0]) if sel else None

    def _wa_show_positions(self, wa_id=None):
        wa_id = wa_id or self._wa_selected_id()
        if not wa_id:
            return
        with self._conn() as con:
            hdr = con.execute("SELECT nummer, quelle, status, COALESCE(rechnungsnummer,'') "
                              "FROM tbl_gdp_warenausgang WHERE id=?", (wa_id,)).fetchone()
            pos = con.execute(
                "SELECT pzn, artikelname, charge, verfall, menge, ek FROM tbl_gdp_warenausgang_pos "
                "WHERE wa_id=? ORDER BY artikelname", (wa_id,)).fetchall()
        win = tk.Toplevel(self)
        win.title(f"Warenausgang {hdr[0] if hdr else wa_id}")
        win.configure(bg=SHELL_BG)
        win.geometry("680x420")
        rnr_txt = f"  ·  Rechnung {hdr[3]}" if (hdr and hdr[3]) else ""
        tk.Label(win, bg=ACCENT, fg="#FFFFFF", font=(theme.FONT, 12, "bold"), anchor="w",
                 text=f"  {hdr[0] if hdr else ''} · {hdr[1] if hdr else ''} · {hdr[2] if hdr else ''}{rnr_txt}"
                 ).pack(fill="x", ipady=8)
        wrap = tk.Frame(win, bg=SHELL_BG)
        wrap.pack(fill="both", expand=True, padx=12, pady=12)
        tree = _make_tree(wrap, ("PZN", "Artikel", "Charge", "Verfall", "Menge", "EK"),
                          (80, 240, 90, 80, 60, 70), anchors={"Menge": "center", "EK": "e"})
        for i, (pzn, name, charge, verf, menge, ek) in enumerate(pos):
            tree.insert("", "end", values=(pzn or "", name or "", charge or "", verf or "", menge,
                                           f"{ek:.2f}" if ek else ""), tags=("even" if i % 2 else "odd",))

    def _wa_storno(self):
        wa_id = self._wa_selected_id()
        if not wa_id:
            messagebox.showinfo("Storno", "Bitte einen Warenausgang waehlen.", parent=self)
            return
        with self._conn() as con:
            st = con.execute("SELECT status FROM tbl_gdp_warenausgang WHERE id=?", (wa_id,)).fetchone()
        if not st or st[0] != "avisiert":
            messagebox.showinfo("Storno", "Nur avisierte (noch nicht bestaetigte) Warenausgaenge "
                                "koennen storniert werden.", parent=self)
            return
        if not messagebox.askyesno("Storno", f"Warenausgang #{wa_id} stornieren?", parent=self):
            return
        with self._conn() as con:
            con.execute("UPDATE tbl_gdp_warenausgang SET status='storniert' WHERE id=?", (wa_id,))
            con.commit()
        self._log("Warenausgang", "Storniert", wa_id, "")
        self._refresh_warenausgang()
        self._refresh_wareneingang()

    def _wa_new(self):
        """Manueller Warenausgang mit beliebig vielen Positionen."""
        win = tk.Toplevel(self)
        win.title("Neuer Warenausgang")
        win.configure(bg=SHELL_BG)
        win.geometry("740x580")
        win.transient(self.winfo_toplevel())
        tk.Label(win, bg=ACCENT, fg="#FFFFFF", font=(theme.FONT, 13, "bold"), anchor="w",
                 text="  Neuer Warenausgang (Avis an den Verkaufs-Wareneingang)").pack(fill="x", ipady=9)
        head = tk.Frame(win, bg=SHELL_BG)
        head.pack(fill="x", padx=16, pady=(12, 4))
        quelle_var = tk.StringVar(value="Produktion")
        rnr_var = tk.StringVar()
        bem_var = tk.StringVar()
        rnr_lbl = "Rechnungsnr *:" if self._rechnung_pflicht() else "Rechnungsnr:"
        for lbl, var, w in (("Quelle:", quelle_var, 18), (rnr_lbl, rnr_var, 16),
                            ("Bemerkung:", bem_var, 24)):
            tk.Label(head, text=lbl, bg=SHELL_BG, fg=TEXT, font=(theme.FONT, 9, "bold")).pack(side="left", padx=(0, 4))
            tk.Entry(head, textvariable=var, width=w).pack(side="left", padx=(0, 12))

        addc = tk.Frame(win, bg=CARD_ALT, highlightbackground=BORDER, highlightthickness=1)
        addc.pack(fill="x", padx=16, pady=8)
        tk.Label(addc, text="Position hinzufuegen:", bg=CARD_ALT, fg=ACCENT,
                 font=(theme.FONT, 9, "bold")).pack(anchor="w", padx=10, pady=(8, 0))
        cur_art = {"pzn": None, "name": None}
        sb = SearchBox(
            addc.master if False else addc,
            fetch=lambda t: [(f"{pzn} · {name}", (pzn, name)) for pzn, name in self._search_nmg(t)],
            on_select=lambda p: cur_art.update(pzn=p[0], name=p[1]) or
                        art_lbl.config(text=f"{p[0]} · {p[1]}", fg=ACCENT), height=4)
        sb.pack(fill="x", padx=10, pady=2)
        art_lbl = tk.Label(addc, text="Kein Artikel gewaehlt", bg=CARD_ALT, fg=MUTED,
                           font=(theme.FONT, 9, "italic"))
        art_lbl.pack(anchor="w", padx=10)
        prow = tk.Frame(addc, bg=CARD_ALT)
        prow.pack(fill="x", padx=10, pady=(2, 8))
        ch_v, vf_v, mn_v, ek_v = (tk.StringVar() for _ in range(4))
        for lbl, var, w in (("Charge", ch_v, 12), ("Verfall", vf_v, 10), ("Menge", mn_v, 6), ("EK", ek_v, 8)):
            tk.Label(prow, text=lbl + ":", bg=CARD_ALT, font=(theme.FONT, 9, "bold")).pack(side="left", padx=(0, 3))
            tk.Entry(prow, textvariable=var, width=w).pack(side="left", padx=(0, 10))

        positions = []
        ptree_wrap = tk.Frame(win, bg=SHELL_BG)
        ptree_wrap.pack(fill="both", expand=True, padx=16, pady=(0, 8))
        ptree = _make_tree(ptree_wrap, ("PZN", "Artikel", "Charge", "Verfall", "Menge", "EK"),
                           (80, 230, 90, 80, 60, 70), anchors={"Menge": "center", "EK": "e"}, height=8)

        def add_pos():
            if not cur_art["pzn"]:
                messagebox.showwarning("Position", "Bitte zuerst einen Artikel waehlen.", parent=win)
                return
            try:
                menge = int(mn_v.get().strip())
                if menge <= 0:
                    raise ValueError
            except ValueError:
                messagebox.showwarning("Position", "Menge muss eine positive Zahl sein.", parent=win)
                return
            verfall = _normalize_verfall(vf_v.get()) or vf_v.get().strip()
            try:
                ek = float(ek_v.get().strip().replace(",", ".")) if ek_v.get().strip() else None
            except ValueError:
                ek = None
            rec = (cur_art["pzn"], cur_art["name"], ch_v.get().strip(), verfall, menge, ek)
            positions.append(rec)
            ptree.insert("", "end", values=(rec[0], rec[1], rec[2], rec[3], rec[4],
                                            f"{ek:.2f}" if ek else ""))
            cur_art.update(pzn=None, name=None)
            art_lbl.config(text="Kein Artikel gewaehlt", fg=MUTED)
            sb.clear()
            for v in (ch_v, vf_v, mn_v, ek_v):
                v.set("")

        def del_pos():
            sel = ptree.selection()
            if not sel:
                return
            idx = ptree.index(sel[0])
            ptree.delete(sel[0])
            if 0 <= idx < len(positions):
                positions.pop(idx)

        btnrow = tk.Frame(addc, bg=CARD_ALT)
        btnrow.pack(fill="x", padx=10, pady=(0, 8))
        theme.PillButton(btnrow, "+ Position", add_pos, kind="accent", font_size=9, padx=12, pady=4).pack(side="left")
        theme.PillButton(btnrow, "− Position entfernen", del_pos, kind="neutral", font_size=9, padx=12, pady=4).pack(side="left", padx=6)

        def save_wa():
            if not positions:
                messagebox.showwarning("Warenausgang", "Bitte mindestens eine Position hinzufuegen.", parent=win)
                return
            rnr = rnr_var.get().strip()
            if self._rechnung_pflicht() and not rnr:
                messagebox.showwarning("Warenausgang",
                    "Rechnungsnummer ist als Pflichtfeld gesetzt (Einstellungen).", parent=win)
                return
            self._wa_create(quelle_var.get().strip() or "Manuell", bem_var.get().strip(),
                            positions, rechnungsnummer=rnr)
            win.destroy()

        foot = tk.Frame(win, bg=SHELL_BG)
        foot.pack(fill="x", padx=16, pady=(0, 14))
        theme.PillButton(foot, "Warenausgang speichern & avisieren", save_wa, kind="success",
                         font_size=10, padx=16, pady=7).pack(side="right")
        theme.PillButton(foot, "Abbrechen", win.destroy, kind="neutral",
                         font_size=10, padx=16, pady=7).pack(side="right", padx=8)
        win.grab_set()

    def _wa_create(self, quelle, bemerkung, positions, rechnungsnummer=""):
        """Legt einen avisierten Warenausgang mit Positionen an."""
        jetzt = datetime.now().isoformat(timespec="seconds")
        with self._conn() as con:
            nummer = self._next_nummer(con, "WA")
            cur = con.execute(
                """INSERT INTO tbl_gdp_warenausgang
                   (datum,nummer,quelle,ziel,status,bemerkung,rechnungsnummer,erstellt_von)
                   VALUES(?,?,?, 'Verkaufsbestand', 'avisiert', ?,?, ?)""",
                (jetzt, nummer, quelle, bemerkung, rechnungsnummer, ME()))
            wa_id = cur.lastrowid
            for pzn, name, charge, verfall, menge, ek in positions:
                con.execute(
                    "INSERT INTO tbl_gdp_warenausgang_pos(wa_id,pzn,artikelname,charge,verfall,menge,ek) "
                    "VALUES(?,?,?,?,?,?,?)", (wa_id, pzn, name, charge, verfall, menge, ek))
            con.commit()
        self._log("Warenausgang", "Avisiert", wa_id,
                  f"{nummer}: {len(positions)} Position(en) -> Verkaufs-Wareneingang")
        self._refresh_warenausgang()
        self._refresh_wareneingang()
        messagebox.showinfo("Warenausgang",
                            f"Warenausgang {nummer} avisiert. Erscheint im Verkaufs-Wareneingang "
                            "zur Bestaetigung.", parent=self)
        return wa_id

    def _wa_import(self):
        path = filedialog.askopenfilename(
            parent=self, title="Warenausgangs-/Lieferliste waehlen (Excel, CSV oder TXT)",
            filetypes=[("Tabellen", "*.xlsx *.xlsm *.csv *.txt"), ("Alle Dateien", "*.*")])
        if not path:
            return
        try:
            positions, rechnung = self._parse_lieferliste(path)
        except Exception as e:
            messagebox.showerror("Import", f"Datei konnte nicht gelesen werden:\n{e}", parent=self)
            return
        if not positions:
            messagebox.showinfo("Import", "Keine gueltigen Positionen (PZN + Menge) gefunden.", parent=self)
            return
        if self._rechnung_pflicht() and not rechnung:
            messagebox.showwarning("Import",
                "Rechnungsnummer ist Pflicht, aber in der Liste wurde keine Spalte "
                "'Rechnung/Rechnungsnummer' gefunden.", parent=self)
            return
        from pathlib import Path as _P
        self._wa_create("Import: " + _P(path).name, "Importierte Liste", positions,
                        rechnungsnummer=rechnung)

    def _parse_lieferliste(self, path):
        """Liest eine Liste (Excel/CSV/TXT) -> (positions, rechnungsnummer).
        positions = [(pzn,name,charge,verfall,menge,ek), ...]. Nutzt dieselbe
        Format-/Spaltenerkennung wie der Kasse-Import."""
        from .file_loader import load_table, find_column
        from .kasse_import import _cell, _to_int, _to_float
        table = load_table(path)
        headers = list(table.headers)
        c_pzn = find_column(headers, ["PZN", "PZN-Code", "Artikel-PZN"], ["pzn"])
        c_charge = find_column(headers, ["Charge", "Chargennummer", "Lot", "Los"], ["charge", "lot"])
        c_verfall = find_column(headers, ["Verfall", "Verfalldatum", "Haltbarkeit", "MHD", "Verwendbar bis"],
                                ["verfall", "mhd", "haltbar"])
        c_menge = find_column(headers, ["Menge", "Anzahl", "Stück", "Stueck", "Bestand"],
                              ["menge", "anzahl", "stueck", "stck"])
        c_ek = find_column(headers, ["EK", "Einkaufspreis", "EK-Preis", "Einkauf"], ["ek", "einkauf"])
        c_name = find_column(headers, ["Artikel", "Artikelname", "Bezeichnung", "Name"], ["artikel", "bezeich"])
        c_rechnung = find_column(headers, ["Rechnungsnummer", "Rechnung", "Rechnungsnr", "RE-Nr", "Belegnummer"],
                                 ["rechnung", "belegnr"])
        if c_pzn is None or c_menge is None:
            raise ValueError("Es werden mindestens eine PZN- und eine Mengen-Spalte benoetigt.")
        out = []
        rechnung = ""
        with self._conn() as con:
            for row in table.rows:
                pzn = _cell(row, c_pzn)
                menge = _to_int(_cell(row, c_menge))
                if not pzn or not menge or menge <= 0:
                    continue
                name = _cell(row, c_name) if c_name is not None else ""
                if not name:
                    r = con.execute("SELECT artikelname FROM tbl_nmg_stamm WHERE pzn=? LIMIT 1", (pzn,)).fetchone()
                    name = r[0] if r else ""
                verfall = _normalize_verfall(_cell(row, c_verfall)) or _cell(row, c_verfall)
                if not rechnung and c_rechnung is not None:
                    rechnung = _cell(row, c_rechnung)
                out.append((pzn, name, _cell(row, c_charge), verfall, menge, _to_float(_cell(row, c_ek))))
        return out, rechnung

    def _wa_confirm(self, wa_id, override_positions=None):
        """Bestaetigt einen avisierten Warenausgang am Verkaufs-Wareneingang:
        bucht die Positionen in den Verkaufsbestand (tbl_lagerbestand), legt einen
        Wareneingang (art='bestand') an und markiert den Avis als bestaetigt.
        override_positions (z.B. aus einer Anlieferliste) ersetzen die avisierten."""
        with self._conn() as con:
            hdr = con.execute("SELECT nummer, quelle, status FROM tbl_gdp_warenausgang WHERE id=?",
                              (wa_id,)).fetchone()
            if not hdr:
                return False
            if hdr[2] != "avisiert":
                messagebox.showinfo("Bestaetigen", f"Warenausgang {hdr[0]} ist bereits '{hdr[2]}'.", parent=self)
                return False
            positions = override_positions if override_positions is not None else con.execute(
                "SELECT pzn, artikelname, charge, verfall, menge, ek FROM tbl_gdp_warenausgang_pos "
                "WHERE wa_id=?", (wa_id,)).fetchall()
            # EK aus dem Avis (stammt z.B. aus der eRechnung-Uebernahme im
            # Produktionsbestand) erhalten, falls die Anlieferliste keinen Preis
            # enthaelt - sonst ginge der Einkaufspreis Richtung Verkaufsbestand verloren.
            ek_lookup = {}
            if override_positions is not None:
                for apzn, _an, ach, _av, _am, aek in con.execute(
                        "SELECT pzn, artikelname, charge, verfall, menge, ek "
                        "FROM tbl_gdp_warenausgang_pos WHERE wa_id=?", (wa_id,)).fetchall():
                    if aek:
                        ek_lookup[((apzn or "").strip(), (ach or "").strip())] = aek
                        ek_lookup.setdefault((apzn or "").strip(), aek)
            jetzt = datetime.now().isoformat(timespec="seconds")
            cur = con.execute(
                "INSERT INTO tbl_wareneingang(datum,lieferant,lieferschein,bearbeiter,art) "
                "VALUES(?,?,?,?, 'bestand')", (jetzt, hdr[1] or "Produktion", hdr[0], ME()))
            we_id = cur.lastrowid
            total = 0
            for pzn, name, charge, verfall, menge, ek in positions:
                if not menge or menge <= 0:
                    continue
                if not ek:  # Anlieferliste ohne Preis -> EK aus dem Avis nachziehen
                    ek = (ek_lookup.get(((pzn or "").strip(), (charge or "").strip()))
                          or ek_lookup.get((pzn or "").strip()))
                con.execute(
                    "INSERT INTO tbl_wareneingang_positionen(we_id,pzn,artikelname,charge,verfall,menge,ek) "
                    "VALUES(?,?,?,?,?,?,?)", (we_id, pzn, name, charge, verfall, menge, ek))
                upd = con.execute(
                    "UPDATE tbl_lagerbestand SET menge=menge+?, ek=COALESCE(?,ek), aktualisiert_am=? "
                    "WHERE pzn=? AND COALESCE(charge,'')=? AND COALESCE(verfall,'')=?",
                    (menge, ek, jetzt, pzn, charge or "", verfall or ""))
                if upd.rowcount == 0:
                    con.execute(
                        "INSERT INTO tbl_lagerbestand(pzn,artikelname,charge,verfall,menge,ek,aktualisiert_am) "
                        "VALUES(?,?,?,?,?,?,?)", (pzn, name, charge, verfall, menge, ek, jetzt))
                total += menge
            con.execute("UPDATE tbl_gdp_warenausgang SET status='bestaetigt', bestaetigt_am=?, "
                        "bestaetigt_von=?, we_id=? WHERE id=?", (jetzt, ME(), we_id, wa_id))
            con.commit()
        self._log("Wareneingang", "Avis bestaetigt", we_id,
                  f"{hdr[0]}: {total} St in den Verkaufsbestand gebucht")
        self._refresh_wareneingang()
        self._refresh_warenausgang()
        messagebox.showinfo("Bestaetigt",
                            f"Warenausgang {hdr[0]} bestaetigt: {total} St im Verkaufsbestand. "
                            "Damit verkaufbar (Kasse).", parent=self)
        return True

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

    # ===================================================== WARENBEWEGUNGEN
    def _build_bewegungen(self, parent):
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(2, weight=1)
        bar = tk.Frame(parent, bg=SHELL_BG)
        bar.grid(row=0, column=0, sticky="ew", pady=(0, 6))
        tk.Label(bar, text="Suche:", bg=SHELL_BG, fg=TEXT, font=(theme.FONT, 10)).pack(side="left")
        self._bw_var = tk.StringVar()
        e = tk.Entry(bar, textvariable=self._bw_var, font=(theme.FONT, 11), width=24,
                     relief="solid", bd=1)
        e.pack(side="left", padx=8)
        e.bind("<Return>", lambda _e: self._refresh_bewegungen())
        tk.Label(bar, text="Richtung:", bg=SHELL_BG, fg=MUTED, font=(theme.FONT, 9)).pack(side="left", padx=(6, 2))
        self._bw_richtung = tk.StringVar(value="Alle")
        cb1 = ttk.Combobox(bar, textvariable=self._bw_richtung, state="readonly", width=10,
                           style="NMG.TCombobox", values=["Alle", "Eingang", "Ausgang"])
        cb1.pack(side="left")
        cb1.bind("<<ComboboxSelected>>", lambda _e: self._refresh_bewegungen())
        tk.Label(bar, text="Bereich:", bg=SHELL_BG, fg=MUTED, font=(theme.FONT, 9)).pack(side="left", padx=(6, 2))
        self._bw_bereich = tk.StringVar(value="Alle")
        cb2 = ttk.Combobox(bar, textvariable=self._bw_bereich, state="readonly", width=12,
                           style="NMG.TCombobox", values=["Alle", "Verkauf", "Produktion"])
        cb2.pack(side="left")
        cb2.bind("<<ComboboxSelected>>", lambda _e: self._refresh_bewegungen())
        theme.PillButton(bar, "\U0001F50E  Suchen", self._refresh_bewegungen,
                         kind="primary", font_size=10).pack(side="left", padx=8)
        theme.PillButton(bar, "⬇  Export CSV", self._bw_export, kind="neutral",
                         font_size=10).pack(side="left")

        self._bw_count = tk.Label(parent, text="", bg=SHELL_BG, fg=MUTED, font=(theme.FONT, 9))
        self._bw_count.grid(row=1, column=0, sticky="w", pady=(0, 4))

        outer, body = _card(parent)
        outer.grid(row=2, column=0, sticky="nsew")
        cols = ("Datum", "Richtung", "Bereich", "Beleg", "Partner", "PZN", "Artikel",
                "Charge", "Verfall", "Menge", "Status")
        self._bw_tree = _make_tree(
            body, cols, (130, 75, 90, 95, 150, 75, 200, 90, 75, 55, 100),
            anchors={"Menge": "center"})
        self._bw_tree.tag_configure("ein", background="#EAF6EE")
        self._bw_tree.tag_configure("aus", background="#EAF1FB")
        self._refreshers["bewegungen"] = self._refresh_bewegungen

    def _movements(self, con, term=None, richtung="Alle", bereich="Alle"):
        """Vereinte Warenbewegungen aus drei Quellen -> Liste von
        (datum,richtung,bereich,beleg,partner,pzn,artikel,charge,verfall,menge,ek,info)."""
        like = f"%{term}%" if term else None
        rows = []
        want_ein = richtung in ("Alle", "Eingang")
        want_aus = richtung in ("Alle", "Ausgang")
        # 1) Wareneingaenge (art bestand=Verkauf, produktion=Produktion)
        if want_ein and _table_exists(con, "tbl_wareneingang"):
            has_art = "art" in {r[1] for r in con.execute("PRAGMA table_info(tbl_wareneingang)")}
            art_sel = "COALESCE(w.art,'bestand')" if has_art else "'bestand'"
            q = (f"SELECT w.datum, 'Eingang', "
                 f"CASE {art_sel} WHEN 'produktion' THEN 'Produktion' ELSE 'Verkauf' END, "
                 "COALESCE(w.lieferschein,''), COALESCE(w.lieferant,''), p.pzn, p.artikelname, "
                 "p.charge, p.verfall, p.menge, p.ek, '' "
                 "FROM tbl_wareneingang_positionen p JOIN tbl_wareneingang w ON w.id=p.we_id")
            params = []
            if like:
                q += (" WHERE (p.pzn LIKE ? OR p.artikelname LIKE ? OR p.charge LIKE ? "
                      "OR COALESCE(w.lieferant,'') LIKE ? OR COALESCE(w.lieferschein,'') LIKE ?)")
                params += [like] * 5
            rows.extend(con.execute(q, params).fetchall())
        # 2) Warenausgaenge aus der Produktion
        if want_aus and _table_exists(con, "tbl_gdp_warenausgang"):
            q = ("SELECT w.datum, 'Ausgang', 'Produktion', w.nummer, COALESCE(w.quelle,''), "
                 "p.pzn, p.artikelname, p.charge, p.verfall, p.menge, p.ek, COALESCE(w.status,'') "
                 "FROM tbl_gdp_warenausgang_pos p JOIN tbl_gdp_warenausgang w ON w.id=p.wa_id")
            params = []
            if like:
                q += (" WHERE (p.pzn LIKE ? OR p.artikelname LIKE ? OR p.charge LIKE ? "
                      "OR COALESCE(w.nummer,'') LIKE ? OR COALESCE(w.rechnungsnummer,'') LIKE ?)")
                params += [like] * 5
            rows.extend(con.execute(q, params).fetchall())
        # 3) Verkauf an Apotheken (Kasse)
        if want_aus and _table_exists(con, "tbl_bestellpositionen") and _table_exists(con, "tbl_bestellungen"):
            q = ("SELECT b.datum, 'Ausgang', 'Verkauf', ('Best-'||b.id), COALESCE(b.apotheke,''), "
                 "p.pzn, p.artikelname, p.charge, p.verfall, p.menge, NULL, COALESCE(b.status,'') "
                 "FROM tbl_bestellpositionen p JOIN tbl_bestellungen b ON b.id=p.bestell_id "
                 "WHERE COALESCE(p.bestellart,'Bestellung')='Bestellung'")
            params = []
            if like:
                q += (" AND (p.pzn LIKE ? OR p.artikelname LIKE ? OR p.charge LIKE ? "
                      "OR COALESCE(b.apotheke,'') LIKE ?)")
                params += [like] * 4
            rows.extend(con.execute(q, params).fetchall())
        if bereich != "Alle":
            rows = [r for r in rows if r[2] == bereich]
        rows.sort(key=lambda r: (r[0] or ""), reverse=True)
        return rows

    def _refresh_bewegungen(self):
        t = getattr(self, "_bw_tree", None)
        if t is None:
            return
        t.delete(*t.get_children())
        term = self._bw_var.get().strip()
        with self._conn() as con:
            rows = self._movements(con, term or None, self._bw_richtung.get(), self._bw_bereich.get())
        ein = aus = 0
        for i, r in enumerate(rows[:2000]):
            (datum, richtung, bereich, beleg, partner, pzn, artikel, charge, verfall, menge, ek, info) = r
            if richtung == "Eingang":
                ein += 1
            else:
                aus += 1
            t.insert("", "end", values=((datum or "")[:16], richtung, bereich, beleg or "",
                                        partner or "", pzn or "", artikel or "", charge or "",
                                        verfall or "", menge, info or ""),
                     tags=("ein" if richtung == "Eingang" else "aus",))
        gezeigt = min(len(rows), 2000)
        mehr = f" (von {len(rows)})" if len(rows) > 2000 else ""
        self._bw_count.config(
            text=f"{gezeigt} Bewegung(en){mehr} · {ein} Eingang · {aus} Ausgang"
                 + (f" · Suche: '{term}'" if term else ""))

    def _bw_export(self):
        term = self._bw_var.get().strip()
        with self._conn() as con:
            rows = self._movements(con, term or None, self._bw_richtung.get(), self._bw_bereich.get())
        if not rows:
            messagebox.showinfo("Export", "Keine Bewegungen zum Exportieren.", parent=self)
            return
        path = filedialog.asksaveasfilename(
            parent=self, title="Warenbewegungen exportieren", defaultextension=".csv",
            initialfile=f"Warenbewegungen_{_heute_iso()}.csv", filetypes=[("CSV", "*.csv")])
        if not path:
            return
        with open(path, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f, delimiter=";")
            w.writerow(["Datum", "Richtung", "Bereich", "Beleg", "Partner", "PZN", "Artikel",
                        "Charge", "Verfall", "Menge", "EK", "Status"])
            for r in rows:
                w.writerow([(r[0] or "")[:19], r[1], r[2], r[3], r[4], r[5], r[6], r[7],
                            r[8], r[9], r[10] if r[10] is not None else "", r[11]])
        messagebox.showinfo("Export", f"{len(rows)} Bewegung(en) gespeichert:\n{path}", parent=self)

    # ===================================================== BESTANDSDIFFERENZEN
    BD_GRUENDE = ["Inventur", "Bruch / Beschaedigung", "Schwund / Diebstahl",
                  "Fund / Mehrbestand", "Buchungsfehler", "Verfall entsorgt", "Sonstiges"]

    def _build_bestandsdiff(self, parent):
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(2, weight=1)
        bar = tk.Frame(parent, bg=SHELL_BG)
        bar.grid(row=0, column=0, sticky="ew", pady=(0, 6))
        theme.PillButton(bar, "➕  Differenz erfassen", self._bd_new,
                         kind="primary", font_size=10).pack(side="left")
        tk.Label(bar, text="Bereich:", bg=SHELL_BG, fg=MUTED, font=(theme.FONT, 9)).pack(side="left", padx=(10, 2))
        self._bd_filter = tk.StringVar(value="Alle")
        cb = ttk.Combobox(bar, textvariable=self._bd_filter, state="readonly", width=12,
                          style="NMG.TCombobox", values=["Alle", "Verkauf", "Produktion"])
        cb.pack(side="left")
        cb.bind("<<ComboboxSelected>>", lambda _e: self._refresh_bestandsdiff())
        theme.PillButton(bar, "↻  Aktualisieren", self._refresh_bestandsdiff,
                         kind="neutral", font_size=10).pack(side="left", padx=8)
        theme.PillButton(bar, "⬇  Export CSV", self._bd_export, kind="neutral",
                         font_size=10).pack(side="left")

        self._bd_summary = tk.Label(parent, text="", bg=SHELL_BG, fg=MUTED, font=(theme.FONT, 9))
        self._bd_summary.grid(row=1, column=0, sticky="w", pady=(0, 4))

        outer, body = _card(parent)
        outer.grid(row=2, column=0, sticky="nsew")
        cols = ("Datum", "Bereich", "PZN", "Artikel", "Charge", "Verfall",
                "Vorher", "Differenz", "Nachher", "Grund", "Bearbeiter")
        self._bd_tree = _make_tree(
            body, cols, (130, 90, 75, 190, 90, 75, 60, 75, 60, 140, 100),
            anchors={"Vorher": "center", "Differenz": "center", "Nachher": "center"})
        self._bd_tree.tag_configure("plus", foreground=OK_GREEN)
        self._bd_tree.tag_configure("minus", foreground=DANGER)
        self._refreshers["bestandsdiff"] = self._refresh_bestandsdiff

    def _refresh_bestandsdiff(self):
        t = getattr(self, "_bd_tree", None)
        if t is None:
            return
        t.delete(*t.get_children())
        flt = self._bd_filter.get()
        where, params = ("", ())
        if flt != "Alle":
            where, params = "WHERE bereich=?", (flt,)
        with self._conn() as con:
            rows = con.execute(
                f"""SELECT datum, bereich, pzn, artikelname, charge, verfall,
                          menge_vorher, menge_diff, menge_nachher, grund, bearbeiter
                   FROM tbl_gdp_bestandsdiff {where} ORDER BY id DESC""", params).fetchall()
        plus = minus = 0
        for r in rows:
            diff = r[7] or 0
            if diff > 0:
                plus += diff
            else:
                minus += diff
            diff_txt = f"+{diff}" if diff > 0 else str(diff)
            t.insert("", "end",
                     values=((r[0] or "")[:16], r[1] or "", r[2] or "", r[3] or "", r[4] or "",
                             r[5] or "", r[6], diff_txt, r[8], r[9] or "", r[10] or ""),
                     tags=("plus" if diff > 0 else ("minus" if diff < 0 else ""),))
        self._bd_summary.config(
            text=f"{len(rows)} Korrektur(en) · Summe Mehrbestand +{plus} · Summe Fehlbestand {minus}")

    def _bd_new(self):
        """Erfasst eine manuelle Bestandsdifferenz im Bestand des aktuellen Modus.
        Zeigt den Soll-Bestand; Ist-Bestand ODER +/- Differenz eingebbar (verknuepft)."""
        if self._we_produktion_aktiv():
            bereich, tabelle = "Produktion", "tbl_gdp_produktionsbestand"
        else:
            bereich, tabelle = "Verkauf", "tbl_lagerbestand"
        with self._conn() as con:
            if not _table_exists(con, tabelle):
                messagebox.showinfo("Bestandsdifferenz", f"Kein {bereich}-Bestand vorhanden.", parent=self)
                return
            lines = con.execute(
                f"SELECT id, pzn, artikelname, charge, verfall, menge FROM {tabelle} "
                "ORDER BY artikelname").fetchall()
        if not lines:
            messagebox.showinfo("Bestandsdifferenz",
                                f"Im {bereich}-Bestand sind keine Artikel vorhanden.", parent=self)
            return
        labels = [f"{name} | Ch {ch or '-'} | {pzn}" for (_id, pzn, name, ch, vf, menge) in lines]
        amap = {labels[i]: lines[i] for i in range(len(lines))}

        win = tk.Toplevel(self)
        win.title("Bestandsdifferenz erfassen")
        win.configure(bg=SHELL_BG)
        win.resizable(False, False)
        win.transient(self.winfo_toplevel())
        tk.Label(win, bg=ACCENT, fg="#FFFFFF", font=(theme.FONT, 13, "bold"), anchor="w",
                 text=f"  Bestandsdifferenz erfassen ({bereich})").pack(fill="x", ipady=9)
        body = tk.Frame(win, bg=SHELL_BG)
        body.pack(fill="both", expand=True, padx=18, pady=14)

        def row(label):
            r = tk.Frame(body, bg=SHELL_BG)
            r.pack(fill="x", pady=4)
            tk.Label(r, text=label, bg=SHELL_BG, fg=TEXT, width=20, anchor="w",
                     font=(theme.FONT, 10)).pack(side="left")
            return r

        art_var = tk.StringVar(value=labels[0])
        r0 = row("Artikel / Charge")
        ttk.Combobox(r0, textvariable=art_var, values=labels, state="readonly",
                     style="NMG.TCombobox", width=34).pack(side="left", fill="x", expand=True)

        rs = row("Soll-Bestand (System)")
        soll_lbl = tk.Label(rs, text="–", bg=SHELL_BG, fg=ACCENT, font=(theme.FONT, 13, "bold"))
        soll_lbl.pack(side="left")

        ist_var = tk.StringVar()
        ri = row("Ist-Bestand (gezaehlt)")
        ist_entry = tk.Entry(ri, textvariable=ist_var, font=(theme.FONT, 11), width=10,
                             relief="solid", bd=1)
        ist_entry.pack(side="left")
        tk.Label(ri, text="Stueck", bg=SHELL_BG, fg=MUTED, font=(theme.FONT, 9)).pack(side="left", padx=6)

        diff_var = tk.StringVar()
        rd = row("oder Differenz (+/-)")
        diff_entry = tk.Entry(rd, textvariable=diff_var, font=(theme.FONT, 11), width=8,
                              relief="solid", bd=1)
        diff_entry.pack(side="left")
        # Schnell-Schritt: meist geht der Bestand manuell nach unten -> "-" prominent
        theme.PillButton(rd, "−1", lambda: step(-1), kind="danger",
                         font_size=11, padx=10, pady=2).pack(side="left", padx=(8, 2))
        theme.PillButton(rd, "+1", lambda: step(1), kind="neutral",
                         font_size=11, padx=10, pady=2).pack(side="left", padx=2)
        tk.Label(rd, text="oder Zahl eintippen", bg=SHELL_BG, fg=MUTED,
                 font=(theme.FONT, 9)).pack(side="left", padx=6)

        grund_var = tk.StringVar(value="Inventur")
        rg = row("Grund")
        ttk.Combobox(rg, textvariable=grund_var, values=self.BD_GRUENDE, state="readonly",
                     style="NMG.TCombobox", width=24).pack(side="left")
        bem_var = tk.StringVar()
        rb = row("Bemerkung")
        tk.Entry(rb, textvariable=bem_var, font=(theme.FONT, 10), width=30,
                 relief="solid", bd=1).pack(side="left", fill="x", expand=True)

        ergebnis = tk.Label(body, text="", bg=SHELL_BG, fg=MUTED, font=(theme.FONT, 10, "bold"))
        ergebnis.pack(anchor="w", pady=(8, 0))

        state = {"soll": 0}

        def _i(s):
            s = (s or "").strip().replace("+", "")
            try:
                return int(s)
            except ValueError:
                return None

        def update_ergebnis():
            soll = state["soll"]
            ist = _i(ist_var.get())
            if ist is None:
                ergebnis.config(text="", fg=MUTED)
                return
            diff = ist - soll
            ergebnis.config(
                text=f"Ergebnis: {soll} → {ist}   ({'+' if diff > 0 else ''}{diff} Stueck)",
                fg=(OK_GREEN if diff > 0 else DANGER if diff < 0 else MUTED))

        def on_art(*_a):
            line = amap[art_var.get()]
            state["soll"] = line[5] or 0
            soll_lbl.config(text=str(state["soll"]))
            ist_var.set(str(state["soll"]))
            diff_var.set("0")
            update_ergebnis()

        def on_ist(_e=None):
            ist = _i(ist_var.get())
            if ist is not None:
                diff_var.set(f"{ist - state['soll']:+d}")
            update_ergebnis()

        def on_diff(_e=None):
            diff = _i(diff_var.get())
            if diff is not None:
                ist_var.set(str(state["soll"] + diff))
            update_ergebnis()

        def step(delta):
            cur = (_i(diff_var.get()) or 0) + delta
            diff_var.set(f"{cur:+d}")
            ist_var.set(str(state["soll"] + cur))
            update_ergebnis()

        art_var.trace_add("write", on_art)
        ist_entry.bind("<KeyRelease>", on_ist)
        diff_entry.bind("<KeyRelease>", on_diff)
        on_art()

        def save():
            line = amap[art_var.get()]
            lid, pzn, name, charge, verfall, vorher = line
            vorher = vorher or 0
            ist = _i(ist_var.get())
            if ist is None:
                messagebox.showwarning("Bestandsdifferenz", "Bitte Ist-Bestand oder Differenz eingeben.", parent=win)
                return
            if ist < 0:
                messagebox.showwarning("Bestandsdifferenz", "Der Ist-Bestand darf nicht negativ sein.", parent=win)
                return
            diff = ist - vorher
            if diff == 0:
                messagebox.showinfo("Bestandsdifferenz",
                                    "Ist-Bestand entspricht dem Soll - keine Differenz.", parent=win)
                return
            jetzt = datetime.now().isoformat(timespec="seconds")
            with self._conn() as con:
                con.execute(f"UPDATE {tabelle} SET menge=?, aktualisiert_am=? WHERE id=?",
                            (ist, jetzt, lid))
                con.execute(
                    """INSERT INTO tbl_gdp_bestandsdiff
                       (datum,bereich,pzn,artikelname,charge,verfall,menge_vorher,menge_diff,
                        menge_nachher,grund,bemerkung,bearbeiter)
                       VALUES(?,?,?,?,?,?,?,?,?,?,?,?)""",
                    (jetzt, bereich, pzn, name, charge, verfall, vorher, diff, ist,
                     grund_var.get(), bem_var.get().strip(), ME()))
                con.commit()
            self._log("Bestandsdifferenz", f"{bereich}: {'+' if diff > 0 else ''}{diff}", lid,
                      f"{name} Ch {charge}: {vorher} -> {ist} ({grund_var.get()})")
            self._refresh_bestandsdiff()
            if hasattr(self, "_bw_tree"):
                self._refresh_bewegungen()
            win.destroy()
            messagebox.showinfo("Bestandsdifferenz",
                                f"Korrektur gebucht: {name} Ch {charge}\n"
                                f"{vorher} → {ist} ({'+' if diff > 0 else ''}{diff} St).", parent=self)

        foot = tk.Frame(win, bg=SHELL_BG)
        foot.pack(fill="x", padx=18, pady=(0, 16))
        theme.PillButton(foot, "Korrektur buchen", save, kind="success",
                         font_size=10, padx=16, pady=7).pack(side="right")
        theme.PillButton(foot, "Abbrechen", win.destroy, kind="neutral",
                         font_size=10, padx=16, pady=7).pack(side="right", padx=8)
        try:
            x = self.winfo_rootx() + 90
            y = self.winfo_rooty() + 70
            win.geometry(f"+{x}+{y}")
        except Exception:
            pass
        win.grab_set()
        ist_entry.focus_set()

    def _bd_export(self):
        flt = self._bd_filter.get()
        where, params = ("", ())
        if flt != "Alle":
            where, params = "WHERE bereich=?", (flt,)
        with self._conn() as con:
            rows = con.execute(
                f"""SELECT datum,bereich,pzn,artikelname,charge,verfall,menge_vorher,menge_diff,
                          menge_nachher,grund,bemerkung,bearbeiter
                   FROM tbl_gdp_bestandsdiff {where} ORDER BY id DESC""", params).fetchall()
        if not rows:
            messagebox.showinfo("Export", "Keine Bestandsdifferenzen zum Exportieren.", parent=self)
            return
        path = filedialog.asksaveasfilename(
            parent=self, title="Bestandsdifferenzen exportieren", defaultextension=".csv",
            initialfile=f"Bestandsdifferenzen_{_heute_iso()}.csv", filetypes=[("CSV", "*.csv")])
        if not path:
            return
        with open(path, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f, delimiter=";")
            w.writerow(["Datum", "Bereich", "PZN", "Artikel", "Charge", "Verfall", "Vorher",
                        "Differenz", "Nachher", "Grund", "Bemerkung", "Bearbeiter"])
            w.writerows(rows)
        messagebox.showinfo("Export", f"{len(rows)} Korrektur(en) gespeichert.", parent=self)

    # ===================================================== EINSTELLUNGEN
    MODUS_INFO = {
        "verkauf": ("\U0001F6D2  Verkauf",
                    "Dieser Arbeitsplatz nimmt Fertigware in den Verkaufsbestand, "
                    "bestaetigt avisierte Lieferungen aus der Produktion und bearbeitet "
                    "Retouren, Retourenbestand und Kundenqualifizierung.",
                    "Sichtbar: Wareneingang (→ Verkaufsbestand) · Retouren · Retourenbestand · "
                    "Kundenqualifizierung · Rueckverfolgung"),
        "produktion": ("\U0001F3ED  Produktion",
                       "Dieser Arbeitsplatz nimmt eingekaufte Ware fuer die Produktion an "
                       "(getrennter Produktionsbestand) und erzeugt beim Produzieren einen "
                       "Warenausgang, der an den Verkauf avisiert wird.",
                       "Sichtbar: Wareneingang (Einkauf → Produktion) · Produktionsbestand · "
                       "Warenausgang / Avis · Rueckverfolgung"),
    }

    def _build_einstellungen(self, parent):
        parent.columnconfigure(0, weight=1)
        outer, body = _card(parent, "Betriebsmodus", "Verkauf oder Produktion - pro Arbeitsplatz")
        outer.grid(row=0, column=0, sticky="ew", pady=(0, 12))
        tk.Label(body, text="Lege fest, ob dieser Arbeitsplatz im Verkauf oder in der Produktion "
                            "arbeitet. Das Programm blendet automatisch die passenden Bereiche ein "
                            "und aktiviert den richtigen Wareneingang. Beide Modi teilen sich "
                            "dieselbe Datenbank - so reicht die Produktion ihren Warenausgang an "
                            "den Verkauf weiter.",
                 bg=BG, fg=MUTED, font=(theme.FONT, 10), justify="left", wraplength=760).pack(anchor="w", pady=(0, 10))

        self._set_modus = tk.StringVar(value=self._modus())

        def modus_row(key):
            icon_title, desc, sicht = self.MODUS_INFO[key]
            rowf = tk.Frame(body, bg=CARD_ALT, highlightbackground=BORDER, highlightthickness=1)
            rowf.pack(fill="x", pady=4)
            rb = tk.Radiobutton(rowf, variable=self._set_modus, value=key, bg=CARD_ALT,
                                activebackground=CARD_ALT, selectcolor=CARD,
                                command=self._on_modus_change)
            rb.pack(side="left", padx=(10, 6), pady=12)
            txt = tk.Frame(rowf, bg=CARD_ALT)
            txt.pack(side="left", fill="x", expand=True, pady=8)
            tk.Label(txt, text=icon_title, bg=CARD_ALT, fg=INK,
                     font=(theme.FONT, 12, "bold")).pack(anchor="w")
            tk.Label(txt, text=desc, bg=CARD_ALT, fg=MUTED, font=(theme.FONT, 9),
                     justify="left", wraplength=700).pack(anchor="w")
            tk.Label(txt, text=sicht, bg=CARD_ALT, fg=ACCENT, font=(theme.FONT, 8),
                     justify="left", wraplength=700).pack(anchor="w", pady=(3, 0))

        modus_row("verkauf")
        modus_row("produktion")
        self._set_hint = tk.Label(body, text="", bg=BG, fg=ACCENT, font=(theme.FONT, 9, "bold"))
        self._set_hint.pack(anchor="w", pady=(8, 0))

        # ---- Felder / Pflichtangaben ----
        outer2, body2 = _card(parent, "Felder", "Pflichtangaben beim Warenausgang")
        outer2.grid(row=1, column=0, sticky="ew")
        self._set_rnr_pflicht = tk.IntVar(value=1 if self._rechnung_pflicht() else 0)
        rowf = tk.Frame(body2, bg=CARD_ALT, highlightbackground=BORDER, highlightthickness=1)
        rowf.pack(fill="x", pady=4)
        tk.Checkbutton(rowf, variable=self._set_rnr_pflicht, bg=CARD_ALT, activebackground=CARD_ALT,
                       command=self._on_rnr_pflicht_change).pack(side="left", padx=(10, 6), pady=10)
        txt = tk.Frame(rowf, bg=CARD_ALT)
        txt.pack(side="left", fill="x", expand=True, pady=6)
        tk.Label(txt, text="Rechnungsnummer ist Pflichtfeld", bg=CARD_ALT, fg=INK,
                 font=(theme.FONT, 11, "bold")).pack(anchor="w")
        tk.Label(txt, text="Wenn aktiv, muss beim Anlegen/Importieren eines Warenausgangs eine "
                           "Rechnungsnummer angegeben werden. Sonst ist sie optional.",
                 bg=CARD_ALT, fg=MUTED, font=(theme.FONT, 9), justify="left",
                 wraplength=700).pack(anchor="w")
        self._refreshers["einstellungen"] = self._refresh_einstellungen

    def _on_rnr_pflicht_change(self):
        self._set_setting("rechnungsnummer_pflicht", self._set_rnr_pflicht.get())

    def _refresh_einstellungen(self):
        if hasattr(self, "_set_modus"):
            self._set_modus.set(self._modus())
        if hasattr(self, "_set_rnr_pflicht"):
            self._set_rnr_pflicht.set(1 if self._rechnung_pflicht() else 0)

    def _on_modus_change(self):
        modus = self._set_modus.get()
        self._set_setting("betriebsmodus", modus)
        if hasattr(self, "_set_hint"):
            self._set_hint.config(
                text=f"Aktiver Modus: {self.MODUS_INFO[modus][0].strip()}. "
                     "Die Bereiche wurden umgestellt.")
        self._apply_nav_visibility()
        self._we_sync_art_selector()
        self._refresh_wareneingang()

    # ===================================================== PROTOKOLL
    LOG_MODULE = ["alle", "Wareneingang", "Warenausgang", "Produktionsbestand",
                  "Bestandsdifferenz", "Retoure", "Retourenbestand", "Rueckruf",
                  "Qualifizierung"]
    LOG_ZEITRAUM = ["Alle", "Dieser Monat", "Dieses Quartal", "Dieses Jahr",
                    "Vorjahr", "Manuell"]

    def _build_protokoll(self, parent):
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(2, weight=1)
        bar = tk.Frame(parent, bg=SHELL_BG)
        bar.grid(row=0, column=0, sticky="ew", pady=(0, 4))
        tk.Label(bar, text="Modul:", bg=SHELL_BG, fg=TEXT, font=(theme.FONT, 10)).pack(side="left")
        self._log_filter = tk.StringVar(value="alle")
        cb = ttk.Combobox(bar, textvariable=self._log_filter, state="readonly", width=18,
                          style="NMG.TCombobox", values=self.LOG_MODULE)
        cb.pack(side="left", padx=8)
        cb.bind("<<ComboboxSelected>>", lambda _e: self._refresh_protokoll())
        tk.Label(bar, text="Zeitraum:", bg=SHELL_BG, fg=TEXT, font=(theme.FONT, 10)).pack(side="left", padx=(8, 2))
        self._log_zeitraum = tk.StringVar(value="Alle")
        cbz = ttk.Combobox(bar, textvariable=self._log_zeitraum, state="readonly", width=15,
                           style="NMG.TCombobox", values=self.LOG_ZEITRAUM)
        cbz.pack(side="left", padx=4)
        cbz.bind("<<ComboboxSelected>>", lambda _e: self._on_zeitraum_change())
        # Manuell: von/bis (nur bei 'Manuell' sichtbar)
        self._log_von = tk.StringVar()
        self._log_bis = tk.StringVar()
        self._log_manual = tk.Frame(bar, bg=SHELL_BG)
        tk.Label(self._log_manual, text="von", bg=SHELL_BG, fg=MUTED, font=(theme.FONT, 9)).pack(side="left")
        ev = tk.Entry(self._log_manual, textvariable=self._log_von, width=11, relief="solid", bd=1)
        ev.pack(side="left", padx=3)
        tk.Label(self._log_manual, text="bis", bg=SHELL_BG, fg=MUTED, font=(theme.FONT, 9)).pack(side="left")
        eb = tk.Entry(self._log_manual, textvariable=self._log_bis, width=11, relief="solid", bd=1)
        eb.pack(side="left", padx=3)
        for e in (ev, eb):
            e.bind("<Return>", lambda _e: self._refresh_protokoll())
        theme.PillButton(self._log_manual, "Anwenden", self._refresh_protokoll,
                         kind="neutral", font_size=9, padx=10, pady=2).pack(side="left", padx=4)
        theme.PillButton(bar, "⬇  Export CSV", self._log_export, kind="neutral",
                         font_size=10).pack(side="right")

        self._log_count = tk.Label(parent, text="", bg=SHELL_BG, fg=MUTED, font=(theme.FONT, 9))
        self._log_count.grid(row=1, column=0, sticky="w", pady=(0, 4))

        outer, body = _card(parent)
        outer.grid(row=2, column=0, sticky="nsew")
        self._log_tree = _make_tree(body, ("Zeitpunkt", "Bearbeiter", "Modul", "Aktion", "Details"),
                                    (150, 110, 130, 190, 250))
        self._refreshers["protokoll"] = self._refresh_protokoll

    def _on_zeitraum_change(self):
        if self._log_zeitraum.get() == "Manuell":
            self._log_manual.pack(side="left", padx=(6, 0))
            # sinnvolle Vorbelegung: aktueller Monat
            if not self._log_von.get():
                self._log_von.set(HEUTE().replace(day=1).isoformat())
                self._log_bis.set(_heute_iso())
        else:
            self._log_manual.pack_forget()
        self._refresh_protokoll()

    def _zeitraum_grenzen(self):
        """Liefert (start_iso, end_iso) als 'JJJJ-MM-TT' fuer den gewaehlten
        Zeitraum, oder (None, None) fuer 'Alle'."""
        preset = self._log_zeitraum.get()
        t = HEUTE()
        def monatsende(y, m):
            return (date(y, 12, 31) if m == 12 else date(y, m + 1, 1) - timedelta(days=1))
        if preset == "Dieser Monat":
            return t.replace(day=1).isoformat(), monatsende(t.year, t.month).isoformat()
        if preset == "Dieses Quartal":
            sm = ((t.month - 1) // 3) * 3 + 1
            return date(t.year, sm, 1).isoformat(), monatsende(t.year, sm + 2).isoformat()
        if preset == "Dieses Jahr":
            return date(t.year, 1, 1).isoformat(), date(t.year, 12, 31).isoformat()
        if preset == "Vorjahr":
            return date(t.year - 1, 1, 1).isoformat(), date(t.year - 1, 12, 31).isoformat()
        if preset == "Manuell":
            return (self._log_von.get().strip() or None), (self._log_bis.get().strip() or None)
        return None, None

    def _log_where(self):
        """Baut (clauses, params) aus Modul- und Zeitraum-Filter."""
        clauses, params = [], []
        modul = self._log_filter.get()
        if modul != "alle":
            clauses.append("modul=?")
            params.append(modul)
        start, end = self._zeitraum_grenzen()
        if start:
            clauses.append("substr(zeitpunkt,1,10) >= ?")
            params.append(start[:10])
        if end:
            clauses.append("substr(zeitpunkt,1,10) <= ?")
            params.append(end[:10])
        return clauses, params

    def _refresh_protokoll(self):
        t = getattr(self, "_log_tree", None)
        if t is None:
            return
        t.delete(*t.get_children())
        clauses, params = self._log_where()
        where = ("WHERE " + " AND ".join(clauses)) if clauses else ""
        with self._conn() as con:
            rows = con.execute(
                f"SELECT zeitpunkt,bearbeiter,modul,aktion,details FROM tbl_gdp_log "
                f"{where} ORDER BY id DESC LIMIT 1000", params).fetchall()
        for i, r in enumerate(rows):
            t.insert("", "end", values=r, tags=("even" if i % 2 else "odd",))
        zr = self._log_zeitraum.get()
        start, end = self._zeitraum_grenzen()
        spanne = f" · {start} bis {end}" if start or end else ""
        self._log_count.config(text=f"{len(rows)} Eintrag/Einträge · Zeitraum: {zr}{spanne}")

    def _log_export(self):
        clauses, params = self._log_where()
        where = ("WHERE " + " AND ".join(clauses)) if clauses else ""
        with self._conn() as con:
            rows = con.execute(
                f"SELECT zeitpunkt,bearbeiter,modul,aktion,details FROM tbl_gdp_log "
                f"{where} ORDER BY id", params).fetchall()
        if not rows:
            messagebox.showinfo("Export", "Keine Protokoll-Eintraege im gewaehlten Filter.", parent=self)
            return
        path = filedialog.asksaveasfilename(
            parent=self, title="Protokoll exportieren", defaultextension=".csv",
            initialfile=f"GDP_Protokoll_{_heute_iso()}.csv", filetypes=[("CSV", "*.csv")])
        if not path:
            return
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
