"""NMG Meldungen - GDP-Meldewesen, Kuehlsachenkontrolle & Selbstinspektion.

Eigenstaendige App (eigenes Fenster / Taskleisten-Icon, AUMID NMG.Meldungen),
teilt sich die NMGone-Datenbank (app/config.py). Sie buendelt die Qualitaets-
und Ueberwachungsaufgaben, die vorher unuebersichtlich in der Wareneingang-App
(GDP) steckten, an einem Ort:

Module (Sidebar):
  - Uebersicht           : offene Meldungen + Pflichten auf einen Blick (Ampel)
  - Meldungen            : Abweichungs-/Qualitaetsmeldungen mit Workflow + CAPA
  - Kuehlsachenkontrolle : Temperaturueberwachung mit Soll-Grenzen + Massnahmen
  - Selbstinspektion     : GDP-Checklisten mit Massnahmen (CAPA)
  - Protokoll            : revisionssicheres Aenderungs-Log (geteilt mit GDP)

Die eigentliche Datenhaltung fuer Temperatur/Inspektion liegt in den geteilten
GDP-Tabellen (tbl_gdp_temperatur, tbl_gdp_messpunkt, tbl_gdp_inspektion, ...);
die UI dafuer ist von der GDP-App hierher gewandert. Die Meldungen selbst
liegen in tbl_gdp_meldung.

run_standalone() wird von start_meldungen.py und NMGone.exe --meldungen genutzt.
"""
from __future__ import annotations

import csv
from .i18n import T as _T
import os
import sqlite3
from datetime import datetime, timedelta
from pathlib import Path
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

from .config import DB_PATH, ASSETS_DIR
from . import theme

# Wiederverwendbare Bausteine + geteilte GDP-Tabellen aus der GDP-App.
from .gdp_app import (
    ensure_gdp_tables, _card, _make_tree, FormDialog,
    _heute_iso, HEUTE, ME,
    BG, SHELL_BG, ACCENT, ACCENT_LIGHT, BORDER,
    TEXT, MUTED, OK_GREEN, WARN, DANGER,
    SIDEBAR, SIDEBAR_ACTIVE, SIDEBAR_TEXT, SIDEBAR_MUTED,
)

try:
    from . import tour
except Exception:  # tour ist optional
    tour = None


# Status -> Treeview-Tag fuer Meldungen
MELD_STATUS_TAG = {
    "Offen": "warn", "In Bearbeitung": "warn",
    "Erledigt": "ok", "Verworfen": "alert",
}

MELD_TYPEN = [
    "Abweichung", "Qualitaetsmangel", "Temperaturbruch", "Reklamation",
    "Faelschungsverdacht / securPharm", "Transportschaden", "Lieferengpass",
    "Sonstiges",
]
MELD_PRIO = ["Niedrig", "Mittel", "Hoch", "Kritisch"]
PRIO_TAG = {"Kritisch": "alert", "Hoch": "warn"}


def ensure_meldung_tables(db_path=DB_PATH):
    """Geteilte GDP-Tabellen + die Meldungs-Tabelle idempotent anlegen."""
    ensure_gdp_tables(db_path)
    with sqlite3.connect(db_path) as con:
        con.executescript(
            """
            CREATE TABLE IF NOT EXISTS tbl_gdp_meldung(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                datum TEXT, typ TEXT, titel TEXT, prioritaet TEXT,
                pzn TEXT, charge TEXT, betrifft TEXT, beschreibung TEXT,
                status TEXT DEFAULT 'Offen', massnahme TEXT,
                verantwortlich TEXT, faellig_am TEXT, erledigt_am TEXT,
                gemeldet_von TEXT, erstellt_am TEXT DEFAULT CURRENT_TIMESTAMP
            );
            """
        )
        con.commit()


# ════════════════════════════════════════════════════════════════════════════
class MeldungenPanel(tk.Frame):
    """Hauptoberflaeche der Meldungen-App (Sidebar + Seiten)."""

    NAV = (
        ("uebersicht",  "Übersicht",            "\U0001F4CA"),
        ("meldungen",   "Meldungen",            "\U0001F514"),
        ("kuehlkette",  "Kühlsachenkontrolle",  "\U0001F321"),
        ("inspektion",  "Selbstinspektion",     "✅"),
        ("protokoll",   "Protokoll",            "\U0001F4DD"),
    )
    TITEL = {k: t for k, t, _ in NAV}
    UNTERTITEL = {
        "uebersicht": "Offene Meldungen und Überwachungspflichten auf einen Blick.",
        "meldungen": "Abweichungen, Qualitätsmängel & Reklamationen von der Meldung bis zur Maßnahme.",
        "kuehlkette": "Temperaturüberwachung der Kühlsachen mit Soll-Grenzen und Maßnahmen.",
        "inspektion": "GDP-Selbstinspektionen mit Checklisten und Maßnahmen (CAPA).",
        "protokoll": "Revisionssicheres Protokoll aller Vorgänge (geteilt mit Wareneingang).",
    }

    def __init__(self, master, db_path=DB_PATH, on_close=None, nmgone_action=None):
        super().__init__(master, bg=SHELL_BG)
        self.db_path = db_path
        self._on_close = on_close or (lambda: self.winfo_toplevel().destroy())
        self._nmgone_action = nmgone_action
        ensure_meldung_tables(db_path)
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

        # ---- Sidebar (zentrale theme.Sidebar) ----
        self.sidebar = theme.Sidebar(self, width=256, title="Meldungen",
                                     subtitle="& Qualität")
        self.sidebar.grid(row=0, column=0, sticky="ns")
        self._app_icon = theme.load_icon(ASSETS_DIR / "Meldungen.ico", 56)
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
        self.sidebar.add_footer_note(_T('Datenbank:\n{p0}', p0=Path(self.db_path).name))

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
            "meldungen": self._build_meldungen,
            "kuehlkette": self._build_kuehlkette,
            "inspektion": self._build_inspektion,
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
            messagebox.showerror("NMGone", _T('NMGone konnte nicht gestartet werden:\n{p0}', p0=exc), parent=self)

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

        outer_a, self._alert_body = _card(body, "Offene Pflichten", "Was als nächstes ansteht")
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
        offene_status = "('Erledigt','Verworfen')"
        with self._conn() as con:
            meld_offen = con.execute(
                "SELECT COUNT(*) FROM tbl_gdp_meldung WHERE status NOT IN " + offene_status).fetchone()[0]
            kritisch = con.execute(
                "SELECT COUNT(*) FROM tbl_gdp_meldung WHERE prioritaet='Kritisch' "
                "AND status NOT IN " + offene_status).fetchone()[0]
            ueberfaellig = con.execute(
                "SELECT COUNT(*) FROM tbl_gdp_meldung WHERE faellig_am IS NOT NULL AND faellig_am<>'' "
                "AND faellig_am < ? AND status NOT IN " + offene_status, (_heute_iso(),)).fetchone()[0]
            temp_offen = con.execute(
                "SELECT COUNT(*) FROM tbl_gdp_temperatur WHERE status='Abweichung' AND behoben=0").fetchone()[0]
            insp_faellig = con.execute(
                "SELECT COUNT(*) FROM tbl_gdp_inspektion WHERE naechste_faellig IS NOT NULL "
                "AND naechste_faellig <= ? AND status<>'abgeschlossen'", (_heute_iso(),)).fetchone()[0]

        for w in self._kpi_row.winfo_children():
            w.destroy()
        tiles = [
            (meld_offen, "Offene Meldungen", WARN if meld_offen else OK_GREEN),
            (kritisch, "Kritische offen", DANGER if kritisch else OK_GREEN),
            (ueberfaellig, "Meldungen überfällig", DANGER if ueberfaellig else OK_GREEN),
            (temp_offen, "Temperatur-Abweichungen offen", DANGER if temp_offen else OK_GREEN),
            (insp_faellig, "Inspektionen fällig", WARN if insp_faellig else OK_GREEN),
        ]
        for i, (v, lbl, col) in enumerate(tiles):
            self._kpi_row.columnconfigure(i, weight=1)
            self._kpi_tile(self._kpi_row, v, lbl, col).grid(row=0, column=i, sticky="nsew", padx=4)

        for w in self._alert_body.winfo_children():
            w.destroy()
        alerts = []
        if kritisch:
            alerts.append((DANGER, f"{kritisch} kritische Meldung(en) offen – sofort bearbeiten.", "meldungen"))
        if temp_offen:
            alerts.append((DANGER, f"{temp_offen} offene Temperatur-Abweichung(en) – Maßnahme dokumentieren.", "kuehlkette"))
        if ueberfaellig:
            alerts.append((WARN, f"{ueberfaellig} Meldung(en) sind überfällig.", "meldungen"))
        if insp_faellig:
            alerts.append((WARN, f"{insp_faellig} Selbstinspektion(en) fällig.", "inspektion"))
        if meld_offen and not kritisch and not ueberfaellig:
            alerts.append((WARN, f"{meld_offen} Meldung(en) in Bearbeitung.", "meldungen"))
        if not alerts:
            tk.Label(self._alert_body, text="✔  Alles erledigt – keine offenen Meldungen oder Pflichten.",
                     bg=BG, fg=OK_GREEN, font=(theme.FONT, 11, "bold")).pack(anchor="w", pady=8)
        for col, text, target in alerts:
            row = tk.Frame(self._alert_body, bg=BG)
            row.pack(fill="x", pady=3)
            tk.Frame(row, bg=col, width=5).pack(side="left", fill="y")
            tk.Label(row, text=text, bg=BG, fg=TEXT, font=(theme.FONT, 10),
                     anchor="w", justify="left", wraplength=420).pack(side="left", padx=8)
            theme.PillButton(row, "öffnen", lambda t=target: self._show_view(t),
                             kind="neutral", font_size=9, padx=10, pady=4).pack(side="right")

    # =============================================================== MELDUNGEN
    def _build_meldungen(self, parent):
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(1, weight=1)
        bar = tk.Frame(parent, bg=SHELL_BG)
        bar.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        theme.PillButton(bar, "➕  Neue Meldung", self._meld_new,
                         kind="primary", font_size=10).pack(side="left")
        theme.PillButton(bar, "▶  In Bearbeitung", lambda: self._meld_status("In Bearbeitung"),
                         kind="neutral", font_size=10).pack(side="left", padx=6)
        theme.PillButton(bar, "✅  Erledigt", lambda: self._meld_status("Erledigt"),
                         kind="success", font_size=10).pack(side="left", padx=6)
        theme.PillButton(bar, "\U0001F527  Maßnahme / CAPA", self._meld_massnahme,
                         kind="accent", font_size=10).pack(side="left", padx=6)
        theme.PillButton(bar, "⬇  Export CSV", self._meld_export,
                         kind="neutral", font_size=10).pack(side="left", padx=6)

        filt = tk.Frame(parent, bg=SHELL_BG)
        filt.grid(row=0, column=0, sticky="e")
        tk.Label(filt, text="Status:", bg=SHELL_BG, fg=TEXT, font=(theme.FONT, 10)).pack(side="left")
        self._meld_filter = tk.StringVar(value="offen")
        cb = ttk.Combobox(filt, textvariable=self._meld_filter, state="readonly", width=16,
                          style="NMG.TCombobox",
                          values=["offen", "alle", "Offen", "In Bearbeitung", "Erledigt", "Verworfen"])
        cb.pack(side="left", padx=8)
        cb.bind("<<ComboboxSelected>>", lambda _e: self._refresh_meldungen())

        outer, body = _card(parent)
        outer.grid(row=1, column=0, sticky="nsew")
        cols = ("Datum", "Typ", "Titel", "Prio", "Betrifft", "Verantwortlich", "Fällig", "Status")
        self._meld_tree = _make_tree(body, cols, (85, 130, 200, 70, 140, 120, 85, 110),
                                     anchors={"Prio": "center"})
        self._meld_tree.bind("<Double-1>", lambda _e: self._meld_detail())
        self._refreshers["meldungen"] = self._refresh_meldungen

    def _refresh_meldungen(self):
        t = self._meld_tree
        t.delete(*t.get_children())
        flt = self._meld_filter.get()
        sql = ("SELECT id,datum,typ,titel,prioritaet,betrifft,verantwortlich,faellig_am,status "
               "FROM tbl_gdp_meldung")
        params = ()
        if flt == "offen":
            sql += " WHERE status NOT IN ('Erledigt','Verworfen')"
        elif flt != "alle":
            sql += " WHERE status=?"
            params = (flt,)
        sql += " ORDER BY id DESC"
        with self._conn() as con:
            rows = con.execute(sql, params).fetchall()
        heute = _heute_iso()
        for i, (mid, datum, typ, titel, prio, betrifft, verant, faellig, status) in enumerate(rows):
            tag = MELD_STATUS_TAG.get(status, "even" if i % 2 else "odd")
            if status not in ("Erledigt", "Verworfen"):
                if faellig and faellig < heute:
                    tag = "alert"
                elif prio == "Kritisch":
                    tag = "alert"
            t.insert("", "end", iid=str(mid),
                     values=(datum or "", typ or "", titel or "", prio or "", betrifft or "",
                             verant or "", faellig or "", status or ""),
                     tags=(tag,))
        if not rows:
            t.insert("", "end", values=("-", "Keine Meldungen für diesen Filter", "", "", "", "", "", ""))

    def _meld_selected_id(self):
        sel = self._meld_tree.selection()
        if not sel:
            return None
        try:
            return int(sel[0])
        except ValueError:
            return None

    def _meld_new(self):
        def save(v):
            faellig = v["faellig_am"].strip()
            with self._conn() as con:
                cur = con.execute(
                    """INSERT INTO tbl_gdp_meldung
                       (datum,typ,titel,prioritaet,pzn,charge,betrifft,beschreibung,
                        status,verantwortlich,faellig_am,gemeldet_von)
                       VALUES(?,?,?,?,?,?,?,?,?,?,?,?)""",
                    (_heute_iso(), v["typ"], v["titel"], v["prioritaet"], v["pzn"], v["charge"],
                     v["betrifft"], v["beschreibung"], "Offen", v["verantwortlich"], faellig, ME()))
                mid = cur.lastrowid
                con.commit()
            self._log("Meldung", "Meldung angelegt", mid, f"{v['typ']}: {v['titel']}")
            self._refresh_meldungen()

        FormDialog(self, "Neue Meldung", [
            ("typ", "Art der Meldung", "combo", MELD_TYPEN, MELD_TYPEN[0]),
            ("titel", "Titel / Kurzfassung", "text", ""),
            ("prioritaet", "Priorität", "combo", MELD_PRIO, "Mittel"),
            ("betrifft", "Betrifft (Artikel/Lieferant/Apotheke)", "text", ""),
            ("pzn", "PZN (optional)", "text", ""),
            ("charge", "Charge (optional)", "text", ""),
            ("verantwortlich", "Verantwortlich", "text", ME()),
            ("faellig_am", "Fällig bis (JJJJ-MM-TT)", "text",
             (HEUTE() + timedelta(days=14)).isoformat()),
            ("beschreibung", "Beschreibung", "multiline", ""),
        ], save, width=500)

    def _meld_status(self, neu):
        mid = self._meld_selected_id()
        if not mid:
            messagebox.showinfo("Meldung", "Bitte zuerst eine Meldung in der Liste wählen.", parent=self)
            return
        erledigt = _heute_iso() if neu in ("Erledigt", "Verworfen") else None
        with self._conn() as con:
            if neu == "Erledigt":
                offen = con.execute(
                    "SELECT COALESCE(massnahme,'') FROM tbl_gdp_meldung WHERE id=?", (mid,)).fetchone()
                if offen is not None and not (offen[0] or "").strip():
                    if not messagebox.askyesno(
                            "Erledigt",
                            "Zu dieser Meldung ist noch keine Maßnahme/CAPA dokumentiert.\n"
                            "Trotzdem als erledigt markieren?", parent=self):
                        return
            con.execute("UPDATE tbl_gdp_meldung SET status=?, erledigt_am=? WHERE id=?",
                        (neu, erledigt, mid))
            con.commit()
        self._log("Meldung", f"Status: {neu}", mid, "")
        self._refresh_meldungen()

    def _meld_massnahme(self):
        mid = self._meld_selected_id()
        if not mid:
            messagebox.showinfo("Maßnahme", "Bitte zuerst eine Meldung wählen.", parent=self)
            return
        with self._conn() as con:
            row = con.execute("SELECT massnahme, verantwortlich FROM tbl_gdp_meldung WHERE id=?",
                              (mid,)).fetchone()
        alt_massn = (row[0] if row else "") or ""
        alt_verant = (row[1] if row else "") or ME()

        def save(v):
            with self._conn() as con:
                con.execute(
                    "UPDATE tbl_gdp_meldung SET massnahme=?, verantwortlich=?, "
                    "status=CASE WHEN status='Offen' THEN 'In Bearbeitung' ELSE status END WHERE id=?",
                    (v["massnahme"], v["verantwortlich"], mid))
                con.commit()
            self._log("Meldung", "Maßnahme/CAPA dokumentiert", mid, "")
            self._refresh_meldungen()

        FormDialog(self, "Maßnahme / CAPA", [
            ("massnahme", "Maßnahme / CAPA", "multiline", alt_massn),
            ("verantwortlich", "Verantwortlich", "text", alt_verant),
        ], save, width=480)

    def _meld_detail(self):
        mid = self._meld_selected_id()
        if not mid:
            return
        with self._conn() as con:
            cols = [c[1] for c in con.execute("PRAGMA table_info(tbl_gdp_meldung)")]
            row = con.execute("SELECT * FROM tbl_gdp_meldung WHERE id=?", (mid,)).fetchone()
            if not row:
                return
            log = con.execute(
                "SELECT zeitpunkt, bearbeiter, aktion, details FROM tbl_gdp_log "
                "WHERE modul='Meldung' AND bezug_id=? ORDER BY id", (mid,)).fetchall()
        d = dict(zip(cols, row))
        win = tk.Toplevel(self)
        win.title(f"Meldung #{mid}")
        win.configure(bg=SHELL_BG)
        win.geometry("580x560")
        tk.Label(win, bg=ACCENT, fg="#FFFFFF", anchor="w", font=(theme.FONT, 12, "bold"),
                 text=f"  {d['typ']} #{mid}  ·  {d['status']}  ·  {d['prioritaet']}").pack(fill="x", ipady=8)
        info = tk.Frame(win, bg=SHELL_BG)
        info.pack(fill="x", padx=16, pady=10)
        pairs = [("Datum", d["datum"]), ("Titel", d["titel"]),
                 ("Betrifft", d["betrifft"]), ("PZN", d["pzn"] or "-"),
                 ("Charge", d["charge"] or "-"), ("Verantwortlich", d["verantwortlich"] or "-"),
                 ("Fällig", d["faellig_am"] or "-"), ("Erledigt am", d["erledigt_am"] or "-"),
                 ("Gemeldet von", d["gemeldet_von"] or "-")]
        for i, (k, val) in enumerate(pairs):
            r, c = divmod(i, 2)
            cell = tk.Frame(info, bg=SHELL_BG)
            cell.grid(row=r, column=c, sticky="w", padx=6, pady=2)
            tk.Label(cell, text=f"{k}: ", bg=SHELL_BG, fg=MUTED, font=(theme.FONT, 9, "bold")).pack(side="left")
            tk.Label(cell, text=str(val), bg=SHELL_BG, fg=TEXT, font=(theme.FONT, 9)).pack(side="left")
        for titel, val in (("Beschreibung", d["beschreibung"]), ("Maßnahme / CAPA", d["massnahme"])):
            tk.Label(win, text=titel, bg=SHELL_BG, fg=ACCENT,
                     font=(theme.FONT, 10, "bold")).pack(anchor="w", padx=16, pady=(8, 0))
            tk.Label(win, text=(val or "-"), bg=SHELL_BG, fg=TEXT, font=(theme.FONT, 9),
                     anchor="w", justify="left", wraplength=540).pack(anchor="w", padx=16)
        tk.Label(win, text="Verlauf", bg=SHELL_BG, fg=ACCENT,
                 font=(theme.FONT, 10, "bold")).pack(anchor="w", padx=16, pady=(8, 2))
        lw = tk.Frame(win, bg=SHELL_BG)
        lw.pack(fill="both", expand=True, padx=16, pady=(0, 12))
        tree = _make_tree(lw, ("Zeit", "Wer", "Aktion", "Details"), (130, 90, 160, 150), height=5)
        for i, lr in enumerate(log):
            tree.insert("", "end", values=lr, tags=("even" if i % 2 else "odd",))

    def _meld_export(self):
        path = filedialog.asksaveasfilename(
            parent=self, title="Meldungen exportieren", defaultextension=".csv",
            initialfile=f"Meldungen_{_heute_iso()}.csv", filetypes=[("CSV", "*.csv")])
        if not path:
            return
        with self._conn() as con:
            rows = con.execute(
                "SELECT datum,typ,titel,prioritaet,betrifft,pzn,charge,verantwortlich,"
                "faellig_am,status,erledigt_am,massnahme,beschreibung FROM tbl_gdp_meldung "
                "ORDER BY id").fetchall()
        with open(path, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f, delimiter=";")
            w.writerow(["Datum", "Typ", "Titel", "Prioritaet", "Betrifft", "PZN", "Charge",
                        "Verantwortlich", "Faellig", "Status", "Erledigt am", "Massnahme", "Beschreibung"])
            w.writerows(rows)
        messagebox.showinfo("Export", _T('{p0} Meldung(en) gespeichert:\n{p1}', p0=len(rows), p1=path), parent=self)

    # ===================================================== KÜHLSACHENKONTROLLE
    def _build_kuehlkette(self, parent):
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(2, weight=1)
        bar = tk.Frame(parent, bg=SHELL_BG)
        bar.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        theme.PillButton(bar, "\U0001F321  Messung erfassen", self._temp_new,
                         kind="primary", font_size=10).pack(side="left")
        theme.PillButton(bar, "➕  Messpunkt", self._messpunkt_new,
                         kind="neutral", font_size=10).pack(side="left", padx=6)
        theme.PillButton(bar, "\U0001F527  Abweichung bearbeiten", self._temp_massnahme,
                         kind="accent", font_size=10).pack(side="left", padx=6)

        mp_outer, mp_body = _card(parent, "Messpunkte", "Soll-Temperaturbereiche")
        mp_outer.grid(row=1, column=0, sticky="ew", pady=(0, 8))
        self._mp_tree = _make_tree(mp_body, ("Messpunkt", "Typ", "Soll min", "Soll max", "Letzte", "Status"),
                                   (210, 120, 70, 70, 90, 110), height=5,
                                   anchors={"Soll min": "center", "Soll max": "center", "Letzte": "center"})
        outer, body = _card(parent, "Messungen", "Neueste zuerst")
        outer.grid(row=2, column=0, sticky="nsew")
        self._temp_tree = _make_tree(body, ("Zeit", "Messpunkt", "Temp C", "Status", "Erfasst von", "Massnahme"),
                                     (140, 200, 70, 100, 110, 180),
                                     anchors={"Temp C": "center"})
        self._refreshers["kuehlkette"] = self._refresh_kuehlkette

    def _refresh_kuehlkette(self):
        with self._conn() as con:
            mps = con.execute(
                "SELECT id, name, typ, soll_min, soll_max FROM tbl_gdp_messpunkt WHERE aktiv=1 ORDER BY id"
            ).fetchall()
            last = {row[0]: (row[1], row[2]) for row in con.execute(
                "SELECT messpunkt_id, temp_c, status FROM tbl_gdp_temperatur t WHERE id IN "
                "(SELECT MAX(id) FROM tbl_gdp_temperatur GROUP BY messpunkt_id)")}
            temps = con.execute(
                """SELECT t.zeitpunkt, m.name, t.temp_c, t.status, t.erfasst_von, t.massnahme, t.behoben
                   FROM tbl_gdp_temperatur t LEFT JOIN tbl_gdp_messpunkt m ON m.id=t.messpunkt_id
                   ORDER BY t.id DESC LIMIT 200""").fetchall()
        mt = self._mp_tree
        mt.delete(*mt.get_children())
        for i, (mid, name, typ, smin, smax) in enumerate(mps):
            lt, lst = last.get(mid, (None, None))
            tag = "even" if i % 2 else "odd"
            if lst == "Abweichung":
                tag = "alert"
            elif lst == "ok":
                tag = "ok"
            mt.insert("", "end", iid=f"mp{mid}",
                      values=(name, typ, smin, smax,
                              f"{lt:.1f}" if lt is not None else "-", lst or "keine"),
                      tags=(tag,))
        tt = self._temp_tree
        tt.delete(*tt.get_children())
        for i, (zeit, name, temp, status, von, massn, behoben) in enumerate(temps):
            if status == "Abweichung" and not behoben:
                tag = "alert"
            elif status == "Abweichung":
                tag = "warn"
            else:
                tag = "even" if i % 2 else "odd"
            mdisp = (massn or "") + (" ✔" if behoben else "")
            tt.insert("", "end", values=(zeit or "", name or "?", f"{temp:.1f}" if temp is not None else "",
                                         status or "", von or "", mdisp),
                      tags=(tag,))

    def _temp_new(self):
        with self._conn() as con:
            mps = con.execute("SELECT id, name, soll_min, soll_max FROM tbl_gdp_messpunkt "
                              "WHERE aktiv=1 ORDER BY id").fetchall()
        if not mps:
            messagebox.showinfo("Messung", "Bitte zuerst einen Messpunkt anlegen.", parent=self)
            return
        labels = [f"{n}" for _i, n, _a, _b in mps]
        mp_map = {n: (mid, smin, smax) for mid, n, smin, smax in mps}

        def save(v):
            mid, smin, smax = mp_map[v["messpunkt"]]
            temp = v["temp_c"]
            status = "ok" if (smin <= temp <= smax) else "Abweichung"
            with self._conn() as con:
                cur = con.execute(
                    "INSERT INTO tbl_gdp_temperatur(messpunkt_id,zeitpunkt,temp_c,status,erfasst_von,notiz,behoben) "
                    "VALUES(?,?,?,?,?,?,0)",
                    (mid, datetime.now().isoformat(timespec="minutes"), temp, status, ME(), v["notiz"]))
                tid = cur.lastrowid
                con.commit()
            self._log("Kuehlkette", f"Messung {status}", tid, f"{v['messpunkt']}: {temp} C")
            if status == "Abweichung":
                messagebox.showwarning("Abweichung",
                    _T('{p0} C liegt außerhalb {p1}-{p2} C!\nBitte Maßnahme dokumentieren (Abweichung bearbeiten).', p0=temp, p1=smin, p2=smax), parent=self)
            self._refresh_kuehlkette()

        FormDialog(self, "Temperatur-Messung erfassen", [
            ("messpunkt", "Messpunkt", "combo", labels, labels[0]),
            ("temp_c", "Temperatur (C)", "float", "5.0"),
            ("notiz", "Notiz", "multiline", ""),
        ], save)

    def _messpunkt_new(self):
        def save(v):
            with self._conn() as con:
                con.execute("INSERT INTO tbl_gdp_messpunkt(name,typ,soll_min,soll_max,aktiv) "
                            "VALUES(?,?,?,?,1)",
                            (v["name"], v["typ"], v["soll_min"], v["soll_max"]))
                con.commit()
            self._log("Kuehlkette", "Messpunkt angelegt", None, v["name"])
            self._refresh_kuehlkette()
        FormDialog(self, "Neuer Messpunkt", [
            ("name", "Bezeichnung", "text", "Kuehlschrank 3"),
            ("typ", "Typ", "combo", ["Kuehlschrank", "Lager", "Transport", "Tiefkuehl"], "Kuehlschrank"),
            ("soll_min", "Soll min (C)", "float", "2.0"),
            ("soll_max", "Soll max (C)", "float", "8.0"),
        ], save)

    def _temp_massnahme(self):
        with self._conn() as con:
            offen = con.execute(
                "SELECT id, temp_c, status FROM tbl_gdp_temperatur WHERE status='Abweichung' AND behoben=0 "
                "ORDER BY id DESC").fetchall()
        if not offen:
            messagebox.showinfo("Maßnahme", "Keine offenen Temperatur-Abweichungen.", parent=self)
            return
        tid, temp, _ = offen[0]

        def save(v):
            with self._conn() as con:
                con.execute("UPDATE tbl_gdp_temperatur SET massnahme=?, behoben=? WHERE id=?",
                            (v["massnahme"], v["behoben"], tid))
                con.commit()
            self._log("Kuehlkette", "Massnahme dokumentiert", tid, v["massnahme"])
            self._refresh_kuehlkette()
        FormDialog(self, f"Abweichung bearbeiten ({temp} C)", [
            ("massnahme", "Maßnahme / CAPA", "multiline", ""),
            ("behoben", "Als behoben markieren", "check", 1),
        ], save)

    # ===================================================== SELBSTINSPEKTION
    INSPEKTION_TEMPLATE = [
        ("Raeumlichkeiten", "Lager sauber, trocken, abschliessbar?"),
        ("Raeumlichkeiten", "Wareneingang/-ausgang getrennt?"),
        ("Temperatur", "Kuehlkette luekenlos dokumentiert?"),
        ("Temperatur", "Messgeraete kalibriert?"),
        ("Dokumentation", "Chargen rueckverfolgbar (Kunde<->Charge)?"),
        ("Dokumentation", "Lieferantenqualifizierung aktuell?"),
        ("Kunden", "Nur lizenzierte Apotheken beliefert?"),
        ("Retouren", "Retouren-Verfahren eingehalten?"),
        ("Retouren", "Vernichtung dokumentiert?"),
        ("Personal", "Schulungen GDP aktuell?"),
        ("Faelschungsschutz", "securPharm / Verifizierung aktiv?"),
        ("Selbstinspektion", "Massnahmen aus letzter Inspektion erledigt?"),
    ]

    def _build_inspektion(self, parent):
        parent.columnconfigure(0, weight=1)
        parent.columnconfigure(1, weight=2)
        parent.rowconfigure(1, weight=1)
        bar = tk.Frame(parent, bg=SHELL_BG)
        bar.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 8))
        theme.PillButton(bar, "➕  Neue Inspektion (Vorlage)", self._insp_new,
                         kind="primary", font_size=10).pack(side="left")
        theme.PillButton(bar, "✅  Punkt: ok / Abweichung", self._insp_toggle,
                         kind="neutral", font_size=10).pack(side="left", padx=6)
        theme.PillButton(bar, "\U0001F527  Maßnahme", self._insp_massnahme,
                         kind="accent", font_size=10).pack(side="left", padx=6)
        theme.PillButton(bar, "\U0001F4C5  Abschließen", self._insp_close,
                         kind="success", font_size=10).pack(side="left", padx=6)

        left_outer, left_body = _card(parent, "Inspektionen")
        left_outer.grid(row=1, column=0, sticky="nsew", padx=(0, 6))
        self._insp_tree = _make_tree(left_body, ("Datum", "Titel", "Status", "Fällig"),
                                     (80, 150, 100, 90), height=16)
        self._insp_tree.bind("<<TreeviewSelect>>", lambda _e: self._insp_load_punkte())

        right_outer, right_body = _card(parent, "Prüfpunkte", "Doppelklick wechselt ok/Abweichung/n.z.")
        right_outer.grid(row=1, column=1, sticky="nsew", padx=(6, 0))
        self._insp_punkt_tree = _make_tree(right_body, ("Kategorie", "Prüfpunkt", "Ergebnis", "Maßnahme"),
                                           (120, 280, 100, 160))
        self._insp_punkt_tree.bind("<Double-1>", lambda _e: self._insp_toggle())
        self._refreshers["inspektion"] = self._refresh_inspektion

    def _refresh_inspektion(self):
        t = self._insp_tree
        prev = t.selection()
        t.delete(*t.get_children())
        with self._conn() as con:
            rows = con.execute(
                "SELECT id, datum, titel, status, naechste_faellig FROM tbl_gdp_inspektion ORDER BY id DESC"
            ).fetchall()
        for i, (iid, datum, titel, status, faellig) in enumerate(rows):
            tag = "even" if i % 2 else "odd"
            if faellig and faellig <= _heute_iso() and status != "abgeschlossen":
                tag = "warn"
            t.insert("", "end", iid=str(iid), values=(datum or "", titel or "", status or "", faellig or ""),
                     tags=(tag,))
        if prev and t.exists(prev[0]):
            t.selection_set(prev[0])
        elif rows:
            t.selection_set(str(rows[0][0]))
        self._insp_load_punkte()

    def _insp_sel_id(self):
        sel = self._insp_tree.selection()
        return int(sel[0]) if sel else None

    def _insp_load_punkte(self):
        t = self._insp_punkt_tree
        t.delete(*t.get_children())
        iid = self._insp_sel_id()
        if not iid:
            return
        with self._conn() as con:
            rows = con.execute(
                "SELECT id, kategorie, frage, ergebnis, massnahme FROM tbl_gdp_inspektion_punkt "
                "WHERE inspektion_id=? ORDER BY id", (iid,)).fetchall()
        for i, (pid, kat, frage, erg, massn) in enumerate(rows):
            tag = "even" if i % 2 else "odd"
            if erg == "Abweichung":
                tag = "alert"
            elif erg == "ok":
                tag = "ok"
            t.insert("", "end", iid=str(pid), values=(kat or "", frage or "", erg or "offen", massn or ""),
                     tags=(tag,))

    def _insp_new(self):
        def save(v):
            with self._conn() as con:
                cur = con.execute(
                    "INSERT INTO tbl_gdp_inspektion(datum,titel,typ,status,durchgefuehrt_von,naechste_faellig,bemerkung) "
                    "VALUES(?,?,?,?,?,?,?)",
                    (_heute_iso(), v["titel"], v["typ"], "laeuft", ME(),
                     (HEUTE() + timedelta(days=365)).isoformat(), v["bemerkung"]))
                iid = cur.lastrowid
                con.executemany(
                    "INSERT INTO tbl_gdp_inspektion_punkt(inspektion_id,kategorie,frage,ergebnis) VALUES(?,?,?,'offen')",
                    [(iid, kat, frage) for kat, frage in self.INSPEKTION_TEMPLATE])
                con.commit()
            self._log("Inspektion", "Inspektion angelegt", iid, v["titel"])
            self._refresh_inspektion()
        FormDialog(self, "Neue Selbstinspektion", [
            ("titel", "Titel", "text", f"GDP-Selbstinspektion {HEUTE().year}"),
            ("typ", "Typ", "combo", ["Selbstinspektion", "Audit", "Behoerdeninspektion"], "Selbstinspektion"),
            ("bemerkung", "Bemerkung", "multiline", ""),
        ], save)

    def _insp_toggle(self):
        sel = self._insp_punkt_tree.selection()
        if not sel:
            return
        pid = int(sel[0])
        order = {"offen": "ok", "ok": "Abweichung", "Abweichung": "n.z.", "n.z.": "offen"}
        with self._conn() as con:
            cur = con.execute("SELECT ergebnis FROM tbl_gdp_inspektion_punkt WHERE id=?", (pid,)).fetchone()
            new = order.get((cur[0] or "offen"), "ok")
            con.execute("UPDATE tbl_gdp_inspektion_punkt SET ergebnis=? WHERE id=?", (new, pid))
            con.commit()
        self._insp_load_punkte()

    def _insp_massnahme(self):
        sel = self._insp_punkt_tree.selection()
        if not sel:
            messagebox.showinfo("Maßnahme", "Bitte einen Prüfpunkt wählen.", parent=self)
            return
        pid = int(sel[0])

        def save(v):
            with self._conn() as con:
                con.execute("UPDATE tbl_gdp_inspektion_punkt SET massnahme=? WHERE id=?",
                            (v["massnahme"], pid))
                con.commit()
            self._insp_load_punkte()
        FormDialog(self, "Maßnahme zum Prüfpunkt", [
            ("massnahme", "Maßnahme / CAPA", "multiline", ""),
        ], save)

    def _insp_close(self):
        iid = self._insp_sel_id()
        if not iid:
            return
        with self._conn() as con:
            offen = con.execute(
                "SELECT COUNT(*) FROM tbl_gdp_inspektion_punkt WHERE inspektion_id=? AND ergebnis='offen'",
                (iid,)).fetchone()[0]
        if offen:
            if not messagebox.askyesno("Abschließen",
                    _T('Es sind noch {p0} Prüfpunkt(e) offen. Trotzdem abschließen?', p0=offen), parent=self):
                return
        self._insp_set(iid)

    def _insp_set(self, iid):
        with self._conn() as con:
            con.execute("UPDATE tbl_gdp_inspektion SET status='abgeschlossen' WHERE id=?", (iid,))
            con.commit()
        self._log("Inspektion", "Inspektion abgeschlossen", iid, "")
        self._refresh_inspektion()

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
                          values=["alle", "Meldung", "Kuehlkette", "Inspektion"])
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
                    "WHERE modul IN ('Meldung','Kuehlkette','Inspektion') "
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
            initialfile=f"Meldungen_Protokoll_{_heute_iso()}.csv", filetypes=[("CSV", "*.csv")])
        if not path:
            return
        with self._conn() as con:
            rows = con.execute(
                "SELECT zeitpunkt,bearbeiter,modul,aktion,details FROM tbl_gdp_log "
                "WHERE modul IN ('Meldung','Kuehlkette','Inspektion') ORDER BY id").fetchall()
        with open(path, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f, delimiter=";")
            w.writerow(["Zeitpunkt", "Bearbeiter", "Modul", "Aktion", "Details"])
            w.writerows(rows)
        messagebox.showinfo("Export", _T('{p0} Protokoll-Zeile(n) gespeichert.', p0=len(rows)), parent=self)


# ════════════════════════════════════════════════════════════════════════════
def run_standalone():
    """Startet die Meldungen-App als eigenstaendiges Fenster (eigenes Icon).
    Wird von start_meldungen.py und von NMGone.exe --meldungen genutzt."""
    if os.name == "nt":
        try:
            import ctypes
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("NMG.Meldungen")
        except Exception:
            pass
    try:
        from .migrations import run_migrations
        run_migrations()
    except Exception:
        pass
    root = tk.Tk()
    root.title("NMG Meldungen")
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
    for ico in ("Meldungen.ico", "GDP.ico"):
        try:
            root.iconbitmap(str(ASSETS_DIR / ico))
            break
        except Exception:
            continue
    MeldungenPanel(root, on_close=root.destroy).pack(fill="both", expand=True)
    if tour is not None:
        try:
            tour.maybe_show(root, "meldungen", _meldungen_tour_steps())
        except Exception:
            pass
    root.mainloop()


def _meldungen_tour_steps():
    return [
        ("Meldungen & Qualität",
         "Diese App bündelt das GDP-Meldewesen: Abweichungen, Qualitätsmängel und "
         "Reklamationen an einem Ort - statt verstreut im Wareneingang."),
        ("Meldungen erfassen",
         "Lege eine Meldung mit Typ, Priorität und Verantwortlichem an und verfolge "
         "sie von 'offen' über 'in Bearbeitung' bis 'erledigt' - inkl. Maßnahme/CAPA."),
        ("Kühlsachenkontrolle & Inspektion",
         "Temperaturen überwachen, Abweichungen mit Maßnahme dokumentieren und "
         "GDP-Selbstinspektionen per Checkliste durchführen."),
    ]


if __name__ == "__main__":
    import sys
    sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
    run_standalone()
