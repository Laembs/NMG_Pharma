"""NMG Kasse-App (ehem. Bestell-App) - gemeinsame Oberflaeche fuer NMGone
(Toplevel) und die eigenstaendige Kasse-.exe (eigenes Fenster / Taskleisten-Icon).

Warenkreislauf: Wareneingang -> Lagerbestand -> Verkauf. UI in KassePanel
(tk.Frame) mit zwei Reitern:
  - Verkauf      = Warenausgang an Apotheke (globale Kunden-/Artikelsuche,
                   Bestand/Charge-Auswahl, Rabatt-Kaskade, Speichern -> Lager ab)
  - Wareneingang = NMG-Artikel mit Charge/Verfall/Menge ins Lager buchen

Datenmodell siehe migrations.py.
"""
from __future__ import annotations

import getpass
import os
import sqlite3
import subprocess
import sys
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog, filedialog

from .config import DB_PATH, ASSETS_DIR, BASE_DIR
from . import kasse_import
from . import auftrag

BG = "#ffffff"           # Inhaltsflaechen (Karten)
SHELL_BG = "#f5f7fb"     # Fensterhintergrund wie NMGone
ACCENT = "#0b4a86"
NAV_SEL = "#e8f1fb"      # Auswahl-Highlight wie NMGone


def _table_exists(con, name) -> bool:
    return con.execute(
        "SELECT 1 FROM sqlite_master WHERE type='table' AND name=?", (name,)
    ).fetchone() is not None


def ensure_kasse_tables(db_path=DB_PATH):
    """Alle Kasse-Tabellen idempotent anlegen (spiegelt migrations.py). WAL fuer
    parallelen Zugriff von NMGone und Kasse-.exe."""
    with sqlite3.connect(db_path) as con:
        try:
            con.execute("PRAGMA journal_mode=WAL")
        except sqlite3.Error:
            pass
        con.execute(
            """CREATE TABLE IF NOT EXISTS tbl_bestellungen(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                datum TEXT, kundennummer TEXT, apotheke TEXT,
                bestellart TEXT DEFAULT 'Bestellung',
                lieferzeit TEXT, liefertermin TEXT,
                status TEXT DEFAULT 'offen', notizen TEXT, bearbeiter TEXT,
                erstellt_am TEXT DEFAULT CURRENT_TIMESTAMP, geaendert_am TEXT
            )"""
        )
        con.execute(
            """CREATE TABLE IF NOT EXISTS tbl_bestellpositionen(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                bestell_id INTEGER NOT NULL,
                pzn TEXT, artikelname TEXT, df TEXT, pck TEXT, apu REAL,
                menge INTEGER DEFAULT 1, rabatt_prozent REAL, rabatt_quelle TEXT,
                charge TEXT, verfall TEXT,
                FOREIGN KEY(bestell_id) REFERENCES tbl_bestellungen(id) ON DELETE CASCADE
            )"""
        )
        for col in ("charge", "verfall"):
            if col not in {r[1] for r in con.execute("PRAGMA table_info(tbl_bestellpositionen)")}:
                con.execute(f"ALTER TABLE tbl_bestellpositionen ADD COLUMN {col} TEXT")
        con.execute(
            """CREATE TABLE IF NOT EXISTS tbl_wareneingang(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                datum TEXT, lieferant TEXT, lieferschein TEXT,
                bearbeiter TEXT, notizen TEXT,
                erstellt_am TEXT DEFAULT CURRENT_TIMESTAMP
            )"""
        )
        con.execute(
            """CREATE TABLE IF NOT EXISTS tbl_wareneingang_positionen(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                we_id INTEGER NOT NULL,
                pzn TEXT, artikelname TEXT, charge TEXT, verfall TEXT,
                menge INTEGER DEFAULT 0, ek REAL,
                FOREIGN KEY(we_id) REFERENCES tbl_wareneingang(id) ON DELETE CASCADE
            )"""
        )
        con.execute(
            """CREATE TABLE IF NOT EXISTS tbl_lagerbestand(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                pzn TEXT, artikelname TEXT, charge TEXT, verfall TEXT,
                menge INTEGER DEFAULT 0, aktualisiert_am TEXT,
                UNIQUE(pzn, charge, verfall)
            )"""
        )
        con.commit()


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

    def set_text(self, t):
        self.var.set(t)

    def clear(self):
        self.var.set("")
        self._hide()


class KassePanel(tk.Frame):
    """Kasse-Oberflaeche: Reiter Verkauf + Wareneingang."""

    def __init__(self, master, db_path=DB_PATH, on_close=None, nmgone_action=None):
        super().__init__(master, bg=SHELL_BG)
        self.db_path = db_path
        self._on_close = on_close or (lambda: self.winfo_toplevel().destroy())
        # Aktion fuer den "NMGone oeffnen"-Button: aus NMGone heraus = Fenster nach
        # vorn holen; standalone = NMGone-Programm starten (Default unten).
        self._nmgone_action = nmgone_action
        ensure_kasse_tables(db_path)
        self.vk_kunde = None          # (kundennummer, apotheke)
        self.vk_cur = None            # aktuell gewaehlter Artikel (dict)
        self.vk_positions = []        # Liste der Verkaufspositionen (dicts)
        self.we_pzn = None
        self._build()

    # =============================================================== Datenbank
    def _conn(self):
        return sqlite3.connect(self.db_path)

    def _search_nmg(self, text, limit=25):
        like = f"%{text}%"
        with self._conn() as con:
            return con.execute(
                "SELECT pzn, artikelname FROM tbl_nmg_stamm "
                "WHERE pzn LIKE ? OR artikelname LIKE ? ORDER BY artikelname LIMIT ?",
                (like, like, limit),
            ).fetchall()

    def _artikel_details(self, pzn):
        with self._conn() as con:
            nmg = con.execute(
                "SELECT artikelname, apu, menge, einheit FROM tbl_nmg_stamm WHERE pzn=? LIMIT 1",
                (pzn,),
            ).fetchone()
            if not nmg:
                return None
            artikelname, apu, menge, einheit = nmg
            df = pck = None
            if _table_exists(con, "tbl_artikelstamm"):
                a = con.execute(
                    "SELECT df, pck FROM tbl_artikelstamm WHERE pzn=? LIMIT 1", (pzn,)
                ).fetchone()
                if a:
                    df, pck = a
            pck = pck or " ".join(str(x) for x in (menge, einheit) if x)
            return {"pzn": pzn, "artikelname": artikelname, "df": df or "",
                    "pck": pck or "", "apu": apu}

    def _search_kunden(self, text, limit=25):
        like = f"%{text}%"
        with self._conn() as con:
            if not _table_exists(con, "tbl_kunden_center"):
                return []
            have = {r[1] for r in con.execute("PRAGMA table_info(tbl_kunden_center)")}
            felder = [c for c in ("kundennummer", "kundenname", "plz", "ort",
                                  "strasse", "inhaber", "ansprechpartner") if c in have]
            if not felder:
                return []
            where = " OR ".join(f"COALESCE({c},'') LIKE ?" for c in felder)
            rows = con.execute(
                f"SELECT kundennummer, kundenname, "
                f"COALESCE(plz,''), COALESCE(ort,'') FROM tbl_kunden_center "
                f"WHERE {where} ORDER BY kundenname LIMIT ?",
                tuple(like for _ in felder) + (limit,),
            ).fetchall()
            return rows

    def _resolve_rabatt(self, kundennummer, pzn):
        """Kaskade: PK-Kondition -> NMG-Rabatt -> 0. Liefert (prozent, quelle)."""
        with self._conn() as con:
            if kundennummer and _table_exists(con, "tbl_pk_konditionen"):
                r = con.execute(
                    "SELECT rabatt_prozent FROM tbl_pk_konditionen "
                    "WHERE kundennummer=? AND pzn=? AND rabatt_prozent IS NOT NULL LIMIT 1",
                    (kundennummer, pzn),
                ).fetchone()
                if r and r[0] is not None:
                    return float(r[0]), "PK"
            r = con.execute(
                "SELECT rabatt FROM nmg_rabatte WHERE nmg_pzn=? AND rabatt IS NOT NULL LIMIT 1",
                (pzn,),
            ).fetchone()
            if r and r[0] is not None:
                val = float(r[0])
                # nmg_rabatte.rabatt ist ein Bruch (0.2 = 20 %); >1 = schon Prozent.
                return (val * 100 if val <= 1 else val), "NMG"
        return 0.0, "manuell"

    def _lager_chargen(self, pzn):
        with self._conn() as con:
            return con.execute(
                "SELECT charge, verfall, menge FROM tbl_lagerbestand "
                "WHERE pzn=? AND menge > 0 ORDER BY verfall",
                (pzn,),
            ).fetchall()

    def _save_verkauf(self, header, positions):
        with self._conn() as con:
            cur = con.execute(
                "INSERT INTO tbl_bestellungen(datum,kundennummer,apotheke,bestellart,"
                "lieferzeit,liefertermin,status,bearbeiter) VALUES(?,?,?,?,?,?,?,?)",
                (header["datum"], header["kundennummer"], header["apotheke"],
                 header["bestellart"], header["lieferzeit"], header["liefertermin"],
                 "offen", getpass.getuser()),
            )
            bid = cur.lastrowid
            for p in positions:
                con.execute(
                    "INSERT INTO tbl_bestellpositionen(bestell_id,pzn,artikelname,df,pck,"
                    "apu,menge,rabatt_prozent,rabatt_quelle,charge,verfall) "
                    "VALUES(?,?,?,?,?,?,?,?,?,?,?)",
                    (bid, p["pzn"], p["artikelname"], p["df"], p["pck"], p["apu"],
                     p["menge"], p["rabatt"], p["rabatt_quelle"], p["charge"], p["verfall"]),
                )
                # Lager abbuchen, wenn eine konkrete Charge gewaehlt wurde (locker:
                # nie unter 0, keine Sperre).
                if p["charge"] or p["verfall"]:
                    con.execute(
                        "UPDATE tbl_lagerbestand SET menge=MAX(0, menge - ?), aktualisiert_am=? "
                        "WHERE pzn=? AND COALESCE(charge,'')=? AND COALESCE(verfall,'')=?",
                        (p["menge"], datetime.now().isoformat(timespec="seconds"),
                         p["pzn"], p["charge"], p["verfall"]),
                    )
            con.commit()
            return bid

    # ================================================================== Aufbau
    def _load_logo(self):
        p = ASSETS_DIR / "NMGone.png"
        if not p.exists():
            return None
        try:
            raw = tk.PhotoImage(file=str(p))
            factor = max(1, int(max(raw.width() / 200, raw.height() / 66) + 0.999))
            self._logo_img = raw.subsample(factor, factor)
            return self._logo_img
        except Exception:
            return None

    def _build(self):
        self.configure(bg=SHELL_BG)
        self.columnconfigure(1, weight=1)
        self.rowconfigure(0, weight=1)

        # ---------------- linke Menueleiste (NMGone-Stil) ----------------
        left = tk.Frame(self, bg=BG, width=240)
        left.grid(row=0, column=0, sticky="ns")
        left.grid_propagate(False)
        left.rowconfigure(2, weight=1)

        logo_box = tk.Frame(left, bg=BG, height=84)
        logo_box.grid(row=0, column=0, sticky="ew", padx=8, pady=(12, 4))
        logo_box.grid_propagate(False)
        logo = self._load_logo()
        if logo:
            tk.Label(logo_box, image=logo, bg=BG).pack(expand=True)
        else:
            tk.Label(logo_box, text="NMGone", font=("Arial", 17, "bold"),
                     fg=ACCENT, bg=BG).pack(expand=True)
        tk.Label(left, text="K A S S E", font=("Arial", 10, "bold"),
                 fg="#8aa0bb", bg=BG).grid(row=1, column=0, sticky="w", padx=18, pady=(0, 8))

        nav = tk.Frame(left, bg=BG)
        nav.grid(row=2, column=0, sticky="new", padx=8)
        self._nav_buttons = {}
        for key, text, icon in (("verkauf", "Verkauf", "🛒"),
                                 ("wareneingang", "Wareneingang", "📦")):
            b = tk.Button(nav, text=f"  {icon}   {text}", anchor="w", relief="flat",
                          bg=BG, fg="#11304d", font=("Arial", 11), bd=0,
                          activebackground=NAV_SEL, cursor="hand2",
                          command=lambda k=key: self._show_view(k))
            b.pack(fill="x", pady=2, ipady=4)
            self._nav_buttons[key] = b

        bottom = tk.Frame(left, bg=BG)
        bottom.grid(row=3, column=0, sticky="ew", padx=8, pady=(0, 6))
        tk.Button(bottom, text="🏠  NMGone öffnen", command=self._open_nmgone,
                  bg="#d8e2ee", fg=ACCENT, relief="flat", font=("Arial", 10, "bold"),
                  padx=10, pady=6, cursor="hand2").pack(fill="x", pady=(0, 6))
        tk.Button(bottom, text="Schließen", command=self._on_close, relief="flat",
                  bg="#eef1f5", fg="#444", padx=10, pady=4, cursor="hand2").pack(fill="x")
        tk.Label(left, text=f"Datenbank:\n{Path(self.db_path).name}", justify="left",
                 bg=BG, fg="#666", font=("Arial", 9)).grid(row=4, column=0, sticky="w", padx=16, pady=12)

        # ---------------- Hauptbereich ----------------
        main = tk.Frame(self, bg=SHELL_BG)
        main.grid(row=0, column=1, sticky="nsew")
        main.columnconfigure(0, weight=1)
        main.rowconfigure(1, weight=1)

        head = tk.Frame(main, bg=SHELL_BG)
        head.grid(row=0, column=0, sticky="ew", padx=22, pady=(16, 0))
        self._view_title = tk.Label(head, text="Verkauf", font=("Arial", 16, "bold"),
                                    fg=ACCENT, bg=SHELL_BG)
        self._view_title.pack(anchor="w")

        page = tk.Frame(main, bg=BG, highlightbackground="#d8e2ee", highlightthickness=1)
        page.grid(row=1, column=0, sticky="nsew", padx=22, pady=12)
        page.columnconfigure(0, weight=1)
        page.rowconfigure(0, weight=1)

        self._views = {}
        for key, builder in (("verkauf", self._build_verkauf),
                             ("wareneingang", self._build_wareneingang)):
            frame = tk.Frame(page, bg=BG)
            frame.grid(row=0, column=0, sticky="nsew")
            builder(frame)
            self._views[key] = frame

        self._show_view("verkauf")

    def _show_view(self, key):
        self._views[key].tkraise()
        self._view_title.config(text={"verkauf": "Verkauf",
                                      "wareneingang": "Wareneingang"}.get(key, key))
        for k, b in self._nav_buttons.items():
            if k == key:
                b.config(bg=NAV_SEL, fg=ACCENT, font=("Arial", 11, "bold"))
            else:
                b.config(bg=BG, fg="#11304d", font=("Arial", 11))

    def _open_nmgone(self):
        # Aus NMGone heraus: Hauptfenster nach vorn holen.
        if self._nmgone_action:
            try:
                self._nmgone_action()
                return
            except Exception:
                pass
        # Standalone: NMGone-Programm starten.
        try:
            if getattr(sys, "frozen", False):
                exe = Path(sys.executable).parent / "NMGone.exe"
                if not exe.exists():
                    raise FileNotFoundError("NMGone.exe nicht im Programmordner gefunden.")
                subprocess.Popen([str(exe)])
            else:
                subprocess.Popen([sys.executable, str(BASE_DIR / "start.py")])
        except Exception as e:
            messagebox.showerror("NMGone", f"NMGone konnte nicht geöffnet werden:\n{e}",
                                 parent=self.winfo_toplevel())

    # ================================================================= Verkauf
    def _build_verkauf(self, parent):
        kopf = tk.LabelFrame(parent, text="Kunde & Lieferung", bg=BG, fg=ACCENT,
                             font=("Arial", 10, "bold"), padx=12, pady=8)
        kopf.pack(fill="x", padx=8, pady=(10, 6))

        krow = tk.Frame(kopf, bg=BG)
        krow.pack(fill="x")
        khead = tk.Frame(krow, bg=BG)
        khead.pack(fill="x")
        tk.Label(khead, text="Kunde suchen:", bg=BG, font=("Arial", 10, "bold")).pack(side="left")
        tk.Button(khead, text="📥 Kundenliste importieren", command=self._import_kunden,
                  font=("Arial", 8), padx=6, pady=1).pack(side="right")
        self.vk_kunde_search = SearchBox(
            krow,
            fetch=lambda t: [(f"{knr or '—'}  ·  {name}  ·  {plz} {ort}", (knr, name))
                             for knr, name, plz, ort in self._search_kunden(t)],
            on_select=self._vk_pick_kunde, height=5,
        )
        self.vk_kunde_search.pack(fill="x", pady=(2, 4))
        self.vk_kunde_label = tk.Label(kopf, text="Kein Kunde gewählt.", bg=BG,
                                       fg="#999", font=("Arial", 9, "italic"))
        self.vk_kunde_label.pack(anchor="w")

        opt = tk.Frame(kopf, bg=BG)
        opt.pack(fill="x", pady=(6, 0))
        tk.Label(opt, text="Bestellart:", bg=BG, font=("Arial", 9, "bold")).pack(side="left")
        self.vk_art_var = tk.StringVar(value="Bestellung")
        ttk.Combobox(opt, textvariable=self.vk_art_var, width=13, state="readonly",
                     values=("Bestellung", "Vorbestellung", "abgesagt")).pack(side="left", padx=(4, 14))
        tk.Label(opt, text="Liefern:", bg=BG, font=("Arial", 9, "bold")).pack(side="left")
        self.vk_lieferzeit_var = tk.StringVar(value="10 Uhr")
        ttk.Combobox(opt, textvariable=self.vk_lieferzeit_var, width=8,
                     values=("10 Uhr", "12 Uhr", "Frei")).pack(side="left", padx=(4, 14))
        tk.Label(opt, text="Termin:", bg=BG, font=("Arial", 9, "bold")).pack(side="left")
        self.vk_termin_var = tk.StringVar()
        tk.Entry(opt, textvariable=self.vk_termin_var, width=12).pack(side="left", padx=(4, 0))

        # Artikel hinzufuegen
        addf = tk.LabelFrame(parent, text="Artikel hinzufügen (nur NMG)", bg=BG, fg=ACCENT,
                             font=("Arial", 10, "bold"), padx=12, pady=8)
        addf.pack(fill="x", padx=8, pady=(0, 6))
        self.vk_artikel_search = SearchBox(
            addf,
            fetch=lambda t: [(f"{pzn}  ·  {name}", pzn) for pzn, name in self._search_nmg(t)],
            on_select=self._vk_pick_artikel, height=5,
        )
        self.vk_artikel_search.pack(fill="x", pady=(2, 4))

        line = tk.Frame(addf, bg=BG)
        line.pack(fill="x", pady=(2, 0))
        self.vk_detail_label = tk.Label(line, text="—", bg=BG, fg="#444",
                                        font=("Arial", 9), anchor="w")
        self.vk_detail_label.pack(side="left", fill="x", expand=True)

        line2 = tk.Frame(addf, bg=BG)
        line2.pack(fill="x", pady=(6, 0))
        tk.Label(line2, text="Charge/Bestand:", bg=BG, font=("Arial", 9, "bold")).pack(side="left")
        self.vk_charge_var = tk.StringVar()
        self.vk_charge_cmb = ttk.Combobox(line2, textvariable=self.vk_charge_var,
                                          width=34, state="readonly")
        self.vk_charge_cmb.pack(side="left", padx=(4, 12))
        tk.Label(line2, text="Rabatt %:", bg=BG, font=("Arial", 9, "bold")).pack(side="left")
        self.vk_rabatt_var = tk.StringVar()
        tk.Entry(line2, textvariable=self.vk_rabatt_var, width=7).pack(side="left", padx=(4, 12))
        tk.Label(line2, text="Menge:", bg=BG, font=("Arial", 9, "bold")).pack(side="left")
        self.vk_menge_var = tk.StringVar(value="1")
        tk.Entry(line2, textvariable=self.vk_menge_var, width=6).pack(side="left", padx=(4, 12))
        tk.Button(line2, text="+ Position", command=self._vk_add_position,
                  bg="#11823b", fg="white", font=("Arial", 9, "bold"),
                  padx=10, pady=2).pack(side="left")

        # Positionsliste
        pos = tk.LabelFrame(parent, text="Positionen", bg=BG, fg=ACCENT,
                            font=("Arial", 10, "bold"), padx=12, pady=8)
        pos.pack(fill="both", expand=True, padx=8, pady=(0, 6))
        pcols = ("pzn", "artikel", "df", "pck", "apu", "menge", "rabatt", "quelle", "charge", "verfall")
        ph = {"pzn": "PZN", "artikel": "Artikel", "df": "DF", "pck": "PCK", "apu": "APU",
              "menge": "Menge", "rabatt": "Rab%", "quelle": "Quelle", "charge": "Charge", "verfall": "Verfall"}
        tf = tk.Frame(pos, bg=BG)
        tf.pack(fill="both", expand=True)
        tree = ttk.Treeview(tf, columns=pcols, show="headings", height=6)
        for c in pcols:
            tree.heading(c, text=ph[c])
            tree.column(c, width=190 if c == "artikel" else 58, anchor="w")
        tree.pack(side="left", fill="both", expand=True)
        sb = tk.Scrollbar(tf, orient="vertical", command=tree.yview)
        sb.pack(side="right", fill="y")
        tree.configure(yscrollcommand=sb.set)
        tree.bind("<Delete>", self._vk_remove_position)
        self.vk_pos_tree = tree

        afoot = tk.Frame(parent, bg=BG)
        afoot.pack(fill="x", padx=8, pady=(0, 8))
        tk.Label(afoot, text="Entf = markierte Position löschen", bg=BG, fg="#999",
                 font=("Arial", 8)).pack(side="left")
        tk.Button(afoot, text="Verkauf speichern", command=self._vk_save,
                  bg=ACCENT, fg="white", font=("Arial", 10, "bold"),
                  padx=14, pady=4).pack(side="right")

    def _vk_pick_kunde(self, payload):
        knr, name = payload
        self.vk_kunde = (knr, name)
        self.vk_kunde_label.config(
            text=f"Gewählt: {name}  ({knr or 'ohne Nr.'})", fg=ACCENT)

    def _vk_pick_artikel(self, pzn):
        det = self._artikel_details(pzn)
        if not det:
            return
        knr = self.vk_kunde[0] if self.vk_kunde else None
        prozent, quelle = self._resolve_rabatt(knr, pzn)
        det["rabatt"] = prozent
        det["rabatt_quelle"] = quelle
        self.vk_cur = det
        self.vk_detail_label.config(
            text=f"{det['artikelname']}  ·  DF {det['df'] or '—'}  ·  PCK {det['pck'] or '—'}  "
                 f"·  APU {det['apu'] if det['apu'] is not None else '—'}")
        self.vk_rabatt_var.set(f"{prozent:.0f}")
        # Chargen/Bestand
        chargen = self._lager_chargen(pzn)
        self._vk_charge_map = [("", "")]
        labels = ["(ohne Charge / kein Bestand)"]
        for charge, verfall, menge in chargen:
            labels.append(f"{charge or '—'}  ·  Verf {verfall or '—'}  ·  Bestand {menge}")
            self._vk_charge_map.append((charge or "", verfall or ""))
        self.vk_charge_cmb.config(values=labels)
        self.vk_charge_cmb.current(1 if chargen else 0)
        self.vk_menge_var.set("1")

    def _vk_add_position(self, *_):
        if not self.vk_cur:
            messagebox.showwarning("Verkauf", "Bitte zuerst einen Artikel wählen.",
                                   parent=self.winfo_toplevel())
            return
        try:
            menge = int(self.vk_menge_var.get().strip())
            if menge <= 0:
                raise ValueError
        except ValueError:
            messagebox.showwarning("Verkauf", "Menge muss eine positive Zahl sein.",
                                   parent=self.winfo_toplevel())
            return
        try:
            rabatt = float(str(self.vk_rabatt_var.get()).replace(",", ".") or 0)
        except ValueError:
            rabatt = 0.0
        idx = self.vk_charge_cmb.current()
        charge, verfall = self._vk_charge_map[idx] if 0 <= idx < len(self._vk_charge_map) else ("", "")
        p = {
            "pzn": self.vk_cur["pzn"], "artikelname": self.vk_cur["artikelname"],
            "df": self.vk_cur["df"], "pck": self.vk_cur["pck"], "apu": self.vk_cur["apu"],
            "menge": menge, "rabatt": rabatt, "rabatt_quelle": self.vk_cur["rabatt_quelle"],
            "charge": charge, "verfall": verfall,
        }
        self.vk_positions.append(p)
        self.vk_pos_tree.insert(
            "", "end",
            values=(p["pzn"], p["artikelname"], p["df"], p["pck"],
                    p["apu"] if p["apu"] is not None else "", p["menge"],
                    f"{rabatt:.0f}", p["rabatt_quelle"], p["charge"], p["verfall"]))
        # Eingabe zuruecksetzen
        self.vk_cur = None
        self.vk_artikel_search.clear()
        self.vk_detail_label.config(text="—")
        self.vk_charge_cmb.config(values=[])
        self.vk_charge_var.set("")
        self.vk_rabatt_var.set("")
        self.vk_menge_var.set("1")

    def _vk_remove_position(self, _e):
        sel = self.vk_pos_tree.selection()
        if not sel:
            return
        for item in sel:
            idx = self.vk_pos_tree.index(item)
            self.vk_pos_tree.delete(item)
            if 0 <= idx < len(self.vk_positions):
                del self.vk_positions[idx]

    def _vk_save(self):
        top = self.winfo_toplevel()
        if not self.vk_kunde:
            messagebox.showwarning("Verkauf", "Bitte einen Kunden wählen.", parent=top)
            return
        if not self.vk_positions:
            messagebox.showwarning("Verkauf", "Keine Positionen erfasst.", parent=top)
            return
        lieferzeit = self.vk_lieferzeit_var.get()
        header = {
            "datum": datetime.now().strftime("%Y-%m-%d"),
            "kundennummer": self.vk_kunde[0], "apotheke": self.vk_kunde[1],
            "bestellart": self.vk_art_var.get(), "lieferzeit": lieferzeit,
            "liefertermin": self.vk_termin_var.get().strip(),
        }
        bid = self._save_verkauf(header, self.vk_positions)
        anzahl = len(self.vk_positions)
        # Reset
        self.vk_positions = []
        self.vk_pos_tree.delete(*self.vk_pos_tree.get_children())
        self.vk_kunde = None
        self.vk_kunde_search.clear()
        self.vk_kunde_label.config(text="Kein Kunde gewählt.", fg="#999")
        self._refresh_lager()
        self._auftrag_dialog(bid, anzahl)

    # ==================================================== Auftragsbestaetigung
    def _auftrag_dialog(self, bestell_id, anzahl):
        top = self.winfo_toplevel()
        win = tk.Toplevel(top)
        win.title("Auftragsbestätigung")
        win.configure(bg=BG)
        win.transient(top)
        win.resizable(False, False)
        tk.Label(win, text=f"✓ Verkauf #{bestell_id} gespeichert",
                 font=("Arial", 12, "bold"), fg="#11823b", bg=BG).pack(padx=20, pady=(16, 2))
        tk.Label(win, text=f"{anzahl} Position(en) · Lagerbestand abgebucht.",
                 font=("Arial", 9), fg="#666", bg=BG).pack(padx=20, pady=(0, 12))

        btns = tk.Frame(win, bg=BG)
        btns.pack(padx=20, pady=(0, 8))
        tk.Button(btns, text="🖨  Drucken / Vorschau", width=22,
                  command=lambda: self._auftrag_drucken(bestell_id, win),
                  bg=ACCENT, fg="white", font=("Arial", 10, "bold"), pady=4).grid(row=0, column=0, padx=4, pady=3)
        tk.Button(btns, text="📧  Per E-Mail senden", width=22,
                  command=lambda: self._auftrag_mail(bestell_id, win),
                  bg="#3867b7", fg="white", font=("Arial", 10, "bold"), pady=4).grid(row=1, column=0, padx=4, pady=3)
        tk.Button(btns, text="Vorlage bearbeiten", width=22,
                  command=self._auftrag_vorlage_oeffnen, pady=3).grid(row=2, column=0, padx=4, pady=3)
        tk.Button(win, text="Schließen", command=win.destroy, padx=14, pady=3).pack(pady=(2, 14))
        win.lift()
        win.focus_force()

    def _auftrag_drucken(self, bestell_id, parent):
        try:
            path = auftrag.render(self.db_path, bestell_id)
            os.startfile(str(path))  # type: ignore[attr-defined]
        except Exception as e:
            messagebox.showerror("Auftragsbestätigung", f"Konnte nicht erstellt werden:\n{e}", parent=parent)

    def _auftrag_mail(self, bestell_id, parent):
        try:
            path = auftrag.render(self.db_path, bestell_id)
        except Exception as e:
            messagebox.showerror("Auftragsbestätigung", f"Konnte nicht erstellt werden:\n{e}", parent=parent)
            return
        to = auftrag.kunde_email(self.db_path, bestell_id)
        if not to:
            to = simpledialog.askstring("E-Mail", "Keine E-Mail beim Kunden hinterlegt.\n"
                                        "Empfänger-Adresse eingeben:", parent=parent) or ""
        try:
            html = path.read_text(encoding="utf-8")
            auftrag.send_via_outlook(to.strip(), f"Auftragsbestätigung #{bestell_id}", html, path)
        except Exception as e:
            messagebox.showerror("E-Mail", f"Outlook konnte nicht geöffnet werden:\n{e}", parent=parent)

    def _auftrag_vorlage_oeffnen(self):
        try:
            os.startfile(str(auftrag.template_path()))  # type: ignore[attr-defined]
        except Exception as e:
            messagebox.showerror("Vorlage", f"Vorlage konnte nicht geöffnet werden:\n{e}",
                                 parent=self.winfo_toplevel())

    # ============================================================ Wareneingang
    def _build_wareneingang(self, parent):
        form = tk.LabelFrame(parent, text="Artikel ins Lager buchen (nur NMG)", bg=BG,
                             fg=ACCENT, font=("Arial", 10, "bold"), padx=12, pady=8)
        form.pack(fill="x", padx=8, pady=(10, 6))

        whead = tk.Frame(form, bg=BG)
        whead.pack(fill="x")
        tk.Label(whead, text="Artikel suchen:", bg=BG, font=("Arial", 9, "bold")).pack(side="left")
        tk.Button(whead, text="📥 Wareneingangs-Liste importieren", command=self._import_wareneingang,
                  font=("Arial", 8), padx=6, pady=1).pack(side="right")
        self.we_search = SearchBox(
            form,
            fetch=lambda t: [(f"{pzn}  ·  {name}", (pzn, name)) for pzn, name in self._search_nmg(t)],
            on_select=self._we_pick_artikel, height=5,
        )
        self.we_search.pack(fill="x", pady=(2, 4))
        self.we_artikel_label = tk.Label(form, text="Kein Artikel gewählt.", bg=BG,
                                         fg="#999", font=("Arial", 9, "italic"))
        self.we_artikel_label.pack(anchor="w")

        row = tk.Frame(form, bg=BG)
        row.pack(fill="x", pady=(6, 0))
        self.we_charge_var = tk.StringVar()
        self.we_verfall_var = tk.StringVar()
        self.we_menge_var = tk.StringVar()
        self.we_ek_var = tk.StringVar()
        for label, var, w in (("Charge", self.we_charge_var, 12), ("Verfall", self.we_verfall_var, 10),
                              ("Menge", self.we_menge_var, 6), ("EK €", self.we_ek_var, 8)):
            tk.Label(row, text=label + ":", bg=BG, font=("Arial", 9, "bold")).pack(side="left", padx=(0, 4))
            tk.Entry(row, textvariable=var, width=w).pack(side="left", padx=(0, 12))
        tk.Button(row, text="Einbuchen", command=self._wareneingang_buchen,
                  bg=ACCENT, fg="white", font=("Arial", 10, "bold"),
                  padx=12, pady=3).pack(side="left")

        lager = tk.LabelFrame(parent, text="Lagerbestand  (Doppelklick = Bestand korrigieren)",
                              bg=BG, fg=ACCENT, font=("Arial", 10, "bold"), padx=12, pady=8)
        lager.pack(fill="both", expand=True, padx=8, pady=(0, 8))
        lcols = ("pzn", "artikel", "charge", "verfall", "menge")
        lh = {"pzn": "PZN", "artikel": "Artikel", "charge": "Charge",
              "verfall": "Verfall", "menge": "Bestand"}
        tf = tk.Frame(lager, bg=BG)
        tf.pack(fill="both", expand=True)
        tree = ttk.Treeview(tf, columns=lcols, show="headings", height=10)
        for c in lcols:
            tree.heading(c, text=lh[c])
            tree.column(c, width=240 if c == "artikel" else 90, anchor="w")
        tree.pack(side="left", fill="both", expand=True)
        sb = tk.Scrollbar(tf, orient="vertical", command=tree.yview)
        sb.pack(side="right", fill="y")
        tree.configure(yscrollcommand=sb.set)
        tree.bind("<Double-1>", self._lager_korrektur)
        self.we_lager_tree = tree
        self._refresh_lager()

    def _we_pick_artikel(self, payload):
        pzn, name = payload
        self.we_pzn = pzn
        self.we_artikel_label.config(text=f"Gewählt: {pzn} · {name}", fg=ACCENT)

    def _refresh_lager(self):
        self.we_lager_tree.delete(*self.we_lager_tree.get_children())
        with self._conn() as con:
            rows = con.execute(
                "SELECT pzn, artikelname, charge, verfall, menge FROM tbl_lagerbestand "
                "WHERE menge <> 0 ORDER BY artikelname, verfall"
            ).fetchall()
        for r in rows:
            self.we_lager_tree.insert("", "end", iid=None,
                                      values=tuple("" if v is None else v for v in r))

    def _wareneingang_buchen(self):
        top = self.winfo_toplevel()
        pzn = self.we_pzn
        if not pzn:
            messagebox.showwarning("Wareneingang", "Bitte zuerst einen NMG-Artikel wählen.", parent=top)
            return
        charge = self.we_charge_var.get().strip()
        verfall = self.we_verfall_var.get().strip()
        try:
            menge = int(self.we_menge_var.get().strip())
            if menge <= 0:
                raise ValueError
        except ValueError:
            messagebox.showwarning("Wareneingang", "Menge muss eine positive Zahl sein.", parent=top)
            return
        ek = None
        if self.we_ek_var.get().strip():
            try:
                ek = float(self.we_ek_var.get().strip().replace(",", "."))
            except ValueError:
                ek = None

        with self._conn() as con:
            row = con.execute(
                "SELECT artikelname FROM tbl_nmg_stamm WHERE pzn=? LIMIT 1", (pzn,)
            ).fetchone()
            if not row:
                messagebox.showwarning("Wareneingang",
                                       f"PZN {pzn} ist kein NMG-Artikel.", parent=top)
                return
            artikelname = row[0]
            jetzt = datetime.now().isoformat(timespec="seconds")
            cur = con.execute(
                "INSERT INTO tbl_wareneingang(datum, lieferant, bearbeiter) VALUES(?,?,?)",
                (jetzt, "NMG", getpass.getuser()))
            we_id = cur.lastrowid
            con.execute(
                "INSERT INTO tbl_wareneingang_positionen(we_id,pzn,artikelname,charge,verfall,menge,ek) "
                "VALUES(?,?,?,?,?,?,?)",
                (we_id, pzn, artikelname, charge, verfall, menge, ek))
            upd = con.execute(
                "UPDATE tbl_lagerbestand SET menge = menge + ?, aktualisiert_am = ? "
                "WHERE pzn=? AND COALESCE(charge,'')=? AND COALESCE(verfall,'')=?",
                (menge, jetzt, pzn, charge, verfall))
            if upd.rowcount == 0:
                con.execute(
                    "INSERT INTO tbl_lagerbestand(pzn,artikelname,charge,verfall,menge,aktualisiert_am) "
                    "VALUES(?,?,?,?,?,?)",
                    (pzn, artikelname, charge, verfall, menge, jetzt))
            con.commit()

        self.we_pzn = None
        self.we_search.clear()
        self.we_artikel_label.config(text="Kein Artikel gewählt.", fg="#999")
        for var in (self.we_charge_var, self.we_verfall_var, self.we_menge_var, self.we_ek_var):
            var.set("")
        self._refresh_lager()

    def _lager_korrektur(self, _e):
        top = self.winfo_toplevel()
        sel = self.we_lager_tree.selection()
        if not sel:
            return
        vals = self.we_lager_tree.item(sel[0], "values")
        pzn, _art, charge, verfall, menge = vals
        neu = simpledialog.askinteger(
            "Bestand korrigieren",
            f"Neuer Bestand für\nPZN {pzn} · Charge {charge or '—'} · Verf {verfall or '—'}\n"
            f"(0 entfernt die Zeile):",
            parent=top, initialvalue=int(menge), minvalue=0)
        if neu is None:
            return
        with self._conn() as con:
            if neu == 0:
                con.execute(
                    "DELETE FROM tbl_lagerbestand WHERE pzn=? AND COALESCE(charge,'')=? "
                    "AND COALESCE(verfall,'')=?", (pzn, charge, verfall))
            else:
                con.execute(
                    "UPDATE tbl_lagerbestand SET menge=?, aktualisiert_am=? WHERE pzn=? "
                    "AND COALESCE(charge,'')=? AND COALESCE(verfall,'')=?",
                    (neu, datetime.now().isoformat(timespec="seconds"), pzn, charge, verfall))
            con.commit()
        self._refresh_lager()

    # ==================================================================== Import
    def _pick_file(self):
        return filedialog.askopenfilename(
            title="Datei wählen (Excel, CSV oder TXT)",
            filetypes=[("Tabellen", "*.xlsx *.xlsm *.csv *.txt"),
                       ("Excel", "*.xlsx *.xlsm"), ("CSV", "*.csv"),
                       ("Text", "*.txt"), ("Alle Dateien", "*.*")],
            parent=self.winfo_toplevel())

    def _import_kunden(self):
        top = self.winfo_toplevel()
        path = self._pick_file()
        if not path:
            return
        try:
            r = kasse_import.import_kunden(self.db_path, path)
        except Exception as e:
            messagebox.showerror("Kunden-Import", self._import_fehlertext(path, e), parent=top)
            return
        messagebox.showinfo(
            "Kunden-Import",
            f"Quelle: {r['quelle']} · {r['gelesen']} Zeilen gelesen.\n"
            f"Erkannte Spalten: {', '.join(r['spalten'])}\n\n"
            f"Neu: {r['neu']} · Aktualisiert: {r['aktualisiert']} · "
            f"Übersprungen: {r['uebersprungen']}", parent=top)

    def _import_wareneingang(self):
        top = self.winfo_toplevel()
        path = self._pick_file()
        if not path:
            return
        try:
            r = kasse_import.import_wareneingang(self.db_path, path)
        except Exception as e:
            messagebox.showerror("Wareneingang-Import", self._import_fehlertext(path, e), parent=top)
            return
        self._refresh_lager()
        messagebox.showinfo(
            "Wareneingang-Import",
            f"Quelle: {r['quelle']} · {r['gelesen']} Zeilen gelesen.\n\n"
            f"Neue Chargen: {r['neu_chargen']} · Bestand erhöht: {r['erhoehte_chargen']}\n"
            f"Kein NMG-Artikel: {r['kein_nmg']} · Übersprungen: {r['uebersprungen']}", parent=top)

    @staticmethod
    def _import_fehlertext(path, e):
        msg = str(e)
        if path.lower().endswith(".pdf"):
            return ("PDF wird derzeit nicht unterstützt (keine PDF-Bibliothek installiert). "
                    "Bitte als Excel, CSV oder TXT speichern.\n\nDetail: " + msg)
        return "Import fehlgeschlagen:\n" + msg
