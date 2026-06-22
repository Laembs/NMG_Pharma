"""NMG Kasse-App (ehem. Bestell-App) - gemeinsame Oberflaeche fuer NMGone
(Toplevel) und die eigenstaendige Kasse-.exe (eigenes Fenster / Taskleisten-Icon).

Warenkreislauf: Wareneingang -> Lagerbestand -> Verkauf. Die UI liegt in
KassePanel (tk.Frame) mit zwei Reitern:
  - Verkauf      = Warenausgang an Apotheke (Geruest; Phasen C folgen)
  - Wareneingang = NMG-Artikel mit Charge/Verfall/Menge ins Lager buchen

Datenmodell siehe migrations.py: tbl_bestellungen/_positionen (Verkauf),
tbl_wareneingang/_positionen, tbl_lagerbestand.
"""
from __future__ import annotations

import getpass
import sqlite3
from datetime import datetime
import tkinter as tk
from tkinter import ttk, messagebox

from .config import DB_PATH

BG = "#ffffff"
ACCENT = "#0b4a86"


def ensure_kasse_tables(db_path=DB_PATH):
    """Alle Kasse-Tabellen idempotent anlegen (spiegelt migrations.py).

    Aktiviert WAL, damit NMGone und die Kasse-.exe gleichzeitig auf dieselbe
    SQLite-Datei zugreifen koennen.
    """
    with sqlite3.connect(db_path) as con:
        try:
            con.execute("PRAGMA journal_mode=WAL")
        except sqlite3.Error:
            pass
        # Verkauf (ehem. Bestellung)
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
        # Wareneingang
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
        # Lagerbestand: eine Zeile pro PZN x Charge x Verfall
        con.execute(
            """CREATE TABLE IF NOT EXISTS tbl_lagerbestand(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                pzn TEXT, artikelname TEXT, charge TEXT, verfall TEXT,
                menge INTEGER DEFAULT 0, aktualisiert_am TEXT,
                UNIQUE(pzn, charge, verfall)
            )"""
        )
        con.commit()


class KassePanel(tk.Frame):
    """Komplette Kasse-Oberflaeche: Reiter Verkauf + Wareneingang."""

    def __init__(self, master, db_path=DB_PATH, on_close=None):
        super().__init__(master, bg=BG)
        self.db_path = db_path
        self._on_close = on_close or (lambda: self.winfo_toplevel().destroy())
        ensure_kasse_tables(db_path)
        self._build()

    # ------------------------------------------------------------------ Aufbau
    def _build(self):
        header = tk.Frame(self, bg=BG)
        header.pack(fill="x", padx=16, pady=(14, 6))
        tk.Label(header, text="Kasse", font=("Arial", 16, "bold"),
                 fg=ACCENT, bg=BG).pack(anchor="w")
        tk.Label(header, text="Verkauf an Apotheken · Wareneingang ins Lager.",
                 font=("Arial", 9), fg="#666", bg=BG).pack(anchor="w", pady=(2, 0))

        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True, padx=16, pady=(4, 8))
        verkauf = tk.Frame(nb, bg=BG)
        wareneingang = tk.Frame(nb, bg=BG)
        nb.add(verkauf, text="  Verkauf  ")
        nb.add(wareneingang, text="  Wareneingang  ")
        self._build_verkauf(verkauf)
        self._build_wareneingang(wareneingang)

        footer = tk.Frame(self, bg=BG)
        footer.pack(fill="x", padx=16, pady=(0, 14))
        tk.Button(footer, text="Schließen", command=self._on_close,
                  font=("Arial", 10), padx=14, pady=4).pack(side="right")

    # ----------------------------------------------------------------- Verkauf
    def _build_verkauf(self, parent):
        kopf = tk.LabelFrame(parent, text="Kunde & Lieferung", bg=BG, fg=ACCENT,
                             font=("Arial", 10, "bold"), padx=12, pady=10)
        kopf.pack(fill="x", padx=8, pady=(10, 8))

        krow = tk.Frame(kopf, bg=BG)
        krow.pack(fill="x", pady=(0, 8))
        tk.Label(krow, text="Kunde:", width=12, anchor="w", bg=BG,
                 font=("Arial", 10, "bold")).pack(side="left")
        self.vk_kunde_var = tk.StringVar()
        tk.Entry(krow, textvariable=self.vk_kunde_var, width=44).pack(side="left")
        tk.Label(krow, text="(globale Kundensuche folgt)", bg=BG, fg="#999",
                 font=("Arial", 8)).pack(side="left", padx=8)

        brow = tk.Frame(kopf, bg=BG)
        brow.pack(fill="x", pady=(0, 8))
        tk.Label(brow, text="Bestellart:", width=12, anchor="w", bg=BG,
                 font=("Arial", 10, "bold")).pack(side="left")
        self.vk_art_var = tk.StringVar(value="Bestellung")
        for txt in ("Bestellung", "Vorbestellung", "abgesagt"):
            tk.Radiobutton(brow, text=txt, value=txt, variable=self.vk_art_var,
                           bg=BG, font=("Arial", 9)).pack(side="left", padx=(0, 10))

        lrow = tk.Frame(kopf, bg=BG)
        lrow.pack(fill="x")
        tk.Label(lrow, text="Liefern:", width=12, anchor="w", bg=BG,
                 font=("Arial", 10, "bold")).pack(side="left")
        self.vk_lieferzeit_var = tk.StringVar(value="10 Uhr")
        for txt in ("10 Uhr", "12 Uhr", "Frei"):
            tk.Radiobutton(lrow, text=txt, value=txt, variable=self.vk_lieferzeit_var,
                           bg=BG, font=("Arial", 9)).pack(side="left", padx=(0, 8))
        self.vk_lieferzeit_frei_var = tk.StringVar()
        tk.Entry(lrow, textvariable=self.vk_lieferzeit_frei_var, width=10).pack(side="left", padx=(0, 16))
        tk.Label(lrow, text="Termin:", bg=BG, font=("Arial", 10, "bold")).pack(side="left")
        self.vk_termin_var = tk.StringVar()
        tk.Entry(lrow, textvariable=self.vk_termin_var, width=12).pack(side="left", padx=(4, 0))

        pos = tk.LabelFrame(parent, text="Positionen (nur NMG-Artikel)", bg=BG, fg=ACCENT,
                            font=("Arial", 10, "bold"), padx=12, pady=10)
        pos.pack(fill="both", expand=True, padx=8, pady=(0, 8))
        arow = tk.Frame(pos, bg=BG)
        arow.pack(fill="x", pady=(0, 8))
        tk.Label(arow, text="Artikel:", bg=BG, font=("Arial", 10, "bold")).pack(side="left")
        self.vk_artikel_var = tk.StringVar()
        tk.Entry(arow, textvariable=self.vk_artikel_var, width=40).pack(side="left", padx=(6, 8))
        tk.Label(arow, text="(NMG-Artikelsuche + Bestand/Charge folgen)", bg=BG, fg="#999",
                 font=("Arial", 8)).pack(side="left")

        pcols = ("pzn", "artikel", "df", "pck", "apu", "menge", "rabatt", "charge", "verfall")
        pheads = {"pzn": "PZN", "artikel": "Artikel", "df": "DF", "pck": "PCK", "apu": "APU",
                  "menge": "Menge", "rabatt": "Rabatt %", "charge": "Charge", "verfall": "Verfall"}
        tf = tk.Frame(pos, bg=BG)
        tf.pack(fill="both", expand=True)
        tree = ttk.Treeview(tf, columns=pcols, show="headings", height=7)
        for c in pcols:
            tree.heading(c, text=pheads[c])
            tree.column(c, width=200 if c == "artikel" else 64, anchor="w")
        tree.pack(side="left", fill="both", expand=True)
        sb = tk.Scrollbar(tf, orient="vertical", command=tree.yview)
        sb.pack(side="right", fill="y")
        tree.configure(yscrollcommand=sb.set)
        self.vk_pos_tree = tree

        afoot = tk.Frame(parent, bg=BG)
        afoot.pack(fill="x", padx=8, pady=(0, 8))
        tk.Button(afoot, text="Verkauf speichern", command=self._verkauf_todo,
                  bg=ACCENT, fg="white", font=("Arial", 10, "bold"),
                  padx=14, pady=4).pack(side="right")

    def _verkauf_todo(self):
        messagebox.showinfo(
            "Kasse",
            "Verkauf speichern folgt (globale Kundensuche, NMG-Artikelsuche mit "
            "Bestand/Charge, Rabatt-Kaskade). Aktuell Geruest.",
            parent=self.winfo_toplevel(),
        )

    # ------------------------------------------------------------ Wareneingang
    def _build_wareneingang(self, parent):
        form = tk.LabelFrame(parent, text="Artikel ins Lager buchen", bg=BG, fg=ACCENT,
                             font=("Arial", 10, "bold"), padx=12, pady=10)
        form.pack(fill="x", padx=8, pady=(10, 8))

        row = tk.Frame(form, bg=BG)
        row.pack(fill="x")
        self.we_pzn_var = tk.StringVar()
        self.we_charge_var = tk.StringVar()
        self.we_verfall_var = tk.StringVar()
        self.we_menge_var = tk.StringVar()
        for label, var, w in (("PZN", self.we_pzn_var, 12), ("Charge", self.we_charge_var, 12),
                              ("Verfall", self.we_verfall_var, 10), ("Menge", self.we_menge_var, 6)):
            tk.Label(row, text=label + ":", bg=BG, font=("Arial", 10, "bold")).pack(side="left", padx=(0, 4))
            tk.Entry(row, textvariable=var, width=w).pack(side="left", padx=(0, 12))
        tk.Button(row, text="Einbuchen", command=self._wareneingang_buchen,
                  bg=ACCENT, fg="white", font=("Arial", 10, "bold"),
                  padx=12, pady=3).pack(side="left")
        tk.Label(form, text="Nur NMG-Artikel (PZN aus tbl_nmg_stamm). Verfall frei als Text.",
                 bg=BG, fg="#999", font=("Arial", 8)).pack(anchor="w", pady=(6, 0))

        lager = tk.LabelFrame(parent, text="Lagerbestand", bg=BG, fg=ACCENT,
                              font=("Arial", 10, "bold"), padx=12, pady=10)
        lager.pack(fill="both", expand=True, padx=8, pady=(0, 8))
        lcols = ("pzn", "artikel", "charge", "verfall", "menge")
        lheads = {"pzn": "PZN", "artikel": "Artikel", "charge": "Charge",
                  "verfall": "Verfall", "menge": "Bestand"}
        tf = tk.Frame(lager, bg=BG)
        tf.pack(fill="both", expand=True)
        tree = ttk.Treeview(tf, columns=lcols, show="headings", height=10)
        for c in lcols:
            tree.heading(c, text=lheads[c])
            tree.column(c, width=240 if c == "artikel" else 90, anchor="w")
        tree.pack(side="left", fill="both", expand=True)
        sb = tk.Scrollbar(tf, orient="vertical", command=tree.yview)
        sb.pack(side="right", fill="y")
        tree.configure(yscrollcommand=sb.set)
        self.we_lager_tree = tree
        self._refresh_lager()

    def _refresh_lager(self):
        self.we_lager_tree.delete(*self.we_lager_tree.get_children())
        with sqlite3.connect(self.db_path) as con:
            rows = con.execute(
                "SELECT pzn, artikelname, charge, verfall, menge FROM tbl_lagerbestand "
                "WHERE menge <> 0 ORDER BY artikelname, verfall"
            ).fetchall()
        for r in rows:
            self.we_lager_tree.insert("", "end", values=tuple("" if v is None else v for v in r))

    def _wareneingang_buchen(self):
        pzn = self.we_pzn_var.get().strip()
        charge = self.we_charge_var.get().strip()
        verfall = self.we_verfall_var.get().strip()
        menge_raw = self.we_menge_var.get().strip()
        top = self.winfo_toplevel()
        if not pzn:
            messagebox.showwarning("Wareneingang", "Bitte eine PZN eingeben.", parent=top)
            return
        try:
            menge = int(menge_raw)
            if menge <= 0:
                raise ValueError
        except ValueError:
            messagebox.showwarning("Wareneingang", "Menge muss eine positive Zahl sein.", parent=top)
            return

        with sqlite3.connect(self.db_path) as con:
            row = con.execute(
                "SELECT artikelname FROM tbl_nmg_stamm WHERE pzn = ? LIMIT 1", (pzn,)
            ).fetchone()
            if not row:
                messagebox.showwarning(
                    "Wareneingang",
                    f"PZN {pzn} ist kein NMG-Artikel (nicht in tbl_nmg_stamm).",
                    parent=top,
                )
                return
            artikelname = row[0]
            jetzt = datetime.now().isoformat(timespec="seconds")
            bearbeiter = getpass.getuser()
            # Wareneingang-Beleg + Position (Historie)
            cur = con.execute(
                "INSERT INTO tbl_wareneingang(datum, lieferant, bearbeiter) VALUES(?,?,?)",
                (jetzt, "NMG", bearbeiter),
            )
            we_id = cur.lastrowid
            con.execute(
                "INSERT INTO tbl_wareneingang_positionen(we_id, pzn, artikelname, charge, verfall, menge) "
                "VALUES(?,?,?,?,?,?)",
                (we_id, pzn, artikelname, charge, verfall, menge),
            )
            # Lagerbestand hochbuchen (Upsert pro pzn/charge/verfall)
            upd = con.execute(
                "UPDATE tbl_lagerbestand SET menge = menge + ?, aktualisiert_am = ? "
                "WHERE pzn = ? AND COALESCE(charge,'') = ? AND COALESCE(verfall,'') = ?",
                (menge, jetzt, pzn, charge, verfall),
            )
            if upd.rowcount == 0:
                con.execute(
                    "INSERT INTO tbl_lagerbestand(pzn, artikelname, charge, verfall, menge, aktualisiert_am) "
                    "VALUES(?,?,?,?,?,?)",
                    (pzn, artikelname, charge, verfall, menge, jetzt),
                )
            con.commit()

        for var in (self.we_pzn_var, self.we_charge_var, self.we_verfall_var, self.we_menge_var):
            var.set("")
        self._refresh_lager()
