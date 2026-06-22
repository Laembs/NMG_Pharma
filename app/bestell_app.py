"""NMG Bestell-App - gemeinsame Oberflaeche fuer NMGone (Toplevel) und die
eigenstaendige Bestell-.exe (eigenes Hauptfenster / eigenes Taskleisten-Icon).

Die UI lebt in BestellPanel (einem tk.Frame), damit sie in beiden Kontexten
identisch ist. Datenmodell siehe migrations.py: tbl_bestellungen (Kopf) +
tbl_bestellpositionen (Zeilen).
"""
from __future__ import annotations

import sqlite3
import tkinter as tk
from tkinter import ttk, messagebox

from .config import DB_PATH


def ensure_bestell_tables(db_path=DB_PATH):
    """Bestell-Tabellen idempotent anlegen (spiegelt migrations.py).

    Aktiviert WAL, damit NMGone und die Bestell-.exe gleichzeitig auf dieselbe
    SQLite-Datei zugreifen koennen, ohne sich zu blockieren.
    """
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
                FOREIGN KEY(bestell_id) REFERENCES tbl_bestellungen(id) ON DELETE CASCADE
            )"""
        )
        con.commit()


class BestellPanel(tk.Frame):
    """Komplette Bestell-Oberflaeche.

    Phase 1: Kopfbereich (Kunde / Bestellart / Lieferung) + Positionen-Tabelle
    als Geruest. Kundensuche (Phase 3), NMG-Artikelsuche + Rabatt-Kaskade
    (Phase 4) und Speichern/Filter (Phase 5) folgen.
    """

    def __init__(self, master, db_path=DB_PATH, on_close=None):
        super().__init__(master, bg="#ffffff")
        self.db_path = db_path
        self._on_close = on_close or (lambda: self.winfo_toplevel().destroy())
        ensure_bestell_tables(db_path)
        self._build()

    def _build(self):
        # --- Kopf ---
        header = tk.Frame(self, bg="#ffffff")
        header.pack(fill="x", padx=16, pady=(14, 6))
        tk.Label(header, text="Bestellung", font=("Arial", 16, "bold"),
                 fg="#0b4a86", bg="#ffffff").pack(anchor="w")
        tk.Label(header, text="NMG-Artikel pro Apotheke bestellen.",
                 font=("Arial", 9), fg="#666", bg="#ffffff").pack(anchor="w", pady=(2, 0))

        kopf = tk.LabelFrame(self, text="Kunde & Lieferung", bg="#ffffff",
                             fg="#0b4a86", font=("Arial", 10, "bold"), padx=12, pady=10)
        kopf.pack(fill="x", padx=16, pady=(4, 8))

        krow = tk.Frame(kopf, bg="#ffffff")
        krow.pack(fill="x", pady=(0, 8))
        tk.Label(krow, text="Kunde:", width=12, anchor="w",
                 bg="#ffffff", font=("Arial", 10, "bold")).pack(side="left")
        self.kunde_var = tk.StringVar()
        tk.Entry(krow, textvariable=self.kunde_var, width=44).pack(side="left")
        tk.Label(krow, text="(Vorschlagssuche folgt in Phase 3)",
                 bg="#ffffff", fg="#999", font=("Arial", 8)).pack(side="left", padx=8)

        brow = tk.Frame(kopf, bg="#ffffff")
        brow.pack(fill="x", pady=(0, 8))
        tk.Label(brow, text="Bestellart:", width=12, anchor="w",
                 bg="#ffffff", font=("Arial", 10, "bold")).pack(side="left")
        self.art_var = tk.StringVar(value="Bestellung")
        for txt in ("Bestellung", "Vorbestellung", "abgesagt"):
            tk.Radiobutton(brow, text=txt, value=txt, variable=self.art_var,
                           bg="#ffffff", font=("Arial", 9)).pack(side="left", padx=(0, 10))

        lrow = tk.Frame(kopf, bg="#ffffff")
        lrow.pack(fill="x")
        tk.Label(lrow, text="Liefern:", width=12, anchor="w",
                 bg="#ffffff", font=("Arial", 10, "bold")).pack(side="left")
        self.lieferzeit_var = tk.StringVar(value="10 Uhr")
        for txt in ("10 Uhr", "12 Uhr", "Frei"):
            tk.Radiobutton(lrow, text=txt, value=txt, variable=self.lieferzeit_var,
                           bg="#ffffff", font=("Arial", 9)).pack(side="left", padx=(0, 8))
        self.lieferzeit_frei_var = tk.StringVar()
        tk.Entry(lrow, textvariable=self.lieferzeit_frei_var, width=10).pack(side="left", padx=(0, 16))
        tk.Label(lrow, text="Termin:", bg="#ffffff", font=("Arial", 10, "bold")).pack(side="left")
        self.termin_var = tk.StringVar()
        tk.Entry(lrow, textvariable=self.termin_var, width=12).pack(side="left", padx=(4, 0))

        # --- Positionen ---
        pos = tk.LabelFrame(self, text="Positionen (nur NMG-Artikel)", bg="#ffffff",
                            fg="#0b4a86", font=("Arial", 10, "bold"), padx=12, pady=10)
        pos.pack(fill="both", expand=True, padx=16, pady=(0, 8))

        arow = tk.Frame(pos, bg="#ffffff")
        arow.pack(fill="x", pady=(0, 8))
        tk.Label(arow, text="Artikel:", bg="#ffffff", font=("Arial", 10, "bold")).pack(side="left")
        self.artikel_var = tk.StringVar()
        tk.Entry(arow, textvariable=self.artikel_var, width=40).pack(side="left", padx=(6, 8))
        tk.Label(arow, text="(NMG-Artikelsuche + Auto-Fill folgen in Phase 4)",
                 bg="#ffffff", fg="#999", font=("Arial", 8)).pack(side="left")

        pcols = ("pzn", "artikel", "df", "pck", "apu", "menge", "rabatt", "quelle")
        pheads = {"pzn": "PZN", "artikel": "Artikel", "df": "DF", "pck": "PCK",
                  "apu": "APU", "menge": "Menge", "rabatt": "Rabatt %", "quelle": "Quelle"}
        tree_frame = tk.Frame(pos, bg="#ffffff")
        tree_frame.pack(fill="both", expand=True)
        tree = ttk.Treeview(tree_frame, columns=pcols, show="headings", height=8)
        for c in pcols:
            tree.heading(c, text=pheads[c])
            w = 220 if c == "artikel" else (70 if c in ("pzn", "rabatt", "quelle") else 55)
            tree.column(c, width=w, anchor="w")
        tree.pack(side="left", fill="both", expand=True)
        sb = tk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
        sb.pack(side="right", fill="y")
        tree.configure(yscrollcommand=sb.set)
        self.pos_tree = tree

        # --- Aktionen ---
        footer = tk.Frame(self, bg="#ffffff")
        footer.pack(fill="x", padx=16, pady=(0, 14))

        tk.Button(footer, text="Speichern", command=self._save_todo, bg="#0b4a86",
                  fg="white", font=("Arial", 10, "bold"), padx=14, pady=4).pack(side="right", padx=(8, 0))
        tk.Button(footer, text="Schließen", command=self._on_close,
                  font=("Arial", 10), padx=14, pady=4).pack(side="right")

    def _save_todo(self):
        messagebox.showinfo(
            "Bestell-App",
            "Speichern wird in Phase 5 ergaenzt.\n\nPhase 1 liefert das Fenster-"
            "Skelett mit fertigem Datenmodell (tbl_bestellungen + tbl_bestellpositionen).",
            parent=self.winfo_toplevel(),
        )
