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


def _eur(v):
    if v is None:
        return "—"
    return f"{v:,.2f} €".replace(",", "X").replace(".", ",").replace("X", ".")


def _pos_netto(p):
    """Netto-Zeilensumme einer Position (APU * Menge * (1 - Rabatt%))."""
    if p.get("apu") is None:
        return 0.0
    return p["apu"] * (p.get("menge") or 0) * (1 - (p.get("rabatt") or 0) / 100.0)


def _normalize_uhrzeit(t):
    """Leer -> '', gueltige Uhrzeit -> 'HH:MM', ungueltig -> None.
    Akzeptiert '14:30', '14.30', '14', '14 Uhr'."""
    t = str(t or "").strip().lower().replace("uhr", "").replace(".", ":").strip()
    if not t:
        return ""
    teile = t.split(":")
    try:
        h = int(teile[0])
        m = int(teile[1]) if len(teile) > 1 and teile[1] != "" else 0
        if 0 <= h < 24 and 0 <= m < 60:
            return f"{h:02d}:{m:02d}"
    except (ValueError, IndexError):
        pass
    return None


def _make_treeview_sortable(tree):
    """Macht jede Spalte einer Treeview per Klick auf den Spaltenkopf sortierbar
    (auf/ab umschaltend). Zahlen/Waehrung/Prozent werden numerisch sortiert,
    Datum/Text alphabetisch. iids bleiben erhalten (Row-Maps bleiben gueltig)."""
    state = {"col": None, "desc": False}

    def _key(val):
        s = str(val).strip()
        if not s:
            return (2, "")
        cand = s.replace("€", "").replace("%", "").replace(" ", "").strip()
        if "," in cand:
            cand = cand.replace(".", "").replace(",", ".")
        if cand.count(".") <= 1 and "-" not in cand and "/" not in cand:
            try:
                return (0, float(cand))
            except ValueError:
                pass
        return (1, s.lower())

    def sort_by(col):
        desc = (state["col"] == col) and not state["desc"]
        state["col"], state["desc"] = col, desc
        data = [(tree.set(iid, col), iid) for iid in tree.get_children("")]
        data.sort(key=lambda x: _key(x[0]), reverse=desc)
        for i, (_, iid) in enumerate(data):
            tree.move(iid, "", i)
        for c in tree["columns"]:
            txt = tree.heading(c, "text").rstrip(" ▲▼")
            tree.heading(c, text=txt + (" ▼" if desc else " ▲") if c == col else txt)

    for col in tree["columns"]:
        tree.heading(col, command=lambda c=col: sort_by(c))


def _normalize_verfall(t):
    """Leer -> ''. Gueltiges Verfalldatum (MM/JJ, MM/JJJJ, auch '.' als Trenner,
    einstelliger Monat) -> 'MM/JJ' bzw. 'MM/JJJJ' mit aufgefuelltem Monat.
    Ungueltig (z.B. Monat 13) -> None."""
    t = str(t or "").strip().replace(".", "/").replace("-", "/")
    if not t:
        return ""
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


def _next_liefertermin(now=None):
    """Liefertermin-Vorschlag: vor 15 Uhr -> naechster Liefertag, ab 15 Uhr ->
    uebernaechster. Sonntage werden uebersprungen. Samstag ist nur dann ein
    Liefertag, wenn Freitag vor 15 Uhr gebucht wurde (-> Samstag); sonst zaehlt
    nur Mo-Fr (Fr ab 15 -> Di)."""
    from datetime import datetime as _dt, timedelta
    now = now or _dt.now()
    vor15 = now.hour < 15
    # Sonderfall: Freitag vor 15 Uhr -> Samstag-Lieferung.
    if now.weekday() == 4 and vor15:
        return (now + timedelta(days=1)).strftime("%d.%m.%Y")
    schritte = 1 if vor15 else 2
    d = now
    gezaehlt = 0
    while gezaehlt < schritte:
        d = d + timedelta(days=1)
        if d.weekday() < 5:  # Mo-Fr (Samstag nur im Sonderfall oben)
            gezaehlt += 1
    return d.strftime("%d.%m.%Y")


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
                erstellt_am TEXT DEFAULT CURRENT_TIMESTAMP, geaendert_am TEXT,
                msk_erfasst INTEGER DEFAULT 0, msk_von TEXT, msk_am TEXT
            )"""
        )
        _bcols = {r[1] for r in con.execute("PRAGMA table_info(tbl_bestellungen)")}
        for col, ddl in (("msk_erfasst", "msk_erfasst INTEGER DEFAULT 0"),
                         ("msk_von", "msk_von TEXT"), ("msk_am", "msk_am TEXT")):
            if col not in _bcols:
                con.execute(f"ALTER TABLE tbl_bestellungen ADD COLUMN {ddl}")
        con.execute(
            """CREATE TABLE IF NOT EXISTS tbl_bestellpositionen(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                bestell_id INTEGER NOT NULL,
                pzn TEXT, artikelname TEXT, df TEXT, pck TEXT, apu REAL,
                menge INTEGER DEFAULT 1, rabatt_prozent REAL, rabatt_quelle TEXT,
                charge TEXT, verfall TEXT,
                bestellart TEXT DEFAULT 'Bestellung', lieferzeit TEXT, liefertermin TEXT,
                FOREIGN KEY(bestell_id) REFERENCES tbl_bestellungen(id) ON DELETE CASCADE
            )"""
        )
        _have = {r[1] for r in con.execute("PRAGMA table_info(tbl_bestellpositionen)")}
        for col, ddl in (("charge", "charge TEXT"), ("verfall", "verfall TEXT"),
                         ("bestellart", "bestellart TEXT DEFAULT 'Bestellung'"),
                         ("lieferzeit", "lieferzeit TEXT"), ("liefertermin", "liefertermin TEXT")):
            if col not in _have:
                con.execute(f"ALTER TABLE tbl_bestellpositionen ADD COLUMN {ddl}")
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
        con.execute(
            """CREATE TABLE IF NOT EXISTS tbl_kasse_log(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                zeitpunkt TEXT DEFAULT CURRENT_TIMESTAMP,
                bearbeiter TEXT, aktion TEXT, bestell_id INTEGER, kunde TEXT, details TEXT
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
        self.vk_positions = []        # Liste ALLER Verkaufspositionen (dicts)
        self._vk_tree_map = {}        # Tree-iid -> Position (nur Bestellungen sichtbar)
        self.we_pzn = None
        self._build()

    # =============================================================== Datenbank
    def _conn(self):
        return sqlite3.connect(self.db_path)

    def _table_exists_now(self, name):
        with self._conn() as con:
            return _table_exists(con, name)

    def _log(self, aktion, bestell_id=None, kunde="", details=""):
        """Schreibt einen Eintrag ins Aenderungs-Protokoll (wer/was/wann)."""
        try:
            with self._conn() as con:
                con.execute(
                    "INSERT INTO tbl_kasse_log(zeitpunkt,bearbeiter,aktion,bestell_id,kunde,details) "
                    "VALUES(?,?,?,?,?,?)",
                    (datetime.now().isoformat(timespec="seconds"), getpass.getuser(),
                     aktion, bestell_id, kunde, details))
                con.commit()
            if hasattr(self, "log_tree"):
                self._refresh_log()
        except Exception:
            pass

    def _restyle_buttons(self, widget):
        """Vereinheitlicht alle tk.Button im NMGone-Flat-Stil: farbige Aktions-
        Buttons behalten ihre Farbe (nur flach + Hand-Cursor), neutrale Standard-
        Buttons bekommen den hellblauen Sekundaer-Look. Aendert nur die Optik."""
        for ch in widget.winfo_children():
            if isinstance(ch, tk.Button):
                try:
                    bg = str(ch.cget("background"))
                    # Nur die grauen Standard-Buttons (Win95-Look) umstylen; weisse
                    # (Nav) und farbige (Aktions-) Buttons behalten ihren Hintergrund.
                    plain = bg.startswith("System") or bg.lower() in (
                        "#f0f0f0", "#ececec", "#e1e1e1", "#d9d9d9")
                    if plain:
                        ch.configure(bg="#e8eef5", fg="#11304d", activebackground="#d8e2ee",
                                     activeforeground=ACCENT, relief="flat", bd=0,
                                     cursor="hand2", highlightthickness=0, padx=10, pady=3)
                    else:
                        ch.configure(relief="flat", bd=0, cursor="hand2",
                                     activebackground=bg, highlightthickness=0)
                except tk.TclError:
                    pass
            self._restyle_buttons(ch)

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

    def _bestand_for(self, pzn, charge, verfall):
        with self._conn() as con:
            r = con.execute(
                "SELECT menge FROM tbl_lagerbestand WHERE pzn=? AND COALESCE(charge,'')=? "
                "AND COALESCE(verfall,'')=?", (pzn, charge or "", verfall or "")).fetchone()
            return r[0] if r else 0

    def _bestand_total(self, pzn):
        with self._conn() as con:
            r = con.execute(
                "SELECT COALESCE(SUM(menge),0) FROM tbl_lagerbestand WHERE pzn=?", (pzn,)).fetchone()
            return r[0] if r else 0

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
                    "apu,menge,rabatt_prozent,rabatt_quelle,charge,verfall,"
                    "bestellart,lieferzeit,liefertermin) "
                    "VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
                    (bid, p["pzn"], p["artikelname"], p["df"], p["pck"], p["apu"],
                     p["menge"], p["rabatt"], p["rabatt_quelle"], p["charge"], p["verfall"],
                     p.get("bestellart", "Bestellung"), p.get("lieferzeit", ""), p.get("liefertermin", "")),
                )
                # Lager NUR fuer echte Bestellungen abbuchen (Vorbestellung/abgesagt
                # ziehen noch nichts ab). Locker: nie unter 0, keine Sperre.
                if p.get("bestellart", "Bestellung") == "Bestellung" and (p["charge"] or p["verfall"]):
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
                                 ("vorbestellungen", "Vorbestellungen", "🕓"),
                                 ("verkaeufe", "Verkäufe", "🧾"),
                                 ("artikel", "Artikel", "🔍"),
                                 ("wareneingang", "Wareneingang", "📦"),
                                 ("protokoll", "Protokoll", "📝")):
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
                             ("vorbestellungen", self._build_vorbestellungen),
                             ("verkaeufe", self._build_verkaeufe),
                             ("artikel", self._build_artikel),
                             ("wareneingang", self._build_wareneingang),
                             ("protokoll", self._build_protokoll)):
            frame = tk.Frame(page, bg=BG)
            frame.grid(row=0, column=0, sticky="nsew")
            builder(frame)
            self._views[key] = frame

        self._show_view("verkauf")
        self._restyle_buttons(self)

    def _show_view(self, key):
        self._views[key].tkraise()
        self._view_title.config(text={"verkauf": "Verkauf",
                                      "vorbestellungen": "Vorbestellungen",
                                      "verkaeufe": "Verkäufe",
                                      "artikel": "Artikel-Übersicht",
                                      "wareneingang": "Wareneingang",
                                      "protokoll": "Änderungs-Protokoll"}.get(key, key))
        if key == "vorbestellungen":
            self._refresh_vorbestellungen()
        elif key == "verkaeufe":
            self._refresh_verkaeufe()
        elif key == "artikel":
            self._refresh_artikel()
        elif key == "protokoll":
            self._refresh_log()
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
        kopf = tk.LabelFrame(parent, text="Kunde", bg=BG, fg=ACCENT,
                             font=("Arial", 10, "bold"), padx=12, pady=8)
        kopf.pack(fill="x", padx=8, pady=(10, 6))

        khead = tk.Frame(kopf, bg=BG)
        khead.pack(fill="x")
        tk.Label(khead, text="Kunde suchen:", bg=BG, font=("Arial", 10, "bold")).pack(side="left")
        tk.Button(khead, text="✕ Kunde entfernen", command=self._vk_clear_kunde,
                  font=("Arial", 8), padx=6, pady=1).pack(side="left", padx=(10, 0))
        tk.Button(khead, text="📥 Kundenliste importieren", command=self._import_kunden,
                  font=("Arial", 8), padx=6, pady=1).pack(side="right")
        tk.Button(khead, text="📊 Top-Artikel (12 Mon.)", command=self._show_kunde_top,
                  font=("Arial", 8), padx=6, pady=1).pack(side="right", padx=(0, 6))
        self.vk_kunde_search = SearchBox(
            kopf,
            fetch=lambda t: [(f"{knr or '—'}  ·  {name}  ·  {plz} {ort}", (knr, name, plz, ort))
                             for knr, name, plz, ort in self._search_kunden(t)],
            on_select=self._vk_pick_kunde, height=5,
        )
        self.vk_kunde_search.pack(fill="x", pady=(2, 6))

        # Kundenkarte: alle relevanten Daten, sobald ein Kunde gewaehlt ist.
        self.vk_kunde_card = tk.Frame(kopf, bg="#f2f6fb", highlightbackground="#dde7f1",
                                      highlightthickness=1)
        self.vk_kunde_card.pack(fill="x")
        self.vk_kunde_card_label = tk.Label(self.vk_kunde_card, text="Kein Kunde gewählt.",
                                            bg="#f2f6fb", fg="#888", justify="left", anchor="w",
                                            font=("Arial", 9, "italic"))
        self.vk_kunde_card_label.pack(fill="x", padx=10, pady=8)

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
                                          width=30, state="readonly")
        self.vk_charge_cmb.pack(side="left", padx=(4, 12))
        tk.Label(line2, text="Rabatt %:", bg=BG, font=("Arial", 9, "bold")).pack(side="left")
        self.vk_rabatt_var = tk.StringVar()
        tk.Entry(line2, textvariable=self.vk_rabatt_var, width=6).pack(side="left", padx=(4, 12))
        tk.Label(line2, text="Menge:", bg=BG, font=("Arial", 9, "bold")).pack(side="left")
        self.vk_menge_var = tk.StringVar(value="1")
        tk.Entry(line2, textvariable=self.vk_menge_var, width=5).pack(side="left", padx=(4, 12))

        # Liefervorgabe PRO Position (je Artikel eigene Vorgabe moeglich).
        line3 = tk.Frame(addf, bg=BG)
        line3.pack(fill="x", pady=(6, 0))
        tk.Label(line3, text="Liefervorgabe:", bg=BG, font=("Arial", 9, "bold")).pack(side="left")
        self.vk_pos_art_var = tk.StringVar(value="Bestellung")
        ttk.Combobox(line3, textvariable=self.vk_pos_art_var, width=12, state="readonly",
                     values=("Bestellung", "Vorbestellung", "abgesagt")).pack(side="left", padx=(4, 8))
        self.vk_pos_lieferzeit_var = tk.StringVar(value="Frei")
        ttk.Combobox(line3, textvariable=self.vk_pos_lieferzeit_var, width=6, state="readonly",
                     values=("Frei", "10 Uhr", "12 Uhr")).pack(side="left", padx=(0, 4))
        self.vk_pos_uhrzeit_var = tk.StringVar()
        tk.Entry(line3, textvariable=self.vk_pos_uhrzeit_var, width=7).pack(side="left", padx=(0, 2))
        tk.Label(line3, text="hh:mm", bg=BG, fg="#999", font=("Arial", 8)).pack(side="left", padx=(0, 8))
        tk.Label(line3, text="Termin:", bg=BG, font=("Arial", 9, "bold")).pack(side="left")
        self.vk_pos_termin_var = tk.StringVar(value=_next_liefertermin())
        tk.Entry(line3, textvariable=self.vk_pos_termin_var, width=11).pack(side="left", padx=(4, 12))
        tk.Button(line3, text="+ Position", command=self._vk_add_position,
                  bg="#11823b", fg="white", font=("Arial", 9, "bold"),
                  padx=10, pady=2).pack(side="left")

        # Positionsliste
        pos = tk.LabelFrame(parent, text="Positionen", bg=BG, fg=ACCENT,
                            font=("Arial", 10, "bold"), padx=12, pady=8)
        pos.pack(fill="both", expand=True, padx=8, pady=(0, 6))
        pcols = ("pzn", "artikel", "df", "pck", "apu", "menge", "rabatt", "charge", "verfall", "liefer")
        ph = {"pzn": "PZN", "artikel": "Artikel", "df": "DF", "pck": "PCK", "apu": "APU",
              "menge": "Menge", "rabatt": "Rab%", "charge": "Charge", "verfall": "Verfall",
              "liefer": "Liefervorgabe"}
        tf = tk.Frame(pos, bg=BG)
        tf.pack(fill="both", expand=True)
        tree = ttk.Treeview(tf, columns=pcols, show="headings", height=6)
        for c in pcols:
            tree.heading(c, text=ph[c])
            if c == "artikel":
                w = 170
            elif c == "liefer":
                w = 150
            else:
                w = 54
            tree.column(c, width=w, anchor="w")
        tree.pack(side="left", fill="both", expand=True)
        sb = tk.Scrollbar(tf, orient="vertical", command=tree.yview)
        sb.pack(side="right", fill="y")
        tree.configure(yscrollcommand=sb.set)
        _make_treeview_sortable(tree)
        # Vorbestellungs-/abgesagte Zeilen optisch absetzen (zaehlen nicht zur Summe).
        tree.tag_configure("vorbestellung", background="#fff8e1", foreground="#7a6000")
        tree.bind("<Delete>", self._vk_remove_position)
        tree.bind("<Double-1>", self._vk_edit_position)
        self.vk_pos_tree = tree

        afoot = tk.Frame(parent, bg=BG)
        afoot.pack(fill="x", padx=8, pady=(0, 8))
        tk.Button(afoot, text="🗑 Position löschen", command=lambda: self._vk_remove_position(None),
                  font=("Arial", 9), padx=8, pady=2).pack(side="left")
        tk.Label(afoot, text="  (oder Entf · Doppelklick = bearbeiten)",
                 bg=BG, fg="#999", font=("Arial", 8)).pack(side="left")
        tk.Button(afoot, text="Verkauf speichern", command=self._vk_save,
                  bg=ACCENT, fg="white", font=("Arial", 10, "bold"),
                  padx=14, pady=4).pack(side="right")
        self._vk_total_label = tk.Label(afoot, text="Gesamt (netto): 0,00 €", bg=BG,
                                        fg=ACCENT, font=("Arial", 13, "bold"))
        self._vk_total_label.pack(side="right", padx=(0, 18))

    def _kunde_details(self, knr):
        with self._conn() as con:
            if not knr or not _table_exists(con, "tbl_kunden_center"):
                return {}
            have = {r[1] for r in con.execute("PRAGMA table_info(tbl_kunden_center)")}
            want = [c for c in ("kundenname", "plz", "ort", "strasse", "inhaber",
                                "telefon", "email") if c in have]
            if not want:
                return {}
            row = con.execute(
                f"SELECT {','.join(want)} FROM tbl_kunden_center WHERE kundennummer=? LIMIT 1",
                (knr,)).fetchone()
            return dict(zip(want, row)) if row else {}

    def _offene_vorbestellungen(self, knr):
        if not knr:
            return 0
        with self._conn() as con:
            return con.execute(
                "SELECT COUNT(*) FROM tbl_bestellpositionen p "
                "JOIN tbl_bestellungen b ON b.id = p.bestell_id "
                "WHERE p.bestellart='Vorbestellung' AND b.kundennummer=?", (knr,)).fetchone()[0]

    def _kunde_top_artikel(self, knr, monate=12, limit=10):
        """Top-Artikel eines Kunden nach gekaufter Menge in den letzten N Monaten.
        Nur echte Bestellungen, keine stornierten Verkaeufe."""
        from datetime import date, timedelta
        cutoff = (date.today() - timedelta(days=int(monate * 30.4))).isoformat()
        with self._conn() as con:
            return con.execute(
                "SELECT p.pzn, p.artikelname, SUM(p.menge), COUNT(DISTINCT b.id), "
                "SUM(p.apu*p.menge*(1-COALESCE(p.rabatt_prozent,0)/100.0)) "
                "FROM tbl_bestellpositionen p JOIN tbl_bestellungen b ON b.id=p.bestell_id "
                "WHERE b.kundennummer=? AND p.bestellart='Bestellung' "
                "AND COALESCE(b.status,'offen')<>'storniert' AND b.datum >= ? "
                "GROUP BY p.pzn, p.artikelname ORDER BY SUM(p.menge) DESC, SUM(p.apu*p.menge) DESC "
                "LIMIT ?", (knr, cutoff, limit)).fetchall()

    def _show_kunde_top(self):
        top = self.winfo_toplevel()
        if not self.vk_kunde:
            messagebox.showinfo("Top-Artikel", "Bitte zuerst einen Kunden wählen.", parent=top)
            return
        knr = self.vk_kunde.get("kundennummer")
        rows = self._kunde_top_artikel(knr)
        win = tk.Toplevel(top)
        win.title("Top-Artikel")
        win.configure(bg=BG)
        win.transient(top)
        tk.Label(win, text=f"Top-Artikel · {self.vk_kunde.get('kundenname', '')}",
                 font=("Arial", 13, "bold"), fg=ACCENT, bg=BG).pack(padx=18, pady=(14, 2), anchor="w")
        tk.Label(win, text="Meistgekaufte Artikel der letzten 12 Monate (nur abgeschlossene "
                          "Verkäufe, ohne Stornos).", font=("Arial", 9), fg="#555",
                 bg=BG).pack(padx=18, anchor="w", pady=(0, 8))
        if not rows:
            tk.Label(win, text="Keine Käufe in den letzten 12 Monaten.", font=("Arial", 10),
                     fg="#888", bg=BG).pack(padx=18, pady=(0, 14))
        else:
            tf = tk.Frame(win, bg=BG)
            tf.pack(fill="both", expand=True, padx=18, pady=(0, 8))
            cols = ("rang", "pzn", "artikel", "menge", "anzahl", "umsatz")
            heads = {"rang": "#", "pzn": "PZN", "artikel": "Artikel", "menge": "Menge",
                     "anzahl": "Bestellungen", "umsatz": "Umsatz netto"}
            tree = ttk.Treeview(tf, columns=cols, show="headings", height=min(10, len(rows)))
            for c in cols:
                tree.heading(c, text=heads[c])
                if c == "artikel":
                    w = 240
                elif c == "rang":
                    w = 30
                else:
                    w = 90
                tree.column(c, width=w, anchor="w")
            tree.pack(side="left", fill="both", expand=True)
            sb = tk.Scrollbar(tf, orient="vertical", command=tree.yview)
            sb.pack(side="right", fill="y")
            tree.configure(yscrollcommand=sb.set)
            _make_treeview_sortable(tree)
            for i, (pzn, art, menge, anzahl, umsatz) in enumerate(rows, 1):
                tree.insert("", "end", values=(i, pzn, art, int(menge or 0), anzahl, _eur(umsatz)))
        tk.Button(win, text="Schließen", command=win.destroy, padx=14, pady=4).pack(pady=(2, 14))
        self._restyle_buttons(win)
        win.lift()
        win.focus_force()

    def _vk_pick_kunde(self, payload):
        knr, name, plz, ort = payload
        det = self._kunde_details(knr)
        self.vk_kunde = {
            "kundennummer": knr, "kundenname": det.get("kundenname") or name,
            "plz": det.get("plz") or plz, "ort": det.get("ort") or ort,
            "strasse": det.get("strasse", ""), "inhaber": det.get("inhaber", ""),
            "telefon": det.get("telefon", ""), "email": det.get("email", ""),
        }
        self._render_kunde_card(self.vk_kunde)
        # Hinweis, wenn der Kunde noch offene Vorbestellungen hat.
        n = self._offene_vorbestellungen(knr)
        if n:
            if messagebox.askyesno(
                "Offene Vorbestellungen",
                f"{self.vk_kunde['kundenname']} hat {n} offene Vorbestellung(en).\n\n"
                "Jetzt im Reiter „Vorbestellungen“ ansehen und ggf. als Verkauf übernehmen?",
                parent=self.winfo_toplevel()):
                self._vb_filter_var.set(self.vk_kunde["kundenname"] or knr)
                self._show_view("vorbestellungen")

    def _render_kunde_card(self, k):
        if not k:
            self.vk_kunde_card_label.config(text="Kein Kunde gewählt.", fg="#888",
                                            font=("Arial", 9, "italic"))
            return
        titel = k.get("kundenname") or k.get("kundennummer") or "—"
        z2 = " · ".join(x for x in [
            f"Nr. {k['kundennummer']}" if k.get("kundennummer") else "",
            f"Inhaber: {k['inhaber']}" if k.get("inhaber") else ""] if x)
        adr = " · ".join(x for x in [
            k.get("strasse", ""), f"{k.get('plz', '')} {k.get('ort', '')}".strip()] if x and x.strip())
        kontakt = " · ".join(x for x in [
            f"Tel. {k['telefon']}" if k.get("telefon") else "",
            k.get("email", "")] if x)
        txt = titel
        for z in (z2, adr, kontakt):
            if z:
                txt += "\n" + z
        self.vk_kunde_card_label.config(text=txt, fg="#11304d", font=("Arial", 9))

    def _vk_clear_kunde(self):
        """Gewaehlten Kunden im Verkauf mit einem Klick entfernen."""
        self.vk_kunde = None
        self.vk_kunde_search.clear()
        self._render_kunde_card(None)

    def _vk_pick_artikel(self, pzn):
        det = self._artikel_details(pzn)
        if not det:
            return
        knr = self.vk_kunde["kundennummer"] if self.vk_kunde else None
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

    @staticmethod
    def _liefer_text(p):
        teile = [p.get("bestellart", "Bestellung")]
        if p.get("lieferzeit"):
            teile.append(p["lieferzeit"])
        if p.get("liefertermin"):
            teile.append(p["liefertermin"])
        return " · ".join(teile)

    def _vk_row_values(self, p):
        return (p["pzn"], p["artikelname"], p["df"], p["pck"],
                p["apu"] if p["apu"] is not None else "", p["menge"],
                f"{p['rabatt']:.0f}", p["charge"], p["verfall"], self._liefer_text(p))

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
        top = self.winfo_toplevel()
        # Lieferzeit: bei "Frei" darf eine Uhrzeit (hh:mm) eingegeben werden.
        lieferzeit = self.vk_pos_lieferzeit_var.get()
        if lieferzeit == "Frei":
            uz = _normalize_uhrzeit(self.vk_pos_uhrzeit_var.get())
            if uz is None:
                messagebox.showwarning("Uhrzeit", "Bitte eine gültige Uhrzeit im Format hh:mm "
                                       "eingeben (z. B. 14:30) oder das Feld leer lassen.", parent=top)
                return
            if uz:
                lieferzeit = uz + " Uhr"
        bestellart = self.vk_pos_art_var.get()
        idx = self.vk_charge_cmb.current()
        charge, verfall = self._vk_charge_map[idx] if 0 <= idx < len(self._vk_charge_map) else ("", "")

        # #4: Bestand pruefen, wenn es eine echte Bestellung ist (nicht schon Vorbestellung).
        if bestellart == "Bestellung":
            verfuegbar = self._bestand_for(self.vk_cur["pzn"], charge, verfall) if (charge or verfall) \
                else self._bestand_total(self.vk_cur["pzn"])
            if verfuegbar < menge:
                als_vb = messagebox.askyesno(
                    "Bestand reicht nicht",
                    f"Bestand reicht nicht (verfügbar {verfuegbar}, benötigt {menge}).\n\n"
                    "Diesen Artikel als Vorbestellung aufnehmen?", parent=top)
                if als_vb:
                    bestellart = "Vorbestellung"
                else:
                    if not messagebox.askyesno(
                            "Bestellung ausführen?",
                            "Sind Sie sicher, dass Sie die Bestellung trotz fehlendem Bestand "
                            "ausführen wollen?", parent=top):
                        bestellart = "Vorbestellung"

        p = {
            "pzn": self.vk_cur["pzn"], "artikelname": self.vk_cur["artikelname"],
            "df": self.vk_cur["df"], "pck": self.vk_cur["pck"], "apu": self.vk_cur["apu"],
            "menge": menge, "rabatt": rabatt, "rabatt_quelle": self.vk_cur["rabatt_quelle"],
            "charge": charge, "verfall": verfall,
            "bestellart": bestellart,
            "lieferzeit": lieferzeit,
            "liefertermin": self.vk_pos_termin_var.get().strip(),
        }
        self.vk_positions.append(p)
        # Eingabe zuruecksetzen
        self.vk_cur = None
        self.vk_artikel_search.clear()
        self.vk_detail_label.config(text="—")
        self.vk_charge_cmb.config(values=[])
        self.vk_charge_var.set("")
        self.vk_rabatt_var.set("")
        self.vk_menge_var.set("1")
        self._vk_render_positions()
        if p["bestellart"] != "Bestellung":
            messagebox.showinfo("Vorbestellung", "Als Vorbestellung in der Liste markiert (gelb) – "
                                "wird NICHT berechnet und wandert beim Speichern in die "
                                "Vorbestellungen.", parent=self.winfo_toplevel())

    def _vk_render_positions(self):
        """Zeigt ALLE Positionen; Vorbestellungen/abgesagte werden farblich abgesetzt
        und zaehlen NICHT zur Summe (sie wandern beim Speichern in die Vorbestellungen).
        Map iid -> Position fuer Loeschen/Bearbeiten."""
        self.vk_pos_tree.delete(*self.vk_pos_tree.get_children())
        self._vk_tree_map = {}
        for pos in self.vk_positions:
            tags = () if pos.get("bestellart", "Bestellung") == "Bestellung" else ("vorbestellung",)
            iid = self.vk_pos_tree.insert("", "end", tags=tags, values=self._vk_row_values(pos))
            self._vk_tree_map[iid] = pos
        self._vk_update_total()

    def _vk_remove_position(self, _e):
        sel = self.vk_pos_tree.selection()
        if not sel:
            return
        ids = {id(self._vk_tree_map.get(i)) for i in sel}
        self.vk_positions = [x for x in self.vk_positions if id(x) not in ids]
        self._vk_render_positions()

    def _vk_update_total(self):
        # Nur echte Bestellungen zaehlen; Vorbestellungen/abgesagte nicht.
        gesamt = sum(_pos_netto(p) for p in self.vk_positions
                     if p.get("bestellart", "Bestellung") == "Bestellung")
        self._vk_total_label.config(text=f"Gesamt (netto): {_eur(gesamt)}")

    def _vk_edit_position(self, _e=None):
        sel = self.vk_pos_tree.selection()
        if not sel:
            return
        p = self._vk_tree_map.get(sel[0])
        if p is None:
            return
        self._charge_dialog(p, sel[0])

    def _charge_dialog(self, p, item):
        top = self.winfo_toplevel()
        win = tk.Toplevel(top)
        win.title("Position bearbeiten")
        win.configure(bg=BG)
        win.transient(top)
        win.resizable(False, False)

        tk.Label(win, text=p["artikelname"] or p["pzn"], font=("Arial", 13, "bold"),
                 fg=ACCENT, bg=BG).pack(padx=18, pady=(14, 2), anchor="w")
        tk.Label(win, text=f"PZN {p['pzn']}  ·  DF {p['df'] or '—'}  ·  PCK {p['pck'] or '—'}  "
                          f"·  APU {_eur(p['apu'])}", font=("Arial", 9), fg="#555",
                 bg=BG).pack(padx=18, anchor="w")
        tk.Label(win, text=f"Aktuell gewählt: Charge {p['charge'] or '—'}  ·  Verfall {p['verfall'] or '—'}",
                 font=("Arial", 9, "italic"), fg="#888", bg=BG).pack(padx=18, pady=(2, 8), anchor="w")

        tk.Label(win, text="Verfügbare Chargen dieses Artikels (Doppelklick = übernehmen):",
                 font=("Arial", 9, "bold"), bg=BG).pack(padx=18, anchor="w")
        tf = tk.Frame(win, bg=BG)
        tf.pack(fill="both", padx=18, pady=(4, 8))
        cols = ("charge", "verfall", "bestand")
        tree = ttk.Treeview(tf, columns=cols, show="headings", height=6, selectmode="browse")
        for c, t, w in (("charge", "Charge", 150), ("verfall", "Verfall", 120), ("bestand", "Bestand", 80)):
            tree.heading(c, text=t)
            tree.column(c, width=w, anchor="w")
        tree.pack(side="left", fill="both", expand=True)
        sb = tk.Scrollbar(tf, orient="vertical", command=tree.yview)
        sb.pack(side="right", fill="y")
        tree.configure(yscrollcommand=sb.set)
        _make_treeview_sortable(tree)

        rowmap = {}
        first = tree.insert("", "end", values=("(ohne Charge)", "", ""))
        rowmap[first] = ("", "")
        for charge, verfall, menge in self._lager_chargen(p["pzn"]):
            iid = tree.insert("", "end", values=(charge or "—", verfall or "—", menge))
            rowmap[iid] = (charge or "", verfall or "")

        mrow = tk.Frame(win, bg=BG)
        mrow.pack(fill="x", padx=18, pady=(0, 6))
        tk.Label(mrow, text="Menge:", font=("Arial", 9, "bold"), bg=BG).pack(side="left")
        menge_var = tk.StringVar(value=str(p["menge"]))
        tk.Entry(mrow, textvariable=menge_var, width=6).pack(side="left", padx=(6, 0))

        lrow = tk.Frame(win, bg=BG)
        lrow.pack(fill="x", padx=18, pady=(0, 8))
        tk.Label(lrow, text="Liefervorgabe:", font=("Arial", 9, "bold"), bg=BG).pack(side="left")
        art_var = tk.StringVar(value=p.get("bestellart", "Bestellung"))
        ttk.Combobox(lrow, textvariable=art_var, width=13, state="readonly",
                     values=("Bestellung", "Vorbestellung", "abgesagt")).pack(side="left", padx=(4, 8))
        lz_var = tk.StringVar(value=p.get("lieferzeit", "") or "10 Uhr")
        ttk.Combobox(lrow, textvariable=lz_var, width=8,
                     values=("10 Uhr", "12 Uhr", "Frei")).pack(side="left", padx=(0, 8))
        tk.Label(lrow, text="Termin:", font=("Arial", 9, "bold"), bg=BG).pack(side="left")
        termin_var = tk.StringVar(value=p.get("liefertermin", ""))
        tk.Entry(lrow, textvariable=termin_var, width=11).pack(side="left", padx=(4, 0))

        def uebernehmen(_e=None):
            sel = tree.selection()
            if sel:
                p["charge"], p["verfall"] = rowmap.get(sel[0], (p["charge"], p["verfall"]))
            try:
                m = int(menge_var.get().strip())
                if m > 0:
                    p["menge"] = m
            except ValueError:
                pass
            p["bestellart"] = art_var.get()
            p["lieferzeit"] = lz_var.get()
            p["liefertermin"] = termin_var.get().strip()
            self._vk_render_positions()
            win.destroy()

        tree.bind("<Double-1>", uebernehmen)
        btns = tk.Frame(win, bg=BG)
        btns.pack(padx=18, pady=(0, 14), anchor="e")
        tk.Button(btns, text="Übernehmen", command=uebernehmen, bg=ACCENT, fg="white",
                  font=("Arial", 10, "bold"), padx=12, pady=4).pack(side="right", padx=(8, 0))
        tk.Button(btns, text="Abbrechen", command=win.destroy, padx=12, pady=4).pack(side="right")
        self._restyle_buttons(win)
        win.lift()
        win.focus_force()

    def _vk_save(self):
        top = self.winfo_toplevel()
        if not self.vk_kunde:
            messagebox.showwarning("Verkauf", "Bitte einen Kunden wählen.", parent=top)
            return
        if not self.vk_positions:
            messagebox.showwarning("Verkauf", "Keine Positionen erfasst.", parent=top)
            return
        # Kopf-Liefervorgabe aus den Positionen ableiten (einheitlich oder "gemischt").
        arten = {p.get("bestellart", "Bestellung") for p in self.vk_positions}
        zeiten = {p.get("lieferzeit", "") for p in self.vk_positions}
        termine = {p.get("liefertermin", "") for p in self.vk_positions}
        header = {
            "datum": datetime.now().strftime("%Y-%m-%d"),
            "kundennummer": self.vk_kunde["kundennummer"], "apotheke": self.vk_kunde["kundenname"],
            "bestellart": next(iter(arten)) if len(arten) == 1 else "gemischt",
            "lieferzeit": next(iter(zeiten)) if len(zeiten) == 1 else "",
            "liefertermin": next(iter(termine)) if len(termine) == 1 else "",
        }
        bid = self._save_verkauf(header, self.vk_positions)
        anzahl = len(self.vk_positions)
        bestellungen = sum(1 for p in self.vk_positions if p.get("bestellart") == "Bestellung")
        vorbestellungen = sum(1 for p in self.vk_positions if p.get("bestellart") == "Vorbestellung")
        # Aus Vorbestellungen uebernommene Positionen: Original-Vorbestellung um die
        # uebernommene Menge reduzieren (Rest bleibt Vorbestellung); nur bei voller
        # Uebernahme loeschen.
        herkunft = [(p["_vb_source"], p.get("_vb_orig_menge", p["menge"]), p["menge"])
                    for p in self.vk_positions if p.get("_vb_source")]
        if herkunft:
            with self._conn() as con:
                for src, orig, genommen in herkunft:
                    rest = (orig or 0) - (genommen or 0)
                    if rest > 0:
                        con.execute("UPDATE tbl_bestellpositionen SET menge=? WHERE id=?", (rest, src))
                    else:
                        con.execute("DELETE FROM tbl_bestellpositionen WHERE id=?", (src,))
                con.commit()
        self._log("Verkauf gespeichert", bid, self.vk_kunde["kundenname"],
                  f"{bestellungen} Bestellung(en), {vorbestellungen} Vorbestellung(en)"
                  + (f", {len(herkunft)} aus Vorbestellung übernommen" if herkunft else ""))
        # Reset
        self.vk_positions = []
        self._vk_render_positions()
        self.vk_kunde = None
        self.vk_kunde_search.clear()
        self._render_kunde_card(None)
        self._refresh_lager()
        if hasattr(self, "vb_tree"):
            self._refresh_vorbestellungen()
        if hasattr(self, "vh_tree"):
            self._refresh_verkaeufe()
        # #1: Auftragsbestaetigung nur, wenn echte Bestellungen dabei sind.
        if bestellungen:
            self._auftrag_dialog(bid, anzahl, vorbestellungen)
        else:
            messagebox.showinfo("Vorbestellung",
                                f"{vorbestellungen} Vorbestellung(en) gespeichert "
                                "(keine Auftragsbestätigung).", parent=top)

    # ==================================================== Auftragsbestaetigung
    def _auftrag_dialog(self, bestell_id, anzahl, vorbestellungen=0):
        top = self.winfo_toplevel()
        win = tk.Toplevel(top)
        win.title("Auftragsbestätigung")
        win.configure(bg=BG)
        win.transient(top)
        win.resizable(False, False)
        tk.Label(win, text=f"✓ Verkauf #{bestell_id} gespeichert",
                 font=("Arial", 12, "bold"), fg="#11823b", bg=BG).pack(padx=20, pady=(16, 2))
        info = f"{anzahl} Position(en) · Lagerbestand abgebucht."
        if vorbestellungen:
            info += (f"\n{vorbestellungen} davon Vorbestellung(en) – noch nicht abgebucht, "
                     "im Reiter „Vorbestellungen“ bestätigen, wenn Ware da ist.")
        tk.Label(win, text=info, font=("Arial", 9), fg="#666", bg=BG,
                 justify="left").pack(padx=20, pady=(0, 12))

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
        self._restyle_buttons(win)
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

    # ========================================================== Vorbestellungen
    def _build_vorbestellungen(self, parent):
        bar = tk.Frame(parent, bg=BG)
        bar.pack(fill="x", padx=8, pady=(10, 4))
        self.vb_info = tk.Label(bar, text="", bg=BG, fg=ACCENT, font=("Arial", 11, "bold"))
        self.vb_info.pack(side="left")
        tk.Button(bar, text="🔄 Aktualisieren", command=self._refresh_vorbestellungen,
                  font=("Arial", 9), padx=8, pady=2).pack(side="right")

        srow = tk.Frame(parent, bg=BG)
        srow.pack(fill="x", padx=8, pady=(0, 4))
        tk.Label(srow, text="Suche (Kunde / Artikel / PZN):", bg=BG,
                 font=("Arial", 9, "bold")).pack(side="left")
        self._vb_filter_var = tk.StringVar()
        e = tk.Entry(srow, textvariable=self._vb_filter_var, width=26)
        e.pack(side="left", padx=(6, 8))
        e.bind("<KeyRelease>", lambda _e: self._refresh_vorbestellungen())
        tk.Button(srow, text="✕", command=lambda: (self._vb_filter_var.set(""),
                  self._refresh_vorbestellungen()), font=("Arial", 8), padx=6).pack(side="left")
        tk.Label(srow, text="Status:", bg=BG, font=("Arial", 9, "bold")).pack(side="left", padx=(12, 0))
        self._vb_status_var = tk.StringVar(value="Offen")
        ttk.Combobox(srow, textvariable=self._vb_status_var, width=11, state="readonly",
                     values=("Offen", "Abgesagt", "Alle")).pack(side="left", padx=(4, 0))
        self._vb_status_var.trace_add("write", lambda *_: self._refresh_vorbestellungen())
        tk.Button(srow, text="📥 Vorbestellungen importieren", command=self._import_vorbestellungen_datei,
                  font=("Arial", 8), padx=6, pady=1).pack(side="right")
        tk.Label(parent, text="Disposition: Kunden anrufen, dann Doppelklick → bearbeiten, "
                              "„Als Verkauf bestätigen“ oder stornieren.",
                 bg=BG, fg="#666", font=("Arial", 9)).pack(anchor="w", padx=8, pady=(0, 4))

        tf = tk.Frame(parent, bg=BG)
        tf.pack(fill="both", expand=True, padx=8, pady=(0, 8))
        cols = ("datum", "kunde", "telefon", "pzn", "artikel", "menge", "lieferzeit", "termin")
        heads = {"datum": "Datum", "kunde": "Kunde", "telefon": "Telefon", "pzn": "PZN",
                 "artikel": "Artikel", "menge": "Menge", "lieferzeit": "Lieferzeit", "termin": "Termin"}
        tree = ttk.Treeview(tf, columns=cols, show="headings")
        for c in cols:
            tree.heading(c, text=heads[c])
            if c == "artikel":
                w = 190
            elif c in ("kunde", "telefon"):
                w = 130
            else:
                w = 76
            tree.column(c, width=w, anchor="w")
        tree.pack(side="left", fill="both", expand=True)
        sb = tk.Scrollbar(tf, orient="vertical", command=tree.yview)
        sb.pack(side="right", fill="y")
        tree.configure(yscrollcommand=sb.set)
        _make_treeview_sortable(tree)
        tree.bind("<Double-1>", self._vorbestellung_dialog)
        self.vb_tree = tree
        self._vb_rowmap = {}

    def _refresh_vorbestellungen(self):
        if not hasattr(self, "vb_tree"):
            return
        self.vb_tree.delete(*self.vb_tree.get_children())
        self._vb_rowmap = {}
        filt = (self._vb_filter_var.get() if hasattr(self, "_vb_filter_var") else "").strip().lower()
        status_f = self._vb_status_var.get() if hasattr(self, "_vb_status_var") else "Offen"
        if status_f == "Abgesagt":
            arten = ("abgesagt",)
        elif status_f == "Alle":
            arten = ("Vorbestellung", "abgesagt")
        else:
            arten = ("Vorbestellung",)
        platzhalter = ",".join("?" for _ in arten)
        has_kunden = self._table_exists_now("tbl_kunden_center")
        tel = "COALESCE(k.telefon,'')" if has_kunden else "''"
        join = ("LEFT JOIN tbl_kunden_center k ON k.kundennummer = b.kundennummer"
                if has_kunden else "")
        with self._conn() as con:
            rows = con.execute(
                f"SELECT p.id, b.datum, COALESCE(b.apotheke,''), {tel}, p.pzn, p.artikelname, "
                f"p.menge, COALESCE(p.lieferzeit,''), COALESCE(p.liefertermin,'') "
                f"FROM tbl_bestellpositionen p JOIN tbl_bestellungen b ON b.id = p.bestell_id "
                f"{join} WHERE p.bestellart IN ({platzhalter}) ORDER BY p.liefertermin, b.datum",
                arten
            ).fetchall()
        gezeigt = 0
        for r in rows:
            if filt and filt not in " ".join(str(x).lower() for x in (r[2], r[4], r[5])):
                continue
            iid = self.vb_tree.insert("", "end", values=r[1:])
            self._vb_rowmap[iid] = r[0]
            gezeigt += 1
        suffix = f" (gefiltert aus {len(rows)})" if filt else ""
        label = {"Abgesagt": "abgesagte", "Alle": "Vorbestellung(en) gesamt"}.get(status_f, "offene")
        if status_f == "Alle":
            self.vb_info.config(text=f"{gezeigt} {label}{suffix}")
        else:
            self.vb_info.config(text=f"{gezeigt} {label} Vorbestellung(en){suffix}")

    def _vorbestellung_dialog(self, _e=None):
        sel = self.vb_tree.selection()
        if not sel:
            return
        pid = self._vb_rowmap.get(sel[0])
        if not pid:
            return
        with self._conn() as con:
            row = con.execute(
                "SELECT p.pzn, p.artikelname, p.df, p.pck, p.apu, p.menge, p.charge, p.verfall, "
                "COALESCE(p.lieferzeit,''), COALESCE(p.liefertermin,''), COALESCE(b.apotheke,'') "
                "FROM tbl_bestellpositionen p JOIN tbl_bestellungen b ON b.id = p.bestell_id "
                "WHERE p.id=?", (pid,)).fetchone()
        if not row:
            return
        pzn, artikel, df, pck, apu, menge, charge, verfall, lieferzeit, termin, apotheke = row

        top = self.winfo_toplevel()
        win = tk.Toplevel(top)
        win.title("Vorbestellung")
        win.configure(bg=BG)
        win.transient(top)
        win.resizable(False, False)
        tk.Label(win, text=artikel or pzn, font=("Arial", 13, "bold"), fg=ACCENT,
                 bg=BG).pack(padx=18, pady=(14, 2), anchor="w")
        tk.Label(win, text=f"PZN {pzn}  ·  DF {df or '—'}  ·  PCK {pck or '—'}  ·  APU {_eur(apu)}",
                 font=("Arial", 9), fg="#555", bg=BG).pack(padx=18, anchor="w")
        tk.Label(win, text=f"Kunde: {apotheke or '—'}", font=("Arial", 9), fg="#555",
                 bg=BG).pack(padx=18, pady=(0, 8), anchor="w")

        erow = tk.Frame(win, bg=BG)
        erow.pack(fill="x", padx=18, pady=(0, 6))
        tk.Label(erow, text="Menge:", font=("Arial", 9, "bold"), bg=BG).pack(side="left")
        menge_var = tk.StringVar(value=str(menge))
        tk.Entry(erow, textvariable=menge_var, width=6).pack(side="left", padx=(4, 12))
        lz_var = tk.StringVar(value=lieferzeit or "10 Uhr")
        ttk.Combobox(erow, textvariable=lz_var, width=8,
                     values=("10 Uhr", "12 Uhr", "Frei")).pack(side="left", padx=(0, 8))
        tk.Label(erow, text="Termin:", font=("Arial", 9, "bold"), bg=BG).pack(side="left")
        termin_var = tk.StringVar(value=termin)
        tk.Entry(erow, textvariable=termin_var, width=11).pack(side="left", padx=(4, 0))

        tk.Label(win, text="Charge wählen (Doppelklick oder markieren – wird beim Übernehmen "
                          "in den Verkauf mitgenommen):",
                 font=("Arial", 9, "bold"), bg=BG).pack(padx=18, anchor="w")
        tf = tk.Frame(win, bg=BG)
        tf.pack(fill="both", padx=18, pady=(4, 8))
        ctree = ttk.Treeview(tf, columns=("charge", "verfall", "bestand"), show="headings",
                             height=4, selectmode="browse")
        for c, t, w in (("charge", "Charge", 150), ("verfall", "Verfall", 120), ("bestand", "Bestand", 80)):
            ctree.heading(c, text=t)
            ctree.column(c, width=w, anchor="w")
        ctree.pack(side="left", fill="both", expand=True)
        sb = tk.Scrollbar(tf, orient="vertical", command=ctree.yview)
        sb.pack(side="right", fill="y")
        ctree.configure(yscrollcommand=sb.set)
        _make_treeview_sortable(ctree)
        rowmap = {}
        f0 = ctree.insert("", "end", values=("(ohne Charge)", "", ""))
        rowmap[f0] = ("", "")
        for ch, vf, m in self._lager_chargen(pzn):
            iid = ctree.insert("", "end", values=(ch or "—", vf or "—", m))
            rowmap[iid] = (ch or "", vf or "")

        def _menge():
            try:
                v = int(menge_var.get().strip())
                return v if v > 0 else menge
            except ValueError:
                return menge

        def _charge():
            sel = ctree.selection()
            return rowmap.get(sel[0], ("", "")) if sel else ("", "")

        def speichern():
            with self._conn() as con:
                con.execute("UPDATE tbl_bestellpositionen SET menge=?, lieferzeit=?, liefertermin=? WHERE id=?",
                            (_menge(), lz_var.get(), termin_var.get().strip(), pid))
                con.commit()
            self._log("Vorbestellung geändert", kunde=apotheke,
                      details=f"{artikel or pzn}, Menge {_menge()}")
            self._refresh_vorbestellungen()
            win.destroy()

        def uebernehmen():
            # In den Verkauf laden (NICHT direkt verkaufen). Gewaehlte Charge mitnehmen.
            ch, vf = _charge()
            if self._vb_in_verkauf(pid, _menge(), lz_var.get(), termin_var.get().strip(), ch, vf):
                win.destroy()

        ctree.bind("<Double-1>", lambda _e: uebernehmen())

        def stornieren():
            if not messagebox.askyesno(
                    "Vorbestellung stornieren",
                    f"Vorbestellung „{artikel or pzn}“ für {apotheke or 'diesen Kunden'} "
                    "wirklich absagen?", parent=win):
                return
            with self._conn() as con:
                con.execute("UPDATE tbl_bestellpositionen SET bestellart='abgesagt' WHERE id=?", (pid,))
                con.commit()
            self._log("Vorbestellung abgesagt", kunde=apotheke, details=f"{artikel or pzn}")
            self._refresh_vorbestellungen()
            win.destroy()

        btns = tk.Frame(win, bg=BG)
        btns.pack(fill="x", padx=18, pady=(0, 14))
        tk.Button(btns, text="➡ In Verkauf übernehmen", command=uebernehmen, bg="#11823b",
                  fg="white", font=("Arial", 10, "bold"), padx=12, pady=4).pack(side="left")
        tk.Button(btns, text="Speichern", command=speichern, bg=ACCENT, fg="white",
                  font=("Arial", 10, "bold"), padx=12, pady=4).pack(side="left", padx=(8, 0))
        tk.Button(btns, text="Stornieren", command=stornieren, padx=12, pady=4).pack(side="right")
        self._restyle_buttons(win)
        win.lift()
        win.focus_force()

    def _vb_in_verkauf(self, pid, menge, lieferzeit, termin, charge="", verfall=""):
        """Laedt eine Vorbestellung als Position in den Verkauf-Reiter. Das Original
        wird erst beim Speichern des Verkaufs entfernt (_vb_source).
        Liefert True, wenn uebernommen; False bei Abbruch (z.B. anderer Kunde)."""
        with self._conn() as con:
            row = con.execute(
                "SELECT p.pzn, p.artikelname, p.df, p.pck, p.apu, p.rabatt_prozent, p.rabatt_quelle, "
                "b.kundennummer, COALESCE(b.apotheke,''), p.menge "
                "FROM tbl_bestellpositionen p JOIN tbl_bestellungen b ON b.id=p.bestell_id "
                "WHERE p.id=?", (pid,)).fetchone()
        if not row:
            return False
        pzn, artikel, df, pck, apu, rabatt, rquelle, knr, apotheke, orig_menge = row
        # Konflikt: im Verkauf liegt schon ein ANDERER Kunde.
        if (self.vk_kunde and self.vk_kunde.get("kundennummer") != knr) or \
           (self.vk_positions and not self.vk_kunde):
            messagebox.showwarning(
                "Unterschiedlicher Kunde im Verkauf",
                f"Im Verkauf liegt bereits ein anderer Kunde "
                f"({(self.vk_kunde or {}).get('kundenname', 'unbekannt')}).\n\n"
                "Bitte den laufenden Auftrag erst löschen oder abschließen und dann "
                "die Vorbestellung holen.", parent=self.winfo_toplevel())
            return False
        # Kunde setzen.
        det = self._kunde_details(knr)
        self.vk_kunde = {
            "kundennummer": knr, "kundenname": det.get("kundenname") or apotheke,
            "plz": det.get("plz", ""), "ort": det.get("ort", ""), "strasse": det.get("strasse", ""),
            "inhaber": det.get("inhaber", ""), "telefon": det.get("telefon", ""),
            "email": det.get("email", ""),
        }
        self._render_kunde_card(self.vk_kunde)
        p = {
            "pzn": pzn, "artikelname": artikel, "df": df or "", "pck": pck or "", "apu": apu,
            "menge": menge, "rabatt": rabatt or 0, "rabatt_quelle": rquelle or "manuell",
            "charge": charge or "", "verfall": verfall or "", "bestellart": "Bestellung",
            "lieferzeit": lieferzeit, "liefertermin": termin,
            "_vb_source": pid, "_vb_orig_menge": orig_menge or 0,
        }
        self.vk_positions.append(p)
        self._vk_render_positions()
        self._show_view("verkauf")
        messagebox.showinfo("Vorbestellung",
                            "In den Verkauf übernommen. Charge/Bestand wählen, ggf. weitere Artikel "
                            "hinzufügen, dann „Verkauf speichern“.", parent=self.winfo_toplevel())
        return True

    # =============================================================== Verkaeufe
    def _build_verkaeufe(self, parent):
        bar = tk.Frame(parent, bg=BG)
        bar.pack(fill="x", padx=8, pady=(10, 4))
        tk.Label(bar, text="Suche (Kunde):", bg=BG, font=("Arial", 9, "bold")).pack(side="left")
        self._vh_filter_var = tk.StringVar()
        e = tk.Entry(bar, textvariable=self._vh_filter_var, width=24)
        e.pack(side="left", padx=(6, 8))
        e.bind("<KeyRelease>", lambda _e: self._refresh_verkaeufe())
        tk.Label(bar, text="Status:", bg=BG, font=("Arial", 9, "bold")).pack(side="left")
        self._vh_status_var = tk.StringVar(value="Alle")
        ttk.Combobox(bar, textvariable=self._vh_status_var, width=11, state="readonly",
                     values=("Alle", "Aktiv", "Storniert")).pack(side="left", padx=(4, 0))
        self._vh_status_var.trace_add("write", lambda *_: self._refresh_verkaeufe())
        tk.Button(bar, text="📥 Verkäufe importieren", command=self._import_verkaeufe_datei,
                  font=("Arial", 8), padx=6, pady=1).pack(side="left", padx=(12, 0))
        tk.Button(bar, text="🔄 Aktualisieren", command=self._refresh_verkaeufe,
                  font=("Arial", 9), padx=8, pady=2).pack(side="right")
        self.vh_info = tk.Label(bar, text="", bg=BG, fg=ACCENT, font=("Arial", 9, "bold"))
        self.vh_info.pack(side="right", padx=(0, 12))

        # Zeitraum-Filter
        zrow = tk.Frame(parent, bg=BG)
        zrow.pack(fill="x", padx=8, pady=(0, 4))
        tk.Label(zrow, text="Zeitraum:", bg=BG, font=("Arial", 9, "bold")).pack(side="left")
        self._vh_zeit_var = tk.StringVar(value="Alle")
        ttk.Combobox(zrow, textvariable=self._vh_zeit_var, width=16, state="readonly",
                     values=("Alle", "Aktueller Monat", "Vormonat", "Aktuelles Quartal",
                             "Aktuelles Jahr", "Frei")).pack(side="left", padx=(4, 12))
        self._vh_zeit_var.trace_add("write", lambda *_: self._refresh_verkaeufe())
        tk.Label(zrow, text="von:", bg=BG, font=("Arial", 9)).pack(side="left")
        self._vh_von_var = tk.StringVar()
        ev = tk.Entry(zrow, textvariable=self._vh_von_var, width=11)
        ev.pack(side="left", padx=(3, 8))
        tk.Label(zrow, text="bis:", bg=BG, font=("Arial", 9)).pack(side="left")
        self._vh_bis_var = tk.StringVar()
        eb = tk.Entry(zrow, textvariable=self._vh_bis_var, width=11)
        eb.pack(side="left", padx=(3, 6))
        tk.Label(zrow, text="(TT.MM.JJJJ – nur bei „Frei“)", bg=BG, fg="#999",
                 font=("Arial", 8)).pack(side="left")
        ev.bind("<KeyRelease>", lambda _e: self._vh_zeit_var.get() == "Frei" and self._refresh_verkaeufe())
        eb.bind("<KeyRelease>", lambda _e: self._vh_zeit_var.get() == "Frei" and self._refresh_verkaeufe())
        tk.Label(zrow, text="MSK:", bg=BG, font=("Arial", 9, "bold")).pack(side="left", padx=(14, 0))
        self._vh_msk_var = tk.StringVar(value="Alle")
        ttk.Combobox(zrow, textvariable=self._vh_msk_var, width=14, state="readonly",
                     values=("Alle", "MSK offen", "MSK erfasst")).pack(side="left", padx=(4, 0))
        self._vh_msk_var.trace_add("write", lambda *_: self._refresh_verkaeufe())

        # MSK-Schnellmarkierung der ausgewaehlten Zeilen.
        mrow = tk.Frame(parent, bg=BG)
        mrow.pack(fill="x", padx=8, pady=(0, 2))
        tk.Label(mrow, text="Markierte Verkäufe:", bg=BG, font=("Arial", 9, "bold")).pack(side="left")
        tk.Button(mrow, text="✓ In MSK erfasst", command=lambda: self._msk_markieren(True),
                  bg="#11823b", fg="white", font=("Arial", 9, "bold"), padx=8, pady=2).pack(side="left", padx=(6, 4))
        tk.Button(mrow, text="↺ MSK offen", command=lambda: self._msk_markieren(False),
                  font=("Arial", 9), padx=8, pady=2).pack(side="left")
        tk.Label(mrow, text="(mehrere mit Strg/Shift markierbar · Doppelklick = Details)",
                 bg=BG, fg="#999", font=("Arial", 8)).pack(side="left", padx=(10, 0))

        tf = tk.Frame(parent, bg=BG)
        tf.pack(fill="both", expand=True, padx=8, pady=(0, 8))
        cols = ("datum", "kunde", "bearbeiter", "msk", "status", "pos", "summe")
        heads = {"datum": "Datum", "kunde": "Kunde", "bearbeiter": "Erfasst von",
                 "msk": "MSK", "status": "Status", "pos": "Pos.", "summe": "Summe (netto)"}
        tree = ttk.Treeview(tf, columns=cols, show="headings", selectmode="extended")
        for c in cols:
            tree.heading(c, text=heads[c])
            w = 180 if c == "kunde" else (110 if c in ("summe",) else (90 if c in ("bearbeiter", "msk", "status") else 70))
            tree.column(c, width=w, anchor="w")
        tree.tag_configure("msk_erfasst", background="#e7f4ea")
        tree.pack(side="left", fill="both", expand=True)
        sb = tk.Scrollbar(tf, orient="vertical", command=tree.yview)
        sb.pack(side="right", fill="y")
        tree.configure(yscrollcommand=sb.set)
        _make_treeview_sortable(tree)
        tree.bind("<Double-1>", self._verkauf_detail_dialog)
        self.vh_tree = tree
        self._vh_rowmap = {}

    def _zeitraum_grenzen(self, modus):
        """Liefert (von_iso, bis_iso) fuer den gewaehlten Zeitraum; (None,None)=Alle."""
        import calendar
        from datetime import date as _date
        h = _date.today()
        if modus == "Aktueller Monat":
            von, bis = h.replace(day=1), h.replace(day=calendar.monthrange(h.year, h.month)[1])
        elif modus == "Vormonat":
            jahr, monat = (h.year, h.month - 1) if h.month > 1 else (h.year - 1, 12)
            von, bis = _date(jahr, monat, 1), _date(jahr, monat, calendar.monthrange(jahr, monat)[1])
        elif modus == "Aktuelles Quartal":
            q = (h.month - 1) // 3
            endm = q * 3 + 3
            von = _date(h.year, q * 3 + 1, 1)
            bis = _date(h.year, endm, calendar.monthrange(h.year, endm)[1])
        elif modus == "Aktuelles Jahr":
            von, bis = _date(h.year, 1, 1), _date(h.year, 12, 31)
        elif modus == "Frei":
            return (kasse_import._parse_datum(self._vh_von_var.get()),
                    kasse_import._parse_datum(self._vh_bis_var.get()))
        else:
            return None, None
        return von.isoformat(), bis.isoformat()

    def _refresh_verkaeufe(self):
        if not hasattr(self, "vh_tree"):
            return
        self.vh_tree.delete(*self.vh_tree.get_children())
        self._vh_rowmap = {}
        filt = self._vh_filter_var.get().strip().lower()
        status_f = self._vh_status_var.get() if hasattr(self, "_vh_status_var") else "Alle"
        zeit_f = self._vh_zeit_var.get() if hasattr(self, "_vh_zeit_var") else "Alle"
        msk_f = self._vh_msk_var.get() if hasattr(self, "_vh_msk_var") else "Alle"
        von, bis = self._zeitraum_grenzen(zeit_f)
        with self._conn() as con:
            rows = con.execute(
                "SELECT b.id, b.datum, COALESCE(b.apotheke,''), COALESCE(b.status,'offen'), "
                "COALESCE(b.bearbeiter,''), COALESCE(b.msk_erfasst,0), COUNT(p.id), "
                "COALESCE(SUM(CASE WHEN p.bestellart='Bestellung' "
                "  THEN p.apu*p.menge*(1-COALESCE(p.rabatt_prozent,0)/100.0) ELSE 0 END),0) "
                "FROM tbl_bestellungen b LEFT JOIN tbl_bestellpositionen p ON p.bestell_id=b.id "
                "GROUP BY b.id "
                "HAVING SUM(CASE WHEN p.bestellart='Bestellung' THEN 1 ELSE 0 END) > 0 "
                "ORDER BY b.id DESC").fetchall()
        summe_ges = 0.0
        gezeigt = storniert_anz = msk_offen_anz = 0
        for bid, datum, apotheke, status, bearbeiter, msk, anz, summe in rows:
            if filt and filt not in str(apotheke).lower():
                continue
            if status_f == "Storniert" and status != "storniert":
                continue
            if status_f == "Aktiv" and status == "storniert":
                continue
            if msk_f == "MSK offen" and msk:
                continue
            if msk_f == "MSK erfasst" and not msk:
                continue
            if von and (datum or "") < von:
                continue
            if bis and (datum or "") > bis:
                continue
            # MSK nur fuer aktive (nicht stornierte) Verkaeufe relevant.
            offen = (not msk) and status != "storniert"
            erfasst = bool(msk) and status != "storniert"
            msk_txt = "✓ erfasst" if msk else ("offen" if status != "storniert" else "–")
            iid = self.vh_tree.insert("", "end", tags=("msk_erfasst",) if erfasst else (),
                                      values=(datum, apotheke, bearbeiter, msk_txt, status,
                                              anz, _eur(summe)))
            self._vh_rowmap[iid] = bid
            if status != "storniert":
                summe_ges += summe or 0
            else:
                storniert_anz += 1
            if offen:
                msk_offen_anz += 1
            gezeigt += 1
        self.vh_info.config(text=f"{gezeigt} Verkäufe · ⚠ {msk_offen_anz} MSK offen · "
                                 f"{storniert_anz} storniert · Summe {_eur(summe_ges)}")

    def _msk_markieren(self, erfasst, bids=None):
        if bids is None:
            sel = self.vh_tree.selection()
            if not sel:
                messagebox.showinfo("MSK", "Bitte erst einen oder mehrere Verkäufe in der Liste "
                                    "markieren.", parent=self.winfo_toplevel())
                return
            bids = [self._vh_rowmap.get(i) for i in sel if self._vh_rowmap.get(i)]
        jetzt = datetime.now().isoformat(timespec="seconds")
        von = getpass.getuser()
        with self._conn() as con:
            for bid in bids:
                if erfasst:
                    con.execute("UPDATE tbl_bestellungen SET msk_erfasst=1, msk_von=?, msk_am=? "
                                "WHERE id=?", (von, jetzt, bid))
                else:
                    con.execute("UPDATE tbl_bestellungen SET msk_erfasst=0, msk_von=NULL, msk_am=NULL "
                                "WHERE id=?", (bid,))
            con.commit()
        for bid in bids:
            self._log("MSK erfasst" if erfasst else "MSK offen gesetzt", bid)
        self._refresh_verkaeufe()

    def _verkauf_detail_dialog(self, _e=None):
        sel = self.vh_tree.selection()
        if not sel:
            return
        bid = self._vh_rowmap.get(sel[0])
        if not bid:
            return
        with self._conn() as con:
            h = con.execute("SELECT datum, COALESCE(apotheke,''), COALESCE(status,'offen'), "
                            "COALESCE(bearbeiter,''), COALESCE(msk_erfasst,0), COALESCE(msk_von,''), "
                            "COALESCE(msk_am,'') FROM tbl_bestellungen WHERE id=?", (bid,)).fetchone()
            positions = con.execute(
                "SELECT pzn, artikelname, menge, COALESCE(rabatt_prozent,0), apu, "
                "COALESCE(charge,''), COALESCE(verfall,''), COALESCE(bestellart,'Bestellung') "
                "FROM tbl_bestellpositionen WHERE bestell_id=? ORDER BY id", (bid,)).fetchall()
        if not h:
            return
        datum, apotheke, status, bearbeiter, msk, msk_von, msk_am = h
        top = self.winfo_toplevel()
        win = tk.Toplevel(top)
        win.title(f"Verkauf #{bid}")
        win.configure(bg=BG)
        win.transient(top)
        tk.Label(win, text=f"Verkauf #{bid} · {apotheke}", font=("Arial", 13, "bold"),
                 fg=ACCENT, bg=BG).pack(padx=18, pady=(14, 2), anchor="w")
        tk.Label(win, text=f"Datum {datum} · Status: {status} · Erfasst von: {bearbeiter or '–'}",
                 font=("Arial", 9), fg="#555", bg=BG).pack(padx=18, anchor="w")
        msk_txt = (f"MSK: ✓ erfasst von {msk_von or '?'} am {(msk_am or '')[:16].replace('T', ' ')}"
                   if msk else "MSK: ⚠ noch nicht erfasst")
        tk.Label(win, text=msk_txt, font=("Arial", 9, "bold"),
                 fg=("#11823b" if msk else "#a35a00"), bg=BG).pack(padx=18, anchor="w", pady=(0, 8))
        tf = tk.Frame(win, bg=BG)
        tf.pack(fill="both", expand=True, padx=18, pady=(0, 8))
        cols = ("pzn", "artikel", "menge", "rabatt", "charge", "verfall", "art")
        heads = {"pzn": "PZN", "artikel": "Artikel", "menge": "Menge", "rabatt": "Rab%",
                 "charge": "Charge", "verfall": "Verfall", "art": "Art"}
        tree = ttk.Treeview(tf, columns=cols, show="headings", height=8)
        for c in cols:
            tree.heading(c, text=heads[c])
            tree.column(c, width=200 if c == "artikel" else 80, anchor="w")
        tree.pack(side="left", fill="both", expand=True)
        sb = tk.Scrollbar(tf, orient="vertical", command=tree.yview)
        sb.pack(side="right", fill="y")
        tree.configure(yscrollcommand=sb.set)
        _make_treeview_sortable(tree)
        gesamt = 0.0
        for pzn, art, menge, rab, apu, ch, vf, ba in positions:
            tree.insert("", "end", values=(pzn, art, menge, f"{rab:.0f}", ch, vf, ba))
            if ba == "Bestellung" and apu is not None:
                gesamt += apu * menge * (1 - rab / 100.0)
        tk.Label(win, text=f"Gesamt (netto): {_eur(gesamt)}", font=("Arial", 11, "bold"),
                 fg=ACCENT, bg=BG).pack(padx=18, anchor="e")

        def stornieren():
            if status == "storniert":
                return
            if not messagebox.askyesno("Stornieren", f"Verkauf #{bid} wirklich stornieren?\n"
                                       "Der Lagerbestand wird zurückgebucht.", parent=win):
                return
            jetzt = datetime.now().isoformat(timespec="seconds")
            with self._conn() as con:
                con.execute("UPDATE tbl_bestellungen SET status='storniert', geaendert_am=? WHERE id=?",
                            (jetzt, bid))
                # Bestand zurueckbuchen fuer gelieferte Positionen mit Charge.
                for pzn, art, menge, rab, apu, ch, vf, ba in positions:
                    if ba == "Bestellung" and (ch or vf):
                        upd = con.execute(
                            "UPDATE tbl_lagerbestand SET menge=menge+?, aktualisiert_am=? "
                            "WHERE pzn=? AND COALESCE(charge,'')=? AND COALESCE(verfall,'')=?",
                            (menge, jetzt, pzn, ch, vf))
                        if upd.rowcount == 0:
                            con.execute(
                                "INSERT INTO tbl_lagerbestand(pzn,artikelname,charge,verfall,menge,aktualisiert_am) "
                                "VALUES(?,?,?,?,?,?)", (pzn, art, ch, vf, menge, jetzt))
                con.commit()
            self._log("Verkauf storniert", bid, apotheke, "Bestand zurückgebucht")
            self._refresh_verkaeufe()
            self._refresh_lager()
            win.destroy()
            messagebox.showinfo("Storniert", f"Verkauf #{bid} storniert, Bestand zurückgebucht.",
                                parent=top)

        def msk_toggle():
            self._msk_markieren(not bool(msk), bids=[bid])
            win.destroy()

        btns = tk.Frame(win, bg=BG)
        btns.pack(fill="x", padx=18, pady=(4, 14))
        if status != "storniert":
            tk.Button(btns, text="Stornieren", command=stornieren, bg="#a32d2d", fg="white",
                      font=("Arial", 10, "bold"), padx=12, pady=4).pack(side="left")
            if msk:
                tk.Button(btns, text="↺ MSK-Markierung aufheben", command=msk_toggle,
                          padx=12, pady=4).pack(side="left", padx=(8, 0))
            else:
                tk.Button(btns, text="✓ In MSK erfasst", command=msk_toggle, bg="#11823b",
                          fg="white", font=("Arial", 10, "bold"), padx=12, pady=4).pack(side="left", padx=(8, 0))
        # Auftragsbestaetigung jederzeit erneut erzeugen.
        tk.Button(btns, text="🖨 Auftragsbestätigung", command=lambda: self._auftrag_drucken(bid, win),
                  bg=ACCENT, fg="white", font=("Arial", 10, "bold"), padx=12, pady=4).pack(side="left", padx=(8, 0))
        tk.Button(btns, text="📧 Per E-Mail", command=lambda: self._auftrag_mail(bid, win),
                  bg="#3867b7", fg="white", font=("Arial", 10, "bold"), padx=12, pady=4).pack(side="left", padx=(6, 0))
        tk.Button(btns, text="Schließen", command=win.destroy, padx=12, pady=4).pack(side="right")
        self._restyle_buttons(win)
        win.lift()
        win.focus_force()

    # ================================================================ Artikel
    def _build_artikel(self, parent):
        bar = tk.Frame(parent, bg=BG)
        bar.pack(fill="x", padx=8, pady=(10, 4))
        tk.Label(bar, text="Suche (PZN / Artikel):", bg=BG, font=("Arial", 9, "bold")).pack(side="left")
        self._art_filter_var = tk.StringVar()
        e = tk.Entry(bar, textvariable=self._art_filter_var, width=28)
        e.pack(side="left", padx=(6, 12))
        e.bind("<KeyRelease>", lambda _e: self._refresh_artikel())
        self._art_nur_bestand = tk.BooleanVar(value=False)
        tk.Radiobutton(bar, text="Alle", variable=self._art_nur_bestand, value=False, bg=BG,
                       font=("Arial", 9), command=self._refresh_artikel).pack(side="left")
        tk.Radiobutton(bar, text="Nur mit Bestand", variable=self._art_nur_bestand, value=True, bg=BG,
                       font=("Arial", 9), command=self._refresh_artikel).pack(side="left", padx=(4, 0))
        self.art_info = tk.Label(bar, text="", bg=BG, fg=ACCENT, font=("Arial", 9, "bold"))
        self.art_info.pack(side="right")

        tk.Label(parent, text="Doppelklick auf einen Artikel: alle Chargen und Verfälle anzeigen.",
                 bg=BG, fg="#666", font=("Arial", 9)).pack(anchor="w", padx=8, pady=(0, 4))
        tf = tk.Frame(parent, bg=BG)
        tf.pack(fill="both", expand=True, padx=8, pady=(0, 8))
        cols = ("pzn", "artikel", "df", "pck", "apu", "bestand")
        heads = {"pzn": "PZN", "artikel": "Artikel", "df": "DF", "pck": "PCK",
                 "apu": "APU €", "bestand": "Bestand"}
        tree = ttk.Treeview(tf, columns=cols, show="headings")
        for c in cols:
            tree.heading(c, text=heads[c])
            tree.column(c, width=250 if c == "artikel" else 80, anchor="w")
        tree.pack(side="left", fill="both", expand=True)
        sb = tk.Scrollbar(tf, orient="vertical", command=tree.yview)
        sb.pack(side="right", fill="y")
        tree.configure(yscrollcommand=sb.set)
        _make_treeview_sortable(tree)
        tree.bind("<Double-1>", self._artikel_chargen_view)
        self.art_tree = tree

    def _refresh_artikel(self):
        if not hasattr(self, "art_tree"):
            return
        self.art_tree.delete(*self.art_tree.get_children())
        like = f"%{self._art_filter_var.get().strip()}%"
        nur_bestand = self._art_nur_bestand.get()
        hat_stamm = self._table_exists_now("tbl_artikelstamm")
        with self._conn() as con:
            if hat_stamm:
                sql = ("SELECT n.pzn, n.artikelname, COALESCE(a.df,''), COALESCE(a.pck,''), n.apu, "
                       "(SELECT COALESCE(SUM(menge),0) FROM tbl_lagerbestand l WHERE l.pzn=n.pzn) "
                       "FROM tbl_nmg_stamm n LEFT JOIN tbl_artikelstamm a ON a.pzn=n.pzn "
                       "WHERE (n.pzn LIKE ? OR n.artikelname LIKE ?) ORDER BY n.artikelname")
            else:
                sql = ("SELECT n.pzn, n.artikelname, '', '', n.apu, "
                       "(SELECT COALESCE(SUM(menge),0) FROM tbl_lagerbestand l WHERE l.pzn=n.pzn) "
                       "FROM tbl_nmg_stamm n "
                       "WHERE (n.pzn LIKE ? OR n.artikelname LIKE ?) ORDER BY n.artikelname")
            rows = con.execute(sql, (like, like)).fetchall()
        gezeigt = 0
        for pzn, art, df, pck, apu, bestand in rows:
            if nur_bestand and (bestand or 0) <= 0:
                continue
            self.art_tree.insert("", "end", values=(pzn, art, df, pck, _eur(apu), bestand))
            gezeigt += 1
        self.art_info.config(text=f"{gezeigt} Artikel")

    def _artikel_chargen_view(self, _e=None):
        sel = self.art_tree.selection()
        if not sel:
            return
        vals = self.art_tree.item(sel[0], "values")
        pzn, art = vals[0], vals[1]
        top = self.winfo_toplevel()
        win = tk.Toplevel(top)
        win.title("Chargen / Verfall")
        win.configure(bg=BG)
        win.transient(top)
        tk.Label(win, text=art or pzn, font=("Arial", 13, "bold"), fg=ACCENT,
                 bg=BG).pack(padx=18, pady=(14, 2), anchor="w")
        tk.Label(win, text=f"PZN {pzn}", font=("Arial", 9), fg="#555", bg=BG).pack(padx=18, anchor="w")
        tf = tk.Frame(win, bg=BG)
        tf.pack(fill="both", expand=True, padx=18, pady=(8, 8))
        tree = ttk.Treeview(tf, columns=("charge", "verfall", "bestand"), show="headings", height=8)
        for c, t, w in (("charge", "Charge", 160), ("verfall", "Verfall", 130), ("bestand", "Bestand", 90)):
            tree.heading(c, text=t)
            tree.column(c, width=w, anchor="w")
        tree.pack(side="left", fill="both", expand=True)
        sb = tk.Scrollbar(tf, orient="vertical", command=tree.yview)
        sb.pack(side="right", fill="y")
        tree.configure(yscrollcommand=sb.set)
        _make_treeview_sortable(tree)
        chargen = self._lager_chargen(pzn)
        if chargen:
            for ch, vf, m in chargen:
                tree.insert("", "end", values=(ch or "—", vf or "—", m))
        else:
            tree.insert("", "end", values=("(kein Bestand)", "", ""))
        tk.Button(win, text="Schließen", command=win.destroy, padx=14, pady=4).pack(pady=(0, 14))
        self._restyle_buttons(win)
        win.lift()
        win.focus_force()

    # =============================================================== Protokoll
    def _build_protokoll(self, parent):
        bar = tk.Frame(parent, bg=BG)
        bar.pack(fill="x", padx=8, pady=(10, 4))
        tk.Label(bar, text="Suche (Mitarbeiter / Aktion / Kunde / Auftrag-Nr / Detail):",
                 bg=BG, font=("Arial", 9, "bold")).pack(side="left")
        self._log_filter_var = tk.StringVar()
        e = tk.Entry(bar, textvariable=self._log_filter_var, width=24)
        e.pack(side="left", padx=(6, 8))
        e.bind("<KeyRelease>", lambda _e: self._refresh_log())
        tk.Label(bar, text="Aktion:", bg=BG, font=("Arial", 9, "bold")).pack(side="left")
        self._log_aktion_var = tk.StringVar(value="Alle")
        self._log_aktion_cmb = ttk.Combobox(bar, textvariable=self._log_aktion_var, width=22,
                                            state="readonly", values=("Alle",))
        self._log_aktion_cmb.pack(side="left", padx=(4, 0))
        self._log_aktion_var.trace_add("write", lambda *_: self._refresh_log())
        tk.Button(bar, text="🔄 Aktualisieren", command=self._refresh_log,
                  font=("Arial", 9), padx=8, pady=2).pack(side="right")
        self.log_info = tk.Label(bar, text="", bg=BG, fg=ACCENT, font=("Arial", 9, "bold"))
        self.log_info.pack(side="right", padx=(0, 12))
        tk.Label(parent, text="Wer hat was wann geändert (neueste oben). Spalten sortierbar – "
                              "z. B. nach Mitarbeiter oder Auftrag, um Fehler/Änderungen zu finden.",
                 bg=BG, fg="#666", font=("Arial", 9)).pack(anchor="w", padx=8, pady=(0, 4))

        tf = tk.Frame(parent, bg=BG)
        tf.pack(fill="both", expand=True, padx=8, pady=(0, 8))
        cols = ("zeit", "bearbeiter", "aktion", "auftrag", "kunde", "details")
        heads = {"zeit": "Zeitpunkt", "bearbeiter": "Mitarbeiter", "aktion": "Aktion",
                 "auftrag": "Auftrag", "kunde": "Kunde", "details": "Details"}
        tree = ttk.Treeview(tf, columns=cols, show="headings")
        for c in cols:
            tree.heading(c, text=heads[c])
            w = 140 if c == "zeit" else (250 if c == "details" else (
                110 if c in ("bearbeiter", "aktion", "kunde") else 70))
            tree.column(c, width=w, anchor="w")
        tree.pack(side="left", fill="both", expand=True)
        sb = tk.Scrollbar(tf, orient="vertical", command=tree.yview)
        sb.pack(side="right", fill="y")
        tree.configure(yscrollcommand=sb.set)
        _make_treeview_sortable(tree)
        self.log_tree = tree

    def _refresh_log(self):
        if not hasattr(self, "log_tree"):
            return
        self.log_tree.delete(*self.log_tree.get_children())
        filt = self._log_filter_var.get().strip().lower()
        aktion_f = self._log_aktion_var.get() if hasattr(self, "_log_aktion_var") else "Alle"
        with self._conn() as con:
            if not _table_exists(con, "tbl_kasse_log"):
                self.log_info.config(text="0 Einträge")
                return
            alle_aktionen = [r[0] for r in con.execute(
                "SELECT DISTINCT COALESCE(aktion,'') FROM tbl_kasse_log WHERE aktion<>'' ORDER BY aktion")]
            rows = con.execute(
                "SELECT zeitpunkt, COALESCE(bearbeiter,''), COALESCE(aktion,''), bestell_id, "
                "COALESCE(kunde,''), COALESCE(details,'') FROM tbl_kasse_log "
                "ORDER BY id DESC LIMIT 2000").fetchall()
        # Aktions-Dropdown aktuell halten.
        werte = ("Alle",) + tuple(alle_aktionen)
        if self._log_aktion_cmb.cget("values") != werte:
            self._log_aktion_cmb.config(values=werte)
        gezeigt = 0
        for z, bearb, aktion, bid, kunde, details in rows:
            auftrag = f"#{bid}" if bid else ""
            if aktion_f != "Alle" and aktion != aktion_f:
                continue
            if filt and filt not in " ".join(str(x).lower() for x in
                                             (bearb, aktion, auftrag, kunde, details)):
                continue
            zt = str(z).replace("T", " ")[:16]
            self.log_tree.insert("", "end", values=(zt, bearb, aktion, auftrag, kunde, details))
            gezeigt += 1
        suffix = f" · Aktion: {aktion_f}" if aktion_f != "Alle" else ""
        self.log_info.config(text=f"{gezeigt} Einträge{suffix}")

    # ============================================================ Wareneingang
    def _build_wareneingang(self, parent):
        rbar = tk.Frame(parent, bg=BG)
        rbar.pack(fill="x", padx=8, pady=(10, 0))
        tk.Button(rbar, text="🕓  Vorbestellungs-Report (offene Vorbestellungen / Kunden anrufen)",
                  command=lambda: self._show_view("vorbestellungen"), bg="#d8e2ee", fg=ACCENT,
                  relief="flat", font=("Arial", 9, "bold"), padx=10, pady=4,
                  cursor="hand2").pack(side="left")

        form = tk.LabelFrame(parent, text="Artikel ins Lager buchen (nur NMG)", bg=BG,
                             fg=ACCENT, font=("Arial", 10, "bold"), padx=12, pady=8)
        form.pack(fill="x", padx=8, pady=(8, 6))

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
        _make_treeview_sortable(tree)
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
        verfall = _normalize_verfall(self.we_verfall_var.get())
        if verfall is None:
            messagebox.showwarning("Wareneingang", "Verfalldatum ungültig. Bitte als MM/JJ oder "
                                   "MM/JJJJ eingeben (Monat 1–12), z. B. 01/2027.", parent=top)
            return
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
        self._log("Wareneingang", details=f"{artikelname} (PZN {pzn}) +{menge} · "
                  f"Charge {charge or '–'} · Verf {verfall or '–'}")

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
        if neu == int(menge):
            return
        if not messagebox.askyesno(
                "Bestand korrigieren",
                f"Bestand für PZN {pzn} · Charge {charge or '–'} · Verf {verfall or '–'}\n"
                f"wirklich von {menge} auf {neu} ändern?"
                + ("\n\n(0 entfernt die Lagerzeile.)" if neu == 0 else ""), parent=top):
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
        self._log("Bestandskorrektur", details=f"PZN {pzn} · Charge {charge or '–'} · "
                  f"{menge} -> {neu}")
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

    def _import_historie(self, als_vorbestellung):
        top = self.winfo_toplevel()
        path = self._pick_file()
        if not path:
            return
        try:
            r = kasse_import.import_verkaeufe(self.db_path, path, als_vorbestellung=als_vorbestellung)
        except Exception as e:
            messagebox.showerror("Import", self._import_fehlertext(path, e), parent=top)
            return
        if hasattr(self, "vh_tree"):
            self._refresh_verkaeufe()
        if hasattr(self, "vb_tree"):
            self._refresh_vorbestellungen()
        was = "Vorbestellungen" if als_vorbestellung else "Verkäufe"
        self._log(f"{was} importiert",
                  details=f"{r['bestellungen']} {was}, {r['positionen']} Positionen aus {r['quelle']}")
        unklar = (f"\n⚠ {r['datum_unklar']} Datum/Daten nicht erkannt – auf heute gesetzt."
                  if r.get("datum_unklar") else "")
        messagebox.showinfo(
            f"{was}-Import",
            f"Quelle: {r['quelle']} · {r['gelesen']} Zeilen gelesen.\n"
            f"Erkannte Zusatzspalten: {', '.join(r['spalten']) or '–'}\n\n"
            f"Angelegt: {r['bestellungen']} {was} mit {r['positionen']} Position(en) · "
            f"Übersprungen: {r['uebersprungen']}{unklar}", parent=top)

    def _import_verkaeufe_datei(self):
        self._import_historie(False)

    def _import_vorbestellungen_datei(self):
        self._import_historie(True)

    @staticmethod
    def _import_fehlertext(path, e):
        msg = str(e)
        if path.lower().endswith(".pdf"):
            return ("PDF wird derzeit nicht unterstützt (keine PDF-Bibliothek installiert). "
                    "Bitte als Excel, CSV oder TXT speichern.\n\nDetail: " + msg)
        return "Import fehlgeschlagen:\n" + msg
