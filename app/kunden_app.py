#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""NMG Kunden-App (eigenstaendig) - CRM fuer Apotheken-Kunden.

Eigenes Fenster / Taskleisten-Icon (run_standalone, AUMID NMG.Kunden), teilt
sich die Datenbank mit NMGone (app/config.py · tbl_kunden_center). Die KASSE
greift auf denselben Kundenstamm zu -> Schema/Spalten bleiben unveraendert,
diese App liest und schreibt nur die bekannten Felder.

Drei Bereiche (Sidebar):
  * Kundenliste  - editierbar: Suche, Ampel (letzte Analyse), ABC-Klasse,
                   Neu/Bearbeiten/Loeschen. Detail-Dialog mit Reitern
                   (Stammdaten, Artikel-Rabatte, Analysen, Verkaufte Artikel,
                   Notizen) + E-Mail-Versand. 1:1 aus dem NMGone-Kunden-Center.
  * ABC-Analyse  - Pareto ueber alle Bedarfsanalysen, je Apotheke nach
                   Rabatt-Potenzial (read-only, aus der Testoberflaeche).
  * Landkarte    - Offline-Deutschlandkarte (app/geo_de.py), Kunden als
                   ABC-farbige Punkte, Klick -> Kontakt-Steckbrief.

Start:  python start_kunden.py     (bzw. NMGone.exe --kunden / Cockpit-Kachel)
"""
from __future__ import annotations

import calendar
import getpass
import os
import re
import sys
import sqlite3
import webbrowser
from datetime import date, datetime
from pathlib import Path
from urllib.parse import quote_plus

import tkinter as tk
from tkinter import ttk, messagebox, simpledialog

from .config import DB_PATH, ASSETS_DIR, fenstertitel
from . import theme
from . import geo_de

# ── Palette (aus dem zentralen Theme) ─────────────────────────────────────────
PRIMARY = theme.PRIMARY
PRIMARY_DARK = theme.PRIMARY_DARK
ACCENT = theme.ACCENT
SUCCESS = theme.SUCCESS
WARNING = theme.WARNING
DANGER = theme.DANGER
BG = theme.BG
CARD = theme.CARD
CARD_ALT = theme.CARD_ALT
INK = theme.INK
MUTED = theme.MUTED
BORDER = theme.BORDER
SIDEBAR = theme.SIDEBAR
SIDEBAR_ACTIVE = theme.SIDEBAR_ACTIVE
SIDEBAR_TEXT = theme.SIDEBAR_TEXT
SIDEBAR_MUTED = theme.SIDEBAR_MUTED
SELECT_BG = theme.SELECT_BG
FONT = theme.FONT

ABC_COLORS = {"A": SUCCESS, "B": WARNING, "C": DANGER, None: "#8A97A5"}


# ══════════════════════════════════════════════════════════════════════════════
#  Datenzugriff
# ══════════════════════════════════════════════════════════════════════════════
def _con():
    """Schreib-/Lese-Verbindung (Kundenstamm wird hier gepflegt)."""
    return sqlite3.connect(DB_PATH)


def _db_ro():
    """Read-only-Verbindung fuer Auswertungen/Bestellungen."""
    try:
        con = sqlite3.connect(f"file:{DB_PATH}?mode=ro", uri=True)
    except Exception:
        con = sqlite3.connect(DB_PATH)
    con.row_factory = sqlite3.Row
    return con


def _table_exists(con, name: str) -> bool:
    return con.execute(
        "SELECT 1 FROM sqlite_master WHERE type='table' AND name=?", (name,)
    ).fetchone() is not None


def ensure_tables():
    """Legt tbl_kunden_center an (falls fremde Erst-DB) und ergaenzt die
    Kunden-App-Felder. Idempotent - identisch zur NMGone-Migration, damit App
    und Hauptprogramm denselben Stamm sehen."""
    with sqlite3.connect(DB_PATH) as con:
        con.execute("""CREATE TABLE IF NOT EXISTS tbl_kunden_center (
                id INTEGER PRIMARY KEY AUTOINCREMENT, kundennummer TEXT,
                kundenname TEXT, kundentyp TEXT, ansprechpartner TEXT,
                telefon TEXT, email TEXT, status TEXT DEFAULT 'aktiv',
                notizen TEXT, erstellt_am TEXT DEFAULT CURRENT_TIMESTAMP,
                geaendert_am TEXT, bearbeiter TEXT)""")
        existing = {r[1] for r in con.execute("PRAGMA table_info(tbl_kunden_center)").fetchall()}
        for col, typedef in [
            ("plz", "TEXT"), ("ort", "TEXT"), ("strasse", "TEXT"),
            ("inhaber", "TEXT"), ("ansprechpartner2", "TEXT"),
            ("kundentyp", "TEXT"), ("ansprechpartner", "TEXT"),
            ("msk_kundennummer", "TEXT"), ("hausnummer", "TEXT"),
            ("inhaber_titel", "TEXT"), ("inhaber_anrede", "TEXT"),
            ("inhaber_vorname", "TEXT"), ("inhaber_zuname", "TEXT"),
            ("besteller_name", "TEXT"), ("besteller_durchwahl", "TEXT"),
            ("besteller_email", "TEXT"), ("rechnungsemail", "TEXT"),
            ("rechnungsart", "TEXT"), ("quartalsverguetung", "TEXT"),
        ]:
            if col not in existing:
                con.execute(f"ALTER TABLE tbl_kunden_center ADD COLUMN {col} {typedef}")
        con.execute("UPDATE tbl_kunden_center SET kundentyp='ZW' WHERE kundentyp='ZF'")
        con.commit()


def _dq_label(dq) -> str:
    """Anzeige-Label fuer eine gespeicherte datenquelle (intern bleibt PK='NMG')."""
    val = (dq or "NMG")
    if val == "NMG":
        return "PK"
    if val in ("ZW", "ZF"):
        return "ZW"
    return str(val)


def _eur(v) -> str:
    if v in (None, ""):
        return "–"
    try:
        return f"{float(v):,.2f} €".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return str(v)


# ── Namens-Abgleich (Kundenstamm <-> Bedarfsanalysen) ─────────────────────────
def _norm_name(s) -> str:
    t = str(s or "").lower()
    t = t.replace("ä", "ae").replace("ö", "oe").replace("ü", "ue").replace("ß", "ss")
    return re.sub(r"[^a-z0-9]+", " ", t).strip()


_STOP = {"apotheke", "apo", "die", "der", "das", "am", "an", "zur", "zum", "und",
         "monate", "bench", "bench2", "hr", "herr", "frau", "12", "gmbh"}


def _name_tokens(s) -> set[str]:
    return {w for w in _norm_name(s).split() if len(w) >= 4 and w not in _STOP}


def abc_ranking() -> list[dict]:
    """ABC-Analyse aller Apotheken aus den Bedarfsanalysen, nach Rabatt-Potenzial.

    Mehrfach-Analysen derselben Apotheke -> juengste zaehlt. Pareto: A bis 80 %,
    B bis 95 %, C Rest des kumulierten Potenzials.
    """
    con = _db_ro()
    try:
        if not _table_exists(con, "tbl_auswertungen"):
            return []
        rows = con.execute(
            "SELECT a.id, a.apotheke, a.datum, a.gesamt_absatz, a.nmg_treffer, "
            "a.anzahl_positionen, COALESCE(SUM(p.nmg_rabatt_gesamt),0) AS pot "
            "FROM tbl_auswertungen a "
            "LEFT JOIN tbl_auswertungspositionen p ON p.auswertung_id=a.id "
            "GROUP BY a.id ORDER BY a.id DESC").fetchall()
    finally:
        con.close()
    best: dict[str, dict] = {}
    for r in rows:
        key = _norm_name(r["apotheke"])
        if not key or key in best:
            continue
        best[key] = {"apotheke": r["apotheke"], "datum": r["datum"],
                     "absatz": r["gesamt_absatz"] or 0, "treffer": r["nmg_treffer"] or 0,
                     "positionen": r["anzahl_positionen"] or 0, "pot": r["pot"] or 0.0}
    items = sorted(best.values(), key=lambda d: d["pot"], reverse=True)
    total = sum(d["pot"] for d in items) or 1.0
    kum = 0.0
    for d in items:
        kum += d["pot"]
        anteil = kum / total
        d["klasse"] = "A" if anteil <= 0.80 else ("B" if anteil <= 0.95 else "C")
        d["anteil_kum"] = anteil
    return items


def abc_map() -> dict[str, dict]:
    """Token-Index der ABC-Apotheken fuer das Zuordnen zu Stammkunden."""
    return {a["apotheke"]: a for a in abc_ranking()}


def _abc_fuer(kundenname, ranking) -> dict | None:
    ktok = _name_tokens(kundenname)
    if not ktok:
        return None
    for a in ranking:
        if ktok & _name_tokens(a["apotheke"]):
            return a
    return None


def kunden_master() -> list[dict]:
    """Stammkunden aus tbl_kunden_center inkl. zugeordneter ABC-Klasse (Fuzzy)."""
    con = _db_ro()
    try:
        if not _table_exists(con, "tbl_kunden_center"):
            return []
        rows = con.execute("SELECT * FROM tbl_kunden_center ORDER BY kundenname").fetchall()
        kunden = [dict(r) for r in rows]
    finally:
        con.close()
    ranking = abc_ranking()
    for k in kunden:
        treffer = _abc_fuer(k.get("kundenname"), ranking)
        if treffer:
            k["abc"] = treffer["klasse"]
            k["abc_pot"] = treffer["pot"]
            k["abc_quelle"] = treffer["apotheke"]
        else:
            k["abc"] = None
            k["abc_pot"] = None
        k["latlon"] = geo_de.plz_to_latlon(k.get("plz"))
    return kunden


def _zeitraum_cutoff(monate: int | None) -> str | None:
    if not monate:
        return None
    t = date.today()
    m = t.month - monate
    y = t.year
    while m <= 0:
        m += 12
        y -= 1
    d = min(t.day, calendar.monthrange(y, m)[1])
    return f"{y:04d}-{m:02d}-{d:02d}"


def _bestell_basis(con):
    return _table_exists(con, "tbl_bestellungen") and _table_exists(con, "tbl_bestellpositionen")


def kunde_umsatz(kundennummer: str, monate: int | None = None) -> dict:
    out = {"umsatz": 0.0, "bestellungen": 0, "positionen": 0, "von": None, "bis": None}
    if not kundennummer:
        return out
    con = _db_ro()
    try:
        if not _bestell_basis(con):
            return out
        cut = _zeitraum_cutoff(monate)
        where = "b.kundennummer=? AND COALESCE(b.status,'')<>'storniert'"
        args = [str(kundennummer)]
        if cut:
            where += " AND b.datum>=?"
            args.append(cut)
        row = con.execute(
            f"SELECT COALESCE(SUM(p.apu*p.menge*(1-COALESCE(p.rabatt_prozent,0))),0) umsatz, "
            f"COUNT(DISTINCT b.id) best, COUNT(p.id) pos, MIN(b.datum) von, MAX(b.datum) bis "
            f"FROM tbl_bestellungen b JOIN tbl_bestellpositionen p ON p.bestell_id=b.id "
            f"WHERE {where}", args).fetchone()
        out.update(umsatz=row["umsatz"] or 0.0, bestellungen=row["best"] or 0,
                   positionen=row["pos"] or 0, von=row["von"], bis=row["bis"])
    finally:
        con.close()
    return out


def kunde_top_artikel(kundennummer: str, monate: int | None = None,
                      limit: int | None = None) -> list[dict]:
    if not kundennummer:
        return []
    con = _db_ro()
    try:
        if not _bestell_basis(con):
            return []
        cut = _zeitraum_cutoff(monate)
        where = "b.kundennummer=? AND COALESCE(b.status,'')<>'storniert'"
        args: list = [str(kundennummer)]
        if cut:
            where += " AND b.datum>=?"
            args.append(cut)
        sql = (
            "SELECT p.pzn, MAX(p.artikelname) name, SUM(p.menge) menge, "
            "SUM(p.apu*p.menge*(1-COALESCE(p.rabatt_prozent,0))) umsatz, "
            "COUNT(DISTINCT b.id) bestellungen "
            "FROM tbl_bestellungen b JOIN tbl_bestellpositionen p ON p.bestell_id=b.id "
            f"WHERE {where} GROUP BY p.pzn ORDER BY umsatz DESC")
        if limit:
            sql += " LIMIT ?"
            args.append(limit)
        return [dict(r) for r in con.execute(sql, args).fetchall()]
    finally:
        con.close()


def kunde_vorbestellungen(kundennummer: str) -> list[dict]:
    if not kundennummer:
        return []
    con = _db_ro()
    try:
        if not _bestell_basis(con):
            return []
        return [dict(r) for r in con.execute(
            "SELECT b.datum, b.status, p.pzn, p.artikelname, p.menge, p.apu, "
            "p.liefertermin, p.lieferzeit "
            "FROM tbl_bestellungen b JOIN tbl_bestellpositionen p ON p.bestell_id=b.id "
            "WHERE b.kundennummer=? AND p.bestellart='Vorbestellung' "
            "AND COALESCE(b.status,'')<>'storniert' "
            "ORDER BY b.datum DESC", (str(kundennummer),)).fetchall()]
    finally:
        con.close()


def kunden_ampel(kundennummer, kundenname):
    """Ampel: gruen < 6 Monate, gelb 6-9, rot > 9 Monate oder nie (letzte Analyse)."""
    try:
        con = _db_ro()
        try:
            row = con.execute(
                "SELECT MAX(datum) FROM tbl_auswertungen "
                "WHERE (kundennummer=? AND kundennummer<>'') OR (kundenname=? AND kundenname<>'')",
                (kundennummer or "", kundenname or "")).fetchone()
        finally:
            con.close()
        letzte = row[0] if row and row[0] else None
        if not letzte:
            return "🔴", "–"
        dt = datetime.strptime(str(letzte)[:10], "%Y-%m-%d")
        tage = (datetime.now() - dt).days
        datum_str = dt.strftime("%d.%m.%Y")
        if tage < 180:
            return "🟢", datum_str
        if tage < 270:
            return "🟡", datum_str
        return "🔴", datum_str
    except Exception:
        return "⚪", "–"


# ══════════════════════════════════════════════════════════════════════════════
#  UI-Bausteine
# ══════════════════════════════════════════════════════════════════════════════
class Card(tk.Frame):
    def __init__(self, master, padding=18, **kw):
        super().__init__(master, bg=CARD, highlightbackground=BORDER,
                         highlightthickness=1, bd=0, **kw)
        self.inner = tk.Frame(self, bg=CARD)
        self.inner.pack(fill="both", expand=True, padx=padding, pady=padding)


class PillButton(tk.Label):
    def __init__(self, master, text, command, kind="primary", **kw):
        colors = {
            "primary": (PRIMARY, PRIMARY_DARK, "#FFFFFF"),
            "accent": (ACCENT, "#1A6FA6", "#FFFFFF"),
            "success": (SUCCESS, "#0D6630", "#FFFFFF"),
            "ghost": (CARD, "#EEF3F8", PRIMARY),
        }
        self._base, self._hover, fg = colors.get(kind, colors["primary"])
        super().__init__(master, text=text, bg=self._base, fg=fg,
                         font=(FONT, 10, "bold"), padx=18, pady=9, cursor="hand2", **kw)
        self._command = command
        if kind == "ghost":
            self.config(highlightbackground=BORDER, highlightthickness=1)
        self.bind("<Button-1>", lambda e: self._command and self._command())
        self.bind("<Enter>", lambda e: self.config(bg=self._hover))
        self.bind("<Leave>", lambda e: self.config(bg=self._base))


# ══════════════════════════════════════════════════════════════════════════════
#  Hauptpanel
# ══════════════════════════════════════════════════════════════════════════════
class KundenPanel(tk.Frame):
    def __init__(self, master, on_close=None):
        super().__init__(master, bg=BG)
        self.on_close = on_close
        self.bearbeiter = getpass.getuser()
        self._active = "liste"
        ensure_tables()
        self._build_sidebar()
        self._content = tk.Frame(self, bg=BG)
        self._content.pack(side="left", fill="both", expand=True)
        self.show_page("liste")

    # ----- Layout -----------------------------------------------------------
    def _build_sidebar(self):
        self.sidebar = theme.Sidebar(self, width=240, title="Kunden",
                                     subtitle="CRM · Apotheken-Kunden")
        self.sidebar.pack(side="left", fill="y")
        for key, icon, label in [
            ("liste", "📋", "Kundenliste"),
            ("abc", "🏆", "ABC-Analyse"),
            ("karte", "🗺️", "Landkarte"),
        ]:
            self.sidebar.add_item(key, icon, label, lambda k=key: self.show_page(k))
        self.sidebar.add_footer_note("Gemeinsamer Kundenstamm\nmit NMGone & Kasse.")

    def show_page(self, key):
        self._active = key
        self.sidebar.set_active(key)
        for child in self._content.winfo_children():
            child.destroy()
        {"liste": self._page_liste, "abc": self._page_abc,
         "karte": self._page_karte}.get(key, self._page_liste)()

    def _page_header(self, parent, title, subtitle):
        head = tk.Frame(parent, bg=BG)
        head.pack(fill="x", padx=30, pady=(24, 8))
        tk.Label(head, text=title, bg=BG, fg=INK, font=(FONT, 18, "bold")).pack(anchor="w")
        tk.Label(head, text=subtitle, bg=BG, fg=MUTED, font=(FONT, 11)).pack(anchor="w", pady=(2, 0))

    # ══════════════════════════════════════════════════════════════════════
    #  Seite: Kundenliste (editierbar)
    # ══════════════════════════════════════════════════════════════════════
    def _page_liste(self):
        page = tk.Frame(self._content, bg=BG)
        page.pack(fill="both", expand=True)
        self._page_header(page, "Kundenliste",
                          "Apotheken verwalten · Analysen einsehen · per E-Mail versenden.")

        body = tk.Frame(page, bg=BG)
        body.pack(fill="both", expand=True, padx=30, pady=(0, 20))

        toolbar = tk.Frame(body, bg=BG)
        toolbar.pack(fill="x", pady=(0, 8))
        tk.Label(toolbar, text="Suche:", bg=BG, fg=PRIMARY,
                 font=(FONT, 10, "bold")).pack(side="left")
        self._liste_search = tk.StringVar()
        ent = tk.Entry(toolbar, textvariable=self._liste_search, width=28,
                       relief="flat", highlightbackground=BORDER, highlightthickness=1)
        ent.pack(side="left", padx=(6, 12), ipady=4)

        PillButton(toolbar, "＋  Neu", self._kunde_neu, kind="primary").pack(side="left", padx=(0, 6))
        PillButton(toolbar, "Öffnen / Bearbeiten", self._kunde_open, kind="ghost").pack(side="left", padx=(0, 6))
        PillButton(toolbar, "Löschen", self._kunde_del, kind="ghost").pack(side="left", padx=(0, 6))
        PillButton(toolbar, "↻", lambda: self._liste_reload(), kind="ghost").pack(side="left")

        tree_frame = tk.Frame(body, bg=CARD, highlightbackground=BORDER, highlightthickness=1)
        tree_frame.pack(fill="both", expand=True)
        cols = ("ampel", "abc", "kundennummer", "kundenname", "plz", "ort",
                "status", "letzte_analyse", "email")
        heads = {"ampel": "🚦", "abc": "ABC", "kundennummer": "Kundennummer",
                 "kundenname": "Apothekenname", "plz": "PLZ", "ort": "Ort",
                 "status": "Status", "letzte_analyse": "Letzte Analyse", "email": "E-Mail"}
        widths = {"ampel": 34, "abc": 40, "kundennummer": 110, "kundenname": 220,
                  "plz": 60, "ort": 120, "status": 80, "letzte_analyse": 100, "email": 170}
        self._liste_tree = ttk.Treeview(tree_frame, columns=cols, show="headings",
                                        selectmode="browse")
        for c in cols:
            self._liste_tree.heading(c, text=heads[c])
            self._liste_tree.column(c, width=widths[c],
                                    anchor="center" if c in ("ampel", "abc") else "w",
                                    stretch=(c == "kundenname"))
        self._liste_tree.pack(side="left", fill="both", expand=True)
        sb = ttk.Scrollbar(tree_frame, orient="vertical", command=self._liste_tree.yview)
        sb.pack(side="right", fill="y")
        self._liste_tree.configure(yscrollcommand=sb.set)
        self._liste_tree.bind("<Double-1>", lambda e: self._kunde_open())

        self._liste_info = tk.Label(body, text="", bg=BG, fg=MUTED, font=(FONT, 9))
        self._liste_info.pack(anchor="w", pady=(6, 0))

        self._liste_rows: dict[str, dict] = {}
        self._liste_search.trace_add("write", lambda *_: self._liste_reload())
        self._liste_reload()

    def _liste_reload(self):
        tree = getattr(self, "_liste_tree", None)
        if tree is None:
            return
        flt = self._liste_search.get().strip().lower()
        for it in tree.get_children():
            tree.delete(it)
        self._liste_rows.clear()
        try:
            kunden = kunden_master()
        except Exception as exc:
            self._liste_info.config(text=f"Kunden konnten nicht geladen werden: {exc}")
            return
        gezeigt = 0
        for k in kunden:
            knr = str(k.get("kundennummer") or "")
            kname = str(k.get("kundenname") or "")
            amp, letzte = kunden_ampel(knr, kname)
            vals = {"ampel": amp, "abc": k.get("abc") or "–", "kundennummer": knr,
                    "kundenname": kname, "plz": str(k.get("plz") or ""),
                    "ort": str(k.get("ort") or ""), "status": str(k.get("status") or ""),
                    "letzte_analyse": letzte, "email": str(k.get("email") or "")}
            row_vals = tuple(vals[c] for c in tree["columns"])
            if flt and flt not in " ".join(str(v) for v in row_vals).lower():
                continue
            iid = tree.insert("", "end", values=row_vals)
            self._liste_rows[iid] = k
            gezeigt += 1
        self._liste_info.config(text=f"{gezeigt} von {len(kunden)} Kunden angezeigt")

    def _selected_kunde(self):
        sel = self._liste_tree.selection()
        return self._liste_rows.get(sel[0]) if sel else None

    def _kunde_neu(self):
        open_kunden_dialog(self, None, self.bearbeiter, on_saved=self._liste_reload)

    def _kunde_open(self):
        k = self._selected_kunde()
        if not k:
            messagebox.showinfo("Kunden", "Bitte zuerst einen Kunden auswählen.", parent=self)
            return
        open_kunden_dialog(self, k, self.bearbeiter, on_saved=self._liste_reload)

    def _kunde_del(self):
        k = self._selected_kunde()
        if not k:
            messagebox.showinfo("Kunden", "Bitte zuerst einen Kunden auswählen.", parent=self)
            return
        if not messagebox.askyesno("Kunden",
                                   f"Kunde '{k.get('kundenname','')}' wirklich löschen?", parent=self):
            return
        try:
            with _con() as con:
                con.execute("DELETE FROM tbl_kunden_center WHERE id=?", (k["id"],))
                con.commit()
        except Exception as exc:
            messagebox.showerror("Kunden", f"Löschen fehlgeschlagen:\n{exc}", parent=self)
        self._liste_reload()

    # ══════════════════════════════════════════════════════════════════════
    #  Seite: ABC-Analyse
    # ══════════════════════════════════════════════════════════════════════
    def _page_abc(self):
        page = tk.Frame(self._content, bg=BG)
        page.pack(fill="both", expand=True)
        self._page_header(page, "ABC-Analyse",
                          "Apotheken nach Rabatt-Potenzial aus den Bedarfsanalysen einstufen.")
        body = tk.Frame(page, bg=BG)
        body.pack(fill="both", expand=True, padx=30, pady=(0, 20))

        try:
            ranking = abc_ranking()
        except Exception as exc:
            tk.Label(body, text=f"ABC-Analyse fehlgeschlagen:\n{exc}",
                     bg=BG, fg=DANGER, font=(FONT, 12)).pack(anchor="w", pady=20)
            return
        if not ranking:
            tk.Label(body, text="Keine Bedarfsanalysen für eine ABC-Analyse vorhanden.",
                     bg=BG, fg=MUTED, font=(FONT, 12)).pack(anchor="w", pady=20)
            return

        summe = sum(a["pot"] for a in ranking)
        zaehler = {"A": 0, "B": 0, "C": 0}
        for a in ranking:
            zaehler[a["klasse"]] += 1
        tk.Label(body, text="A = die wertvollen 80 %, B = nächste 15 %, C = restliche 5 % "
                            "des Gesamtpotenzials (Pareto).",
                 bg=BG, fg=MUTED, font=(FONT, 11), justify="left", wraplength=860).pack(anchor="w", pady=(0, 10))
        sumrow = tk.Frame(body, bg=BG)
        sumrow.pack(fill="x", pady=(0, 12))
        for kl in ("A", "B", "C"):
            t = tk.Frame(sumrow, bg=CARD, highlightbackground=BORDER, highlightthickness=1)
            t.pack(side="left", padx=(0 if kl == "A" else 10, 0))
            tk.Label(t, text=f"  Klasse {kl}  ", bg=ABC_COLORS[kl], fg="#FFFFFF",
                     font=(FONT, 11, "bold")).pack(anchor="w")
            tk.Label(t, text=f"{zaehler[kl]} Apotheken", bg=CARD, fg=INK,
                     font=(FONT, 13, "bold")).pack(anchor="w", padx=14, pady=(6, 10))
        tk.Label(body, text=f"Gesamtpotenzial: {_eur(summe)}", bg=BG, fg=PRIMARY,
                 font=(FONT, 12, "bold")).pack(anchor="w", pady=(0, 10))

        head = tk.Frame(body, bg=BG)
        head.pack(fill="x", padx=(0, 8))
        for txt, w in [("ABC", 5), ("Apotheke", 32), ("Absatz", 9), ("NMG", 6),
                       ("Rabatt-Potenzial", 16), ("kum. %", 8)]:
            tk.Label(head, text=txt, bg=BG, fg=MUTED, font=(FONT, 10, "bold"),
                     width=w, anchor="w").pack(side="left")
        canvas = tk.Canvas(body, bg=BG, highlightthickness=0)
        scroll = ttk.Scrollbar(body, orient="vertical", command=canvas.yview)
        inner = tk.Frame(canvas, bg=BG)
        win = canvas.create_window((0, 0), window=inner, anchor="nw")
        inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind("<Configure>", lambda e: canvas.itemconfig(win, width=e.width))
        canvas.configure(yscrollcommand=scroll.set)
        canvas.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")
        canvas.bind_all("<MouseWheel>", lambda e: canvas.yview_scroll(int(-e.delta / 120), "units"))
        for i, a in enumerate(ranking):
            bg = CARD if i % 2 == 0 else CARD_ALT
            row = tk.Frame(inner, bg=bg, highlightbackground=BORDER, highlightthickness=1)
            row.pack(fill="x", pady=1, padx=(0, 8))
            tk.Label(row, text=f" {a['klasse']} ", bg=ABC_COLORS[a["klasse"]], fg="#FFFFFF",
                     font=(FONT, 10, "bold"), width=3).pack(side="left", padx=(6, 6), pady=6)
            tk.Label(row, text=str(a["apotheke"])[:34], bg=bg, fg=INK, font=(FONT, 11, "bold"),
                     width=30, anchor="w").pack(side="left")
            tk.Label(row, text=f"{a['absatz']:.0f}", bg=bg, fg=INK, font=(FONT, 10),
                     width=9, anchor="w").pack(side="left")
            tk.Label(row, text=str(a["treffer"]), bg=bg, fg=SUCCESS, font=(FONT, 10),
                     width=6, anchor="w").pack(side="left")
            tk.Label(row, text=_eur(a["pot"]), bg=bg, fg=PRIMARY, font=(FONT, 10, "bold"),
                     width=16, anchor="w").pack(side="left")
            tk.Label(row, text=f"{a['anteil_kum']*100:.0f}%", bg=bg, fg=MUTED, font=(FONT, 10),
                     width=8, anchor="w").pack(side="left")

    # ══════════════════════════════════════════════════════════════════════
    #  Seite: Landkarte
    # ══════════════════════════════════════════════════════════════════════
    def _page_karte(self):
        page = tk.Frame(self._content, bg=BG)
        page.pack(fill="both", expand=True)
        self._page_header(page, "Landkarte",
                          "Kunden nach Region · Punkt anklicken für die Kontaktdaten.")
        body = tk.Frame(page, bg=BG)
        body.pack(fill="both", expand=True, padx=30, pady=(0, 20))

        try:
            kunden = kunden_master()
        except Exception as exc:
            tk.Label(body, text=f"Kunden konnten nicht geladen werden:\n{exc}",
                     bg=BG, fg=DANGER, font=(FONT, 12)).pack(anchor="w", pady=20)
            return
        platzierbar = [k for k in kunden if k.get("latlon")]
        mapcol = tk.Frame(body, bg=BG)
        mapcol.pack(side="left", fill="both", expand=True)
        self._karte_detailcol = tk.Frame(body, bg=BG, width=360)
        self._karte_detailcol.pack(side="left", fill="both", padx=(16, 0))
        self._karte_detailcol.pack_propagate(False)
        info = "Punkt anklicken für Kontaktdaten." if platzierbar else "Keine Kunden mit erkennbarer PLZ."
        tk.Label(self._karte_detailcol, text=info, bg=BG, fg=MUTED,
                 font=(FONT, 11), wraplength=320, justify="left").pack(anchor="w", pady=18)

        W, H = 560, 700
        cv = tk.Canvas(mapcol, bg="#EAF1F8", width=W, height=H, highlightthickness=1,
                       highlightbackground=BORDER)
        cv.pack(anchor="nw")
        pts = []
        for lat, lon in geo_de.GERMANY_OUTLINE:
            x, y = geo_de.project(lat, lon, W, H)
            pts += [x, y]
        cv.create_polygon(pts, fill="#FBFCFE", outline="#9DB6CE", width=2)
        for i, kl in enumerate(("A", "B", "C")):
            cv.create_oval(16, 16 + i * 22, 28, 28 + i * 22, fill=ABC_COLORS[kl], outline="")
            cv.create_text(36, 22 + i * 22, text=f"Klasse {kl}", anchor="w",
                           fill=INK, font=(FONT, 9))
        belegt = []
        for k in platzierbar:
            lat, lon = k["latlon"]
            x, y = geo_de.project(lat, lon, W, H)
            col = ABC_COLORS.get(k.get("abc"))
            r = 7
            dot = cv.create_oval(x - r, y - r, x + r, y + r, fill=col, outline="#FFFFFF", width=2)
            ly = y - 12
            for (bx, by) in belegt:
                if abs(bx - x) < 70 and abs(by - ly) < 14:
                    ly += 16
            belegt.append((x, ly))
            if x > W * 0.66:
                lx, anchor = x - 10, "e"
            else:
                lx, anchor = x + 10, "w"
            label = cv.create_text(lx, ly, text=k.get("kundenname") or "", anchor=anchor,
                                   fill=INK, font=(FONT, 9, "bold"))
            for item in (dot, label):
                cv.tag_bind(item, "<Button-1>", lambda e, kk=k: self._karte_show(kk))
                cv.tag_bind(item, "<Enter>", lambda e: cv.config(cursor="hand2"))
                cv.tag_bind(item, "<Leave>", lambda e: cv.config(cursor=""))

    def _karte_show(self, kunde):
        col = self._karte_detailcol
        for c in col.winfo_children():
            c.destroy()
        self._render_steckbrief(col, kunde)

    # ----- Kontakt-Steckbrief (read-only, fuer die Karte) -------------------
    def _render_steckbrief(self, parent, k):
        card = Card(parent, padding=16)
        card.pack(fill="x")
        klasse = k.get("abc")
        top = tk.Frame(card.inner, bg=CARD)
        top.pack(fill="x")
        tk.Label(top, text=f" {klasse or '–'} ", bg=ABC_COLORS.get(klasse), fg="#FFFFFF",
                 font=(FONT, 12, "bold")).pack(side="left", padx=(0, 10))
        tk.Label(top, text=k.get("kundenname") or "—", bg=CARD, fg=INK,
                 font=(FONT, 14, "bold"), wraplength=280, justify="left").pack(side="left")
        if k.get("kundennummer"):
            tk.Label(card.inner, text=f"Kundennr. {k['kundennummer']}"
                     + (f"  ·  {k.get('kundentyp')}" if k.get("kundentyp") else ""),
                     bg=CARD, fg=MUTED, font=(FONT, 10)).pack(anchor="w", pady=(4, 0))

        def field(label, value, action=None, action_label=None):
            if not value:
                return
            r = tk.Frame(card.inner, bg=CARD)
            r.pack(fill="x", pady=3)
            tk.Label(r, text=label, bg=CARD, fg=MUTED, font=(FONT, 9), anchor="w").pack(anchor="w")
            line = tk.Frame(r, bg=CARD)
            line.pack(fill="x")
            tk.Label(line, text=str(value), bg=CARD, fg=INK, font=(FONT, 11),
                     anchor="w", justify="left", wraplength=240).pack(side="left")
            if action and action_label:
                PillButton(line, action_label, action, kind="ghost").pack(side="right")

        adresse = " ".join(str(x) for x in [k.get("strasse"), k.get("hausnummer")] if x).strip()
        ort_zeile = " ".join(str(x) for x in [k.get("plz"), k.get("ort")] if x).strip()
        voll_adr = (adresse + ", " + ort_zeile).strip(", ")
        inhaber = (k.get("inhaber") or " ".join(
            str(x) for x in [k.get("inhaber_vorname"), k.get("inhaber_zuname")] if x).strip())

        tk.Frame(card.inner, bg=BORDER, height=1).pack(fill="x", pady=(10, 8))
        field("Inhaber / Ansprechpartner", inhaber or k.get("ansprechpartner"))
        field("Telefon", k.get("telefon"),
              action=(lambda v=k.get("telefon"): self._copy_text(v, "Telefonnummer")) if k.get("telefon") else None,
              action_label="Kopieren" if k.get("telefon") else None)
        field("E-Mail", k.get("email"),
              action=(lambda v=k.get("email"): _mailto(v)) if k.get("email") else None,
              action_label="Schreiben" if k.get("email") else None)
        field("Adresse", voll_adr or None,
              action=(lambda v=voll_adr: _open_maps(v)) if voll_adr else None,
              action_label="Karte" if voll_adr else None)
        if k.get("abc_pot"):
            field("Rabatt-Potenzial (aus Analyse)", _eur(k["abc_pot"]))

        knr = k.get("kundennummer")
        if knr:
            try:
                u = kunde_umsatz(knr)
            except Exception:
                u = None
            tk.Frame(card.inner, bg=BORDER, height=1).pack(fill="x", pady=(10, 8))
            if u and u["bestellungen"]:
                ums = tk.Frame(card.inner, bg=SELECT_BG)
                ums.pack(fill="x")
                tk.Label(ums, text=f"Umsatz gesamt: {_eur(u['umsatz'])}", bg=SELECT_BG,
                         fg=PRIMARY, font=(FONT, 12, "bold")).pack(anchor="w", padx=12, pady=(8, 0))
                tk.Label(ums, text=f"{u['bestellungen']} Bestellungen · {u['von']} bis {u['bis']}",
                         bg=SELECT_BG, fg=MUTED, font=(FONT, 10)).pack(anchor="w", padx=12, pady=(0, 8))
                PillButton(card.inner, "📊  Umsatz & Artikel ansehen",
                           lambda kk=k: open_umsatz_dialog(self, kk), kind="accent").pack(anchor="w", pady=(8, 0))
            else:
                tk.Label(card.inner, text="Noch keine Bestellungen erfasst.", bg=CARD,
                         fg=MUTED, font=(FONT, 10)).pack(anchor="w")
        PillButton(card.inner, "✏  Kunde bearbeiten",
                   lambda kk=k: open_kunden_dialog(self, kk, self.bearbeiter,
                                                   on_saved=self._liste_reload),
                   kind="ghost").pack(anchor="w", pady=(10, 0))

    def _copy_text(self, value, was="Text"):
        try:
            self.clipboard_clear()
            self.clipboard_append(str(value))
            messagebox.showinfo("Kopiert", f"{was} in die Zwischenablage kopiert.", parent=self)
        except Exception:
            pass


# ── Kontakt-Aktionen (modul-global) ───────────────────────────────────────────
def _mailto(email):
    try:
        webbrowser.open(f"mailto:{email}")
    except Exception:
        pass


def _open_maps(adresse):
    try:
        webbrowser.open(f"https://www.google.com/maps/search/?api=1&query={quote_plus(adresse)}")
    except Exception:
        pass


# ══════════════════════════════════════════════════════════════════════════════
#  Umsatz-Detailfenster
# ══════════════════════════════════════════════════════════════════════════════
def open_umsatz_dialog(parent, kunde):
    """Eigenes Fenster: Umsatz mit Zeitraeumen, Top-Artikel, Vorbestellungen."""
    knr = kunde.get("kundennummer")
    win = tk.Toplevel(parent)
    win.title(f"Umsatz – {kunde.get('kundenname') or knr}")
    win.geometry("720x720")
    win.configure(bg=BG)
    win.transient(parent.winfo_toplevel())

    head = tk.Frame(win, bg=PRIMARY)
    head.pack(fill="x")
    tk.Label(head, text=kunde.get("kundenname") or "—", bg=PRIMARY, fg="#FFFFFF",
             font=(FONT, 17, "bold")).pack(anchor="w", padx=20, pady=(14, 2))
    tk.Label(head, text=f"Kundennr. {knr or '—'} · {kunde.get('ort') or ''}", bg=PRIMARY,
             fg="#CFE2F3", font=(FONT, 10)).pack(anchor="w", padx=20, pady=(0, 14))

    state = {"zeitraum": 0, "show_alle": False}
    bar = tk.Frame(win, bg=BG)
    bar.pack(fill="x", padx=20, pady=(12, 4))
    tk.Label(bar, text="Zeitraum:", bg=BG, fg=MUTED, font=(FONT, 10)).pack(side="left", padx=(0, 8))
    zr_buttons = {}
    body = tk.Frame(win, bg=BG)
    body.pack(fill="both", expand=True, padx=20, pady=(4, 16))

    def render():
        for c in body.winfo_children():
            c.destroy()
        mz = state["zeitraum"] or None
        u = kunde_umsatz(knr, mz)
        grid = tk.Frame(body, bg=BG)
        grid.pack(fill="x", pady=(0, 12))
        for i, (lab, val, col) in enumerate([
                ("Umsatz", _eur(u["umsatz"]), PRIMARY),
                ("Bestellungen", str(u["bestellungen"]), ACCENT),
                ("Artikelpositionen", str(u["positionen"]), SUCCESS)]):
            t = tk.Frame(grid, bg=CARD, highlightbackground=BORDER, highlightthickness=1)
            t.grid(row=0, column=i, sticky="nsew", padx=(0 if i == 0 else 10, 0))
            grid.columnconfigure(i, weight=1)
            tk.Label(t, text=lab, bg=CARD, fg=MUTED, font=(FONT, 10)).pack(anchor="w", padx=14, pady=(12, 0))
            tk.Label(t, text=val, bg=CARD, fg=col, font=(FONT, 19, "bold")).pack(anchor="w", padx=14, pady=(2, 12))

        alle = kunde_top_artikel(knr, mz, limit=None)
        anzeige = alle if state["show_alle"] else alle[:10]
        titel = tk.Frame(body, bg=BG)
        titel.pack(fill="x")
        tk.Label(titel, text=("Alle Artikel" if state["show_alle"] else "Top 10 Artikel")
                 + f"  ({len(alle)} insgesamt)", bg=BG, fg=INK, font=(FONT, 13, "bold")).pack(side="left")
        if len(alle) > 10:
            def _toggle():
                state["show_alle"] = not state["show_alle"]
                render()
            PillButton(titel, "Top 10 anzeigen" if state["show_alle"] else f"Alle {len(alle)} anzeigen",
                       _toggle, kind="ghost").pack(side="right")

        if not alle:
            tk.Label(body, text="Keine Artikel im Zeitraum.", bg=BG, fg=MUTED,
                     font=(FONT, 11)).pack(anchor="w", pady=12)
        else:
            hdr = tk.Frame(body, bg=BG)
            hdr.pack(fill="x", pady=(8, 2))
            for txt, w in [("PZN", 10), ("Artikel", 34), ("Menge", 7), ("Best.", 6), ("Umsatz", 13)]:
                tk.Label(hdr, text=txt, bg=BG, fg=MUTED, font=(FONT, 9, "bold"),
                         width=w, anchor="w").pack(side="left")
            cv = tk.Canvas(body, bg=BG, highlightthickness=0, height=240)
            sc = ttk.Scrollbar(body, orient="vertical", command=cv.yview)
            inner = tk.Frame(cv, bg=BG)
            w_ = cv.create_window((0, 0), window=inner, anchor="nw")
            inner.bind("<Configure>", lambda e: cv.configure(scrollregion=cv.bbox("all")))
            cv.bind("<Configure>", lambda e: cv.itemconfig(w_, width=e.width))
            cv.configure(yscrollcommand=sc.set)
            cv.pack(side="top", fill="both", expand=True)
            sc.pack(side="right", fill="y")
            for i, a in enumerate(anzeige):
                bg = CARD if i % 2 == 0 else CARD_ALT
                row = tk.Frame(inner, bg=bg)
                row.pack(fill="x", pady=1)
                tk.Label(row, text=a["pzn"], bg=bg, fg=INK, font=(FONT, 9), width=10, anchor="w").pack(side="left")
                tk.Label(row, text=str(a["name"] or "")[:36], bg=bg, fg=INK, font=(FONT, 10),
                         width=34, anchor="w").pack(side="left")
                tk.Label(row, text=f"{a['menge']:.0f}", bg=bg, fg=INK, font=(FONT, 9), width=7, anchor="w").pack(side="left")
                tk.Label(row, text=str(a["bestellungen"]), bg=bg, fg=MUTED, font=(FONT, 9), width=6, anchor="w").pack(side="left")
                tk.Label(row, text=_eur(a["umsatz"]), bg=bg, fg=PRIMARY, font=(FONT, 9, "bold"),
                         width=13, anchor="w").pack(side="left")

        vb = kunde_vorbestellungen(knr)
        vcard = tk.Frame(body, bg=BG)
        vcard.pack(fill="x", pady=(14, 0))
        tk.Label(vcard, text=f"🕓  Offene Vorbestellungen ({len(vb)})", bg=BG, fg=WARNING,
                 font=(FONT, 13, "bold")).pack(anchor="w", pady=(0, 4))
        if not vb:
            tk.Label(vcard, text="Keine offenen Vorbestellungen.", bg=BG, fg=MUTED,
                     font=(FONT, 10)).pack(anchor="w")
        else:
            for v in vb[:8]:
                txt = f"   {v.get('datum') or ''}  ·  {str(v.get('artikelname') or '')[:40]}  ·  {v.get('menge') or 0} St"
                if v.get("liefertermin"):
                    txt += f"  ·  Liefertermin {v['liefertermin']}"
                tk.Label(vcard, text=txt, bg=BG, fg=INK, font=(FONT, 10)).pack(anchor="w")

    def set_zeitraum(monate):
        state["zeitraum"] = monate
        for m, b in zr_buttons.items():
            active = (m == monate)
            b.config(bg=PRIMARY if active else CARD, fg="#FFFFFF" if active else PRIMARY)
        render()

    for label, monate in [("Alles", 0), ("12 Monate", 12), ("6 Monate", 6), ("3 Monate", 3)]:
        b = tk.Label(bar, text=label, bg=CARD, fg=PRIMARY, font=(FONT, 10, "bold"),
                     padx=12, pady=6, cursor="hand2", highlightbackground=BORDER, highlightthickness=1)
        b.pack(side="left", padx=(0, 6))
        b.bind("<Button-1>", lambda e, m=monate: set_zeitraum(m))
        zr_buttons[monate] = b
    set_zeitraum(0)


# ══════════════════════════════════════════════════════════════════════════════
#  Detail-/Bearbeiten-Dialog (Stammdaten, Rabatte, Analysen, Verkaufte, Notizen)
# ══════════════════════════════════════════════════════════════════════════════
def open_kunden_dialog(parent, kunden_row=None, bearbeiter=None, on_saved=None):
    """Kompaktes Kunden-Formular mit Reitern. 1:1 aus dem NMGone-Kunden-Center,
    arbeitet auf demselben tbl_kunden_center."""
    ensure_tables()
    bearbeiter = bearbeiter or getpass.getuser()
    is_new = kunden_row is None
    initial = dict(kunden_row) if kunden_row else {}
    title = "Neuer Kunde" if is_new else f"Kunde: {initial.get('kundenname','') or initial.get('kundennummer','')}"
    top = parent.winfo_toplevel()

    win = tk.Toplevel(parent)
    win.title(title)
    win.configure(bg=BG)
    win.transient(top)
    win.resizable(True, True)
    win.minsize(720, 500)
    w, h = 820, 600
    try:
        top.update_idletasks()
        x = top.winfo_rootx() + max(0, (top.winfo_width() - w) // 2)
        y = top.winfo_rooty() + max(0, (top.winfo_height() - h) // 2)
        win.geometry(f"{w}x{h}+{x}+{y}")
    except Exception:
        win.geometry(f"{w}x{h}")
    win.grab_set()

    win.columnconfigure(0, weight=1)
    win.rowconfigure(2, weight=1)

    vars_ = {}

    # ── Kopf: Typ-Auswahl (PK/ZW) + Kundennummer + MSK ──────────────────────
    head = tk.Frame(win, bg=BG)
    head.grid(row=0, column=0, sticky="ew", padx=18, pady=(14, 8))

    typ_var = tk.StringVar(value=str(initial.get("kundentyp", "") or ""))
    vars_["kundentyp"] = typ_var
    tk.Label(head, text="Typ:", bg=BG, fg=PRIMARY, font=(FONT, 11, "bold")).pack(side="left")
    typ_btns = {}

    def _set_typ(val):
        typ_var.set(val)
        for v, b in typ_btns.items():
            on = (v == val)
            b.configure(bg=(PRIMARY if on else "#e8eef5"), fg=("white" if on else "#11304d"))
        _apply_typ_defaults(val)

    for val in ("PK", "ZW"):
        b = tk.Button(head, text=val, width=6, relief="flat", cursor="hand2",
                      bg="#e8eef5", fg="#11304d", font=(FONT, 11, "bold"),
                      activebackground="#d8e2ee", command=lambda v=val: _set_typ(v))
        b.pack(side="left", padx=(8, 0))
        typ_btns[val] = b

    tk.Label(head, text="Kundennummer *", bg=BG, fg="#c00",
             font=(FONT, 11, "bold")).pack(side="left", padx=(22, 6))
    knr_var = tk.StringVar(value=str(initial.get("kundennummer", "") or ""))
    vars_["kundennummer"] = knr_var
    tk.Entry(head, textvariable=knr_var, width=16, font=(FONT, 11)).pack(side="left")

    tk.Label(head, text="MSK:", bg=BG, fg=PRIMARY, font=(FONT, 11, "bold")).pack(side="left", padx=(16, 6))
    msk_var = tk.StringVar()
    vars_["msk_kundennummer"] = msk_var
    tk.Entry(head, textvariable=msk_var, width=16, state="readonly",
             readonlybackground="#eef2f8", fg="#555").pack(side="left")

    def _sync_msk(*_a):
        k = knr_var.get().strip()
        msk_var.set(("216" + k) if k else "")
    knr_var.trace_add("write", _sync_msk)
    _sync_msk()

    tk.Frame(win, bg="#d8e2ee", height=1).grid(row=1, column=0, sticky="ew", padx=18)

    nb = ttk.Notebook(win)
    nb.grid(row=2, column=0, sticky="nsew", padx=18, pady=(8, 6))

    def _mk_section(parent, rowref, text, top=10):
        tk.Label(parent, text=text, bg="#ffffff", fg=PRIMARY,
                 font=(FONT, 11, "bold")).grid(row=rowref[0], column=0, columnspan=2,
                                               sticky="w", pady=(top, 2))
        rowref[0] += 1

    def _mk_field(parent, rowref, key, label, required=False):
        fg = "#c00" if required else "#11304d"
        tk.Label(parent, text=label, bg="#ffffff", fg=fg,
                 font=(FONT, 9, "bold")).grid(row=rowref[0], column=0, sticky="w", pady=3, padx=(0, 8))
        var = tk.StringVar(value=str(initial.get(key, "") or ""))
        vars_[key] = var
        tk.Entry(parent, textvariable=var).grid(row=rowref[0], column=1, sticky="ew", pady=3)
        rowref[0] += 1
        return var

    def _mk_combo(parent, rowref, key, label, values, default=""):
        tk.Label(parent, text=label, bg="#ffffff", fg="#11304d",
                 font=(FONT, 9, "bold")).grid(row=rowref[0], column=0, sticky="w", pady=3, padx=(0, 8))
        var = tk.StringVar(value=(str(initial.get(key, "") or "") or default))
        vars_[key] = var
        ttk.Combobox(parent, textvariable=var, values=values, state="readonly").grid(
            row=rowref[0], column=1, sticky="ew", pady=3)
        rowref[0] += 1
        return var

    # ---- Reiter: Stammdaten ----
    tab_stamm = tk.Frame(nb, bg="#ffffff")
    nb.add(tab_stamm, text="  Stammdaten  ")
    tab_stamm.columnconfigure(0, weight=1, uniform="cols")
    tab_stamm.columnconfigure(1, weight=1, uniform="cols")

    colL = tk.Frame(tab_stamm, bg="#ffffff")
    colL.grid(row=0, column=0, sticky="nsew", padx=(10, 12), pady=8)
    colL.columnconfigure(1, weight=1)
    colR = tk.Frame(tab_stamm, bg="#ffffff")
    colR.grid(row=0, column=1, sticky="nsew", padx=(12, 10), pady=8)
    colR.columnconfigure(1, weight=1)

    rL = [0]
    _mk_section(colL, rL, "Adresse", top=0)
    _mk_field(colL, rL, "kundenname", "Apothekenname *", required=True)
    _mk_field(colL, rL, "strasse", "Straße")
    _mk_field(colL, rL, "hausnummer", "Hausnummer")
    tk.Label(colL, text="PLZ *", bg="#ffffff", fg="#c00",
             font=(FONT, 9, "bold")).grid(row=rL[0], column=0, sticky="w", pady=3, padx=(0, 8))
    _plz_box = tk.Frame(colL, bg="#ffffff")
    _plz_box.grid(row=rL[0], column=1, sticky="ew", pady=3)
    plz_var = tk.StringVar(value=str(initial.get("plz", "") or ""))
    vars_["plz"] = plz_var
    tk.Entry(_plz_box, textvariable=plz_var, width=10).pack(side="left")
    _plz_status = tk.Label(_plz_box, text="", bg="#ffffff", font=(FONT, 8))
    _plz_status.pack(side="left", padx=(6, 0))
    rL[0] += 1
    _mk_field(colL, rL, "ort", "Ort")

    _mk_section(colL, rL, "Inhaber")
    _mk_field(colL, rL, "inhaber_titel", "Titel")
    _mk_combo(colL, rL, "inhaber_anrede", "Anrede", ["", "Frau", "Herr", "Divers"])
    _mk_field(colL, rL, "inhaber_vorname", "Vorname")
    _mk_field(colL, rL, "inhaber_zuname", "Zuname")

    def _check_plz(*_a):
        try:
            from .plz_lookup import is_valid_plz, lookup_ort
        except Exception:
            return
        p = plz_var.get().strip()
        if not p:
            _plz_status.config(text="", fg="#555")
            return
        if is_valid_plz(p):
            _plz_status.config(text="✓ gültig", fg="#127a2e")
            ort = lookup_ort(p)
            cur = vars_["ort"].get().strip()
            if ort and (not cur or cur == _check_plz.last):
                vars_["ort"].set(ort)
                _check_plz.last = ort
        else:
            _plz_status.config(text="✗ nicht gefunden", fg="#c00")
    _check_plz.last = ""
    try:
        from .plz_lookup import lookup_ort as _lo0
        _ip = str(initial.get("plz", "") or "").strip()
        if _ip and _lo0(_ip) and _lo0(_ip) == str(initial.get("ort", "") or "").strip():
            _check_plz.last = _lo0(_ip)
    except Exception:
        pass
    plz_var.trace_add("write", _check_plz)
    _check_plz()

    rR = [0]
    _mk_section(colR, rR, "Kontakt", top=0)
    _mk_field(colR, rR, "telefon", "Telefon")
    _mk_field(colR, rR, "email", "E-Mail")
    _mk_field(colR, rR, "rechnungsemail", "Rechnungs-E-Mail")

    _mk_section(colR, rR, "Verantw. Besteller (Rückfragen)")
    _mk_field(colR, rR, "besteller_name", "Name (leer = Inhaber)")
    _mk_field(colR, rR, "besteller_durchwahl", "Durchwahl")
    _mk_field(colR, rR, "besteller_email", "E-Mail")

    _mk_section(colR, rR, "Abrechnung")
    _mk_combo(colR, rR, "rechnungsart", "Rechnungsart",
              ["Sofortige Rechnung", "Monatlich", "Quartalsrechnung"],
              default="Sofortige Rechnung")
    _mk_combo(colR, rR, "status", "Status", ["aktiv", "inaktiv"], default="aktiv")

    def _truthy(v):
        return str(v or "").strip().lower() in ("ja", "yes", "1", "true", "x")
    quart_touched = [False]
    quart_bv = tk.BooleanVar(value=_truthy(initial.get("quartalsverguetung")))
    tk.Checkbutton(colR, text="Quartalsvergütung Partnerprogramm", variable=quart_bv,
                   bg="#ffffff", fg="#11304d", activebackground="#ffffff",
                   font=(FONT, 9, "bold"), anchor="w",
                   command=lambda: quart_touched.__setitem__(0, True)).grid(
        row=rR[0], column=0, columnspan=2, sticky="w", pady=(6, 3))
    rR[0] += 1

    def _apply_typ_defaults(val):
        if is_new and not quart_touched[0]:
            quart_bv.set(val == "PK")

    if str(initial.get("kundentyp", "") or "") in ("PK", "ZW"):
        _set_typ(initial["kundentyp"])

    # ---- Reiter: Artikel-Rabatte ----
    tab_rab = tk.Frame(nb, bg="#ffffff")
    nb.add(tab_rab, text="  Artikel-Rabatte  ")
    tab_rab.columnconfigure(0, weight=1)
    tab_rab.rowconfigure(0, weight=1)
    rab_tree = ttk.Treeview(tab_rab, columns=("pzn", "artikel", "rabatt"),
                            show="headings", selectmode="browse")
    for col, head_t, wdt in [("pzn", "PZN", 90), ("artikel", "Artikel", 260), ("rabatt", "Rabatt %", 80)]:
        rab_tree.heading(col, text=head_t)
        rab_tree.column(col, width=wdt, anchor=("e" if col == "rabatt" else "w"))
    rab_tree.grid(row=0, column=0, sticky="nsew", padx=(8, 0), pady=8)
    rab_sb = tk.Scrollbar(tab_rab, orient="vertical", command=rab_tree.yview)
    rab_sb.grid(row=0, column=1, sticky="ns", pady=8)
    rab_tree.configure(yscrollcommand=rab_sb.set)

    nmg_artikel = []
    try:
        with _db_ro() as con:
            if _table_exists(con, "tbl_nmg_stamm"):
                nmg_artikel = con.execute(
                    "SELECT pzn, artikelname FROM tbl_nmg_stamm ORDER BY artikelname").fetchall()
    except Exception:
        nmg_artikel = []
    nmg_artikel = [(str(p), a or "") for p, a in nmg_artikel]
    nmg_by_pzn = {p: a for p, a in nmg_artikel}

    def _rab_load():
        for it in rab_tree.get_children():
            rab_tree.delete(it)
        knr = knr_var.get().strip()
        if not knr:
            return
        try:
            with _db_ro() as con:
                if not _table_exists(con, "tbl_pk_konditionen"):
                    return
                for pzn, rab in con.execute(
                        "SELECT pzn, rabatt_prozent FROM tbl_pk_konditionen "
                        "WHERE kundennummer=? ORDER BY pzn", (knr,)).fetchall():
                    art = nmg_by_pzn.get(str(pzn), "")
                    rab_tree.insert("", "end", values=(pzn, art, f"{rab:g}" if rab is not None else ""))
        except Exception:
            pass
    _rab_load()

    def _rab_add():
        if not knr_var.get().strip():
            messagebox.showinfo(title, "Bitte zuerst die Kundennummer eingeben.", parent=win)
            return
        if not nmg_artikel:
            messagebox.showinfo(title, "Keine NMG-Artikel verfügbar.", parent=win)
            return
        dlg = tk.Toplevel(win)
        dlg.title("Artikel-Rabatt hinzufügen")
        dlg.configure(bg=BG)
        dlg.transient(win)
        dlg.grab_set()
        tk.Label(dlg, text="Artikel", bg=BG, fg=PRIMARY,
                 font=(FONT, 10, "bold")).grid(row=0, column=0, sticky="w", padx=12, pady=(12, 4))
        disp = [f"{a}  ({p})" for p, a in nmg_artikel]
        art_var = tk.StringVar()
        cb = ttk.Combobox(dlg, textvariable=art_var, values=disp, width=44, state="readonly")
        cb.grid(row=0, column=1, padx=12, pady=(12, 4))
        tk.Label(dlg, text="Rabatt %", bg=BG, fg=PRIMARY,
                 font=(FONT, 10, "bold")).grid(row=1, column=0, sticky="w", padx=12, pady=4)
        rab_var = tk.StringVar()
        tk.Entry(dlg, textvariable=rab_var, width=10).grid(row=1, column=1, sticky="w", padx=12, pady=4)

        def _ok():
            idx = cb.current()
            if idx < 0:
                messagebox.showinfo(title, "Bitte einen Artikel wählen.", parent=dlg)
                return
            pzn = str(nmg_artikel[idx][0])
            art = nmg_artikel[idx][1] or ""
            try:
                rab = float(rab_var.get().strip().replace(",", "."))
            except ValueError:
                messagebox.showinfo(title, "Rabatt muss eine Zahl sein.", parent=dlg)
                return
            for it in rab_tree.get_children():
                if str(rab_tree.set(it, "pzn")) == pzn:
                    rab_tree.set(it, "rabatt", f"{rab:g}")
                    dlg.destroy()
                    return
            rab_tree.insert("", "end", values=(pzn, art, f"{rab:g}"))
            dlg.destroy()
        tk.Button(dlg, text="Übernehmen", command=_ok, bg=PRIMARY, fg="white",
                  relief="flat", padx=14, pady=6).grid(row=2, column=1, sticky="e", padx=12, pady=12)

    def _rab_del():
        for it in rab_tree.selection():
            rab_tree.delete(it)

    rab_btns = tk.Frame(tab_rab, bg="#ffffff")
    rab_btns.grid(row=1, column=0, columnspan=2, sticky="ew", padx=8, pady=(0, 8))
    tk.Button(rab_btns, text="➕ Rabatt", command=_rab_add, bg="#3867b7", fg="white",
              relief="flat", padx=12, pady=6).pack(side="left")
    tk.Button(rab_btns, text="➖ Entfernen", command=_rab_del, padx=10, pady=6).pack(side="left", padx=(8, 0))

    # ---- Reiter: Analysen ----
    tab_ana = tk.Frame(nb, bg="#ffffff")
    nb.add(tab_ana, text="  Analysen  ")
    tab_ana.columnconfigure(0, weight=1)
    tab_ana.rowconfigure(0, weight=1)
    ana_tree = ttk.Treeview(tab_ana, columns=("datum", "typ", "apotheke", "treffer"),
                            show="headings", selectmode="browse")
    for col, head_t, wdt in [("datum", "Datum", 90), ("typ", "Typ", 50),
                             ("apotheke", "Name", 200), ("treffer", "Treffer", 60)]:
        ana_tree.heading(col, text=head_t)
        ana_tree.column(col, width=wdt, anchor="w")
    ana_tree.grid(row=0, column=0, sticky="nsew", padx=(8, 0), pady=8)
    ana_sb = tk.Scrollbar(tab_ana, orient="vertical", command=ana_tree.yview)
    ana_sb.grid(row=0, column=1, sticky="ns", pady=8)
    ana_tree.configure(yscrollcommand=ana_sb.set)

    ana_rows = {}

    def load_analysen():
        for item in ana_tree.get_children():
            ana_tree.delete(item)
        ana_rows.clear()
        knr = knr_var.get().strip()
        kname = vars_["kundenname"].get().strip()
        if not knr and not kname:
            return
        try:
            with _db_ro() as con:
                if not _table_exists(con, "tbl_auswertungen"):
                    return
                rows = con.execute(
                    "SELECT id, datum, COALESCE(datenquelle,'PK') as datenquelle, "
                    "apotheke, quelldatei, ausgabedatei, nmg_treffer, kundennummer, kundenname "
                    "FROM tbl_auswertungen "
                    "WHERE (kundennummer=? AND kundennummer<>'') "
                    "OR (kundenname=? AND kundenname<>'') "
                    "ORDER BY datetime(datum) DESC LIMIT 30", (knr, kname)).fetchall()
            for row in rows:
                dq = _dq_label(row["datenquelle"])
                datum = str(row["datum"] or "")[:10]
                iid = ana_tree.insert("", "end", values=(datum, dq, row["apotheke"] or "", row["nmg_treffer"] or 0))
                ana_rows[iid] = dict(row)
        except Exception:
            pass

    load_analysen()

    def send_email_analyse():
        sel = ana_tree.selection()
        if not sel:
            messagebox.showinfo(title, "Bitte zuerst eine Analyse auswählen.", parent=win)
            return
        row = ana_rows.get(sel[0], {})
        email_addr = vars_["email"].get().strip()
        if not email_addr:
            if messagebox.askyesno(title, "Keine E-Mail-Adresse hinterlegt.\nJetzt E-Mail-Adresse eingeben?", parent=win):
                new_mail = simpledialog.askstring(title, "E-Mail-Adresse eingeben:", parent=win)
                if new_mail:
                    vars_["email"].set(new_mail.strip())
                    email_addr = new_mail.strip()
                else:
                    return
            else:
                return
        ausgabe = row.get("ausgabedatei", "") or ""
        anhang = ""
        if ausgabe and Path(ausgabe).exists():
            anhang = ausgabe
        else:
            try:
                for f in (Path(__file__).resolve().parent.parent / "gespeicherte_analysen").rglob("*.xlsx"):
                    if row.get("apotheke") and str(row["apotheke"]).lower() in str(f).lower():
                        anhang = str(f)
                        break
            except Exception:
                pass
        try:
            if sys.platform.startswith("win"):
                import win32com.client
                outlook = win32com.client.Dispatch("Outlook.Application")
                mail = outlook.CreateItem(0)
                mail.To = email_addr
                mail.Subject = f"Analyse – {row.get('apotheke','')}"
                mail.Body = f"Anbei die Auswertung vom {str(row.get('datum',''))[:10]}."
                if anhang:
                    mail.Attachments.Add(anhang)
                mail.Display(True)
            else:
                messagebox.showinfo(title, "Outlook-Integration nur unter Windows verfügbar.", parent=win)
        except Exception as exc:
            messagebox.showerror(title, f"Outlook konnte nicht geöffnet werden:\n{exc}", parent=win)

    ana_btns = tk.Frame(tab_ana, bg="#ffffff")
    ana_btns.grid(row=1, column=0, columnspan=2, sticky="ew", padx=8, pady=(0, 8))
    tk.Button(ana_btns, text="📧 Analyse per E-Mail senden", command=send_email_analyse,
              bg="#3867b7", fg="white", relief="flat", padx=12, pady=6).pack(side="left")
    tk.Button(ana_btns, text="🔄 Aktualisieren", command=load_analysen, padx=10, pady=6).pack(side="left", padx=(8, 0))

    # ---- Reiter: Verkaufte Artikel ----
    tab_vk = tk.Frame(nb, bg="#ffffff")
    nb.add(tab_vk, text="  Verkaufte Artikel  ")
    tab_vk.columnconfigure(0, weight=1)
    tab_vk.rowconfigure(0, weight=1)
    vk_tree = ttk.Treeview(tab_vk, columns=("datum", "pzn", "artikel", "menge", "rabatt", "status"),
                           show="headings", selectmode="browse")
    for col, head_t, wdt, anc in [("datum", "Datum", 90, "w"), ("pzn", "PZN", 80, "w"),
                                  ("artikel", "Artikel", 210, "w"), ("menge", "Menge", 60, "e"),
                                  ("rabatt", "Rabatt %", 70, "e"), ("status", "Status", 90, "w")]:
        vk_tree.heading(col, text=head_t)
        vk_tree.column(col, width=wdt, anchor=anc)
    vk_tree.grid(row=0, column=0, sticky="nsew", padx=(8, 0), pady=8)
    vk_sb = tk.Scrollbar(tab_vk, orient="vertical", command=vk_tree.yview)
    vk_sb.grid(row=0, column=1, sticky="ns", pady=8)
    vk_tree.configure(yscrollcommand=vk_sb.set)
    vk_info = tk.Label(tab_vk, text="", bg="#ffffff", fg="#555", font=(FONT, 9))
    vk_info.grid(row=1, column=0, columnspan=2, sticky="w", padx=10, pady=(0, 6))

    def load_verkaufte():
        for it in vk_tree.get_children():
            vk_tree.delete(it)
        knr = knr_var.get().strip()
        if not knr:
            vk_info.config(text="Keine Kundennummer angegeben.")
            return
        try:
            with _db_ro() as con:
                if not _table_exists(con, "tbl_bestellpositionen"):
                    vk_info.config(text="Noch keine Verkaufsdaten vorhanden.")
                    return
                rows = con.execute(
                    "SELECT b.datum, p.pzn, p.artikelname, p.menge, p.rabatt_prozent, b.status "
                    "FROM tbl_bestellpositionen p "
                    "JOIN tbl_bestellungen b ON b.id = p.bestell_id "
                    "WHERE b.kundennummer = ? "
                    "ORDER BY datetime(b.datum) DESC, p.id DESC LIMIT 300", (knr,)).fetchall()
            gesamt = 0
            for datum, pzn, art, menge, rab, status in rows:
                vk_tree.insert("", "end", values=(
                    str(datum or "")[:10], pzn or "", art or "",
                    menge if menge is not None else "",
                    f"{rab:g}" if rab is not None else "", status or ""))
                gesamt += (menge or 0)
            vk_info.config(text=(f"{len(rows)} Positionen · Gesamtmenge {gesamt}"
                                 if rows else "Für diesen Kunden sind keine Verkäufe erfasst."))
        except Exception as exc:
            vk_info.config(text=f"Fehler beim Laden: {exc}")
    load_verkaufte()

    # ---- Reiter: Notizen ----
    tab_notiz = tk.Frame(nb, bg="#ffffff")
    nb.add(tab_notiz, text="  Notizen  ")
    tab_notiz.columnconfigure(0, weight=1)
    tab_notiz.rowconfigure(0, weight=1)
    notiz_txt = tk.Text(tab_notiz, wrap="word", relief="flat",
                        highlightbackground="#d8e2ee", highlightthickness=1)
    notiz_txt.insert("1.0", str(initial.get("notizen", "") or ""))
    notiz_txt.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
    notiz_sb = tk.Scrollbar(tab_notiz, orient="vertical", command=notiz_txt.yview)
    notiz_sb.grid(row=0, column=1, sticky="ns", pady=10)
    notiz_txt.configure(yscrollcommand=notiz_sb.set)

    # ── Button-Leiste ───────────────────────────────────────────────────────
    bar = tk.Frame(win, bg="#e8edf5", highlightbackground="#c5d3e8", highlightthickness=1)
    bar.grid(row=3, column=0, sticky="ew")

    def save_kunde():
        if typ_var.get() not in ("PK", "ZW"):
            messagebox.showinfo(title, "Bitte oben den Typ (PK oder ZW) wählen.", parent=win)
            return
        knr = knr_var.get().strip()
        kname = vars_["kundenname"].get().strip()
        plz = vars_["plz"].get().strip()
        if not knr:
            messagebox.showinfo(title, "Kundennummer ist ein Pflichtfeld.", parent=win)
            return
        if not kname:
            messagebox.showinfo(title, "Apothekenname ist ein Pflichtfeld.", parent=win)
            return
        if not plz:
            messagebox.showinfo(title, "PLZ ist ein Pflichtfeld.", parent=win)
            return
        try:
            from .plz_lookup import is_valid_plz
            if not is_valid_plz(plz) and not messagebox.askyesno(
                    title, f"Die PLZ {plz} wurde nicht gefunden.\nTrotzdem speichern?", parent=win):
                return
        except Exception:
            pass
        data = {k: v.get().strip() for k, v in vars_.items()}
        data["msk_kundennummer"] = ("216" + knr) if knr else ""
        _teile = [data.get("inhaber_titel", ""), data.get("inhaber_anrede", ""),
                  data.get("inhaber_vorname", ""), data.get("inhaber_zuname", "")]
        data["inhaber"] = " ".join(t for t in (x.strip() for x in _teile) if t)
        data["quartalsverguetung"] = "ja" if quart_bv.get() else "nein"
        data["notizen"] = notiz_txt.get("1.0", "end").strip()
        data["bearbeiter"] = bearbeiter

        cols = ["kundennummer", "msk_kundennummer", "kundenname", "kundentyp",
                "strasse", "hausnummer", "plz", "ort",
                "inhaber_titel", "inhaber_anrede", "inhaber_vorname", "inhaber_zuname", "inhaber",
                "telefon", "email", "rechnungsemail",
                "besteller_name", "besteller_durchwahl", "besteller_email",
                "rechnungsart", "quartalsverguetung", "status", "notizen", "bearbeiter"]
        try:
            with _con() as con:
                if is_new:
                    con.execute(
                        f"INSERT INTO tbl_kunden_center({','.join(cols)}) "
                        f"VALUES({','.join(':' + c for c in cols)})", data)
                else:
                    data["id"] = kunden_row["id"]
                    _sets = ",".join(f"{c}=:{c}" for c in cols)
                    con.execute(
                        f"UPDATE tbl_kunden_center SET {_sets},"
                        f"geaendert_am=CURRENT_TIMESTAMP WHERE id=:id", data)
                # Artikel-Rabatte: die Kunden-Maske ist alleiniger Editor je Kunde.
                con.execute("""CREATE TABLE IF NOT EXISTS tbl_pk_konditionen(
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    kundennummer TEXT, kundenname TEXT, pzn TEXT, rabatt_prozent REAL,
                    gueltigkeit TEXT, quelle TEXT, importdatum TEXT,
                    letzte_aktualisierung TEXT DEFAULT CURRENT_TIMESTAMP,
                    UNIQUE(kundennummer, pzn))""")
                con.execute("DELETE FROM tbl_pk_konditionen WHERE kundennummer=?", (knr,))
                _jetzt = datetime.now().isoformat(timespec="seconds")
                for it in rab_tree.get_children():
                    pzn = str(rab_tree.set(it, "pzn")).strip()
                    _raw = str(rab_tree.set(it, "rabatt")).strip().replace(",", ".")
                    if not pzn:
                        continue
                    try:
                        rabv = float(_raw) if _raw else None
                    except ValueError:
                        rabv = None
                    con.execute(
                        "INSERT OR REPLACE INTO tbl_pk_konditionen"
                        "(kundennummer,kundenname,pzn,rabatt_prozent,quelle,letzte_aktualisierung) "
                        "VALUES(?,?,?,?,?,?)", (knr, kname, pzn, rabv, "Kunden", _jetzt))
                con.commit()
            # Bestehende Analysen diesem Kunden zuordnen (Name -> Kundennummer).
            with _con() as con:
                cols_a = {r[1] for r in con.execute("PRAGMA table_info(tbl_auswertungen)").fetchall()}
                for col in ("kundennummer", "kundenname"):
                    if col not in cols_a:
                        con.execute(f"ALTER TABLE tbl_auswertungen ADD COLUMN {col} TEXT")
                con.execute(
                    "UPDATE tbl_auswertungen SET kundennummer=?, kundenname=? "
                    "WHERE (apotheke=? OR kundenname=?) AND (kundennummer IS NULL OR kundennummer='')",
                    (data["kundennummer"], data["kundenname"], data["kundenname"], data["kundenname"]))
                con.commit()
        except Exception as exc:
            messagebox.showerror(title, f"Speichern fehlgeschlagen:\n{exc}", parent=win)
            return
        win.destroy()
        if on_saved:
            try:
                on_saved()
            except Exception:
                pass

    tk.Button(bar, text="Abbrechen", command=win.destroy, padx=14, pady=7).pack(side="right", padx=(8, 12), pady=6)
    tk.Button(bar, text="✔  Speichern", command=save_kunde, bg=PRIMARY, fg="white",
              relief="flat", font=(FONT, 11, "bold"), padx=18, pady=7).pack(side="right", pady=6)


# ══════════════════════════════════════════════════════════════════════════════
#  Standalone-Start
# ══════════════════════════════════════════════════════════════════════════════
def run_standalone():
    """Startet die Kunden-App als eigenstaendiges Fenster (eigenes Taskleisten-Icon)."""
    if os.name == "nt":
        try:
            import ctypes
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("NMG.Kunden")
        except Exception:
            pass
    try:
        from .migrations import run_migrations
        run_migrations()
    except Exception:
        pass
    root = tk.Tk()
    root.title(fenstertitel("NMGone · Kunden"))
    root.geometry("1180x760")
    root.minsize(1000, 660)
    # Im Vollbild (maximiert) starten. 'zoomed' = Windows; sonst -zoomed/Bildschirmgroesse.
    try:
        root.state("zoomed")
    except tk.TclError:
        try:
            root.attributes("-zoomed", True)
        except tk.TclError:
            root.geometry(f"{root.winfo_screenwidth()}x{root.winfo_screenheight()}+0+0")
    root.configure(bg=BG)
    try:
        theme.apply_theme(root)
        theme.apply_widget_defaults(root)
    except Exception:
        pass
    for ico in ("Kunden.ico", "NMGone.ico"):
        try:
            root.iconbitmap(str(ASSETS_DIR / ico))
            break
        except Exception:
            continue
    KundenPanel(root, on_close=root.destroy).pack(fill="both", expand=True)
    root.mainloop()


def main():
    run_standalone()


if __name__ == "__main__":
    sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
    main()
