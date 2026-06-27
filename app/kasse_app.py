"""NMG Kasse-App (ehem. Bestell-App) - gemeinsame Oberflaeche fuer NMGone
(Toplevel) und die eigenstaendige Kasse-.exe (eigenes Fenster / Taskleisten-Icon).

Warenkreislauf: Wareneingang -> Lagerbestand -> Verkauf. Der Wareneingang
(Ware annehmen) wird in der App "Wareneingang & Retouren" (gdp_app.py) gebucht;
die Kasse zeigt den resultierenden Lagerbestand in der Artikel-Uebersicht und
verkauft daraus. UI in KassePanel (tk.Frame), u. a.:
  - Verkauf = Warenausgang an Apotheke (globale Kunden-/Artikelsuche,
              Bestand/Charge-Auswahl, Rabatt-Kaskade, Speichern -> Lager ab)
  - Artikel = Artikelstamm + Lagerbestand (Doppelklick auf Charge = Bestand
              korrigieren)

Datenmodell siehe migrations.py.
"""
from __future__ import annotations

import getpass
from .i18n import T as _T
import os
import sqlite3
import subprocess
import sys
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog, filedialog

from .config import DB_PATH, ASSETS_DIR, BASE_DIR, DEMO_SUFFIX
from . import kasse_import
from . import online_verkaeufe
from . import auftrag
from . import lieferschein
from . import defektmeldung
from . import kasse_reports
from . import einstellungen
from . import theme
from . import tour

# Redesign: Palette aus dem zentralen Theme beziehen (gemeinsamer Look mit NMGone).
BG = theme.CARD          # Inhaltsflaechen (Karten)
SHELL_BG = theme.BG      # Fensterhintergrund wie NMGone
ACCENT = theme.PRIMARY
ACCENT_DARK = theme.PRIMARY_DARK  # Hover/aktiv fuer Primaer-Aktionen
ACCENT_LIGHT = theme.SELECT_BG    # Akzent-Flaeche (Auswahl, Linien, Strips)
NAV_SEL = ACCENT_LIGHT   # Auswahl-Highlight wie NMGone
NAV_HOVER = "#F2F6FB"    # Maus-ueber in der Navigation
BORDER = theme.BORDER    # dezente Rahmenlinie
HEAD_BG = "#EEF3F8"      # Tabellenkopf-Hintergrund
TEXT = theme.INK         # Standard-Textfarbe
MUTED = theme.MUTED      # Sekundaer-/Hinweistext
OK_GREEN = theme.SUCCESS # Bestaetigen/Speichern (gruen)

# Dunkle Sidebar (einheitlich mit NMGone / Report)
SIDEBAR = theme.SIDEBAR
SIDEBAR_ACTIVE = theme.SIDEBAR_ACTIVE
SIDEBAR_TEXT = theme.SIDEBAR_TEXT
SIDEBAR_MUTED = theme.SIDEBAR_MUTED


def _make_card(parent, title=None, subtitle=None):
    """Moderne Karte (weisse Flaeche mit dezenter 1px-Linie und optionalem
    Akzent-Titelkopf). Liefert (outer, body): outer wird vom Aufrufer gepackt/
    gegrid'et, die Inhalte kommen in body. Ersetzt die altmodischen LabelFrames."""
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
    body.pack(fill="both", expand=True, padx=14, pady=(10, 12))
    return outer, body


def _eur(v):
    if v is None:
        return "—"
    return f"{v:,.2f} €".replace(",", "X").replace(".", ",").replace("X", ".")


def _pos_netto(p):
    """Netto-Zeilensumme einer Position (APU * Menge * (1 - Rabatt%))."""
    if p.get("apu") is None:
        return 0.0
    return p["apu"] * (p.get("menge") or 0) * (1 - (p.get("rabatt") or 0) / 100.0)


def _make_date_entry(parent, default_today=True):
    """Kalender-Datumsfeld (tkcalendar.DateEntry) wenn verfuegbar, sonst einfaches
    Entry als Fallback. Wert immer ueber _read_date lesen (liefert 'YYYY-MM-DD')."""
    try:
        from tkcalendar import DateEntry
        return DateEntry(parent, date_pattern="yyyy-mm-dd", width=12,
                         background=ACCENT, foreground="white", borderwidth=1)
    except Exception:
        from datetime import date as _d
        e = tk.Entry(parent, width=12)
        if default_today:
            e.insert(0, _d.today().isoformat())
        return e


def _read_date(widget) -> str:
    """Datum eines _make_date_entry-Widgets als 'YYYY-MM-DD' (leer wenn nichts)."""
    try:
        return widget.get_date().isoformat()      # DateEntry
    except Exception:
        return (widget.get() or "").strip()


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
    Auch ohne Trenner: '1226' -> '12/26', '0127' -> '01/27', '122026' -> '12/2026'.
    Ungueltig (z.B. Monat 13) -> None."""
    t = str(t or "").strip().replace(".", "/").replace("-", "/").replace(" ", "")
    if not t:
        return ""
    # Reine Ziffernfolge ohne Trenner -> Monat/Jahr nach Laenge aufteilen.
    if "/" not in t and t.isdigit():
        if len(t) in (3, 4):       # M+JJ / MM+JJ
            t = f"{t[:-2]}/{t[-2:]}"
        elif len(t) in (5, 6):     # M+JJJJ / MM+JJJJ
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
                menge INTEGER DEFAULT 0, ek REAL, aktualisiert_am TEXT,
                UNIQUE(pzn, charge, verfall)
            )"""
        )
        # EK (Einkaufspreis) je Lagerzeile - nachruesten fuer bestehende DBs.
        if "ek" not in {r[1] for r in con.execute("PRAGMA table_info(tbl_lagerbestand)")}:
            con.execute("ALTER TABLE tbl_lagerbestand ADD COLUMN ek REAL")
        con.execute(
            """CREATE TABLE IF NOT EXISTS tbl_kasse_log(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                zeitpunkt TEXT DEFAULT CURRENT_TIMESTAMP,
                bearbeiter TEXT, aktion TEXT, bestell_id INTEGER, kunde TEXT, details TEXT
            )"""
        )
        # Tagesabschluss-Register (laufende Nr + festgeschriebene Kennzahlen pro Tag).
        con.execute(
            """CREATE TABLE IF NOT EXISTS tbl_kasse_tagesabschluss(
                datum TEXT PRIMARY KEY, nr INTEGER, erzeugt_am TEXT, erzeugt_von TEXT,
                anzahl INTEGER, pakete INTEGER, brutto REAL, rabatt REAL, netto REAL
            )"""
        )
        # In der App editierbare Einstellungen (Texte, Firmendaten, Parameter).
        con.execute(
            """CREATE TABLE IF NOT EXISTS tbl_kasse_einstellungen(
                schluessel TEXT PRIMARY KEY, wert TEXT, geaendert_am TEXT, bearbeiter TEXT
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
    """Kasse-Oberflaeche: Verkauf + Artikel/Lager (Wareneingang siehe gdp_app)."""

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
        self.dm_kunde = None          # Defektmeldung: gewaehlte Apotheke
        self.dm_cur = None            # Defektmeldung: aktuell gewaehlter Artikel
        self.dm_positions = []        # Defektmeldung: Artikelliste
        self._online_verkaeufe = []        # Cache: Online-Verkäufe
        self._online_vorbestellungen = []  # Cache: Online-Vorbestellungen
        self._online_uebernahme_id = None  # läuft. Übernahme -> nach Speichern markieren
        self._online_cache_laden()         # zuletzt geladene Online-Bestellungen sofort zeigen
        self._build()
        # Automatischer Tagesabschluss: verpasste Vortage nachholen + Timer fuer heute.
        self._auto_tagesabschluss_setup()
        # Online-Verkäufe (Handy/Web) automatisch alle 5 Minuten nachladen.
        self._online_auto_setup()
        # Bestand alle 2 Minuten zum Server pushen (PC ist die Bestands-Quelle).
        self._online_push_setup()

    # =============================================================== Datenbank
    def _conn(self):
        # busy_timeout: bei parallelem Zugriff (NMGone + Kasse-.exe auf dieselbe
        # DB, Ziel Mehrbenutzer) nicht sofort mit "database is locked" abbrechen,
        # sondern bis zu 30 s auf die Sperre warten. WAL ist bereits aktiv
        # (ensure_kasse_tables).
        con = sqlite3.connect(self.db_path, timeout=30)
        con.execute("PRAGMA busy_timeout=30000")
        return con

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

    def _unterwegs_total(self, pzn):
        """Klammer-Bestand: Menge aus Warenausgaengen, die geladen/vorgemerkt
        sind (GDP-Status 'geladen'), aber noch NICHT bei MSK bestaetigt wurden.
        Nur Anzeige – diese Menge ist NICHT verkaufbar, sie signalisiert nur
        eintreffenden Nachschub."""
        with self._conn() as con:
            if not _table_exists(con, "tbl_gdp_warenausgang"):
                return 0
            r = con.execute(
                """SELECT COALESCE(SUM(p.menge),0)
                     FROM tbl_gdp_warenausgang_pos p
                     JOIN tbl_gdp_warenausgang w ON w.id = p.wa_id
                    WHERE p.pzn = ? AND w.status = 'geladen'
                      AND w.ziel = 'Verkaufsbestand'""",
                (pzn,)).fetchone()
            return int(r[0]) if r and r[0] else 0

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

    def _setup_style(self):
        """Zentrales ttk-Theme fuer die Kasse. Bewusst EIGENE, benannte Styles
        (Kasse.*), damit ein eingebettetes NMGone (gemeinsamer ttk-Style-Pool)
        nicht mit-veraendert wird. Modernisiert Tabellen und Comboboxen."""
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure("Kasse.Treeview", background=BG, fieldbackground=BG,
                        foreground=TEXT, borderwidth=0, rowheight=28,
                        font=(theme.FONT, 10))
        style.configure("Kasse.Treeview.Heading", font=(theme.FONT, 10, "bold"),
                        background=HEAD_BG, foreground=ACCENT, relief="flat",
                        padding=(8, 6), borderwidth=0)
        style.map("Kasse.Treeview",
                  background=[("selected", ACCENT_LIGHT)],
                  foreground=[("selected", ACCENT)])
        style.map("Kasse.Treeview.Heading",
                  background=[("active", "#e0e9f4")])
        style.configure("Kasse.TCombobox", fieldbackground=BG, background=BG,
                        bordercolor=BORDER, arrowcolor=ACCENT, padding=2)
        # Reiter (Notebook) im Kasse-Look: flache Tabs, aktiver Tab im Akzentblau.
        style.configure("Kasse.TNotebook", background=BG, borderwidth=0, tabmargins=(2, 6, 2, 0))
        style.configure("Kasse.TNotebook.Tab", font=(theme.FONT, 10, "bold"),
                        padding=(18, 9), background=HEAD_BG, foreground=MUTED,
                        borderwidth=0)
        style.map("Kasse.TNotebook.Tab",
                  background=[("selected", ACCENT), ("active", "#e0e9f4")],
                  foreground=[("selected", "#FFFFFF"), ("active", ACCENT)])

    def _metric_cards(self, parent, items):
        """Reihe von Kennzahl-Karten im Tagesabschluss-Stil. items = Liste aus
        (key, label) oder (key, label, fg). Liefert (frame, {key: wert-Label})."""
        row = tk.Frame(parent, bg=BG)
        row.pack(fill="x", padx=8, pady=(6, 8))
        out = {}
        for item in items:
            key, label = item[0], item[1]
            fg = item[2] if len(item) > 2 else ACCENT
            c = tk.Frame(row, bg="#f2f6fb", highlightbackground="#dde7f1", highlightthickness=1)
            c.pack(side="left", fill="both", expand=True, padx=(0, 8))
            tk.Label(c, text=label, bg="#f2f6fb", fg="#666",
                     font=(theme.FONT, 9)).pack(anchor="w", padx=10, pady=(8, 0))
            val = tk.Label(c, text="–", bg="#f2f6fb", fg=fg, font=(theme.FONT, 14, "bold"))
            val.pack(anchor="w", padx=10, pady=(0, 8))
            out[key] = val
        return row, out

    def _build(self):
        self.configure(bg=SHELL_BG)
        self.columnconfigure(1, weight=1)
        self.rowconfigure(0, weight=1)
        self._setup_style()
        self._current_view = None

        # ---------------- linke Menueleiste (zentrale theme.Sidebar) ----------
        self.sidebar = theme.Sidebar(self, width=248, title="Kasse",
                                     subtitle="Verkauf & Lager")
        self.sidebar.grid(row=0, column=0, sticky="ns")
        self._app_icon = theme.load_icon(ASSETS_DIR / "Kasse.ico", 60)
        if self._app_icon:
            self.sidebar.set_logo(self._app_icon)
        self.sidebar.add_section("Kasse")
        for key, text, icon in (("uebersicht", "Übersicht", "📊"),
                                 ("verkauf", "Verkauf", "🛒"),
                                 ("vorbestellungen", "Vorbestellungen", "🕓"),
                                 ("verkaeufe", "Verkäufe", "🧾"),
                                 ("artikel", "Artikel", "🔍"),
                                 ("defektmeldung", "Defektmeldung", "⚠"),
                                 ("auswertung", "Auswertung", "📈"),
                                 ("protokoll", "Protokoll", "📝"),
                                 ("einstellungen", "Einstellungen", "⚙")):
            self.sidebar.add_item(key, icon, text, lambda k=key: self._show_view(k))

        foot = self.sidebar.footer()
        tk.Button(foot, text="🏠  NMGone öffnen", command=self._open_nmgone,
                  bg=SIDEBAR_ACTIVE, fg="#FFFFFF", relief="flat", font=(theme.FONT, 10, "bold"),
                  activebackground="#1B5085", activeforeground="#FFFFFF",
                  padx=10, pady=7, cursor="hand2").pack(fill="x", padx=10, pady=(0, 6))
        tk.Button(foot, text="Schließen", command=self._on_close, relief="flat",
                  bg="#0E3454", fg=SIDEBAR_TEXT, activebackground="#15466E",
                  activeforeground="#FFFFFF", padx=10, pady=5, cursor="hand2").pack(fill="x", padx=10)
        self.sidebar.add_footer_note(_T('Datenbank:\n{p0}', p0=Path(self.db_path).name))

        # ---------------- Hauptbereich ----------------
        main = tk.Frame(self, bg=SHELL_BG)
        main.grid(row=0, column=1, sticky="nsew")
        main.columnconfigure(0, weight=1)
        main.rowconfigure(1, weight=1)

        head = tk.Frame(main, bg=SHELL_BG)
        head.grid(row=0, column=0, sticky="ew", padx=22, pady=(18, 0))
        self._view_title = tk.Label(head, text="Verkauf", font=(theme.FONT, 18, "bold"),
                                    fg=ACCENT, bg=SHELL_BG)
        self._view_title.pack(anchor="w")
        self._view_subtitle = tk.Label(head, text="", font=(theme.FONT, 10),
                                       fg=MUTED, bg=SHELL_BG)
        self._view_subtitle.pack(anchor="w", pady=(1, 0))
        tk.Frame(head, bg=ACCENT_LIGHT, height=3).pack(fill="x", pady=(10, 0))

        page = tk.Frame(main, bg=BG, highlightbackground=BORDER, highlightthickness=1)
        page.grid(row=1, column=0, sticky="nsew", padx=22, pady=12)
        page.columnconfigure(0, weight=1)
        page.rowconfigure(0, weight=1)

        self._views = {}
        for key, builder in (("uebersicht", self._build_uebersicht),
                             ("verkauf", self._build_verkauf),
                             ("vorbestellungen", self._build_vorbestellungen),
                             ("verkaeufe", self._build_verkaeufe),
                             ("artikel", self._build_artikel),
                             ("defektmeldung", self._build_defektmeldung),
                             ("auswertung", self._build_auswertung),
                             ("protokoll", self._build_protokoll),
                             ("einstellungen", self._build_einstellungen)):
            frame = tk.Frame(page, bg=BG)
            frame.grid(row=0, column=0, sticky="nsew")
            builder(frame)
            self._views[key] = frame

        self._show_view("uebersicht")
        self._restyle_buttons(self)

    def _show_view(self, key):
        self._current_view = key
        self._views[key].tkraise()
        titel = {"uebersicht": "Übersicht",
                 "verkauf": "Verkauf",
                 "vorbestellungen": "Vorbestellungen",
                 "verkaeufe": "Verkäufe",
                 "artikel": "Artikel-Übersicht",
                 "defektmeldung": "Defektmeldung",
                 "auswertung": "Auswertung",
                 "protokoll": "Änderungs-Protokoll",
                 "einstellungen": "Einstellungen"}.get(key, key)
        untertitel = {"uebersicht": "Tagesgeschäft, offene Aufgaben und Lager auf einen Blick",
                      "verkauf": "Warenausgang an die Apotheke erfassen",
                      "vorbestellungen": "Offene Vorbestellungen disponieren – Kunden anrufen",
                      "verkaeufe": "Abgeschlossene Verkäufe & MSK-Status",
                      "artikel": "Artikelstamm und aktueller Lagerbestand",
                      "defektmeldung": "Nichtverfügbarkeit für die Apotheke bescheinigen",
                      "auswertung": "Verfall, Inventur und Umsatz/Tagesabschluss",
                      "protokoll": "Wer hat was wann geändert",
                      "einstellungen": "Firmendaten, Dokument-Texte und Parameter anpassen"}.get(key, "")
        self._view_title.config(text=titel)
        self._view_subtitle.config(text=untertitel)
        if key == "uebersicht":
            self._refresh_uebersicht()
        elif key == "vorbestellungen":
            self._refresh_vorbestellungen()
        elif key == "verkaeufe":
            self._refresh_verkaeufe()
        elif key == "artikel":
            self._refresh_artikel()
        elif key == "auswertung":
            self._refresh_auswertung()
        elif key == "protokoll":
            self._refresh_log()
        self.sidebar.set_active(key)

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
            messagebox.showerror("NMGone", _T('NMGone konnte nicht geöffnet werden:\n{p0}', p0=e),
                                 parent=self.winfo_toplevel())

    # =============================================================== Übersicht
    def _build_uebersicht(self, parent):
        """Startseite: Tagesgeschäft, offene Aufgaben und Lager auf einen Blick."""
        bar = tk.Frame(parent, bg=BG)
        bar.pack(fill="x", padx=8, pady=(10, 0))
        self._db_greeting = tk.Label(bar, text="", bg=BG, fg="#666", font=(theme.FONT, 10))
        self._db_greeting.pack(side="left")
        tk.Button(bar, text="🔄 Aktualisieren", command=self._refresh_uebersicht,
                  bg="#d8e2ee", fg=ACCENT, relief="flat", font=(theme.FONT, 9, "bold"),
                  padx=10, pady=4, cursor="hand2").pack(side="right")

        # Reihe 1: Tagesgeschäft & offene Aufgaben
        _, self._db_cards = self._metric_cards(parent, [
            ("umsatz_heute", "Umsatz heute (netto)"),
            ("verkaeufe_heute", "Verkäufe heute"),
            ("offene_auftraege", "Offene Aufträge", "#b45309"),
            ("offene_vorbestellungen", "Offene Vorbestellungen", "#b45309"),
        ])
        # Reihe 2: Lager & Verfall
        _, self._db_cards2 = self._metric_cards(parent, [
            ("verkaufswert", "Verkaufswert (APU)"),
            ("lagerwert", "Lagerwert (EK)"),
            ("bald_anzahl", "Bald ablaufend", "#b45309"),
            ("abgelaufen_anzahl", "Abgelaufen", "#b00020"),
        ])

        # Zwei Panels nebeneinander: Verfall-Warnungen | letzte Verkäufe
        mid = tk.Frame(parent, bg=BG)
        mid.pack(fill="both", expand=True, padx=8, pady=(2, 10))
        mid.columnconfigure(0, weight=1, uniform="db")
        mid.columnconfigure(1, weight=1, uniform="db")
        mid.rowconfigure(0, weight=1)

        # -- Panel links: Handlungsbedarf Verfall --
        left = tk.Frame(mid, bg="#f2f6fb", highlightbackground="#dde7f1", highlightthickness=1)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 6))
        lhead = tk.Frame(left, bg="#f2f6fb")
        lhead.pack(fill="x", padx=10, pady=(8, 2))
        tk.Label(lhead, text="⚠  Verfall im Blick", bg="#f2f6fb", fg=ACCENT,
                 font=(theme.FONT, 11, "bold")).pack(side="left")
        tk.Button(lhead, text="→ Auswertung", command=lambda: self._show_view("auswertung"),
                  bg="#f2f6fb", fg=ACCENT, relief="flat", font=(theme.FONT, 9, "bold"),
                  cursor="hand2", bd=0).pack(side="right")
        ltf = tk.Frame(left, bg="#f2f6fb")
        ltf.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        vcols = ("artikel", "charge", "verfall", "menge", "status")
        self._db_verfall_tree = ttk.Treeview(ltf, columns=vcols, show="headings",
                                              height=8, style="Kasse.Treeview")
        for c, t, w in (("artikel", "Artikel", 170), ("charge", "Charge", 90),
                        ("verfall", "Verfall", 70), ("menge", "Menge", 55),
                        ("status", "Status", 90)):
            self._db_verfall_tree.heading(c, text=t)
            self._db_verfall_tree.column(c, width=w, anchor="w")
        self._db_verfall_tree.tag_configure("abgelaufen", background="#fde2e2")
        self._db_verfall_tree.tag_configure("bald", background="#fff4d6")
        self._db_verfall_tree.pack(side="left", fill="both", expand=True)
        vsb = tk.Scrollbar(ltf, orient="vertical", command=self._db_verfall_tree.yview)
        vsb.pack(side="right", fill="y")
        self._db_verfall_tree.configure(yscrollcommand=vsb.set)

        # -- Panel rechts: letzte Verkäufe --
        right = tk.Frame(mid, bg="#f2f6fb", highlightbackground="#dde7f1", highlightthickness=1)
        right.grid(row=0, column=1, sticky="nsew", padx=(6, 0))
        rhead = tk.Frame(right, bg="#f2f6fb")
        rhead.pack(fill="x", padx=10, pady=(8, 2))
        tk.Label(rhead, text="🧾  Letzte Verkäufe", bg="#f2f6fb", fg=ACCENT,
                 font=(theme.FONT, 11, "bold")).pack(side="left")
        tk.Button(rhead, text="→ Verkäufe", command=lambda: self._show_view("verkaeufe"),
                  bg="#f2f6fb", fg=ACCENT, relief="flat", font=(theme.FONT, 9, "bold"),
                  cursor="hand2", bd=0).pack(side="right")
        rtf = tk.Frame(right, bg="#f2f6fb")
        rtf.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        lcols = ("nr", "datum", "apotheke", "status", "netto")
        self._db_letzte_tree = ttk.Treeview(rtf, columns=lcols, show="headings",
                                             height=8, style="Kasse.Treeview")
        for c, t, w in (("nr", "Nr", 50), ("datum", "Datum", 80),
                        ("apotheke", "Apotheke", 170), ("status", "Status", 70),
                        ("netto", "Netto €", 80)):
            self._db_letzte_tree.heading(c, text=t)
            self._db_letzte_tree.column(c, width=w, anchor="w")
        self._db_letzte_tree.pack(side="left", fill="both", expand=True)
        rsb = tk.Scrollbar(rtf, orient="vertical", command=self._db_letzte_tree.yview)
        rsb.pack(side="right", fill="y")
        self._db_letzte_tree.configure(yscrollcommand=rsb.set)
        self._db_letzte_tree.bind("<Double-1>", self._db_open_letzter)
        tk.Label(right, text="Doppelklick öffnet den Verkauf.", bg="#f2f6fb", fg="#888",
                 font=(theme.FONT, 8)).pack(anchor="w", padx=10, pady=(0, 8))

    def _db_open_letzter(self, _e=None):
        sel = self._db_letzte_tree.selection()
        if not sel:
            return
        try:
            bid = int(self._db_letzte_tree.item(sel[0], "values")[0])
        except (ValueError, IndexError):
            return
        self._verkauf_detail_window(bid)

    def _refresh_uebersicht(self):
        if not hasattr(self, "_db_cards"):
            return
        try:
            k = kasse_reports.dashboard_kennzahlen(self.db_path)
        except Exception as exc:
            self._db_greeting.config(text=_T('Kennzahlen konnten nicht geladen werden: {p0}', p0=exc))
            return
        std = datetime.now().hour
        gruss = "Guten Morgen" if std < 11 else "Guten Tag" if std < 18 else "Guten Abend"
        self._db_greeting.config(
            text=_T('{p0} – Überblick für {p1:%d.%m.%Y}, {p2} Packungen heute verkauft.', p0=gruss, p1=datetime.now(), p2=k['pakete_heute']))
        warn = k.get("warn_tage", 90)
        self._db_cards["umsatz_heute"].config(text=_eur(k["umsatz_heute"]))
        self._db_cards["verkaeufe_heute"].config(text=str(k["verkaeufe_heute"]))
        self._db_cards["offene_auftraege"].config(text=str(k["offene_auftraege"]))
        self._db_cards["offene_vorbestellungen"].config(text=str(k["offene_vorbestellungen"]))
        self._db_cards2["verkaufswert"].config(text=_eur(k["verkaufswert"]))
        self._db_cards2["lagerwert"].config(text=_eur(k["lagerwert"]))
        self._db_cards2["bald_anzahl"].config(text=str(k["bald_anzahl"]))
        self._db_cards2["abgelaufen_anzahl"].config(text=str(k["abgelaufen_anzahl"]))

        self._db_verfall_tree.delete(*self._db_verfall_tree.get_children())
        verfall = list(k["abgelaufen"]) + list(k["bald"])
        if verfall:
            for r in verfall:
                label = "Abgelaufen" if r["status"] == "abgelaufen" else f"≤{warn} Tage"
                self._db_verfall_tree.insert(
                    "", "end",
                    values=(r["artikelname"] or r["pzn"], r["charge"] or "—",
                            r["verfall"] or "—", r["menge"], label),
                    tags=(r["status"],))
        else:
            self._db_verfall_tree.insert("", "end", values=("Keine Verfall-Warnungen 👍", "", "", "", ""))

        self._db_letzte_tree.delete(*self._db_letzte_tree.get_children())
        if k["letzte_verkaeufe"]:
            for bid, datum, apotheke, status, netto in k["letzte_verkaeufe"]:
                self._db_letzte_tree.insert(
                    "", "end",
                    values=(bid, (datum or "")[:10], apotheke or "—", status, _eur(netto)))
        else:
            self._db_letzte_tree.insert("", "end", values=("", "", "Noch keine Verkäufe", "", ""))

    # ================================================================= Verkauf
    def _build_verkauf(self, parent):
        # Auftrag-Nr nachschlagen (read-only Detailansicht, KEIN neuer Auftrag).
        nrbar = tk.Frame(parent, bg=BG)
        nrbar.pack(fill="x", padx=8, pady=(10, 0))
        tk.Label(nrbar, text="Auftrag-Nr anzeigen:", bg=BG, font=(theme.FONT, 9, "bold")).pack(side="left")
        self._vk_nr_var = tk.StringVar()
        en = tk.Entry(nrbar, textvariable=self._vk_nr_var, width=8)
        en.pack(side="left", padx=(4, 4))
        en.bind("<Return>", lambda _e: self._vk_open_auftrag_nr())
        tk.Button(nrbar, text="Anzeigen", command=self._vk_open_auftrag_nr,
                  font=(theme.FONT, 8), padx=6, pady=1).pack(side="left")
        tk.Label(nrbar, text="(nur ansehen – legt keinen Auftrag an)", bg=BG, fg="#999",
                 font=(theme.FONT, 8)).pack(side="left", padx=(8, 0))

        kopf_card, kopf = _make_card(parent, "Kunde")
        kopf_card.pack(fill="x", padx=8, pady=(10, 6))

        khead = tk.Frame(kopf, bg=BG)
        khead.pack(fill="x")
        tk.Label(khead, text="Kunde suchen:", bg=BG, font=(theme.FONT, 10, "bold")).pack(side="left")
        tk.Button(khead, text="✕ Kunde entfernen", command=self._vk_clear_kunde,
                  font=(theme.FONT, 8), padx=6, pady=1).pack(side="left", padx=(10, 0))
        tk.Button(khead, text="📥 Kundenliste importieren", command=self._import_kunden,
                  font=(theme.FONT, 8), padx=6, pady=1).pack(side="right")
        tk.Button(khead, text="📊 Gekaufte Artikel", command=self._show_kunde_top,
                  font=(theme.FONT, 8), padx=6, pady=1).pack(side="right", padx=(0, 6))
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
                                            font=(theme.FONT, 9, "italic"))
        self.vk_kunde_card_label.pack(fill="x", padx=10, pady=8)

        # Artikel hinzufuegen
        addf_card, addf = _make_card(parent, "Artikel hinzufügen", "nur NMG")
        addf_card.pack(fill="x", padx=8, pady=(0, 6))
        self.vk_artikel_search = SearchBox(
            addf,
            fetch=lambda t: [(f"{pzn}  ·  {name}", pzn) for pzn, name in self._search_nmg(t)],
            on_select=self._vk_pick_artikel, height=5,
        )
        self.vk_artikel_search.pack(fill="x", pady=(2, 4))

        line = tk.Frame(addf, bg=BG)
        line.pack(fill="x", pady=(2, 0))
        self.vk_detail_label = tk.Label(line, text="—", bg=BG, fg="#444",
                                        font=(theme.FONT, 9), anchor="w")
        self.vk_detail_label.pack(side="left", fill="x", expand=True)

        line2 = tk.Frame(addf, bg=BG)
        line2.pack(fill="x", pady=(6, 0))
        tk.Label(line2, text="Charge/Bestand:", bg=BG, font=(theme.FONT, 9, "bold")).pack(side="left")
        self.vk_charge_var = tk.StringVar()
        self.vk_charge_cmb = ttk.Combobox(line2, textvariable=self.vk_charge_var,
                                          width=30, state="readonly")
        self.vk_charge_cmb.pack(side="left", padx=(4, 12))
        tk.Label(line2, text="Rabatt %:", bg=BG, font=(theme.FONT, 9, "bold")).pack(side="left")
        self.vk_rabatt_var = tk.StringVar()
        tk.Entry(line2, textvariable=self.vk_rabatt_var, width=6).pack(side="left", padx=(4, 12))
        tk.Label(line2, text="Menge:", bg=BG, font=(theme.FONT, 9, "bold")).pack(side="left")
        self.vk_menge_var = tk.StringVar(value="1")
        tk.Entry(line2, textvariable=self.vk_menge_var, width=5).pack(side="left", padx=(4, 12))

        # Liefervorgabe PRO Position (je Artikel eigene Vorgabe moeglich).
        line3 = tk.Frame(addf, bg=BG)
        line3.pack(fill="x", pady=(6, 0))
        tk.Label(line3, text="Liefervorgabe:", bg=BG, font=(theme.FONT, 9, "bold")).pack(side="left")
        self.vk_pos_art_var = tk.StringVar(value="Bestellung")
        ttk.Combobox(line3, textvariable=self.vk_pos_art_var, width=12, state="readonly",
                     values=("Bestellung", "Vorbestellung", "abgesagt")).pack(side="left", padx=(4, 8))
        self.vk_pos_lieferzeit_var = tk.StringVar(value="Frei")
        ttk.Combobox(line3, textvariable=self.vk_pos_lieferzeit_var, width=6, state="readonly",
                     values=("Frei", "10 Uhr", "12 Uhr")).pack(side="left", padx=(0, 4))
        self.vk_pos_uhrzeit_var = tk.StringVar()
        tk.Entry(line3, textvariable=self.vk_pos_uhrzeit_var, width=7).pack(side="left", padx=(0, 2))
        tk.Label(line3, text="hh:mm", bg=BG, fg="#999", font=(theme.FONT, 8)).pack(side="left", padx=(0, 8))
        tk.Label(line3, text="Termin:", bg=BG, font=(theme.FONT, 9, "bold")).pack(side="left")
        self.vk_pos_termin_var = tk.StringVar(value=_next_liefertermin())
        tk.Entry(line3, textvariable=self.vk_pos_termin_var, width=11).pack(side="left", padx=(4, 12))
        tk.Button(line3, text="+ Position", command=self._vk_add_position,
                  bg="#11823b", fg="white", font=(theme.FONT, 9, "bold"),
                  padx=10, pady=2).pack(side="left")

        # Positionsliste
        pos_card, pos = _make_card(parent, "Positionen")
        pos_card.pack(fill="both", expand=True, padx=8, pady=(0, 6))
        pcols = ("pzn", "artikel", "df", "pck", "apu", "menge", "rabatt", "charge", "verfall", "liefer")
        ph = {"pzn": "PZN", "artikel": "Artikel", "df": "DF", "pck": "PCK", "apu": "APU",
              "menge": "Menge", "rabatt": "Rab%", "charge": "Charge", "verfall": "Verfall",
              "liefer": "Liefervorgabe"}
        tf = tk.Frame(pos, bg=BG)
        tf.pack(fill="both", expand=True)
        tree = ttk.Treeview(tf, columns=pcols, show="headings", height=6,
                            style="Kasse.Treeview")
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
        # Freie (nicht bestandsgefuehrte) Positionen dezent hervorheben.
        tree.tag_configure("frei", background="#eef3f8", foreground="#11304d")
        tree.bind("<Delete>", self._vk_remove_position)
        tree.bind("<Double-1>", self._vk_edit_position)
        self.vk_pos_tree = tree

        afoot = tk.Frame(parent, bg=BG)
        afoot.pack(fill="x", padx=8, pady=(0, 8))
        tk.Button(afoot, text="➕ Freie Position", command=self._vk_freie_position_dialog,
                  bg="#3867b7", fg="white", font=(theme.FONT, 9, "bold"),
                  padx=10, pady=2).pack(side="left")
        tk.Button(afoot, text="🗑 Position löschen", command=lambda: self._vk_remove_position(None),
                  font=(theme.FONT, 9), padx=8, pady=2).pack(side="left", padx=(8, 0))
        tk.Label(afoot, text="  (oder Entf · Doppelklick = bearbeiten)",
                 bg=BG, fg="#999", font=(theme.FONT, 8)).pack(side="left")
        tk.Button(afoot, text="Verkauf speichern", command=self._vk_save,
                  bg=ACCENT, fg="white", font=(theme.FONT, 10, "bold"),
                  padx=14, pady=4).pack(side="right")
        self._vk_total_label = tk.Label(afoot, text="Gesamt (netto): 0,00 €", bg=BG,
                                        fg=ACCENT, font=(theme.FONT, 13, "bold"))
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

    def _kunde_top_artikel(self, knr, monate=12, limit=None):
        """Vom Kunden gekaufte Artikel nach Menge in den letzten N Monaten.
        Nur echte Bestellungen, keine stornierten Verkaeufe. limit=None -> alle."""
        from datetime import date, timedelta
        cutoff = (date.today() - timedelta(days=int(monate * 30.4))).isoformat()
        sql = (
            "SELECT p.pzn, p.artikelname, SUM(p.menge), COUNT(DISTINCT b.id), "
            "SUM(p.apu*p.menge*(1-COALESCE(p.rabatt_prozent,0)/100.0)) "
            "FROM tbl_bestellpositionen p JOIN tbl_bestellungen b ON b.id=p.bestell_id "
            "WHERE b.kundennummer=? AND p.bestellart='Bestellung' "
            "AND COALESCE(b.status,'offen')<>'storniert' AND b.datum >= ? "
            "GROUP BY p.pzn, p.artikelname ORDER BY SUM(p.menge) DESC, SUM(p.apu*p.menge) DESC")
        params = [knr, cutoff]
        if limit is not None:
            sql += " LIMIT ?"
            params.append(limit)
        with self._conn() as con:
            return con.execute(sql, params).fetchall()

    def _show_kunde_top(self):
        top = self.winfo_toplevel()
        if not self.vk_kunde:
            messagebox.showinfo("Top-Artikel", "Bitte zuerst einen Kunden wählen.", parent=top)
            return
        knr = self.vk_kunde.get("kundennummer")
        win = tk.Toplevel(top)
        win.title("Gekaufte Artikel")
        win.configure(bg=BG)
        win.transient(top)
        win.geometry("680x460")
        tk.Label(win, text=_T('Gekaufte Artikel · {p0}', p0=self.vk_kunde.get('kundenname', '')),
                 font=(theme.FONT, 13, "bold"), fg=ACCENT, bg=BG).pack(padx=18, pady=(14, 2), anchor="w")

        # Zeitraum-Auswahl: 3 / 6 / 12 Monate oder freie Eingabe.
        zrow = tk.Frame(win, bg=BG)
        zrow.pack(fill="x", padx=18, pady=(2, 6))
        tk.Label(zrow, text="Zeitraum:", font=(theme.FONT, 9, "bold"), bg=BG).pack(side="left")
        zeit_var = tk.StringVar(value="12 Monate")
        ttk.Combobox(zrow, textvariable=zeit_var, width=12, state="readonly",
                     values=("3 Monate", "6 Monate", "12 Monate", "Frei")).pack(side="left", padx=(4, 8))
        tk.Label(zrow, text="frei (Monate):", font=(theme.FONT, 9), bg=BG).pack(side="left")
        frei_var = tk.StringVar(value="12")
        tk.Entry(zrow, textvariable=frei_var, width=5).pack(side="left", padx=(4, 8))
        info_lbl = tk.Label(win, text="", font=(theme.FONT, 9), fg="#555", bg=BG)
        info_lbl.pack(padx=18, anchor="w", pady=(0, 6))

        tf = tk.Frame(win, bg=BG)
        tf.pack(fill="both", expand=True, padx=18, pady=(0, 8))
        cols = ("rang", "pzn", "artikel", "menge", "anzahl", "umsatz")
        heads = {"rang": "#", "pzn": "PZN", "artikel": "Artikel", "menge": "Menge",
                 "anzahl": "Bestellungen", "umsatz": "Umsatz netto"}
        tree = ttk.Treeview(tf, columns=cols, show="headings", style="Kasse.Treeview")
        for c in cols:
            tree.heading(c, text=heads[c])
            w = 240 if c == "artikel" else (30 if c == "rang" else 90)
            tree.column(c, width=w, anchor="w")
        tree.pack(side="left", fill="both", expand=True)
        sb = tk.Scrollbar(tf, orient="vertical", command=tree.yview)
        sb.pack(side="right", fill="y")
        tree.configure(yscrollcommand=sb.set)
        _make_treeview_sortable(tree)

        def _monate():
            if zeit_var.get() == "Frei":
                try:
                    m = int(frei_var.get().strip())
                    return m if m > 0 else 12
                except ValueError:
                    return 12
            return {"3 Monate": 3, "6 Monate": 6, "12 Monate": 12}.get(zeit_var.get(), 12)

        def _refresh(*_):
            monate = _monate()
            rows = self._kunde_top_artikel(knr, monate=monate, limit=None)
            tree.delete(*tree.get_children())
            for i, (pzn, art, menge, anzahl, umsatz) in enumerate(rows, 1):
                tree.insert("", "end", values=(i, pzn, art, int(menge or 0), anzahl, _eur(umsatz)))
            info_lbl.config(text=(f"{len(rows)} Artikel in den letzten {monate} Monaten "
                                  "(nur abgeschlossene Verkäufe, ohne Stornos)."
                                  if rows else f"Keine Käufe in den letzten {monate} Monaten."))

        zeit_var.trace_add("write", _refresh)
        frei_var.trace_add("write", lambda *_: zeit_var.get() == "Frei" and _refresh())
        _refresh()
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
                _T('{p0} hat {p1} offene Vorbestellung(en).\n\nJetzt im Reiter „Vorbestellungen“ ansehen und ggf. als Verkauf übernehmen?', p0=self.vk_kunde['kundenname'], p1=n),
                parent=self.winfo_toplevel()):
                self._vb_filter_var.set(self.vk_kunde["kundenname"] or knr)
                self._show_view("vorbestellungen")

    def _render_kunde_card(self, k):
        if not k:
            self.vk_kunde_card_label.config(text="Kein Kunde gewählt.", fg="#888",
                                            font=(theme.FONT, 9, "italic"))
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
        self.vk_kunde_card_label.config(text=txt, fg="#11304d", font=(theme.FONT, 9))

    def _vk_clear_kunde(self):
        """Gewaehlten Kunden im Verkauf mit einem Klick entfernen."""
        self.vk_kunde = None
        self.vk_kunde_search.clear()
        self._render_kunde_card(None)

    def _vk_open_auftrag_nr(self):
        """Im Verkauf eine bestehende Auftrag-Nr nur ANZEIGEN (kein neuer Auftrag)."""
        bid = self._auftrag_id_eingabe(self._vk_nr_var.get())
        if bid is None:
            messagebox.showinfo("Auftrag anzeigen", "Bitte eine Auftrag-Nr (Zahl) eingeben.",
                                parent=self.winfo_toplevel())
            return
        with self._conn() as con:
            ok = con.execute("SELECT 1 FROM tbl_bestellungen WHERE id=?", (bid,)).fetchone()
        if not ok:
            messagebox.showinfo("Auftrag anzeigen", _T('Kein Auftrag #{p0} gefunden.', p0=bid),
                                parent=self.winfo_toplevel())
            return
        self._verkauf_detail_window(bid)

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
            text=_T('{p0}  ·  DF {p1}  ·  PCK {p2}  ·  APU {p3}', p0=det['artikelname'], p1=det['df'] or '—', p2=det['pck'] or '—', p3=det['apu'] if det['apu'] is not None else '—'))
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

        liefertermin = self.vk_pos_termin_var.get().strip()

        def _mk(menge_, bestellart_):
            """Baut eine Positions-Zeile aus dem aktuell gewaehlten Artikel."""
            return {
                "pzn": self.vk_cur["pzn"], "artikelname": self.vk_cur["artikelname"],
                "df": self.vk_cur["df"], "pck": self.vk_cur["pck"], "apu": self.vk_cur["apu"],
                "menge": menge_, "rabatt": rabatt, "rabatt_quelle": self.vk_cur["rabatt_quelle"],
                "charge": charge, "verfall": verfall,
                "bestellart": bestellart_,
                "lieferzeit": lieferzeit,
                "liefertermin": liefertermin,
            }

        # #3: Bestand pruefen, wenn es eine echte Bestellung ist (nicht schon Vorbestellung).
        # Reicht der Bestand nicht: verfuegbare Menge sofort liefern, Rest auf Wunsch
        # als Vorbestellung; oder nur den verfuegbaren Bestand abverkaufen.
        neue_positionen = []
        if bestellart == "Bestellung":
            verfuegbar = self._bestand_for(self.vk_cur["pzn"], charge, verfall) if (charge or verfall) \
                else self._bestand_total(self.vk_cur["pzn"])
            verf = max(int(verfuegbar or 0), 0)
            if verf < menge:
                rest = menge - verf
                # Klammer-Bestand (unterwegs, noch nicht bei MSK bestaetigt):
                # NICHT verkaufbar, aber als Hinweis einblenden – der Rest ist
                # dann eine Vorbestellung, die durch eintreffende Ware gedeckt ist.
                unterwegs = self._unterwegs_total(self.vk_cur["pzn"])
                if unterwegs > 0:
                    gedeckt = min(unterwegs, rest)
                    uw_hinweis = _T('\n\nℹ {p0} St. unterwegs (in Klammern, noch nicht bei MSK bestätigt) – davon decken {p1} St. die Vorbestellung.', p0=unterwegs, p1=gedeckt)
                else:
                    uw_hinweis = ""
                antwort = messagebox.askyesnocancel(
                    "Bestand reicht nicht",
                    _T('Bestand reicht nicht (verfügbar {p0}, benötigt {p1}).\n\nJA:  {p2} sofort liefern und die restlichen {p3} als Vorbestellung aufnehmen.\nNEIN:  nur den verfügbaren Bestand ({p4}) abverkaufen.\nABBRECHEN:  nichts hinzufügen.', p0=verf, p1=menge, p2=verf, p3=rest, p4=verf) + uw_hinweis, parent=top)
                if antwort is None:
                    return
                if antwort:  # JA -> Split: Bestand als Bestellung + Rest als Vorbestellung
                    if verf > 0:
                        neue_positionen.append(_mk(verf, "Bestellung"))
                    neue_positionen.append(_mk(rest, "Vorbestellung"))
                else:        # NEIN -> nur den verfuegbaren Bestand
                    if verf <= 0:
                        messagebox.showinfo(
                            "Kein Bestand",
                            "Kein Bestand vorhanden – es wurde nichts hinzugefügt.\n"
                            "Tipp: bei „Liefervorgabe“ direkt „Vorbestellung“ wählen.",
                            parent=top)
                        return
                    neue_positionen.append(_mk(verf, "Bestellung"))
        if not neue_positionen:
            neue_positionen.append(_mk(menge, bestellart))

        self.vk_positions.extend(neue_positionen)
        # Eingabe zuruecksetzen
        self.vk_cur = None
        self.vk_artikel_search.clear()
        self.vk_detail_label.config(text="—")
        self.vk_charge_cmb.config(values=[])
        self.vk_charge_var.set("")
        self.vk_rabatt_var.set("")
        self.vk_menge_var.set("1")
        self._vk_render_positions()
        vb_neu = [q for q in neue_positionen if q["bestellart"] != "Bestellung"]
        if vb_neu:
            if len(neue_positionen) > 1:
                # Split: Bestellung + Vorbestellung aus einer Eingabe entstanden.
                hinweis = (f"Aufgeteilt: {neue_positionen[0]['menge']} als Bestellung, "
                           f"{vb_neu[0]['menge']} als Vorbestellung (gelb).\n"
                           "Die Vorbestellung wird NICHT berechnet und wandert beim "
                           "Speichern in die Vorbestellungen.")
            else:
                hinweis = ("Als Vorbestellung in der Liste markiert (gelb) – wird NICHT "
                           "berechnet und wandert beim Speichern in die Vorbestellungen.")
            messagebox.showinfo("Vorbestellung", hinweis, parent=self.winfo_toplevel())

    def _vk_render_positions(self):
        """Zeigt ALLE Positionen; Vorbestellungen/abgesagte werden farblich abgesetzt
        und zaehlen NICHT zur Summe (sie wandern beim Speichern in die Vorbestellungen).
        Map iid -> Position fuer Loeschen/Bearbeiten."""
        self.vk_pos_tree.delete(*self.vk_pos_tree.get_children())
        self._vk_tree_map = {}
        for pos in self.vk_positions:
            if pos.get("bestellart", "Bestellung") != "Bestellung":
                tags = ("vorbestellung",)
            elif pos.get("_frei"):
                tags = ("frei",)
            else:
                tags = ()
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
        self._vk_total_label.config(text=_T('Gesamt (netto): {p0}', p0=_eur(gesamt)))

    def _vk_edit_position(self, _e=None):
        sel = self.vk_pos_tree.selection()
        if not sel:
            return
        p = self._vk_tree_map.get(sel[0])
        if p is None:
            return
        if p.get("_frei"):
            self._vk_freie_position_dialog(edit=p)
        else:
            self._charge_dialog(p, sel[0])

    def _vk_freie_position_dialog(self, edit=None):
        """Freie (nicht bestandsgefuehrte) Position erfassen/bearbeiten: freie
        Bezeichnung + Einzelpreis + Menge + Rabatt. Wird gespeichert und erscheint
        mit Preis auf der Auftragsbestaetigung; bucht KEINEN Lagerbestand ab."""
        top = self.winfo_toplevel()
        win = tk.Toplevel(top)
        win.title("Freie Position")
        win.configure(bg=BG)
        win.transient(top)
        win.resizable(False, False)
        tk.Label(win, text="Freie Position", font=(theme.FONT, 13, "bold"),
                 fg=ACCENT, bg=BG).pack(padx=18, pady=(14, 2), anchor="w")
        tk.Label(win, text="Nicht bestandsgeführt – erscheint mit Preis auf der "
                 "Auftragsbestätigung.", font=(theme.FONT, 9), fg="#555",
                 bg=BG).pack(padx=18, anchor="w", pady=(0, 8))
        form = tk.Frame(win, bg=BG)
        form.pack(fill="x", padx=18)
        e = edit or {}
        tk.Label(form, text="Bezeichnung:", bg=BG, font=(theme.FONT, 9, "bold")).grid(
            row=0, column=0, sticky="w", pady=3)
        bez_var = tk.StringVar(value=e.get("artikelname", ""))
        tk.Entry(form, textvariable=bez_var, width=38).grid(row=0, column=1, sticky="w", pady=3, padx=(8, 0))
        tk.Label(form, text="Menge:", bg=BG, font=(theme.FONT, 9, "bold")).grid(
            row=1, column=0, sticky="w", pady=3)
        menge_var = tk.StringVar(value=str(e.get("menge", 1)))
        tk.Entry(form, textvariable=menge_var, width=8).grid(row=1, column=1, sticky="w", pady=3, padx=(8, 0))
        tk.Label(form, text="Einzelpreis € (APU):", bg=BG, font=(theme.FONT, 9, "bold")).grid(
            row=2, column=0, sticky="w", pady=3)
        apu0 = e.get("apu")
        preis_var = tk.StringVar(value=("" if apu0 is None else f"{apu0:.2f}".replace(".", ",")))
        tk.Entry(form, textvariable=preis_var, width=10).grid(row=2, column=1, sticky="w", pady=3, padx=(8, 0))
        tk.Label(form, text="Rabatt %:", bg=BG, font=(theme.FONT, 9, "bold")).grid(
            row=3, column=0, sticky="w", pady=3)
        rab_var = tk.StringVar(value=str(int(e.get("rabatt", 0) or 0)))
        tk.Entry(form, textvariable=rab_var, width=8).grid(row=3, column=1, sticky="w", pady=3, padx=(8, 0))

        def uebernehmen():
            bez = bez_var.get().strip()
            if not bez:
                messagebox.showwarning("Freie Position", "Bitte eine Bezeichnung eingeben.", parent=win)
                return
            try:
                menge = int(menge_var.get().strip())
                if menge <= 0:
                    raise ValueError
            except ValueError:
                messagebox.showwarning("Freie Position", "Menge muss eine positive Zahl sein.", parent=win)
                return
            try:
                preis = float(preis_var.get().strip().replace(",", ".")) if preis_var.get().strip() else 0.0
            except ValueError:
                messagebox.showwarning("Freie Position", "Einzelpreis ist keine gültige Zahl.", parent=win)
                return
            try:
                rabatt = float(rab_var.get().strip().replace(",", ".") or 0)
            except ValueError:
                rabatt = 0.0
            if edit is not None:
                edit.update({"artikelname": bez, "menge": menge, "apu": preis, "rabatt": rabatt})
            else:
                self.vk_positions.append({
                    "pzn": "", "artikelname": bez, "df": "", "pck": "", "apu": preis,
                    "menge": menge, "rabatt": rabatt, "rabatt_quelle": "frei",
                    "charge": "", "verfall": "",
                    "bestellart": "Bestellung", "lieferzeit": "", "liefertermin": "",
                    "_frei": True,
                })
            self._vk_render_positions()
            win.destroy()

        btns = tk.Frame(win, bg=BG)
        btns.pack(padx=18, pady=(12, 14), anchor="e")
        tk.Button(btns, text="Übernehmen", command=uebernehmen, bg=ACCENT, fg="white",
                  font=(theme.FONT, 10, "bold"), padx=12, pady=4).pack(side="right", padx=(8, 0))
        tk.Button(btns, text="Abbrechen", command=win.destroy, padx=12, pady=4).pack(side="right")
        self._restyle_buttons(win)
        win.lift()
        win.focus_force()

    def _charge_dialog(self, p, item):
        top = self.winfo_toplevel()
        win = tk.Toplevel(top)
        win.title("Position bearbeiten")
        win.configure(bg=BG)
        win.transient(top)
        win.resizable(False, False)

        tk.Label(win, text=p["artikelname"] or p["pzn"], font=(theme.FONT, 13, "bold"),
                 fg=ACCENT, bg=BG).pack(padx=18, pady=(14, 2), anchor="w")
        tk.Label(win, text=_T('PZN {p0}  ·  DF {p1}  ·  PCK {p2}  ·  APU {p3}', p0=p['pzn'], p1=p['df'] or '—', p2=p['pck'] or '—', p3=_eur(p['apu'])), font=(theme.FONT, 9), fg="#555",
                 bg=BG).pack(padx=18, anchor="w")
        tk.Label(win, text=_T('Aktuell gewählt: Charge {p0}  ·  Verfall {p1}', p0=p['charge'] or '—', p1=p['verfall'] or '—'),
                 font=(theme.FONT, 9, "italic"), fg="#888", bg=BG).pack(padx=18, pady=(2, 8), anchor="w")

        tk.Label(win, text="Verfügbare Chargen dieses Artikels (Doppelklick = übernehmen):",
                 font=(theme.FONT, 9, "bold"), bg=BG).pack(padx=18, anchor="w")
        tf = tk.Frame(win, bg=BG)
        tf.pack(fill="both", padx=18, pady=(4, 8))
        cols = ("charge", "verfall", "bestand")
        tree = ttk.Treeview(tf, columns=cols, show="headings", height=6,
                            selectmode="browse", style="Kasse.Treeview")
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
        tk.Label(mrow, text="Menge:", font=(theme.FONT, 9, "bold"), bg=BG).pack(side="left")
        menge_var = tk.StringVar(value=str(p["menge"]))
        tk.Entry(mrow, textvariable=menge_var, width=6).pack(side="left", padx=(6, 0))

        lrow = tk.Frame(win, bg=BG)
        lrow.pack(fill="x", padx=18, pady=(0, 8))
        tk.Label(lrow, text="Liefervorgabe:", font=(theme.FONT, 9, "bold"), bg=BG).pack(side="left")
        art_var = tk.StringVar(value=p.get("bestellart", "Bestellung"))
        ttk.Combobox(lrow, textvariable=art_var, width=13, state="readonly",
                     values=("Bestellung", "Vorbestellung", "abgesagt")).pack(side="left", padx=(4, 8))
        lz_var = tk.StringVar(value=p.get("lieferzeit", "") or "10 Uhr")
        ttk.Combobox(lrow, textvariable=lz_var, width=8,
                     values=("10 Uhr", "12 Uhr", "Frei")).pack(side="left", padx=(0, 8))
        tk.Label(lrow, text="Termin:", font=(theme.FONT, 9, "bold"), bg=BG).pack(side="left")
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
                  font=(theme.FONT, 10, "bold"), padx=12, pady=4).pack(side="right", padx=(8, 0))
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
        # Bearbeitung: zugrunde liegenden Original-Auftrag VOR dem Speichern
        # zurückbuchen + entfernen (der korrigierte Auftrag wird neu angelegt).
        for old in {p["_edit_src"] for p in self.vk_positions if p.get("_edit_src")}:
            self._verkauf_original_ersetzen(old)
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
        # Online-Übernahme: die zugrunde liegenden Online-Aufträge am Server als
        # übernommen markieren (fallen dann aus den Online-Listen). An die
        # Positionen gebunden -> nur was wirklich gespeichert wurde.
        for aid in {p["_online_src"] for p in self.vk_positions if p.get("_online_src")}:
            self._online_mark_uebernommen(aid)
        # Bestand nach dem Verkauf sofort zum Server pushen (frische Reservierung).
        self._online_push_lager()
        # Reset
        self.vk_positions = []
        self._vk_render_positions()
        self.vk_kunde = None
        self.vk_kunde_search.clear()
        self._render_kunde_card(None)
        self._refresh_artikel()
        if hasattr(self, "vb_tree"):
            self._refresh_vorbestellungen()
        if hasattr(self, "vh_tree"):
            self._refresh_verkaeufe()
        # #1: Auftragsbestaetigung nur, wenn echte Bestellungen dabei sind.
        if bestellungen:
            self._auftrag_dialog(bid, anzahl, vorbestellungen)
        else:
            messagebox.showinfo("Vorbestellung",
                                _T('{p0} Vorbestellung(en) gespeichert (keine Auftragsbestätigung).', p0=vorbestellungen), parent=top)

    # ==================================================== Auftragsbestaetigung
    def _auftrag_dialog(self, bestell_id, anzahl, vorbestellungen=0):
        top = self.winfo_toplevel()
        win = tk.Toplevel(top)
        win.title("Auftragsbestätigung")
        win.configure(bg=BG)
        win.transient(top)
        win.resizable(False, False)
        tk.Label(win, text=_T('✓ Verkauf #{p0} gespeichert', p0=bestell_id),
                 font=(theme.FONT, 12, "bold"), fg="#11823b", bg=BG).pack(padx=20, pady=(16, 2))
        info = f"{anzahl} Position(en) · Lagerbestand abgebucht."
        if vorbestellungen:
            info += (f"\n{vorbestellungen} davon Vorbestellung(en) – noch nicht abgebucht, "
                     "im Reiter „Vorbestellungen“ bestätigen, wenn Ware da ist.")
        tk.Label(win, text=info, font=(theme.FONT, 9), fg="#666", bg=BG,
                 justify="left").pack(padx=20, pady=(0, 12))

        btns = tk.Frame(win, bg=BG)
        btns.pack(padx=20, pady=(0, 8))
        tk.Button(btns, text="🖨  Drucken / Vorschau", width=22,
                  command=lambda: self._auftrag_drucken(bestell_id, win),
                  bg=ACCENT, fg="white", font=(theme.FONT, 10, "bold"), pady=4).grid(row=0, column=0, padx=4, pady=3)
        tk.Button(btns, text="📧  Per E-Mail senden", width=22,
                  command=lambda: self._auftrag_mail(bestell_id, win),
                  bg="#3867b7", fg="white", font=(theme.FONT, 10, "bold"), pady=4).grid(row=1, column=0, padx=4, pady=3)
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
            messagebox.showerror("Auftragsbestätigung", _T('Konnte nicht erstellt werden:\n{p0}', p0=e), parent=parent)

    def _lieferschein_drucken(self, bestell_id, parent):
        # Vor dem Erzeugen abfragen, ob der Auftrag in MSK erfasst wurde.
        with self._conn() as con:
            row = con.execute("SELECT COALESCE(msk_erfasst,0) FROM tbl_bestellungen WHERE id=?",
                              (bestell_id,)).fetchone()
        msk = bool(row[0]) if row else False
        if not msk:
            antwort = messagebox.askyesnocancel(
                "MSK-Erfassung",
                _T('Wurde Auftrag #{p0} in MSK erfasst?\n\nJa:  als „in MSK erfasst“ markieren und Lieferschein erzeugen.\nNein:  Lieferschein trotzdem erzeugen (ohne MSK-Markierung).\nAbbrechen:  nichts tun.', p0=bestell_id), parent=parent)
            if antwort is None:
                return
            if antwort:
                # MSK setzen, aber den Lieferschein erzeugen wir gleich hier selbst.
                self._msk_markieren(True, bids=[bestell_id], mit_lieferschein=False)
        try:
            path = lieferschein.render(self.db_path, bestell_id)
            self._log("Lieferschein erzeugt", bestell_id, details=str(path.name))
            os.startfile(str(path))  # type: ignore[attr-defined]
        except Exception as e:
            messagebox.showerror("Lieferschein", _T('Konnte nicht erstellt werden:\n{p0}', p0=e), parent=parent)

    def _auftrag_verlauf_window(self, bid):
        """Audit-Verlauf eines Auftrags: wer hat was wann gemacht (aus tbl_kasse_log)."""
        top = self.winfo_toplevel()
        kopf, eintraege = kasse_reports.auftrag_historie(self.db_path, bid)
        win = tk.Toplevel(top)
        win.title(f"Verlauf · Auftrag #{bid}")
        win.configure(bg=BG)
        win.transient(top)
        win.geometry("740x430")
        apotheke = kopf[2] if kopf else ""
        tk.Label(win, text=_T('Verlauf · Auftrag #{p0} · {p1}', p0=bid, p1=apotheke),
                 font=(theme.FONT, 13, "bold"), fg=ACCENT, bg=BG).pack(padx=18, pady=(14, 2), anchor="w")
        tk.Label(win, text="Wer hat was wann mit diesem Auftrag gemacht.",
                 font=(theme.FONT, 9), fg="#555", bg=BG).pack(padx=18, anchor="w", pady=(0, 8))
        tf = tk.Frame(win, bg=BG)
        tf.pack(fill="both", expand=True, padx=18, pady=(0, 8))
        cols = ("zeit", "bearbeiter", "aktion", "details")
        heads = {"zeit": "Zeitpunkt", "bearbeiter": "Mitarbeiter", "aktion": "Aktion",
                 "details": "Details"}
        tree = ttk.Treeview(tf, columns=cols, show="headings", style="Kasse.Treeview")
        for c in cols:
            tree.heading(c, text=heads[c])
            tree.column(c, width=300 if c == "details" else (140 if c == "zeit" else 120),
                        anchor="w")
        tree.pack(side="left", fill="both", expand=True)
        sb = tk.Scrollbar(tf, orient="vertical", command=tree.yview)
        sb.pack(side="right", fill="y")
        tree.configure(yscrollcommand=sb.set)
        _make_treeview_sortable(tree)
        for z, b, a, d in eintraege:
            tree.insert("", "end", values=(str(z).replace("T", " ")[:16], b, a, d))
        if not eintraege:
            tree.insert("", "end", values=("—", "", "keine Protokoll-Einträge", ""))
        btns = tk.Frame(win, bg=BG)
        btns.pack(fill="x", padx=18, pady=(0, 14))
        tk.Button(btns, text="📄 Als PDF",
                  command=lambda: self._run_report(
                      lambda: kasse_reports.auftrag_historie_pdf(self.db_path, bid)),
                  bg=ACCENT, fg="white", font=(theme.FONT, 10, "bold"),
                  padx=12, pady=4).pack(side="left")
        tk.Button(btns, text="Schließen", command=win.destroy, padx=12, pady=4).pack(side="right")
        self._restyle_buttons(win)
        win.lift()
        win.focus_force()

    def _auftrag_mail(self, bestell_id, parent):
        try:
            path = auftrag.render(self.db_path, bestell_id)
        except Exception as e:
            messagebox.showerror("Auftragsbestätigung", _T('Konnte nicht erstellt werden:\n{p0}', p0=e), parent=parent)
            return
        to = auftrag.kunde_email(self.db_path, bestell_id)
        if not to:
            to = simpledialog.askstring("E-Mail", "Keine E-Mail beim Kunden hinterlegt.\n"
                                        "Empfänger-Adresse eingeben:", parent=parent) or ""
        try:
            html = path.read_text(encoding="utf-8")
            auftrag.send_via_outlook(to.strip(), f"Auftragsbestätigung #{bestell_id}", html, path)
        except Exception as e:
            messagebox.showerror("E-Mail", _T('Outlook konnte nicht geöffnet werden:\n{p0}', p0=e), parent=parent)

    def _auftrag_vorlage_oeffnen(self):
        try:
            os.startfile(str(auftrag.template_path()))  # type: ignore[attr-defined]
        except Exception as e:
            messagebox.showerror("Vorlage", _T('Vorlage konnte nicht geöffnet werden:\n{p0}', p0=e),
                                 parent=self.winfo_toplevel())

    # ========================================================== Vorbestellungen
    def _build_vorbestellungen(self, parent):
        bar = tk.Frame(parent, bg=BG)
        bar.pack(fill="x", padx=8, pady=(10, 4))
        self.vb_info = tk.Label(bar, text="", bg=BG, fg=ACCENT, font=(theme.FONT, 11, "bold"))
        self.vb_info.pack(side="left")
        tk.Button(bar, text="➡ Mehrere in Verkauf übernehmen", command=self._vb_bulk_dialog,
                  bg="#11823b", fg="white", font=(theme.FONT, 9, "bold"), padx=10, pady=2).pack(
                      side="left", padx=(16, 0))
        tk.Button(bar, text="🌐 Online laden", command=self._online_verkaeufe_laden,
                  bg="#0B4A86", fg="white", font=(theme.FONT, 9, "bold"), padx=10, pady=2).pack(
                      side="left", padx=(8, 0))
        tk.Button(bar, text="🔄 Aktualisieren", command=self._refresh_vorbestellungen,
                  font=(theme.FONT, 9), padx=8, pady=2).pack(side="right")

        srow = tk.Frame(parent, bg=BG)
        srow.pack(fill="x", padx=8, pady=(0, 4))
        tk.Label(srow, text="Suche (Kunde / Artikel / PZN):", bg=BG,
                 font=(theme.FONT, 9, "bold")).pack(side="left")
        self._vb_filter_var = tk.StringVar()
        e = tk.Entry(srow, textvariable=self._vb_filter_var, width=26)
        e.pack(side="left", padx=(6, 8))
        e.bind("<KeyRelease>", lambda _e: self._refresh_vorbestellungen())
        tk.Button(srow, text="✕", command=lambda: (self._vb_filter_var.set(""),
                  self._refresh_vorbestellungen()), font=(theme.FONT, 8), padx=6).pack(side="left")
        tk.Label(srow, text="Status:", bg=BG, font=(theme.FONT, 9, "bold")).pack(side="left", padx=(12, 0))
        self._vb_status_var = tk.StringVar(value="Offen")
        ttk.Combobox(srow, textvariable=self._vb_status_var, width=11, state="readonly",
                     values=("Offen", "Abgesagt", "Alle")).pack(side="left", padx=(4, 0))
        self._vb_status_var.trace_add("write", lambda *_: self._refresh_vorbestellungen())
        tk.Button(srow, text="📥 Vorbestellungen importieren", command=self._import_vorbestellungen_datei,
                  font=(theme.FONT, 8), padx=6, pady=1).pack(side="right")
        tk.Label(parent, text="Disposition: Kunden anrufen, dann Doppelklick → bearbeiten, "
                              "„Als Verkauf bestätigen“ oder stornieren.",
                 bg=BG, fg="#666", font=(theme.FONT, 9)).pack(anchor="w", padx=8, pady=(0, 4))

        tf = tk.Frame(parent, bg=BG)
        tf.pack(fill="both", expand=True, padx=8, pady=(0, 8))
        cols = ("datum", "kunde", "telefon", "pzn", "artikel", "menge", "lieferzeit", "termin")
        heads = {"datum": "Datum", "kunde": "Kunde", "telefon": "Telefon", "pzn": "PZN",
                 "artikel": "Artikel", "menge": "Menge", "lieferzeit": "Lieferzeit", "termin": "Termin"}
        tree = ttk.Treeview(tf, columns=cols, show="headings", style="Kasse.Treeview")
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
        tree.tag_configure("online", background="#eef4fb")
        _make_treeview_sortable(tree)
        tree.bind("<Double-1>", self._vorbestellung_dialog)
        self.vb_tree = tree
        self._vb_rowmap = {}
        self._vb_online_map = {}

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
        # Online-Vorbestellungen (vom Server) zusätzlich einblenden – read-only,
        # blau hinterlegt, „(online)". Kein Eintrag in _vb_rowmap -> Doppelklick
        # öffnet keinen lokalen Bearbeiten-Dialog.
        online_anz = 0
        if status_f != "Abgesagt":
            for v in getattr(self, "_online_vorbestellungen", []):
                kunde = v.get("kunde_name") or ""
                pzn = v.get("pzn") or ""
                artikel = v.get("bezeichnung") or ""
                if filt and filt not in " ".join(str(x).lower() for x in (kunde, pzn, artikel)):
                    continue
                iid = self.vb_tree.insert(
                    "", "end", tags=("online",),
                    values=(v.get("datum", ""), f"{kunde} (online)", "", pzn, artikel,
                            v.get("menge", 0), "", v.get("liefertermin") or ""))
                self._vb_online_map[iid] = v.get("auftrag")
                online_anz += 1
                gezeigt += 1
        suffix = f" (gefiltert aus {len(rows)})" if filt else ""
        zusatz = f" · 🌐 {online_anz} online" if online_anz else ""
        label = {"Abgesagt": "abgesagte", "Alle": "Vorbestellung(en) gesamt"}.get(status_f, "offene")
        if status_f == "Alle":
            self.vb_info.config(text=f"{gezeigt} {label}{suffix}{zusatz}")
        else:
            self.vb_info.config(text=_T('{p0} {p1} Vorbestellung(en){p2}', p0=gezeigt, p1=label, p2=suffix) + zusatz)

    def _vorbestellung_dialog(self, _e=None):
        sel = self.vb_tree.selection()
        if not sel:
            return
        if sel[0] in getattr(self, "_vb_online_map", {}):
            aid = self._vb_online_map[sel[0]]
            if aid:
                self._online_detail_window(aid)
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
        tk.Label(win, text=artikel or pzn, font=(theme.FONT, 13, "bold"), fg=ACCENT,
                 bg=BG).pack(padx=18, pady=(14, 2), anchor="w")
        tk.Label(win, text=_T('PZN {p0}  ·  DF {p1}  ·  PCK {p2}  ·  APU {p3}', p0=pzn, p1=df or '—', p2=pck or '—', p3=_eur(apu)),
                 font=(theme.FONT, 9), fg="#555", bg=BG).pack(padx=18, anchor="w")
        tk.Label(win, text=_T('Kunde: {p0}', p0=apotheke or '—'), font=(theme.FONT, 9), fg="#555",
                 bg=BG).pack(padx=18, pady=(0, 8), anchor="w")

        erow = tk.Frame(win, bg=BG)
        erow.pack(fill="x", padx=18, pady=(0, 6))
        tk.Label(erow, text="Menge:", font=(theme.FONT, 9, "bold"), bg=BG).pack(side="left")
        menge_var = tk.StringVar(value=str(menge))
        tk.Entry(erow, textvariable=menge_var, width=6).pack(side="left", padx=(4, 12))
        lz_var = tk.StringVar(value=lieferzeit or "10 Uhr")
        ttk.Combobox(erow, textvariable=lz_var, width=8,
                     values=("10 Uhr", "12 Uhr", "Frei")).pack(side="left", padx=(0, 8))
        tk.Label(erow, text="Termin:", font=(theme.FONT, 9, "bold"), bg=BG).pack(side="left")
        termin_var = tk.StringVar(value=termin)
        tk.Entry(erow, textvariable=termin_var, width=11).pack(side="left", padx=(4, 0))

        tk.Label(win, text="Charge wählen (Doppelklick oder markieren – wird beim Übernehmen "
                          "in den Verkauf mitgenommen):",
                 font=(theme.FONT, 9, "bold"), bg=BG).pack(padx=18, anchor="w")
        tf = tk.Frame(win, bg=BG)
        tf.pack(fill="both", padx=18, pady=(4, 8))
        ctree = ttk.Treeview(tf, columns=("charge", "verfall", "bestand"), show="headings",
                             height=4, selectmode="browse", style="Kasse.Treeview")
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
                    _T('Vorbestellung „{p0}“ für {p1} wirklich absagen?', p0=artikel or pzn, p1=apotheke or 'diesen Kunden'), parent=win):
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
                  fg="white", font=(theme.FONT, 10, "bold"), padx=12, pady=4).pack(side="left")
        tk.Button(btns, text="Speichern", command=speichern, bg=ACCENT, fg="white",
                  font=(theme.FONT, 10, "bold"), padx=12, pady=4).pack(side="left", padx=(8, 0))
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
                _T('Im Verkauf liegt bereits ein anderer Kunde ({p0}).\n\nBitte den laufenden Auftrag erst löschen oder abschließen und dann die Vorbestellung holen.', p0=(self.vk_kunde or {}).get('kundenname', 'unbekannt')), parent=self.winfo_toplevel())
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

    def _vb_bulk_dialog(self):
        """Sammelfenster: alle offenen Vorbestellungen mit Bestand; je Zeile eine
        Übernehmen-Menge (Doppelklick = ändern, 0 = überspringen). „Übernehmen“ lädt
        alle Zeilen mit Menge>0 in den Verkauf (nur EIN Kunde gleichzeitig)."""
        top = self.winfo_toplevel()
        with self._conn() as con:
            rows = con.execute(
                "SELECT p.id, COALESCE(b.kundennummer,''), COALESCE(b.apotheke,''), p.pzn, "
                "p.artikelname, p.menge, "
                "(SELECT COALESCE(SUM(menge),0) FROM tbl_lagerbestand l WHERE l.pzn=p.pzn) "
                "FROM tbl_bestellpositionen p JOIN tbl_bestellungen b ON b.id=p.bestell_id "
                "WHERE p.bestellart='Vorbestellung' ORDER BY b.apotheke, p.liefertermin, b.datum"
            ).fetchall()
        if not rows:
            messagebox.showinfo("Vorbestellungen", "Keine offenen Vorbestellungen vorhanden.", parent=top)
            return

        win = tk.Toplevel(top)
        win.title("Vorbestellungen übernehmen")
        win.configure(bg=BG)
        win.transient(top)
        win.geometry("820x520")
        tk.Label(win, text="Vorbestellungen in den Verkauf übernehmen", font=(theme.FONT, 13, "bold"),
                 fg=ACCENT, bg=BG).pack(padx=18, pady=(14, 2), anchor="w")
        tk.Label(win, text="Übernehmen-Menge je Zeile per Doppelklick ändern (0 = überspringen). "
                          "Es kann immer nur EIN Kunde gleichzeitig übernommen werden.",
                 font=(theme.FONT, 9), fg="#555", bg=BG).pack(padx=18, anchor="w")

        srow = tk.Frame(win, bg=BG)
        srow.pack(fill="x", padx=18, pady=(6, 4))
        tk.Label(srow, text="Filter (Kunde / Artikel / PZN):", bg=BG, font=(theme.FONT, 9, "bold")).pack(side="left")
        filt_var = tk.StringVar()
        tk.Entry(srow, textvariable=filt_var, width=26).pack(side="left", padx=(6, 8))

        tf = tk.Frame(win, bg=BG)
        tf.pack(fill="both", expand=True, padx=18, pady=(0, 6))
        cols = ("kunde", "pzn", "artikel", "offen", "bestand", "uebernehmen")
        heads = {"kunde": "Kunde", "pzn": "PZN", "artikel": "Artikel", "offen": "Vorbestellt",
                 "bestand": "Bestand", "uebernehmen": "Übernehmen"}
        tree = ttk.Treeview(tf, columns=cols, show="headings", style="Kasse.Treeview")
        for c in cols:
            tree.heading(c, text=heads[c])
            w = 200 if c == "artikel" else (150 if c == "kunde" else 90)
            tree.column(c, width=w, anchor="w")
        tree.tag_configure("skip", foreground="#999")
        tree.pack(side="left", fill="both", expand=True)
        sb = tk.Scrollbar(tf, orient="vertical", command=tree.yview)
        sb.pack(side="right", fill="y")
        tree.configure(yscrollcommand=sb.set)
        _make_treeview_sortable(tree)

        # state: iid -> {pid, knr, kunde, pzn, artikel, offen, bestand, take}
        state = {}
        for pid, knr, apotheke, pzn, artikel, menge, bestand in rows:
            state[pid] = {"pid": pid, "knr": knr, "kunde": apotheke, "pzn": pzn,
                          "artikel": artikel, "offen": int(menge or 0),
                          "bestand": int(bestand or 0), "take": int(menge or 0)}

        def _fill():
            tree.delete(*tree.get_children())
            f = filt_var.get().strip().lower()
            for pid, d in state.items():
                if f and f not in " ".join(str(x).lower() for x in (d["kunde"], d["pzn"], d["artikel"])):
                    continue
                tags = () if d["take"] > 0 else ("skip",)
                tree.insert("", "end", iid=str(pid), tags=tags,
                            values=(d["kunde"], d["pzn"], d["artikel"], d["offen"],
                                    d["bestand"], d["take"]))

        filt_var.trace_add("write", lambda *_: _fill())

        def _edit_take(_e=None):
            sel = tree.selection()
            if not sel:
                return
            d = state.get(int(sel[0]))
            if not d:
                return
            neu = simpledialog.askinteger(
                "Übernehmen-Menge",
                _T('{p0}\nVorbestellt {p1} · Bestand {p2}\n\nWie viel übernehmen? (0 = überspringen)', p0=d['artikel'] or d['pzn'], p1=d['offen'], p2=d['bestand']),
                parent=win, initialvalue=d["take"], minvalue=0, maxvalue=d["offen"])
            if neu is None:
                return
            d["take"] = neu
            _fill()
        tree.bind("<Double-1>", _edit_take)

        def _set_all(full):
            for d in state.values():
                d["take"] = d["offen"] if full else 0
            _fill()

        def _uebernehmen():
            gewaehlt = [(d["pid"], d["take"], d["knr"], d["kunde"])
                        for d in state.values() if d["take"] > 0]
            if not gewaehlt:
                messagebox.showinfo("Übernehmen", "Keine Zeile mit Menge > 0 gewählt.", parent=win)
                return
            knrs = {g[2] for g in gewaehlt}
            if len(knrs) > 1:
                kunden = ", ".join(sorted({g[3] or g[2] for g in gewaehlt}))
                messagebox.showwarning(
                    "Mehrere Kunden",
                    _T('Es können nur Vorbestellungen EINES Kunden gleichzeitig übernommen werden.\n\nAktuell gewählt: {p0}\n\nBitte über den Filter auf einen Kunden eingrenzen.', p0=kunden),
                    parent=win)
                return
            if self._vb_load_many([(g[0], g[1]) for g in gewaehlt]):
                win.destroy()

        btns = tk.Frame(win, bg=BG)
        btns.pack(fill="x", padx=18, pady=(0, 14))
        tk.Button(btns, text="➡ Übernehmen", command=_uebernehmen, bg="#11823b", fg="white",
                  font=(theme.FONT, 10, "bold"), padx=12, pady=4).pack(side="left")
        tk.Button(btns, text="Alle = volle Menge", command=lambda: _set_all(True),
                  padx=10, pady=4).pack(side="left", padx=(8, 0))
        tk.Button(btns, text="Alle = 0", command=lambda: _set_all(False),
                  padx=10, pady=4).pack(side="left", padx=(6, 0))
        tk.Button(btns, text="Abbrechen", command=win.destroy, padx=12, pady=4).pack(side="right")
        _fill()
        self._restyle_buttons(win)
        win.lift()
        win.focus_force()

    def _vb_load_many(self, pairs):
        """Laedt mehrere Vorbestellungen (Liste (pid, menge)) in den Verkauf-Reiter.
        Alle muessen zum selben Kunden gehoeren. Originale werden erst beim Speichern
        des Verkaufs reduziert/geloescht (_vb_source). True = uebernommen."""
        top = self.winfo_toplevel()
        geladen = []
        with self._conn() as con:
            for pid, menge in pairs:
                row = con.execute(
                    "SELECT p.pzn, p.artikelname, p.df, p.pck, p.apu, p.rabatt_prozent, "
                    "p.rabatt_quelle, b.kundennummer, COALESCE(b.apotheke,''), p.menge, "
                    "COALESCE(p.lieferzeit,''), COALESCE(p.liefertermin,'') "
                    "FROM tbl_bestellpositionen p JOIN tbl_bestellungen b ON b.id=p.bestell_id "
                    "WHERE p.id=?", (pid,)).fetchone()
                if row:
                    geladen.append((pid, menge, row))
        if not geladen:
            return False
        knrs = {g[2][7] for g in geladen}
        if len(knrs) > 1:
            messagebox.showwarning("Mehrere Kunden",
                                   "Nur Vorbestellungen eines Kunden können gemeinsam übernommen werden.",
                                   parent=top)
            return False
        knr = next(iter(knrs))
        # Konflikt mit bereits laufendem Verkauf (anderer Kunde)?
        if (self.vk_kunde and self.vk_kunde.get("kundennummer") != knr) or \
           (self.vk_positions and not self.vk_kunde):
            messagebox.showwarning(
                "Unterschiedlicher Kunde im Verkauf",
                _T('Im Verkauf liegt bereits ein anderer Kunde ({p0}).\n\nBitte den laufenden Auftrag erst löschen oder abschließen.', p0=(self.vk_kunde or {}).get('kundenname', 'unbekannt')), parent=top)
            return False
        apotheke = geladen[0][2][8]
        det = self._kunde_details(knr)
        self.vk_kunde = {
            "kundennummer": knr, "kundenname": det.get("kundenname") or apotheke,
            "plz": det.get("plz", ""), "ort": det.get("ort", ""), "strasse": det.get("strasse", ""),
            "inhaber": det.get("inhaber", ""), "telefon": det.get("telefon", ""),
            "email": det.get("email", ""),
        }
        self._render_kunde_card(self.vk_kunde)
        for pid, menge, row in geladen:
            pzn, artikel, df, pck, apu, rabatt, rquelle, _knr, _apo, orig_menge, lz, termin = row
            self.vk_positions.append({
                "pzn": pzn, "artikelname": artikel, "df": df or "", "pck": pck or "", "apu": apu,
                "menge": menge, "rabatt": rabatt or 0, "rabatt_quelle": rquelle or "manuell",
                "charge": "", "verfall": "", "bestellart": "Bestellung",
                "lieferzeit": lz, "liefertermin": termin,
                "_vb_source": pid, "_vb_orig_menge": orig_menge or 0,
            })
        self._vk_render_positions()
        self._show_view("verkauf")
        messagebox.showinfo("Vorbestellungen",
                            _T('{p0} Vorbestellung(en) in den Verkauf übernommen. Charge/Bestand je Position wählen, dann „Verkauf speichern“.', p0=len(geladen)), parent=top)
        return True

    # =============================================================== Verkaeufe
    def _build_verkaeufe(self, parent):
        bar = tk.Frame(parent, bg=BG)
        bar.pack(fill="x", padx=8, pady=(10, 4))
        tk.Label(bar, text="Suche (Kunde / Auftrag-Nr):", bg=BG, font=(theme.FONT, 9, "bold")).pack(side="left")
        self._vh_filter_var = tk.StringVar()
        e = tk.Entry(bar, textvariable=self._vh_filter_var, width=24)
        e.pack(side="left", padx=(6, 8))
        e.bind("<KeyRelease>", lambda _e: self._refresh_verkaeufe())
        e.bind("<Return>", lambda _e: self._verkaeufe_open_nr())
        tk.Button(bar, text="Auftrag öffnen", command=self._verkaeufe_open_nr,
                  font=(theme.FONT, 8), padx=6, pady=1).pack(side="left", padx=(0, 8))
        tk.Label(bar, text="Status:", bg=BG, font=(theme.FONT, 9, "bold")).pack(side="left")
        self._vh_status_var = tk.StringVar(value="Alle")
        ttk.Combobox(bar, textvariable=self._vh_status_var, width=11, state="readonly",
                     values=("Alle", "Aktiv", "Storniert")).pack(side="left", padx=(4, 0))
        self._vh_status_var.trace_add("write", lambda *_: self._refresh_verkaeufe())
        tk.Button(bar, text="📥 Verkäufe importieren", command=self._import_verkaeufe_datei,
                  font=(theme.FONT, 8), padx=6, pady=1).pack(side="left", padx=(12, 0))
        tk.Button(bar, text="🌐 Online laden", command=self._online_verkaeufe_laden,
                  bg="#0B4A86", fg="white", font=(theme.FONT, 9, "bold"),
                  padx=10, pady=2).pack(side="left", padx=(8, 0))
        tk.Button(bar, text="🔄 Aktualisieren", command=self._refresh_verkaeufe,
                  font=(theme.FONT, 9), padx=8, pady=2).pack(side="right")
        self.vh_info = tk.Label(bar, text="", bg=BG, fg=ACCENT, font=(theme.FONT, 9, "bold"))
        self.vh_info.pack(side="right", padx=(0, 12))

        # Zeitraum-Filter
        zrow = tk.Frame(parent, bg=BG)
        zrow.pack(fill="x", padx=8, pady=(0, 4))
        tk.Label(zrow, text="Zeitraum:", bg=BG, font=(theme.FONT, 9, "bold")).pack(side="left")
        self._vh_zeit_var = tk.StringVar(value="Alle")
        ttk.Combobox(zrow, textvariable=self._vh_zeit_var, width=16, state="readonly",
                     values=("Alle", "Aktueller Monat", "Vormonat", "Aktuelles Quartal",
                             "Aktuelles Jahr", "Frei")).pack(side="left", padx=(4, 12))
        self._vh_zeit_var.trace_add("write", lambda *_: self._refresh_verkaeufe())
        tk.Label(zrow, text="von:", bg=BG, font=(theme.FONT, 9)).pack(side="left")
        self._vh_von_var = tk.StringVar()
        ev = tk.Entry(zrow, textvariable=self._vh_von_var, width=11)
        ev.pack(side="left", padx=(3, 8))
        tk.Label(zrow, text="bis:", bg=BG, font=(theme.FONT, 9)).pack(side="left")
        self._vh_bis_var = tk.StringVar()
        eb = tk.Entry(zrow, textvariable=self._vh_bis_var, width=11)
        eb.pack(side="left", padx=(3, 6))
        tk.Label(zrow, text="(TT.MM.JJJJ – nur bei „Frei“)", bg=BG, fg="#999",
                 font=(theme.FONT, 8)).pack(side="left")
        ev.bind("<KeyRelease>", lambda _e: self._vh_zeit_var.get() == "Frei" and self._refresh_verkaeufe())
        eb.bind("<KeyRelease>", lambda _e: self._vh_zeit_var.get() == "Frei" and self._refresh_verkaeufe())
        tk.Label(zrow, text="MSK:", bg=BG, font=(theme.FONT, 9, "bold")).pack(side="left", padx=(14, 0))
        self._vh_msk_var = tk.StringVar(value="Alle")
        ttk.Combobox(zrow, textvariable=self._vh_msk_var, width=14, state="readonly",
                     values=("Alle", "MSK offen", "MSK erfasst")).pack(side="left", padx=(4, 0))
        self._vh_msk_var.trace_add("write", lambda *_: self._refresh_verkaeufe())

        # MSK-Schnellmarkierung der ausgewaehlten Zeilen.
        mrow = tk.Frame(parent, bg=BG)
        mrow.pack(fill="x", padx=8, pady=(0, 2))
        tk.Label(mrow, text="Markierte Verkäufe:", bg=BG, font=(theme.FONT, 9, "bold")).pack(side="left")
        tk.Button(mrow, text="✓ In MSK erfasst", command=lambda: self._msk_markieren(True),
                  bg="#11823b", fg="white", font=(theme.FONT, 9, "bold"), padx=8, pady=2).pack(side="left", padx=(6, 4))
        tk.Button(mrow, text="↺ MSK offen", command=lambda: self._msk_markieren(False),
                  font=(theme.FONT, 9), padx=8, pady=2).pack(side="left")
        tk.Label(mrow, text="(mehrere mit Strg/Shift markierbar · Doppelklick = Details)",
                 bg=BG, fg="#999", font=(theme.FONT, 8)).pack(side="left", padx=(10, 0))

        tf = tk.Frame(parent, bg=BG)
        tf.pack(fill="both", expand=True, padx=8, pady=(0, 8))
        cols = ("nr", "datum", "kunde", "bearbeiter", "msk", "status", "pos", "summe")
        heads = {"nr": "Auftrag", "datum": "Datum", "kunde": "Kunde", "bearbeiter": "Erfasst von",
                 "msk": "MSK", "status": "Status", "pos": "Pos.", "summe": "Summe (netto)"}
        tree = ttk.Treeview(tf, columns=cols, show="headings", selectmode="extended",
                            style="Kasse.Treeview")
        for c in cols:
            tree.heading(c, text=heads[c])
            w = 180 if c == "kunde" else (110 if c in ("summe",) else (
                90 if c in ("bearbeiter", "msk", "status") else (60 if c == "nr" else 70)))
            tree.column(c, width=w, anchor="w")
        tree.tag_configure("msk_erfasst", background="#e7f4ea")
        tree.tag_configure("online", background="#eef4fb")
        tree.pack(side="left", fill="both", expand=True)
        sb = tk.Scrollbar(tf, orient="vertical", command=tree.yview)
        sb.pack(side="right", fill="y")
        tree.configure(yscrollcommand=sb.set)
        _make_treeview_sortable(tree)
        tree.bind("<Double-1>", self._verkauf_detail_dialog)
        self.vh_tree = tree
        self._vh_rowmap = {}
        self._vh_online_map = {}    # Tree-iid -> Online-Auftrag-Id (für Doppelklick)
        self._online_verkaeufe = getattr(self, "_online_verkaeufe", [])

    def _online_auto_setup(self):
        """Startet das automatische Nachladen der Online-Bestellungen (alle 90 s),
        damit sich Multiarbeitsplätze gegenseitig aktuell halten."""
        self._online_auto_ms = 90 * 1000
        self._online_auto_tick()

    def _online_cache_path(self):
        import os
        return os.path.join(os.path.dirname(self.db_path), "online_cache.json")

    def _online_cache_laden(self):
        """Zuletzt geladene Online-Bestellungen aus dem lokalen Cache holen – sofort
        sichtbar nach Neustart, auch ohne Netz."""
        import json
        try:
            with open(self._online_cache_path(), encoding="utf-8") as fh:
                data = json.load(fh)
            self._online_verkaeufe = data.get("verkaeufe") or []
            self._online_vorbestellungen = data.get("vorbestellungen") or []
        except Exception:
            pass

    def _online_cache_speichern(self):
        """Online-Bestellungen lokal sichern (atomar). Liegt neben der DB – teilen
        sich Arbeitsplätze den DB-Ordner, teilen sie sich auch diesen Stand."""
        import json, os, tempfile
        path = self._online_cache_path()
        try:
            fd, tmp = tempfile.mkstemp(dir=os.path.dirname(path), suffix=".tmp")
            with os.fdopen(fd, "w", encoding="utf-8") as fh:
                json.dump({"verkaeufe": self._online_verkaeufe,
                           "vorbestellungen": self._online_vorbestellungen},
                          fh, ensure_ascii=False)
            os.replace(tmp, path)
        except Exception:
            pass

    def _online_auto_tick(self):
        """Periodischer Hintergrund-Abruf; plant sich selbst neu ein."""
        self._online_fetch_all_async(show_errors=False)
        try:
            self.after(self._online_auto_ms, self._online_auto_tick)
        except Exception:
            pass   # Fenster geschlossen -> Timer endet still

    def _online_verkaeufe_laden(self):
        """Button „🌐 Online laden": Verkäufe UND Vorbestellungen sofort nachladen."""
        for lbl in ("vh_info", "vb_info"):
            if hasattr(self, lbl):
                getattr(self, lbl).config(text="🌐 lädt …")
                getattr(self, lbl).update_idletasks()
        self._online_fetch_all_async(show_errors=True)

    def _online_fetch_all_async(self, show_errors):
        """Holt Online-Verkäufe + -Vorbestellungen im Hintergrund-Thread; das
        Ergebnis wird im Tk-Hauptthread per after(0, …) übernommen."""
        import threading

        def worker():
            verk, f1 = online_verkaeufe.fetch()
            vorb, f2 = online_verkaeufe.fetch_vorbestellungen()
            try:
                self.after(0, lambda: self._online_apply_all(verk, f1, vorb, f2, show_errors))
            except Exception:
                pass   # Fenster bereits zu

        threading.Thread(target=worker, daemon=True).start()

    def _online_apply_all(self, verk, f1, vorb, f2, show_errors):
        """Übernimmt beide Abruf-Ergebnisse (läuft im Tk-Hauptthread)."""
        if not f1:
            self._online_verkaeufe = verk
        if not f2:
            self._online_vorbestellungen = vorb
        if not (f1 and f2):           # mind. eine Liste erfolgreich -> Cache aktualisieren
            self._online_cache_speichern()
        fehler = f1 or f2
        if fehler and show_errors:
            messagebox.showwarning(
                "Online-Daten", _T('Konnte die Online-Daten nicht (vollständig) laden:\n{p0}', p0=fehler),
                parent=self.winfo_toplevel())
        if hasattr(self, "vh_tree"):
            self._refresh_verkaeufe()
        if hasattr(self, "vb_tree"):
            self._refresh_vorbestellungen()

    def _online_push_setup(self):
        """Startet das regelmäßige Pushen des Bestands zum Server (alle 2 Min)."""
        self._online_push_ms = 2 * 60 * 1000
        self._online_push_tick()

    def _online_push_tick(self):
        self._online_push_lager()
        try:
            self.after(self._online_push_ms, self._online_push_tick)
        except Exception:
            pass   # Fenster geschlossen -> Timer endet still

    def _online_push_lager(self):
        """Aktuellen Bestand der PC-Kasse im Hintergrund zum Server pushen, damit
        das Handy „verfügbar = Bestand − offene Online-Bestellungen" anzeigt."""
        try:
            with self._conn() as con:
                bestand = [{"pzn": r[0], "bezeichnung": r[1] or "", "charge": r[2] or "",
                            "verfall": r[3] or "", "menge": r[4] or 0}
                           for r in con.execute(
                    "SELECT pzn, artikelname, charge, verfall, menge FROM tbl_lagerbestand "
                    "WHERE pzn IS NOT NULL AND pzn<>''").fetchall()]
        except Exception:
            return
        seen = {}
        for b in bestand:
            seen.setdefault(b["pzn"], b["bezeichnung"])
        artikel = [{"pzn": p, "bezeichnung": n} for p, n in seen.items()]
        import threading
        threading.Thread(
            target=lambda: online_verkaeufe.push_lager(artikel, bestand),
            daemon=True).start()

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
        self._vh_online_map = {}
        filt = self._vh_filter_var.get().strip().lower()
        status_f = self._vh_status_var.get() if hasattr(self, "_vh_status_var") else "Alle"
        zeit_f = self._vh_zeit_var.get() if hasattr(self, "_vh_zeit_var") else "Alle"
        msk_f = self._vh_msk_var.get() if hasattr(self, "_vh_msk_var") else "Alle"
        von, bis = self._zeitraum_grenzen(zeit_f)
        with self._conn() as con:
            rows = con.execute(
                "SELECT b.id, b.datum, COALESCE(b.apotheke,''), COALESCE(b.status,'offen'), "
                "COALESCE(b.bearbeiter,''), COALESCE(b.msk_erfasst,0), "
                "SUM(CASE WHEN p.bestellart='Bestellung' THEN 1 ELSE 0 END), "
                "COALESCE(SUM(CASE WHEN p.bestellart='Bestellung' "
                "  THEN p.apu*p.menge*(1-COALESCE(p.rabatt_prozent,0)/100.0) ELSE 0 END),0) "
                "FROM tbl_bestellungen b LEFT JOIN tbl_bestellpositionen p ON p.bestell_id=b.id "
                "GROUP BY b.id "
                "HAVING SUM(CASE WHEN p.bestellart='Bestellung' THEN 1 ELSE 0 END) > 0 "
                "ORDER BY b.id DESC").fetchall()
        summe_ges = 0.0
        gezeigt = storniert_anz = msk_offen_anz = 0
        for bid, datum, apotheke, status, bearbeiter, msk, anz, summe in rows:
            if filt and filt not in str(apotheke).lower() and filt not in str(bid) \
                    and filt.lstrip("#") != str(bid):
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
                                      values=(f"#{bid}", datum, apotheke, bearbeiter, msk_txt, status,
                                              anz, _eur(summe)))
            self._vh_rowmap[iid] = bid
            if status != "storniert":
                summe_ges += summe or 0
            else:
                storniert_anz += 1
            if offen:
                msk_offen_anz += 1
            gezeigt += 1
        # Online erfasste Verkäufe (vom Server) zusätzlich einblenden – als
        # Nur-Lese-Zeilen mit „(online)". Kein Eintrag in _vh_rowmap -> Doppelklick
        # bleibt für sie wirkungslos (keine lokale Detailansicht).
        online_anz = 0
        for v in getattr(self, "_online_verkaeufe", []):
            vid = str(v.get("id"))
            datum = v.get("datum") or ""
            kunde = v.get("kunde_name") or ""
            erf = (v.get("erfasst_von") or "").strip()
            bearb = f"{erf} (online)" if erf else "(online)"
            if filt and filt not in kunde.lower() and filt != vid and filt.lstrip("#") != vid:
                continue
            if status_f == "Storniert":      # Online-Verkäufe sind nie storniert
                continue
            if msk_f != "Alle":              # MSK gilt nur für lokale Verkäufe
                continue
            if von and datum < von:
                continue
            if bis and datum > bis:
                continue
            iid = self.vh_tree.insert("", "end", tags=("online",),
                                      values=(f"#{vid}", datum, kunde, bearb, "–",
                                              v.get("status") or "offen",
                                              v.get("positionen", 0), "—"))
            self._vh_online_map[iid] = int(vid)
            online_anz += 1
            gezeigt += 1
        zusatz = f" · 🌐 {online_anz} online" if online_anz else ""
        self.vh_info.config(text=_T('{p0} Verkäufe{p1} · ⚠ {p2} MSK offen · {p3} storniert · Summe {p4}', p0=gezeigt, p1=zusatz, p2=msk_offen_anz, p3=storniert_anz, p4=_eur(summe_ges)))

    def _auftrag_id_eingabe(self, text):
        """Auftrag-Nr aus einer Eingabe lesen ('#12' / '12'); None wenn keine Zahl."""
        t = str(text or "").strip().lstrip("#").strip()
        return int(t) if t.isdigit() else None

    def _verkaeufe_open_nr(self):
        """Auftrag-Nr aus dem Suchfeld der Verkäufe-Liste oeffnen (Detailansicht)."""
        bid = self._auftrag_id_eingabe(self._vh_filter_var.get())
        if bid is None:
            messagebox.showinfo("Auftrag öffnen", "Bitte eine Auftrag-Nr (Zahl) eingeben.",
                                parent=self.winfo_toplevel())
            return
        with self._conn() as con:
            ok = con.execute("SELECT 1 FROM tbl_bestellungen WHERE id=?", (bid,)).fetchone()
        if not ok:
            messagebox.showinfo("Auftrag öffnen", _T('Kein Auftrag #{p0} gefunden.', p0=bid),
                                parent=self.winfo_toplevel())
            return
        self._verkauf_detail_window(bid)

    def _msk_markieren(self, erfasst, bids=None, mit_lieferschein=True):
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
        # Nach MSK-Erfassung ist der Verkauf zur Lieferung freigegeben -> Lieferschein.
        # mit_lieferschein=False, wenn der Aufrufer selbst schon einen erzeugt
        # (manueller Lieferschein-Button mit eigener MSK-Abfrage).
        if erfasst and mit_lieferschein:
            self._lieferscheine_erzeugen(bids)

    def _lieferscheine_erzeugen(self, bids):
        """Bietet nach der MSK-Erfassung an, fuer die betroffenen Verkaeufe einen
        Lieferschein zu erzeugen und zu oeffnen. Nur Verkaeufe mit echten
        (gelieferten) Bestell-Positionen erhalten einen Lieferschein."""
        top = self.winfo_toplevel()
        liefer_bids = []
        with self._conn() as con:
            for bid in bids:
                n = con.execute(
                    "SELECT COUNT(*) FROM tbl_bestellpositionen WHERE bestell_id=? "
                    "AND COALESCE(bestellart,'Bestellung')='Bestellung'", (bid,)).fetchone()[0]
                if n:
                    liefer_bids.append(bid)
        if not liefer_bids:
            return
        if len(liefer_bids) == 1:
            frage = (f"Verkauf #{liefer_bids[0]} ist in MSK erfasst.\n\n"
                     "Jetzt den Lieferschein erzeugen und öffnen?")
        else:
            frage = (f"{len(liefer_bids)} Verkäufe sind in MSK erfasst.\n\n"
                     "Jetzt die Lieferscheine erzeugen und öffnen?")
        if not messagebox.askyesno("Lieferschein", frage, parent=top):
            return
        erzeugt, fehler = [], []
        for bid in liefer_bids:
            try:
                path = lieferschein.render(self.db_path, bid)
                erzeugt.append((bid, path))
                self._log("Lieferschein erzeugt", bid, details=str(path.name))
            except Exception as e:
                fehler.append((bid, str(e)))
        for _bid, path in erzeugt:
            try:
                os.startfile(str(path))  # type: ignore[attr-defined]
            except Exception:
                pass
        if fehler:
            txt = "\n".join(f"#{b}: {e}" for b, e in fehler)
            messagebox.showerror("Lieferschein",
                                 _T('{p0} erzeugt, {p1} fehlgeschlagen:\n{p2}', p0=len(erzeugt), p1=len(fehler), p2=txt),
                                 parent=top)

    def _verkauf_detail_dialog(self, _e=None):
        sel = self.vh_tree.selection()
        if not sel:
            return
        if sel[0] in getattr(self, "_vh_online_map", {}):
            self._online_detail_window(self._vh_online_map[sel[0]])
            return
        bid = self._vh_rowmap.get(sel[0])
        if not bid:
            return
        self._verkauf_detail_window(bid)

    def _online_info_text(self, kopf):
        lt = ("    Liefertermin: " + kopf["liefertermin"]) if kopf.get("liefertermin") else ""
        return (f"Auftrag #{kopf.get('id')}    Datum: {kopf.get('datum', '')}{lt}\n"
                f"Kunde: {kopf.get('kunde_name', '')}\n"
                f"Erfasst von: {kopf.get('erfasst_von', '')} (online)    "
                f"Status: {kopf.get('status', '')}")

    def _online_detail_window(self, auftrag_id):
        """Öffnet das Detailfenster SOFORT (Kopf aus dem Cache + „lädt …") und füllt
        die Positionen nach, sobald der Server geantwortet hat – kein Warten aufs
        Fenster mehr."""
        top = self.winfo_toplevel()
        kopf0 = {"id": auftrag_id}
        for v in (self._online_verkaeufe + self._online_vorbestellungen):
            if str(v.get("id") or v.get("auftrag") or "") == str(auftrag_id):
                kopf0 = {"id": auftrag_id, "datum": v.get("datum", ""),
                         "kunde_name": v.get("kunde_name", ""), "status": v.get("status", ""),
                         "erfasst_von": v.get("erfasst_von", ""),
                         "liefertermin": v.get("liefertermin", "")}
                break

        win = tk.Toplevel(top)
        win.title(f"Online-Auftrag #{auftrag_id}")
        win.configure(bg=BG)
        win.geometry("620x540")
        win.minsize(480, 360)
        info_var = tk.StringVar(value=self._online_info_text(kopf0))
        tk.Label(win, textvariable=info_var, bg=BG, justify="left", anchor="w",
                 font=(theme.FONT, 10)).pack(fill="x", padx=12, pady=(12, 8))
        # Buttons + Zusammenfassung unten verankern (immer sichtbar).
        state = {"daten": None}
        btns = tk.Frame(win, bg=BG)
        btns.pack(side="bottom", fill="x", padx=12, pady=(6, 12))
        ueber_btn = tk.Button(
            btns, text="➡ In Kasse übernehmen", state="disabled",
            command=lambda: state["daten"] and self._online_uebernehmen(auftrag_id, state["daten"], win),
            bg="#11823b", fg="white", font=(theme.FONT, 11, "bold"), padx=14, pady=6)
        ueber_btn.pack(side="left")
        tk.Button(btns, text="Schließen", command=win.destroy,
                  font=(theme.FONT, 10), padx=12, pady=6).pack(side="right")
        summary_var = tk.StringVar(value="lädt …")
        tk.Label(win, textvariable=summary_var, bg=BG, fg="#555", anchor="w",
                 font=(theme.FONT, 9)).pack(side="bottom", fill="x", padx=12, pady=(2, 0))
        # Positionsliste füllt den restlichen Platz.
        tf = tk.Frame(win, bg=BG)
        tf.pack(fill="both", expand=True, padx=12, pady=(0, 6))
        cols = ("art", "pzn", "bez", "menge")
        heads = {"art": "Art", "pzn": "PZN", "bez": "Bezeichnung", "menge": "Menge"}
        tree = ttk.Treeview(tf, columns=cols, show="headings", height=8, style="Kasse.Treeview")
        for c in cols:
            tree.heading(c, text=heads[c])
            tree.column(c, width=(140 if c == "art" else (250 if c == "bez" else 90)), anchor="w")
        tree.tag_configure("vor", background="#fdf3e3")
        tree.tag_configure("best", background="#e9f5ec")
        tree.insert("", "end", values=("lädt …", "", "", ""))
        sb = tk.Scrollbar(tf, orient="vertical", command=tree.yview)
        sb.pack(side="right", fill="y")
        tree.configure(yscrollcommand=sb.set)
        tree.pack(side="left", fill="both", expand=True)

        def fuellen(daten, fehler):
            try:
                if not win.winfo_exists():
                    return
            except Exception:
                return
            tree.delete(*tree.get_children())
            if fehler or not daten:
                summary_var.set(_T('Konnte nicht laden: {p0}', p0=fehler or 'unbekannt'))
                return
            state["daten"] = daten
            kopf = daten.get("kopf") or kopf0
            pos = daten.get("positionen", [])
            info_var.set(self._online_info_text(kopf))
            for p in pos:
                art = str(p.get("bestellart", "Bestellung"))
                tag = "vor" if art.lower().startswith("vor") else "best"
                tree.insert("", "end", tags=(tag,),
                            values=(art, p.get("pzn", ""), p.get("bezeichnung", ""), p.get("menge", 0)))
            nv = sum(1 for p in pos if str(p.get("bestellart", "")).lower().startswith("vor"))
            nb = len(pos) - nv
            summary_var.set(_T('{p0} Bestellung (lieferbar, grün)  ·  {p1} Vorbestellung (orange)',
                               p0=nb, p1=nv))
            ueber_btn.config(state="normal")

        import threading

        def worker():
            daten, fehler = online_verkaeufe.fetch_detail(auftrag_id)
            try:
                self.after(0, lambda: fuellen(daten, fehler))
            except Exception:
                pass

        threading.Thread(target=worker, daemon=True).start()

    def _online_uebernehmen(self, auftrag_id, daten, win=None):
        """Lädt einen Online-Auftrag vorausgefüllt in den Verkauf-Reiter; beim
        Speichern wird er am Server als übernommen markiert (siehe _vk_save)."""
        top = self.winfo_toplevel()
        # Konflikt: im Verkauf liegt schon ein laufender Auftrag.
        if self.vk_positions:
            messagebox.showwarning(
                "Verkauf nicht leer",
                "Im Verkauf liegt bereits ein Auftrag. Bitte erst abschließen oder "
                "löschen, dann den Online-Auftrag übernehmen.", parent=top)
            return
        kopf = daten.get("kopf", {})
        positionen = daten.get("positionen", [])
        if not positionen:
            messagebox.showinfo("Online-Auftrag", "Dieser Auftrag hat keine Positionen.", parent=top)
            return
        # Kunde: über den Namen einen lokalen Kunden suchen, sonst „lose" übernehmen.
        name = kopf.get("kunde_name") or ""
        knr, det = self._kunde_by_name(name)
        self.vk_kunde = {
            "kundennummer": knr, "kundenname": det.get("kundenname") or name or "—",
            "plz": det.get("plz", ""), "ort": det.get("ort", ""), "strasse": det.get("strasse", ""),
            "inhaber": det.get("inhaber", ""), "telefon": det.get("telefon", ""),
            "email": det.get("email", ""),
        }
        self._render_kunde_card(self.vk_kunde)
        # Positionen: lokal per PZN anreichern (APU/DF/PCK); Aufteilung aus dem
        # Online-Auftrag übernehmen (Bestellung/Vorbestellung je Position).
        fehlend = []
        for pos in positionen:
            pzn = str(pos.get("pzn") or "").strip()
            det_a = self._artikel_details(pzn) if pzn else None
            if not det_a:
                fehlend.append(pos.get("bezeichnung") or pzn or "?")
            self.vk_positions.append({
                "pzn": pzn,
                "artikelname": (det_a or {}).get("artikelname") or pos.get("bezeichnung") or pzn,
                "df": (det_a or {}).get("df", ""), "pck": (det_a or {}).get("pck", ""),
                "apu": (det_a or {}).get("apu", 0) or 0,
                "menge": pos.get("menge", 0), "rabatt": 0, "rabatt_quelle": "online",
                "charge": "", "verfall": "",
                "bestellart": pos.get("bestellart", "Bestellung"),
                "lieferzeit": "", "liefertermin": kopf.get("liefertermin", "") or "",
                "_online_src": auftrag_id,
            })
        self._vk_render_positions()
        self._show_view("verkauf")
        if win is not None:
            try:
                win.destroy()
            except Exception:
                pass
        hinweis = ("Online-Auftrag in den Verkauf übernommen. Kunde/Charge/Bestand prüfen, "
                   "ggf. Mengen anpassen, dann „Verkauf speichern“.")
        if not knr:
            hinweis += f"\n\nHinweis: Kein lokaler Kunde zu „{name}“ gefunden – bitte Kunden prüfen/wählen."
        if fehlend:
            hinweis += "\n\nNicht im lokalen Stamm (APU=0, bitte prüfen): " + ", ".join(fehlend[:8])
        messagebox.showinfo("Online-Auftrag übernommen", hinweis, parent=top)

    def _kunde_by_name(self, name):
        """Sucht einen lokalen Kunden über den (Apotheken-)Namen. Gibt
        (kundennummer, details) oder (None, {})."""
        name = (name or "").strip()
        if not name:
            return None, {}
        with self._conn() as con:
            if not _table_exists(con, "tbl_kunden_center"):
                return None, {}
            row = con.execute(
                "SELECT kundennummer FROM tbl_kunden_center "
                "WHERE LOWER(kundenname)=LOWER(?) LIMIT 1", (name,)).fetchone()
        if not row:
            return None, {}
        knr = row[0]
        return knr, self._kunde_details(knr)

    def _online_mark_uebernommen(self, auftrag_id):
        """Markiert den Online-Auftrag am Server als übernommen (Hintergrund) und
        aktualisiert danach die Online-Listen."""
        import threading

        def worker():
            ok, fehler = online_verkaeufe.mark_uebernommen(auftrag_id, von=getpass.getuser())
            try:
                self.after(0, lambda: self._online_fetch_all_async(show_errors=False))
            except Exception:
                pass
            if not ok:
                try:
                    self.after(0, lambda: messagebox.showwarning(
                        "Online-Auftrag",
                        f"Verkauf gespeichert, aber der Online-Auftrag #{auftrag_id} konnte nicht "
                        f"als übernommen markiert werden:\n{fehler}\n\nEr erscheint ggf. erneut "
                        "in der Online-Liste.", parent=self.winfo_toplevel()))
                except Exception:
                    pass

        threading.Thread(target=worker, daemon=True).start()

    def _verkauf_detail_window(self, bid):
        """Read-only Detailansicht eines Verkaufs (Stornieren/MSK/Auftragsbestaetigung).
        Wird per Doppelklick in „Verkäufe“ UND per Auftrag-Nr-Suche aufgerufen -
        erzeugt selbst KEINEN neuen Auftrag."""
        if not bid:
            return
        with self._conn() as con:
            h = con.execute("SELECT datum, COALESCE(apotheke,''), COALESCE(status,'offen'), "
                            "COALESCE(bearbeiter,''), COALESCE(msk_erfasst,0), COALESCE(msk_von,''), "
                            "COALESCE(msk_am,'') FROM tbl_bestellungen WHERE id=?", (bid,)).fetchone()
            # Nur echte Bestellungen anzeigen - Vorbestellungen gehoeren in den
            # Reiter "Vorbestellungen", nicht in die Verkaufs-Detailansicht.
            positions = con.execute(
                "SELECT pzn, artikelname, menge, COALESCE(rabatt_prozent,0), apu, "
                "COALESCE(charge,''), COALESCE(verfall,''), COALESCE(bestellart,'Bestellung') "
                "FROM tbl_bestellpositionen WHERE bestell_id=? "
                "AND COALESCE(bestellart,'Bestellung')='Bestellung' ORDER BY id", (bid,)).fetchall()
        if not h:
            return
        datum, apotheke, status, bearbeiter, msk, msk_von, msk_am = h
        top = self.winfo_toplevel()
        win = tk.Toplevel(top)
        win.title(f"Verkauf #{bid}")
        win.configure(bg=BG)
        win.transient(top)
        tk.Label(win, text=_T('Verkauf #{p0} · {p1}', p0=bid, p1=apotheke), font=(theme.FONT, 13, "bold"),
                 fg=ACCENT, bg=BG).pack(padx=18, pady=(14, 2), anchor="w")
        tk.Label(win, text=_T('Datum {p0} · Status: {p1} · Erfasst von: {p2}', p0=datum, p1=status, p2=bearbeiter or '–'),
                 font=(theme.FONT, 9), fg="#555", bg=BG).pack(padx=18, anchor="w")
        msk_txt = (f"MSK: ✓ erfasst von {msk_von or '?'} am {(msk_am or '')[:16].replace('T', ' ')}"
                   if msk else "MSK: ⚠ noch nicht erfasst")
        tk.Label(win, text=msk_txt, font=(theme.FONT, 9, "bold"),
                 fg=("#11823b" if msk else "#a35a00"), bg=BG).pack(padx=18, anchor="w", pady=(0, 8))
        tf = tk.Frame(win, bg=BG)
        tf.pack(fill="both", expand=True, padx=18, pady=(0, 8))
        cols = ("pzn", "artikel", "menge", "rabatt", "charge", "verfall", "art")
        heads = {"pzn": "PZN", "artikel": "Artikel", "menge": "Menge", "rabatt": "Rab%",
                 "charge": "Charge", "verfall": "Verfall", "art": "Art"}
        tree = ttk.Treeview(tf, columns=cols, show="headings", height=8,
                            style="Kasse.Treeview")
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
        tk.Label(win, text=_T('Gesamt (netto): {p0}', p0=_eur(gesamt)), font=(theme.FONT, 11, "bold"),
                 fg=ACCENT, bg=BG).pack(padx=18, anchor="e")

        def stornieren():
            if status == "storniert":
                return
            if not messagebox.askyesno("Stornieren", _T('Verkauf #{p0} wirklich stornieren?\nDer Lagerbestand wird zurückgebucht.', p0=bid), parent=win):
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
                                "INSERT INTO tbl_lagerbestand(pzn,artikelname,charge,verfall,menge,ek,aktualisiert_am) "
                                "VALUES(?,?,?,?,?,?,?)", (pzn, art, ch, vf, menge, apu, jetzt))
                con.commit()
            self._log("Verkauf storniert", bid, apotheke, "Bestand zurückgebucht")
            self._refresh_verkaeufe()
            self._refresh_artikel()
            win.destroy()
            messagebox.showinfo("Storniert", _T('Verkauf #{p0} storniert, Bestand zurückgebucht.', p0=bid),
                                parent=top)

        def msk_toggle():
            self._msk_markieren(not bool(msk), bids=[bid])
            win.destroy()

        btns = tk.Frame(win, bg=BG)
        btns.pack(fill="x", padx=18, pady=(4, 14))
        if status != "storniert":
            if not msk:
                tk.Button(btns, text="✏ Bearbeiten",
                          command=lambda: self._verkauf_bearbeiten(bid, win),
                          bg="#b06a00", fg="white", font=(theme.FONT, 10, "bold"),
                          padx=12, pady=4).pack(side="left", padx=(0, 8))
            tk.Button(btns, text="Stornieren", command=stornieren, bg="#a32d2d", fg="white",
                      font=(theme.FONT, 10, "bold"), padx=12, pady=4).pack(side="left")
            if msk:
                tk.Button(btns, text="↺ MSK-Markierung aufheben", command=msk_toggle,
                          padx=12, pady=4).pack(side="left", padx=(8, 0))
            else:
                tk.Button(btns, text="✓ In MSK erfasst", command=msk_toggle, bg="#11823b",
                          fg="white", font=(theme.FONT, 10, "bold"), padx=12, pady=4).pack(side="left", padx=(8, 0))
        # Auftragsbestaetigung jederzeit erneut erzeugen.
        tk.Button(btns, text="🖨 Auftragsbestätigung", command=lambda: self._auftrag_drucken(bid, win),
                  bg=ACCENT, fg="white", font=(theme.FONT, 10, "bold"), padx=12, pady=4).pack(side="left", padx=(8, 0))
        # Lieferschein jederzeit erneut erzeugen (wird sonst bei MSK-Erfassung erstellt).
        tk.Button(btns, text="📋 Lieferschein", command=lambda: self._lieferschein_drucken(bid, win),
                  bg="#1f7a4d", fg="white", font=(theme.FONT, 10, "bold"), padx=12, pady=4).pack(side="left", padx=(6, 0))
        # Wer hat was wann mit diesem Auftrag gemacht (Audit-Verlauf).
        tk.Button(btns, text="🕘 Verlauf", command=lambda: self._auftrag_verlauf_window(bid),
                  bg="#5a4b8a", fg="white", font=(theme.FONT, 10, "bold"), padx=12, pady=4).pack(side="left", padx=(6, 0))
        tk.Button(btns, text="📧 Per E-Mail", command=lambda: self._auftrag_mail(bid, win),
                  bg="#3867b7", fg="white", font=(theme.FONT, 10, "bold"), padx=12, pady=4).pack(side="left", padx=(6, 0))
        tk.Button(btns, text="Schließen", command=win.destroy, padx=12, pady=4).pack(side="right")
        self._restyle_buttons(win)
        win.lift()
        win.focus_force()

    def _verkauf_bearbeiten(self, bid, win=None):
        """Lädt einen noch nicht MSK-erfassten Verkauf zum Bearbeiten in den
        Verkauf-Reiter. Das Original bleibt bestehen, bis der korrigierte Auftrag
        gespeichert wird (dann wird es ersetzt, Bestand korrekt verrechnet)."""
        top = self.winfo_toplevel()
        if self.vk_positions:
            messagebox.showwarning(
                "Verkauf nicht leer",
                "Im Verkauf liegt bereits ein Auftrag. Bitte erst abschließen oder "
                "löschen, dann den Verkauf zum Bearbeiten laden.", parent=top)
            return
        with self._conn() as con:
            h = con.execute(
                "SELECT kundennummer, COALESCE(apotheke,''), COALESCE(msk_erfasst,0), "
                "COALESCE(status,'offen') FROM tbl_bestellungen WHERE id=?", (bid,)).fetchone()
            if not h:
                return
            knr, apotheke, msk, status = h
            if msk:
                messagebox.showinfo(
                    "Bearbeiten", "Dieser Verkauf ist bereits in MSK erfasst und kann nicht "
                    "mehr bearbeitet werden.", parent=top)
                return
            if status == "storniert":
                return
            pos = con.execute(
                "SELECT pzn, artikelname, COALESCE(df,''), COALESCE(pck,''), apu, menge, "
                "COALESCE(rabatt_prozent,0), COALESCE(rabatt_quelle,'manuell'), "
                "COALESCE(charge,''), COALESCE(verfall,''), COALESCE(bestellart,'Bestellung'), "
                "COALESCE(lieferzeit,''), COALESCE(liefertermin,'') "
                "FROM tbl_bestellpositionen WHERE bestell_id=? ORDER BY id", (bid,)).fetchall()
        if not messagebox.askyesno(
                "Bearbeiten",
                f"Verkauf #{bid} zum Bearbeiten in den Verkauf laden?\n\n"
                "Das Original bleibt erhalten, bis du den korrigierten Auftrag speicherst.",
                parent=top):
            return
        det = self._kunde_details(knr)
        self.vk_kunde = {
            "kundennummer": knr, "kundenname": det.get("kundenname") or apotheke or "—",
            "plz": det.get("plz", ""), "ort": det.get("ort", ""), "strasse": det.get("strasse", ""),
            "inhaber": det.get("inhaber", ""), "telefon": det.get("telefon", ""),
            "email": det.get("email", ""),
        }
        self._render_kunde_card(self.vk_kunde)
        for (pzn, art, df, pck, apu, menge, rab, rq, ch, vf, ba, lz, lt) in pos:
            self.vk_positions.append({
                "pzn": pzn, "artikelname": art, "df": df, "pck": pck, "apu": apu or 0,
                "menge": menge, "rabatt": rab or 0, "rabatt_quelle": rq or "manuell",
                "charge": ch, "verfall": vf, "bestellart": ba or "Bestellung",
                "lieferzeit": lz, "liefertermin": lt, "_edit_src": bid,
            })
        self._vk_render_positions()
        self._show_view("verkauf")
        if win is not None:
            try:
                win.destroy()
            except Exception:
                pass
        messagebox.showinfo(
            "Bearbeiten",
            f"Verkauf #{bid} in den Verkauf geladen. Anpassen (Mengen, Charge, Artikel), "
            "dann „Verkauf speichern“ – der korrigierte Auftrag ersetzt den alten.", parent=top)

    def _verkauf_original_ersetzen(self, bid):
        """Bucht den Bestand des Original-Auftrags zurück und entfernt ihn – wird
        beim Speichern eines bearbeiteten Verkaufs aufgerufen (siehe _vk_save)."""
        jetzt = datetime.now().isoformat(timespec="seconds")
        with self._conn() as con:
            rows = con.execute(
                "SELECT pzn, artikelname, menge, COALESCE(charge,''), COALESCE(verfall,''), "
                "COALESCE(bestellart,'Bestellung'), apu FROM tbl_bestellpositionen "
                "WHERE bestell_id=?", (bid,)).fetchall()
            for pzn, art, menge, ch, vf, ba, apu in rows:
                if ba == "Bestellung" and (ch or vf):
                    upd = con.execute(
                        "UPDATE tbl_lagerbestand SET menge=menge+?, aktualisiert_am=? "
                        "WHERE pzn=? AND COALESCE(charge,'')=? AND COALESCE(verfall,'')=?",
                        (menge, jetzt, pzn, ch, vf))
                    if upd.rowcount == 0:
                        con.execute(
                            "INSERT INTO tbl_lagerbestand(pzn,artikelname,charge,verfall,menge,ek,aktualisiert_am) "
                            "VALUES(?,?,?,?,?,?,?)", (pzn, art, ch, vf, menge, apu, jetzt))
            con.execute("DELETE FROM tbl_bestellpositionen WHERE bestell_id=?", (bid,))
            con.execute("DELETE FROM tbl_bestellungen WHERE id=?", (bid,))
            con.commit()
        self._log("Verkauf bearbeitet", bid, details="Original ersetzt, Bestand zurückgebucht")

    # ================================================================ Artikel
    def _build_artikel(self, parent):
        bar = tk.Frame(parent, bg=BG)
        bar.pack(fill="x", padx=8, pady=(10, 4))
        tk.Label(bar, text="Suche (PZN / Artikel):", bg=BG, font=(theme.FONT, 9, "bold")).pack(side="left")
        self._art_filter_var = tk.StringVar()
        e = tk.Entry(bar, textvariable=self._art_filter_var, width=28)
        e.pack(side="left", padx=(6, 12))
        e.bind("<KeyRelease>", lambda _e: self._refresh_artikel())
        self._art_nur_bestand = tk.BooleanVar(value=False)
        tk.Radiobutton(bar, text="Alle", variable=self._art_nur_bestand, value=False, bg=BG,
                       font=(theme.FONT, 9), command=self._refresh_artikel).pack(side="left")
        tk.Radiobutton(bar, text="Nur mit Bestand", variable=self._art_nur_bestand, value=True, bg=BG,
                       font=(theme.FONT, 9), command=self._refresh_artikel).pack(side="left", padx=(4, 0))
        self.art_info = tk.Label(bar, text="", bg=BG, fg=ACCENT, font=(theme.FONT, 9, "bold"))
        self.art_info.pack(side="right")

        tk.Label(parent, text="Doppelklick auf einen Artikel: alle Chargen und Verfälle anzeigen.",
                 bg=BG, fg="#666", font=(theme.FONT, 9)).pack(anchor="w", padx=8, pady=(0, 4))
        tf = tk.Frame(parent, bg=BG)
        tf.pack(fill="both", expand=True, padx=8, pady=(0, 8))
        cols = ("pzn", "artikel", "df", "pck", "apu", "bestand")
        heads = {"pzn": "PZN", "artikel": "Artikel", "df": "DF", "pck": "PCK",
                 "apu": "APU €", "bestand": "Bestand"}
        tree = ttk.Treeview(tf, columns=cols, show="headings", style="Kasse.Treeview")
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

        # Summenleiste unten: passt sich der aktuellen Suche/Filterung an.
        summe_bar = tk.Frame(parent, bg=ACCENT_LIGHT, highlightbackground=BORDER,
                             highlightthickness=1)
        summe_bar.pack(fill="x", padx=8, pady=(0, 8))
        self.art_summe = tk.Label(summe_bar, text="", bg=ACCENT_LIGHT, fg=ACCENT,
                                  font=(theme.FONT, 11, "bold"), anchor="w")
        self.art_summe.pack(side="left", padx=12, pady=6)

    def _refresh_artikel(self):
        if not hasattr(self, "art_tree"):
            return
        self.art_tree.delete(*self.art_tree.get_children())
        like = f"%{self._art_filter_var.get().strip()}%"
        nur_bestand = self._art_nur_bestand.get()
        hat_stamm = self._table_exists_now("tbl_artikelstamm")
        hat_wa = self._table_exists_now("tbl_gdp_warenausgang")
        with self._conn() as con:
            bestand_sub = "(SELECT COALESCE(SUM(menge),0) FROM tbl_lagerbestand l WHERE l.pzn=n.pzn)"
            lagerwert_sub = ("(SELECT COALESCE(SUM(COALESCE(ek,0)*menge),0) "
                             "FROM tbl_lagerbestand l WHERE l.pzn=n.pzn)")
            # Klammer-Bestand (unterwegs, geladen aber noch nicht bei MSK bestaetigt).
            unterwegs_sub = (
                "(SELECT COALESCE(SUM(p.menge),0) FROM tbl_gdp_warenausgang_pos p "
                "JOIN tbl_gdp_warenausgang w ON w.id=p.wa_id "
                "WHERE p.pzn=n.pzn AND w.status='geladen' AND w.ziel='Verkaufsbestand')"
            ) if hat_wa else "0"
            if hat_stamm:
                sql = (f"SELECT n.pzn, n.artikelname, COALESCE(a.df,''), COALESCE(a.pck,''), n.apu, "
                       f"{bestand_sub}, {lagerwert_sub}, {unterwegs_sub} "
                       "FROM tbl_nmg_stamm n LEFT JOIN tbl_artikelstamm a ON a.pzn=n.pzn "
                       "WHERE (n.pzn LIKE ? OR n.artikelname LIKE ?) ORDER BY n.artikelname")
            else:
                sql = (f"SELECT n.pzn, n.artikelname, '', '', n.apu, {bestand_sub}, {lagerwert_sub}, "
                       f"{unterwegs_sub} FROM tbl_nmg_stamm n "
                       "WHERE (n.pzn LIKE ? OR n.artikelname LIKE ?) ORDER BY n.artikelname")
            rows = con.execute(sql, (like, like)).fetchall()
        gezeigt = 0
        summe_bestand = 0
        summe_wert = 0.0
        summe_lagerwert = 0.0
        for pzn, art, df, pck, apu, bestand, lagerwert, unterwegs in rows:
            if nur_bestand and (bestand or 0) <= 0 and (unterwegs or 0) <= 0:
                continue
            # Eintreffender Nachschub in Klammern: "12 (+30)" – nur Anzeige.
            bestand_disp = f"{bestand} (+{unterwegs})" if (unterwegs or 0) > 0 else bestand
            self.art_tree.insert("", "end", values=(pzn, art, df, pck, _eur(apu), bestand_disp))
            gezeigt += 1
            summe_bestand += bestand or 0
            summe_wert += (apu or 0) * (bestand or 0)
            summe_lagerwert += lagerwert or 0
        self.art_info.config(text=_T('{p0} Artikel', p0=gezeigt))
        umfang = "Auswahl" if self._art_filter_var.get().strip() else "gesamt"
        self.art_summe.config(
            text=_T('Bestand {p0}: {p1} Packungen   ·   Verkaufswert (APU × Bestand): {p2}   ·   Lagerwert (EK × Bestand): {p3}', p0=umfang, p1=summe_bestand, p2=_eur(summe_wert), p3=_eur(summe_lagerwert)))

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
        tk.Label(win, text=art or pzn, font=(theme.FONT, 13, "bold"), fg=ACCENT,
                 bg=BG).pack(padx=18, pady=(14, 2), anchor="w")
        tk.Label(win, text=_T('PZN {p0}', p0=pzn), font=(theme.FONT, 9), fg="#555", bg=BG).pack(padx=18, anchor="w")
        tk.Label(win, text="Doppelklick auf eine Charge: Bestand korrigieren.",
                 font=(theme.FONT, 9), fg="#666", bg=BG).pack(padx=18, anchor="w", pady=(2, 0))
        tf = tk.Frame(win, bg=BG)
        tf.pack(fill="both", expand=True, padx=18, pady=(8, 8))
        tree = ttk.Treeview(tf, columns=("charge", "verfall", "bestand"), show="headings",
                            height=8, style="Kasse.Treeview")
        for c, t, w in (("charge", "Charge", 160), ("verfall", "Verfall", 130), ("bestand", "Bestand", 90)):
            tree.heading(c, text=t)
            tree.column(c, width=w, anchor="w")
        tree.pack(side="left", fill="both", expand=True)
        sb = tk.Scrollbar(tf, orient="vertical", command=tree.yview)
        sb.pack(side="right", fill="y")
        tree.configure(yscrollcommand=sb.set)
        _make_treeview_sortable(tree)
        # iid -> (charge, verfall, menge) mit Rohwerten (fuer die Korrektur).
        row_map = {}

        def _fill():
            tree.delete(*tree.get_children())
            row_map.clear()
            chargen = self._lager_chargen(pzn)
            if chargen:
                for ch, vf, m in chargen:
                    iid = tree.insert("", "end", values=(ch or "—", vf or "—", m))
                    row_map[iid] = (ch or "", vf or "", m)
            else:
                tree.insert("", "end", values=("(kein Bestand)", "", ""))

        def _korr(_ev=None):
            sel2 = tree.selection()
            if not sel2 or sel2[0] not in row_map:
                return
            ch, vf, m = row_map[sel2[0]]
            if self._lager_korrektur_dialog(pzn, ch, vf, int(m)):
                _fill()
                self._refresh_artikel()

        _fill()
        tree.bind("<Double-1>", _korr)
        tk.Button(win, text="Schließen", command=win.destroy, padx=14, pady=4).pack(pady=(0, 14))
        self._restyle_buttons(win)
        win.lift()
        win.focus_force()

    # ============================================================ Defektmeldung
    def _build_defektmeldung(self, parent):
        tk.Label(parent, text="Für Artikel ohne (ausreichenden) Bestand eine "
                 "Nichtverfügbarkeits-Bescheinigung für die Apotheke erzeugen.",
                 bg=BG, fg="#666", font=(theme.FONT, 9)).pack(anchor="w", padx=8, pady=(10, 2))

        kopf_card, kopf = _make_card(parent, "Apotheke")
        kopf_card.pack(fill="x", padx=8, pady=(6, 6))
        khead = tk.Frame(kopf, bg=BG)
        khead.pack(fill="x")
        tk.Label(khead, text="Apotheke suchen:", bg=BG, font=(theme.FONT, 10, "bold")).pack(side="left")
        tk.Button(khead, text="✕ entfernen", command=self._dm_clear_kunde,
                  font=(theme.FONT, 8), padx=6, pady=1).pack(side="left", padx=(10, 0))
        self.dm_kunde_search = SearchBox(
            kopf,
            fetch=lambda t: [(f"{knr or '—'}  ·  {name}  ·  {plz} {ort}", (knr, name, plz, ort))
                             for knr, name, plz, ort in self._search_kunden(t)],
            on_select=self._dm_pick_kunde, height=5,
        )
        self.dm_kunde_search.pack(fill="x", pady=(2, 6))
        self.dm_kunde_card = tk.Frame(kopf, bg="#f2f6fb", highlightbackground="#dde7f1",
                                      highlightthickness=1)
        self.dm_kunde_card.pack(fill="x")
        self.dm_kunde_card_label = tk.Label(self.dm_kunde_card, text="Keine Apotheke gewählt.",
                                            bg="#f2f6fb", fg="#888", justify="left", anchor="w",
                                            font=(theme.FONT, 9, "italic"))
        self.dm_kunde_card_label.pack(fill="x", padx=10, pady=8)

        add_card, addf = _make_card(parent, "Artikel hinzufügen", "nur NMG")
        add_card.pack(fill="x", padx=8, pady=(0, 6))
        self.dm_artikel_search = SearchBox(
            addf,
            fetch=lambda t: [(f"{pzn}  ·  {name}", pzn) for pzn, name in self._search_nmg(t)],
            on_select=self._dm_pick_artikel, height=5,
        )
        self.dm_artikel_search.pack(fill="x", pady=(2, 4))
        line = tk.Frame(addf, bg=BG)
        line.pack(fill="x", pady=(2, 0))
        self.dm_detail_label = tk.Label(line, text="—", bg=BG, fg="#444",
                                        font=(theme.FONT, 9), anchor="w")
        self.dm_detail_label.pack(side="left", fill="x", expand=True)
        line2 = tk.Frame(addf, bg=BG)
        line2.pack(fill="x", pady=(6, 0))
        tk.Label(line2, text="Menge:", bg=BG, font=(theme.FONT, 9, "bold")).pack(side="left")
        self.dm_menge_var = tk.StringVar(value="1")
        tk.Entry(line2, textvariable=self.dm_menge_var, width=6).pack(side="left", padx=(4, 12))
        tk.Button(line2, text="+ Position", command=self._dm_add_position,
                  bg="#11823b", fg="white", font=(theme.FONT, 9, "bold"),
                  padx=10, pady=2).pack(side="left")

        pos_card, pos = _make_card(parent, "Nicht verfügbare Artikel")
        pos_card.pack(fill="both", expand=True, padx=8, pady=(0, 6))
        cols = ("pzn", "artikel", "menge", "bestand")
        heads = {"pzn": "PZN", "artikel": "Artikel", "menge": "Menge", "bestand": "akt. Bestand"}
        tf = tk.Frame(pos, bg=BG)
        tf.pack(fill="both", expand=True)
        tree = ttk.Treeview(tf, columns=cols, show="headings", height=5, style="Kasse.Treeview")
        for c in cols:
            tree.heading(c, text=heads[c])
            tree.column(c, width=240 if c == "artikel" else 90,
                        anchor="e" if c in ("menge", "bestand") else "w")
        tree.pack(side="left", fill="both", expand=True)
        sb = tk.Scrollbar(tf, orient="vertical", command=tree.yview)
        sb.pack(side="right", fill="y")
        tree.configure(yscrollcommand=sb.set)
        _make_treeview_sortable(tree)
        tree.bind("<Delete>", self._dm_remove_position)
        self.dm_pos_tree = tree
        self._dm_tree_map = {}

        grund_row = tk.Frame(parent, bg=BG)
        grund_row.pack(fill="x", padx=8, pady=(0, 4))
        tk.Label(grund_row, text="Grund:", bg=BG, font=(theme.FONT, 9, "bold")).pack(side="left")
        self.dm_grund_var = tk.StringVar(value="nicht lieferbar")
        ttk.Combobox(grund_row, textvariable=self.dm_grund_var, width=26, state="readonly",
                     values=("nicht lieferbar", "nicht vorrätig / kein Bestand",
                             "vorübergehend nicht verfügbar", "Herstellerengpass",
                             "Sonstiges")).pack(side="left", padx=(4, 8))
        tk.Label(grund_row, text="Zusatz:", bg=BG, font=(theme.FONT, 9, "bold")).pack(side="left")
        self.dm_grund_frei = tk.StringVar()
        tk.Entry(grund_row, textvariable=self.dm_grund_frei, width=30).pack(side="left", padx=(4, 0))

        foot = tk.Frame(parent, bg=BG)
        foot.pack(fill="x", padx=8, pady=(0, 8))
        tk.Button(foot, text="🗑 Position löschen", command=lambda: self._dm_remove_position(None),
                  font=(theme.FONT, 9), padx=8, pady=2).pack(side="left")
        tk.Button(foot, text="⚠ Defektmeldung erzeugen", command=self._defektmeldung_erzeugen,
                  bg="#a32d2d", fg="white", font=(theme.FONT, 10, "bold"),
                  padx=14, pady=4).pack(side="right")

    def _dm_pick_kunde(self, payload):
        knr, name, plz, ort = payload
        det = self._kunde_details(knr)
        self.dm_kunde = {
            "kundennummer": knr, "kundenname": det.get("kundenname") or name,
            "plz": det.get("plz") or plz, "ort": det.get("ort") or ort,
            "strasse": det.get("strasse", ""), "inhaber": det.get("inhaber", ""),
            "email": det.get("email", ""),
        }
        k = self.dm_kunde
        zeilen = " · ".join(x for x in [
            f"Nr. {k['kundennummer']}" if k.get("kundennummer") else "",
            k.get("inhaber", ""), k.get("strasse", ""),
            f"{k.get('plz','')} {k.get('ort','')}".strip()] if x)
        self.dm_kunde_card_label.config(
            text=f"{k.get('kundenname') or '—'}\n{zeilen}", fg=ACCENT,
            font=(theme.FONT, 9))

    def _dm_clear_kunde(self):
        self.dm_kunde = None
        self.dm_kunde_search.clear()
        self.dm_kunde_card_label.config(text="Keine Apotheke gewählt.", fg="#888",
                                        font=(theme.FONT, 9, "italic"))

    def _dm_pick_artikel(self, pzn):
        det = self._artikel_details(pzn)
        if not det:
            return
        self.dm_cur = det
        bestand = self._bestand_total(pzn)
        hinweis = "kein Bestand" if bestand <= 0 else f"⚠ noch {bestand} auf Lager"
        self.dm_detail_label.config(
            text=_T('{p0}  ·  PZN {p1}  ·  {p2}', p0=det['artikelname'], p1=pzn, p2=hinweis),
            fg="#a35a00" if bestand > 0 else ACCENT)

    def _dm_add_position(self, *_):
        if not self.dm_cur:
            messagebox.showwarning("Defektmeldung", "Bitte zuerst einen Artikel wählen.",
                                   parent=self.winfo_toplevel())
            return
        try:
            menge = int(self.dm_menge_var.get().strip())
            if menge <= 0:
                raise ValueError
        except ValueError:
            messagebox.showwarning("Defektmeldung", "Menge muss eine positive Zahl sein.",
                                   parent=self.winfo_toplevel())
            return
        pzn = self.dm_cur["pzn"]
        bestand = self._bestand_total(pzn)
        self.dm_positions.append({"pzn": pzn, "artikelname": self.dm_cur["artikelname"],
                                  "menge": menge, "bestand": bestand})
        self.dm_cur = None
        self.dm_artikel_search.clear()
        self.dm_detail_label.config(text="—", fg="#444")
        self.dm_menge_var.set("1")
        self._dm_render_positions()

    def _dm_render_positions(self):
        self.dm_pos_tree.delete(*self.dm_pos_tree.get_children())
        self._dm_tree_map = {}
        for p in self.dm_positions:
            iid = self.dm_pos_tree.insert("", "end", values=(
                p["pzn"], p["artikelname"], p["menge"], p["bestand"]))
            self._dm_tree_map[iid] = p

    def _dm_remove_position(self, _e):
        sel = self.dm_pos_tree.selection()
        if not sel:
            return
        ids = {id(self._dm_tree_map.get(i)) for i in sel}
        self.dm_positions = [x for x in self.dm_positions if id(x) not in ids]
        self._dm_render_positions()

    def _defektmeldung_erzeugen(self):
        top = self.winfo_toplevel()
        if not self.dm_kunde:
            messagebox.showwarning("Defektmeldung", "Bitte eine Apotheke wählen.", parent=top)
            return
        if not self.dm_positions:
            messagebox.showwarning("Defektmeldung", "Keine Artikel erfasst.", parent=top)
            return
        grund = self.dm_grund_var.get()
        if self.dm_grund_frei.get().strip():
            grund = f"{grund} – {self.dm_grund_frei.get().strip()}"
        try:
            path = defektmeldung.render(self.db_path, knr=self.dm_kunde["kundennummer"],
                                        kunde=self.dm_kunde, positionen=self.dm_positions,
                                        grund=grund)
        except Exception as e:
            messagebox.showerror("Defektmeldung", _T('Konnte nicht erstellt werden:\n{p0}', p0=e), parent=top)
            return
        self._log("Defektmeldung erzeugt", kunde=self.dm_kunde.get("kundenname", ""),
                  details=f"{len(self.dm_positions)} Artikel · {grund}")
        try:
            os.startfile(str(path))  # type: ignore[attr-defined]
        except Exception:
            messagebox.showinfo("Defektmeldung", _T('Gespeichert:\n{p0}', p0=path), parent=top)
        # Nach dem Erzeugen die Liste leeren (Apotheke bleibt gewaehlt).
        self.dm_positions = []
        self._dm_render_positions()

    # ============================================================= Einstellungen
    def _build_einstellungen(self, parent):
        wrap = tk.Frame(parent, bg=BG)
        wrap.pack(fill="both", expand=True, padx=10, pady=10)

        # --- Firmendaten (Kopf aller Dokumente) ---
        fcard, f = _make_card(wrap, "Firmendaten",
                              "erscheinen im Kopf von Auftragsbestätigung, Lieferschein, Defektmeldung")
        fcard.pack(fill="x", pady=(0, 10))
        r1 = tk.Frame(f, bg=BG)
        r1.pack(fill="x", pady=(2, 4))
        tk.Label(r1, text="Firma:", bg=BG, width=14, anchor="w",
                 font=(theme.FONT, 9, "bold")).pack(side="left")
        self._set_firma_name = tk.StringVar(value=einstellungen.get(self.db_path, "firma_name"))
        tk.Entry(r1, textvariable=self._set_firma_name, width=50).pack(side="left", fill="x", expand=True)
        r2 = tk.Frame(f, bg=BG)
        r2.pack(fill="x", pady=(2, 4))
        tk.Label(r2, text="Adresse:", bg=BG, width=14, anchor="nw",
                 font=(theme.FONT, 9, "bold")).pack(side="left", anchor="n")
        self._set_firma_adresse_txt = tk.Text(r2, height=3, width=50, wrap="word",
                                              font=(theme.FONT, 10))
        self._set_firma_adresse_txt.pack(side="left", fill="x", expand=True)
        self._set_firma_adresse_txt.insert("1.0", einstellungen.get(self.db_path, "firma_adresse"))
        r3 = tk.Frame(f, bg=BG)
        r3.pack(fill="x", pady=(2, 2))
        tk.Label(r3, text="Kontakt:", bg=BG, width=14, anchor="w",
                 font=(theme.FONT, 9, "bold")).pack(side="left")
        self._set_firma_kontakt = tk.StringVar(value=einstellungen.get(self.db_path, "firma_kontakt"))
        tk.Entry(r3, textvariable=self._set_firma_kontakt, width=50).pack(side="left", fill="x", expand=True)

        # --- Defektmeldung-Rechtstext ---
        dcard, d = _make_card(wrap, "Defektmeldung – Rechtstext",
                              "erscheint unten auf der Defektmeldung")
        dcard.pack(fill="both", expand=True, pady=(0, 10))
        tk.Label(d, text="Hier die fachlich/rechtlich geprüfte Formulierung samt gesetzlicher "
                 "Grundlagen eintragen:", bg=BG, fg="#666",
                 font=(theme.FONT, 9)).pack(anchor="w", pady=(0, 4))
        self._set_rechtstext_txt = tk.Text(d, height=8, wrap="word", font=(theme.FONT, 10))
        self._set_rechtstext_txt.pack(fill="both", expand=True)
        self._set_rechtstext_txt.insert("1.0", einstellungen.get(self.db_path, "defekt_rechtstext"))

        # --- Parameter ---
        pcard, pf = _make_card(wrap, "Tagesabschluss")
        pcard.pack(fill="x", pady=(0, 10))
        pr = tk.Frame(pf, bg=BG)
        pr.pack(fill="x")
        tk.Label(pr, text="Automatisch erzeugen um (Uhrzeit, 0–23):", bg=BG,
                 font=(theme.FONT, 9, "bold")).pack(side="left")
        self._set_stunde_var = tk.StringVar(
            value=str(einstellungen.get_int(self.db_path, "tagesabschluss_stunde", 18)))
        tk.Entry(pr, textvariable=self._set_stunde_var, width=5).pack(side="left", padx=(6, 4))
        tk.Label(pr, text="Uhr", bg=BG, fg="#666", font=(theme.FONT, 9)).pack(side="left")

        btnrow = tk.Frame(wrap, bg=BG)
        btnrow.pack(fill="x")
        tk.Button(btnrow, text="💾 Einstellungen speichern", command=self._einstellungen_speichern,
                  bg=ACCENT, fg="white", font=(theme.FONT, 10, "bold"),
                  padx=14, pady=5).pack(side="left")
        tk.Label(btnrow, text="  Gilt für neu erzeugte Dokumente. Leere Firmenfelder nutzen "
                 "den Standardtext.", bg=BG, fg="#999", font=(theme.FONT, 8)).pack(side="left")

    def _einstellungen_speichern(self):
        top = self.winfo_toplevel()
        stunde_text = self._set_stunde_var.get().strip()
        try:
            stunde = int(stunde_text)
            if not (0 <= stunde <= 23):
                raise ValueError
        except ValueError:
            messagebox.showwarning("Einstellungen", "Die Tagesabschluss-Uhrzeit muss eine "
                                   "Zahl zwischen 0 und 23 sein.", parent=top)
            return
        werte = {
            "firma_name": self._set_firma_name.get().strip(),
            "firma_adresse": self._set_firma_adresse_txt.get("1.0", "end").strip(),
            "firma_kontakt": self._set_firma_kontakt.get().strip(),
            "defekt_rechtstext": self._set_rechtstext_txt.get("1.0", "end").strip(),
            "tagesabschluss_stunde": str(stunde),
        }
        try:
            einstellungen.set_many(self.db_path, werte)
        except Exception as e:
            messagebox.showerror("Einstellungen", _T('Konnte nicht gespeichert werden:\n{p0}', p0=e), parent=top)
            return
        self._log("Einstellungen geändert", details="Firmendaten/Texte/Tagesabschluss-Uhrzeit")
        # Auto-Tagesabschluss mit (ggf. neuer) Uhrzeit neu planen.
        if getattr(self, "_auto_ta_job", None):
            try:
                self.after_cancel(self._auto_ta_job)
            except Exception:
                pass
        self._auto_tagesabschluss_schedule()
        messagebox.showinfo("Einstellungen", "Gespeichert.", parent=top)

    # =============================================================== Auswertung
    def _build_auswertung(self, parent):
        nb = ttk.Notebook(parent, style="Kasse.TNotebook")
        nb.pack(fill="both", expand=True, padx=8, pady=(10, 8))
        umsatz = tk.Frame(nb, bg=BG)
        tagesabschluss = tk.Frame(nb, bg=BG)
        verfall = tk.Frame(nb, bg=BG)
        inventur = tk.Frame(nb, bg=BG)
        nb.add(umsatz, text="  Umsatz  ")
        nb.add(tagesabschluss, text="  Tagesabschluss  ")
        nb.add(verfall, text="  Verfall  ")
        nb.add(inventur, text="  Inventur  ")
        self._build_aw_umsatz(umsatz)
        self._build_aw_tagesabschluss(tagesabschluss)
        self._build_aw_verfall(verfall)
        self._build_aw_inventur(inventur)

    def _build_aw_umsatz(self, parent):
        bar = tk.Frame(parent, bg=BG)
        bar.pack(fill="x", padx=4, pady=(10, 4))
        tk.Label(bar, text="Gruppierung:", bg=BG, font=(theme.FONT, 9, "bold")).pack(side="left")
        self._aw_gran_var = tk.StringVar(value="Tag")
        cb = ttk.Combobox(bar, textvariable=self._aw_gran_var, width=8, state="readonly",
                          values=("Tag", "Monat", "Jahr"))
        cb.pack(side="left", padx=(4, 12))
        cb.bind("<<ComboboxSelected>>", lambda _e: self._refresh_aw_umsatz())
        tk.Label(bar, text="von:", bg=BG, font=(theme.FONT, 9, "bold")).pack(side="left")
        self._aw_von_var = tk.StringVar()
        tk.Entry(bar, textvariable=self._aw_von_var, width=11).pack(side="left", padx=(4, 8))
        tk.Label(bar, text="bis:", bg=BG, font=(theme.FONT, 9, "bold")).pack(side="left")
        self._aw_bis_var = tk.StringVar()
        tk.Entry(bar, textvariable=self._aw_bis_var, width=11).pack(side="left", padx=(4, 4))
        tk.Label(bar, text="(JJJJ-MM-TT, optional)", bg=BG, fg="#999",
                 font=(theme.FONT, 8)).pack(side="left", padx=(0, 8))
        tk.Button(bar, text="Aktualisieren", command=self._refresh_aw_umsatz,
                  font=(theme.FONT, 9), padx=8, pady=2).pack(side="left")

        _, self._aw_umsatz_cards = self._metric_cards(parent, [
            ("anzahl", "Anzahl Verkäufe"), ("pakete", "Anzahl Packungen"),
            ("brutto", "APU Brutto"), ("rabatt", "Rabatt (Netto)"), ("netto", "APU Netto")])

        tf = tk.Frame(parent, bg=BG)
        tf.pack(fill="both", expand=True, padx=4, pady=(0, 6))
        cols = ("periode", "anzahl", "pakete", "brutto", "rabatt", "netto")
        heads = {"periode": "Zeitraum", "anzahl": "Anzahl Verkäufe",
                 "pakete": "Anzahl Packungen", "brutto": "APU Brutto",
                 "rabatt": "Rabatt (Netto)", "netto": "APU Netto"}
        tree = ttk.Treeview(tf, columns=cols, show="headings", style="Kasse.Treeview")
        for c in cols:
            tree.heading(c, text=heads[c])
            tree.column(c, width=110 if c == "periode" else 100,
                        anchor="w" if c == "periode" else "e")
        tree.pack(side="left", fill="both", expand=True)
        sb = tk.Scrollbar(tf, orient="vertical", command=tree.yview)
        sb.pack(side="right", fill="y")
        tree.configure(yscrollcommand=sb.set)
        _make_treeview_sortable(tree)
        tree.bind("<Double-1>", self._aw_umsatz_tagesabschluss_aus_zeile)
        self.aw_umsatz_tree = tree

        foot = tk.Frame(parent, bg=BG)
        foot.pack(fill="x", padx=4, pady=(0, 8))
        tk.Button(foot, text="📄 Tabelle als PDF",
                  command=lambda: self._run_report(
                      lambda: kasse_reports.umsatz_pdf(
                          self.db_path, self._aw_gran_var.get(),
                          self._aw_von_var.get().strip() or None,
                          self._aw_bis_var.get().strip() or None)),
                  bg=ACCENT, fg="white", font=(theme.FONT, 9, "bold"),
                  padx=10, pady=3).pack(side="left")
        tk.Label(foot, text="  (Doppelklick auf einen Tag öffnet den Tagesabschluss)",
                 bg=BG, fg="#999", font=(theme.FONT, 8)).pack(side="left")

    def _build_aw_tagesabschluss(self, parent):
        bar = tk.Frame(parent, bg=BG)
        bar.pack(fill="x", padx=8, pady=(12, 6))
        tk.Label(bar, text="Datum:", bg=BG, font=(theme.FONT, 10, "bold")).pack(side="left")
        self._ta_date = _make_date_entry(bar, default_today=True)
        self._ta_date.pack(side="left", padx=(6, 10))
        tk.Button(bar, text="Anzeigen", command=self._refresh_aw_tagesabschluss,
                  bg=ACCENT, fg="white", font=(theme.FONT, 9, "bold"),
                  padx=10, pady=2).pack(side="left")
        self.ta_nr_label = tk.Label(bar, text="", bg=BG, fg="#5a4b8a",
                                    font=(theme.FONT, 11, "bold"))
        self.ta_nr_label.pack(side="left", padx=(16, 0))

        # Kennzahl-Karten (Verkäufe, Packungen, APU, Rabatt, Umsatz).
        _, self._ta_cards = self._metric_cards(parent, [
            ("anzahl", "Menge der Verkäufe"), ("pakete", "Verkaufte Packungen"),
            ("brutto", "APU-Summe"), ("rabatt", "Rabatt-Summe"), ("netto", "Umsatz-Summe")])

        tf = tk.Frame(parent, bg=BG)
        tf.pack(fill="both", expand=True, padx=8, pady=(0, 6))
        cols = ("pzn", "artikel", "menge", "apu", "umsatz")
        heads = {"pzn": "PZN", "artikel": "Artikel", "menge": "Menge",
                 "apu": "APU-Summe", "umsatz": "Umsatz netto"}
        tree = ttk.Treeview(tf, columns=cols, show="headings", style="Kasse.Treeview")
        for c in cols:
            tree.heading(c, text=heads[c])
            tree.column(c, width=260 if c == "artikel" else 110,
                        anchor="w" if c in ("pzn", "artikel") else "e")
        tree.pack(side="left", fill="both", expand=True)
        sb = tk.Scrollbar(tf, orient="vertical", command=tree.yview)
        sb.pack(side="right", fill="y")
        tree.configure(yscrollcommand=sb.set)
        _make_treeview_sortable(tree)
        self.ta_tree = tree

        tk.Button(parent, text="🧾 Tagesabschluss als PDF (vergibt Nr.)",
                  command=self._tagesabschluss_pdf_aktuell,
                  bg="#1f7a4d", fg="white", font=(theme.FONT, 9, "bold"),
                  padx=10, pady=3).pack(anchor="w", padx=8, pady=(0, 8))

    def _refresh_aw_tagesabschluss(self):
        if not hasattr(self, "ta_tree"):
            return
        tag = _read_date(self._ta_date) or datetime.now().strftime("%Y-%m-%d")
        self._ta_tag = tag
        data = kasse_reports.tagesabschluss_data(self.db_path, tag)
        self._ta_cards["anzahl"].config(text=str(data["anzahl"]))
        self._ta_cards["pakete"].config(text=str(data["pakete"]))
        self._ta_cards["brutto"].config(text=_eur(data["brutto"]))
        self._ta_cards["rabatt"].config(text=_eur(data["rabatt"]))
        self._ta_cards["netto"].config(text=_eur(data["netto"]))
        nr = kasse_reports.tagesabschluss_nr(self.db_path, tag, assign=False)
        try:
            tag_disp = datetime.strptime(tag, "%Y-%m-%d").strftime("%d.%m.%Y")
        except ValueError:
            tag_disp = tag
        self.ta_nr_label.config(text=(f"Tagesabschluss Nr. {nr} · {tag_disp}" if nr
                                      else f"{tag_disp} · noch nicht erzeugt"))
        self.ta_tree.delete(*self.ta_tree.get_children())
        for p in data["positionen"]:
            self.ta_tree.insert("", "end", values=(
                p[0], p[1], int(p[2] or 0), _eur(p[3]), _eur(p[5])))

    def _tagesabschluss_pdf_aktuell(self):
        tag = getattr(self, "_ta_tag", None) or _read_date(self._ta_date) \
            or datetime.now().strftime("%Y-%m-%d")
        self._tagesabschluss_erzeugen(tag, zeige_leer=True)
        self._refresh_aw_tagesabschluss()

    def _build_aw_verfall(self, parent):
        bar = tk.Frame(parent, bg=BG)
        bar.pack(fill="x", padx=4, pady=(10, 4))
        tk.Label(bar, text="Zeitraum:", bg=BG, font=(theme.FONT, 9, "bold")).pack(side="left")
        self._aw_verfall_monate_var = tk.StringVar(value="3 Monate")
        cb = ttk.Combobox(bar, textvariable=self._aw_verfall_monate_var, width=10,
                          state="readonly",
                          values=("3 Monate", "6 Monate", "9 Monate", "12 Monate", "Alle"))
        cb.pack(side="left", padx=(4, 12))
        cb.bind("<<ComboboxSelected>>", lambda _e: self._refresh_aw_verfall())
        tk.Button(bar, text="🔄 Aktualisieren", command=self._refresh_aw_verfall,
                  font=(theme.FONT, 9), padx=8, pady=2).pack(side="right")
        tk.Label(parent, text="Rot = abgelaufen · Gelb = läuft im gewählten Zeitraum ab. "
                 "„Alle“ zeigt den ganzen Bestand.",
                 bg=BG, fg="#666", font=(theme.FONT, 9)).pack(anchor="w", padx=6, pady=(0, 2))
        _, self._aw_verfall_cards = self._metric_cards(parent, [
            ("ab", "Abgelaufen", "#a32d2d"), ("bald", "Läuft bald ab", "#a35a00"),
            ("bestand", "Bestand"), ("wert", "Verkaufswert"), ("lager", "Lagerwert")])

        tf = tk.Frame(parent, bg=BG)
        tf.pack(fill="both", expand=True, padx=4, pady=(0, 6))
        cols = ("pzn", "artikel", "charge", "verfall", "menge", "status")
        heads = {"pzn": "PZN", "artikel": "Artikel", "charge": "Charge",
                 "verfall": "Verfall", "menge": "Bestand", "status": "Status"}
        tree = ttk.Treeview(tf, columns=cols, show="headings", style="Kasse.Treeview")
        for c in cols:
            tree.heading(c, text=heads[c])
            tree.column(c, width=230 if c == "artikel" else 95, anchor="w")
        tree.pack(side="left", fill="both", expand=True)
        sb = tk.Scrollbar(tf, orient="vertical", command=tree.yview)
        sb.pack(side="right", fill="y")
        tree.configure(yscrollcommand=sb.set)
        _make_treeview_sortable(tree)
        tree.tag_configure("abgelaufen", background="#fadbd8", foreground="#922")
        tree.tag_configure("bald", background="#fcf3cf", foreground="#7a6000")
        self.aw_verfall_tree = tree

        tk.Button(parent, text="📄 Verfall-Report als PDF",
                  command=lambda: self._run_report(lambda: kasse_reports.verfall_pdf(
                      self.db_path, monate=self._aw_verfall_monate_var.get())),
                  bg=ACCENT, fg="white", font=(theme.FONT, 9, "bold"),
                  padx=10, pady=3).pack(anchor="w", padx=4, pady=(0, 8))

    def _build_aw_inventur(self, parent):
        bar = tk.Frame(parent, bg=BG)
        bar.pack(fill="x", padx=4, pady=(10, 4))
        tk.Label(bar, text="Inventur-Zählliste zum Ausdrucken (Soll-Bestand je Charge "
                 "+ Spalte zum Eintragen der Zählmenge).", bg=BG, fg="#666",
                 font=(theme.FONT, 9)).pack(side="left")
        tk.Button(bar, text="🔄 Aktualisieren", command=self._refresh_aw_inventur,
                  font=(theme.FONT, 9), padx=8, pady=2).pack(side="right")
        _, self._aw_inv_cards = self._metric_cards(parent, [
            ("positionen", "Positionen"), ("menge", "Gesamtmenge"),
            ("wert", "Verkaufswert"), ("lager", "Lagerwert")])

        tf = tk.Frame(parent, bg=BG)
        tf.pack(fill="both", expand=True, padx=4, pady=(0, 6))
        cols = ("pzn", "artikel", "charge", "verfall", "menge", "ek", "lagerwert")
        heads = {"pzn": "PZN", "artikel": "Artikel", "charge": "Charge",
                 "verfall": "Verfall", "menge": "Soll-Bestand", "ek": "EK €",
                 "lagerwert": "Lagerwert €"}
        tree = ttk.Treeview(tf, columns=cols, show="headings", style="Kasse.Treeview")
        for c in cols:
            tree.heading(c, text=heads[c])
            tree.column(c, width=230 if c == "artikel" else 95,
                        anchor="e" if c in ("ek", "lagerwert") else "w")
        tree.pack(side="left", fill="both", expand=True)
        sb = tk.Scrollbar(tf, orient="vertical", command=tree.yview)
        sb.pack(side="right", fill="y")
        tree.configure(yscrollcommand=sb.set)
        _make_treeview_sortable(tree)
        self.aw_inv_tree = tree

        tk.Button(parent, text="📄 Inventurliste als PDF",
                  command=lambda: self._run_report(lambda: kasse_reports.inventur_pdf(self.db_path)),
                  bg=ACCENT, fg="white", font=(theme.FONT, 9, "bold"),
                  padx=10, pady=3).pack(anchor="w", padx=4, pady=(0, 8))

    def _refresh_auswertung(self):
        self._refresh_aw_umsatz()
        self._refresh_aw_tagesabschluss()
        self._refresh_aw_verfall()
        self._refresh_aw_inventur()

    def _refresh_aw_umsatz(self):
        if not hasattr(self, "aw_umsatz_tree"):
            return
        rows = kasse_reports.umsatz_rows(
            self.db_path, self._aw_gran_var.get(),
            self._aw_von_var.get().strip() or None, self._aw_bis_var.get().strip() or None)
        self.aw_umsatz_tree.delete(*self.aw_umsatz_tree.get_children())
        for r in rows:
            self.aw_umsatz_tree.insert("", "end", values=(
                r["periode"], r["anzahl"], r["pakete"], _eur(r["brutto"]),
                _eur(r["rabatt"]), _eur(r["netto"])))
        c = self._aw_umsatz_cards
        c["anzahl"].config(text=str(sum(r["anzahl"] for r in rows)))
        c["pakete"].config(text=str(sum(r["pakete"] for r in rows)))
        c["brutto"].config(text=_eur(sum(r["brutto"] for r in rows)))
        c["rabatt"].config(text=_eur(sum(r["rabatt"] for r in rows)))
        c["netto"].config(text=_eur(sum(r["netto"] for r in rows)))

    def _refresh_aw_verfall(self):
        if not hasattr(self, "aw_verfall_tree"):
            return
        monate = self._aw_verfall_monate_var.get()
        warn_tage = kasse_reports.VERFALL_MONATE.get(monate, 90)
        rows = kasse_reports.verfall_rows(
            self.db_path, warn_tage=warn_tage if warn_tage is not None else 90)
        if warn_tage is not None:
            # Nur abgelaufene + im gewaehlten Zeitraum ablaufende anzeigen.
            rows = [r for r in rows if r["status"] in ("abgelaufen", "bald")]
        self.aw_verfall_tree.delete(*self.aw_verfall_tree.get_children())
        ab = bald = 0
        for r in rows:
            if r["status"] == "abgelaufen":
                tags, stat, ab = ("abgelaufen",), "abgelaufen", ab + 1
            elif r["status"] == "bald":
                tags, stat, bald = ("bald",), f"in {r['tage']} Tagen", bald + 1
            elif r["status"] == "ok":
                tags, stat = (), "ok"
            else:
                tags, stat = (), "—"
            self.aw_verfall_tree.insert("", "end", tags=tags, values=(
                r["pzn"], r["artikelname"], r["charge"] or "—", r["verfall"] or "—",
                r["menge"], stat))
        c = self._aw_verfall_cards
        c["ab"].config(text=str(ab))
        c["bald"].config(text=str(bald))
        c["bestand"].config(text=str(sum(r["menge"] or 0 for r in rows)))
        c["wert"].config(text=_eur(sum(r["wert"] for r in rows)))
        c["lager"].config(text=_eur(sum(r["lagerwert"] for r in rows)))

    def _refresh_aw_inventur(self):
        if not hasattr(self, "aw_inv_tree"):
            return
        rows = kasse_reports.inventur_rows(self.db_path)
        self.aw_inv_tree.delete(*self.aw_inv_tree.get_children())
        for r in rows:
            self.aw_inv_tree.insert("", "end", values=(
                r["pzn"], r["artikelname"], r["charge"] or "—", r["verfall"] or "—", r["menge"],
                _eur(r["ek"]), _eur(r["lagerwert"]) if r["ek"] is not None else "—"))
        c = self._aw_inv_cards
        c["positionen"].config(text=str(len(rows)))
        c["menge"].config(text=str(sum(r["menge"] or 0 for r in rows)))
        c["wert"].config(text=_eur(sum(r["wert"] for r in rows)))
        c["lager"].config(text=_eur(sum(r["lagerwert"] for r in rows)))

    def _run_report(self, make_fn):
        """Erzeugt ein Report-PDF (make_fn -> Pfad) und oeffnet es. Fehlt fpdf2,
        gibt es einen verstaendlichen Hinweis statt eines Tracebacks."""
        top = self.winfo_toplevel()
        try:
            path = make_fn()
        except ImportError:
            messagebox.showerror("PDF", "Für PDF-Reports wird die Bibliothek „fpdf2“ benötigt "
                                 "(pip install fpdf2).", parent=top)
            return
        except Exception as e:
            messagebox.showerror("PDF", _T('Report konnte nicht erstellt werden:\n{p0}', p0=e), parent=top)
            return
        self._log("Report erzeugt", details=str(Path(path).name))
        try:
            os.startfile(str(path))  # type: ignore[attr-defined]
        except Exception:
            messagebox.showinfo("PDF", _T('Report gespeichert:\n{p0}', p0=path), parent=top)

    def _aw_umsatz_tagesabschluss_aus_zeile(self, _e=None):
        """Doppelklick auf eine Tages-Zeile -> Tagesabschluss-PDF fuer diesen Tag."""
        sel = self.aw_umsatz_tree.selection()
        if not sel:
            return
        periode = self.aw_umsatz_tree.item(sel[0], "values")[0]
        # Nur bei Tages-Granularitaet ist die Periode ein vollstaendiges Datum.
        if len(str(periode)) == 10 and str(periode)[4] == "-":
            self._tagesabschluss_erzeugen(str(periode), zeige_leer=True)
        else:
            messagebox.showinfo("Tagesabschluss", "Bitte zuerst auf „Tag“ gruppieren, dann den "
                                "gewünschten Tag doppelklicken.", parent=self.winfo_toplevel())

    def _tagesabschluss_erzeugen(self, tag, zeige_leer=False, oeffnen=True):
        """Erzeugt den Tagesabschluss-PDF fuer 'tag' (YYYY-MM-DD). Bei zeige_leer=False
        wird nichts erzeugt, wenn es an dem Tag keinen Umsatz gab (fuer den
        automatischen Lauf)."""
        data = kasse_reports.tagesabschluss_data(self.db_path, tag)
        if not zeige_leer and data["anzahl"] == 0:
            return None
        if zeige_leer and data["anzahl"] == 0:
            if not messagebox.askyesno(
                    "Tagesabschluss",
                    _T('Für {p0} sind keine Verkäufe erfasst.\nTrotzdem einen (leeren) Tagesabschluss erzeugen?', p0=tag), parent=self.winfo_toplevel()):
                return None
        if oeffnen:
            self._run_report(lambda: kasse_reports.tagesabschluss_pdf(self.db_path, tag))
        else:
            try:
                kasse_reports.tagesabschluss_pdf(self.db_path, tag)
                self._log("Tagesabschluss (automatisch)", details=tag)
            except Exception:
                return None
        return tag

    AUTO_TAGESABSCHLUSS_STUNDE = 18   # Standard-Uhrzeit (per Einstellungen aenderbar)

    def _ta_stunde(self):
        """Uhrzeit des automatischen Tagesabschlusses aus den Einstellungen (0-23)."""
        s = einstellungen.get_int(self.db_path, "tagesabschluss_stunde",
                                  self.AUTO_TAGESABSCHLUSS_STUNDE)
        return s if 0 <= s <= 23 else self.AUTO_TAGESABSCHLUSS_STUNDE

    def _auto_tagesabschluss_setup(self):
        """Beim Start verpasste Tagesabschluesse nachholen (Tage mit Umsatz ohne
        gespeichertes PDF - Vortage immer, heute nur wenn 18 Uhr schon vorbei) und
        den naechsten automatischen Lauf exakt auf 18:00 Uhr planen."""
        try:
            jetzt = datetime.now()
            heute = jetzt.strftime("%Y-%m-%d")
            faellig = jetzt.hour >= self._ta_stunde()
            for tag in kasse_reports.tage_mit_umsatz(self.db_path):
                # Erledigt = es gibt bereits eine Tagesabschluss-Nr fuer den Tag.
                if kasse_reports.tagesabschluss_nr(self.db_path, tag) is not None:
                    continue
                if tag < heute or (tag == heute and faellig):
                    self._tagesabschluss_erzeugen(tag, zeige_leer=False, oeffnen=False)
        except Exception:
            pass
        self._auto_tagesabschluss_schedule()

    def _auto_tagesabschluss_schedule(self):
        """Plant den naechsten automatischen Lauf punktgenau auf 18:00 Uhr."""
        from datetime import timedelta
        try:
            now = datetime.now()
            ziel = now.replace(hour=self._ta_stunde(), minute=0,
                               second=0, microsecond=0)
            if now >= ziel:
                ziel += timedelta(days=1)   # 18 Uhr ist heute schon vorbei -> morgen
            ms = max(1000, int((ziel - now).total_seconds() * 1000))
            self._auto_ta_job = self.after(ms, self._auto_tagesabschluss_run)
        except Exception:
            pass

    def _auto_tagesabschluss_run(self):
        """Um 18 Uhr: heutigen Tagesabschluss erzeugen (vergibt die Nr.), wenn es
        Umsatz gab und noch keiner existiert; danach den naechsten Tag planen."""
        try:
            heute = datetime.now().strftime("%Y-%m-%d")
            if kasse_reports.tagesabschluss_nr(self.db_path, heute) is None:
                self._tagesabschluss_erzeugen(heute, zeige_leer=False, oeffnen=False)
        except Exception:
            pass
        self._auto_tagesabschluss_schedule()

    # =============================================================== Protokoll
    def _build_protokoll(self, parent):
        bar = tk.Frame(parent, bg=BG)
        bar.pack(fill="x", padx=8, pady=(10, 4))
        tk.Label(bar, text="Suche (Mitarbeiter / Aktion / Kunde / Auftrag-Nr / Detail):",
                 bg=BG, font=(theme.FONT, 9, "bold")).pack(side="left")
        self._log_filter_var = tk.StringVar()
        e = tk.Entry(bar, textvariable=self._log_filter_var, width=24)
        e.pack(side="left", padx=(6, 8))
        e.bind("<KeyRelease>", lambda _e: self._refresh_log())
        tk.Label(bar, text="Aktion:", bg=BG, font=(theme.FONT, 9, "bold")).pack(side="left")
        self._log_aktion_var = tk.StringVar(value="Alle")
        self._log_aktion_cmb = ttk.Combobox(bar, textvariable=self._log_aktion_var, width=22,
                                            state="readonly", values=("Alle",))
        self._log_aktion_cmb.pack(side="left", padx=(4, 0))
        self._log_aktion_var.trace_add("write", lambda *_: self._refresh_log())
        tk.Button(bar, text="🔄 Aktualisieren", command=self._refresh_log,
                  font=(theme.FONT, 9), padx=8, pady=2).pack(side="right")
        tk.Button(bar, text="📄 Als PDF", command=self._protokoll_pdf_export,
                  font=(theme.FONT, 9), padx=8, pady=2).pack(side="right", padx=(0, 6))
        self.log_info = tk.Label(bar, text="", bg=BG, fg=ACCENT, font=(theme.FONT, 9, "bold"))
        self.log_info.pack(side="right", padx=(0, 12))
        tk.Label(parent, text="Wer hat was wann geändert (neueste oben). Spalten sortierbar – "
                              "z. B. nach Mitarbeiter oder Auftrag, um Fehler/Änderungen zu finden.",
                 bg=BG, fg="#666", font=(theme.FONT, 9)).pack(anchor="w", padx=8, pady=(0, 4))

        tf = tk.Frame(parent, bg=BG)
        tf.pack(fill="both", expand=True, padx=8, pady=(0, 8))
        cols = ("zeit", "bearbeiter", "aktion", "auftrag", "kunde", "details")
        heads = {"zeit": "Zeitpunkt", "bearbeiter": "Mitarbeiter", "aktion": "Aktion",
                 "auftrag": "Auftrag", "kunde": "Kunde", "details": "Details"}
        tree = ttk.Treeview(tf, columns=cols, show="headings", style="Kasse.Treeview")
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
        self.log_info.config(text=_T('{p0} Einträge{p1}', p0=gezeigt, p1=suffix))

    def _protokoll_pdf_export(self):
        """Exportiert die aktuell angezeigten (gefilterten) Protokoll-Zeilen als PDF -
        z. B. nach Auftrag-Nr oder Mitarbeiter gefiltert: wer hat was gemacht."""
        zeilen = [self.log_tree.item(i, "values") for i in self.log_tree.get_children()]
        if not zeilen:
            messagebox.showinfo("Protokoll", "Keine Einträge zum Exportieren.",
                                parent=self.winfo_toplevel())
            return
        filt = self._log_filter_var.get().strip()
        aktion_f = self._log_aktion_var.get() if hasattr(self, "_log_aktion_var") else "Alle"
        teile = []
        if aktion_f and aktion_f != "Alle":
            teile.append(f"Aktion: {aktion_f}")
        if filt:
            teile.append(f"Filter: {filt}")
        untertitel = " · ".join(teile)
        self._run_report(lambda: kasse_reports.protokoll_pdf(
            zeilen, untertitel=untertitel or None))

    # ====================================================== Bestandskorrektur
    def _lager_korrektur_dialog(self, pzn, charge, verfall, menge):
        """Bestand einer Lagerzeile (PZN/Charge/Verfall) manuell korrigieren.
        Grund ist Pflicht (revisionssicher protokolliert). Buchen von Ware
        erfolgt in der App 'Wareneingang & Retouren'. -> True wenn geaendert."""
        top = self.winfo_toplevel()
        neu = simpledialog.askinteger(
            "Bestand korrigieren",
            _T('Neuer Bestand für\nPZN {p0} · Charge {p1} · Verf {p2}\n(0 entfernt die Zeile):', p0=pzn, p1=charge or '—', p2=verfall or '—'),
            parent=top, initialvalue=int(menge), minvalue=0)
        if neu is None or neu == int(menge):
            return False
        # Pflicht-Kommentar: WARUM wird manuell korrigiert? (Abbrechen = keine Aenderung.)
        grund = simpledialog.askstring(
            "Bestandskorrektur – Grund (Pflicht)",
            f"PZN {pzn} · Charge {charge or '–'} · Verf {verfall or '–'}\n"
            f"Bestand von {menge} auf {neu} ändern"
            + (" (0 entfernt die Lagerzeile).\n\n" if neu == 0 else ".\n\n")
            + "Bitte den Grund eingeben (Pflicht):", parent=top)
        if grund is None:
            return False
        grund = grund.strip()
        if not grund:
            messagebox.showwarning("Bestandskorrektur",
                                   "Ein Grund (Kommentar) ist Pflicht – Korrektur abgebrochen.",
                                   parent=top)
            return False
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
                  f"{menge} -> {neu} · Grund: {grund}")
        return True

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
            _T('Quelle: {p0} · {p1} Zeilen gelesen.\nErkannte Spalten: {p2}\n\nNeu: {p3} · Aktualisiert: {p4} · Übersprungen: {p5}', p0=r['quelle'], p1=r['gelesen'], p2=', '.join(r['spalten']), p3=r['neu'], p4=r['aktualisiert'], p5=r['uebersprungen']), parent=top)

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
            _T('{p0}-Import', p0=was),
            _T('Quelle: {p0} · {p1} Zeilen gelesen.\nErkannte Zusatzspalten: {p2}\n\nAngelegt: {p3} {p4} mit {p5} Position(en) · Übersprungen: {p6}{p7}', p0=r['quelle'], p1=r['gelesen'], p2=', '.join(r['spalten']) or '–', p3=r['bestellungen'], p4=was, p5=r['positionen'], p6=r['uebersprungen'], p7=unklar), parent=top)

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


def run_standalone():
    """Startet die Kasse als eigenstaendiges Fenster (eigenes Taskleisten-Icon).
    Wird von start_kasse.py und von NMGone.exe --kasse genutzt."""
    if os.name == "nt":
        try:
            import ctypes
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("NMG.Kasse")
        except Exception:
            pass
    try:
        from .migrations import run_migrations
        run_migrations()
    except Exception:
        pass
    root = tk.Tk()
    root.title(f"NMG Kasse{DEMO_SUFFIX}")
    root.geometry("1040x660")          # Groesse im wiederhergestellten Zustand
    root.minsize(860, 580)
    # Im Vollbild (maximiert) starten. 'zoomed' = Windows; sonst -zoomed/Bildschirmgroesse.
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
        root.iconbitmap(str(ASSETS_DIR / "kasse.ico"))
    except Exception:
        pass
    KassePanel(root, on_close=root.destroy).pack(fill="both", expand=True)
    tour.maybe_show(root, "kasse", tour.kasse_steps())
    root.mainloop()
