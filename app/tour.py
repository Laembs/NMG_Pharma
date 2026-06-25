"""Gefuehrte Einfuehrungs-Tour ("Wizard") fuer die NMGone-Programme.

Eine Tour ist eine Folge von Folien (TourStep): Titel + Text (+ optional Bild).
Der Dialog fuehrt den Anwender mit Zurueck / Weiter / Ueberspringen durch.

Drei Abschalt-Ebenen (Reihenfolge der Auswertung):

  1. HART AUS (pro Build/Version):  version.json -> "tour_enabled": false
     ODER die Konstanten HARD_DISABLED hier. Wenn aus, erscheint gar nichts
     und es sollte auch kein Menue-Eintrag verdrahtet werden.

  2. ANWENDER AUS:  Flag in der DB (tbl_tour_state, key '__user_enabled__').
     Setzt der Anwender "nicht mehr automatisch zeigen", startet keine Tour
     mehr von selbst. Per Menue laesst sie sich trotzdem manuell aufrufen.

  3. PRO VERSION EINMAL:  jede Tour merkt sich (tour_id + Version), ob sie in
     dieser Version schon gesehen wurde. Neue Version -> Tour kommt wieder.

Verwendung in einer App (nach dem Aufbau der Oberflaeche):

    from . import tour
    tour.maybe_show(self, "nmgone", tour.nmgone_steps())

und optional ein Menue-/Hilfe-Eintrag, der erzwungen startet:

    tour.start(self, "nmgone", tour.nmgone_steps())
"""
from __future__ import annotations

import json
import sqlite3
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

import tkinter as tk

from . import theme
from .config import BASE_DIR, DB_PATH


# --------------------------------------------------------------------------
# Ebene 1: HART AUS (im Code festverdrahtet, gewinnt ueber alles andere)
# --------------------------------------------------------------------------
# Auf True setzen, um die Tour fuer einen Build komplett zu entfernen.
HARD_DISABLED = False


def _version() -> str:
    """Aktuelle Anzeige-Version aus version.json (z. B. 'V2.0 SP4').

    Dient als Schluessel fuer "pro Version einmal". Faellt auf 'app'+'version'
    zurueck und im Fehlerfall auf '?', damit die Tour nie crasht.
    """
    try:
        data = json.loads((BASE_DIR / "version.json").read_text(encoding="utf-8"))
        return str(data.get("version_display") or data.get("version") or "?")
    except Exception:
        return "?"


def _tour_enabled_in_build() -> bool:
    """version.json -> "tour_enabled" (Default True). False = Build-weit aus."""
    if HARD_DISABLED:
        return False
    try:
        data = json.loads((BASE_DIR / "version.json").read_text(encoding="utf-8"))
        return bool(data.get("tour_enabled", True))
    except Exception:
        return True


# --------------------------------------------------------------------------
# Persistenz (gemeinsame DB, damit NMGone + Kasse-/Faktura-.exe sich einig sind)
# --------------------------------------------------------------------------
_USER_FLAG_KEY = "__user_enabled__"


def _connect():
    con = sqlite3.connect(DB_PATH)
    con.execute(
        "CREATE TABLE IF NOT EXISTS tbl_tour_state("
        "tour_id TEXT NOT NULL, version TEXT NOT NULL, seen_at TEXT, "
        "PRIMARY KEY(tour_id, version))")
    return con


def user_enabled() -> bool:
    """Ebene 2: Hat der Anwender die automatische Tour erlaubt? (Default ja)."""
    try:
        with _connect() as con:
            row = con.execute(
                "SELECT seen_at FROM tbl_tour_state WHERE tour_id=? AND version=?",
                (_USER_FLAG_KEY, "*")).fetchone()
    except sqlite3.Error:
        return True
    if row is None:
        return True
    return str(row[0]) != "0"


def set_user_enabled(enabled: bool) -> None:
    try:
        with _connect() as con:
            con.execute(
                "INSERT INTO tbl_tour_state(tour_id, version, seen_at) VALUES(?,?,?) "
                "ON CONFLICT(tour_id, version) DO UPDATE SET seen_at=excluded.seen_at",
                (_USER_FLAG_KEY, "*", "1" if enabled else "0"))
    except sqlite3.Error:
        pass


def _seen(tour_id: str, version: str) -> bool:
    try:
        with _connect() as con:
            row = con.execute(
                "SELECT 1 FROM tbl_tour_state WHERE tour_id=? AND version=?",
                (tour_id, version)).fetchone()
    except sqlite3.Error:
        return False
    return row is not None


def _mark_seen(tour_id: str, version: str) -> None:
    try:
        with _connect() as con:
            con.execute(
                "INSERT OR IGNORE INTO tbl_tour_state(tour_id, version, seen_at) "
                "VALUES(?,?,?)", (tour_id, version, datetime.now().isoformat(timespec="seconds")))
    except sqlite3.Error:
        pass


# --------------------------------------------------------------------------
# Tour-Inhalt
# --------------------------------------------------------------------------
@dataclass
class TourStep:
    titel: str
    text: str
    bild: str | None = None   # optionaler Pfad zu einem PNG


class _Wizard(tk.Toplevel):
    """Modaler Schritt-fuer-Schritt-Dialog ueber dem Elternfenster."""

    def __init__(self, parent, tour_id: str, steps: list[TourStep],
                 show_optout: bool = True):
        super().__init__(parent)
        self._tour_id = tour_id
        self._steps = steps
        self._idx = 0
        self._img_ref = None  # PhotoImage am Leben halten

        self.title("Einfuehrung")
        self.configure(bg=theme.CARD)
        self.transient(parent)
        self.resizable(False, False)
        try:
            self.iconbitmap("")  # Standard-Icon vermeiden, falls keins gesetzt
        except Exception:
            pass

        # Kopf
        head = tk.Frame(self, bg=theme.PRIMARY)
        head.pack(fill="x")
        self._title_lbl = tk.Label(head, text="", bg=theme.PRIMARY, fg="white",
                                   font=theme.H2, anchor="w", padx=20, pady=14)
        self._title_lbl.pack(fill="x")

        # Inhalt
        body = tk.Frame(self, bg=theme.CARD, padx=24, pady=18)
        body.pack(fill="both", expand=True)
        self._img_lbl = tk.Label(body, bg=theme.CARD)
        self._text_lbl = tk.Label(body, text="", bg=theme.CARD, fg=theme.INK,
                                  font=theme.BODY, justify="left", wraplength=460,
                                  anchor="nw")
        self._text_lbl.pack(fill="both", expand=True)

        # Fortschritt
        self._dots = tk.Label(body, text="", bg=theme.CARD, fg=theme.FAINT,
                              font=theme.SMALL, anchor="w")
        self._dots.pack(fill="x", pady=(14, 0))

        # Fuss: optionaler Opt-out + Navigation
        foot = tk.Frame(self, bg=theme.CARD_ALT, padx=20, pady=14)
        foot.pack(fill="x")

        self._optout_var = tk.IntVar(value=0)
        if show_optout:
            tk.Checkbutton(
                foot, text="Diese Einfuehrung nicht mehr automatisch zeigen",
                variable=self._optout_var, bg=theme.CARD_ALT, fg=theme.MUTED,
                activebackground=theme.CARD_ALT, font=theme.SMALL,
                selectcolor=theme.CARD, anchor="w", bd=0, highlightthickness=0,
            ).pack(side="left")

        self._next_btn = _btn(foot, "Weiter ›", self._next, primary=True)
        self._next_btn.pack(side="right")
        self._back_btn = _btn(foot, "‹ Zurueck", self._back)
        self._back_btn.pack(side="right", padx=(0, 8))
        _btn(foot, "Ueberspringen", self._skip).pack(side="right", padx=(0, 8))

        self._render()
        self.protocol("WM_DELETE_WINDOW", self._skip)
        self.bind("<Escape>", lambda e: self._skip())
        self.bind("<Return>", lambda e: self._next())
        self._center(parent)
        self.grab_set()
        self.focus_set()

    # -- Navigation -----------------------------------------------------
    def _render(self):
        step = self._steps[self._idx]
        self._title_lbl.config(text=step.titel)
        self._text_lbl.config(text=step.text)

        self._img_lbl.pack_forget()
        self._img_ref = None
        if step.bild:
            p = Path(step.bild)
            if p.exists():
                try:
                    self._img_ref = tk.PhotoImage(file=str(p))
                    self._img_lbl.config(image=self._img_ref)
                    self._img_lbl.pack(before=self._text_lbl, pady=(0, 14))
                except Exception:
                    self._img_ref = None

        n = len(self._steps)
        self._dots.config(text=f"Schritt {self._idx + 1} von {n}   "
                               + "  ".join("●" if i == self._idx else "○"
                                           for i in range(n)))
        self._back_btn.config(state="normal" if self._idx > 0 else "disabled")
        self._next_btn.config(text="Fertig ✓" if self._idx == n - 1 else "Weiter ›")

    def _next(self):
        if self._idx < len(self._steps) - 1:
            self._idx += 1
            self._render()
        else:
            self._finish()

    def _back(self):
        if self._idx > 0:
            self._idx -= 1
            self._render()

    def _skip(self):
        self._finish()

    def _finish(self):
        # In dieser Version als gesehen markieren + ggf. Anwender-Opt-out setzen.
        _mark_seen(self._tour_id, _version())
        if self._optout_var.get():
            set_user_enabled(False)
        try:
            self.grab_release()
        except Exception:
            pass
        self.destroy()

    def _center(self, parent):
        self.update_idletasks()
        w, h = self.winfo_reqwidth(), self.winfo_reqheight()
        try:
            px, py = parent.winfo_rootx(), parent.winfo_rooty()
            pw, ph = parent.winfo_width(), parent.winfo_height()
            x = px + max(0, (pw - w) // 2)
            y = py + max(0, (ph - h) // 3)
        except Exception:
            x = (self.winfo_screenwidth() - w) // 2
            y = (self.winfo_screenheight() - h) // 3
        self.geometry(f"+{x}+{y}")


def _btn(parent, text, command, primary=False):
    if primary:
        opts = dict(bg=theme.PRIMARY, fg="white", activebackground=theme.PRIMARY_DARK,
                    activeforeground="white")
    else:
        opts = dict(bg="#EDF1F6", fg=theme.PRIMARY, activebackground="#E2E9F1",
                    activeforeground=theme.PRIMARY)
    return tk.Button(parent, text=text, command=command, relief="flat", bd=0,
                     font=(theme.FONT, 10, "bold"), padx=16, pady=7, cursor="hand2",
                     **opts)


# --------------------------------------------------------------------------
# Oeffentliche API
# --------------------------------------------------------------------------
def maybe_show(parent, tour_id: str, steps: list[TourStep], delay_ms: int = 400) -> None:
    """Startet die Tour automatisch, wenn alle drei Ebenen es erlauben.

    Wertet aus: Build-Schalter (1) -> Anwender-Flag (2) -> pro Version einmal (3).
    Wird verzoegert per .after() gestartet, damit das Hauptfenster zuerst steht.
    """
    if not steps or not _tour_enabled_in_build() or not user_enabled():
        return
    if _seen(tour_id, _version()):
        return

    def _go():
        try:
            _Wizard(parent, tour_id, steps, show_optout=True)
        except Exception:
            pass

    try:
        parent.after(delay_ms, _go)
    except Exception:
        _go()


def start(parent, tour_id: str, steps: list[TourStep]) -> None:
    """Erzwingt die Tour (z. B. aus dem Hilfe-Menue), ignoriert "schon gesehen".

    Respektiert weiterhin den harten Build-Schalter: ist die Tour Build-weit
    aus, sollte ohnehin kein Menue-Eintrag verdrahtet sein.
    """
    if not steps or not _tour_enabled_in_build():
        return
    try:
        _Wizard(parent, tour_id, steps, show_optout=False)
    except Exception:
        pass


def reset(tour_id: str | None = None) -> None:
    """Setzt den "gesehen"-Status zurueck (zum Testen / "Tour erneut zeigen")."""
    try:
        with _connect() as con:
            if tour_id is None:
                con.execute("DELETE FROM tbl_tour_state WHERE tour_id <> ?",
                            (_USER_FLAG_KEY,))
            else:
                con.execute("DELETE FROM tbl_tour_state WHERE tour_id=?", (tour_id,))
    except sqlite3.Error:
        pass


# --------------------------------------------------------------------------
# Tour-Inhalte je Programm (Start-Fassung - Texte frei anpassbar)
# --------------------------------------------------------------------------
def nmgone_steps() -> list[TourStep]:
    return [
        TourStep("Willkommen bei NMGone",
                 "Diese kurze Einfuehrung zeigt dir in wenigen Schritten, wie du "
                 "dich im Programm zurechtfindest. Du kannst sie jederzeit mit "
                 "„Ueberspringen“ beenden."),
        TourStep("Navigation",
                 "Links findest du die Bereiche des Programms – z. B. "
                 "Bedarfsanalyse, Auswertungen, Kunden und Personal. Ein Klick "
                 "wechselt den Bereich."),
        TourStep("Bedarfsanalyse",
                 "Das Herzstueck: Hier liest du Apotheken-Daten ein und erzeugst "
                 "Auswertungen sowie den Kurzbericht. Die Schritte fuehren dich "
                 "von oben nach unten durch."),
        TourStep("Hilfe & Einstellungen",
                 "Ueber das Hilfe-Menue erreichst du das bebilderte Handbuch und "
                 "kannst diese Einfuehrung jederzeit erneut starten. In den "
                 "Einstellungen schaltest du sie auch ganz ab."),
        TourStep("Los geht’s",
                 "Das war’s fuer den Einstieg. Probiere die Bereiche aus – "
                 "bei Fragen hilft das Handbuch im Hilfe-Menue weiter."),
    ]


def kasse_steps() -> list[TourStep]:
    return [
        TourStep("Willkommen in der NMG Kasse",
                 "Diese Einfuehrung zeigt dir den Ablauf: vom Lagerbestand bis "
                 "zum Verkauf."),
        TourStep("Artikel & Lager",
                 "Unter 'Artikel' siehst du Artikelstamm und aktuellen "
                 "Lagerbestand. Per Doppelklick auf eine Charge korrigierst du "
                 "den Bestand. Ware angenommen wird in der App 'Wareneingang & "
                 "Retouren'."),
        TourStep("Verkauf",
                 "Artikel suchen, in den Warenkorb legen, kassieren – der "
                 "Tagesabschluss fasst die Verkaeufe zusammen."),
        TourStep("Fertig",
                 "Du kannst diese Einfuehrung im Menue jederzeit erneut starten "
                 "oder in den Einstellungen abschalten."),
    ]


def faktura_steps() -> list[TourStep]:
    return [
        TourStep("Willkommen in NMG Faktura",
                 "Hier erstellst du Rechnungen und Gutschriften. Die Einfuehrung "
                 "zeigt die wichtigsten Schritte."),
        TourStep("Beleg anlegen",
                 "Waehle Kunde und Positionen, das Programm berechnet Summen und "
                 "Steuer automatisch."),
        TourStep("Layout & Druck",
                 "Ueber den Layout-Editor passt du das Aussehen an und erzeugst "
                 "das fertige PDF."),
        TourStep("Fertig",
                 "Diese Einfuehrung laesst sich im Menue erneut starten oder ganz "
                 "abschalten."),
    ]


def personal_steps() -> list[TourStep]:
    return [
        TourStep("Willkommen im Mitarbeiter-Board",
                 "Hier verwaltest du Mitarbeiter, Arbeitsbereiche und "
                 "Abwesenheiten."),
        TourStep("Mitarbeiter & Felder",
                 "Lege Mitarbeiter an und pflege eigene Felder. Das Organigramm "
                 "bildet die Vorgesetzten-Struktur ab."),
        TourStep("Abwesenheiten",
                 "Urlaub und Abwesenheiten traegst du zentral ein und behaeltst "
                 "den Ueberblick."),
        TourStep("Fertig",
                 "Die Einfuehrung kannst du jederzeit erneut starten oder "
                 "abschalten."),
    ]
