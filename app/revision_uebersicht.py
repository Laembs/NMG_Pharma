# -*- coding: utf-8 -*-
"""Revisions-Uebersicht – die Testoberflaeche "Was wurde veraendert".

Ein eigenes, ISOLIERTES Fenster (das Hauptprogramm bleibt unberuehrt), das auf
einen Blick zeigt, was in dieser Revision dazugekommen ist: die neuen Apps, die
Aenderungen, der Aufraeum-Stand und die Roadmap fuer die naechste „Revolution“.

Inhalt kommt aus app/revision_data.py (gleiche Quelle wie das PDF-Handout).
Optik aus app/theme.py (gemeinsamer Look der ganzen Programm-Familie).

Start:  python start_revision.py   (oder „Revisions-Uebersicht starten.bat“)
"""
from __future__ import annotations

import os
import sys
import subprocess
from pathlib import Path

import tkinter as tk
from tkinter import ttk, messagebox

from . import theme
from . import revision_data as RD

ROOT = Path(__file__).resolve().parent.parent
PDF_PATH = ROOT / "docs" / "Handout_Revision_Uebersicht.pdf"

STATUS_BADGE = {
    "neu":       ("NEU",       theme.SUCCESS),
    "erweitert": ("ERWEITERT", theme.ACCENT),
    "stabil":    ("STABIL",    theme.MUTED),
}
AUFWAND_BADGE = {
    "klein":  ("kleiner Aufwand",  theme.SUCCESS),
    "mittel": ("mittlerer Aufwand", theme.WARNING),
    "gross":  ("grosser Aufwand",  theme.DANGER),
}


# ── Scrollbarer Seiteninhalt ─────────────────────────────────────────────────
class ScrollPage(tk.Frame):
    """Vertikal scrollbarer Inhaltsbereich. Inhalte in .body packen."""

    def __init__(self, master):
        super().__init__(master, bg=theme.BG)
        self._canvas = tk.Canvas(self, bg=theme.BG, highlightthickness=0, bd=0)
        vsb = ttk.Scrollbar(self, orient="vertical", command=self._canvas.yview)
        self._canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        self._canvas.pack(side="left", fill="both", expand=True)

        self.body = tk.Frame(self._canvas, bg=theme.BG)
        self._win = self._canvas.create_window((0, 0), window=self.body, anchor="nw")
        self.body.bind("<Configure>",
                       lambda e: self._canvas.configure(scrollregion=self._canvas.bbox("all")))
        self._canvas.bind("<Configure>",
                          lambda e: self._canvas.itemconfigure(self._win, width=e.width))
        # Mausrad (Windows/Mac liefern <MouseWheel>)
        self._canvas.bind_all("<MouseWheel>", self._on_wheel)

    def _on_wheel(self, event):
        try:
            self._canvas.yview_scroll(int(-event.delta / 120), "units")
        except Exception:
            pass

    def destroy(self):
        try:
            self._canvas.unbind_all("<MouseWheel>")
        except Exception:
            pass
        super().destroy()


def _chip(parent, text, color, bg=theme.CARD):
    f = tk.Frame(parent, bg=color)
    tk.Label(f, text=f" {text} ", bg=color, fg="#FFFFFF",
             font=(theme.FONT, 9, "bold")).pack(padx=1, pady=1)
    return f


def _bullet(parent, title, detail, color=theme.PRIMARY):
    row = tk.Frame(parent, bg=theme.CARD)
    row.pack(fill="x", pady=(8, 0))
    tk.Label(row, text="●", bg=theme.CARD, fg=color, font=(theme.FONT, 11)).pack(
        side="left", anchor="n", padx=(0, 8))
    txt = tk.Frame(row, bg=theme.CARD)
    txt.pack(side="left", fill="x", expand=True)
    tk.Label(txt, text=title, bg=theme.CARD, fg=theme.INK,
             font=(theme.FONT, 11, "bold"), anchor="w", justify="left").pack(anchor="w")
    if detail:
        tk.Label(txt, text=detail, bg=theme.CARD, fg=theme.MUTED, font=(theme.FONT, 10),
                 anchor="w", justify="left", wraplength=720).pack(anchor="w")


# ── Hauptfenster ─────────────────────────────────────────────────────────────
class RevisionUebersicht(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("NMGone – Revisions-Übersicht (Was wurde verändert)")
        self.geometry("1200x780")
        self.minsize(1020, 660)
        self.configure(bg=theme.BG)
        theme.apply_theme(self)
        theme.apply_widget_defaults(self)
        try:
            self.iconbitmap(str(ROOT / "assets" / "GDP.ico"))
        except Exception:
            pass

        self.sidebar = theme.Sidebar(self, title="NMGone",
                                     subtitle="Revisions-Übersicht")
        self.sidebar.pack(side="left", fill="y")
        self.sidebar.add_section("Revision " + RD.REVISION["version"])
        nav = [
            ("ueberblick", "📋", "Überblick", self._page_ueberblick),
            ("familie",    "🧩", "Programm-Familie", self._page_familie),
            ("aenderungen","✨", "Was wurde verändert", self._page_aenderungen),
            ("aufraeumen", "🧹", "Aufgeräumt", self._page_aufraeumen),
            ("roadmap",    "🚀", "Nächste Revolution", self._page_roadmap),
        ]
        self._pages = {k: b for k, _, _, b in [(k, i, l, fn) for k, i, l, fn in nav]}
        for key, icon, label, _ in nav:
            self.sidebar.add_item(key, icon, label,
                                  lambda k=key: self.show(k))
        self.sidebar.add_footer_note("Nur eine Übersicht.\nHauptprogramm bleibt unberührt.")

        # Fuss-Aktion: PDF oeffnen
        foot = tk.Frame(self.sidebar._foot, bg=theme.SIDEBAR)
        foot.pack(fill="x", padx=16, pady=(0, 6))
        theme.PillButton(foot, "📄  PDF-Handout öffnen", self._open_pdf,
                         kind="accent", font_size=10).pack(fill="x")

        self.content = tk.Frame(self, bg=theme.BG)
        self.content.pack(side="left", fill="both", expand=True)
        self._current = None
        self.show("ueberblick")

    # ----- Navigation -------------------------------------------------------
    def show(self, key):
        if self._current is not None:
            self._current.destroy()
        self.sidebar.set_active(key)
        page = ScrollPage(self.content)
        page.pack(fill="both", expand=True)
        self._current = page
        self._pages[key](page.body)

    def _header(self, parent, title, subtitle):
        h = theme.page_header(parent, title, subtitle)
        h.pack(fill="x", padx=34, pady=(28, 6))

    # ----- Seite: Überblick -------------------------------------------------
    def _page_ueberblick(self, body):
        self._header(body, RD.REVISION["titel"], RD.REVISION["untertitel"])

        # Claim-Karte
        c = theme.Card(body)
        c.pack(fill="x", padx=34, pady=(8, 16))
        tk.Label(c, text="Worum es geht", bg=theme.CARD, fg=theme.PRIMARY,
                 font=(theme.FONT, 13, "bold")).pack(anchor="w")
        tk.Label(c, text=RD.REVISION["claim"], bg=theme.CARD, fg=theme.INK,
                 font=(theme.FONT, 12), justify="left", wraplength=820).pack(
            anchor="w", pady=(6, 2))
        tk.Label(c, text=f"Stand {RD.REVISION['datum']} · Version {RD.REVISION['version']}",
                 bg=theme.CARD, fg=theme.MUTED, font=(theme.FONT, 10)).pack(anchor="w")

        # Kennzahlen-Reihe
        tk.Label(body, text="Auf einen Blick", bg=theme.BG, fg=theme.INK,
                 font=(theme.FONT, 14, "bold")).pack(anchor="w", padx=34, pady=(6, 6))
        grid = tk.Frame(body, bg=theme.BG)
        grid.pack(fill="x", padx=28)
        for i, (zahl, titel, sub) in enumerate(RD.REVISION["kennzahlen"]):
            k = tk.Frame(grid, bg=theme.CARD, highlightbackground=theme.BORDER,
                         highlightthickness=1)
            k.grid(row=0, column=i, sticky="nsew", padx=6, pady=4, ipadx=8, ipady=10)
            grid.columnconfigure(i, weight=1)
            tk.Label(k, text=zahl, bg=theme.CARD, fg=theme.PRIMARY,
                     font=(theme.FONT, 30, "bold")).pack(anchor="w", padx=14, pady=(6, 0))
            tk.Label(k, text=titel, bg=theme.CARD, fg=theme.INK,
                     font=(theme.FONT, 11, "bold")).pack(anchor="w", padx=14)
            tk.Label(k, text=sub, bg=theme.CARD, fg=theme.MUTED, font=(theme.FONT, 9),
                     justify="left", wraplength=220).pack(anchor="w", padx=14, pady=(2, 8))

        # Hinweiskarte
        info = theme.Card(body)
        info.pack(fill="x", padx=34, pady=(18, 28))
        tk.Label(info, text="💡  So nutzt du diese Übersicht", bg=theme.CARD,
                 fg=theme.SUCCESS, font=(theme.FONT, 12, "bold")).pack(anchor="w")
        for t in ("Links durch die Punkte klicken – jede Seite erklärt einen Bereich.",
                  "„Programm-Familie“ zeigt alle Apps und was neu daran ist.",
                  "„Nächste Revolution“ ist der Fahrplan für deine Rückkehr.",
                  "Unten links das PDF-Handout öffnen, um alles in Ruhe nachzulesen."):
            tk.Label(info, text="•  " + t, bg=theme.CARD, fg=theme.INK,
                     font=(theme.FONT, 11), justify="left", wraplength=820).pack(
                anchor="w", pady=(4, 0))

    # ----- Seite: Programm-Familie -----------------------------------------
    def _page_familie(self, body):
        self._header(body, "Programm-Familie",
                     "Eigenständige Apps rund um eine gemeinsame Datenbank.")
        wrap = tk.Frame(body, bg=theme.BG)
        wrap.pack(fill="both", expand=True, padx=28, pady=(6, 28))
        cols = 2
        for i, app in enumerate(RD.APPS):
            card = tk.Frame(wrap, bg=theme.CARD, highlightbackground=theme.BORDER,
                            highlightthickness=1)
            card.grid(row=i // cols, column=i % cols, sticky="nsew", padx=6, pady=6)
            wrap.columnconfigure(i % cols, weight=1)

            top = tk.Frame(card, bg=theme.CARD)
            top.pack(fill="x", padx=16, pady=(14, 2))
            tk.Label(top, text=app["icon"], bg=theme.CARD, fg=app["farbe"],
                     font=(theme.FONT, 22)).pack(side="left")
            tk.Label(top, text=app["name"], bg=theme.CARD, fg=theme.INK,
                     font=(theme.FONT, 13, "bold")).pack(side="left", padx=(8, 0))
            label, color = STATUS_BADGE.get(app["status"], STATUS_BADGE["stabil"])
            _chip(top, label, color).pack(side="right")

            tk.Label(card, text=app["zweck"], bg=theme.CARD, fg=theme.MUTED,
                     font=(theme.FONT, 10), justify="left", wraplength=360).pack(
                anchor="w", padx=16, pady=(4, 0))
            box = tk.Frame(card, bg=theme.CARD_ALT)
            box.pack(fill="x", padx=16, pady=(8, 6))
            tk.Label(box, text="Neu/Geändert: " + app["neu"], bg=theme.CARD_ALT,
                     fg=theme.INK, font=(theme.FONT, 9), justify="left",
                     wraplength=350).pack(anchor="w", padx=8, pady=6)
            tk.Label(card, text="▸ " + app["start"], bg=theme.CARD, fg=theme.FAINT,
                     font=(theme.MONO, 9)).pack(anchor="w", padx=16, pady=(0, 12))

    # ----- Seite: Was wurde verändert --------------------------------------
    def _page_aenderungen(self, body):
        self._header(body, "Was wurde verändert",
                     "Alle Änderungen dieser Revision – nach Bereich gruppiert.")
        for kategorie, eintraege in RD.AENDERUNGEN:
            c = theme.Card(body)
            c.pack(fill="x", padx=34, pady=(8, 4))
            tk.Label(c, text=kategorie, bg=theme.CARD, fg=theme.PRIMARY,
                     font=(theme.FONT, 13, "bold")).pack(anchor="w")
            for titel, detail in eintraege:
                _bullet(c, titel, detail)
        tk.Frame(body, bg=theme.BG, height=20).pack()

    # ----- Seite: Aufgeräumt ------------------------------------------------
    def _page_aufraeumen(self, body):
        self._header(body, "Aufgeräumt & geprüft",
                     "Was bereits erledigt ist – und was du noch freigeben kannst.")
        done = theme.Card(body)
        done.pack(fill="x", padx=34, pady=(8, 10))
        tk.Label(done, text="✓  Bereits erledigt", bg=theme.CARD, fg=theme.SUCCESS,
                 font=(theme.FONT, 13, "bold")).pack(anchor="w")
        for t in RD.AUFGERAEUMT:
            _bullet(done, t, "", color=theme.SUCCESS)

        rec = theme.Card(body)
        rec.pack(fill="x", padx=34, pady=(6, 24))
        tk.Label(rec, text="⚠  Empfohlen – aber NICHT automatisch gelöscht",
                 bg=theme.CARD, fg=theme.WARNING, font=(theme.FONT, 13, "bold")).pack(anchor="w")
        tk.Label(rec, text="Diese Dinge liegen nur lokal herum (alle in .gitignore). "
                           "Ich habe sie bewusst stehen lassen – du entscheidest, wenn du zurück bist.",
                 bg=theme.CARD, fg=theme.MUTED, font=(theme.FONT, 10),
                 justify="left", wraplength=820).pack(anchor="w", pady=(2, 6))
        tbl = tk.Frame(rec, bg=theme.CARD)
        tbl.pack(fill="x")
        heads = ("Pfad", "Größe", "Empfehlung")
        widths = (26, 12, 64)
        hr = tk.Frame(tbl, bg=theme.CARD)
        hr.pack(fill="x")
        for h, w in zip(heads, widths):
            tk.Label(hr, text=h, bg=theme.CARD, fg=theme.FAINT, width=w, anchor="w",
                     font=(theme.FONT, 9, "bold")).pack(side="left")
        for pfad, groesse, empf in RD.AUFRAEUM_EMPFEHLUNG:
            r = tk.Frame(tbl, bg=theme.CARD)
            r.pack(fill="x", pady=1)
            tk.Label(r, text=pfad, bg=theme.CARD, fg=theme.INK, width=widths[0],
                     anchor="w", font=(theme.MONO, 9)).pack(side="left")
            tk.Label(r, text=groesse, bg=theme.CARD, fg=theme.PRIMARY, width=widths[1],
                     anchor="w", font=(theme.FONT, 9, "bold")).pack(side="left")
            tk.Label(r, text=empf, bg=theme.CARD, fg=theme.MUTED, width=widths[2],
                     anchor="w", font=(theme.FONT, 9), justify="left").pack(side="left")

    # ----- Seite: Roadmap ---------------------------------------------------
    def _page_roadmap(self, body):
        self._header(body, "Nächste Revolution",
                     "Der Fahrplan – sortiert, damit ihr direkt loslegen könnt.")
        for i, (titel, beschr, aufwand) in enumerate(RD.ROADMAP, start=1):
            c = theme.Card(body)
            c.pack(fill="x", padx=34, pady=(8, 4))
            head = tk.Frame(c, bg=theme.CARD)
            head.pack(fill="x")
            tk.Label(head, text=f"{i}.", bg=theme.CARD, fg=theme.PRIMARY,
                     font=(theme.FONT, 14, "bold")).pack(side="left", padx=(0, 8))
            tk.Label(head, text=titel, bg=theme.CARD, fg=theme.INK,
                     font=(theme.FONT, 13, "bold")).pack(side="left")
            label, color = AUFWAND_BADGE.get(aufwand, AUFWAND_BADGE["mittel"])
            _chip(head, label, color).pack(side="right")
            tk.Label(c, text=beschr, bg=theme.CARD, fg=theme.MUTED,
                     font=(theme.FONT, 11), justify="left", wraplength=820).pack(
                anchor="w", pady=(4, 0))
        tk.Frame(body, bg=theme.BG, height=20).pack()

    # ----- PDF öffnen -------------------------------------------------------
    def _open_pdf(self):
        if not PDF_PATH.exists():
            messagebox.showinfo(
                "PDF-Handout",
                "Das PDF wurde noch nicht erzeugt.\n\n"
                "Bitte einmal ausführen:\n    python scripts/build_handout_revision.py\n\n"
                f"Es entsteht unter:\n{PDF_PATH}", parent=self)
            return
        try:
            if os.name == "nt":
                os.startfile(str(PDF_PATH))  # type: ignore[attr-defined]
            elif sys.platform == "darwin":
                subprocess.Popen(["open", str(PDF_PATH)])
            else:
                subprocess.Popen(["xdg-open", str(PDF_PATH)])
        except Exception as exc:
            messagebox.showerror("PDF-Handout", f"Konnte das PDF nicht öffnen:\n{exc}",
                                 parent=self)


def run_standalone():
    if os.name == "nt":
        try:
            import ctypes
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("NMG.Revision")
        except Exception:
            pass
    RevisionUebersicht().mainloop()


def main():
    run_standalone()


if __name__ == "__main__":
    main()
