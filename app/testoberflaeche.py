"""Testoberflaeche (Preview) fuer NMGone – moderne UI, isoliert vom Original.

Ziel: ein professioneller, leicht verstaendlicher Entwurf der Neuen Auswertung,
der die echte Engine (create_vorlage_export) nutzt, aber das Original-GUI NICHT
veraendert. Vier Schwerpunkte:

  1. Sichtbarer Fortschritt – animierter Balken, Stoppuhr, Stufen-Checkliste,
     echte Positionszahl aus der DB (man sieht, dass gearbeitet wird).
  2. Modernes, professionelles Design – Karten, klare Typografie, Sidebar.
  3. Hilfe mit Bildern und Erklaerungen.
  4. .xls-Unterstuetzung – alte Excel-Dateien werden automatisch nach .xlsx
     konvertiert (xlrd -> openpyxl) und dann von der echten Engine verarbeitet.

Start:  python start_testoberflaeche.py
"""
from __future__ import annotations

import threading
import time
import sqlite3
from pathlib import Path

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from .config import DB_PATH
from .testui_assets import ensure_help_images

# ── Palette ──────────────────────────────────────────────────────────────────
PRIMARY = "#0B4A86"
PRIMARY_DARK = "#083A6B"
ACCENT = "#208ACD"
SUCCESS = "#11823B"
WARNING = "#C88200"
DANGER = "#B3261E"
BG = "#F4F6F9"
CARD = "#FFFFFF"
INK = "#1F2933"
MUTED = "#626E7D"
BORDER = "#DCE2E8"
SIDEBAR = "#0B2C4A"
SIDEBAR_ACTIVE = "#13456E"

FONT = "Segoe UI"

SUPPORTED = {".xlsx", ".xlsm", ".xls", ".csv", ".txt"}


# ── .xls-Bruecke ─────────────────────────────────────────────────────────────
def konvertiere_xls_zu_xlsx(xls_path: Path) -> Path:
    """Wandelt eine alte .xls-Datei in eine temporaere .xlsx, damit die
    bestehende Engine (die nur xlsx/xlsm/csv/txt liest) sie verarbeiten kann.
    Liest mit xlrd, schreibt mit openpyxl. Liefert den Pfad der neuen Datei.
    """
    import xlrd
    from openpyxl import Workbook

    import tempfile

    book = xlrd.open_workbook(str(xls_path))
    sheet = book.sheet_by_index(0)
    wb = Workbook()
    ws = wb.active
    ws.title = (sheet.name or "Tabelle1")[:31]
    for r in range(sheet.nrows):
        ws.append([sheet.cell_value(r, c) for c in range(sheet.ncols)])
    out = Path(tempfile.gettempdir()) / f"_xls_konvertiert_{xls_path.stem[:40]}_{int(time.time())}.xlsx"
    wb.save(out)
    return out


def zeilen_schaetzen(path: Path) -> int:
    """Grobe Zeilenzahl der Quelldatei fuer die Fortschrittsschaetzung."""
    suffix = path.suffix.lower()
    try:
        if suffix in {".xlsx", ".xlsm"}:
            from openpyxl import load_workbook
            wb = load_workbook(path, read_only=True)
            return max(0, wb.active.max_row - 1)
        if suffix == ".xls":
            import xlrd
            return max(0, xlrd.open_workbook(str(path)).sheet_by_index(0).nrows - 1)
        with open(path, "r", encoding="utf-8", errors="replace") as fh:
            return max(0, sum(1 for _ in fh) - 1)
    except Exception:
        return 0


# ── Wiederverwendbare UI-Bausteine ───────────────────────────────────────────
class Card(tk.Frame):
    """Weisse Karte mit dezentem Rahmen."""
    def __init__(self, master, padding=20, **kw):
        super().__init__(master, bg=CARD, highlightbackground=BORDER,
                         highlightthickness=1, bd=0, **kw)
        self.inner = tk.Frame(self, bg=CARD)
        self.inner.pack(fill="both", expand=True, padx=padding, pady=padding)


class PillButton(tk.Label):
    """Flacher, moderner Button mit Hover-Effekt."""
    def __init__(self, master, text, command, kind="primary", **kw):
        colors = {
            "primary": (PRIMARY, "#0D5596", "#FFFFFF"),
            "accent": (ACCENT, "#2A97DA", "#FFFFFF"),
            "success": (SUCCESS, "#159247", "#FFFFFF"),
            "ghost": (CARD, "#EEF3F8", PRIMARY),
        }
        self._base, self._hover, fg = colors.get(kind, colors["primary"])
        super().__init__(master, text=text, bg=self._base, fg=fg,
                         font=(FONT, 11, "bold"), padx=22, pady=11, cursor="hand2", **kw)
        self._command = command
        self._enabled = True
        if kind == "ghost":
            self.config(highlightbackground=BORDER, highlightthickness=1)
        self.bind("<Button-1>", self._click)
        self.bind("<Enter>", lambda e: self._enabled and self.config(bg=self._hover))
        self.bind("<Leave>", lambda e: self._enabled and self.config(bg=self._base))

    def _click(self, _e):
        if self._enabled and self._command:
            self._command()

    def set_enabled(self, on: bool):
        self._enabled = on
        self.config(bg=self._base if on else "#C3CCD6",
                    fg="#FFFFFF" if on else "#8A97A5",
                    cursor="hand2" if on else "arrow")


# ── Hauptfenster ─────────────────────────────────────────────────────────────
class TestOberflaeche(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("NMGone – Testoberfläche (Vorschau)")
        self.geometry("1180x760")
        self.minsize(1040, 680)
        self.configure(bg=BG)
        self._busy = False
        self._selected_file: Path | None = None
        self._start_ts = 0.0
        self._erwartete_zeilen = 0

        self._assets = ensure_help_images(Path(__file__).resolve().parent.parent / "testui_assets")
        self._imgcache: dict[str, tk.PhotoImage] = {}

        self._build_sidebar()
        self._build_content()
        self.show_page("auswertung")

    # ----- Layout -----------------------------------------------------------
    def _build_sidebar(self):
        bar = tk.Frame(self, bg=SIDEBAR, width=248)
        bar.pack(side="left", fill="y")
        bar.pack_propagate(False)

        head = tk.Frame(bar, bg=SIDEBAR)
        head.pack(fill="x", pady=(26, 8), padx=22)
        tk.Label(head, text="NMGone", bg=SIDEBAR, fg="#FFFFFF",
                 font=(FONT, 21, "bold")).pack(anchor="w")
        tk.Label(head, text="Testoberfläche · Vorschau", bg=SIDEBAR, fg="#8FB4D6",
                 font=(FONT, 10)).pack(anchor="w")

        tk.Frame(bar, bg="#16395C", height=1).pack(fill="x", padx=18, pady=(14, 10))

        self._nav_buttons: dict[str, tk.Label] = {}
        items = [
            ("auswertung", "📊", "Neue Auswertung"),
            ("hilfe", "❓", "Hilfe & Anleitung"),
            ("info", "ℹ️", "Über diese Vorschau"),
        ]
        for key, icon, label in items:
            b = tk.Label(bar, text=f"   {icon}   {label}", bg=SIDEBAR, fg="#D7E6F4",
                         font=(FONT, 12), anchor="w", padx=14, pady=12, cursor="hand2")
            b.pack(fill="x", padx=12, pady=2)
            b.bind("<Button-1>", lambda e, k=key: self.show_page(k))
            b.bind("<Enter>", lambda e, bb=b, k=key: bb.config(bg=SIDEBAR_ACTIVE) if self._active != k else None)
            b.bind("<Leave>", lambda e, bb=b, k=key: bb.config(bg=SIDEBAR if self._active != k else SIDEBAR_ACTIVE))
            self._nav_buttons[key] = b

        foot = tk.Label(bar, text="Original bleibt unberührt.\nNur eine Vorschau.",
                        bg=SIDEBAR, fg="#6E92B4", font=(FONT, 9), justify="left")
        foot.pack(side="bottom", anchor="w", padx=24, pady=18)
        self._active = "auswertung"

    def _build_content(self):
        self.content = tk.Frame(self, bg=BG)
        self.content.pack(side="left", fill="both", expand=True)
        self._pages: dict[str, tk.Frame] = {}

    def show_page(self, key: str):
        if self._busy and key != "auswertung":
            # Waehrend einer laufenden Auswertung Navigation erlauben, aber Seite neu bauen ist ok.
            pass
        self._active = key
        for k, b in self._nav_buttons.items():
            b.config(bg=SIDEBAR_ACTIVE if k == key else SIDEBAR,
                     fg="#FFFFFF" if k == key else "#D7E6F4")
        for child in self.content.winfo_children():
            child.destroy()
        builder = {"auswertung": self._page_auswertung,
                   "hilfe": self._page_hilfe,
                   "info": self._page_info}.get(key, self._page_auswertung)
        builder()

    def _page_header(self, parent, title, subtitle):
        head = tk.Frame(parent, bg=BG)
        head.pack(fill="x", padx=36, pady=(30, 10))
        tk.Label(head, text=title, bg=BG, fg=INK, font=(FONT, 26, "bold")).pack(anchor="w")
        tk.Label(head, text=subtitle, bg=BG, fg=MUTED, font=(FONT, 13)).pack(anchor="w", pady=(2, 0))

    def _img(self, name: str, max_w: int | None = None) -> tk.PhotoImage | None:
        p = self._assets.get(name)
        if not p or not Path(p).exists():
            return None
        if name in self._imgcache:
            return self._imgcache[name]
        try:
            img = tk.PhotoImage(file=str(p))
            if max_w and img.width() > max_w:
                factor = max(1, round(img.width() / max_w))
                img = img.subsample(factor, factor)
            self._imgcache[name] = img
            return img
        except Exception:
            return None

    # ----- Seite: Neue Auswertung ------------------------------------------
    def _page_auswertung(self):
        page = tk.Frame(self.content, bg=BG)
        page.pack(fill="both", expand=True)
        self._page_header(page, "Neue Auswertung",
                          "Datei wählen, starten – der Fortschritt ist jederzeit sichtbar.")

        body = tk.Frame(page, bg=BG)
        body.pack(fill="both", expand=True, padx=36, pady=(6, 28))

        # Schritt 1: Datei
        c1 = Card(body)
        c1.pack(fill="x", pady=(0, 16))
        tk.Label(c1.inner, text="1 · Datei auswählen", bg=CARD, fg=PRIMARY,
                 font=(FONT, 14, "bold")).pack(anchor="w")
        row = tk.Frame(c1.inner, bg=CARD)
        row.pack(fill="x", pady=(12, 0))
        self._file_var = tk.StringVar(value="Noch keine Datei gewählt")
        drop = tk.Frame(row, bg="#F0F4F8", highlightbackground=BORDER, highlightthickness=1)
        drop.pack(side="left", fill="x", expand=True, ipady=14)
        tk.Label(drop, textvariable=self._file_var, bg="#F0F4F8", fg=INK,
                 font=(FONT, 11)).pack(side="left", padx=16)
        self._file_meta = tk.Label(drop, text="", bg="#F0F4F8", fg=MUTED, font=(FONT, 10))
        self._file_meta.pack(side="right", padx=16)
        PillButton(row, "Durchsuchen …", self._choose_file, kind="ghost").pack(side="left", padx=(12, 0))
        tk.Label(c1.inner, text="Unterstützt: .xlsx · .xls · .xlsm · .csv · .txt   (alte .xls werden automatisch konvertiert)",
                 bg=CARD, fg=MUTED, font=(FONT, 10)).pack(anchor="w", pady=(10, 0))

        # Schritt 2: Name + Start
        c2 = Card(body)
        c2.pack(fill="x", pady=(0, 16))
        tk.Label(c2.inner, text="2 · Bezeichnung & Start", bg=CARD, fg=PRIMARY,
                 font=(FONT, 14, "bold")).pack(anchor="w")
        row2 = tk.Frame(c2.inner, bg=CARD)
        row2.pack(fill="x", pady=(12, 0))
        tk.Label(row2, text="Name der Auswertung", bg=CARD, fg=MUTED, font=(FONT, 10)).pack(anchor="w")
        self._name_var = tk.StringVar(value="Testauswertung")
        ent = tk.Entry(row2, textvariable=self._name_var, font=(FONT, 12),
                       relief="flat", highlightbackground=BORDER, highlightthickness=1)
        ent.pack(side="left", fill="x", expand=True, ipady=7, pady=(4, 0))
        self._start_btn = PillButton(row2, "▶  Auswertung starten", self._start_auswertung, kind="success")
        self._start_btn.pack(side="left", padx=(12, 0), pady=(4, 0))

        # Schritt 3: Fortschritt (anfangs versteckt)
        self._progress_card = Card(body)
        self._build_progress_widgets(self._progress_card.inner)

    def _build_progress_widgets(self, parent):
        self._prog_title = tk.Label(parent, text="Auswertung läuft …", bg=CARD, fg=INK,
                                    font=(FONT, 15, "bold"))
        self._prog_title.pack(anchor="w")
        toprow = tk.Frame(parent, bg=CARD)
        toprow.pack(fill="x", pady=(2, 12))
        self._prog_sub = tk.Label(toprow, text="", bg=CARD, fg=MUTED, font=(FONT, 11))
        self._prog_sub.pack(side="left")
        self._prog_timer = tk.Label(toprow, text="00:00", bg=CARD, fg=ACCENT, font=(FONT, 13, "bold"))
        self._prog_timer.pack(side="right")

        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except Exception:
            pass
        style.configure("NMG.Horizontal.TProgressbar", troughcolor="#E8EDF2",
                        background=ACCENT, thickness=20, borderwidth=0)
        self._pbar = ttk.Progressbar(parent, style="NMG.Horizontal.TProgressbar",
                                     mode="indeterminate", length=100)
        self._pbar.pack(fill="x")

        self._stage_frame = tk.Frame(parent, bg=CARD)
        self._stage_frame.pack(fill="x", pady=(14, 0))
        self._stage_labels: list[tk.Label] = []
        stages = ["Datei eingelesen", "Stammdaten geladen", "Positionen abgleichen", "Excel schreiben", "Fertig"]
        for i, s in enumerate(stages):
            lab = tk.Label(self._stage_frame, text=f"○  {s}", bg=CARD, fg=MUTED, font=(FONT, 11))
            lab.grid(row=0, column=i, padx=(0, 26), sticky="w")
            self._stage_labels.append(lab)

        self._result_frame = tk.Frame(parent, bg=CARD)
        self._result_frame.pack(fill="x", pady=(16, 0))

    def _set_stage(self, idx: int, state: str):
        marks = {"done": ("✓", SUCCESS), "active": ("●", ACCENT), "todo": ("○", MUTED)}
        labels = ["Datei eingelesen", "Stammdaten geladen", "Positionen abgleichen", "Excel schreiben", "Fertig"]
        for i, lab in enumerate(self._stage_labels):
            if i < idx:
                m, c = marks["done"]
            elif i == idx:
                m, c = marks[state] if state in marks else marks["active"]
            else:
                m, c = marks["todo"]
            lab.config(text=f"{m}  {labels[i]}", fg=c,
                       font=(FONT, 11, "bold" if i == idx else "normal"))

    # ----- Datei-Auswahl ----------------------------------------------------
    def _choose_file(self):
        if self._busy:
            return
        path = filedialog.askopenfilename(
            title="Rohdatei wählen",
            filetypes=[("Alle unterstützten", "*.xlsx *.xlsm *.xls *.csv *.txt"),
                       ("Excel neu", "*.xlsx *.xlsm"), ("Excel alt", "*.xls"),
                       ("CSV/Text", "*.csv *.txt"), ("Alle Dateien", "*.*")])
        if not path:
            return
        p = Path(path)
        if p.suffix.lower() not in SUPPORTED:
            messagebox.showwarning("Format", f"'{p.suffix}' wird nicht unterstützt.\nErlaubt: xlsx, xls, xlsm, csv, txt.")
            return
        self._selected_file = p
        self._file_var.set(p.name)
        n = zeilen_schaetzen(p)
        self._erwartete_zeilen = n
        tag = "alte Excel – wird konvertiert" if p.suffix.lower() == ".xls" else p.suffix.lower().lstrip(".").upper()
        self._file_meta.config(text=f"{n} Zeilen · {tag}")
        if not self._name_var.get().strip() or self._name_var.get() == "Testauswertung":
            self._name_var.set(p.stem[:40])

    # ----- Start / Hintergrund ---------------------------------------------
    def _start_auswertung(self):
        if self._busy:
            return
        if not self._selected_file:
            messagebox.showinfo("Datei fehlt", "Bitte zuerst eine Datei auswählen.")
            return
        name = self._name_var.get().strip() or "Testauswertung"

        self._busy = True
        self._start_btn.set_enabled(False)
        self._progress_card.pack(fill="x", pady=(0, 8))
        for f in self._result_frame.winfo_children():
            f.destroy()
        self._prog_title.config(text="Auswertung läuft …", fg=INK)
        self._prog_sub.config(text=f"Verarbeite {self._selected_file.name}")
        self._pbar.config(mode="indeterminate")
        self._pbar.start(12)
        self._start_ts = time.time()
        self._set_stage(0, "active")
        self._tick_timer()
        self._poll_positions()

        threading.Thread(target=self._worker, args=(self._selected_file, name), daemon=True).start()

    def _worker(self, file: Path, name: str):
        """Laeuft im Hintergrund: ggf. .xls konvertieren, dann echte Engine."""
        from .exporter import create_vorlage_export
        try:
            src = file
            if file.suffix.lower() == ".xls":
                self.after(0, lambda: self._prog_sub.config(text="Alte .xls-Datei wird konvertiert …"))
                src = konvertiere_xls_zu_xlsx(file)
            self.after(0, lambda: self._set_stage(1, "active"))
            out = create_vorlage_export(str(src), name)
            self.after(0, lambda: self._finish_ok(Path(out)))
        except Exception as exc:
            self.after(0, lambda e=exc: self._finish_err(e))

    def _tick_timer(self):
        if not self._busy:
            return
        el = int(time.time() - self._start_ts)
        self._prog_timer.config(text=f"{el // 60:02d}:{el % 60:02d}")
        self.after(500, self._tick_timer)

    def _latest_position_count(self) -> int:
        try:
            with sqlite3.connect(DB_PATH) as con:
                row = con.execute(
                    "SELECT COUNT(*) FROM tbl_auswertungspositionen "
                    "WHERE auswertung_id = (SELECT MAX(id) FROM tbl_auswertungen)"
                ).fetchone()
                return int(row[0]) if row else 0
        except Exception:
            return 0

    def _poll_positions(self):
        """Echte Rueckmeldung: zeigt, wie viele Positionen schon in der DB sind."""
        if not self._busy:
            return
        cnt = self._latest_position_count()
        if cnt > 0:
            self._set_stage(2, "active")
            if self._erwartete_zeilen:
                pct = min(99, int(cnt / max(1, self._erwartete_zeilen) * 100))
                self._prog_sub.config(text=f"{cnt} von ~{self._erwartete_zeilen} Positionen abgeglichen ({pct}%)")
            else:
                self._prog_sub.config(text=f"{cnt} Positionen abgeglichen …")
        self.after(400, self._poll_positions)

    def _finish_ok(self, out: Path):
        self._busy = False
        self._pbar.stop()
        self._pbar.config(mode="determinate", value=100)
        self._set_stage(4, "done")
        self._prog_title.config(text="✓  Auswertung fertig", fg=SUCCESS)
        self._prog_sub.config(text="Die Excel-Datei wurde erstellt.")
        self._start_btn.set_enabled(True)
        for f in self._result_frame.winfo_children():
            f.destroy()
        box = tk.Frame(self._result_frame, bg="#EAF6EF", highlightbackground="#BfE3CD", highlightthickness=1)
        box.pack(fill="x")
        tk.Label(box, text=f"📄  {out.name}", bg="#EAF6EF", fg=INK, font=(FONT, 11, "bold")).pack(side="left", padx=16, pady=12)
        PillButton(box, "Ordner öffnen", lambda: self._open_folder(out.parent), kind="success").pack(side="right", padx=12, pady=8)
        PillButton(box, "Excel öffnen", lambda: self._open_file(out), kind="ghost").pack(side="right", pady=8)

    def _finish_err(self, exc: Exception):
        self._busy = False
        self._pbar.stop()
        self._set_stage(2, "todo")
        self._prog_title.config(text="✗  Auswertung fehlgeschlagen", fg=DANGER)
        self._prog_sub.config(text=str(exc)[:120])
        self._start_btn.set_enabled(True)
        for f in self._result_frame.winfo_children():
            f.destroy()
        msg = str(exc)
        hint = ""
        if "nicht erkannt" in msg.lower() or "format" in msg.lower():
            hint = "\n\nTipp: Im Original gibt es den Rohdaten-Formatassistenten für unbekannte Spalten."
        tk.Label(self._result_frame, text=msg[:300] + hint, bg=CARD, fg=DANGER,
                 font=(FONT, 10), justify="left", wraplength=820).pack(anchor="w")

    # ----- Seite: Hilfe -----------------------------------------------------
    def _page_hilfe(self):
        page = tk.Frame(self.content, bg=BG)
        page.pack(fill="both", expand=True)
        self._page_header(page, "Hilfe & Anleitung",
                          "In drei Schritten zur fertigen Auswertung – mit Bildern erklärt.")

        canvas = tk.Canvas(page, bg=BG, highlightthickness=0)
        scroll = ttk.Scrollbar(page, orient="vertical", command=canvas.yview)
        inner = tk.Frame(canvas, bg=BG)
        inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=inner, anchor="nw")
        canvas.configure(yscrollcommand=scroll.set)
        canvas.pack(side="left", fill="both", expand=True, padx=(36, 0), pady=(0, 20))
        scroll.pack(side="right", fill="y")
        canvas.bind_all("<MouseWheel>", lambda e: canvas.yview_scroll(int(-e.delta / 120), "units"))

        sections = [
            ("prozess_flow.png",
             "Wie eine Auswertung entsteht",
             "Du wählst eine Datei und startest. Den Rest macht das Programm automatisch: "
             "Es erkennt die Spalten, gleicht jede PZN gegen Stammdaten und Biosimilars ab, "
             "berechnet Absatz und Umsatz und schreibt am Ende eine fertige Excel-Datei."),
            ("formate.png",
             "Welche Dateien funktionieren",
             "Moderne Excel-Dateien (.xlsx), alte Excel-Dateien (.xls), sowie CSV- und Text-"
             "tabellen. Alte .xls-Dateien werden automatisch in das neue Format umgewandelt – "
             "du musst nichts vorbereiten."),
            ("fortschritt.png",
             "Du siehst, dass gearbeitet wird",
             "Während der Auswertung laufen ein Fortschrittsbalken und eine Stoppuhr. Die "
             "Stufen-Checkliste zeigt, wo das Programm gerade steht, und die Zahl der bereits "
             "abgeglichenen Positionen steigt sichtbar an."),
        ]
        for img_name, title, text in sections:
            card = Card(inner, padding=22)
            card.pack(fill="x", pady=(0, 18), padx=(0, 8))
            tk.Label(card.inner, text=title, bg=CARD, fg=PRIMARY, font=(FONT, 17, "bold")).pack(anchor="w")
            tk.Label(card.inner, text=text, bg=CARD, fg=INK, font=(FONT, 11),
                     justify="left", wraplength=820).pack(anchor="w", pady=(6, 14))
            img = self._img(img_name, max_w=860)
            if img:
                tk.Label(card.inner, image=img, bg=CARD).pack(anchor="w")
            else:
                tk.Label(card.inner, text="(Bild konnte nicht geladen werden)", bg=CARD, fg=MUTED).pack(anchor="w")

    # ----- Seite: Info ------------------------------------------------------
    def _page_info(self):
        page = tk.Frame(self.content, bg=BG)
        page.pack(fill="both", expand=True)
        self._page_header(page, "Über diese Vorschau",
                          "Was diese Testoberfläche ist – und was nicht.")
        card = Card(page, padding=24)
        card.pack(fill="x", padx=36, pady=(6, 0))
        punkte = [
            ("Isoliert vom Original", "Diese Oberfläche ist ein eigenes Fenster und verändert das bestehende NMGone-Programm nicht."),
            ("Echte Engine", "Die Auswertung nutzt dieselbe Berechnung wie das Original (create_vorlage_export) – die Ergebnisse sind echt."),
            ("Sichtbarer Fortschritt", "Balken, Stoppuhr, Stufen und die live steigende Positionszahl zeigen, dass gearbeitet wird."),
            (".xls wird gelesen", "Alte Excel-Dateien werden automatisch konvertiert und ausgewertet."),
            ("Zweck", "Entwurf für ein moderneres, professionelleres und leichter verständliches Design."),
        ]
        for t, b in punkte:
            r = tk.Frame(card.inner, bg=CARD)
            r.pack(fill="x", pady=7)
            tk.Label(r, text="●", bg=CARD, fg=ACCENT, font=(FONT, 12)).pack(side="left", anchor="n", padx=(0, 10))
            col = tk.Frame(r, bg=CARD)
            col.pack(side="left", fill="x", expand=True)
            tk.Label(col, text=t, bg=CARD, fg=INK, font=(FONT, 12, "bold")).pack(anchor="w")
            tk.Label(col, text=b, bg=CARD, fg=MUTED, font=(FONT, 11), justify="left", wraplength=820).pack(anchor="w")


def main():
    app = TestOberflaeche()
    app.mainloop()


if __name__ == "__main__":
    main()
