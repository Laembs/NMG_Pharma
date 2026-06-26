"""Zentrales Design-System fuer NMGone (gui.py, kasse_app.py, report_app.py).

Eine einzige Quelle fuer Palette, Schrift, ttk-Styles und wiederverwendbare
Widgets, damit alle Fenster denselben modernen, professionellen Look haben.
Optik abgeleitet aus der Testoberflaeche (app/testoberflaeche.py), die der
User als Vorbild gewaehlt hat.

Verwendung:
    from . import theme
    theme.apply_theme(root)          # ttk-Styles global setzen
    bar = theme.Sidebar(root)        # dunkle Navigation
    card = theme.Card(parent)        # weisse Karte
    theme.PillButton(parent, "OK", cmd, kind="success")
    theme.page_header(parent, "Titel", "Untertitel")
"""
from __future__ import annotations

import tkinter as tk
from tkinter import ttk
from tkinter import font as tkfont

try:                                   # Pillow ist bereits Projekt-Abhaengigkeit
    from PIL import Image, ImageTk, ImageDraw
    _PIL_OK = True
except Exception:                      # ohne PIL: eckiger Fallback (siehe PillButton)
    _PIL_OK = False


# ── Palette ──────────────────────────────────────────────────────────────────
PRIMARY = "#0B4A86"
PRIMARY_DARK = "#083A6B"
ACCENT = "#208ACD"
ACCENT_DARK = "#1A6FA6"
SUCCESS = "#11823B"
SUCCESS_DARK = "#0D6630"
WARNING = "#C88200"
DANGER = "#B3261E"
DANGER_DARK = "#8E1E18"
PURPLE = "#6B4FB3"

BG = "#F4F6F9"           # App-Hintergrund / Arbeitsflaeche
CARD = "#FFFFFF"         # Karten / Formulare
CARD_ALT = "#F7FAFD"     # leicht abgesetzte Flaeche
INK = "#1F2933"          # Haupttext
MUTED = "#626E7D"        # Sekundaertext
FAINT = "#9AA5B1"        # Hinweise
BORDER = "#DCE2E8"       # dezente Rahmen
DIVIDER = "#EBEFF3"

SIDEBAR = "#0B2C4A"      # dunkle Navigation
SIDEBAR_DARK = "#082238"
SIDEBAR_ACTIVE = "#13456E"
SIDEBAR_TEXT = "#D7E6F4"
SIDEBAR_MUTED = "#6E92B4"
SELECT_BG = "#E8F1FB"    # Auswahl in Listen/Tabellen

# Legacy-Aliase: erleichtern den Sweep alter inline-Hexes im Bestand.
LEGACY = {
    "#f5f7fb": BG, "#f8fbff": CARD_ALT, "#ffffff": CARD, "#d8e2ee": BORDER,
    "#0b4a86": PRIMARY, "#11823b": SUCCESS, "#0a7d2c": SUCCESS, "#9b1c1c": DANGER,
    "#8b5a00": WARNING, "#6b4fb3": PURPLE, "#e8f1fb": SELECT_BG, "#666": MUTED,
    "#333": INK, "#f0f4fa": CARD_ALT,
}


# ── Schrift ──────────────────────────────────────────────────────────────────
FONT = "Segoe UI"
MONO = "Consolas"


def font(size: int, bold: bool = False, italic: bool = False):
    style = " ".join(s for s in (("bold" if bold else ""), ("italic" if italic else "")) if s)
    return (FONT, size, style) if style else (FONT, size)


H1 = (FONT, 18, "bold")
H2 = (FONT, 15, "bold")
SECTION = (FONT, 13, "bold")
BODY = (FONT, 11)
BODY_BOLD = (FONT, 11, "bold")
SMALL = (FONT, 10)
TINY = (FONT, 9)
BTN = (FONT, 11, "bold")


# ── ttk-Style global setzen ──────────────────────────────────────────────────
def apply_theme(root) -> ttk.Style:
    """Konfiguriert ttk einmalig app-weit. Idempotent pro Root."""
    style = ttk.Style(root)
    try:
        style.theme_use("clam")
    except Exception:
        pass

    # Basis-Klassen mitthemen: so erben ALLE ttk-Widgets (auch ohne style=)
    # automatisch den modernen Look - zentraler Chokepoint statt pro Widget.
    for tv in ("Treeview", "NMG.Treeview"):
        style.configure(tv, background=CARD, fieldbackground=CARD, foreground=INK,
                        borderwidth=0, rowheight=30, font=(FONT, 10))
        style.map(tv, background=[("selected", SELECT_BG)], foreground=[("selected", PRIMARY)])
    for th in ("Treeview.Heading", "NMG.Treeview.Heading"):
        style.configure(th, background="#EEF3F8", foreground=PRIMARY, relief="flat",
                        font=(FONT, 10, "bold"), padding=6)
        style.map(th, background=[("active", "#E2EAF2")])

    for cb in ("TCombobox", "NMG.TCombobox"):
        style.configure(cb, fieldbackground=CARD, background=CARD, foreground=INK,
                        arrowcolor=PRIMARY, bordercolor=BORDER, relief="flat", padding=5)
    style.configure("TEntry", fieldbackground=CARD, foreground=INK,
                    bordercolor=BORDER, relief="flat", padding=4)
    style.configure("TButton", background="#EDF1F6", foreground=PRIMARY, relief="flat",
                    font=(FONT, 10, "bold"), padding=(12, 6), borderwidth=0)
    style.map("TButton", background=[("active", "#E2E9F1")])
    for sb in ("Vertical.TScrollbar", "Horizontal.TScrollbar"):
        style.configure(sb, troughcolor=BG, background="#C7D2DD", bordercolor=BG,
                        arrowcolor=PRIMARY, relief="flat")

    style.configure("NMG.Vertical.TScrollbar", troughcolor=BG, background="#C7D2DD",
                    bordercolor=BG, arrowcolor=PRIMARY, relief="flat")
    style.configure("NMG.Horizontal.TProgressbar", troughcolor="#E8EDF2",
                    background=ACCENT, thickness=18, borderwidth=0)
    style.configure("NMG.TCombobox", fieldbackground=CARD, background=CARD,
                    foreground=INK, arrowcolor=PRIMARY, bordercolor=BORDER, padding=5)
    style.configure("NMG.TNotebook", background=BG, borderwidth=0)
    style.configure("NMG.TNotebook.Tab", background="#E3EAF1", foreground=MUTED,
                    padding=(16, 8), font=(FONT, 10, "bold"))
    style.map("NMG.TNotebook.Tab",
              background=[("selected", CARD)], foreground=[("selected", PRIMARY)])
    return style


def load_icon(path, size=64):
    """Laedt ein .ico/.png als scharfes Tk-Bild in Zielgroesse (PIL Lanczos).
    Rueckgabe muss vom Aufrufer referenziert werden (sonst GC). None bei Fehler.
    """
    try:
        from PIL import Image, ImageTk
        im = Image.open(path)
        im.load()
        im = im.convert("RGBA")
        if im.size != (size, size):
            im = im.resize((size, size), Image.LANCZOS)
        return ImageTk.PhotoImage(im)
    except Exception:
        return None


def apply_widget_defaults(root):
    """Default-Optik fuer klassische tk-Widgets ueber die Option-Datenbank.

    Wirkt nur auf Widgets, die danach erzeugt werden und die Option NICHT
    explizit setzen. Dadurch bekommen ungestylte Buttons (z.B. 'Datei
    auswaehlen') automatisch den flachen Look, waehrend farbige Buttons
    (primary/danger) ihre explizite Farbe behalten.
    """
    root.option_add("*Button.relief", "flat")
    root.option_add("*Button.background", "#E7EDF4")
    root.option_add("*Button.foreground", PRIMARY)
    root.option_add("*Button.activeBackground", "#DCE6F0")
    root.option_add("*Button.activeForeground", PRIMARY)
    root.option_add("*Button.cursor", "hand2")
    root.option_add("*Button.borderWidth", 1)
    root.option_add("*Button.padX", 12)
    root.option_add("*Button.padY", 6)


def style_treeview(tree: ttk.Treeview):
    """Setzt eine bestehende Treeview auf den NMG-Stil."""
    try:
        tree.configure(style="NMG.Treeview")
    except Exception:
        pass


# ── Widgets ──────────────────────────────────────────────────────────────────
class Card(tk.Frame):
    """Weisse Karte mit dezentem Rahmen und Innenabstand."""
    def __init__(self, master, padding=18, bg=CARD, **kw):
        super().__init__(master, bg=bg, highlightbackground=BORDER,
                         highlightthickness=1, bd=0, **kw)
        self._bg = bg
        self.inner = tk.Frame(self, bg=bg)
        self.inner.pack(fill="both", expand=True, padx=padding, pady=padding)


DISABLED_FILL = "#C3CCD6"
DISABLED_FG = "#8A97A5"


def _darken(hex_color, factor=0.84):
    """Dunklere Variante einer Hex-Farbe (fuer den Hover-Zustand)."""
    try:
        h = hex_color.lstrip("#")
        r, g, b = (int(h[i:i + 2], 16) for i in (0, 2, 4))
        return "#%02X%02X%02X" % (int(r * factor), int(g * factor), int(b * factor))
    except Exception:
        return hex_color

# Cache fuer abgerundete Hintergrundbilder (haelt zugleich die PhotoImage-Refs,
# sonst wuerden sie vom Garbage Collector eingesammelt und verschwinden).
_ROUND_CACHE: dict = {}


def _rounded_image(w, h, radius, fill, outline, bg):
    """Liefert ein abgerundetes Rechteck als PhotoImage (Pillenform/Card-Ecke).

    Die Ecken werden in der Eltern-Farbe ``bg`` gefuellt, damit der Button auf
    seinem Untergrund nahtlos rund wirkt. 3x-Supersampling fuer saubere Kanten.
    """
    w = max(1, int(w)); h = max(1, int(h))
    radius = max(0, min(int(radius), w // 2, h // 2))
    key = (w, h, radius, fill, outline, bg)
    img = _ROUND_CACHE.get(key)
    if img is not None:
        return img
    if len(_ROUND_CACHE) > 400:        # Deckel gegen Speicherwachstum beim Resize
        _ROUND_CACHE.clear()
    s = 3
    im = Image.new("RGB", (w * s, h * s), bg)
    d = ImageDraw.Draw(im)
    line = s if outline else 0
    d.rounded_rectangle([line // 2, line // 2, w * s - 1 - line // 2, h * s - 1 - line // 2],
                        radius=radius * s, fill=fill, outline=outline,
                        width=max(1, line))
    im = im.resize((w, h), Image.LANCZOS)
    photo = ImageTk.PhotoImage(im)
    _ROUND_CACHE[key] = photo
    return photo


class PillButton(tk.Label):
    """Flacher, moderner Button mit abgerundeten Ecken und Hover-Effekt.

    Bleibt ein ``tk.Label`` (volle Label-API: ``config(text=...)``, ``pack``,
    ``cget`` …), bekommt aber als Hintergrund ein per Pillow gezeichnetes,
    abgerundetes Rechteck. Das Bild wird bei Groessen-/Textaenderung neu
    erzeugt, sodass auch ``pack(fill="x")``-Buttons rund bleiben. Ohne Pillow
    faellt das Widget automatisch auf die fruehere eckige Optik zurueck.
    """
    KINDS = {
        "primary": (PRIMARY, "#0D5596", "#FFFFFF"),
        "accent": (ACCENT, ACCENT_DARK, "#FFFFFF"),
        "success": (SUCCESS, SUCCESS_DARK, "#FFFFFF"),
        "danger": (DANGER, DANGER_DARK, "#FFFFFF"),
        "warning": (WARNING, "#A86E00", "#FFFFFF"),
        "ghost": (CARD, "#EEF3F8", PRIMARY),
        "neutral": ("#EDF1F6", "#E2E9F1", PRIMARY),
    }
    RADIUS = 10

    def __init__(self, master, text, command=None, kind="primary",
                 font_size=11, padx=20, pady=10, color=None, **kw):
        if color is not None:                  # freie Grundfarbe (z.B. aus gui)
            base, hover, fg = color, _darken(color), "#FFFFFF"
        else:
            base, hover, fg = self.KINDS.get(kind, self.KINDS["primary"])
        self._base, self._hover, self._fg = base, hover, fg
        self._outline = BORDER if (color is None and kind in ("ghost", "neutral")) else None
        self._padx, self._pady = padx, pady
        self._command = command
        self._enabled = True
        self._hovering = False
        self._last_key = None
        try:
            self._parent_bg = master.cget("background")
        except Exception:
            self._parent_bg = CARD
        # PIL kann nur Hex-Farben parsen; System-Farbnamen -> heller Fallback.
        if not (isinstance(self._parent_bg, str) and self._parent_bg.startswith("#")):
            self._parent_bg = CARD

        if _PIL_OK:
            super().__init__(master, text=text, bg=self._parent_bg, fg=fg,
                             font=(FONT, font_size, "bold"), bd=0,
                             highlightthickness=0, compound="center",
                             cursor="hand2", **kw)
            self._render()
            self.bind("<Configure>", lambda e: self._render())
        else:
            # Fallback ohne Pillow: alte eckige Optik.
            super().__init__(master, text=text, bg=base, fg=fg,
                             font=(FONT, font_size, "bold"), padx=padx, pady=pady,
                             cursor="hand2", **kw)
            if self._outline:
                self.config(highlightbackground=BORDER, highlightthickness=1)
        self.bind("<Button-1>", self._click)
        self.bind("<Enter>", self._enter)
        self.bind("<Leave>", self._leave)

    # -- Rendering ----------------------------------------------------------
    def _natural_size(self):
        f = tkfont.Font(font=self.cget("font"))
        return (f.measure(self.cget("text") or "") + 2 * self._padx,
                f.metrics("linespace") + 2 * self._pady)

    def _render(self):
        if not _PIL_OK:
            return
        w, h = self.winfo_width(), self.winfo_height()
        if w <= 1 or h <= 1:
            w, h = self._natural_size()
        # Auf ein Raster ABrunden: das Bild bleibt damit immer <= der zugeteilten
        # Flaeche (kann den Container nie aufblaehen -> keine Oszillation) und es
        # entstehen nur wenige verschiedene Bilder beim Resize (Cache-schonend).
        w = max(8, (w // 6) * 6)
        h = max(8, (h // 4) * 4)
        if not self._enabled:
            fill, outline = DISABLED_FILL, None
        else:
            fill = self._hover if self._hovering else self._base
            outline = self._outline
        key = (w, h, fill, outline, self._parent_bg)
        if key == self._last_key:
            return
        self._last_key = key
        self._img = _rounded_image(w, h, self.RADIUS, fill, outline, self._parent_bg)
        self.config(image=self._img)

    def _enter(self, _e):
        if self._enabled:
            self._hovering = True
            if _PIL_OK:
                self._render()
            else:
                self.config(bg=self._hover)

    def _leave(self, _e):
        self._hovering = False
        if self._enabled:
            if _PIL_OK:
                self._render()
            else:
                self.config(bg=self._base)

    def _click(self, _e):
        if self._enabled and self._command:
            self._command()

    def configure_command(self, command):
        self._command = command

    def set_enabled(self, on: bool):
        self._enabled = on
        self.config(fg=self._fg if on else DISABLED_FG,
                    cursor="hand2" if on else "arrow")
        if _PIL_OK:
            self._render()
        else:
            self.config(bg=self._base if on else DISABLED_FILL)


class Sidebar(tk.Frame):
    """Dunkle, vertikale Navigation mit Logo-Kopf, Eintraegen und Fuss.

    add_item(key, icon, label, command, active=False) -> registriert einen
    klickbaren Eintrag. set_active(key) hebt den aktiven Eintrag hervor.
    """
    # Helle, auf dem Navy-Grund gut lesbare Icon-Farben (zyklisch je Eintrag),
    # damit die Sidebar-Icons farbig wirken wie die Kacheln im Cockpit.
    ICON_COLORS = ["#4FA8E0", "#46C28A", "#F0A93B", "#E8736B",
                   "#B59BE6", "#3FC4C4", "#E6C04F"]

    def __init__(self, master, width=250, title="NMGone", subtitle="", **kw):
        super().__init__(master, bg=SIDEBAR, width=width, **kw)
        self.pack_propagate(False)
        self._items: dict[str, dict] = {}
        self._active: str | None = None
        self._icon_i = 0

        head = tk.Frame(self, bg=SIDEBAR)
        head.pack(fill="x", pady=(18, 6), padx=16)
        # Logo sitzt direkt auf dem dunklen Grund - die weisse Outline in
        # NMGone.png sorgt fuer Kontrast (frueher: weisse Karte noetig).
        self._logo_card = tk.Frame(head, bg=SIDEBAR)
        self._title_lbl = tk.Label(head, text=title, bg=SIDEBAR, fg="#FFFFFF",
                                   font=(FONT, 16, "bold"))
        self._title_lbl.pack(anchor="w")
        self._subtitle_lbl = None
        if subtitle:
            self._subtitle_lbl = tk.Label(head, text=subtitle, bg=SIDEBAR,
                                          fg=SIDEBAR_MUTED, font=(FONT, 9))
            self._subtitle_lbl.pack(anchor="w")
        tk.Frame(self, bg="#16395C", height=1).pack(fill="x", padx=16, pady=(12, 8))

        # scrollbarer Bereich fuer viele Eintraege
        self._body = tk.Frame(self, bg=SIDEBAR)
        self._body.pack(fill="both", expand=True)

        self._foot = tk.Frame(self, bg=SIDEBAR)
        self._foot.pack(side="bottom", fill="x")

    def set_logo(self, image):
        """Zeigt das Logo oben (statt des Titeltextes) direkt auf dem dunklen
        Sidebar-Grund; die weisse Outline in NMGone.png liefert den Kontrast."""
        try:
            self._title_lbl.pack_forget()
            if self._subtitle_lbl is not None:
                self._logo_card.pack(fill="x", pady=(0, 4), before=self._subtitle_lbl)
            else:
                self._logo_card.pack(fill="x", pady=(0, 4))
            lbl = tk.Label(self._logo_card, image=image, bg=SIDEBAR)
            lbl.image = image
            lbl.pack(padx=12, pady=8)
        except Exception:
            pass

    def add_section(self, label: str):
        tk.Label(self._body, text=label.upper(), bg=SIDEBAR, fg=SIDEBAR_MUTED,
                 font=(FONT, 8, "bold"), anchor="w").pack(fill="x", padx=22, pady=(12, 2))

    def add_item(self, key, icon, label, command, active=False, color=None):
        """Klickbarer Eintrag: farbiges Icon + Text (getrennte Labels, damit das
        Icon eine eigene Farbe tragen kann). ``color`` setzt die Icon-Farbe
        explizit; ohne Angabe wird zyklisch aus ICON_COLORS gewaehlt."""
        if color is None:
            color = self.ICON_COLORS[self._icon_i % len(self.ICON_COLORS)]
            self._icon_i += 1
        row = tk.Frame(self._body, bg=SIDEBAR, cursor="hand2")
        row.pack(fill="x", padx=10, pady=1)
        ic = tk.Label(row, text=icon, bg=SIDEBAR, fg=color, font=(FONT, 12),
                      width=2, anchor="center")
        ic.pack(side="left", padx=(10, 4), pady=8)
        tx = tk.Label(row, text=label, bg=SIDEBAR, fg=SIDEBAR_TEXT, font=(FONT, 10),
                      anchor="w")
        tx.pack(side="left", fill="x", expand=True, pady=8)
        rec = {"row": row, "icon": ic, "text": tx, "color": color, "sub": False}
        self._items[key] = rec
        for w in (row, ic, tx):
            w.bind("<Button-1>", lambda e, k=key, c=command: (self.set_active(k), c and c()))
            w.bind("<Enter>", lambda e, k=key: self._on_enter(k))
            w.bind("<Leave>", lambda e, k=key: self._on_leave(k))
        if active:
            self.set_active(key)
        return row

    def add_subitem(self, key, label, command):
        """Eingerückter Unterpunkt (für aufklappbare Gruppen). Wird NICHT automatisch
        gepackt - der Aufrufer blendet ihn über pack()/pack_forget() ein/aus."""
        if not hasattr(self, "_subkeys"):
            self._subkeys = set()
        self._subkeys.add(key)
        b = tk.Label(self._body, text=f"          {label}", bg=SIDEBAR, fg=SIDEBAR_MUTED,
                     font=(FONT, 10), anchor="w", padx=12, pady=6, cursor="hand2")
        rec = {"row": b, "icon": None, "text": b, "color": None, "sub": True}
        self._items[key] = rec
        b.bind("<Button-1>", lambda e, k=key, c=command: (self.set_active(k), c and c()))
        b.bind("<Enter>", lambda e, k=key: self._on_enter(k))
        b.bind("<Leave>", lambda e, k=key: self._on_leave(k))
        return b

    def footer(self):
        """Fuss-Container (unten in der Sidebar) fuer eigene Buttons/Widgets."""
        return self._foot

    def body(self):
        """Navigations-Container (oben) fuer eigene Widgets, z.B. ein Modus-Badge
        oberhalb der Eintraege. Wird vor add_item() befuellt."""
        return self._body

    def add_footer_note(self, text):
        tk.Label(self._foot, text=text, bg=SIDEBAR, fg=SIDEBAR_MUTED,
                 font=(FONT, 9), justify="left", anchor="w").pack(anchor="w", padx=22, pady=14)

    # -- Hover/Active ohne Flackern (Bg auf Zeile + Kind-Labels setzen) ------
    def _row_bg(self, rec, bg):
        rec["row"].config(bg=bg)
        if rec["icon"] is not None:
            rec["icon"].config(bg=bg)
        rec["text"].config(bg=bg)

    def _on_enter(self, key):
        if self._active != key:
            self._row_bg(self._items[key], SIDEBAR_ACTIVE)

    def _on_leave(self, key):
        if self._active != key:
            self._row_bg(self._items[key], SIDEBAR)

    def set_active(self, key):
        self._active = key
        for k, rec in self._items.items():
            on = (k == key)
            self._row_bg(rec, SIDEBAR_ACTIVE if on else SIDEBAR)
            size = 10
            if rec["sub"]:
                fg = "#FFFFFF" if on else SIDEBAR_MUTED
            else:
                fg = "#FFFFFF" if on else SIDEBAR_TEXT
            rec["text"].config(fg=fg, font=(FONT, size, "bold" if on else "normal"))
            if rec["icon"] is not None:           # Icon behaelt immer seine Farbe
                rec["icon"].config(fg=rec["color"])


def page_header(parent, title, subtitle="", bg=BG):
    """Standard-Seitenkopf: grosser Titel + optionaler Untertitel."""
    head = tk.Frame(parent, bg=bg)
    tk.Label(head, text=title, bg=bg, fg=INK, font=(FONT, 18, "bold")).pack(anchor="w")
    if subtitle:
        tk.Label(head, text=subtitle, bg=bg, fg=MUTED, font=(FONT, 11)).pack(anchor="w", pady=(2, 0))
    return head


def tile(parent, icon, title, desc, button_text, command, color=PRIMARY):
    """Moderne Dashboard-Kachel als Karte mit Icon, Text und Aktion."""
    card = tk.Frame(parent, bg=CARD, highlightbackground=BORDER, highlightthickness=1)
    tk.Label(card, text=icon, font=(FONT, 32), bg=CARD, fg=color).pack(pady=(20, 4))
    tk.Label(card, text=title, font=(FONT, 14, "bold"), bg=CARD, fg=INK).pack()
    tk.Label(card, text=desc, wraplength=190, justify="center", bg=CARD,
             fg=MUTED, font=(FONT, 10)).pack(padx=14, pady=10)
    PillButton(card, button_text + "  →", command, kind="primary",
               font_size=11, padx=14, pady=8).pack(fill="x", padx=18, pady=(6, 18))
    return card


def divider(parent, bg=BG):
    return tk.Frame(parent, bg=DIVIDER, height=1)
