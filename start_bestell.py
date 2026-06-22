"""Eigenstaendiger Start der NMG Bestell-App.

Laeuft als eigener Prozess mit eigenem Taskleisten-Icon, teilt sich aber die
Datenbank mit NMGone (siehe app/config.py -> ProgramData/NMGone). Wird als
zweite .exe gepackt (NMGone_Bestellung) und bekommt eine eigene Verknuepfung.
"""
import os
import sys

# Repo-Root in den Pfad, damit das 'app'-Paket im Dev-Modus gefunden wird.
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))


def main():
    # Eigene Taskleisten-Identitaet -> eigene Gruppe + eigenes Icon, getrennt
    # von NMGone (sonst gruppiert Windows beide unter NMGone).
    if os.name == "nt":
        try:
            import ctypes
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("NMG.Bestellung")
        except Exception:
            pass

    import tkinter as tk
    from app.config import ASSETS_DIR
    from app.bestell_app import BestellPanel

    # Schema sicherstellen (DB wird beim Import von app.config angelegt/geseedet).
    try:
        from app.migrations import run_migrations
        run_migrations()
    except Exception:
        pass

    root = tk.Tk()
    root.title("NMG Bestellung")
    root.geometry("900x600")
    root.minsize(760, 520)
    root.configure(bg="#ffffff")
    try:
        root.iconbitmap(str(ASSETS_DIR / "bestell.ico"))
    except Exception:
        pass

    BestellPanel(root, on_close=root.destroy).pack(fill="both", expand=True)
    root.mainloop()


if __name__ == "__main__":
    main()
