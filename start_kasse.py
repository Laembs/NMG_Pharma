"""Eigenstaendiger Start der NMG Kasse-App (Verkauf + Wareneingang).

Eigener Prozess mit eigenem Taskleisten-Icon, teilt sich die Datenbank mit
NMGone (app/config.py -> ProgramData/NMGone). Wird als zweite .exe gepackt
(NMGone_Kasse) und bekommt eine eigene Verknuepfung.
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
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("NMG.Kasse")
        except Exception:
            pass

    import tkinter as tk
    from app.config import ASSETS_DIR
    from app.kasse_app import KassePanel

    try:
        from app.migrations import run_migrations
        run_migrations()
    except Exception:
        pass

    root = tk.Tk()
    root.title("NMG Kasse")
    root.geometry("980x640")
    root.minsize(820, 560)
    root.configure(bg="#ffffff")
    try:
        root.iconbitmap(str(ASSETS_DIR / "kasse.ico"))
    except Exception:
        pass

    KassePanel(root, on_close=root.destroy).pack(fill="both", expand=True)
    root.mainloop()


if __name__ == "__main__":
    main()
