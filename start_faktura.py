"""Eigenstaendiger Start der NMG Faktura-App (Rechnungen & Gutschriften).

Eigener Prozess mit eigenem Taskleisten-Icon, teilt sich die Datenbank mit
NMGone (app/config.py -> ProgramData/NMGone). Die Logik liegt in
app.faktura_app.run_standalone.
"""
import os
import sys

# Repo-Root in den Pfad, damit das 'app'-Paket im Dev-Modus gefunden wird.
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))


def main():
    from app.faktura_app import run_standalone
    run_standalone()


if __name__ == "__main__":
    main()
