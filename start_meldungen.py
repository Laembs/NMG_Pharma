"""Eigenstaendiger Start der NMG Meldungen-App (GDP-Meldewesen,
Kuehlsachenkontrolle & Selbstinspektion).

Eigener Prozess mit eigenem Taskleisten-Icon, teilt sich die Datenbank mit
NMGone (app/config.py -> ProgramData/NMGone). Die Logik liegt in
app.meldungen_app.run_standalone (wird auch von NMGone.exe --meldungen genutzt).
"""
import os
import sys

# Repo-Root in den Pfad, damit das 'app'-Paket im Dev-Modus gefunden wird.
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))


def main():
    from app.meldungen_app import run_standalone
    run_standalone()


if __name__ == "__main__":
    main()
