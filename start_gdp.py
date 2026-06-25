"""Eigenstaendiger Start der NMG GDP-App (Wareneingang, Chargen-Rueckverfolgung,
Kuehlkette & Retouren).

Eigener Prozess mit eigenem Taskleisten-Icon, teilt sich die Datenbank mit
NMGone (app/config.py -> ProgramData/NMGone). Die Logik liegt in
app.gdp_app.run_standalone (wird auch von NMGone.exe --gdp genutzt).
"""
import os
import sys

# Repo-Root in den Pfad, damit das 'app'-Paket im Dev-Modus gefunden wird.
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))


def main():
    from app.gdp_app import run_standalone
    run_standalone()


if __name__ == "__main__":
    main()
