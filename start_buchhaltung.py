"""Eigenstaendiger Start der NMG Buchhaltungs-App (Vorerfassung & Export
an die Buchhaltung / das Steuerbuero, eRechnung-Belege, Kontenrahmen).

Eigener Prozess mit eigenem Taskleisten-Icon, teilt sich die Datenbank mit
NMGone (app/config.py -> ProgramData/NMGone). Die Logik liegt in
app.buchhaltung_app.run_standalone (wird auch von NMGone.exe --buchhaltung genutzt).
"""
import os
import sys

# Repo-Root in den Pfad, damit das 'app'-Paket im Dev-Modus gefunden wird.
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))


def main():
    from app.buchhaltung_app import run_standalone
    run_standalone()


if __name__ == "__main__":
    main()
