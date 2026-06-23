"""Eigenstaendiger Start des NMG Auswertungs-/Report-Moduls.

Eigener Prozess mit eigenem Taskleisten-Icon, teilt sich die Datenbank mit
NMGone (app/config.py -> ProgramData/NMGone). Die eigentliche Logik liegt in
app.report_app.run_standalone (wird auch von NMGone.exe --report genutzt).

Das Modul liest die Daten nur (Verkaeufe/Kunden/Artikel) und aendert nichts am
Kundenstamm.
"""
import os
import sys

# Repo-Root in den Pfad, damit das 'app'-Paket im Dev-Modus gefunden wird.
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))


def main():
    from app.report_app import run_standalone
    run_standalone()


if __name__ == "__main__":
    main()
