"""Eigenstaendiger Start der NMGone-Hilfe ("Handbuch").

Eigener Prozess mit eigenem Taskleisten-Icon. Die eigentliche Logik liegt in
app.hilfe_app.run_standalone (wird auch von NMGone.exe --hilfe genutzt). Das
Modul zeigt nur das bebilderte Handbuch - es liest und schreibt keine Daten.
"""
import os
import sys

# Repo-Root in den Pfad, damit das 'app'-Paket im Dev-Modus gefunden wird.
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))


def main():
    from app.hilfe_app import run_standalone
    run_standalone()


if __name__ == "__main__":
    main()
