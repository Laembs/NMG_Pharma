"""Eigenstaendiger Start von NMGone (Analyse-Programm fuer den Vertrieb).

Eigener Prozess mit eigenem Taskleisten-Icon. Die Logik liegt in app.gui.main.
Wird auch vom Cockpit (start_cockpit.py) genutzt, um das NMGone-Modul "Analysen"
zu oeffnen.
"""
import os
import sys

# Repo-Root in den Pfad, damit das 'app'-Paket im Dev-Modus gefunden wird.
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))


def main():
    from app.gui import main as run
    run()


if __name__ == "__main__":
    main()
