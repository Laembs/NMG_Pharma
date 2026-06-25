"""Eigenstaendiger Start der NMG Einkauf-App (Beschaffung EU-Ausland & Margen).

Eigener Prozess mit eigenem Taskleisten-Icon, teilt sich die Datenbank mit
NMGone (app/config.py). Die Logik liegt in app.einkauf_app.run_standalone.
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))


def main():
    from app.einkauf_app import run_standalone
    run_standalone()


if __name__ == "__main__":
    main()
