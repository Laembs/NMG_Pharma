"""Eigenstaendiger Start der NMG Kunden-App (CRM fuer Apotheken-Kunden).

Eigener Prozess mit eigenem Taskleisten-Icon, teilt sich die Datenbank mit
NMGone (app/config.py · tbl_kunden_center). Die Logik liegt in
app.kunden_app.run_standalone.
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))


def main():
    from app.kunden_app import run_standalone
    run_standalone()


if __name__ == "__main__":
    main()
