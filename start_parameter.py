"""Eigenstaendiger Start der NMGone Parameter- & Berechtigungs-App.

Eigener Prozess, teilt sich die Datenbank mit NMGone (app/config.py).
Die Logik liegt in app.parameter_app.run_standalone.
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))


def main():
    from app.parameter_app import run_standalone
    run_standalone()


if __name__ == "__main__":
    main()
