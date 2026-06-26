"""TESTUMGEBUNG der Parameter- & Berechtigungs-App.

Startet die App mit:
  * eigener Test-Datenbank (data/nmg_parameter_testumgebung.sqlite) –
    die echte NMGone-Datenbank wird NICHT berührt,
  * Demo-Mitarbeitern inkl. zugewiesener Rollen und Beispiel-Ausnahmen,
  * sofort offenem Admin-Modus (volle Rechte, keine PIN).

Zum Ausprobieren aller Funktionen ohne Risiko.
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))


def main():
    from app.parameter_app import run_testmode
    run_testmode()


if __name__ == "__main__":
    main()
