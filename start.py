import sys

if __name__ == "__main__":
    # NMGone.exe --kasse startet direkt die Kasse (eigene Verknuepfung im Setup).
    if "--kasse" in sys.argv:
        from app.kasse_app import run_standalone
        run_standalone()
    # NMGone.exe --faktura startet direkt die Faktura-App.
    elif "--faktura" in sys.argv:
        from app.faktura_app import run_standalone
        run_standalone()
    # NMGone.exe --report startet direkt das Auswertungsmodul.
    elif "--report" in sys.argv:
        from app.report_app import run_standalone
        run_standalone()
    # NMGone.exe --hilfe startet direkt das Hilfe-/Handbuch-Modul.
    elif "--hilfe" in sys.argv:
        from app.hilfe_app import run_standalone
        run_standalone()
    # NMGone.exe --gdp startet direkt die GDP-App (Wareneingang & Retouren).
    elif "--gdp" in sys.argv:
        from app.gdp_app import run_standalone
        run_standalone()
    # NMGone.exe --einkauf startet direkt die Einkauf-App (Beschaffung & Margen).
    elif "--einkauf" in sys.argv:
        from app.einkauf_app import run_standalone
        run_standalone()
    # NMGone.exe --meldungen startet direkt die Meldungen-App (GDP-Meldewesen).
    elif "--meldungen" in sys.argv:
        from app.meldungen_app import run_standalone
        run_standalone()
    # NMGone.exe --personal startet direkt die Mitarbeiter-/Personal-App.
    elif "--personal" in sys.argv:
        from app.personal_app import run_standalone
        run_standalone()
    # NMGone.exe --parameter startet direkt die Parameter-/Berechtigungs-App.
    elif "--parameter" in sys.argv:
        from app.parameter_app import run_standalone
        run_standalone()
    else:
        from app.gui import main
        main()
