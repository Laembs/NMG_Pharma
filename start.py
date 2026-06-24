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
    else:
        from app.gui import main
        main()
