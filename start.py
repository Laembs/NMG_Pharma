import sys

if __name__ == "__main__":
    # NMGone.exe --kasse startet direkt die Kasse (eigene Verknuepfung im Setup).
    if "--kasse" in sys.argv:
        from app.kasse_app import run_standalone
        run_standalone()
    else:
        from app.gui import main
        main()
