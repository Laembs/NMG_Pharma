@echo off
REM Startet die NMGone-Revisions-Uebersicht (Testoberflaeche).
REM Zeigt auf einen Blick, was in dieser Revision veraendert wurde:
REM neue Apps, Aenderungen, Aufraeum-Stand und die Roadmap.
REM Das Hauptprogramm (start.py) bleibt davon unberuehrt.
cd /d "%~dp0"
echo ==================================================
echo   NMGone - Revisions-Uebersicht (Was wurde veraendert)
echo   Neue Apps . Aenderungen . Aufraeumen . Roadmap
echo ==================================================
python start_revision.py %*
pause
