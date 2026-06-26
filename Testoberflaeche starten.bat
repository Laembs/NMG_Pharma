@echo off
REM Startet die NMGone-Testoberflaeche (Vorschau) als eigenes Fenster.
REM Das Hauptprogramm (start.py) bleibt davon unberuehrt.
REM Genutzt wird die normale Datenbank data\nmg_startdatenbank.sqlite,
REM damit echte Kunden, Bestellungen und Auswertungen sichtbar sind.
cd /d "%~dp0"
echo ============================================
echo   NMGone - Testoberflaeche (Vorschau)
echo   Kunden ^& ABC, Karte, Markt-Insights, ...
echo ============================================
python start_testoberflaeche.py %*
pause
