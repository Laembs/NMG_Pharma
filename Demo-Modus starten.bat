@echo off
REM Startet NMGone mit der DEMO-Datenbank (data\nmg_demodatenbank.sqlite).
REM Die Produktiv-DB nmg_startdatenbank.sqlite bleibt unberuehrt.
cd /d "%~dp0"
set NMGONE_DEMO=1
echo ============================================
echo   NMGone wird im DEMO-MODUS gestartet
echo   Datenbank: data\nmg_demodatenbank.sqlite
echo ============================================
python start.py %*
pause
