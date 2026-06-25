@echo off
REM Startet die Meldungen-App (GDP-Meldewesen, Kuehlsachenkontrolle &
REM Selbstinspektion) direkt im DEMO-MODUS mit der Demo-Datenbank inkl.
REM Demodaten. Die Produktiv-DB nmg_startdatenbank.sqlite bleibt unberuehrt.
cd /d "%~dp0"
set NMGONE_DEMO=1
echo ==================================================
echo   Meldungen (GDP-Meldewesen) - DEMO-MODUS
echo   Datenbank: data\nmg_demodatenbank.sqlite
echo ==================================================
python start_meldungen.py %*
pause
