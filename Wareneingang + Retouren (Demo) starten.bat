@echo off
REM Startet die GDP-App (Wareneingang, Chargen-Rueckverfolgung, Kuehlkette &
REM Retouren) direkt im DEMO-MODUS mit der Demo-Datenbank inkl. Demodaten.
REM Die Produktiv-DB nmg_startdatenbank.sqlite bleibt unberuehrt.
cd /d "%~dp0"
set NMGONE_DEMO=1
echo ==================================================
echo   Wareneingang ^& Retouren (GDP) - DEMO-MODUS
echo   Datenbank: data\nmg_demodatenbank.sqlite
echo ==================================================
python start_gdp.py %*
pause
