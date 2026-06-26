@echo off
REM Startet die Parameter- & Berechtigungs-App als TESTUMGEBUNG:
REM eigene Test-Datenbank, Demo-Mitarbeiter, Admin-Modus sofort offen (volle Rechte).
REM Die echte NMGone-Datenbank wird dabei NICHT veraendert.
cd /d "%~dp0"
echo ==================================================
echo   Parameter ^& Berechtigungen - TESTUMGEBUNG
echo   (eigene Test-DB, volle Rechte, Demo-Mitarbeiter)
echo ==================================================
python start_parameter_test.py %*
pause
