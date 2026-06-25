@echo off
REM Startet die Parameter- & Berechtigungs-App (wer darf was) direkt.
REM Teilt sich die Datenbank mit NMGone (app/config.py).
cd /d "%~dp0"
echo ==================================================
echo   Parameter ^& Berechtigungen - wer darf was
echo ==================================================
python start_parameter.py %*
pause
