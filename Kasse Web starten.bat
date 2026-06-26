@echo off
title NMGone Kasse (Web) Starter
echo ============================================================
echo   NMGone Kasse (Web) wird als eigenes Fenster gestartet...
echo ============================================================
echo.
echo   Login (Pilot-Demo):
echo     Firma:    muster-pharma-gmbh
echo     Benutzer: admin
echo     Passwort: demo123
echo.
cd /d "%~dp0"
python "%~dp0start_kasse_web.py"
