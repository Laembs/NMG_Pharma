@echo off
title NMGone Cloud-Test Starter
echo ============================================================
echo   NMGone Cloud-Test wird gestartet...
echo ============================================================
echo.

REM Fenster 1: der Test-Server
start "NMGone Test-Server" cmd /k python "C:\nmg_analyse_3_1_0 - Kopie\cloud_test_server.py"

REM kurz warten, damit der Server zuerst laeuft
timeout /t 2 >nul

REM Fenster 2: der Cloudflare-Tunnel
start "Cloudflare Tunnel" cmd /k cloudflared tunnel --url http://localhost:8000

echo Es muessten sich ZWEI Fenster geoeffnet haben:
echo   1) Test-Server
echo   2) Cloudflare Tunnel  -- hier steht die https://...trycloudflare.com Adresse
echo.
echo Beide Fenster offen lassen, solange die Seite erreichbar sein soll.
echo Zum Stoppen: in beiden Fenstern Strg + C  oder die Fenster schliessen.
echo.
echo Dieses Fenster kannst du jetzt schliessen.
pause
