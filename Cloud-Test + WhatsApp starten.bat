@echo off
title NMGone Cloud-Test (mit WhatsApp)
echo ============================================================
echo   NMGone Cloud-Test wird gestartet (mit WhatsApp-Adresse)...
echo ============================================================
echo.

REM Fenster 1: der Test-Server
start "NMGone Test-Server" cmd /k python "C:\nmg_analyse_3_1_0 - Kopie\cloud_test_server.py"

REM kurz warten, damit der Server zuerst laeuft
timeout /t 2 >nul

REM Fenster 2: Tunnel + WhatsApp-Benachrichtigung
start "Cloudflare Tunnel + WhatsApp" cmd /k python "C:\nmg_analyse_3_1_0 - Kopie\start_tunnel_notify.py"

echo Zwei Fenster muessten sich geoeffnet haben.
echo Die Adresse kommt automatisch per WhatsApp (wenn whatsapp_config.txt ausgefuellt ist)
echo und steht ausserdem im Tunnel-Fenster.
echo.
echo Dieses Fenster kannst du schliessen.
pause
