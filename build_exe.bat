@echo off
set APPNAME=NMG_Analyse
python -m PyInstaller --noconfirm --windowed --name %APPNAME% --add-data "data;data" --add-data "assets;assets" start.py
echo.
echo Fertig. Die EXE liegt im Ordner dist\%APPNAME%
pause
