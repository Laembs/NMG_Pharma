@echo off
set APPNAME=NMGone
python -m PyInstaller --noconfirm --windowed --name %APPNAME% --icon assets\nmg_logo.ico --add-data "assets;assets" start.py
echo.
echo Fertig. Die EXE liegt im Ordner dist\%APPNAME%
pause
