@echo off
setlocal EnableExtensions EnableDelayedExpansion
chcp 65001 >nul

echo =====================================================
echo   NMG Analyse Setup 3.1.0
echo =====================================================
echo.

set "DEFAULT_APP=%LOCALAPPDATA%\NMG Analyse"
set "DEFAULT_DATA=%PROGRAMDATA%\NMG Analyse"
if not exist "%PROGRAMDATA%" set "DEFAULT_DATA=%LOCALAPPDATA%\NMG Analyse Daten"

echo Installationsordner Programm:
echo   %DEFAULT_APP%
echo Datenordner:
echo   %DEFAULT_DATA%
echo.
set /p "APPDIR=Programmordner verwenden oder eigenen Pfad eingeben [Enter = Standard]: "
if "%APPDIR%"=="" set "APPDIR=%DEFAULT_APP%"
set /p "DATADIR=Datenordner verwenden oder eigenen Pfad eingeben [Enter = Standard]: "
if "%DATADIR%"=="" set "DATADIR=%DEFAULT_DATA%"

echo.
echo Programm wird installiert nach:
echo   %APPDIR%
echo Daten werden gespeichert unter:
echo   %DATADIR%
echo.
pause

mkdir "%APPDIR%" 2>nul
mkdir "%DATADIR%" 2>nul
mkdir "%DATADIR%\data" 2>nul
mkdir "%DATADIR%\ausgaben" 2>nul
mkdir "%DATADIR%\gespeicherte_analysen" 2>nul
mkdir "%DATADIR%\backups" 2>nul
mkdir "%DATADIR%\updates" 2>nul
mkdir "%DATADIR%\logs" 2>nul

echo Kopiere Programmdateien...
xcopy "%~dp0..\app" "%APPDIR%\app" /E /I /Y >nul
xcopy "%~dp0..\assets" "%APPDIR%\assets" /E /I /Y >nul
copy "%~dp0..\start.py" "%APPDIR%\start.py" /Y >nul
copy "%~dp0..\start.bat" "%APPDIR%\start.bat" /Y >nul
copy "%~dp0..\requirements.txt" "%APPDIR%\requirements.txt" /Y >nul
copy "%~dp0..\version.json" "%APPDIR%\version.json" /Y >nul
copy "%~dp0..\README.md" "%APPDIR%\README.md" /Y >nul

if not exist "%APPDIR%\install_config.json" (
  echo {> "%APPDIR%\install_config.json"
  echo   "data_root": "%DATADIR:\=\\%" >> "%APPDIR%\install_config.json"
  echo }>> "%APPDIR%\install_config.json"
)

if not exist "%DATADIR%\data\nmg_startdatenbank.sqlite" (
  echo Ersteinrichtung: kopiere Startdatenbank und Stammdaten...
  xcopy "%~dp0..\data" "%DATADIR%\data" /E /I /Y >nul
) else (
  echo Bestehende Datenbank gefunden. Nutzerdaten bleiben erhalten.
)

if not exist "%DATADIR%\updates" mkdir "%DATADIR%\updates"
xcopy "%~dp0..\updates" "%DATADIR%\updates" /E /I /Y >nul

echo Desktop-Verknuepfung wird erstellt...
powershell -NoProfile -ExecutionPolicy Bypass -Command "$s=(New-Object -ComObject WScript.Shell).CreateShortcut([Environment]::GetFolderPath('Desktop')+'\\NMG Analyse.lnk'); $s.TargetPath='%APPDIR%\\start.bat'; $s.WorkingDirectory='%APPDIR%'; $s.IconLocation='%APPDIR%\\assets\\nmg_logo.ico'; $s.Save()" >nul 2>nul

echo.
echo Installation abgeschlossen.
echo Starten Sie NMG Analyse ueber die Desktop-Verknuepfung oder:
echo   %APPDIR%\start.bat
echo.
pause
endlocal
