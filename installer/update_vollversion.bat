@echo off
setlocal EnableExtensions EnableDelayedExpansion
chcp 65001 >nul

echo =====================================================
echo   NMG Analyse Vollversions-Update 3.0.1
echo =====================================================
echo.
set "DEFAULT_APP=%LOCALAPPDATA%\NMG Analyse"
set /p "APPDIR=Installationsordner der bestehenden Version [Enter = %DEFAULT_APP%]: "
if "%APPDIR%"=="" set "APPDIR=%DEFAULT_APP%"

if not exist "%APPDIR%\app" (
  echo Keine bestehende Installation gefunden unter:
  echo %APPDIR%
  echo Bitte install_nmg_analyse.bat fuer eine Neuinstallation verwenden.
  pause
  exit /b 1
)

for /f "tokens=2 delims=:" %%a in ('findstr /i "data_root" "%APPDIR%\install_config.json" 2^>nul') do set "RAW=%%a"
set "DATADIR="
if defined RAW (
  set "DATADIR=!RAW:,= !"
  set "DATADIR=!DATADIR:"=!"
  set "DATADIR=!DATADIR:  = !"
)
if not defined DATADIR set "DATADIR=%LOCALAPPDATA%\NMG Analyse Daten"

set "STAMP=%DATE:~-4%%DATE:~3,2%%DATE:~0,2%_%TIME:~0,2%%TIME:~3,2%%TIME:~6,2%"
set "STAMP=%STAMP: =0%"
set "BACKUP=%DATADIR%\backups\Vollupdate_3_0_1_vorher_%STAMP%"
mkdir "%BACKUP%" 2>nul

echo Sicherheitskopie wird erstellt...
if exist "%DATADIR%\data" xcopy "%DATADIR%\data" "%BACKUP%\data" /E /I /Y >nul
if exist "%DATADIR%\gespeicherte_analysen" xcopy "%DATADIR%\gespeicherte_analysen" "%BACKUP%\gespeicherte_analysen" /E /I /Y >nul

echo Programmdateien werden aktualisiert. Daten bleiben erhalten...
xcopy "%~dp0..\app" "%APPDIR%\app" /E /I /Y >nul
xcopy "%~dp0..\assets" "%APPDIR%\assets" /E /I /Y >nul
copy "%~dp0..\start.py" "%APPDIR%\start.py" /Y >nul
copy "%~dp0..\start.bat" "%APPDIR%\start.bat" /Y >nul
copy "%~dp0..\requirements.txt" "%APPDIR%\requirements.txt" /Y >nul
copy "%~dp0..\version.json" "%APPDIR%\version.json" /Y >nul

if exist "%~dp0..\updates" xcopy "%~dp0..\updates" "%DATADIR%\updates" /E /I /Y >nul

echo.
echo Update abgeschlossen.
echo Sicherheitskopie:
echo   %BACKUP%
echo.
echo NMG Analyse wird gestartet...
start "" "%APPDIR%\start.bat"
endlocal
