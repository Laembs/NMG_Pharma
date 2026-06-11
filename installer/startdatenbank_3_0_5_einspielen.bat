@echo off
setlocal EnableExtensions EnableDelayedExpansion
chcp 65001 >nul

echo =====================================================
echo   NMG Analyse - Startdatenbank 3.0.5 einspielen
echo =====================================================
echo.
echo ACHTUNG: Diese Routine sichert die vorhandene Datenbank und kopiert danach

echo die ausgelieferte aktuelle Startdatenbank in den Datenordner.
echo.
set "DATADIR=%PROGRAMDATA%\NMG Analyse"
if not exist "%DATADIR%" mkdir "%DATADIR%"
if not exist "%DATADIR%\data" mkdir "%DATADIR%\data"
if not exist "%DATADIR%\backups" mkdir "%DATADIR%\backups"

for /f "tokens=1-4 delims=. " %%a in ("%date%") do set TODAY=%%d%%c%%b
for /f "tokens=1-3 delims=:, " %%a in ("%time%") do set NOW=%%a%%b%%c
set "BK=%DATADIR%\backups\vor_startdaten_3_0_5_%TODAY%_%NOW%"
mkdir "%BK%" 2>nul

if exist "%DATADIR%\data\nmg_startdatenbank.sqlite" (
  copy "%DATADIR%\data\nmg_startdatenbank.sqlite" "%BK%\nmg_startdatenbank.sqlite" /Y >nul
)

echo Kopiere aktuelle Startdaten...
xcopy "%~dp0..\data" "%DATADIR%\data" /E /I /Y >nul

echo.
echo Fertig. Backup der alten Datenbank liegt hier:
echo %BK%
echo.
pause
endlocal
