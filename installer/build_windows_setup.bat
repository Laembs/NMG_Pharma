@echo off
setlocal EnableExtensions
chcp 65001 >nul

echo =====================================================
echo   NMG Analyse - Windows Setup bauen
echo   Version 3.0.1
echo =====================================================
echo.

set "ISCC=%ProgramFiles(x86)%\Inno Setup 6\ISCC.exe"
if not exist "%ISCC%" set "ISCC=%ProgramFiles%\Inno Setup 6\ISCC.exe"

if not exist "%ISCC%" (
  echo Inno Setup 6 wurde nicht gefunden.
  echo Bitte Inno Setup 6 installieren und diese Datei erneut starten.
  echo Download: https://jrsoftware.org/isdl.php
  pause
  exit /b 1
)

if not exist "%~dp0..\dist_setup" mkdir "%~dp0..\dist_setup"
"%ISCC%" "%~dp0NMG_Analyse_Setup_3_0_1.iss"

if errorlevel 1 (
  echo.
  echo Setup konnte nicht gebaut werden.
  pause
  exit /b 1
)

echo.
echo Fertig. Setup-Datei liegt hier:
echo   %~dp0..\dist_setup\NMG_Analyse_Setup_3_0_1.exe
echo.
pause
endlocal
