@echo off
chcp 65001 >nul
echo =====================================================
echo   NMGone Installer Build V1.1 SP16 (1.1.16)
echo =====================================================
echo.

echo [1/3] Python-Abhaengigkeiten sicherstellen...
pip install -r requirements.txt pyinstaller >nul

echo [2/3] PyInstaller: dist\NMGone\NMGone.exe bauen...
python -m PyInstaller --noconfirm --windowed --name "NMGone" --icon assets\NMGone.ico --add-data "assets;assets" --collect-all babel --collect-all tkcalendar start.py
if errorlevel 1 (
  echo PyInstaller fehlgeschlagen.
  pause
  exit /b 1
)

echo.
echo EXE liegt unter dist\NMGone\NMGone.exe
echo.

echo [3/3] Inno Setup: Setup.exe bauen...
set "ISCC=%ProgramFiles(x86)%\Inno Setup 6\ISCC.exe"
if not exist "%ISCC%" set "ISCC=%ProgramFiles%\Inno Setup 6\ISCC.exe"

if exist "%ISCC%" (
  "%ISCC%" installer\NMGone_Setup_1_1_16.iss
  if errorlevel 1 (
    echo Inno Setup fehlgeschlagen.
    pause
    exit /b 1
  )
  echo.
  echo Setup-Datei: dist_setup\NMGone_Setup_1_1_16.exe
) else (
  echo Inno Setup 6 nicht gefunden. Installiere via: winget install JRSoftware.InnoSetup
  echo Skript-Vorlage liegt unter installer\NMGone_Setup_1_1_16.iss
)
pause
