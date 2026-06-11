@echo off
chcp 65001 >nul
echo =====================================================
echo   NMG Analyse Installer Build 3.1.0 Recovery
echo =====================================================
echo.
pip install -r requirements.txt pyinstaller
python -m PyInstaller --noconfirm --windowed --name "NMG_Analyse" --add-data "data;data" --add-data "assets;assets" start.py

echo.
echo EXE liegt unter dist\NMG_Analyse\
echo.
if exist "%ProgramFiles(x86)%\Inno Setup 6\ISCC.exe" (
  echo Inno Setup gefunden. Setup.exe wird gebaut...
  "%ProgramFiles(x86)%\Inno Setup 6\ISCC.exe" installer\NMG_Analyse_Setup_3_1_0.iss
) else (
  echo Inno Setup nicht gefunden.
  echo Die Vorlage liegt unter installer\NMG_Analyse_Setup_3_1_0.iss
)
pause
