NMG Analyse 3.0.1 – echter Windows-Installer

Ziel-Datei nach dem Build:
  dist_setup\NMG_Analyse_Setup_3_0_1.exe

So wird die Setup.exe erstellt:
1. Inno Setup 6 auf dem Entwickler-PC installieren.
2. Diese ZIP entpacken.
3. installer\build_windows_setup.bat starten.
4. Danach liegt die Setup.exe im Ordner dist_setup.

Für Mitarbeiterinnen ist später nur diese Datei wichtig:
  NMG_Analyse_Setup_3_0_1.exe

Doppelklick genügt. Bestehende Daten bleiben erhalten:
- Datenbank
- gespeicherte Analysen
- Backups
- Updates
- Logs

Programmdateien werden aktualisiert.
Nutzerdaten liegen getrennt unter:
  C:\ProgramData\NMG Analyse
