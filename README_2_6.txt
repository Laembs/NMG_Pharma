NMG Analyse 2.6.0
==================

Schwerpunkt dieser Version:
- Installations- und Update-System
- sichtbare Versionsverwaltung
- Datenbank-Schema-Version 1.1
- Update-Manager im Programm
- automatisches Backup vor jedem Update
- sichere Datenbank-Migrationen
- Updatepakete mit Endung .nmgupdate

Start:
1. ZIP entpacken
2. start.bat oder python start.py starten

Update installieren:
1. Im Programm links "Update installieren" anklicken
2. .nmgupdate-Datei auswählen
3. Bestätigen
4. Programm neu starten

Updatepaket erzeugen:
- create_update_package.bat ausführen
- Ergebnis im Ordner updates

Geschützte Bereiche bei Updates:
- data / Datenbank
- backups
- gespeicherte_analysen
- ausgaben
- updates
- logs

Diese Ordner werden durch Updates nicht überschrieben.
