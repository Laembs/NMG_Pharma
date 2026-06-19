#define MyAppName "NMGone"
#define MyAppDisplayName "NMGone V1.1 SP12"
#define MyAppVersion "1.1.12"
#define MyAppPublisher "Laembs Software"
#define MyAppExeName "NMGone.exe"
#define MyAppId "{{8A2C6F40-1A77-4C2E-9F33-NMGONE000100}"

[Setup]
AppId={#MyAppId}
AppName={#MyAppDisplayName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL=https://github.com/Laembs/NMG_Pharma
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=
DefaultDirName={localappdata}\Programs\NMGone
DefaultGroupName=NMGone
DisableProgramGroupPage=yes
OutputDir=..\dist_setup
OutputBaseFilename=NMGone_Setup_1_1_12
SetupIconFile=..\assets\NMGone.ico
Compression=lzma
SolidCompression=yes
WizardStyle=modern
UninstallDisplayIcon={app}\NMGone.exe
VersionInfoVersion=1.1.12.0
VersionInfoCompany=Laembs Software
VersionInfoDescription=NMGone Setup
VersionInfoProductName=NMGone
VersionInfoProductVersion=1.1.12
CloseApplications=yes
RestartApplications=no

[Languages]
Name: "german"; MessagesFile: "compiler:Languages\German.isl"

[Dirs]
Name: "{commonappdata}\NMGone"
Name: "{commonappdata}\NMGone\data"
Name: "{commonappdata}\NMGone\ausgaben"
Name: "{commonappdata}\NMGone\gespeicherte_analysen"
Name: "{commonappdata}\NMGone\gespeicherte_analysen\PK"
Name: "{commonappdata}\NMGone\gespeicherte_analysen\ZF"
Name: "{commonappdata}\NMGone\backups"
Name: "{commonappdata}\NMGone\updates"
Name: "{commonappdata}\NMGone\logs"

[InstallDelete]
Type: files; Name: "{userdesktop}\NMGone.lnk"
Type: files; Name: "{userstartmenu}\Programs\NMGone\NMGone.lnk"
Type: files; Name: "{userprograms}\NMGone\NMGone.lnk"

[Files]
Source: "..\dist\NMGone\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{userprograms}\NMGone"; Filename: "{app}\{#MyAppExeName}"; WorkingDir: "{app}"; IconFilename: "{app}\_internal\assets\NMGone.ico"
Name: "{userdesktop}\NMGone"; Filename: "{app}\{#MyAppExeName}"; WorkingDir: "{app}"; IconFilename: "{app}\_internal\assets\NMGone.ico"

[Run]
Filename: "{cmd}"; Parameters: "/C echo {{> ""{app}\install_config.json"" & echo   ""data_root"": ""{commonappdata}\NMGone"" >> ""{app}\install_config.json"" & echo }>> ""{app}\install_config.json"""; Flags: runhidden
Filename: "{app}\{#MyAppExeName}"; Description: "NMGone starten"; Flags: nowait postinstall skipifsilent shellexec

[Code]
function GetUninstallString(RootKey: Integer): String;
var
  RegKey: String;
  Cmd: String;
begin
  RegKey := 'Software\Microsoft\Windows\CurrentVersion\Uninstall\{#MyAppId}_is1';
  Cmd := '';
  if not RegQueryStringValue(RootKey, RegKey, 'UninstallString', Cmd) then
    Cmd := '';
  Result := Cmd;
end;

function RunOldUninstaller(RootKey: Integer): Boolean;
var
  Cmd: String;
  Filename, Params: String;
  Quote: Char;
  Pos1: Integer;
  ResultCode: Integer;
begin
  Result := False;
  Cmd := GetUninstallString(RootKey);
  if Cmd = '' then Exit;
  Quote := '"';
  if Cmd[1] = Quote then
  begin
    Pos1 := Pos(Quote, Copy(Cmd, 2, Length(Cmd))) + 1;
    Filename := Copy(Cmd, 2, Pos1 - 2);
    Params := Trim(Copy(Cmd, Pos1 + 1, Length(Cmd)));
  end else begin
    Filename := Cmd;
    Params := '';
  end;
  Exec(Filename, '/VERYSILENT /SUPPRESSMSGBOXES /NORESTART', '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
  Result := True;
end;

function InitializeSetup(): Boolean;
var
  Msg: String;
begin
  RunOldUninstaller(HKCU);
  Msg := 'NMGone V1.1 SP12 wird installiert.' + #13#10#13#10 +
         'Service Pack 12 zu V1.1.0 - grosses Bug-Fix-Sammlung.' + #13#10#13#10 +
         'Gespeicherte Analysen:' + #13#10 +
         '- Refresh nach Admin-Loeschen funktioniert wieder' + #13#10 +
         '- "Item end not found"-Fehler bei Seitenwechsel weg' + #13#10 +
         '- Manuelle Imports laufen jetzt mit Busy-Dialog' + #13#10#13#10 +
         'Dashboard:' + #13#10 +
         '- Editor-Fenster groesser, Speicher-Button immer sichtbar' + #13#10 +
         '- "Alles abwaehlen" funktioniert jetzt wirklich' + #13#10 +
         '  (vorher: kam beim naechsten Oeffnen alles zurueck)' + #13#10 +
         '- Info-Bereich verschwindet komplett bei allen abgewaehlt' + #13#10#13#10 +
         'Globale Suche auf Startseite:' + #13#10 +
         '- Ergebnis-Tabelle nur sichtbar wenn Treffer da' + #13#10 +
         '- Initial nur 3 Zeilen hoch (Toplevel-Tool bleibt gross)' + #13#10#13#10 +
         'Kalender:' + #13#10 +
         '- Popup-Kalender schliesst beim Klick ausserhalb' + #13#10#13#10 +
         'Neue Ordner-Struktur:' + #13#10 +
         '- ausgaben/Produktanalyse/<Jahr>/Q<n>/' + #13#10 +
         '- ausgaben/PK/<Jahr>/Q<n>/  (Auswertungen)' + #13#10 +
         '- ausgaben/Abweichungsanalyse/<Jahr>/Q<n>/' + #13#10 +
         '- importierte_analysen/PK|ZF/<Jahr>/Q<n>/  (manuelle Imports' + #13#10 +
         '  inklusive Original-Datei-Kopie)';
  MsgBox(Msg, mbInformation, MB_OK);
  Result := True;
end;
