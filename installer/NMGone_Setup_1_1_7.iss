#define MyAppName "NMGone"
#define MyAppDisplayName "NMGone V1.1 SP7"
#define MyAppVersion "1.1.7"
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
OutputBaseFilename=NMGone_Setup_1_1_7
SetupIconFile=..\assets\NMGone.ico
Compression=lzma
SolidCompression=yes
WizardStyle=modern
UninstallDisplayIcon={app}\NMGone.exe
VersionInfoVersion=1.1.7.0
VersionInfoCompany=Laembs Software
VersionInfoDescription=NMGone Setup
VersionInfoProductName=NMGone
VersionInfoProductVersion=1.1.7
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
  Msg := 'NMGone V1.1 SP7 wird installiert.' + #13#10#13#10 +
         'Service Pack 7 zu V1.1.0.' + #13#10#13#10 +
         'Startseite - Globale Suche statt Austauschsuche:' + #13#10 +
         '- altes Austausch-Suche-Widget entfernt' + #13#10 +
         '  (Vergleichs-Suche ist seit V1.1.0 das eigene Tool)' + #13#10 +
         '- neues Such-Widget sucht parallel in:' + #13#10 +
         '  - Kunden (PK + ZF, nach Nummer / Name / Apotheke)' + #13#10 +
         '  - gespeicherten Analysen (Apotheke / Datum / Kunde)' + #13#10 +
         '  - Artikeln (delegiert an die Vergleichs-Suche)' + #13#10 +
         '- Doppelklick auf Treffer verzweigt:' + #13#10 +
         '  - Kunde -> Kunden-Center mit Vorbefuellung' + #13#10 +
         '  - Analyse -> Excel oeffnen wenn vorhanden, sonst Liste' + #13#10 +
         '  - Artikel -> Vergleichs-Suche mit Vorbefuellung' + #13#10 +
         '- Worker-Thread + Debounce, UI bleibt klickbar';
  MsgBox(Msg, mbInformation, MB_OK);
  Result := True;
end;
