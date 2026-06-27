#define MyAppName "NMGone"
#define MyAppDisplayName "NMGone V2.2 SP1"
#define MyAppVersion "2.2.1"
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
OutputBaseFilename=NMGone_Setup_2_2_1
SetupIconFile=..\assets\NMGone.ico
Compression=lzma
SolidCompression=yes
WizardStyle=modern
UninstallDisplayIcon={app}\NMGone.exe
VersionInfoVersion=2.2.1.0
VersionInfoCompany=Laembs Software
VersionInfoDescription=NMGone Setup
VersionInfoProductName=NMGone
VersionInfoProductVersion=2.2.0
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
Name: "{commonappdata}\NMGone\gespeicherte_analysen\ZW"
Name: "{commonappdata}\NMGone\backups"
Name: "{commonappdata}\NMGone\updates"
Name: "{commonappdata}\NMGone\logs"
Name: "{commonappdata}\NMGone\vorlagen"

[InstallDelete]
Type: files; Name: "{userdesktop}\NMGone.lnk"
Type: files; Name: "{userstartmenu}\Programs\NMGone\NMGone.lnk"
Type: files; Name: "{userprograms}\NMGone\NMGone.lnk"

[Files]
Source: "..\dist\NMGone\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
; AppUserModelID muss zur prozess-internen ID passen (gui.main: NMG.NMGone,
; kasse run_standalone: NMG.Kasse) -> getrennte Taskleisten-Gruppen + Icons.
Name: "{userprograms}\NMGone"; Filename: "{app}\{#MyAppExeName}"; WorkingDir: "{app}"; IconFilename: "{app}\_internal\assets\NMGone.ico"; AppUserModelID: "NMG.NMGone"
Name: "{userdesktop}\NMGone"; Filename: "{app}\{#MyAppExeName}"; WorkingDir: "{app}"; IconFilename: "{app}\_internal\assets\NMGone.ico"; AppUserModelID: "NMG.NMGone"
; NMG Kasse als eigene Verknuepfung (gleiche EXE mit --kasse, eigenes Icon).
Name: "{userprograms}\NMG Kasse"; Filename: "{app}\{#MyAppExeName}"; Parameters: "--kasse"; WorkingDir: "{app}"; IconFilename: "{app}\_internal\assets\kasse.ico"; AppUserModelID: "NMG.Kasse"
Name: "{userdesktop}\NMG Kasse"; Filename: "{app}\{#MyAppExeName}"; Parameters: "--kasse"; WorkingDir: "{app}"; IconFilename: "{app}\_internal\assets\kasse.ico"; AppUserModelID: "NMG.Kasse"
; NMG Wareneingang & Retouren (gleiche EXE mit --gdp, eigenes Icon).
Name: "{userprograms}\NMG Wareneingang & Retouren"; Filename: "{app}\{#MyAppExeName}"; Parameters: "--gdp"; WorkingDir: "{app}"; IconFilename: "{app}\_internal\assets\GDP.ico"; AppUserModelID: "NMG.GDP"
Name: "{userdesktop}\NMG Wareneingang & Retouren"; Filename: "{app}\{#MyAppExeName}"; Parameters: "--gdp"; WorkingDir: "{app}"; IconFilename: "{app}\_internal\assets\GDP.ico"; AppUserModelID: "NMG.GDP"

[Run]
Filename: "{cmd}"; Parameters: "/C echo {{> ""{app}\install_config.json"" & echo   ""data_root"": ""{commonappdata}\NMGone"" >> ""{app}\install_config.json"" & echo }>> ""{app}\install_config.json"""; Flags: runhidden
; Auto-Restart nach Update (ohne "postinstall" => startet IMMER nach Setup,
; skipifsilent schliesst den silent-Install aus).
Filename: "{app}\{#MyAppExeName}"; Flags: nowait skipifsilent shellexec

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
  Msg := 'NMGone V2.2 SP1 wird installiert.' + #13#10#13#10 +
         'Neuer Meilenstein - fasst den V2.1-Zyklus zusammen und bringt' + #13#10 +
         'eine neue Welle eigenstaendiger Apps:' + #13#10 +
         '- Buchhaltung: Vorerfassung + DATEV-Export ans Steuerbuero' + #13#10 +
         '- eRechnung-Empfang (XRechnung/ZUGFeRD)' + #13#10 +
         '- Kunden (CRM): Steckbrief, ABC-Analyse, Offline-Deutschlandkarte' + #13#10 +
         '- Zweite Programmsprache Slowakisch' + #13#10 +
         '- Revisions-Uebersicht "Was wurde veraendert" + neue Testoberflaeche' + #13#10 +
         '- Ausgebautes Cockpit, Berechtigungen, Theme und Hilfe';
  MsgBox(Msg, mbInformation, MB_OK);
  Result := True;
end;
