#define MyAppName "NMGone"
#define MyAppDisplayName "NMGone V1.1 SP19"
#define MyAppVersion "1.1.19"
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
OutputBaseFilename=NMGone_Setup_1_1_19
SetupIconFile=..\assets\NMGone.ico
Compression=lzma
SolidCompression=yes
WizardStyle=modern
UninstallDisplayIcon={app}\NMGone.exe
VersionInfoVersion=1.1.19.0
VersionInfoCompany=Laembs Software
VersionInfoDescription=NMGone Setup
VersionInfoProductName=NMGone
VersionInfoProductVersion=1.1.19
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
; V1.1 SP19: Auto-Restart nach Update. Vorher mit "postinstall" => User
; musste am Wizard-Ende ein Haekchen klicken. Jetzt ohne "postinstall" =>
; NMGone wird IMMER nach Setup-Abschluss gestartet (skipifsilent schliesst
; den silent-Install aus).
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
  Msg := 'NMGone V1.1 SP19 wird installiert.' + #13#10#13#10 +
         'Service Pack 19 zu V1.1.0.' + #13#10#13#10 +
         'Neue Auswertung 11x schneller:' + #13#10 +
         '- trace_lookup_source mit Index-Cache statt Pro-Zeile-Scan' + #13#10 +
         '- 414 Zeilen jetzt in ~16 s statt 3 Min' + #13#10#13#10 +
         'Bug-Fix Neue Auswertung (SP18-Regression):' + #13#10 +
         '- Doppelklick-Schutz auf "Auswertung starten"' + #13#10 +
         '- Status-Box-Fallback wenn Sidebar eingeklappt ist' + #13#10 +
         '- Exceptions im Background-Job werden geloggt + angezeigt' + #13#10#13#10 +
         'NMG-Rabatte-Uebersicht (neue Apps-Kachel):' + #13#10 +
         '- Tabelle aller Rabatte mit Filter' + #13#10 +
         '- Statistik (Anzahl, Min/Max/Avg, letzte Aktualisierung)' + #13#10 +
         '- Diff zum letzten Stand (neu/geaendert/entfernt)' + #13#10 +
         '- Verlauf pro PZN ueber alle Snapshots' + #13#10#13#10 +
         'Auto-Restart nach Setup:' + #13#10 +
         '- NMGone startet nach Update IMMER automatisch (kein Haekchen mehr noetig)';
  MsgBox(Msg, mbInformation, MB_OK);
  Result := True;
end;
