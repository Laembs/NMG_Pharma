#define MyAppName "NMGone"
#define MyAppDisplayName "NMGone V1.0 SP9"
#define MyAppVersion "1.0.9"
#define MyAppPublisher "Laembs Software"
#define MyAppExeName "NMGone.exe"
#define MyAppId "{{8A2C6F40-1A77-4C2E-9F33-NMGONE000100}"

[Setup]
AppId={#MyAppId}
AppName={#MyAppDisplayName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL=https://github.com/Laembs/NMG_Pharma
; SP9: ab jetzt zwingend nach C:\Program Files\NMGone.
; Vorher war PrivilegesRequired=lowest + Override=dialog -> User konnte
; einmal mit UAC nach Program Files installieren, beim Update ohne UAC nach
; AppData. Resultat: Desktop-Verknuepfung zeigte aufs falsche Verzeichnis.
; admin-elevation erzwingt einen stabilen Pfad.
PrivilegesRequired=admin
DefaultDirName={autopf}\NMGone
DefaultGroupName=NMGone
DisableProgramGroupPage=yes
OutputDir=..\dist_setup
OutputBaseFilename=NMGone_Setup_1_0_9
SetupIconFile=..\assets\NMGone.ico
Compression=lzma
SolidCompression=yes
WizardStyle=modern
UninstallDisplayIcon={app}\NMGone.exe
VersionInfoVersion=1.0.9.0
VersionInfoCompany=Laembs Software
VersionInfoDescription=NMGone Setup
VersionInfoProductName=NMGone
VersionInfoProductVersion=1.0.9
CloseApplications=yes
RestartApplications=no

[Languages]
Name: "german"; MessagesFile: "compiler:Languages\German.isl"

[Tasks]
Name: "desktopicon"; Description: "Desktop-Verknüpfung erstellen"; GroupDescription: "Zusätzliche Symbole:"; Flags: checkedonce

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

[Files]
Source: "..\dist\NMGone\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{autoprograms}\NMGone"; Filename: "{app}\{#MyAppExeName}"; WorkingDir: "{app}"; IconFilename: "{app}\_internal\assets\NMGone.ico"
Name: "{autodesktop}\NMGone"; Filename: "{app}\{#MyAppExeName}"; WorkingDir: "{app}"; IconFilename: "{app}\_internal\assets\NMGone.ico"; Tasks: desktopicon

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

function RunOldUninstaller(RootKey: Integer; Label_: String): Boolean;
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
  // Parsen: "C:\...\unins000.exe" /STANDARD-SWITCHES (Inno-Style)
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
  Log('SP9 cleanup: alte ' + Label_ + '-Installation gefunden, deinstalliere: ' + Filename);
  Exec(Filename, '/VERYSILENT /SUPPRESSMSGBOXES /NORESTART', '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
  Result := True;
end;

function InitializeSetup(): Boolean;
var
  CleanedHKCU, CleanedHKLM: Boolean;
  Msg: String;
begin
  CleanedHKCU := RunOldUninstaller(HKCU, 'Benutzer (AppData)');
  CleanedHKLM := RunOldUninstaller(HKLM, 'Admin (Program Files)');
  Msg := 'NMGone V1.0 SP9 wird in C:\Program Files\NMGone installiert.' + #13#10#13#10 +
         'Vorhandene Daten unter C:\ProgramData\NMGone bleiben erhalten.';
  if CleanedHKCU or CleanedHKLM then
    Msg := Msg + #13#10#13#10 + 'Eine aeltere Installation wurde zuvor entfernt.';
  MsgBox(Msg, mbInformation, MB_OK);
  Result := True;
end;
