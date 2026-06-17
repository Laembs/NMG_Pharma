#define MyAppName "NMGone"
#define MyAppDisplayName "NMGone V1.0 SP12"
#define MyAppVersion "1.0.12"
#define MyAppPublisher "Laembs Software"
#define MyAppExeName "NMGone.exe"
#define MyAppId "{{8A2C6F40-1A77-4C2E-9F33-NMGONE000100}"

[Setup]
AppId={#MyAppId}
AppName={#MyAppDisplayName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL=https://github.com/Laembs/NMG_Pharma
; SP12: zurueck auf User-Mode-Install. Mitarbeiter-PCs sind oft so gesperrt,
; dass UAC ohne Admin-Passwort nicht freigegeben werden kann. Mit lowest
; und explizitem Pfad (kein {autopf}) bleibt der Pfad stabil und nichts
; driftet zwischen Updates - das Problem das SP9 loesen sollte ist hier
; auch geloest, nur ohne Admin.
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=
DefaultDirName={localappdata}\Programs\NMGone
DefaultGroupName=NMGone
DisableProgramGroupPage=yes
OutputDir=..\dist_setup
OutputBaseFilename=NMGone_Setup_1_0_12
SetupIconFile=..\assets\NMGone.ico
Compression=lzma
SolidCompression=yes
WizardStyle=modern
UninstallDisplayIcon={app}\NMGone.exe
VersionInfoVersion=1.0.12.0
VersionInfoCompany=Laembs Software
VersionInfoDescription=NMGone Setup
VersionInfoProductName=NMGone
VersionInfoProductVersion=1.0.12
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

; SP12: alte User-Verknuepfungen entfernen vor Neuinstall (admin/Program-Files-
; Reste muessen die IT manuell ueber "Einstellungen -> Apps" entfernen, dafuer
; haben wir keine Rechte).
[InstallDelete]
Type: files; Name: "{userdesktop}\NMGone.lnk"
Type: files; Name: "{userstartmenu}\Programs\NMGone\NMGone.lnk"
Type: files; Name: "{userprograms}\NMGone\NMGone.lnk"

[Files]
Source: "..\dist\NMGone\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

; SP12: Shortcut in User-Bereich (kein admin = kein commondesktop).
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
  CleanedHKCU: Boolean;
  HasAdminInstall: Boolean;
  Msg: String;
begin
  // Alte User-Installation entfernen (HKCU). Eine HKLM-Installation
  // koennen wir hier nicht entfernen - ohne admin keine Schreibrechte
  // in Program Files. Stattdessen geben wir dem Anwender einen Hinweis.
  CleanedHKCU := RunOldUninstaller(HKCU);
  HasAdminInstall := FileExists(ExpandConstant('{commonpf}\NMGone\NMGone.exe')) or
                     FileExists(ExpandConstant('{commonpf32}\NMGone\NMGone.exe'));

  Msg := 'NMGone V1.0 SP12 wird in deinem Benutzer-Profil installiert:' + #13#10 +
         '   C:\Users\<dein Login>\AppData\Local\Programs\NMGone' + #13#10#13#10 +
         'Keine Admin-Rechte noetig.';
  if CleanedHKCU then
    Msg := Msg + #13#10#13#10 + 'Eine aeltere Benutzer-Installation wurde entfernt.';
  if HasAdminInstall then
    Msg := Msg + #13#10#13#10 +
           'HINWEIS: Es existiert noch eine alte Admin-Installation unter' + #13#10 +
           '   C:\Program Files\NMGone oder C:\Program Files (x86)\NMGone' + #13#10 +
           'Diese bitte ueber "Windows-Einstellungen -> Apps -> NMGone -> Deinstallieren"' + #13#10 +
           'manuell entfernen (braucht Admin-Recht).';
  MsgBox(Msg, mbInformation, MB_OK);
  Result := True;
end;
