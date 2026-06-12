#define MyAppName "NMGone"
#define MyAppDisplayName "NMGone V1.0"
#define MyAppVersion "1.0.0"
#define MyAppPublisher "NMG Pharma"
#define MyAppExeName "NMGone.exe"

[Setup]
AppId={{8A2C6F40-1A77-4C2E-9F33-NMGONE000100}}
AppName={#MyAppDisplayName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL=https://www.nmg-pharma.de
DefaultDirName={autopf}\NMGone
DefaultGroupName=NMGone
DisableProgramGroupPage=yes
OutputDir=..\dist_setup
OutputBaseFilename=NMGone_Setup_1_0_0
SetupIconFile=..\assets\nmg_logo.ico
Compression=lzma
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=dialog
UninstallDisplayIcon={app}\NMGone.exe
VersionInfoVersion=1.0.0.0
VersionInfoCompany=NMG Pharma
VersionInfoDescription=NMGone Setup
VersionInfoProductName=NMGone
VersionInfoProductVersion=1.0.0

[Languages]
Name: "german"; MessagesFile: "compiler:Languages\German.isl"

[Tasks]
Name: "desktopicon"; Description: "Desktop-Verknüpfung erstellen"; GroupDescription: "Zusätzliche Symbole:"; Flags: checkedonce

[Dirs]
Name: "{commonappdata}\NMGone"
Name: "{commonappdata}\NMGone\data"
Name: "{commonappdata}\NMGone\ausgaben"
Name: "{commonappdata}\NMGone\gespeicherte_analysen"
Name: "{commonappdata}\NMGone\backups"
Name: "{commonappdata}\NMGone\updates"
Name: "{commonappdata}\NMGone\logs"

[Files]
; Das komplette PyInstaller-Output-Verzeichnis. Enthaelt NMGone.exe + alle
; eingebetteten Python-Laufzeit-DLLs und Pakete. Python muss auf dem
; Zielrechner NICHT installiert sein.
Source: "..\dist\NMGone\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{autoprograms}\NMGone"; Filename: "{app}\{#MyAppExeName}"; WorkingDir: "{app}"; IconFilename: "{app}\_internal\assets\nmg_logo.ico"
Name: "{autodesktop}\NMGone"; Filename: "{app}\{#MyAppExeName}"; WorkingDir: "{app}"; IconFilename: "{app}\_internal\assets\nmg_logo.ico"; Tasks: desktopicon

[Run]
; install_config.json zeigt der App, wo Nutzerdaten liegen sollen.
; Hinweis: {{ ist Inno-Setup-Escape fuer { ; } muss nicht escapt werden.
Filename: "{cmd}"; Parameters: "/C echo {{> ""{app}\install_config.json"" & echo   ""data_root"": ""{commonappdata}\NMGone"" >> ""{app}\install_config.json"" & echo }>> ""{app}\install_config.json"""; Flags: runhidden
Filename: "{app}\{#MyAppExeName}"; Description: "NMGone starten"; Flags: nowait postinstall skipifsilent

[Code]
function InitializeSetup(): Boolean;
var
  Msg: String;
begin
  Msg := 'NMGone V1.0 wird installiert oder aktualisiert.' + #13#10#13#10 +
         'Vorhandene Daten unter C:\ProgramData\NMGone bleiben erhalten.';
  MsgBox(Msg, mbInformation, MB_OK);
  Result := True;
end;
