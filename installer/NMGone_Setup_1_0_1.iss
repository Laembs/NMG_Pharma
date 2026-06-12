#define MyAppName "NMGone"
#define MyAppDisplayName "NMGone V1.0 SP1"
#define MyAppVersion "1.0.1"
#define MyAppPublisher "NMG Pharma"
#define MyAppExeName "NMGone.exe"

[Setup]
; Gleiche AppId wie V1.0: damit V1.0 SP1 ueber V1.0 drueberinstalliert
; statt parallel zu landen. Daten unter C:\ProgramData\NMGone bleiben.
AppId={{8A2C6F40-1A77-4C2E-9F33-NMGONE000100}}
AppName={#MyAppDisplayName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL=https://www.nmg-pharma.de
DefaultDirName={autopf}\NMGone
DefaultGroupName=NMGone
DisableProgramGroupPage=yes
OutputDir=..\dist_setup
OutputBaseFilename=NMGone_Setup_1_0_1
SetupIconFile=..\assets\NMGone.ico
Compression=lzma
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=dialog
UninstallDisplayIcon={app}\NMGone.exe
VersionInfoVersion=1.0.1.0
VersionInfoCompany=NMG Pharma
VersionInfoDescription=NMGone Setup
VersionInfoProductName=NMGone
VersionInfoProductVersion=1.0.1

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
Source: "..\dist\NMGone\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{autoprograms}\NMGone"; Filename: "{app}\{#MyAppExeName}"; WorkingDir: "{app}"; IconFilename: "{app}\_internal\assets\NMGone.ico"
Name: "{autodesktop}\NMGone"; Filename: "{app}\{#MyAppExeName}"; WorkingDir: "{app}"; IconFilename: "{app}\_internal\assets\NMGone.ico"; Tasks: desktopicon

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
  Msg := 'NMGone V1.0 SP1 wird installiert oder ueber bestehende V1.0 aktualisiert.' + #13#10#13#10 +
         'Vorhandene Daten unter C:\ProgramData\NMGone bleiben erhalten.' + #13#10 +
         'Neu in SP1: 4 Import-Bugs behoben, Daten landen ab jetzt in den' + #13#10 +
         'richtigen Tabellen, Admin-Modus per Strg+Alt+A versteckt.';
  MsgBox(Msg, mbInformation, MB_OK);
  Result := True;
end;
