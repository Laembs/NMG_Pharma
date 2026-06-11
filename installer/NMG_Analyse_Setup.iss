#define MyAppName "NMG Analyse"
#define MyAppVersion "3.1.0"
#define MyAppPublisher "NMG Pharma"
#define MyAppExeName "start.bat"

[Setup]
AppId={{4F03D930-7E50-4DB2-9B16-NMGANALYSE310}}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL=https://www.nmg-pharma.de
DefaultDirName={autopf}\NMG Analyse
DefaultGroupName=NMG Analyse
DisableProgramGroupPage=yes
OutputDir=..\dist_setup
OutputBaseFilename=NMG_Analyse_Setup_3_1_0
SetupIconFile=..\assets\nmg_logo.ico
Compression=lzma
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=lowest
UninstallDisplayIcon={app}\assets\nmg_logo.ico
VersionInfoVersion=3.1.0.0
VersionInfoCompany=NMG Pharma
VersionInfoDescription=NMG Analyse Setup
VersionInfoProductName=NMG Analyse
VersionInfoProductVersion=3.1.0

[Languages]
Name: "german"; MessagesFile: "compiler:Languages\German.isl"

[Tasks]
Name: "desktopicon"; Description: "Desktop-Verknüpfung erstellen"; GroupDescription: "Zusätzliche Symbole:"; Flags: checkedonce

[Dirs]
Name: "{commonappdata}\NMG Analyse"
Name: "{commonappdata}\NMG Analyse\data"
Name: "{commonappdata}\NMG Analyse\ausgaben"
Name: "{commonappdata}\NMG Analyse\gespeicherte_analysen"
Name: "{commonappdata}\NMG Analyse\backups"
Name: "{commonappdata}\NMG Analyse\updates"
Name: "{commonappdata}\NMG Analyse\logs"

[Files]
Source: "..\app\*"; DestDir: "{app}\app"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "..\assets\*"; DestDir: "{app}\assets"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "..\start.py"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\start.bat"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\requirements.txt"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\version.json"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\README.md"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\updates\*"; DestDir: "{commonappdata}\NMG Analyse\updates"; Flags: ignoreversion recursesubdirs createallsubdirs
; Startdaten nur bei Erstinstallation kopieren. Bestehende Daten bleiben erhalten.
Source: "..\data\*"; DestDir: "{commonappdata}\NMG Analyse\data"; Flags: ignoreversion recursesubdirs createallsubdirs onlyifdoesntexist

[Icons]
Name: "{autoprograms}\NMG Analyse"; Filename: "{app}\start.bat"; WorkingDir: "{app}"; IconFilename: "{app}\assets\nmg_logo.ico"
Name: "{autodesktop}\NMG Analyse"; Filename: "{app}\start.bat"; WorkingDir: "{app}"; IconFilename: "{app}\assets\nmg_logo.ico"; Tasks: desktopicon

[Run]
Filename: "{cmd}"; Parameters: "/C echo {{> ""{app}\install_config.json"" & echo   ""data_root"": ""{commonappdata}\NMG Analyse"" >> ""{app}\install_config.json"" & echo }}>> ""{app}\install_config.json"""; Flags: runhidden
Filename: "{app}\start.bat"; Description: "NMG Analyse starten"; Flags: nowait postinstall skipifsilent

[Code]
function InitializeSetup(): Boolean;
var
  Msg: String;
begin
  Msg := 'NMG Analyse 3.1.0 Recovery wird installiert oder aktualisiert.' + #13#10#13#10 +
         'Vorhandene Daten unter ProgramData bleiben erhalten. Bei Erstinstallation werden die Recovery-Startdaten aus Prototyp 2.6 kopiert.';
  MsgBox(Msg, mbInformation, MB_OK);
  Result := True;
end;
