#define MyAppName "Etiquettes CSN"
#define MyAppVersion "1.0.0"
#define MyAppPublisher "TonOrganisation"
#define MyAppExeName "EtiquettesCSN.exe"
[Setup]
AppId={{8C2B0B8F-3CC2-4B80-9E34-0E2F1E1D67AA}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={userappdata}\EtiquettesCSN
DisableDirPage=yes
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
OutputBaseFilename=Setup_EtiquettesCSN
Compression=lzma
SolidCompression=yes
WizardStyle=modern
[Languages]
Name: "french"; MessagesFile: "compiler:Languages\French.isl"
[Dirs]
Name: "{app}\data"
Name: "{app}\src\app\templates"
[Files]
Source: "dist\EtiquettesCSN.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "src\app\templates\*"; DestDir: "{app}\src\app\templates"; Flags: ignoreversion recursesubdirs createallsubdirs
[InstallDelete]
Type: files; Name: "{app}\data\app.db"
[Icons]
Name: "{userdesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon
[Tasks]
Name: "desktopicon"; Description: "Créer un raccourci sur le Bureau"; GroupDescription: "Raccourcis :"; Flags: unchecked
[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "Lancer {#MyAppName}"; Flags: nowait postinstall skipifsilent
