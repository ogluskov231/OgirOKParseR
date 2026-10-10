#define MyAppName "ОгрОК_UA"
#define MyAppVersion "0.6"
#define MyAppExeName "OgrOK_UA.exe"

[Setup]
AppId={{8F0E9F8D-7A4E-4D8E-9F2C-0601A0B0C0D1}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
DefaultDirName={autopf}\OgrOK_UA
DefaultGroupName={#MyAppName}
OutputDir=..\installer_output
OutputBaseFilename=OgrOK_UA_Setup_v{#MyAppVersion}
Compression=lzma
SolidCompression=yes
WizardStyle=modern
ArchitecturesInstallIn64BitMode=x64compatible
SetupIconFile=..\assets\icon.ico
UninstallDisplayIcon={app}\OgrOK_UA.exe

[Files]
Source: "..\dist\OgrOK_UA\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Dirs]
Name: "{app}\input\documents"
Name: "{app}\output"

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Tasks]
Name: "desktopicon"; Description: "Створити ярлик на робочому столі"; GroupDescription: "Додаткові ярлики:"

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "Запустити {#MyAppName}"; Flags: nowait postinstall skipifsilent
