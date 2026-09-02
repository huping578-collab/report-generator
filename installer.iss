#define MyAppName "报告生成工具"
#ifndef AppVersion
  #error AppVersion must be supplied by build-release.bat
#endif

[Setup]
AppId={{D4B6D72C-6E6B-4DAB-9A1B-2B4C8C7F1F6A}
AppName={#MyAppName}
AppVersion={#AppVersion}
AppVerName={#MyAppName} {#AppVersion}
AppPublisher=京炜交通
DefaultDirName={localappdata}\Programs\报告生成工具
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
PrivilegesRequired=lowest
ArchitecturesAllowed=x64compatible
OutputDir=dist
OutputBaseFilename=报告生成工具-Setup
SetupIconFile=assets\app.ico
UninstallDisplayIcon={app}\报告生成工具.exe
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
CloseApplications=yes
RestartApplications=no
VersionInfoVersion={#AppVersion}

[Files]
Source: "dist\报告生成工具\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs; Excludes: "templates\*;logs\*"
Source: "templates\*"; DestDir: "{app}\templates"; Flags: ignoreversion recursesubdirs createallsubdirs uninsneveruninstall; Excludes: "*.docx"
Source: "dist\updater\updater.exe"; DestDir: "{app}"; DestName: "更新程序.exe"; Flags: ignoreversion

[Dirs]
Name: "{app}\logs"

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\报告生成工具.exe"; WorkingDir: "{app}"
Name: "{group}\检查报告生成工具更新"; Filename: "{app}\更新程序.exe"; WorkingDir: "{app}"
Name: "{commondesktop}\{#MyAppName}"; Filename: "{app}\报告生成工具.exe"; WorkingDir: "{app}"

[Run]
Filename: "{app}\报告生成工具.exe"; Description: "启动 {#MyAppName}"; Flags: nowait postinstall skipifsilent
