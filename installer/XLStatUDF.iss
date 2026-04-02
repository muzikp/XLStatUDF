#define MyAppName "XLStatUDF"
#ifndef MyAppVersion
  #define MyAppVersion "1.0.0"
#endif

[Setup]
AppId={{C2E3810A-3B2E-4E61-B0D1-4E0BB59D81C2}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher=XLStatUDF
DefaultDirName={localappdata}\Programs\XLStatUDF
DefaultGroupName=XLStatUDF
UninstallDisplayIcon={app}\XLStatUDF-packed.xll
OutputDir=..\artifacts\installer
OutputBaseFilename=XLStatUDF_Setup
Compression=lzma2
SolidCompression=yes
PrivilegesRequired=lowest
ArchitecturesInstallIn64BitMode=x64compatible
WizardStyle=modern
DisableProgramGroupPage=yes

[Languages]
Name: "czech"; MessagesFile: "compiler:Languages\Czech.isl"
Name: "english"; MessagesFile: "compiler:Default.isl"

[Files]
Source: "XLStatUDF-packed.xll"; DestDir: "{userappdata}\Microsoft\Excel\XLSTART"; DestName: "XLStatUDF-packed.xll"; Flags: ignoreversion

[Icons]
Name: "{autoprograms}\XLStatUDF\Odinstalovat XLStatUDF"; Filename: "{uninstallexe}"; Languages: czech
Name: "{autoprograms}\XLStatUDF\Uninstall XLStatUDF"; Filename: "{uninstallexe}"; Languages: english

[Run]
Filename: "{userappdata}\Microsoft\Excel\XLSTART"; Flags: postinstall shellexec skipifsilent unchecked; Description: "Otevrit slozku XLSTART"; Languages: czech
Filename: "{userappdata}\Microsoft\Excel\XLSTART"; Flags: postinstall shellexec skipifsilent unchecked; Description: "Open XLSTART folder"; Languages: english

[UninstallDelete]
Type: files; Name: "{userappdata}\Microsoft\Excel\XLSTART\XLStatUDF-packed.xll"
