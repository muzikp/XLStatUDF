#define MyAppName "XLStatUDF"
#ifndef MyAppVersion
  #define MyAppVersion "1.0.0"
#endif

#ifndef InstallerCulture
  #define InstallerCulture "CS"
#endif

#ifndef AddInSource
  #define AddInSource "XLStatUDF-packed.xll"
#endif

#if InstallerCulture == "EN"
  #define InstallerLanguageName "english"
  #define InstallerMessagesFile "compiler:Default.isl"
  #define InstallerSuffix "EN"
  #define InstallerOpenXlStartText "Open XLSTART folder"
  #define InstallerUninstallShortcutText "Uninstall XLStatUDF"
#else
  #define InstallerLanguageName "czech"
  #define InstallerMessagesFile "compiler:Languages\\Czech.isl"
  #define InstallerSuffix "CS"
  #define InstallerOpenXlStartText "Otevrit slozku XLSTART"
  #define InstallerUninstallShortcutText "Odinstalovat XLStatUDF"
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
OutputBaseFilename=XLStatUDF_{#InstallerSuffix}_Setup
Compression=lzma2
SolidCompression=yes
PrivilegesRequired=lowest
ArchitecturesInstallIn64BitMode=x64compatible
WizardStyle=modern
DisableProgramGroupPage=yes
ShowLanguageDialog=no

[Languages]
Name: "{#InstallerLanguageName}"; MessagesFile: "{#InstallerMessagesFile}"

[Files]
Source: "{#AddInSource}"; DestDir: "{userappdata}\Microsoft\Excel\XLSTART"; DestName: "XLStatUDF-packed.xll"; Flags: ignoreversion

[Icons]
Name: "{autoprograms}\XLStatUDF\{#InstallerUninstallShortcutText}"; Filename: "{uninstallexe}"

[Run]
Filename: "{userappdata}\Microsoft\Excel\XLSTART"; Flags: postinstall shellexec skipifsilent unchecked; Description: "{#InstallerOpenXlStartText}"

[UninstallDelete]
Type: files; Name: "{userappdata}\Microsoft\Excel\XLSTART\XLStatUDF-packed.xll"
