[Setup]
AppName=XLStatUDF
AppVersion=1.0.0
DefaultDirName={autopf}\XLStatUDF
DefaultGroupName=XLStatUDF
OutputBaseFilename=XLStatUDF_Setup
Compression=lzma2
SolidCompression=yes
PrivilegesRequired=lowest

[Files]
Source: "XLStatUDF-packed.xll"; DestDir: "{userappdata}\Microsoft\Excel\XLSTART"; Flags: ignoreversion

[Icons]
Name: "{group}\Odinstalovat XLStatUDF"; Filename: "{uninstallexe}"

[Code]
procedure RegisterAddin();
var
  RegKey: string;
  AddinPath: string;
begin
  RegKey := 'Software\Microsoft\Office\16.0\Excel\Options';
  AddinPath := ExpandConstant('{userappdata}\Microsoft\Excel\XLSTART\XLStatUDF-packed.xll');
  RegWriteStringValue(HKCU, RegKey, 'OPEN', '/R "' + AddinPath + '"');
end;

procedure CurStepChanged(CurStep: TSetupStep);
begin
  if CurStep = ssPostInstall then
    RegisterAddin();
end;
