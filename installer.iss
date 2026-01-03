; ---------------------------------------------
; BasicAppLauncher Installer Script
; ---------------------------------------------

[Setup]
AppName=BasicAppLauncher
AppVersion=1.0
DefaultDirName={pf}\BasicAppLauncher
DefaultGroupName=BasicAppLauncher
OutputBaseFilename=BasicAppLauncherInstaller
Compression=lzma
SolidCompression=yes
AllowNoIcons=yes
DisableProgramGroupPage=no
LicenseFile=EULA.txt

[Files]
Source: "main.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "config.json"; DestDir: "{app}"; Flags: ignoreversion
Source: "CascadiaCode.ttf"; DestDir: "{app}"; Flags: ignoreversion
Source: "readme.txt"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\BasicAppLauncher"; Filename: "{app}\main.exe"; IconFilename: "{app}\main.exe"

[Tasks]
Name: "startup"; Description: "Run BasicAppLauncher on Windows startup"; GroupDescription: "Additional Tasks"; Flags: unchecked

[Run]
Filename: "{app}\main.exe"; Description: "Launch BasicAppLauncher"; Flags: nowait postinstall skipifsilent

[Code]
procedure CurStepChanged(CurStep: TSetupStep);
begin
  if (CurStep = ssPostInstall) and IsTaskSelected('startup') then
  begin
    RegWriteStringValue(
      HKCU,
      'Software\Microsoft\Windows\CurrentVersion\Run',
      'BasicAppLauncher',
      '"' + ExpandConstant('{app}\main.exe') + '"'
    );
  end;
end;
