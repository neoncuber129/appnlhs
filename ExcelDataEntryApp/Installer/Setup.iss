; Inno Setup script - Excel Data Entry
#define MyAppName "Excel Data Entry"
#define MyAppVersion "1.0.0"
#define MyAppPublisher "VMMU"
#define MyAppExeName "ExcelDataEntryApp.exe"
#define MyAppURL ""

[Setup]
AppId={{8D5D0F0D-AC35-4471-BB78-611FAC1FFE28}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppVerName={#MyAppName} {#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
AllowNoIcons=yes
DisableProgramGroupPage=yes
OutputDir=out
OutputBaseFilename=ExcelDataEntrySetup
SetupIconFile=app.ico
Compression=lzma2/ultra64
SolidCompression=yes
WizardStyle=modern
WizardSizePercent=120
WizardImageFile=wizard-image.bmp
WizardSmallImageFile=wizard-small.bmp
PrivilegesRequired=admin
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
UninstallDisplayIcon={app}\{#MyAppExeName}
UninstallDisplayName={#MyAppName}
VersionInfoVersion={#MyAppVersion}
VersionInfoCompany={#MyAppPublisher}
VersionInfoProductName={#MyAppName}
VersionInfoDescription=Cài đặt {#MyAppName}
ShowLanguageDialog=no
DisableWelcomePage=no
DisableDirPage=no
DisableFinishedPage=no
CloseApplications=force
RestartApplications=no
MinVersion=10.0

[Languages]
Name: "vietnamese"; MessagesFile: "Vietnamese.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: checkedonce

[Files]
Source: "payload\{#MyAppExeName}"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{autoprograms}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(MyAppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent

[Code]
function InitializeSetup(): Boolean;
begin
  Result := True;
end;
