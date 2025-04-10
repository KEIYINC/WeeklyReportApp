#define MyAppName "Weekly Report App"
#define MyAppVersion "1.0"
#define MyAppPublisher "Your Company Name"
#define MyAppExeName "WeeklyReportApp.exe"

[Setup]
AppId={{YOUR-APP-GUID}}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
AllowNoIcons=yes
OutputDir=Output
OutputBaseFilename=WeeklyReportApp-Setup
Compression=lzma
SolidCompression=yes
WizardStyle=modern

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"
Name: "turkish"; MessagesFile: "compiler:Languages\Turkish.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
Source: "bin\Release\net6.0-windows\win-x64\publish\{#MyAppExeName}"; DestDir: "{app}"; Flags: ignoreversion
Source: "bin\Release\net6.0-windows\win-x64\publish\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{group}\{cm:UninstallProgram,{#MyAppName}}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(MyAppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent

[Code]
procedure CurStepChanged(CurStep: TSetupStep);
var
  ResultCode: Integer;
begin
  if CurStep = ssPostInstall then
  begin
    // Create scheduled task
    Exec('schtasks', '/create /tn "WeeklyReportApp" /tr "' + ExpandConstant('{app}\{#MyAppExeName}') + '" /sc weekly /d FRI /st 18:00 /ru ' + GetUserNameString() + ' /f', '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
  end;
end; 