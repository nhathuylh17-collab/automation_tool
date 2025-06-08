#define MyAppVersion "v2.2.9"
#define MyAppName "automation_tool"
#define MyAppPublisher "MEK Automation - Maersk, Co.op."
#define MyAppURL "https://github.com/HuyGiaMsk/automation_tool"
#define MyAppExeName "automation_tool.exe"
#define MyAppAssocName MyAppName + " File"
#define MyAppAssocExt ".exe"
#define MyAppAssocKey StringChange(MyAppAssocName, " ", "") + MyAppAssocExt

[Setup]
AppId={{47E17F47-6018-4CDB-9778-0364E4561CBE}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
AppUpdatesURL={#MyAppURL}
UsePreviousAppDir=yes
DefaultDirName={sd}\{#MyAppName}
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
ChangesAssociations=yes
DisableProgramGroupPage=yes
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=commandline
OutputBaseFilename=automation_tool_installer
OutputDir=dist\
Compression=lzma
SolidCompression=yes
WizardStyle=modern

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
Source: "dist\{#MyAppExeName}"; DestDir: "{app}"; Flags: ignoreversion;
Source: "dist\input\*"; DestDir: "{app}\input"; Flags: ignoreversion;
Source: "dist\release_notes\*"; DestDir: "{app}\release_notes"; Flags: ignoreversion;

[Registry]
Root: HKA; Subkey: "Software\Classes\{#MyAppAssocExt}\OpenWithProgids"; ValueType: string; ValueName: "{#MyAppAssocKey}"; ValueData: ""; Flags: uninsdeletevalue
Root: HKA; Subkey: "Software\Classes\{#MyAppAssocKey}"; ValueType: string; ValueName: ""; ValueData: "{#MyAppAssocName}"; Flags: uninsdeletekey
Root: HKA; Subkey: "Software\Classes\{#MyAppAssocKey}\DefaultIcon"; ValueType: string; ValueName: ""; ValueData: "{app}\{#MyAppExeName},0"
Root: HKA; Subkey: "Software\Classes\{#MyAppAssocKey}\shell\open\command"; ValueType: string; ValueName: ""; ValueData: """{app}\{#MyAppExeName}"" ""%1"""
Root: HKA; Subkey: "Software\Classes\Applications\{#MyAppExeName}\SupportedTypes"; ValueType: string; ValueName: ".myp"; ValueData: ""

[Icons]
Name: "{autoprograms}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(MyAppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent

[Code]
var
  SkipScreens: Boolean;

function ShouldSkipPage(PageID: Integer): Boolean;
begin
  if SkipScreens then
  begin
    case PageID of
      wpSelectTasks: Result := True;
      wpReady: Result := True;
      wpFinished: Result := True;
    else
      Result := False;
    end;
  end
  else
    Result := False;
end;

function InitializeSetup(): Boolean;
var
  I: Integer;
  Param: String;
begin
  SkipScreens := False;

  for I := 1 to ParamCount do
  begin
    Param := ParamStr(I);
    if CompareText(Param, '/NOSCREENS') = 0 then
    begin
      SkipScreens := True;
      Break;
    end;
  end;
  Result := True;
end;