; Copilot Clown - Excel AI Add-in Installer
; Inno Setup Script
; Download Inno Setup from: https://jrsoftware.org/isinfo.php

#define MyAppName "Copilot Clown"
#define MyAppVersion "2.0.0"
#define MyAppPublisher "Copilot Clown"
#define MyAppURL "https://github.com/Benidrissa/copilot_clown"

[Setup]
AppId={{B3A7C2D1-F8E6-4A90-BC5D-1234567890AB}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
DefaultDirName={localappdata}\CopilotClown
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
OutputDir=Output
OutputBaseFilename=CopilotClownSetup
Compression=lzma2
SolidCompression=yes
PrivilegesRequired=lowest
CloseApplications=force
CloseApplicationsFilter=*.exe
RestartApplications=no
UninstallDisplayName={#MyAppName}

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Files]
; The packed .xll add-in (64-bit)
Source: "..\CopilotClown\bin\Release\net48\publish\CopilotClown64-packed.xll"; DestDir: "{app}"; DestName: "CopilotClown64-packed.xll"; Flags: ignoreversion
; Also include 32-bit for older Excel installations
Source: "..\CopilotClown\bin\Release\net48\publish\CopilotClown-packed.xll"; DestDir: "{app}"; DestName: "CopilotClown-packed.xll"; Flags: ignoreversion

[Registry]
; Add as trusted location for Excel (Office 16.0 = Office 2016/2019/365)
Root: HKCU; Subkey: "Software\Microsoft\Office\16.0\Excel\Security\Trusted Locations\LocationCopilotClown"; ValueType: string; ValueName: "Path"; ValueData: "{app}\"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Microsoft\Office\16.0\Excel\Security\Trusted Locations\LocationCopilotClown"; ValueType: dword; ValueName: "AllowSubFolders"; ValueData: "1"

; Also add for Office 15.0 (Office 2013) just in case
Root: HKCU; Subkey: "Software\Microsoft\Office\15.0\Excel\Security\Trusted Locations\LocationCopilotClown"; ValueType: string; ValueName: "Path"; ValueData: "{app}\"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Microsoft\Office\15.0\Excel\Security\Trusted Locations\LocationCopilotClown"; ValueType: dword; ValueName: "AllowSubFolders"; ValueData: "1"

[Icons]
Name: "{group}\Uninstall {#MyAppName}"; Filename: "{uninstallexe}"

[Code]
// Find the next available OPEN key slot in Excel Options
function FindNextOpenSlot(): String;
var
  I: Integer;
  KeyName: String;
begin
  // Excel uses OPEN, OPEN1, OPEN2, etc. for auto-loaded add-ins
  // First check if OPEN exists
  if not RegValueExists(HKCU, 'Software\Microsoft\Office\16.0\Excel\Options', 'OPEN') then
  begin
    Result := 'OPEN';
    Exit;
  end;
  // Check OPEN1 through OPEN20
  for I := 1 to 20 do
  begin
    KeyName := 'OPEN' + IntToStr(I);
    if not RegValueExists(HKCU, 'Software\Microsoft\Office\16.0\Excel\Options', KeyName) then
    begin
      Result := KeyName;
      Exit;
    end;
  end;
  // Fallback
  Result := 'OPEN99';
end;

// Check if the add-in is already registered
function IsAddinRegistered(): Boolean;
var
  I: Integer;
  KeyName: String;
  Value: String;
begin
  Result := False;
  // Check OPEN, OPEN1, OPEN2, ...
  if RegQueryStringValue(HKCU, 'Software\Microsoft\Office\16.0\Excel\Options', 'OPEN', Value) then
    if Pos('CopilotClown', Value) > 0 then begin Result := True; Exit; end;
  for I := 1 to 20 do
  begin
    KeyName := 'OPEN' + IntToStr(I);
    if RegQueryStringValue(HKCU, 'Software\Microsoft\Office\16.0\Excel\Options', KeyName, Value) then
      if Pos('CopilotClown', Value) > 0 then begin Result := True; Exit; end;
  end;
end;

procedure CurStepChanged(CurStep: TSetupStep);
var
  SlotName: String;
  XllPath: String;
begin
  if CurStep = ssPostInstall then
  begin
    // Register the add-in to auto-load in Excel
    if not IsAddinRegistered() then
    begin
      SlotName := FindNextOpenSlot();
      XllPath := ExpandConstant('{app}\CopilotClown64-packed.xll');
      RegWriteStringValue(HKCU, 'Software\Microsoft\Office\16.0\Excel\Options', SlotName, XllPath);
    end;
  end;
end;

procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
var
  I: Integer;
  KeyName: String;
  Value: String;
begin
  if CurUninstallStep = usPostUninstall then
  begin
    // Remove the add-in registration from Excel Options
    if RegQueryStringValue(HKCU, 'Software\Microsoft\Office\16.0\Excel\Options', 'OPEN', Value) then
      if Pos('CopilotClown', Value) > 0 then
        RegDeleteValue(HKCU, 'Software\Microsoft\Office\16.0\Excel\Options', 'OPEN');
    for I := 1 to 20 do
    begin
      KeyName := 'OPEN' + IntToStr(I);
      if RegQueryStringValue(HKCU, 'Software\Microsoft\Office\16.0\Excel\Options', KeyName, Value) then
        if Pos('CopilotClown', Value) > 0 then
          RegDeleteValue(HKCU, 'Software\Microsoft\Office\16.0\Excel\Options', KeyName);
    end;
    // Clean up settings
    DelTree(ExpandConstant('{localappdata}\CopilotClown'), True, True, True);
  end;
end;

[Messages]
FinishedLabel=Setup has finished installing [name] on your computer.%n%nThe =USEAI() function will be available next time you open Excel.%n%nUse Alt+F8 > ShowAISettings to configure your API keys.
