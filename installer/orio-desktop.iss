[Setup]
AppId={{9D2D3D5F-5B39-4B75-9C15-4B09B54F8E91}
AppName=Orio AI Desktop
AppVersion={#AppVersion}
AppPublisher=Orio AI
DefaultDirName={autopf}\\Orio AI Desktop
DefaultGroupName=Orio AI Desktop
DisableProgramGroupPage=yes
OutputDir=output
OutputBaseFilename=orio-desktop-setup
Compression=lzma
SolidCompression=yes
WizardStyle=modern

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Files]
Source: "..\\publish\\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\\Orio AI Desktop"; Filename: "{app}\\AiInterviewAssistant.exe"
Name: "{commondesktop}\\Orio AI Desktop"; Filename: "{app}\\AiInterviewAssistant.exe"; Tasks: desktopicon

[Tasks]
Name: "desktopicon"; Description: "Create a &desktop shortcut"; GroupDescription: "Additional icons:"; Flags: unchecked

[Registry]
; Register orioai:// protocol for this user
Root: HKCU; Subkey: "Software\Classes\orioai"; ValueType: string; ValueName: ""; ValueData: "URL:Orio AI Session"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Classes\orioai"; ValueType: string; ValueName: "URL Protocol"; ValueData: ""
Root: HKCU; Subkey: "Software\Classes\orioai\shell\open\command"; ValueType: string; ValueName: ""; ValueData: """{app}\AiInterviewAssistant.exe"" ""%1"""; Flags: uninsdeletekey

[Run]
Filename: "{app}\\AiInterviewAssistant.exe"; Description: "Launch Orio AI Desktop"; Flags: nowait postinstall skipifsilent

