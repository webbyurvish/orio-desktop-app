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

[Run]
Filename: "{app}\\AiInterviewAssistant.exe"; Description: "Launch Orio AI Desktop"; Flags: nowait postinstall skipifsilent

