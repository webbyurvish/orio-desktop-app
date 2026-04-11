; Code signing (SmartScreen / "Windows protected your PC"):
; Unsigned installers almost always show Microsoft Defender SmartScreen until you sign with a
; trusted Authenticode certificate from a public CA (e.g. DigiCert, Sectigo, SSL.com).
; 1) Publish: dotnet publish -c Release -r win-x64 --self-contained false (or your pipeline output).
; 2) Sign AiInterviewAssistant.exe with signtool BEFORE compiling this script (see DEPLOYMENT_RUNBOOK).
; 3) Install Windows SDK for signtool.exe; obtain a .pfx (Standard or EV code signing cert).
; 4) Inno Setup: Tools → Configure Sign Tools → add a tool named e.g. "orio" with a command like:
;      "C:\Program Files (x86)\Windows Kits\10\bin\10.0.22621.0\x64\signtool.exe" sign /fd sha256 /tr http://timestamp.digicert.com /td sha256 /f "C:\path\cert.pfx" /p "PFX_PASSWORD" $f
; 5) Uncomment the next two lines after configuring the sign tool name to match:
;SignTool=orio
;SignedUninstaller=yes

[Setup]
AppId={{9D2D3D5F-5B39-4B75-9C15-4B09B54F8E91}
AppName=Smeed AI Desktop
AppVersion={#AppVersion}
AppPublisher=Smeed AI
DefaultDirName={autopf}\\Smeed AI Desktop
DefaultGroupName=Smeed AI Desktop
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
Name: "{group}\\Smeed AI Desktop"; Filename: "{app}\\AiInterviewAssistant.exe"
Name: "{commondesktop}\\Smeed AI Desktop"; Filename: "{app}\\AiInterviewAssistant.exe"; Tasks: desktopicon

[Tasks]
Name: "desktopicon"; Description: "Create a &desktop shortcut"; GroupDescription: "Additional icons:"; Flags: unchecked

[Registry]
; Register orioai:// protocol for this user
Root: HKCU; Subkey: "Software\Classes\orioai"; ValueType: string; ValueName: ""; ValueData: "URL:Smeed AI Session"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Classes\orioai"; ValueType: string; ValueName: "URL Protocol"; ValueData: ""
Root: HKCU; Subkey: "Software\Classes\orioai\shell\open\command"; ValueType: string; ValueName: ""; ValueData: """{app}\AiInterviewAssistant.exe"" ""%1"""; Flags: uninsdeletekey

[Run]
Filename: "{app}\\AiInterviewAssistant.exe"; Description: "Launch Smeed AI Desktop"; Flags: nowait postinstall skipifsilent

