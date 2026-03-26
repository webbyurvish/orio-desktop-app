@echo off
setlocal
REM Register orioai:// so that clicking "Desktop App" in the dashboard launches this app.
REM Run this from: (a) the folder that contains AiInterviewAssistant.exe, OR (b) the "Desktop App" project folder.

set "EXE_PATH="
if exist "%~dp0AiInterviewAssistant.exe" set "EXE_PATH=%~dp0AiInterviewAssistant.exe"
if not defined EXE_PATH if exist "%~dp0bin\Debug\net8.0-windows\AiInterviewAssistant.exe" set "EXE_PATH=%~dp0bin\Debug\net8.0-windows\AiInterviewAssistant.exe"
if not defined EXE_PATH if exist "%~dp0bin\x64\Debug\net8.0-windows\AiInterviewAssistant.exe" set "EXE_PATH=%~dp0bin\x64\Debug\net8.0-windows\AiInterviewAssistant.exe"
if not defined EXE_PATH if exist "%~dp0bin\Release\net8.0-windows\AiInterviewAssistant.exe" set "EXE_PATH=%~dp0bin\Release\net8.0-windows\AiInterviewAssistant.exe"

if not defined EXE_PATH (
  echo AiInterviewAssistant.exe not found.
  echo Run this script from the folder that contains the .exe, or from the "Desktop App" project folder.
  exit /b 1
)

echo Registering orioai:// with: %EXE_PATH%
reg add "HKCU\Software\Classes\orioai" /ve /d "URL:Orio AI Session" /f
reg add "HKCU\Software\Classes\orioai" /v "URL Protocol" /d "" /f
reg add "HKCU\Software\Classes\orioai\shell\open\command" /ve /d "\"%EXE_PATH%\" \"%%1\"" /f
echo Done. Try "Desktop App" in the dashboard again.
endlocal
