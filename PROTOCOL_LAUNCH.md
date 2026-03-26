# Desktop App – Launch from dashboard (orioai://)

The dashboard can open this app with a specific **call session** and **API token** so conversation is saved to that session.

## How it works

1. In the dashboard, when you click **Desktop App** in the "Choose Platform" modal (after creating a session or when opening a session from the list), the browser navigates to a URL like:
   ```
   orioai://start?sessionId=<GUID>&apiBaseUrl=<...>&token=<JWT>
   ```
2. If the `orioai://` protocol is registered on your PC, Windows starts this app and passes that URL as a command-line argument.
3. The app parses the URL and uses that **session id**, **API base URL**, and **token** for the rest of the run. All transcript and AI answers are then sent to that session.

## Register the protocol (one-time)

So that Windows knows to run this app when you open a `parakeetai://` link:

### Option A – Run the batch file (easiest)

1. Build the app so `AiInterviewAssistant.exe` exists (e.g. in `bin\Debug\net8.0-windows\` or next to the script).
2. Copy `RegisterProtocol.bat` into the **same folder** as `AiInterviewAssistant.exe`.
3. Right‑click `RegisterProtocol.bat` → **Run as administrator** (or run from a normal Command Prompt; `HKCU` does not require admin).

If the exe is elsewhere, edit the script and set `EXE_PATH` to the full path of `AiInterviewAssistant.exe`.

### Option B – Registry by hand

1. Win+R → `regedit` → Enter.
2. Under `HKEY_CURRENT_USER\Software\Classes`, create:
   - Key: `orioai`
     - Default value: `URL:Orio AI Session`
     - String value: `URL Protocol` = (empty)
   - Key: `orioai\shell\open\command`
     - Default value: `"C:\Path\To\AiInterviewAssistant.exe" "%1"`
       (Replace with the real path to your exe; `%1` is the full `orioai://...` URL.)

## After install (for end users)

When you distribute the app (e.g. installer or zip):

- Run the same registration once per user (e.g. from your installer or a “Register protocol” shortcut that runs `RegisterProtocol.bat` with the installed exe path).
- Then “Desktop App” in the dashboard will start the app with the correct session and token for that user.

## Local dev (no protocol)

- You can still run the app from Visual Studio or the command line; it will use `appsettings.json` (e.g. `CallSessionId`, `ApiBaseUrl`, `ApiBearerToken`).
- To test the protocol, run `RegisterProtocol.bat` so the exe path points to your build output, then click “Desktop App” in the dashboard.
