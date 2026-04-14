# Install / Uninstall

## Install

Open PowerShell in this folder and run:

```powershell
powershell -ExecutionPolicy Bypass -File ".\Install-PaperCutAutoLogin.ps1"
```

Or double-click `Install.bat`.

This will:

- stop immediately without changing anything if PaperCut MF Client is not installed
- prompt to install AutoHotkey v2 with `winget` if AutoHotkey is missing
- prompt to save the PaperCut credential if needed
- change the existing PaperCut startup shortcut to use `--minimized --silent`
- add the AutoHotkey watcher to your user startup via `HKCU\...\Run`
- restart PaperCut and the watcher immediately

## Uninstall

Open PowerShell in this folder and run:

```powershell
powershell -ExecutionPolicy Bypass -File ".\Uninstall-PaperCutAutoLogin.ps1"
```

Or double-click `Uninstall.bat`.

This will:

- remove the AutoHotkey autorun entry
- stop the running AutoHotkey watcher
- restore the original PaperCut startup shortcut from backup when available
- remove the saved PaperCut credential
- leave this folder in place and tell you it is safe to delete manually
