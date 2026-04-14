# Install / Uninstall

## Install

Run PowerShell as Administrator:

```powershell
powershell -ExecutionPolicy Bypass -File "C:\Users\Edison\PaperCut-AHK\Install-PaperCutAutoLogin.ps1"
```

This will:

- stop immediately without changing anything if PaperCut MF Client is not installed
- prompt to install AutoHotkey v2 with `winget` if AutoHotkey is missing
- prompt to save the PaperCut credential if needed
- change the existing PaperCut startup shortcut to use `--minimized --silent`
- add the AutoHotkey watcher to your user startup via `HKCU\...\Run`
- restart PaperCut and the watcher immediately

## Uninstall

Run PowerShell as Administrator:

```powershell
powershell -ExecutionPolicy Bypass -File "C:\Users\Edison\PaperCut-AHK\Uninstall-PaperCutAutoLogin.ps1"
```

This will:

- remove the AutoHotkey autorun entry
- stop the running AutoHotkey watcher
- restore the original PaperCut startup shortcut from backup when available
- remove the saved PaperCut credential
- leave the `C:\Users\Edison\PaperCut-AHK` folder in place and tell you it is safe to delete manually
