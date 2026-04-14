# PaperCut AutoHotkey Setup

## Files

- `Save-PaperCutCredential.ps1` saves your PaperCut credential encrypted for your Windows account.
- `Get-PaperCutCredential.ps1` reads the encrypted credential for the AutoHotkey script.
- `PaperCutAutoLogin.ahk` watches for the PaperCut `Login` popup and fills it automatically. It also closes `Balance for ...` windows.

## First-time setup

For a fresh install, run:

```powershell
powershell -ExecutionPolicy Bypass -File ".\Install-PaperCutAutoLogin.ps1"
```

Or just double-click:

- `Install.bat`

If you only want to update the saved PaperCut password, run:

```powershell
powershell -ExecutionPolicy Bypass -File ".\Save-PaperCutCredential.ps1"
```

Or double-click:

- `Set Password.bat`

## Notes

- The script targets the Java dialog title `Login`, class `SunAwtDialog`, and process `pc-client.exe`.
- Press `Ctrl+Alt+C` while hovering over the PaperCut popup to show the current mouse position relative to the window. This helps if we need to tune the click coordinates.
