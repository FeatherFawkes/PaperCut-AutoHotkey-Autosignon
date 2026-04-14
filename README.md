# PaperCut AutoHotkey Setup

## Files

- `Save-PaperCutCredential.ps1` saves your PaperCut credential encrypted for your Windows account.
- `Get-PaperCutCredential.ps1` reads the encrypted credential for the AutoHotkey script.
- `PaperCutAutoLogin.ahk` watches for the PaperCut `Login` popup and fills it automatically. It also closes `Balance for ...` windows.

## First-time setup

Run:

```powershell
powershell -ExecutionPolicy Bypass -File "C:\Users\Edison\PaperCut-AHK\Save-PaperCutCredential.ps1"
```

Then start:

```powershell
"C:\Program Files\AutoHotkey\v2\AutoHotkey64.exe" "C:\Users\Edison\PaperCut-AHK\PaperCutAutoLogin.ahk"
```

## Notes

- The script targets the Java dialog title `Login`, class `SunAwtDialog`, and process `pc-client.exe`.
- The click positions are window-relative and based on the PaperCut dialog layout you showed.
- Press `Ctrl+Alt+C` while hovering over the PaperCut popup to show the current mouse position relative to the window. This helps if we need to tune the click coordinates.
