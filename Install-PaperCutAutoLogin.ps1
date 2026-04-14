[CmdletBinding()]
param(
    [switch]$StartNow = $true
)

$ErrorActionPreference = 'Stop'

$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$startupShortcut = 'C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Startup\PaperCut MF Client.lnk'
$backupPath = Join-Path $root 'install-backup.json'
$ahkScript = Join-Path $root 'PaperCutAutoLogin.ahk'
$saveCredentialScript = Join-Path $root 'Save-PaperCutCredential.ps1'
$credentialPath = Join-Path $env:LOCALAPPDATA 'PaperCutAHK\papercut-credential.xml'
$runValueName = 'PaperCutAutoLoginAHK'
$paperCutExe = 'C:\Program Files\PaperCut MF Client\pc-client.exe'
$paperCutArgs = '--minimized --silent'

function Get-AutoHotkeyExe {
    $paths = @(
        'C:\Users\Edison\AppData\Local\Programs\AutoHotkey\v2\AutoHotkey64.exe',
        'C:\Users\Edison\AppData\Local\Programs\AutoHotkey\v2\AutoHotkey32.exe',
        'C:\Program Files\AutoHotkey\v2\AutoHotkey64.exe',
        'C:\Program Files\AutoHotkey\v2\AutoHotkey32.exe',
        'C:\Program Files (x86)\AutoHotkey\v2\AutoHotkey64.exe',
        'C:\Program Files (x86)\AutoHotkey\v2\AutoHotkey32.exe'
    )

    foreach ($path in $paths) {
        if (Test-Path -LiteralPath $path) {
            return $path
        }
    }

    return $null
}

if (-not (Test-Path -LiteralPath $paperCutExe)) {
    Write-Warning "PaperCut MF Client was not found at $paperCutExe"
    Write-Host 'Nothing was installed.'
    exit 0
}

if (-not (Test-Path -LiteralPath $ahkScript)) {
    throw "AutoHotkey script not found at $ahkScript"
}

if (-not (Test-Path -LiteralPath $startupShortcut)) {
    throw "PaperCut startup shortcut not found at $startupShortcut"
}

$ahkExe = Get-AutoHotkeyExe
if (-not $ahkExe) {
    $choice = Read-Host 'AutoHotkey v2 is not installed. Install it now with winget? (Y/N)'
    if ($choice -notmatch '^(?i)y(es)?$') {
        Write-Warning 'AutoHotkey is required. Installation cancelled.'
        exit 1
    }

    $winget = Get-Command winget -ErrorAction SilentlyContinue
    if (-not $winget) {
        throw 'winget was not found. Install AutoHotkey v2 manually and rerun this script.'
    }

    & $winget.Source install --id AutoHotkey.AutoHotkey --exact --accept-package-agreements --accept-source-agreements
    $ahkExe = Get-AutoHotkeyExe
    if (-not $ahkExe) {
        throw 'AutoHotkey installation did not complete successfully.'
    }
}

if (-not (Test-Path -LiteralPath $credentialPath)) {
    & powershell -ExecutionPolicy Bypass -File $saveCredentialScript
}

$shell = New-Object -ComObject WScript.Shell
$shortcut = $shell.CreateShortcut($startupShortcut)

$backup = [pscustomobject]@{
    StartupShortcut = $startupShortcut
    TargetPath = $shortcut.TargetPath
    Arguments = $shortcut.Arguments
    WorkingDirectory = $shortcut.WorkingDirectory
    IconLocation = $shortcut.IconLocation
    RunValueName = $runValueName
}
$backup | ConvertTo-Json | Set-Content -Path $backupPath -Encoding UTF8

$shortcut.TargetPath = $paperCutExe
$shortcut.Arguments = $paperCutArgs
$shortcut.WorkingDirectory = Split-Path -Parent $paperCutExe
$shortcut.IconLocation = "$paperCutExe,0"
$shortcut.Save()

$runCommand = ('"{0}" "{1}"' -f $ahkExe, $ahkScript)
New-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Run' -Name $runValueName -Value $runCommand -PropertyType String -Force | Out-Null

if ($StartNow) {
    Get-Process -Name 'AutoHotkey64','AutoHotkey32','pc-client' -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
    Start-Sleep -Milliseconds 300
    Start-Process -FilePath $ahkExe -ArgumentList ('"{0}"' -f $ahkScript) | Out-Null
    Start-Process -FilePath $paperCutExe -ArgumentList $paperCutArgs -WorkingDirectory (Split-Path -Parent $paperCutExe) | Out-Null
}

Write-Host 'PaperCut auto login installed.'
Write-Host 'Startup shortcut updated to run PaperCut minimized and silent.'
Write-Host 'AutoHotkey watcher added to HKCU Run.'
