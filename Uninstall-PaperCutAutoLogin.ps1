[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'

function Test-IsAdministrator {
    $currentIdentity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($currentIdentity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Wait-BeforeExit {
    Read-Host 'Press Enter to close this window' | Out-Null
}

if (-not (Test-IsAdministrator)) {
    Start-Process -FilePath 'powershell.exe' -Verb RunAs -ArgumentList @(
        '-ExecutionPolicy', 'Bypass',
        '-File', ('"{0}"' -f $MyInvocation.MyCommand.Path)
    ) -Wait
    exit 0
}

$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$backupPath = Join-Path $root 'install-backup.json'
$credentialPath = Join-Path $env:LOCALAPPDATA 'PaperCutAHK\papercut-credential.xml'
$credentialDir = Split-Path -Parent $credentialPath
$runPath = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Run'
$defaultRunValueName = 'PaperCutAutoLoginAHK'

if (Test-Path -LiteralPath $backupPath) {
    $backup = Get-Content -Path $backupPath -Raw | ConvertFrom-Json

    if (Test-Path -LiteralPath $backup.StartupShortcut) {
        $shell = New-Object -ComObject WScript.Shell
        $shortcut = $shell.CreateShortcut($backup.StartupShortcut)
        $shortcut.TargetPath = $backup.TargetPath
        $shortcut.Arguments = $backup.Arguments
        $shortcut.WorkingDirectory = $backup.WorkingDirectory
        $shortcut.IconLocation = $backup.IconLocation
        $shortcut.Save()
    }

    $runValueName = if ($backup.RunValueName) { $backup.RunValueName } else { $defaultRunValueName }
}
else {
    $runValueName = $defaultRunValueName
}

Remove-ItemProperty -Path $runPath -Name $runValueName -ErrorAction SilentlyContinue
Get-Process -Name 'AutoHotkey64','AutoHotkey32' -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue

if (Test-Path -LiteralPath $credentialPath) {
    Remove-Item -LiteralPath $credentialPath -Force -ErrorAction SilentlyContinue
}

if (Test-Path -LiteralPath $credentialDir) {
    Remove-Item -LiteralPath $credentialDir -Force -Recurse -ErrorAction SilentlyContinue
}

if (Test-Path -LiteralPath $backupPath) {
    Remove-Item -LiteralPath $backupPath -Force -ErrorAction SilentlyContinue
}

Write-Host 'PaperCut auto login uninstalled.'
Write-Host 'Startup shortcut restored when backup was available.'
Write-Host 'Saved credential removed.'
Write-Host "It is now safe to delete $root if you no longer want the scripts."
Write-Host 'Uninstall complete.'
Wait-BeforeExit
