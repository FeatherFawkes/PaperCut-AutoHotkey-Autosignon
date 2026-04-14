[CmdletBinding()]
param(
    [string]$Path = "$env:LOCALAPPDATA\PaperCutAHK\papercut-credential.xml"
)

$dir = Split-Path -Parent $Path
if (-not (Test-Path -LiteralPath $dir)) {
    New-Item -ItemType Directory -Path $dir | Out-Null
}

$credential = Get-Credential -Message "Enter your PaperCut username and password."
$credential | Export-Clixml -Path $Path

Write-Host "Saved encrypted credential to $Path"
Write-Host "Only this Windows account can decrypt it."
