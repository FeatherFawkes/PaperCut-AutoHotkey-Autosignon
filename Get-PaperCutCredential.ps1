[CmdletBinding()]
param(
    [string]$Path = "$env:LOCALAPPDATA\PaperCutAHK\papercut-credential.xml"
)

if (-not (Test-Path -LiteralPath $Path)) {
    throw "Credential file not found at $Path"
}

$credential = Import-Clixml -Path $Path
$username = $credential.UserName
$password = $credential.GetNetworkCredential().Password

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
Write-Output $username
Write-Output $password
