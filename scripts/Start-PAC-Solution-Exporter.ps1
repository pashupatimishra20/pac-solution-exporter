[CmdletBinding()]
param(
    [int]$Port = 4141
)

$ErrorActionPreference = 'Stop'
$AppRoot = Split-Path -Parent $PSScriptRoot

function Test-Command {
    param([Parameter(Mandatory)][string]$Name)
    return [bool](Get-Command $Name -ErrorAction SilentlyContinue)
}

if (-not (Test-Command node)) {
    throw "Node.js was not found. Run .\scripts\Install-Prerequisites.ps1 -InstallMissing first."
}

if (-not (Test-Command npm)) {
    throw "npm was not found. Reinstall Node.js LTS, then reopen PowerShell."
}

$url = "http://127.0.0.1:$Port"

try {
    $health = Invoke-RestMethod -Uri "$url/api/health" -TimeoutSec 2
    if ($health.ok) {
        Start-Process $url
        Write-Host "PAC Solution Exporter is already running at $url"
        exit 0
    }
} catch {
    # Server is not running yet.
}

$command = "Set-Location -LiteralPath '$($AppRoot.Replace("'", "''"))'; `$env:PORT='$Port'; npm start"

Start-Process powershell.exe -ArgumentList @(
    '-NoExit',
    '-ExecutionPolicy', 'Bypass',
    '-Command', $command
) -WorkingDirectory $AppRoot

Start-Sleep -Seconds 2
Start-Process $url
Write-Host "PAC Solution Exporter started at $url"
