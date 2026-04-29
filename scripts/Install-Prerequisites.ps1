[CmdletBinding()]
param(
    [switch]$InstallMissing
)

$ErrorActionPreference = 'Stop'

function Test-Command {
    param([Parameter(Mandatory)][string]$Name)
    return [bool](Get-Command $Name -ErrorAction SilentlyContinue)
}

function Write-Step {
    param([Parameter(Mandatory)][string]$Message)
    Write-Host ""
    Write-Host "== $Message" -ForegroundColor Cyan
}

Write-Step "Checking required tools"
$hasNode = Test-Command node
$hasNpm = Test-Command npm
$hasPac = Test-Command pac
$hasWinget = Test-Command winget

Write-Host ("Node.js: " + ($(if ($hasNode) { "found ($(& node -v))" } else { "missing" })))
Write-Host ("npm:     " + ($(if ($hasNpm) { "found ($(& npm -v))" } else { "missing" })))
Write-Host ("PAC CLI: " + ($(if ($hasPac) { "found" } else { "missing" })))

if (-not $InstallMissing) {
    Write-Host ""
    Write-Host "Run with -InstallMissing to install missing components." -ForegroundColor Yellow
    Write-Host "Example: .\scripts\Install-Prerequisites.ps1 -InstallMissing"
    exit 0
}

if ((-not $hasNode -or -not $hasNpm) -and -not $hasWinget) {
    throw "Node.js is missing and winget was not found. Install Node.js LTS from https://nodejs.org/en/download, then reopen PowerShell."
}

if (-not $hasNode -or -not $hasNpm) {
    Write-Step "Installing Node.js LTS with winget"
    winget install -e --id OpenJS.NodeJS.LTS
    Write-Host "Close and reopen PowerShell if node/npm are still not found after install." -ForegroundColor Yellow
}

if (-not $hasPac) {
    Write-Step "Installing Power Platform CLI Windows MSI"
    $msi = Join-Path $env:TEMP 'powerapps-cli.msi'
    Invoke-WebRequest 'https://aka.ms/PowerAppsCLI' -OutFile $msi
    Start-Process msiexec.exe -Wait -ArgumentList "/i `"$msi`" /passive"
    Write-Host "Close and reopen PowerShell if pac is still not found after install." -ForegroundColor Yellow
}

Write-Step "Done"
Write-Host "Verify with: node -v; npm -v; pac"
