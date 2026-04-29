[CmdletBinding()]
param(
    [string]$ShortcutName = 'PAC Solution Exporter'
)

$ErrorActionPreference = 'Stop'
$AppRoot = Split-Path -Parent $PSScriptRoot
$Launcher = Join-Path $AppRoot 'Start-PAC-Solution-Exporter.cmd'

if (-not (Test-Path -LiteralPath $Launcher)) {
    throw "Launcher not found: $Launcher"
}

$desktop = [Environment]::GetFolderPath('Desktop')
$shortcutPath = Join-Path $desktop "$ShortcutName.lnk"

$shell = New-Object -ComObject WScript.Shell
$shortcut = $shell.CreateShortcut($shortcutPath)
$shortcut.TargetPath = $Launcher
$shortcut.WorkingDirectory = $AppRoot
$shortcut.Description = 'Start PAC Solution Exporter'
$shortcut.Save()

Write-Host "Created desktop shortcut: $shortcutPath"
