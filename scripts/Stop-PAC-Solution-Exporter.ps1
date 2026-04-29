[CmdletBinding()]
param(
    [int]$Port = 4141
)

$ErrorActionPreference = 'Stop'
$connections = Get-NetTCPConnection -LocalPort $Port -State Listen -ErrorAction SilentlyContinue

if (-not $connections) {
    Write-Host "PAC Solution Exporter is not listening on port $Port."
    exit 0
}

$processIds = $connections | Select-Object -ExpandProperty OwningProcess -Unique
foreach ($processId in $processIds) {
    $process = Get-Process -Id $processId -ErrorAction SilentlyContinue
    if (-not $process) {
        continue
    }

    if ($process.ProcessName -notin @('node', 'nodejs')) {
        Write-Host "Port $Port is used by $($process.ProcessName) (PID $processId). Not stopping it automatically." -ForegroundColor Yellow
        continue
    }

    Stop-Process -Id $processId -Force
    Write-Host "Stopped PAC Solution Exporter server on port $Port (PID $processId)."
}
