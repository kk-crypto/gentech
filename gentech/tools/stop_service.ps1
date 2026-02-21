param(
    [string]$ServiceName = "GentecBillingService"
)

$ErrorActionPreference = "Stop"

$service = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
if (-not $service) {
    throw "Service '$ServiceName' is not installed."
}

if ($service.Status -ne "Stopped") {
    Stop-Service -Name $ServiceName -Force
}

Get-Service -Name $ServiceName | Select-Object Name, Status
