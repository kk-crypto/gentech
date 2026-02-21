param(
    [string]$ServiceName = "GentecBillingService"
)

$ErrorActionPreference = "Stop"

$service = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
if (-not $service) {
    throw "Service '$ServiceName' is not installed. Run tools\\deploy_windows.ps1 first."
}

if ($service.Status -ne "Running") {
    Start-Service -Name $ServiceName
}

Get-Service -Name $ServiceName | Select-Object Name, Status
