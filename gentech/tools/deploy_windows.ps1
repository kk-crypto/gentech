param(
    [int]$Port = 5000,
    [string]$PythonVersion = "3.11"
)

$ErrorActionPreference = "Stop"

function Test-IsAdmin {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

if (-not (Test-IsAdmin)) {
    throw "Run this script in an elevated PowerShell window (Run as Administrator)."
}

$ToolsDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$ProjectRoot = Split-Path -Parent $ToolsDir
$VenvPath = Join-Path $ProjectRoot ".venv"
$VenvPython = Join-Path $VenvPath "Scripts\\python.exe"
$RequirementsFile = Join-Path $ProjectRoot "requirements.txt"
$ServiceScript = Join-Path $ToolsDir "gentec_service.py"
$ServiceConfigPath = Join-Path $ProjectRoot "data\\service_config.json"
$ServiceName = "GentecBillingService"

function Resolve-Python {
    if (Test-Path $VenvPython) {
        return $VenvPython
    }

    if (Get-Command py -ErrorAction SilentlyContinue) {
        try {
            $resolved = (& py -$PythonVersion -c "import sys; print(sys.executable)" 2>$null | Select-Object -First 1).Trim()
            if ($resolved) {
                return $resolved
            }
        } catch {}
    }

    if (Get-Command python -ErrorAction SilentlyContinue) {
        return (Get-Command python).Source
    }

    return $null
}

$PythonExe = Resolve-Python
if (-not $PythonExe) {
    if (-not (Get-Command winget -ErrorAction SilentlyContinue)) {
        throw "Python not found and winget is unavailable. Install Python 3.11+ manually and rerun."
    }

    Write-Host "Installing Python $PythonVersion via winget..."
    winget install -e --id Python.Python.3.11 --accept-package-agreements --accept-source-agreements

    $PythonExe = Resolve-Python
    if (-not $PythonExe) {
        throw "Python installation completed but executable was not found. Restart terminal and rerun."
    }
}

if (-not (Test-Path $VenvPython)) {
    Write-Host "Creating virtual environment..."
    & $PythonExe -m venv $VenvPath
}

Write-Host "Installing dependencies..."
& $VenvPython -m pip install --upgrade pip
& $VenvPython -m pip install -r $RequirementsFile
& $VenvPython -m pip install pywin32

Write-Host "Writing service configuration..."
New-Item -ItemType Directory -Path (Split-Path -Parent $ServiceConfigPath) -Force | Out-Null
@{ port = $Port } | ConvertTo-Json | Set-Content -Path $ServiceConfigPath -Encoding UTF8

if (Get-Service -Name $ServiceName -ErrorAction SilentlyContinue) {
    Write-Host "Updating existing service..."
    & $VenvPython $ServiceScript stop | Out-Null
    & $VenvPython $ServiceScript remove | Out-Null
}

Write-Host "Installing Windows service..."
& $VenvPython $ServiceScript --startup auto install

Write-Host "Starting service..."
& $VenvPython $ServiceScript start

Write-Host ""
Write-Host "Deployment complete."
Write-Host "Service Name : $ServiceName"
Write-Host "Port         : $Port"
Write-Host "App URL      : http://localhost:$Port"
Write-Host "Logs         : $ProjectRoot\\data\\service_logs"
