<#
    start_gui.ps1 - One-click launcher for the pdf2excel GUI.

    Creates the virtual environment (if missing), installs the required
    packages (only when requirements change), then starts the Tkinter GUI.

    Usage:
        Right-click -> "Run with PowerShell"
        or from a terminal:  .\start_gui.ps1
        Force a reinstall:   .\start_gui.ps1 -Reinstall
#>

param(
    [switch]$Reinstall
)

$ErrorActionPreference = 'Stop'

# Always work from the folder this script lives in.
Set-Location -Path $PSScriptRoot

$VenvDir    = Join-Path $PSScriptRoot 'venv'
$VenvPython = Join-Path $VenvDir 'Scripts\python.exe'
$ReqFile    = Join-Path $PSScriptRoot 'requirements.txt'
$Stamp      = Join-Path $VenvDir '.requirements.stamp'   # hash of last-installed requirements
$Gui        = Join-Path $PSScriptRoot 'pdf_to_excel_gui.py'

function Find-Python {
    foreach ($cmd in 'py', 'python', 'python3') {
        $exe = Get-Command $cmd -ErrorAction SilentlyContinue
        if ($exe) {
            if ($cmd -eq 'py') { return @('py', '-3') }
            return @($exe.Source)
        }
    }
    return $null
}

# --- 1. Ensure the virtual environment exists -------------------------------
if (-not (Test-Path $VenvPython)) {
    Write-Host '[setup] Creating virtual environment...' -ForegroundColor Cyan
    $py = Find-Python
    if (-not $py) {
        Write-Error 'Python was not found on PATH. Install Python 3.10+ from https://python.org and re-run.'
        exit 1
    }
    & $py[0] $py[1..($py.Count - 1)] -m venv $VenvDir
    if (-not (Test-Path $VenvPython)) {
        Write-Error 'Failed to create the virtual environment.'
        exit 1
    }
}

# --- 2. Install / update packages only when requirements have changed -------
$reqHash = (Get-FileHash -Path $ReqFile -Algorithm SHA256).Hash
$installed = (Test-Path $Stamp) -and ((Get-Content $Stamp -Raw).Trim() -eq $reqHash)

if ($Reinstall -or -not $installed) {
    Write-Host '[setup] Installing dependencies (this may take a minute)...' -ForegroundColor Cyan
    & $VenvPython -m pip install --upgrade pip --quiet
    & $VenvPython -m pip install -r $ReqFile
    if ($LASTEXITCODE -ne 0) {
        Write-Error 'pip failed to install the requirements.'
        exit 1
    }
    Set-Content -Path $Stamp -Value $reqHash -Encoding utf8
} else {
    Write-Host '[setup] Dependencies already up to date.' -ForegroundColor DarkGray
}

# --- 3. Launch the GUI ------------------------------------------------------
Write-Host '[run] Starting the pdf2excel GUI...' -ForegroundColor Green
& $VenvPython $Gui
