Param(
    [switch]$RunChecks
)

$ErrorActionPreference = "Stop"

$projectRoot = Resolve-Path (Join-Path $PSScriptRoot "..")
Set-Location $projectRoot

$venvPath = Join-Path $projectRoot ".venv"
$activateScript = Join-Path $venvPath "Scripts\\Activate.ps1"

if (-not (Test-Path $venvPath)) {
    Write-Host "[bootstrap] Creating virtual environment..."
    python -m venv .venv
}

if (-not (Test-Path $activateScript)) {
    throw "Activation script not found at $activateScript"
}

Write-Host "[bootstrap] Activating virtual environment..."
. $activateScript

Write-Host "[bootstrap] Installing project and dev dependencies..."
python -m pip install -r requirements-dev.txt

if ($RunChecks) {
    Write-Host "[bootstrap] Running smoke checks..."
    zoemzone
    python -m pytest -q
}

Write-Host "[bootstrap] Done."
Write-Host "Run '.\\.venv\\Scripts\\Activate.ps1' in new terminals before working."
