Param(
    [string]$InterpreterPath = "C:\Users\krame015\AppData\Local\ESRI\conda\envs\arcgispro-py3-clone\python.exe",
    [switch]$RunChecks
)

$ErrorActionPreference = "Stop"

$projectRoot = Resolve-Path (Join-Path $PSScriptRoot "..")
Set-Location $projectRoot

if (-not (Test-Path $InterpreterPath)) {
    throw "ArcPy interpreter not found at $InterpreterPath"
}

Write-Host "[bootstrap-arcpy] Using interpreter: $InterpreterPath"

Write-Host "[bootstrap-arcpy] Verifying ArcPy import..."
& $InterpreterPath -c "import arcpy; print(arcpy.__file__)"

Write-Host "[bootstrap-arcpy] Installing project and dev dependencies..."
& $InterpreterPath -m pip install -r requirements-dev.txt

if ($RunChecks) {
    Write-Host "[bootstrap-arcpy] Running smoke checks..."
    & $InterpreterPath -m zoemzone.main
    & $InterpreterPath -m pytest -q
}

Write-Host "[bootstrap-arcpy] Done."
