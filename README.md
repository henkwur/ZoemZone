# ZoemZone

A minimal Python starter project scaffolded in VS Code.

## Quickstart

Run all commands below in the same PowerShell terminal.

## ArcPy setup (ArcGIS Pro clone)

This project is configured to use:

`C:\Users\krame015\AppData\Local\ESRI\conda\envs\arcgispro-py3-clone\python.exe`

One-command setup with ArcPy:

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
.\scripts\bootstrap-arcpy.ps1 -RunChecks
```

Manual ArcPy commands:

```powershell
$ArcPyPython = "C:\Users\krame015\AppData\Local\ESRI\conda\envs\arcgispro-py3-clone\python.exe"
& $ArcPyPython -m pip install -r requirements-dev.txt
& $ArcPyPython -m pytest -q
```

### One-command bootstrap (recommended)

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
.\scripts\bootstrap.ps1 -RunChecks
```

1. Create and activate a virtual environment:

   ```powershell
   python -m venv .venv
   Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
   .\.venv\Scripts\Activate.ps1
   ```

2. Install in editable mode with dev dependencies:

   ```powershell
   python -m pip install -r requirements-dev.txt
   ```

3. Run the app:

   ```powershell
   zoemzone
   ```

4. Run tests:

   ```powershell
   python -m pytest
   ```

## Shapefile GUI export

Launch the GUI to select a subfolder, list `.shp` files, and export selected files to Excel:

```powershell
zoemzone-shp-gui
```
