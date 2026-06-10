# ZoemZone

A Python project for spatial data analysis using ArcPy and ArcGIS Pro.

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

Launch the GUI tool to analyse shapefiles and export summaries to Excel:

```powershell
zoemzone-shp-gui
```

### What it does

1. **Select input subfolder** — browse to a folder inside `E:\2026\ZoemZoneLimburg` (default).
2. **Scan shapefiles** — lists all `.shp` files in the selected folder automatically on open.
3. **Select a shapefile** — clicking a file in the list loads its available fields.
4. **Select a field** — choose any non-geometry field from the dropdown to summarise.
5. **Select output folder** — browse to the destination folder for the Excel file.
6. **Export** — writes an Excel file named `<shapefile>_<field>_summary.xlsx` containing:
   - All unique values of the selected field
   - Record count per value
   - Columns auto-fitted to their content width
7. **Output folder opens automatically** after a successful export.

### Persistent settings

The last used input folder, output folder, and selected field are saved to:

`%USERPROFILE%\.zoemzone_shp_gui_settings.json`

These are restored automatically on next launch.

### Data folder

Input data is expected under:

`E:\2026\ZoemZoneLimburg`
