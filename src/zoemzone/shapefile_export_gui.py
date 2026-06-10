"""Tkinter GUI for listing shapefiles and exporting details to Excel."""

from __future__ import annotations

import json
import os
import re
from pathlib import Path

import arcpy
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk


DEFAULT_INPUT_FOLDER = Path(r"E:\2026\ZoemZoneLimburg")
SETTINGS_FILE = Path.home() / ".zoemzone_shp_gui_settings.json"


class ShapefileExportApp:
    """Desktop app for selecting shapefiles and exporting metadata."""

    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("ZoemZone Shapefile to Excel")
        self.root.geometry("900x560")

        self.input_folder_var = tk.StringVar(value=str(DEFAULT_INPUT_FOLDER))
        self.output_folder_var = tk.StringVar(value=str(DEFAULT_INPUT_FOLDER))
        self.field_var = tk.StringVar()
        self.status_var = tk.StringVar(value="Select a subfolder to start.")

        self.last_output_folder = DEFAULT_INPUT_FOLDER
        self._load_settings()
        self._build_ui()

        if Path(self.input_folder_var.get()).is_dir():
            self.scan_shapefiles()

        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    def _load_settings(self) -> None:
        if not SETTINGS_FILE.exists():
            return

        try:
            settings = json.loads(SETTINGS_FILE.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError):
            return

        input_folder = settings.get("input_folder")
        output_folder = settings.get("output_folder")
        selected_field = settings.get("selected_field")

        if isinstance(input_folder, str) and input_folder.strip():
            self.input_folder_var.set(input_folder)

        if isinstance(output_folder, str) and output_folder.strip():
            output_path = Path(output_folder)
            self.last_output_folder = output_path
            self.output_folder_var.set(str(output_path))

        if isinstance(selected_field, str) and selected_field.strip():
            self.field_var.set(selected_field)

    def _save_settings(self) -> None:
        settings = {
            "input_folder": self.input_folder_var.get().strip(),
            "output_folder": str(self.last_output_folder),
            "selected_field": self.field_var.get().strip(),
        }

        try:
            SETTINGS_FILE.write_text(
                json.dumps(settings, indent=2),
                encoding="utf-8",
            )
        except OSError:
            # Saving preferences is best-effort and should not block the app.
            pass

    def _on_close(self) -> None:
        output_folder = self.output_folder_var.get().strip()
        if output_folder:
            self.last_output_folder = Path(output_folder)
        self._save_settings()
        self.root.destroy()

    def _open_output_folder(self, folder: Path) -> None:
        try:
            os.startfile(str(folder))
        except OSError:
            self.status_var.set(f"Exported, but could not open folder: {folder}")

    def _build_ui(self) -> None:
        container = ttk.Frame(self.root, padding=12)
        container.pack(fill=tk.BOTH, expand=True)

        input_row = ttk.Frame(container)
        input_row.pack(fill=tk.X, pady=(0, 8))

        ttk.Label(input_row, text="Input subfolder:").pack(side=tk.LEFT)
        ttk.Entry(input_row, textvariable=self.input_folder_var).pack(
            side=tk.LEFT, fill=tk.X, expand=True, padx=8
        )
        ttk.Button(
            input_row, text="Browse...", command=self.select_input_folder
        ).pack(side=tk.LEFT)
        ttk.Button(input_row, text="Scan", command=self.scan_shapefiles).pack(
            side=tk.LEFT, padx=(8, 0)
        )

        list_row = ttk.Frame(container)
        list_row.pack(fill=tk.BOTH, expand=True, pady=(0, 8))

        ttk.Label(
            list_row,
            text="Shapefiles (.shp) in selected folder:",
        ).pack(anchor=tk.W)

        listbox_frame = ttk.Frame(list_row)
        listbox_frame.pack(fill=tk.BOTH, expand=True, pady=(4, 0))

        self.listbox = tk.Listbox(
            listbox_frame,
            selectmode=tk.BROWSE,
            exportselection=False,
            height=15,
        )
        self.listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        scrollbar = ttk.Scrollbar(
            listbox_frame, orient=tk.VERTICAL, command=self.listbox.yview
        )
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.listbox.configure(yscrollcommand=scrollbar.set)
        self.listbox.bind("<<ListboxSelect>>", self._on_shapefile_selected)

        field_row = ttk.Frame(container)
        field_row.pack(fill=tk.X, pady=(0, 8))

        ttk.Label(field_row, text="Field for summary:").pack(side=tk.LEFT)
        self.field_combo = ttk.Combobox(
            field_row,
            textvariable=self.field_var,
            state="readonly",
        )
        self.field_combo.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=8)
        self.field_combo.bind("<<ComboboxSelected>>", self._on_field_selected)

        output_row = ttk.Frame(container)
        output_row.pack(fill=tk.X, pady=(0, 8))

        ttk.Label(output_row, text="Excel output folder:").pack(side=tk.LEFT)
        ttk.Entry(output_row, textvariable=self.output_folder_var).pack(
            side=tk.LEFT, fill=tk.X, expand=True, padx=8
        )
        ttk.Button(
            output_row, text="Browse...", command=self.select_output_folder
        ).pack(side=tk.LEFT)

        action_row = ttk.Frame(container)
        action_row.pack(fill=tk.X, pady=(0, 8))
        ttk.Button(
            action_row,
            text="Export unique values summary to Excel",
            command=self.export_selected,
        ).pack(side=tk.LEFT)

        ttk.Label(
            container,
            textvariable=self.status_var,
            foreground="#1f5f2f",
        ).pack(anchor=tk.W)

    def select_input_folder(self) -> None:
        start_dir = self.input_folder_var.get().strip() or str(DEFAULT_INPUT_FOLDER)
        folder = filedialog.askdirectory(
            title="Select subfolder with shapefiles",
            initialdir=start_dir,
        )
        if folder:
            self.input_folder_var.set(folder)
            self._save_settings()
            self.scan_shapefiles()

    def _selected_shapefile_path(self) -> Path | None:
        selected = self.listbox.curselection()
        if not selected:
            return None

        input_folder = Path(self.input_folder_var.get().strip())
        return input_folder / self.listbox.get(selected[0])

    def _on_shapefile_selected(self, _event: tk.Event[tk.Misc] | None = None) -> None:
        shapefile = self._selected_shapefile_path()
        if shapefile is None:
            self.field_combo["values"] = []
            self.field_var.set("")
            return

        fields = self._list_exportable_fields(shapefile)
        self.field_combo["values"] = fields

        if self.field_var.get() in fields:
            return

        if fields:
            self.field_var.set(fields[0])
            self._save_settings()
        else:
            self.field_var.set("")

    def _on_field_selected(self, _event: tk.Event[tk.Misc] | None = None) -> None:
        self._save_settings()

    def select_output_folder(self) -> None:
        output_folder = filedialog.askdirectory(
            title="Select output folder",
            initialdir=str(self.last_output_folder),
        )
        if output_folder:
            self.output_folder_var.set(output_folder)
            self.last_output_folder = Path(output_folder)
            self._save_settings()

    def scan_shapefiles(self) -> None:
        self.listbox.delete(0, tk.END)
        input_folder = Path(self.input_folder_var.get().strip())

        if not input_folder.exists() or not input_folder.is_dir():
            self.status_var.set("Select a valid input subfolder.")
            return

        shapefiles = sorted(input_folder.glob("*.shp"))
        for shp in shapefiles:
            self.listbox.insert(tk.END, shp.name)

        if shapefiles:
            self.status_var.set(f"Found {len(shapefiles)} shapefile(s).")
            self.listbox.selection_clear(0, tk.END)
            self.listbox.selection_set(0)
            self._on_shapefile_selected()
        else:
            self.field_combo["values"] = []
            self.field_var.set("")
            self.status_var.set("No .shp files found in the selected folder.")

    def _list_exportable_fields(self, shapefile: Path) -> list[str]:
        return [
            field.name
            for field in arcpy.ListFields(str(shapefile))
            if field.type not in {"Geometry", "Raster", "Blob"}
        ]

    def _field_summary_rows(
        self, shapefile: Path, field_name: str
    ) -> pd.DataFrame:
        values: list[object] = []
        with arcpy.da.SearchCursor(str(shapefile), [field_name]) as cursor:
            for row in cursor:
                value = row[0]
                values.append("<NULL>" if value is None else value)

        series = pd.Series(values, name=field_name)
        summary = series.value_counts(dropna=False).rename_axis(field_name).reset_index(name="count")
        return summary

    def export_selected(self) -> None:
        selected_shp = self._selected_shapefile_path()
        if selected_shp is None:
            messagebox.showwarning("No selection", "Select one shapefile.")
            return

        selected_field = self.field_var.get().strip()
        if not selected_field:
            messagebox.showwarning("No field", "Select a field for the summary.")
            return

        output_folder = self.output_folder_var.get().strip()
        if not output_folder:
            messagebox.showwarning("No output", "Select an Excel output folder.")
            return

        try:
            summary_df = self._field_summary_rows(selected_shp, selected_field)
            output_dir = Path(output_folder)
            output_dir.mkdir(parents=True, exist_ok=True)
            safe_field = re.sub(r"[^A-Za-z0-9_-]", "_", selected_field)
            output = output_dir / f"{selected_shp.stem}_{safe_field}_summary.xlsx"

            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                summary_df.to_excel(writer, index=False, sheet_name="Summary")
                worksheet = writer.sheets["Summary"]
                for col_cells in worksheet.columns:
                    max_len = max(
                        len(str(cell.value)) if cell.value is not None else 0
                        for cell in col_cells
                    )
                    worksheet.column_dimensions[
                        col_cells[0].column_letter
                    ].width = max_len + 4
            self.last_output_folder = output_dir
            self.output_folder_var.set(str(output_dir))
            self._save_settings()
            messagebox.showinfo(
                "Export complete",
                f"Exported {len(summary_df)} unique value(s) for field '{selected_field}' to:\n{output}",
            )
            self.status_var.set(
                f"Exported unique values summary for field '{selected_field}'."
            )
            self._open_output_folder(output_dir)
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror("Export failed", str(exc))
            self.status_var.set("Export failed. See error message.")


def run() -> None:
    root = tk.Tk()
    ShapefileExportApp(root)
    root.mainloop()


if __name__ == "__main__":
    run()
