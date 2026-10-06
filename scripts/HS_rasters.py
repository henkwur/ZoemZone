"""
Rasterize shapefile/feature class data with GUI.

Rasterizes all shapefiles in an input folder to GeoTIFF format.
Supports point, line, and polygon feature types.
"""

import arcpy
import logging
import os
import time
from datetime import datetime
import re
import traceback
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinter import scrolledtext
import threading


class RasterizeGUI:
    """GUI for rasterizing HSI data."""

    def __init__(self, root):
        """Initialize the GUI."""
        self.root = root
        self.root.title("HSI Rasterization Tool")
        self.root.geometry("1200x900")
        self.root.minsize(1000, 700)

        # Default paths
        self.default_input = r"E:\2026\ZoemZoneLimburg\HSI"
        self.default_merge_input = r"E:\2026\ZoemZoneLimburg\HSI_rasters"
        self.default_output = r"E:\2026\ZoemZoneLimburg\HSI_rasters"
        self.default_merge_output = r"E:\2026\ZoemZoneLimburg\Limburg_HSI_2026"
        self.default_snap = r"E:\2026\ZoemZoneLimburg\BNL2024\BNL2024.tif"

        # Variables
        self.input_folder = tk.StringVar(value=self.default_input)
        self.merge_input_folder = tk.StringVar(value=self.default_merge_input)
        self.output_folder = tk.StringVar(value=self.default_output)
        self.merge_output_folder = tk.StringVar(value=self.default_merge_output)
        self.snap_file = tk.StringVar(value=self.default_snap)
        self.default_log_folder = Path.home() / ".zoemzone" / "logs"
        self.log_folder = tk.StringVar(value=str(self.default_log_folder))
        self.cellsize = tk.StringVar(value="2.5")
        self.attribute = tk.StringVar(value="HSI_min")

        # Feature type selections
        self.include_points = tk.BooleanVar(value=True)
        self.include_lines = tk.BooleanVar(value=True)
        self.include_polygons = tk.BooleanVar(value=True)
        self.test_mode = tk.BooleanVar(value=False)

        # File counts
        self.file_counts = {"Point": 0, "Polyline": 0, "Polygon": 0}
        self.cancel_requested = False
        self.status_var = tk.StringVar(value="Idle")

        self.file_logger = logging.getLogger(f"zoemzone.hs_rasters.{id(self)}")
        self.file_logger.setLevel(logging.INFO)
        self.file_logger.propagate = False
        self._configure_file_logger(self.default_log_folder)

        self._create_widgets()
        self.root.resizable(True, True)
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

        # Allow the main UI area to expand with window resizing.
        self.root.grid_rowconfigure(0, weight=1)
        self.root.grid_columnconfigure(0, weight=1)

    def _create_widgets(self):
        """Create GUI widgets."""
        # Scrollable content keeps the lower controls reachable on smaller screens.
        content_frame = ttk.Frame(self.root)
        content_frame.grid(row=0, column=0, sticky="nsew")
        content_frame.grid_rowconfigure(0, weight=1)
        content_frame.grid_columnconfigure(0, weight=1)

        canvas = tk.Canvas(content_frame, highlightthickness=0)
        scrollbar = ttk.Scrollbar(content_frame, orient=tk.VERTICAL, command=canvas.yview)
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")

        main_frame = ttk.Frame(canvas, padding="10")
        canvas_window = canvas.create_window((0, 0), window=main_frame, anchor="nw")

        def update_scroll_region(_event=None):
            canvas.configure(scrollregion=canvas.bbox("all"))

        def resize_content(event):
            canvas.itemconfigure(canvas_window, width=event.width)

        main_frame.bind("<Configure>", update_scroll_region)
        canvas.bind("<Configure>", resize_content)
        canvas.bind("<MouseWheel>", lambda event: canvas.yview_scroll(-int(event.delta / 120), "units"))

        main_frame.grid_columnconfigure(1, weight=1)
        main_frame.grid_rowconfigure(14, weight=1)

        # Input folder
        ttk.Label(main_frame, text="Input Folder:", font=("Arial", 10, "bold")).grid(
            row=0, column=0, sticky=tk.W, pady=5
        )
        ttk.Entry(main_frame, textvariable=self.input_folder, width=60).grid(
            row=0, column=1, padx=5
        )
        ttk.Button(
            main_frame, text="Browse", command=self._browse_input_folder
        ).grid(row=0, column=2, padx=5)

        # Rasterization output folder
        ttk.Label(main_frame, text="Raster Output Folder:", font=("Arial", 10, "bold")).grid(
            row=1, column=0, sticky=tk.W, pady=5
        )
        ttk.Entry(main_frame, textvariable=self.output_folder, width=60).grid(
            row=1, column=1, padx=5
        )
        ttk.Button(
            main_frame, text="Browse", command=self._browse_output_folder
        ).grid(row=1, column=2, padx=5)

        # Merge input folder
        ttk.Label(main_frame, text="Merge Input Folder:", font=("Arial", 10, "bold")).grid(
            row=2, column=0, sticky=tk.W, pady=5
        )
        ttk.Entry(main_frame, textvariable=self.merge_input_folder, width=60).grid(
            row=2, column=1, padx=5
        )
        ttk.Button(
            main_frame, text="Browse", command=self._browse_merge_input_folder
        ).grid(row=2, column=2, padx=5)

        # Merge output folder
        ttk.Label(main_frame, text="Merge Output Folder:", font=("Arial", 10, "bold")).grid(
            row=3, column=0, sticky=tk.W, pady=5
        )
        ttk.Entry(main_frame, textvariable=self.merge_output_folder, width=60).grid(
            row=3, column=1, padx=5
        )
        ttk.Button(
            main_frame, text="Browse", command=self._browse_merge_output_folder
        ).grid(row=3, column=2, padx=5)

        # Snap file
        ttk.Label(main_frame, text="Snap File (Extent):", font=("Arial", 10, "bold")).grid(
            row=4, column=0, sticky=tk.W, pady=5
        )
        ttk.Entry(main_frame, textvariable=self.snap_file, width=60).grid(
            row=4, column=1, padx=5
        )
        ttk.Button(
            main_frame, text="Browse", command=self._browse_snap_file
        ).grid(row=4, column=2, padx=5)

        # Cell size
        ttk.Label(main_frame, text="Cell Size (m):", font=("Arial", 10, "bold")).grid(
            row=5, column=0, sticky=tk.W, pady=5
        )
        ttk.Entry(main_frame, textvariable=self.cellsize, width=20).grid(
            row=5, column=1, sticky=tk.W, padx=5
        )

        # Attribute field
        ttk.Label(main_frame, text="Attribute Field:", font=("Arial", 10, "bold")).grid(
            row=6, column=0, sticky=tk.W, pady=5
        )
        self.attribute_selector = ttk.Combobox(
            main_frame,
            textvariable=self.attribute,
            values=("HSI_min", "HSI_max"),
            state="readonly",
            width=18,
        )
        self.attribute_selector.grid(
            row=6, column=1, sticky=tk.W, padx=5
        )

        # Separator
        ttk.Separator(main_frame, orient=tk.HORIZONTAL).grid(
            row=7, column=0, columnspan=3, sticky="ew", pady=10
        )

        # Feature type selection
        ttk.Label(main_frame, text="Input Feature Types:", font=("Arial", 10, "bold")).grid(
            row=8, column=0, sticky=tk.W, pady=5
        )

        types_frame = ttk.Frame(main_frame)
        types_frame.grid(row=9, column=0, columnspan=3, sticky="ew", padx=5, pady=5)

        ttk.Checkbutton(
            types_frame, text="Points", variable=self.include_points, command=self._update_file_counts
        ).pack(anchor=tk.W)
        ttk.Checkbutton(
            types_frame, text="Lines", variable=self.include_lines, command=self._update_file_counts
        ).pack(anchor=tk.W)
        ttk.Checkbutton(
            types_frame, text="Polygons", variable=self.include_polygons, command=self._update_file_counts
        ).pack(anchor=tk.W)
        ttk.Checkbutton(
            types_frame,
            text="Test mode (process only first selected file)",
            variable=self.test_mode,
        ).pack(anchor=tk.W, pady=(4, 0))

        # File counts display
        ttk.Label(main_frame, text="File Counts:", font=("Arial", 10, "bold")).grid(
            row=10, column=0, sticky=tk.W, pady=(10, 5)
        )

        counts_frame = ttk.Frame(main_frame)
        counts_frame.grid(row=11, column=0, columnspan=3, sticky="ew", padx=5, pady=5)

        self.points_label = ttk.Label(counts_frame, text="Points: 0")
        self.points_label.pack(anchor=tk.W)

        self.lines_label = ttk.Label(counts_frame, text="Lines: 0")
        self.lines_label.pack(anchor=tk.W)

        self.polygons_label = ttk.Label(counts_frame, text="Polygons: 0")
        self.polygons_label.pack(anchor=tk.W)

        self.total_label = ttk.Label(counts_frame, text="Total: 0", font=("Arial", 10, "bold"))
        self.total_label.pack(anchor=tk.W)

        # Progress bar
        ttk.Label(main_frame, text="Progress:", font=("Arial", 10, "bold")).grid(
            row=12, column=0, sticky=tk.W, pady=(10, 5)
        )

        self.progress_var = tk.IntVar()
        self.progress_bar = ttk.Progressbar(
            main_frame, variable=self.progress_var, maximum=100, mode="determinate"
        )
        self.progress_bar.grid(row=13, column=0, columnspan=3, sticky="ew", padx=5, pady=5)

        self.progress_label = ttk.Label(main_frame, text="Ready")
        self.progress_label.grid(row=14, column=0, columnspan=3, sticky=tk.W, padx=5)

        # Status display
        ttk.Label(main_frame, text="Status:", font=("Arial", 10, "bold")).grid(
            row=15, column=0, sticky=tk.W, pady=(5, 5)
        )
        self.status_label = ttk.Label(main_frame, textvariable=self.status_var)
        self.status_label.grid(row=15, column=1, columnspan=2, sticky=tk.W, padx=5, pady=(5, 5))

        # Log folder
        ttk.Label(main_frame, text="Log Folder:", font=("Arial", 10, "bold")).grid(
            row=16, column=0, sticky=tk.W, pady=(5, 5)
        )
        ttk.Entry(main_frame, textvariable=self.log_folder, width=60).grid(
            row=16, column=1, padx=5
        )
        ttk.Button(
            main_frame, text="Browse", command=self._browse_log_folder
        ).grid(row=16, column=2, padx=5)

        # Processing log
        ttk.Label(main_frame, text="Processing Log:", font=("Arial", 10, "bold")).grid(
            row=18, column=0, sticky=tk.W, pady=(10, 5)
        )
        self.process_text = scrolledtext.ScrolledText(
            main_frame, height=14, wrap=tk.WORD, font=("Consolas", 10)
        )
        self.process_text.grid(row=19, column=0, columnspan=3, sticky="nsew", padx=5, pady=5)
        self.process_text.insert(tk.END, "Ready.\n")
        self.process_text.configure(state=tk.DISABLED)

        # Error log
        ttk.Label(main_frame, text="Error Log (copy/paste enabled):", font=("Arial", 10, "bold")).grid(
            row=20, column=0, sticky=tk.W, pady=(10, 5)
        )
        self.error_text = scrolledtext.ScrolledText(main_frame, height=7, wrap=tk.WORD)
        self.error_text.grid(row=21, column=0, columnspan=3, sticky="ew", padx=5, pady=5)
        self.error_text.insert(tk.END, "No errors yet.\n")
        self.error_text.configure(state=tk.DISABLED)

        # Buttons frame
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=22, column=0, columnspan=3, sticky="ew", pady=20)

        ttk.Button(
            button_frame, text="Scan Folder", command=self._update_file_counts
        ).pack(side=tk.LEFT, padx=5)

        ttk.Button(
            button_frame, text="Start Rasterization", command=self._start_rasterization_thread
        ).pack(side=tk.LEFT, padx=5)

        ttk.Button(
            button_frame, text="Merge Rasters", command=self._start_merge_thread
        ).pack(side=tk.LEFT, padx=5)

        ttk.Button(
            button_frame, text="Cancel", command=self._cancel_process
        ).pack(side=tk.LEFT, padx=5)

        ttk.Button(
            button_frame, text="Clear Logs", command=self._clear_logs
        ).pack(side=tk.LEFT, padx=5)

        ttk.Button(
            button_frame, text="Exit", command=self._on_close
        ).pack(side=tk.LEFT, padx=5)

    def _log_error(self, title: str, exc: Exception):
        """Write detailed errors to a copyable textbox."""
        details = traceback.format_exc()
        message = f"{title}: {exc}\n{details}\n"
        self.file_logger.error(message.rstrip())
        self.error_text.configure(state=tk.NORMAL)
        if self.error_text.get("1.0", tk.END).strip() == "No errors yet.":
            self.error_text.delete("1.0", tk.END)
        self.error_text.insert(tk.END, message)
        self.error_text.see(tk.END)
        self.error_text.configure(state=tk.DISABLED)

    def _log_process(self, message: str):
        """Write process messages to the processing log textbox."""
        self.file_logger.info(message)
        self.process_text.configure(state=tk.NORMAL)
        if self.process_text.get("1.0", tk.END).strip() == "Ready.":
            self.process_text.delete("1.0", tk.END)
        self.process_text.insert(tk.END, message + "\n")
        self.process_text.see(tk.END)
        self.process_text.configure(state=tk.DISABLED)
        self.process_text.update_idletasks()

    def _set_status(self, value: str):
        """Update the GUI status label safely."""
        self.status_var.set(value)
        self.status_label.update_idletasks()

    def _configure_file_logger(self, folder: Path):
        """Create a timestamped log file in the selected folder."""
        folder.mkdir(parents=True, exist_ok=True)
        for handler in self.file_logger.handlers[:]:
            handler.close()
            self.file_logger.removeHandler(handler)

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        self.log_file = folder / f"hs_rasters_{timestamp}.log"
        file_handler = logging.FileHandler(self.log_file, encoding="utf-8")
        file_handler.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(message)s"))
        self.file_logger.addHandler(file_handler)

    def _browse_log_folder(self):
        """Select the folder for the current log file."""
        folder = filedialog.askdirectory(
            initialdir=self.log_folder.get(),
            title="Select Log Folder"
        )
        if folder:
            self.log_folder.set(folder)
            self._configure_file_logger(Path(folder))
            self._log_process(f"LOG   file: {self.log_file}")

    def _clear_logs(self):
        """Clear both error and processing logs."""
        self.error_text.configure(state=tk.NORMAL)
        self.error_text.delete("1.0", tk.END)
        self.error_text.insert(tk.END, "No errors yet.\n")
        self.error_text.configure(state=tk.DISABLED)

        self.process_text.configure(state=tk.NORMAL)
        self.process_text.delete("1.0", tk.END)
        self.process_text.insert(tk.END, "Ready.\n")
        self.process_text.configure(state=tk.DISABLED)

    def _cancel_process(self):
        """Request cancellation of the current raster or merge run."""
        self.cancel_requested = True
        self._log_process("CANCEL requested by user")
        self.progress_label.config(text="Cancel requested")
        self._set_status("Cancelling")

    def _on_close(self):
        """Handle window close with a shutdown log entry."""
        try:
            self._log_process("APP shutting down")
            self._set_status("Closing")
        except Exception:
            pass
        for handler in self.file_logger.handlers[:]:
            handler.close()
            self.file_logger.removeHandler(handler)
        self.root.destroy()

    @staticmethod
    def _shape_type(path: Path) -> str:
        """Return lowercase ArcGIS geometry type for a feature dataset."""
        return str(arcpy.Describe(str(path)).shapeType).lower()

    @staticmethod
    def _delete_if_exists(path: str):
        """Delete a GIS dataset safely if it exists."""
        try:
            if arcpy.Exists(path):
                arcpy.management.Delete(path)
        except Exception:
            pass

    @staticmethod
    def _sanitize_raster_name(name: str) -> str:
        """Replace characters that are unsafe for raster dataset names."""
        sanitized_name = re.sub(r"[^A-Za-z0-9_]+", "_", name)
        sanitized_name = re.sub(r"_+", "_", sanitized_name).strip("_")
        return sanitized_name or "raster"

    @staticmethod
    def _resolve_raster_attribute(input_file: str, requested_attribute: str) -> str:
        """Use HSI_min when a selected HSI_max field is absent."""
        field_names = {field.name.casefold() for field in arcpy.ListFields(input_file)}
        if requested_attribute.casefold() == "hsi_max" and "hsi_max" not in field_names:
            return "HSI_min"
        return requested_attribute

    def _browse_input_folder(self):
        """Browse for input folder."""
        folder = filedialog.askdirectory(
            initialdir=self.input_folder.get(),
            title="Select Input Folder"
        )
        if folder:
            self.input_folder.set(folder)
            self._update_file_counts()

    def _browse_output_folder(self):
        """Browse for output folder."""
        folder = filedialog.askdirectory(
            initialdir=self.output_folder.get(),
            title="Select Output Folder"
        )
        if folder:
            self.output_folder.set(folder)

    def _browse_merge_input_folder(self):
        """Browse for merge input folder."""
        folder = filedialog.askdirectory(
            initialdir=self.merge_input_folder.get(),
            title="Select Merge Input Folder"
        )
        if folder:
            self.merge_input_folder.set(folder)

    def _browse_merge_output_folder(self):
        """Browse for merge output folder."""
        folder = filedialog.askdirectory(
            initialdir=self.merge_output_folder.get(),
            title="Select Merge Output Folder"
        )
        if folder:
            self.merge_output_folder.set(folder)

    def _browse_snap_file(self):
        """Browse for snap file."""
        file = filedialog.askopenfilename(
            initialdir=os.path.dirname(self.snap_file.get()),
            title="Select Snap File",
            filetypes=[("TIFF Files", "*.tif"), ("All Files", "*.*")]
        )
        if file:
            self.snap_file.set(file)

    def _update_file_counts(self):
        """Count input files by feature type."""
        input_path = Path(self.input_folder.get())

        self._log_process(f"SCAN  checking input folder: {input_path}")

        if not input_path.exists():
            self.error_text.configure(state=tk.NORMAL)
            self.error_text.delete("1.0", tk.END)
            self.error_text.insert(tk.END, f"Input folder does not exist: {input_path}\n")
            self.error_text.configure(state=tk.DISABLED)
            self._log_process(f"SCAN  failed: input folder not found")
            return

        self.file_counts = {"Point": 0, "Polyline": 0, "Polygon": 0}

        try:
            # List all shapefiles
            shapefiles = list(input_path.glob("*.shp"))
            self._log_process(f"SCAN  found {len(shapefiles)} shapefile(s)")

            for shp_file in shapefiles:
                try:
                    shape_type = self._shape_type(shp_file)
                    self._log_process(f"SCAN  file: {shp_file.name} | type: {shape_type}")
                    if shape_type == "point":
                        self.file_counts["Point"] += 1
                    elif shape_type == "polyline":
                        self.file_counts["Polyline"] += 1
                    elif shape_type == "polygon":
                        self.file_counts["Polygon"] += 1
                except Exception as e:
                    self._log_error(f"Error reading {shp_file}", e)

            # Update labels
            self.points_label.config(text=f"Points: {self.file_counts['Point']}")
            self.lines_label.config(text=f"Lines: {self.file_counts['Polyline']}")
            self.polygons_label.config(text=f"Polygons: {self.file_counts['Polygon']}")

            total = sum(self.file_counts.values())
            self.total_label.config(text=f"Total: {total}")
            self._log_process(
                "SCAN  counts -> points: {points}, lines: {lines}, polygons: {polygons}, total: {total}".format(
                    points=self.file_counts["Point"],
                    lines=self.file_counts["Polyline"],
                    polygons=self.file_counts["Polygon"],
                    total=total,
                )
            )

        except Exception as e:
            self._log_error("Error scanning folder", e)

    def _start_rasterization_thread(self):
        """Start rasterization in a separate thread."""
        thread = threading.Thread(target=self._rasterize_all)
        thread.daemon = True
        thread.start()

    def _start_merge_thread(self):
        """Start raster merging in a separate thread."""
        thread = threading.Thread(target=self._merge_rasters)
        thread.daemon = True
        thread.start()

    def _rasterize_all(self):
        """Rasterize all input files."""
        self.cancel_requested = False
        self._set_status("Rasterizing")
        try:
            input_path = Path(self.input_folder.get())
            output_path = Path(self.output_folder.get())
            snap_file = self.snap_file.get()
            cellsize = float(self.cellsize.get())
            attribute = self.attribute.get()

            # Validate inputs
            if not input_path.exists():
                raise FileNotFoundError(f"Input folder does not exist: {input_path}")
            if not os.path.exists(snap_file):
                raise FileNotFoundError(f"Snap file does not exist: {snap_file}")

            # Create output folder
            output_path.mkdir(parents=True, exist_ok=True)
            self._log_process(f"RUN   rasterization started")
            self._log_process(f"RUN   input folder: {input_path}")
            self._log_process(f"RUN   output folder: {output_path}")
            self._log_process(f"RUN   snap file: {snap_file}")
            self._log_process(f"RUN   cell size: {cellsize}")
            self._log_process(f"RUN   attribute: {attribute}")
            self._log_process(f"RUN   source folder for features: {input_path}")
            self._log_process(f"RUN   raster output folder: {output_path}")
            self._log_process(f"RUN   snap alignment file: {snap_file}")
            self._log_process(
                "RUN   types enabled -> points: {points}, lines: {lines}, polygons: {polygons}, test mode: {test}".format(
                    points=self.include_points.get(),
                    lines=self.include_lines.get(),
                    polygons=self.include_polygons.get(),
                    test=self.test_mode.get(),
                )
            )
        except Exception as e:
            self._log_error("Input validation/setup failed", e)
            self.progress_label.config(text="Error occurred")
            return

        # Collect files to process
        shapefiles = list(input_path.glob("*.shp"))
        files_to_process = []

        for shp_file in shapefiles:
            try:
                shape_type = self._shape_type(shp_file)
                if shape_type == "point" and self.include_points.get():
                    files_to_process.append(str(shp_file))
                elif shape_type == "polyline" and self.include_lines.get():
                    files_to_process.append(str(shp_file))
                elif shape_type == "polygon" and self.include_polygons.get():
                    files_to_process.append(str(shp_file))
            except Exception as e:
                self._log_error(f"Error reading {shp_file}", e)

        self._log_process(f"RUN   selected {len(files_to_process)} file(s) for rasterization")

        if self.test_mode.get() and files_to_process:
            self._log_process("RUN   test mode active: limiting to first selected file")
            files_to_process = files_to_process[:1]

        if not files_to_process:
            self.progress_label.config(text="No files to process")
            self.error_text.configure(state=tk.NORMAL)
            self.error_text.insert(tk.END, "No files to process with selected types.\n")
            self.error_text.see(tk.END)
            self.error_text.configure(state=tk.DISABLED)
            self._log_process("RUN   no files matched the selected feature types")
            return

        # Process files
        total_files = len(files_to_process)
        self.progress_var.set(0)
        rasterized_files = []

        try:
            snap_desc = arcpy.Describe(snap_file)
            self._log_process(f"RUN   output coordinate system: {snap_desc.spatialReference.name}")
            self._log_process("RUN   output will be snapped to the reference grid")
            for idx, input_file in enumerate(files_to_process):
                if self.cancel_requested:
                    self._log_process("RUN   rasterization cancelled before next file")
                    self.progress_label.config(text="Rasterization cancelled")
                    self._set_status("Cancelled")
                    break
                with arcpy.EnvManager(
                    workspace=str(output_path),
                    overwriteOutput=True,
                    snapRaster=snap_file,
                    extent=input_file,
                    cellSize=cellsize,
                    outputCoordinateSystem=snap_desc.spatialReference,
                ):
                    input_name = Path(input_file).stem
                    raster_attribute = self._resolve_raster_attribute(input_file, attribute)
                    if raster_attribute != attribute:
                        self._log_process(
                            f"FIELD  {Path(input_file).name} | {attribute} missing, using HSI_min"
                        )
                    safe_name = self._sanitize_raster_name(input_name)
                    if safe_name != input_name:
                        self._log_process(
                            f"NAME  {Path(input_file).name} | remapped to {safe_name}.tif"
                        )
                    safe_attribute = self._sanitize_raster_name(raster_attribute)
                    output_name = f"{safe_name}_{safe_attribute}.tif"
                    output_raster = str(output_path / output_name)
                    self._log_process(f"FILE  {Path(input_file).name} -> {Path(output_raster).name}")
                    self._log_process(
                        f"FILE  applying snap raster and input extent for {Path(input_file).name}"
                    )

                    if os.path.exists(output_raster):
                        self._log_process(
                            f"SKIP  {Path(input_file).name} | output exists: {Path(output_raster).name}"
                        )
                        progress_pct = int((idx + 1) / total_files * 100)
                        self.progress_var.set(progress_pct)
                        self.progress_label.config(
                            text=(
                                f"Skipped existing: {Path(input_file).name} "
                                f"({idx + 1}/{total_files})"
                            )
                        )
                        self.root.update()
                        continue

                    # Update progress label
                    self.progress_label.config(
                        text=f"Processing: {Path(input_file).name} ({idx + 1}/{total_files})"
                    )
                    self.root.update()

                    # Rasterize
                    start_clock = datetime.now()
                    start_time = time.perf_counter()
                    self._log_process(
                        f"START {Path(input_file).name} | {start_clock.strftime('%Y-%m-%d %H:%M:%S')}"
                    )
                    self._log_process("STEP  generating raster from feature class")
                    self._rasterize_file(
                        input_file, output_raster, raster_attribute
                    )
                    rasterized_files.append(Path(input_file).name)
                    elapsed_seconds = time.perf_counter() - start_time
                    end_clock = datetime.now()
                    elapsed_text = f"{elapsed_seconds:.1f}s"

                    self._log_process(
                        "END   {name} | {end} | duration {duration}".format(
                            name=Path(input_file).name,
                            end=end_clock.strftime('%Y-%m-%d %H:%M:%S'),
                            duration=elapsed_text,
                        )
                    )

                    # Update progress bar
                    progress_pct = int((idx + 1) / total_files * 100)
                    self.progress_var.set(progress_pct)
                    self.progress_label.config(
                        text=(
                            f"Done: {Path(input_file).name} in {elapsed_text} "
                            f"({idx + 1}/{total_files})"
                        )
                    )
                    self.root.update()

            if rasterized_files:
                self._log_process("RUN   rasterized files:")
                for rasterized_name in rasterized_files:
                    self._log_process(f"DONE  {rasterized_name}")
            else:
                self._log_process("RUN   no files were rasterized")

            if self.cancel_requested:
                messagebox.showinfo("Cancelled", "Rasterization was cancelled by the user.")
                self._set_status("Cancelled")
            else:
                self.progress_label.config(text="Completed successfully!")
                self._set_status("Idle")
                messagebox.showinfo("Success", f"Rasterized {total_files} file(s)")

        except Exception as e:
            self._log_error("Rasterization failed", e)
            self.progress_label.config(text="Error occurred")
            self._set_status("Error")

    def _rasterize_file(
        self, input_file: str, output_raster: str, attribute: str
    ):
        """Rasterize a single file."""
        temp_feature = output_raster.replace(".tif", "__tmp_feature.tif")
        self._delete_if_exists(temp_feature)

        try:
            # Keep intermediate processing lightweight: no pyramids/statistics on temp data.
            self._log_process(f"STEP  preparing intermediate raster: {Path(temp_feature).name}")
            with arcpy.EnvManager(pyramid="NONE", rasterStatistics="NONE"):
                # Create raster using snap/cell alignment from environment settings.
                self._log_process(f"STEP  FeatureToRaster using field '{attribute}'")
                arcpy.conversion.FeatureToRaster(
                    in_features=input_file,
                    field=attribute,
                    out_raster=temp_feature,
                )

                # Enforce floating-point TIFF output in one step to final output.
                self._log_process(f"STEP  CopyRaster -> {Path(output_raster).name} as 32-bit float TIFF")
                arcpy.management.CopyRaster(
                    in_raster=temp_feature,
                    out_rasterdataset=output_raster,
                    nodata_value="-9999",
                    pixel_type="32_BIT_FLOAT",
                    format="TIFF",
                )

            # Validate that the output is a raster dataset and compute statistics.
            self._log_process(f"STEP  validating output raster: {Path(output_raster).name}")
            out_desc = arcpy.Describe(output_raster)
            if str(getattr(out_desc, "dataType", "")).lower() != "rasterdataset":
                raise RuntimeError(f"Output is not a raster dataset: {output_raster}")

            # Build pyramids only for the final output.
            self._log_process(f"STEP  building pyramids for {Path(output_raster).name}")
            arcpy.management.BuildPyramids(output_raster)
            self._log_process(f"STEP  calculating statistics for {Path(output_raster).name}")
            arcpy.management.CalculateStatistics(output_raster)
            self._log_process(f"STEP  final output ready: {Path(output_raster).name}")
        finally:
            self._delete_if_exists(temp_feature)
            self._log_process(f"STEP  cleaned temporary file: {Path(temp_feature).name}")

    def _merge_rasters(self):
        """Merge all TIFF rasters in the output folder into Limburg_HSI.tif."""
        self.cancel_requested = False
        self._set_status("Merging")
        try:
            input_path = Path(self.merge_input_folder.get())
            output_path = Path(self.merge_output_folder.get())
            snap_file = self.snap_file.get()
            cellsize = float(self.cellsize.get())
            attribute = self.attribute.get()
        except Exception as e:
            self._log_error("Merge setup failed", e)
            self.progress_label.config(text="Error occurred")
            return

        self._log_process(f"MERGE checking folder: {input_path}")

        if not input_path.exists():
            self._log_error("Merge setup failed", FileNotFoundError(f"Raster folder does not exist: {input_path}"))
            self.progress_label.config(text="Error occurred")
            self._log_process("MERGE failed: source folder missing")
            return

        if not os.path.exists(snap_file):
            self._log_error("Merge setup failed", FileNotFoundError(f"Snap file does not exist: {snap_file}"))
            self.progress_label.config(text="Error occurred")
            self._log_process("MERGE failed: snap file missing")
            return

        output_raster = output_path / f"Limburg_HSI_{attribute}.tif"
        raster_files = sorted(
            str(path)
            for path in input_path.glob("*.tif")
            if path.is_file() and path.name.lower() != output_raster.name.lower()
        )
        raster_files.extend(
            str(path)
            for path in input_path.glob("*.tiff")
            if path.is_file() and path.name.lower() != output_raster.name.lower()
        )

        for raster_file in raster_files:
            self._log_process(f"MERGE file: {Path(raster_file).name}")

        if not raster_files:
            self._log_process(f"MERGE no rasters found in {input_path}")
            self.progress_label.config(text="No rasters to merge")
            return

        output_path.mkdir(parents=True, exist_ok=True)
        self._log_process(f"MERGE found {len(raster_files)} raster file(s)")
        self._log_process(f"MERGE source folder: {input_path}")
        self._log_process(f"MERGE output folder: {output_path}")
        self._log_process(f"MERGE output target: {output_raster.name}")
        self._log_process(f"MERGE attribute field: {attribute}")

        if output_raster.exists() or arcpy.Exists(str(output_raster)):
            self._log_process(f"SKIP merge | output exists: {output_raster.name}")
            self.progress_label.config(text="Merge output already exists")
            return

        self.progress_var.set(0)
        self.progress_label.config(text=f"Merging {len(raster_files)} rasters...")
        self.root.update()

        try:
            snap_desc = arcpy.Describe(snap_file)
            self._log_process(
                f"MERGE START | {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} | {len(raster_files)} inputs"
            )
            self._log_process(f"MERGE snap raster: {snap_file}")
            self._log_process(f"MERGE cell size: {cellsize}")
            self._log_process(f"MERGE mosaic method: MAXIMUM")

            with arcpy.EnvManager(
                overwriteOutput=True,
                snapRaster=snap_file,
                cellSize=cellsize,
                outputCoordinateSystem=snap_desc.spatialReference,
            ):
                if self.cancel_requested:
                    self._log_process("MERGE cancelled before MosaicToNewRaster")
                    self.progress_label.config(text="Merge cancelled")
                    self._set_status("Cancelled")
                    return
                self._log_process("MERGE running MosaicToNewRaster")
                arcpy.management.MosaicToNewRaster(
                    input_rasters=raster_files,
                    output_location=str(output_path),
                    raster_dataset_name_with_extension=output_raster.name,
                    coordinate_system_for_the_raster=snap_desc.spatialReference,
                    pixel_type="32_BIT_FLOAT",
                    cellsize=cellsize,
                    number_of_bands=1,
                    mosaic_method="MAXIMUM",
                    mosaic_colormap_mode="FIRST",
                )

            self._log_process(f"MERGE building pyramids for {output_raster.name}")
            arcpy.management.BuildPyramids(str(output_raster))
            self._log_process(f"MERGE calculating statistics for {output_raster.name}")
            arcpy.management.CalculateStatistics(str(output_raster))

            self._log_process(f"MERGE validating output raster type")
            out_desc = arcpy.Describe(str(output_raster))
            if str(getattr(out_desc, "dataType", "")).lower() != "rasterdataset":
                raise RuntimeError(f"Merged output is not a raster dataset: {output_raster}")

            self.progress_var.set(100)
            self.progress_label.config(text=f"Merge complete: {output_raster.name}")
            self._log_process(
                f"MERGE END   | {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} | output {output_raster.name}"
            )
            if self.cancel_requested:
                messagebox.showinfo("Cancelled", "Merge was cancelled by the user.")
                self._set_status("Cancelled")
            else:
                self._set_status("Idle")
                messagebox.showinfo("Success", f"Created {output_raster.name}")

        except Exception as e:
            self._log_error("Merge failed", e)
            self.progress_label.config(text="Error occurred")
            self._set_status("Error")


def main():
    """Main entry point."""
    root = tk.Tk()
    app = RasterizeGUI(root)
    app._log_process(f"APP started | {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    root.mainloop()


if __name__ == "__main__":
    main()
