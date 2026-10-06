"""Read Excel files and create a merged Excel summary.

Rules implemented:
- First column is treated as id1.
- Second column is treated as id2, unless its header is 'count'.
- The column named 'HSI_min' is used as lookup value.
- Source Excel filename is included for each output row.
"""

from __future__ import annotations

import argparse
from datetime import datetime
import glob
import os
from pathlib import Path
import subprocess
import sys
from typing import Iterable


DEFAULT_INPUT_PATH = Path(r"E:/2026/ZoemZoneLimburg/OutputDataInfo")
IGNORED_INPUT_FILENAMES = {
    "summsry.xlsx",
    "summary.xlsx",
    "brunssum_ruw gras_fid_summary.xlsx",
}


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Read Excel files and write merged rows to an Excel summary file.",
    )
    parser.add_argument(
        "--input",
        nargs="+",
        default=[str(DEFAULT_INPUT_PATH)],
        help="Excel files or glob patterns (for example: data/*.xlsx).",
    )
    parser.add_argument(
        "--output",
        default=None,
        help="Optional output Excel path. Default: <input-folder>/summsry.xlsx",
    )
    return parser


def _expand_inputs(patterns: Iterable[str]) -> list[Path]:
    files: list[Path] = []
    for pattern in patterns:
        path = Path(pattern)
        if path.exists():
            if path.is_file() and path.suffix.lower() in {".xlsx", ".xlsm"}:
                files.append(path)
                continue

            if path.is_dir():
                for extension in ("*.xlsx", "*.xlsm"):
                    files.extend(sorted(path.rglob(extension)))
                continue

        # Resolve both relative and absolute glob patterns.
        matched = sorted(glob.glob(pattern, recursive=True))
        for item_text in matched:
            item = Path(item_text)
            if item.is_file() and item.suffix.lower() in {".xlsx", ".xlsm"}:
                files.append(item)

    unique_files = sorted({file.resolve() for file in files})
    return unique_files


def _normalize_header(value: object) -> str:
    return str(value or "").strip().lower()


def _cell_to_text(value: object) -> str:
    if value is None:
        return ""
    return str(value).strip()


def _extract_rows_from_file(path: Path) -> list[tuple[str, str, object, str]]:
    try:
        from openpyxl import load_workbook
    except ImportError as exc:
        raise RuntimeError(
            "Missing dependency 'openpyxl'. Install it with: pip install openpyxl"
        ) from exc

    workbook = load_workbook(filename=path, data_only=True, read_only=True)
    extracted_rows: list[tuple[str, str, object, str]] = []
    found_hsi_column = False

    for sheet in workbook.worksheets:
        rows = sheet.iter_rows(values_only=True)
        header_row = next(rows, None)
        if not header_row:
            continue

        headers = [_normalize_header(value) for value in header_row]
        hsi_index = next(
            (index for index, header in enumerate(headers) if header == "hsi_min"),
            -1,
        )
        if hsi_index < 0:
            continue
        found_hsi_column = True

        second_is_count = len(headers) >= 2 and headers[1] == "count"

        for row in rows:
            if not any(cell is not None and _cell_to_text(cell) != "" for cell in row):
                continue

            id1 = _cell_to_text(row[0] if len(row) > 0 else "")
            if not id1:
                continue

            id2 = ""
            if not second_is_count and len(row) > 1:
                id2 = _cell_to_text(row[1])

            hsi_value = row[hsi_index] if len(row) > hsi_index else None
            extracted_rows.append((id1, id2, hsi_value, path.name))

    workbook.close()

    if not found_hsi_column:
        return []
    return extracted_rows


def _default_output_path(input_files: list[Path]) -> Path:
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"summary_{timestamp}.xlsx"

    if len(input_files) == 1:
        only = input_files[0]
        if only.is_dir():
            return only / filename
        return only.parent / filename

    parents = {item.parent for item in input_files}
    if len(parents) == 1:
        return next(iter(parents)) / filename
    return Path.cwd() / filename


def _input_folder_label(input_files: list[Path]) -> str:
    parents = {item.parent for item in input_files}
    if len(parents) == 1:
        folder = next(iter(parents))
        return f"{folder.name} ({folder})"
    return "multiple-folders"


def _write_output_excel(
    rows: list[tuple[str, str, object, str]],
    output_path: Path,
    input_file_count: int,
    input_folder: str,
    run_timestamp: str,
    excluded_files: list[tuple[str, str]],
) -> None:
    try:
        from openpyxl import Workbook
        from openpyxl.utils import get_column_letter
    except ImportError as exc:
        raise RuntimeError(
            "Missing dependency 'openpyxl'. Install it with: pip install openpyxl"
        ) from exc

    workbook = Workbook()
    sheet = workbook.active
    if sheet is None:
        raise RuntimeError("Could not create an active worksheet in output workbook.")
    sheet.title = "summary"

    sheet.append(
        [
            f"input_files={input_file_count}",
            f"folder={input_folder}",
            f"timestamp={run_timestamp}",
            "",
        ]
    )
    sheet.append(["id1", "id2", "HSI_min", "source_excel"])
    for id1, id2, hsi_min, source_file in rows:
        sheet.append([id1, id2, hsi_min, source_file])

    sheet.append(["", "", "", ""])
    sheet.append(["Skipped files", "Reason", "", ""])
    for filename, reason in sorted(excluded_files, key=lambda item: item[0].lower()):
        sheet.append([filename, reason, "", ""])

    # Autosize columns based on the maximum text length in each column.
    for column_index, column_cells in enumerate(sheet.columns, start=1):
        max_len = 0
        col_letter = get_column_letter(column_index)
        for cell in column_cells:
            value = "" if cell.value is None else str(cell.value)
            if len(value) > max_len:
                max_len = len(value)
        sheet.column_dimensions[col_letter].width = max(10, min(max_len + 2, 80))

    output_path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(output_path)
    workbook.close()


def _open_output_file(path: Path) -> None:
    """Open the generated output file with the system default application."""
    try:
        if sys.platform.startswith("win"):
            os.startfile(str(path))  # type: ignore[attr-defined]
        elif sys.platform == "darwin":
            subprocess.Popen(["open", str(path)])
        else:
            subprocess.Popen(["xdg-open", str(path)])
    except Exception as exc:
        print(f"Warning: could not open output file automatically: {exc}")


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()

    files = _expand_inputs(args.input)
    if not files:
        parser.error("No Excel files found. Check --input paths or patterns.")

    excluded_files: list[tuple[str, str]] = []
    usable_files: list[Path] = []

    for item in files:
        lower_name = item.name.lower()
        if item.name.startswith("~$"):
            excluded_files.append((item.name, "Excel temporary/lock file"))
            continue
        if lower_name in IGNORED_INPUT_FILENAMES:
            excluded_files.append((item.name, "Ignored by script configuration"))
            continue
        if lower_name.startswith("hsi_updates"):
            excluded_files.append((item.name, "Ignored HSI updates file"))
            continue
        if lower_name.startswith("summsry") or lower_name.startswith("summary"):
            excluded_files.append((item.name, "Generated summary output file"))
            continue
        usable_files.append(item)

    if not usable_files:
        parser.error("Only ignored files were found; no source Excel files to process.")

    run_timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    input_folder = _input_folder_label(usable_files)
    output_path = Path(args.output).resolve() if args.output else _default_output_path(usable_files)

    merged_rows: list[tuple[str, str, object, str]] = []
    for path in usable_files:
        file_rows = _extract_rows_from_file(path)
        if not file_rows:
            excluded_files.append(
                (
                    path.name,
                    "No usable rows (missing HSI_min column or no valid id1 values)",
                )
            )
            continue
        merged_rows.extend(file_rows)

    merged_rows.sort(key=lambda item: (item[0], item[1]))
    _write_output_excel(
        merged_rows,
        output_path,
        input_file_count=len(usable_files),
        input_folder=input_folder,
        run_timestamp=run_timestamp,
        excluded_files=excluded_files,
    )
    _open_output_file(output_path)

    print(f"Rows written: {len(merged_rows)}")
    print(f"Output file: {output_path}")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
