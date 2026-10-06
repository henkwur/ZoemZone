"""Join an Excel score column to a feature class using a 2-field composite key.

ArcGIS Pro joins are usually configured on a single key field. This script creates
an intermediate text key on both datasets from two attributes and performs the
join using that key.
"""

from __future__ import annotations

import argparse
import json
from pathlib import Path
import re
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import uuid


SETTINGS_FILE = Path.home() / ".joinmultiple_settings.json"
DEFAULT_OUTPUT_FOLDER = Path(r"E:/2026/ZoemZoneLimburg/HSI")


def _sanitize_filename_part(value: str) -> str:
    safe = re.sub(r'[<>:"/\\|?*]', "", value).strip()
    safe = re.sub(r"\s+", "_", safe)
    return safe


def _field_name_key(name: str) -> str:
    """Normalize field names for relaxed matching across sources."""
    return re.sub(r"[\s_]+", "", name).lower()


def _extract_gemeente_prefix_from_input(input_path: str) -> str:
    parent_name = Path(input_path).parent.name.strip()
    match = re.search(r"(?i)\bgemeente[\s_-]*(.+)$", parent_name)
    if not match:
        return ""

    gemeente_name = _sanitize_filename_part(match.group(1))
    if not gemeente_name:
        return ""
    return f"{gemeente_name}_"


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description=(
            "Join one score field from an Excel table to a feature class using "
            "a composite key of two fields."
        ),
    )
    parser.add_argument(
        "--gui",
        action="store_true",
        help="Launch GUI mode.",
    )
    parser.add_argument(
        "--input-features",
        help="Path to input feature class or shapefile.",
    )
    parser.add_argument(
        "--output-features",
        help="Path to output feature class or shapefile.",
    )
    parser.add_argument(
        "--shp-field-1",
        help="First key field in the feature class.",
    )
    parser.add_argument(
        "--shp-field-2",
        help="Second key field in the feature class.",
    )
    parser.add_argument(
        "--excel-file",
        help="Path to the Excel file (.xlsx).",
    )
    parser.add_argument(
        "--excel-sheet",
        help="Excel sheet name to read (without trailing $).",
    )
    parser.add_argument(
        "--excel-field-1",
        help="First key field in the Excel sheet.",
    )
    parser.add_argument(
        "--excel-field-2",
        help="Second key field in the Excel sheet.",
    )
    parser.add_argument(
        "--score-field",
        help="Field in the Excel sheet containing the score to join.",
    )
    parser.add_argument(
        "--keep-join-key",
        action="store_true",
        help="Keep temporary composite key fields in output.",
    )
    parser.add_argument(
        "--case-sensitive",
        action="store_true",
        help="Use case-sensitive matching (default is case-insensitive).",
    )
    return parser


def _required_cli_args(args: argparse.Namespace) -> list[str]:
    required = [
        "input_features",
        "output_features",
        "shp_field_1",
        "excel_file",
        "excel_sheet",
        "excel_field_1",
        "score_field",
    ]
    missing = [name for name in required if not getattr(args, name, None)]
    return missing


def _validate_join_field_setup(args: argparse.Namespace) -> str | None:
    shp2 = bool((getattr(args, "shp_field_2", "") or "").strip())
    excel2 = bool((getattr(args, "excel_field_2", "") or "").strip())
    if shp2 != excel2:
        return (
            "Second join field must be set on both SHP and Excel, or left empty on both."
        )
    return None


def _has_any_cli_join_arg(args: argparse.Namespace) -> bool:
    join_args = [
        "input_features",
        "output_features",
        "shp_field_1",
        "shp_field_2",
        "excel_file",
        "excel_sheet",
        "excel_field_1",
        "excel_field_2",
        "score_field",
    ]
    return any(bool(getattr(args, name, None)) for name in join_args)


def _assert_field_exists(table: str, field_name: str, table_label: str) -> None:
    import arcpy  # type: ignore

    names = {f.name for f in arcpy.ListFields(table)}
    if field_name not in names:
        raise ValueError(
            f"Field '{field_name}' not found in {table_label}: {table}"
        )


def _add_text_field_if_missing(table: str, field_name: str, length: int = 512) -> None:
    import arcpy  # type: ignore

    names = {f.name for f in arcpy.ListFields(table)}
    if field_name in names:
        return
    arcpy.management.AddField(table, field_name, "TEXT", field_length=length)


def _normalize_nulls_to_zero(table: str, field_names: list[str]) -> None:
    """Convert NULL values to zero in the provided fields.

    Numeric fields get 0, string-like fields get "0".
    """
    import arcpy  # type: ignore

    if not field_names:
        return

    field_meta = {f.name: f for f in arcpy.ListFields(table)}
    valid_fields = [name for name in field_names if name in field_meta]
    if not valid_fields:
        return

    numeric_types = {"SmallInteger", "Integer", "Single", "Double", "BigInteger"}
    text_types = {"String", "GUID"}

    with arcpy.da.UpdateCursor(table, valid_fields) as cursor:
        for row in cursor:
            changed = False
            for idx, field_name in enumerate(valid_fields):
                if row[idx] is not None:
                    continue

                field_type = field_meta[field_name].type
                if field_type in numeric_types:
                    row[idx] = 0
                    changed = True
                elif field_type in text_types:
                    row[idx] = "0"
                    changed = True
                else:
                    # Fallback to "0" for unsupported/null-prone types.
                    row[idx] = "0"
                    changed = True

            if changed:
                cursor.updateRow(row)


def _export_input_features_with_fields(
    input_features: str,
    output_features: str,
    selected_fields: list[str],
) -> None:
    """Export input features keeping only selected attribute fields.

    This avoids schema conversion issues by excluding unsupported attributes
    and only carrying the fields needed for the join.
    """
    import arcpy  # type: ignore

    output_path = Path(output_features)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    all_fields = arcpy.ListFields(input_features)
    keep = set(selected_fields)

    field_info = arcpy.FieldInfo()
    for fld in all_fields:
        visible = "VISIBLE" if fld.name in keep else "HIDDEN"
        field_info.addField(fld.name, fld.name, visible, "NONE")

    tmp_layer = f"tmp_export_{uuid.uuid4().hex[:8]}"
    arcpy.management.MakeFeatureLayer(
        in_features=input_features,
        out_layer=tmp_layer,
        field_info=field_info,
    )
    try:
        arcpy.management.CopyFeatures(tmp_layer, str(output_path))
    finally:
        try:
            arcpy.management.Delete(tmp_layer)
        except Exception:
            pass

    dataset_exists = bool(arcpy.Exists(str(output_path)))
    file_exists = output_path.exists()
    if not dataset_exists and not file_exists:
        messages = arcpy.GetMessages()
        raise RuntimeError(
            "Output shapefile was not created. "
            f"Expected: {output_path}. ArcPy messages: {messages}"
        )


def _calculate_composite_key(
    table: str,
    field_1: str,
    field_2: str | None,
    key_field: str,
    case_sensitive: bool,
) -> None:
    import arcpy  # type: ignore

    def _norm(value: object) -> str:
        if value is None:
            return ""
        text = str(value).strip()
        if not case_sensitive:
            text = text.upper()
        return text

    fields = [field_1]
    if field_2:
        fields.append(field_2)
    fields.append(key_field)

    with arcpy.da.UpdateCursor(table, fields) as cursor:
        for row in cursor:
            first_value = row[0]
            second_value = row[1] if field_2 else None
            row[-1] = f"{_norm(first_value)}||{_norm(second_value)}"
            cursor.updateRow(row)


def _has_duplicate_keys(table: str, key_field: str) -> bool:
    import arcpy  # type: ignore

    seen = set()
    with arcpy.da.SearchCursor(table, [key_field]) as cursor:
        for (key_value,) in cursor:
            if key_value in seen:
                return True
            seen.add(key_value)
    return False


def _build_key_value(v1: object, v2: object, case_sensitive: bool) -> str:
    left = "" if v1 is None else str(v1).strip()
    right = "" if v2 is None else str(v2).strip()
    if not case_sensitive:
        left = left.upper()
        right = right.upper()
    return f"{left}||{right}"


def preview_join(args: argparse.Namespace) -> dict[str, int]:
    """Compute preview stats for a 2-field join without writing output."""
    import arcpy  # type: ignore

    setup_error = _validate_join_field_setup(args)
    if setup_error:
        raise ValueError(setup_error)

    use_second_field = bool((args.shp_field_2 or "").strip())

    _assert_field_exists(args.input_features, args.shp_field_1, "feature class")
    if use_second_field:
        _assert_field_exists(args.input_features, args.shp_field_2, "feature class")

    excel_table = _excel_to_table(args.excel_file, args.excel_sheet)
    _assert_field_exists(excel_table, args.excel_field_1, "Excel table")
    if use_second_field:
        _assert_field_exists(excel_table, args.excel_field_2, "Excel table")
    _assert_field_exists(excel_table, args.score_field, "Excel table")

    try:
        excel_keys = set()
        duplicate_excel_keys = 0
        excel_cursor_fields = [args.excel_field_1]
        if use_second_field:
            excel_cursor_fields.append(args.excel_field_2)
        with arcpy.da.SearchCursor(
            excel_table,
            excel_cursor_fields,
        ) as cursor:
            for row in cursor:
                v1 = row[0]
                v2 = row[1] if use_second_field else None
                key = _build_key_value(v1, v2, args.case_sensitive)
                if key in excel_keys:
                    duplicate_excel_keys += 1
                excel_keys.add(key)

        total_features = 0
        matched_features = 0
        unique_feature_keys = set()
        feature_cursor_fields = [args.shp_field_1]
        if use_second_field:
            feature_cursor_fields.append(args.shp_field_2)
        with arcpy.da.SearchCursor(
            args.input_features,
            feature_cursor_fields,
        ) as cursor:
            for row in cursor:
                v1 = row[0]
                v2 = row[1] if use_second_field else None
                total_features += 1
                key = _build_key_value(v1, v2, args.case_sensitive)
                unique_feature_keys.add(key)
                if key in excel_keys:
                    matched_features += 1

        return {
            "total_features": total_features,
            "unique_feature_keys": len(unique_feature_keys),
            "excel_rows": sum(1 for _ in arcpy.da.SearchCursor(excel_table, [args.excel_field_1])),
            "unique_excel_keys": len(excel_keys),
            "duplicate_excel_keys": duplicate_excel_keys,
            "matched_features": matched_features,
            "unmatched_features": total_features - matched_features,
        }
    finally:
        try:
            arcpy.management.Delete(excel_table)
        except Exception:
            pass


def _list_joinable_fields(table: str) -> list[str]:
    import arcpy  # type: ignore

    skip_types = {"OID", "Geometry", "Blob", "Raster"}
    return [
        fld.name
        for fld in arcpy.ListFields(table)
        if fld.type not in skip_types and not fld.required
    ]


def _excel_to_table(excel_file: str, excel_sheet: str | None) -> str:
    import arcpy  # type: ignore

    kwargs = {
        "Input_Excel_File": excel_file,
        "Output_Table": f"in_memory\\excel_scores_{uuid.uuid4().hex[:8]}",
    }
    if excel_sheet and excel_sheet.strip():
        kwargs["Sheet"] = excel_sheet.strip()
    return str(arcpy.conversion.ExcelToTable(**kwargs)[0])


def _read_excel_sheet_names(excel_file: str) -> list[str]:
    try:
        import pandas as pd  # type: ignore
    except Exception:
        return ["Sheet1"]

    try:
        xls = pd.ExcelFile(excel_file)
        return [str(s) for s in xls.sheet_names if str(s)]
    except Exception:
        return ["Sheet1"]


def _field_names(table: str) -> list[str]:
    import arcpy  # type: ignore

    return [f.name for f in arcpy.ListFields(table)]


def _resolve_field_name(table: str, requested_name: str) -> str | None:
    """Resolve a requested field name against actual table fields.

    Handles common differences like spaces vs underscores, case differences,
    and shapefile truncation behavior.
    """
    names = _field_names(table)
    if requested_name in names:
        return requested_name

    lower_lookup = {name.lower(): name for name in names}
    by_lower = lower_lookup.get(requested_name.lower())
    if by_lower:
        return by_lower

    requested_key = _field_name_key(requested_name)
    by_key = [name for name in names if _field_name_key(name) == requested_key]
    if len(by_key) == 1:
        return by_key[0]

    requested_shp = re.sub(r"\s+", "_", requested_name).upper()[:10]
    shp_exact = [name for name in names if name.upper() == requested_shp]
    if len(shp_exact) == 1:
        return shp_exact[0]

    shp_prefix = [name for name in names if name.upper().startswith(requested_shp[:8])]
    if len(shp_prefix) == 1:
        return shp_prefix[0]

    return None


def _require_resolved_field_name(table: str, requested_name: str, table_label: str) -> str:
    resolved = _resolve_field_name(table, requested_name)
    if resolved:
        return resolved

    available = ", ".join(_field_names(table))
    raise ValueError(
        f"Field '{requested_name}' not found in {table_label}: {table}. "
        f"Available fields: {available}"
    )


def _resolve_joined_score_field(
    output_features: str,
    requested_score_field: str,
    before_fields: set[str],
) -> str:
    """Resolve actual joined field name after JoinField.

    For shapefile outputs, ArcGIS may truncate field names to 10 chars.
    """
    after_fields = _field_names(output_features)
    if requested_score_field in after_fields:
        return requested_score_field

    new_fields = [name for name in after_fields if name not in before_fields]
    if len(new_fields) == 1:
        return new_fields[0]

    req_upper = requested_score_field.upper()
    truncated_upper = req_upper[:10]
    candidates = [
        name
        for name in new_fields
        if name.upper() == truncated_upper
        or name.upper().startswith(truncated_upper[:8])
        or name.upper().startswith(req_upper[:8])
    ]
    if len(candidates) == 1:
        return candidates[0]

    raise RuntimeError(
        "Could not determine joined score field name in output. "
        f"Requested: {requested_score_field}. New fields: {new_fields}"
    )


def _resolve_existing_output_score_field(
    output_features: str,
    requested_score_field: str,
) -> str | None:
    """Resolve score field name from an already-joined output feature class."""
    names = _field_names(output_features)
    if requested_score_field in names:
        return requested_score_field

    req_upper = requested_score_field.upper()
    truncated_upper = req_upper[:10]
    candidates = [
        name
        for name in names
        if name.upper() == truncated_upper
        or name.upper().startswith(truncated_upper[:8])
        or name.upper().startswith(req_upper[:8])
    ]
    if len(candidates) == 1:
        return candidates[0]
    return None


def _desired_output_score_field_name(input_score_field: str) -> str:
    if input_score_field.strip().upper().startswith("HSI"):
        return "HSI_min"
    return input_score_field


def _get_field(table: str, field_name: str):
    import arcpy  # type: ignore

    for fld in arcpy.ListFields(table):
        if fld.name == field_name:
            return fld
    return None



def _add_or_update_target_score_field(
    table: str,
    source_field: str,
    target_field: str,
) -> str:
    """Create/update target score field and copy values from source field."""
    import arcpy  # type: ignore

    if _field_name_key(source_field) == _field_name_key(target_field):
        resolved_same = _resolve_field_name(table, source_field)
        return resolved_same or source_field

    source = _get_field(table, source_field)
    if source is None:
        raise RuntimeError(f"Source score field not found: {source_field}")

    target = _get_field(table, target_field)
    if target is None:
        if source.type == "Integer":
            arcpy.management.AddField(table, target_field, "LONG")
        elif source.type == "SmallInteger":
            arcpy.management.AddField(table, target_field, "SHORT")
        elif source.type == "Double":
            arcpy.management.AddField(table, target_field, "DOUBLE")
        elif source.type == "Single":
            arcpy.management.AddField(table, target_field, "FLOAT")
        elif source.type == "Date":
            arcpy.management.AddField(table, target_field, "DATE")
        else:
            arcpy.management.AddField(
                table,
                target_field,
                "TEXT",
                field_length=max(1, source.length or 50),
            )

    resolved_target = _resolve_field_name(table, target_field) or target_field

    arcpy.management.CalculateField(
        in_table=table,
        field=resolved_target,
        expression=f"!{source_field}!",
        expression_type="PYTHON3",
    )

    source_now = _get_field(table, source_field)
    if (
        source_now is not None
        and not source_now.required
        and _field_name_key(source_field) != _field_name_key(resolved_target)
    ):
        try:
            arcpy.management.DeleteField(table, [source_field])
        except Exception:
            pass

    return resolved_target

    try:
        xls = pd.ExcelFile(excel_file)
        return [s for s in xls.sheet_names if s]
    except Exception:
        return ["Sheet1"]

    return ["Sheet1"]


def run_join(args: argparse.Namespace) -> int:
    try:
        import arcpy  # type: ignore
    except Exception as exc:
        print("ArcPy import failed.")
        print(f"Reason: {exc}")
        return 1

    arcpy.env.overwriteOutput = True

    input_features = args.input_features
    output_features = args.output_features
    excel_file = args.excel_file
    excel_sheet = args.excel_sheet
    setup_error = _validate_join_field_setup(args)
    if setup_error:
        print(f"Input error: {setup_error}")
        return 2

    use_second_field = bool((args.shp_field_2 or "").strip())

    print("Creating output features with selected key fields...")
    selected_input_fields = [args.shp_field_1]
    if use_second_field:
        selected_input_fields.append(args.shp_field_2)
    _export_input_features_with_fields(
        input_features=input_features,
        output_features=output_features,
        selected_fields=selected_input_fields,
    )

    excel_table = _excel_to_table(excel_file, excel_sheet)

    shp_field_1 = _require_resolved_field_name(
        output_features,
        args.shp_field_1,
        "feature class",
    )
    shp_field_2 = ""
    if use_second_field:
        shp_field_2 = _require_resolved_field_name(
            output_features,
            args.shp_field_2,
            "feature class",
        )
    excel_field_1 = _require_resolved_field_name(
        excel_table,
        args.excel_field_1,
        "Excel table",
    )
    excel_field_2 = ""
    if use_second_field:
        excel_field_2 = _require_resolved_field_name(
            excel_table,
            args.excel_field_2,
            "Excel table",
        )
    score_field = _require_resolved_field_name(
        excel_table,
        args.score_field,
        "Excel table",
    )

    # Normalize NULLs in selected fields before calculating keys and joining.
    input_fields_to_normalize = [shp_field_1]
    excel_fields_to_normalize = [excel_field_1, score_field]
    if use_second_field:
        input_fields_to_normalize.append(shp_field_2)
        excel_fields_to_normalize.append(excel_field_2)
    _normalize_nulls_to_zero(output_features, input_fields_to_normalize)
    _normalize_nulls_to_zero(excel_table, excel_fields_to_normalize)

    feature_key = "JOIN_KEY2"
    excel_key = "JOIN_KEY2"

    print("Building composite keys...")
    _add_text_field_if_missing(output_features, feature_key)
    _add_text_field_if_missing(excel_table, excel_key)
    _calculate_composite_key(
        table=output_features,
        field_1=shp_field_1,
        field_2=shp_field_2 if use_second_field else None,
        key_field=feature_key,
        case_sensitive=args.case_sensitive,
    )
    _calculate_composite_key(
        table=excel_table,
        field_1=excel_field_1,
        field_2=excel_field_2 if use_second_field else None,
        key_field=excel_key,
        case_sensitive=args.case_sensitive,
    )

    if _has_duplicate_keys(excel_table, excel_key):
        print(
            "Error: duplicate key combinations found in Excel table. "
            "Ensure each key pair is unique before joining."
        )
        return 2

    print("Joining score field to output features...")
    before_fields = set(_field_names(output_features))
    arcpy.management.JoinField(
        in_data=output_features,
        in_field=feature_key,
        join_table=excel_table,
        join_field=excel_key,
        fields=[score_field],
    )

    resolved_score_field = _resolve_joined_score_field(
        output_features=output_features,
        requested_score_field=score_field,
        before_fields=before_fields,
    )

    final_score_field = _add_or_update_target_score_field(
        table=output_features,
        source_field=resolved_score_field,
        target_field=_desired_output_score_field_name(score_field),
    )

    joined_count = 0
    with arcpy.da.SearchCursor(output_features, [final_score_field]) as cursor:
        for (value,) in cursor:
            if value is not None and str(value).strip() != "":
                joined_count += 1

    # Keep only the two selected key fields and the joined score field.
    keep_fields = {shp_field_1, final_score_field}
    if use_second_field:
        keep_fields.add(shp_field_2)
    if args.keep_join_key:
        keep_fields.add(feature_key)

    keep_keys = {_field_name_key(name) for name in keep_fields}

    removable_fields = [
        field.name
        for field in arcpy.ListFields(output_features)
        if not field.required and _field_name_key(field.name) not in keep_keys
    ]
    if removable_fields:
        arcpy.management.DeleteField(output_features, removable_fields)

    if not args.keep_join_key:
        try:
            arcpy.management.DeleteField(output_features, [feature_key])
        except Exception:
            pass

    print("Done.")
    print(f"Output: {output_features}")
    print(f"Joined score field in output: {final_score_field}")
    print(f"Records with joined score: {joined_count}")
    return 0


class JoinMultipleApp:
    """Tkinter GUI for composite-key joins between shapefile and Excel."""

    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("Join Score By 2 Fields")
        self.root.geometry("900x470")

        self.input_features_var = tk.StringVar()
        self.output_features_var = tk.StringVar()
        self.shp_field_1_var = tk.StringVar()
        self.shp_field_2_var = tk.StringVar()
        self.excel_file_var = tk.StringVar()
        self.excel_sheet_var = tk.StringVar(value="Sheet1")
        self.excel_field_1_var = tk.StringVar()
        self.excel_field_2_var = tk.StringVar()
        self.score_field_var = tk.StringVar()
        self.excel_display_to_actual: dict[str, str] = {}
        self.keep_join_key_var = tk.BooleanVar(value=False)
        self.case_sensitive_var = tk.BooleanVar(value=False)
        self.status_var = tk.StringVar(value="Step 1: select Excel. Step 2: select input shapefile.")
        self.output_score_field_var = tk.StringVar(value="Output score field: -")
        self.match_indicator_var = tk.StringVar(
            value="Field match: waiting for Excel and SHP field selections."
        )
        self.last_shp_folder = Path.cwd()
        self.last_excel_folder = Path.cwd()
        self.last_output_folder = DEFAULT_OUTPUT_FOLDER

        self._load_settings()
        self._build_ui()
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    def _to_excel_display_name(self, actual_name: str) -> str:
        return actual_name.replace("_", " ")

    def _to_excel_actual_name(self, display_name: str) -> str:
        if not display_name:
            return ""
        return self.excel_display_to_actual.get(display_name, display_name)

    def _set_excel_field_mapping(self, actual_fields: list[str]) -> list[str]:
        self.excel_display_to_actual = {}
        display_fields: list[str] = []
        used_display_names: set[str] = set()

        for actual in actual_fields:
            base_display = self._to_excel_display_name(actual)
            display = base_display
            suffix = 2
            # Keep display names unique for combobox values.
            while display in used_display_names:
                display = f"{base_display} ({suffix})"
                suffix += 1
            used_display_names.add(display)
            self.excel_display_to_actual[display] = actual
            display_fields.append(display)

        return display_fields

    def _load_settings(self) -> None:
        if not SETTINGS_FILE.exists():
            return

        try:
            settings = json.loads(SETTINGS_FILE.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError):
            return

        shp_folder = settings.get("last_shp_folder")
        excel_folder = settings.get("last_excel_folder")

        if isinstance(shp_folder, str) and shp_folder.strip():
            self.last_shp_folder = Path(shp_folder)
        if isinstance(excel_folder, str) and excel_folder.strip():
            self.last_excel_folder = Path(excel_folder)

    def _save_settings(self) -> None:
        settings = {
            "last_shp_folder": str(self.last_shp_folder),
            "last_excel_folder": str(self.last_excel_folder),
            "last_output_folder": str(self.last_output_folder),
        }
        try:
            SETTINGS_FILE.write_text(json.dumps(settings, indent=2), encoding="utf-8")
        except OSError:
            pass

    def _on_close(self) -> None:
        self._save_settings()
        self.root.destroy()

    def _build_ui(self) -> None:
        container = ttk.Frame(self.root, padding=12)
        container.pack(fill=tk.BOTH, expand=True)

        ttk.Label(container, text="Step 1: Excel source").pack(anchor=tk.W)

        excel_row = ttk.Frame(container)
        excel_row.pack(fill=tk.X, pady=(0, 8))
        ttk.Label(excel_row, text="Excel file:").pack(side=tk.LEFT)
        ttk.Entry(excel_row, textvariable=self.excel_file_var).pack(
            side=tk.LEFT, fill=tk.X, expand=True, padx=8
        )
        ttk.Button(
            excel_row,
            text="Browse...",
            command=self._select_excel_file,
        ).pack(side=tk.LEFT)

        sheet_row = ttk.Frame(container)
        sheet_row.pack(fill=tk.X, pady=(0, 8))
        ttk.Label(sheet_row, text="Excel sheet:").pack(side=tk.LEFT)
        self.excel_sheet_combo = ttk.Combobox(
            sheet_row,
            textvariable=self.excel_sheet_var,
            state="readonly",
            width=34,
        )
        self.excel_sheet_combo.pack(side=tk.LEFT, padx=8)
        self.excel_sheet_combo.bind("<<ComboboxSelected>>", self._on_excel_sheet_changed)

        excel_fields_row = ttk.Frame(container)
        excel_fields_row.pack(fill=tk.X, pady=(0, 8))
        ttk.Label(excel_fields_row, text="Excel field 1:").pack(side=tk.LEFT)
        self.excel_field_1_combo = ttk.Combobox(
            excel_fields_row,
            textvariable=self.excel_field_1_var,
            state="readonly",
            width=20,
        )
        self.excel_field_1_combo.pack(side=tk.LEFT, padx=(8, 12))
        self.excel_field_1_combo.bind(
            "<<ComboboxSelected>>",
            self._on_excel_key_fields_changed,
        )

        ttk.Label(excel_fields_row, text="Excel field 2:").pack(side=tk.LEFT)
        self.excel_field_2_combo = ttk.Combobox(
            excel_fields_row,
            textvariable=self.excel_field_2_var,
            state="readonly",
            width=20,
        )
        self.excel_field_2_combo.pack(side=tk.LEFT, padx=(8, 12))
        self.excel_field_2_combo.bind(
            "<<ComboboxSelected>>",
            self._on_excel_key_fields_changed,
        )

        ttk.Label(excel_fields_row, text="Score field:").pack(side=tk.LEFT)
        self.score_field_combo = ttk.Combobox(
            excel_fields_row,
            textvariable=self.score_field_var,
            state="readonly",
            width=20,
        )
        self.score_field_combo.pack(side=tk.LEFT, padx=8)

        ttk.Label(container, text="Step 2: Input shapefile").pack(anchor=tk.W)

        shp_row = ttk.Frame(container)
        shp_row.pack(fill=tk.X, pady=(0, 8))
        ttk.Label(shp_row, text="Input shapefile:").pack(side=tk.LEFT)
        ttk.Entry(shp_row, textvariable=self.input_features_var).pack(
            side=tk.LEFT, fill=tk.X, expand=True, padx=8
        )
        ttk.Button(
            shp_row,
            text="Browse...",
            command=self._select_input_features,
        ).pack(side=tk.LEFT)

        shp_fields_row = ttk.Frame(container)
        shp_fields_row.pack(fill=tk.X, pady=(0, 8))
        ttk.Label(shp_fields_row, text="SHP field 1:").pack(side=tk.LEFT)
        self.shp_field_1_combo = ttk.Combobox(
            shp_fields_row,
            textvariable=self.shp_field_1_var,
            state="readonly",
            width=24,
        )
        self.shp_field_1_combo.pack(side=tk.LEFT, padx=(8, 16))
        self.shp_field_1_combo.bind("<<ComboboxSelected>>", self._on_shp_fields_changed)
        ttk.Label(shp_fields_row, text="SHP field 2:").pack(side=tk.LEFT)
        self.shp_field_2_combo = ttk.Combobox(
            shp_fields_row,
            textvariable=self.shp_field_2_var,
            state="readonly",
            width=24,
        )
        self.shp_field_2_combo.pack(side=tk.LEFT, padx=8)
        self.shp_field_2_combo.bind("<<ComboboxSelected>>", self._on_shp_fields_changed)

        out_row = ttk.Frame(container)
        out_row.pack(fill=tk.X, pady=(0, 8))
        ttk.Label(out_row, text="Output shapefile:").pack(side=tk.LEFT)
        ttk.Entry(out_row, textvariable=self.output_features_var).pack(
            side=tk.LEFT, fill=tk.X, expand=True, padx=8
        )
        ttk.Button(
            out_row,
            text="Browse...",
            command=self._select_output_features,
        ).pack(side=tk.LEFT)

        options_row = ttk.Frame(container)
        options_row.pack(fill=tk.X, pady=(0, 8))
        ttk.Checkbutton(
            options_row,
            text="Keep temporary join key",
            variable=self.keep_join_key_var,
        ).pack(side=tk.LEFT)
        ttk.Checkbutton(
            options_row,
            text="Case sensitive matching",
            variable=self.case_sensitive_var,
        ).pack(side=tk.LEFT, padx=16)

        action_row = ttk.Frame(container)
        action_row.pack(fill=tk.X, pady=(8, 8))
        ttk.Button(
            action_row,
            text="Preview Match",
            command=self._preview_join,
        ).pack(side=tk.LEFT)
        ttk.Button(
            action_row,
            text="Run Join",
            command=self._run_join,
        ).pack(side=tk.LEFT, padx=(8, 0))

        ttk.Label(container, textvariable=self.output_score_field_var).pack(anchor=tk.W)
        ttk.Label(container, textvariable=self.match_indicator_var).pack(anchor=tk.W)
        ttk.Label(container, textvariable=self.status_var).pack(anchor=tk.W)

    def _select_input_features(self) -> None:
        path = filedialog.askopenfilename(
            title="Select input shapefile",
            initialdir=str(self.last_shp_folder),
            filetypes=[
                ("Shapefile", "*.shp"),
                ("All files", "*.*"),
            ],
        )
        if not path:
            return

        self.input_features_var.set(path)
        self.last_shp_folder = Path(path).parent
        self._save_settings()
        self._set_default_output_path(path)
        self._populate_shapefile_fields(path)

    def _select_output_features(self) -> None:
        default_name = "joined_HSI.shp"
        if self.input_features_var.get().strip():
            selected_input = self.input_features_var.get().strip()
            prefix = _extract_gemeente_prefix_from_input(selected_input)
            default_name = f"{prefix}{Path(selected_input).stem}_HSI.shp"

        path = filedialog.asksaveasfilename(
            title="Select output shapefile",
            defaultextension=".shp",
            initialdir=str(self.last_output_folder),
            initialfile=default_name,
            filetypes=[("Shapefile", "*.shp")],
        )
        if path:
            self.output_features_var.set(path)
            self.last_output_folder = Path(path).parent
            self._save_settings()

    def _select_excel_file(self) -> None:
        path = filedialog.askopenfilename(
            title="Select Excel file",
            initialdir=str(self.last_excel_folder),
            filetypes=[("Excel", "*.xlsx;*.xls"), ("All files", "*.*")],
        )
        if not path:
            return

        self.excel_file_var.set(path)
        self.last_excel_folder = Path(path).parent
        self._save_settings()
        self._populate_excel_sheets(path)
        self._populate_excel_fields()

    def _on_excel_sheet_changed(self, _event: tk.Event[tk.Misc] | None = None) -> None:
        self._populate_excel_fields()

    def _set_default_output_path(self, input_path: str) -> None:
        prefix = _extract_gemeente_prefix_from_input(input_path)
        output_folder = self.last_output_folder if self.last_output_folder else DEFAULT_OUTPUT_FOLDER
        self.output_features_var.set(
            str(output_folder / f"{prefix}{Path(input_path).stem}_HSI.shp")
        )
        self.last_output_folder = output_folder
        self._save_settings()

    def _populate_shapefile_fields(self, feature_path: str) -> None:
        try:
            fields = _list_joinable_fields(feature_path)
        except Exception as exc:
            self.status_var.set(f"Could not read shapefile fields: {exc}")
            return

        self.shp_field_1_combo["values"] = fields
        self.shp_field_2_combo["values"] = [""] + fields
        if fields:
            self.shp_field_1_var.set(fields[0])
            self.shp_field_2_var.set(fields[1] if len(fields) > 1 else fields[0])
        self._sync_shp_fields_from_excel()
        self._update_match_indicator()

    def _populate_excel_sheets(self, excel_file: str) -> None:
        sheets = _read_excel_sheet_names(excel_file)
        self.excel_sheet_combo["values"] = sheets
        if sheets:
            self.excel_sheet_var.set(sheets[0])

    def _populate_excel_fields(self) -> None:
        excel_file = self.excel_file_var.get().strip()
        if not excel_file:
            return

        sheet = self.excel_sheet_var.get().strip()
        temp_table = ""
        arcpy_mod = None
        try:
            import arcpy  # type: ignore

            arcpy_mod = arcpy
            temp_table = _excel_to_table(excel_file, sheet)
            actual_fields = _list_joinable_fields(temp_table)
        except Exception as exc:
            self.status_var.set(f"Could not read Excel fields: {exc}")
            return
        finally:
            if temp_table and arcpy_mod is not None:
                try:
                    arcpy_mod.management.Delete(temp_table)
                except Exception:
                    pass

        fields = self._set_excel_field_mapping(actual_fields)

        self.excel_field_1_combo["values"] = fields
        self.excel_field_2_combo["values"] = [""] + fields
        self.score_field_combo["values"] = fields

        if actual_fields:
            # Requested default behavior:
            # - 4 fields: use 1st + 2nd as join keys, 4th as score
            # - 3 fields: use 1st as join key, no 2nd join key, 3rd as score
            self.excel_field_1_var.set(self._to_excel_display_name(actual_fields[0]))

            if len(actual_fields) >= 4:
                self.excel_field_2_var.set(self._to_excel_display_name(actual_fields[1]))
                self.score_field_var.set(self._to_excel_display_name(actual_fields[3]))
            elif len(actual_fields) == 3:
                self.excel_field_2_var.set("")
                self.score_field_var.set(self._to_excel_display_name(actual_fields[2]))
            elif len(actual_fields) == 2:
                self.excel_field_2_var.set(self._to_excel_display_name(actual_fields[1]))
                self.score_field_var.set(self._to_excel_display_name(actual_fields[1]))
            else:
                self.excel_field_2_var.set("")
                self.score_field_var.set(self._to_excel_display_name(actual_fields[0]))

        self._sync_shp_fields_from_excel()
        self._update_match_indicator()

    def _on_excel_key_fields_changed(self, _event: tk.Event[tk.Misc] | None = None) -> None:
        self._sync_shp_fields_from_excel()
        self._update_match_indicator()

    def _on_shp_fields_changed(self, _event: tk.Event[tk.Misc] | None = None) -> None:
        self._update_match_indicator()

    def _update_match_indicator(self) -> None:
        excel1 = self._to_excel_actual_name(self.excel_field_1_var.get().strip())
        excel2 = self._to_excel_actual_name(self.excel_field_2_var.get().strip())
        shp1 = self.shp_field_1_var.get().strip()
        shp2 = self.shp_field_2_var.get().strip()

        if not (excel1 or excel2 or shp1 or shp2):
            self.match_indicator_var.set(
                "Field match: waiting for Excel and SHP field selections."
            )
            return

        def _status(excel_name: str, shp_name: str) -> str:
            if not excel_name:
                return "Excel not selected"
            if not shp_name:
                return "SHP not selected"
            if _field_name_key(excel_name) == _field_name_key(shp_name):
                return "matched"
            return f"mismatch ({excel_name} -> {shp_name})"

        s1 = _status(excel1, shp1)
        if not excel2 and not shp2:
            s2 = "not used"
        else:
            s2 = _status(excel2, shp2)
        self.match_indicator_var.set(f"Field match: 1={s1}; 2={s2}")

    def _sync_shp_fields_from_excel(self) -> None:
        shp_values = [str(v) for v in self.shp_field_1_combo["values"]]
        if not shp_values:
            return

        shp_lookup = {name.lower(): name for name in shp_values if name}
        shp_key_lookup = {_field_name_key(name): name for name in shp_values if name}
        excel1 = self._to_excel_actual_name(self.excel_field_1_var.get().strip())
        excel2 = self._to_excel_actual_name(self.excel_field_2_var.get().strip())

        mapped_1 = None
        mapped_2 = None
        if excel1:
            mapped_1 = shp_lookup.get(excel1.lower())
            if mapped_1 is None:
                mapped_1 = shp_key_lookup.get(_field_name_key(excel1))
        if excel2:
            mapped_2 = shp_lookup.get(excel2.lower())
            if mapped_2 is None:
                mapped_2 = shp_key_lookup.get(_field_name_key(excel2))

        if mapped_1:
            self.shp_field_1_var.set(mapped_1)
        if excel2 and mapped_2:
            self.shp_field_2_var.set(mapped_2)
        if not excel2:
            self.shp_field_2_var.set("")
        self._update_match_indicator()

    def _run_join(self) -> None:
        args = argparse.Namespace(
            input_features=self.input_features_var.get().strip(),
            output_features=self.output_features_var.get().strip(),
            shp_field_1=self.shp_field_1_var.get().strip(),
            shp_field_2=self.shp_field_2_var.get().strip(),
            excel_file=self.excel_file_var.get().strip(),
            excel_sheet=self.excel_sheet_var.get().strip(),
            excel_field_1=self._to_excel_actual_name(self.excel_field_1_var.get().strip()),
            excel_field_2=self._to_excel_actual_name(self.excel_field_2_var.get().strip()),
            score_field=self._to_excel_actual_name(self.score_field_var.get().strip()),
            keep_join_key=self.keep_join_key_var.get(),
            case_sensitive=self.case_sensitive_var.get(),
        )

        missing = _required_cli_args(args)
        if missing:
            messagebox.showerror(
                "Missing input",
                "Please fill all required inputs before running the join.\n"
                + "Missing: "
                + ", ".join(missing),
            )
            return

        setup_error = _validate_join_field_setup(args)
        if setup_error:
            messagebox.showerror("Invalid join fields", setup_error)
            return

        self.status_var.set("Running join...")
        self.output_score_field_var.set("Output score field: -")
        self.root.update_idletasks()
        try:
            code = run_join(args)
        except Exception as exc:
            self.status_var.set("Join failed. Check error details.")
            messagebox.showerror("Join failed", str(exc))
            return

        if code == 0:
            resolved = _resolve_existing_output_score_field(
                args.output_features,
                _desired_output_score_field_name(args.score_field),
            )
            if resolved:
                self.output_score_field_var.set(f"Output score field: {resolved}")
            else:
                self.output_score_field_var.set("Output score field: (could not resolve)")
            self.status_var.set("Join completed successfully.")
            messagebox.showinfo("Success", "Join completed successfully.")
            return

        self.status_var.set("Join failed. Check console output for details.")
        messagebox.showerror(
            "Join failed",
            "Join failed. See terminal output for details.",
        )

    def _preview_join(self) -> None:
        args = argparse.Namespace(
            input_features=self.input_features_var.get().strip(),
            output_features=self.output_features_var.get().strip() or "unused_for_preview",
            shp_field_1=self.shp_field_1_var.get().strip(),
            shp_field_2=self.shp_field_2_var.get().strip(),
            excel_file=self.excel_file_var.get().strip(),
            excel_sheet=self.excel_sheet_var.get().strip(),
            excel_field_1=self._to_excel_actual_name(self.excel_field_1_var.get().strip()),
            excel_field_2=self._to_excel_actual_name(self.excel_field_2_var.get().strip()),
            score_field=self._to_excel_actual_name(self.score_field_var.get().strip()),
            keep_join_key=self.keep_join_key_var.get(),
            case_sensitive=self.case_sensitive_var.get(),
        )

        missing = [
            name
            for name in [
                "input_features",
                "shp_field_1",
                "excel_file",
                "excel_sheet",
                "excel_field_1",
                "score_field",
            ]
            if not getattr(args, name, None)
        ]
        if missing:
            messagebox.showerror(
                "Missing input",
                "Please fill all required preview inputs first.\nMissing: "
                + ", ".join(missing),
            )
            return

        setup_error = _validate_join_field_setup(args)
        if setup_error:
            messagebox.showerror("Invalid join fields", setup_error)
            return

        self.status_var.set("Calculating preview...")
        self.root.update_idletasks()

        try:
            stats = preview_join(args)
        except Exception as exc:
            self.status_var.set("Preview failed.")
            messagebox.showerror("Preview failed", str(exc))
            return

        match_pct = 0.0
        if stats["total_features"] > 0:
            match_pct = (stats["matched_features"] / stats["total_features"]) * 100.0

        self.status_var.set(
            f"Preview ready: {stats['matched_features']}/{stats['total_features']} matched ({match_pct:.1f}%)."
        )

        message = (
            "Preview results\n\n"
            f"Total features: {stats['total_features']}\n"
            f"Unique feature key pairs: {stats['unique_feature_keys']}\n"
            f"Excel rows: {stats['excel_rows']}\n"
            f"Unique Excel key pairs: {stats['unique_excel_keys']}\n"
            f"Duplicate Excel key pairs: {stats['duplicate_excel_keys']}\n"
            f"Matched features: {stats['matched_features']}\n"
            f"Unmatched features: {stats['unmatched_features']}\n"
            f"Match rate: {match_pct:.1f}%"
        )
        if stats["duplicate_excel_keys"] > 0:
            message += (
                "\n\nWarning: duplicate key pairs found in Excel. "
                "Join may be ambiguous; make key pairs unique first."
            )
        messagebox.showinfo("Join Preview", message)


def launch_gui() -> int:
    root = tk.Tk()
    JoinMultipleApp(root)
    root.mainloop()
    return 0


def main() -> int:
    args = build_parser().parse_args()

    # VS Code Run button starts the file without arguments.
    # In that case, open GUI mode by default.
    if args.gui or not _has_any_cli_join_arg(args):
        return launch_gui()

    missing = _required_cli_args(args)
    if missing:
        print(
            "Missing required arguments for CLI mode: "
            + ", ".join(missing)
            + ". Use --gui to launch the interface."
        )
        return 2

    setup_error = _validate_join_field_setup(args)
    if setup_error:
        print(f"Input error: {setup_error}")
        return 2

    try:
        return run_join(args)
    except ValueError as exc:
        print(f"Input error: {exc}")
        return 2
    except Exception as exc:
        print(f"Unhandled error: {exc}")
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
