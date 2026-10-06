"""Set HSI_max = 1 for rows where id1 or id2 contains grass/mowing terms.

Rules:
- HSI_max is set to 1 when id1 or id2 contains any of the INCLUDE_TERMS.
- HSI_max is NOT set (or cleared if already 1) when id1 or id2 contains any of the EXCLUDE_TERMS.
- EXCLUDE_TERMS take priority over INCLUDE_TERMS.
"""

from __future__ import annotations

import argparse
from collections import defaultdict
import datetime
import logging
from pathlib import Path
import re
from typing import Any

import arcpy
from openpyxl import load_workbook


DEFAULT_SUMMARY_PATH = Path(r"E:/2026/ZoemZoneLimburg/HSI_max/summary_20260721_145240_metHSI_max.xlsx")
DEFAULT_SHAPEFILE_FOLDER = Path(r"E:/2026/ZoemZoneLimburg/HSI")

INCLUDE_TERMS = ("berm", "gras", "gazon", "maai")
EXCLUDE_TERMS = ("kunstgras", "sport", "intensief")

HEADER_ROW = 2

SHAPEFILE_OVERRIDES = {
    "Eijsden-Margraten_EM groen, Lijnvormige elementen_Element_summary.xlsx": "Eijsden-Margraten_EM groen_LijnEl_HSI.shp",
    "Eijsden-Margraten_EM groen, Vlakvormige elementen_Element_summary.xlsx": "Eijsden-Margraten_EM groen_VlakEl_HSI.shp",
    "Gulpen-Wittem_beplanting_objecttype_summary.xlsx": "Gulpen_Beplanting_HSI.shp",
    "Gulpen-Wittem_Gulpen_Gras_objecttype_x_maaifreque_summary.xlsx": "Gulpen_Gras_HSI.shp",
    "Gulpen-Wittem_haag_objecttype_summary.xlsx": "Gulpen_Haag_HSI.shp",
    "Horst aan de Maas_Bermtypen_GRASTYPE_summary.xlsx": "Horst_Bermtypen_HSI.shp",
    "Horst aan de Maas_Faunabermen_bermtype_2_summary.xlsx": "Horst_Faunabermen_HSI.shp",
    "Horst aan de Maas_groenbeheer_vlakobjecten_Type_gedet_summary.xlsx": "Horst_groenbeheer_vlakobjecten_HSI.shp",
    "Leudal_ul_gi_gras_kruidachtigen_6507e7e7_f527_40bf_99a5_a9362a133e9f_v_Maairegime_1x_summary.xlsx": "Leuldal_gras_kruidachtigen_1xmaai_HSI.shp",
    "Leudal_ul_gi_gras_kruidachtigen_a408ee7e_f87f_43f1_9eb0_8452cf9783b2_v_Maairegime_2x_summary.xlsx": "Leuldal_gras_kruidachtigen_2xmaai_HSI.shp",
    "Roermond_gb_gras_v_Gras_categ_x_Maairegime_summary.xlsx": "Roermond_gb_gras_v_HSI.shp",
    "Meerssen_grassen_BEHEERGROE_summary.xlsx": "Meersen_grassen_HSI.shp",
    "Peel en Maas 07042026_Haag_inrichtendV_CT_7IMGEO__summary.xlsx": "PeelEnMaas_Haag_inrichtendV_HSI.shp",
    "Peel&Maas_groenobject_vlakobject_export_TYPE_GEDET_summary.xlsx": "PeelEnMaas_groenobject_vlakobject_export_HSI.shp",
    "Roerdaelen_BGT_INBEHEER_BGT_Vlakken_InBeheer_TYPERING_x_TYPE_BOR1_summary.xlsx": "Roerdaelen_BGT_Vlakken_InBeheer_HSI.shp",
    "Sittard-Geleen_Sittard_Beplantingen_BEHEERGROE_summary.xlsx": "Sittard_Beplantingen_HSI.shp",
    "Sittard-Geleen_grassen_BEHEERGROE_summary.xlsx": "Sittard_grassen_HSI.shp",
    "Sittard-Geleen_hagen_polygons_BEHEERGROE_summary.xlsx": "Sittard_hagen_polygons_HSI.shp",
    "Vaals_Gras_objecttype_summary.xlsx": "Vaals_Gras_HSI.shp",
}

LOGGER = logging.getLogger("zoemzone.hsi_max_grassen")


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Set HSI_max = 1 for grass/mowing rows in a summary workbook.",
    )
    parser.add_argument(
        "--summary",
        default=str(DEFAULT_SUMMARY_PATH),
        help="Path to the summary workbook with an HSI_max column.",
    )
    parser.add_argument(
        "--shapefile-folder",
        default=str(DEFAULT_SHAPEFILE_FOLDER),
        help="Folder containing the original shapefiles to update.",
    )
    parser.add_argument(
        "--log",
        default=None,
        help="Optional log file. Defaults to HSI_updates_<timestamp>.log beside the summary.",
    )
    return parser


def _cell_to_text(value: object) -> str:
    return "" if value is None else str(value).strip()


def _value_key(value: object) -> str:
    """Normalize values, including order-independent semicolon lists."""
    text = _cell_to_text(value).casefold()
    parts = [part.strip() for part in text.split(";") if part.strip()]
    return ";".join(sorted(parts))


def _filename_key(path: str) -> str:
    stem = Path(path).stem.casefold()
    stem = re.sub(r"_(summary|layer)$", "", stem)
    stem = re.sub(r"_hsi$", "", stem)
    return re.sub(r"[^a-z0-9]+", "", stem)


def _read_summary_rows(summary_path: Path) -> list[tuple[str, str, Any, str]]:
    wb = load_workbook(filename=summary_path)
    ws = wb.active
    if ws is None:
        raise RuntimeError("No active worksheet found in workbook.")

    headers = {
        _cell_to_text(ws.cell(row=HEADER_ROW, column=index).value).casefold(): index
        for index in range(1, ws.max_column + 1)
    }
    required = {"id1", "id2", "hsi_max", "source_excel"}
    missing = required - headers.keys()
    if missing:
        raise ValueError(f"Summary is missing columns: {', '.join(sorted(missing))}")

    rows = []
    for row_index in range(HEADER_ROW + 1, ws.max_row + 1):
        source_excel = _cell_to_text(ws.cell(row_index, headers["source_excel"]).value)
        if not source_excel:
            continue
        rows.append(
            (
                _cell_to_text(ws.cell(row_index, headers["id1"]).value),
                _cell_to_text(ws.cell(row_index, headers["id2"]).value),
                ws.cell(row_index, headers["hsi_max"]).value,
                source_excel,
            )
        )
    wb.close()
    return rows


def update_hsi_max(summary_path: Path) -> None:
    wb = load_workbook(filename=summary_path)
    ws = wb.active
    if ws is None:
        raise RuntimeError("No active worksheet found in workbook.")

    set_to_1 = 0
    cleared = 0
    copied_from_hsi_min = 0

    for row_index in range(HEADER_ROW + 1, ws.max_row + 1):
        id1 = str(ws.cell(row_index, 1).value or "").lower()
        id2 = str(ws.cell(row_index, 2).value or "").lower()
        combined = id1 + " " + id2

        hsi_cell: Any = ws.cell(row_index, 4)
        old_value = hsi_cell.value
        if any(term in combined for term in EXCLUDE_TERMS):
            if hsi_cell.value == 1:
                hsi_cell.value = None
                cleared += 1
        elif any(term in combined for term in INCLUDE_TERMS):
            hsi_cell: Any = ws.cell(row_index, 4)
            if hsi_cell.value != 1:
                hsi_cell.value = 1
                set_to_1 += 1

        hsi_cell: Any = ws.cell(row_index, 4)
        if hsi_cell.value in (None, ""):
            hsi_min_value = ws.cell(row_index, 3).value
            if hsi_min_value not in (None, ""):
                hsi_cell.value = hsi_min_value
                copied_from_hsi_min += 1

        if hsi_cell.value != old_value:
            LOGGER.info(
                "SUMMARY row=%s id1=%r id2=%r HSI_max: %r -> %r",
                row_index,
                id1,
                id2,
                old_value,
                hsi_cell.value,
            )

    wb.save(summary_path)
    print(
        f"Done. {set_to_1} rows set to HSI_max=1, {cleared} rows cleared "
        f"(excluded), {copied_from_hsi_min} rows copied from HSI_min."
    )


def _resolve_shapefile(source_excel: str, shapefile_folder: Path) -> Path:
    override = SHAPEFILE_OVERRIDES.get(Path(source_excel).name)
    if override:
        path = shapefile_folder / override
        if path.is_file():
            return path

    source_stem = Path(source_excel).stem
    if source_stem.casefold().endswith("_summary"):
        source_stem = source_stem[:-len("_summary")]
    source_key = _filename_key(source_stem)
    source_compact_key = re.sub(r"[^a-z0-9]+", "", source_stem.casefold())
    source_compact_key = re.sub(r"(?:summary|layer)$", "", source_compact_key)

    candidates = []
    for path in shapefile_folder.glob("*.shp"):
        shapefile_stem = path.stem
        if shapefile_stem.casefold().endswith("_hsi"):
            shapefile_stem = shapefile_stem[:-len("_hsi")]
        shapefile_key = _filename_key(shapefile_stem)
        shapefile_compact_key = re.sub(r"[^a-z0-9]+", "", shapefile_stem.casefold())
        if source_key.startswith(shapefile_key) or source_compact_key.startswith(shapefile_compact_key):
            candidates.append(path)

    # Prefer the longest prefix so similarly named source layers resolve correctly.
    candidates.sort(key=lambda path: len(_filename_key(path.stem)), reverse=True)
    if len(candidates) == 1:
        return candidates[0]
    if candidates and len(_filename_key(candidates[0].stem)) > len(_filename_key(candidates[1].stem)):
        return candidates[0]
    if not candidates:
        raise FileNotFoundError(f"No shapefile match for summary source: {source_excel}")
    raise RuntimeError(f"Multiple equally specific shapefile matches for {source_excel}: {candidates}")


def _update_shapefile(shapefile: Path, rows: list[tuple[str, str, Any, str]]) -> tuple[int, int]:
    fields = {field.name: field for field in arcpy.ListFields(str(shapefile))}
    if "HSI_max" not in fields:
        arcpy.management.AddField(str(shapefile), "HSI_max", "DOUBLE")
        LOGGER.info("SHAPEFILE %s added field HSI_max type=Double", shapefile)
    elif fields["HSI_max"].type != "Double":
        raise TypeError(f"Field HSI_max in {shapefile} is not type Double")

    value_map: dict[tuple[str, str], float | None] = {}
    for id1, id2, hsi_max, _ in rows:
        key = (_value_key(id1), _value_key(id2))
        if hsi_max not in (None, ""):
            value_map[key] = float(hsi_max)
        else:
            value_map.setdefault(key, None)

    source_fields = [
        field.name
        for field in arcpy.ListFields(str(shapefile))
        if field.type not in {"OID", "Geometry"} and field.name not in {"HSI_min", "HSI_max"}
    ]
    updated = 0
    unmatched = 0
    cursor_fields = source_fields + ["HSI_min", "HSI_max"]
    with arcpy.da.UpdateCursor(str(shapefile), cursor_fields) as cursor:
        for feature_index, row in enumerate(cursor, start=1):
            attribute_keys = {_value_key(value) for value in row[:-2] if _cell_to_text(value)}
            matches = [
                key
                for key in value_map
                if key[0] in attribute_keys and (not key[1] or key[1] in attribute_keys)
            ]
            if matches:
                matching_values: list[float] = [
                    value for key in matches
                    if (value := value_map[key]) is not None
                ]
                new_value = max(matching_values) if matching_values else row[-2]
                if new_value is None:
                    unmatched += 1
                    continue
                if row[-1] != new_value:
                    old_value = row[-1]
                    row[-1] = float(new_value)
                    cursor.updateRow(row)
                    updated += 1
                    LOGGER.info(
                        "SHAPEFILE %s feature=%s HSI_max: %r -> %r",
                        shapefile,
                        feature_index,
                        old_value,
                        new_value,
                    )
            else:
                unmatched += 1
    return updated, unmatched


def update_shapefiles(summary_path: Path, shapefile_folder: Path) -> None:
    rows_by_source: dict[str, list[tuple[str, str, Any, str]]] = defaultdict(list)
    for row in _read_summary_rows(summary_path):
        rows_by_source[row[3]].append(row)

    total_updated = 0
    total_unmatched = 0
    for source_excel, rows in sorted(rows_by_source.items()):
        shapefile = _resolve_shapefile(source_excel, shapefile_folder)
        updated, unmatched = _update_shapefile(shapefile, rows)
        total_updated += updated
        total_unmatched += unmatched
        print(f"{shapefile.name}: {updated} features updated, {unmatched} unmatched")

    print(f"Shapefiles complete. {total_updated} features updated, {total_unmatched} unmatched.")


def main() -> None:
    args = build_parser().parse_args()
    summary_path = Path(args.summary)
    if not summary_path.is_file():
        raise FileNotFoundError(f"Summary file not found: {summary_path}")

    log_path = Path(args.log) if args.log else summary_path.parent / (
        f"HSI_updates_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
    )
    log_path.parent.mkdir(parents=True, exist_ok=True)
    logging.basicConfig(
        filename=log_path,
        level=logging.INFO,
        format="%(asctime)s %(levelname)s %(message)s",
        encoding="utf-8",
    )
    LOGGER.info("START summary=%s shapefile_folder=%s", summary_path, args.shapefile_folder)
    update_hsi_max(summary_path)
    update_shapefiles(summary_path, Path(args.shapefile_folder))
    LOGGER.info("END")
    print(f"Log written to: {log_path}")


if __name__ == "__main__":
    main()
