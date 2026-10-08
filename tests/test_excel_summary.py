from pathlib import Path

from scripts.excel_summary import _default_output_path


def test_default_output_path_uses_first_input_folder_for_multiple_folders() -> None:
    input_files = [
        Path("data/first/input.xlsx"),
        Path("other/second/input.xlsx"),
    ]

    output_path = _default_output_path(input_files)

    assert output_path.parent == input_files[0].parent
