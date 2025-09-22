from pathlib import Path
import sys


ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import pytest
from openpyxl import Workbook, load_workbook

from merge_excel_files import merge_excel_files


def _create_workbook(path: Path, rows: list[list[object]]) -> None:
    workbook = Workbook()
    sheet = workbook.active
    for row in rows:
        sheet.append(row)
    workbook.save(path)


def test_merge_excel_files_creates_sheet_per_file(tmp_path: Path) -> None:
    first = tmp_path / "alpha.xlsx"
    second = tmp_path / "bravo.xlsx"
    _create_workbook(first, [["header", "value"], ["row", 1]])
    _create_workbook(second, [["only", "row"]])

    output = tmp_path / "combined.xlsx"
    merged = merge_excel_files(tmp_path, output)

    assert [entry.sheet_name for entry in merged] == ["alpha", "bravo"]

    workbook = load_workbook(output)
    try:
        assert workbook.sheetnames == ["alpha", "bravo"]
        alpha_rows = [tuple(row) for row in workbook["alpha"].iter_rows(values_only=True)]
        bravo_rows = [tuple(row) for row in workbook["bravo"].iter_rows(values_only=True)]
    finally:
        workbook.close()

    assert alpha_rows == [("header", "value"), ("row", 1)]
    assert bravo_rows == [("only", "row")]


def test_merge_excel_files_handles_duplicate_names(tmp_path: Path) -> None:
    long_name = "a" * 40
    first = tmp_path / f"{long_name}1.xlsx"
    second = tmp_path / f"{long_name}2.xlsx"
    _create_workbook(first, [[1]])
    _create_workbook(second, [[2]])

    output = tmp_path / "out.xlsx"
    merged = merge_excel_files(tmp_path, output)

    sheet_names = [entry.sheet_name for entry in merged]
    assert sheet_names[0] == "a" * 31
    assert sheet_names[1].startswith("a" * 27)
    assert sheet_names[1].endswith("_2")

    workbook = load_workbook(output)
    try:
        first_value = workbook[sheet_names[0]].cell(1, 1).value
        second_value = workbook[sheet_names[1]].cell(1, 1).value
    finally:
        workbook.close()

    assert first_value == 1
    assert second_value == 2


def test_merge_excel_files_fails_when_no_files(tmp_path: Path) -> None:
    output = tmp_path / "merged.xlsx"
    with pytest.raises(ValueError):
        merge_excel_files(tmp_path, output)