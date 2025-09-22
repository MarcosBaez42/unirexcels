"""Utility to merge multiple Excel workbooks into a single file.

This module provides a command line interface and a callable helper
function that copies the first worksheet from every Excel file in a
folder into a consolidated workbook.  Each worksheet in the output is
named after the corresponding source file (without its extension).
"""
from __future__ import annotations

import argparse
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Iterator, List, Sequence

from openpyxl import Workbook, load_workbook
from openpyxl.worksheet.worksheet import Worksheet

MAX_SHEET_NAME_LENGTH = 31
_INVALID_SHEET_CHARS = set("[]:*?/\\")


@dataclass(frozen=True)
class MergedSheet:
    """Represents the relationship between a source file and a sheet name."""

    source: Path
    sheet_name: str


def merge_excel_files(
    source_directory: Path,
    output_path: Path,
    pattern: str = "*.xlsx",
    recursive: bool = False,
    values_only: bool = False,
) -> List[MergedSheet]:
    """Merge Excel workbooks stored inside *source_directory* into one file.

    Parameters
    ----------
    source_directory:
        Folder that contains the files to merge.
    output_path:
        Path where the consolidated workbook should be written.  The parent
        directory will be created automatically if it does not exist.
    pattern:
        Glob expression used to filter the files inside ``source_directory``.
        Defaults to ``"*.xlsx"`` which will match the most common Excel
        workbooks.  The comparison is case-sensitive.
    recursive:
        When ``True`` the pattern is applied recursively.
    values_only:
        When ``True`` formulas in the source sheets are replaced by their
        last cached value.  When ``False`` formulas are preserved.

    Returns
    -------
    List[MergedSheet]
        A list describing which files were merged and the sheet names used
        in the output workbook.  The sheets are ordered according to the
        file names used for the merge.

    Raises
    ------
    FileNotFoundError
        If ``source_directory`` does not exist.
    ValueError
        If no Excel files matching ``pattern`` are found.
    """

    if not source_directory.exists():
        raise FileNotFoundError(f"Source directory '{source_directory}' does not exist")
    if not source_directory.is_dir():
        raise NotADirectoryError(f"'{source_directory}' is not a directory")

    resolved_output = output_path.resolve()

    candidates = _collect_excel_files(source_directory, pattern, recursive)
    files = [path for path in candidates if path.resolve() != resolved_output]

    if not files:
        raise ValueError(
            f"No Excel files matching pattern '{pattern}' were found in '{source_directory}'"
        )

    workbook = Workbook()
    # Remove the default sheet created by openpyxl; it will be replaced with
    # the sheets copied from the source files.
    if workbook.sheetnames:
        workbook.remove(workbook.active)

    existing_names: set[str] = set()
    merged_sheets: List[MergedSheet] = []

    for file_path in files:
        sheet_title = _build_sheet_title(file_path.stem, existing_names)
        source_wb = load_workbook(filename=file_path, data_only=values_only)
        try:
            source_sheet = source_wb.worksheets[0]
            target_sheet = workbook.create_sheet(title=sheet_title)
            _copy_sheet_contents(source_sheet, target_sheet)
        finally:
            source_wb.close()
        merged_sheets.append(MergedSheet(source=file_path, sheet_name=sheet_title))

    output_path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(output_path)

    return merged_sheets


def _collect_excel_files(directory: Path, pattern: str, recursive: bool) -> List[Path]:
    search_method = directory.rglob if recursive else directory.glob
    files = sorted(path for path in search_method(pattern) if path.is_file())
    return files


def _build_sheet_title(source_stem: str, existing_names: set[str]) -> str:
    base_name = _sanitize_sheet_base(source_stem)
    for candidate in _generate_sheet_name_candidates(base_name):
        if candidate not in existing_names:
            existing_names.add(candidate)
            return candidate
    raise RuntimeError("Unable to determine a unique sheet name")


def _sanitize_sheet_base(raw_name: str) -> str:
    sanitized = "".join(
        "_" if ch in _INVALID_SHEET_CHARS or ord(ch) < 32 else ch for ch in raw_name
    ).strip()
    sanitized = sanitized.rstrip("'")
    return sanitized or "Sheet"


def _generate_sheet_name_candidates(base_name: str) -> Iterator[str]:
    trimmed = base_name[:MAX_SHEET_NAME_LENGTH]
    if trimmed:
        yield trimmed
    else:
        yield "Sheet"

    counter = 2
    while True:
        suffix = f"_{counter}"
        trimmed_length = MAX_SHEET_NAME_LENGTH - len(suffix)
        prefix = base_name[:trimmed_length].rstrip()
        if not prefix:
            prefix = base_name[:trimmed_length]
        if not prefix:
            prefix = "Sheet"[:trimmed_length]
        candidate = f"{prefix}{suffix}"[:MAX_SHEET_NAME_LENGTH]
        if candidate:
            yield candidate
        counter += 1


def _copy_sheet_contents(source: Worksheet, target: Worksheet) -> None:
    if source.max_row == 1 and source.max_column == 1 and source.cell(1, 1).value is None:
        # Leave the sheet empty if the source is empty.
        return

    for row in source.iter_rows():
        target.append([cell.value for cell in row])

    for merged_range in source.merged_cells.ranges:
        target.merge_cells(str(merged_range))


def _parse_args(args: Sequence[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Merge multiple Excel workbooks into a single file where each file "
            "becomes its own worksheet."
        )
    )
    parser.add_argument(
        "source_directory",
        type=Path,
        help="Folder containing the Excel files to merge",
    )
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        default=Path("combined.xlsx"),
        help="Path to the consolidated Excel workbook",
    )
    parser.add_argument(
        "-p",
        "--pattern",
        default="*.xlsx",
        help="Glob pattern used to filter Excel files (default: '*.xlsx')",
    )
    parser.add_argument(
        "-r",
        "--recursive",
        action="store_true",
        help="Search for Excel files recursively",
    )
    parser.add_argument(
        "--values-only",
        action="store_true",
        help="Copy the last calculated values instead of formulas",
    )
    return parser.parse_args(args)


def main(argv: Sequence[str] | None = None) -> int:
    args = _parse_args(argv or sys.argv[1:])
    try:
        merged = merge_excel_files(
            source_directory=args.source_directory,
            output_path=args.output,
            pattern=args.pattern,
            recursive=args.recursive,
            values_only=args.values_only,
        )
    except Exception as exc:  # pragma: no cover - CLI error handling
        message = f"Error: {exc}\n"
        sys.stderr.write(message)
        return 1

    file_count = len(merged)
    plural = "s" if file_count != 1 else ""
    sys.stdout.write(
        f"Merged {file_count} Excel file{plural} into '{args.output}'.\n"
    )
    return 0


if __name__ == "__main__":  # pragma: no cover - CLI entry point
    sys.exit(main())