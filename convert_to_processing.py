"""
Convert selected tabs from an input Excel workbook into cleaned dataframes
and save them to the processing/ folder.

Interactive flow:
  1. Choose input file
  2. Review proposed default tabs
  3. Accept, change, or add tabs
  4. Write one multi-sheet workbook to processing/
"""

from __future__ import annotations

import re
import sys
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path

import pandas as pd

from project_paths import INPUT_DIR, PROCESSING_DIR, ensure_workspace_dirs

EXCEL_SUFFIXES = {".xlsx", ".xls", ".xlsm"}

DEFAULT_SHEET_EXACT = ("Additional Info", "Base Freight Rates", "Service Fees")
DEFAULT_SHEET_SUBSTRINGS = ("bcl", "fcl")

HEADER_MARKERS = (
    "original naming convention",
    "origin city",
    "destination city",
    "service type",
    "charge code supplier",
)

SECTION_TITLE_PATTERN = re.compile(r"gor25\s+.+\s+common\s+rating", re.IGNORECASE)


@dataclass(frozen=True)
class SheetSelection:
    file_path: Path
    sheet_name: str

    @property
    def label(self) -> str:
        return f"{self.file_path.name} -> {self.sheet_name}"


@dataclass
class SheetMetadata:
    section_titles: list[str] = field(default_factory=list)
    valid_from: object = None
    valid_to: object = None


def list_input_files() -> list[Path]:
    return [
        path
        for path in sorted(INPUT_DIR.iterdir())
        if path.is_file()
        and path.suffix.lower() in EXCEL_SUFFIXES
        and not path.name.startswith("~$")
    ]


def parse_selection(raw: str, max_index: int) -> list[int]:
    """Parse '1,3-5' into zero-based indices."""
    raw = raw.strip().lower()
    if raw in {"all", "*"}:
        return list(range(max_index))

    indices: set[int] = set()
    for part in re.split(r"\s*,\s*", raw):
        if not part:
            continue
        if "-" in part:
            start_s, end_s = part.split("-", 1)
            start = int(start_s) - 1
            end = int(end_s) - 1
            if start > end or start < 0 or end >= max_index:
                raise ValueError(f"Invalid range: {part}")
            indices.update(range(start, end + 1))
        else:
            idx = int(part) - 1
            if idx < 0 or idx >= max_index:
                raise ValueError(f"Invalid index: {part}")
            indices.add(idx)
    return sorted(indices)


def prompt_selection(title: str, items: list[str], allow_empty: bool = False) -> list[int]:
    if not items:
        return []

    print(f"\n{title}")
    for i, item in enumerate(items, start=1):
        print(f"  {i}. {item}")

    hint = "Enter numbers (e.g. 1,3 or 1-3) or 'all'"

    while True:
        raw = input(f"{hint}: ").strip()
        if not raw and allow_empty:
            return []
        if not raw:
            print("Please enter at least one choice.")
            continue
        try:
            chosen = parse_selection(raw, len(items))
            if chosen or allow_empty:
                return chosen
            print("Please enter at least one choice.")
        except ValueError as exc:
            print(f"Invalid input: {exc}")


def ask_yes_no(prompt: str) -> bool:
    while True:
        answer = input(f"{prompt} (y/n): ").strip().lower()
        if answer in {"y", "yes"}:
            return True
        if answer in {"n", "no"}:
            return False
        print("Please enter y or n.")


def propose_default_tabs(sheet_names: list[str]) -> list[str]:
    """Match default exact tabs and tabs containing bcl/fcl."""
    by_lower = {name.lower(): name for name in sheet_names}
    selected: list[str] = []
    seen_lower: set[str] = set()

    for exact_name in DEFAULT_SHEET_EXACT:
        actual = by_lower.get(exact_name.lower())
        if actual is not None and actual.lower() not in seen_lower:
            selected.append(actual)
            seen_lower.add(actual.lower())

    for sheet_name in sheet_names:
        lower_name = sheet_name.lower()
        if lower_name in seen_lower:
            continue
        if any(substr in lower_name for substr in DEFAULT_SHEET_SUBSTRINGS):
            selected.append(sheet_name)
            seen_lower.add(lower_name)

    return selected


def select_input_file(files: list[Path], *, auto: bool = False) -> Path:
    if not files:
        print(f"No Excel files found in: {INPUT_DIR}")
        sys.exit(1)

    if auto:
        if len(files) == 1:
            print(f"\nAuto mode: using input file {files[0].name}")
        else:
            print(f"\nAuto mode: using first input file {files[0].name}")
        return files[0]

    print("\nAvailable input files:")
    for i, path in enumerate(files, start=1):
        print(f"  {i}. {path.name}")

    indices = prompt_selection("Select file to convert:", [f.name for f in files])
    if len(indices) != 1:
        print("Please select exactly one file.")
        return select_input_file(files)
    return files[indices[0]]


def print_all_tabs(sheet_names: list[str], selected: list[str]) -> None:
    print("\nAll available tabs:")
    for i, name in enumerate(sheet_names, start=1):
        marker = " [selected]" if name in selected else ""
        print(f"  {i}. {name}{marker}")


def print_current_selection(selected: list[str]) -> None:
    if selected:
        print("\nCurrent tab selection:")
        for name in selected:
            print(f"  - {name}")
    else:
        print("\nNo tabs selected yet.")


def select_tabs_interactive(file_path: Path, sheet_names: list[str]) -> list[str]:
    selected = propose_default_tabs(sheet_names)

    while True:
        print_all_tabs(sheet_names, selected)
        print_current_selection(selected)
        print("\nTab selection options:")
        print("  1. Accept current selection and convert")
        print("  2. Change tabs (replace selection)")
        print("  3. Add tabs to current selection")

        choice = input("Choose option (1-3): ").strip()

        if choice == "1":
            if not selected:
                print("Select at least one tab before converting.")
                continue
            return selected

        if choice == "2":
            indices = prompt_selection(
                f"Choose sheet(s) for {file_path.name}:",
                sheet_names,
            )
            selected = [sheet_names[i] for i in indices]
            continue

        if choice == "3":
            remaining = [name for name in sheet_names if name not in selected]
            if not remaining:
                print("All tabs are already selected.")
                continue
            indices = prompt_selection(
                "Choose additional sheet(s) to add:",
                remaining,
            )
            for idx in indices:
                tab_name = remaining[idx]
                if tab_name not in selected:
                    selected.append(tab_name)
            continue

        print("Invalid choice. Enter 1, 2, or 3.")


def _cell_text(value: object) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip()


def _format_metadata_date(value: object) -> object:
    if pd.isna(value):
        return pd.NA
    if isinstance(value, datetime):
        return value.date().isoformat()
    if isinstance(value, date):
        return value.isoformat()
    return value


def _row_texts(row: pd.Series) -> list[str]:
    return [_cell_text(value).lower() for value in row]


def _is_section_title(value: object) -> bool:
    text = _cell_text(value)
    return bool(text and SECTION_TITLE_PATTERN.search(text))


def _find_section_title_in_row(row: pd.Series) -> str | None:
    for value in row:
        text = _cell_text(value)
        if _is_section_title(text):
            return text
    return None


def extract_sheet_metadata(df_raw: pd.DataFrame, max_scan_rows: int = 50) -> SheetMetadata:
    metadata = SheetMetadata()
    scan_limit = min(len(df_raw), max_scan_rows)

    for row_idx in range(scan_limit):
        row = df_raw.iloc[row_idx]
        for col_idx, value in enumerate(row):
            text = _cell_text(value)
            lower = text.lower()

            if _is_section_title(text) and text not in metadata.section_titles:
                metadata.section_titles.append(text)

            if lower.startswith("from (yy"):
                if col_idx + 1 < len(row):
                    candidate = row.iloc[col_idx + 1]
                    if pd.notna(candidate):
                        metadata.valid_from = candidate

            if lower.startswith("to (yy"):
                if col_idx + 1 < len(row):
                    candidate = row.iloc[col_idx + 1]
                    if pd.notna(candidate):
                        metadata.valid_to = candidate

    return metadata


def _is_additional_info_header_row(row: pd.Series) -> bool:
    texts = set(_row_texts(row))
    return "origin city" in texts or "destination city" in texts


def _normalize_headers(headers: list[object]) -> list[str]:
    cleaned: list[str] = []
    seen: dict[str, int] = {}

    for idx, header in enumerate(headers, start=1):
        value = _cell_text(header)
        if not value or value.lower() == "nan":
            value = f"column_{idx}"

        base = value
        count = seen.get(base, 0)
        if count:
            value = f"{base}_{count + 1}"
        seen[base] = count + 1
        cleaned.append(value)

    return cleaned


def _drop_empty_rows(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df

    non_empty_mask = df.apply(
        lambda row: any(_cell_text(value) for value in row),
        axis=1,
    )
    return df.loc[non_empty_mask].copy()


def _drop_empty_columns(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df

    keep_columns = [
        column
        for column in df.columns
        if any(_cell_text(value) for value in df[column])
    ]
    return df.loc[:, keep_columns].copy()


def _row_contains_marker(row: pd.Series) -> bool:
    texts = set(_row_texts(row))
    return any(marker in texts for marker in HEADER_MARKERS)


def _attach_validity_columns(df: pd.DataFrame, metadata: SheetMetadata) -> pd.DataFrame:
    if metadata.valid_from is None and metadata.valid_to is None:
        return df

    enriched = df.copy()
    enriched.insert(0, "Valid To", _format_metadata_date(metadata.valid_to))
    enriched.insert(0, "Valid From", _format_metadata_date(metadata.valid_from))
    return enriched


def clean_additional_info_df(df_raw: pd.DataFrame, metadata: SheetMetadata) -> pd.DataFrame:
    sections: list[pd.DataFrame] = []
    current_title: str | None = metadata.section_titles[0] if metadata.section_titles else None
    current_headers: list[str] | None = None
    current_rows: list[list[object]] = []

    def flush_section() -> None:
        nonlocal current_rows, current_headers, current_title
        if current_headers is None or not current_rows:
            current_rows = []
            return

        section_df = pd.DataFrame(current_rows, columns=current_headers)
        section_df.insert(0, "Section Title", current_title or pd.NA)
        sections.append(section_df)
        current_rows = []

    for _, row in df_raw.iterrows():
        title = _find_section_title_in_row(row)
        if title is not None:
            flush_section()
            current_title = title
            current_headers = None
            continue

        if _is_additional_info_header_row(row):
            flush_section()
            current_headers = _normalize_headers(row.tolist())
            continue

        if current_headers is None:
            continue

        if any(_cell_text(value) for value in row):
            current_rows.append(row.tolist())

    flush_section()

    if not sections:
        return clean_standard_tab_df(df_raw, metadata)

    combined = pd.concat(sections, ignore_index=True)
    combined = _drop_empty_rows(combined)
    combined = _drop_empty_columns(combined)
    return combined.reset_index(drop=True)


def clean_standard_tab_df(df_raw: pd.DataFrame, metadata: SheetMetadata) -> pd.DataFrame:
    """Remove preamble rows above the header and fully empty rows/columns."""
    if df_raw.empty:
        return df_raw.copy()

    header_row_idx = find_header_row_index(df_raw)
    if header_row_idx is None:
        cleaned = _drop_empty_rows(df_raw)
        cleaned.columns = _normalize_headers(list(cleaned.columns))
        cleaned = cleaned.reset_index(drop=True)
        return _attach_validity_columns(cleaned, metadata)

    headers = _normalize_headers(df_raw.iloc[header_row_idx].tolist())
    data_start = header_row_idx + 1
    if data_start < len(df_raw) and _is_internal_field_row(df_raw.iloc[data_start]):
        data_start += 1

    df = df_raw.iloc[data_start:].copy()
    df.columns = headers
    df = _drop_empty_rows(df)
    df = _drop_empty_columns(df)
    df = df.reset_index(drop=True)
    return _attach_validity_columns(df, metadata)


def clean_tab_df(df_raw: pd.DataFrame, sheet_name: str | None = None) -> pd.DataFrame:
    metadata = extract_sheet_metadata(df_raw)
    if sheet_name and sheet_name.strip().lower() == "additional info":
        return clean_additional_info_df(df_raw, metadata)
    return clean_standard_tab_df(df_raw, metadata)


def find_header_row_index(df_raw: pd.DataFrame, max_scan_rows: int = 30) -> int | None:
    scan_limit = min(len(df_raw), max_scan_rows)
    for row_idx in range(scan_limit):
        if _row_contains_marker(df_raw.iloc[row_idx]):
            return row_idx
    return None


def _is_internal_field_row(row: pd.Series) -> bool:
    """Skip Geodis internal field-name row directly below display headers."""
    first_values = [_cell_text(value).lower() for value in row[:6]]
    return any(value in {"item_name", "service_type", "origin_region"} for value in first_values)


def tab_to_df(file_path: Path, sheet_name: str) -> pd.DataFrame:
    raw = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
    return clean_tab_df(raw, sheet_name=sheet_name)


MAX_SHEET_NAME_LEN = 31


def sanitize_sheet_part(text: str) -> str:
    return re.sub(r"[\[\]:*?/\\]", "_", text).strip() or "Sheet"


def output_sheet_name(sheet_name: str, used: set[str]) -> str:
    base = sanitize_sheet_part(sheet_name)
    if base not in used:
        used.add(base)
        return base

    for n in range(2, 1000):
        suffix = f"_{n}"
        candidate = sanitize_sheet_part(sheet_name)[: MAX_SHEET_NAME_LEN - len(suffix)] + suffix
        if candidate not in used:
            used.add(candidate)
            return candidate

    raise RuntimeError(f"Could not create a unique sheet name for: {sheet_name}")


def collect_frames(
    file_path: Path,
    sheet_names: list[str],
) -> list[tuple[str, pd.DataFrame]]:
    frames: list[tuple[str, pd.DataFrame]] = []
    used_names: set[str] = set()

    for sheet_name in sheet_names:
        try:
            raw = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
            metadata = extract_sheet_metadata(raw)
            df = clean_tab_df(raw, sheet_name=sheet_name)
            removed_rows = len(raw) - len(df)
        except Exception as exc:
            print(f"Skipping {sheet_name}: could not read sheet ({exc})")
            continue

        label = output_sheet_name(sheet_name, used_names)
        frames.append((label, df))
        meta_parts: list[str] = []
        if metadata.section_titles:
            meta_parts.append(f"sections: {', '.join(metadata.section_titles)}")
        if metadata.valid_from is not None or metadata.valid_to is not None:
            meta_parts.append(
                "validity: "
                f"{_format_metadata_date(metadata.valid_from)} to "
                f"{_format_metadata_date(metadata.valid_to)}"
            )
        meta_summary = f"; {'; '.join(meta_parts)}" if meta_parts else ""
        print(
            f"  Loaded: {sheet_name} -> '{label}' "
            f"({len(df)} rows, {len(df.columns)} columns; removed {removed_rows} preamble/empty rows"
            f"{meta_summary})"
        )

    return frames


def save_to_processing(file_path: Path, frames: list[tuple[str, pd.DataFrame]]) -> Path:
    output_path = PROCESSING_DIR / f"{file_path.stem}_extracted.xlsx"

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        for sheet_name, df in frames:
            df.to_excel(writer, sheet_name=sheet_name, index=False)
            print(f"  Wrote tab: {sheet_name}")

    return output_path


def run_convert(*, auto: bool = False, file_path: Path | None = None) -> Path:
    ensure_workspace_dirs()
    files = list_input_files()
    selected_file = file_path or select_input_file(files, auto=auto)

    if not selected_file.exists():
        raise FileNotFoundError(f"Input file not found: {selected_file}")

    try:
        workbook = pd.ExcelFile(selected_file)
    except Exception as exc:
        raise RuntimeError(f"Could not open {selected_file.name}: {exc}") from exc

    if auto:
        selected_tabs = propose_default_tabs(workbook.sheet_names)
        if not selected_tabs:
            raise RuntimeError(
                f"No default tabs matched in {selected_file.name}. "
                "Run without --auto to choose tabs manually."
            )
        print(f"Auto mode: converting tabs: {', '.join(selected_tabs)}")
    else:
        selected_tabs = select_tabs_interactive(selected_file, workbook.sheet_names)

    frames = collect_frames(selected_file, selected_tabs)

    if not frames:
        raise RuntimeError("Nothing to save. No sheets could be loaded.")

    output_path = save_to_processing(selected_file, frames)
    print(f"\nSaved {len(frames)} sheet(s) to: {output_path}")
    return output_path


def main() -> int:
    try:
        run_convert()
        return 0
    except KeyboardInterrupt:
        print("\nOperation cancelled.")
        return 130
    except Exception as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
