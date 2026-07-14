"""
Build Postal Code Zones txt from Additional Info common-rating tables.

Output columns (tab-separated):
  Name | Country | Postal Code | Excluded
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill

from build_matrix import DATA_START_ROW, SHIPMENT_COLUMNS, cell_text
from country_codes import to_country_iso
from project_paths import OUTPUT_DIR, PROCESSING_DIR, ensure_workspace_dirs
from us_ca_city_states import format_us_ca_postal_city

ORIGIN_COMMON_RATING_SECTION = "GOR25 Origin Common Rating"
DESTINATION_COMMON_RATING_SECTION = "GOR25 Destination Common Rating"

POSTAL_ZONE_COLUMNS = ("Name", "Country", "Postal Code", "Excluded")
POSTAL_CODE_ZONES_SUFFIX = "_postal_code_zones.txt"

DESTINATION_HIGHLIGHT_FILL = PatternFill("solid", fgColor="DAEEF3")
DESTINATION_HIGHLIGHT_FONT = Font(underline="single", bold=True)

DESTINATION_RAW_COLUMN = "Destination City"
DESTINATION_LABEL_COLUMN = "Destination City "
DESTINATION_LABEL_SUFFIX = " (Destination)"


@dataclass(frozen=True)
class PostalCodeZone:
    name: str
    country: str
    postal_code: str
    excluded: str = "-"

    def as_txt_row(self) -> str:
        return "\t".join([self.name, self.country, self.postal_code, self.excluded])


def load_additional_info(file_path: Path) -> pd.DataFrame:
    workbook = pd.ExcelFile(file_path)
    if "Additional Info" not in workbook.sheet_names:
        raise RuntimeError(f"No 'Additional Info' sheet found in {file_path.name}")
    return pd.read_excel(file_path, sheet_name="Additional Info")


def _unique_preserve_order(values: list[str]) -> list[str]:
    seen: set[str] = set()
    result: list[str] = []
    for value in values:
        text = cell_text(value)
        if not text:
            continue
        key = text.casefold()
        if key in seen:
            continue
        seen.add(key)
        result.append(text)
    return result


def _normalize_country(value: object) -> str:
    text = cell_text(value)
    if not text:
        return ""
    iso = to_country_iso(text)
    if iso:
        return iso.upper()
    return text.upper()


def _build_zones_from_section(
    section_df: pd.DataFrame,
    *,
    city_column: str,
    country_column: str,
    suffix: str,
) -> list[PostalCodeZone]:
    if section_df.empty:
        return []

    grouped: dict[tuple[str, str], list[str]] = {}
    for _, row in section_df.iterrows():
        common_rating_city = cell_text(row.get("Common Rating City"))
        country = _normalize_country(row.get(country_column))
        city = cell_text(row.get(city_column))
        if not common_rating_city or not country or not city:
            continue

        key = (common_rating_city, country)
        grouped.setdefault(key, []).append(city)

    zones: list[PostalCodeZone] = []
    for (common_rating_city, country), cities in grouped.items():
        unique_cities = _unique_preserve_order(cities)
        postal_cities = [common_rating_city, *unique_cities]
        postal_cities = _unique_preserve_order(postal_cities)
        postal_values = [
            format_us_ca_postal_city(city, country, common_rating_city) for city in postal_cities
        ]
        zones.append(
            PostalCodeZone(
                name=f"{common_rating_city} ({suffix})",
                country=country,
                postal_code=", ".join(postal_values),
            )
        )

    return zones


def build_postal_code_zones(additional_info_df: pd.DataFrame) -> list[PostalCodeZone]:
    origin_df = additional_info_df[
        additional_info_df["Section Title"].astype(str).eq(ORIGIN_COMMON_RATING_SECTION)
    ]
    destination_df = additional_info_df[
        additional_info_df["Section Title"].astype(str).eq(DESTINATION_COMMON_RATING_SECTION)
    ]

    zones = [
        *_build_zones_from_section(
            origin_df,
            city_column="Origin City",
            country_column="Origin Country",
            suffix="Origin",
        ),
        *_build_zones_from_section(
            destination_df,
            city_column="Destination City",
            country_column="Destination Country",
            suffix="Destination",
        ),
    ]
    zones.sort(key=lambda zone: (zone.country, zone.name.casefold()))
    return zones


def _base_city_from_destination_label(label: str) -> str:
    text = cell_text(label)
    if text.endswith(DESTINATION_LABEL_SUFFIX):
        return text[: -len(DESTINATION_LABEL_SUFFIX)].strip()
    return text


def matrix_destination_labels(matrix_df: pd.DataFrame) -> list[tuple[str, str]]:
    """Return unique (Destination City label, country) pairs from the matrix."""
    seen: set[str] = set()
    results: list[tuple[str, str]] = []

    for _, row in matrix_df.iterrows():
        if cell_text(row.get("Service Type")) == "SPECIAL":
            continue

        label = cell_text(row.get(DESTINATION_LABEL_COLUMN))
        country = _normalize_country(row.get("Destination Country"))
        if not label or not country:
            continue

        key = label.casefold()
        if key in seen:
            continue
        seen.add(key)
        results.append((label, country))

    return results


def build_final_postal_zones(
    additional_info_df: pd.DataFrame,
    matrix_df: pd.DataFrame,
) -> list[PostalCodeZone]:
    all_zones = build_postal_code_zones(additional_info_df)
    zones_by_name = {zone.name: zone for zone in all_zones}

    destination_zones: list[PostalCodeZone] = []
    for label, country in matrix_destination_labels(matrix_df):
        if label in zones_by_name:
            destination_zones.append(zones_by_name[label])
            continue

        base_city = _base_city_from_destination_label(label)
        postal_code = (
            format_us_ca_postal_city(base_city, country, base_city)
            if country in {"US", "CA"}
            else base_city
        )
        destination_zones.append(
            PostalCodeZone(
                name=label,
                country=country,
                postal_code=postal_code,
            )
        )

    destination_zones.sort(key=lambda zone: (zone.country, zone.name.casefold()))
    return destination_zones


def destination_zone_names(zones: list[PostalCodeZone]) -> set[str]:
    return {zone.name for zone in zones if zone.name.endswith("(Destination)")}


def write_postal_code_zones_txt(zones: list[PostalCodeZone], output_path: Path) -> Path:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    lines = ["\t".join(POSTAL_ZONE_COLUMNS)]
    lines.extend(zone.as_txt_row() for zone in zones)
    output_path.write_text("\n".join(lines) + "\n", encoding="utf-8")
    return output_path


def _second_destination_city_column_index() -> int:
    return SHIPMENT_COLUMNS.index(DESTINATION_LABEL_COLUMN) + 1


def highlight_destination_cities_in_matrix(
    matrix_path: Path,
    destination_zone_names_set: set[str],
) -> int:
    if not destination_zone_names_set:
        return 0

    workbook = load_workbook(matrix_path)
    worksheet = workbook["Rate card"]
    col_index = _second_destination_city_column_index()
    highlighted = 0

    for row_index in range(DATA_START_ROW, worksheet.max_row + 1):
        cell = worksheet.cell(row_index, col_index)
        value = cell_text(cell.value)
        if not value or value not in destination_zone_names_set:
            continue

        cell.fill = DESTINATION_HIGHLIGHT_FILL
        cell.font = DESTINATION_HIGHLIGHT_FONT
        highlighted += 1

    workbook.save(matrix_path)
    return highlighted


def default_postal_zones_path(source_file: Path) -> Path:
    stem = source_file.stem.replace("_extracted", "")
    return OUTPUT_DIR / f"{stem}{POSTAL_CODE_ZONES_SUFFIX}"


def apply_postal_code_zones(
    *,
    source_file: Path,
    matrix_path: Path,
    matrix_df: pd.DataFrame,
    output_path: Path | None = None,
) -> tuple[Path, int, int, int]:
    additional_info_df = load_additional_info(source_file)
    all_zones = build_postal_code_zones(additional_info_df)
    final_zones = build_final_postal_zones(additional_info_df, matrix_df)

    txt_path = output_path or default_postal_zones_path(source_file)
    write_postal_code_zones_txt(final_zones, txt_path)

    highlight_names = destination_zone_names(final_zones)
    highlighted = highlight_destination_cities_in_matrix(matrix_path, highlight_names)
    return txt_path, len(final_zones), len(all_zones), highlighted


def run_build_postal_code_zones(
    *,
    source_file: Path,
    matrix_path: Path,
    matrix_df: pd.DataFrame,
    output_path: Path | None = None,
) -> Path:
    txt_path, used_count, total_count, highlighted = apply_postal_code_zones(
        source_file=source_file,
        matrix_path=matrix_path,
        matrix_df=matrix_df,
        output_path=output_path,
    )
    print(f"  Wrote {used_count} postal code zone(s) from {total_count} common-rating zone(s)")
    print(f"  Highlighted {highlighted} destination label cell(s) in matrix")
    print(f"  Saved postal code zones to: {txt_path}")
    return txt_path


def list_extracted_files() -> list[Path]:
    if not PROCESSING_DIR.is_dir():
        return []
    return sorted(PROCESSING_DIR.glob("*_extracted.xlsx"), key=lambda path: path.stat().st_mtime, reverse=True)
