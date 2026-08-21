"""
Caterpillar Pre-On Carriage — Rate Agreement update pipeline.

Reads yellow-highlighted changes from input/Rate Card/, applies them to
input/previous RA/ workbooks, and writes updated files to output/.

Usage (local):
  python run_pipeline.py
  python run_pipeline.py --auto

Legacy matrix build (optional):
  python run_pipeline.py --build-matrix --auto --start-date 01.01.2026
"""

from __future__ import annotations

import argparse
import os
import sys
from dataclasses import dataclass
from pathlib import Path

_COLAB_CODE_DIRS = (
    Path("/content/CAT-PreOnCarriage"),
    Path("/content/CAT-Pre-On-Carriage"),
)


def _resolve_code_dir() -> Path:
    try:
        return Path(__file__).resolve().parent
    except NameError:
        pass
    for path in _COLAB_CODE_DIRS:
        if path.is_dir():
            return path
    return Path.cwd()


_CODE_DIR = _resolve_code_dir()
if str(_CODE_DIR) not in sys.path:
    sys.path.insert(0, str(_CODE_DIR))

_PIPELINE_MODULES = (
    "project_paths",
    "number_utils",
    "build_matrix",
    "update_ra",
)


def _bootstrap_paths() -> object:
    for module_name in _PIPELINE_MODULES:
        sys.modules.pop(module_name, None)

    import project_paths

    project_paths.configure_paths_from_env()
    return project_paths


_project_paths = _bootstrap_paths()
configure_paths_from_env = _project_paths.configure_paths_from_env
print_path_config = _project_paths.print_path_config

from update_ra import UpdateRAResult, run_update_ra  # noqa: E402


@dataclass(frozen=True)
class PipelineResult:
    ra_update: UpdateRAResult


def run_pipeline(
    *,
    auto: bool = False,
    rate_card_file: Path | None = None,
    previous_ra_files: list[Path] | None = None,
    output_dir: Path | None = None,
) -> PipelineResult:
    configure_paths_from_env()
    print_path_config()

    print("\n=== Update previous RA from Rate Card ===")
    ra_update = run_update_ra(
        auto=auto,
        rate_card_file=rate_card_file,
        previous_ra_files=previous_ra_files,
        output_dir=output_dir,
    )

    print("\n=== Pipeline complete ===")
    for item in ra_update.outputs:
        print(f"  Updated RA:          {item.output_path}")
        print(f"    Source:            {item.source_ra_path.name}")
        print(
            f"    Updates: {item.updates_applied} | "
            f"Cells highlighted: {item.cells_highlighted} | "
            f"New lanes: {item.lanes_added} | "
            f"Expired highlighted: {item.expired_highlighted} | "
            f"Extended highlighted: {item.extended_highlighted}"
        )

    return PipelineResult(ra_update=ra_update)


def _parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Update previous RA workbooks from a highlighted Rate Card.",
    )
    parser.add_argument(
        "--auto",
        action="store_true",
        help="Use the first rate card file without prompts.",
    )
    parser.add_argument(
        "--rate-card",
        type=Path,
        default=None,
        help="Optional rate card workbook path.",
    )
    parser.add_argument(
        "--previous-ra",
        type=Path,
        nargs="+",
        default=None,
        help="Optional previous RA workbook path(s).",
    )
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=None,
        help="Optional output directory.",
    )
    return parser.parse_args()


def _env_flag(name: str) -> bool:
    return os.environ.get(name, "").strip().lower() in {"1", "true", "yes", "on"}


def _running_in_notebook() -> bool:
    if "colab_kernel_launcher" in Path(sys.argv[0]).name:
        return True
    if any(arg == "-f" for arg in sys.argv):
        return True
    return "ipykernel" in sys.modules or "IPython" in sys.modules


def main() -> int:
    try:
        args = _parse_args()
        run_pipeline(
            auto=args.auto,
            rate_card_file=args.rate_card,
            previous_ra_files=args.previous_ra,
            output_dir=args.output_dir,
        )
        return 0
    except KeyboardInterrupt:
        print("\nOperation cancelled.")
        return 130
    except Exception as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    if _running_in_notebook():
        run_pipeline(auto=_env_flag("CAT_PRECARR_AUTO"))
    else:
        raise SystemExit(main())
