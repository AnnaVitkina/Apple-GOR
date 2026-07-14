"""Project folder paths for GOR rate conversion."""

from __future__ import annotations

import os
from pathlib import Path

_DEFAULT_ROOT = Path(__file__).resolve().parent

ROOT = _DEFAULT_ROOT
INPUT_DIR = ROOT / "input"
PROCESSING_DIR = ROOT / "processing"
OUTPUT_DIR = ROOT / "output"


def _path_from_env(name: str) -> Path | None:
    value = os.environ.get(name)
    if not value:
        return None
    return Path(value).expanduser()


def configure_paths(
    *,
    root: Path | str | None = None,
    input_dir: Path | str | None = None,
    processing_dir: Path | str | None = None,
    output_dir: Path | str | None = None,
) -> None:
    """Override data folder locations (for Colab / Google Drive)."""
    global ROOT, INPUT_DIR, PROCESSING_DIR, OUTPUT_DIR

    if root is not None:
        ROOT = Path(root).expanduser().resolve()
    if input_dir is not None:
        INPUT_DIR = Path(input_dir).expanduser().resolve()
    if processing_dir is not None:
        PROCESSING_DIR = Path(processing_dir).expanduser().resolve()
    if output_dir is not None:
        OUTPUT_DIR = Path(output_dir).expanduser().resolve()


def configure_paths_from_env() -> None:
    """Apply GOR_* environment variables when set."""
    root = _path_from_env("GOR_ROOT")
    input_dir = _path_from_env("GOR_INPUT_DIR")
    processing_dir = _path_from_env("GOR_PROCESSING_DIR")
    output_dir = _path_from_env("GOR_OUTPUT_DIR")

    if any(path is not None for path in (root, input_dir, processing_dir, output_dir)):
        configure_paths(
            root=root or ROOT,
            input_dir=input_dir or (root / "input" if root is not None else INPUT_DIR),
            processing_dir=processing_dir or (root / "processing" if root is not None else PROCESSING_DIR),
            output_dir=output_dir or (root / "output" if root is not None else OUTPUT_DIR),
        )


def ensure_workspace_dirs() -> None:
    INPUT_DIR.mkdir(parents=True, exist_ok=True)
    PROCESSING_DIR.mkdir(parents=True, exist_ok=True)
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)


def print_path_config() -> None:
    print("GOR paths:")
    print(f"  Root:       {ROOT}")
    print(f"  Input:      {INPUT_DIR}")
    print(f"  Processing: {PROCESSING_DIR}")
    print(f"  Output:     {OUTPUT_DIR}")
