"""Helpers for choosing writable output directories."""

from pathlib import Path
import sys


def _is_app_bundle_path(path: Path) -> bool:
    return any(part.endswith(".app") for part in path.parts)


def default_output_dir_for_base_dir(
    base_dir: Path,
    app_name: str,
    home_dir: Path | None = None,
    is_frozen: bool | None = None,
) -> Path:
    """Use a user-writable default when running from a packaged app."""
    base_path = Path(base_dir).expanduser()
    frozen = bool(getattr(sys, "frozen", False)) if is_frozen is None else is_frozen

    if frozen or _is_app_bundle_path(base_path):
        home_path = Path.home() if home_dir is None else Path(home_dir).expanduser()
        documents_dir = home_path / "Documents"
        if home_dir is not None or documents_dir.exists():
            return documents_dir / app_name / "output"
        return home_path / app_name / "output"

    return base_path / "output"
