"""Merge many Word (.docx) files into one."""

from importlib.metadata import PackageNotFoundError, version

try:
    __version__ = version("docx-stitcher")
except PackageNotFoundError:  # running from a source checkout without installing
    __version__ = "0.0.0"

from .core import (  # noqa: E402
    StitchError,
    collect_inputs,
    find_docx_files,
    sort_files,
    stitch,
)

__all__ = ["StitchError", "__version__", "collect_inputs", "find_docx_files", "main", "sort_files", "stitch"]


def main() -> None:
    from .cli import main as cli_main

    raise SystemExit(cli_main())
