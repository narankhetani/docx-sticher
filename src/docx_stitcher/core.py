"""Finding, ordering and merging .docx files.

This module has no UI code, so the CLI, the GUI and your own scripts can all use it.
"""

from __future__ import annotations

import fnmatch
import re
from collections.abc import Callable, Iterable, Sequence
from pathlib import Path
from typing import Literal

DEFAULT_OUTPUT_NAME = "merged.docx"

SortMode = Literal["number", "name", "modified", "none"]
SORT_MODES: tuple[SortMode, ...] = ("number", "name", "modified", "none")

ProgressCallback = Callable[[int, int, Path], None]

_DIGITS = re.compile(r"(\d+)")
# " (1)" / " - Copy" style suffixes that browsers and file managers add to duplicates
_COPY_SUFFIX = re.compile(r"(\s*\(\d+\)|\s*-\s*copy(\s*\(\d+\))?|\s+copy(\s+\d+)?)$", re.IGNORECASE)


class StitchError(Exception):
    """A problem the user can fix, with a message fit to show them."""


def is_candidate(path: Path) -> bool:
    """True for real .docx files, skipping Word lock files (~$x.docx) and hidden files."""
    name = path.name
    return path.suffix.lower() == ".docx" and not name.startswith(("~$", ".")) and path.is_file()


def find_docx_files(
    folder: Path | str,
    *,
    recursive: bool = False,
    exclude: Iterable[str] = (),
    skip_copies: bool = False,
) -> list[Path]:
    """Return the .docx files in *folder* (unordered).

    *exclude* holds glob patterns matched against file names, e.g. ``"draft*"``.
    """
    folder = Path(folder)
    if not folder.is_dir():
        raise StitchError(f"Not a folder: {folder}")
    patterns = list(exclude)
    entries = folder.rglob("*") if recursive else folder.iterdir()
    return [
        p
        for p in entries
        if is_candidate(p)
        and not (skip_copies and looks_like_copy(p))
        and not any(fnmatch.fnmatch(p.name, pat) for pat in patterns)
    ]


def natural_key(path: Path) -> tuple:
    """Sort key that orders "file2" before "file10"."""
    parts = _DIGITS.split(path.stem.casefold())
    return tuple(int(part) if part.isdigit() else part for part in parts), str(path)


def looks_like_copy(path: Path) -> bool:
    """True for names like ``part_3 (1).docx`` or ``part_3 - Copy.docx``."""
    return bool(_COPY_SUFFIX.search(path.stem))


def trailing_number(path: Path) -> int | None:
    """The last number in the file name: ``report_007.docx`` -> 7, ``cover.docx`` -> None.

    Copy markers are ignored, so ``part_3 (1).docx`` -> 3.
    """
    numbers = _DIGITS.findall(_COPY_SUFFIX.sub("", path.stem))
    return int(numbers[-1]) if numbers else None


def _number_key(path: Path) -> tuple:
    number = trailing_number(path)
    # Files without a number go first (e.g. a cover page), then by number, then by name.
    return (number is not None, number or 0, natural_key(path))


def sort_files(files: Iterable[Path], mode: SortMode = "number", *, reverse: bool = False) -> list[Path]:
    """Order files for merging.

    - ``number``:   by the last number in the file name (``part_2`` before ``part_10``)
    - ``name``:     alphabetically, treating digit runs as numbers
    - ``modified``: oldest modification time first
    - ``none``:     keep the given order
    """
    files = list(files)
    if mode == "number":
        files.sort(key=_number_key)
    elif mode == "name":
        files.sort(key=natural_key)
    elif mode == "modified":
        files.sort(key=lambda p: (p.stat().st_mtime, natural_key(p)))
    elif mode != "none":
        raise ValueError(f"Unknown sort mode {mode!r}; expected one of {', '.join(SORT_MODES)}")
    if reverse:
        files.reverse()
    return files


def collect_inputs(
    inputs: Sequence[Path | str],
    *,
    sort: SortMode | None = None,
    reverse: bool = False,
    recursive: bool = False,
    exclude: Iterable[str] = (),
    skip_copies: bool = False,
    output: Path | None = None,
) -> list[Path]:
    """Expand a mix of folders and files into an ordered, de-duplicated list of .docx files.

    With ``sort=None`` inputs keep the order given and each folder's contents are
    sorted by number; any other *sort* orders the whole list. The *output* file is
    never included, so re-running is safe.
    """
    exclude = list(exclude)
    found: list[Path] = []
    for raw in inputs:
        path = Path(raw).expanduser()
        if path.is_dir():
            batch = find_docx_files(path, recursive=recursive, exclude=exclude, skip_copies=skip_copies)
            found.extend(sort_files(batch, "number"))
        elif path.is_file():
            if path.suffix.lower() != ".docx":
                raise StitchError(f"Not a .docx file: {path}")
            found.append(path)
        else:
            raise StitchError(f"No such file or folder: {path}")

    skip = output.resolve() if output else None
    seen: set[Path] = set()
    unique: list[Path] = []
    for path in found:
        resolved = path.resolve()
        if resolved != skip and resolved not in seen:
            seen.add(resolved)
            unique.append(path)

    return sort_files(unique, sort or "none", reverse=reverse)


def default_output_for(inputs: Sequence[Path | str]) -> Path:
    """``merged.docx`` inside the first folder given, or next to the first file."""
    if not inputs:
        return Path.cwd() / DEFAULT_OUTPUT_NAME
    first = Path(inputs[0]).expanduser()
    base = first if first.is_dir() else first.parent
    return base / DEFAULT_OUTPUT_NAME


def next_free_path(path: Path) -> Path:
    """``merged.docx`` -> ``merged (2).docx`` -> ``merged (3).docx`` ... whichever is free."""
    if not path.exists():
        return path
    n = 2
    while (candidate := path.with_name(f"{path.stem} ({n}){path.suffix}")).exists():
        n += 1
    return candidate


def _open(path: Path):
    from docx import Document
    from docx.opc.exceptions import PackageNotFoundError

    try:
        return Document(str(path))
    except PackageNotFoundError as exc:
        raise StitchError(f"Could not open {path.name}: file is missing or not a valid .docx") from exc
    except Exception as exc:  # zipfile.BadZipFile, KeyError, ValueError from corrupt or .doc files
        raise StitchError(f"Could not read {path.name}: {exc}") from exc


def stitch(
    files: Sequence[Path | str],
    output: Path | str,
    *,
    page_breaks: bool = True,
    overwrite: bool = False,
    on_progress: ProgressCallback | None = None,
) -> Path:
    """Merge *files* in order into *output* and return its path.

    Uses docxcompose, so images, styles, numbering, footnotes and tables survive the merge.
    """
    from docxcompose.composer import Composer

    files = [Path(f) for f in files]
    output = Path(output)
    if not files:
        raise StitchError("There are no .docx files to merge.")
    if output.suffix.lower() != ".docx":
        raise StitchError(f"The output file must end in .docx: {output.name}")
    if output.exists() and not overwrite:
        raise StitchError(f"{output} already exists. Choose another name or allow overwriting.")
    if output.resolve() in {f.resolve() for f in files}:
        raise StitchError(f"The output file {output.name} is also one of the inputs.")

    total = len(files)
    if on_progress:
        on_progress(0, total, files[0])
    master = _open(files[0])
    composer = Composer(master)
    for index, path in enumerate(files[1:], start=1):
        if on_progress:
            on_progress(index, total, path)
        doc = _open(path)
        if page_breaks:
            master.add_page_break()
        try:
            composer.append(doc)
        except Exception as exc:
            raise StitchError(f"Could not merge {path.name}: {exc}") from exc

    output.parent.mkdir(parents=True, exist_ok=True)
    try:
        composer.save(str(output))
    except PermissionError as exc:
        raise StitchError(f"Could not save {output}: permission denied (is it open in Word?)") from exc
    if on_progress:
        on_progress(total, total, output)
    return output
