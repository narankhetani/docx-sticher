"""Finding, ordering and merging .docx files.

This module has no UI code, so the CLI, the GUI and your own scripts can all use it.
"""

from __future__ import annotations

import fnmatch
import os
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


def _style_id_names(styles) -> list[tuple[str, str | None]]:
    """``(style_id, UI name)`` per style, without building a Style object for each.

    Iterating ``doc.styles`` the normal way parses each style's type through an enum;
    docxcompose does that for every document, which costs more than the merge itself.
    """
    from docx.styles import BabelFish

    pairs = []
    for element in styles.element.style_lst:
        name = element.name_val
        pairs.append((element.styleId, None if name is None else BabelFish.internal2ui(name)))
    return pairs


class _CachedStyles:
    """The document's styles, with the id/name listing cached until a style is added."""

    def __init__(self, styles):
        self._styles = styles
        self._count = -1
        self._cached: list = []

    def __iter__(self):
        from collections import namedtuple

        count = len(self._styles.element.style_lst)
        if count != self._count:
            style = namedtuple("style", "style_id name")
            self._cached = [style(sid, name) for sid, name in _style_id_names(self._styles)]
            self._count = count
        return iter(self._cached)

    def __getattr__(self, name):
        return getattr(self._styles, name)

    def __getitem__(self, key):
        return self._styles[key]

    def __contains__(self, name):
        return name in self._styles

    def __len__(self):
        return len(self._styles)


class _DocWithCachedStyles:
    """The merged document, handing out :class:`_CachedStyles` instead of fresh Styles."""

    def __init__(self, doc):
        self._doc = doc
        self._styles = _CachedStyles(doc.styles)

    @property
    def styles(self):
        return self._styles

    def __getattr__(self, name):
        return getattr(self._doc, name)


def _fast_composer(master):
    """A docxcompose Composer that stays fast when merging hundreds of documents.

    The stock ``Composer.insert`` handles every top-level body element on its own:
    it re-lists all styles, runs a dozen XPath queries and inserts by index into an
    ever-growing body, then renumbers ids across the whole merged body after every
    document. That makes a 700-page merge take minutes. This subclass runs the
    order-independent steps once per document, keeps the per-element numbering
    steps (list restarts depend on element order), and renumbers ids once on save.
    """
    from copy import deepcopy

    from docx.oxml import OxmlElement
    from docx.oxml.section import CT_SectPr
    from docxcompose.composer import Composer
    from docxcompose.properties import CustomProperties

    class FastComposer(Composer):
        def _create_style_id_mapping(self, doc):
            self._style_id2name = dict(_style_id_names(doc.styles))
            self._style_name2id = {name: sid for sid, name in _style_id_names(self.doc.styles)}

        def insert(self, index, doc, remove_property_fields=True):
            self.reset_reference_mapping()
            self._current_preserved_styles = {}

            if remove_property_fields:
                cprops = CustomProperties(doc)
                for name in cprops.keys():  # noqa: SIM118 - not a dict, has no __iter__
                    cprops.dissolve_fields(name)

            self._create_style_id_mapping(doc)
            self.retain_formatting_from_default_styles(doc)

            # Every step below searches descendants (".//"), so running it on a wrapper
            # holding the whole body does the same work as once per element.
            batch = OxmlElement("w:body")
            for element in doc.element.body:
                if not isinstance(element, CT_SectPr):
                    batch.append(deepcopy(element))
            self.add_referenced_parts(doc.part, self.doc.part, batch)
            self.add_styles(doc, batch)
            self.add_images(doc, batch)
            self.add_diagrams(doc, batch)
            self.add_shapes(doc, batch)
            self.add_footnotes(doc, batch)
            self.remove_header_and_footer_references(doc, batch)

            body = self.doc.element.body
            anchor = body[index] if index < len(body) else None
            elements = list(batch)
            for element in elements:
                if anchor is None:
                    body.append(element)
                else:
                    anchor.addprevious(element)
            for element in elements:
                self.add_numberings(doc, element)
                self.restart_first_numbering(doc, element)

            self.add_styles_from_other_parts(doc)
            # Both are no-ops for single-section documents, but count the merged
            # body's sections first, which gets slower with every document added.
            if len(doc.sections) > 1:
                self.fix_section_types(doc)
                self.fix_header_and_footers(doc)

        def save(self, filename):
            Composer.renumber_bookmarks(self)
            Composer.renumber_docpr_ids(self)
            Composer.renumber_nvpicpr_ids(self)
            super().save(filename)

    return FastComposer(_DocWithCachedStyles(master))


def strip_images(doc) -> int:
    """Remove pictures from a document in place and return how many were removed.

    Drops DrawingML pictures (``w:drawing``), the older VML kind (``w:pict``) and
    embedded objects (``w:object``). Text, styles, lists and tables are untouched;
    the runs that held a picture stay behind, empty.
    """
    from docx.opc.constants import RELATIONSHIP_TYPE as RT

    removed = 0
    for element in doc.element.body.xpath(".//w:drawing | .//w:pict | .//w:object"):
        parent = element.getparent()
        if parent is not None:
            parent.remove(element)
            removed += 1
    if removed:
        # Drop the now-unused image relationships, or the picture files would still
        # be saved into the merged document and keep it large.
        part = doc.part
        for rId in [r for r, rel in part.rels.items() if rel.reltype == RT.IMAGE]:
            part.drop_rel(rId)
    return removed


def check_docx(path: Path) -> str | None:
    """Why *path* can't be merged, or None if it looks like a complete .docx.

    Only reads the zip directory, so checking hundreds of files takes about a second.
    Catches the common case of a file that was copied or downloaded part way.
    """
    import zipfile

    try:
        with zipfile.ZipFile(path) as package:
            names = set(package.namelist())
    except zipfile.BadZipFile:
        return "not a complete .docx - it looks like it was copied or downloaded only part way"
    except OSError as exc:
        return exc.strerror or str(exc)
    if "word/document.xml" not in names:
        return "not a Word document (no word/document.xml inside)"
    return None


def _check_saved(path: Path) -> None:
    """Fail unless *path* is a complete .docx, so a cut-short save is never kept.

    A merge that dies part way through still leaves a file on disk. It looks fine in
    Explorer but has no zip directory, and Word refuses to open it.
    """
    import zipfile

    try:
        with zipfile.ZipFile(path) as package:
            names = set(package.namelist())
    except (zipfile.BadZipFile, OSError) as exc:
        raise StitchError(
            f"The merged file was not written completely ({path.stat().st_size / 1e6:.0f} MB so far). "
            "The merge may have run out of memory or disk space. Nothing was overwritten."
        ) from exc
    missing = {"[Content_Types].xml", "word/document.xml"} - names
    if missing:
        raise StitchError(f"The merged file is incomplete: {', '.join(sorted(missing))} is missing.")


def stitch(
    files: Sequence[Path | str],
    output: Path | str,
    *,
    page_breaks: bool = True,
    keep_images: bool = True,
    skip_unreadable: bool = False,
    overwrite: bool = False,
    on_progress: ProgressCallback | None = None,
    on_skip: Callable[[Path, str], None] | None = None,
) -> Path:
    """Merge *files* in order into *output* and return its path.

    Uses docxcompose, so images, styles, numbering, footnotes and tables survive the
    merge. With ``keep_images=False`` pictures are left out, which makes the merged
    file much smaller; everything else is merged the same way.
    """
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

    # Check every file up front: on a folder of hundreds this takes about a second,
    # instead of failing minutes into the merge.
    broken = [(f, why) for f in files if (why := check_docx(f)) is not None]
    if broken and not skip_unreadable:
        names = ", ".join(f.name for f, _ in broken[:3])
        more = f" and {len(broken) - 3} more" if len(broken) > 3 else ""
        raise StitchError(
            f"{len(broken)} file(s) can't be read: {names}{more}.\n"
            f"The first one is {broken[0][1]}.\n"
            "Copy those files again, or skip them and merge the rest."
        )
    if broken:
        unusable = {f for f, _ in broken}
        files = [f for f in files if f not in unusable]
        for path, why in broken:
            if on_skip:
                on_skip(path, why)
        if not files:
            raise StitchError("None of the files could be read.")

    total = len(files)
    if on_progress:
        on_progress(0, total, files[0])
    master = _open(files[0])
    if not keep_images:
        strip_images(master)
    composer = _fast_composer(master)
    for index, path in enumerate(files[1:], start=1):
        if on_progress:
            on_progress(index, total, path)
        doc = _open(path)
        if not keep_images:
            strip_images(doc)
        if page_breaks:
            master.add_page_break()
        try:
            composer.append(doc)
        except Exception as exc:
            raise StitchError(f"Could not merge {path.name}: {exc}") from exc

    output.parent.mkdir(parents=True, exist_ok=True)
    # Save beside the output first: if the merge is cut short (out of memory, the app
    # closed, a full disk, a sync client) a half-written file never lands as the result.
    # The leading dot keeps the part file out of any later merge of the same folder.
    part = output.with_name(f".{output.stem}.part-{os.getpid()}{output.suffix}")
    try:
        composer.save(str(part))
        _check_saved(part)
        os.replace(part, output)
    except PermissionError as exc:
        part.unlink(missing_ok=True)
        raise StitchError(f"Could not save {output}: permission denied (is it open in Word?)") from exc
    except OSError as exc:
        part.unlink(missing_ok=True)
        raise StitchError(f"Could not save {output}: {exc.strerror or exc}") from exc
    except BaseException:  # MemoryError, KeyboardInterrupt, anything else mid-save
        part.unlink(missing_ok=True)
        raise
    if on_progress:
        on_progress(total, total, output)
    return output
