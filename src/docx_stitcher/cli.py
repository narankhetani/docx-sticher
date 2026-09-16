"""Command-line interface. Run with no arguments to open the desktop app."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from . import __version__
from .core import (
    SORT_MODES,
    StitchError,
    collect_inputs,
    default_output_for,
    looks_like_copy,
    next_free_path,
    stitch,
)

EPILOG = """\
examples:
  docx-stitcher                         open the desktop app
  docx-stitcher ~/Documents/chapters    merge a folder into chapters/merged.docx
  docx-stitcher intro.docx body.docx -o book.docx
  docx-stitcher chapters --dry-run      show the merge order without writing anything
  docx-stitcher chapters --skip-copies --exclude "draft*" --sort name
"""


def display(path: Path) -> str:
    """A short, readable form of *path*: relative to the current folder or ~ when possible."""
    path = path.absolute()
    for base, prefix in ((Path.cwd(), ""), (Path.home(), "~/")):
        if path.is_relative_to(base):
            return prefix + str(path.relative_to(base)) if path != base else (prefix or ".")
    return str(path)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="docx-stitcher",
        description="Merge many Word (.docx) files into one, in the right order.",
        epilog=EPILOG,
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument(
        "inputs",
        nargs="*",
        type=Path,
        metavar="PATH",
        help="folders and/or .docx files to merge (none: open the desktop app)",
    )
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        help="where to save the result (default: merged.docx in the first folder)",
    )
    parser.add_argument(
        "-s",
        "--sort",
        choices=SORT_MODES,
        help="number: last number in the file name (default for folders); "
        "name: A-Z, part2 before part10; modified: oldest first; "
        "none: exactly the order given (default for files)",
    )
    parser.add_argument("--reverse", action="store_true", help="reverse the final order")
    parser.add_argument("-r", "--recursive", action="store_true", help="include .docx files in subfolders")
    parser.add_argument(
        "-x",
        "--exclude",
        action="append",
        default=[],
        metavar="GLOB",
        help='skip file names matching a pattern, e.g. "draft*" (repeatable)',
    )
    parser.add_argument(
        "--skip-copies",
        action="store_true",
        help="skip duplicate downloads like 'part_3 (1).docx' or 'part_3 - Copy.docx'",
    )
    parser.add_argument(
        "--no-page-breaks", action="store_true", help="don't start each document on a new page"
    )
    parser.add_argument(
        "--no-images", action="store_true", help="leave pictures out, for a much smaller merged file"
    )
    parser.add_argument(
        "--skip-broken",
        action="store_true",
        help="merge the rest instead of stopping when a file can't be read",
    )
    overwrite = parser.add_mutually_exclusive_group()
    overwrite.add_argument(
        "-f", "--force", action="store_true", help="overwrite the output file if it exists"
    )
    overwrite.add_argument(
        "--rename", action="store_true", help="if the output exists, save as 'merged (2).docx' instead"
    )
    parser.add_argument("-n", "--dry-run", action="store_true", help="list the files in merge order and exit")
    parser.add_argument("-q", "--quiet", action="store_true", help="only print errors")
    parser.add_argument("--gui", action="store_true", help="open the desktop app, pre-filled with PATHs")
    parser.add_argument("-V", "--version", action="version", version=f"%(prog)s {__version__}")
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)

    if args.gui or not args.inputs:
        from .gui import run

        return run(args.inputs)

    say = (lambda *_a, **_k: None) if args.quiet else print
    progress_shown = False
    try:
        output = (args.output or default_output_for(args.inputs)).expanduser()
        files = collect_inputs(
            args.inputs,
            sort=args.sort,
            reverse=args.reverse,
            recursive=args.recursive,
            exclude=args.exclude,
            skip_copies=args.skip_copies,
            output=output,
        )
        if not files:
            raise StitchError("No .docx files found.")

        width = len(str(len(files)))
        folders = {p.parent for p in files}
        where = f" from {display(folders.pop())}" if len(folders) == 1 else ""
        say(f"{len(files)} file{'s' if len(files) != 1 else ''} in merge order{where}:")
        for i, path in enumerate(files, 1):
            say(f"  {i:>{width}}. {path.name if where else display(path)}")
        copies = [p.name for p in files if looks_like_copy(p)]
        if copies:
            say(f"\nnote: {len(copies)} file(s) look like duplicate copies, e.g. '{copies[0]}'.")
            say("      Add --skip-copies to leave them out.")
        if args.dry_run:
            say(f"\nWould save to {display(output)}")
            return 0

        if args.rename:
            output = next_free_path(output)
        elif output.exists() and not args.force:
            raise StitchError(
                f"{display(output)} already exists. Use --force to overwrite or --rename to keep both."
            )

        live = not args.quiet and sys.stdout.isatty()

        def progress(done: int, total: int, path: Path) -> None:
            nonlocal progress_shown
            if live and done < total:
                progress_shown = True
                print(f"\r\033[KMerging {done + 1}/{total}: {path.name}", end="", flush=True)

        def skipped(path: Path, why: str) -> None:
            print(f"skipped {path.name}: {why}", file=sys.stderr)

        stitch(
            files,
            output,
            page_breaks=not args.no_page_breaks,
            keep_images=not args.no_images,
            skip_unreadable=args.skip_broken,
            overwrite=True,
            on_progress=progress,
            on_skip=skipped,
        )
        if progress_shown:
            print("\r\033[K", end="")
        say(f"\nSaved {display(output)}")
        return 0
    except StitchError as exc:
        if progress_shown:
            print()
        print(f"error: {exc}", file=sys.stderr)
        return 1
    except KeyboardInterrupt:
        print("\ncancelled", file=sys.stderr)
        return 130


if __name__ == "__main__":
    sys.exit(main())
