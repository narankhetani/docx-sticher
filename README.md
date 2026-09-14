# DOCX Stitcher

Merge many Word (`.docx`) files into one, in the right order. Use the desktop app or the command line.

- **Drag and drop** files or folders, check the order, click **Stitch**.
- **Smart ordering**: `part_2` comes before `part_10`, and zero-padded names (`007`) work too. Files
  without a number go first (a cover page, say). You can also sort by name or date, or drag rows into any order.
- **Keeps your formatting**: images, tables, styles, numbered lists and footnotes carry over (merging uses
  [docxcompose](https://github.com/4teamwork/docxcompose)).
- **Safe**: never overwrites without asking, skips Word lock files (`~$...docx`) and never merges the
  output file back into itself.
- **Finds duplicates**: files like `part_3 (1).docx` or `part_3 - Copy.docx` are flagged, so you can leave them out.

## Install on Windows

1. Press the **Start** button, type `PowerShell` and open it.
2. Paste this line and press **Enter**:

   ```powershell
   powershell -ExecutionPolicy ByPass -c "irm https://raw.githubusercontent.com/narankhetani/docx-sticher/main/install.ps1 | iex"
   ```

That's it. The app opens when setup finishes. After that you can start it from:

- the **Start Menu** or the **Desktop** shortcut, or
- by right-clicking a folder of `.docx` files and choosing **Send to > DOCX Stitcher**.

No admin rights or git needed. Python is downloaded automatically by [uv](https://docs.astral.sh/uv/).
To **update**, run the same line again. To **uninstall**:

```powershell
powershell -ExecutionPolicy ByPass -c "irm https://raw.githubusercontent.com/narankhetani/docx-sticher/main/uninstall.ps1 | iex"
```

## Install on macOS / Linux

Install [uv](https://docs.astral.sh/uv/getting-started/installation/):

```sh
curl -LsSf https://astral.sh/uv/install.sh | sh
```

Then either run it straight from GitHub with nothing to install:

```sh
uvx --from https://github.com/narankhetani/docx-sticher/archive/refs/heads/main.zip docx-stitcher
```

or install it once, so `docx-stitcher` is always on your PATH:

```sh
uv tool install https://github.com/narankhetani/docx-sticher/archive/refs/heads/main.zip
```

## Desktop app

```sh
docx-stitcher                   # open the app
docx-stitcher --gui ~/chapters  # open the app with a folder already loaded
docx-stitcher-app               # same as the first line, without a console window on Windows
```

1. Drop a folder or some `.docx` files onto the window, or use **Add folder...** / **Add files...**.
2. Check the order. Drag rows, or use **Move up** / **Move down** (`Cmd/Ctrl + Up/Down`). You can also
   re-sort by number, name or date. `Delete` removes the selected rows.
3. Choose where to save (by default, `merged.docx` in the first folder), then click **Stitch**.
4. Click **Open** or **Show in folder** when it's done.

## Command line

```sh
docx-stitcher ~/chapters                          # -> ~/chapters/merged.docx
docx-stitcher ~/chapters --dry-run                # show the order, write nothing
docx-stitcher intro.docx body.docx -o book.docx   # files are merged in the order you give them
docx-stitcher ~/chapters --skip-copies --exclude "draft*"
docx-stitcher ~/chapters --sort modified --no-page-breaks --rename
```

| Option | What it does |
| --- | --- |
| `-o, --output FILE` | Where to save (default: `merged.docx` in the first folder) |
| `-s, --sort MODE` | `number` (default for folders), `name`, `modified`, or `none` (default for files) |
| `--reverse` | Reverse the final order |
| `-r, --recursive` | Include subfolders |
| `-x, --exclude GLOB` | Skip matching file names (can be repeated) |
| `--skip-copies` | Skip duplicates like `name (1).docx` / `name - Copy.docx` |
| `--no-page-breaks` | Don't start each document on a new page |
| `-f, --force` / `--rename` | Overwrite an existing output file / save as `merged (2).docx` instead |
| `-n, --dry-run` | List the merge order and exit |
| `-q, --quiet` | Only print errors |
| `--gui` | Open the desktop app with the given paths |

## Use it from Python

```python
from docx_stitcher import collect_inputs, stitch

files = collect_inputs(["chapters"], skip_copies=True)
stitch(files, "book.docx", page_breaks=True)
```

## Development

```sh
uv sync                 # create .venv with all dependencies
uv run docx-stitcher    # run from source
uv run pytest           # tests
uv run ruff check . && uv run ruff format .
```

Code layout: `src/docx_stitcher/core.py` finds, orders and merges files (no UI code), `cli.py` is the
command line and `gui.py` is the Tkinter desktop app.

### Troubleshooting

- **`Can't find a usable init.tcl`** when opening the app: older Python builds from uv ship a broken Tcl/Tk.
  Run `uv python upgrade`, then `uv sync --reinstall`.
- **Linux**: drag and drop needs `tkdnd`, which ships with `tkinterdnd2`. If it can't load, the app still
  works; use the **Add** buttons instead.
- **`.doc` files** (the old Word format) aren't supported. Open them in Word and save as `.docx` first.
