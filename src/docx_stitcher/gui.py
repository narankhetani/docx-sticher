"""Desktop app: add files, check the order, click Stitch."""

from __future__ import annotations

import contextlib
import os
import queue
import subprocess
import sys
import threading
import tkinter as tk
from collections.abc import Sequence
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from . import __version__
from .core import (
    DEFAULT_OUTPUT_NAME,
    StitchError,
    find_docx_files,
    is_candidate,
    looks_like_copy,
    next_free_path,
    sort_files,
    stitch,
)

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
except ImportError:  # drag and drop is a nice-to-have
    TkinterDnD = None

SORT_LABELS = {
    "By number in file name": "number",
    "By name (A-Z)": "name",
    "By date modified": "modified",
}
MOD = "Command" if sys.platform == "darwin" else "Control"


def reveal(path: Path, *, select: bool) -> None:
    """Open a file, or show it in Finder / Explorer / the file manager."""
    if sys.platform == "darwin":
        subprocess.Popen(["open", "-R", str(path)] if select else ["open", str(path)])
    elif sys.platform == "win32":
        if select:
            subprocess.Popen(["explorer", f"/select,{path}"])
        else:
            os.startfile(path)  # type: ignore[attr-defined]
    else:
        subprocess.Popen(["xdg-open", str(path.parent if select else path)])


class StitcherApp:
    def __init__(self, root: tk.Tk, initial: Sequence[Path] = (), *, dnd: bool = False) -> None:
        self.root = root
        self.paths: dict[str, Path] = {}
        self.output_chosen = False
        self.last_output: Path | None = None
        self.events: queue.Queue = queue.Queue()
        self.busy = False
        self._drag_iid: str | None = None

        root.title(f"DOCX Stitcher {__version__}")
        root.geometry("820x600")
        root.minsize(560, 440)
        self._build(dnd)
        self._bind_keys()
        self.add_paths(initial)
        self._refresh()

    # ---- layout ---------------------------------------------------------------------------

    def _build(self, dnd: bool) -> None:
        outer = ttk.Frame(self.root, padding=16)
        outer.pack(fill="both", expand=True)

        ttk.Label(outer, text="Merge Word documents", font=("TkDefaultFont", 18, "bold")).pack(anchor="w")
        hint = "Drop .docx files or folders below" if dnd else "Add .docx files or a folder"
        ttk.Label(
            outer, text=f"{hint}, check the order (drag to rearrange), then click Stitch.", foreground="gray"
        ).pack(anchor="w", pady=(2, 12))

        bar = ttk.Frame(outer)
        bar.pack(fill="x")
        ttk.Button(bar, text="Add folder...", command=self.ask_folder).pack(side="left")
        ttk.Button(bar, text="Add files...", command=self.ask_files).pack(side="left", padx=(6, 0))
        ttk.Button(bar, text="Reverse", command=self.reverse).pack(side="right")
        ttk.Button(bar, text="Sort", command=self.sort).pack(side="right", padx=6)
        self.sort_choice = tk.StringVar(value=next(iter(SORT_LABELS)))
        ttk.Combobox(
            bar, textvariable=self.sort_choice, values=list(SORT_LABELS), state="readonly", width=22
        ).pack(side="right")

        table = ttk.Frame(outer)
        table.pack(fill="both", expand=True, pady=(10, 6))
        self.tree = ttk.Treeview(table, columns=("n", "name", "folder", "modified"), show="headings")
        for col, title, width, stretch in (
            ("n", "#", 44, False),
            ("name", "File", 260, True),
            ("folder", "Folder", 260, True),
            ("modified", "Modified", 130, False),
        ):
            self.tree.heading(col, text=title)
            self.tree.column(col, width=width, stretch=stretch, anchor="e" if col == "n" else "w")
        scroll = ttk.Scrollbar(table, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scroll.set)
        self.tree.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")
        self.empty_label = ttk.Label(
            table,
            text=("Drop .docx files or a folder here" if dnd else 'No files yet - click "Add folder..."'),
            foreground="gray",
            font=("TkDefaultFont", 14),
        )
        self.tree.bind("<<TreeviewSelect>>", lambda _e: self._refresh())
        self.tree.bind("<ButtonPress-1>", self._drag_start, add="+")
        self.tree.bind("<B1-Motion>", self._drag_motion, add="+")
        self.tree.bind("<ButtonRelease-1>", lambda _e: setattr(self, "_drag_iid", None), add="+")
        self.tree.bind("<Double-1>", self._open_selected)
        if dnd:
            for widget in (self.tree, self.empty_label):
                widget.drop_target_register(DND_FILES)
                widget.dnd_bind("<<Drop>>", self._on_drop)

        edit = ttk.Frame(outer)
        edit.pack(fill="x")
        self.btn_up = ttk.Button(edit, text="Move up", command=lambda: self.move(-1))
        self.btn_down = ttk.Button(edit, text="Move down", command=lambda: self.move(1))
        self.btn_remove = ttk.Button(edit, text="Remove", command=self.remove_selected)
        self.btn_clear = ttk.Button(edit, text="Clear all", command=self.clear)
        for i, btn in enumerate((self.btn_up, self.btn_down, self.btn_remove, self.btn_clear)):
            btn.pack(side="left", padx=(0 if i == 0 else 6, 0))
        self.count_label = ttk.Label(edit, foreground="gray")
        self.count_label.pack(side="right")

        options = ttk.Frame(outer)
        options.pack(fill="x", pady=(14, 0))
        self.page_breaks = tk.BooleanVar(value=True)
        ttk.Checkbutton(options, text="Start each document on a new page", variable=self.page_breaks).pack(
            anchor="w"
        )
        self.keep_images = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            options, text="Keep pictures (uncheck for a smaller file)", variable=self.keep_images
        ).pack(anchor="w")
        save = ttk.Frame(outer)
        save.pack(fill="x", pady=(8, 0))
        ttk.Label(save, text="Save as:").pack(side="left")
        self.output = tk.StringVar()
        entry = ttk.Entry(save, textvariable=self.output)
        entry.pack(side="left", fill="x", expand=True, padx=6)
        entry.bind("<Key>", lambda _e: setattr(self, "output_chosen", True))
        ttk.Button(save, text="Change...", command=self.ask_output).pack(side="left")

        footer = ttk.Frame(outer)
        footer.pack(fill="x", pady=(16, 0))
        self.stitch_btn = ttk.Button(footer, text="Stitch", command=self.start_stitch, default="active")
        self.stitch_btn.pack(side="right", ipadx=12, ipady=4)
        self.show_btn = ttk.Button(footer, text="Show in folder", command=lambda: self._reveal(select=True))
        self.open_btn = ttk.Button(footer, text="Open", command=lambda: self._reveal(select=False))
        self.status = ttk.Label(footer, text="", wraplength=480, justify="left")
        self.status.pack(side="left", fill="x", expand=True)
        self.progress = ttk.Progressbar(footer, mode="determinate", length=160)

    def _bind_keys(self) -> None:
        for key in ("<Delete>", "<BackSpace>"):
            self.tree.bind(key, lambda _e: self.remove_selected())
        self.tree.bind(f"<{MOD}-Up>", lambda _e: self.move(-1) or "break")
        self.tree.bind(f"<{MOD}-Down>", lambda _e: self.move(1) or "break")
        self.tree.bind(f"<{MOD}-a>", lambda _e: self.tree.selection_set(self.tree.get_children()))
        self.root.bind(f"<{MOD}-o>", lambda _e: self.ask_files())
        self.root.bind(f"<{MOD}-Return>", lambda _e: self.start_stitch())

    # ---- list management ------------------------------------------------------------------

    def add_paths(self, paths: Sequence[Path | str]) -> None:
        added = skipped = 0
        copies: list[str] = []
        known = {p.resolve() for p in self.paths.values()}
        for raw in paths:
            path = Path(raw).expanduser()
            if path.is_dir():
                try:
                    batch = sort_files(find_docx_files(path), "number")
                except (StitchError, OSError) as exc:
                    messagebox.showerror("Can't read folder", str(exc), parent=self.root)
                    continue
                if not self.output_chosen and not self.paths:
                    self.output.set(str(path / DEFAULT_OUTPUT_NAME))
            elif is_candidate(path):
                batch = [path]
            else:
                skipped += 1
                continue
            for file in batch:
                resolved = file.resolve()
                if resolved in known or self._is_output(file):
                    continue
                known.add(resolved)
                iid = self._insert(file)
                added += 1
                if looks_like_copy(file):
                    copies.append(iid)

        if not self.output_chosen and self.paths and not self.output.get():
            first = next(iter(self.paths.values()))
            self.output.set(str(first.parent / DEFAULT_OUTPUT_NAME))
        note = f"Added {added} file{'s' if added != 1 else ''}."
        if skipped:
            note += f" Skipped {skipped} (not .docx)."
        if copies:
            # Pre-select likely duplicates ("x (1).docx") so one click on Remove drops them.
            self.tree.selection_set(copies)
            self.tree.see(copies[0])
            what = (
                "1 looks like a duplicate copy" if len(copies) == 1 else f"{len(copies)} look like duplicates"
            )
            note += f" {what} (selected) - click Remove to drop."
        self._refresh()
        if paths:
            self._set_status(note)

    def _is_output(self, path: Path) -> bool:
        out = self.output.get().strip()
        return bool(out) and Path(out).expanduser().resolve() == path.resolve()

    def _insert(self, path: Path) -> str:
        try:
            modified = datetime.fromtimestamp(path.stat().st_mtime).strftime("%Y-%m-%d %H:%M")
        except OSError:
            modified = ""
        iid = self.tree.insert("", "end", values=("", path.name, str(path.parent), modified))
        self.paths[iid] = path
        return iid

    def ordered_paths(self) -> list[Path]:
        return [self.paths[iid] for iid in self.tree.get_children()]

    def _reorder(self, paths: list[Path]) -> None:
        by_path = {id(p): iid for iid, p in self.paths.items()}
        for index, path in enumerate(paths):
            self.tree.move(by_path[id(path)], "", index)
        self._refresh()

    def sort(self) -> None:
        try:
            self._reorder(sort_files(self.ordered_paths(), SORT_LABELS[self.sort_choice.get()]))
        except OSError as exc:
            messagebox.showerror("Can't sort", str(exc), parent=self.root)

    def reverse(self) -> None:
        self._reorder(self.ordered_paths()[::-1])

    def move(self, step: int) -> None:
        selected = list(self.tree.selection())
        if not selected:
            return
        children = list(self.tree.get_children())
        indices = sorted((children.index(i) for i in selected), reverse=step > 0)
        for index in indices:
            target = index + step
            if 0 <= target < len(children) and children[target] not in selected:
                children[index], children[target] = children[target], children[index]
                self.tree.move(children[target], "", target)
        self.tree.see(selected[0])
        self._refresh()

    def remove_selected(self) -> None:
        for iid in self.tree.selection():
            self.tree.delete(iid)
            del self.paths[iid]
        self._refresh()

    def clear(self) -> None:
        self.tree.delete(*self.tree.get_children())
        self.paths.clear()
        if not self.output_chosen:
            self.output.set("")
        self._refresh()

    def _refresh(self) -> None:
        children = self.tree.get_children()
        for n, iid in enumerate(children, 1):
            self.tree.set(iid, "n", n)
        if children:
            self.empty_label.place_forget()
        else:
            self.empty_label.place(relx=0.5, rely=0.5, anchor="center")
        has_sel = bool(self.tree.selection())
        idle = "disabled" if self.busy else "!disabled"
        for btn, enabled in (
            (self.btn_up, has_sel),
            (self.btn_down, has_sel),
            (self.btn_remove, has_sel),
            (self.btn_clear, bool(children)),
        ):
            btn.state(["!disabled" if enabled and not self.busy else "disabled"])
        count = len(children)
        self.count_label.config(text=f"{count} file{'s' if count != 1 else ''}" if count else "")
        self.stitch_btn.config(text=f"Stitch {count} files" if count > 1 else "Stitch")
        self.stitch_btn.state([idle if count else "disabled"])

    # ---- drag and drop --------------------------------------------------------------------

    def _drag_start(self, event: tk.Event) -> None:
        self._drag_iid = self.tree.identify_row(event.y) or None

    def _drag_motion(self, event: tk.Event) -> None:
        target = self.tree.identify_row(event.y)
        if self._drag_iid and target and target != self._drag_iid and not self.busy:
            self.tree.move(self._drag_iid, "", self.tree.index(target))
            self._refresh()

    def _on_drop(self, event) -> str:
        self.add_paths(self.root.tk.splitlist(event.data))
        return event.action

    def _open_selected(self, event: tk.Event) -> None:
        iid = self.tree.identify_row(event.y)
        if iid:
            reveal(self.paths[iid], select=False)

    # ---- dialogs --------------------------------------------------------------------------

    def ask_folder(self) -> None:
        folder = filedialog.askdirectory(
            parent=self.root, title="Choose a folder of .docx files", mustexist=True
        )
        if folder:
            self.add_paths([folder])

    def ask_files(self) -> None:
        files = filedialog.askopenfilenames(
            parent=self.root, title="Choose .docx files", filetypes=[("Word documents", "*.docx")]
        )
        if files:
            self.add_paths(files)

    def ask_output(self) -> None:
        current = (
            Path(self.output.get()).expanduser() if self.output.get() else Path.home() / DEFAULT_OUTPUT_NAME
        )
        chosen = filedialog.asksaveasfilename(
            parent=self.root,
            title="Save merged document as",
            initialdir=str(current.parent),
            initialfile=current.name,
            defaultextension=".docx",
            filetypes=[("Word document", "*.docx")],
        )
        if chosen:
            self.output.set(chosen)
            self.output_chosen = True

    # ---- stitching ------------------------------------------------------------------------

    def start_stitch(self) -> None:
        files = self.ordered_paths()
        if self.busy or not files:
            return
        raw = self.output.get().strip()
        if not raw:
            self.ask_output()
            raw = self.output.get().strip()
            if not raw:
                return
        output = Path(raw).expanduser()
        if output.suffix.lower() != ".docx":
            output = output.with_name(output.name + ".docx")
        if output.exists():
            answer = messagebox.askyesnocancel(
                "File already exists",
                f"{output.name} already exists.\n\nYes: replace it\nNo: keep both (save as a new name)",
                parent=self.root,
            )
            if answer is None:
                return
            if answer is False:
                output = next_free_path(output)
        self.output.set(str(output))

        self.busy = True
        self.last_output = None
        self.open_btn.pack_forget()
        self.show_btn.pack_forget()
        self.progress.config(maximum=len(files), value=0)
        self.progress.pack(side="right", padx=12)
        self._refresh()
        threading.Thread(
            target=self._worker,
            args=(files, output, self.page_breaks.get(), self.keep_images.get()),
            daemon=True,
        ).start()
        self.root.after(50, self._poll)

    def _worker(self, files: list[Path], output: Path, page_breaks: bool, keep_images: bool) -> None:
        try:
            result = stitch(
                files,
                output,
                page_breaks=page_breaks,
                keep_images=keep_images,
                overwrite=True,
                on_progress=lambda done, total, path: self.events.put(("progress", done, total, path)),
            )
            self.events.put(("done", result))
        except StitchError as exc:
            self.events.put(("error", str(exc)))
        except Exception as exc:  # keep the app alive on unexpected failures
            self.events.put(("error", f"Unexpected error: {exc!r}"))

    def _poll(self) -> None:
        try:
            while True:
                kind, *data = self.events.get_nowait()
                if kind == "progress":
                    done, total, path = data
                    self.progress.config(value=done)
                    if done < total:
                        self._set_status(f"Merging {done + 1} of {total}: {path.name}")
                else:
                    self._finish(kind, data[0])
                    return
        except queue.Empty:
            pass
        self.root.after(50, self._poll)

    def _finish(self, kind: str, payload) -> None:
        self.busy = False
        self.progress.pack_forget()
        self._refresh()
        if kind == "done":
            self.last_output = payload
            self._set_status(f"Saved {payload.name}")
            self.show_btn.pack(side="right", padx=(0, 12))
            self.open_btn.pack(side="right", padx=(0, 6))
        else:
            self._set_status("Merge failed.")
            messagebox.showerror("Merge failed", payload, parent=self.root)

    def _reveal(self, *, select: bool) -> None:
        if self.last_output:
            reveal(self.last_output, select=select)

    def _set_status(self, text: str) -> None:
        self.status.config(text=text)


def _prepare_windows() -> None:
    """Crisp text on high-DPI screens and our own taskbar icon instead of Python's."""
    import ctypes

    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("DocxStitcher")
    except (AttributeError, OSError):
        pass


def run(initial: Sequence[Path] = ()) -> int:
    if sys.platform == "win32":
        _prepare_windows()
    dnd = TkinterDnD is not None
    try:
        root = TkinterDnD.Tk() if dnd else tk.Tk()
    except (tk.TclError, RuntimeError) as exc:
        if dnd and "tkdnd" in str(exc).lower():
            dnd, root = False, tk.Tk()
        else:
            message = (
                f"Couldn't start the desktop app: {exc}\n"
                "Try `uv python upgrade` (older uv Python builds can't find Tcl/Tk),\n"
                "or use the command line: docx-stitcher --help"
            )
            print(message, file=sys.stderr)
            if sys.platform == "win32":  # launched from a shortcut there is no console to print to
                import ctypes

                ctypes.windll.user32.MessageBoxW(None, message, "DOCX Stitcher", 0x10)
            return 1
    with contextlib.suppress(tk.TclError):
        root.iconphoto(True, tk.PhotoImage(file=str(Path(__file__).with_name("assets") / "icon.png")))
    StitcherApp(root, initial, dnd=dnd)
    root.lift()
    root.mainloop()
    return 0


def main() -> None:
    raise SystemExit(run([Path(arg) for arg in sys.argv[1:]]))
