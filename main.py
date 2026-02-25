"""
main.py – Excel Range Screenshot Manager
=========================================
A tkinter GUI for recording Excel file/sheet/range combinations and
capturing them as PNG screenshots for later use.

Usage:
    python main.py

Requirements:
    pip install pywin32 Pillow openpyxl
"""

from __future__ import annotations

import subprocess
import threading
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from typing import List, Optional

from capture import CaptureError, ExcelCapture
from data_store import DataStore, Entry

# ---------------------------------------------------------------------------
# Optional PIL for thumbnail preview
# ---------------------------------------------------------------------------
try:
    from PIL import Image, ImageTk

    HAS_PIL = True
except ImportError:
    HAS_PIL = False

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------
APP_TITLE = "Excel Range Screenshot Manager"
DATA_FILE = "entries.json"
SCREENSHOTS_DIR = "screenshots"
MIN_WIDTH, MIN_HEIGHT = 780, 480


# ===========================================================================
# Add / Edit dialog
# ===========================================================================


class EntryDialog(tk.Toplevel):
    """Modal dialog for creating or editing an Entry."""

    def __init__(self, parent: tk.Tk, entry: Optional[Entry] = None) -> None:
        super().__init__(parent)
        self.result: Optional[Entry] = None
        self._entry = entry

        self.title("Edit Entry" if entry else "Add Entry")
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()

        self._build()
        if entry:
            self._populate(entry)

        # Centre over parent
        self.update_idletasks()
        pw, ph = parent.winfo_width(), parent.winfo_height()
        px, py = parent.winfo_rootx(), parent.winfo_rooty()
        w, h = self.winfo_reqwidth(), self.winfo_reqheight()
        self.geometry(f"+{px + (pw - w) // 2}+{py + (ph - h) // 2}")

    # ------------------------------------------------------------------
    # Build
    # ------------------------------------------------------------------

    def _build(self) -> None:
        p = {"padx": 8, "pady": 5}
        frm = ttk.Frame(self, padding=16)
        frm.grid(sticky="nsew")
        self.columnconfigure(0, weight=1)

        # Row 0 – Name
        ttk.Label(frm, text="Name *").grid(row=0, column=0, sticky="w", **p)
        self._name_var = tk.StringVar()
        ttk.Entry(frm, textvariable=self._name_var, width=42).grid(
            row=0, column=1, columnspan=2, sticky="ew", **p
        )

        # Row 1 – Excel file
        ttk.Label(frm, text="Excel File *").grid(row=1, column=0, sticky="w", **p)
        self._file_var = tk.StringVar()
        self._file_entry = ttk.Entry(frm, textvariable=self._file_var, width=37)
        self._file_entry.grid(row=1, column=1, sticky="ew", **p)
        ttk.Button(frm, text="Browse…", width=9, command=self._browse).grid(
            row=1, column=2, **p
        )
        self._file_var.trace_add("write", lambda *_: self._schedule_sheet_refresh())

        # Row 2 – Sheet
        ttk.Label(frm, text="Sheet *").grid(row=2, column=0, sticky="w", **p)
        self._sheet_var = tk.StringVar()
        self._sheet_cb = ttk.Combobox(
            frm, textvariable=self._sheet_var, width=30, state="readonly"
        )
        self._sheet_cb.grid(row=2, column=1, sticky="ew", **p)
        self._refresh_btn = ttk.Button(
            frm, text="↻", width=3, command=self._load_sheets
        )
        self._refresh_btn.grid(row=2, column=2, sticky="w", padx=(0, 8), pady=5)

        # Row 3 – Range
        ttk.Label(frm, text="Cell Range *").grid(row=3, column=0, sticky="w", **p)
        self._range_var = tk.StringVar()
        ttk.Entry(frm, textvariable=self._range_var, width=20).grid(
            row=3, column=1, sticky="w", **p
        )
        ttk.Label(frm, text="e.g.  A1:H20", foreground="#888").grid(
            row=3, column=2, sticky="w", **p
        )

        # Row 4 – Notes
        ttk.Label(frm, text="Notes").grid(row=4, column=0, sticky="nw", **p)
        self._notes = tk.Text(frm, width=40, height=3, font=("TkDefaultFont", 9))
        self._notes.grid(row=4, column=1, columnspan=2, sticky="ew", **p)

        # Row 5 – Buttons
        btn_row = ttk.Frame(frm)
        btn_row.grid(row=5, column=0, columnspan=3, pady=(8, 0))
        ttk.Button(btn_row, text="OK", width=10, command=self._ok).pack(
            side="left", padx=4
        )
        ttk.Button(btn_row, text="Cancel", width=10, command=self.destroy).pack(
            side="left", padx=4
        )

        frm.columnconfigure(1, weight=1)
        self._refresh_after_id: Optional[str] = None

    def _populate(self, entry: Entry) -> None:
        self._name_var.set(entry.name)
        self._file_var.set(entry.file_path)
        self._range_var.set(entry.cell_range)
        self._notes.insert("1.0", entry.notes or "")
        # Load sheets then select the stored sheet
        self._load_sheets(preselect=entry.sheet_name)

    # ------------------------------------------------------------------
    # Events
    # ------------------------------------------------------------------

    def _browse(self) -> None:
        path = filedialog.askopenfilename(
            parent=self,
            title="Select Excel File",
            filetypes=[
                ("Excel files", "*.xlsx *.xlsm *.xls *.xlsb"),
                ("All files", "*.*"),
            ],
        )
        if path:
            self._file_var.set(path)

    def _schedule_sheet_refresh(self) -> None:
        """Debounce sheet refresh while the user types in the path field."""
        if self._refresh_after_id:
            self.after_cancel(self._refresh_after_id)
        self._refresh_after_id = self.after(400, self._load_sheets)

    def _load_sheets(self, preselect: str = "") -> None:
        path = self._file_var.get().strip()
        if not path or not Path(path).exists():
            return
        names = ExcelCapture.get_sheet_names(path)
        self._sheet_cb["values"] = names
        if names:
            target = preselect if preselect in names else names[0]
            self._sheet_var.set(target)

    # ------------------------------------------------------------------
    # Submit
    # ------------------------------------------------------------------

    def _ok(self) -> None:
        name = self._name_var.get().strip()
        file_path = self._file_var.get().strip()
        sheet = self._sheet_var.get().strip()
        cell_range = self._range_var.get().strip().upper()
        notes = self._notes.get("1.0", "end-1c").strip()

        # Validation
        errors: List[str] = []
        if not name:
            errors.append("• Name is required.")
        if not file_path:
            errors.append("• Excel file is required.")
        elif not Path(file_path).exists():
            errors.append(f"• File not found:\n  {file_path}")
        if not sheet:
            errors.append("• Sheet name is required.")
        if not cell_range:
            errors.append("• Cell range is required.")

        if errors:
            messagebox.showwarning("Validation", "\n".join(errors), parent=self)
            return

        if self._entry:
            self.result = Entry(
                id=self._entry.id,
                name=name,
                file_path=file_path,
                sheet_name=sheet,
                cell_range=cell_range,
                last_capture_path=self._entry.last_capture_path,
                last_capture_time=self._entry.last_capture_time,
                notes=notes,
            )
        else:
            self.result = Entry.create(name, file_path, sheet, cell_range, notes)

        self.destroy()


# ===========================================================================
# Main application window
# ===========================================================================


class MainApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("1050x600")
        self.minsize(MIN_WIDTH, MIN_HEIGHT)

        self._store = DataStore(DATA_FILE)
        self._capturer = ExcelCapture(SCREENSHOTS_DIR)
        self._preview_ref: Optional[object] = None  # prevent GC of PhotoImage

        self._build_ui()
        self._refresh_tree()
        self._set_status("Ready")

    # ------------------------------------------------------------------
    # UI construction
    # ------------------------------------------------------------------

    def _build_ui(self) -> None:
        self._build_toolbar()
        self._build_body()
        self._build_statusbar()

        # Keyboard shortcuts
        self.bind("<Control-n>", lambda _e: self._add())
        self.bind("<Delete>", lambda _e: self._delete())
        self.bind("<F5>", lambda _e: self._capture_selected())
        self.bind("<F2>", lambda _e: self._edit())

    # --- Toolbar -------------------------------------------------------

    def _build_toolbar(self) -> None:
        tb = ttk.Frame(self, relief="groove", padding=(6, 4))
        tb.pack(side="top", fill="x")

        def sep():
            ttk.Separator(tb, orient="vertical").pack(
                side="left", fill="y", padx=6, pady=2
            )

        def btn(label, cmd, tip=""):
            b = ttk.Button(tb, text=label, command=cmd)
            b.pack(side="left", padx=2)
            if tip:
                _tooltip(b, tip)
            return b

        btn("＋ Add", self._add, "Ctrl+N")
        btn("✎ Edit", self._edit, "F2 or double-click")
        btn("✕ Delete", self._delete, "Delete key")
        sep()
        self._cap_btn = btn("⬛ Capture", self._capture_selected, "F5")
        self._cap_all_btn = btn("⬛ Capture All", self._capture_all)
        sep()
        btn("📂 Open Folder", self._open_folder)

    # --- Body (paned) ---------------------------------------------------

    def _build_body(self) -> None:
        pane = ttk.PanedWindow(self, orient="horizontal")
        pane.pack(fill="both", expand=True, padx=4, pady=(0, 2))

        # Left: entry list
        left = ttk.Frame(pane)
        pane.add(left, weight=3)
        self._build_tree(left)

        # Right: preview + info
        right = ttk.Frame(pane, width=280)
        pane.add(right, weight=1)
        self._build_preview(right)

    def _build_tree(self, parent: ttk.Frame) -> None:
        cols = ("name", "file", "sheet", "range", "captured")
        self._tree = ttk.Treeview(
            parent, columns=cols, show="headings", selectmode="browse"
        )
        col_defs = [
            ("name",     "Name",          200, True),
            ("file",     "File",          200, True),
            ("sheet",    "Sheet",          90, False),
            ("range",    "Range",          90, False),
            ("captured", "Last Captured", 140, False),
        ]
        for cid, heading, width, stretch in col_defs:
            self._tree.heading(
                cid, text=heading, command=lambda c=cid: self._sort_tree(c)
            )
            self._tree.column(cid, width=width, stretch=stretch, minwidth=60)

        vsb = ttk.Scrollbar(parent, orient="vertical", command=self._tree.yview)
        hsb = ttk.Scrollbar(parent, orient="horizontal", command=self._tree.xview)
        self._tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self._tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        parent.rowconfigure(0, weight=1)
        parent.columnconfigure(0, weight=1)

        self._tree.bind("<<TreeviewSelect>>", self._on_tree_select)
        self._tree.bind("<Double-1>", lambda _e: self._edit())

        # Context menu
        self._ctx = tk.Menu(self, tearoff=False)
        self._ctx.add_command(label="Edit",        command=self._edit)
        self._ctx.add_command(label="Capture",     command=self._capture_selected)
        self._ctx.add_command(label="Open Screenshot", command=self._open_screenshot)
        self._ctx.add_separator()
        self._ctx.add_command(label="Delete",      command=self._delete)
        self._tree.bind("<Button-3>", self._show_ctx)

    def _build_preview(self, parent: ttk.Frame) -> None:
        ttk.Label(parent, text="Preview", font=("TkDefaultFont", 9, "bold")).pack(
            anchor="w", padx=6, pady=(6, 0)
        )
        ttk.Separator(parent, orient="horizontal").pack(fill="x", padx=6, pady=(2, 0))

        self._preview_lbl = ttk.Label(
            parent, text="No selection", anchor="center", relief="sunken"
        )
        self._preview_lbl.pack(fill="both", expand=True, padx=6, pady=6)

        # Info text box
        self._info = tk.Text(
            parent,
            height=7,
            state="disabled",
            font=("Consolas", 8),
            relief="flat",
            background="#f4f4f4",
            wrap="word",
        )
        self._info.pack(fill="x", padx=6, pady=(0, 6))

    # --- Status bar ----------------------------------------------------

    def _build_statusbar(self) -> None:
        bar = ttk.Frame(self, relief="sunken")
        bar.pack(side="bottom", fill="x")
        self._status_var = tk.StringVar()
        self._count_var = tk.StringVar()
        ttk.Label(bar, textvariable=self._status_var, anchor="w").pack(
            side="left", padx=6
        )
        ttk.Label(bar, textvariable=self._count_var, anchor="e").pack(
            side="right", padx=6
        )

    # ------------------------------------------------------------------
    # Tree helpers
    # ------------------------------------------------------------------

    def _refresh_tree(self, select_id: str = "") -> None:
        self._tree.delete(*self._tree.get_children())
        entries = self._store.entries
        for e in entries:
            self._tree.insert(
                "",
                "end",
                iid=e.id,
                values=(
                    e.name,
                    e.display_file(),
                    e.sheet_name,
                    e.cell_range,
                    e.display_capture_time(),
                ),
            )
        n = len(entries)
        self._count_var.set(f"{n} entr{'y' if n == 1 else 'ies'}")
        if select_id:
            try:
                self._tree.selection_set(select_id)
                self._tree.see(select_id)
            except Exception:
                pass

    def _selected_id(self) -> Optional[str]:
        sel = self._tree.selection()
        return sel[0] if sel else None

    def _sort_tree(self, col: str) -> None:
        items = [
            (self._tree.set(iid, col).lower(), iid)
            for iid in self._tree.get_children()
        ]
        items.sort()
        for idx, (_, iid) in enumerate(items):
            self._tree.move(iid, "", idx)

    # ------------------------------------------------------------------
    # CRUD
    # ------------------------------------------------------------------

    def _add(self) -> None:
        dlg = EntryDialog(self)
        self.wait_window(dlg)
        if dlg.result:
            self._store.add(dlg.result)
            self._refresh_tree(select_id=dlg.result.id)
            self._set_status(f"Added: {dlg.result.name}")

    def _edit(self) -> None:
        eid = self._selected_id()
        if not eid:
            messagebox.showinfo("Edit", "Select an entry first.", parent=self)
            return
        entry = self._store.get(eid)
        if not entry:
            return
        dlg = EntryDialog(self, entry=entry)
        self.wait_window(dlg)
        if dlg.result:
            self._store.update(dlg.result)
            self._refresh_tree(select_id=dlg.result.id)
            self._set_status(f"Updated: {dlg.result.name}")

    def _delete(self) -> None:
        eid = self._selected_id()
        if not eid:
            return
        entry = self._store.get(eid)
        if not entry:
            return
        if messagebox.askyesno(
            "Confirm Delete",
            f"Delete entry '{entry.name}'?",
            icon="warning",
            parent=self,
        ):
            self._store.delete(eid)
            self._refresh_tree()
            self._clear_preview()
            self._set_status(f"Deleted: {entry.name}")

    # ------------------------------------------------------------------
    # Capture
    # ------------------------------------------------------------------

    def _capture_selected(self) -> None:
        eid = self._selected_id()
        if not eid:
            messagebox.showinfo("Capture", "Select an entry first.", parent=self)
            return
        entry = self._store.get(eid)
        if entry:
            self._run_capture([entry])

    def _capture_all(self) -> None:
        entries = self._store.entries
        if not entries:
            messagebox.showinfo("Capture All", "No entries saved yet.", parent=self)
            return
        self._run_capture(entries)

    def _run_capture(self, entries: List[Entry]) -> None:
        n = len(entries)
        self._set_status(f"Capturing {n} item{'s' if n > 1 else ''}…")
        self._cap_btn.config(state="disabled")
        self._cap_all_btn.config(state="disabled")

        def worker() -> None:
            results = []
            for e in entries:
                try:
                    path = self._capturer.capture(
                        e.file_path, e.sheet_name, e.cell_range, e.name
                    )
                    e.last_capture_path = path
                    e.last_capture_time = datetime.now().strftime("%Y-%m-%d %H:%M")
                    self._store.update(e)
                    results.append((e, path, None))
                except CaptureError as exc:
                    results.append((e, None, str(exc)))
                except Exception as exc:
                    results.append((e, None, f"Unexpected error: {exc}"))
            self.after(0, lambda: self._capture_done(results))

        threading.Thread(target=worker, daemon=True).start()

    def _capture_done(self, results: list) -> None:
        self._cap_btn.config(state="normal")
        self._cap_all_btn.config(state="normal")
        self._refresh_tree(select_id=self._selected_id())
        self._on_tree_select()  # refresh preview

        errors = [(e.name, msg) for e, _, msg in results if msg]
        ok_count = sum(1 for _, p, _ in results if p)

        if errors:
            detail = "\n\n".join(f"• {n}:\n  {m}" for n, m in errors)
            messagebox.showerror(
                "Capture Errors",
                f"{len(errors)} error(s):\n\n{detail}",
                parent=self,
            )
            self._set_status(
                f"Captured {ok_count}, failed {len(errors)}"
            )
        else:
            self._set_status(f"Captured {ok_count} item{'s' if ok_count != 1 else ''} successfully")

    # ------------------------------------------------------------------
    # Preview panel
    # ------------------------------------------------------------------

    def _on_tree_select(self, _event=None) -> None:
        eid = self._selected_id()
        if not eid:
            self._clear_preview()
            return
        entry = self._store.get(eid)
        if not entry:
            return
        self._update_info(entry)
        cap = entry.last_capture_path
        if cap and Path(cap).exists():
            self._show_preview(cap)
        else:
            self._preview_lbl.config(image="", text="No screenshot yet.\nPress F5 to capture.")
            self._preview_ref = None

    def _show_preview(self, path: str) -> None:
        if not HAS_PIL:
            self._preview_lbl.config(text="Install Pillow\nfor image preview.")
            return
        try:
            w = max(self._preview_lbl.winfo_width() - 4, 200)
            h = max(self._preview_lbl.winfo_height() - 4, 200)
            img = Image.open(path)
            img.thumbnail((w, h), Image.LANCZOS)
            photo = ImageTk.PhotoImage(img)
            self._preview_lbl.config(image=photo, text="")
            self._preview_ref = photo  # prevent GC
        except Exception as exc:
            self._preview_lbl.config(image="", text=f"Preview error:\n{exc}")
            self._preview_ref = None

    def _clear_preview(self) -> None:
        self._preview_lbl.config(image="", text="No selection")
        self._preview_ref = None
        self._info.config(state="normal")
        self._info.delete("1.0", "end")
        self._info.config(state="disabled")

    def _update_info(self, entry: Entry) -> None:
        self._info.config(state="normal")
        self._info.delete("1.0", "end")
        lines = [
            f"Name:  {entry.name}",
            f"File:  {entry.display_file()}",
            f"Sheet: {entry.sheet_name}",
            f"Range: {entry.cell_range}",
            f"Cap:   {entry.display_capture_time()}",
        ]
        if entry.notes:
            lines.append(f"Notes: {entry.notes}")
        self._info.insert("1.0", "\n".join(lines))
        self._info.config(state="disabled")

    # ------------------------------------------------------------------
    # Misc
    # ------------------------------------------------------------------

    def _open_folder(self) -> None:
        folder = Path(SCREENSHOTS_DIR).resolve()
        folder.mkdir(parents=True, exist_ok=True)
        subprocess.Popen(["explorer", str(folder)])

    def _open_screenshot(self) -> None:
        eid = self._selected_id()
        if not eid:
            return
        entry = self._store.get(eid)
        if entry and entry.last_capture_path and Path(entry.last_capture_path).exists():
            subprocess.Popen(["explorer", str(Path(entry.last_capture_path).resolve())])
        else:
            messagebox.showinfo("Open", "No screenshot available for this entry.", parent=self)

    def _show_ctx(self, event: tk.Event) -> None:
        iid = self._tree.identify_row(event.y)
        if iid:
            self._tree.selection_set(iid)
            self._ctx.post(event.x_root, event.y_root)

    def _set_status(self, msg: str) -> None:
        self._status_var.set(msg)


# ===========================================================================
# Tooltip helper (no external dependency)
# ===========================================================================


def _tooltip(widget: tk.Widget, text: str) -> None:
    tip_win: list = [None]  # mutable container for closure

    def show(_event: tk.Event) -> None:
        if tip_win[0]:
            return
        x = widget.winfo_rootx() + 20
        y = widget.winfo_rooty() + widget.winfo_height() + 4
        win = tk.Toplevel(widget)
        win.wm_overrideredirect(True)
        win.wm_geometry(f"+{x}+{y}")
        ttk.Label(
            win,
            text=text,
            background="#ffffcc",
            relief="solid",
            borderwidth=1,
            font=("TkDefaultFont", 8),
        ).pack()
        tip_win[0] = win

    def hide(_event: tk.Event) -> None:
        if tip_win[0]:
            tip_win[0].destroy()
            tip_win[0] = None

    widget.bind("<Enter>", show)
    widget.bind("<Leave>", hide)


# ===========================================================================
# Entry point
# ===========================================================================

if __name__ == "__main__":
    app = MainApp()
    app.mainloop()
