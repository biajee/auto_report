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

import queue as _queue
import subprocess
import threading
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from typing import List, Optional

from capture import CaptureError, ExcelCapture
from data_store import DataStore, Entry
from ppt_export import PPTExportError, PPTExporter

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
        range_frm = ttk.Frame(frm)
        range_frm.grid(row=3, column=1, sticky="ew", padx=8, pady=5)
        ttk.Entry(range_frm, textvariable=self._range_var, width=16).pack(side="left")
        ttk.Label(range_frm, text="e.g. A1:H20", foreground="#888").pack(
            side="left", padx=(6, 0)
        )
        ttk.Button(
            frm, text="📍 Pick in Excel", command=self._pick_range
        ).grid(row=3, column=2, padx=8, pady=5)

        # Row 4 – Notes
        ttk.Label(frm, text="Notes").grid(row=4, column=0, sticky="nw", **p)
        self._notes = tk.Text(frm, width=40, height=3, font=("TkDefaultFont", 9))
        self._notes.grid(row=4, column=1, columnspan=2, sticky="ew", **p)

        # Row 5 – PowerPoint Destination
        ppt_frm = ttk.LabelFrame(
            frm, text=" PowerPoint Destination (optional) ", padding=(8, 6)
        )
        ppt_frm.grid(row=5, column=0, columnspan=3, sticky="ew", padx=8, pady=(6, 0))

        pp = {"padx": 4, "pady": 3}

        # PPT file
        ttk.Label(ppt_frm, text="File:").grid(row=0, column=0, sticky="w", **pp)
        self._pptx_file_var = tk.StringVar()
        ttk.Entry(ppt_frm, textvariable=self._pptx_file_var, width=30).grid(
            row=0, column=1, columnspan=2, sticky="ew", **pp
        )
        ttk.Button(ppt_frm, text="Browse…", width=8, command=self._browse_ppt).grid(
            row=0, column=3, **pp
        )
        self._pptx_info_var = tk.StringVar(value="")
        ttk.Label(
            ppt_frm, textvariable=self._pptx_info_var,
            foreground="#666", font=("TkDefaultFont", 8),
        ).grid(row=0, column=4, sticky="w", **pp)
        self._pptx_file_var.trace_add("write", lambda *_: self._on_pptx_changed())

        # Slide number + pick button
        ttk.Label(ppt_frm, text="Slide:").grid(row=1, column=0, sticky="w", **pp)
        self._pptx_slide_var = tk.StringVar(value="1")
        ttk.Spinbox(
            ppt_frm, textvariable=self._pptx_slide_var,
            from_=1, to=999, width=6,
        ).grid(row=1, column=1, sticky="w", **pp)
        ttk.Button(
            ppt_frm, text="📍 Pick Position in PowerPoint",
            command=self._pick_ppt_pos,
        ).grid(row=1, column=2, columnspan=3, sticky="w", **pp)

        # Position + size on one row
        ps_frm = ttk.Frame(ppt_frm)
        ps_frm.grid(row=2, column=0, columnspan=5, sticky="w", **pp)

        def _num_field(parent, label, var_name, default):
            ttk.Label(parent, text=label).pack(side="left")
            var = tk.StringVar(value=str(default))
            setattr(self, var_name, var)
            ttk.Entry(parent, textvariable=var, width=6).pack(side="left", padx=(2, 0))
            ttk.Label(parent, text="in").pack(side="left", padx=(2, 10))

        _num_field(ps_frm, "Left:",   "_pptx_left_var",   0.5)
        _num_field(ps_frm, "Top:",    "_pptx_top_var",    0.5)
        _num_field(ps_frm, "Width:",  "_pptx_width_var",  9.0)
        _num_field(ps_frm, "Height:", "_pptx_height_var", 6.5)

        ppt_frm.columnconfigure(1, weight=1)

        # Row 6 – Buttons
        btn_row = ttk.Frame(frm)
        btn_row.grid(row=6, column=0, columnspan=3, pady=(10, 0))
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
        # PPT fields
        self._pptx_file_var.set(entry.pptx_file or "")
        self._pptx_slide_var.set(str(entry.pptx_slide))
        self._pptx_left_var.set(str(entry.pptx_left))
        self._pptx_top_var.set(str(entry.pptx_top))
        self._pptx_width_var.set(str(entry.pptx_width))
        self._pptx_height_var.set(str(entry.pptx_height))
        if entry.pptx_file:
            self._on_pptx_changed()

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
    # PPT helpers
    # ------------------------------------------------------------------

    def _browse_ppt(self) -> None:
        path = filedialog.askopenfilename(
            parent=self,
            title="Select PowerPoint File",
            filetypes=[
                ("PowerPoint files", "*.pptx *.ppt"),
                ("All files", "*.*"),
            ],
        )
        if path:
            self._pptx_file_var.set(path)

    def _on_pptx_changed(self) -> None:
        path = self._pptx_file_var.get().strip()
        if not path or not Path(path).exists():
            self._pptx_info_var.set("")
            return

        def load() -> None:
            count, w, h = PPTExporter.get_slide_info(path)
            info = f"{count} slide{'s' if count != 1 else ''}  ({w} × {h} in)"
            self.after(0, lambda: self._pptx_info_var.set(info))

        threading.Thread(target=load, daemon=True).start()

    @staticmethod
    def _f(var: tk.StringVar, default: float) -> float:
        try:
            return max(0.0, float(var.get()))
        except ValueError:
            return default

    @staticmethod
    def _i(var: tk.StringVar, default: int) -> int:
        try:
            return max(1, int(var.get()))
        except ValueError:
            return default

    # ------------------------------------------------------------------
    # Range picker
    # ------------------------------------------------------------------

    def _pick_range(self) -> None:
        file_path = self._file_var.get().strip()
        if not file_path:
            messagebox.showwarning("Pick Range", "Select an Excel file first.", parent=self)
            return
        if not Path(file_path).exists():
            messagebox.showwarning(
                "Pick Range", f"File not found:\n{file_path}", parent=self
            )
            return
        # Release grab so user can freely click in Excel
        self.grab_release()
        try:
            dlg = RangePickerDialog(self, file_path, initial_sheet=self._sheet_var.get())
            self.wait_window(dlg)
            if dlg.result_range:
                self._range_var.set(dlg.result_range)
            if dlg.result_sheet:
                sheets = list(self._sheet_cb["values"])
                if dlg.result_sheet not in sheets:
                    self._load_sheets(preselect=dlg.result_sheet)
                else:
                    self._sheet_var.set(dlg.result_sheet)
        finally:
            self.grab_set()

    def _pick_ppt_pos(self) -> None:
        pptx_file = self._pptx_file_var.get().strip()
        if not pptx_file:
            messagebox.showwarning(
                "Pick Position", "Select a PowerPoint file first.", parent=self
            )
            return
        if not Path(pptx_file).exists():
            messagebox.showwarning(
                "Pick Position", f"File not found:\n{pptx_file}", parent=self
            )
            return
        try:
            slide_num = max(1, int(self._pptx_slide_var.get() or 1))
        except ValueError:
            slide_num = 1

        self.grab_release()
        try:
            dlg = PPTPosPickerDialog(
                self,
                pptx_file,
                slide_number=slide_num,
                init_left=self._f(self._pptx_left_var, 0.5),
                init_top=self._f(self._pptx_top_var, 0.5),
                init_width=self._f(self._pptx_width_var, 9.0),
                init_height=self._f(self._pptx_height_var, 6.5),
            )
            self.wait_window(dlg)
            if dlg.result is not None:
                l, t, w, h = dlg.result
                self._pptx_left_var.set(f"{l:.2f}")
                self._pptx_top_var.set(f"{t:.2f}")
                self._pptx_width_var.set(f"{w:.2f}")
                self._pptx_height_var.set(f"{h:.2f}")
        finally:
            self.grab_set()

    # ------------------------------------------------------------------
    # Submit
    # ------------------------------------------------------------------

    def _ok(self) -> None:
        name = self._name_var.get().strip()
        file_path = self._file_var.get().strip()
        sheet = self._sheet_var.get().strip()
        cell_range = self._range_var.get().strip().upper()
        notes = self._notes.get("1.0", "end-1c").strip()

        # PPT fields
        pptx_file = self._pptx_file_var.get().strip() or None
        pptx_slide = self._i(self._pptx_slide_var, 1)
        pptx_left   = self._f(self._pptx_left_var,   0.5)
        pptx_top    = self._f(self._pptx_top_var,    0.5)
        pptx_width  = self._f(self._pptx_width_var,  9.0)
        pptx_height = self._f(self._pptx_height_var, 6.5)

        # Validation
        errors: List[str] = []
        if not name:
            errors.append("• Name is required.")
        if not file_path:
            errors.append("• Excel file is required.")
        elif not Path(file_path).exists():
            errors.append(f"• Excel file not found:\n  {file_path}")
        if not sheet:
            errors.append("• Sheet name is required.")
        if not cell_range:
            errors.append("• Cell range is required.")
        if pptx_file and not Path(pptx_file).exists():
            errors.append(f"• PowerPoint file not found:\n  {pptx_file}")

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
                pptx_file=pptx_file,
                pptx_slide=pptx_slide,
                pptx_left=pptx_left,
                pptx_top=pptx_top,
                pptx_width=pptx_width,
                pptx_height=pptx_height,
            )
        else:
            e = Entry.create(name, file_path, sheet, cell_range, notes)
            e.pptx_file   = pptx_file
            e.pptx_slide  = pptx_slide
            e.pptx_left   = pptx_left
            e.pptx_top    = pptx_top
            e.pptx_width  = pptx_width
            e.pptx_height = pptx_height
            self.result = e

        self.destroy()


# ===========================================================================
# Range picker dialog – opens Excel, shows live selection, returns range
# ===========================================================================


class RangePickerDialog(tk.Toplevel):
    """
    Opens the Excel file visibly and polls the active selection every 400 ms.
    The user clicks a cell range in Excel, then clicks Confirm here.

    result_sheet and result_range are set on confirm, None on cancel.
    """

    def __init__(self, parent, file_path: str, initial_sheet: str = "") -> None:
        super().__init__(parent)
        self.title("Pick Range in Excel")
        self.resizable(False, False)
        self.transient(parent)
        self.attributes("-topmost", True)   # float above Excel

        self.result_sheet: Optional[str] = None
        self.result_range: Optional[str] = None

        self._file_path = file_path
        self._initial_sheet = initial_sheet
        self._last_sel: tuple = ()          # (sheet, range) last seen
        self._stop = threading.Event()
        self._q: _queue.Queue = _queue.Queue()
        self._poll_id: Optional[str] = None

        self._build()
        self.protocol("WM_DELETE_WINDOW", self._cancel)

        # Position: top-right so it doesn't cover the spreadsheet
        self.update_idletasks()
        sw = self.winfo_screenwidth()
        w = self.winfo_reqwidth()
        self.geometry(f"+{sw - w - 30}+30")

        threading.Thread(target=self._worker, daemon=True).start()
        self._drain()

    # ------------------------------------------------------------------
    # UI
    # ------------------------------------------------------------------

    def _build(self) -> None:
        frm = ttk.Frame(self, padding=16)
        frm.pack(fill="both", expand=True)

        ttk.Label(
            frm,
            text="Switch to Excel and select a range.\nThen click Confirm.",
            justify="left",
            font=("TkDefaultFont", 9),
        ).pack(anchor="w", pady=(0, 10))

        ttk.Label(frm, text="Current selection:").pack(anchor="w")
        self._sel_var = tk.StringVar(value="opening Excel…")
        ttk.Label(
            frm,
            textvariable=self._sel_var,
            font=("Consolas", 11, "bold"),
            foreground="#0055aa",
            relief="sunken",
            padding=(10, 5),
            width=28,
            anchor="center",
        ).pack(fill="x", pady=(3, 14))

        btn_row = ttk.Frame(frm)
        btn_row.pack()
        self._confirm_btn = ttk.Button(
            btn_row, text="✓  Confirm", width=13,
            command=self._confirm, state="disabled",
        )
        self._confirm_btn.pack(side="left", padx=4)
        ttk.Button(btn_row, text="Cancel", width=10, command=self._cancel).pack(
            side="left", padx=4
        )

        self._status_var = tk.StringVar(value="Opening Excel…")
        ttk.Label(
            frm, textvariable=self._status_var,
            foreground="#999", font=("TkDefaultFont", 8),
        ).pack(pady=(12, 0))

    # ------------------------------------------------------------------
    # Background worker (all COM calls live here)
    # ------------------------------------------------------------------

    def _worker(self) -> None:
        try:
            import pythoncom
            pythoncom.CoInitialize()
        except Exception:
            pass

        xl = None
        wb = None
        try:
            try:
                import win32com.client
            except ImportError:
                self._q.put(("err", "pywin32 not installed.\nRun: pip install pywin32"))
                return

            xl = win32com.client.DispatchEx("Excel.Application")
            xl.Visible = True
            xl.DisplayAlerts = False
            try:
                xl.WindowState = -4137   # xlMaximized
            except Exception:
                pass

            abs_path = str(Path(self._file_path).resolve())
            wb = xl.Workbooks.Open(abs_path, ReadOnly=True)

            if self._initial_sheet:
                try:
                    wb.Worksheets(self._initial_sheet).Activate()
                except Exception:
                    pass

            # Bring Excel window to front
            try:
                xl.ActiveWindow.Activate()
            except Exception:
                pass

            self._q.put(("ready", None))

            # Poll loop
            while not self._stop.is_set():
                try:
                    addr = xl.Selection.Address.replace("$", "")
                    sheet = xl.ActiveSheet.Name
                    self._q.put(("sel", (sheet, addr)))
                except Exception:
                    pass
                self._stop.wait(0.4)

        except Exception as exc:
            self._q.put(("err", str(exc)))
        finally:
            self._q.put(("done", None))
            try:
                if wb:
                    wb.Close(SaveChanges=False)
            except Exception:
                pass
            try:
                if xl:
                    xl.Quit()
            except Exception:
                pass
            try:
                import pythoncom
                pythoncom.CoUninitialize()
            except Exception:
                pass

    # ------------------------------------------------------------------
    # Main-thread queue drain
    # ------------------------------------------------------------------

    def _drain(self) -> None:
        try:
            while True:
                kind, data = self._q.get_nowait()
                if kind == "ready":
                    self._status_var.set("Select a range in Excel, then click Confirm.")
                elif kind == "sel":
                    sheet, addr = data
                    self._sel_var.set(f"{sheet}!{addr}")
                    self._last_sel = (sheet, addr)
                    self._confirm_btn.config(state="normal")
                elif kind == "err":
                    messagebox.showerror("Excel Error", data, parent=self)
                    self._cleanup()
                    return
                elif kind == "done":
                    return
        except _queue.Empty:
            pass
        self._poll_id = self.after(200, self._drain)

    # ------------------------------------------------------------------
    # Actions
    # ------------------------------------------------------------------

    def _confirm(self) -> None:
        if self._last_sel:
            self.result_sheet, self.result_range = self._last_sel
        self._cleanup()

    def _cancel(self) -> None:
        self._cleanup()

    def _cleanup(self) -> None:
        self._stop.set()
        if self._poll_id:
            self.after_cancel(self._poll_id)
            self._poll_id = None
        self.destroy()


# ===========================================================================
# PPT position picker – open PowerPoint with a draggable placeholder shape
# ===========================================================================


class PPTPosPickerDialog(tk.Toplevel):
    """
    Opens the PPTX in a visible PowerPoint window, inserts an orange
    placeholder rectangle on the target slide, and polls its position
    every 400 ms.  When the user clicks Confirm the (left, top, width,
    height) tuple (in inches) is stored in self.result.
    """

    _PTS = 72.0  # points per inch (PowerPoint uses points internally)

    def __init__(
        self,
        parent,
        pptx_path: str,
        slide_number: int = 1,
        init_left: float = 0.5,
        init_top: float = 0.5,
        init_width: float = 9.0,
        init_height: float = 6.5,
    ) -> None:
        super().__init__(parent)
        self.title("Pick Position in PowerPoint")
        self.resizable(False, False)
        self.transient(parent)
        self.attributes("-topmost", True)

        self.result: Optional[tuple] = None   # (left, top, width, height) inches

        self._pptx_path = pptx_path
        self._slide_number = slide_number
        self._init = (init_left, init_top, init_width, init_height)
        self._last: tuple = (init_left, init_top, init_width, init_height)
        self._stop = threading.Event()
        self._q: _queue.Queue = _queue.Queue()
        self._poll_id: Optional[str] = None

        self._build()
        self.protocol("WM_DELETE_WINDOW", self._cancel)

        # Top-right corner so it floats above PowerPoint
        self.update_idletasks()
        sw = self.winfo_screenwidth()
        w = self.winfo_reqwidth()
        self.geometry(f"+{sw - w - 30}+30")

        threading.Thread(target=self._worker, daemon=True).start()
        self._drain()

    # ------------------------------------------------------------------
    # UI
    # ------------------------------------------------------------------

    def _build(self) -> None:
        frm = ttk.Frame(self, padding=16)
        frm.pack(fill="both", expand=True)

        ttk.Label(
            frm,
            text="Move and resize the orange rectangle\nin PowerPoint, then click Confirm.",
            justify="left",
            font=("TkDefaultFont", 9),
        ).pack(anchor="w", pady=(0, 10))

        # Live position display
        grid = ttk.Frame(frm, relief="sunken", padding=10)
        grid.pack(fill="x", pady=(0, 12))

        lbl_font = ("Consolas", 10)
        val_font = ("Consolas", 11, "bold")

        self._left_var   = tk.StringVar(value="—")
        self._top_var    = tk.StringVar(value="—")
        self._width_var  = tk.StringVar(value="—")
        self._height_var = tk.StringVar(value="—")

        for row, (label, var) in enumerate([
            ("Left:",   self._left_var),
            ("Top:",    self._top_var),
            ("Width:",  self._width_var),
            ("Height:", self._height_var),
        ]):
            ttk.Label(grid, text=label, font=lbl_font, width=8, anchor="e").grid(
                row=row // 2, column=(row % 2) * 2, sticky="e", padx=(4, 2), pady=3
            )
            ttk.Label(grid, textvariable=var, font=val_font,
                      foreground="#0055aa", width=10).grid(
                row=row // 2, column=(row % 2) * 2 + 1, sticky="w", pady=3
            )

        # Buttons
        btn_row = ttk.Frame(frm)
        btn_row.pack()
        self._confirm_btn = ttk.Button(
            btn_row, text="✓  Confirm", width=13,
            command=self._confirm, state="disabled",
        )
        self._confirm_btn.pack(side="left", padx=4)
        ttk.Button(btn_row, text="Cancel", width=10, command=self._cancel).pack(
            side="left", padx=4
        )

        self._status_var = tk.StringVar(value="Opening PowerPoint…")
        ttk.Label(
            frm, textvariable=self._status_var,
            foreground="#999", font=("TkDefaultFont", 8),
        ).pack(pady=(12, 0))

    # ------------------------------------------------------------------
    # Background worker (all COM calls)
    # ------------------------------------------------------------------

    def _worker(self) -> None:
        try:
            import pythoncom
            pythoncom.CoInitialize()
        except Exception:
            pass

        ppt = None
        prs = None
        shape = None

        try:
            try:
                import win32com.client
            except ImportError:
                self._q.put(("err", "pywin32 not installed.\nRun: pip install pywin32"))
                return

            ppt = win32com.client.DispatchEx("PowerPoint.Application")
            ppt.Visible = True
            ppt.DisplayAlerts = 0  # ppAlertsNone – suppress save prompts

            abs_path = str(Path(self._pptx_path).resolve())
            # Untitled=True opens as an unnamed copy – original file is never modified
            prs = ppt.Presentations.Open(abs_path, ReadOnly=False, Untitled=True)

            n = prs.Slides.Count
            sn = min(max(1, self._slide_number), n)
            slide = prs.Slides(sn)
            prs.Windows(1).View.GotoSlide(sn)

            # Bring PowerPoint to front
            try:
                ppt.ActiveWindow.Activate()
            except Exception:
                pass

            # Insert placeholder rectangle (positions in points)
            P = self._PTS
            l, t, w, h = self._init
            shape = slide.Shapes.AddShape(1, l * P, t * P, w * P, h * P)
            shape.Name = "_img_placeholder_"

            # Style: orange border, transparent fill, centred label
            shape.Line.ForeColor.RGB = 255 + 140 * 256        # orange
            shape.Line.Weight = 2.5
            shape.Line.DashStyle = 2                           # dashed
            shape.Fill.ForeColor.RGB = 255 + 165 * 256        # light orange
            shape.Fill.Transparency = 0.75

            tf = shape.TextFrame
            tf.TextRange.Text = "IMAGE PLACEHOLDER\nMove & resize to set position"
            tf.TextRange.Font.Size = 12
            tf.TextRange.Font.Bold = True
            tf.TextRange.Font.Color.RGB = 255 + 69 * 256      # orange-red
            tf.VerticalAnchor = 3                              # msoAnchorMiddle
            tf.TextRange.ParagraphFormat.Alignment = 2        # ppAlignCenter

            self._q.put(("ready", None))

            # Poll loop
            while not self._stop.is_set():
                try:
                    pos = (
                        round(shape.Left  / P, 3),
                        round(shape.Top   / P, 3),
                        round(shape.Width / P, 3),
                        round(shape.Height/ P, 3),
                    )
                    self._q.put(("pos", pos))
                except Exception:
                    pass
                self._stop.wait(0.4)

        except Exception as exc:
            self._q.put(("err", str(exc)))
        finally:
            self._q.put(("done", None))
            # Clean up: delete placeholder, close copy without saving
            try:
                if shape:
                    shape.Delete()
            except Exception:
                pass
            try:
                if prs:
                    prs.Close()
            except Exception:
                pass
            try:
                if ppt:
                    ppt.Quit()
            except Exception:
                pass
            try:
                import pythoncom
                pythoncom.CoUninitialize()
            except Exception:
                pass

    # ------------------------------------------------------------------
    # Main-thread queue drain
    # ------------------------------------------------------------------

    def _drain(self) -> None:
        try:
            while True:
                kind, data = self._q.get_nowait()
                if kind == "ready":
                    self._status_var.set(
                        f"Slide {self._slide_number}  —  move & resize the orange rectangle"
                    )
                elif kind == "pos":
                    l, t, w, h = data
                    self._last = data
                    self._left_var.set(f"{l:.2f} in")
                    self._top_var.set(f"{t:.2f} in")
                    self._width_var.set(f"{w:.2f} in")
                    self._height_var.set(f"{h:.2f} in")
                    self._confirm_btn.config(state="normal")
                elif kind == "err":
                    messagebox.showerror("PowerPoint Error", data, parent=self)
                    self._cleanup()
                    return
                elif kind == "done":
                    return
        except _queue.Empty:
            pass
        self._poll_id = self.after(200, self._drain)

    # ------------------------------------------------------------------
    # Actions
    # ------------------------------------------------------------------

    def _confirm(self) -> None:
        self.result = self._last
        self._cleanup()

    def _cancel(self) -> None:
        self._cleanup()

    def _cleanup(self) -> None:
        self._stop.set()
        if self._poll_id:
            self.after_cancel(self._poll_id)
            self._poll_id = None
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
        self._paste_btn = btn("📊 Paste to PPT", self._paste_selected)
        self._paste_all_btn = btn("📊 Paste All to PPT", self._paste_all)
        sep()
        self._run_all_btn = btn("▶ Run All", self._run_all, "Capture all + Paste all to PPT")
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
        cols = ("name", "file", "sheet", "range", "ppt", "captured")
        self._tree = ttk.Treeview(
            parent, columns=cols, show="headings", selectmode="browse"
        )
        col_defs = [
            ("name",     "Name",          180, True),
            ("file",     "File",          170, True),
            ("sheet",    "Sheet",          80, False),
            ("range",    "Range",          80, False),
            ("ppt",      "PPT Dest",      160, False),
            ("captured", "Last Captured", 130, False),
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
                    e.display_ppt_dest(),
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
        self._set_action_btns(False)

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
        self._set_action_btns(True)
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
    # PPT export
    # ------------------------------------------------------------------

    def _paste_selected(self) -> None:
        eid = self._selected_id()
        if not eid:
            messagebox.showinfo("Paste to PPT", "Select an entry first.", parent=self)
            return
        entry = self._store.get(eid)
        if not entry:
            return
        if not entry.has_ppt_dest():
            messagebox.showinfo(
                "Paste to PPT",
                f"'{entry.name}' has no PowerPoint destination.\n"
                "Edit the entry to add one.",
                parent=self,
            )
            return
        if not entry.last_capture_path or not Path(entry.last_capture_path).exists():
            if messagebox.askyesno(
                "No Screenshot",
                "No screenshot exists yet for this entry.\nCapture it first?",
                parent=self,
            ):
                self._run_capture_then_paste([entry])
            return
        self._run_paste([entry])

    def _paste_all(self) -> None:
        entries = [e for e in self._store.entries if e.has_ppt_dest()]
        if not entries:
            messagebox.showinfo(
                "Paste All to PPT",
                "No entries have a PowerPoint destination set.\n"
                "Edit entries to add one.",
                parent=self,
            )
            return
        missing = [
            e for e in entries
            if not e.last_capture_path or not Path(e.last_capture_path).exists()
        ]
        if missing:
            names = "\n".join(f"  • {e.name}" for e in missing)
            if messagebox.askyesno(
                "Missing Screenshots",
                f"These entries have no screenshot yet:\n{names}\n\n"
                "Capture all missing ones first, then paste?",
                parent=self,
            ):
                self._run_capture_then_paste(entries)
                return
        self._run_paste(entries)

    def _run_all(self) -> None:
        """Capture every entry, then paste those with a PPT destination."""
        entries = self._store.entries
        if not entries:
            messagebox.showinfo("Run All", "No entries to process.", parent=self)
            return
        self._run_capture_then_paste(entries)

    # ------------------------------------------------------------------
    # Background workers
    # ------------------------------------------------------------------

    def _run_capture_then_paste(self, entries: list) -> None:
        n = len(entries)
        self._set_status(f"Capturing {n} item{'s' if n > 1 else ''}…")
        self._set_action_btns(False)

        def worker() -> None:
            cap_results = []
            for e in entries:
                try:
                    path = self._capturer.capture(
                        e.file_path, e.sheet_name, e.cell_range, e.name
                    )
                    e.last_capture_path = path
                    e.last_capture_time = datetime.now().strftime("%Y-%m-%d %H:%M")
                    self._store.update(e)
                    cap_results.append((e, path, None))
                except CaptureError as exc:
                    cap_results.append((e, None, str(exc)))
                except Exception as exc:
                    cap_results.append((e, None, f"Unexpected: {exc}"))

            # Build paste jobs for entries that captured OK and have a PPT dest
            jobs = [
                {
                    "pptx_path":    e.pptx_file,
                    "slide_number": e.pptx_slide,
                    "image_path":   path,
                    "left":   e.pptx_left,
                    "top":    e.pptx_top,
                    "width":  e.pptx_width,
                    "height": e.pptx_height,
                    "entry":  e,
                }
                for e, path, err in cap_results
                if err is None and e.has_ppt_dest()
            ]
            try:
                ppt_results = PPTExporter().paste_batch(jobs) if jobs else []
            except PPTExportError as exc:
                ppt_results = [(j, str(exc)) for j in jobs]
            except Exception as exc:
                ppt_results = [(j, f"Unexpected error: {exc}") for j in jobs]

            self.after(0, lambda: self._run_all_done(cap_results, ppt_results))

        threading.Thread(target=worker, daemon=True).start()

    def _run_paste(self, entries: list) -> None:
        n = len(entries)
        self._set_status(f"Pasting {n} item{'s' if n > 1 else ''} to PPT…")
        self._set_action_btns(False)

        def worker() -> None:
            jobs = [
                {
                    "pptx_path":    e.pptx_file,
                    "slide_number": e.pptx_slide,
                    "image_path":   e.last_capture_path,
                    "left":   e.pptx_left,
                    "top":    e.pptx_top,
                    "width":  e.pptx_width,
                    "height": e.pptx_height,
                    "entry":  e,
                }
                for e in entries
            ]
            try:
                results = PPTExporter().paste_batch(jobs)
            except PPTExportError as exc:
                results = [(j, str(exc)) for j in jobs]
            except Exception as exc:
                results = [(j, f"Unexpected error: {exc}") for j in jobs]
            self.after(0, lambda: self._paste_done(results))

        threading.Thread(target=worker, daemon=True).start()

    def _run_all_done(self, cap_results: list, ppt_results: list) -> None:
        self._set_action_btns(True)
        self._refresh_tree(select_id=self._selected_id())
        self._on_tree_select()

        cap_errors = [(e.name, msg) for e, _, msg in cap_results if msg]
        ppt_errors = [(j["entry"].name, msg) for j, msg in ppt_results if msg]
        ok_cap = sum(1 for _, p, _ in cap_results if p)
        ok_ppt = sum(1 for _, msg in ppt_results if not msg)

        lines = []
        if ok_cap:
            lines.append(f"Captured:  {ok_cap} item{'s' if ok_cap != 1 else ''}")
        if ok_ppt:
            lines.append(f"Pasted:    {ok_ppt} item{'s' if ok_ppt != 1 else ''} to PPT")

        if cap_errors or ppt_errors:
            all_errs = (
                [f"[Capture] {n}: {m}" for n, m in cap_errors]
                + [f"[PPT]     {n}: {m}" for n, m in ppt_errors]
            )
            messagebox.showerror(
                "Errors",
                "\n\n".join(all_errs),
                parent=self,
            )
            self._set_status(f"Done with {len(all_errs)} error(s)")
        else:
            summary = "  |  ".join(lines) if lines else "Nothing to do"
            self._set_status(summary)

    def _paste_done(self, results: list) -> None:
        self._set_action_btns(True)
        errors = [(j["entry"].name, msg) for j, msg in results if msg]
        ok = sum(1 for _, msg in results if not msg)
        if errors:
            detail = "\n\n".join(f"• {n}:\n  {m}" for n, m in errors)
            messagebox.showerror("Paste Errors", detail, parent=self)
            self._set_status(f"Pasted {ok}, failed {len(errors)}")
        else:
            self._set_status(f"Pasted {ok} item{'s' if ok != 1 else ''} to PPT successfully")

    def _set_action_btns(self, enabled: bool) -> None:
        state = "normal" if enabled else "disabled"
        for b in (
            self._cap_btn, self._cap_all_btn,
            self._paste_btn, self._paste_all_btn,
            self._run_all_btn,
        ):
            b.config(state=state)

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
