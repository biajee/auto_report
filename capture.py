"""
capture.py – Screenshot an Excel cell range using the Excel COM API.

Requirements:
    pip install pywin32 Pillow openpyxl
"""

from __future__ import annotations

import time
from datetime import datetime
from pathlib import Path
from typing import Callable, List, Optional


class CaptureError(Exception):
    """Raised when a screenshot operation fails."""


_Noop = lambda msg, level="info": None   # default no-op logger


class ExcelCapture:
    """Captures Excel ranges as PNG images via the Windows COM interface."""

    def __init__(self, output_dir: str = "screenshots") -> None:
        self.output_dir = Path(output_dir)
        self.output_dir.mkdir(parents=True, exist_ok=True)

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------

    def capture(
        self,
        file_path: str,
        sheet_name: str,
        cell_range: str,
        name: str,
        dropdowns: Optional[list] = None,
        log: Optional[Callable] = None,
    ) -> str:
        """
        Open the Excel file, optionally set dropdown cell values and
        recalculate, then copy *cell_range* as a bitmap PNG.

        *dropdowns* is a list of {"cell": "B2", "value": "Q1"} dicts.
        Excel only recalculates if at least one value actually changed.

        *log(msg, level)* is called with progress messages.
        level is one of: "info", "ok", "err", "dim", "head".

        Returns the absolute path of the saved PNG.
        Raises CaptureError on any failure.
        """
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        safe = "".join(c for c in name if c.isalnum() or c in " -_").strip() or "capture"
        out_path = self.output_dir / f"{safe}_{timestamp}.png"

        self._capture_via_com(
            str(Path(file_path).resolve()),
            sheet_name,
            cell_range.upper(),
            str(out_path),
            dropdowns or [],
            log or _Noop,
        )
        return str(out_path)

    @staticmethod
    def get_sheet_names(file_path: str) -> List[str]:
        """Return sheet names from an Excel workbook (openpyxl first, COM fallback)."""
        path = str(Path(file_path).resolve())

        try:
            import openpyxl  # noqa: PLC0415
            wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
            names = wb.sheetnames
            wb.close()
            return names
        except Exception:
            pass

        try:
            import win32com.client  # noqa: PLC0415
            xl = win32com.client.DispatchEx("Excel.Application")
            xl.Visible = False
            xl.DisplayAlerts = False
            try:
                wb = xl.Workbooks.Open(path, ReadOnly=True)
                names = [wb.Worksheets(i + 1).Name for i in range(wb.Worksheets.Count)]
                wb.Close(False)
                return names
            finally:
                xl.Quit()
        except Exception:
            return []

    # ------------------------------------------------------------------
    # Private implementation
    # ------------------------------------------------------------------

    def _capture_via_com(
        self,
        abs_path: str,
        sheet_name: str,
        cell_range: str,
        out_path: str,
        dropdowns: list,
        log: Callable,
    ) -> None:
        try:
            import win32com.client  # noqa: PLC0415
        except ImportError:
            raise CaptureError("pywin32 is not installed.\nFix: pip install pywin32")
        try:
            from PIL import ImageGrab  # noqa: PLC0415
        except ImportError:
            raise CaptureError("Pillow is not installed.\nFix: pip install Pillow")

        if not Path(abs_path).exists():
            raise CaptureError(f"File not found:\n{abs_path}")

        log(f"Opening  {Path(abs_path).name}", "info")

        xl = win32com.client.DispatchEx("Excel.Application")
        xl.Visible = False
        xl.DisplayAlerts = False

        try:
            wb = xl.Workbooks.Open(abs_path, ReadOnly=False)

            try:
                ws = wb.Worksheets(sheet_name)
            except Exception:
                available = [wb.Worksheets(i + 1).Name for i in range(wb.Worksheets.Count)]
                raise CaptureError(
                    f"Sheet '{sheet_name}' not found.\n"
                    f"Available: {', '.join(available)}"
                )

            # --- dropdown values & conditional recalculation -----------
            if dropdowns:
                log(f"Checking {len(dropdowns)} dropdown cell(s)…", "info")
                changed = self._apply_dropdowns(ws, dropdowns, log)
                if changed:
                    log("Recalculating sheet…", "info")
                    ws.Calculate()
                    time.sleep(1.0)
                    log("Calculation complete", "ok")
                else:
                    log("All values unchanged — skipping recalculation", "dim")

            # --- screenshot --------------------------------------------
            log(f"Capturing  {cell_range}", "info")
            try:
                rng = ws.Range(cell_range)
            except Exception:
                raise CaptureError(f"Invalid cell range: '{cell_range}'")

            rng.CopyPicture(Appearance=1, Format=2)   # xlScreen, xlBitmap
            time.sleep(0.5)

            img = ImageGrab.grabclipboard()
            if img is None:
                raise CaptureError(
                    "Clipboard empty after CopyPicture — range may be empty."
                )

            img.save(out_path, "PNG")
            log(f"Saved  {Path(out_path).name}", "ok")
            wb.Close(SaveChanges=False)

        except CaptureError:
            raise
        except Exception as exc:
            raise CaptureError(f"Unexpected error: {exc}") from exc
        finally:
            try:
                xl.Quit()
            except Exception:
                pass

    @staticmethod
    def _apply_dropdowns(ws, dropdowns: list, log: Callable) -> bool:
        """Set dropdown cell values. Returns True if any value changed."""
        changed = False
        errors: list = []

        for dd in dropdowns:
            cell   = str(dd.get("cell",  "")).strip().upper()
            target = str(dd.get("value", "")).strip()
            if not cell:
                continue
            try:
                current = str(ws.Range(cell).Value or "").strip()
                if current != target:
                    ws.Range(cell).Value = target
                    log(f"  {cell}:  '{current}'  →  '{target}'", "info")
                    changed = True
                else:
                    log(f"  {cell}:  '{current}'  (unchanged)", "dim")
            except Exception as exc:
                errors.append(f"  {cell}: {exc}")

        if errors:
            raise CaptureError(
                "Could not set dropdown cells:\n" + "\n".join(errors)
            )

        return changed
