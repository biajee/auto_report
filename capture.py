"""
capture.py – Screenshot an Excel cell range using the Excel COM API.

Requirements:
    pip install pywin32 Pillow openpyxl
"""

from __future__ import annotations

import time
from datetime import datetime
from pathlib import Path
from typing import List


class CaptureError(Exception):
    """Raised when a screenshot operation fails."""


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
    ) -> str:
        """
        Open the Excel file, copy *cell_range* from *sheet_name* as a
        bitmap picture, and save it as a timestamped PNG.

        Returns the absolute path of the saved image.
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
        )
        return str(out_path)

    @staticmethod
    def get_sheet_names(file_path: str) -> List[str]:
        """
        Return the sheet names from an Excel workbook.

        Tries openpyxl first (fast, no Excel needed), then falls back to
        the COM interface if the file format is not supported by openpyxl.
        """
        path = str(Path(file_path).resolve())

        # Fast path: openpyxl (works for .xlsx / .xlsm)
        try:
            import openpyxl  # noqa: PLC0415

            wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
            names = wb.sheetnames
            wb.close()
            return names
        except Exception:
            pass

        # Fallback: COM (works for .xls, .xlsb, password-protected, etc.)
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
    ) -> None:
        # --- dependency checks ----------------------------------------
        try:
            import win32com.client  # noqa: PLC0415
        except ImportError:
            raise CaptureError(
                "pywin32 is not installed.\n"
                "Fix: pip install pywin32"
            )
        try:
            from PIL import ImageGrab  # noqa: PLC0415
        except ImportError:
            raise CaptureError(
                "Pillow is not installed.\n"
                "Fix: pip install Pillow"
            )

        if not Path(abs_path).exists():
            raise CaptureError(f"File not found:\n{abs_path}")

        # --- open Excel -----------------------------------------------
        # DispatchEx always creates a new Excel process so we never
        # interfere with the user's open workbooks.
        xl = win32com.client.DispatchEx("Excel.Application")
        xl.Visible = False
        xl.DisplayAlerts = False

        try:
            wb = xl.Workbooks.Open(abs_path, ReadOnly=True)

            try:
                ws = wb.Worksheets(sheet_name)
            except Exception:
                available = [wb.Worksheets(i + 1).Name for i in range(wb.Worksheets.Count)]
                raise CaptureError(
                    f"Sheet '{sheet_name}' not found.\n"
                    f"Available sheets: {', '.join(available)}"
                )

            try:
                rng = ws.Range(cell_range)
            except Exception:
                raise CaptureError(f"Invalid cell range: '{cell_range}'")

            # CopyPicture(Appearance=xlScreen=1, Format=xlBitmap=2)
            rng.CopyPicture(Appearance=1, Format=2)
            time.sleep(0.5)  # give Excel time to write to clipboard

            img = ImageGrab.grabclipboard()
            if img is None:
                raise CaptureError(
                    "Clipboard was empty after CopyPicture.\n"
                    "The range may be empty or Excel could not render it."
                )

            img.save(out_path, "PNG")
            wb.Close(SaveChanges=False)

        except CaptureError:
            raise
        except Exception as exc:
            raise CaptureError(f"Unexpected error during capture: {exc}") from exc
        finally:
            try:
                xl.Quit()
            except Exception:
                pass
