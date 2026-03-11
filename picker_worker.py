"""
picker_worker.py – Range picker subprocess.

Mirrors the logic from main.py's RangePickerDialog._worker() exactly.
Communicates via a shared JSON state file (no stdout pipe buffering issues).

Arguments:
    sys.argv[1]  abs path to Excel file
    sys.argv[2]  sheet name (may be empty string)
    sys.argv[3]  path to state JSON file  (worker writes here)
    sys.argv[4]  path to stop sentinel file (Flask creates this to signal stop)
"""

import json
import sys
import threading
import time
from pathlib import Path


def write_state(state_path: Path, data: dict) -> None:
    tmp = state_path.with_suffix(".tmp")
    tmp.write_text(json.dumps(data), encoding="utf-8")
    tmp.replace(state_path)


def main() -> None:
    if len(sys.argv) < 5:
        sys.exit("picker_worker: wrong number of arguments")

    file_path  = sys.argv[1]
    sheet_name = sys.argv[2]
    state_path = Path(sys.argv[3])
    stop_path  = Path(sys.argv[4])

    write_state(state_path, {"status": "starting"})

    # Initialise COM – same as main.py _worker()
    try:
        import pythoncom
        pythoncom.CoInitialize()
    except Exception:
        pass

    xl = None
    wb = None
    stop = threading.Event()

    try:
        try:
            import win32com.client
        except ImportError as exc:
            write_state(state_path, {"status": "error", "error": str(exc)})
            return

        xl = win32com.client.DispatchEx("Excel.Application")
        xl.Visible = True
        xl.DisplayAlerts = False
        try:
            xl.WindowState = -4137   # xlMaximized – same as main.py
        except Exception:
            pass

        abs_path = str(Path(file_path).resolve())
        wb = xl.Workbooks.Open(abs_path, ReadOnly=True)

        if sheet_name:
            try:
                wb.Worksheets(sheet_name).Activate()
            except Exception:
                pass

        # Bring Excel to front – same as main.py
        try:
            xl.ActiveWindow.Activate()
        except Exception:
            pass

        write_state(state_path, {"status": "ready", "sheet": "", "range": ""})

        # Poll loop – mirrors main.py exactly:
        #   addr = xl.Selection.Address.replace("$", "")   ← property, not method
        #   sheet = xl.ActiveSheet.Name
        #   stop.wait(0.4)
        while not stop_path.exists():
            try:
                addr  = xl.Selection.Address.replace("$", "")
                sheet = xl.ActiveSheet.Name
                write_state(state_path, {
                    "status": "ready",
                    "sheet":  sheet,
                    "range":  addr,
                })
            except Exception as exc:
                write_state(state_path, {
                    "status": "ready",
                    "sheet":  "",
                    "range":  "",
                    "err":    str(exc),
                })
            time.sleep(0.4)

    except Exception as exc:
        write_state(state_path, {"status": "error", "error": str(exc)})
    finally:
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
            pythoncom.CoUninitialize()
        except Exception:
            pass
        try:
            stop_path.unlink(missing_ok=True)
        except Exception:
            pass


if __name__ == "__main__":
    main()
