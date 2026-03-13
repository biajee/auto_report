"""
app.py – Finance Team Report Manager (Web Interface)
========================================================
A Flask web server that exposes the Excel screenshot / PowerPoint export
functionality through a browser UI.

Usage:
    python app.py
    Then open http://localhost:5000 in your browser.

Requirements:
    pip install flask pywin32 Pillow openpyxl python-pptx
"""

from __future__ import annotations

import json
import os
import queue
import subprocess
import sys
import threading
import uuid
from pathlib import Path
from typing import Iterator

from flask import (
    Flask,
    Response,
    jsonify,
    render_template,
    request,
    send_from_directory,
)

from capture import CaptureError, ExcelCapture
from data_store import DataStore, ExcelTask, PowerBIEntry, Report
from ppt_export import PPTExportError, PPTExporter

# ---------------------------------------------------------------------------
# App setup
# ---------------------------------------------------------------------------

app = Flask(__name__)
store = DataStore("entries.json")
capturer = ExcelCapture("screenshots")
exporter = PPTExporter()

SCREENSHOTS_DIR = Path("screenshots")
SCREENSHOTS_DIR.mkdir(exist_ok=True)

# ---------------------------------------------------------------------------
# Background job registry  (job_id -> {"status", "log", "done", "error"})
# ---------------------------------------------------------------------------

_jobs: dict[str, dict] = {}
_jobs_lock = threading.Lock()

# ---------------------------------------------------------------------------
# Range-picker session registry
# (sid -> {"last": {sheet, range}, "stop": Event, "error": str|None})
# ---------------------------------------------------------------------------

_pick_sessions: dict[str, dict] = {}
_pick_lock = threading.Lock()


def _new_job() -> str:
    tid = str(uuid.uuid4())
    with _jobs_lock:
        _jobs[tid] = {"log": [], "done": False, "error": None}
    return tid


def _job_log(tid: str, msg: str, level: str = "info") -> None:
    with _jobs_lock:
        _jobs[tid]["log"].append({"msg": msg, "level": level})


def _job_done(tid: str, error: str | None = None) -> None:
    with _jobs_lock:
        _jobs[tid]["done"] = True
        _jobs[tid]["error"] = error


# ---------------------------------------------------------------------------
# Routes – pages
# ---------------------------------------------------------------------------


@app.route("/")
def index():
    return render_template("index.html")


# ---------------------------------------------------------------------------
# Routes – static screenshots
# ---------------------------------------------------------------------------


@app.route("/screenshots/<path:filename>")
def screenshot(filename: str):
    return send_from_directory(SCREENSHOTS_DIR, filename)


# ---------------------------------------------------------------------------
# Routes – Excel tasks CRUD
# ---------------------------------------------------------------------------


@app.route("/api/tasks", methods=["GET"])
def list_tasks():
    return jsonify([_task_dict(e) for e in store.tasks])


@app.route("/api/tasks", methods=["POST"])
def create_task():
    data = request.json or {}
    try:
        e = ExcelTask.create(
            name=data["name"],
            file_path=data["file_path"],
            sheet_name=data["sheet_name"],
            cell_range=data["cell_range"],
            notes=data.get("notes", ""),
        )
        e.dropdowns   = data.get("dropdowns", [])
        e.pptx_file   = data.get("pptx_file") or None
        e.pptx_slide  = int(data.get("pptx_slide", 1))
        e.pptx_left   = float(data.get("pptx_left", 0.5))
        e.pptx_top    = float(data.get("pptx_top", 0.5))
        e.pptx_width  = float(data.get("pptx_width", 9.0))
        e.pptx_height = float(data.get("pptx_height", 6.5))
        e.pptx_output = data.get("pptx_output") or None
        store.add_task(e)
        return jsonify(_task_dict(e)), 201
    except KeyError as exc:
        return jsonify({"error": f"Missing field: {exc}"}), 400


@app.route("/api/tasks/<task_id>", methods=["PUT"])
def update_task(task_id: str):
    e = store.get_task(task_id)
    if not e:
        return jsonify({"error": "Not found"}), 404
    data = request.json or {}
    updated = ExcelTask(
        id=e.id,
        name=data.get("name", e.name),
        file_path=data.get("file_path", e.file_path),
        sheet_name=data.get("sheet_name", e.sheet_name),
        cell_range=data.get("cell_range", e.cell_range),
        last_capture_path=e.last_capture_path,
        last_capture_time=e.last_capture_time,
        notes=data.get("notes", e.notes),
        dropdowns=data.get("dropdowns", e.dropdowns),
        pptx_file=data.get("pptx_file") or None,
        pptx_slide=int(data.get("pptx_slide", e.pptx_slide)),
        pptx_left=float(data.get("pptx_left", e.pptx_left)),
        pptx_top=float(data.get("pptx_top", e.pptx_top)),
        pptx_width=float(data.get("pptx_width", e.pptx_width)),
        pptx_height=float(data.get("pptx_height", e.pptx_height)),
        pptx_output=data.get("pptx_output") or None,
    )
    store.update_task(updated)
    return jsonify(_task_dict(updated))


@app.route("/api/tasks/<task_id>", methods=["DELETE"])
def delete_task(task_id: str):
    try:
        store.delete_task(task_id)
        return "", 204
    except KeyError:
        return jsonify({"error": "Not found"}), 404


@app.route("/api/tasks/<task_id>/clone", methods=["POST"])
def clone_task(task_id: str):
    e = store.get_task(task_id)
    if not e:
        return jsonify({"error": "Not found"}), 404
    from dataclasses import replace
    cloned = replace(e, id=str(__import__("uuid").uuid4()),
                     name=f"Copy of {e.name}",
                     last_capture_path=None,
                     last_capture_time=None)
    store.add_task(cloned)
    return jsonify(_task_dict(cloned)), 201


# ---------------------------------------------------------------------------
# Routes – Excel / PPT helpers
# ---------------------------------------------------------------------------


@app.route("/api/browse", methods=["GET"])
def browse_file():
    """Open a native Windows file-picker dialog and return the chosen path."""
    file_type = request.args.get("type", "excel")  # "excel" or "pptx"

    import tkinter as tk
    from tkinter import filedialog

    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)

    try:
        if file_type == "pptx":
            path = filedialog.askopenfilename(
                title="Select PowerPoint File",
                filetypes=[
                    ("PowerPoint files", "*.pptx *.ppt"),
                    ("All files", "*.*"),
                ],
            )
        else:
            path = filedialog.askopenfilename(
                title="Select Excel File",
                filetypes=[
                    ("Excel files", "*.xlsx *.xlsm *.xls *.xlsb"),
                    ("All files", "*.*"),
                ],
            )
    finally:
        root.destroy()

    # filedialog returns "" if the user cancelled
    return jsonify({"path": path or ""})


@app.route("/api/pick-range/start", methods=["POST"])
def pick_range_start():
    """Spawn picker_worker.py; communicate via temp files (no pipe buffering)."""
    import tempfile

    data       = request.json or {}
    file_path  = data.get("file_path", "").strip()
    sheet_name = data.get("sheet_name", "").strip()

    if not file_path:
        return jsonify({"error": "file_path required"}), 400

    abs_path = str(Path(file_path).resolve())
    if not Path(abs_path).exists():
        return jsonify({"error": f"File not found: {file_path}"}), 400

    worker = Path(__file__).parent / "picker_worker.py"

    # Two temp files: state (worker writes) and stop sentinel (Flask creates)
    fd, state_path = tempfile.mkstemp(suffix=".json", prefix="picker_state_")
    os.close(fd)
    stop_path = state_path + ".stop"

    try:
        proc = subprocess.Popen(
            [sys.executable, str(worker), abs_path, sheet_name,
             state_path, stop_path],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
    except Exception as exc:
        return jsonify({"error": f"Could not start picker: {exc}"}), 500

    sid = str(uuid.uuid4())
    with _pick_lock:
        _pick_sessions[sid] = {
            "state_path": state_path,
            "stop_path":  stop_path,
            "proc":       proc,
        }

    return jsonify({"session_id": sid}), 202


@app.route("/api/pick-range/<sid>/stream")
def pick_range_stream(sid: str):
    """SSE: poll the state file every 400 ms and push changes."""
    import time

    def generate() -> Iterator[str]:
        last_sent = None
        while True:
            with _pick_lock:
                session = _pick_sessions.get(sid)
            if session is None:
                yield f"data: {json.dumps({'done': True})}\n\n"
                return

            state = {}
            try:
                raw = Path(session["state_path"]).read_text(encoding="utf-8")
                state = json.loads(raw)
            except Exception:
                pass   # file not written yet — keep waiting

            status = state.get("status", "starting")

            if status == "error":
                yield f"data: {json.dumps({'error': state.get('error', 'Unknown error')})}\n\n"
                return

            if status == "ready":
                payload = {
                    "sheet": state.get("sheet", ""),
                    "range": state.get("range", ""),
                    "err":   state.get("err", ""),
                }
                if payload != last_sent:
                    yield f"data: {json.dumps(payload)}\n\n"
                    last_sent = dict(payload)
                else:
                    yield ": keep-alive\n\n"
            else:
                # still starting
                yield f"data: {json.dumps({'waiting': True})}\n\n"

            time.sleep(0.4)

    return Response(generate(), mimetype="text/event-stream",
                    headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"})


@app.route("/api/pick-range/<sid>/stop", methods=["POST"])
def pick_range_stop(sid: str):
    """Create the stop sentinel, wait for the worker, return last selection."""
    with _pick_lock:
        session = _pick_sessions.pop(sid, None)
    if not session:
        return jsonify({"error": "Session not found"}), 404

    # Signal worker to quit
    stop_path  = Path(session["stop_path"])
    state_path = Path(session["state_path"])
    try:
        stop_path.touch()
    except Exception:
        pass

    proc = session.get("proc")
    if proc:
        try:
            proc.wait(timeout=6)
        except Exception:
            proc.terminate()

    # Read final state
    last = {"sheet": "", "range": ""}
    try:
        state = json.loads(state_path.read_text(encoding="utf-8"))
        last  = {"sheet": state.get("sheet", ""), "range": state.get("range", "")}
    except Exception:
        pass

    # Clean up temp files
    for p in (state_path, stop_path):
        try:
            p.unlink(missing_ok=True)
        except Exception:
            pass

    return jsonify(last)


@app.route("/api/range-size", methods=["GET"])
def get_range_size():
    """Return the pixel width/height (in points) of an Excel cell range."""
    file_path  = request.args.get("file_path",  "")
    sheet_name = request.args.get("sheet_name", "")
    cell_range = request.args.get("cell_range", "")
    if not file_path or not sheet_name or not cell_range:
        return jsonify({"error": "file_path, sheet_name, cell_range required"}), 400
    try:
        w, h = ExcelCapture.get_range_size(file_path, sheet_name, cell_range)
        return jsonify({"width": w, "height": h})
    except Exception as exc:
        return jsonify({"error": str(exc)}), 500


@app.route("/api/sheets", methods=["GET"])
def get_sheets():
    file_path = request.args.get("file_path", "")
    if not file_path:
        return jsonify({"error": "file_path required"}), 400
    names = ExcelCapture.get_sheet_names(file_path)
    return jsonify({"sheets": names})


@app.route("/api/slide-info", methods=["GET"])
def get_slide_info():
    pptx_path = request.args.get("pptx_path", "")
    if not pptx_path:
        return jsonify({"error": "pptx_path required"}), 400
    count, w, h = PPTExporter.get_slide_info(pptx_path)
    return jsonify({"slide_count": count, "width": w, "height": h})


# ---------------------------------------------------------------------------
# Routes – capture
# ---------------------------------------------------------------------------


@app.route("/api/tasks/<task_id>/capture", methods=["POST"])
def capture_task(task_id: str):
    e = store.get_task(task_id)
    if not e:
        return jsonify({"error": "Not found"}), 404
    tid = _new_job()

    def run():
        from datetime import datetime
        log = lambda msg, level="info": _job_log(tid, msg, level)
        try:
            path, buf = capturer.capture(
                e.file_path, e.sheet_name, e.cell_range, e.name,
                dropdowns=e.dropdowns, log=log,
            )
            e.last_capture_path = path
            e.last_capture_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            store.update_task(e)

            # If a PPT destination is configured, paste directly from the
            # in-memory buffer (clipboard image) — no file read-back needed
            if e.pptx_file:
                out_name = Path(e.pptx_output).name if e.pptx_output else Path(e.pptx_file).name
                log(f"Pasting → {out_name}  slide {e.pptx_slide}", "info")
                exporter.paste_image(
                    e.pptx_file, e.pptx_slide, buf,
                    e.pptx_left, e.pptx_top, e.pptx_width, e.pptx_height,
                    output_path=e.pptx_output,
                )
                log(f"Pasted to slide {e.pptx_slide}", "ok")

            _job_done(tid)
        except (CaptureError, PPTExportError) as exc:
            log(str(exc), "err")
            _job_done(tid, str(exc))
        except Exception as exc:
            log(f"Unexpected error: {exc}", "err")
            _job_done(tid, str(exc))

    threading.Thread(target=run, daemon=True).start()
    return jsonify({"task_id": tid}), 202


@app.route("/api/capture-all", methods=["POST"])
def capture_all():
    excel_tasks = store.tasks
    if not excel_tasks:
        return jsonify({"error": "No tasks"}), 400
    tid = _new_job()

    def run():
        from datetime import datetime
        log = lambda msg, level="info": _job_log(tid, msg, level)

        # Group tasks by PPT file so each PPTX is opened/saved only once
        ppt_jobs: list = []

        for e in excel_tasks:
            log(f"── {e.name}", "head")
            try:
                path, buf = capturer.capture(
                    e.file_path, e.sheet_name, e.cell_range, e.name,
                    dropdowns=e.dropdowns, log=log,
                )
                e.last_capture_path = path
                e.last_capture_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                store.update_task(e)

                if e.pptx_file:
                    ppt_jobs.append({
                        "pptx_path":    e.pptx_file,
                        "slide_number": e.pptx_slide,
                        "image_source": buf,
                        "left":         e.pptx_left,
                        "top":          e.pptx_top,
                        "width":        e.pptx_width,
                        "height":       e.pptx_height,
                        "entry":        e,
                        "output_path":  e.pptx_output,
                    })
            except CaptureError as exc:
                log(str(exc), "err")

        if ppt_jobs:
            log("── Pasting to PowerPoint…", "head")
            try:
                results = exporter.paste_batch(ppt_jobs, log=log)
                errors = [err for _, err in results if err]
                if errors:
                    _job_done(tid, "\n".join(errors))
                    return
            except PPTExportError as exc:
                log(str(exc), "err")
                _job_done(tid, str(exc))
                return

        _job_done(tid)

    threading.Thread(target=run, daemon=True).start()
    return jsonify({"task_id": tid}), 202


# ---------------------------------------------------------------------------
# Routes – export to PowerPoint (uses saved PNG path as fallback)
# ---------------------------------------------------------------------------


@app.route("/api/tasks/<task_id>/export", methods=["POST"])
def export_task(task_id: str):
    """Re-export using the saved PNG (for when you want to re-paste without recapturing)."""
    e = store.get_task(task_id)
    if not e:
        return jsonify({"error": "Not found"}), 404
    if not e.pptx_file:
        return jsonify({"error": "No PowerPoint destination set"}), 400
    if not e.last_capture_path:
        return jsonify({"error": "No screenshot captured yet — capture first"}), 400
    tid = _new_job()

    def run():
        try:
            _job_log(tid, f"Exporting '{e.name}' → slide {e.pptx_slide}", "info")
            exporter.paste_image(
                e.pptx_file, e.pptx_slide, e.last_capture_path,
                e.pptx_left, e.pptx_top, e.pptx_width, e.pptx_height,
                output_path=e.pptx_output,
            )
            out_name = Path(e.pptx_output).name if e.pptx_output else Path(e.pptx_file).name
            _job_log(tid, f"Saved {out_name}", "ok")
            _job_done(tid)
        except PPTExportError as exc:
            _job_log(tid, str(exc), "err")
            _job_done(tid, str(exc))
        except Exception as exc:
            _job_log(tid, f"Unexpected error: {exc}", "err")
            _job_done(tid, str(exc))

    threading.Thread(target=run, daemon=True).start()
    return jsonify({"task_id": tid}), 202


@app.route("/api/export-all", methods=["POST"])
def export_all():
    """Re-export all tasks using their saved PNGs."""
    excel_tasks = [e for e in store.tasks if e.pptx_file and e.last_capture_path]
    if not excel_tasks:
        return jsonify({"error": "No tasks with both a screenshot and a PPT destination"}), 400
    tid = _new_job()

    def run():
        jobs = [
            {
                "pptx_path":    e.pptx_file,
                "slide_number": e.pptx_slide,
                "image_source": e.last_capture_path,
                "left":         e.pptx_left,
                "top":          e.pptx_top,
                "width":        e.pptx_width,
                "height":       e.pptx_height,
                "entry":        e,
                "output_path":  e.pptx_output,
            }
            for e in excel_tasks
        ]
        try:
            results = exporter.paste_batch(
                jobs, log=lambda msg, level="info": _job_log(tid, msg, level),
            )
            errors = [err for _, err in results if err]
            _job_done(tid, "\n".join(errors) if errors else None)
        except PPTExportError as exc:
            _job_log(tid, str(exc), "err")
            _job_done(tid, str(exc))

    threading.Thread(target=run, daemon=True).start()
    return jsonify({"task_id": tid}), 202


# ---------------------------------------------------------------------------
# Routes – Power BI tasks CRUD + stub capture
# ---------------------------------------------------------------------------


def _pbi_dict(e: PowerBIEntry) -> dict:
    return {
        "id":                e.id,
        "name":              e.name,
        "url":               e.url,
        "report_page":       e.report_page,
        "notes":             e.notes,
        "dropdowns":         e.dropdowns,
        "filters":           e.filters,
        "buttons":           e.buttons,
        "crop_enabled": e.crop_enabled,
        "crop_left":    e.crop_left,
        "crop_top":     e.crop_top,
        "crop_width":   e.crop_width,
        "crop_height":  e.crop_height,
        "last_capture_path": e.last_capture_path,
        "last_capture_time": e.last_capture_time,
        "last_capture_url":  (
            "/screenshots/" + Path(e.last_capture_path).name
            if e.last_capture_path and Path(e.last_capture_path).exists() else None
        ),
        "pptx_file":      e.pptx_file,
        "pptx_file_name": Path(e.pptx_file).name if e.pptx_file else "",
        "pptx_slide":     e.pptx_slide,
        "pptx_left":      e.pptx_left,
        "pptx_top":       e.pptx_top,
        "pptx_width":     e.pptx_width,
        "pptx_height":    e.pptx_height,
        "pptx_output":      e.pptx_output,
        "pptx_output_name": Path(e.pptx_output).name if e.pptx_output else "",
    }


def _pbi_from_data(data: dict, existing: PowerBIEntry = None) -> PowerBIEntry:
    def _f(k, d): return float(data.get(k, getattr(existing, k, d)) if existing else data.get(k, d))
    def _i(k, d): return int(data.get(k, getattr(existing, k, d)) if existing else data.get(k, d))
    return PowerBIEntry(
        id           = existing.id if existing else str(__import__("uuid").uuid4()),
        name         = data.get("name",        getattr(existing, "name",        "")),
        url          = data.get("url",         getattr(existing, "url",         "")),
        report_page  = data.get("report_page", getattr(existing, "report_page", "")),
        notes        = data.get("notes",       getattr(existing, "notes",       "")),
        dropdowns    = data.get("dropdowns",   getattr(existing, "dropdowns",   [])),
        filters      = data.get("filters",     getattr(existing, "filters",     [])),
        buttons      = data.get("buttons",     getattr(existing, "buttons",     [])),
        crop_enabled  = bool(data.get("crop_enabled", getattr(existing, "crop_enabled", False))),
        crop_left     = int(data.get("crop_left",   getattr(existing, "crop_left",   0))),
        crop_top      = int(data.get("crop_top",    getattr(existing, "crop_top",    0))),
        crop_width    = int(data.get("crop_width",  getattr(existing, "crop_width",  0))),
        crop_height   = int(data.get("crop_height", getattr(existing, "crop_height", 0))),
        last_capture_path = getattr(existing, "last_capture_path", None),
        last_capture_time = getattr(existing, "last_capture_time", None),
        pptx_file    = data.get("pptx_file") or None,
        pptx_slide   = _i("pptx_slide", 1),
        pptx_left    = _f("pptx_left",  0.5),
        pptx_top     = _f("pptx_top",   0.5),
        pptx_width   = _f("pptx_width", 9.0),
        pptx_height  = _f("pptx_height",6.5),
        pptx_output  = data.get("pptx_output") or None,
    )


@app.route("/api/pbi-tasks", methods=["GET"])
def list_pbi_tasks():
    return jsonify([_pbi_dict(e) for e in store.pbi_tasks])


@app.route("/api/pbi-tasks", methods=["POST"])
def create_pbi_task():
    data = request.json or {}
    if not data.get("name") or not data.get("url"):
        return jsonify({"error": "name and url are required"}), 400
    e = _pbi_from_data(data)
    store.add_pbi_task(e)
    return jsonify(_pbi_dict(e)), 201


@app.route("/api/pbi-tasks/<task_id>", methods=["PUT"])
def update_pbi_task(task_id: str):
    existing = store.get_pbi_task(task_id)
    if not existing:
        return jsonify({"error": "Not found"}), 404
    updated = _pbi_from_data(request.json or {}, existing)
    store.update_pbi_task(updated)
    return jsonify(_pbi_dict(updated))


@app.route("/api/pbi-tasks/<task_id>", methods=["DELETE"])
def delete_pbi_task(task_id: str):
    try:
        store.delete_pbi_task(task_id)
        return "", 204
    except KeyError:
        return jsonify({"error": "Not found"}), 404


@app.route("/api/pbi-tasks/<task_id>/clone", methods=["POST"])
def clone_pbi_task(task_id: str):
    e = store.get_pbi_task(task_id)
    if not e:
        return jsonify({"error": "Not found"}), 404
    from dataclasses import replace
    cloned = replace(e, id=str(__import__("uuid").uuid4()),
                     name=f"Copy of {e.name}",
                     last_capture_path=None,
                     last_capture_time=None)
    store.add_pbi_task(cloned)
    return jsonify(_pbi_dict(cloned)), 201


@app.route("/api/pbi-tasks/<task_id>/capture", methods=["POST"])
def capture_pbi_task(task_id: str):
    """Stub — replace body with real Power BI capture when ready."""
    e = store.get_pbi_task(task_id)
    if not e:
        return jsonify({"error": "Not found"}), 404
    tid = _new_job()

    def run():
        log = lambda msg, level="info": _job_log(tid, msg, level)
        log(f"Power BI capture: {e.name}", "head")
        log(f"URL: {e.url}", "info")
        if e.report_page:
            log(f"Page: {e.report_page}", "info")
        # ── placeholder – user will supply real implementation ──────────
        log("Power BI capture not yet implemented.", "err")
        _job_done(tid, "Not implemented")

    threading.Thread(target=run, daemon=True).start()
    return jsonify({"task_id": tid}), 202


# ---------------------------------------------------------------------------
# Routes – reports CRUD
# ---------------------------------------------------------------------------


def _report_dict(r: Report) -> dict:
    tasks_out = []
    for t in r.tasks:
        kind = t.get("type")
        tid  = t.get("id")
        if kind == "excel":
            e = store.get_task(tid)
            if e:
                tasks_out.append({"type": "excel",   **_task_dict(e)})
        elif kind == "powerbi":
            e = store.get_pbi_task(tid)
            if e:
                tasks_out.append({"type": "powerbi", **_pbi_dict(e)})
    return {
        "id":    r.id,
        "name":  r.name,
        "notes": r.notes,
        "tasks": r.tasks,
        "task_details": tasks_out,
    }


@app.route("/api/reports", methods=["GET"])
def list_reports():
    return jsonify([_report_dict(r) for r in store.reports])


@app.route("/api/reports", methods=["POST"])
def create_report():
    data = request.json or {}
    r = Report.create(name=data.get("name", "Untitled"), notes=data.get("notes", ""))
    r.tasks = data.get("tasks", [])
    store.add_report(r)
    return jsonify(_report_dict(r)), 201


@app.route("/api/reports/<report_id>", methods=["PUT"])
def update_report(report_id: str):
    r = store.get_report(report_id)
    if not r:
        return jsonify({"error": "Not found"}), 404
    data = request.json or {}
    r.name  = data.get("name",  r.name)
    r.notes = data.get("notes", r.notes)
    r.tasks = data.get("tasks", r.tasks)
    store.update_report(r)
    return jsonify(_report_dict(r))


@app.route("/api/reports/<report_id>", methods=["DELETE"])
def delete_report(report_id: str):
    try:
        store.delete_report(report_id)
        return "", 204
    except KeyError:
        return jsonify({"error": "Not found"}), 404


@app.route("/api/reports/<report_id>/run", methods=["POST"])
def run_report(report_id: str):
    r = store.get_report(report_id)
    if not r:
        return jsonify({"error": "Not found"}), 404
    if not r.tasks:
        return jsonify({"error": "Report has no tasks"}), 400

    tid = _new_job()

    def run():
        from datetime import datetime
        log      = lambda msg, level="info": _job_log(tid, msg, level)
        ppt_jobs: list = []

        for task in r.tasks:
            kind    = task.get("type")
            task_id = task.get("id")

            # ── Excel ────────────────────────────────────────────────────
            if kind == "excel":
                e = store.get_task(task_id)
                if not e:
                    log(f"Excel task {task_id} not found — skipped", "err")
                    continue
                log(f"── [Excel] {e.name}", "head")
                try:
                    path, buf = capturer.capture(
                        e.file_path, e.sheet_name, e.cell_range, e.name,
                        dropdowns=e.dropdowns, log=log,
                    )
                    e.last_capture_path = path
                    e.last_capture_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    store.update_task(e)
                    if e.pptx_file:
                        ppt_jobs.append({
                            "pptx_path":    e.pptx_file,
                            "slide_number": e.pptx_slide,
                            "image_source": buf,
                            "left": e.pptx_left, "top":    e.pptx_top,
                            "width": e.pptx_width, "height": e.pptx_height,
                            "entry": e,
                            "output_path":  e.pptx_output,
                        })
                except CaptureError as exc:
                    log(str(exc), "err")

            # ── Power BI ─────────────────────────────────────────────────
            elif kind == "powerbi":
                e = store.get_pbi_task(task_id)
                if not e:
                    log(f"Power BI task {task_id} not found — skipped", "err")
                    continue
                log(f"── [Power BI] {e.name}", "head")
                # ── placeholder: call real PBI capture here ───────────────
                log("Power BI capture not yet implemented — skipped", "err")

            else:
                log(f"Unknown task type '{kind}' — skipped", "err")

        if ppt_jobs:
            log("── Pasting to PowerPoint…", "head")
            try:
                results = exporter.paste_batch(ppt_jobs, log=log)
                errors  = [err for _, err in results if err]
                if errors:
                    _job_done(tid, "\n".join(errors))
                    return
            except PPTExportError as exc:
                log(str(exc), "err")
                _job_done(tid, str(exc))
                return

        _job_done(tid)

    threading.Thread(target=run, daemon=True).start()
    return jsonify({"task_id": tid}), 202


# ---------------------------------------------------------------------------
# Routes – PPT position picker
# (session: start → stream SSE → [goto] → stop)
# ---------------------------------------------------------------------------

_ppt_pick_sessions: dict[str, dict] = {}
_ppt_pick_lock = threading.Lock()


@app.route("/api/ppt-pick/start", methods=["POST"])
def ppt_pick_start():
    data        = request.json or {}
    pptx_file   = data.get("pptx_file", "").strip()
    slide_num   = int(data.get("slide_number", 1))
    init_left   = float(data.get("left",   0.5))
    init_top    = float(data.get("top",    0.5))
    init_width  = float(data.get("width",  9.0))
    init_height = float(data.get("height", 6.5))

    if not pptx_file:
        return jsonify({"error": "pptx_file required"}), 400
    abs_path = str(Path(pptx_file).resolve())
    if not Path(abs_path).exists():
        return jsonify({"error": f"File not found: {pptx_file}"}), 400

    sid        = str(uuid.uuid4())
    stop_event = threading.Event()
    cmdq: queue.Queue = queue.Queue()
    state: dict = {
        "status":    "starting",
        "cur_slide": slide_num,
        "total":     0,
        "left":      init_left,
        "top":       init_top,
        "width":     init_width,
        "height":    init_height,
        "error":     None,
    }

    def worker() -> None:
        _PTS = 72.0

        def _style(s) -> None:
            s.Name = "_img_placeholder_"
            s.Line.ForeColor.RGB = 255 + 140 * 256
            s.Line.Weight = 2.5
            s.Line.DashStyle = 2
            s.Fill.ForeColor.RGB = 255 + 165 * 256
            s.Fill.Transparency = 0.75
            tf = s.TextFrame
            tf.TextRange.Text = "IMAGE PLACEHOLDER\nMove & resize to set position"
            tf.TextRange.Font.Size = 12
            tf.TextRange.Font.Bold = True
            tf.TextRange.Font.Color.RGB = 255 + 69 * 256
            tf.VerticalAnchor = 3
            tf.TextRange.ParagraphFormat.Alignment = 2

        ppt = shape = prs = None
        try:
            try:
                import pythoncom
                pythoncom.CoInitialize()
            except Exception:
                pass
            try:
                import win32com.client
            except ImportError:
                state["status"] = "error"
                state["error"]  = "pywin32 not installed. Run: pip install pywin32"
                return

            # GetObject returns the already-open presentation regardless of
            # which PowerPoint instance holds it.  If the file isn't open yet
            # it opens it in an existing or new instance automatically.
            try:
                prs = win32com.client.GetObject(abs_path)
                ppt = prs.Application
            except Exception:
                prs = None
                ppt = None

            if prs is None:
                # File not open yet — attach to a running instance or start one
                try:
                    ppt = win32com.client.GetActiveObject("PowerPoint.Application")
                except Exception:
                    ppt = win32com.client.Dispatch("PowerPoint.Application")
                    ppt.DisplayAlerts = 0
                prs = ppt.Presentations.Open(abs_path, ReadOnly=False)

            ppt.Visible = True

            n_slides = prs.Slides.Count
            cur      = min(max(1, slide_num), n_slides)
            P        = _PTS

            slide = prs.Slides(cur)
            prs.Windows(1).View.GotoSlide(cur)
            try:
                ppt.ActiveWindow.Activate()
            except Exception:
                pass

            shape = slide.Shapes.AddShape(
                1, init_left * P, init_top * P, init_width * P, init_height * P
            )
            _style(shape)
            state.update({"status": "ready", "cur_slide": cur, "total": n_slides})

            while not stop_event.is_set():
                # handle goto commands
                try:
                    while True:
                        cmd, arg = cmdq.get_nowait()
                        if cmd == "goto":
                            target = min(max(1, arg), n_slides)
                            if target != cur:
                                try:
                                    gl = shape.Left;  gt = shape.Top
                                    gw = shape.Width; gh = shape.Height
                                    shape.Delete()
                                except Exception:
                                    gl = init_left  * P; gt = init_top    * P
                                    gw = init_width * P; gh = init_height * P
                                cur   = target
                                slide = prs.Slides(cur)
                                prs.Windows(1).View.GotoSlide(cur)
                                shape = slide.Shapes.AddShape(1, gl, gt, gw, gh)
                                _style(shape)
                                state["cur_slide"] = cur
                except queue.Empty:
                    pass

                # poll position
                try:
                    state.update({
                        "left":   round(shape.Left   / P, 3),
                        "top":    round(shape.Top    / P, 3),
                        "width":  round(shape.Width  / P, 3),
                        "height": round(shape.Height / P, 3),
                    })
                except Exception:
                    pass

                stop_event.wait(0.4)

        except Exception as exc:
            state["status"] = "error"
            state["error"]  = str(exc)
        finally:
            state["status"] = "done"
            # Remove the placeholder shape but leave PowerPoint open
            try:
                if shape:
                    shape.Delete()
            except Exception:
                pass
            try:
                import pythoncom
                pythoncom.CoUninitialize()
            except Exception:
                pass

    t = threading.Thread(target=worker, daemon=True)
    t.start()
    with _ppt_pick_lock:
        _ppt_pick_sessions[sid] = {
            "state":      state,
            "stop_event": stop_event,
            "cmdq":       cmdq,
            "thread":     t,
        }
    return jsonify({"session_id": sid}), 202


@app.route("/api/ppt-pick/<sid>/stream")
def ppt_pick_stream(sid: str):
    """SSE: send position/slide updates until done or cancelled."""
    import time

    def generate() -> Iterator[str]:
        last_payload = None
        while True:
            with _ppt_pick_lock:
                session = _ppt_pick_sessions.get(sid)
            if session is None:
                yield f"data: {json.dumps({'done': True})}\n\n"
                return
            state  = session["state"]
            status = state.get("status", "starting")

            if status == "error":
                yield f"data: {json.dumps({'error': state.get('error', 'Unknown')})}\n\n"
                return
            if status in ("ready", "done"):
                payload = {
                    "cur_slide": state["cur_slide"],
                    "total":     state["total"],
                    "left":      state["left"],
                    "top":       state["top"],
                    "width":     state["width"],
                    "height":    state["height"],
                    "status":    status,
                }
                if payload != last_payload:
                    yield f"data: {json.dumps(payload)}\n\n"
                    last_payload = dict(payload)
                else:
                    yield ": keep-alive\n\n"
                if status == "done":
                    return
            else:
                yield f"data: {json.dumps({'waiting': True})}\n\n"
            time.sleep(0.4)

    return Response(generate(), mimetype="text/event-stream",
                    headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"})


@app.route("/api/ppt-pick/<sid>/goto", methods=["POST"])
def ppt_pick_goto(sid: str):
    with _ppt_pick_lock:
        session = _ppt_pick_sessions.get(sid)
    if not session:
        return jsonify({"error": "Session not found"}), 404
    slide = int((request.json or {}).get("slide", 1))
    session["cmdq"].put(("goto", slide))
    return jsonify({"ok": True})


@app.route("/api/ppt-pick/<sid>/stop", methods=["POST"])
def ppt_pick_stop(sid: str):
    with _ppt_pick_lock:
        session = _ppt_pick_sessions.pop(sid, None)
    if not session:
        return jsonify({"error": "Session not found"}), 404
    session["stop_event"].set()
    try:
        session["thread"].join(timeout=8)
    except Exception:
        pass
    s = session["state"]
    return jsonify({
        "slide":  s["cur_slide"],
        "left":   s["left"],
        "top":    s["top"],
        "width":  s["width"],
        "height": s["height"],
    })


# ---------------------------------------------------------------------------
# Routes – job polling (Server-Sent Events)
# ---------------------------------------------------------------------------


@app.route("/api/jobs/<job_id>/stream")
def job_stream(job_id: str):
    """SSE endpoint — streams log lines until the background job completes."""

    def generate() -> Iterator[str]:
        sent = 0
        while True:
            with _jobs_lock:
                job = _jobs.get(job_id)
            if job is None:
                yield "data: {}\n\n"
                return
            logs = job["log"]
            while sent < len(logs):
                item = logs[sent]
                yield f"data: {json.dumps(item)}\n\n"
                sent += 1
            if job["done"]:
                yield f"data: {json.dumps({'done': True, 'error': job['error']})}\n\n"
                return
            import time
            time.sleep(0.3)

    return Response(generate(), mimetype="text/event-stream",
                    headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"})


@app.route("/api/jobs/<job_id>", methods=["GET"])
def job_status(job_id: str):
    with _jobs_lock:
        job = _jobs.get(job_id)
    if job is None:
        return jsonify({"error": "Not found"}), 404
    return jsonify(job)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def _task_dict(e: ExcelTask) -> dict:
    d = {
        "id":                e.id,
        "name":              e.name,
        "file_path":         e.file_path,
        "file_name":         Path(e.file_path).name if e.file_path else "",
        "sheet_name":        e.sheet_name,
        "cell_range":        e.cell_range,
        "notes":             e.notes,
        "dropdowns":         e.dropdowns,
        "last_capture_path": e.last_capture_path,
        "last_capture_time": e.last_capture_time,
        "last_capture_url":  None,
        "pptx_file":         e.pptx_file,
        "pptx_file_name":    Path(e.pptx_file).name if e.pptx_file else "",
        "pptx_slide":        e.pptx_slide,
        "pptx_left":         e.pptx_left,
        "pptx_top":          e.pptx_top,
        "pptx_width":        e.pptx_width,
        "pptx_height":       e.pptx_height,
        "pptx_output":       e.pptx_output,
        "pptx_output_name":  Path(e.pptx_output).name if e.pptx_output else "",
    }
    if e.last_capture_path and Path(e.last_capture_path).exists():
        d["last_capture_url"] = "/screenshots/" + Path(e.last_capture_path).name
    return d


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    print("Starting Finance Team Report Manager web server…")
    print("Open http://localhost:5000 in your browser.")
    app.run(host="127.0.0.1", port=5000, debug=False, threaded=True)
