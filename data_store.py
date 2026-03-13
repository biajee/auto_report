"""
data_store.py – Persistent storage for tasks (Excel & Power BI) and reports.
Uses a local JSON file; no database required.
"""

from __future__ import annotations

import json
import uuid
from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import List, Optional


# ---------------------------------------------------------------------------
# Excel task
# ---------------------------------------------------------------------------

@dataclass
class ExcelTask:
    id: str
    name: str
    file_path: str
    sheet_name: str
    cell_range: str
    last_capture_path: Optional[str] = None
    last_capture_time: Optional[str] = None
    notes: str = ""

    # Dropdown cells to set before capture: [{"cell": "B2", "value": "Q1"}]
    dropdowns: list = field(default_factory=list)

    # PowerPoint destination (all optional)
    pptx_file: Optional[str] = None
    pptx_slide: int = 1
    pptx_left: float = 0.5
    pptx_top: float = 0.5
    pptx_width: float = 9.0
    pptx_height: float = 6.5
    pptx_output: Optional[str] = None

    @classmethod
    def create(cls, name, file_path, sheet_name, cell_range, notes="") -> "ExcelTask":
        return cls(
            id=str(uuid.uuid4()),
            name=name, file_path=file_path,
            sheet_name=sheet_name, cell_range=cell_range, notes=notes,
        )

    def display_file(self) -> str:
        return Path(self.file_path).name if self.file_path else ""

    def display_capture_time(self) -> str:
        return self.last_capture_time or "Never"

    def has_ppt_dest(self) -> bool:
        return bool(self.pptx_file)

    def display_ppt_dest(self) -> str:
        if not self.pptx_file:
            return "—"
        return f"{Path(self.pptx_file).name}  sl.{self.pptx_slide}"


# ---------------------------------------------------------------------------
# Power BI entry
# ---------------------------------------------------------------------------

@dataclass
class PowerBIEntry:
    """
    A screenshot task targeting a Power BI report page in a browser.

    Interactions applied before the screenshot is taken:
      dropdowns – list of {"label": "Region",  "value": "Asia"}
      filters   – list of {"name":  "Year",    "value": "2024"}
      buttons   – list of {"label": "Apply"}
    """
    id: str
    name: str
    url: str
    report_page: str = ""
    notes: str = ""

    # Interactions (ordered – applied top to bottom)
    dropdowns: list = field(default_factory=list)
    filters:   list = field(default_factory=list)
    buttons:   list = field(default_factory=list)

    # Crop region applied after screenshot (pixels; all 0 = no crop)
    crop_enabled: bool = False
    crop_left:    int  = 0
    crop_top:     int  = 0
    crop_width:   int  = 0   # 0 = full width from crop_left
    crop_height:  int  = 0   # 0 = full height from crop_top

    # Capture results
    last_capture_path: Optional[str] = None
    last_capture_time: Optional[str] = None

    # PowerPoint destination
    pptx_file:   Optional[str] = None
    pptx_slide:  int   = 1
    pptx_left:   float = 0.5
    pptx_top:    float = 0.5
    pptx_width:  float = 9.0
    pptx_height: float = 6.5
    pptx_output: Optional[str] = None

    @classmethod
    def create(cls, name: str, url: str, report_page: str = "",
               notes: str = "") -> "PowerBIEntry":
        return cls(id=str(uuid.uuid4()), name=name, url=url,
                   report_page=report_page, notes=notes)


# ---------------------------------------------------------------------------
# Report  (ordered mix of Excel and Power BI tasks)
# ---------------------------------------------------------------------------

@dataclass
class Report:
    """
    A named, ordered group of tasks that run together as one operation.

    tasks – ordered list of {"type": "excel"|"powerbi", "id": "<task id>"}
    """
    id: str
    name: str
    tasks: List[dict] = field(default_factory=list)
    notes: str = ""

    @classmethod
    def create(cls, name: str, notes: str = "") -> "Report":
        return cls(id=str(uuid.uuid4()), name=name, notes=notes)


# ---------------------------------------------------------------------------
# DataStore
# ---------------------------------------------------------------------------

class DataStore:
    """Load / save all data to a single JSON file atomically."""

    _VERSION = 3

    def __init__(self, data_file: str = "entries.json") -> None:
        self._path = Path(data_file)
        self._tasks:     List[ExcelTask]     = []
        self._pbi_tasks: List[PowerBIEntry]  = []
        self._reports:   List[Report]        = []
        self._load()

    # ── Excel tasks ────────────────────────────────────────────────────────

    @property
    def tasks(self) -> List[ExcelTask]:
        return list(self._tasks)

    def add_task(self, task: ExcelTask) -> None:
        self._tasks.append(task)
        self._save()

    def update_task(self, task: ExcelTask) -> None:
        for i, e in enumerate(self._tasks):
            if e.id == task.id:
                self._tasks[i] = task
                self._save()
                return
        raise KeyError(f"Task not found: {task.id}")

    def delete_task(self, task_id: str) -> None:
        before = len(self._tasks)
        self._tasks = [e for e in self._tasks if e.id != task_id]
        if len(self._tasks) == before:
            raise KeyError(f"Task not found: {task_id}")
        self._remove_task_from_reports("excel", task_id)
        self._save()

    def get_task(self, task_id: str) -> Optional[ExcelTask]:
        return next((e for e in self._tasks if e.id == task_id), None)

    # ── Power BI tasks ─────────────────────────────────────────────────────

    @property
    def pbi_tasks(self) -> List[PowerBIEntry]:
        return list(self._pbi_tasks)

    def add_pbi_task(self, task: PowerBIEntry) -> None:
        self._pbi_tasks.append(task)
        self._save()

    def update_pbi_task(self, task: PowerBIEntry) -> None:
        for i, e in enumerate(self._pbi_tasks):
            if e.id == task.id:
                self._pbi_tasks[i] = task
                self._save()
                return
        raise KeyError(f"PowerBIEntry not found: {task.id}")

    def delete_pbi_task(self, task_id: str) -> None:
        before = len(self._pbi_tasks)
        self._pbi_tasks = [e for e in self._pbi_tasks if e.id != task_id]
        if len(self._pbi_tasks) == before:
            raise KeyError(f"PowerBIEntry not found: {task_id}")
        self._remove_task_from_reports("powerbi", task_id)
        self._save()

    def get_pbi_task(self, task_id: str) -> Optional[PowerBIEntry]:
        return next((e for e in self._pbi_tasks if e.id == task_id), None)

    # ── Reports ────────────────────────────────────────────────────────────

    @property
    def reports(self) -> List[Report]:
        return list(self._reports)

    def add_report(self, report: Report) -> None:
        self._reports.append(report)
        self._save()

    def update_report(self, report: Report) -> None:
        for i, r in enumerate(self._reports):
            if r.id == report.id:
                self._reports[i] = report
                self._save()
                return
        raise KeyError(f"Report not found: {report.id}")

    def delete_report(self, report_id: str) -> None:
        before = len(self._reports)
        self._reports = [r for r in self._reports if r.id != report_id]
        if len(self._reports) == before:
            raise KeyError(f"Report not found: {report_id}")
        self._save()

    def get_report(self, report_id: str) -> Optional[Report]:
        return next((r for r in self._reports if r.id == report_id), None)

    # ── Private ────────────────────────────────────────────────────────────

    def _remove_task_from_reports(self, task_type: str, task_id: str) -> None:
        for r in self._reports:
            r.tasks = [t for t in r.tasks
                       if not (t.get("type") == task_type and t.get("id") == task_id)]

    def _load(self) -> None:
        if not self._path.exists():
            return
        try:
            with open(self._path, "r", encoding="utf-8") as fh:
                raw = json.load(fh)

            # Accept both old "entries" key and new "tasks" key
            self._tasks = [ExcelTask(**item)
                           for item in (raw.get("tasks") or raw.get("entries", []))]
            # Accept both old "pbi_entries" key and new "pbi_tasks" key
            self._pbi_tasks = [PowerBIEntry(**item)
                               for item in (raw.get("pbi_tasks") or raw.get("pbi_entries", []))]

            # Migrate old report format (entry_ids / pbi_entry_ids → tasks)
            reports = []
            for item in raw.get("reports", []):
                if "tasks" not in item:
                    tasks = (
                        [{"type": "excel",   "id": i} for i in item.get("entry_ids", [])] +
                        [{"type": "powerbi", "id": i} for i in item.get("pbi_entry_ids", [])]
                    )
                    item = {k: v for k, v in item.items()
                            if k not in ("entry_ids", "pbi_entry_ids")}
                    item["tasks"] = tasks
                reports.append(Report(**item))
            self._reports = reports

        except Exception as exc:
            print(f"[DataStore] Could not load {self._path}: {exc}")
            self._tasks = []
            self._pbi_tasks = []
            self._reports = []

    def _save(self) -> None:
        payload = {
            "version":   self._VERSION,
            "tasks":     [asdict(e) for e in self._tasks],
            "pbi_tasks": [asdict(e) for e in self._pbi_tasks],
            "reports":   [asdict(r) for r in self._reports],
        }
        self._path.parent.mkdir(parents=True, exist_ok=True)
        tmp = self._path.with_suffix(".tmp")
        with open(tmp, "w", encoding="utf-8") as fh:
            json.dump(payload, fh, indent=2, ensure_ascii=False)
        tmp.replace(self._path)
