"""
data_store.py – Persistent storage for Excel range entries.
Uses a local JSON file; no database required.
"""

from __future__ import annotations

import json
import uuid
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import List, Optional


@dataclass
class Entry:
    id: str
    name: str
    file_path: str
    sheet_name: str
    cell_range: str
    last_capture_path: Optional[str] = None
    last_capture_time: Optional[str] = None
    notes: str = ""

    # ------------------------------------------------------------------
    # Factory
    # ------------------------------------------------------------------

    @classmethod
    def create(
        cls,
        name: str,
        file_path: str,
        sheet_name: str,
        cell_range: str,
        notes: str = "",
    ) -> "Entry":
        return cls(
            id=str(uuid.uuid4()),
            name=name,
            file_path=file_path,
            sheet_name=sheet_name,
            cell_range=cell_range,
            notes=notes,
        )

    # ------------------------------------------------------------------
    # Display helpers
    # ------------------------------------------------------------------

    def display_file(self) -> str:
        return Path(self.file_path).name if self.file_path else ""

    def display_capture_time(self) -> str:
        return self.last_capture_time or "Never"


class DataStore:
    """Load / save entries to a JSON file atomically."""

    _VERSION = 1

    def __init__(self, data_file: str = "entries.json") -> None:
        self._path = Path(data_file)
        self._entries: List[Entry] = []
        self._load()

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------

    @property
    def entries(self) -> List[Entry]:
        return list(self._entries)

    def add(self, entry: Entry) -> None:
        self._entries.append(entry)
        self._save()

    def update(self, entry: Entry) -> None:
        for i, e in enumerate(self._entries):
            if e.id == entry.id:
                self._entries[i] = entry
                self._save()
                return
        raise KeyError(f"Entry not found: {entry.id}")

    def delete(self, entry_id: str) -> None:
        before = len(self._entries)
        self._entries = [e for e in self._entries if e.id != entry_id]
        if len(self._entries) == before:
            raise KeyError(f"Entry not found: {entry_id}")
        self._save()

    def get(self, entry_id: str) -> Optional[Entry]:
        for e in self._entries:
            if e.id == entry_id:
                return e
        return None

    # ------------------------------------------------------------------
    # Private helpers
    # ------------------------------------------------------------------

    def _load(self) -> None:
        if not self._path.exists():
            return
        try:
            with open(self._path, "r", encoding="utf-8") as fh:
                raw = json.load(fh)
            self._entries = [Entry(**item) for item in raw.get("entries", [])]
        except Exception as exc:
            print(f"[DataStore] Could not load {self._path}: {exc}")
            self._entries = []

    def _save(self) -> None:
        payload = {
            "version": self._VERSION,
            "entries": [asdict(e) for e in self._entries],
        }
        self._path.parent.mkdir(parents=True, exist_ok=True)
        # Write to a temp file then rename for atomicity
        tmp = self._path.with_suffix(".tmp")
        with open(tmp, "w", encoding="utf-8") as fh:
            json.dump(payload, fh, indent=2, ensure_ascii=False)
        tmp.replace(self._path)
