"""
progress_manager.py - Persistent state and progress manager for UK NSM Downloader.
Tracks processing state in progress.json for pause, resume, and crash recovery.
"""
import os
import json
import threading
from typing import Dict, Any, Optional, List
from config import (
    DEFAULT_STATE_FILE,
    STATUS_SUCCESS,
    STATUS_SKIPPED_ALREADY_DOWNLOADED,
    STATUS_SKIPPED_NO_ANNUAL_REPORT,
    STATUS_SKIPPED_NO_PDF,
    STATUS_REVIEW_REQUIRED
)
from logger import get_logger

logger = get_logger()

class ProgressManager:
    """Thread-safe manager for tracking download execution state and progress."""
    def __init__(self, state_file=DEFAULT_STATE_FILE):
        self.state_file = str(state_file)
        self.lock = threading.Lock()
        self.state = {
            "total_records": 0,
            "processed_count": 0,
            "success_count": 0,
            "skipped_count": 0,
            "failed_count": 0,
            "review_count": 0,
            "start_time": None,
            "last_updated": None,
            "records": {}  # Key: str(uni_code) -> record state
        }
        self.load_state()

    def load_state(self) -> bool:
        """Loads state from JSON file if present."""
        with self.lock:
            if not os.path.exists(self.state_file):
                return False
            try:
                with open(self.state_file, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    self.state.update(data)
                logger.info(f"Loaded progress state: {self.state['processed_count']} records previously processed.")
                return True
            except Exception as e:
                logger.warning(f"Could not load state file {self.state_file}: {e}")
                return False

    def save_state(self):
        """Atomically saves current state to disk."""
        with self.lock:
            try:
                os.makedirs(os.path.dirname(self.state_file), exist_ok=True)
                temp_file = f"{self.state_file}.tmp"
                with open(temp_file, "w", encoding="utf-8") as f:
                    json.dump(self.state, f, indent=2, default=str)
                
                if os.path.exists(self.state_file):
                    os.replace(temp_file, self.state_file)
                else:
                    os.rename(temp_file, self.state_file)
            except Exception as e:
                logger.warning(f"Could not save state file: {e}")

    def reset_state(self, total_count: int = 0):
        """Resets all metrics and state for a fresh run."""
        with self.lock:
            self.state = {
                "total_records": total_count,
                "processed_count": 0,
                "success_count": 0,
                "skipped_count": 0,
                "failed_count": 0,
                "review_count": 0,
                "start_time": None,
                "last_updated": None,
                "records": {}
            }
        self.save_state()

    def is_record_completed(self, uni_code: str) -> bool:
        """Checks if a record with uni_code has already been processed successfully or skipped."""
        code_key = str(uni_code).strip()
        with self.lock:
            rec = self.state["records"].get(code_key)
            if not rec:
                return False
            status = rec.get("status")
            return status in [STATUS_SUCCESS, STATUS_SKIPPED_ALREADY_DOWNLOADED]

    def record_result(self, record_data: Dict[str, Any]):
        """
        Updates metrics and stores the result of a processed record.
        """
        code_key = str(record_data.get("uni_code", "")).strip()
        status = record_data.get("status", "")

        with self.lock:
            is_new = code_key not in self.state["records"]
            self.state["records"][code_key] = record_data

            if is_new:
                self.state["processed_count"] += 1
                if status == STATUS_SUCCESS:
                    self.state["success_count"] += 1
                elif status in [STATUS_SKIPPED_ALREADY_DOWNLOADED, STATUS_SKIPPED_NO_ANNUAL_REPORT, STATUS_SKIPPED_NO_PDF]:
                    self.state["skipped_count"] += 1
                elif status == STATUS_REVIEW_REQUIRED:
                    self.state["review_count"] += 1
                else:
                    self.state["failed_count"] += 1

        self.save_state()

    def get_metrics(self) -> Dict[str, Any]:
        """Returns snapshot of current metrics."""
        with self.lock:
            return {
                "total": self.state["total_records"],
                "processed": self.state["processed_count"],
                "success": self.state["success_count"],
                "skipped": self.state["skipped_count"],
                "failed": self.state["failed_count"],
                "review": self.state["review_count"]
            }

    def get_all_saved_records(self) -> List[Dict[str, Any]]:
        """Returns list of all saved records."""
        with self.lock:
            return list(self.state["records"].values())
