"""
progress_manager.py - Persistent state, pause/resume coordination, and progress tracking.
"""
import os
import json
import threading
from typing import Dict, Any, Optional
from datetime import datetime
from config import (
    DEFAULT_STATE_FILE,
    STATUS_SUCCESS,
    STATUS_SKIPPED_ALREADY_DOWNLOADED,
    STATUS_SKIPPED_NO_MATCHING_10K,
    STATUS_FAILED_COMPANY_NOT_FOUND,
    STATUS_FAILED_BROWSER_ERROR,
    STATUS_FAILED_DOWNLOAD_ERROR,
    STATUS_FAILED_INVALID_HTML,
    STATUS_REVIEW_REQUIRED
)
from logger import get_logger

logger = get_logger()

class ProgressManager:
    """
    Manages process state, atomic saves to JSON, and pause/resume threading events.
    """
    def __init__(self, state_file=DEFAULT_STATE_FILE):
        self.state_file = str(state_file)
        self.lock = threading.Lock()
        
        # Threading events for control flow
        self._pause_event = threading.Event()
        self._pause_event.set()  # Set means NOT paused (running)
        self._stop_event = threading.Event()  # Set means stop requested

        # State data
        self.state: Dict[str, Any] = {
            "total_records": 0,
            "processed_records": 0,
            "success_count": 0,
            "skipped_count": 0,
            "failed_count": 0,
            "review_count": 0,
            "last_updated": None,
            "records": {}  # unique_code -> record data dict
        }
        self.load_state()

    def reset_state(self, total_records: int):
        with self.lock:
            self.state = {
                "total_records": total_records,
                "processed_records": 0,
                "success_count": 0,
                "skipped_count": 0,
                "failed_count": 0,
                "review_count": 0,
                "last_updated": datetime.now().isoformat(),
                "records": {}
            }
            self._save_atomic()

    def load_state(self) -> bool:
        """Loads state from JSON file if available."""
        with self.lock:
            if os.path.exists(self.state_file):
                try:
                    with open(self.state_file, "r", encoding="utf-8") as f:
                        data = json.load(f)
                        self.state.update(data)
                        logger.info(f"Loaded existing progress state: {self.state.get('processed_records', 0)}/{self.state.get('total_records', 0)} processed.")
                        return True
                except Exception as e:
                    logger.warning(f"Could not load state file: {e}")
            return False

    def save_record_result(self, record_code: str, result_dict: Dict[str, Any]):
        """Updates record status and updates aggregate counts."""
        with self.lock:
            status = result_dict.get("download_status", "")
            prev_record = self.state["records"].get(record_code)
            
            # If not previously counted as processed
            if not prev_record:
                self.state["processed_records"] += 1

            self.state["records"][record_code] = result_dict
            self.state["last_updated"] = datetime.now().isoformat()

            # Recalculate summary counts
            success = 0
            skipped = 0
            failed = 0
            review = 0

            for r in self.state["records"].values():
                st = r.get("download_status", "")
                if st == STATUS_SUCCESS:
                    success += 1
                elif st in [STATUS_SKIPPED_ALREADY_DOWNLOADED, STATUS_SKIPPED_NO_MATCHING_10K]:
                    skipped += 1
                elif st in [STATUS_FAILED_COMPANY_NOT_FOUND, STATUS_FAILED_BROWSER_ERROR, 
                            STATUS_FAILED_DOWNLOAD_ERROR, STATUS_FAILED_INVALID_HTML]:
                    failed += 1
                elif st == STATUS_REVIEW_REQUIRED:
                    review += 1

            self.state["success_count"] = success
            self.state["skipped_count"] = skipped
            self.state["failed_count"] = failed
            self.state["review_count"] = review

            self._save_atomic()

    def _save_atomic(self):
        """Atomically saves state to avoid corruption."""
        os.makedirs(os.path.dirname(self.state_file), exist_ok=True)
        temp_file = self.state_file + ".tmp"
        try:
            with open(temp_file, "w", encoding="utf-8") as f:
                json.dump(self.state, f, indent=2)
            if os.path.exists(self.state_file):
                try:
                    os.remove(self.state_file)
                except OSError:
                    pass
            os.replace(temp_file, self.state_file)
        except Exception as e:
            logger.error(f"Failed to save state file: {e}")

    def is_already_completed(self, unique_code: str) -> bool:
        """Returns True if record has already been successfully downloaded."""
        with self.lock:
            r = self.state["records"].get(unique_code)
            if r and r.get("download_status") in [STATUS_SUCCESS, STATUS_SKIPPED_ALREADY_DOWNLOADED]:
                return True
            return False

    def get_record(self, unique_code: str) -> Optional[Dict[str, Any]]:
        with self.lock:
            return self.state["records"].get(unique_code)

    def get_summary_stats(self) -> Dict[str, Any]:
        with self.lock:
            total = self.state.get("total_records", 0)
            processed = self.state.get("processed_records", 0)
            pct = (processed / total * 100) if total > 0 else 0.0
            return {
                "total": total,
                "processed": processed,
                "success": self.state.get("success_count", 0),
                "skipped": self.state.get("skipped_count", 0),
                "failed": self.state.get("failed_count", 0),
                "review": self.state.get("review_count", 0),
                "completion_pct": pct
            }

    # Pause / Resume / Stop Controls
    def pause(self):
        logger.info("Automation pause requested.")
        self._pause_event.clear()

    def resume(self):
        logger.info("Automation resumed.")
        self._pause_event.set()

    def stop(self):
        logger.info("Automation stop requested.")
        self._stop_event.set()
        self._pause_event.set()  # Unblock if paused so thread can exit

    def is_paused(self) -> bool:
        return not self._pause_event.is_set()

    def is_stopped(self) -> bool:
        return self._stop_event.is_set()

    def wait_if_paused(self):
        """Blocks execution while paused. Checks stop event periodically."""
        while not self._pause_event.is_set():
            if self._stop_event.is_set():
                break
            self._pause_event.wait(timeout=0.5)
