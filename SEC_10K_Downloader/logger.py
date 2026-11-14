"""
logger.py - Application logging system with file, console, and GUI callback handlers.
"""
import logging
from logging.handlers import RotatingFileHandler
from datetime import datetime
import threading
from typing import Callable, Optional
from config import DEFAULT_LOG_FILE

_logger = None
_lock = threading.Lock()
_gui_callback: Optional[Callable[[str, str], None]] = None

class GuiLogHandler(logging.Handler):
    """Custom logging handler that forwards log records to a GUI callback function."""
    def __init__(self, callback: Optional[Callable[[str, str], None]] = None):
        super().__init__()
        self.callback = callback

    def emit(self, record):
        try:
            msg = self.format(record)
            if self.callback:
                self.callback(msg, record.levelname)
        except Exception:
            self.handleError(record)

def setup_logger(log_file=DEFAULT_LOG_FILE, gui_callback: Optional[Callable[[str, str], None]] = None) -> logging.Logger:
    global _logger, _gui_callback
    with _lock:
        if _logger is not None:
            if gui_callback:
                _gui_callback = gui_callback
                for h in _logger.handlers:
                    if isinstance(h, GuiLogHandler):
                        h.callback = gui_callback
            return _logger

        _gui_callback = gui_callback
        _logger = logging.getLogger("SEC_10K_Downloader")
        _logger.setLevel(logging.INFO)
        _logger.handlers.clear()

        # Format: [10:05:12] Message
        gui_formatter = logging.Formatter("[%(asctime)s] %(message)s", datefmt="%H:%M:%S")
        file_formatter = logging.Formatter(
            "%(asctime)s [%(levelname)s] [%(filename)s:%(lineno)d]: %(message)s",
            datefmt="%Y-%m-%d %H:%M:%S"
        )

        # 1. Console Handler
        console_handler = logging.StreamHandler()
        console_handler.setFormatter(gui_formatter)
        console_handler.setLevel(logging.INFO)
        _logger.addHandler(console_handler)

        # 2. Rotating File Handler (max 10MB, keep 5 backups)
        file_handler = RotatingFileHandler(
            str(log_file), maxBytes=10 * 1024 * 1024, backupCount=5, encoding="utf-8"
        )
        file_handler.setFormatter(file_formatter)
        file_handler.setLevel(logging.DEBUG)
        _logger.addHandler(file_handler)

        # 3. GUI Handler
        if gui_callback:
            gui_h = GuiLogHandler(gui_callback)
            gui_h.setFormatter(gui_formatter)
            gui_h.setLevel(logging.INFO)
            _logger.addHandler(gui_h)

        return _logger

def get_logger() -> logging.Logger:
    global _logger
    if _logger is None:
        return setup_logger()
    return _logger

def set_gui_callback(callback: Callable[[str, str], None]):
    global _logger, _gui_callback
    _gui_callback = callback
    if _logger:
        has_gui_handler = False
        for h in _logger.handlers:
            if isinstance(h, GuiLogHandler):
                h.callback = callback
                has_gui_handler = True
        if not has_gui_handler:
            gui_formatter = logging.Formatter("[%(asctime)s] %(message)s", datefmt="%H:%M:%S")
            gui_h = GuiLogHandler(callback)
            gui_h.setFormatter(gui_formatter)
            gui_h.setLevel(logging.INFO)
            _logger.addHandler(gui_h)
