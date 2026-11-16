"""
logger.py - Logging setup with rotating file handler and GUI callback queue.
"""
import os
import sys
import logging
from logging.handlers import RotatingFileHandler
from typing import Callable, Optional
from config import DEFAULT_LOG_FILE

_logger: Optional[logging.Logger] = None
_gui_callback: Optional[Callable[[str, str], None]] = None

class GuiLogHandler(logging.Handler):
    """Custom logging handler that routes formatted log records to a GUI callback."""
    def __init__(self, callback: Callable[[str, str], None]):
        super().__init__()
        self.callback = callback

    def emit(self, record):
        try:
            msg = self.format(record)
            level = record.levelname
            if self.callback:
                self.callback(msg, level)
        except Exception:
            self.handleError(record)

def get_logger(log_file=DEFAULT_LOG_FILE, level=logging.INFO) -> logging.Logger:
    """Configures and returns the application singleton logger."""
    global _logger
    if _logger is not None:
        return _logger

    _logger = logging.getLogger("UK_NSM_Downloader")
    _logger.setLevel(level)
    _logger.propagate = False

    formatter = logging.Formatter(
        "[%(asctime)s] [%(levelname)-7s] %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S"
    )

    # Console Handler
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setFormatter(formatter)
    _logger.addHandler(console_handler)

    # Rotating File Handler (Max 10 MB, up to 5 backups)
    try:
        os.makedirs(os.path.dirname(str(log_file)), exist_ok=True)
        file_handler = RotatingFileHandler(
            str(log_file),
            maxBytes=10 * 1024 * 1024,
            backupCount=5,
            encoding="utf-8"
        )
        file_handler.setFormatter(formatter)
        _logger.addHandler(file_handler)
    except Exception as e:
        print(f"Warning: Could not initialize file logger: {e}", file=sys.stderr)

    return _logger

def set_gui_callback(callback: Callable[[str, str], None]):
    """Registers a callback for the GUI log display (msg, level)."""
    global _gui_callback, _logger
    _gui_callback = callback
    if _logger is None:
        get_logger()
    
    # Remove existing GuiLogHandlers if any
    for h in list(_logger.handlers):
        if isinstance(h, GuiLogHandler):
            _logger.removeHandler(h)

    gui_handler = GuiLogHandler(callback)
    gui_formatter = logging.Formatter(
        "[%(asctime)s] [%(levelname)s] %(message)s",
        datefmt="%H:%M:%S"
    )
    gui_handler.setFormatter(gui_formatter)
    _logger.addHandler(gui_handler)
