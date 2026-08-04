"""
Word Page Exporter Pro - Logging Module
Provides structured logging with console output, file logging, and GUI event listeners.
"""

import os
import sys
import logging
import tempfile
from datetime import datetime
from typing import Callable, List, Optional


class AppLogger:
    """Centralized logging manager with event listener support for GUI."""
    
    _instance: Optional['AppLogger'] = None

    def __init__(self, log_to_file: bool = True, log_dir: Optional[str] = None):
        self.listeners: List[Callable[[str, str, str], None]] = []
        self.logger = logging.getLogger("WordPageExporterPro")
        self.logger.setLevel(logging.DEBUG)
        
        # Prevent duplicate handlers
        if not self.logger.handlers:
            # Console handler
            console_handler = logging.StreamHandler(sys.stdout)
            console_handler.setLevel(logging.INFO)
            formatter = logging.Formatter("[%(asctime)s] [%(levelname)s] %(message)s", "%H:%M:%S")
            console_handler.setFormatter(formatter)
            self.logger.addHandler(console_handler)

            # Optional file handler
            if log_to_file:
                if not log_dir:
                    log_dir = os.getenv(
                        "WORD_EXPORTER_LOG_DIR",
                        os.path.join(tempfile.gettempdir(), "word_exporter_pro", "logs"),
                    )
                try:
                    os.makedirs(log_dir, exist_ok=True)
                    log_file = os.path.join(log_dir, f"exporter_{datetime.now().strftime('%Y%m%d')}.log")
                    file_handler = logging.FileHandler(log_file, encoding="utf-8")
                    file_handler.setLevel(logging.DEBUG)
                    file_formatter = logging.Formatter("[%(asctime)s] [%(levelname)s] %(name)s: %(message)s")
                    file_handler.setFormatter(file_formatter)
                    self.logger.addHandler(file_handler)
                except OSError:
                    # Container platforms still collect stdout; a file-log
                    # permission problem must not prevent application startup.
                    pass

    @classmethod
    def get_logger(cls) -> 'AppLogger':
        if cls._instance is None:
            cls._instance = AppLogger()
        return cls._instance

    def add_listener(self, listener: Callable[[str, str, str], None]):
        """Register a callback for log updates: listener(timestamp, level, message)"""
        if listener not in self.listeners:
            self.listeners.append(listener)

    def remove_listener(self, listener: Callable[[str, str, str], None]):
        if listener in self.listeners:
            self.listeners.remove(listener)

    def _emit(self, level: str, message: str):
        timestamp = datetime.now().strftime("%H:%M:%S")
        for listener in self.listeners:
            try:
                listener(timestamp, level, message)
            except Exception as e:
                print(f"Error in log listener: {e}", file=sys.stderr)

    def info(self, message: str):
        self.logger.info(message)
        self._emit("INFO", message)

    def warning(self, message: str):
        self.logger.warning(message)
        self._emit("WARNING", message)

    def error(self, message: str):
        self.logger.error(message)
        self._emit("ERROR", message)

    def success(self, message: str):
        self.logger.info(f"SUCCESS: {message}")
        self._emit("SUCCESS", message)

    def debug(self, message: str):
        self.logger.debug(message)
        self._emit("DEBUG", message)


def get_logger() -> AppLogger:
    return AppLogger.get_logger()
