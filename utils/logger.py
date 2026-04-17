from __future__ import annotations
import logging
from pathlib import Path
from PySide6.QtCore import QObject, Signal

class LogEmitter(QObject):
    log_signal = Signal(str)

class GuiLogHandler(logging.Handler):
    def __init__(self, emitter: LogEmitter):
        super().__init__()
        self.emitter = emitter
        self.setFormatter(logging.Formatter('[%(asctime)s] %(message)s', '%H:%M:%S'))

    def emit(self, record: logging.LogRecord):
        msg = self.format(record)
        self.emitter.log_signal.emit(msg)

log_file = Path("app.log")
logging.basicConfig(
    level=logging.INFO,
    format='[%(asctime)s] %(levelname)s - %(message)s',
    datefmt='%H:%M:%S',
    handlers=[
        logging.FileHandler(log_file, encoding='utf-8'),
        logging.StreamHandler()
    ]
)

logger = logging.getLogger("AppLogger")
gui_emitter = LogEmitter()
logger.addHandler(GuiLogHandler(gui_emitter))