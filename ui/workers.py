from __future__ import annotations
from pathlib import Path
from typing import Dict
from PySide6.QtCore import QThread, Signal

from core.data_processor import DataProcessor
from core.pdf_engine import PdfEngine
from utils.logger import logger

class MainWorker(QThread):
    finished_sig = Signal(bool, str) 
    progress_sig = Signal(int, str)

    def __init__(self, config: Dict, mode: int = 0):
        super().__init__()
        self.config = config
        self.mode = mode

    def run(self) -> None:
        logger.info("="*30); logger.info("🚀 BẮT ĐẦU CHẠY")
        processor, engine = DataProcessor(self.config), PdfEngine(self.config)
        success, pdf_path = True, None
        csv_file = processor.csv_dir / "filtered_email_tables.csv"
        cb = lambda val, txt: self.progress_sig.emit(val, txt)

        if self.mode in [0, 1]:
            tonkho_file = processor.fetch_inventory(use_existing=(self.mode == 1), progress_cb=cb)
            if tonkho_file:
                mail_file = processor.fetch_email(progress_cb=cb)
                if mail_file:
                    if not processor.merge_and_calculate(mail_file, tonkho_file, progress_cb=cb): success = False
                else: success = False
            else: success = False

        if success or self.mode == 2:
            if csv_file.exists():
                pdf_path = engine.generate_pdf(csv_file, progress_cb=cb)
                if not pdf_path: success = False
            else:
                logger.error("❌ Không tìm thấy dữ liệu cũ"); success = False; cb(0, "Lỗi dữ liệu")
                
        self.finished_sig.emit(success, str(pdf_path) if pdf_path else "")

class DemoWorker(QThread):
    finished_sig = Signal(bool, str)
    progress_sig = Signal(int, str)

    def __init__(self, config: Dict):
        super().__init__()
        self.config = config

    def run(self) -> None:
        engine = PdfEngine(self.config)
        pdf_path = engine.generate_demo_pdf(progress_cb=lambda val, txt: self.progress_sig.emit(val, txt))
        self.finished_sig.emit(bool(pdf_path), str(pdf_path) if pdf_path else "")