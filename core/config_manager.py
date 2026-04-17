from __future__ import annotations
import json
from pathlib import Path
from typing import Any, Dict
from utils.logger import logger

class ConfigManager:
    DEFAULT_CONFIG = {
        'erp_user': '', 'erp_pass': '',
        'mail_user': '', 'mail_pass': '',
        'target_kho': '046_01',
        'mail_subjects': 'SP.073636',
        'pdf_format': 'A4',
        'libre_path': '',
        'replacement_image': '',
        'use_image_replace': False
    }

    def __init__(self, config_path: Path | str = "config.json"):
        self.config_path = Path(config_path)
        self.config: Dict[str, Any] = self.DEFAULT_CONFIG.copy()
        self.load()

    def load(self) -> None:
        if self.config_path.exists():
            try:
                data = json.loads(self.config_path.read_text(encoding='utf-8'))
                self.config.update(data)
            except Exception as e:
                logger.error(f"Lỗi đọc config: {e}")
        else:
            self.save()

    def save(self) -> None:
        self.config_path.write_text(json.dumps(self.config, indent=4, ensure_ascii=False), encoding='utf-8')

    def get(self, key: str, default: Any = None) -> Any:
        return self.config.get(key, default)

    def set(self, key: str, value: Any) -> None:
        self.config[key] = value