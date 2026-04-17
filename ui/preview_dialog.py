from __future__ import annotations
from pathlib import Path
from PySide6.QtWidgets import QDialog, QVBoxLayout, QLabel
from PySide6.QtCore import QUrl
from PySide6.QtGui import Qt

try:
    from PySide6.QtWebEngineWidgets import QWebEngineView
    from PySide6.QtWebEngineCore import QWebEngineSettings
    HAS_WEBENGINE = True
except ImportError:
    HAS_WEBENGINE = False

class PreviewDialog(QDialog):
    def __init__(self, pdf_path: Path, parent=None):
        super().__init__(parent)
        self.setWindowTitle("👀 Xem trước PDF")
        self.resize(800, 900)
        layout = QVBoxLayout(self)
        
        if HAS_WEBENGINE:
            self.viewer = QWebEngineView()
            self.viewer.settings().setAttribute(QWebEngineSettings.WebAttribute.PluginsEnabled, True)
            self.viewer.settings().setAttribute(QWebEngineSettings.WebAttribute.PdfViewerEnabled, True)
            self.viewer.load(QUrl.fromLocalFile(str(pdf_path.resolve())))
            layout.addWidget(self.viewer)
        else:
            lbl = QLabel(f"Thiếu module QtWebEngine.\nFile: {pdf_path}")
            lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
            layout.addWidget(lbl)