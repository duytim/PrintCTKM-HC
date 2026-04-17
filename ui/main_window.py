from __future__ import annotations
import threading
from pathlib import Path
from PySide6.QtWidgets import (QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
                               QLabel, QLineEdit, QPushButton, QTextEdit, 
                               QTabWidget, QComboBox, QFileDialog, QMessageBox, 
                               QGroupBox, QCheckBox, QFormLayout, QProgressBar, QApplication)
from PySide6.QtCore import Qt
from PySide6.QtGui import QIcon, QFont

from core.config_manager import ConfigManager
# XÓA IMPORT: Không import thư viện xử lý nặng (MainWorker, DemoWorker) ở đầu file nữa
from utils.logger import gui_emitter

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Tool Bảng Giá Pro v3.0 - Enterprise Edition")
        self.resize(900, 780)
        if Path("logo.ico").exists(): self.setWindowIcon(QIcon("logo.ico"))
            
        self.cfg_mgr = ConfigManager()
        self.last_pdf: str = ""
        self.init_ui()
        gui_emitter.log_signal.connect(self.log.append)
        
        # CHẠY LUỒNG NGẦM TẢI THƯ VIỆN NGAY KHI GIAO DIỆN VỪA KHỞI TẠO XONG
        threading.Thread(target=self.preload_libraries, daemon=True).start()

    def preload_libraries(self) -> None:
        """Hàm chạy ngầm để tải trước các module nặng (pandas, requests, docxtpl, pypdf...) vào Cache của Python"""
        try:
            self.statusBar().showMessage("Hệ thống đang khởi động các module lõi dưới nền...")
            
            # Tải ngầm workers. Toàn bộ thời gian chờ 30s sẽ diễn ra ở đây mà không làm đơ giao diện
            import ui.workers
            
            self.statusBar().showMessage("Sẵn sàng hoạt động.")
        except Exception as e:
            self.statusBar().showMessage(f"Lỗi khởi động nền: {e}")

    def init_ui(self) -> None:
        w = QWidget(); self.setCentralWidget(w); layout = QVBoxLayout(w)
        tabs = QTabWidget(); layout.addWidget(tabs)

        t1 = QWidget(); l1 = QVBoxLayout(t1); l1.setSpacing(15)
        
        gb_config = QGroupBox("Cài Đặt Chạy"); gb_config.setFont(QFont("Arial", 10, QFont.Bold))
        fl = QFormLayout(gb_config); fl.setVerticalSpacing(12)
        
        self.txt_sub = QLineEdit(self.cfg_mgr.get('mail_subjects')); fl.addRow("Từ Khóa Mail:", self.txt_sub)
        self.txt_kho = QLineEdit(self.cfg_mgr.get('target_kho')); fl.addRow("Mã Kho Đích:", self.txt_kho)
        
        hl1 = QHBoxLayout()
        self.cbb_fmt = QComboBox(); self.cbb_fmt.addItems(["A4", "A5", "A7"]); self.cbb_fmt.setCurrentText(self.cfg_mgr.get('pdf_format'))
        hl1.addWidget(self.cbb_fmt)
        btn_tpl = QPushButton("📝 Mở Phôi"); btn_tpl.clicked.connect(self.open_tpl)
        btn_xls = QPushButton("📂 Mở Excel Cũ"); btn_xls.clicked.connect(self.open_xls)
        hl1.addWidget(btn_tpl); hl1.addWidget(btn_xls); fl.addRow("Định dạng In:", hl1)
        
        hl2 = QHBoxLayout()
        self.chk_img = QCheckBox("Thay Banner"); self.chk_img.setChecked(self.cfg_mgr.get('use_image_replace'))
        self.txt_img = QLineEdit(self.cfg_mgr.get('replacement_image'))
        btn_img = QPushButton("Chọn Ảnh..."); btn_img.clicked.connect(self.pick_img)
        self.btn_demo = QPushButton("👁️ Xem Demo"); self.btn_demo.setStyleSheet("background: #17a2b8; color: white; font-weight: bold;")
        self.btn_demo.clicked.connect(self.run_demo)
        
        hl2.addWidget(self.chk_img); hl2.addWidget(self.txt_img); hl2.addWidget(btn_img); hl2.addWidget(self.btn_demo)
        fl.addRow("Banner (Tùy chọn):", hl2); l1.addWidget(gb_config)

        self.chk_img.stateChanged.connect(self.check_demo_btn_visibility)
        self.txt_img.textChanged.connect(self.check_demo_btn_visibility)
        self.check_demo_btn_visibility()

        gb_actions = QGroupBox("Chức Năng Chính"); gb_actions.setFont(QFont("Arial", 10, QFont.Bold))
        vbox_actions = QVBoxLayout(gb_actions); vbox_actions.setSpacing(10)
        
        self.btns = []
        b1 = QPushButton("🚀 1. CHẠY TOÀN BỘ QUY TRÌNH (Tải Kho -> Mail -> In PDF)")
        b1.setStyleSheet("background: #0056b3; color: white; font-size: 14px; font-weight: bold; padding: 12px; border-radius: 4px;")
        b1.clicked.connect(lambda: self.run_task(0))
        
        b2 = QPushButton("🖨️ 2. IN PDF (Đã có Tồn Kho sẵn)")
        b2.setStyleSheet("background: #fd7e14; color: white; font-weight: bold; padding: 8px; border-radius: 4px;")
        b2.clicked.connect(lambda: self.run_task(1))
        
        b3 = QPushButton("♻️ 3. CHỈ TẠO LẠI PDF (Sử dụng dữ liệu cuối cùng)")
        b3.setStyleSheet("background: #28a745; color: white; font-weight: bold; padding: 8px; border-radius: 4px;")
        b3.clicked.connect(lambda: self.run_task(2))

        self.btn_open_last = QPushButton("📄 Mở PDF Cuối Cùng")
        self.btn_open_last.setStyleSheet("background: #6c757d; color: white; font-weight: bold; padding: 8px; border-radius: 4px;")
        self.btn_open_last.clicked.connect(self.open_last_pdf)
        
        for b in [b1, b2, b3, self.btn_open_last]:
            vbox_actions.addWidget(b)
            if b != self.btn_open_last: self.btns.append(b)

        l1.addWidget(gb_actions); l1.addStretch(); tabs.addTab(t1, "⚙️ Vận Hành")

        t2 = QWidget(); l2 = QFormLayout(t2)
        self.u_erp = QLineEdit(self.cfg_mgr.get('erp_user')); l2.addRow("ERP User:", self.u_erp)
        self.p_erp = QLineEdit(self.cfg_mgr.get('erp_pass')); self.p_erp.setEchoMode(QLineEdit.EchoMode.Password); l2.addRow("ERP Pass:", self.p_erp)
        self.u_mail = QLineEdit(self.cfg_mgr.get('mail_user')); l2.addRow("Mail User:", self.u_mail)
        self.p_mail = QLineEdit(self.cfg_mgr.get('mail_pass')); self.p_mail.setEchoMode(QLineEdit.EchoMode.Password); l2.addRow("Mail Pass:", self.p_mail)
        hl3 = QHBoxLayout(); self.txt_lib = QLineEdit(self.cfg_mgr.get('libre_path')); b_lib = QPushButton("Tìm Soft"); b_lib.clicked.connect(self.pick_lib)
        hl3.addWidget(self.txt_lib); hl3.addWidget(b_lib); l2.addRow("LibreOffice Path:", hl3)
        b_save = QPushButton("💾 Lưu Cấu Hình"); b_save.clicked.connect(lambda: (self.save_cfg(), QMessageBox.information(self,"OK","Đã lưu!")))
        l2.addRow(b_save); tabs.addTab(t2, "🔧 Cấu Hình")

        layout.addWidget(QLabel("Nhật ký làm việc:"))
        self.log = QTextEdit(); self.log.setReadOnly(True)
        self.log.setStyleSheet("background: #1e1e1e; color: #00ff00; font-family: Consolas;")
        layout.addWidget(self.log)

        self.progress_bar = QProgressBar(); self.progress_bar.setValue(0)
        self.progress_bar.setAlignment(Qt.AlignmentFlag.AlignCenter); self.progress_bar.setFormat("Sẵn sàng (0%)")
        layout.addWidget(self.progress_bar)
        self.statusBar().showMessage("Sẵn sàng hoạt động.")

    def check_demo_btn_visibility(self) -> None:
        self.btn_demo.setVisible(self.chk_img.isChecked() and bool(self.txt_img.text().strip()))

    def save_cfg(self) -> None:
        for k, v in [('erp_user', self.u_erp), ('erp_pass', self.p_erp), ('mail_user', self.u_mail), ('mail_pass', self.p_mail),
                     ('target_kho', self.txt_kho), ('mail_subjects', self.txt_sub), ('pdf_format', self.cbb_fmt),
                     ('libre_path', self.txt_lib), ('replacement_image', self.txt_img)]:
            self.cfg_mgr.set(k, v.currentText() if isinstance(v, QComboBox) else v.text())
        self.cfg_mgr.set('use_image_replace', self.chk_img.isChecked()); self.cfg_mgr.save()

    def pick_img(self) -> None:
        f, _ = QFileDialog.getOpenFileName(self, "Chọn ảnh Banner", "", "Image (*.png *.jpg *.jpeg)")
        if f: self.txt_img.setText(f)
        
    def pick_lib(self) -> None:
        f, _ = QFileDialog.getOpenFileName(self, "Chọn soffice.exe", "", "Exe (*.exe)")
        if f: self.txt_lib.setText(f)

    def open_tpl(self) -> None:
        f = Path("templates") / f"{self.cbb_fmt.currentText()}-Auto.docx"
        if f.exists(): __import__('os').startfile(f)

    def open_xls(self) -> None:
        f = Path("Data_CSV") / "filtered_email_tables.csv"
        if f.exists(): __import__('os').startfile(f)

    def open_last_pdf(self) -> None:
        if self.last_pdf and Path(self.last_pdf).exists(): __import__('os').startfile(self.last_pdf)

    def update_progress(self, val: int, text: str) -> None:
        self.progress_bar.setValue(val); self.progress_bar.setFormat(f"{text} ({val}%)")

    def lock_ui(self, lock: bool) -> None:
        for b in self.btns: b.setEnabled(not lock)
        self.btn_demo.setEnabled(not lock)

    def run_demo(self) -> None:
        self.save_cfg(); self.log.clear(); self.lock_ui(True); self.statusBar().showMessage("Đang tạo Demo...")
        
        # Ép giao diện vẽ lại ngay để hiện thông báo ngay lập tức
        QApplication.processEvents()
        
        # LAZY IMPORT (Nhờ pre-load chạy ngầm, lệnh này lấy cache siêu tốc chỉ mất ~0.001s)
        from ui.workers import DemoWorker
        
        self.worker = DemoWorker(self.cfg_mgr.config)
        self.worker.progress_sig.connect(self.update_progress)
        self.worker.finished_sig.connect(self.on_finished); self.worker.start()

    def run_task(self, mode: int) -> None:
        self.save_cfg(); self.log.clear(); self.lock_ui(True); self.statusBar().showMessage("Đang chạy...")
        
        # Ép giao diện vẽ lại ngay để hiện thông báo ngay lập tức
        QApplication.processEvents()
        
        # LAZY IMPORT (Nhờ pre-load chạy ngầm, lệnh này lấy cache siêu tốc chỉ mất ~0.001s)
        from ui.workers import MainWorker
        
        self.worker = MainWorker(self.cfg_mgr.config, mode)
        self.worker.progress_sig.connect(self.update_progress)
        self.worker.finished_sig.connect(self.on_finished); self.worker.start()

    def on_finished(self, success: bool, pdf_path: str) -> None:
        self.lock_ui(False); self.statusBar().showMessage("Sẵn sàng." if success else "Có lỗi!")
        if success and pdf_path: self.last_pdf = pdf_path; __import__('os').startfile(pdf_path)