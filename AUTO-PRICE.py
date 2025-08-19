#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
AUTO-PRICE - Ứng dụng tự động hóa bảng giá A4 và A5
Kết hợp tính năng từ A4-AUTO.py và A5-AUTO.py với giao diện thống nhất
"""

import pandas as pd
from docxtpl import DocxTemplate
import os
from pypdf import PdfWriter
import subprocess
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, font
import json
import time
import threading
import warnings
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import platform
import shutil

# Tắt warnings không cần thiết
warnings.filterwarnings("ignore", category=UserWarning, module="docxcompose")
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

class PDFConverter:
    """Lớp chuyển đổi PDF với nhiều công cụ khác nhau"""
    
    def __init__(self, config):
        self.config = config
        self.available_tools = self.detect_available_tools()
        self.current_tool = self.select_best_tool()
    
    def detect_available_tools(self):
        """Phát hiện các công cụ chuyển đổi PDF có sẵn"""
        tools = {}
        
        # 1. Kiểm tra python-docx2pdf (CÔNG CỤ TỐT NHẤT)
        try:
            import docx2pdf
            tools['docx2pdf'] = {
                'name': 'python-docx2pdf',
                'priority': 1,
                'available': True,
                'description': '🚀 THỦ VIỆN PYTHON THUẦN TÚY - NHANH VÀ NHẸ NHẤT (Khuyến nghị)'
            }
        except ImportError:
            tools['docx2pdf'] = {
                'name': 'python-docx2pdf',
                'priority': 1,
                'available': False,
                'description': 'Cần cài đặt: pip install docx2pdf (Khuyến nghị)'
            }
        
        # 2. Kiểm tra LibreOffice (hỗ trợ cả bản cài đặt và Portable)
        configured_path = self.config.get('libreoffice_path', r"C:\\Program Files\\LibreOffice\\program\\soffice.exe")
        env_path = os.environ.get('LIBREOFFICE_PATH')
        candidate_paths = [
            configured_path,
            env_path if env_path else '',
            os.path.join(os.getcwd(), 'LibreOfficePortable', 'App', 'libreoffice', 'program', 'soffice.exe'),
            os.path.join(os.getcwd(), 'LibreOfficePortable', 'App', 'libreoffice', 'program', 'soffice.com'),
            os.path.join(os.getcwd(), 'libreoffice', 'program', 'soffice.exe'),
            os.path.join(os.getcwd(), 'libreoffice', 'program', 'soffice.com'),
            os.path.join(os.getcwd(), 'bin', 'LibreOffice', 'program', 'soffice.exe'),
            os.path.join(os.getcwd(), 'bin', 'LibreOffice', 'program', 'soffice.com'),
        ]
        detected_lo_path = next((p for p in candidate_paths if p and os.path.exists(p)), None)
        if detected_lo_path:
            tools['libreoffice'] = {
                'name': 'LibreOffice',
                'priority': 2,
                'available': True,
                'description': 'LibreOffice/LibreOffice Portable (fallback, không cần cài đặt nếu dùng Portable)',
                'path': detected_lo_path
            }
        else:
            tools['libreoffice'] = {
                'name': 'LibreOffice',
                'priority': 2,
                'available': False,
                'description': 'Không tìm thấy LibreOffice. Có thể dùng LibreOffice Portable (giải nén cạnh ứng dụng) hoặc chọn đường dẫn thủ công.'
            }
        
        return tools
    
    def select_best_tool(self):
        """Chọn công cụ tốt nhất có sẵn"""
        available_tools = [tool for tool, info in self.available_tools.items() 
                          if info['available']]
        
        if not available_tools:
            raise Exception("Không tìm thấy công cụ chuyển đổi PDF nào!")
        
        # Chọn công cụ có priority thấp nhất (tốt nhất)
        best_tool = min(available_tools, 
                       key=lambda x: self.available_tools[x]['priority'])
        
        return best_tool
    
    def convert_to_pdf(self, word_file):
        """Chuyển đổi Word sang PDF sử dụng công cụ được chọn"""
        try:
            # Ưu tiên sử dụng python-docx2pdf (công cụ tốt nhất)
            if self.current_tool == 'docx2pdf':
                result = self._convert_with_docx2pdf(word_file)
                if result:
                    return result
                else:
                    print("⚠️ python-docx2pdf thất bại, thử fallback...")
                    return self._fallback_convert(word_file)
            elif self.current_tool == 'libreoffice':
                return self._convert_with_libreoffice(word_file)
            else:
                raise Exception(f"Công cụ không được hỗ trợ: {self.current_tool}")
                
        except Exception as e:
            print(f"❌ Lỗi với {self.current_tool}: {e}")
            # Thử fallback sang công cụ khác
            return self._fallback_convert(word_file)
    
    def _convert_with_docx2pdf(self, word_file):
        """Chuyển đổi sử dụng python-docx2pdf (CÔNG CỤ TỐT NHẤT)"""
        try:
            import docx2pdf
            pdf_file = os.path.splitext(word_file)[0] + ".pdf"
            
            # Tối ưu hóa cho docx2pdf
            print(f"🔄 Đang chuyển đổi với python-docx2pdf (công cụ tốt nhất)...")
            
            # Chuyển đổi với timeout và error handling tốt hơn
            docx2pdf.convert(word_file, pdf_file)
            
            # Kiểm tra kết quả
            if os.path.exists(pdf_file):
                file_size = os.path.getsize(pdf_file)
                print(f"✅ Chuyển đổi thành công: {os.path.basename(pdf_file)} ({file_size:,} bytes)")
                return pdf_file
            else:
                print(f"❌ Không tạo được file PDF: {pdf_file}")
                return None
                
        except Exception as e:
            print(f"❌ Lỗi docx2pdf: {e}")
            return None
    
    def _convert_with_libreoffice(self, word_file):
        """Chuyển đổi sử dụng LibreOffice (fallback)"""
        try:
            libreoffice_path = self.available_tools['libreoffice']['path']
            
            result = subprocess.run([
                libreoffice_path,
                "--headless", "--convert-to", "pdf",
                word_file, "--outdir", os.path.dirname(word_file)
            ], check=True, capture_output=True, timeout=60)
            
            pdf_file = os.path.splitext(word_file)[0] + ".pdf"
            
            if os.path.exists(pdf_file):
                return pdf_file
            else:
                # Thử tìm file trong thư mục con
                subfolder = os.path.join(os.path.dirname(word_file), "In_PDF")
                if os.path.exists(subfolder):
                    subfolder_pdf = os.path.join(subfolder, os.path.basename(pdf_file))
                    if os.path.exists(subfolder_pdf):
                        return subfolder_pdf
                
                return None
                
        except subprocess.TimeoutExpired:
            print(f"Timeout LibreOffice: {word_file}")
            return None
        except Exception as e:
            print(f"Lỗi LibreOffice: {e}")
            return None
    
    def _fallback_convert(self, word_file):
        """Fallback sang công cụ khác nếu công cụ chính lỗi"""
        available_tools = [tool for tool, info in self.available_tools.items() 
                          if info['available'] and tool != self.current_tool]
        
        # Sắp xếp theo độ ưu tiên (docx2pdf trước)
        available_tools.sort(key=lambda x: self.available_tools[x]['priority'])
        
        for tool in available_tools:
            try:
                print(f"🔄 Thử fallback sang {tool}...")
                if tool == 'docx2pdf':
                    result = self._convert_with_docx2pdf(word_file)
                elif tool == 'libreoffice':
                    result = self._convert_with_libreoffice(word_file)
                
                if result:
                    print(f"✅ Fallback thành công với {tool}")
                    return result
                    
            except Exception as e:
                print(f"❌ Fallback {tool} thất bại: {e}")
                continue
        
        print("❌ Tất cả công cụ fallback đều thất bại!")
        return None
    
    def get_tool_info(self):
        """Lấy thông tin về công cụ hiện tại"""
        return self.available_tools.get(self.current_tool, {})
    
    def get_all_tools_info(self):
        """Lấy thông tin về tất cả công cụ"""
        return self.available_tools

class ConfigManager:
    def __init__(self, config_file="config.json"):
        self.config_file = config_file
        self.default_config = {
            "a4_excel_file": "A4-Auto.xlsx",
            "a4_word_template": "A4-Auto.docx",
            "a5_excel_file": "A5-AUTO.xlsx",
            "a5_word_template": "A5-AUTO.docx",
            "output_folder": "In_PDF",
            "libreoffice_path": r"C:\\Program Files\\LibreOffice\\program\\soffice.exe",
            "batch_size": 10,
            "default_format": "A5",  # A4 hoặc A5
            "preferred_tool": ""  # ''=auto; 'docx2pdf' hoặc 'libreoffice'
        }
        self.config = self.load_config()
    
    def load_config(self):
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    return {**self.default_config, **json.load(f)}
            else:
                self.save_config(self.default_config)
                return self.default_config
        except Exception as e:
            print(f"Lỗi khi đọc config: {e}")
            return self.default_config
    
    def save_config(self, config):
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(config, f, indent=2, ensure_ascii=False)
        except Exception as e:
            print(f"Lỗi khi lưu config: {e}")
    
    def update_config(self, key, value):
        self.config[key] = value
        self.save_config(self.config)

class ProgressTracker:
    def __init__(self, total_steps):
        self.total_steps = total_steps
        self.current_step = 0
    
    def update(self, step=1):
        self.current_step += step
    
    def get_progress(self):
        return (self.current_step / self.total_steps) * 100 if self.total_steps > 0 else 0

class AutoPriceGUI:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("AUTO-PRICE - Tự động hóa bảng giá A4 & A5")
        self.root.geometry("900x800")
        self.root.resizable(True, True)
        
        self.config = ConfigManager()
        self.converter = PDFConverter(self.config.config)
        self.setup_ui()
        self.refresh_tools()
        
    def setup_ui(self):
        # Frame chính
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Tiêu đề
        title_label = ttk.Label(main_frame, text="AUTO-PRICE - Tự động hóa bảng giá", 
                               font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # Chọn format
        format_frame = ttk.LabelFrame(main_frame, text="Chọn định dạng", padding="10")
        format_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        
        self.format_var = tk.StringVar(value=self.config.config["default_format"])
        ttk.Radiobutton(format_frame, text="A4 - Một sản phẩm/trang", 
                       variable=self.format_var, value="A4", 
                       command=self.on_format_change).grid(row=0, column=0, padx=(0, 20))
        ttk.Radiobutton(format_frame, text="A5 - Hai sản phẩm/trang", 
                       variable=self.format_var, value="A5", 
                       command=self.on_format_change).grid(row=0, column=1)
        
        # Cấu hình file A4
        self.a4_frame = ttk.LabelFrame(main_frame, text="Cấu hình A4", padding="10")
        self.a4_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # Excel file A4
        ttk.Label(self.a4_frame, text="File Excel A4:").grid(row=0, column=0, sticky=tk.W)
        self.a4_excel_var = tk.StringVar(value=self.config.config["a4_excel_file"])
        a4_excel_entry = ttk.Entry(self.a4_frame, textvariable=self.a4_excel_var, width=50)
        a4_excel_entry.grid(row=0, column=1, padx=(10, 5))
        ttk.Button(self.a4_frame, text="Chọn", command=self.browse_a4_excel).grid(row=0, column=2)
        ttk.Button(self.a4_frame, text="Mở", command=self.open_a4_excel).grid(row=0, column=3, padx=(5, 0))
        
        # Word template A4
        ttk.Label(self.a4_frame, text="Template Word A4:").grid(row=1, column=0, sticky=tk.W, pady=(10, 0))
        self.a4_template_var = tk.StringVar(value=self.config.config["a4_word_template"])
        a4_template_entry = ttk.Entry(self.a4_frame, textvariable=self.a4_template_var, width=50)
        a4_template_entry.grid(row=1, column=1, padx=(10, 5), pady=(10, 0))
        ttk.Button(self.a4_frame, text="Chọn", command=self.browse_a4_template).grid(row=1, column=2, pady=(10, 0))
        ttk.Button(self.a4_frame, text="Mở", command=self.open_a4_template).grid(row=1, column=3, pady=(10, 0), padx=(5, 0))
        
        # Cấu hình file A5
        self.a5_frame = ttk.LabelFrame(main_frame, text="Cấu hình A5", padding="10")
        self.a5_frame.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # Excel file A5
        ttk.Label(self.a5_frame, text="File Excel A5:").grid(row=0, column=0, sticky=tk.W)
        self.a5_excel_var = tk.StringVar(value=self.config.config["a5_excel_file"])
        a5_excel_entry = ttk.Entry(self.a5_frame, textvariable=self.a5_excel_var, width=50)
        a5_excel_entry.grid(row=0, column=1, padx=(10, 5))
        ttk.Button(self.a5_frame, text="Chọn", command=self.browse_a5_excel).grid(row=0, column=2)
        ttk.Button(self.a5_frame, text="Mở", command=self.open_a5_excel).grid(row=0, column=3, padx=(5, 0))
        
        # Word template A5
        ttk.Label(self.a5_frame, text="Template Word A5:").grid(row=1, column=0, sticky=tk.W, pady=(10, 0))
        self.a5_template_var = tk.StringVar(value=self.config.config["a5_word_template"]) 
        a5_template_entry = ttk.Entry(self.a5_frame, textvariable=self.a5_template_var, width=50)
        a5_template_entry.grid(row=1, column=1, padx=(10, 5), pady=(10, 0))
        ttk.Button(self.a5_frame, text="Chọn", command=self.browse_a5_template).grid(row=1, column=2, pady=(10, 0))
        ttk.Button(self.a5_frame, text="Mở", command=self.open_a5_template).grid(row=1, column=3, pady=(10, 0), padx=(5, 0))
        
        # Output folder
        output_frame = ttk.LabelFrame(main_frame, text="Thư mục xuất", padding="10")
        output_frame.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        
        ttk.Label(output_frame, text="Thư mục xuất:").grid(row=0, column=0, sticky=tk.W)
        self.output_var = tk.StringVar(value=self.config.config["output_folder"])
        output_entry = ttk.Entry(output_frame, textvariable=self.output_var, width=50)
        output_entry.grid(row=0, column=1, padx=(10, 5))
        ttk.Button(output_frame, text="Chọn", command=self.browse_output).grid(row=0, column=2)
        ttk.Button(output_frame, text="Mở", command=self.open_output_folder).grid(row=0, column=3, padx=(5, 0))
        
        # PDF Converter Tools
        config_frame = ttk.LabelFrame(main_frame, text="Công cụ chuyển đổi PDF", padding="10")
        config_frame.grid(row=5, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # Tool selection
        ttk.Label(config_frame, text="Công cụ chuyển đổi:").grid(row=0, column=0, sticky=tk.W)
        self.tool_var = tk.StringVar(value=self.config.config.get("preferred_tool", ""))
        self.tool_combo = ttk.Combobox(config_frame, textvariable=self.tool_var, width=25, state="readonly")
        self.tool_combo.grid(row=0, column=1, padx=(10, 5), sticky=tk.W)
        self.tool_combo.bind("<<ComboboxSelected>>", self.on_tool_change)
        ttk.Button(config_frame, text="Làm mới", command=self.refresh_tools).grid(row=0, column=2)
        
        # Tool info
        self.tool_info_var = tk.StringVar(value="Đang kiểm tra công cụ...")
        ttk.Label(config_frame, textvariable=self.tool_info_var, wraplength=600).grid(row=1, column=0, columnspan=3, sticky=tk.W, pady=(5, 0))
        
        # LibreOffice path (hỗ trợ cả bản Portable)
        ttk.Label(config_frame, text="Đường dẫn LibreOffice (có thể là Portable):").grid(row=2, column=0, sticky=tk.W, pady=(10, 0))
        self.libreoffice_var = tk.StringVar(value=self.config.config["libreoffice_path"])
        libreoffice_entry = ttk.Entry(config_frame, textvariable=self.libreoffice_var, width=50)
        libreoffice_entry.grid(row=2, column=1, padx=(10, 5), pady=(10, 0))
        ttk.Button(config_frame, text="Chọn", command=self.browse_libreoffice).grid(row=2, column=2, pady=(10, 0))
        
        # Progress bar
        progress_frame = ttk.LabelFrame(main_frame, text="Tiến trình xử lý", padding="10")
        progress_frame.grid(row=6, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(progress_frame, variable=self.progress_var, maximum=100, length=760, mode='determinate')
        self.progress_bar.grid(row=0, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 5))
        
        self.status_var = tk.StringVar(value="Sẵn sàng")
        ttk.Label(progress_frame, textvariable=self.status_var).grid(row=1, column=0, columnspan=3)
        
        # Log
        log_frame = ttk.LabelFrame(main_frame, text="Log", padding="10")
        log_frame.grid(row=7, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        
        self.log_text = tk.Text(log_frame, height=8, width=100)
        scrollbar = ttk.Scrollbar(log_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        # Buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=8, column=0, columnspan=3, pady=(10, 0))
        
        self.start_button = ttk.Button(button_frame, text="Bắt đầu xử lý", command=self.start_processing)
        self.start_button.grid(row=0, column=0, padx=(0, 10))
        
        ttk.Button(button_frame, text="Thông tin công cụ", command=self.show_tools_info).grid(row=0, column=1, padx=(0, 10))
        ttk.Button(button_frame, text="Lưu cấu hình", command=self.save_config).grid(row=0, column=2, padx=(0, 10))
        ttk.Button(button_frame, text="Thoát", command=self.root.quit).grid(row=0, column=3)
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.columnconfigure(2, weight=1)
        main_frame.rowconfigure(7, weight=1)
        self.a4_frame.columnconfigure(1, weight=1)
        self.a5_frame.columnconfigure(1, weight=1)
        output_frame.columnconfigure(1, weight=1)
        config_frame.columnconfigure(1, weight=1)
        progress_frame.columnconfigure(0, weight=1)
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        
        # Cập nhật hiển thị format
        self.on_format_change()
    
    def on_format_change(self):
        """Cập nhật hiển thị khi thay đổi format"""
        current_format = self.format_var.get()
        if current_format == "A4":
            self.a4_frame.configure(text="Cấu hình A4 (Đang sử dụng)")
            self.a5_frame.configure(text="Cấu hình A5")
        else:
            self.a4_frame.configure(text="Cấu hình A4")
            self.a5_frame.configure(text="Cấu hình A5 (Đang sử dụng)")
    
    def log(self, message):
        """Ghi log vào text widget"""
        timestamp = time.strftime("%H:%M:%S")
        self.log_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.log_text.see(tk.END)
        self.root.update_idletasks()
    
    def browse_a4_excel(self):
        filename = filedialog.askopenfilename(
            title="Chọn file Excel A4",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if filename:
            self.a4_excel_var.set(filename)
    
    def browse_a4_template(self):
        filename = filedialog.askopenfilename(
            title="Chọn template Word A4",
            filetypes=[("Word files", "*.docx *.doc"), ("All files", "*.*")]
        )
        if filename:
            self.a4_template_var.set(filename)
    
    def _open_path(self, path):
        try:
            if not path:
                messagebox.showwarning("Cảnh báo", "Đường dẫn trống.")
                return
            if not os.path.exists(path):
                messagebox.showerror("Lỗi", f"Không tìm thấy: {path}")
                return
            if platform.system() == 'Windows':
                os.startfile(path)
            elif platform.system() == 'Darwin':
                subprocess.Popen(['open', path])
            else:
                subprocess.Popen(['xdg-open', path])
        except Exception as e:
            messagebox.showerror("Lỗi", f"Không thể mở file: {e}")

    def open_a4_excel(self):
        self._open_path(self.a4_excel_var.get())

    def open_a4_template(self):
        self._open_path(self.a4_template_var.get())
    
    def browse_a5_excel(self):
        filename = filedialog.askopenfilename(
            title="Chọn file Excel A5",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if filename:
            self.a5_excel_var.set(filename)
    
    def browse_a5_template(self):
        filename = filedialog.askopenfilename(
            title="Chọn template Word A5",
            filetypes=[("Word files", "*.docx *.doc"), ("All files", "*.*")]
        )
        if filename:
            self.a5_template_var.set(filename)
    
    def open_a5_excel(self):
        self._open_path(self.a5_excel_var.get())

    def open_a5_template(self):
        self._open_path(self.a5_template_var.get())
    
    def browse_output(self):
        folder = filedialog.askdirectory(title="Chọn thư mục xuất")
        if folder:
            self.output_var.set(folder)
    
    def open_output_folder(self):
        """Mở thư mục xuất trong Explorer (tạo nếu chưa tồn tại)"""
        try:
            folder = self.output_var.get()
            if not folder:
                messagebox.showwarning("Cảnh báo", "Đường dẫn thư mục xuất đang trống.")
                return
            if not os.path.exists(folder):
                os.makedirs(folder, exist_ok=True)
            self._open_path(folder)
        except Exception as e:
            messagebox.showerror("Lỗi", f"Không thể mở thư mục xuất: {e}")
    
    def refresh_tools(self):
        """Làm mới danh sách công cụ chuyển đổi"""
        try:
            # Cập nhật converter với config mới
            config = {
                "libreoffice_path": self.libreoffice_var.get()
            }
            self.converter = PDFConverter(config)
            
            # Lấy thông tin công cụ
            all_tools = self.converter.get_all_tools_info()
            available_tools = []
            
            for tool_name, tool_info in all_tools.items():
                if tool_info['available']:
                    available_tools.append(tool_name)
            
            # Cập nhật combobox
            self.tool_combo['values'] = available_tools
            # Nếu user có lưu preferred_tool và nó khả dụng thì dùng nó
            preferred_tool = self.config.config.get("preferred_tool", "")
            if preferred_tool and preferred_tool in available_tools:
                self.converter.current_tool = preferred_tool
                self.tool_combo.set(preferred_tool)
            elif available_tools:
                # Không có preferred hợp lệ: dùng auto chọn tốt nhất
                self.tool_combo.set(self.converter.current_tool)
            
            # Hiển thị thông tin công cụ hiện tại
            current_tool_info = self.converter.get_tool_info()
            if current_tool_info:
                if current_tool_info['name'] == 'python-docx2pdf':
                    self.tool_info_var.set(f"🎯 Đang sử dụng: {current_tool_info['name']} - {current_tool_info['description']}")
                else:
                    self.tool_info_var.set(f"⚠️ Đang sử dụng: {current_tool_info['name']} - {current_tool_info['description']}")
            else:
                self.tool_info_var.set("❌ Không tìm thấy công cụ chuyển đổi nào!")
                
        except Exception as e:
            self.tool_info_var.set(f"Lỗi khi kiểm tra công cụ: {e}")

    def on_tool_change(self, event=None):
        """Xử lý khi người dùng chọn công cụ chuyển đổi trong combobox"""
        try:
            selected_tool = self.tool_var.get()
            if selected_tool in self.converter.available_tools and self.converter.available_tools[selected_tool]['available']:
                self.converter.current_tool = selected_tool
                # Lưu lại lựa chọn này vào config đang giữ trong bộ nhớ
                self.config.update_config("preferred_tool", selected_tool)
                current_tool_info = self.converter.get_tool_info()
                if current_tool_info['name'] == 'python-docx2pdf':
                    self.tool_info_var.set(f"🎯 Đang sử dụng: {current_tool_info['name']} - {current_tool_info['description']}")
                else:
                    self.tool_info_var.set(f"⚠️ Đang sử dụng: {current_tool_info['name']} - {current_tool_info['description']}")
                self.log(f"Đã chọn công cụ chuyển đổi: {selected_tool}")
            else:
                messagebox.showwarning("Cảnh báo", "Công cụ được chọn không khả dụng.")
        except Exception as e:
            messagebox.showerror("Lỗi", f"Không thể thay đổi công cụ: {e}")
    
    def browse_libreoffice(self):
        filename = filedialog.askopenfilename(
            title="Chọn file thực thi LibreOffice (soffice.exe hoặc soffice.com)",
            filetypes=[("Executable files", "*.exe *.com"), ("All files", "*.*")]
        )
        if filename:
            self.libreoffice_var.set(filename)
            self.refresh_tools()  # Làm mới sau khi thay đổi đường dẫn
    
    def save_config(self):
        """Lưu cấu hình hiện tại"""
        config = {
            "a4_excel_file": self.a4_excel_var.get(),
            "a4_word_template": self.a4_template_var.get(),
            "a5_excel_file": self.a5_excel_var.get(),
            "a5_word_template": self.a5_template_var.get(),
            "output_folder": self.output_var.get(),
            "libreoffice_path": self.libreoffice_var.get(),
            "batch_size": self.config.config["batch_size"],
            "default_format": self.format_var.get(),
            "preferred_tool": self.tool_var.get() or self.converter.current_tool
        }
        self.config.save_config(config)
        messagebox.showinfo("Thông báo", "Đã lưu cấu hình thành công!")
    
    def start_processing(self):
        """Bắt đầu xử lý trong thread riêng"""
        self.start_button.config(state='disabled')
        self.progress_var.set(0)
        self.status_var.set("Đang xử lý...")
        self.log_text.delete(1.0, tk.END)
        
        # Chạy xử lý trong thread riêng để không block GUI
        thread = threading.Thread(target=self.process_data)
        thread.daemon = True
        thread.start()
    
    def process_data(self):
        """Xử lý dữ liệu chính"""
        try:
            # Cập nhật config
            config = {
                "a4_excel_file": self.a4_excel_var.get(),
                "a4_word_template": self.a4_template_var.get(),
                "a5_excel_file": self.a5_excel_var.get(),
                "a5_word_template": self.a5_template_var.get(),
                "output_folder": self.output_var.get(),
                "libreoffice_path": self.libreoffice_var.get(),
                "batch_size": self.config.config["batch_size"],
                "default_format": self.format_var.get(),
                "preferred_tool": self.tool_var.get() # Lấy công cụ được chọn từ config
            }
            
            # Cập nhật converter với config mới
            self.converter.config = config
            
            current_format = self.format_var.get()
            self.log(f"Bắt đầu xử lý định dạng {current_format}...")
            
            if current_format == "A4":
                self.process_a4_format(config)
            else:
                self.process_a5_format(config)
            
        except Exception as e:
            self.log(f"Lỗi: {str(e)}")
            self.status_var.set("Lỗi")
            messagebox.showerror("Lỗi", f"Có lỗi xảy ra: {str(e)}")
        
        finally:
            self.start_button.config(state='normal')
    
    def process_a4_format(self, config):
        """Xử lý định dạng A4 - một sản phẩm/trang"""
        try:
            self.log(f"Đang xử lý định dạng A4...")
            self.log(f"File Excel: {config['a4_excel_file']}")
            self.log(f"Template: {config['a4_word_template']}")
            self.log(f"Thư mục xuất: {config['output_folder']}")
            
            # Kiểm tra file tồn tại
            if not os.path.exists(config['a4_excel_file']):
                raise FileNotFoundError(f"Không tìm thấy file Excel A4: {config['a4_excel_file']}")
            if not os.path.exists(config['a4_word_template']):
                raise FileNotFoundError(f"Không tìm thấy template A4: {config['a4_word_template']}")
            
            # Tạo thư mục xuất
            os.makedirs(config['output_folder'], exist_ok=True)
            
            # Đọc Excel
            self.log("Đang đọc file Excel A4...")
            df = pd.read_excel(config['a4_excel_file'])
            self.log(f"Đã đọc {len(df)} dòng dữ liệu")
            
            # Tính toán tổng số bước
            total_steps = len(df) + 2  # Word + PDF + Merge
            self.progress_tracker = ProgressTracker(total_steps)
            
            # Xử lý tạo file Word
            self.log("Đang tạo file Word A4...")
            word_files = self.process_a4_word_files(df, config)
            
            if not word_files:
                raise Exception("Không tạo được file Word nào")
            
            self.log(f"Đã tạo {len(word_files)} file Word")
            
            # Chuyển đổi PDF
            self.log("Đang chuyển đổi sang PDF...")
            pdf_files = self.convert_to_pdf_sequential(word_files, config)
            
            if not pdf_files:
                raise Exception("Không chuyển đổi được file PDF nào")
            
            self.log(f"Đã chuyển đổi {len(pdf_files)} file PDF")
            
            # Gộp PDF
            self.log("Đang gộp file PDF...")
            merged_pdf = self.merge_pdfs(pdf_files, config['output_folder'], "A4-Auto-Tong.pdf")
            
            # Dọn dẹp
            self.cleanup_files(word_files + pdf_files)
            
            self.log("Hoàn thành! File PDF đã được gộp thành công!")
            self.status_var.set("Hoàn thành")
            self.progress_var.set(100)
            
            messagebox.showinfo("Thành công", f"Đã xử lý xong! File PDF: {merged_pdf}")
            
        except Exception as e:
            self.log(f"Lỗi xử lý A4: {str(e)}")
            raise
    
    def process_a5_format(self, config):
        """Xử lý định dạng A5 - hai sản phẩm/trang"""
        try:
            self.log(f"Đang xử lý định dạng A5...")
            self.log(f"File Excel: {config['a5_excel_file']}")
            self.log(f"Template: {config['a5_word_template']}")
            self.log(f"Thư mục xuất: {config['output_folder']}")
            
            # Kiểm tra file tồn tại
            if not os.path.exists(config['a5_excel_file']):
                raise FileNotFoundError(f"Không tìm thấy file Excel A5: {config['a5_excel_file']}")
            if not os.path.exists(config['a5_word_template']):
                raise FileNotFoundError(f"Không tìm thấy template A5: {config['a5_word_template']}")
            
            # Tạo thư mục xuất
            os.makedirs(config['output_folder'], exist_ok=True)
            
            # Đọc Excel
            self.log("Đang đọc file Excel A5...")
            df = pd.read_excel(config['a5_excel_file'])
            self.log(f"Đã đọc {len(df)} dòng dữ liệu")
            
            # Tính toán tổng số bước
            total_steps = len(df) // 2 + 2  # Word + PDF + Merge
            self.progress_tracker = ProgressTracker(total_steps)
            
            # Xử lý tuần tự tạo file Word
            self.log("Đang tạo file Word A5...")
            word_files = self.process_a5_word_files_sequential(df, config)
            
            if not word_files:
                raise Exception("Không tạo được file Word nào")
            
            self.log(f"Đã tạo {len(word_files)} file Word")
            
            # Chuyển đổi PDF tuần tự
            self.log("Đang chuyển đổi sang PDF...")
            pdf_files = self.convert_to_pdf_sequential(word_files, config)
            
            if not pdf_files:
                raise Exception("Không chuyển đổi được file PDF nào")
            
            self.log(f"Đã chuyển đổi {len(pdf_files)} file PDF")
            
            # Gộp PDF
            self.log("Đang gộp file PDF...")
            merged_pdf = self.merge_pdfs(pdf_files, config['output_folder'], "A5-Auto-Tong.pdf")
            
            # Dọn dẹp
            self.cleanup_files(word_files + pdf_files)
            
            self.log("Hoàn thành! File PDF đã được gộp thành công!")
            self.status_var.set("Hoàn thành")
            self.progress_var.set(100)
            
            messagebox.showinfo("Thành công", f"Đã xử lý xong! File PDF: {merged_pdf}")
            
        except Exception as e:
            self.log(f"Lỗi xử lý A5: {str(e)}")
            raise
    
    def process_a4_word_files(self, df, config):
        """Xử lý file Word A4 - một sản phẩm/trang"""
        word_files = []
        errors = []
        
        # Xử lý từng dòng tuần tự
        for i, row in df.iterrows():
            try:
                # Load template
                doc = DocxTemplate(config['a4_word_template'])
                
                # Tạo context
                context = self.create_a4_context(row)
                
                # Render template
                doc.render(context)
                
                # Lưu file Word
                temp_word_file = os.path.join(config['output_folder'], f"temp_output_{i + 1}.docx")
                doc.save(temp_word_file)
                
                if os.path.exists(temp_word_file):
                    word_files.append(temp_word_file)
                    self.log(f"Đã tạo: {os.path.basename(temp_word_file)}")
                else:
                    error_msg = f"Không thể tạo file Word: {temp_word_file}"
                    errors.append(error_msg)
                    self.log(f"Lỗi: {error_msg}")
                
                # Cập nhật progress
                self.progress_tracker.update()
                progress = self.progress_tracker.get_progress()
                self.progress_var.set(progress)
                self.status_var.set(f"Đang tạo Word A4... {len(word_files)}/{len(df)}")
                
            except Exception as e:
                error_msg = f"Lỗi khi xử lý dòng {i+1}: {e}"
                errors.append(error_msg)
                self.log(f"Lỗi: {error_msg}")
        
        if errors:
            self.log(f"Có {len(errors)} lỗi xảy ra")
        
        return word_files
    
    def process_a5_word_files_sequential(self, df, config):
        """Xử lý file Word A5 tuần tự - hai sản phẩm/trang"""
        word_files = []
        errors = []
        
        # Xử lý từng cặp 2 dòng tuần tự
        for i in range(0, len(df), 2):
            try:
                # Load template
                doc = DocxTemplate(config['a5_word_template'])
                
                # Lấy dữ liệu
                row_1 = df.iloc[i]
                row_2 = df.iloc[i + 1] if i + 1 < len(df) else None
                
                # Tạo context
                context = self.create_a5_context(row_1, row_2)
                
                # Render template
                doc.render(context)
                
                # Lưu file Word
                temp_word_file = os.path.join(config['output_folder'], f"temp_output_{i//2 + 1}.docx")
                doc.save(temp_word_file)
                
                if os.path.exists(temp_word_file):
                    word_files.append(temp_word_file)
                    self.log(f"Đã tạo: {os.path.basename(temp_word_file)}")
                else:
                    error_msg = f"Không thể tạo file Word: {temp_word_file}"
                    errors.append(error_msg)
                    self.log(f"Lỗi: {error_msg}")
                
                # Cập nhật progress
                self.progress_tracker.update()
                progress = self.progress_tracker.get_progress()
                self.progress_var.set(progress)
                self.status_var.set(f"Đang tạo Word A5... {len(word_files)}/{len(df)//2}")
                
            except Exception as e:
                error_msg = f"Lỗi khi xử lý cặp dòng {i+1} và {i+2}: {e}"
                errors.append(error_msg)
                self.log(f"Lỗi: {error_msg}")
        
        if errors:
            self.log(f"Có {len(errors)} lỗi xảy ra")
        
        return word_files
    
    def create_a4_context(self, row):
        """Tạo context cho template A4"""
        return {
            "NganhHang": self.limit_string(row['NganhHang'], max_length=29),
            "Hang": row['Hang'] if pd.notnull(row['Hang']) else "",
            "SAP": row['SAP'] if pd.notnull(row['SAP']) else "",
            "Model": row['Model'] if pd.notnull(row['Model']) else "",
            "GiaNiemYet": self.format_currency(row['GiaNiemYet']),
            "GiaKM": self.format_currency(row['GiaKM']),
            "G": self.format_percentage(row['G']),
            "Qua": row['Qua'] if pd.notnull(row['Qua']) else "",
            "ThoiGian": row['ThoiGian']
        }
    
    def create_a5_context(self, row_1, row_2):
        """Tạo context cho template A5"""
        context = {}
        
        # Dòng 1
        context.update({
            "NganhHang": self.limit_string(row_1['NganhHang'], max_length=31),
            "Hang": row_1['Hang'] if pd.notnull(row_1['Hang']) else "",
            "SAP": row_1['SAP'] if pd.notnull(row_1['SAP']) else "",
            "Model": row_1['Model'] if pd.notnull(row_1['Model']) else "",
            "GiaNiemYet": self.format_currency(row_1['GiaNiemYet']),
            "GiaKM": self.format_currency(row_1['GiaKm']),
            "G": self.format_percentage(row_1['G']),
            "Qua": row_1['Qua'] if pd.notnull(row_1['Qua']) else "",
            "ThoiGian": row_1['ThoiGian']
        })
        
        # Dòng 2 (nếu có)
        if row_2 is not None:
            context.update({
                "NganhHang1": self.limit_string(row_2['NganhHang'], max_length=31),
                "Hang1": row_2['Hang'] if pd.notnull(row_2['Hang']) else "",
                "SAP1": row_2['SAP'] if pd.notnull(row_2['SAP']) else "",
                "Model1": row_2['Model'] if pd.notnull(row_2['Model']) else "",
                "GiaNiemYet1": self.format_currency(row_2['GiaNiemYet']),
                "GiaKM1": self.format_currency(row_2['GiaKm']),
                "G1": self.format_percentage(row_2['G']),
                "Qua1": row_2['Qua'] if pd.notnull(row_2['Qua']) else "",
                "ThoiGian1": row_2['ThoiGian']
            })
        else:
            context.update({
                "NganhHang1": "", "Hang1": "", "SAP1": "", "Model1": "",
                "GiaNiemYet1": "", "GiaKM1": "", "G1": "", "Qua1": "", "ThoiGian1": ""
            })
        
        return context
    
    def format_currency(self, value):
        """Format currency"""
        if pd.isna(value):
            return ""
        try:
            return "{:,.0f}đ".format(float(str(value).replace(',', '')))
        except (ValueError, TypeError):
            return str(value)
    
    def format_percentage(self, value):
        """Format percentage"""
        if pd.isna(value):
            return ""
        try:
            return "{:.0f}%".format(float(str(value).replace('%', '')))
        except (ValueError, TypeError):
            return str(value)
    
    def limit_string(self, text, max_length=33):
        """Limit string"""
        if pd.isna(text):
            return ""
        text = str(text)
        return text[:max_length-3] + "..." if len(text) > max_length else text
    
    def convert_to_pdf_sequential(self, word_files, config):
        """Chuyển đổi PDF tuần tự"""
        pdf_files = []
        
        for word_file in word_files:
            try:
                pdf_file = self.converter.convert_to_pdf(word_file)
                if pdf_file:
                    pdf_files.append(pdf_file)
                    self.log(f"Đã chuyển đổi: {os.path.basename(pdf_file)}")
                else:
                    self.log(f"Lỗi chuyển đổi: {os.path.basename(word_file)}")
                
                # Cập nhật progress
                self.progress_tracker.update()
                progress = self.progress_tracker.get_progress()
                self.progress_var.set(progress)
                self.status_var.set(f"Đang chuyển đổi PDF... {len(pdf_files)}/{len(word_files)}")
                
            except Exception as e:
                self.log(f"Lỗi chuyển đổi {word_file}: {e}")
        
        return pdf_files
    
    def merge_pdfs(self, pdf_files, output_folder, output_name):
        """Gộp các file PDF"""
        try:
            merger = PdfWriter()
            for pdf in pdf_files:
                merger.append(pdf)
            
            merged_pdf = os.path.join(output_folder, output_name)
            merger.write(merged_pdf)
            merger.close()
            
            self.log(f"Đã gộp {len(pdf_files)} file PDF thành: {merged_pdf}")
            return merged_pdf
            
        except Exception as e:
            self.log(f"Lỗi khi gộp PDF: {e}")
            raise
    
    def cleanup_files(self, files):
        """Dọn dẹp file tạm"""
        for file_path in files:
            try:
                if os.path.exists(file_path):
                    os.remove(file_path)
                    self.log(f"Đã xóa: {os.path.basename(file_path)}")
            except Exception as e:
                self.log(f"Lỗi khi xóa {file_path}: {e}")
    
    def run(self):
        """Chạy GUI"""
        self.root.mainloop()

    # Đã xóa preview_template theo yêu cầu
    
    def show_tools_info(self):
        """Hiển thị thông tin chi tiết về các công cụ chuyển đổi"""
        try:
            tools_window = tk.Toplevel(self.root)
            tools_window.title("Thông tin công cụ chuyển đổi PDF")
            tools_window.geometry("600x500")
            tools_window.resizable(True, True)
            
            # Frame chính
            main_frame = ttk.Frame(tools_window, padding="10")
            main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
            
            # Tiêu đề
            title_label = ttk.Label(main_frame, text="Công cụ chuyển đổi PDF", font=("Arial", 14, "bold"))
            title_label.grid(row=0, column=0, columnspan=2, pady=(0, 10))
            
            # Text widget để hiển thị thông tin
            text_widget = tk.Text(main_frame, wrap=tk.WORD, height=20)
            scrollbar = ttk.Scrollbar(main_frame, orient=tk.VERTICAL, command=text_widget.yview)
            text_widget.configure(yscrollcommand=scrollbar.set)
            
            text_widget.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
            scrollbar.grid(row=1, column=1, sticky=(tk.N, tk.S))
            
            # Lấy thông tin công cụ
            all_tools = self.converter.get_all_tools_info()
            current_tool = self.converter.current_tool
            
            # Hiển thị thông tin
            info_text = "=== THÔNG TIN CÔNG CỤ CHUYỂN ĐỔI PDF ===\n\n"
            
            for tool_name, tool_info in all_tools.items():
                status = "✓ CÓ SẴN" if tool_info['available'] else "✗ KHÔNG CÓ"
                current = " (ĐANG SỬ DỤNG)" if tool_name == current_tool else ""
                
                info_text += f"🔧 {tool_info['name']} {status}{current}\n"
                info_text += f"   Mô tả: {tool_info['description']}\n"
                info_text += f"   Độ ưu tiên: {tool_info['priority']}\n\n"
            
            info_text += "\n=== HƯỚNG DẪN CÀI ĐẶT ===\n\n"
            info_text += "1. python-docx2pdf (🚀 Khuyến nghị):\n"
            info_text += "   pip install docx2pdf\n\n"
            info_text += "2. LibreOffice / LibreOffice Portable (Fallback):\n"
            info_text += "   - Dùng bản Portable: Chỉ cần giải nén cạnh ứng dụng (thư mục 'LibreOfficePortable').\n"
            info_text += "   - Hoặc đặt biến môi trường LIBREOFFICE_PATH trỏ tới 'soffice.exe' hoặc 'soffice.com'.\n"
            info_text += "   - Hoặc chọn đường dẫn thủ công bằng nút 'Chọn'.\n\n"
            
            info_text += "=== LƯU Ý ===\n"
            info_text += "- Ưu tiên python-docx2pdf. Nếu lỗi sẽ fallback sang LibreOffice.\n"
            
            text_widget.insert(tk.END, info_text)
            text_widget.config(state=tk.DISABLED)
            
            # Button đóng
            ttk.Button(main_frame, text="Đóng", command=tools_window.destroy).grid(row=2, column=0, columnspan=2, pady=(10, 0))
            
            # Configure grid weights
            tools_window.columnconfigure(0, weight=1)
            tools_window.rowconfigure(0, weight=1)
            main_frame.columnconfigure(0, weight=1)
            main_frame.rowconfigure(1, weight=1)
            
        except Exception as e:
            messagebox.showerror("Lỗi", f"Không thể hiển thị thông tin công cụ: {e}")

def main():
    """Hàm chính"""
    try:
        app = AutoPriceGUI()
        app.run()
    except Exception as e:
        print(f"Lỗi khởi động GUI: {e}")
        # Fallback về command line nếu GUI lỗi
        run_command_line()

def run_command_line():
    """Chạy phiên bản command line cũ"""
    print("Khởi động phiên bản command line...")
    print("Vui lòng sử dụng A4-AUTO.py hoặc A5-AUTO.py cho command line")

if __name__ == "__main__":
    main()
