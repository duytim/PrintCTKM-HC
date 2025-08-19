# 🚀 AUTO-PRICE - Tự động hóa bảng giá A4 & A5

![Python](https://img.shields.io/badge/Python-3.8+-3776AB?logo=python&logoColor=white)
![License](https://img.shields.io/badge/License-MIT-blue)
![Status](https://img.shields.io/badge/Status-Active-green)

**AUTO-PRICE** là một ứng dụng Python mạnh mẽ giúp tự động hóa việc tạo bảng giá định dạng **A4** (một sản phẩm/trang) và **A5** (hai sản phẩm/trang) từ dữ liệu Excel và template Word. Với giao diện GUI thân thiện, ứng dụng hỗ trợ tạo file Word, chuyển đổi sang PDF, và gộp thành một file PDF duy nhất, tiết kiệm thời gian và công sức.

---

## 📋 Mục lục
- [Giới thiệu](#-giới-thiệu)
- [Tính năng](#-tính-năng)
- [Yêu cầu hệ thống](#-yêu-cầu-hệ-thống)
- [Cài đặt](#-cài-đặt)
- [Cách sử dụng](#-cách-sử-dụng)
- [Cấu hình](#-cấu-hình)
- [Build file thực thi (.exe)](#-build-file-thực-thi-exe)
- [Cấu trúc thư mục](#-cấu-trúc-thư-mục)
- [My Stack](#-my-stack)
- [Inspiration](#-inspiration)
- [Future Ideas](#-future-ideas)
- [Lưu ý](#-lưu-ý)
- [Hỗ trợ](#-hỗ-trợ)
- [Giấy phép](#-giấy-phép)

---

## 🌟 Giới thiệu
**AUTO-PRICE** là giải pháp tự động hóa tối ưu cho việc tạo bảng giá chuyên nghiệp. Ứng dụng sử dụng dữ liệu từ file Excel, kết hợp với template Word để tạo ra các file Word, sau đó chuyển đổi sang PDF và gộp lại. Giao diện GUI được xây dựng bằng Tkinter, giúp người dùng dễ dàng cấu hình và theo dõi tiến trình.

---

## ✨ Tính năng
- **Hỗ trợ định dạng linh hoạt**:
  - **A4**: Một sản phẩm mỗi trang, lý tưởng cho hiển thị chi tiết.
  - **A5**: Hai sản phẩm mỗi trang, tối ưu cho in ấn tiết kiệm.
- **Giao diện người dùng trực quan**:
  - Chọn file Excel, template Word, và thư mục xuất qua GUI.
  - Hiển thị tiến trình xử lý với thanh tiến độ và log chi tiết.
  - Lưu và tải cấu hình từ file `config.json`.
- **Chuyển đổi PDF thông minh**:
  - Ưu tiên `python-docx2pdf` (nhanh, nhẹ, hiệu quả).
  - Fallback sang LibreOffice (hỗ trợ bản cài đặt và Portable).
- **Xử lý dữ liệu mạnh mẽ**:
  - Đọc dữ liệu từ Excel (`pandas`).
  - Render template Word (`docxtpl`).
  - Chuyển đổi và gộp PDF (`pypdf`).
  - Tự động dọn dẹp file tạm sau khi hoàn thành.

---

## 🛠 Yêu cầu hệ thống
- **Python**: Phiên bản 3.8 hoặc cao hơn.
- **Hệ điều hành**: Windows, macOS, hoặc Linux.
- **Phần mềm bổ trợ** (tùy chọn):
  - **Microsoft Word** hoặc **LibreOffice** (bản cài đặt hoặc Portable) để chuyển đổi PDF.
- **Thư viện Python**:
  - Cài đặt từ `requirements.txt`:
    ```bash
    pip install -r requirements.txt
    ```

---

## ⚙️ Cài đặt
1. **Clone repository**:
   ```bash
   git clone <repository_url>
   cd <repository_directory>
   ```

2. **Cài đặt thư viện**:
   Tạo môi trường ảo và cài đặt các thư viện:
   ```bash
   python -m venv venv
   source venv/bin/activate  # Linux/macOS
   venv\Scripts\activate     # Windows
   pip install -r requirements.txt
   ```

3. **Chuẩn bị file cần thiết**:
   - **File Excel** (`A4-Auto.xlsx`, `A5-AUTO.xlsx`): Chứa dữ liệu sản phẩm với các cột như `NganhHang`, `Hang`, `SAP`, `Model`, `GiaNiemYet`, `GiaKM`, `G`, `Qua`, `ThoiGian`.
   - **Template Word** (`A4-Auto.docx`, `A5-AUTO.docx`): Template chứa các placeholder (ví dụ: `{{NganhHang}}`, `{{Hang}}`,...).
   - **Thư mục xuất** (`In_PDF`): Nơi lưu file PDF gộp và file tạm.

4. **(Tùy chọn) Cấu hình LibreOffice**:
   - Đặt thư mục `LibreOfficePortable` cạnh `AUTO-PRICE.py` hoặc chỉ định đường dẫn `soffice.exe`/`soffice.com` trong GUI.

---

## 📖 Cách sử dụng
1. **Chạy ứng dụng**:
   ```bash
   python AUTO-PRICE.py
   ```

2. **Giao diện GUI**:
   - Chọn định dạng (**A4** hoặc **A5**).
   - Cấu hình file Excel, template Word, thư mục xuất, và công cụ chuyển đổi PDF.
   - Nhấn **Lưu cấu hình** để lưu thiết lập.
   - Nhấn **Bắt đầu xử lý** để chạy quy trình.

3. **Quy trình xử lý**:
   - Đọc dữ liệu từ Excel.
   - Tạo file Word từ template.
   - Chuyển đổi Word sang PDF.
   - Gộp PDF thành file duy nhất (`A4-Auto-Tong.pdf` hoặc `A5-Auto-Tong.pdf`).
   - Dọn dẹp file tạm.

4. **Kiểm tra kết quả**:
   - File PDF gộp nằm trong thư mục `In_PDF`.
   - Theo dõi log trong GUI để kiểm tra chi tiết.

---

## 🔧 Cấu hình
File `config.json` lưu trữ các thiết lập mặc định:

```json
{
  "a4_excel_file": "A4-Auto.xlsx",
  "a4_word_template": "A4-Auto.docx",
  "a5_excel_file": "A5-AUTO.xlsx",
  "a5_word_template": "A5-AUTO.docx",
  "output_folder": "In_PDF",
  "libreoffice_path": "C:\\Program Files\\LibreOffice\\program\\soffice.exe",
  "batch_size": 10,
  "default_format": "A5",
  "preferred_tool": ""
}
```

- **preferred_tool**: Để trống (`""`) để tự động chọn công cụ tốt nhất, hoặc chỉ định `"docx2pdf"` hoặc `"libreoffice"`.

---

## 📦 Build file thực thi (.exe)
1. **Cài đặt PyInstaller**:
   ```bash
   pip install pyinstaller
   ```

2. **Build file .exe**:
   ```bash
   pyinstaller --onefile --noconsole --name AUTO-PRICE AUTO-PRICE.py
   ```

3. **Build với tài nguyên**:
   Nếu cần thêm file Excel/Word:
   ```bash
   pyinstaller --onefile --noconsole --name AUTO-PRICE --add-data "A4-Auto.xlsx;." --add-data "A4-Auto.docx;." --add-data "A5-AUTO.xlsx;." --add-data "A5-AUTO.docx;." AUTO-PRICE.py
   ```

4. **Build tự động trên GitHub**:
   - Xem file `.github/workflows/build.yml` để build `.exe` trên GitHub Actions.
   - File thực thi được lưu trong **Artifacts** của workflow.

---

## 📂 Cấu trúc thư mục
```
<repository_directory>/
├── AUTO-PRICE.py           # File chính
├── requirements.txt        # Danh sách thư viện
├── A4-Auto.xlsx           # File Excel cho A4
├── A4-Auto.docx           # Template Word cho A4
├── A5-AUTO.xlsx           # File Excel cho A5
├── A5-AUTO.docx           # Template Word cho A5
├── In_PDF/                # Thư mục xuất (tạo tự động)
├── config.json            # File cấu hình (tạo tự động)
├── LibreOfficePortable/   # (Tùy chọn) LibreOffice Portable
└── .github/workflows/     # Thư mục chứa workflow
    └── build.yml
```

---

## 🛠 My Stack
Công nghệ và công cụ được sử dụng trong dự án:
- **Python 3.8+**: Ngôn ngữ chính, mạnh mẽ và linh hoạt.
- **Tkinter**: Xây dựng giao diện GUI đơn giản, tích hợp sẵn trong Python.
- **Pandas**: Xử lý dữ liệu Excel hiệu quả.
- **Docxtpl**: Render template Word với dữ liệu động.
- **Pypdf**: Gộp và quản lý file PDF.
- **Docx2pdf**: Chuyển đổi Word sang PDF nhanh chóng.
- **LibreOffice**: Công cụ chuyển đổi PDF dự phòng (hỗ trợ Portable).
- **PyInstaller**: Build file thực thi `.exe` cho Windows.
- **GitHub Actions**: Tự động hóa build và triển khai.

---

## 💡 Inspiration
Dự án được truyền cảm hứng từ nhu cầu thực tế trong việc tự động hóa quy trình tạo bảng giá bán lẻ. Mục tiêu là:
- Giảm thiểu công việc thủ công lặp đi lặp lại.
- Tăng tính chính xác và đồng nhất trong định dạng bảng giá.
- Tạo ra công cụ dễ sử dụng, phù hợp với người dùng không chuyên về lập trình.
- Tích hợp nhiều công cụ chuyển đổi PDF để đảm bảo tính tương thích và linh hoạt.

---

## 🚀 Future Ideas
Những ý tưởng để cải thiện và mở rộng AUTO-PRICE trong tương lai:
- **Hỗ trợ thêm định dạng**: Thêm các kích thước trang khác như A3, Letter, hoặc tùy chỉnh.
- **Tích hợp xem trước template**: Cho phép xem trước template Word ngay trong GUI.
- **Hỗ trợ đa ngôn ngữ**: Thêm giao diện và tài liệu tiếng Anh, hỗ trợ người dùng quốc tế.
- **Tối ưu hiệu suất**: Thêm xử lý song song (multithreading) để tăng tốc độ với dữ liệu lớn.
- **Tích hợp đám mây**: Hỗ trợ tải file Excel/template từ Google Drive hoặc Dropbox.
- **Kiểm tra dữ liệu**: Thêm chức năng kiểm tra định dạng Excel trước khi xử lý.
- **Giao diện hiện đại hơn**: Nâng cấp GUI với thư viện như `customtkinter` hoặc `PyQt`.

---

## ⚠️ Lưu ý
- **Công cụ chuyển đổi PDF**:
  - **python-docx2pdf** (khuyến nghị): Cần Microsoft Word hoặc LibreOffice.
  - **LibreOffice**: Đặt thư mục Portable cạnh file `.py` hoặc chỉ định đường dẫn trong GUI.
- **File Excel**:
  - Đảm bảo các cột khớp với placeholder trong template Word.
  - Kiểm tra dữ liệu để tránh lỗi định dạng.
- **Hiệu suất**:
  - Với dữ liệu lớn, điều chỉnh `batch_size` trong `config.json`.
- **Khắc phục lỗi**:
  - Theo dõi log trong GUI để xác định vấn đề.
  - Đảm bảo file Excel, Word template, và thư mục xuất tồn tại.

---

## 📧 Hỗ trợ
Nếu gặp vấn đề hoặc cần hỗ trợ:
- Kiểm tra log trong GUI để xem chi tiết lỗi.
- Tạo issue trên repository GitHub.
- Liên hệ qua email hoặc kênh hỗ trợ của dự án.

---

## 📜 Giấy phép
Dự án được phát triển dưới [Giấy phép MIT](LICENSE). Vui lòng đọc file `LICENSE` để biết thêm chi tiết.