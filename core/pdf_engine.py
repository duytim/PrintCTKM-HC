from __future__ import annotations
import shutil
import subprocess
import zipfile
import concurrent.futures
from datetime import datetime
from pathlib import Path
from typing import List, Dict, Optional, Callable, Tuple, Any

import pandas as pd
from docxtpl import DocxTemplate
from pypdf import PdfWriter
from utils.logger import logger

def render_single_docx(args: Tuple[Path, Path, List[Dict[str, Any]], str, int, List[str]]) -> Optional[Path]:
    tpl_path, out_path, batch_records, fmt, step, keys = args
    try:
        doc = DocxTemplate(str(tpl_path))
        ctx = {f"{k}{str(idx) if idx > 0 and fmt != 'A4' else ''}": "" for idx in range(step) for k in keys}
        
        for idx, row in enumerate(batch_records):
            sf = str(idx) if idx > 0 and fmt != 'A4' else ""
            def money(v): return "{:,.0f}đ".format(float(v)) if v else ""
            def pct(v): return "{:.0f}%".format(float(v)) if v else ""

            ctx[f"Ten{sf}"] = str(row.get('Tên', ''))[:35]
            ctx[f"Ma{sf}"] = str(row.get('Mã', ''))
            ctx[f"Model{sf}"] = str(row.get('Model', ''))
            ctx[f"GiaNY{sf}"] = money(row.get('Giá NY'))
            ctx[f"GiaKM{sf}"] = money(row.get('Giá KM'))
            ctx[f"G{sf}"] = pct(row.get('% Giảm giá'))
            ctx[f"Qua{sf}"] = str(row.get('Tên SP1', ''))
            ctx[f"QuaA{sf}"] = str(row.get('Tên SP2', ''))
            ctx[f"QuaB{sf}"] = str(row.get('Tên SP3', ''))
            ctx[f"ThoiGian{sf}"] = str(row.get('Thời gian', datetime.now().strftime("%d/%m")))

        doc.render(ctx)
        doc.save(out_path)
        return out_path
    except Exception as e:
        logger.warning(f"Lỗi thread render {out_path.name}: {e}")
        return None

class PdfEngine:
    def __init__(self, config: Dict):
        self.config = config
        self.workspace = Path.cwd()
        self.out_dir = self.workspace / "In_PDF"
        self.out_dir.mkdir(exist_ok=True)

    def find_libreoffice(self) -> Optional[Path]:
        cfg_path = self.config.get('libre_path', '')
        if cfg_path and Path(cfg_path).exists(): return Path(cfg_path)
        possibles = [
            self.workspace / "LibreOfficePortable" / "App" / "libreoffice" / "program" / "soffice.exe",
            Path(r"C:\Program Files\LibreOffice\program\soffice.exe"),
            Path(r"C:\Program Files (x86)\LibreOffice\program\soffice.exe")
        ]
        for p in possibles:
            if p.exists(): return p
        return None

    def overwrite_image(self, template_path: Path, image_path: Path) -> bool:
        tmp_zip = template_path.with_suffix('.docx.tmp')
        try:
            img_data = image_path.read_bytes()
            replaced = False
            with zipfile.ZipFile(template_path, 'r') as zin, zipfile.ZipFile(tmp_zip, 'w') as zout:
                for item in zin.infolist():
                    if item.filename.startswith('word/media/') and item.filename.lower().endswith(('.png','.jpg','.jpeg')) and not replaced:
                        zout.writestr(item, img_data); replaced = True
                    else: zout.writestr(item, zin.read(item.filename))
            shutil.move(tmp_zip, template_path)
            return replaced
        except Exception:
            if tmp_zip.exists(): tmp_zip.unlink()
            return False

    def _process_docx_to_pdf(self, df: pd.DataFrame, fmt: str, out_name: str, do_rep: bool, rep_img: Path, progress_cb: Optional[Callable[[int, str], None]] = None) -> Optional[Path]:
        libre = self.find_libreoffice()
        tpl = self.workspace / "templates" / f"{fmt}-Auto.docx"
        
        if not libre: logger.error("❌ Không tìm thấy LibreOffice."); return None
        if not tpl.exists(): logger.error(f"❌ Thiếu Template: {tpl}"); return None
        if do_rep and rep_img.exists(): self.overwrite_image(tpl, rep_img)

        try:
            step = 1 if fmt == 'A4' else (2 if fmt == 'A5' else 8)
            keys = ["Ten", "Ma", "Model", "GiaNY", "GiaKM", "G", "Qua", "QuaA", "QuaB", "ThoiGian", "Hang"]
            
            records = df.to_dict('records')
            tasks = []
            word_lst = []

            for i in range(0, len(records), step):
                batch = records[i : i + step]
                w_out = self.out_dir / f"temp_{fmt}_{i}.docx"
                tasks.append((tpl, w_out, batch, fmt, step, keys))
                word_lst.append(w_out)

            if progress_cb: progress_cb(75, f"Đang gen {len(word_lst)} file Word đa luồng...")
            with concurrent.futures.ThreadPoolExecutor(max_workers=6) as executor:
                list(executor.map(render_single_docx, tasks))

            word_lst = [w for w in word_lst if w.exists()]
            if not word_lst: return None

            if progress_cb: progress_cb(85, "Đang convert sang PDF...")
            subprocess.run([
                str(libre), "--headless", "--convert-to", "pdf", 
                "--outdir", str(self.out_dir.resolve())
            ] + [str(w.resolve()) for w in word_lst], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)

            pdf_lst = [w.with_suffix('.pdf') for w in word_lst if w.with_suffix('.pdf').exists()]

            final_pdf = None
            if pdf_lst:
                if progress_cb: progress_cb(95, "Đang gộp PDF...")
                merger = PdfWriter()
                for p in pdf_lst: merger.append(str(p))
                final_pdf = self.out_dir / out_name
                merger.write(str(final_pdf))
                merger.close()
                logger.info(f"🎉 XONG! File: {final_pdf.name}")

            for f in word_lst + pdf_lst:
                try: f.unlink()
                except: pass

            if progress_cb: progress_cb(100, "Tạo PDF hoàn tất!")
            return final_pdf

        except Exception as e:
            logger.error(f"❌ Lỗi PDF Process: {e}")
            return None

    def generate_pdf(self, csv_file: Path, progress_cb: Optional[Callable[[int, str], None]] = None) -> Optional[Path]:
        logger.info("\n🔵 BƯỚC 4: TẠO PDF (LibreOffice)")
        if progress_cb: progress_cb(70, "Chuẩn bị tạo PDF...")
        fmt = self.config.get('pdf_format', 'A4')
        return self._process_docx_to_pdf(pd.read_csv(csv_file, encoding='utf-8-sig').fillna(""), fmt, f"BangGia_Final_{fmt}.pdf", self.config.get('use_image_replace', False), Path(self.config.get('replacement_image', '')), progress_cb)

    def generate_demo_pdf(self, progress_cb: Optional[Callable[[int, str], None]] = None) -> Optional[Path]:
        logger.info("\n🔵 TẠO FILE DEMO BẢNG GIÁ")
        if progress_cb: progress_cb(20, "Đang khởi tạo dữ liệu Demo...")
        fmt = self.config.get('pdf_format', 'A4')
        df = pd.DataFrame({
            'Mã': ['DL.044842', 'DL.044844', 'DL.044852', 'DL.044927'],
            'Tên': ['Tủ đông TOSHIBA', 'Máy giặt TOSHIBA', 'Máy giặt AQUA', 'Điều Hòa AQUA'],
            'Model': ['GR-RC390CM-PMV', 'TW-T21B120UWV', 'AWM10-B2158L', 'AQA-RUV10RB5'],
            'Giá NY': [7990000, 9890000, 8090000, 9990000], 'Giá KM': [6490000, 7990000, 5990000, 6990000],
            'Thời gian': ['09/04-12/04']*4,
            'Tên SP1': ['Phiếu 300k', '', '', 'Phiếu 300k'],
            'Tên SP2': ['Công lắp đặt', '', '', 'Công lắp đặt'],
            'Tên SP3': ['', '', '', '']
        })
        df['% Giảm giá'] = round((1 - df['Giá KM']/df['Giá NY'])*100, 0)
        return self._process_docx_to_pdf(df, fmt, f"Demo_BangGia_{fmt}.pdf", self.config.get('use_image_replace', False), Path(self.config.get('replacement_image', '')), progress_cb)