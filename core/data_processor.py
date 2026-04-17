from __future__ import annotations
import csv
import email
import json
import re
from datetime import datetime
from email import policy
from email.header import decode_header
from pathlib import Path
from typing import Optional, Dict, Callable

import pandas as pd
import requests
from bs4 import BeautifulSoup
from utils.logger import logger

class DataProcessor:
    def __init__(self, config: Dict):
        self.config = config
        self.workspace = Path.cwd()
        self.csv_dir = self.workspace / "Data_CSV"
        self.csv_dir.mkdir(exist_ok=True)

    def fetch_inventory(self, use_existing: bool = False, progress_cb: Optional[Callable[[int, str], None]] = None) -> Optional[Path]:
        if progress_cb: progress_cb(5, "Đang chuẩn bị tải tồn kho...")
        target = self.config.get('target_kho', 'KHO_01').replace(' ', '_')
        
        if use_existing:
            logger.info("🔵 BƯỚC 1: TÌM FILE TỒN KHO CÓ SẴN")
            files = list(self.csv_dir.glob(f"Tonkho_{target}*.csv")) or list(self.csv_dir.glob("Tonkho_*.csv"))
            if not files:
                logger.error("❌ Không tìm thấy file Tồn kho trong Data_CSV/.")
                return None
            latest_file = max(files, key=lambda f: f.stat().st_ctime)
            logger.info(f"✅ Sử dụng file: {latest_file.name}")
            if progress_cb: progress_cb(30, "Đã lấy tồn kho có sẵn!")
            return latest_file

        logger.info("🔵 BƯỚC 1: TẢI TỒN KHO TỪ ERP")
        user, pwd = self.config.get('erp_user', ''), self.config.get('erp_pass', '')
        if not user or not pwd:
            logger.error("❌ LỖI: Thiếu tài khoản ERP.")
            return None

        try:
            s = requests.Session()
            s.verify = False
            headers = {"User-Agent": "Mozilla/5.0"}
            
            if progress_cb: progress_cb(10, "Đang kết nối đến ERP...")
            r = s.get("https://192.168.1.88/ords/f?p=137:LOGIN_DESKTOP::::::", headers=headers)
            soup = BeautifulSoup(r.text, 'html.parser')
            
            p_inst = soup.find('input', {'id': 'pInstance'})['value'] if soup.find('input', {'id': 'pInstance'}) else ""
            p_sub = soup.find('input', {'id': 'pPageSubmissionId'})['value'] if soup.find('input', {'id': 'pPageSubmissionId'}) else ""
            p_prot = soup.find('input', {'id': 'pPageItemsProtected'})['value'] if soup.find('input', {'id': 'pPageItemsProtected'}) else ""
            p_salt = soup.find('input', {'id': 'pSalt'})['value'] if soup.find('input', {'id': 'pSalt'}) else ""
            
            data = {
                "p_flow_id": "137", "p_flow_step_id": "101", "p_instance": p_inst,
                "p_request": "LOGIN", "p_reload_on_submit": "S", "p_page_submission_id": p_sub,
                "p_json": json.dumps({
                    "pageItems": {"itemsToSubmit": [{"n": "P101_USERNAME", "v": user}, {"n": "P101_PASSWORD", "v": pwd}], "protected": p_prot},
                    "salt": p_salt or p_sub
                })
            }
            s.post("https://192.168.1.88/ords/wwv_flow.accept", headers=headers, data=data)
            
            if progress_cb: progress_cb(20, "Đang tải file CSV từ ERP...")
            r_csv = s.get(f"https://192.168.1.88/ords/f?p=137:19:{p_inst}:CSV_N::::", headers=headers)
            
            raw_csv = self.csv_dir / "temp_raw.csv"
            raw_csv.write_text(r_csv.text, encoding='utf-8-sig')
            
            if progress_cb: progress_cb(25, f"Đang lọc dữ liệu cho {target}...")
            df = pd.read_csv(raw_csv, encoding='utf-8-sig')
            df['SL'] = pd.to_numeric(df['SL'], errors='coerce').fillna(0)
            
            df_filt = df[(df['Kho'] == target) & (df['SL'] > 0)]
            if df_filt.empty:
                logger.error(f"❌ Kho {target} không có tồn.")
                return None
            
            ts = datetime.now().strftime("%Y%m%d_%H%M")
            out_file = self.csv_dir / f"Tonkho_{target}_{ts}.csv"
            df_filt.to_csv(out_file, index=False, encoding='utf-8-sig')
            
            if raw_csv.exists(): raw_csv.unlink()
            logger.info(f"✅ Đã lưu: {out_file.name}")
            if progress_cb: progress_cb(30, "Tải tồn kho hoàn tất!")
            return out_file

        except Exception as e:
            logger.error(f"❌ Lỗi ERP: {e}")
            return None

    def fetch_email(self, progress_cb: Optional[Callable[[int, str], None]] = None) -> Optional[Path]:
        logger.info("\n🔵 BƯỚC 2: QUÉT EMAIL (IMAP HYBRID)")
        if progress_cb: progress_cb(35, "Đang kết nối IMAP Server...")
        
        u, p = self.config.get('mail_user', ''), self.config.get('mail_pass', '')
        sub_list = [s.strip() for s in self.config.get('mail_subjects', '').split(',') if s.strip()]
        
        if not sub_list:
            logger.error("❌ Chưa nhập từ khóa mail.")
            return None

        try:
            import imaplib
            mail = imaplib.IMAP4_SSL("mail.hc.com.vn", 993) 
            mail.login(u, p)
            mail.select('inbox')
            
            _, messages = mail.search(None, 'ALL')
            mail_ids = messages[0].split()
            
            if not mail_ids: return None

            scan_limit = 50
            latest_ids = mail_ids[-scan_limit:]
            if progress_cb: progress_cb(40, f"Đang check Header {len(latest_ids)} email...")
            
            cnt = 0
            csv_out = self.csv_dir / "email_tables.csv"
            
            with open(csv_out, 'w', newline='', encoding='utf-8-sig') as f:
                w = csv.writer(f); head_ok = False
                
                for num in reversed(latest_ids):
                    try:
                        _, header_data = mail.fetch(num, '(RFC822.HEADER)')
                        for response_part in header_data:
                            if isinstance(response_part, tuple):
                                msg = email.message_from_bytes(response_part[1], policy=policy.default)
                                sb = msg.get('Subject', '')
                                
                                decoded_list = decode_header(sb)
                                d_sb = ""
                                for text, encoding in decoded_list:
                                    d_sb += text.decode(encoding if encoding else "utf-8", 'ignore') if isinstance(text, bytes) else str(text)

                                if any(k in d_sb for k in sub_list):
                                    cnt += 1
                                    if progress_cb: progress_cb(40 + min(10, cnt*2), f"Đang bóc tách mail {cnt}...")
                                    logger.info(f"➡️ Bóc tách: {d_sb}")
                                    
                                    _, body_data = mail.fetch(num, '(RFC822)')
                                    for bp in body_data:
                                        if isinstance(bp, tuple):
                                            full_msg = email.message_from_bytes(bp[1], policy=policy.default)
                                            body = ""
                                            for pt in full_msg.walk():
                                                if pt.get_content_type() in ['text/html', 'text/plain']:
                                                    body = pt.get_payload(decode=True).decode('utf-8', 'ignore')
                                                    break
                                            
                                            t_km = datetime.now().strftime("%d/%m")
                                            if body:
                                                soup = BeautifulSoup(body, 'html.parser')
                                                s_match = re.search(r"Từ ngày/giờ:.*?(\d{1,2}/\d{1,2})", soup.get_text(), re.IGNORECASE)
                                                e_match = re.search(r"Đến ngày/giờ:.*?(\d{1,2}/\d{1,2})", soup.get_text(), re.IGNORECASE)
                                                if s_match and e_match: t_km = f"{s_match.group(1)}-{e_match.group(1)}"

                                                tbl = soup.find('table', {'border': '1'})
                                                if tbl:
                                                    cols = ['Tiêu đề', 'Thời gian', 'Mã', 'Tên', 'Model', 'Serial', 'TT', 'Giá NY', 'Giá KM', 
                                                            'Mã SP1', 'Tên SP1', 'Giá trị1', 'Giá trị KM1', 'Giá bán1',
                                                            'Mã SP2', 'Tên SP2', 'Giá trị2', 'Giá trị KM2', 'Giá bán2',
                                                            'Mã SP3', 'Tên SP3', 'Giá trị3', 'Giá trị KM3', 'Giá bán3',
                                                            'Mã SP4', 'Tên SP4', 'Giá trị4', 'Giá trị KM4', 'Giá bán4',
                                                            'Mã SP5', 'Tên SP5', 'Giá trị5', 'Giá trị KM5', 'Giá bán5',
                                                            '$Thẻ', '$CK trực tiếp', '%CK duyệt MAX', 'Giá sau KM', 'NVBH']
                                                    if not head_ok: w.writerow(cols); head_ok = True
                                                    for tr in tbl.find_all('tr')[1:]:
                                                        tds = [c.get_text(strip=True).replace('\xa0', ' ') for c in tr.find_all('td')]
                                                        row = ([d_sb, t_km] + tds)[:len(cols)]
                                                        row.extend(['']*(len(cols)-len(row)))
                                                        w.writerow(row)
                    except Exception as loop_e:
                        continue
            mail.logout()
            
            if cnt > 0:
                if progress_cb: progress_cb(55, "Quét email hoàn tất!")
                return csv_out
            else:
                logger.warning("⚠️ Không thấy email nào chứa từ khóa.")
                return None
                
        except Exception as e:
            logger.error(f"❌ Lỗi Mail (IMAP): {e}")
            return None

    def merge_and_calculate(self, mail_csv: Path, tonkho_csv: Path, progress_cb: Optional[Callable[[int, str], None]] = None) -> Optional[Path]:
        logger.info("\n🔵 BƯỚC 3: TÍNH GIÁ")
        if progress_cb: progress_cb(60, "Đang xử lý Pandas...")
        try:
            dm = pd.read_csv(mail_csv, encoding='utf-8-sig')
            dt = pd.read_csv(tonkho_csv, encoding='utf-8-sig')
            df = dm[dm['Mã'].isin(dt['Mã SP'])]
            
            if df.empty:
                logger.warning("⚠️ Không có mã hàng nào khớp.")
                return None

            out_cols = ['Mã', 'Tên', 'Model', 'Giá NY', 'Giá KM', 'Thời gian', 'Tên SP1', 'Tên SP2', 'Tên SP3']
            res = df[[c for c in out_cols if c in df.columns]].copy()

            def clean(v): 
                try: return float(re.sub(r'[^\d]', '', str(v).replace('VND','').split('.')[0]))
                except: return None

            res['Giá NY'] = res['Giá NY'].apply(clean)
            res['Giá KM'] = res['Giá KM'].apply(clean)
            res['% Giảm giá'] = res.apply(lambda r: round((1 - r['Giá KM']/r['Giá NY'])*100, 0) if r['Giá NY'] else 0, axis=1)
            
            f_out = self.csv_dir / "filtered_email_tables.csv"
            res.to_csv(f_out, encoding='utf-8-sig', index=False)
            if progress_cb: progress_cb(70, "Xử lý dữ liệu hoàn tất!")
            return f_out
        except Exception as e:
            logger.error(f"❌ Lỗi tính toán: {e}")
            return None