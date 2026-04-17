import imaplib
import email
from email.header import decode_header
from bs4 import BeautifulSoup

# --- CẤU HÌNH TEST ---
MAIL_SERVER = "mail.hc.com.vn"
MAIL_USER = "duynguyen1@hc.com.vn"
MAIL_PASS = "Vhc123@"
SEARCH_SUBJECT = "SP.077460" # Hoặc từ khóa bạn muốn test

def test_imap_search():
    print(f"🔄 Đang kết nối đến {MAIL_SERVER}...")
    try:
        # Khởi tạo kết nối bảo mật SSL
        mail = imaplib.IMAP4_SSL(MAIL_SERVER, 993) # Đổi port thành 993 nếu 995 không hoạt động với IMAP của HC
        mail.login(MAIL_USER, MAIL_PASS)
        print("✅ Đăng nhập thành công!")

        # Chọn hộp thư đến (INBOX)
        mail.select('inbox')
        
        print(f"🔍 Đang nhờ Server tìm kiếm thư có tiêu đề chứa: '{SEARCH_SUBJECT}'...")
        # Lọc thư chứa từ khóa (Có thể thêm cờ UNSEEN nếu chỉ muốn tìm thư chưa đọc)
        status, search_data = mail.search(None, f'(SUBJECT "{SEARCH_SUBJECT}")')
        
        if status != 'OK':
            print("❌ Lỗi khi tìm kiếm trên server.")
            return

        mail_ids = search_data[0].split()
        if not mail_ids:
            print("⚠️ Không tìm thấy email nào khớp với từ khóa.")
            return
            
        print(f"📬 Tìm thấy {len(mail_ids)} email. Đang lấy email mới nhất...")
        
        # Lấy ID của email mới nhất (nằm ở cuối danh sách)
        latest_email_id = mail_ids[-1]
        
        # Chỉ Fetch nội dung của đúng email này (nhanh hơn Fetch POP3 toàn bộ rất nhiều)
        status, msg_data = mail.fetch(latest_email_id, '(RFC822)')
        
        for response_part in msg_data:
            if isinstance(response_part, tuple):
                msg = email.message_from_bytes(response_part[1])
                
                # Giải mã Subject
                subject, encoding = decode_header(msg["Subject"])[0]
                if isinstance(subject, bytes):
                    subject = subject.decode(encoding if encoding else "utf-8")
                
                print("\n" + "="*50)
                print(f"📩 TIÊU ĐỀ: {subject}")
                print("="*50)

                # Lấy nội dung Body
                if msg.is_multipart():
                    for part in msg.walk():
                        if part.get_content_type() == "text/html":
                            body = part.get_payload(decode=True).decode('utf-8', 'ignore')
                            soup = BeautifulSoup(body, 'html.parser') # Cài 'lxml' sẽ nhanh hơn
                            print("📝 TRÍCH XUẤT 500 KÝ TỰ ĐẦU TIÊN CỦA HTML:")
                            print(soup.get_text()[:500] + "...\n")
                            break
                else:
                    body = msg.get_payload(decode=True).decode('utf-8', 'ignore')
                    print("📝 TRÍCH XUẤT NỘI DUNG:")
                    print(body[:500] + "...\n")
                    
        mail.logout()
        print("🔌 Đã ngắt kết nối an toàn.")

    except Exception as e:
        print(f"❌ Xảy ra lỗi: {e}")

if __name__ == "__main__":
    test_imap_search()