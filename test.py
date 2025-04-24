import pandas as pd
from datetime import datetime, timedelta, time
import os
from dotenv import load_dotenv
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from typing import List, Dict, Optional, Tuple, Set

# --- Constants ---
# File names
DEFAULT_ATTENDANCE_FILE = "attendance.xlsx"
DEFAULT_LEAVE_REQUESTS_FILE = "leave_requests.txt"
DEFAULT_EMAILS_FILE = "emails.csv"
DEFAULT_EMAIL_TEMPLATE_FILE = "Mau_Email.txt"

# Excel structure (Adjust if your structure differs)
HEADER_ROW_INDEX = 3 # Row index (0-based) where the date numbers are found
DATA_START_ROW_INDEX = 6 # First row index (0-based) containing employee data
NAME_COLUMN_INDEX = 0    # Column index (0-based) for Fullname
ROW_INCREMENT = 2      # Number of rows per employee entry

# Email settings & Content
SMTP_DEFAULT_SERVER = 'smtp.gmail.com'
SMTP_DEFAULT_PORT = 587
EMAIL_SUBJECT = "Thông báo vi phạm nội quy CLB Tiếng Anh"
DAYS_TO_HANDLE_DEFAULT = 7 # Deadline for handling the violation

# Violation details (Consider making these configurable if they change often)
VIOLATION_LATE = "Đi muộn"
VIOLATION_ABSENT = "Vắng không phép"
FINE_LATE = "5,000"  # Format as string for direct insertion
FINE_ABSENT = "20,000" # Format as string for direct insertion
COUNT_DEFAULT = "1" # Default violation count

# --- Functions ---

def danh_gia_di_muon_vang(
    ngay_nhap: int,
    gio_nhap_str: str,
    ten_file_excel: str = DEFAULT_ATTENDANCE_FILE
) -> Dict[str, List[str]]:
    """
    Đánh giá danh sách đi muộn và vắng dựa trên file Excel điểm danh.

    Args:
        ngay_nhap: Ngày cần kiểm tra (ví dụ: 2).
        gio_nhap_str: Giờ vào làm chuẩn dạng chuỗi (ví dụ: "18:00").
        ten_file_excel: Tên file Excel chứa dữ liệu điểm danh.

    Returns:
        Một dictionary chứa hai danh sách: 'di_muon' và 'vang'.
        Giá trị trong 'vang' có thể chứa thông báo lỗi nếu file/ngày không tìm thấy.
    """
    try:
        # Đọc file Excel, không dùng header mặc định vì cấu trúc phức tạp
        df = pd.read_excel(ten_file_excel, header=None)
    except FileNotFoundError:
        return {"di_muon": [], "vang": [f"Lỗi: Không tìm thấy file: {ten_file_excel}"]}
    except Exception as e:
        return {"di_muon": [], "vang": [f"Lỗi khi đọc file Excel: {e}"]}

    danh_sach_di_muon: List[str] = []
    danh_sach_vang: List[str] = []

    try:
        # Chuyển đổi giờ chuẩn sang đối tượng time một lần duy nhất
        gio_so_sanh: time = datetime.strptime(gio_nhap_str, '%H:%M').time()
    except ValueError:
        return {"di_muon": [], "vang": [f"Lỗi: Định dạng giờ nhập vào không hợp lệ: {gio_nhap_str}"]}

    # --- Tìm cột ngày và cột 'In' ---
    cot_ngay: Optional[int] = None
    ngay_str = str(ngay_nhap) # Chuyển ngày sang chuỗi để so sánh

    if HEADER_ROW_INDEX >= df.shape[0]:
         return {"di_muon": [], "vang": [f"Lỗi: File Excel không có hàng tiêu đề ngày (hàng index {HEADER_ROW_INDEX})."]}

    # Tìm cột chứa ngày nhập trên hàng HEADER_ROW_INDEX
    for col_idx in range(df.shape[1]):
        # Thêm kiểm tra pd.notna để tránh lỗi nếu ô bị trống
        cell_value = df.iloc[HEADER_ROW_INDEX, col_idx]
        if pd.notna(cell_value) and str(cell_value).strip() == ngay_str:
            cot_ngay = col_idx
            break

    if cot_ngay is None:
        return {"di_muon": [], "vang": [f"Không tìm thấy ngày {ngay_nhap} trên hàng {HEADER_ROW_INDEX + 1} trong file."]}

    # Xác định cột 'In' (giả định nó nằm ngay sau cột ngày)
    cot_in = cot_ngay + 1
    if cot_in >= df.shape[1]:
         return {"di_muon": [], "vang": [f"Lỗi: Không tìm thấy cột 'In' dự kiến tại index {cot_in} sau cột ngày {ngay_nhap}."]}

    # --- Lặp qua các hàng dữ liệu nhân viên ---
    row_idx = DATA_START_ROW_INDEX
    while row_idx < df.shape[0]:
        # Lấy tên nhân viên (kiểm tra ô tên không rỗng)
        ten_nhan_vien_raw = df.iloc[row_idx, NAME_COLUMN_INDEX]
        if pd.isna(ten_nhan_vien_raw) or str(ten_nhan_vien_raw).strip() == "":
             # Bỏ qua hàng nếu không có tên nhân viên, có thể là dòng trống hoặc cuối file
             row_idx += ROW_INCREMENT
             continue
        ten_nhan_vien = str(ten_nhan_vien_raw).strip()

        # Kiểm tra giá trị ở cột 'In'
        # Đảm bảo row_idx và cot_in nằm trong giới hạn của DataFrame
        if row_idx < df.shape[0] and cot_in < df.shape[1]:
            gio_vao_cell = df.iloc[row_idx, cot_in]

            if pd.notna(gio_vao_cell):
                try:
                    # Xử lý cả kiểu datetime và string
                    if isinstance(gio_vao_cell, (datetime, time)):
                        gio_vao = gio_vao_cell.time() # Lấy phần time nếu là datetime
                    else:
                        # Chuyển đổi chuỗi sang time, loại bỏ khoảng trắng thừa
                        gio_vao_str = str(gio_vao_cell).strip()
                        gio_vao = datetime.strptime(gio_vao_str, '%H:%M').time()

                    # So sánh giờ vào với giờ chuẩn
                    if gio_vao > gio_so_sanh:
                        danh_sach_di_muon.append(ten_nhan_vien)

                except ValueError:
                    # Ghi nhận lỗi nếu không thể phân tích cú pháp giờ
                    print(f"Cảnh báo: Không thể phân tích giờ '{gio_vao_cell}' cho {ten_nhan_vien} ở hàng {row_idx + 1}. Bỏ qua.")
                except Exception as e:
                     print(f"Cảnh báo: Lỗi không xác định khi xử lý giờ cho {ten_nhan_vien} ở hàng {row_idx + 1}: {e}. Bỏ qua.")

            else: # Nếu ô giờ vào là NaN/trống
                danh_sach_vang.append(ten_nhan_vien)
        else:
            # Nếu chỉ số hàng hoặc cột 'In' nằm ngoài giới hạn (ví dụ: cuối file)
            print(f"Cảnh báo: Đã đến cuối file hoặc cấu trúc bất thường tại hàng {row_idx + 1} cho {ten_nhan_vien}.")
            # Quyết định dừng hoặc tiếp tục tùy theo logic mong muốn
            # break # Dừng nếu muốn kết thúc khi gặp bất thường

        # Chuyển đến mục nhập của nhân viên tiếp theo
        row_idx += ROW_INCREMENT

    return {"di_muon": danh_sach_di_muon, "vang": danh_sach_vang}


def loai_bo_nguoi_nghi_phep(
    danh_sach_vang: List[str],
    ten_file_leave_requests: str = DEFAULT_LEAVE_REQUESTS_FILE
) -> List[str]:
    """
    Loại bỏ những người đã xin nghỉ phép khỏi danh sách vắng.

    Args:
        danh_sach_vang: Danh sách tên những người được đánh giá là vắng.
        ten_file_leave_requests: Tên file chứa danh sách người xin nghỉ phép.

    Returns:
        Danh sách vắng sau khi lọc.
    """
    try:
        with open(ten_file_leave_requests, "r", encoding="utf-8") as file:
            # Đọc và loại bỏ khoảng trắng, chuyển thành set để tối ưu việc kiểm tra
            nguoi_nghi_phep: Set[str] = {line.strip() for line in file if line.strip()}
    except FileNotFoundError:
        print(f"Cảnh báo: Không tìm thấy file nghỉ phép {ten_file_leave_requests}. Sẽ không loại trừ ai.")
        return danh_sach_vang
    except Exception as e:
         print(f"Lỗi khi đọc file nghỉ phép {ten_file_leave_requests}: {e}. Sẽ không loại trừ ai.")
         return danh_sach_vang

    # Sử dụng list comprehension và kiểm tra trong set (nhanh hơn list)
    danh_sach_vang_sau_loc = [
        ten for ten in danh_sach_vang if ten not in nguoi_nghi_phep
    ]
    return danh_sach_vang_sau_loc


def tao_noi_dung_email(
    danh_sach_vang: List[str],
    danh_sach_di_muon: List[str],
    ten_file_emails: str = DEFAULT_EMAILS_FILE,
    ten_file_mau: str = DEFAULT_EMAIL_TEMPLATE_FILE
) -> Dict[str, str]:
    """
    Tạo nội dung email cá nhân hóa cho người vi phạm.

    Args:
        danh_sach_vang: Danh sách người vắng.
        danh_sach_di_muon: Danh sách người đi muộn.
        ten_file_emails: Tên file CSV chứa thông tin email (cột 'ten', 'email').
        ten_file_mau: Tên file chứa mẫu email.

    Returns:
        Dictionary với key là email người nhận, value là nội dung email.
    """
    # Đọc file mẫu email
    try:
        with open(ten_file_mau, "r", encoding="utf-8") as f:
            mau_email_base = f.read()
    except FileNotFoundError:
        print(f"Lỗi: Không tìm thấy file mẫu email: {ten_file_mau}")
        return {}
    except Exception as e:
        print(f"Lỗi khi đọc file mẫu email: {e}")
        return {}

    # Đọc danh sách email và tạo map để truy cập nhanh
    try:
        df_emails = pd.read_csv(ten_file_emails)
        # Kiểm tra cột cần thiết có tồn tại không
        if 'ten' not in df_emails.columns or 'email' not in df_emails.columns:
            print(f"Lỗi: File {ten_file_emails} phải chứa cột 'ten' và 'email'.")
            return {}
        # Tạo dictionary từ tên sang email, xử lý tên trùng lặp (lấy email đầu tiên)
        email_map = pd.Series(df_emails.email.values, index=df_emails.ten).to_dict()
    except FileNotFoundError:
        print(f"Lỗi: Không tìm thấy file emails: {ten_file_emails}")
        return {}
    except Exception as e:
        print(f"Lỗi khi đọc file emails {ten_file_emails}: {e}")
        return {}

    emails_to_send: Dict[str, str] = {}
    han_xu_ly = (datetime.now() + timedelta(days=DAYS_TO_HANDLE_DEFAULT)).strftime("%d/%m/%Y")

    # Hàm trợ giúp để tạo nội dung email
    def _tao_noi_dung(ten: str, ly_do: str, so_tien: str) -> Optional[str]:
        email_nhan = email_map.get(ten)
        if not email_nhan:
            print(f"Cảnh báo: Không tìm thấy email cho '{ten}' trong {ten_file_emails}.")
            return None
        if not isinstance(email_nhan, str) or '@' not in email_nhan:
             print(f"Cảnh báo: Email không hợp lệ ('{email_nhan}') cho '{ten}'. Bỏ qua.")
             return None

        # Sử dụng f-string hoặc str.format để thay thế dễ đọc hơn
        noi_dung = mau_email_base.replace("[Tên thành viên]", ten)
        noi_dung = noi_dung.replace("[Ví dụ: Đi họp muộn, nghỉ không phép, chưa đóng quỹ…]", ly_do)
        noi_dung = noi_dung.replace("[Số lần]", COUNT_DEFAULT)
        noi_dung = noi_dung.replace("[Số tiền]", so_tien)
        noi_dung = noi_dung.replace("[ngày/tháng/năm]", han_xu_ly)
        return noi_dung

    # Xử lý danh sách đi muộn
    for ten in danh_sach_di_muon:
        noi_dung = _tao_noi_dung(ten, VIOLATION_LATE, FINE_LATE)
        if noi_dung:
             email_nhan = email_map.get(ten) # Lấy lại email đã được xác thực
             if email_nhan:
                emails_to_send[email_nhan] = noi_dung


    # Xử lý danh sách vắng
    for ten in danh_sach_vang:
        noi_dung = _tao_noi_dung(ten, VIOLATION_ABSENT, FINE_ABSENT)
        if noi_dung:
            email_nhan = email_map.get(ten) # Lấy lại email đã được xác thực
            if email_nhan:
                # Kiểm tra nếu người này vừa đi muộn vừa vắng (ít khả năng nhưng đề phòng)
                # Hoặc nếu muốn gửi email vắng thay vì đi muộn nếu có cả 2 lỗi
                emails_to_send[email_nhan] = noi_dung # Ghi đè nếu đã có email đi muộn

    return emails_to_send


def gui_email(emails_data: Dict[str, str]) -> Dict[str, str]:
    """
    Gửi email thông báo vi phạm tới danh sách người nhận.

    Args:
        emails_data: Dictionary với key là email người nhận, value là nội dung email.

    Returns:
        Dictionary với key là email, value là trạng thái gửi ("Thành công" hoặc "Lỗi: ...").
    """
    if not emails_data:
        print("Không có email nào để gửi.")
        return {}

    # Load biến môi trường từ file .env
    load_dotenv()
    EMAIL_ADDRESS = os.getenv('EMAIL_ADDRESS')
    EMAIL_PASSWORD = os.getenv('EMAIL_PASSWORD')
    SMTP_SERVER = os.getenv('SMTP_SERVER', SMTP_DEFAULT_SERVER)
    # Đảm bảo port là số nguyên
    try:
        SMTP_PORT = int(os.getenv('SMTP_PORT', str(SMTP_DEFAULT_PORT)))
    except ValueError:
        print(f"Lỗi: SMTP_PORT trong .env không phải là số. Sử dụng port mặc định {SMTP_DEFAULT_PORT}.")
        SMTP_PORT = SMTP_DEFAULT_PORT

    if not EMAIL_ADDRESS or not EMAIL_PASSWORD:
        print("Lỗi: Thiếu EMAIL_ADDRESS hoặc EMAIL_PASSWORD trong file .env. Không thể gửi email.")
        return {email: "Lỗi: Thiếu cấu hình email gửi" for email in emails_data}

    ket_qua: Dict[str, str] = {}
    server: Optional[smtplib.SMTP] = None # Khởi tạo server là None

    try:
        # Thiết lập kết nối SMTP
        # Sử dụng context manager (with) để đảm bảo server.quit() được gọi
        print(f"Đang kết nối tới {SMTP_SERVER}:{SMTP_PORT}...")
        server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT, timeout=30) # Thêm timeout
        server.ehlo() # Chào hỏi server
        server.starttls() # Bắt đầu mã hóa TLS
        server.ehlo() # Chào hỏi lại sau TLS
        server.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
        print("Kết nối và đăng nhập SMTP thành công.")

        # Gửi email cho từng người
        for email_nhan, noi_dung in emails_data.items():
            try:
                msg = MIMEMultipart()
                msg['From'] = EMAIL_ADDRESS
                msg['To'] = email_nhan
                msg['Subject'] = EMAIL_SUBJECT

                # Đính kèm nội dung email với encoding utf-8
                msg.attach(MIMEText(noi_dung, 'plain', 'utf-8'))

                # Gửi email
                server.send_message(msg)
                ket_qua[email_nhan] = "Thành công"
                print(f"Đã gửi email thành công tới: {email_nhan}")

            except smtplib.SMTPRecipientsRefused:
                error_msg = "Địa chỉ người nhận bị từ chối."
                ket_qua[email_nhan] = f"Lỗi: {error_msg}"
                print(f"Lỗi khi gửi email tới {email_nhan}: {error_msg}")
            except Exception as e:
                error_msg = str(e)
                ket_qua[email_nhan] = f"Lỗi: {error_msg}"
                print(f"Lỗi khi gửi email tới {email_nhan}: {error_msg}")

    except smtplib.SMTPAuthenticationError:
        print("Lỗi: Xác thực SMTP thất bại. Kiểm tra EMAIL_ADDRESS và EMAIL_PASSWORD.")
        ket_qua = {email: "Lỗi: Xác thực SMTP thất bại" for email in emails_data}
    except smtplib.SMTPServerDisconnected:
         print("Lỗi: Mất kết nối đến máy chủ SMTP.")
         ket_qua = {email: "Lỗi: Mất kết nối SMTP" for email in emails_data}
    except ConnectionRefusedError:
         print(f"Lỗi: Kết nối đến {SMTP_SERVER}:{SMTP_PORT} bị từ chối. Kiểm tra địa chỉ/port và tường lửa.")
         ket_qua = {email: "Lỗi: Kết nối SMTP bị từ chối" for email in emails_data}
    except Exception as e:
        print(f"Lỗi kết nối SMTP hoặc lỗi không xác định khác: {str(e)}")
        # Gán lỗi chung cho tất cả nếu lỗi xảy ra trước vòng lặp gửi
        if not ket_qua:
             ket_qua = {email: f"Lỗi SMTP chung: {str(e)}" for email in emails_data}
        # Nếu lỗi xảy ra sau khi đã gửi một số email, chỉ cập nhật trạng thái cho những email chưa gửi
        else:
             for email in emails_data:
                  if email not in ket_qua:
                       ket_qua[email] = f"Lỗi SMTP chung: {str(e)}"

    finally:
        # Đảm bảo đóng kết nối ngay cả khi có lỗi
        if server:
            try:
                server.quit()
                print("Đã đóng kết nối SMTP.")
            except Exception as e:
                print(f"Lỗi khi đóng kết nối SMTP: {e}")

    return ket_qua


def main():
    """Hàm chính điều phối các bước thực thi."""
    # --- Cấu hình đầu vào ---
    # Lấy ngày hôm qua làm ngày mặc định để kiểm tra
    yesterday = datetime.now() - timedelta(days=1)
    ngay_can_kiem_tra_default = yesterday.day
    # Hoặc bạn có thể nhập từ người dùng hoặc tham số dòng lệnh
    try:
      ngay_input = input(f"Nhập ngày cần kiểm tra (mặc định: {ngay_can_kiem_tra_default}): ")
      ngay_can_kiem_tra = int(ngay_input) if ngay_input else ngay_can_kiem_tra_default
    except ValueError:
       print("Ngày nhập không hợp lệ, sử dụng ngày mặc định.")
       ngay_can_kiem_tra = ngay_can_kiem_tra_default

    gio_vao_so_sanh_default = "18:00"
    gio_vao_input = input(f"Nhập giờ vào làm chuẩn (HH:MM, mặc định: {gio_vao_so_sanh_default}): ")
    gio_vao_so_sanh = gio_vao_input if gio_vao_input else gio_vao_so_sanh_default
    # Thêm kiểm tra định dạng giờ ở đây nếu muốn chặt chẽ hơn

    # --- Thực thi ---
    print(f"\n--- Bắt đầu kiểm tra điểm danh ngày {ngay_can_kiem_tra} so với {gio_vao_so_sanh} ---")
    ket_qua_diem_danh = danh_gia_di_muon_vang(ngay_can_kiem_tra, gio_vao_so_sanh)

    # Kiểm tra nếu có lỗi trả về từ hàm danh_gia
    if isinstance(ket_qua_diem_danh["vang"], list) and ket_qua_diem_danh["vang"] and ket_qua_diem_danh["vang"][0].startswith("Lỗi:"):
        print(f"Không thể tiếp tục do lỗi: {ket_qua_diem_danh['vang'][0]}")
        return # Dừng thực thi

    danh_sach_di_muon = ket_qua_diem_danh["di_muon"]
    danh_sach_vang_ban_dau = ket_qua_diem_danh["vang"]

    print(f"\nDanh sách đi muộn ban đầu ({len(danh_sach_di_muon)}): {danh_sach_di_muon}")
    print(f"Danh sách vắng ban đầu ({len(danh_sach_vang_ban_dau)}): {danh_sach_vang_ban_dau}")

    print("\n--- Loại bỏ người nghỉ phép khỏi danh sách vắng ---")
    danh_sach_vang_sau_loc = loai_bo_nguoi_nghi_phep(danh_sach_vang_ban_dau)
    nguoi_bi_loai = len(danh_sach_vang_ban_dau) - len(danh_sach_vang_sau_loc)
    if nguoi_bi_loai > 0:
        print(f"Đã loại bỏ {nguoi_bi_loai} người nghỉ phép.")
    print(f"Danh sách vắng cuối cùng ({len(danh_sach_vang_sau_loc)}): {danh_sach_vang_sau_loc}")

    if not danh_sach_di_muon and not danh_sach_vang_sau_loc:
        print("\nKhông có ai đi muộn hoặc vắng cần gửi email. Kết thúc.")
        return

    print("\n--- Tạo nội dung email ---")
    emails_can_gui = tao_noi_dung_email(danh_sach_vang_sau_loc, danh_sach_di_muon)

    if not emails_can_gui:
        print("Không tạo được nội dung email nào (có thể do lỗi file hoặc không tìm thấy email).")
        return

    # In ra để kiểm tra (tùy chọn)
    # print("\nNội dung email sẽ gửi:")
    # for email, noi_dung in emails_can_gui.items():
    #     print(f"\nTo: {email}")
    #     print("--- Start Content ---")
    #     print(noi_dung)
    #     print("--- End Content ---")
    #     print("-" * 50)

    # Xác nhận trước khi gửi (tùy chọn)
    confirm = input(f"\nTìm thấy {len(emails_can_gui)} email cần gửi. Bạn có muốn gửi không? (y/n): ").lower()
    if confirm != 'y':
        print("Đã hủy gửi email.")
        return

    # Gửi email
    print("\n--- Bắt đầu gửi email ---")
    ket_qua_gui = gui_email(emails_can_gui)

    # In kết quả
    print("\n--- Kết quả gửi email ---")
    thanh_cong_count = 0
    that_bai_count = 0
    for email, trang_thai in ket_qua_gui.items():
        print(f"{email}: {trang_thai}")
        if trang_thai == "Thành công":
            thanh_cong_count += 1
        else:
            that_bai_count +=1
    print(f"\nTổng kết: Gửi thành công {thanh_cong_count} email, thất bại {that_bai_count} email.")


# Chạy hàm main khi script được thực thi trực tiếp
if __name__ == "__main__":
    main()