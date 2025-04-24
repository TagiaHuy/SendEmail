import streamlit as st
import pandas as pd
from datetime import datetime, time, timedelta
import os

# Must be the first Streamlit command
st.set_page_config(
    page_title="ESL HR CMC - Kiểm tra Điểm danh & Gửi Email",
    layout="wide"
)

# Import các hàm từ file script gốc của bạn
# Đảm bảo file attendance_checker.py nằm cùng thư mục với app.py
try:
    from attendance_checker import (
        danh_gia_di_muon_vang,
        loai_bo_nguoi_nghi_phep,
        tao_noi_dung_email,
        gui_email,
        luu_log,  # Add this import
        DEFAULT_ATTENDANCE_FILE,
        DEFAULT_LEAVE_REQUESTS_FILE,
        DEFAULT_EMAILS_FILE,
        DEFAULT_EMAIL_TEMPLATE_FILE,
        VIOLATION_LATE,
        VIOLATION_ABSENT,
        FINE_LATE,
        FINE_ABSENT
    )
    functions_loaded = True
except ImportError as e:
    st.error(f"Lỗi import hàm từ 'attendance_checker.py': {e}")
    st.error("Hãy đảm bảo file 'attendance_checker.py' chứa các hàm cần thiết và nằm cùng thư mục với 'app.py'")
    functions_loaded = False
except Exception as e:
    st.error(f"Lỗi không xác định khi import: {e}")
    functions_loaded = False


# Helper functions
def save_uploaded_file(uploaded_file, destination):
    """Helper function to save uploaded files safely"""
    if uploaded_file:
        try:
            with open(destination, "wb") as f:
                f.write(uploaded_file.getvalue())
            return True
        except Exception as e:
            st.sidebar.error(f"Lỗi khi lưu file: {str(e)}")
            return False
    return False

def create_file_upload_section(label, default_value, file_types, key):
    """
    Create a file upload section with a single column layout.
    
    Args:
        label (str): Label for the file section
        default_value (str): Default file path
        file_types (list): Allowed file extensions
        key (str): Unique key for Streamlit widgets
        
    Returns:
        str: Selected or uploaded file path
    """
    st.sidebar.markdown(f"**{label}:**")
    
    # Show current file path
    file_path = st.sidebar.text_input(
        "Đường dẫn hiện tại",
        value=default_value,
        key=f"path_{key}"
    )
    
    # File uploader
    uploaded_file = st.sidebar.file_uploader(
        "Tải lên file mới",
        type=file_types,
        key=key,
        help=f"Các định dạng hỗ trợ: {', '.join(file_types)}"
    )
    
    # Handle file upload
    if uploaded_file:
        if save_uploaded_file(uploaded_file, default_value):
            st.sidebar.success(f"✅ Đã tải lên: {uploaded_file.name}")
            file_path = default_value
        else:
            st.sidebar.error("❌ Tải lên thất bại")
    
    st.sidebar.markdown("---")
    return file_path

def display_results_table(data, title):
    """Display a consistent format for results tables"""
    st.write(f"**{title}:**")
    if data:
        st.dataframe(pd.DataFrame(data, columns=["Tên"]))
        return len(data)
    else:
        st.info(f"Không có {title.lower()}")
        return 0

def initialize_app():
    """Initialize the Streamlit app with custom header and styling"""
    st.markdown("""
        <style>
        .block-container {
            padding-top: 1rem;
            padding-bottom: 0rem;
        }
        .main > div {
            padding-left: 2rem;
            padding-right: 2rem;
        }
        h1 {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
        }
        </style>
    """, unsafe_allow_html=True)
    
    # Create header with logo and title
    col1, col2 = st.columns([1, 4])
    
    with col1:
        try:
            st.image(
                "assets/logo.jpg",
                width=1600,
                output_format="JPG"
            )
        except:
            st.error("Không tìm thấy file logo trong thư mục assets/")
    
    with col2:
        st.markdown("""
            <h1 style='color: #2E86C1; margin-bottom: 0px;'>
                ESL HR CMC
            </h1>
            <h3 style='color: #5D6D7E; margin-top: 0px;'>
                Công cụ Kiểm tra Điểm danh và Gửi Email Thông báo
            </h3>
        """, unsafe_allow_html=True)
    
    # Add a separator
    st.markdown("<hr>", unsafe_allow_html=True)

def setup_sidebar():
    st.sidebar.header("Cấu hình Đầu vào")
    
    # Date and time inputs
    yesterday = datetime.now() - timedelta(days=1)
    ngay_can_kiem_tra = st.sidebar.number_input(
        "Ngày cần kiểm tra",
        min_value=1, max_value=31,
        value=yesterday.day, step=1
    )
    
    gio_vao_so_sanh_time = st.sidebar.time_input(
        "Giờ vào làm chuẩn",
        value=time(18, 0)
    )
    
    # File paths
    st.sidebar.markdown("---")
    st.sidebar.subheader("Đường dẫn File")
    
    file_configs = {
        "Excel điểm danh": (DEFAULT_ATTENDANCE_FILE, ['xlsx', 'xls']),
        "danh sách nghỉ phép": (DEFAULT_LEAVE_REQUESTS_FILE, ['txt']),
        "CSV emails": (DEFAULT_EMAILS_FILE, ['csv']),
        "mẫu Email": (DEFAULT_EMAIL_TEMPLATE_FILE, ['txt'])
    }
    
    file_paths = {
        key: create_file_upload_section(key, default, types, f"upload_{key}")
        for key, (default, types) in file_configs.items()
    }
    
    return ngay_can_kiem_tra, gio_vao_so_sanh_time, file_paths

def process_attendance(ngay, gio, file_paths):
    """Process attendance and return results"""
    with st.spinner("Đang đọc và phân tích file điểm danh..."):
        ket_qua = danh_gia_di_muon_vang(
            ngay_nhap=ngay,
            gio_nhap_str=gio.strftime('%H:%M'),
            ten_file_excel=file_paths["Excel điểm danh"]
        )
        
        if isinstance(ket_qua["vang"], list) and ket_qua["vang"] and ket_qua["vang"][0].startswith("Lỗi:"):
            return None, ket_qua["vang"][0]
            
        return ket_qua, None

def view_log_history(log_file="email_logs.txt"):
    """
    Display the log history in a formatted way using Streamlit.
    
    Args:
        log_file (str): Path to the log file
    """
    try:
        with open(log_file, "r", encoding="utf-8") as f:
            logs = f.read().split("="*50)
            logs = [log.strip() for log in logs if log.strip()]
            
        if not logs:
            st.info("Chưa có lịch sử gửi email.")
            return
            
        for log in logs:
            with st.expander(f"Log {log.split('Thời gian ghi log: ')[1].split('\n')[0]}", expanded=False):
                st.text(log)
                
    except FileNotFoundError:
        st.info("Chưa có file log.")
    except Exception as e:
        st.error(f"Lỗi khi đọc log: {str(e)}")

def main():
    if not functions_loaded:
        st.warning("Không thể tải các chức năng xử lý. Vui lòng kiểm tra lỗi import.")
        return
        
    # Initialize session state if needed
    if 'processed_data' not in st.session_state:
        st.session_state.processed_data = None
    
    # Setup UI
    initialize_app()
    ngay, gio, file_paths = setup_sidebar()
    col1, col2 = st.columns(2)
    
    # Main processing
    with col1:
        st.subheader("Kết quả Phân tích Điểm danh")
        if st.button("📊 Bắt đầu Kiểm tra"):
            results, error = process_attendance(ngay, gio, file_paths)
            if error:
                st.error(error)
            else:
                st.session_state.processed_data = results
                danh_sach_di_muon = results.get("di_muon", [])
                danh_sach_vang_ban_dau = results.get("vang", [])
                
                display_results_table(danh_sach_di_muon, "Danh sách đi muộn")
                display_results_table(danh_sach_vang_ban_dau, "Danh sách vắng (trước khi lọc)")
                
                with st.spinner("Đang lọc danh sách người nghỉ phép..."):
                    danh_sach_vang_sau_loc = loai_bo_nguoi_nghi_phep(
                        danh_sach_vang=danh_sach_vang_ban_dau,
                        ten_file_leave_requests=file_paths["danh sách nghỉ phép"]
                    )
                
                removed_count = display_results_table(danh_sach_vang_sau_loc, "Danh sách vắng (sau khi lọc người nghỉ phép)")
                if removed_count > 0:
                    st.caption(f"Đã loại bỏ {len(danh_sach_vang_ban_dau) - removed_count} người có trong danh sách nghỉ phép.")
                
                st.session_state.processed_data = {
                    "di_muon": danh_sach_di_muon,
                    "vang_sau_loc": danh_sach_vang_sau_loc
                }
                
                if danh_sach_di_muon or danh_sach_vang_sau_loc:
                    with st.spinner("Đang tạo nội dung email..."):
                        emails_can_gui = tao_noi_dung_email(
                            danh_sach_vang=danh_sach_vang_sau_loc,
                            danh_sach_di_muon=danh_sach_di_muon,
                            ten_file_emails=file_paths["CSV emails"],
                            ten_file_mau=file_paths["mẫu Email"]
                        )
                        st.session_state.emails_can_gui = emails_can_gui
                        st.success(f"Đã tạo xong nội dung cho {len(emails_can_gui)} email.")
                else:
                    st.session_state.emails_can_gui = {}
                    st.info("Không có vi phạm nào cần tạo email.")
    
    # Email sending section
    with col2:
        st.subheader("Gửi Email Thông báo")
        
        # Add tabs for sending and viewing history
        tab1, tab2 = st.tabs(["Gửi Email", "Lịch sử"])
        
        with tab1:
            if 'emails_can_gui' in st.session_state and st.session_state.emails_can_gui:
                emails_to_send = st.session_state.emails_can_gui
                st.info(f"Tìm thấy {len(emails_to_send)} email đã được tạo sẵn.")
                
                with st.expander("Xem trước danh sách email sẽ gửi"):
                    for email_addr, content in emails_to_send.items():
                        st.markdown(f"**Tới:** {email_addr}")
                        st.markdown("---")
                
                if st.button("✉️ Gửi tất cả Email", key="send_email_button"):
                    with st.spinner("Đang gửi email... Vui lòng đợi."):
                        from dotenv import load_dotenv
                        load_dotenv()
                        
                        ket_qua_gui = gui_email(emails_to_send)
                    
                    st.subheader("Kết quả Gửi Email")
                    all_success = True
                    for email, trang_thai in ket_qua_gui.items():
                        if trang_thai == "Thành công":
                            st.success(f"{email}: {trang_thai}")
                        else:
                            st.error(f"{email}: {trang_thai}")
                            all_success = False
                    
                    # Add logging after sending emails
                    luu_log(
                        ngay_kiem_tra=ngay,
                        gio_so_sanh=gio.strftime('%H:%M'),
                        danh_sach_di_muon=st.session_state.processed_data["di_muon"],
                        danh_sach_vang=st.session_state.processed_data["vang_sau_loc"],
                        ket_qua_gui=ket_qua_gui
                    )
                    
                    if all_success:
                        st.balloons()
                    else:
                        st.warning("Một số email không gửi được. Vui lòng kiểm tra log lỗi.")
        
        # Add history tab
        with tab2:
            st.subheader("Lịch sử Gửi Email")
            if st.button("🔄 Làm mới", key="refresh_history"):
                pass  # The view_log_history function will be called anyway
            view_log_history()

if __name__ == "__main__":
    main()