import streamlit as st
import openpyxl
import os
from datetime import datetime
from logic import lay_danh_sach_sheet, xu_ly_tat_ca_ke_hoach

# 1. Cấu hình trang web
st.set_page_config(
    page_title="Hệ thống Quản lý Kế hoạch Sản xuất",
    page_icon="🏭",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- HÀM TÍNH TOÁN MẬT KHẨU ĐỘNG CHO USER THEO THÁNG ---
def get_dynamic_user_password():
    now = datetime.now()
    thang_str = now.strftime("%m") # Lấy tháng 2 chữ số (vd: 11, 12)
    nam_str = now.strftime("%Y")   # Lấy năm 4 chữ số (vd: 2026)
    return f"pkhwp{thang_str}{nam_str}01"

# Danh sách tài khoản cứng (Admin cố định, User tự động đổi theo tháng)
USERS_CREDENTIALS = {
    "admin": {
        "password": "AdminPassword123@#", # Bạn có thể đổi mật khẩu admin ở đây
        "role": "admin",
        "name": "Quản trị viên"
    },
    "PKHWP@gmail.com": {
        "password": get_dynamic_user_password(),
        "role": "user",
        "name": "Nhân viên Vận hành"
    }
}

# --- QUẢN LÝ TRẠNG THÁI ĐĂNG NHẬP ---
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False
    st.session_state.role = None
    st.session_state.username = None

# --- GIAO DIỆN ĐĂNG NHẬP (LOGIN SCREEN) ---
if not st.session_state.logged_in:
    st.markdown("<br><br>", unsafe_allow_html=True)
    col1, col2, col3 = st.columns([1, 1.2, 1])
    
    with col2:
        st.markdown("<h2 style='text-align: center;'>🔐 ĐĂNG NHẬP HỆ THỐNG</h2>", unsafe_allow_html=True)
        st.markdown("<p style='text-align: center; color: gray;'>Vui lòng đăng nhập để tiếp tục</p>", unsafe_allow_html=True)
        
        with st.form("login_form"):
            input_user = st.text_input("Tên tài khoản (Email / Username)")
            input_pass = st.text_input("Mật khẩu", type="password", autocomplete="new-password")
            submit_btn = st.form_submit_button("Đăng nhập", use_container_width=True)
            
            if submit_btn:
                # Kiểm tra thông tin đăng nhập
                if input_user in USERS_CREDENTIALS and USERS_CREDENTIALS[input_user]["password"] == input_pass:
                    st.session_state.logged_in = True
                    st.session_state.role = USERS_CREDENTIALS[input_user]["role"]
                    st.session_state.username = input_user
                    st.rerun()
                else:
                    st.error("❌ Sai tài khoản hoặc mật khẩu! Vui lòng kiểm tra lại.")
        
        # Gợi ý nhỏ cho user biết quy luật nếu cần
        with st.expander("ℹ️ Hỗ trợ đăng nhập"):
            st.caption(f"• **Tài khoản User (`PKHWP@gmail.com`)**: Mật khẩu thay đổi tự động vào ngày 01 hằng tháng .")

else:
    # --- GIAO DIỆN KHI ĐÃ ĐĂNG NHẬP THÀNH CÔNG ---
    
    # Custom CSS giao diện
    st.markdown("""
    <style>
        .main { background-color: #f8f9fa; }
        .main-title { color: #1e293b; font-weight: 700; margin-bottom: 8px; }
        .sub-title { color: #64748b; font-size: 15px; margin-bottom: 25px; }
        div.stButton > button:first-child {
            background-color: #2563eb; color: white; font-weight: 600;
            border-radius: 8px; padding: 10px 24px; border: none; width: 100%;
        }
        div.stButton > button:first-child:hover { background-color: #1d4ed8; }
    </style>
    """, unsafe_allow_html=True)

    # Sidebar
    with st.sidebar:
        st.image("https://img.icons8.com/fluency/96/factory.png", width=64)
        st.title("XLKH System")
        st.write(f"👤 Xin chào: **{st.session_state.username}**")
        st.caption(f"Phân quyền: `{st.session_state.role.upper()}`")
        
        st.markdown("---")
        
        # Nếu là Admin thì có thêm bảng hướng dẫn mật khẩu tháng hiện tại của User để tiện tra cứu
        if st.session_state.role == "admin":
            st.markdown("### 🛠️ Khu vực Admin")
            st.info(f"🔑 Mật khẩu User tháng này (`{datetime.now().strftime('%m/%Y')}`): \n\n`{get_dynamic_user_password()}`")
            st.markdown("---")
            
        if st.button("🚪 Đăng xuất"):
            st.session_state.logged_in = False
            st.session_state.role = None
            st.session_state.username = None
            st.rerun()

    # Main Content
    st.markdown('<h1 class="main-title">🏭 QUẢN LÝ KẾ HOẠCH SẢN XUẤT</h1>', unsafe_allow_html=True)
    st.markdown('<p class="sub-title">Công cụ chuyển đổi & lập lịch tự động cho dây chuyền Ép vỉ và Pha chế</p>', unsafe_allow_html=True)

    # Khung Upload file
    st.markdown('### 📂 1. Tải file Excel kế hoạch')
    uploaded_file = st.file_uploader("Kéo thả hoặc bấm để chọn file .xlsx / .xls", type=["xlsx", "xls"])

    if uploaded_file is not None:
        temp_file_path = "temp_" + uploaded_file.name
        with open(temp_file_path, "wb") as f:
            f.write(uploaded_file.getbuffer())
        
        st.success(f"✓ Đã nhận file: **{uploaded_file.name}**")
        
        try:
            sheets = lay_danh_sach_sheet(temp_file_path)
            
            st.markdown('---')
            st.markdown('### ⚙️ 2. Cấu hình thông số xử lý')
            
            col1, col2 = st.columns(2)
            with col1:
                selected_sheet = st.selectbox("📋 Chọn Sheet nguồn:", sheets)
            with col2:
                selected_month = st.selectbox("📅 Chọn Tháng xử lý:", [str(i) for i in range(1, 13)], index=9)
                
            st.markdown('#### 🎯 Tùy chọn Sheet đầu ra:')
            opt_col1, opt_col2, opt_col3 = st.columns(3)
            with opt_col1:
                export_ep_vi = st.checkbox("Sheet Ép vỉ", value=True)
            with opt_col2:
                export_pha_che = st.checkbox("Sheet Pha chế", value=True)
            with opt_col3:
                export_lich = st.checkbox("Sheet Lịch SP (Tổng hợp)", value=True)
                
            st.markdown('---')
            
            if st.button("🚀 BẮT ĐẦU XỬ LÝ DỮ LIỆU"):
                if not (export_ep_vi or export_pha_che or export_lich):
                    st.warning("⚠️ Vui lòng chọn ít nhất 1 sheet kết quả để xuất!")
                else:
                    progress_bar = st.progress(0)
                    status_text = st.empty()
                    
                    def update_progress(percent, msg):
                        progress_bar.progress(percent)
                        status_text.text(f"⏳ {msg}")
                    
                    output_file = xu_ly_tat_ca_ke_hoach(
                        temp_file_path,
                        int(selected_month),
                        selected_sheet,
                        export_ep_vi=export_ep_vi,
                        export_pha_che=export_pha_che,
                        export_lich=export_lich,
                        progress_callback=update_progress
                    )
                    
                    status_text.empty()
                    progress_bar.empty()
                    st.balloons()
                    
                    st.success("🎉 **Xử lý hoàn tất thành công!**")
                    
                    with open(output_file, "rb") as f:
                        st.download_button(
                            label="📥 TẢI FILE KẾT QUẢ SẠCH (.XLSX)",
                            data=f,
                            file_name=os.path.basename(output_file),
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            type="primary"
                        )
                        
        except Exception as e:
            st.error(f"❌ Có lỗi xảy ra trong quá trình xử lý: {str(e)}")