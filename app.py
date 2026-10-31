import streamlit as st
import openpyxl
import os
from logic import lay_danh_sach_sheet, xu_ly_tat_ca_ke_hoach

# 1. Cấu hình trang web & Inject CSS tùy biến giao diện
st.set_page_config(
    page_title="Hệ thống Quản lý Kế hoạch Sản xuất",
    page_icon="🏭",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS giao diện hiện đại
st.markdown("""
<style>
    /* Styling chính cho ứng dụng */
    .main {
        background-color: #f8f9fa;
    }
    .stAppHeader {
        background-color: rgba(255, 255, 255, 0);
    }
    .css-card {
        background-color: #ffffff;
        padding: 24px;
        border-radius: 12px;
        box-shadow: 0 4px 12px rgba(0,0,0,0.05);
        margin-bottom: 20px;
        border: 1px solid #e9ecef;
    }
    .main-title {
        color: #1e293b;
        font-weight: 700;
        margin-bottom: 8px;
    }
    .sub-title {
        color: #64748b;
        font-size: 15px;
        margin-bottom: 25px;
    }
    /* Style button chính */
    div.stButton > button:first-child {
        background-color: #2563eb;
        color: white;
        font-weight: 600;
        border-radius: 8px;
        padding: 10px 24px;
        border: none;
        width: 100%;
        transition: all 0.3s ease;
    }
    div.stButton > button:first-child:hover {
        background-color: #1d4ed8;
        box-shadow: 0 4px 12px rgba(37, 99, 235, 0.3);
    }
</style>
""", unsafe_allow_html=True)

# 2. Sidebar hướng dẫn & thông tin
with st.sidebar:
    st.image("https://img.icons8.com/fluency/96/factory.png", width=64)
    st.title("XLKH System")
    st.caption("Phiên bản 2.0 Web Edition")
    st.markdown("---")
    st.markdown("### 📌 Hướng dẫn sử dụng")
    st.markdown("""
    1. **Tải lên** file Excel kế hoạch gốc.
    2. **Chọn Sheet** và **Tháng** cần xử lý.
    3. **Tích chọn** các thông tin muốn xuất.
    4. Nhấn **Bắt đầu xử lý** và tải file kết quả.
    """)
    st.markdown("---")
    st.info("💡 File xuất ra đã được làm sạch metadata, không bị lỗi Excel Repair.")

# 3. Khu vực chính (Main Dashboard)
st.markdown('<h1 class="main-title">🏭 QUẢN LÝ KẾ HOẠCH SẢN XUẤT</h1>', unsafe_allow_html=True)
st.markdown('<p class="sub-title">Công cụ chuyển đổi & lập lịch tự động cho dây chuyền Ép vỉ và Pha chế</p>', unsafe_allow_html=True)

# Khung Upload file
st.markdown('### 📂 1. Tải file Excel kế hoạch')
uploaded_file = st.file_uploader("Kéo thả hoặc bấm để chọn file .xlsx / .xls", type=["xlsx", "xls"])

if uploaded_file is not None:
    # Lưu tạm file upload
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
        
        # Nút thực thi
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
                st.balloons() # Hiệu ứng chúc mừng
                
                st.success("🎉 **Xử lý hoàn tất thành công!**")
                
                # Nút tải file
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