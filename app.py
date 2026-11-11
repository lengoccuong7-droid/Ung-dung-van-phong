import streamlit as st
import pandas as pd
import io
import os
from openpyxl.styles import Alignment

# -----------------------------------------------------------------------------
# CẤU HÌNH THÔNG TIN ỨNG DỤNG
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="Siêu Công Cụ Xử Lý Dữ Liệu Đa Năng", 
    layout="wide", 
    page_icon="📋",
    initial_sidebar_state="expanded"
)

st.title("📋 Siêu Công Cụ Xử Lý Dữ Liệu: Cập Nhật, Chỉnh Sửa & Xuất File")
st.caption("Hỗ trợ mở file cũ, nhập thêm dữ liệu, chỉnh sửa trực tiếp, tự động căn giữa & cập nhật file mới")

# -----------------------------------------------------------------------------
# HÀM BÓC TÁCH DỮ LIỆU ĐA NĂNG (HỖ TRỢ CẢ TEXT KHOẢNG TRẮNG & CSV DẤU PHẨY)
# -----------------------------------------------------------------------------
def smart_parse_input(text_data):
    lines = [line.strip() for line in text_data.strip().split('\n') if line.strip()]
    if not lines:
        return pd.DataFrame()
    
    first_line = lines[0]
    if ',' in first_line and ('Date' in first_line or 'Open' in first_line):
        try:
            df_csv = pd.read_csv(io.StringIO(text_data))
            return df_csv
        except Exception:
            pass
            
    parsed_rows = []
    for line in lines:
        if 'Open' in line or 'Date' in line or 'Ngày' in line:
            continue
            
        parts = line.split()
        if len(parts) >= 6:
            date_str = parts[0]
            try:
                open_p = float(parts[1].replace(',', ''))
                high_p = float(parts[2].replace(',', ''))
                low_p = float(parts[3].replace(',', ''))
                close_p = float(parts[4].replace(',', ''))
                
                vol_str = parts[5].upper().replace(',', '')
                vol = 0.0
                if 'M' in vol_str:
                    vol = float(vol_str.replace('M', '')) * 1_000_000
                elif 'K' in vol_str:
                    vol = float(vol_str.replace('K', '')) * 1_000
                else:
                    vol = float(vol_str)
                    
                parsed_rows.append({
                    'Date': date_str,
                    'Open': open_p,
                    'High': high_p,
                    'Low': low_p,
                    'Close': close_p,
                    'Volume': vol
                })
            except Exception:
                continue
                
    return pd.DataFrame(parsed_rows)

# -----------------------------------------------------------------------------
# GIAO DIỆN THANH BÊN (SIDEBAR) - NẠP DỮ LIỆU
# -----------------------------------------------------------------------------
st.sidebar.header("⚙️ Phương Thức Nhập Dữ Liệu")
input_method = st.sidebar.radio(
    "Chọn cách nạp liệu:",
    ["📁 Mở File Excel / CSV Cũ Sẵn Có", "📝 Dán Văn Bản Trực Tiếp"]
)

df_result = None
original_filename = "du_lieu_cap_nhat"

# =============================================================================
# CÁCH 1: MỞ TỆP CŨ SẴN CÓ ĐỂ CẬP NHẬT/THÊM DỮ LIỆU
# =============================================================================
if input_method == "📁 Mở File Excel / CSV Cũ Sẵn Có":
    uploaded_file = st.sidebar.file_uploader("Chọn file Excel (.xlsx) hoặc CSV cũ của bạn:", type=["xlsx", "xls", "csv"])
    if uploaded_file is not None:
        try:
            file_bytes = io.BytesIO(uploaded_file.getvalue())
            if uploaded_file.name.lower().endswith('.csv'):
                df_result = pd.read_csv(file_bytes)
            else:
                df_result = pd.read_excel(file_bytes, engine='openpyxl')
            
            # Lấy tên file gốc (bỏ đuôi mở rộng) để gợi ý tên lúc lưu lại
            original_filename = os.path.splitext(uploaded_file.name)[0]
            st.sidebar.success(f"✅ Đã mở file: **{uploaded_file.name}** ({len(df_result)} dòng)")
        except Exception as e:
            st.sidebar.error(f"❌ Lỗi đọc tệp: {e}")

# =============================================================================
# CÁCH 2: DÁN VĂN BẢN TRỰC TIẾP
# =============================================================================
else:
    st.subheader("📝 Dán dữ liệu mới vào khung bên dưới:")
    raw_text_input = st.text_area(
        "Dán dữ liệu tại đây:",
        value="Date,Open,High,Low,Close,Volume\n2026-11-09,20.45,20.50,19.85,19.90,25847000",
        height=180
    )
    
    if st.button("🚀 Bóc Tách Dữ Liệu", type="primary", use_container_width=True):
        if raw_text_input.strip():
            df_result = smart_parse_input(raw_text_input)
            if not df_result.empty:
                st.success(f"✅ Đã bóc tách thành công **{len(df_result)} dòng dữ liệu**!")
            else:
                st.warning("⚠️ Không thể đọc được cấu trúc dữ liệu. Vui lòng kiểm tra lại định dạng!")

# =============================================================================
# KHU VỰC HIỂN THỊ, CHỈNH SỬA & LƯU CẬP NHẬT FILE
# =============================================================================
if df_result is not None and not df_result.empty:
    st.markdown("---")
    st.subheader(f"📊 Bảng Dữ Liệu Hiện Tại ({len(df_result)} dòng, {len(df_result.columns)} cột):")
    st.caption("💡 Bạn có thể nhấp đúp vào ô bất kỳ để chỉnh sửa, hoặc kéo xuống dòng cuối cùng để bấm thêm dòng dữ liệu mới.")
    
    # Sử dụng data_editor để cho phép người dùng thêm/sửa trực tiếp trên web
    edited_df = st.data_editor(
        df_result, 
        use_container_width=True, 
        num_rows="dynamic",
        key="data_editor_main"
    )
    
    st.markdown("---")
    st.subheader("💾 LƯU & CẬP NHẬT TỆP")
    
    # Ô nhập tên file (mặc định lấy theo tên file cũ hoặc tên gợi ý)
    custom_filename = st.text_input(
        "Đặt tên file khi lưu lại (Có thể giữ nguyên tên cũ để ghi đè):", 
        value=original_filename
    )
    
    # Định dạng tên file hoàn chỉnh
    excel_filename = f"{custom_filename.strip()}.xlsx" if custom_filename.strip() else "du_lieu_moi.xlsx"
    csv_filename = f"{custom_filename.strip()}.csv" if custom_filename.strip() else "du_lieu_moi.csv"

    # Xử lý tạo nội dung file Excel với định dạng căn giữa và co giãn cột
    excel_buffer = io.BytesIO()
    with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
        edited_df.to_excel(writer, index=False, sheet_name='Du_Lieu_Cap_Nhat')
        
        worksheet = writer.sheets['Du_Lieu_Cap_Nhat']
        center_alignment = Alignment(horizontal='center', vertical='center')
        
        for col in worksheet.columns:
            max_len = 0
            col_letter = col[0].column_letter
            for cell in col:
                cell.alignment = center_alignment
                try:
                    if cell.value is not None:
                        max_len = max(max_len, len(str(cell.value)))
                except:
                    pass
            worksheet.column_dimensions[col_letter].width = max(max_len + 6, 14)
            
    # Tạo nội dung file CSV
    csv_bytes = edited_df.to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig')

    col_x1, col_x2 = st.columns(2)
    
    with col_x1:
        st.download_button(
            label="📊 Tải Xuất File Excel Đã Cập Nhật",
            data=excel_buffer.getvalue(),
            file_name=excel_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            type="primary"
        )
        
    with col_x2:
        st.download_button(
            label="📄 Tải Xuất File CSV Đã Cập Nhật",
            data=csv_bytes,
            file_name=csv_filename,
            mime="text/csv",
            use_container_width=True
        )
else:
    st.info("👋 Hãy chọn file cũ ở thanh bên trái hoặc dán dữ liệu để bắt đầu làm việc nhé!")
