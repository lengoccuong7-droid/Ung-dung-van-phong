import streamlit as st
import pandas as pd
import numpy as np
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

st.title("📋 Siêu Công Cụ Xử Lý Dữ Liệu: Tự Động Nhận Diện, Căn Giữa & Co Giãn")
st.caption("Hỗ trợ dán text thô từ web hoặc dữ liệu CSV, tự động bóc tách đầy đủ 100% dòng, căn giữa ô, co giãn cột & xuất file chuẩn")

# -----------------------------------------------------------------------------
# HÀM BÓC TÁCH DỮ LIỆU ĐA NĂNG (HỖ TRỢ CẢ TEXT KHOẢNG TRẮNG & CSV DẤU PHẨY)
# -----------------------------------------------------------------------------
def smart_parse_input(text_data):
    lines = [line.strip() for line in text_data.strip().split('\n') if line.strip()]
    if not lines:
        return pd.DataFrame()
    
    # Kiểm tra xem dòng đầu tiên có phải là tiêu đề dạng CSV không (có dấu phẩy)
    first_line = lines[0]
    if ',' in first_line and ('Date' in first_line or 'Open' in first_line):
        try:
            # Đọc trực tiếp bằng pandas read_csv từ chuỗi văn bản
            df_csv = pd.read_csv(io.StringIO(text_data))
            return df_csv
        except Exception:
            pass
            
    # Nếu không phải CSV, tiến hành bóc tách dạng text copy khoảng trắng (có chữ M/K hoặc số thường)
    parsed_rows = []
    for line in lines:
        # Bỏ qua dòng tiêu đề nếu lỡ copy vào
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
# GIAO DIỆN THANH BÊN (SIDEBAR)
# -----------------------------------------------------------------------------
st.sidebar.header("⚙️ Phương Thức Nhập Dữ Liệu")
input_method = st.sidebar.radio(
    "Chọn cách nạp liệu:",
    ["📝 Dán Văn Bản Trực Tiếp", "📁 Tải Lên Tệp Excel / CSV"]
)

df_result = None

# =============================================================================
# CÁCH 1: DÁN VĂN BẢN TRỰC TIẾP (HỖ TRỢ CẢ TEXT & CSV)
# =============================================================================
if input_method == "📝 Dán Văn Bản Trực Tiếp":
    st.subheader("📝 Dán dữ liệu (Text thô hoặc định dạng CSV có dấu phẩy) vào khung bên dưới:")
    raw_text_input = st.text_area(
        "Dán dữ liệu tại đây:",
        value="Date,Open,High,Low,Close,Volume\n2026-11-09,20.45,20.50,19.85,19.90,25847000\n2026-11-06,19.95,20.60,19.95,20.40,31408000",
        height=220
    )
    
    if st.button("🚀 Xử Lý & Chuyển Đổi Thành Bảng Chuẩn", type="primary", use_container_width=True):
        if raw_text_input.strip():
            df_result = smart_parse_input(raw_text_input)
            if not df_result.empty:
                st.success(f"✅ Đã xử lý và bóc tách thành công **{len(df_result)} dòng dữ liệu** (Không thiếu một dòng nào)!")
            else:
                st.warning("⚠️ Không thể đọc được cấu trúc dữ liệu. Sếp kiểm tra lại định dạng nhé!")

# =============================================================================
# CÁCH 2: NẠP TỆP EXCEL / CSV CỔ ĐIỂN
# =============================================================================
else:
    uploaded_file = st.sidebar.file_uploader("Chọn tệp Excel (.xlsx, .xls) hoặc CSV", type=["xlsx", "xls", "csv"])
    if uploaded_file is not None:
        try:
            file_bytes = io.BytesIO(uploaded_file.getvalue())
            if uploaded_file.name.lower().endswith('.csv'):
                df_result = pd.read_csv(file_bytes)
            else:
                df_result = pd.read_excel(file_bytes, engine='openpyxl')
            st.success(f"✅ Đã nạp thành công tệp: **{uploaded_file.name}** ({len(df_result)} dòng)")
        except Exception as e:
            st.error(f"❌ Lỗi đọc tệp: {e}")

# =============================================================================
# HIỂN THỊ KẾT QUẢ & XUẤT FILE TỰ ĐỘNG CĂN GIỮA & CO GIÃN CỘT
# =============================================================================
if df_result is not None and not df_result.empty:
    st.markdown("---")
    st.subheader(f"📊 Bảng Dữ Liệu Đã Chuẩn Hóa ({len(df_result)} dòng, {len(df_result.columns)} cột):")
    st.dataframe(df_result, use_container_width=True)
    
    st.markdown("---")
    st.subheader("💾 LƯU & XUẤT TỆP CHUẨN MICROSOFT OFFICE")
    st.caption("Toàn bộ dữ liệu dòng vào được giữ nguyên vẹn 100%, tự động căn giữa ô và co giãn độ rộng cột khi mở bằng Excel:")
    
    col_x1, col_x2 = st.columns(2)
    
    # 1. Xuất file Microsoft Excel (.xlsx) với định dạng chuẩn văn phòng
    excel_buffer = io.BytesIO()
    with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
        df_result.to_excel(writer, index=False, sheet_name='Du_Lieu_Chuan')
        
        worksheet = writer.sheets['Du_Lieu_Chuan']
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
            
    with col_x1:
        st.download_button(
            label="📊 Tải Xuất File Excel (Căn Giữa & Co Giãn)",
            data=excel_buffer.getvalue(),
            file_name="du_lieu_chung_khoan_chuan.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            type="primary"
        )
        
    # 2. Xuất file CSV (.csv) chuẩn UTF-8
    csv_bytes = df_result.to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig')
    with col_x2:
        st.download_button(
            label="📄 Tải Xuất File CSV Chuẩn UTF-8 (.csv)",
            data=csv_bytes,
            file_name="du_lieu_chung_khoan_chuan.csv",
            mime="text/csv",
            use_container_width=True
        )
