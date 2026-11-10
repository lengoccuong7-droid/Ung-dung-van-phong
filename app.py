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
    page_title="Công Cụ Chuyển Đổi Dữ Liệu Tự Động Căn Giữa & Co Giãn", 
    layout="wide", 
    page_icon="📋",
    initial_sidebar_state="expanded"
)

st.title("📋 Siêu Công Cụ Xử Lý Dữ Liệu: Tự Động Căn Giữa, Co Giãn Cột & Xuất File")
st.caption("Dán dữ liệu copy từ web, tự động bóc tách cột, quy đổi đơn vị M/K, tự động căn giữa ô, co giãn kích thước & xuất file Microsoft Excel/CSV")

# -----------------------------------------------------------------------------
# HÀM BÓC TÁCH DỮ LIỆU COPY TRỰC TIẾP TỪ WEB/MẠNG
# -----------------------------------------------------------------------------
def parse_copied_text(text_data):
    lines = text_data.strip().split('\n')
    parsed_rows = []
    
    for line in lines:
        parts = line.strip().split()
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
                    'Ngày': date_str,
                    'Giá Mở': open_p,
                    'Giá Cao': high_p,
                    'Giá Thấp': low_p,
                    'Giá Đóng': close_p,
                    'Khối Lượng': vol
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
    ["📋 Dán Text Copy Từ Mạng", "📁 Tải Lên Tệp Excel / CSV"]
)

df_result = None

# =============================================================================
# CÁCH 1: DÁN TEXT TRỰC TIẾP TỪ MẠNG
# =============================================================================
if input_method == "📋 Dán Text Copy Từ Mạng":
    st.subheader("📝 Dán dữ liệu số liệu copy từ web vào khung bên dưới:")
    raw_text_input = st.text_area(
        "Dán dữ liệu tại đây (Mỗi dòng là một phiên giao dịch):",
        value="23/10 20.10 20.15 19.90 19.95 20.683M\n22/10 20.45 20.60 20.15 20.15 19.083M\n21/10 20.05 20.45 20.00 20.40 32.392M",
        height=200
    )
    
    if st.button("🚀 Xử Lý & Chuyển Đổi Thành Bảng Chuẩn", type="primary", use_container_width=True):
        if raw_text_input.strip():
            df_result = parse_copied_text(raw_text_input)
            if not df_result.empty:
                st.success(f"✅ Đã bóc tách thành công **{len(df_result)} dòng dữ liệu**!")
            else:
                st.warning("⚠️ Không thể đọc được cấu trúc văn bản. Sếp hãy kiểm tra lại định dạng dữ liệu dán vào nhé!")

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
            st.success(f"✅ Đã nạp thành công tệp: {uploaded_file.name}")
        except Exception as e:
            st.error(f"❌ Lỗi đọc tệp: {e}")

# =============================================================================
# HIỂN THỊ KẾT QUẢ & XUẤT FILE TỰ ĐỘNG CĂN GIỮA & CO GIÃN CỘT
# =============================================================================
if df_result is not None and not df_result.empty:
    st.markdown("---")
    st.subheader("📊 Bảng Dữ Liệu Đã Được Chuẩn Hóa:")
    st.dataframe(df_result, use_container_width=True)
    
    st.markdown("---")
    st.subheader("💾 LƯU & XUẤT TỆP CHUẨN MICROSOFT OFFICE (CĂN GIỮA & CO GIÃN)")
    st.caption("Khi mở file Excel tải về, toàn bộ dữ liệu trong ô sẽ tự động căn giữa và co giãn vừa khít nội dung:")
    
    col_x1, col_x2 = st.columns(2)
    
    # 1. Xuất file Microsoft Excel (.xlsx) với tính năng TỰ ĐỘNG CĂN GIỮA & CO GIÃN CỘT
    excel_buffer = io.BytesIO()
    with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
        df_result.to_excel(writer, index=False, sheet_name='Du_Lieu_Chuan')
        
        worksheet = writer.sheets['Du_Lieu_Chuan']
        
        # Thiết lập tự động canh giữa (Center Alignment) cho toàn bộ ô và co giãn cột
        center_alignment = Alignment(horizontal='center', vertical='center')
        
        for col in worksheet.columns:
            max_len = 0
            col_letter = col[0].column_letter
            for cell in col:
                # Gán định dạng căn giữa cho từng ô
                cell.alignment = center_alignment
                try:
                    if cell.value:
                        max_len = max(max_len, len(str(cell.value)))
                except:
                    pass
            # Đặt độ rộng cột vừa khít dữ liệu
            worksheet.column_dimensions[col_letter].width = max(max_len + 6, 14)
            
    with col_x1:
        st.download_button(
            label="📊 Tải Xuất File Excel (Tự Động Căn Giữa & Co Giãn)",
            data=excel_buffer.getvalue(),
            file_name="du_lieu_can_giua_co_gian.xlsx",
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
            file_name="du_lieu_can_giua_co_gian.csv",
            mime="text/csv",
            use_container_width=True
        )
