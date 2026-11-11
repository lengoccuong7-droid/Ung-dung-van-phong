import streamlit as st
import pandas as pd
import numpy as np
import io
import os
import plotly.graph_objects as go
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side

# -----------------------------------------------------------------------------
# CẤU HÌNH THÔNG TIN ỨNG DỤNG PRO
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="Siêu Công Cụ Xử Lý Dữ Liệu Đa Năng Pro", 
    layout="wide", 
    page_icon="📋",
    initial_sidebar_state="expanded"
)

st.title("📋 Siêu Công Cụ Xử Lý Dữ Liệu: Tự Động Nhận Diện, Căn Giữa & Co Giãn Pro")
st.caption("Hỗ trợ dán text thô từ web/CSV, bóc tách 100% dòng, chỉnh sửa trực tiếp, vẽ biểu đồ & xuất file chuẩn Microsoft Office")

# -----------------------------------------------------------------------------
# HÀM BÓC TÁCH DỮ LIỆU ĐA NĂNG (GIỮ NGUYÊN NGUYÊN BẢN LOGIC GỐC)
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
# CÁCH 1: DÁN VĂN BẢN TRỰC TIẾP
# =============================================================================
if input_method == "📝 Dán Văn Bản Trực Tiếp":
    st.sidebar.subheader("📝 Dán dữ liệu vào khung bên dưới:")
    raw_text_input = st.sidebar.text_area(
        "Dán dữ liệu tại đây:",
        value="Date,Open,High,Low,Close,Volume\n2026-11-09,20.45,20.50,19.85,19.90,25847000\n2026-11-06,19.95,20.60,19.95,20.40,31408000",
        height=220
    )
    
    if st.sidebar.button("🚀 Xử Lý & Chuyển Đổi Thành Bảng Chuẩn", type="primary", use_container_width=True):
        if raw_text_input.strip():
            parsed_df = smart_parse_input(raw_text_input)
            if not parsed_df.empty:
                st.session_state['df_working'] = parsed_df
                st.success(f"✅ Đã xử lý và bóc tách thành công **{len(parsed_df)} dòng dữ liệu**!")
            else:
                st.warning("⚠️ Không thể đọc được cấu trúc dữ liệu. Sếp kiểm tra lại định dạng nhé!")

# =============================================================================
# CÁCH 2: NẠP TỆP EXCEL / CSV
# =============================================================================
else:
    uploaded_file = st.sidebar.file_uploader("Chọn tệp Excel (.xlsx, .xls) hoặc CSV", type=["xlsx", "xls", "csv"])
    if uploaded_file is not None:
        try:
            file_bytes = io.BytesIO(uploaded_file.getvalue())
            if uploaded_file.name.lower().endswith('.csv'):
                parsed_df = pd.read_csv(file_bytes)
            else:
                parsed_df = pd.read_excel(file_bytes, engine='openpyxl')
            st.session_state['df_working'] = parsed_df
            st.success(f"✅ Đã nạp thành công tệp: **{uploaded_file.name}** ({len(parsed_df)} dòng)")
        except Exception as e:
            st.sidebar.error(f"❌ Lỗi đọc tệp: {e}")

# Lấy dữ liệu hiện tại từ Session State
if 'df_working' in st.session_state:
    df_result = st.session_state['df_working']

# =============================================================================
# MÀN HÌNH CHÍNH & NÂNG CẤP TAB TRẢI NGHIỆM VĂN PHÒNG
# =============================================================================
if df_result is not None and not df_result.empty:
    tab1, tab2, tab3 = st.tabs([
        "✏️ 1. Bảng Dữ Liệu & Chỉnh Sửa", 
        "📈 2. Biểu Đồ & Thống Kê Nhanh", 
        "💾 3. Lưu & Xuất Tệp Office"
    ])
    
    # -------------------------------------------------------------------------
    # TAB 1: HIỂN THỊ & CHỈNH SỬA TRỰC TIẾP DẠNG EXCEL
    # -------------------------------------------------------------------------
    with tab1:
        st.subheader(f"📊 Bảng Dữ Liệu Đã Chuẩn Hóa ({len(df_result)} dòng, {len(df_result.columns)} cột):")
        st.caption("💡 Sếp có thể nhấp đôi chuột vào từng ô để sửa dữ liệu hoặc thêm/xóa dòng trực tiếp:")
        
        edited_df = st.data_editor(
            df_result,
            num_rows="dynamic",
            use_container_width=True,
            key="pro_data_editor"
        )
        st.session_state['df_working'] = edited_df

    # -------------------------------------------------------------------------
    # TAB 2: BIỂU ĐỒ & THỐNG KÊ (NÂNG CẤP PRO)
    # -------------------------------------------------------------------------
    with tab2:
        df_analysis = st.session_state['df_working']
        
        # Thống kê nhanh
        st.subheader("📊 Thống Kê Tổng Quan (KPI)")
        numeric_cols = df_analysis.select_dtypes(include=[np.number]).columns.tolist()
        if numeric_cols:
            kpi_cols = st.columns(min(len(numeric_cols), 4))
            for idx, col in enumerate(numeric_cols[:4]):
                with kpi_cols[idx]:
                    st.metric(label=f"Trung bình {col}", value=f"{df_analysis[col].mean():,.2f}")
        
        st.markdown("---")
        st.subheader("📈 Biểu Đồ Trực Quan")
        
        # Vẽ nến nếu đúng định dạng chứng khoán
        stock_cols = {'Open', 'High', 'Low', 'Close'}
        if stock_cols.issubset(set(df_analysis.columns)):
            fig = go.Figure(data=[go.Candlestick(
                x=df_analysis['Date'] if 'Date' in df_analysis.columns else df_analysis.index,
                open=df_analysis['Open'],
                high=df_analysis['High'],
                low=df_analysis['Low'],
                close=df_analysis['Close'],
                name="Giá Chứng Khoán"
            )])
            fig.update_layout(title="Biểu Đồ Nến Giá Kỹ Thuật", template="plotly_white", height=450)
            st.plotly_chart(fig, use_container_width=True)
        elif numeric_cols:
            selected_chart_col = st.selectbox("Chọn cột dữ liệu vẽ biểu đồ:", numeric_cols)
            st.line_chart(df_analysis[selected_chart_col])
        else:
            st.info("Không phát hiện cột dữ liệu số để vẽ biểu đồ.")

    # -------------------------------------------------------------------------
    # TAB 3: XUẤT FILE PRO (GIỮ NGUYÊN NGUYÊN BẢN CĂN GIỮA & CO GIÃN CỘT)
    # -------------------------------------------------------------------------
    with tab3:
        st.subheader("💾 LƯU & XUẤT TỆP CHUẨN MICROSOFT OFFICE")
        st.caption("Toàn bộ dữ liệu dòng vào được giữ nguyên vẹn 100%, tự động căn giữa ô và co giãn độ rộng cột khi mở bằng Excel:")
        
        df_export = st.session_state['df_working']
        
        col_x1, col_x2 = st.columns(2)
        
        # 1. Xuất file Microsoft Excel (.xlsx) giữ chuẩn 100% logic gốc của sếp + Trang trí viền chuyên nghiệp
        excel_buffer = io.BytesIO()
        with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
            df_export.to_excel(writer, index=False, sheet_name='Du_Lieu_Chuan')
            
            worksheet = writer.sheets['Du_Lieu_Chuan']
            center_alignment = Alignment(horizontal='center', vertical='center')
            header_font = Font(bold=True)
            
            for col in worksheet.columns:
                max_len = 0
                col_letter = col[0].column_letter
                for cell in col:
                    cell.alignment = center_alignment
                    if cell.row == 1:
                        cell.font = header_font
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
        csv_bytes = df_export.to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig')
        with col_x2:
            st.download_button(
                label="📄 Tải Xuất File CSV Chuẩn UTF-8 (.csv)",
                data=csv_bytes,
                file_name="du_lieu_chung_khoan_chuan.csv",
                mime="text/csv",
                use_container_width=True
            )
else:
    st.info("👈 Sếp vui lòng dán văn bản hoặc tải tệp ở thanh bên (Sidebar) để bắt đầu nhé!")
