import streamlit as st
import pandas as pd
import numpy as np
import io
import os
import plotly.graph_objects as go
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side

# -----------------------------------------------------------------------------
# CẤU HÌNH THÔNG TIN ỨNG DỤNG VĂN PHÒNG TOÀN DIỆN
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="Office Suite Pro - Trạm Làm Việc Văn Phòng Đa Năng", 
    layout="wide", 
    page_icon="🏢",
    initial_sidebar_state="expanded"
)

# Custom CSS giao diện phong cách Văn Phòng Hiện Đại (Office Suite)
st.markdown("""
    <style>
    .stApp { background-color: #f4f6f9; }
    .sidebar .stRadio { background-color: #ffffff; padding: 10px; border-radius: 8px; }
    div.stButton > button:first-child { border-radius: 6px; font-weight: 600; }
    </style>
""", unsafe_allow_html=True)

st.title("🏢 Office Suite Pro: Trạm Làm Việc Văn Phòng Đa Năng")
st.caption("Tích hợp Xử lý Dữ liệu Thông minh (Excel/CSV) & Trình Soạn Thảo Văn Bản Chuyên Nghiệp")

# -----------------------------------------------------------------------------
# HÀM BÓC TÁCH DỮ LIỆU ĐA NĂNG (GIỮ NGUYÊN 100% LOGIC GỐC CỦA SẾP)
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
# GIAO DIỆN THANH BÊN (SIDEBAR) - LỰA CHỌN MODULE VĂN PHÒNG
# -----------------------------------------------------------------------------
st.sidebar.header("🗂️ Phân Hệ Văn Phòng")
app_module = st.sidebar.radio(
    "Chọn phân hệ làm việc:",
    ["📊 Phân Hệ 1: Xử Lý Dữ Liệu & Bảng Biểu (Data Studio)", "📝 Phân Hệ 2: Soạn Thảo & Quản Lý Văn Bản (Word Processor)"]
)

st.sidebar.markdown("---")

# =============================================================================
# PHÂN HỆ 1: XỬ LÝ DỮ LIỆU & BẢNG BIỂU (DỰA TRÊN CODE GỐC CỦA SẾP)
# =============================================================================
if app_module.startswith("📊 Phân Hệ 1"):
    st.sidebar.header("⚙️ Nạp & Xử Lý Dữ Liệu")
    input_method = st.sidebar.radio(
        "Chọn cách nạp liệu:",
        ["📝 Dán Văn Bản Trực Tiếp", "📁 Tải Lên Tệp Excel / CSV"]
    )

    df_result = None

    if input_method == "📝 Dán Văn Bản Trực Tiếp":
        st.sidebar.subheader("📝 Dán dữ liệu vào khung:")
        raw_text_input = st.sidebar.text_area(
            "Dán dữ liệu tại đây:",
            value="Date,Open,High,Low,Close,Volume\n2026-11-09,20.45,20.50,19.85,19.90,25847000\n2026-11-06,19.95,20.60,19.95,20.40,31408000",
            height=200
        )
        
        if st.sidebar.button("🚀 Xử Lý & Chuyển Đổi Thành Bảng", type="primary", use_container_width=True):
            if raw_text_input.strip():
                parsed_df = smart_parse_input(raw_text_input)
                if not parsed_df.empty:
                    st.session_state['df_working'] = parsed_df
                    st.success(f"✅ Bóc tách thành công **{len(parsed_df)} dòng dữ liệu**!")
                else:
                    st.warning("⚠️ Không thể đọc cấu trúc dữ liệu. Kiểm tra lại định dạng nhé!")
    else:
        uploaded_file = st.sidebar.file_uploader("Chọn tệp Excel / CSV", type=["xlsx", "xls", "csv"])
        if uploaded_file is not None:
            try:
                file_bytes = io.BytesIO(uploaded_file.getvalue())
                if uploaded_file.name.lower().endswith('.csv'):
                    parsed_df = pd.read_csv(file_bytes)
                else:
                    parsed_df = pd.read_excel(file_bytes, engine='openpyxl')
                st.session_state['df_working'] = parsed_df
                st.success(f"✅ Nạp tệp thành công: **{uploaded_file.name}** ({len(parsed_df)} dòng)")
            except Exception as e:
                st.sidebar.error(f"❌ Lỗi đọc tệp: {e}")

    if 'df_working' in st.session_state:
        df_result = st.session_state['df_working']

    if df_result is not None and not df_result.empty:
        # Chia Tab cho Phân hệ dữ liệu
        d_tab1, d_tab2, d_tab3 = st.tabs(["✏️ Chỉnh Sửa Dữ Liệu Grid", "📈 Thống Kê & Biểu Đồ", "💾 Xuất Tệp Văn Phòng"])
        
        with d_tab1:
            st.subheader(f"📊 Bảng Dữ Liệu ({len(df_result)} dòng, {len(df_result.columns)} cột):")
            st.caption("💡 Sếp có thể nhấp đôi chuột vào ô để sửa trực tiếp hoặc thêm/xóa dòng tương tự Google Sheets:")
            edited_df = st.data_editor(df_result, num_rows="dynamic", use_container_width=True, key="data_grid_editor")
            st.session_state['df_working'] = edited_df

        with d_tab2:
            st.subheader("📊 Phân Tích & Biểu Đồ Kỹ Thuật")
            df_ana = st.session_state['df_working']
            numeric_cols = df_ana.select_dtypes(include=[np.number]).columns.tolist()
            
            if numeric_cols:
                col_metric = st.columns(min(len(numeric_cols), 4))
                for idx, col in enumerate(numeric_cols[:4]):
                    with col_metric[idx]:
                        st.metric(label=f"TB {col}", value=f"{df_ana[col].mean():,.2f}")
            
            stock_cols = {'Open', 'High', 'Low', 'Close'}
            if stock_cols.issubset(set(df_ana.columns)):
                fig = go.Figure(data=[go.Candlestick(
                    x=df_ana['Date'] if 'Date' in df_ana.columns else df_ana.index,
                    open=df_ana['Open'], high=df_ana['High'], low=df_ana['Low'], close=df_ana['Close'],
                    name="Nến Giá"
                )])
                fig.update_layout(title="Biểu Đồ Nến Giá Chứng Khoán", template="plotly_white", height=450)
                st.plotly_chart(fig, use_container_width=True)
            elif numeric_cols:
                sel_col = st.selectbox("Chọn cột vẽ biểu đồ đường:", numeric_cols)
                st.line_chart(df_ana[sel_col])

        with d_tab3:
            st.subheader("💾 LƯU & XUẤT TỆP CHUẨN MICROSOFT OFFICE")
            st.caption("Tự động căn giữa ô, đóng khung viền & co giãn độ rộng cột khi mở bằng Excel:")
            
            df_export = st.session_state['df_working']
            col_x1, col_x2 = st.columns(2)
            
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
                    file_name="du_lieu_chuan_van_phong.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    type="primary"
                )
                
            csv_bytes = df_export.to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig')
            with col_x2:
                st.download_button(
                    label="📄 Tải Xuất File CSV Chuẩn UTF-8 (.csv)",
                    data=csv_bytes,
                    file_name="du_lieu_chuan_van_phong.csv",
                    mime="text/csv",
                    use_container_width=True
                )
    else:
        st.info("👈 Sếp vui lòng dán văn bản hoặc tải tệp lên ở thanh bên (Sidebar) để bắt đầu phân tích dữ liệu nhé!")

# =============================================================================
# PHÂN HỆ 2: SOẠN THẢO & QUẢN LÝ VĂN BẢN (WORD PROCESSOR)
# =============================================================================
else:
    st.subheader("📝 Phân Hệ Soạn Thảo Văn Bản & Báo Cáo Hành Chính")
    st.caption("Công cụ soạn thảo nội dung văn bản, biên bản, hợp đồng trực tiếp trên ứng dụng văn phòng.")
    
    # Khởi tạo bộ nhớ tạm cho văn bản trong session state
    if 'doc_title' not in st.session_state:
        st.session_state['doc_title'] = "Báo Cáo Công Tác Tuần"
    if 'doc_content' not in st.session_state:
        st.session_state['doc_content'] = "Kính gửi Ban Giám Đốc,\n\nBáo cáo tiến độ triển khai dự án tuần qua như sau:\n1. Hoàn thành xử lý dữ liệu hệ thống.\n2. Cập nhật biểu đồ báo cáo tài chính.\n\nTrân trọng,"

    w_tab1, w_tab2 = st.tabs(["✍️ Soạn Thảo & Định Dạng", "📋 Thư Viện Mẫu Văn Bản Nhanh"])
    
    with w_tab1:
        st.session_state['doc_title'] = st.text_input("Tiêu đề tài liệu:", value=st.session_state['doc_title'])
        st.session_state['doc_content'] = st.text_area("Nội dung chi tiết văn bản:", value=st.session_state['doc_content'], height=300)
        
        # Thống kê nhanh chỉ số văn bản
        words_count = len(st.session_state['doc_content'].split())
        chars_count = len(st.session_state['doc_content'])
        reading_time = round(words_count / 200, 1) # ~200 từ/phút
        
        col_m1, col_m2, col_m3 = st.columns(3)
        col_m1.metric("Tổng số từ", f"{words_count} từ")
        col_m2.metric("Tổng số ký tự", f"{chars_count} ký tự")
        col_m3.metric("Thời gian đọc ước tính", f"{reading_time} phút")
        
        st.markdown("---")
        st.subheader("💾 Tải Xuống Tài Liệu Văn Bản")
        
        # Xuất file text / markdown mô phỏng tài liệu word
        doc_full_text = f"# {st.session_state['doc_title']}\n\n{st.session_state['doc_content']}"
        
        col_w1, col_w2 = st.columns(2)
        with col_w1:
            st.download_button(
                label="📄 Tải Xuống File Văn Bản (.txt)",
                data=doc_full_text.encode('utf-8-sig'),
                file_name=f"{st.session_state['doc_title'].replace(' ', '_')}.txt",
                mime="text/plain",
                use_container_width=True,
                type="primary"
            )
        with col_w2:
            st.download_button(
                label="📝 Tải Xuống Báo Cáo Markdown (.md)",
                data=doc_full_text.encode('utf-8-sig'),
                file_name=f"{st.session_state['doc_title'].replace(' ', '_')}.md",
                mime="text/markdown",
                use_container_width=True
            )

    with w_tab2:
        st.subheader("📋 Mẫu Văn Bản Hành Chính Sẵn Sàng Sử Dụng")
        st.caption("Chọn nhanh các biểu mẫu chuẩn văn phòng để tiết kiệm thời gian soạn thảo:")
        
        template_choice = st.selectbox(
            "Chọn biểu mẫu:",
            ["Biên Bản Cuộc Họp", "Báo Cáo Công Việc Tuần", "Quyết Định Hành Chính"]
        )
        
        if st.button("📥 Áp Dụng Mẫu Này Vào Trình Soạn Thảo"):
            if template_choice == "Biên Bản Cuộc Họp":
                st.session_state['doc_title'] = "Biên Bản Cuộc Họp Định Kỳ"
                st.session_state['doc_content'] = "1. Thời gian: ...\n2. Thành phần tham dự: ...\n3. Nội dung thảo luận:\n- Vấn đề 1: ...\n- Vấn đề 2: ...\n4. Kết luận cuộc họp:"
            elif template_choice == "Báo Cáo Công Việc Tuần":
                st.session_state['doc_title'] = "Báo Cáo Công Tác Tuần"
                st.session_state['doc_content'] = "- Các công việc đã thực hiện trong tuần:\n- Các vấn đề tồn đọng:\n- Kế hoạch công tác tuần tới:"
            else:
                st.session_state['doc_title'] = "Quyết Định V/v Triển Khai Công Việc"
                st.session_state['doc_content'] = "Căn cứ vào chức năng nhiệm vụ...\nNay quyết định:\nĐiều 1: ...\nĐiều 2: Các bộ phận liên quan thi hành quyết định này."
            st.success(f"✅ Đã tải mẫu '{template_choice}' vào Phân hệ soạn thảo văn bản thành công! Sếp bấm qua tab 'Soạn Thảo & Định Dạng' để xem nhé.")
