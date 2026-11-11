import streamlit as st
import pandas as pd
import numpy as np
import io
import os
import plotly.graph_objects as go
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side

# -----------------------------------------------------------------------------
# 1. CẤU HÌNH GIAO DIỆN HỆ THỐNG ALL-IN-ONE OFFICE
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="All-In-One Office Suite Pro", 
    layout="wide", 
    page_icon="🏢",
    initial_sidebar_state="expanded"
)

# Custom CSS cho phong cách Microsoft Ribbon / Office Workspace
st.markdown("""
    <style>
    .stApp { background-color: #f3f4f6; }
    .main .block-container { padding-top: 1.5rem; padding-bottom: 2rem; }
    div[data-testid="stSidebarNav"] { font-weight: bold; }
    .stTabs [data-baseweb="tab-list"] { gap: 8px; }
    .stTabs [data-baseweb="tab"] {
        background-color: #ffffff;
        border-radius: 6px 6px 0 0;
        padding: 8px 16px;
        font-weight: 600;
        border: 1px solid #e5e7eb;
    }
    .stTabs [aria-selected="true"] {
        background-color: #1e40af !important;
        color: white !important;
    }
    </style>
""", unsafe_allow_html=True)

st.title("🏢 All-In-One Office Suite Pro")
st.caption("Hệ sinh thái văn phòng thế hệ mới: Word, Excel & Quản lý Tài liệu trên một nền tảng")

# -----------------------------------------------------------------------------
# 2. HÀM BÓC TÁCH DỮ LIỆU BẢNG TÍNH THÔNG MINH (EXCEL CORE ENGINE)
# -----------------------------------------------------------------------------
def smart_parse_input(text_data):
    lines = [line.strip() for line in text_data.strip().split('\n') if line.strip()]
    if not lines:
        return pd.DataFrame()
    
    # Ưu tiên bóc tách định dạng CSV
    first_line = lines[0]
    if ',' in first_line and ('Date' in first_line or 'Open' in first_line):
        try:
            return pd.read_csv(io.StringIO(text_data))
        except Exception:
            pass
            
    # Bóc tách dữ liệu dạng Text phân cách bằng khoảng trắng
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
# 3. MẠNG ĐIỀU HƯỚNG SIDEBAR (PHÂN HỆ VĂN PHÒNG)
# -----------------------------------------------------------------------------
st.sidebar.header("🗂️ Phân Hệ Văn Phòng")
app_module = st.sidebar.radio(
    "Chọn công cụ làm việc:",
    [
        "📝 Phân Hệ Word (Soạn Thảo Văn Bản)", 
        "📊 Phân Hệ Excel (Bảng Tính & Phân Tích)", 
        "📁 Phân Hệ Quản Lý Tệp (Document Center)"
    ]
)

st.sidebar.markdown("---")
st.sidebar.info("💡 **Trạng thái:** Hệ thống đang kết nối trực tiếp bộ nhớ Cục bộ / Session State.")

# =============================================================================
# PHÂN HỆ 1: WORD - SOẠN THẢO VĂN BẢN CHUYÊN NGHIỆP
# =============================================================================
if app_module.startswith("📝 Phân Hệ Word"):
    st.subheader("📝 Word Processor - Soạn Thảo & Quản Lý Văn Bản")
    st.caption("Tạo báo cáo, biên bản cuộc họp và xuất tài liệu đa định dạng.")

    # Khởi tạo dữ liệu Word trong session
    if 'word_doc_title' not in st.session_state:
        st.session_state['word_doc_title'] = "BÁO CÁO CÔNG TÁC TUẦN"
    if 'word_doc_content' not in st.session_state:
        st.session_state['word_doc_content'] = (
            "Kính gửi: Ban Giám Đốc,\n\n"
            "Báo cáo tiến độ triển khai dự án tuần qua như sau:\n"
            "1. Hoàn thành nâng cấp giao diện All-In-One Office trên nền tảng Python.\n"
            "2. Đã tích hợp bộ bóc tách dữ liệu và vẽ biểu đồ kỹ thuật tự động.\n"
            "3. Chuẩn hóa quy trình xuất file Excel căn giữa và co giãn độ rộng cột.\n\n"
            "Kế hoạch công tác tuần tới:\n"
            "- Kiểm thử hiệu năng hệ thống với tập dữ liệu lớn.\n"
            "- Tích hợp bổ sung trợ lý AI Writer.\n\n"
            "Trân trọng,"
        )

    w_tab1, w_tab2 = st.tabs(["✍️ Khung Soạn Thảo & Định Dạng", "📋 Thư Viện Mẫu Văn Bản"])

    with w_tab1:
        st.session_state['word_doc_title'] = st.text_input("Tiêu đề tài liệu:", value=st.session_state['word_doc_title'])
        st.session_state['word_doc_content'] = st.text_area(
            "Nội dung văn bản (Hỗ trợ Markdown / Plaintext):", 
            value=st.session_state['word_doc_content'], 
            height=320
        )

        # Phân tích các chỉ số tài liệu
        content_text = st.session_state['word_doc_content']
        words_count = len(content_text.split())
        chars_count = len(content_text)
        read_time = round(words_count / 200, 1) if words_count > 0 else 0

        st.markdown("---")
        m_col1, m_col2, m_col3 = st.columns(3)
        m_col1.metric("Số từ", f"{words_count} từ")
        m_col2.metric("Số ký tự", f"{chars_count} ký tự")
        m_col3.metric("Thời gian đọc ước tính", f"{read_time} phút")

        st.markdown("---")
        st.subheader("💾 Xuất Tài Liệu")
        full_word_document = f"# {st.session_state['word_doc_title']}\n\n{st.session_state['word_doc_content']}"

        col_w1, col_w2 = st.columns(2)
        with col_w1:
            st.download_button(
                label="📄 Tải Xuất Văn Bản (.txt)",
                data=full_word_document.encode('utf-8-sig'),
                file_name=f"{st.session_state['word_doc_title'].replace(' ', '_')}.txt",
                mime="text/plain",
                use_container_width=True,
                type="primary"
            )
        with col_w2:
            st.download_button(
                label="📝 Tải Xuất File Markdown (.md)",
                data=full_word_document.encode('utf-8-sig'),
                file_name=f"{st.session_state['word_doc_title'].replace(' ', '_')}.md",
                mime="text/markdown",
                use_container_width=True
            )

    with w_tab2:
        st.subheader("📋 Chọn Mẫu Văn Bản Chuẩn Hành Chính")
        template_choice = st.selectbox(
            "Danh sách biểu mẫu:",
            ["Biên Bản Cuộc Họp", "Báo Cáo Công Việc Tuần", "Quyết Định Hành Chính"]
        )
        if st.button("📥 Nạp Mẫu Này Vào Trình Soạn Thảo"):
            if template_choice == "Biên Bản Cuộc Họp":
                st.session_state['word_doc_title'] = "BIÊN BẢN CUỘC HỌP ĐỊNH KỲ"
                st.session_state['word_doc_content'] = "1. Thời gian: 08:00 AM\n2. Thành phần tham dự: Ban Giám Đốc và các phòng ban.\n3. Nội dung thảo luận:\n- Tiến độ dự án: ...\n- Phân bổ nhân sự: ...\n4. Kết luận cuộc họp:"
            elif template_choice == "Báo Cáo Công Việc Tuần":
                st.session_state['word_doc_title'] = "BÁO CÁO CÔNG TÁC TUẦN"
                st.session_state['word_doc_content'] = "Các công việc hoàn thành:\n- ...\nVấn đề tồn đọng:\n- ...\nKế hoạch tuần tiếp theo:"
            else:
                st.session_state['word_doc_title'] = "QUYẾT ĐỊNH BAN HÀNH QUY CHẾ VĂN PHÒNG"
                st.session_state['word_doc_content'] = "Căn cứ vào thẩm quyền...\nQuyết định:\nĐiều 1: Ban hành quy chế làm việc mới.\nĐiều 2: Quyết định có hiệu lực từ ngày ký."
            st.success(f"✅ Đã tải mẫu '{template_choice}' vào trình soạn thảo!")

# =============================================================================
# PHÂN HỆ 2: EXCEL - BẢNG TÍNH & PHÂN TÍCH DỮ LIỆU
# =============================================================================
elif app_module.startswith("📊 Phân Hệ Excel"):
    st.subheader("📊 Excel Studio - Bảng Tính & Phân Tích Thông Minh")
    st.caption("Nhận diện dữ liệu tự động, chỉnh sửa ô trực tiếp, tính toán và xuất file chuẩn Microsoft Excel.")

    st.sidebar.header("⚙️ Nạp Dữ Liệu Excel")
    input_method = st.sidebar.radio(
        "Chọn hình thức nạp liệu:",
        ["📝 Dán Văn Bản / CSV", "📁 Tải Tệp Excel / CSV"]
    )

    df_excel_data = None

    if input_method == "📝 Dán Văn Bản / CSV":
        raw_input = st.sidebar.text_area(
            "Dán dữ liệu thô vào đây:",
            value="Date,Open,High,Low,Close,Volume\n2026-11-09,20.45,20.50,19.85,19.90,25847000\n2026-11-06,19.95,20.60,19.95,20.40,31408000",
            height=200
        )
        if st.sidebar.button("🚀 Bóc Tách & Nạp Dữ Liệu", type="primary", use_container_width=True):
            if raw_input.strip():
                parsed_df = smart_parse_input(raw_input)
                if not parsed_df.empty:
                    st.session_state['excel_working_df'] = parsed_df
                    st.success(f"✅ Đã xử lý {len(parsed_df)} dòng dữ liệu thành công!")
                else:
                    st.warning("⚠️ Không thể nhận diện cấu trúc. Hãy kiểm tra lại dữ liệu đầu vào!")
    else:
        file_upload = st.sidebar.file_uploader("Chọn file Excel/CSV", type=["xlsx", "xls", "csv"])
        if file_upload is not None:
            try:
                content_bytes = io.BytesIO(file_upload.getvalue())
                if file_upload.name.lower().endswith('.csv'):
                    parsed_df = pd.read_csv(content_bytes)
                else:
                    parsed_df = pd.read_excel(content_bytes, engine='openpyxl')
                st.session_state['excel_working_df'] = parsed_df
                st.success(f"✅ Đã nạp file thành công: {file_upload.name}")
            except Exception as ex:
                st.sidebar.error(f"❌ Lỗi nạp file: {ex}")

    if 'excel_working_df' in st.session_state:
        df_excel_data = st.session_state['excel_working_df']

    if df_excel_data is not None and not df_excel_data.empty:
        ex_tab1, ex_tab2, ex_tab3 = st.tabs([
            "✏️ Chỉnh Sửa Trực Tiếp (Grid Editor)", 
            "📈 Thống Kê & Biểu Đồ Kỹ Thuật", 
            "💾 Xuất File Chuẩn Office"
        ])

        with ex_tab1:
            st.subheader(f"📊 Bảng Dữ Liệu Hiện Tại ({len(df_excel_data)} dòng, {len(df_excel_data.columns)} cột)")
            st.caption("💡 Bạn có thể nhấp đôi vào từng ô để sửa giá trị hoặc chèn/xóa dòng trực tiếp:")
            
            edited_grid = st.data_editor(
                df_excel_data,
                num_rows="dynamic",
                use_container_width=True,
                key="excel_grid_editor"
            )
            st.session_state['excel_working_df'] = edited_grid

        with ex_tab2:
            st.subheader("📈 Phân Tích Chỉ Số & Biểu Đồ Kỹ Thuật")
            df_curr = st.session_state['excel_working_df']
            num_cols = df_curr.select_dtypes(include=[np.number]).columns.tolist()

            if num_cols:
                kpi_box = st.columns(min(len(num_cols), 4))
                for idx, c in enumerate(num_cols[:4]):
                    with kpi_box[idx]:
                        st.metric(label=f"Trung bình {c}", value=f"{df_curr[c].mean():,.2f}")

            # Vẽ biểu đồ nến nếu có đủ cột chứng khoán
            stock_fields = {'Open', 'High', 'Low', 'Close'}
            if stock_fields.issubset(set(df_curr.columns)):
                fig = go.Figure(data=[go.Candlestick(
                    x=df_curr['Date'] if 'Date' in df_curr.columns else df_curr.index,
                    open=df_curr['Open'], high=df_curr['High'], low=df_curr['Low'], close=df_curr['Close'],
                    name="Nến Giá"
                )])
                fig.update_layout(title="Biểu Đồ Nến Kỹ Thuật", template="plotly_white", height=450)
                st.plotly_chart(fig, use_container_width=True)
            elif num_cols:
                chart_target = st.selectbox("Chọn cột vẽ biểu đồ đường:", num_cols)
                st.line_chart(df_curr[chart_target])

        with ex_tab3:
            st.subheader("💾 Lưu & Xuất File Excel Định Dạng Chuyên Nghiệp")
            st.caption("Tự động dập viền, căn giữa dữ liệu & co giãn độ rộng cột khi mở bằng Microsoft Excel:")

            export_df = st.session_state['excel_working_df']
            col_x1, col_x2 = st.columns(2)

            # Đóng gói Excel qua OpenPyXL
            excel_out = io.BytesIO()
            with pd.ExcelWriter(excel_out, engine='openpyxl') as writer:
                export_df.to_excel(writer, index=False, sheet_name='Du_Lieu_Chuan')
                ws = writer.sheets['Du_Lieu_Chuan']
                
                center_align = Alignment(horizontal='center', vertical='center')
                hdr_font = Font(name="Calibri", bold=True, color="FFFFFF")
                hdr_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
                border_thin = Border(
                    left=Side(style='thin', color='D9D9D9'),
                    right=Side(style='thin', color='D9D9D9'),
                    top=Side(style='thin', color='D9D9D9'),
                    bottom=Side(style='thin', color='D9D9D9')
                )

                for col in ws.columns:
                    max_len = 0
                    col_letter = col[0].column_letter
                    for cell in col:
                        if cell.row == 1:
                            cell.fill = hdr_fill
                            cell.font = hdr_font
                            cell.alignment = center_align
                        else:
                            cell.alignment = center_align
                            cell.border = border_thin
                        try:
                            if cell.value is not None:
                                max_len = max(max_len, len(str(cell.value)))
                        except:
                            pass
                    ws.column_dimensions[col_letter].width = max(max_len + 6, 14)

            with col_x1:
                st.download_button(
                    label="📊 Tải File Excel Pro (.xlsx)",
                    data=excel_out.getvalue(),
                    file_name="Bao_Cao_Excel_Pro.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    type="primary"
                )

            csv_data = export_df.to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig')
            with col_x2:
                st.download_button(
                    label="📄 Tải File CSV Standard (.csv)",
                    data=csv_data,
                    file_name="Bao_Cao_Excel_Pro.csv",
                    mime="text/csv",
                    use_container_width=True
                )
    else:
        st.info("👈 Bạn vui lòng dán dữ liệu hoặc tải tệp lên từ Sidebar để bắt đầu làm việc nhé!")

# =============================================================================
# PHÂN HỆ 3: QUẢN LÝ TỆP (DOCUMENT CENTER)
# =============================================================================
else:
    st.subheader("📁 Document Center - Trung Tâm Quản Lý Tài Liệu")
    st.caption("Quản lý danh mục tập tin, giám sát dung lượng và phân loại lưu trữ.")

    col_stat1, col_stat2, col_stat3 = st.columns(3)
    col_stat1.metric("Tài liệu Word", "4 Tệp", "1.2 MB")
    col_stat2.metric("Bảng tính Excel", "8 Tệp", "5.8 MB")
    col_stat3.metric("Lưu trữ Cloud", "7.0 MB / 5.0 GB", "An toàn")

    st.markdown("---")
    st.write("##### 📋 Danh Sách Tài Liệu Trong Kho Lưu Trữ:")

    dummy_files_data = pd.DataFrame([
        {"Tên Tệp": "Bao_Cao_Cong_Tac_Tuan.docx", "Loại": "Word Document", "Kích Thước": "345 KB", "Ngày Cập Nhật": "2026-11-11"},
        {"Tên Tệp": "Thong_Ke_Doanh_Thu_Q4.xlsx", "Loại": "Excel Spreadsheet", "Kích Thước": "1.2 MB", "Ngày Cập Nhật": "2026-11-10"},
        {"Tên Tệp": "Bien_Ban_Hop_Giam_Doc.docx", "Loại": "Word Document", "Kích Thước": "180 KB", "Ngày Cập Nhật": "2026-11-09"},
        {"Tên Tệp": "Du_Lieu_Giao_Dich_VNM.csv", "Loại": "CSV Data File", "Kích Thước": "4.1 MB", "Ngày Cập Nhật": "2026-11-08"}
    ])

    st.dataframe(dummy_files_data, use_container_width=True)
