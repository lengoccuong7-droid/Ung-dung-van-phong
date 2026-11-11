import streamlit as st
import pandas as pd
import numpy as np
import io
import plotly.graph_objects as go
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side

# -----------------------------------------------------------------------------
# 1. CẤU HÌNH GIAO DIỆN PHONG CÁCH MICROSOFT OFFICE & CHỐNG CHỮ MỜ
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="Microsoft Office & Math Suite Pro", 
    layout="wide", 
    page_icon="💼",
    initial_sidebar_state="expanded"
)

st.markdown("""
    <style>
    .stApp { background-color: #f3f4f6; color: #111827 !important; }
    p, span, label, div, h1, h2, h3, h4, h5, h6, .stMarkdown {
        color: #111827 !important;
        font-weight: 500;
    }
    .stCaption p {
        color: #374151 !important;
        font-weight: 600 !important;
    }
    div[data-testid="stMetricValue"], div[data-testid="stMetricLabel"] {
        color: #111827 !important;
    }
    .office-ribbon {
        background-color: #ffffff;
        border: 1px solid #cbd5e1;
        border-radius: 8px;
        padding: 10px 15px;
        margin-bottom: 15px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
        display: flex;
        gap: 15px;
        align-items: center;
    }
    .stTabs [data-baseweb="tab"] {
        background-color: #ffffff;
        border-radius: 6px 6px 0 0;
        padding: 8px 16px;
        font-weight: 700;
        color: #1e3a8a !important;
        border: 1px solid #cbd5e1;
    }
    .stTabs [aria-selected="true"] {
        background-color: #1b5e20 !important;
        color: #ffffff !important;
    }
    </style>
""", unsafe_allow_html=True)

st.title("💼 Microsoft Office & Math Suite Pro (All-In-One)")
st.caption("Mô phỏng trọn vẹn không gian làm việc chuẩn Microsoft Office: Word, Excel, Math Studio & Document Center")

# -----------------------------------------------------------------------------
# 2. HÀM BÓC TÁCH DỮ LIỆU EXCEL ĐA NĂNG
# -----------------------------------------------------------------------------
def smart_parse_input(text_data):
    lines = [line.strip() for line in text_data.strip().split('\n') if line.strip()]
    if not lines:
        return pd.DataFrame()
    
    first_line = lines[0]
    if ',' in first_line and ('Date' in first_line or 'Open' in first_line):
        try:
            return pd.read_csv(io.StringIO(text_data))
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
                vol = float(vol_str.replace('M', '')) * 1_000_000 if 'M' in vol_str else (float(vol_str.replace('K', '')) * 1_000 if 'K' in vol_str else float(vol_str))
                    
                parsed_rows.append({
                    'Date': date_str, 'Open': open_p, 'High': high_p,
                    'Low': low_p, 'Close': close_p, 'Volume': vol
                })
            except Exception:
                continue
    return pd.DataFrame(parsed_rows)

# -----------------------------------------------------------------------------
# 3. ĐIỀU HƯỚNG CHÍNH (SIDEBAR)
# -----------------------------------------------------------------------------
st.sidebar.header("📁 Bộ Ứng Dụng Microsoft Office")
app_module = st.sidebar.radio(
    "Mở ứng dụng:",
    [
        "🟢 Microsoft Excel (Bảng Tính & Phân Tích)",
        "🔵 Microsoft Word (Soạn Thảo Văn Bản)",
        "🧮 Math Studio (Tính Toán & Đồ Thị)",
        "📂 Quản Lý Tệp (OneDrive / Cloud)"
    ]
)

# =============================================================================
# 1. MICROSOFT EXCEL (BẢNG TÍNH & PHÂN TÍCH CHUYÊN NGHIỆP)
# =============================================================================
if app_module.startswith("🟢 Microsoft Excel"):
    st.subheader("🟢 Microsoft Excel - Bảng Tính Chuyên Nghiệp")

    # KHUNG NẠP DỮ LIỆU ĐẶT TRỰC TIẾP TRÊN MÀN HÌNH CHÍNH (RẤT DỄ THẤY)
    with st.expander("📥 MỞ KHUNG NẠP DỮ LIỆU (Tải file hoặc Dán văn bản)", expanded=True):
        col_n1, col_n2 = st.columns(2)
        
        with col_n1:
            st.write("##### 📁 Cách 1: Tải lên tệp Excel / CSV")
            uploaded_f = st.file_uploader("Chọn tệp từ máy tính:", type=["xlsx", "xls", "csv"])
            if uploaded_f is not None:
                try:
                    b_io = io.BytesIO(uploaded_f.getvalue())
                    df_parsed = pd.read_csv(b_io) if uploaded_f.name.lower().endswith('.csv') else pd.read_excel(b_io, engine='openpyxl')
                    st.session_state['excel_data'] = df_parsed
                    st.success(f"✅ Nạp thành công tệp: {uploaded_f.name} ({len(df_parsed)} dòng)")
                except Exception as e:
                    st.error(f"❌ Lỗi đọc tệp: {e}")

        with col_n2:
            st.write("##### 📝 Cách 2: Dán văn bản thô / CSV trực tiếp")
            raw_txt = st.text_area(
                "Dán dữ liệu tại đây:", 
                value="Date,Open,High,Low,Close,Volume\n2026-11-09,20.45,20.50,19.85,19.90,25847000\n2026-11-06,19.95,20.60,19.95,20.40,31408000", 
                height=100
            )
            if st.button("🚀 Xử Lý & Nạp Dữ Liệu Này", type="primary"):
                if raw_txt.strip():
                    df_parsed = smart_parse_input(raw_txt)
                    if not df_parsed.empty:
                        st.session_state['excel_data'] = df_parsed
                        st.success(f"✅ Đã nạp thành công {len(df_parsed)} dòng!")
                    else:
                        st.error("❌ Không thể đọc cấu trúc dữ liệu.")

    # Lấy dữ liệu từ session state nếu có
    if 'excel_data' in st.session_state:
        df_excel = st.session_state['excel_data']
    else:
        df_excel = None

    if df_excel is not None and not df_excel.empty:
        st.markdown("---")
        ex_tab1, ex_tab2, ex_tab3 = st.tabs(["📊 Trang Tính (Grid)", "📈 Biểu Đồ Phân Tích", "💾 Lưu & Xuất File"])

        with ex_tab1:
            st.caption("💡 Nhấp đôi vào ô bất kỳ để chỉnh sửa dữ liệu trực tiếp như trên Excel:")
            edited_table = st.data_editor(df_excel, num_rows="dynamic", use_container_width=True)
            st.session_state['excel_data'] = edited_table

        with ex_tab2:
            st.write("##### 📈 Biểu Đồ Kỹ Thuật Nhanh:")
            numeric_c = df_excel.select_dtypes(include=[np.number]).columns.tolist()
            if {'Open', 'High', 'Low', 'Close'}.issubset(set(df_excel.columns)):
                fig_ex = go.Figure(data=[go.Candlestick(
                    x=df_excel['Date'] if 'Date' in df_excel.columns else df_excel.index,
                    open=df_excel['Open'], high=df_excel['High'], low=df_excel['Low'], close=df_excel['Close']
                )])
                fig_ex.update_layout(title="Biểu Đồ Nến Giá", template="plotly_white", height=400)
                st.plotly_chart(fig_ex, use_container_width=True)
            elif numeric_c:
                sel_c = st.selectbox("Chọn cột vẽ biểu đồ:", numeric_c)
                st.line_chart(df_excel[sel_c])

        with ex_tab3:
            st.write("##### 💾 Xuất File Chuẩn Office:")
            out_buf = io.BytesIO()
            with pd.ExcelWriter(out_buf, engine='openpyxl') as writer:
                df_excel.to_excel(writer, index=False, sheet_name='Sheet1')
                ws = writer.sheets['Sheet1']
                c_align = Alignment(horizontal='center', vertical='center')
                h_font = Font(name="Calibri", bold=True, color="FFFFFF")
                h_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
                b_side = Border(left=Side(style='thin', color='D9D9D9'), right=Side(style='thin', color='D9D9D9'), top=Side(style='thin', color='D9D9D9'), bottom=Side(style='thin', color='D9D9D9'))
                
                for col in ws.columns:
                    max_l = 0
                    c_let = col[0].column_letter
                    for cell in col:
                        if cell.row == 1:
                            cell.fill = h_fill; cell.font = h_font; cell.alignment = c_align
                        else:
                            cell.alignment = c_align; cell.border = b_side
                        try:
                            if cell.value: max_l = max(max_l, len(str(cell.value)))
                        except: pass
                    ws.column_dimensions[c_let].width = max(max_l + 6, 14)

            st.download_button(
                label="📊 Tải Xuất File Excel (.xlsx)",
                data=out_buf.getvalue(),
                file_name="Bang_Tinh_Chuan.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True
            )
    else:
        st.info("👆 Sếp hãy bấm mở khung **'MỞ KHUNG NẠP DỮ LIỆU'** phía trên để tải file hoặc dán dữ liệu vào bảng tính nhé!")

# =============================================================================
# 2. MICROSOFT WORD (SOẠN THẢO VĂN BẢN)
# =============================================================================
elif app_module.startswith("🔵 Microsoft Word"):
    st.subheader("🔵 Microsoft Word - Soạn Thảo Văn Bản Hành Chính")
    
    if 'w_title' not in st.session_state: st.session_state['w_title'] = "BÁO CÁO CÔNG TÁC"
    if 'w_content' not in st.session_state: st.session_state['w_content'] = "Kính gửi Ban Giám Đốc,\n\nNội dung báo cáo chi tiết..."

    st.session_state['w_title'] = st.text_input("Tiêu đề văn bản:", value=st.session_state['w_title'])
    st.session_state['w_content'] = st.text_area("Nội dung tài liệu:", value=st.session_state['w_content'], height=300)

    wc = len(st.session_state['w_content'].split())
    st.markdown(f"**Thống kê:** {wc} từ | Trạng thái: Đã lưu nháp an toàn.")

    doc_full = f"# {st.session_state['w_title']}\n\n{st.session_state['w_content']}"
    st.download_button(
        label="📄 Tải Xuất Tài Liệu Văn Bản (.txt)",
        data=doc_full.encode('utf-8-sig'),
        file_name=f"{st.session_state['w_title'].replace(' ', '_')}.txt",
        mime="text/plain",
        type="primary"
    )

# =============================================================================
# 3. MATH STUDIO (TÍNH TOÁN & ĐỒ THỊ)
# =============================================================================
elif app_module.startswith("🧮 Math Studio"):
    st.subheader("🧮 Math Studio - Trạm Tính Toán & Đồ Thị Khoa Học")
    
    c1, c2 = st.columns(2)
    with c1:
        deg = st.selectbox("Chọn hàm số:", ["Đa thức bậc 2: y = ax² + bx + c", "Đa thức bậc 3: y = ax³ + bx² + cx + d"])
        xmin = st.number_input("X bắt đầu:", value=-10.0)
    with c2:
        xmax = st.number_input("X kết thúc:", value=10.0)
        steps = st.slider("Số điểm khảo sát:", 10, 200, 100)

    ac = st.number_input("Hệ số a:", value=1.0)
    bc = st.number_input("Hệ số b:", value=0.0)
    cc = st.number_input("Hệ số c:", value=-4.0)
    dc = st.number_input("Hệ số d (Bậc 3):", value=0.0)

    if st.button("🚀 Tính Toán & Vẽ Đồ Thị", type="primary"):
        xs = np.linspace(xmin, xmax, steps)
        ys = ac*(xs**2) + bc*xs + cc if "bậc 2" in deg else ac*(xs**3) + bc*(xs**2) + cc*xs + dc
        
        df_m = pd.DataFrame({"X": xs, "Y": ys})
        fig_m = go.Figure(data=go.Scatter(x=xs, y=ys, mode='lines+markers', line=dict(color='green', width=3)))
        fig_m.update_layout(title="Đồ Thị Khảo Sát", template="plotly_white", height=400)
        st.plotly_chart(fig_m, use_container_width=True)
        st.dataframe(df_m, use_container_width=True)

# =============================================================================
# 4. QUẢN LÝ TỆP (DOCUMENT CENTER)
# =============================================================================
else:
    st.subheader("📂 Trung Tâm Quản Lý Tệp (OneDrive / Cloud Storage)")
    st.metric("Dung lượng đám mây", "12.4 MB / 5.0 GB", "Đồng bộ an toàn")
    st.dataframe(pd.DataFrame([
        {"Tên Tệp": "Bao_Cao_Tai_Chinh.xlsx", "Loại": "Microsoft Excel", "Kích Thước": "1.2 MB", "Cập Nhật": "Hôm nay"},
        {"Tên Tệp": "Quy_Che_Noi_Bo.docx", "Loại": "Microsoft Word", "Kích Thước": "340 KB", "Cập Nhật": "Hôm qua"}
    ]), use_container_width=True)
