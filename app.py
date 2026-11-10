import streamlit as st
import pandas as pd
import numpy as np
import io
import os

# -----------------------------------------------------------------------------
# CẤU HÌNH THÔNG TIN ỨNG DỤNG VĂN PHÒNG ĐA NĂNG
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="Công Cụ Chuẩn Hóa Dữ Liệu Văn Phòng Đa Năng", 
    layout="wide", 
    page_icon="💼",
    initial_sidebar_state="expanded"
)

st.title("💼 Công Cụ Chuẩn Hóa & Sửa Lỗi Dữ Liệu Văn Phòng Đa Năng")
st.caption("Tự động sửa lỗi định dạng, làm sạch khoảng trắng, chuẩn hóa số liệu & xuất tệp Microsoft Excel / CSV mượt 100%")

# -----------------------------------------------------------------------------
# HÀM TỰ ĐỘNG LÀM SẠCH VÀ CHUẨN HÓA DỮ LIỆU THÔ VĂN PHÒNG
# -----------------------------------------------------------------------------
def auto_clean_office_dataframe(df):
    """
    Hàm tự động quét và sửa các lỗi phổ biến trong mọi loại tệp văn phòng
    """
    cleaned_df = df.copy()
    
    # 1. Làm sạch tên cột (Xóa khoảng trắng thừa 2 đầu)
    cleaned_df.columns = [str(col).strip() for col in cleaned_df.columns]
    
    # 2. Quét từng cột để sửa định dạng
    for col in cleaned_df.columns:
        # Nếu dữ liệu dạng chuỗi/văn bản
        if cleaned_df[col].dtype == 'object':
            # Xóa khoảng trắng thừa ở từng ô
            cleaned_df[col] = cleaned_df[col].astype(str).str.strip()
            
            # Xử lý các từ viết tắt đơn vị hay gặp (K, M) hoặc dấu phẩy
            temp_str = cleaned_df[col].str.upper().str.replace(',', '', regex=False)
            
            # Kiểm tra xem có thể chuyển thành số không
            def try_numeric(v):
                try:
                    if 'M' in v: return float(v.replace('M', '')) * 1_000_000
                    if 'K' in v: return float(v.replace('K', '')) * 1_000
                    return float(v)
                except:
                    return np.nan

            numeric_converted = temp_str.apply(try_numeric)
            
            # Nếu trên 70% số dòng trong cột chuyển thành số thành công thì ép kiểu số
            if numeric_converted.notna().sum() / max(len(cleaned_df), 1) > 0.7:
                cleaned_df[col] = numeric_converted
            else:
                # Thử chuyển đổi sang định dạng Ngày tháng
                try:
                    date_converted = pd.to_datetime(cleaned_df[col], errors='coerce', dayfirst=True)
                    if date_converted.notna().sum() / max(len(cleaned_df), 1) > 0.7:
                        cleaned_df[col] = date_converted.dt.strftime('%Y-%m-%d')
                except:
                    pass

    return cleaned_df

# -----------------------------------------------------------------------------
# GIAO DIỆN NẠP VÀ THAO TÁC TỆP
# -----------------------------------------------------------------------------
st.sidebar.header("📥 Nạp Tệp Văn Phòng Cần Xử Lý")
uploaded_file = st.sidebar.file_uploader(
    "Chọn tệp Excel (.xlsx, .xls) hoặc CSV bất kỳ", 
    type=["xlsx", "xls", "csv"]
)

if uploaded_file is not None:
    try:
        # Nạp dữ liệu qua bộ nhớ đệm BytesIO để triệt tiêu lỗi đường dẫn/openpyxl
        file_bytes = io.BytesIO(uploaded_file.getvalue())
        
        if uploaded_file.name.lower().endswith('.csv'):
            df_raw = pd.read_csv(file_bytes)
        else:
            df_raw = pd.read_excel(file_bytes, engine='openpyxl')
            
        st.success(f"✅ Đã nạp thành công tệp: **{uploaded_file.name}**")
        
        # Sửa lỗi tự động
        with st.spinner("🔄 Đang quét và tự động chuẩn hóa dữ liệu..."):
            df_cleaned = auto_clean_office_dataframe(df_raw)
            
        # Thống kê nhanh thông số tệp
        col_s1, col_s2, col_s3, col_s4 = st.columns(4)
        col_s1.metric("Tổng Số Dòng", f"{len(df_cleaned):,}")
        col_s2.metric("Tổng Số Cột", f"{len(df_cleaned.columns):,}")
        col_s3.metric("Số Ô Trống (Null)", f"{df_cleaned.isna().sum().sum():,}")
        col_s4.metric("Dòng Trùng Lặp", f"{df_cleaned.duplicated().sum():,}")

        st.markdown("---")
        st.subheader("📋 Bảng Dữ Liệu Đã Được Làm Sạch & Chuẩn Hóa:")
        st.dataframe(df_cleaned, use_container_width=True)
        
        # Nếu có dòng trùng lặp, cho phép bấm nút loại bỏ
        if df_cleaned.duplicated().sum() > 0:
            if st.button("🧹 Loại Bỏ Tất Cả Các Dòng Trùng Lặp", use_container_width=True):
                df_cleaned = df_cleaned.drop_duplicates().reset_index(drop=True)
                st.success("✅ Đã loại bỏ sạch sẽ các dòng trùng lặp!")
                st.rerun()

        st.markdown("---")
        st.subheader("💾 LƯU & XUẤT TỆP MICROSOFT OFFICE / CSV")
        st.caption("Bấm nút bên dưới để tải về tệp sạch chuẩn 100%, mở trên Microsoft Excel không bị lỗi font hay đĩa:")
        
        col_out1, col_out2 = st.columns(2)
        
        # 1. Xuất file Microsoft Excel (.xlsx) chuẩn
        buffer_excel = io.BytesIO()
        with pd.ExcelWriter(buffer_excel, engine='openpyxl') as writer:
            df_cleaned.to_excel(writer, index=False, sheet_name='Du_Lieu_Van_Phong')
        excel_bytes = buffer_excel.getvalue()
        
        out_excel_name = f"chuan_hoa_{os.path.splitext(uploaded_file.name)[0]}.xlsx"
        with col_out1:
            st.download_button(
                label="📊 Xuất File Microsoft Excel (.xlsx) Chuẩn",
                data=excel_bytes,
                file_name=out_excel_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                type="primary"
            )
            
        # 2. Xuất file CSV (.csv) chuẩn UTF-8 (Không lỗi font Tiếng Việt)
        csv_bytes = df_cleaned.to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig')
        out_csv_name = f"chuan_hoa_{os.path.splitext(uploaded_file.name)[0]}.csv"
        with col_out2:
            st.download_button(
                label="📄 Xuất File CSV (.csv) Chuẩn UTF-8",
                data=csv_bytes,
                file_name=out_csv_name,
                mime="text/csv",
                use_container_width=True
            )
            
    except Exception as e:
        st.error(f"❌ Lỗi đọc hoặc xử lý tệp: {e}")
        st.info("💡 Gợi ý: Nếu file bị khóa định dạng từ ứng dụng khác, sếp mở ra bấm 'Save As' dạng CSV rồi nạp lên lại là app đọc sạch mượt 100%!")

else:
    st.info("👆 Sếp chọn hoặc kéo thả bất kỳ tệp Excel/CSV nào (bán hàng, kho, kế toán, danh sách...) vào khung bên trái để ứng dụng làm sạch và xuất file nhé!")
