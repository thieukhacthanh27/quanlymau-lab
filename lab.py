import streamlit as st
import pandas as pd
from datetime import datetime
from streamlit_gsheets import GSheetsConnection

# ==========================================
# CẤU HÌNH CƠ BẢN
# ==========================================
st.set_page_config(page_title="Hệ thống Quản lý Mẫu - Lab GC", layout="wide")

# Đã thay bằng link thật của bạn
SHEET_URL = "https://docs.google.com/spreadsheets/d/1F2wFnxboWTFWDMGUuBDRGB901a5EKgvazHxkCgBjjRU/edit?gid=0#gid=0"

STATUSES = [
    "🔴 1. Chờ xử lý", "🟠 2. Đang xử lý mẫu", "🟡 3. Chờ chạy máy",
    "🔵 4. Đang chạy máy", "🟣 5. Đang tính số liệu", "🟢 6. Lưu kho", "⚫ 7. Đã tiêu hủy"
]

# ==========================================
# HÀM XỬ LÝ DỮ LIỆU BẰNG GOOGLE SHEETS API
# ==========================================
# Khởi tạo kết nối với Google Sheets thông qua Secrets đã cài
conn = st.connection("gsheets", type=GSheetsConnection)

def load_data():
    # Đọc dữ liệu trực tiếp từ Google Sheets, bỏ qua cache (ttl=0) để luôn lấy bản mới nhất
    df = conn.read(spreadsheet=SHEET_URL, ttl=0)
    
    # Nếu bảng trống chưa có gì, tự tạo tiêu đề cột và đẩy lên Google Sheets
    if df.empty or len(df.columns) == 0 or "Mã Mẫu" not in df.columns:
        df = pd.DataFrame(columns=["Mã Mẫu", "Tên Mẫu", "Nền Mẫu", "Chỉ Tiêu", "Trạng Thái", "Người Giữ", "Ghi Chú", "Giờ Nhận"])
        conn.update(spreadsheet=SHEET_URL, data=df)
    
    df['Giờ Nhận'] = pd.to_datetime(df['Giờ Nhận'], errors='coerce')
    return df

def save_data(df):
    df_save = df.copy()
    # Phải biến thời gian thành chuỗi chữ trước khi lưu để Google Sheets không báo lỗi
    df_save['Giờ Nhận'] = df_save['Giờ Nhận'].dt.strftime('%Y-%m-%d %H:%M:%S')
    
    # Lệnh đẩy dữ liệu lên Google Sheets
    conn.update(spreadsheet=SHEET_URL, data=df_save)
    st.cache_data.clear()

if "df" not in st.session_state:
    st.session_state.df = load_data()

# ==========================================
# GIAO DIỆN CHÍNH
# ==========================================
st.title("🧪 Bảng Điều Khiển LIMS - Phòng Lab GC")

# --- BỘ LỌC NGÀY LÀM VIỆC ---
col_date, col_search, col_filter = st.columns([1.5, 1, 1.5])
with col_date:
    selected_date = st.date_input("📅 Chọn Ngày Làm Việc:", datetime.today())
with col_search:
    search_query = st.text_input("🔍 Nhập ID mã mẫu:")
with col_filter:
    filter_status = st.multiselect("Lọc theo trạng thái:", STATUSES, default=[])

st.subheader(f"📋 Danh sách công việc ngày {selected_date.strftime('%d/%m/%Y')}")

# --- XỬ LÝ LOGIC MẪU TỒN ĐỌNG & TRONG NGÀY ---
df_current = st.session_state.df.copy()
df_current["Ngày Nhận"] = df_current["Giờ Nhận"].dt.date

mask_ton_dong = (df_current["Ngày Nhận"] < selected_date) & (~df_current["Trạng Thái"].isin(["🟢 6. Lưu kho", "⚫ 7. Đã tiêu hủy"]))
mask_trong_ngay = (df_current["Ngày Nhận"] == selected_date)

df_display = df_current[mask_ton_dong | mask_trong_ngay].copy()
df_display["Phân Loại"] = "🟢 Nhận trong ngày"
df_display.loc[mask_ton_dong, "Phân Loại"] = "⚠️ TỒN ĐỌNG CHƯA XONG"
df_display = df_display.sort_values(by=["Phân Loại", "Giờ Nhận"], ascending=[False, True])

if search_query:
    df_display = df_display[df_display["Mã Mẫu"].str.contains(search_query, case=False, na=False)]
if filter_status:
    df_display = df_display[df_display["Trạng Thái"].isin(filter_status)]

# --- BẢNG DỮ LIỆU TƯƠNG TÁC ---
st.caption("Mẹo: Mẫu nào làm xong và chuyển trạng thái thành 'Lưu kho', sang ngày mai sẽ tự động biến mất khỏi danh sách chờ.")
edited_df = st.data_editor(
    df_display,
    column_config={
        "Trạng Thái": st.column_config.SelectboxColumn("Trạng Thái Hiện Tại", options=STATUSES, required=True),
        "Nền Mẫu": st.column_config.SelectboxColumn("Nền Mẫu", options=["Khí", "Nước"], required=True),
        "Phân Loại": st.column_config.TextColumn("Phân Loại", disabled=True),
        "Giờ Nhận": st.column_config.DatetimeColumn("Giờ Nhận", format="DD/MM/YYYY HH:mm", disabled=True),
        "Ngày Nhận": None 
    },
    disabled=["Mã Mẫu", "Tên Mẫu", "Chỉ Tiêu", "Phân Loại", "Giờ Nhận"], 
    use_container_width=True,
    num_rows="dynamic",
    key="data_editor"
)

if st.button("💾 Lưu các thay đổi vào Hệ thống"):
    for index, row in edited_df.iterrows():
        st.session_state.df.loc[index, "Trạng Thái"] = row["Trạng Thái"]
        st.session_state.df.loc[index, "Nền Mẫu"] = row["Nền Mẫu"]
        st.session_state.df.loc[index, "Người Giữ"] = row["Người Giữ"]
        st.session_state.df.loc[index, "Ghi Chú"] = row["Ghi Chú"]
    save_data(st.session_state.df)
    st.success("Đã cập nhật cơ sở dữ liệu thành công!")
    st.rerun()

st.divider()

# ==========================================
# THÊM MẪU & XUẤT SEQUENCE
# ==========================================
col_add, col_export = st.columns([1, 1])

with col_add:
    st.subheader("📥 Tiếp nhận mẫu mới")
    with st.form("add_sample_form", clear_on_submit=True):
        new_id = st.text_input("Mã Mẫu (VD: NT-1509-01)*")
        new_name = st.text_input("Tên/Ký hiệu Mẫu")
        new_nen = st.selectbox("Nền Mẫu", ["Nước", "Khí"])
        new_chi_tieu = st.text_input("Chỉ tiêu đo (VD: VOCs, Formaldehyde)")
        new_nguoi = st.text_input("Người tiếp nhận (Ký tên)")
        
        submitted = st.form_submit_button("Thêm Mẫu")
        if submitted and new_id:
            new_row = pd.DataFrame([{
                "Mã Mẫu": new_id, "Tên Mẫu": new_name, "Nền Mẫu": new_nen, 
                "Chỉ Tiêu": new_chi_tieu, "Trạng Thái": STATUSES[0], 
                "Người Giữ": new_nguoi, "Ghi Chú": "", 
                "Giờ Nhận": datetime.now() 
            }])
            st.session_state.df = pd.concat([st.session_state.df, new_row], ignore_index=True)
            save_data(st.session_state.df)
            st.success(f"Đã thêm mẫu {new_id} thành công!")
            st.rerun()

with col_export:
    st.subheader("⚙️ Xuất Sequence MassHunter")
    st.info("Hệ thống sẽ gom TẤT CẢ các mẫu đang ở trạng thái '🟡 3. Chờ chạy máy' để tạo file Sequence CSV.")
    
    if st.button("Tạo File Sequence"):
        df_ready = st.session_state.df[st.session_state.df["Trạng Thái"] == "🟡 3. Chờ chạy máy"]
        if df_ready.empty:
            st.warning("Không có mẫu nào đang chờ chạy máy!")
        else:
            seq_df = pd.DataFrame()
            seq_df['Vial'] = range(1, len(df_ready) + 1)
            seq_df['Sample Name'] = df_ready['Mã Mẫu']
            seq_df['Sample Type'] = 'Sample'
            seq_df['Method'] = df_ready['Chỉ Tiêu'].apply(lambda x: 'VOCs.M' if 'VOC' in str(x).upper() else 'HCHO.M')
            seq_df['Data File'] = datetime.now().strftime("%Y%m%d") + "_" + df_ready['Mã Mẫu']
            
            csv = seq_df.to_csv(index=False).encode('utf-8')
            st.download_button(
                label="📥 Tải xuống Sequence.csv",
                data=csv,
                file_name=f"MassHunter_Seq_{datetime.now().strftime('%Y%m%d')}.csv",
                mime="text/csv",
            )
