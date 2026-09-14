import streamlit as st
import pandas as pd
import os
from datetime import datetime

# ==========================================
# CẤU HÌNH CƠ BẢN
# ==========================================
st.set_page_config(page_title="Hệ thống Quản lý Mẫu - Lab GC", layout="wide")
DATA_FILE = "lab_database.csv"

STATUSES = [
    "🔴 1. Chờ xử lý",
    "🟠 2. Đang xử lý mẫu",
    "🟡 3. Chờ chạy máy",
    "🔵 4. Đang chạy máy",
    "🟣 5. Đang tính số liệu",
    "🟢 6. Lưu kho",
    "⚫ 7. Đã tiêu hủy"
]

# ==========================================
# HÀM XỬ LÝ DỮ LIỆU
# ==========================================
def load_data():
    if not os.path.exists(DATA_FILE):
        df = pd.DataFrame(columns=["Mã Mẫu", "Tên Mẫu", "Nền Mẫu", "Chỉ Tiêu", "Trạng Thái", "Người Giữ", "Ghi Chú", "Giờ Nhận"])
        df.to_csv(DATA_FILE, index=False)
    return pd.read_csv(DATA_FILE)

def save_data(df):
    df.to_csv(DATA_FILE, index=False)

# Khởi tạo session state
if "df" not in st.session_state:
    st.session_state.df = load_data()

# ==========================================
# GIAO DIỆN CHÍNH
# ==========================================
st.title("🧪 Bảng Điều Khiển LIMS - Phòng Lab GC")

# --- BỘ LỌC TÌM KIẾM ---
col1, col2 = st.columns(2)
with col1:
    search_query = st.text_input("🔍 Nhập ID mã mẫu để tìm kiếm:")
with col2:
    filter_status = st.multiselect("Lọc theo trạng thái:", STATUSES, default=[])

# --- BẢNG DỮ LIỆU TƯƠNG TÁC (CHỈNH SỬA TRỰC TIẾP) ---
st.subheader("📋 Bảng theo dõi thời gian thực (Cập nhật tự động)")
st.caption("Mẹo: Click đúp vào ô 'Trạng Thái' hoặc 'Người Giữ' để thay đổi trực tiếp trên bảng.")

df_display = st.session_state.df.copy()

if search_query:
    df_display = df_display[df_display["Mã Mẫu"].str.contains(search_query, case=False, na=False)]
if filter_status:
    df_display = df_display[df_display["Trạng Thái"].isin(filter_status)]

# Cấu hình data editor cho phép edit bằng dropdown
edited_df = st.data_editor(
    df_display,
    column_config={
        "Trạng Thái": st.column_config.SelectboxColumn(
            "Trạng Thái Hiện Tại", options=STATUSES, required=True
        ),
        "Nền Mẫu": st.column_config.SelectboxColumn(
            "Nền Mẫu", options=["Khí", "Nước"], required=True
        ),
    },
    use_container_width=True,
    num_rows="dynamic",
    key="data_editor"
)

# Nút lưu thay đổi
if st.button("💾 Lưu các thay đổi vào Hệ thống"):
    st.session_state.df = edited_df
    save_data(edited_df)
    st.success("Đã cập nhật cơ sở dữ liệu thành công!")

st.divider()

# ==========================================
# TÍNH NĂNG NÂNG CAO & THÊM MẪU
# ==========================================
col_add, col_export = st.columns([1, 1])

# --- KHU VỰC THÊM MẪU MỚI ---
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
                "Giờ Nhận": datetime.now().strftime("%Y-%m-%d %H:%M")
            }])
            st.session_state.df = pd.concat([st.session_state.df, new_row], ignore_index=True)
            save_data(st.session_state.df)
            st.success(f"Đã thêm mẫu {new_id} thành công!")
            st.rerun()

# --- KHU VỰC XUẤT SEQUENCE CHO GC-MS ---
with col_export:
    st.subheader("⚙️ Xuất Sequence MassHunter")
    st.info("Hệ thống sẽ lọc các mẫu đang ở trạng thái '🟡 3. Chờ chạy máy' để tạo file Sequence CSV.")
    
    if st.button("Tạo File Sequence"):
        df_ready = st.session_state.df[st.session_state.df["Trạng Thái"] == "🟡 3. Chờ chạy máy"]
        if df_ready.empty:
            st.warning("Không có mẫu nào đang chờ chạy máy!")
        else:
            seq_df = pd.DataFrame()
            seq_df['Vial'] = range(1, len(df_ready) + 1)
            seq_df['Sample Name'] = df_ready['Mã Mẫu']
            seq_df['Sample Type'] = 'Sample' # Mặc định
            seq_df['Method'] = df_ready['Chỉ Tiêu'].apply(lambda x: 'VOCs.M' if 'VOC' in str(x).upper() else 'HCHO.M')
            seq_df['Data File'] = datetime.now().strftime("%Y%m%d") + "_" + df_ready['Mã Mẫu']
            
            csv = seq_df.to_csv(index=False).encode('utf-8')
            st.download_button(
                label="📥 Tải xuống Sequence.csv",
                data=csv,
                file_name=f"MassHunter_Seq_{datetime.now().strftime('%Y%m%d')}.csv",
                mime="text/csv",
            )
            st.success(f"Đã tạo Sequence cho {len(df_ready)} mẫu.")