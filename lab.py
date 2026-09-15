import streamlit as st
import pandas as pd
from datetime import datetime
from streamlit_gsheets import GSheetsConnection

# ==========================================
# CẤU HÌNH CƠ BẢN
# ==========================================
st.set_page_config(page_title="Hệ thống Quản lý Mẫu - Lab GC", layout="wide")

# BẠN NHỚ DÁN LẠI LINK GOOGLE SHEETS THẬT VÀO ĐÂY:
SHEET_URL = "Dhttps://docs.google.com/spreadsheets/d/1F2wFnxboWTFWDMGUuBDRGB901a5EKgvazHxkCgBjjRU/edit?gid=0#gid=0"

STATUSES = [
    "🔴 1. Chờ xử lý", "🟠 2. Đang xử lý mẫu", "🟡 3. Chờ chạy máy",
    "🔵 4. Đang chạy máy", "🟣 5. Đang tính số liệu", "🟢 6. Lưu kho", "⚫ 7. Đã tiêu hủy"
]

# ==========================================
# HÀM XỬ LÝ DỮ LIỆU BẰNG GOOGLE SHEETS API
# ==========================================
conn = st.connection("gsheets", type=GSheetsConnection)

def load_data():
    df = conn.read(spreadsheet=SHEET_URL, ttl=0)
    
    if df.empty or len(df.columns) == 0 or "Mã Mẫu" not in df.columns:
        df = pd.DataFrame(columns=["Mã Mẫu", "Tên Mẫu", "Nền Mẫu", "Chỉ Tiêu", "Trạng Thái", "Người Giữ", "Ghi Chú", "Giờ Nhận"])
        conn.update(spreadsheet=SHEET_URL, data=df)
    
    df['Giờ Nhận'] = pd.to_datetime(df['Giờ Nhận'], errors='coerce')
    return df

def save_data(df):
    df_save = df.copy()
    df_save['Giờ Nhận'] = df_save['Giờ Nhận'].dt.strftime('%Y-%m-%d %H:%M:%S')
    
    conn.update(spreadsheet=SHEET_URL, data=df_save)
    st.cache_data.clear()

if "df" not in st.session_state:
    st.session_state.df = load_data()

# ==========================================
# GIAO DIỆN CHÍNH
# ==========================================
st.title("🧪 Bảng Điều Khiển LIMS - Phòng Lab GC")

col_date, col_search, col_filter = st.columns([1.5, 1, 1.5])
with col_date:
    selected_date = st.date_input("📅 Chọn Ngày Làm Việc:", datetime.today())
with col_search:
    search_query = st.text_input("🔍 Nhập ID mã mẫu:")
with col_filter:
    filter_status = st.multiselect("Lọc theo trạng thái:", STATUSES, default=[])

st.subheader(f"📋 Danh sách công việc ngày {selected_date.strftime('%d/%m/%Y')}")

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
# THÊM MẪU (THỦ CÔNG & EXCEL) & XUẤT SEQUENCE
# ==========================================
col_add, col_export = st.columns([1, 1])

with col_add:
    st.subheader("📥 Tiếp nhận mẫu mới")
    # Chia làm 2 tab giao diện
    tab_thu_cong, tab_excel = st.tabs(["✍️ Nhập thủ công", "📁 Tải file Excel"])
    
    with tab_thu_cong:
        with st.form("add_sample_form", clear_on_submit=True):
            new_id = st.text_input("Mã Mẫu (VD: NT-1509-01)*")
            new_name = st.text_input("Tên/Ký hiệu Mẫu")
            new_nen = st.selectbox("Nền Mẫu", ["Nước", "Khí"])
            new_chi_tieu = st.text_input("Chỉ tiêu đo (VD: VOCs, Formaldehyde)")
            new_nguoi = st.text_input("Người tiếp nhận (Ký tên)")
            
            if st.form_submit_button("Thêm Mẫu") and new_id:
                new_row = pd.DataFrame([{
                    "Mã Mẫu": new_id, "Tên Mẫu": new_name, "Nền Mẫu": new_nen, 
                    "Chỉ Tiêu": new_chi_tieu, "Trạng Thái": STATUSES[0], 
                    "Người Giữ": new_nguoi, "Ghi Chú": "", "Giờ Nhận": datetime.now() 
                }])
                st.session_state.df = pd.concat([st.session_state.df, new_row], ignore_index=True)
                save_data(st.session_state.df)
                st.success(f"Đã thêm mẫu {new_id} thành công!")
                st.rerun()

    with tab_excel:
        st.info("💡 Hệ thống sẽ tự động quét file và lọc lấy các mã mẫu hợp lệ (KHM).")
        uploaded_file = st.file_uploader("Kéo thả file KetQuaMeThuNghiem...xlsx vào đây", type=["xlsx", "xls"])
        
        if uploaded_file is not None:
            try:
                df_upload = pd.read_excel(uploaded_file, sheet_name=0)
                khm_col, start_row = None, None
                
                # Quét 20 dòng đầu để tìm cột chứa chữ KHM
                for row_idx in range(min(20, len(df_upload))):
                    row_vals = df_upload.iloc[row_idx].values
                    for col_idx, val in enumerate(row_vals):
                        if str(val).strip() == 'KHM':
                            khm_col = df_upload.columns[col_idx]
                            start_row = row_idx + 1
                            break
                    if khm_col: break
                        
                if khm_col and start_row is not None:
                    # Lấy danh sách mẫu, loại bỏ ô trống hoặc rác (chỉ lấy mã > 3 ký tự)
                    raw_samples = df_upload[khm_col].iloc[start_row:].dropna().astype(str).tolist()
                    samples = [s for s in raw_samples if len(s) > 3 and s.lower() != 'nan']
                    
                    if len(samples) > 0:
                        st.success(f"✔️ Đã quét thành công **{len(samples)}** mẫu: {', '.join(samples)}")
                        
                        # Khai báo thông tin chung cho cả lô file Excel này
                        batch_nen = st.selectbox("Nền mẫu chung cho lô này:", ["Nước", "Khí"])
                        batch_chitieu = st.text_input("Chỉ tiêu chung cho lô:", "Chưa xác định")
                        batch_nguoi = st.text_input("Người tiếp nhận (Ký tên):")
                        
                        if st.button("🚀 Thêm hàng loạt vào Hệ thống"):
                            new_rows = []
                            for s in samples:
                                new_rows.append({
                                    "Mã Mẫu": s, "Tên Mẫu": "", "Nền Mẫu": batch_nen, 
                                    "Chỉ Tiêu": batch_chitieu, "Trạng Thái": STATUSES[0], 
                                    "Người Giữ": batch_nguoi, "Ghi Chú": "Import từ Excel", 
                                    "Giờ Nhận": datetime.now()
                                })
                            st.session_state.df = pd.concat([st.session_state.df, pd.DataFrame(new_rows)], ignore_index=True)
                            save_data(st.session_state.df)
                            st.success(f"Tuyệt vời! Đã nạp thành công {len(samples)} mẫu.")
                            st.rerun()
                    else:
                        st.warning("Không tìm thấy mã mẫu nào bên dưới cột 'KHM'.")
                else:
                    st.error("Không tìm thấy ô 'KHM' (Ký hiệu mẫu) trong file. Vui lòng kiểm tra lại!")
            except Exception as e:
                st.error(f"Lỗi đọc file: {e}")

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
