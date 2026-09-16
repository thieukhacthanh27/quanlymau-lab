import streamlit as st
import pandas as pd
from datetime import datetime
from streamlit_gsheets import GSheetsConnection

# ==========================================
# 1. CẤU HÌNH GIAO DIỆN & BRANDING
# ==========================================
st.set_page_config(page_title="LIMS HATICO - Lab GC", page_icon="🔬", layout="wide")

SHEET_URL = "https://docs.google.com/spreadsheets/d/1F2wFnxboWTFWDMGUuBDRGB901a5EKgvazHxkCgBjjRU/edit?usp=sharing"

STATUSES = [
    "🔴 1. Chờ xử lý", "🟠 2. Đang xử lý mẫu", "🟡 3. Chờ chạy máy",
    "🔵 4. Đang chạy máy", "🟣 5. Đang tính số liệu", "🟢 6. Lưu kho", "⚫ 7. Đã tiêu hủy"
]

# ==========================================
# 2. KẾT NỐI & XỬ LÝ DỮ LIỆU
# ==========================================
conn = st.connection("gsheets", type=GSheetsConnection)

def load_data():
    df = conn.read(spreadsheet=SHEET_URL, ttl=0)
    if df.empty or len(df.columns) == 0 or "Mã Mẫu" not in df.columns:
        df = pd.DataFrame(columns=["Mã Mẫu", "Tên Mẻ", "Nền Mẫu", "Chỉ Tiêu", "Trạng Thái", "Người Giữ", "Ghi Chú", "Giờ Nhận"])
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
# 3. THANH ĐIỀU HƯỚNG BÊN TRÁI (SIDEBAR)
# ==========================================
st.sidebar.title("🔬 LIMS HATICO")
st.sidebar.caption("Phần mềm Quản lý Phòng Lab GC-MS")
st.sidebar.divider()

menu = st.sidebar.radio("📌 ĐIỀU HƯỚNG CHÍNH", [
    "🏠 Trang chủ (Tổng quan)", 
    "📥 Quản lý Tiếp nhận", 
    "⚙️ Vận hành GC-MS",
    "🚀 Không gian phát triển"
])

st.sidebar.divider()
st.sidebar.markdown("**Hỗ trợ nhanh:**")
if st.sidebar.button("🔄 Cập nhật dữ liệu tức thì"):
    st.cache_data.clear()
    st.session_state.df = load_data()
    st.rerun()

# --- Tính toán Logic chung ---
df_current = st.session_state.df.copy()
df_current["Ngày Nhận"] = df_current["Giờ Nhận"].dt.date
today_date = datetime.today().date()

# ==========================================
# 4. GIAO DIỆN CÁC TRANG
# ==========================================

if menu == "🏠 Trang chủ (Tổng quan)":
    st.title("📊 Bảng Điều Khiển Trung Tâm")
    
    # --- THỐNG KÊ NHANH (METRICS) ---
    col1, col2, col3, col4 = st.columns(4)
    tong_hom_nay = len(df_current[df_current["Ngày Nhận"] == today_date])
    cho_chay_may = len(df_current[df_current["Trạng Thái"] == "🟡 3. Chờ chạy máy"])
    da_hoan_thanh = len(df_current[(df_current["Ngày Nhận"] == today_date) & (df_current["Trạng Thái"].isin(["🟢 6. Lưu kho", "⚫ 7. Đã tiêu hủy"]))])
    ton_dong = len(df_current[(df_current["Ngày Nhận"] < today_date) & (~df_current["Trạng Thái"].isin(["🟢 6. Lưu kho", "⚫ 7. Đã tiêu hủy"]))])
    
    col1.metric("Tổng mẫu nhận hôm nay", tong_hom_nay, f"Đã xong: {da_hoan_thanh}")
    col2.metric("Đang chờ chạy máy GC", cho_chay_may)
    col3.metric("⚠️ Tồn đọng nguy cấp", ton_dong, delta="-Cần xử lý", delta_color="inverse")
    
    st.divider()
    
    # --- BỘ LỌC TÌM KIẾM ---
    col_date, col_search, col_filter = st.columns([1.5, 1, 1.5])
    with col_date:
        selected_date = st.date_input("📅 Chọn Ngày Làm Việc:", today_date)
    with col_search:
        search_query = st.text_input("🔍 Nhập ID mã mẫu:")
    with col_filter:
        filter_status = st.multiselect("Lọc theo trạng thái:", STATUSES, default=[])

    st.subheader(f"📋 Danh sách công việc ({selected_date.strftime('%d/%m/%Y')})")

    # --- LỌC BẢNG DỮ LIỆU ---
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

    st.caption("Mẹo: Chọn dòng và bấm Delete trên bàn phím để xóa. Sau khi chỉnh sửa, hãy bấm nút Lưu bên dưới.")
    edited_df = st.data_editor(
        df_display,
        column_config={
            "Trạng Thái": st.column_config.SelectboxColumn("Trạng Thái Hiện Tại", options=STATUSES, required=True),
            "Nền Mẫu": st.column_config.SelectboxColumn("Nền Mẫu", options=["Khí", "Nước"], required=True),
            "Phân Loại": st.column_config.TextColumn("Phân Loại", disabled=True),
            "Giờ Nhận": st.column_config.DatetimeColumn("Giờ Nhận", format="DD/MM/YYYY HH:mm", disabled=True),
            "Ngày Nhận": None 
        },
        disabled=["Mã Mẫu", "Tên Mẻ", "Chỉ Tiêu", "Phân Loại", "Giờ Nhận"], 
        use_container_width=True,
        num_rows="dynamic",
        key="data_editor",
        height=400
    )

    if st.button("💾 Lưu các thay đổi vào Hệ thống", type="primary"):
        for index, row in edited_df.iterrows():
            st.session_state.df.loc[index, "Trạng Thái"] = row["Trạng Thái"]
            st.session_state.df.loc[index, "Nền Mẫu"] = row["Nền Mẫu"]
            st.session_state.df.loc[index, "Người Giữ"] = row["Người Giữ"]
            st.session_state.df.loc[index, "Ghi Chú"] = row["Ghi Chú"]
            
        original_indices = df_display.index.tolist()
        remaining_indices = edited_df.index.tolist()
        deleted_indices = list(set(original_indices) - set(remaining_indices))
        
        if deleted_indices:
            st.session_state.df = st.session_state.df.drop(index=deleted_indices)
            st.session_state.df = st.session_state.df.reset_index(drop=True)
            
        save_data(st.session_state.df)
        st.success("Đã đồng bộ lên cơ sở dữ liệu chung!")
        st.rerun()

# ---------------------------------------------------------
elif menu == "📥 Quản lý Tiếp nhận":
    st.title("📥 Khu vực Tiếp nhận mẫu mới")
    
    tab_excel, tab_thu_cong = st.tabs(["📁 Tải file Excel tự động", "✍️ Nhập thủ công (Mẫu lẻ)"])
    
    with tab_excel:
        st.info("💡 Kéo thả file Excel để hệ thống tự động trích xuất Tên Mẻ, Mã Mẫu, **Nhận diện Nền Mẫu** và Chỉ Tiêu.")
        uploaded_file = st.file_uploader("Kéo thả file KetQuaMeThuNghiem...xlsx vào đây", type=["xlsx", "xls"])
        
        if uploaded_file is not None:
            try:
                df_upload = pd.read_excel(uploaded_file, sheet_name=0)
                
                # 1. Quét Tên Mẻ
                ten_me_extract = "Không xác định"
                for r in range(min(5, len(df_upload))):
                    for c in range(len(df_upload.columns)):
                        val = str(df_upload.iloc[r, c]).strip()
                        if val.startswith("Số:"):
                            ten_me_extract = val.replace("Số:", "").strip()
                            break
                    if ten_me_extract != "Không xác định": break
                
                if ten_me_extract == "Không xác định":
                    for col in df_upload.columns:
                        if str(col).startswith("Số:"):
                            ten_me_extract = str(col).replace("Số:", "").strip()
                            break
                
                # 2. Tìm KHM và Trích xuất danh sách Chỉ Tiêu
                khm_col_idx, header_row_idx = None, None
                for r in range(min(20, len(df_upload))):
                    for c in range(len(df_upload.columns)):
                        if str(df_upload.iloc[r, c]).strip() == 'KHM':
                            khm_col_idx = c
                            header_row_idx = r
                            break
                    if khm_col_idx is not None: break
                        
                if khm_col_idx is not None:
                    params_info = []
                    for c in range(khm_col_idx + 1, len(df_upload.columns)):
                        header_val = str(df_upload.iloc[header_row_idx, c]).strip()
                        if header_val.lower() == 'ghi chú' or header_val == 'nan' or header_val == '':
                            break
                        params_info.append((c, header_val))
                    
                    # 3. Nhặt thông tin mẫu và TỰ ĐỘNG PHÂN LOẠI NỀN MẪU
                    samples_data = []
                    for r in range(header_row_idx + 1, len(df_upload)):
                        khm_val = str(df_upload.iloc[r, khm_col_idx]).strip()
                        if len(khm_val) > 3 and khm_val.lower() != 'nan':
                            
                            khm_upper = khm_val.upper()
                            if khm_upper.startswith(("KT", "KKXQ", "KLV")):
                                nen_mau_auto = "Khí"
                            elif khm_upper.startswith(("NS", "NT", "NM", "NU")):
                                nen_mau_auto = "Nước"
                            else:
                                nen_mau_auto = "Chưa xác định"

                            sample_params = []
                            for c, param_name in params_info:
                                cell_val = df_upload.iloc[r, c]
                                if pd.notna(cell_val) and str(cell_val).strip() != '':
                                    sample_params.append(param_name)
                            
                            chuoi_chi_tieu = ", ".join(sample_params) if sample_params else "Chưa xác định"
                            
                            samples_data.append({
                                "Chọn": True, 
                                "Mã Mẫu": khm_val,
                                "Nền Mẫu": nen_mau_auto, 
                                "Chỉ Tiêu": chuoi_chi_tieu
                            })
                    
                    if len(samples_data) > 0:
                        st.success(f"✔️ Quét thành công **{len(samples_data)}** mẫu thuộc mẻ: **{ten_me_extract}**")
                        
                        st.caption("☑️ Máy đã tự đoán Nền Mẫu. Nếu sai, bạn có thể click vào ô để chọn lại trước khi lưu.")
                        
                        df_preview = pd.DataFrame(samples_data)
                        edited_preview = st.data_editor(
                            df_preview,
                            column_config={
                                "Chọn": st.column_config.CheckboxColumn("Nhập mẫu?", default=True),
                                "Mã Mẫu": st.column_config.TextColumn("Mã Mẫu", disabled=True),
                                "Nền Mẫu": st.column_config.SelectboxColumn("Nền Mẫu", options=["Nước", "Khí", "Chưa xác định"]),
                                "Chỉ Tiêu": st.column_config.TextColumn("Chỉ Tiêu", disabled=True)
                            },
                            hide_index=True,
                            use_container_width=True,
                            key="preview_editor"
                        )
                        
                        batch_nguoi = st.selectbox("Người tiếp nhận:", ["Thành", "Kỹ thuật viên 2", "Kỹ thuật viên 3"])
                        
                        selected_samples = edited_preview[edited_preview["Chọn"] == True]
                        
                        if st.button(f"🚀 Lưu {len(selected_samples)} mẫu đã chọn vào Hệ thống", type="primary"):
                            if selected_samples.empty:
                                st.warning("⚠️ Bạn chưa chọn mẫu nào để lưu!")
                            else:
                                new_rows = []
                                for _, row in selected_samples.iterrows():
                                    new_rows.append({
                                        "Mã Mẫu": row["Mã Mẫu"], 
                                        "Tên Mẻ": ten_me_extract, 
                                        "Nền Mẫu": row["Nền Mẫu"], 
                                        "Chỉ Tiêu": row["Chỉ Tiêu"], 
                                        "Trạng Thái": STATUSES[0], 
                                        "Người Giữ": batch_nguoi, 
                                        "Ghi Chú": "Import từ Excel", 
                                        "Giờ Nhận": datetime.now()
                                    })
                                st.session_state.df = pd.concat([st.session_state.df, pd.DataFrame(new_rows)], ignore_index=True)
                                save_data(st.session_state.df)
                                st.success(f"Tuyệt vời! Đã nạp thành công {len(selected_samples)} mẫu.")
                                st.rerun()
                    else: st.warning("Không có mã KHM nào hợp lệ bên dưới ô tiêu đề.")
                else: st.error("Không tìm thấy ô 'KHM' trong file!")
            except Exception as e: st.error(f"Lỗi đọc file: {e}")

    with tab_thu_cong:
        with st.form("add_sample_form", clear_on_submit=True):
            new_id = st.text_input("Mã Mẫu (VD: NT-1509-01)*")
            new_name = st.text_input("Tên Mẻ (VD: 2026.07.017)")
            col_t1, col_t2 = st.columns(2)
            with col_t1: new_nen = st.selectbox("Nền Mẫu", ["Nước", "Khí"])
            with col_t2: new_chi_tieu = st.text_input("Chỉ tiêu đo")
            new_nguoi = st.text_input("Người tiếp nhận (Ký tên)")
            
            if st.form_submit_button("Thêm Mẫu lẻ") and new_id:
                new_row = pd.DataFrame([{
                    "Mã Mẫu": new_id, "Tên Mẻ": new_name, "Nền Mẫu": new_nen, 
                    "Chỉ Tiêu": new_chi_tieu, "Trạng Thái": STATUSES[0], 
                    "Người Giữ": new_nguoi, "Ghi Chú": "", "Giờ Nhận": datetime.now() 
                }])
                st.session_state.df = pd.concat([st.session_state.df, new_row], ignore_index=True)
                save_data(st.session_state.df)
                st.success(f"Đã thêm {new_id}!")

# ---------------------------------------------------------
elif menu == "⚙️ Vận hành GC-MS":
    st.title("⚙️ Điều phối & Vận hành Máy đo")
    
    col_seq, col_import = st.columns(2)
    with col_seq:
        st.subheader("1. Xuất Sequence chạy máy")
        st.info("Gom tự động các mẫu đang ở trạng thái '🟡 3. Chờ chạy máy'.")
        
        df_ready = st.session_state.df[st.session_state.df["Trạng Thái"] == "🟡 3. Chờ chạy máy"]
        st.write(f"Hiện đang có **{len(df_ready)}** mẫu chờ chạy.")
        
        if not df_ready.empty:
            seq_df = pd.DataFrame()
            seq_df['Vial'] = range(1, len(df_ready) + 1)
            seq_df['Sample Name'] = df_ready['Mã Mẫu']
            seq_df['Sample Type'] = 'Sample'
            
            # Logic mới: Tự động phát hiện phương pháp dựa vào tên hóa chất dài
            seq_df['Method'] = df_ready['Chỉ Tiêu'].apply(lambda x: 'VOCs.M' if any(k in str(x).upper() for k in ['VOC', 'BENZEN', 'TOLUEN', 'CHLORO', 'STYREN']) else 'HCHO.M')
            
            seq_df['Data File'] = datetime.now().strftime("%Y%m%d") + "_" + df_ready['Mã Mẫu']
            
            csv = seq_df.to_csv(index=False).encode('utf-8')
            st.download_button("📥 Tải Sequence.csv", data=csv, file_name=f"MassHunter_Seq_{datetime.now().strftime('%Y%m%d')}.csv", mime="text/csv", type="primary")
            
    with col_import:
        st.subheader("2. Xử lý dữ liệu GC (Chờ lập trình)")
        st.markdown(
            """
            *Khu vực này được chuẩn bị sẵn cho giai đoạn 2:*
            *   Tải file kết quả từ MassHunter lên.
            *   Tự động tính toán hàm lượng dựa trên diện tích Peak.
            *   Khớp dữ liệu nồng độ trực tiếp vào Google Sheets.
            """
        )
        st.file_uploader("Kéo thả báo cáo MassHunter vào đây (Tính năng chưa kích hoạt)", disabled=True)

# ---------------------------------------------------------
elif menu == "🚀 Không gian phát triển":
    st.title("🛠️ Tiện ích & Mở rộng tương lai")
    
    st.markdown("Khu vực này để dành cho bạn tự do sáng tạo và code thêm các module tự động hóa khác cho Lab:")
    
    tab1, tab2, tab3 = st.tabs(["📝 In Phiếu Kết Quả", "🧪 Quản lý Hóa chất / Chuẩn", "🏷️ Sinh Mã QR"])
    
    with tab1:
        st.write("Tại đây, bạn có thể lập trình code đọc kết quả từ bảng và điền tự động vào Form Word/Excel mẫu của công ty, sau đó xuất ra file PDF.")
    with tab2:
        st.write("Khu vực theo dõi tồn kho hóa chất, hạn sử dụng dung dịch chuẩn, tự động cảnh báo khi sắp hết hạn.")
    with tab3:
        st.write("Chức năng tạo mã vạch (Barcode) hàng loạt cho các mẫu mẻ mới nhận để dán lên khay lưu kho.")
