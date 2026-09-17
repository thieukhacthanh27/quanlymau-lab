import streamlit as st
import pandas as pd
from datetime import datetime
import re
import difflib
from streamlit_gsheets import GSheetsConnection

# ==========================================
# 1. CẤU HÌNH GIAO DIỆN & BRANDING
# ==========================================
st.set_page_config(page_title="GC HATICO - Lab GC", page_icon="🔬", layout="wide")

SHEET_URL = "https://docs.google.com/spreadsheets/d/1F2wFnxboWTFWDMGUuBDRGB901a5EKgvazHxkCgBjjRU/edit?usp=sharing"

STATUSES = [
    "🔴 1. Chờ xử lý", "🟠 2. Đang xử lý mẫu", "🟡 3. Chờ chạy máy",
    "🔵 4. Đang chạy máy", "🟣 5. Đang tính số liệu", "🟢 6. Lưu kho", "⚫ 7. Đã tiêu hủy"
]

# ==========================================
# 2. KẾT NỐI DATABASE & THƯ VIỆN MDL
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

def load_mdl_config():
    """Tải thư viện MDL từ Tab CauHinh_MDL"""
    try:
        df_mdl = conn.read(spreadsheet=SHEET_URL, worksheet="CauHinh_MDL", ttl=60)
        return df_mdl
    except:
        return pd.DataFrame(columns=["Nền Mẫu", "Tên Chất", "MDL", "Đơn Vị"])

if "df" not in st.session_state:
    st.session_state.df = load_data()
if "df_mdl" not in st.session_state:
    st.session_state.df_mdl = load_mdl_config()

# ==========================================
# 3. HÀM BỔ SUNG: XỬ LÝ SỐ LIỆU & ĐÁNH GIÁ
# ==========================================
def parse_sample_matrix(sample_name):
    """Bộ lọc tự động nhận diện thể tích và Ký hiệu nền mẫu dựa trên mã."""
    name_upper = str(sample_name).upper()
    if 'KT' in name_upper:
        return 24.0, 'KT'
    elif 'KXQ' in name_upper:
        return 4.0, 'KXQ'
    elif 'KLV' in name_upper:
        return 4.0, 'KLV'
    return None, None

def evaluate_result(raw_conc, v_gas, compound_name, nen_mau, v_desorb=1.0, recovery=100.0):
    """Tính toán thực tế và tự động so sánh với Thư viện MDL"""
    if pd.isna(raw_conc) or raw_conc <= 0:
        return "KPH"
        
    result = (raw_conc * v_desorb) / v_gas * (100.0 / recovery)
    
    df_mdl = st.session_state.df_mdl
    if not df_mdl.empty and compound_name and nen_mau:
        mask = (df_mdl["Nền Mẫu"].astype(str).str.upper() == nen_mau.upper()) & \
               (df_mdl["Tên Chất"].astype(str).str.lower() == str(compound_name).lower())
        match = df_mdl[mask]
        
        if not match.empty:
            mdl_str = str(match["MDL"].values[0]).replace(',', '.')
            try:
                mdl_val = float(mdl_str)
                unit = str(match["Đơn Vị"].values[0])
                if pd.isna(unit) or unit == 'nan': unit = ""
                
                if result < mdl_val:
                    return f"KPH (MDL={mdl_val} {unit})"
                else:
                    return f"{round(result, 4)} {unit}"
            except:
                pass
                
    return round(result, 4)

# ==========================================
# 4. THANH ĐIỀU HƯỚNG BÊN TRÁI (SIDEBAR)
# ==========================================
st.sidebar.title("🔬 LIMS HATICO")
st.sidebar.caption("Phần mềm Quản lý Phòng Lab GC-MS")
st.sidebar.divider()

menu = st.sidebar.radio("📌 ĐIỀU HƯỚNG CHÍNH", [
    "🏠 Trang chủ (Tổng quan)", 
    "📥 Quản lý Tiếp nhận", 
    "⚙️ Vận hành GC-MS",
    "🚀 Tiện ích & Cấu hình"
])

st.sidebar.divider()
st.sidebar.markdown("**Hỗ trợ nhanh:**")
if st.sidebar.button("🔄 Cập nhật dữ liệu tức thì"):
    st.cache_data.clear()
    st.session_state.df = load_data()
    st.session_state.df_mdl = load_mdl_config() 
    st.rerun()

df_current = st.session_state.df.copy()
df_current["Ngày Nhận"] = df_current["Giờ Nhận"].dt.date
today_date = datetime.today().date()

# ==========================================
# 5. GIAO DIỆN CÁC TRANG
# ==========================================

if menu == "🏠 Trang chủ (Tổng quan)":
    st.title("📊 Bảng Điều Khiển Trung Tâm")
    
    col1, col2, col3, col4 = st.columns(4)
    tong_hom_nay = len(df_current[df_current["Ngày Nhận"] == today_date])
    cho_chay_may = len(df_current[df_current["Trạng Thái"] == "🟡 3. Chờ chạy máy"])
    ton_dong = len(df_current[(df_current["Ngày Nhận"] < today_date) & (~df_current["Trạng Thái"].isin(["🟢 6. Lưu kho", "⚫ 7. Đã tiêu hủy"]))])
    
    # Cập nhật Logic đếm mẫu hoàn thành, lưu kho và tiêu hủy
    da_luu = len(df_current[df_current["Trạng Thái"] == "🟢 6. Lưu kho"])
    da_huy = len(df_current[df_current["Trạng Thái"] == "⚫ 7. Đã tiêu hủy"])
    tong_hoan_thanh = da_luu + da_huy
    
    col1.metric("📥 Tổng nhận hôm nay", tong_hom_nay)
    col2.metric("⏳ Đang chờ chạy GC", cho_chay_may)
    col3.metric("⚠️ Tồn đọng chưa xử lý", ton_dong, delta="-Cần xử lý", delta_color="inverse")
    col4.metric("✅ Đã hoàn thành", tong_hoan_thanh, f"Lưu kho: {da_luu} | Hủy: {da_huy}", delta_color="off")
    
    st.divider()
    
    col_date, col_status, col_search = st.columns([1, 1.5, 2])
    with col_date: selected_date = st.date_input("📅 Chọn Ngày:", today_date)
    with col_status: filter_status = st.multiselect("Lọc trạng thái:", STATUSES, default=[])
    with col_search: search_query = st.text_input("🔍 Tìm kiếm (Mã mẫu, Tên mẻ, Chỉ tiêu):", placeholder="Gõ 'VOCs', '2026.07.017' hoặc 'NS...'")

    st.subheader(f"📋 Danh sách công việc ({selected_date.strftime('%d/%m/%Y')})")

    mask_ton_dong = (df_current["Ngày Nhận"] < selected_date) & (~df_current["Trạng Thái"].isin(["🟢 6. Lưu kho", "⚫ 7. Đã tiêu hủy"]))
    mask_trong_ngay = (df_current["Ngày Nhận"] == selected_date)

    df_display = df_current[mask_ton_dong | mask_trong_ngay].copy()
    df_display["Phân Loại"] = "🟢 Nhận trong ngày"
    df_display.loc[mask_ton_dong, "Phân Loại"] = "⚠️ TỒN ĐỌNG CHƯA XONG"
    df_display = df_display.sort_values(by=["Phân Loại", "Giờ Nhận"], ascending=[True, True])

    if search_query:
        mask_id = df_display["Mã Mẫu"].astype(str).str.contains(search_query, case=False, na=False)
        mask_me = df_display["Tên Mẻ"].astype(str).str.contains(search_query, case=False, na=False)
        mask_chitieu = df_display["Chỉ Tiêu"].astype(str).str.contains(search_query, case=False, na=False)
        df_display = df_display[mask_id | mask_me | mask_chitieu]
        
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
        use_container_width=True, num_rows="dynamic", key="data_editor", height=400
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
                        if header_val.lower() == 'ghi chú' or header_val == 'nan' or header_val == '': break
                        params_info.append((c, header_val))
                    
                    samples_data = []
                    for r in range(header_row_idx + 1, len(df_upload)):
                        khm_val = str(df_upload.iloc[r, khm_col_idx]).strip()
                        if len(khm_val) > 3 and khm_val.lower() != 'nan':
                            khm_upper = khm_val.upper()
                            if khm_upper.startswith(("KT", "KKXQ", "KLV")): nen_mau_auto = "Khí"
                            elif khm_upper.startswith(("NS", "NT", "NM", "NU")): nen_mau_auto = "Nước"
                            else: nen_mau_auto = "Chưa xác định"

                            sample_params = []
                            for c, param_name in params_info:
                                cell_val = df_upload.iloc[r, c]
                                if pd.notna(cell_val) and str(cell_val).strip() != '':
                                    sample_params.append(param_name)
                            
                            chuoi_chi_tieu = ", ".join(sample_params) if sample_params else "Chưa xác định"
                            samples_data.append({
                                "Chọn": True, "Mã Mẫu": khm_val, "Nền Mẫu": nen_mau_auto, "Chỉ Tiêu": chuoi_chi_tieu
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
                            hide_index=True, use_container_width=True, key="preview_editor"
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
                                        "Mã Mẫu": row["Mã Mẫu"], "Tên Mẻ": ten_me_extract, "Nền Mẫu": row["Nền Mẫu"], 
                                        "Chỉ Tiêu": row["Chỉ Tiêu"], "Trạng Thái": STATUSES[0], "Người Giữ": batch_nguoi, 
                                        "Ghi Chú": "Import từ Excel", "Giờ Nhận": datetime.now()
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
            
            seq_df['Method'] = df_ready['Chỉ Tiêu'].apply(lambda x: 'VOCs.M' if any(k in str(x).upper() for k in ['VOC', 'BENZEN', 'TOLUEN', 'CHLORO', 'STYREN']) else 'HCHO.M')
            seq_df['Data File'] = datetime.now().strftime("%Y%m%d") + "_" + df_ready['Mã Mẫu']
            
            csv = seq_df.to_csv(index=False).encode('utf-8')
            st.download_button("📥 Tải Sequence.csv", data=csv, file_name=f"MassHunter_Seq_{datetime.now().strftime('%Y%m%d')}.csv", mime="text/csv", type="primary")
            
    with col_import:
        st.subheader("2. Xử lý dữ liệu GC-MS (Tự động tính & So sánh MDL)")
        st.info("💡 Web sẽ tự động dùng thông số từ Thư viện MDL để kết luận KPH cho mẫu.")

        gc_file = st.file_uploader("Kéo thả báo cáo kết quả GC (PDF/Excel/CSV)", type=["pdf", "xlsx", "xls", "csv"])

        if gc_file is not None:
            calc_results = []
            try:
                # ĐỌC FILE PDF
                if gc_file.name.endswith('.pdf'):
                    import PyPDF2
                    reader = PyPDF2.PdfReader(gc_file)
                    text = ""
                    for page in reader.pages: text += page.extract_text() + "\n"

                    lines = text.split('\n')
                    current_compound = None
                    known_compounds = ['Benzene', 'Toluene-D8', 'Toluene', 'Ethylbenzene', 'm-Xylene', 'p-Xylene', 'o-Xylene', 'Styrene', 'Acetaldehyde', 'Formaldehyde']

                    for line in lines:
                        parts = line.split()
                        if not parts: continue

                        if line.strip() in known_compounds:
                            current_compound = line.strip()
                            continue

                        if len(parts) >= 5 and parts[0].endswith('.d') and 'Sample' in parts:
                            data_file = parts[0].replace('.d', '')
                            floats = [float(p) for p in parts if p.replace('.', '', 1).isdigit() and p.count('.') <= 1]

                            if len(floats) >= 3 and current_compound:
                                final_conc = floats[-1]
                                v_gas, nen_mau = parse_sample_matrix(data_file)
                                if v_gas is not None:
                                    final_result_str = evaluate_result(final_conc, v_gas, current_compound, nen_mau)
                                    calc_results.append({
                                        "Mã Mẫu": data_file,
                                        "Nền Mẫu": nen_mau,
                                        "Tên Chất": current_compound,
                                        "Nồng độ GC (ng/ml)": final_conc,
                                        "Kết quả Thực Tế": final_result_str
                                    })

                # ĐỌC FILE EXCEL / CSV
                else:
                    if gc_file.name.endswith('.csv'): df_gc = pd.read_csv(gc_file)
                    else: df_gc = pd.read_excel(gc_file)

                    df_gc.columns = [str(c).strip() for c in df_gc.columns]

                    if 'Data File' in df_gc.columns and 'Final Conc.' in df_gc.columns:
                        compound_col = next((c for c in df_gc.columns if c.lower() in ['name', 'compound', 'compound name', 'tên chất']), None)

                        for _, row in df_gc.iterrows():
                            sample_name = str(row['Data File']).replace('.d', '')
                            raw_conc = pd.to_numeric(row['Final Conc.'], errors='coerce')
                            if pd.isna(raw_conc): continue

                            v_gas, nen_mau = parse_sample_matrix(sample_name)
                            if v_gas is not None:
                                comp_name = row[compound_col] if compound_col else "N/A"
                                final_result_str = evaluate_result(raw_conc, v_gas, comp_name, nen_mau)
                                calc_results.append({
                                    "Mã Mẫu": sample_name, "Nền Mẫu": nen_mau, "Tên Chất": comp_name,
                                    "Nồng độ GC (ng/ml)": raw_conc, "Kết quả Thực Tế": final_result_str
                                })
                    else: st.error("❌ File Excel thiếu cột 'Data File' hoặc 'Final Conc.'")

                # HIỂN THỊ KẾT QUẢ ĐÃ ĐÁNH GIÁ MDL
                if calc_results:
                    st.success(f"✔️ Đã xử lý & đối chiếu MDL thành công {len(calc_results)} dữ liệu!")
                    st.dataframe(pd.DataFrame(calc_results), use_container_width=True)
                elif gc_file.name.endswith('.pdf'): st.warning("⚠️ Không tìm thấy thông tin hợp lệ trong PDF.")
            except Exception as e: st.error(f"❌ Lỗi khi đọc file: {e}")

# ---------------------------------------------------------
elif menu == "🚀 Tiện ích & Cấu hình":
    st.title("🛠️ Tiện ích & Cấu hình Hệ thống")
    
    tab_mdl, tab_report, tab_qr = st.tabs(["📚 Quản lý Thư viện MDL", "📝 Lập Biên Bản (Sắp ra mắt)", "🏷️ Sinh Mã QR"])
    
    with tab_mdl:
        st.subheader("Tra cứu thông minh Thư viện MDL")
        
        # 1. KHU VỰC TRA CỨU NHANH (FUZZY SEARCH)
        search_mdl = st.text_input("🔍 Tra cứu MDL sát tên chất (VD: Benzen, Toluen, Xylen):")
        
        mdl_library = st.session_state.df_mdl
        if search_mdl:
            if not mdl_library.empty:
                all_compounds = mdl_library["Tên Chất"].astype(str).unique()
                
                # Thuật toán tìm kiếm sát nghĩa (Fuzzy Matching)
                close_matches = difflib.get_close_matches(search_mdl.lower(), [c.lower() for c in all_compounds], n=10, cutoff=0.4)
                
                # Lọc kết quả: Chứa một phần từ khóa HOẶC nằm trong danh sách khớp sát nghĩa
                mask_contains = mdl_library["Tên Chất"].astype(str).str.contains(search_mdl, case=False, na=False)
                mask_fuzzy = mdl_library["Tên Chất"].astype(str).str.lower().isin(close_matches)
                
                search_result = mdl_library[mask_contains | mask_fuzzy]
                
                if not search_result.empty:
                    st.dataframe(search_result, use_container_width=True, hide_index=True)
                else:
                    st.warning("⚠️ Không tìm thấy chất nào sát với từ khóa bạn nhập trong Thư viện.")
            else:
                st.info("Thư viện MDL hiện đang trống.")
        else:
            if not mdl_library.empty:
                with st.expander("Xem toàn bộ Thư viện MDL hiện tại", expanded=False):
                    st.dataframe(mdl_library, use_container_width=True, hide_index=True)

        st.divider()

        # 2. KHU VỰC CẬP NHẬT THƯ VIỆN
        st.subheader("Cập nhật Thư viện Giới hạn Phát hiện (MDL)")
        st.info("Kéo thả file Excel chứa bảng MDL (như file `Bang_MDL_Khi_Cac_Loai...xlsx`). Hệ thống sẽ tự quét các Sheet và gộp lại rồi bắn lên Google Sheets.")
        
        mdl_file = st.file_uploader("Tải lên file Excel Bảng MDL", type=["xlsx"])
        if mdl_file:
            try:
                xls = pd.ExcelFile(mdl_file)
                mdl_data = []
                
                for sheet in xls.sheet_names:
                    df_sheet = pd.read_excel(xls, sheet_name=sheet)
                    
                    col_ten = next((c for c in df_sheet.columns if 'tên' in str(c).lower() or 'hợp chất' in str(c).lower()), None)
                    col_mdl = next((c for c in df_sheet.columns if 'mdl' in str(c).lower()), None)
                    
                    if col_ten and col_mdl:
                        unit_match = re.search(r'\((.*?)\)', str(col_mdl))
                        unit = unit_match.group(1) if unit_match else "Chưa rõ"
                        nen_mau = "KT" if "thải" in sheet.lower() else ("KXQ" if "xung quanh" in sheet.lower() else ("KLV" if "làm việc" in sheet.lower() else sheet))
                        
                        for _, row in df_sheet.iterrows():
                            if pd.notna(row[col_ten]) and pd.notna(row[col_mdl]):
                                mdl_data.append({
                                    "Nền Mẫu": nen_mau,
                                    "Tên Chất": str(row[col_ten]).strip(),
                                    "MDL": str(row[col_mdl]).replace(',', '.').strip(),
                                    "Đơn Vị": unit
                                })
                
                if mdl_data:
                    df_mdl_new = pd.DataFrame(mdl_data)
                    st.success(f"✔️ Đã quét được {len(df_mdl_new)} chỉ tiêu MDL từ file.")
                    st.dataframe(df_mdl_new, use_container_width=True)
                    
                    if st.button("🚀 Đẩy lên Google Sheets (Tạo/Ghi đè Thư viện)", type="primary"):
                        conn.update(spreadsheet=SHEET_URL, worksheet="CauHinh_MDL", data=df_mdl_new)
                        st.session_state.df_mdl = df_mdl_new
                        st.success("🎉 Đã lưu cấu hình MDL lên Cloud thành công! Tự động áp dụng cho các mẻ tính toán sau.")
                else: st.error("Không tìm thấy cấu trúc bảng hợp lệ (Cột Tên hợp chất / Cột MDL).")
            except Exception as e: st.error(f"Lỗi đọc file: {e}")

    with tab_report: st.write("Khu vực xuất Form Word/Excel.")
    with tab_qr: st.write("Chức năng tạo mã vạch (Barcode) hàng loạt cho các mẫu mẻ mới nhận.")
