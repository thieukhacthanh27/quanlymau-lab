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
# 2. KẾT NỐI DATABASE & THƯ VIỆN KÉP (MDL & LOQ)
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

def load_limit_config():
    """Tải thư viện kép chứa cả MDL và LOQ từ Tab CauHinh_MDL_LOQ"""
    try:
        df_limit = conn.read(spreadsheet=SHEET_URL, worksheet="CauHinh_MDL_LOQ", ttl=60)
        return df_limit
    except:
        return pd.DataFrame(columns=["Nền Mẫu", "Tên Chất", "MDL", "LOQ", "Đơn Vị"])

if "df" not in st.session_state:
    st.session_state.df = load_data()
if "df_limit" not in st.session_state:
    st.session_state.df_limit = load_limit_config()

# ==========================================
# 3. HÀM BỔ SUNG: XỬ LÝ SỐ LIỆU THEO SOP (KHÍ & NƯỚC)
# ==========================================
def parse_sample_matrix(sample_name):
    """Phân loại mẫu thành Khí (thể tích V_gas) hoặc Nước (hệ số 1)"""
    name_upper = str(sample_name).upper()
    
    # Nhóm mẫu Khí
    if 'KT' in name_upper: return 24.0, 'KT', 'Khí'
    elif 'KXQ' in name_upper: return 4.0, 'KXQ', 'Khí'
    elif 'KLV' in name_upper: return 4.0, 'KLV', 'Khí'
    
    # Nhóm mẫu Nước (Sinh hoạt, Thải, Mặt, Ngầm)
    elif any(k in name_upper for k in ['NS', 'NT', 'NM', 'NN']):
        return 1.0, 'NS', 'Nước' 
        
    return None, None, None

def get_dynamic_surrogate_expected(c_do):
    """Tự động nội suy nồng độ C_ban_đầu của chuẩn đồng hành"""
    levels = [1.0, 2.0, 4.0, 5.0, 6.0, 8.0, 10.0, 20.0, 25.0, 50.0, 100.0]
    valid_levels = []
    for lvl in levels:
        rec = (c_do / lvl) * 100.0
        if 70 <= rec <= 130:
            valid_levels.append((lvl, abs(100 - rec)))
    if valid_levels:
        valid_levels.sort(key=lambda x: x[1])
        return valid_levels[0][0]
    return min(levels, key=lambda x: abs(x - c_do))

def get_limit_info(compound_name, nen_mau):
    """Tra cứu song song MDL và LOQ dựa trên tên chất (có Fuzzy Match)"""
    df_limit = st.session_state.df_limit
    if not df_limit.empty and compound_name and nen_mau:
        # Gom nhóm nước chung thành mã NS để dễ quản lý thư viện
        search_nen = "NS" if nen_mau in ['NT', 'NM', 'NN'] else nen_mau
        
        mask_exact = (df_limit["Nền Mẫu"].astype(str).str.upper() == search_nen.upper()) & \
                     (df_limit["Tên Chất"].astype(str).str.lower() == str(compound_name).lower())
        match = df_limit[mask_exact]
        
        # Fuzzy search nếu không khớp 100%
        if match.empty:
            all_comps = df_limit[df_limit["Nền Mẫu"].astype(str).str.upper() == search_nen.upper()]["Tên Chất"].astype(str).tolist()
            close_matches = difflib.get_close_matches(str(compound_name).lower(), [c.lower() for c in all_comps], n=1, cutoff=0.7)
            if close_matches:
                match = df_limit[(df_limit["Nền Mẫu"].astype(str).str.upper() == search_nen.upper()) & (df_limit["Tên Chất"].astype(str).str.lower() == close_matches[0])]

        if not match.empty:
            def parse_val(col_name):
                if col_name in match.columns:
                    val = str(match[col_name].values[0]).replace(',', '.')
                    try: return float(val)
                    except: return None
                return None

            mdl_val = parse_val("MDL")
            loq_val = parse_val("LOQ")
            unit = str(match["Đơn Vị"].values[0]) if "Đơn Vị" in match.columns else ""
            if pd.isna(unit) or unit == 'nan': unit = ""
            
            return mdl_val, loq_val, unit
    return None, None, ""

def evaluate_result(raw_conc, v_param, mdl_val, loq_val, unit, loai_mau, recovery=100.0):
    """Tính toán C_thực theo SOP và so sánh với Giới hạn (MDL/LOQ) tương ứng"""
    if pd.isna(raw_conc) or raw_conc <= 0:
        return "KPH"
        
    if loai_mau == 'Khí':
        # Công thức mẫu Khí: (C_đo * V_giải_hấp) / V_khí * (100 / R%)
        c_thuc_val = (raw_conc * 1.0) / v_param * (100.0 / recovery)
        if mdl_val is not None and c_thuc_val < mdl_val:
            return f"KPH (< MDL: {mdl_val} {unit})"
            
    elif loai_mau == 'Nước':
        # Công thức mẫu Nước chuẩn SOP VOCs-NS: C_thực = C_đo * (100 / R%)
        c_thuc_val = raw_conc * (100.0 / recovery)
        if loq_val is not None and c_thuc_val < loq_val:
            return f"< LOQ ({loq_val} {unit})"
            
    else:
        c_thuc_val = raw_conc

    return f"{round(c_thuc_val, 4)}"

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
    st.session_state.df_limit = load_limit_config() 
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
    with col_search: search_query = st.text_input("🔍 Tìm kiếm (Mã mẫu, Tên mẻ, Chỉ tiêu):")

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
            st.session_state.df = st.session_state.df.drop(index=deleted_indices).reset_index(drop=True)
            
        save_data(st.session_state.df)
        st.success("Đã đồng bộ lên cơ sở dữ liệu chung!")
        st.rerun()

# ---------------------------------------------------------
elif menu == "📥 Quản lý Tiếp nhận":
    st.title("📥 Khu vực Tiếp nhận mẫu mới")
    
    tab_excel, tab_thu_cong = st.tabs(["📁 Tải file Excel tự động", "✍️ Nhập thủ công (Mẫu lẻ)"])
    
    with tab_excel:
        st.info("💡 Kéo thả file Excel để hệ thống tự động trích xuất Tên Mẻ, Mã Mẫu, Nhận diện Nền Mẫu và Chỉ Tiêu.")
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
                
                khm_col_idx, header_row_idx = None, None
                for r in range(min(20, len(df_upload))):
                    for c in range(len(df_upload.columns)):
                        if str(df_upload.iloc[r, c]).strip() == 'KHM':
                            khm_col_idx, header_row_idx = c, r
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
                            elif khm_upper.startswith(("NS", "NT", "NM", "NN")): nen_mau_auto = "Nước"
                            else: nen_mau_auto = "Chưa xác định"

                            sample_params = [param_name for c, param_name in params_info if pd.notna(df_upload.iloc[r, c]) and str(df_upload.iloc[r, c]).strip() != '']
                            chuoi_chi_tieu = ", ".join(sample_params) if sample_params else "Chưa xác định"
                            samples_data.append({"Chọn": True, "Mã Mẫu": khm_val, "Nền Mẫu": nen_mau_auto, "Chỉ Tiêu": chuoi_chi_tieu})
                    
                    if len(samples_data) > 0:
                        st.success(f"✔️ Quét thành công **{len(samples_data)}** mẫu thuộc mẻ: **{ten_me_extract}**")
                        edited_preview = st.data_editor(
                            pd.DataFrame(samples_data),
                            column_config={"Chọn": st.column_config.CheckboxColumn("Nhập mẫu?", default=True), "Mã Mẫu": st.column_config.TextColumn(disabled=True), "Chỉ Tiêu": st.column_config.TextColumn(disabled=True)},
                            hide_index=True, use_container_width=True
                        )
                        batch_nguoi = st.selectbox("Người tiếp nhận:", ["Thành", "Kỹ thuật viên 2", "Kỹ thuật viên 3"])
                        selected_samples = edited_preview[edited_preview["Chọn"] == True]
                        
                        if st.button(f"🚀 Lưu {len(selected_samples)} mẫu đã chọn vào Hệ thống", type="primary"):
                            new_rows = [{"Mã Mẫu": row["Mã Mẫu"], "Tên Mẻ": ten_me_extract, "Nền Mẫu": row["Nền Mẫu"], "Chỉ Tiêu": row["Chỉ Tiêu"], "Trạng Thái": STATUSES[0], "Người Giữ": batch_nguoi, "Ghi Chú": "Import Excel", "Giờ Nhận": datetime.now()} for _, row in selected_samples.iterrows()]
                            st.session_state.df = pd.concat([st.session_state.df, pd.DataFrame(new_rows)], ignore_index=True)
                            save_data(st.session_state.df)
                            st.success("Đã nạp thành công!")
                            st.rerun()
                    else: st.warning("Không có mã KHM nào hợp lệ bên dưới ô tiêu đề.")
                else: st.error("Không tìm thấy ô 'KHM' trong file!")
            except Exception as e: st.error(f"Lỗi đọc file: {e}")

    with tab_thu_cong:
        with st.form("add_sample_form", clear_on_submit=True):
            new_id, new_name = st.text_input("Mã Mẫu (VD: NS-1509-01)*"), st.text_input("Tên Mẻ (VD: 2026.07.017)")
            col_t1, col_t2 = st.columns(2)
            with col_t1: new_nen = st.selectbox("Nền Mẫu", ["Nước", "Khí"])
            with col_t2: new_chi_tieu = st.text_input("Chỉ tiêu đo")
            new_nguoi = st.text_input("Người tiếp nhận (Ký tên)")
            
            if st.form_submit_button("Thêm Mẫu lẻ") and new_id:
                st.session_state.df = pd.concat([st.session_state.df, pd.DataFrame([{"Mã Mẫu": new_id, "Tên Mẻ": new_name, "Nền Mẫu": new_nen, "Chỉ Tiêu": new_chi_tieu, "Trạng Thái": STATUSES[0], "Người Giữ": new_nguoi, "Ghi Chú": "", "Giờ Nhận": datetime.now()}])], ignore_index=True)
                save_data(st.session_state.df)
                st.success(f"Đã thêm {new_id}!")

# ---------------------------------------------------------
elif menu == "⚙️ Vận hành GC-MS":
    st.title("⚙️ Điều phối & Vận hành Máy đo")
    
    col_seq, col_import = st.columns(2)
    with col_seq:
        st.subheader("1. Xuất Sequence chạy máy")
        df_ready = st.session_state.df[st.session_state.df["Trạng Thái"] == "🟡 3. Chờ chạy máy"]
        st.write(f"Hiện đang có **{len(df_ready)}** mẫu chờ chạy.")
        
        if not df_ready.empty:
            seq_df = pd.DataFrame({'Vial': range(1, len(df_ready) + 1), 'Sample Name': df_ready['Mã Mẫu'], 'Sample Type': 'Sample'})
            seq_df['Method'] = df_ready['Chỉ Tiêu'].apply(lambda x: 'VOCs.M' if any(k in str(x).upper() for k in ['VOC', 'BENZEN', 'TOLUEN', 'CHLORO', 'STYREN']) else 'HCHO.M')
            seq_df['Data File'] = datetime.now().strftime("%Y%m%d") + "_" + df_ready['Mã Mẫu']
            st.download_button("📥 Tải Sequence.csv", data=seq_df.to_csv(index=False).encode('utf-8'), file_name=f"MassHunter_Seq_{datetime.now().strftime('%Y%m%d')}.csv", mime="text/csv", type="primary")
            
    with col_import:
        st.subheader("2. Xử lý dữ liệu GC-MS theo SOP")
        st.info("💡 Web phân loại song song mẫu Khí (So với MDL) và mẫu Nước (So với LOQ). Bù trừ R% tự động.")
        
        gc_file = st.file_uploader("Kéo thả báo cáo GC (PDF/Excel/CSV)", type=["pdf", "xlsx", "xls", "csv"])

        if gc_file is not None:
            calc_results = []
            try:
                # --- TIỀN XỬ LÝ DỮ LIỆU TỪ FILE ---
                if gc_file.name.endswith('.pdf'):
                    import PyPDF2
                    text = "".join([page.extract_text() + "\n" for page in PyPDF2.PdfReader(gc_file).pages])
                    pdf_data, current_compound = [], None
                    known_compounds = ['Benzene', 'Toluene-D8', 'BFB', '4-Bromofluorobenzene', 'Toluene', 'Ethylbenzene', 'm-Xylene', 'p-Xylene', 'o-Xylene', 'Styrene']

                    for line in text.split('\n'):
                        parts = line.split()
                        if not parts: continue
                        if line.strip() in known_compounds: 
                            current_compound = line.strip()
                            continue
                        
                        if len(parts) >= 5 and (parts[0].endswith('.d') or parts[0].replace('.d','') in ['10PPM', '2', '4', '6', '8']) and ('Sample' in parts or 'Cal' in parts):
                            data_file = parts[0].replace('.d', '')
                            floats = [float(p) for p in parts if p.replace('.', '', 1).isdigit() and p.count('.') <= 1]
                            if len(floats) >= 3 and current_compound:
                                final_conc = floats[-3] if 'Cal' in parts and len(floats) >= 5 else (floats[-2] if 'Cal' in parts and len(floats) >= 4 else floats[-1])
                                pdf_data.append({"Data File": data_file, "Compound Name": current_compound, "Final Conc.": final_conc})
                    
                    df_gc, compound_col = pd.DataFrame(pdf_data), "Compound Name"
                else:
                    df_gc = pd.read_csv(gc_file) if gc_file.name.endswith('.csv') else pd.read_excel(gc_file)
                    df_gc.columns = [str(c).strip() for c in df_gc.columns]
                    compound_col = next((c for c in df_gc.columns if c.lower() in ['name', 'compound', 'compound name', 'tên chất']), None)

                # --- BƯỚC TÍNH TOÁN & ÁP MỨC GIỚI HẠN ---
                if 'Data File' in df_gc.columns and 'Final Conc.' in df_gc.columns and compound_col:
                    
                    # 1. Quét tìm Độ thu hồi (Recovery) cho từng mẫu dựa trên AI đoán nồng độ
                    surrogate_dict = {}
                    for _, row in df_gc.iterrows():
                        comp_name = str(row[compound_col]).upper()
                        if comp_name in ['TOLUENE-D8', 'TOLUEN-D8', 'BFB', '4-BROMOFLUOROBENZENE']:
                            sample_name = str(row['Data File']).replace('.d', '')
                            raw_conc = pd.to_numeric(row['Final Conc.'], errors='coerce')
                            if pd.notna(raw_conc) and raw_conc > 0:
                                c_exp = get_dynamic_surrogate_expected(raw_conc)
                                recovery = (raw_conc / c_exp) * 100.0
                                surrogate_dict[sample_name] = recovery

                    # 2. Suy luận loại mẫu mặc định nếu không tra được mã
                    default_v_gas, default_nen, default_loai = 24.0, 'KT', 'Khí'
                    for _, row in df_gc.iterrows():
                        sn = str(row['Data File']).upper()
                        if 'KT' in sn: default_v_gas, default_nen, default_loai = 24.0, 'KT', 'Khí'; break
                        elif 'KXQ' in sn: default_v_gas, default_nen, default_loai = 4.0, 'KXQ', 'Khí'; break
                        elif 'KLV' in sn: default_v_gas, default_nen, default_loai = 4.0, 'KLV', 'Khí'; break
                        elif any(k in sn for k in ['NS', 'NT', 'NM', 'NN']): default_v_gas, default_nen, default_loai = 1.0, 'NS', 'Nước'; break

                    # 3. Tính C thực tế cho các chất
                    for _, row in df_gc.iterrows():
                        sample_name, comp_name = str(row['Data File']).replace('.d', ''), str(row[compound_col])
                        upper_name = sample_name.upper()
                        
                        # Chỉ duyệt mẫu hợp lệ
                        if not any(k in upper_name for k in ['KT', 'KXQ', 'KLV', 'NS', 'NT', 'NM', 'NN', 'BL', 'BLANK', 'TC', 'QC']): continue
                        if upper_name in ['1', '2', '4', '5', '6', '8', '10'] or 'PPM' in upper_name: continue
                        if comp_name.upper() in ['TOLUENE-D8', 'TOLUEN-D8', 'BFB', '4-BROMOFLUOROBENZENE']: continue
                            
                        raw_conc = pd.to_numeric(row['Final Conc.'], errors='coerce')
                        if pd.isna(raw_conc): continue

                        v_param, nen_mau, loai_mau = parse_sample_matrix(sample_name)
                        if v_param is None: v_param, nen_mau, loai_mau = default_v_gas, default_nen, default_loai
                            
                        sample_recovery = surrogate_dict.get(sample_name, 100.0)
                        
                        # Tra cứu song song MDL và LOQ
                        mdl_val, loq_val, unit = get_limit_info(comp_name, nen_mau)
                        
                        # Đánh giá C thực tế dựa trên Loại mẫu
                        c_thuc_str = evaluate_result(raw_conc, v_param, mdl_val, loq_val, unit, loai_mau, recovery=sample_recovery)
                        
                        # Hiển thị Limit tham chiếu tương ứng
                        limit_display = ""
                        if loai_mau == 'Khí' and mdl_val is not None: limit_display = f"MDL: {mdl_val} {unit}"
                        elif loai_mau == 'Nước' and loq_val is not None: limit_display = f"LOQ: {loq_val} {unit}"
                        
                        calc_results.append({
                            "Tên mẫu": sample_name, "Tên chỉ tiêu": comp_name, "C đo": round(raw_conc, 4), 
                            "C thực": c_thuc_str, "Giới hạn": limit_display, "R(%)": f"{round(sample_recovery, 1)}%"
                        })

                if calc_results:
                    st.success(f"✔️ Đã xuất {len(calc_results)} dòng kết quả. Tự động áp dụng SOP Khí/Nước.")
                    st.dataframe(pd.DataFrame(calc_results), use_container_width=True, hide_index=True)
                else: 
                    st.warning("⚠️ File không chứa mẫu hợp lệ (KT, KXQ, NS, NT...) hoặc thiếu dữ liệu.")
            except Exception as e: st.error(f"❌ Lỗi xử lý: {e}")

# ---------------------------------------------------------
elif menu == "🚀 Tiện ích & Cấu hình":
    st.title("🛠️ Tiện ích & Cấu hình Hệ thống")
    
    tab_limit, tab_report, tab_qr = st.tabs(["📚 Thư viện Cấu hình MDL/LOQ", "📝 Lập Biên Bản (Sắp ra mắt)", "🏷️ Sinh Mã QR"])
    
    with tab_limit:
        st.subheader("Tra cứu thông minh Thư viện Cấu hình (MDL & LOQ)")
        search_limit = st.text_input("🔍 Tra cứu giới hạn sát tên chất (VD: Benzen, Toluen, Xylen):")
        
        limit_library = st.session_state.df_limit
        if search_limit:
            if not limit_library.empty:
                all_compounds = limit_library["Tên Chất"].astype(str).unique()
                close_matches = difflib.get_close_matches(search_limit.lower(), [c.lower() for c in all_compounds], n=20, cutoff=0.3)
                
                mask_contains = limit_library["Tên Chất"].astype(str).str.contains(search_limit, case=False, na=False)
                mask_fuzzy = limit_library["Tên Chất"].astype(str).str.lower().isin(close_matches)
                
                search_result = limit_library[mask_contains | mask_fuzzy].copy()
                if not search_result.empty:
                    # Sắp xếp theo mức độ sát nghĩa với từ khóa
                    search_result["Score"] = search_result["Tên Chất"].apply(lambda x: difflib.SequenceMatcher(None, search_limit.lower(), str(x).lower()).ratio())
                    st.dataframe(search_result.sort_values(by="Score", ascending=False).drop(columns=["Score"]), use_container_width=True, hide_index=True)
                else: 
                    st.warning("⚠️ Không tìm thấy chất nào sát với từ khóa bạn nhập.")
            else: st.info("Thư viện hiện đang trống.")
        else:
            if not limit_library.empty:
                with st.expander("Xem toàn bộ Thư viện MDL & LOQ hiện tại", expanded=False):
                    st.dataframe(limit_library, use_container_width=True, hide_index=True)

        st.divider()
        st.subheader("Cập nhật Thư viện Giới hạn (Khí & Nước)")
        st.info("Kéo thả file Excel chứa bảng Giới hạn. Code sẽ tự động lùng sục cột 'MDL' và 'LOQ' để gộp lại tải lên Google Sheets (Tab `CauHinh_MDL_LOQ`).")
        
        limit_file = st.file_uploader("Tải lên file Excel Bảng MDL/LOQ", type=["xlsx"])
        if limit_file:
            try:
                xls = pd.ExcelFile(limit_file)
                limit_data = []
                for sheet in xls.sheet_names:
                    df_sheet = pd.read_excel(xls, sheet_name=sheet)
                    col_ten = next((c for c in df_sheet.columns if 'tên' in str(c).lower() or 'hợp chất' in str(c).lower()), None)
                    col_mdl = next((c for c in df_sheet.columns if 'mdl' in str(c).lower()), None)
                    col_loq = next((c for c in df_sheet.columns if 'loq' in str(c).lower()), None)
                    
                    if col_ten and (col_mdl or col_loq):
                        unit = "Chưa rõ"
                        if col_loq:
                            unit_match = re.search(r'\((.*?)\)', str(col_loq))
                            if unit_match: unit = unit_match.group(1)
                        if unit == "Chưa rõ" and col_mdl:
                            unit_match = re.search(r'\((.*?)\)', str(col_mdl))
                            if unit_match: unit = unit_match.group(1)

                        # Phân biệt nền mẫu để gắn Tag
                        nen_mau = "KT" if "thải" in sheet.lower() else ("KXQ" if "xung quanh" in sheet.lower() else ("KLV" if "làm việc" in sheet.lower() else ("NS" if "nước" in sheet.lower() else sheet)))
                        
                        for _, row in df_sheet.iterrows():
                            val_mdl = str(row[col_mdl]).replace(',', '.').strip() if col_mdl and pd.notna(row[col_mdl]) else ""
                            val_loq = str(row[col_loq]).replace(',', '.').strip() if col_loq and pd.notna(row[col_loq]) else ""
                            
                            if pd.notna(row[col_ten]) and (val_mdl or val_loq):
                                limit_data.append({
                                    "Nền Mẫu": nen_mau, "Tên Chất": str(row[col_ten]).strip(), 
                                    "MDL": val_mdl, "LOQ": val_loq, "Đơn Vị": unit
                                })
                
                if limit_data:
                    df_limit_new = pd.DataFrame(limit_data)
                    st.success(f"✔️ Đã quét được {len(df_limit_new)} chỉ tiêu từ file.")
                    st.dataframe(df_limit_new, use_container_width=True)
                    if st.button("🚀 Đẩy lên Google Sheets (Tạo/Ghi đè Thư viện)", type="primary"):
                        conn.update(spreadsheet=SHEET_URL, worksheet="CauHinh_MDL_LOQ", data=df_limit_new)
                        st.session_state.df_limit = df_limit_new
                        st.success("🎉 Đã lưu cấu hình lên Cloud thành công!")
                else: st.error("Không tìm thấy cấu trúc bảng hợp lệ (Cột Tên / Cột MDL / Cột LOQ).")
            except Exception as e: st.error(f"Lỗi đọc file: {e}")

    with tab_report: st.write("Khu vực xuất Form Word/Excel.")
    with tab_qr: st.write("Chức năng tạo mã vạch (Barcode) hàng loạt.")
