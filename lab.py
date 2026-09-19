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
    try:
        df_limit = conn.read(spreadsheet=SHEET_URL, worksheet="CauHinh_MDL_LOQ", ttl=0) 
        return df_limit
    except:
        return pd.DataFrame(columns=["Nền Mẫu", "Tên Chất", "MDL", "LOQ", "Đơn Vị"])

if "df" not in st.session_state:
    st.session_state.df = load_data()
if "df_limit" not in st.session_state:
    st.session_state.df_limit = load_limit_config()

# ==========================================
# 3. HÀM BỔ SUNG: XỬ LÝ SỐ LIỆU THEO SOP
# ==========================================
def parse_sample_matrix(sample_name):
    name_upper = str(sample_name).upper()
    if 'KT' in name_upper: return 24.0, 'KT', 'Khí'
    elif 'KXQ' in name_upper: return 4.0, 'KXQ', 'Khí'
    elif 'KLV' in name_upper: return 4.0, 'KLV', 'Khí'
    elif any(k in name_upper for k in ['NS', 'NT', 'NM', 'NN']): return 1.0, 'NS', 'Nước' 
    return None, None, None

def get_dynamic_surrogate_expected(c_do):
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
    df_limit = st.session_state.df_limit
    if not df_limit.empty and compound_name and nen_mau:
        search_nen = "NS" if nen_mau in ['NT', 'NM', 'NN'] else nen_mau
        
        mask_exact = (df_limit["Nền Mẫu"].astype(str).str.upper() == search_nen.upper()) & \
                     (df_limit["Tên Chất"].astype(str).str.lower() == str(compound_name).lower())
        match = df_limit[mask_exact]
        
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
    if pd.isna(raw_conc) or raw_conc <= 0:
        return "KPH"
        
    if loai_mau == 'Khí':
        c_thuc_val = (raw_conc * 1.0) / v_param * (100.0 / recovery)
        if mdl_val is not None and c_thuc_val < mdl_val:
            return f"KPH (< MDL: {mdl_val} {unit})"
            
    elif loai_mau == 'Nước':
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
st.sidebar.caption("Phần mềm Quản lý Phòng Lab")
st.sidebar.divider()

menu = st.sidebar.radio("📌 ĐIỀU HƯỚNG CHÍNH", [
    "🏠 Trang chủ (Tổng quan)", 
    "📥 Quản lý Tiếp nhận", 
    "⚙️ Vận hành GC-MS (Agilent - VOCs)",
    "🔥 Vận hành GC-FID (Agilent)",
    "🧬 Vận hành Thermo (OCP/OPP/PCB)",
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
    col2.metric("⏳ Đang chờ chạy", cho_chay_may)
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
                            column_config={
                                "Chọn": st.column_config.CheckboxColumn("Nhập mẫu?", default=True),
                                "Mã Mẫu": st.column_config.TextColumn(disabled=True),
                                "Nền Mẫu": st.column_config.SelectboxColumn("Nền Mẫu", options=["Nước", "Khí", "Chưa xác định"]),
                                "Chỉ Tiêu": st.column_config.TextColumn(disabled=True)
                            },
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
            except Exception as e: st.error(f"Lỗi: {e}")

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

elif menu == "⚙️ Vận hành GC-MS (Agilent - VOCs)":
    st.title("⚙️ Điều phối & Vận hành Máy đo (Agilent)")
    
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
        st.info("💡 Hệ thống AI tự quét thư viện LOQ/MDL để nhận diện danh sách các chất cần bóc tách từ file báo cáo.")
        
        gc_file = st.file_uploader("Kéo thả báo cáo GC (PDF/Excel/CSV)", type=["pdf", "xlsx", "xls", "csv"])

        if gc_file is not None:
            calc_results = []
            try:
                dynamic_compounds = []
                if not st.session_state.df_limit.empty and "Tên Chất" in st.session_state.df_limit.columns:
                    dynamic_compounds = st.session_state.df_limit["Tên Chất"].dropna().astype(str).str.strip().tolist()
                
                surrogate_compounds = ['Toluene-D8', 'Toluen-D8', 'BFB', '4-Bromofluorobenzene', 'Chlorobenzene-d5']
                known_compounds_upper = set(c.upper() for c in dynamic_compounds + surrogate_compounds)

                if gc_file.name.endswith('.pdf'):
                    import PyPDF2
                    text = "".join([page.extract_text() + "\n" for page in PyPDF2.PdfReader(gc_file).pages])
                    pdf_data, current_compound = [], None

                    for line in text.split('\n'):
                        parts = line.split()
                        if not parts: continue
                        
                        clean_line = line.strip()
                        if clean_line.upper() in known_compounds_upper: 
                            current_compound = clean_line
                            continue
                        
                        data_file_raw = parts[0].replace('.d', '')
                        if len(parts) >= 5 and (parts[0].endswith('.d') or data_file_raw in ['10PPM', '10a', '2', '4', '6', '8']) and ('Sample' in parts or 'Cal' in parts):
                            data_file = data_file_raw
                            floats = [float(p.replace(',', '.')) for p in parts if p.replace('.', '', 1).replace(',', '', 1).isdigit() and (p.count('.') + p.count(',') <= 1)]
                            
                            if len(floats) >= 3 and current_compound:
                                final_conc = floats[-3] if 'Cal' in parts and len(floats) >= 5 else (floats[-2] if 'Cal' in parts and len(floats) >= 4 else floats[-1])
                                pdf_data.append({"Data File": data_file, "Compound Name": current_compound, "Final Conc.": final_conc})
                    
                    df_gc, compound_col = pd.DataFrame(pdf_data), "Compound Name"
                else:
                    df_gc = pd.read_csv(gc_file) if gc_file.name.endswith('.csv') else pd.read_excel(gc_file)
                    df_gc.columns = [str(c).strip() for c in df_gc.columns]
                    compound_col = next((c for c in df_gc.columns if c.lower() in ['name', 'compound', 'compound name', 'tên chất']), None)

                if 'Data File' in df_gc.columns and 'Final Conc.' in df_gc.columns and compound_col:
                    
                    surrogate_dict = {}
                    for _, row in df_gc.iterrows():
                        comp_name = str(row[compound_col]).upper()
                        if comp_name in ['TOLUENE-D8', 'TOLUEN-D8', 'BFB', '4-BROMOFLUOROBENZENE']:
                            sample_name = str(row['Data File']).replace('.d', '')
                            raw_conc = pd.to_numeric(row['Final Conc.'], errors='coerce')
                            if pd.notna(raw_conc) and raw_conc > 0:
                                c_exp = get_dynamic_surrogate_expected(raw_conc)
                                recovery = (raw_conc / c_exp) * 100.0
                                surrogate_dict[sample_name] = {"recovery": recovery, "c_exp": c_exp, "c_do": raw_conc}

                    default_v_gas, default_nen, default_loai = 24.0, 'KT', 'Khí'
                    for _, row in df_gc.iterrows():
                        sn = str(row['Data File']).upper()
                        if 'KT' in sn: default_v_gas, default_nen, default_loai = 24.0, 'KT', 'Khí'; break
                        elif 'KXQ' in sn: default_v_gas, default_nen, default_loai = 4.0, 'KXQ', 'Khí'; break
                        elif 'KLV' in sn: default_v_gas, default_nen, default_loai = 4.0, 'KLV', 'Khí'; break
                        elif any(k in sn for k in ['NS', 'NT', 'NM', 'NN']): default_v_gas, default_nen, default_loai = 1.0, 'NS', 'Nước'; break

                    for _, row in df_gc.iterrows():
                        sample_name, comp_name = str(row['Data File']).replace('.d', ''), str(row[compound_col])
                        upper_name = sample_name.upper()
                        
                        if not any(k in upper_name for k in ['KT', 'KXQ', 'KLV', 'NS', 'NT', 'NM', 'NN', 'BL', 'BLANK', 'TC', 'QC']): continue
                        if upper_name in ['1', '2', '4', '5', '6', '8', '10', '10A'] or 'PPM' in upper_name: continue
                        if comp_name.upper() in ['TOLUENE-D8', 'TOLUEN-D8', 'BFB', '4-BROMOFLUOROBENZENE']: continue
                            
                        raw_conc = pd.to_numeric(row['Final Conc.'], errors='coerce')
                        if pd.isna(raw_conc): continue

                        v_param, nen_mau, loai_mau = parse_sample_matrix(sample_name)
                        if v_param is None: v_param, nen_mau, loai_mau = default_v_gas, default_nen, default_loai
                            
                        surr_info = surrogate_dict.get(sample_name, {"recovery": 100.0, "c_exp": "", "c_do": ""})
                        sample_recovery = surr_info["recovery"]
                        c_surr_truoc = surr_info["c_exp"]
                        c_surr_sau = surr_info["c_do"]
                        
                        mdl_val, loq_val, unit = get_limit_info(comp_name, nen_mau)
                        
                        c_thuc_str = evaluate_result(raw_conc, v_param, mdl_val, loq_val, unit, loai_mau, recovery=sample_recovery)
                        
                        limit_display = ""
                        if loai_mau == 'Khí' and mdl_val is not None: limit_display = f"MDL: {mdl_val} {unit}"
                        elif loai_mau == 'Nước' and loq_val is not None: limit_display = f"LOQ: {loq_val} {unit}"
                        
                        calc_results.append({
                            "Tên mẫu": sample_name, "Tên chỉ tiêu": comp_name, "C đo": round(raw_conc, 4), 
                            "C thực": c_thuc_str, "Giới hạn": limit_display, "R(%)": f"{round(sample_recovery, 1)}%",
                            "C_surr_truoc": c_surr_truoc, "C_surr_sau": c_surr_sau
                        })

                if calc_results:
                    st.success(f"✔️ Đã xuất {len(calc_results)} dòng kết quả. Tự động áp dụng SOP Khí/Nước.")
                    
                    df_results = pd.DataFrame(calc_results)
                    df_results = df_results.sort_values(by=["Tên mẫu", "Tên chỉ tiêu"]).reset_index(drop=True)
                    
                    display_cols = ["Tên mẫu", "Tên chỉ tiêu", "C đo", "C thực", "Giới hạn", "R(%)"]
                    st.dataframe(df_results[display_cols], use_container_width=True, hide_index=True)
                    
                    csv_results = df_results.to_csv(index=False).encode('utf-8-sig')
                    st.download_button(
                        label="📥 Tải Kết quả (CSV) để Lập Biên Bản",
                        data=csv_results,
                        file_name=f"Ket_Qua_GC_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
                        mime="text/csv",
                        type="primary"
                    )
                else: 
                    st.warning("⚠️ File không chứa mẫu hợp lệ (KT, KXQ, NS, NT...) hoặc thiếu dữ liệu.")
            except Exception as e: st.error(f"❌ Lỗi xử lý: {e}")

elif menu == "🔥 Vận hành GC-FID (Agilent)":
    st.title("🔥 Hệ thống GC-FID (Agilent)")
    st.caption("Module chuyên biệt xử lý dữ liệu từ đầu dò FID")
    
    col_seq, col_import = st.columns(2)
    with col_seq:
        st.subheader("1. Xuất Sequence GC-FID")
        st.info("Sẽ tích hợp thuật toán xuất file Sequence định dạng cho máy GC-FID Agilent.")
        
    with col_import:
        st.subheader("2. Xử lý kết quả GC-FID")
        st.info("Khu vực chờ tích hợp thuật toán đọc file báo cáo từ máy GC-FID.")
        fid_file = st.file_uploader("Kéo thả báo cáo GC-FID (PDF/Excel/CSV/TXT)", type=["pdf", "xlsx", "xls", "csv", "txt"])
        if fid_file:
            st.warning("🚧 Hệ thống đang chờ cập nhật thuật toán bóc tách dữ liệu từ file report FID.")

elif menu == "🧬 Vận hành Thermo (OCP/OPP/PCB)":
    st.title("🧬 Hệ thống Thermo GC-MS")
    st.caption("Module chuyên biệt xử lý dữ liệu OCP, OPP, PCB và Phenol")
    
    col_seq, col_import = st.columns(2)
    with col_seq:
        st.subheader("1. Xuất Sequence Thermo")
        st.info("Sẽ tích hợp thuật toán xuất file Sequence định dạng tương thích phần mềm Thermo.")
        
    with col_import:
        st.subheader("2. Xử lý kết quả Thermo")
        st.info("Khu vực chờ tích hợp thuật toán đọc file xuất từ máy Thermo.")
        thermo_file = st.file_uploader("Kéo thả báo cáo Thermo (PDF/Excel/CSV)", type=["pdf", "xlsx", "xls", "csv"])
        if thermo_file:
            st.warning("🚧 Hệ thống đang chờ cập nhật thuật toán bóc tách dữ liệu từ file report Thermo.")

elif menu == "🚀 Tiện ích & Cấu hình":
    st.title("🛠️ Tiện ích & Cấu hình Hệ thống")
    
    tab_limit, tab_report, tab_qr = st.tabs(["📚 Quản lý Thư viện MDL & LOQ", "📝 Lập Biên Bản", "🏷️ Sinh Mã QR"])
    
    with tab_limit:
        st.subheader("1. Quản lý Thư viện Trực tiếp (Thêm/Sửa Cột & Hàng)")
        st.info("Chỉnh sửa số liệu, xóa hoặc thêm chất trực tiếp trên bảng này. Bạn có thể gõ vào cột LOQ hoặc MDL tùy ý. Sau khi chỉnh sửa, bấm **Lưu thay đổi**.")
        
        df_current_limit = st.session_state.df_limit.copy()
        
        if df_current_limit.empty:
            df_current_limit = pd.DataFrame(columns=["Nền Mẫu", "Tên Chất", "MDL", "LOQ", "Đơn Vị"])
            df_current_limit.loc[0] = ["", "", "", "", ""]
            
        edited_limit = st.data_editor(
            df_current_limit, 
            num_rows="dynamic", 
            use_container_width=True,
            key="limit_editor",
            height=350
        )
        
        if st.button("💾 Lưu thay đổi Thư viện (Ghi đè Tab CauHinh_MDL_LOQ)", type="primary"):
            try:
                st.cache_data.clear() 
                edited_limit = edited_limit[edited_limit["Tên Chất"].str.strip() != ""] 
                
                conn.update(spreadsheet=SHEET_URL, worksheet="CauHinh_MDL_LOQ", data=edited_limit)
                st.session_state.df_limit = edited_limit
                st.success("🎉 Đã lưu thư viện lên Google Sheets thành công!")
            except Exception as e:
                st.error(f"⚠️ Lỗi kết nối Google Sheets: {e}")

        st.divider()
        st.subheader("2. Hoặc Cập nhật hàng loạt từ file Excel")
        st.info("Kéo thả file Excel chứa bảng Giới hạn. Đảm bảo file có cột 'Tên chất' và 'MDL' hoặc 'LOQ'. Hệ thống sẽ gộp dữ liệu mới vào thư viện cũ.")
        
        limit_file = st.file_uploader("Tải lên file Excel Bảng MDL/LOQ", type=["xlsx"])
        if limit_file:
            try:
                xls = pd.ExcelFile(limit_file)
                limit_data = []
                for sheet in xls.sheet_names:
                    df_sheet = pd.read_excel(xls, sheet_name=sheet, header=None)
                    header_idx = -1
                    c_ten, c_mdl, c_loq = None, None, None
                    
                    for r in range(min(20, len(df_sheet))):
                        row_vals = [str(val).lower() for val in df_sheet.iloc[r].values]
                        c_ten_temp = next((i for i, v in enumerate(row_vals) if 'tên' in v or 'hợp chất' in v), None)
                        c_mdl_temp = next((i for i, v in enumerate(row_vals) if 'mdl' in v), None)
                        c_loq_temp = next((i for i, v in enumerate(row_vals) if 'loq' in v), None)
                        
                        if c_ten_temp is not None and (c_mdl_temp is not None or c_loq_temp is not None):
                            header_idx, c_ten, c_mdl, c_loq = r, c_ten_temp, c_mdl_temp, c_loq_temp
                            break
                    
                    if header_idx != -1:
                        unit = "Chưa rõ"
                        if c_loq is not None:
                            unit_match = re.search(r'\((.*?)\)', str(df_sheet.iloc[header_idx, c_loq]))
                            if unit_match: unit = unit_match.group(1)
                        if unit == "Chưa rõ" and c_mdl is not None:
                            unit_match = re.search(r'\((.*?)\)', str(df_sheet.iloc[header_idx, c_mdl]))
                            if unit_match: unit = unit_match.group(1)

                        s_lower = sheet.lower()
                        nen_mau = "KT" if "thải" in s_lower else ("KXQ" if "xung quanh" in s_lower else ("KLV" if "làm việc" in s_lower else ("NS" if "nước" in s_lower or "voc" in s_lower else sheet)))
                        
                        for r in range(header_idx + 1, len(df_sheet)):
                            ten_val = df_sheet.iloc[r, c_ten]
                            if pd.isna(ten_val) or str(ten_val).strip() == "" or str(ten_val).lower() == 'nan': continue
                            
                            val_mdl = str(df_sheet.iloc[r, c_mdl]).replace(',', '.').strip() if c_mdl is not None and pd.notna(df_sheet.iloc[r, c_mdl]) else ""
                            val_loq = str(df_sheet.iloc[r, c_loq]).replace(',', '.').strip() if c_loq is not None and pd.notna(df_sheet.iloc[r, c_loq]) else ""
                            if val_mdl.lower() == 'nan': val_mdl = ""
                            if val_loq.lower() == 'nan': val_loq = ""
                            
                            if val_mdl or val_loq:
                                limit_data.append({"Nền Mẫu": nen_mau, "Tên Chất": str(ten_val).strip(), "MDL": val_mdl, "LOQ": val_loq, "Đơn Vị": unit})
                
                if limit_data:
                    df_limit_new = pd.DataFrame(limit_data)
                    st.success(f"✔️ Đã quét được {len(df_limit_new)} chỉ tiêu từ file.")
                    st.dataframe(df_limit_new, use_container_width=True)
                    
                    if st.button("🚀 Ghi thêm vào Google Sheets (Bổ sung/Cập nhật)", type="primary"):
                        try:
                            st.cache_data.clear() 
                            combined_df = pd.concat([st.session_state.df_limit, df_limit_new], ignore_index=True)
                            combined_df = combined_df.drop_duplicates(subset=['Nền Mẫu', 'Tên Chất'], keep='last').reset_index(drop=True)
                            
                            conn.update(spreadsheet=SHEET_URL, worksheet="CauHinh_MDL_LOQ", data=combined_df)
                            st.session_state.df_limit = combined_df
                            st.success(f"🎉 Đã ghi thêm thành công! Tổng số chỉ tiêu hiện tại trong Thư viện: {len(combined_df)}")
                        except Exception as sheet_err:
                            st.error(f"⚠️ Lỗi kết nối Google Sheets: {sheet_err}")
                else: st.error("Không tìm thấy cấu trúc bảng hợp lệ (Cột Tên / Cột MDL / Cột LOQ).")
            except Exception as e: st.error(f"Lỗi đọc file: {e}")

    with tab_report: 
        st.subheader("📝 Lập Biên Bản Xử Lý Mẫu Tự Động")
        st.info("Tải lên file Word Template (đã gắn thẻ `{{...}}`) hoặc File Excel Template và File CSV kết quả đã tính toán để hệ thống tự động điền số liệu.")
        
        col_tpl, col_data = st.columns(2)
        with col_tpl:
            template_file = st.file_uploader("1. Tải file Mẫu (.docx, .xlsx)", type=["docx", "xlsx"])
        with col_data:
            data_file = st.file_uploader("2. Tải file Kết quả (.csv)", type=["xlsx", "csv"])

        # CẢNH BÁO NẾU NGƯỜI DÙNG CHƯA TẢI ĐỦ 2 FILE
        if template_file is not None and data_file is None:
            st.warning("👉 Bạn cần tải thêm **File Kết Quả (.csv)** vào ô số 2 thì nút Lập Biên Bản mới hiện ra nhé!")
        elif data_file is not None and template_file is None:
            st.warning("👉 Bạn cần tải thêm **File Mẫu (.docx hoặc .xlsx)** vào ô số 1 thì nút Lập Biên Bản mới hiện ra nhé!")

        if template_file and data_file:
            try:
                df_kq = pd.read_excel(data_file) if data_file.name.endswith('.xlsx') else pd.read_csv(data_file)
                
                if "Tên mẫu" not in df_kq.columns or "Tên chỉ tiêu" not in df_kq.columns:
                    st.error("⚠️ File kết quả không đúng chuẩn. Vui lòng dùng đúng file tải về từ phần mềm này.")
                else:
                    # ----- TẠO DANH SÁCH CHO MẪU KHÍ THẢI -----
                    ket_qua_list = []
                    for _, row in df_kq.iterrows():
                        ten_mau_str = str(row.get("Tên mẫu", ""))
                        v_khi = "24,0" if "KT" in ten_mau_str.upper() else ("4,0" if any(k in ten_mau_str.upper() for k in ["KXQ", "KLV"]) else "")
                        h_phantram = str(row.get("R(%)", "")).replace('%', '').strip()
                        
                        c_surr_truoc = str(row.get("C_surr_truoc", "")).replace('.', ',')
                        c_surr_sau = str(row.get("C_surr_sau", "")).replace('.', ',')
                        
                        if c_surr_sau == "" and h_phantram != "":
                            try: c_surr_sau = str(round((float(h_phantram) / 100) * 10.0, 2)).replace('.', ',')
                            except: pass
                            
                        ket_qua_list.append({
                            "ngay": datetime.now().strftime("%d/%m/%Y"),
                            "ten_mau": ten_mau_str,
                            "chi_tieu": str(row.get("Tên chỉ tiêu", "")),
                            "c_surr_truoc": c_surr_truoc,
                            "c_surr_sau": c_surr_sau,
                            "h_phantram": str(h_phantram).replace('.', ','),
                            "v_khi": v_khi,
                            "c_do": str(row.get("C đo", "")).replace('.', ','),
                            "kq_thuc": str(row.get("C thực", "")).replace('.', ',')
                        })

                    # ----- TẠO DANH SÁCH CHO MẪU NƯỚC -----
                    danh_sach_mau = []
                    grouped = df_kq.groupby("Tên mẫu")
                    for ten_mau, group in grouped:
                        first_row = group.iloc[0]
                        mau_dict = {
                            "ngay": datetime.now().strftime("%d/%m/%Y"),
                            "ky_hieu": ten_mau,
                            "c_surr_truoc": str(first_row.get("C_surr_truoc", "")).replace('.', ','), 
                            "c_surr_sau": str(first_row.get("C_surr_sau", "")).replace('.', ','), 
                            "de": str(first_row.get("R(%)", "")).replace('.', ','),
                            "ghi_chu": ""
                        }
                        
                        for _, row in group.iterrows():
                            chi_tieu = str(row.get("Tên chỉ tiêu", "")).upper()
                            slug = re.sub(r'\W+', '', chi_tieu.lower())
                            mau_dict[f"cdo_{slug}"] = str(row.get("C đo", "")).replace('.', ',')
                            mau_dict[f"kq_{slug}"] = str(row.get("C thực", "")).replace('.', ',')
                            
                            if "BENZENE" in chi_tieu or "BENZEN" in chi_tieu:
                                mau_dict["cdo_benzen"] = mau_dict[f"cdo_{slug}"]
                                mau_dict["kq_benzen"] = mau_dict[f"kq_{slug}"]
                            elif "TOLUENE" in chi_tieu or "TOLUEN" in chi_tieu:
                                mau_dict["cdo_toluen"] = mau_dict[f"cdo_{slug}"]
                                mau_dict["kq_toluen"] = mau_dict[f"kq_{slug}"]
                                
                        danh_sach_mau.append(mau_dict)

                    # ----- XỬ LÝ TEMPLATE WORD -----
                    if template_file.name.endswith('.docx'):
                        if st.button("🚀 Bắt đầu Lập Biên Bản Word", type="primary"):
                            try:
                                from docxtpl import DocxTemplate
                                import io
                                
                                doc = DocxTemplate(template_file)
                                doc.render({"danh_sach_mau": danh_sach_mau, "ket_qua": ket_qua_list})
                                
                                bio = io.BytesIO()
                                doc.save(bio)
                                bio.seek(0)
                                
                                st.success("🎉 Biên bản Word đã được tạo thành công!")
                                st.download_button("📥 Tải xuống Biên Bản (.docx)", data=bio, file_name=f"Bien_Ban_{datetime.now().strftime('%Y%m%d_%H%M')}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
                            except ImportError:
                                st.error("⚠️ Hệ thống chưa cài thư viện 'docxtpl'. Hãy thêm 'docxtpl' vào requirements.txt!")
                                
                    # ----- XỬ LÝ TEMPLATE EXCEL -----
                    elif template_file.name.endswith('.xlsx'):
                        st.info("💡 Hướng dẫn: Gõ các thẻ như `{{ row.ten_mau }}`, `{{ row.kq_thuc }}` vào đúng một hàng mẫu. Hệ thống sẽ tự động nhân bản định dạng.")
                        if st.button("🚀 Bắt đầu Lập Biên Bản Excel", type="primary"):
                            import openpyxl
                            from copy import copy
                            import io
                            
                            wb = openpyxl.load_workbook(template_file)
                            ws = wb.active
                            
                            # Tìm vị trí dòng chứa thẻ template
                            template_row_idx = None
                            template_cells = []
                            for r in range(1, ws.max_row + 1):
                                for c in range(1, ws.max_column + 1):
                                    val = str(ws.cell(row=r, column=c).value or "")
                                    if "{{" in val:
                                        template_row_idx = r
                                        template_cells = [ws.cell(row=r, column=col).value for col in range(1, ws.max_column + 1)]
                                        break
                                if template_row_idx:
                                    break
                                    
                            if template_row_idx:
                                # Nhận diện template Mẫu Nước hay Mẫu Khí
                                is_nuoc = any(isinstance(v, str) and ("kq_benzen" in v or "kq_toluen" in v or "mau.ky_hieu" in v) for v in template_cells)
                                data_loop = danh_sach_mau if is_nuoc else ket_qua_list
                                
                                # Lưu lại định dạng (Style) của dòng mẫu
                                original_styles = []
                                for col in range(1, ws.max_column + 1):
                                    cell_obj = ws.cell(row=template_row_idx, column=col)
                                    original_styles.append({
                                        "font": copy(cell_obj.font),
                                        "border": copy(cell_obj.border),
                                        "fill": copy(cell_obj.fill),
                                        "number_format": copy(cell_obj.number_format),
                                        "alignment": copy(cell_obj.alignment)
                                    })
                                
                                current_row = template_row_idx
                                for item in data_loop:
                                    for col_idx, cell_val in enumerate(template_cells, start=1):
                                        new_val = cell_val
                                        if cell_val and isinstance(cell_val, str):
                                            # Dọn dẹp thẻ Word lỡ copy nhầm
                                            new_val = re.sub(r'\{%p.*?%\}', '', new_val).strip()
                                            new_val = re.sub(r'\{%.*?%\}', '', new_val).strip()
                                            
                                            # Thay thế số liệu
                                            matches = re.findall(r'\{\{\s*(?:row\.|mau\.)?(\w+)\s*\}\}', new_val)
                                            for m in matches:
                                                replacement = str(item.get(m, ""))
                                                new_val = re.sub(r'\{\{\s*(?:row\.|mau\.)?' + m + r'\s*\}\}', replacement, new_val)
                                                
                                            # Ép kiểu thành số cho Excel
                                            if isinstance(new_val, str) and re.match(r'^-?\d+(?:,\d+)?$', new_val.strip()):
                                                try:
                                                    new_val = float(new_val.strip().replace(',', '.'))
                                                except: pass
                                            
                                            if new_val == "": new_val = None
                                                
                                        # Ghi giá trị
                                        new_cell = ws.cell(row=current_row, column=col_idx)
                                        new_cell.value = new_val
                                        
                                        # Copy định dạng (Style)
                                        if current_row > template_row_idx:
                                            style = original_styles[col_idx - 1]
                                            new_cell.font = copy(style["font"])
                                            new_cell.border = copy(style["border"])
                                            new_cell.fill = copy(style["fill"])
                                            new_cell.number_format = copy(style["number_format"])
                                            new_cell.alignment = copy(style["alignment"])
                                            
                                    current_row += 1
                                    
                                # Dọn dẹp các thẻ rác ở các dòng bên dưới
                                for r in range(current_row, current_row + 5):
                                    for c in range(1, ws.max_column + 1):
                                        val = str(ws.cell(row=r, column=c).value or "")
                                        if "{%" in val or "{{" in val:
                                            ws.cell(row=r, column=c).value = None
                                            
                                bio = io.BytesIO()
                                wb.save(bio)
                                bio.seek(0)
                                
                                st.success("🎉 Biên bản Excel đã được tạo thành công!")
                                st.download_button("📥 Tải xuống Biên Bản (.xlsx)", data=bio, file_name=f"Bien_Ban_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                            else:
                                st.error("⚠️ Không tìm thấy thẻ {{...}} nào trong file Excel Template. Hãy chắc chắn bạn đã gắn thẻ vào một dòng mẫu.")
            except Exception as e:
                st.error(f"❌ Có lỗi xảy ra trong quá trình xử lý Biên Bản: {e}")

    with tab_qr: st.write("Chức năng tạo mã vạch (Barcode) hàng loạt.")
