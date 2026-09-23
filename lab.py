import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import re
import difflib
import io
import json
import unicodedata
from streamlit_gsheets import GSheetsConnection

# ==========================================
# 1. CẤU HÌNH GIAO DIỆN & THIẾT KẾ (UI/UX)
# ==========================================
st.set_page_config(page_title="GC HATICO - Lab GC", page_icon="🔬", layout="wide")

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');

    html, body, [class*="css"] {
        font-family: 'Inter', sans-serif !important;
    }
    .stApp {
        background-color: #F8FAFC;
    }
    [data-testid="stSidebar"] {
        background-color: #FFFFFF;
        border-right: 1px solid #E2E8F0;
        box-shadow: 2px 0 10px rgba(0,0,0,0.02);
    }
    .main-title {
        background: linear-gradient(135deg, #0F172A 0%, #3B82F6 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        font-weight: 800;
        font-size: 2.2rem;
        padding-bottom: 0.5rem;
    }
    .sub-title {
        color: #1E293B;
        font-weight: 700;
        font-size: 1.4rem;
        margin-top: 1rem;
        margin-bottom: 1rem;
    }
    div[data-testid="stMetric"] {
        background-color: #FFFFFF;
        border: 1px solid #E2E8F0;
        border-radius: 12px;
        padding: 20px 24px;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05), 0 2px 4px -1px rgba(0, 0, 0, 0.03);
        border-left: 5px solid #3B82F6;
        transition: transform 0.2s ease-in-out, box-shadow 0.2s ease-in-out;
    }
    div[data-testid="stMetric"]:hover {
        transform: translateY(-4px);
        box-shadow: 0 10px 15px -3px rgba(37, 99, 235, 0.15);
        border-left: 5px solid #2563EB;
    }
    div[data-testid="stMetricValue"] {
        font-size: 2.2rem !important;
        font-weight: 700 !important;
        color: #0F172A;
    }
    .stButton > button {
        border-radius: 8px;
        font-weight: 600;
        padding: 0.5rem 1rem;
        border: 1px solid #CBD5E1;
        transition: all 0.2s ease;
    }
    .stButton > button[kind="primary"] {
        background: linear-gradient(135deg, #2563EB 0%, #1D4ED8 100%) !important;
        border: none !important;
        color: white !important;
        box-shadow: 0 4px 6px rgba(37, 99, 235, 0.25);
    }
    .stButton > button[kind="primary"]:hover {
        background: linear-gradient(135deg, #1D4ED8 0%, #1E40AF 100%) !important;
        box-shadow: 0 6px 12px rgba(37, 99, 235, 0.4);
        transform: translateY(-2px);
    }
    .stTabs [data-baseweb="tab-list"] {
        background-color: #FFFFFF;
        border-radius: 10px;
        padding: 5px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
        gap: 8px;
    }
    .stTabs [data-baseweb="tab"] {
        border-radius: 6px;
        padding: 10px 20px;
        color: #64748B;
        font-weight: 600;
        border: none;
        transition: all 0.2s;
    }
    .stTabs [data-baseweb="tab"]:hover {
        background-color: #F1F5F9;
        color: #0F172A;
    }
    .stTabs [aria-selected="true"] {
        background-color: #EFF6FF !important;
        color: #2563EB !important;
        box-shadow: 0 1px 2px rgba(0,0,0,0.05);
    }
    [data-testid="stDataFrame"] {
        border-radius: 10px;
        border: 1px solid #E2E8F0;
        box-shadow: 0 2px 4px rgba(0,0,0,0.02);
    }
    [data-testid="stExpander"] {
        background: #FFFFFF;
        border-radius: 10px;
        border: 1px solid #E2E8F0;
    }
    .warning-box {
        background-color: #FEF2F2;
        border-left: 4px solid #EF4444;
        padding: 15px;
        border-radius: 6px;
        margin-bottom: 20px;
    }
</style>
""", unsafe_allow_html=True)

SHEET_URL = "https://docs.google.com/spreadsheets/d/1F2wFnxboWTFWDMGUuBDRGB901a5EKgvazHxkCgBjjRU/edit?usp=sharing"

STATUSES = [
    "🔴 1. Chờ xử lý", "🟠 2. Đang xử lý mẫu", "🟡 3. Chờ chạy máy",
    "🔵 4. Đang chạy máy", "🟣 5. Đang tính số liệu", "🟢 6. Lưu kho", "⚫ 7. Đã tiêu hủy"
]

CHEM_SYSTEMS = ["GC-MS", "GC-FID", "Thermo", "Dùng chung"]
CHEM_TYPES = ["Chất chuẩn (IS/Surrogate)", "Dung môi", "Vật tư tiêu hao", "Khí chuẩn", "Khác"]
CHEM_STATUS = ["🟢 Còn nhiều", "🟡 Sắp hết", "🔴 Đã hết"]

# ==========================================
# 2. KẾT NỐI DATABASE
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

def load_chemical_data():
    try:
        df_chem = conn.read(spreadsheet=SHEET_URL, worksheet="QuanLyHoaChat", ttl=0) 
        if df_chem.empty or len(df_chem.columns) == 0 or "Tên Hóa Chất" not in df_chem.columns:
            raise Exception("Chưa có cấu trúc")
        df_chem['Hạn Sử Dụng'] = pd.to_datetime(df_chem['Hạn Sử Dụng'], errors='coerce').dt.date
        df_chem['Ngày Mở Nắp'] = pd.to_datetime(df_chem['Ngày Mở Nắp'], errors='coerce').dt.date
        return df_chem
    except:
        df_chem = pd.DataFrame(columns=["Hệ Máy", "Phân Loại", "Tên Hóa Chất", "Số Lô (Lot)", "Ngày Mở Nắp", "Hạn Sử Dụng", "Tình Trạng Kho", "Ghi Chú"])
        return df_chem

def save_chemical_data(df_chem):
    df_save = df_chem.copy()
    df_save['Hạn Sử Dụng'] = pd.to_datetime(df_save['Hạn Sử Dụng'], errors='coerce').dt.strftime('%Y-%m-%d')
    df_save['Ngày Mở Nắp'] = pd.to_datetime(df_save['Ngày Mở Nắp'], errors='coerce').dt.strftime('%Y-%m-%d')
    try:
        conn.update(spreadsheet=SHEET_URL, worksheet="QuanLyHoaChat", data=df_save)
    except:
        pass
    st.cache_data.clear()

if "df" not in st.session_state:
    st.session_state.df = load_data()
if "df_limit" not in st.session_state:
    st.session_state.df_limit = load_limit_config()
if "df_chem" not in st.session_state:
    st.session_state.df_chem = load_chemical_data()

# ==========================================
# 3. CÁC HÀM BỔ SUNG & XỬ LÝ SỐ LIỆU
# ==========================================
def clean_str(value):
    return "" if pd.isna(value) else str(value).strip()

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

def lab_norm(value):
    if value is None or pd.isna(value): return ''
    return ''.join(c for c in unicodedata.normalize('NFD',str(value).lower().replace('đ','d')) if unicodedata.category(c)!='Mn').strip()

def lab_local_answer(question, df):
    q = lab_norm(question)
    
    if any(x in q for x in ['c ban dau', 'thu hoi', 'r%', 'csurr', 'duong chuan']):
        return "💡 **R(%) = (C_surr_sau / C_surr_trước) × 100**.\nHệ thống tự động chọn C ban đầu từ danh sách mức cố định sao cho R(%) rơi vào 70–130%."
    if any(x in q for x in ['cong thuc', 'cach tinh']):
        return "💡 **Công thức SOP:**\n- Mẫu Nước: C = C đo × 100/R.\n- Mẫu Khí: C = C đo ÷ V × 100/R.\nHệ thống tự động gợi ý V=24L cho mẫu Khí (KT) và V=4L cho mẫu Xung quanh (KXQ/KLV)."
    if any(x in q for x in ['bien ban', 'huong dan', 'cach nhap', 'loi', 'mdl', 'loq']):
        return "💡 **Hướng dẫn:**\n- Tiếp nhận mẫu ở Tab 2.\n- Tính kết quả tự động ở GC-MS.\n- Tính tay & Pha chuẩn ở Công cụ Phân tích.\n- Lập Biên bản & In tem ở Báo cáo & Lập Biên bản."

    filters = []
    code = re.search(r'\b(?:ns|nt|nm|nn|kt|kxq|kkxq|klv)[.\-]?\d[\w.\-]*', q)
    if code: filters.append(('Mã Mẫu', code.group()))
    batch = re.search(r'\b\d{4}\.\d{2}\.\d{3}\b', q)
    if batch: filters.append(('Tên Mẻ', batch.group()))
    status_words = [('cho chay', '3'), ('dang chay', '4'), ('cho xu ly', '1'), ('dang xu ly', '2'), ('tinh so lieu', '5'), ('luu kho', '6'), ('tieu huy', '7')]
    for phrase, status in status_words:
        if phrase in q: filters.append(('Trạng Thái', status))
    if 'mau nuoc' in q: filters.append(('Nền Mẫu', 'Nước'))
    elif 'mau khi' in q: filters.append(('Nền Mẫu', 'Khí'))

    res_df = df.copy()
    for field, val in filters:
        if field == 'Trạng Thái':
            res_df = res_df[res_df[field].str.contains(val, na=False)]
        else:
            res_df = res_df[res_df[field].astype(str).str.lower().str.contains(val.lower(), na=False)]

    if not filters and not any(x in q for x in ['tong', 'bao nhieu', 'thong ke', 'danh sach', 'nhom theo', 'cua ai', 'ai dang giu']):
        return "🤔 Mình chưa hiểu ý bạn. Bạn có thể hỏi cụ thể:\n- *Tra mẫu NS-150826-004*\n- *Có bao nhiêu mẫu chờ chạy?*\n- *Ai đang giữ nhiều mẫu nhất?*"

    if any(phrase in q for phrase in ['theo nguoi', 'ai dang giu', 'cua ai']):
        users = res_df['Người Giữ'].replace("", "Chưa phân công").dropna().value_counts()
        ans = f"👥 **Phân bổ mẫu ({len(res_df)} mẫu):**\n"
        for user, count in users.items(): ans += f"- **{user}**: {count} mẫu\n"
        return ans
    elif 'theo trang thai' in q:
        stats = res_df['Trạng Thái'].value_counts()
        ans = f"📊 **Trạng thái ({len(res_df)} mẫu):**\n"
        for st_name, count in stats.items(): ans += f"- {st_name}: {count} mẫu\n"
        return ans

    if len(res_df) == 0:
        return "❌ Không tìm thấy mẫu nào khớp với dữ liệu bạn tìm kiếm."
    elif len(res_df) == 1:
        info = res_df.iloc[0]
        return f"🔍 **Đã tìm thấy mẫu {info['Mã Mẫu']}:**\n- **Thuộc mẻ:** {info['Tên Mẻ']}\n- **Nền:** {info['Nền Mẫu']}\n- **Trạng thái:** {info['Trạng Thái']}\n- **Người giữ:** {info['Người Giữ']}\n- **Chỉ tiêu:** {info['Chỉ Tiêu']}"
    else:
        ans = f"📋 **Tìm thấy {len(res_df)} mẫu khớp yêu cầu:**\n"
        for _, row in res_df.head(10).iterrows():
            ans += f"- **{row['Mã Mẫu']}** (Mẻ: {row['Tên Mẻ']} | Trạng thái: {row['Trạng Thái'][:4]})\n"
        if len(res_df) > 10: ans += "*... và nhiều mẫu khác.*"
        return ans

# ==========================================
# 4. THANH ĐIỀU HƯỚNG BÊN TRÁI (SIDEBAR)
# ==========================================
st.sidebar.markdown("<h2 style='text-align: center; color: #1E293B;'>🔬 GC HATICO</h2>", unsafe_allow_html=True)
st.sidebar.caption("<div style='text-align: center; margin-bottom: 20px;'>Phần mềm Quản lý Phòng Lab Tự động</div>", unsafe_allow_html=True)

menu = st.sidebar.radio("📌 ĐIỀU HƯỚNG CHÍNH", [
    "🏠 Trang chủ (Tổng quan)", 
    "📥 Quản lý Tiếp nhận", 
    "⚙️ Vận hành GC-MS",
    "🔥 Vận hành GC-FID",
    "🧬 Vận hành Thermo",
    "🧮 Tiện ích Phân tích",
    "📝 Báo cáo & Lập Biên bản",
    "🧪 Kiểm soát Hóa chất",
    "⚙️ Cấu hình Hệ thống"
])

st.sidebar.divider()

if st.sidebar.button("🔄 Cập nhật hệ thống", use_container_width=True):
    st.cache_data.clear()
    if 'results' in st.session_state:
        st.session_state.results_stale = True
    st.session_state.df = load_data()
    st.session_state.df_limit = load_limit_config() 
    st.session_state.df_chem = load_chemical_data()
    st.rerun()

st.sidebar.divider()

with st.sidebar.popover("💬 Chat với Trợ lý AI", use_container_width=True):
    st.markdown("👋 **Xin chào! Mình là Trợ lý LIMS.**")
    st.caption("Tra cứu tiến độ mẫu, mẻ, phân công hoặc hỏi về SOP.")
    
    if "chat_history" not in st.session_state:
        st.session_state.chat_history = []
        
    chat_container = st.container(height=350)
    with chat_container:
        for msg in st.session_state.chat_history:
            if msg["role"] == "user":
                st.markdown(f"👤 **Bạn:** {msg['content']}")
            else:
                st.info(f"🤖 **Bot:**\n{msg['content']}")
                
    with st.form("chat_form", clear_on_submit=True):
        user_input = st.text_input("Gõ câu hỏi...", placeholder="VD: Mẫu NS-01 đang ở đâu?")
        submit_btn = st.form_submit_button("Gửi")
        if submit_btn and user_input.strip():
            st.session_state.chat_history.append({"role": "user", "content": user_input})
            ans = lab_local_answer(user_input, st.session_state.df)
            st.session_state.chat_history.append({"role": "bot", "content": ans})
            st.rerun()
            
    if st.session_state.chat_history:
        if st.button("🧹 Xóa hội thoại", use_container_width=True):
            st.session_state.chat_history = []
            st.rerun()

df_current = st.session_state.df.copy()
df_current["Ngày Nhận"] = df_current["Giờ Nhận"].dt.date
today_date = datetime.today().date()

# ==========================================
# 5. GIAO DIỆN CÁC TRANG
# ==========================================

if menu == "🏠 Trang chủ (Tổng quan)":
    st.markdown("<h1 class='main-title'>📊 Bảng Điều Khiển Trung Tâm</h1>", unsafe_allow_html=True)
    
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
    col4.metric("✅ Đã hoàn thành", tong_hoan_thanh, f"Lưu/Hủy", delta_color="off")
    
    st.markdown("<div class='sub-title'>Tra cứu & Cập nhật Trạng thái</div>", unsafe_allow_html=True)
    
    with st.container(border=True):
        col_date, col_status, col_search = st.columns([1.2, 1.5, 2])
        with col_date: selected_date = st.date_input("📅 Chọn Ngày Giao Mẫu:", today_date)
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

    st.caption("✨ **Mẹo:** Chọn dòng và bấm `Delete` để xóa. Bấm đúp vào ô để sửa dữ liệu.")
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
        use_container_width=True, num_rows="dynamic", key="data_editor", height=450
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
        st.success("✅ Đã đồng bộ thành công lên cơ sở dữ liệu chung!")
        st.rerun()

elif menu == "📥 Quản lý Tiếp nhận":
    st.markdown("<h1 class='main-title'>📥 Khu vực Tiếp nhận mẫu mới</h1>", unsafe_allow_html=True)
    
    tab_excel, tab_thu_cong = st.tabs(["📁 Nạp qua Excel tự động", "✍️ Nhập Mẫu Lẻ"])
    
    with tab_excel:
        with st.container(border=True):
            st.info("💡 **Tính năng AI:** Tự động loại bỏ các ô bị bôi xám (#808080) (Các chỉ tiêu khách không yêu cầu).")
            uploaded_file = st.file_uploader("Kéo thả file KetQuaMeThuNghiem...xlsx vào đây", type=["xlsx", "xls"])
            
            if uploaded_file is not None:
                try:
                    import openpyxl
                    wb = openpyxl.load_workbook(io.BytesIO(uploaded_file.getvalue()), data_only=True)
                    ws = wb['Kết quả'] if 'Kết quả' in wb.sheetnames else wb.worksheets[0]
                    
                    batch, header_row, khm_col = "Không xác định", None, None
                    
                    for row in ws.iter_rows(min_row=1, max_row=min(30, ws.max_row)):
                        for cell in row:
                            val = clean_str(cell.value)
                            if val.startswith('Số:'):
                                batch = val.replace("Số:", "").strip()
                            if val.upper() == 'KHM':
                                header_row, khm_col = cell.row, cell.column
                                
                    if header_row is not None:
                        params_info = []
                        for col in range(khm_col + 1, ws.max_column + 1):
                            header_val = clean_str(ws.cell(header_row, col).value)
                            if header_val.casefold() == 'ghi chú' or not header_val: break
                            
                            header_val = re.sub(r'^\d+\.\s*', '', header_val)
                            header_val = re.sub(r'\s*\(chọn cái này\)', '', header_val, flags=re.I).strip()
                            params_info.append((col, header_val))
                        
                        samples_data = []
                        for r in range(header_row + 1, ws.max_row + 1):
                            khm_val = clean_str(ws.cell(r, khm_col).value)
                            if len(khm_val) < 3 or khm_val.lower() == 'nan': continue
                            
                            nen_mau_auto = "Khí" if khm_val.upper().startswith(("KT", "KKXQ", "KLV")) else ("Nước" if khm_val.upper().startswith(("NS", "NT", "NM", "NN")) else "Chưa xác định")

                            selected_params = []
                            for col_idx, param_name in params_info:
                                fill = ws.cell(r, col_idx).fill
                                color = fill.fgColor
                                
                                is_gray = fill.patternType == 'solid' and color.type == 'rgb' and str(color.rgb)[-6:].upper() == '808080'
                                if not is_gray: selected_params.append(param_name)
                                    
                            chuoi_chi_tieu = "; ".join(selected_params) if selected_params else "Chưa xác định"
                            samples_data.append({"Chọn": True, "Mã Mẫu": khm_val, "Nền Mẫu": nen_mau_auto, "Chỉ Tiêu": chuoi_chi_tieu})
                        
                        if len(samples_data) > 0:
                            st.success(f"✔️ Quét thành công **{len(samples_data)}** mẫu thuộc mẻ: **{batch}**")
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
                                new_rows = [{"Mã Mẫu": row["Mã Mẫu"], "Tên Mẻ": batch, "Nền Mẫu": row["Nền Mẫu"], "Chỉ Tiêu": row["Chỉ Tiêu"], "Trạng Thái": STATUSES[0], "Người Giữ": batch_nguoi, "Ghi Chú": "Import Excel", "Giờ Nhận": datetime.now()} for _, row in selected_samples.iterrows()]
                                st.session_state.df = pd.concat([st.session_state.df, pd.DataFrame(new_rows)], ignore_index=True)
                                save_data(st.session_state.df)
                                st.success("✅ Đã nạp thành công vào hệ thống!")
                                st.rerun()
                        else: st.warning("Không tìm thấy dữ liệu mẫu hợp lệ bên dưới ô KHM.")
                    else: st.error("Không tìm thấy ô 'KHM' trong file Excel!")
                except Exception as e: st.error(f"Lỗi đọc file: {e}")

    with tab_thu_cong:
        with st.container(border=True):
            with st.form("add_sample_form", clear_on_submit=True):
                st.markdown("<div class='sub-title'>Bổ sung Mẫu Lẻ</div>", unsafe_allow_html=True)
                new_id = st.text_input("Mã Mẫu (VD: NS-1509-01)*")
                new_name = st.text_input("Tên Mẻ (VD: 2026.07.017)")
                col_t1, col_t2 = st.columns(2)
                with col_t1: new_nen = st.selectbox("Nền Mẫu", ["Nước", "Khí"])
                with col_t2: new_chi_tieu = st.text_input("Chỉ tiêu đo")
                new_nguoi = st.text_input("Người tiếp nhận (Ký tên)")
                
                if st.form_submit_button("Thêm Mẫu lẻ", type="primary") and new_id:
                    st.session_state.df = pd.concat([st.session_state.df, pd.DataFrame([{"Mã Mẫu": new_id, "Tên Mẻ": new_name, "Nền Mẫu": new_nen, "Chỉ Tiêu": new_chi_tieu, "Trạng Thái": STATUSES[0], "Người Giữ": new_nguoi, "Ghi Chú": "", "Giờ Nhận": datetime.now()}])], ignore_index=True)
                    save_data(st.session_state.df)
                    st.success(f"✅ Đã thêm mẫu {new_id} thành công!")

elif menu == "⚙️ Vận hành GC-MS":
    st.markdown("<h1 class='main-title'>⚙️ Phân tích & Vận hành Máy đo GC-MS</h1>", unsafe_allow_html=True)
    
    tab_seq, tab_auto = st.tabs(["1. Xuất Sequence Chạy Máy", "2. Xử lý Tự động (File PDF/Excel)"])
    
    with tab_seq:
        with st.container(border=True):
            st.markdown("<div class='sub-title'>Xuất Sequence Chạy Máy</div>", unsafe_allow_html=True)
            df_ready = st.session_state.df[st.session_state.df["Trạng Thái"] == "🟡 3. Chờ chạy máy"]
            st.write(f"Đang có **{len(df_ready)}** mẫu trong hàng chờ.")
            
            if not df_ready.empty:
                seq_df = pd.DataFrame({'Vial': range(1, len(df_ready) + 1), 'Sample Name': df_ready['Mã Mẫu'], 'Sample Type': 'Sample'})
                seq_df['Method'] = df_ready['Chỉ Tiêu'].apply(lambda x: 'VOCs.M' if any(k in str(x).upper() for k in ['VOC', 'BENZEN', 'TOLUEN', 'CHLORO', 'STYREN']) else 'HCHO.M')
                seq_df['Data File'] = datetime.now().strftime("%Y%m%d") + "_" + df_ready['Mã Mẫu']
                st.download_button("📥 Tải File Sequence.csv", data=seq_df.to_csv(index=False).encode('utf-8'), file_name=f"MassHunter_Seq_{datetime.now().strftime('%Y%m%d')}.csv", mime="text/csv", type="primary")
            
    with tab_auto:
        with st.container(border=True):
            st.markdown("<div class='sub-title'>Xử lý Kết quả Hàng loạt (SOP)</div>", unsafe_allow_html=True)
            st.info("💡 Tự động bóc tách số liệu, nội suy nồng độ $C_{surr}$ chuẩn và so khớp Giới hạn MDL/LOQ theo đúng chuẩn phòng Lab.")
            
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
                        st.success(f"✅ Đã xử lý {len(calc_results)} dòng kết quả. Tự động áp dụng tiêu chuẩn Khí/Nước.")
                        
                        df_results = pd.DataFrame(calc_results)
                        df_results = df_results.sort_values(by=["Tên mẫu", "Tên chỉ tiêu"]).reset_index(drop=True)
                        
                        st.session_state.results = df_results
                        st.session_state.results_stale = False
                        
                        display_cols = ["Tên mẫu", "Tên chỉ tiêu", "C đo", "C thực", "Giới hạn", "R(%)"]
                        st.dataframe(df_results[display_cols], use_container_width=True, hide_index=True)
                        
                        csv_results = df_results.to_csv(index=False).encode('utf-8-sig')
                        st.download_button(
                            label="📥 Tải Kết quả (CSV) để Lưu trữ",
                            data=csv_results,
                            file_name=f"Ket_Qua_GC_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
                            mime="text/csv",
                            type="primary"
                        )
                        
                        processed_samples = df_results["Tên mẫu"].unique().tolist()
                        for smp in processed_samples:
                            mask = st.session_state.df["Mã Mẫu"] == smp
                            if mask.any():
                                st.session_state.df.loc[mask, "Trạng Thái"] = "🟣 5. Đang tính số liệu"
                        save_data(st.session_state.df)
                        
                    else: 
                        st.warning("⚠️ Báo cáo không chứa mẫu hợp lệ (KT, KXQ, NS, NT...) hoặc thiếu dữ liệu phân tích.")
                except Exception as e: st.error(f"❌ Lỗi xử lý: {e}")

elif menu == "🔥 Vận hành GC-FID":
    st.markdown("<h1 class='main-title'>🔥 Hệ thống GC-FID (Agilent)</h1>", unsafe_allow_html=True)
    st.caption("Module chuyên biệt xử lý dữ liệu từ đầu dò FID")
    
    col_seq, col_import = st.columns(2)
    with col_seq:
        with st.container(border=True):
            st.markdown("<div class='sub-title'>1. Xuất Sequence GC-FID</div>", unsafe_allow_html=True)
            st.info("Sẽ tích hợp thuật toán xuất file Sequence định dạng cho máy GC-FID Agilent.")
        
    with col_import:
        with st.container(border=True):
            st.markdown("<div class='sub-title'>2. Xử lý kết quả GC-FID</div>", unsafe_allow_html=True)
            st.info("Khu vực chờ tích hợp thuật toán đọc file báo cáo từ máy GC-FID.")
            fid_file = st.file_uploader("Kéo thả báo cáo GC-FID (PDF/Excel/CSV/TXT)", type=["pdf", "xlsx", "xls", "csv", "txt"])
            if fid_file:
                st.warning("🚧 Hệ thống đang chờ cập nhật thuật toán bóc tách dữ liệu từ file report FID.")

elif menu == "🧬 Vận hành Thermo":
    st.markdown("<h1 class='main-title'>🧬 Hệ thống Thermo GC-MS</h1>", unsafe_allow_html=True)
    st.caption("Module chuyên biệt xử lý dữ liệu OCP, OPP, PCB và Phenol")
    
    col_seq, col_import = st.columns(2)
    with col_seq:
        with st.container(border=True):
            st.markdown("<div class='sub-title'>1. Xuất Sequence Thermo</div>", unsafe_allow_html=True)
            st.info("Sẽ tích hợp thuật toán xuất file Sequence định dạng tương thích phần mềm Thermo.")
        
    with col_import:
        with st.container(border=True):
            st.markdown("<div class='sub-title'>2. Xử lý kết quả Thermo</div>", unsafe_allow_html=True)
            st.info("Khu vực chờ tích hợp thuật toán đọc file xuất từ máy Thermo.")
            thermo_file = st.file_uploader("Kéo thả báo cáo Thermo (PDF/Excel/CSV)", type=["pdf", "xlsx", "xls", "csv"])
            if thermo_file:
                st.warning("🚧 Hệ thống đang chờ cập nhật thuật toán bóc tách dữ liệu từ file report Thermo.")

elif menu == "🧮 Tiện ích Phân tích":
    st.markdown("<h1 class='main-title'>🧮 Tiện ích Phân tích & Tính toán</h1>", unsafe_allow_html=True)
    
    tab_manual, tab_calib = st.tabs(["1. Tính toán & Nhập liệu Thủ công", "2. Pha Đường Chuẩn Đa Thành Phần"])
    
    with tab_manual:
        with st.container(border=True):
            st.markdown("<div class='sub-title'>Tính toán & Nhập liệu Thủ công (Gồm tính Tổng)</div>", unsafe_allow_html=True)
            st.info("💡 Điền thông số đo để máy tự tính, hoặc nhập thẳng vào cột 'Kết quả'. Có thể thêm dòng chỉ tiêu bị thiếu. Đánh dấu các chất cần tính dồn để cộng thành một chỉ tiêu Tổng chung.")
            
            valid_samples = df_current[df_current['Chỉ Tiêu'].str.strip() != ""]['Mã Mẫu'].tolist()
            selected_sample_manual = st.selectbox("🔍 Chọn Mã Mẫu để nhập liệu:", ["-- Chọn mẫu --"] + valid_samples)
            
            if selected_sample_manual != "-- Chọn mẫu --":
                sample_info = df_current[df_current['Mã Mẫu'] == selected_sample_manual].iloc[0]
                chi_tieu_raw = str(sample_info['Chỉ Tiêu'])
                chi_tieu_list = [ct.strip() for ct in re.split(r'[;,]', chi_tieu_raw) if ct.strip()]
                
                v_param_def, nen_mau_code, loai_mau_def = parse_sample_matrix(selected_sample_manual)
                if v_param_def is None:
                    v_param_def, nen_mau_code, loai_mau_def = 24.0, 'KT', 'Khí'
                
                st.write(f"**Phân loại:** {loai_mau_def} ({nen_mau_code}) | **Chỉ tiêu theo Database:** {len(chi_tieu_list)}")
                
                manual_data = []
                for ct in chi_tieu_list:
                    mdl_val, loq_val, unit = get_limit_info(ct, nen_mau_code)
                    limit_str = ""
                    if loai_mau_def == 'Khí' and mdl_val is not None: limit_str = f"MDL: {mdl_val} {unit}"
                    elif loai_mau_def == 'Nước' and loq_val is not None: limit_str = f"LOQ: {loq_val} {unit}"

                    manual_data.append({
                        "Cộng Tổng": False,
                        "Chỉ Tiêu": ct,
                        "Giới hạn (Tới hạn)": limit_str,
                        "Độ làm giàu (V)": v_param_def,
                        "C_surr thực": 10.0,
                        "C_surr đo": 10.0,
                        "C_đo chỉ tiêu": 0.0,
                        "Kết quả (Ghi đè)": ""
                    })
                
                df_manual = pd.DataFrame(manual_data)
                
                st.caption("Có thể thêm hàng mới (dấu + ở dưới bảng) nếu mẫu có chỉ tiêu con bị sót.")
                edited_manual = st.data_editor(
                    df_manual,
                    num_rows="dynamic",
                    column_config={
                        "Cộng Tổng": st.column_config.CheckboxColumn("Cộng Tổng?"),
                        "Chỉ Tiêu": st.column_config.TextColumn("Chỉ Tiêu"),
                        "Giới hạn (Tới hạn)": st.column_config.TextColumn(disabled=True),
                        "Độ làm giàu (V)": st.column_config.NumberColumn(format="%.2f"),
                        "C_surr thực": st.column_config.NumberColumn(format="%.3f"),
                        "C_surr đo": st.column_config.NumberColumn(format="%.3f"),
                        "C_đo chỉ tiêu": st.column_config.NumberColumn(format="%.4f"),
                        "Kết quả (Ghi đè)": st.column_config.TextColumn("Kết quả (Tự nhập)")
                    },
                    use_container_width=True,
                    hide_index=True,
                    key=f"manual_editor_{selected_sample_manual}"
                )
                
                st.markdown("---")
                total_param_name = st.text_input("📝 Nhập tên chỉ tiêu Tổng (Chỉ áp dụng nếu có tick chọn 'Cộng Tổng' ở bảng trên):", placeholder="VD: Tổng VOCs, Tổng PCB...")
                
                if st.button("💾 Tính Toán & Lưu Mẫu Này", type="primary"):
                    new_results = []
                    sum_val = 0.0
                    has_sum = False
                    
                    for _, row in edited_manual.iterrows():
                        ct = str(row["Chỉ Tiêu"]).strip()
                        if not ct or ct == "nan": continue
                        
                        v = float(row.get("Độ làm giàu (V)", v_param_def) or v_param_def)
                        c_truoc = float(row.get("C_surr thực", 10.0) or 10.0)
                        c_sau = float(row.get("C_surr đo", 10.0) or 10.0)
                        c_do = float(row.get("C_đo chỉ tiêu", 0.0) or 0.0)
                        override_res = str(row.get("Kết quả (Ghi đè)", "")).strip()
                        is_sum = bool(row.get("Cộng Tổng", False))
                        
                        recovery = (c_sau / c_truoc) * 100.0 if c_truoc > 0 else 100.0
                        mdl_val, loq_val, unit = get_limit_info(ct, nen_mau_code)
                        
                        limit_display = ""
                        if loai_mau_def == 'Khí' and mdl_val is not None: limit_display = f"MDL: {mdl_val} {unit}"
                        elif loai_mau_def == 'Nước' and loq_val is not None: limit_display = f"LOQ: {loq_val} {unit}"
                        
                        c_thuc_str = ""
                        c_thuc_numeric = 0.0
                        
                        if override_res and override_res.lower() != "nan":
                            c_thuc_str = override_res
                            try: c_thuc_numeric = float(override_res.replace(',', '.'))
                            except: c_thuc_numeric = 0.0
                        else:
                            c_thuc_str = evaluate_result(c_do, v, mdl_val, loq_val, unit, loai_mau_def, recovery)
                            if c_thuc_str.startswith("KPH") or c_thuc_str.startswith("<"):
                                c_thuc_numeric = 0.0
                            else:
                                try: c_thuc_numeric = float(c_thuc_str.replace(',', '.'))
                                except: c_thuc_numeric = 0.0
                        
                        if is_sum:
                            has_sum = True
                            sum_val += c_thuc_numeric
                            
                        new_results.append({
                            "Tên mẫu": selected_sample_manual,
                            "Tên chỉ tiêu": ct,
                            "C đo": round(c_do, 4),
                            "C thực": c_thuc_str,
                            "Giới hạn": limit_display,
                            "R(%)": f"{round(recovery, 1)}%",
                            "C_surr_truoc": c_truoc,
                            "C_surr_sau": c_sau
                        })
                    
                    if has_sum and total_param_name:
                        new_results.append({
                            "Tên mẫu": selected_sample_manual,
                            "Tên chỉ tiêu": total_param_name,
                            "C đo": "",
                            "C thực": str(round(sum_val, 4)).replace('.', ','),
                            "Giới hạn": "", 
                            "R(%)": "",
                            "C_surr_truoc": "",
                            "C_surr_sau": ""
                        })
                    
                    new_df = pd.DataFrame(new_results)
                    if 'results' not in st.session_state or st.session_state.results.empty:
                        st.session_state.results = new_df
                    else:
                        existing = st.session_state.results
                        existing = existing[existing["Tên mẫu"] != selected_sample_manual]
                        st.session_state.results = pd.concat([existing, new_df], ignore_index=True)
                    
                    st.session_state.results_stale = False
                    
                    mask = st.session_state.df["Mã Mẫu"] == selected_sample_manual
                    if mask.any():
                        st.session_state.df.loc[mask, "Trạng Thái"] = "🟣 5. Đang tính số liệu"
                        save_data(st.session_state.df)
                        
                    st.success(f"✅ Đã lưu kết quả cho mẫu {selected_sample_manual}! Hệ thống đã cộng dồn vào danh sách tổng chờ xuất Biên bản.")
                    st.rerun()

    with tab_calib:
        with st.container(border=True):
            st.markdown("<div class='sub-title'>🧪 Lập Công Thức Pha Chuẩn Tối Ưu (Smart Mix)</div>", unsafe_allow_html=True)
            st.info("💡 Hệ thống hỗ trợ tối ưu hóa quy trình pha chuẩn: Gộp các Nội chuẩn/Surrogate thành 1 Master Mix để hút 1 lần; Gộp các Mix chuẩn gốc thành Mix làm việc (Working Standard) để tránh sai số pipet nhỏ.")
            
            col_v1, col_v2 = st.columns(2)
            with col_v1:
                v_final = st.number_input("Thể tích định mức mỗi điểm chuẩn (mL)", value=1.0, step=0.1, min_value=0.1)
            with col_v2:
                levels_input = st.text_input("🎯 Nhập dãy nồng độ chuẩn cần pha (ppm):", "0.5, 1, 2, 5, 10, 20")
            
            st.markdown("### ⚙️ Tùy chọn Tối ưu hóa (Optimization)")
            col_opt1, col_opt2 = st.columns(2)
            with col_opt1:
                opt_mix_is_surr = st.toggle("🧪 Gộp IS & Surrogate thành Master Mix", value=True)
                if opt_mix_is_surr:
                    v_spike_is = st.number_input("Thể tích hút Master Mix cho mỗi vial (µL)", value=50.0, step=10.0)
            with col_opt2:
                opt_mix_std = st.toggle("🧪 Gộp các Chuẩn gốc thành Mix Trung gian", value=True)
                if opt_mix_std:
                    c_ws_std = st.number_input("Nồng độ Mix Trung gian cần pha (ppm)", value=100.0, step=10.0)
                    v_ws_total = st.number_input("Thể tích cần pha Mix Trung gian (µL)", value=1000.0, step=100.0)

            st.markdown("**1. Các dung dịch Chuẩn (Mix) thay đổi theo dãy nồng độ:**")
            if "df_mix_default" not in st.session_state:
                st.session_state.df_mix_default = pd.DataFrame([{"Tên Mix Chuẩn": "Mix VOCs", "C_gốc (ppm)": 1000.0}])
            df_mix = st.data_editor(st.session_state.df_mix_default, num_rows="dynamic", use_container_width=True, key="mix_editor")
            
            st.markdown("**2. Các dung dịch Nội chuẩn (IS) & Đồng hành (Surrogate) cố định:**")
            if "df_is_default" not in st.session_state:
                st.session_state.df_is_default = pd.DataFrame([
                    {"Phân Loại": "Nội chuẩn (IS)", "Tên Hợp Chất": "Fluorobenzene", "C_gốc (ppm)": 1000.0, "C_đích mỗi vial (ppm)": 10.0},
                    {"Phân Loại": "Surrogate", "Tên Hợp Chất": "Toluene-D8", "C_gốc (ppm)": 1000.0, "C_đích mỗi vial (ppm)": 10.0}
                ])
            df_is_surr = st.data_editor(st.session_state.df_is_default, num_rows="dynamic", use_container_width=True, key="is_surr_editor", column_config={"Phân Loại": st.column_config.SelectboxColumn(options=["Nội chuẩn (IS)", "Surrogate", "Khác"])})
            
            if st.button("🚀 Tính Toán Bảng Pha Chuẩn & Tối Ưu", type="primary"):
                try:
                    levels = sorted([float(x.strip()) for x in levels_input.split(",") if x.strip()])
                    calib_data = []
                    instructions = []

                    is_surr_columns = {} 
                    total_is_v_per_vial = 0.0
                    
                    if opt_mix_is_surr:
                        v_master_total = (len(levels) + 3) * v_spike_is
                        instructions.append("### 🧪 BƯỚC 1: PHA MASTER MIX NỘI CHUẨN (IS) & SURROGATE")
                        instructions.append(f"*(Pha tổng cộng {v_master_total} µL Master Mix dùng chung cho toàn bộ các điểm chuẩn. Mỗi điểm chuẩn sẽ hút {v_spike_is} µL)*")
                        sum_v = 0.0
                        for _, row in df_is_surr.iterrows():
                            name = str(row.get("Tên Hợp Chất", "")).strip()
                            if not name or name == "nan": continue
                            c_s = float(row.get("C_gốc (ppm)", 0))
                            c_t = float(row.get("C_đích mỗi vial (ppm)", 0))
                            v_stock = (c_t * v_final * 1000 * v_master_total) / (c_s * v_spike_is) if c_s > 0 else 0
                            instructions.append(f"- Hút **{v_stock:.2f} µL** {name} gốc ({c_s} ppm)")
                            sum_v += v_stock
                        v_solv = v_master_total - sum_v
                        instructions.append(f"- Thêm **{v_solv:.2f} µL** dung môi. Lắc đều.")
                        if sum_v > v_master_total:
                            instructions.append("❌ **LỖI:** Thể tích các chất gốc vượt quá tổng thể tích Master Mix. Hãy tăng Thể tích hút mỗi vial.")
                        
                        is_surr_columns["Hút IS/Surr Master Mix (µL)"] = v_spike_is
                        total_is_v_per_vial = v_spike_is
                    else:
                        for _, row in df_is_surr.iterrows():
                            name = str(row.get("Tên Hợp Chất", "")).strip()
                            if not name or name == "nan": continue
                            c_s = float(row.get("C_gốc (ppm)", 0))
                            c_t = float(row.get("C_đích mỗi vial (ppm)", 0))
                            v_ul = (c_t * v_final * 1000) / c_s if c_s > 0 else 0
                            prefix = "IS" if "IS" in str(row.get("Phân Loại", "")) else "Surr"
                            is_surr_columns[f"Hút {prefix}: {name} (µL)"] = v_ul
                            total_is_v_per_vial += v_ul

                    if opt_mix_std:
                        instructions.append("### 🧪 BƯỚC 2: PHA MIX CHUẨN LÀM VIỆC (WORKING STANDARD)")
                        instructions.append(f"*(Gộp các Mix chuẩn gốc thành Mix trung gian duy nhất có nồng độ {c_ws_std} ppm. Thể tích pha: {v_ws_total} µL)*")
                        sum_v = 0.0
                        for _, r in df_mix.iterrows():
                            name = str(r.get("Tên Mix Chuẩn", "")).strip()
                            if not name or name == "nan": continue
                            c_s = float(r.get("C_gốc (ppm)", 0))
                            v_stock = (c_ws_std * v_ws_total) / c_s if c_s > 0 else 0
                            instructions.append(f"- Hút **{v_stock:.2f} µL** {name} ({c_s} ppm)")
                            sum_v += v_stock
                        v_solv = v_ws_total - sum_v
                        instructions.append(f"- Thêm **{v_solv:.2f} µL** dung môi. Lắc đều.")
                        if sum_v > v_ws_total:
                            instructions.append("❌ **LỖI:** Thể tích chuẩn gốc vượt quá thể tích Mix trung gian. Hãy tăng Thể tích pha hoặc giảm nồng độ trung gian.")

                    instructions.append("### 🧪 BƯỚC 3: PHA DÃY CHUẨN VÀO VIAL CUỐI CÙNG")
                    for lvl in levels:
                        row_data = {"Điểm chuẩn": f"Level {lvl} ppm"}
                        total_std_v = 0.0
                        
                        if opt_mix_std:
                            v_ws = (lvl * v_final * 1000) / c_ws_std if c_ws_std > 0 else 0
                            row_data[f"Hút Mix Làm Việc {c_ws_std}ppm (µL)"] = round(v_ws, 2)
                            total_std_v += v_ws
                        else:
                            for _, r in df_mix.iterrows():
                                name = str(r.get("Tên Mix Chuẩn", "")).strip()
                                if not name or name == "nan": continue
                                c_s = float(r.get("C_gốc (ppm)", 0))
                                v_ul = (lvl * v_final * 1000) / c_s if c_s > 0 else 0
                                row_data[f"Hút Mix {name} (µL)"] = round(v_ul, 2)
                                total_std_v += v_ul
                        
                        for k, v in is_surr_columns.items():
                            row_data[k] = round(v, 2)
                            
                        v_dungmoi = (v_final * 1000) - total_std_v - total_is_v_per_vial
                        row_data["Dung môi bù (µL)"] = round(v_dungmoi, 2) if v_dungmoi >= 0 else "Quá thể tích!"
                        row_data["V tổng đích (mL)"] = v_final
                        calib_data.append(row_data)

                    for ins in instructions:
                        st.markdown(ins)
                    
                    df_calib = pd.DataFrame(calib_data)
                    st.dataframe(df_calib, use_container_width=True, hide_index=True)
                    
                    st.download_button("📥 Tải Bảng Pha Chuẩn (CSV)", data=df_calib.to_csv(index=False).encode('utf-8-sig'), file_name=f"Quy_trinh_Pha_chuan_{datetime.now().strftime('%Y%m%d')}.csv", mime="text/csv")
                except Exception as e:
                    st.error(f"Lỗi tính toán: {e}. Vui lòng kiểm tra lại dữ liệu.")

elif menu == "📝 Báo cáo & Lập Biên bản":
    st.markdown("<h1 class='main-title'>📝 Báo cáo & Lập Biên Bản</h1>", unsafe_allow_html=True)
    
    tab_report, tab_qr = st.tabs(["1. Lập Biên Bản Xử Lý Mẫu", "2. Sinh Mã Vạch QR"])

    with tab_report: 
        st.info("💡 **Hệ thống Thông minh:** Tự động lấy kết quả bạn vừa tính ở tab Phân tích để điền vào Biểu mẫu Word/Excel. Mọi định dạng Form, Chữ ký được bảo lưu 100%.")
        
        with st.container(border=True):
            col_tpl, col_data = st.columns(2)
            with col_tpl:
                template_file = st.file_uploader("1. Tải lên Form Mẫu (.docx, .xlsx)", type=["docx", "xlsx"])
            with col_data:
                data_file = st.file_uploader("2. File Số liệu (.csv) [Tùy chọn tải tay]", type=["xlsx", "csv"])

            if template_file:
                try:
                    if data_file is not None:
                        df_kq = pd.read_excel(data_file) if data_file.name.endswith('.xlsx') else pd.read_csv(data_file)
                    else:
                        df_kq = st.session_state.get('results', pd.DataFrame())
                    
                    if df_kq.empty or "Tên mẫu" not in df_kq.columns or "Tên chỉ tiêu" not in df_kq.columns:
                        st.error("⚠️ Hệ thống đang không có số liệu kết quả trong bộ nhớ. Bạn hãy qua tab **Vận hành Máy Đo** hoặc **Tiện ích Phân tích** tính số liệu trước, hoặc thả thủ công file `.csv` kết quả vào ô số 2 nhé!")
                    elif st.session_state.get('results_stale') and data_file is None:
                        st.warning("⚠️ **Cảnh báo lệch số liệu:** Thư viện giới hạn MDL/LOQ vừa bị sửa. Bạn nên quay lại tab phân tích bấm 'Tính lại' để kết quả được cập nhật chuẩn xác nhất!")
                    else:
                        ket_qua_list = []
                        danh_sach_mau = []
                        grouped = df_kq.groupby("Tên mẫu")
                        
                        for ten_mau, group in grouped:
                            first_row = group.iloc[0]
                            ten_mau_str = str(ten_mau)
                            
                            v_khi = "24,0" if "KT" in ten_mau_str.upper() else ("4,0" if any(k in ten_mau_str.upper() for k in ["KXQ", "KLV"]) else "")
                            h_phantram = str(first_row.get("R(%)", "")).replace('%', '').strip()
                            c_surr_truoc = str(first_row.get("C_surr_truoc", "")).replace('.', ',')
                            c_surr_sau = str(first_row.get("C_surr_sau", "")).replace('.', ',')
                            
                            if c_surr_sau == "" and h_phantram != "":
                                try: c_surr_sau = str(round((float(h_phantram) / 100) * 10.0, 2)).replace('.', ',')
                                except: pass

                            mau_dict = {
                                "ngay": datetime.now().strftime("%d/%m/%Y"),
                                "ky_hieu": ten_mau_str,
                                "c_surr_truoc": c_surr_truoc, 
                                "c_surr_sau": c_surr_sau, 
                                "de": h_phantram,
                                "ghi_chu": ""
                            }
                            
                            for _, row in group.iterrows():
                                chi_tieu = str(row.get("Tên chỉ tiêu", ""))
                                c_do = str(row.get("C đo", "")).replace('.', ',')
                                kq_thuc = str(row.get("C thực", "")).replace('.', ',')
                                
                                ket_qua_list.append({
                                    "ngay": datetime.now().strftime("%d/%m/%Y"),
                                    "ten_mau": ten_mau_str,
                                    "chi_tieu": chi_tieu,
                                    "c_surr_truoc": c_surr_truoc,
                                    "c_surr_sau": c_surr_sau,
                                    "h_phantram": h_phantram,
                                    "v_khi": v_khi,
                                    "c_do": c_do,
                                    "kq_thuc": kq_thuc
                                })
                                
                                chi_tieu_upper = chi_tieu.upper()
                                slug = re.sub(r'\W+', '', chi_tieu_upper.lower())
                                mau_dict[f"cdo_{slug}"] = c_do
                                mau_dict[f"kq_{slug}"] = kq_thuc
                                
                                if "BENZENE" in chi_tieu_upper or "BENZEN" in chi_tieu_upper:
                                    mau_dict["cdo_benzen"] = c_do
                                    mau_dict["kq_benzen"] = kq_thuc
                                elif "TOLUENE" in chi_tieu_upper or "TOLUEN" in chi_tieu_upper:
                                    mau_dict["cdo_toluen"] = c_do
                                    mau_dict["kq_toluen"] = kq_thuc
                                    
                            danh_sach_mau.append(mau_dict)

                        if template_file.name.endswith('.docx'):
                            if st.button("🚀 Bơm Dữ Liệu Lập Biên Bản (Word)", type="primary"):
                                try:
                                    from docxtpl import DocxTemplate
                                    import io
                                    
                                    doc = DocxTemplate(template_file)
                                    doc.render({"danh_sach_mau": danh_sach_mau, "ket_qua": ket_qua_list})
                                    
                                    bio = io.BytesIO()
                                    doc.save(bio)
                                    bio.seek(0)
                                    
                                    st.success("🎉 Cập nhật số liệu thành công!")
                                    st.download_button("📥 Tải Xuống Biên Bản (.docx)", data=bio, file_name=f"Bien_Ban_{datetime.now().strftime('%Y%m%d_%H%M')}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
                                except ImportError:
                                    st.error("⚠️ Máy chủ chưa cài thư viện 'docxtpl'. Hãy nhớ thêm 'docxtpl' vào file requirements.txt trên GitHub nhé!")
                                    
                        elif template_file.name.endswith('.xlsx'):
                            if st.button("🚀 Bơm Dữ Liệu Lập Biên Bản (Excel)", type="primary"):
                                import openpyxl
                                from copy import copy
                                import io
                                
                                wb = openpyxl.load_workbook(template_file)
                                ws = wb.active
                                
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
                                    is_nuoc = any(isinstance(v, str) and ("kq_benzen" in v or "kq_toluen" in v or "mau.ky_hieu" in v) for v in template_cells)
                                    data_loop = danh_sach_mau if is_nuoc else ket_qua_list
                                    
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
                                    
                                    if len(data_loop) > 1 and ws.max_row > template_row_idx:
                                        ws.move_range(f"A{template_row_idx+1}:{ws.cell(ws.max_row, ws.max_column).coordinate}", rows=len(data_loop)-1, translate=True)
                                    
                                    current_row = template_row_idx
                                    for item in data_loop:
                                        for col_idx, cell_val in enumerate(template_cells, start=1):
                                            new_val = cell_val
                                            if cell_val and isinstance(cell_val, str):
                                                new_val = re.sub(r'\{%p.*?%\}', '', new_val).strip()
                                                new_val = re.sub(r'\{%.*?%\}', '', new_val).strip()
                                                
                                                matches = re.findall(r'\{\{\s*(?:row\.|mau\.)?(\w+)\s*\}\}', new_val)
                                                for m in matches:
                                                    replacement = str(item.get(m, ""))
                                                    new_val = re.sub(r'\{\{\s*(?:row\.|mau\.)?' + m + r'\s*\}\}', replacement, new_val)
                                                    
                                                if isinstance(new_val, str) and re.match(r'^-?\d+(?:,\d+)?$', new_val.strip()):
                                                    try:
                                                        new_val = float(new_val.strip().replace(',', '.'))
                                                    except: pass
                                                
                                                if new_val == "": new_val = None
                                                    
                                            new_cell = ws.cell(row=current_row, column=col_idx)
                                            new_cell.value = new_val
                                            
                                            if current_row >= template_row_idx:
                                                style = original_styles[col_idx - 1]
                                                new_cell.font = copy(style["font"])
                                                new_cell.border = copy(style["border"])
                                                new_cell.fill = copy(style["fill"])
                                                new_cell.number_format = copy(style["number_format"])
                                                new_cell.alignment = copy(style["alignment"])
                                                
                                        current_row += 1
                                                
                                    bio = io.BytesIO()
                                    wb.save(bio)
                                    bio.seek(0)
                                    
                                    st.success("🎉 Cập nhật số liệu thành công!")
                                    st.download_button("📥 Tải Xuống Biên Bản (.xlsx)", data=bio, file_name=f"Bien_Ban_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                                else:
                                    st.error("⚠️ Không tìm thấy vị trí chèn số liệu. Bạn hãy ghi đúng các thẻ (VD: `{{ row.ten_mau }}`) vào 1 dòng duy nhất trên Form Excel nhé.")
                except Exception as e:
                    st.error(f"❌ Có lỗi xảy ra trong quá trình xử lý: {e}")

    with tab_qr: 
        st.markdown("<div class='sub-title'>🏷️ Sinh Mã Vạch QR Tự Động</div>", unsafe_allow_html=True)
        st.write("Mô đun in tem dán mã vạch (Barcode/QR code) hàng loạt đang chờ tích hợp.")

elif menu == "🧪 Kiểm soát Hóa chất":
    st.markdown("<h1 class='main-title'>🧪 Quản lý Hóa chất & Vật tư tiêu hao</h1>", unsafe_allow_html=True)
    st.caption("Module kiểm soát chất chuẩn, dung môi và vật tư riêng biệt cho 3 hệ máy.")
    
    df_chem = st.session_state.df_chem.copy()
    
    today = datetime.now().date()
    df_chem['Hạn Sử Dụng'] = pd.to_datetime(df_chem['Hạn Sử Dụng'], errors='coerce').dt.date
    
    warnings = []
    for idx, row in df_chem.iterrows():
        exp_date = row['Hạn Sử Dụng']
        if pd.notna(exp_date):
            days_left = (exp_date - today).days
            if days_left < 0:
                warnings.append(f"❌ **ĐÃ HẾT HẠN:** {row['Tên Hóa Chất']} (Hệ: {row['Hệ Máy']}, Lô: {row['Số Lô (Lot)']}) - Hết hạn từ {exp_date.strftime('%d/%m/%Y')}.")
            elif days_left <= 30:
                warnings.append(f"⚠️ **SẮP HẾT HẠN:** {row['Tên Hóa Chất']} (Hệ: {row['Hệ Máy']}) - Còn lại {days_left} ngày (EXP: {exp_date.strftime('%d/%m/%Y')}).")
        
        if str(row['Tình Trạng Kho']) == "🔴 Đã hết":
            warnings.append(f"🛒 **HẾT HÀNG TRONG KHO:** {row['Tên Hóa Chất']} ({row['Hệ Máy']}). Cần lên kế hoạch mua sắm (PO) ngay!")

    if warnings:
        st.markdown("<div class='warning-box'><strong>🚨 DANH SÁCH CẢNH BÁO CẦN LƯU Ý:</strong><br>", unsafe_allow_html=True)
        for w in warnings:
            st.markdown(w)
        st.markdown("</div>", unsafe_allow_html=True)
    
    with st.container(border=True):
        tab_all, tab_gcms, tab_gcfid, tab_thermo = st.tabs(["Tất cả Hóa chất", "🔬 GC-MS", "🔥 GC-FID", "🧬 Thermo"])
        
        def render_chem_editor(filter_system=None):
            if filter_system:
                mask = df_chem["Hệ Máy"] == filter_system
                df_view = df_chem[mask].copy()
            else:
                df_view = df_chem.copy()
                
            edited_chem = st.data_editor(
                df_view,
                num_rows="dynamic",
                use_container_width=True,
                column_config={
                    "Hệ Máy": st.column_config.SelectboxColumn("Hệ Máy", options=CHEM_SYSTEMS, required=True),
                    "Phân Loại": st.column_config.SelectboxColumn("Phân Loại", options=CHEM_TYPES, required=True),
                    "Tên Hóa Chất": st.column_config.TextColumn("Tên Hóa Chất / Vật Tư", required=True),
                    "Ngày Mở Nắp": st.column_config.DateColumn("Ngày Mở Nắp", format="YYYY-MM-DD"),
                    "Hạn Sử Dụng": st.column_config.DateColumn("Hạn Sử Dụng (EXP)", format="YYYY-MM-DD"),
                    "Tình Trạng Kho": st.column_config.SelectboxColumn("Tình Trạng Kho", options=CHEM_STATUS)
                },
                key=f"chem_editor_{filter_system if filter_system else 'all'}",
                height=400
            )
            return edited_chem

        with tab_all: edited_all = render_chem_editor()
        with tab_gcms: edited_gcms = render_chem_editor("GC-MS")
        with tab_gcfid: edited_gcfid = render_chem_editor("GC-FID")
        with tab_thermo: edited_thermo = render_chem_editor("Thermo")
            
        st.caption("✨ **Mẹo:** Thêm, sửa, xóa các hóa chất trực tiếp trên bảng. Hệ thống sẽ tự động cập nhật cảnh báo khi bạn lưu lại.")

        if st.button("💾 Lưu Cập nhật Kho Hóa chất", type="primary"):
            st.session_state.df_chem = edited_all
            save_chemical_data(st.session_state.df_chem)
            st.success("🎉 Đã lưu danh mục Hóa chất & Vật tư thành công!")
            st.rerun()

elif menu == "⚙️ Cấu hình Hệ thống":
    st.markdown("<h1 class='main-title'>⚙️ Cấu hình & Quản trị Hệ thống</h1>", unsafe_allow_html=True)
    
    st.markdown("<div class='sub-title'>Quản lý Thư viện MDL & LOQ</div>", unsafe_allow_html=True)
    if st.session_state.get('results_stale'):
        st.warning("⚠️ Thư viện đã bị thay đổi! Vui lòng quay lại tab Vận hành máy đo và bấm 'Tính lại' để có kết quả mới nhất.")
        
    st.info("💡 Bạn có thể chỉnh sửa, thêm, xóa các mức giới hạn trực tiếp trên bảng. Nhớ ấn **Lưu thay đổi**.")
    
    with st.container(border=True):
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
        
        if st.button("💾 Lưu thay đổi Thư viện", type="primary"):
            try:
                st.cache_data.clear() 
                edited_limit = edited_limit[edited_limit["Tên Chất"].str.strip() != ""] 
                
                conn.update(spreadsheet=SHEET_URL, worksheet="CauHinh_MDL_LOQ", data=edited_limit)
                st.session_state.df_limit = edited_limit
                if 'results' in st.session_state:
                    st.session_state.results_stale = True
                st.success("🎉 Đã lưu thư viện lên Google Sheets thành công!")
                st.rerun()
            except Exception as e:
                st.error(f"⚠️ Lỗi kết nối Google Sheets: {e}")

    with st.container(border=True):
        st.markdown("<div class='sub-title'>Cập nhật Hàng loạt (Import Excel)</div>", unsafe_allow_html=True)
        limit_file = st.file_uploader("Kéo thả file Bảng giới hạn (.xlsx)", type=["xlsx"])
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
                    st.success(f"✔️ Đã nhận diện {len(df_limit_new)} chỉ tiêu mới.")
                    st.dataframe(df_limit_new, use_container_width=True)
                    
                    if st.button("🚀 Trộn và Lưu vào Thư viện chung", type="primary"):
                        try:
                            st.cache_data.clear() 
                            combined_df = pd.concat([st.session_state.df_limit, df_limit_new], ignore_index=True)
                            combined_df = combined_df.drop_duplicates(subset=['Nền Mẫu', 'Tên Chất'], keep='last').reset_index(drop=True)
                            
                            conn.update(spreadsheet=SHEET_URL, worksheet="CauHinh_MDL_LOQ", data=combined_df)
                            st.session_state.df_limit = combined_df
                            if 'results' in st.session_state:
                                st.session_state.results_stale = True
                            st.success(f"🎉 Ghi đè thành công! Tổng bộ nhớ: {len(combined_df)} chỉ tiêu.")
                            st.rerun()
                        except Exception as sheet_err:
                            st.error(f"⚠️ Lỗi kết nối: {sheet_err}")
                else: st.error("Không tìm thấy cấu trúc bảng hợp lệ (Cột Tên / Cột MDL / Cột LOQ).")
            except Exception as e: st.error(f"Lỗi đọc file: {e}")
