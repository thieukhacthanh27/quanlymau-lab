import streamlit as st
import pandas as pd
from datetime import datetime
from streamlit_gsheets import GSheetsConnection

# ==========================================
# 1. CẤU HÌNH GIAO DIỆN & BRANDING
# ==========================================
# Đổi biểu tượng tab và tiêu đề phần mềm
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
import React, { useState, useEffect } from 'react';
import { View, Text, StyleSheet, TouchableOpacity, Alert } from 'react-native';
import { registerBarcodeScanner } from '../services/DataWedgeService';
import database from '../database';
import { SAMPLE_STATUS } from '../database/models/Sample';

export default function Step1_PendingStorage({ currentUser }) {
  const [sampleBarcode, setSampleBarcode] = useState(null);
  const [locationBarcode, setLocationBarcode] = useState(null);
  const [isSaving, setIsSaving] = useState(false);

  // Lắng nghe sự kiện từ máy quét mã vạch công nghiệp Zebra
  useEffect(() => {
    const unsubscribe = registerBarcodeScanner((data, type) => {
      // Phân loại mã vạch dựa trên tiền tố (Prefix) được quy định trong SOP
      if (data.startsWith('LOC-')) {
        setLocationBarcode(data); // Ví dụ: LOC-FRZ1-SH2-BX4-A1
      } else {
        setSampleBarcode(data);   // Mã 2D DataMatrix trên lọ vial 2mL
      }
    });
    return () => unsubscribe();
  }, []);

  // Ghi dữ liệu vào WatermelonDB (Ưu tiên ngoại tuyến)
  const handleSaveSample = async () => {
    if (!sampleBarcode || !locationBarcode) {
      Alert.alert('Lỗi', 'Vui lòng quét đủ Mã mẫu và Vị trí lưu trữ.');
      return;
    }

    setIsSaving(true);
    try {
      await database.write(async () => {
        const samplesCollection = database.collections.get('samples');
        
        // Tạo bản ghi mẫu mới ở Bước 1
        await samplesCollection.create(sample => {
          sample.barcodeId = sampleBarcode;
          sample.status = SAMPLE_STATUS.STEP_1_PENDING;
          sample.locationId = locationBarcode;
          sample.isFlaggedOos = false;
        });

        // Ghi Audit Trail tự động cho thao tác tiếp nhận
        const logsCollection = database.collections.get('audit_logs');
        await logsCollection.create(log => {
          log.record_id = sampleBarcode;
          log.action = 'ACCESSIONING_STEP_1';
          log.user_id = currentUser.id;
          log.timestamp = Date.now();
        });
      });

      Alert.alert('Thành công', 'Đã lưu mẫu vào CSDL cục bộ!');
      setSampleBarcode(null);
      setLocationBarcode(null);
    } catch (error) {
      Alert.alert('Lỗi lưu trữ', error.message);
    } finally {
      setIsSaving(false);
    }
  };

  return (
    <View style={styles.container}>
      <Text style={styles.title}>Bước 1: Tiếp Nhận & Lưu Kho</Text>
      
      <View style={styles.scanCard}>
        <Text style={styles.label}>Mã lọ mẫu (DataMatrix 2D):</Text>
        <Text style={styles.value}>{sampleBarcode || 'Đang chờ quét...'}</Text>
      </View>

      <View style={styles.scanCard}>
        <Text style={styles.label}>Tọa độ lưu trữ (Tủ/Kệ/Hộp):</Text>
        <Text style={styles.value}>{locationBarcode || 'Đang chờ quét...'}</Text>
      </View>

      <TouchableOpacity 
        style={[styles.button, (!sampleBarcode || !locationBarcode) && styles.disabled]} 
        onPress={handleSaveSample}
        disabled={!sampleBarcode || !locationBarcode || isSaving}
      >
        <Text style={styles.buttonText}>{isSaving ? 'Đang lưu...' : 'Xác Nhận Lưu'}</Text>
      </TouchableOpacity>
    </View>
  );
}

const styles = StyleSheet.create({
  container: { flex: 1, padding: 20, backgroundColor: '#f5f7fa' },
  title: { fontSize: 22, fontWeight: 'bold', marginBottom: 20, color: '#2c3e50' },
  scanCard: { backgroundColor: '#fff', padding: 15, borderRadius: 8, marginBottom: 15, elevation: 2 },
  label: { fontSize: 14, color: '#7f8c8d', marginBottom: 5 },
  value: { fontSize: 16, fontWeight: '600', color: '#2980b9' },
  button: { backgroundColor: '#27ae60', padding: 15, borderRadius: 8, alignItems: 'center' },
  disabled: { backgroundColor: '#95a5a6' },
  buttonText: { color: '#fff', fontWeight: 'bold', fontSize: 16 }
});
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
            seq_df['Method'] = df_ready['Chỉ Tiêu'].apply(lambda x: 'VOCs.M' if 'VOC' in str(x).upper() else 'HCHO.M')
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
