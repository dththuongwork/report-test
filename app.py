import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import json
import gspread
from google.oauth2.service_account import Credentials
import datetime

st.set_page_config(page_title="Gigaversal Sales Dashboard", layout="wide", page_icon="📊")

# ---------------------------------------------------------
# HÀM XỬ LÝ DỮ LIỆU CỐT LÕI
# ---------------------------------------------------------
def normalize_cols(df):
    if df.empty: return df
    df = df.copy()
    df.columns = df.columns.astype(str).str.strip().str.lower()
    rename_dict = {
        'tên sản phẩm': 'Tên sản phẩm',
        'tên hàng': 'Tên sản phẩm',
        'sku': 'SKU',
        'mã hàng': 'SKU',
        'khu': 'Khu',
        'tên khu': 'Khu',
        'khu game': 'Khu',
        'loại': 'Loại',
        'phân loại': 'Loại',
        'ngày': 'Ngày',
        'thời gian': 'Ngày',
        'số lượng': 'Số lượng',
        'sl': 'Số lượng',
        'thành tiền': 'Thành tiền',
        'doanh thu chi tiết': 'Thành tiền',
        'trạng thái thanh toán': 'Trạng thái thanh toán',
        'trạng thái tt': 'Trạng thái thanh toán'
    }
    df.rename(columns=rename_dict, inplace=True)
    return df

def parse_money(val):
    if pd.isna(val): return 0.0
    if isinstance(val, (int, float)): return float(val)
    val = str(val).strip()
    if val in ('', 'nan', '-'): return 0.0
    
    # Xoá dấu phẩy ngăn cách hàng nghìn (ví dụ 1,500,000)
    val = val.replace(',', '')
    # Kiểm tra xử lý dấu chấm
    if val.count('.') > 1:
        # Nếu có nhiều dấu chấm (1.500.000), đây chắc chắn là phân cách hàng nghìn, xoá đi.
        val = val.replace('.', '')
    elif val.count('.') == 1:
        parts = val.split('.')
        if len(parts[1]) == 3: 
            # VD: 1500.000 -> Có thể là 1 triệu 5, xoá dấu chấm
            val = val.replace('.', '')
        # Ngược lại nếu len != 3, ví dụ 1500000.0 thì giữ nguyên dấu chấm thập phân
    try: return float(val)
    except: return 0.0

@st.cache_data(ttl=3600)
def load_data_from_google_drive(url):
    try:
        # Trích xuất ID từ link chia sẻ
        file_id = url.split('/d/')[1].split('/')[0]
        download_url = f'https://drive.google.com/uc?id={file_id}&export=download'
        
        # Đọc trực tiếp bằng Pandas không cần thư viện phụ trợ
        xls = pd.ExcelFile(download_url)
        sheets = xls.sheet_names
        
        df_kiot = pd.read_excel(xls, 'Kiot Viet') if 'Kiot Viet' in sheets else pd.DataFrame()
        df_har = pd.read_excel(xls, 'Haravan') if 'Haravan' in sheets else pd.DataFrame()
        map_kiot = pd.read_excel(xls, 'Khu Kiot Viet') if 'Khu Kiot Viet' in sheets else pd.DataFrame()
        map_har = pd.read_excel(xls, 'Khu Haravan') if 'Khu Haravan' in sheets else pd.DataFrame()
        
        return df_kiot, df_har, map_kiot, map_har
    except Exception as e:
        st.error(f"Lỗi khi tải từ Google Drive: {e}")
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

@st.cache_data(ttl=3600)
def process_data(df_kiot, df_har, map_kiot, map_har, loai_bo_qua):
    df_kiot = normalize_cols(df_kiot)
    df_har = normalize_cols(df_har)
    map_kiot = normalize_cols(map_kiot)
    map_har = normalize_cols(map_har)

    # 1. CHỈ LẤY ĐƠN ĐÃ THANH TOÁN CỦA HARAVAN
    if 'Trạng thái thanh toán' in df_har.columns:
        df_har = df_har[df_har['Trạng thái thanh toán'].astype(str).str.strip().str.lower() == 'paid']

    if 'Ngày' in df_kiot.columns: df_kiot['Ngày'] = pd.to_datetime(df_kiot['Ngày'], dayfirst=True, errors='coerce')
    if 'Ngày' in df_har.columns: df_har['Ngày'] = pd.to_datetime(df_har['Ngày'], errors='coerce')

    df_raw = pd.concat([df_kiot, df_har], ignore_index=True)
    df_map = pd.concat([map_kiot, map_har], ignore_index=True)

    if df_raw.empty or 'Tên sản phẩm' not in df_raw.columns:
        return pd.DataFrame(), pd.DataFrame()

    df_raw['Tên gốc'] = df_raw['Tên sản phẩm'] 
    df_raw['Tên chuẩn'] = df_raw['Tên sản phẩm'].astype(str).str.strip().str.lower()
    
    df_raw = df_raw[df_raw['Tên chuẩn'] != '']
    df_raw = df_raw[df_raw['Tên chuẩn'] != 'nan']
    df_raw = df_raw[~df_raw['Tên chuẩn'].str.match(r'^\d{1,2}:\d{2}(:\d{2})?$')]
    if 'SKU' in df_raw.columns:
        df_raw = df_raw[~df_raw['SKU'].astype(str).str.strip().str.match(r'^\d{4}-\d{2}-\d{2}$')]

    if 'Số lượng' not in df_raw.columns: df_raw['Số lượng'] = 0
    if 'Thành tiền' not in df_raw.columns: df_raw['Thành tiền'] = 0
    df_raw['Số lượng'] = pd.to_numeric(df_raw['Số lượng'], errors='coerce').fillna(0)
    df_raw['Thành tiền'] = df_raw['Thành tiền'].apply(parse_money)

    for col in ['Tên sản phẩm', 'Khu', 'Loại']:
        if col not in df_map.columns: df_map[col] = 'Chưa phân loại'

    df_map['Tên chuẩn'] = df_map['Tên sản phẩm'].astype(str).str.strip().str.lower()
    df_map = df_map[df_map['Tên chuẩn'] != '']
    df_map = df_map.drop_duplicates(subset=['Tên chuẩn'])

    df_merged = df_raw.merge(df_map[['Tên chuẩn', 'Khu', 'Loại']], on='Tên chuẩn', how='left')
    df_merged['Khu'] = df_merged['Khu'].fillna('Chưa phân loại').replace('', 'Chưa phân loại')
    df_merged['Loại'] = df_merged['Loại'].fillna('Chưa phân loại').replace('', 'Chưa phân loại')

    missing_map = df_merged[df_merged['Khu'] == 'Chưa phân loại'].copy()

    if 'Ngày' in df_merged.columns:
        df_merged = df_merged.dropna(subset=['Ngày'])
        df_merged['Loại_lower'] = df_merged['Loại'].astype(str).str.lower()
        df_filtered = df_merged[~df_merged['Loại_lower'].isin(loai_bo_qua)].copy()
        
        df_filtered['Ngày báo cáo'] = df_filtered['Ngày'].dt.date
        df_filtered['Năm'] = df_filtered['Ngày'].dt.isocalendar().year
        df_filtered['Tháng'] = df_filtered['Ngày'].dt.month
        df_filtered['Tuần (T2-CN)'] = df_filtered['Ngày'].dt.isocalendar().week
        
        df_report = df_filtered.groupby(
            ['Năm', 'Tháng', 'Tuần (T2-CN)', 'Ngày báo cáo', 'Khu', 'Loại'], as_index=False
        ).agg({'Số lượng': 'sum', 'Thành tiền': 'sum'})
        df_report.rename(columns={'Số lượng': 'Số lượt bán', 'Thành tiền': 'Doanh thu'}, inplace=True)
        return df_report, missing_map
    return pd.DataFrame(), missing_map

# ---------------------------------------------------------
# PHÂN LOẠI KHU VỰC VÀO TABS
# ---------------------------------------------------------
khu_gioi_tre = ["Xu", "Bắn cung", "Súng", "VR Game", "Chụp hình", "Bowling"]
khu_thieu_nhi = ["Alpha Games", "Rainbow Fun", "Jungle", "Xe Tuần Lộc", "Led Sàn cá", "Light City", "Combo RB + JU", "Combo Light City + RB + JU", "Combo Tuần Lộc + Sàn Cá", "Tô Tượng"]
khu_12d = ["The Lost World", "Infinity World", "Fly Over The World", "Universal Guardians", "Combo Light City + 12D+", "Combo Universal Guardians + 12D+", "Combo 2 trong 3 Game 12D+", "Combo 5 khu công nghệ"]
khu_vangogh = ["Vé Van Gogh", "Souvenir Van Gogh"]

def get_khu_group(khu_name):
    khu_name_lower = str(khu_name).lower()
    for item in khu_gioi_tre:
        if item.lower() in khu_name_lower: return "Khu Giới trẻ"
    for item in khu_thieu_nhi:
        if item.lower() in khu_name_lower: return "Khu thiếu nhi"
    for item in khu_12d:
        if item.lower() in khu_name_lower: return "Khu 12D"
    for item in khu_vangogh:
        if item.lower() in khu_name_lower: return "Khu VanGogh"
    return "Khu Khác"

# ---------------------------------------------------------
# GIAO DIỆN
# ---------------------------------------------------------
st.title("🚀 Gigaversal Sales Dashboard")

st.sidebar.header("⚙️ Nguồn dữ liệu")
data_mode = st.sidebar.radio("Chọn cách lấy dữ liệu", ["Tự động tải từ Google Drive (Public)", "Upload file Excel"])

loai_bo_qua_input = st.sidebar.text_input("Các loại vé/nhóm cần bỏ qua (cách nhau bởi dấu phẩy)", "vé đoàn, phụ kiện")
loai_bo_qua = [x.strip().lower() for x in loai_bo_qua_input.split(',')]

df_report = pd.DataFrame()
missing_map = pd.DataFrame()

if data_mode == "Tự động tải từ Google Drive (Public)":
    st.info("Hệ thống sẽ tự động đọc file Excel trực tiếp từ link Google Drive của bạn (Miễn là link được thiết lập Bất kỳ ai có liên kết đều xem được).")
    gg_drive_url = st.text_input("Link Google Drive chứa file Excel", "https://docs.google.com/spreadsheets/d/1v0j1mf6KbLf7ws1klEpBOi-6hv_xfv_9/edit?usp=sharing")
    
    if st.button("Tải dữ liệu mới nhất"):
        with st.spinner("Đang tải dữ liệu từ Google Drive..."):
            df_kiot, df_har, map_kiot, map_har = load_data_from_google_drive(gg_drive_url)
            df_report, missing_map = process_data(df_kiot, df_har, map_kiot, map_har, loai_bo_qua)
            if not df_report.empty:
                st.success("Tải và xử lý dữ liệu thành công!")

else:
    st.info("Vui lòng tải file Excel của bạn lên")
    uploaded_file = st.file_uploader("Chọn file Excel (.xlsx)", type=["xlsx"])
    if uploaded_file is not None:
        try:
            xls = pd.ExcelFile(uploaded_file)
            sheets = xls.sheet_names
            
            df_kiot = pd.read_excel(xls, 'Kiot Viet') if 'Kiot Viet' in sheets else pd.DataFrame()
            df_har = pd.read_excel(xls, 'Haravan') if 'Haravan' in sheets else pd.DataFrame()
            map_kiot = pd.read_excel(xls, 'Khu Kiot Viet') if 'Khu Kiot Viet' in sheets else pd.DataFrame()
            map_har = pd.read_excel(xls, 'Khu Haravan') if 'Khu Haravan' in sheets else pd.DataFrame()

            df_report, missing_map = process_data(df_kiot, df_har, map_kiot, map_har, loai_bo_qua)
            st.success("Tải và xử lý dữ liệu thành công!")
        except Exception as e:
            st.error(f"Lỗi khi đọc file: {e}")

# ---------------------------------------------------------
# HIỂN THỊ DASHBOARD
# ---------------------------------------------------------
if not df_report.empty:
    st.divider()
    
    st.sidebar.header("🔍 Bộ lọc Dashboard")
    min_date = df_report['Ngày báo cáo'].min()
    max_date = df_report['Ngày báo cáo'].max()
    
    date_range = st.sidebar.date_input("Chọn thời gian", [min_date, max_date], min_value=min_date, max_value=max_date)
    
    all_loai = sorted(list(df_report['Loại'].unique()))
    selected_loai = st.sidebar.multiselect("Chọn Loại", all_loai, default=all_loai)
    
    if len(date_range) == 2:
        df_final = df_report[
            (df_report['Ngày báo cáo'] >= date_range[0]) & 
            (df_report['Ngày báo cáo'] <= date_range[1]) &
            (df_report['Loại'].isin(selected_loai))
        ]
    else:
        df_final = df_report.copy()
        
    col1, col2 = st.columns(2)
    tong_luot = df_final['Số lượt bán'].sum()
    tong_tien = df_final['Doanh thu'].sum()
    
    col1.metric("🎫 Tổng Số Lượt Bán Toàn Khu", f"{tong_luot:,.0f}")
    col2.metric("💰 Tổng Doanh Thu (VNĐ)", f"{tong_tien:,.0f} đ")
    
    st.divider()
    
    # Phân nhóm dữ liệu vào các Tab
    groups = ["Khu Giới trẻ", "Khu thiếu nhi", "Khu 12D", "Khu VanGogh", "Khu Khác"]
    tabs = st.tabs(groups)
    
    def plot_dual_axis(df, x_col, title):
        fig = go.Figure()
        
        # Hàm format hiển thị số (ẩn đi nếu số bằng 0 để bớt rối)
        fmt_func = lambda x: f"{x:,.0f}" if x > 0 else ""
        
        fig.add_trace(go.Bar(
            x=df[x_col], y=df['Số lượt bán'], name="Số lượt bán", yaxis="y1", 
            marker_color="#1f77b4",
            text=df['Số lượt bán'].apply(fmt_func), textposition='auto'
        ))
        
        fig.add_trace(go.Scatter(
            x=df[x_col], y=df['Doanh thu'], name="Doanh thu", yaxis="y2", 
            mode="lines+markers+text", 
            line=dict(color="#ff7f0e", width=3),
            text=df['Doanh thu'].apply(fmt_func), textposition='top center', textfont=dict(color="#ff7f0e", size=11)
        ))
        
        fig.update_layout(
            title=title,
            xaxis=dict(title=x_col),
            yaxis=dict(title="Số lượt bán", side="left", showgrid=False),
            yaxis2=dict(title="Doanh thu (VNĐ)", side="right", overlaying="y", showgrid=True),
            legend=dict(x=0.01, y=1.1, orientation="h"),
            margin=dict(t=60)
        )
        return fig

    # Xử lý hiển thị từng tab
    for i, group_name in enumerate(groups):
        with tabs[i]:
            # Lọc các dòng thuộc group này
            df_group = df_final[df_final['Khu'].apply(get_khu_group) == group_name]
            
            if df_group.empty:
                st.info(f"Chưa có dữ liệu cho {group_name} trong thời gian này.")
                continue
                
            khu_list = sorted(list(df_group['Khu'].unique()))
            selected_khu = st.multiselect(f"Chọn khu vực chi tiết trong {group_name}:", khu_list, default=khu_list, key=f"ms_{i}")
            
            for khu in selected_khu:
                df_khu = df_group[df_group['Khu'] == khu]
                if df_khu.empty: continue
                
                st.markdown(f"### 🎯 {khu}")
                
                # Điền đầy đủ các ngày trống bằng 0
                df_ngay = df_khu.groupby('Ngày báo cáo').agg({'Số lượt bán':'sum', 'Doanh thu':'sum'}).reset_index()
                if not df_ngay.empty:
                    df_ngay['Ngày báo cáo'] = pd.to_datetime(df_ngay['Ngày báo cáo'])
                    idx = pd.date_range(pd.to_datetime(date_range[0]), pd.to_datetime(date_range[1]))
                    df_ngay = df_ngay.set_index('Ngày báo cáo').reindex(idx, fill_value=0).reset_index()
                    df_ngay.rename(columns={'index': 'Ngày báo cáo'}, inplace=True)
                    df_ngay['Ngày báo cáo'] = df_ngay['Ngày báo cáo'].dt.strftime('%d/%m')
                
                # Tuần
                df_khu_tuan = df_khu.copy()
                df_khu_tuan['Tuần-Năm'] = df_khu_tuan['Tuần (T2-CN)'].astype(str) + "/" + df_khu_tuan['Năm'].astype(str)
                df_tuan = df_khu_tuan.groupby(['Năm', 'Tuần (T2-CN)', 'Tuần-Năm']).agg({'Số lượt bán':'sum', 'Doanh thu':'sum'}).reset_index()
                df_tuan = df_tuan.sort_values(['Năm', 'Tuần (T2-CN)'])
                
                # Tháng
                df_khu_thang = df_khu.copy()
                df_khu_thang['Tháng-Năm'] = df_khu_thang['Tháng'].astype(str) + "/" + df_khu_thang['Năm'].astype(str)
                df_thang = df_khu_thang.groupby(['Năm', 'Tháng', 'Tháng-Năm']).agg({'Số lượt bán':'sum', 'Doanh thu':'sum'}).reset_index()
                df_thang = df_thang.sort_values(['Năm', 'Tháng'])
                
                # Năm
                df_nam = df_khu.groupby('Năm').agg({'Số lượt bán':'sum', 'Doanh thu':'sum'}).reset_index()
                
                c1, c2, c3, c4 = st.tabs(["📅 Ngày", "📆 Tuần", "🗓️ Tháng", "🎉 Năm"])
                
                with c1: st.plotly_chart(plot_dual_axis(df_ngay, 'Ngày báo cáo', f"Doanh thu & Lượt bán Ngày - {khu}"), use_container_width=True)
                with c2: st.plotly_chart(plot_dual_axis(df_tuan, 'Tuần-Năm', f"Doanh thu & Lượt bán Tuần - {khu}"), use_container_width=True)
                with c3: st.plotly_chart(plot_dual_axis(df_thang, 'Tháng-Năm', f"Doanh thu & Lượt bán Tháng - {khu}"), use_container_width=True)
                with c4: st.plotly_chart(plot_dual_axis(df_nam, 'Năm', f"Doanh thu & Lượt bán Năm - {khu}"), use_container_width=True)

    if not missing_map.empty:
        st.divider()
        st.warning(f"⚠️ Phát hiện {missing_map['Tên gốc'].nunique()} mã sản phẩm chưa được phân khu!")
        cols = ['Tên gốc']
        if 'SKU' in missing_map.columns: cols.insert(0, 'SKU')
        st.dataframe(missing_map[cols].drop_duplicates())
