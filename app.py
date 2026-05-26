import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import json
import gspread
from google.oauth2.service_account import Credentials
import datetime

st.set_page_config(page_title="Gigaversal Sales Dashboard", layout="wide", page_icon="📊")

# ---------------------------------------------------------
# HÀM XỬ LÝ DỮ LIỆU CỐT LÕI (TÁI SỬ DỤNG CODE CŨ)
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
        'doanh thu chi tiết': 'Thành tiền' 
    }
    df.rename(columns=rename_dict, inplace=True)
    return df

def parse_money(val):
    val = str(val).strip()
    if val in ('', 'nan', '-'): return 0.0
    val = val.replace(',', '') 
    if val.endswith('.00'): val = val[:-3] 
    val = val.replace('.', '') 
    try: return float(val)
    except: return 0.0

@st.cache_data(ttl=3600)
def process_data(df_kiot, df_har, map_kiot, map_har, loai_bo_qua):
    df_kiot = normalize_cols(df_kiot)
    df_har = normalize_cols(df_har)
    map_kiot = normalize_cols(map_kiot)
    map_har = normalize_cols(map_har)

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
    df_merged['Khu'] = df_merged['Khu'].fillna('Chưa phân loại')
    df_merged['Loại'] = df_merged['Loại'].fillna('Chưa phân loại')

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
# GIAO DIỆN
# ---------------------------------------------------------
st.title("🚀 Gigaversal Sales Dashboard")
st.markdown("Dashboard báo cáo doanh thu động (thay thế Google Sheets)")

st.sidebar.header("⚙️ Nguồn dữ liệu")
data_mode = st.sidebar.radio("Chọn cách lấy dữ liệu", ["Upload file Excel", "Kết nối Google Sheets (cần File JSON)"])

loai_bo_qua_input = st.sidebar.text_input("Các loại cần bỏ qua (cách nhau bởi dấu phẩy)", "vé đoàn, phụ kiện")
loai_bo_qua = [x.strip().lower() for x in loai_bo_qua_input.split(',')]

df_report = pd.DataFrame()
missing_map = pd.DataFrame()

if data_mode == "Upload file Excel":
    st.info("Vui lòng tải file Excel của bạn lên (File Excel chứa các tab Kiot Viet, Haravan, Khu Kiot Viet, Khu Haravan)")
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

else:
    st.info("Để kết nối Google Sheets, bạn cần tải file Service Account (JSON) từ Google Cloud và dán link file Google Sheets.")
    gg_url = st.text_input("Link Google Sheets", "https://docs.google.com/spreadsheets/d/1JJ_Gmf35ZhUQBESgp3DOMtZMtrJBghu6pAS9zub37Y4/edit")
    json_file = st.file_uploader("Upload file credentials.json", type=["json"])
    
    if st.button("Tải dữ liệu từ Google Sheets"):
        creds = None
        
        # Thử lấy credentials từ Streamlit Secrets (Khi chạy online)
        try:
            if "gcp_service_account" in st.secrets:
                creds_dict = st.secrets["gcp_service_account"]
                creds = Credentials.from_service_account_info(creds_dict, scopes=['https://www.googleapis.com/auth/spreadsheets'])
        except Exception:
            pass
            
        # Nếu không có Secrets, lấy từ file JSON upload
        if creds is None and json_file is not None:
            creds_dict = json.load(json_file)
            creds = Credentials.from_service_account_info(creds_dict, scopes=['https://www.googleapis.com/auth/spreadsheets'])
            
        if creds is not None and gg_url:
            try:
                gc = gspread.authorize(creds)
                sh = gc.open_by_url(gg_url)
                
                def get_ws(name):
                    try: return pd.DataFrame(sh.worksheet(name).get_all_records())
                    except: return pd.DataFrame()
                
                with st.spinner("Đang tải dữ liệu từ Google Sheets..."):
                    df_kiot = get_ws('Kiot Viet')
                    df_har = get_ws('Haravan')
                    map_kiot = get_ws('Khu Kiot Viet')
                    map_har = get_ws('Khu Haravan')
                    
                    df_report, missing_map = process_data(df_kiot, df_har, map_kiot, map_har, loai_bo_qua)
                    st.success("Tải và xử lý dữ liệu thành công!")
            except Exception as e:
                st.error(f"Lỗi kết nối: {e}")
        else:
            st.warning("Vui lòng nhập Link và Upload file JSON.")

# ---------------------------------------------------------
# HIỂN THỊ DASHBOARD
# ---------------------------------------------------------
if not df_report.empty:
    st.divider()
    
    # Bộ lọc Filter
    st.sidebar.header("🔍 Bộ lọc Dashboard")
    min_date = df_report['Ngày báo cáo'].min()
    max_date = df_report['Ngày báo cáo'].max()
    
    date_range = st.sidebar.date_input("Chọn thời gian", [min_date, max_date], min_value=min_date, max_value=max_date)
    
    all_khu = sorted(list(df_report['Khu'].unique()))
    selected_khu = st.sidebar.multiselect("Chọn Khu", all_khu, default=all_khu)
    
    all_loai = sorted(list(df_report['Loại'].unique()))
    selected_loai = st.sidebar.multiselect("Chọn Loại", all_loai, default=all_loai)
    
    # Lọc data
    if len(date_range) == 2:
        df_final = df_report[
            (df_report['Ngày báo cáo'] >= date_range[0]) & 
            (df_report['Ngày báo cáo'] <= date_range[1]) &
            (df_report['Khu'].isin(selected_khu)) &
            (df_report['Loại'].isin(selected_loai))
        ]
    else:
        df_final = df_report.copy()
        
    # Thẻ Tổng quan
    col1, col2 = st.columns(2)
    tong_luot = df_final['Số lượt bán'].sum()
    tong_tien = df_final['Doanh thu'].sum()
    
    col1.metric("🎫 Tổng Số Lượt Bán", f"{tong_luot:,.0f}")
    col2.metric("💰 Tổng Doanh Thu (VNĐ)", f"{tong_tien:,.0f} đ")
    
    st.divider()
    st.subheader("📈 Phân tích từng khu vực (Mỗi khu vực 4 biểu đồ: Ngày, Tuần, Tháng, Năm)")
    
    # Nhóm theo Khu để vẽ biểu đồ
    for khu in selected_khu:
        df_khu = df_final[df_final['Khu'] == khu]
        if df_khu.empty: continue
        
        st.markdown(f"### 🎯 Khu: {khu}")
        
        # Nhóm theo Ngày
        df_ngay = df_khu.groupby('Ngày báo cáo').agg({'Số lượt bán':'sum', 'Doanh thu':'sum'}).reset_index()
        # Nhóm theo Tuần
        df_khu['Tuần-Năm'] = df_khu['Tuần (T2-CN)'].astype(str) + "/" + df_khu['Năm'].astype(str)
        df_tuan = df_khu.groupby(['Năm', 'Tuần (T2-CN)', 'Tuần-Năm']).agg({'Số lượt bán':'sum', 'Doanh thu':'sum'}).reset_index()
        df_tuan = df_tuan.sort_values(['Năm', 'Tuần (T2-CN)'])
        # Nhóm theo Tháng
        df_khu['Tháng-Năm'] = df_khu['Tháng'].astype(str) + "/" + df_khu['Năm'].astype(str)
        df_thang = df_khu.groupby(['Năm', 'Tháng', 'Tháng-Năm']).agg({'Số lượt bán':'sum', 'Doanh thu':'sum'}).reset_index()
        df_thang = df_thang.sort_values(['Năm', 'Tháng'])
        # Nhóm theo Năm
        df_nam = df_khu.groupby('Năm').agg({'Số lượt bán':'sum', 'Doanh thu':'sum'}).reset_index()
        
        # Vẽ biểu đồ
        tab1, tab2, tab3, tab4 = st.tabs(["📅 Theo Ngày", "📆 Theo Tuần", "🗓️ Theo Tháng", "🎉 Theo Năm"])
        
        def plot_dual_axis(df, x_col, title):
            fig = go.Figure()
            fig.add_trace(go.Bar(x=df[x_col], y=df['Số lượt bán'], name="Số lượt bán", yaxis="y1", marker_color="#1f77b4"))
            fig.add_trace(go.Scatter(x=df[x_col], y=df['Doanh thu'], name="Doanh thu", yaxis="y2", mode="lines+markers", line=dict(color="#ff7f0e", width=3)))
            fig.update_layout(
                title=title,
                xaxis=dict(title=x_col),
                yaxis=dict(title="Số lượt bán", side="left", showgrid=False),
                yaxis2=dict(title="Doanh thu (VNĐ)", side="right", overlaying="y", showgrid=True),
                legend=dict(x=0.01, y=1.1, orientation="h")
            )
            return fig
            
        with tab1:
            st.plotly_chart(plot_dual_axis(df_ngay, 'Ngày báo cáo', f"Doanh thu & Lượt bán theo Ngày - {khu}"), use_container_width=True)
        with tab2:
            st.plotly_chart(plot_dual_axis(df_tuan, 'Tuần-Năm', f"Doanh thu & Lượt bán theo Tuần - {khu}"), use_container_width=True)
        with tab3:
            st.plotly_chart(plot_dual_axis(df_thang, 'Tháng-Năm', f"Doanh thu & Lượt bán theo Tháng - {khu}"), use_container_width=True)
        with tab4:
            st.plotly_chart(plot_dual_axis(df_nam, 'Năm', f"Doanh thu & Lượt bán theo Năm - {khu}"), use_container_width=True)

    if not missing_map.empty:
        st.divider()
        st.warning(f"⚠️ Phát hiện {missing_map['Tên gốc'].nunique()} mã sản phẩm chưa được phân khu!")
        cols = ['Tên gốc']
        if 'SKU' in missing_map.columns: cols.insert(0, 'SKU')
        st.dataframe(missing_map[cols].drop_duplicates())
