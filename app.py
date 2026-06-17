import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import unicodedata
import os
import json
from datetime import timedelta
import base64

st.set_page_config(page_title="Gigaversal Sales Dashboard", layout="wide", page_icon="📊")

# ---------------------------------------------------------
# HỆ THỐNG LƯU TRỮ LOCAL & GITHUB
# ---------------------------------------------------------
MAPPING_FILE = "mapping.json"
NHOM_FILE = "nhom_mapping.json"

NHOM_MAC_DINH = {
    "Khu Giới trẻ": ["Xu", "Bắn cung", "Bắn súng", "Súng", "VR Game", "Chụp hình", "Bowling"],
    "Khu thiếu nhi": ["Alpha Games", "Rainbow Fun", "Jungle", "Xe Tuần Lộc", "Led Sàn cá", "Light City", "Combo RB + JU", "Combo Light City + RB + JU", "Combo Tuần Lộc + Sàn Cá", "Tô Tượng"],
    "Khu 12D": ["The Lost World", "Infinity World", "Fly Over The World", "Universal Guardians", "Combo Light City + 12D+", "Combo Universal Guardians + 12D+", "Combo 2 trong 3 Game 12D+", "Combo 5 khu công nghệ"],
    "Khu VanGogh": ["Vé Van Gogh", "Souvenir Van Gogh"]
}

def load_json(filepath, default):
    if os.path.exists(filepath):
        try:
            with open(filepath, 'r', encoding='utf-8') as f:
                return json.load(f)
        except: pass
    return default

def save_json(filepath, data):
    with open(filepath, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

def is_github_connected():
    return "github_token" in st.secrets and "github_repo" in st.secrets

def push_to_github_if_configured(filename):
    if not is_github_connected(): return False
    try:
        token = st.secrets["github_token"]
        repo = st.secrets["github_repo"]
        import requests
        
        if not os.path.exists(filename): return False
        with open(filename, 'rb') as f: content = f.read()
        
        url = f"https://api.github.com/repos/{repo}/contents/{filename}"
        headers = {"Authorization": f"token {token}"}
        
        r_get = requests.get(url, headers=headers)
        sha = r_get.json().get('sha', None) if r_get.status_code == 200 else None
        
        data = {
            "message": f"Auto-save {filename} from Dashboard UI",
            "content": base64.b64encode(content).decode("utf-8")
        }
        if sha: data["sha"] = sha
        res = requests.put(url, headers=headers, json=data)
        return res.status_code in [200, 201]
    except:
        return False

# ---------------------------------------------------------
# HÀM XỬ LÝ DỮ LIỆU CỐT LÕI
# ---------------------------------------------------------
def clean_text(text):
    if pd.isna(text): return ""
    text = str(text).strip().lower()
    text = unicodedata.normalize('NFC', text)
    return " ".join(text.split())

def normalize_cols(df):
    if df.empty: return df
    df = df.copy()
    df.columns = [clean_text(c) for c in df.columns]
    rename_dict = {
        'tên sản phẩm': 'Tên sản phẩm',
        'tên hàng': 'Tên sản phẩm',
        'sku': 'SKU',
        'mã hàng': 'SKU',
        'khu': 'Khu',
        'tên khu': 'Khu',
        'khu game': 'Khu',
        'nhóm': 'Nhóm',
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
    val = val.replace(',', '')
    if val.count('.') > 1: val = val.replace('.', '')
    elif val.count('.') == 1:
        parts = val.split('.')
        if len(parts[1]) == 3: val = val.replace('.', '')
    try: return float(val)
    except: return 0.0

@st.cache_data(ttl=86400*30) 
def load_data_from_google_drive(url):
    try:
        file_id = url.split('/d/')[1].split('/')[0]
        download_url = f'https://drive.google.com/uc?id={file_id}&export=download'
        
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

def get_week_label(date_obj):
    start = date_obj - timedelta(days=date_obj.weekday())
    end = start + timedelta(days=6)
    return f"{start.strftime('%d/%m')} - {end.strftime('%d/%m')}"

@st.cache_data(ttl=86400*30)
def process_data(df_kiot, df_har, map_kiot, map_har, loai_bo_qua, haravan_status):
    df_kiot = normalize_cols(df_kiot)
    df_har = normalize_cols(df_har)
    map_kiot = normalize_cols(map_kiot)
    map_har = normalize_cols(map_har)

    # ĐÁNH DẤU NGUỒN DỮ LIỆU
    if not df_kiot.empty: df_kiot['Nguồn'] = 'Kiot Viet'
    if not df_har.empty: df_har['Nguồn'] = 'Haravan'

    # BỘ LỌC HARAVAN TÙY CHỈNH
    if haravan_status and 'Trạng thái thanh toán' in df_har.columns:
        keywords = [k.strip().lower() for k in haravan_status.split(',')]
        mask = df_har['Trạng thái thanh toán'].astype(str).str.lower().apply(lambda x: any(k in x for k in keywords))
        df_har = df_har[mask]

    if 'Ngày' in df_kiot.columns: df_kiot['Ngày'] = pd.to_datetime(df_kiot['Ngày'], dayfirst=True, errors='coerce')
    if 'Ngày' in df_har.columns: df_har['Ngày'] = pd.to_datetime(df_har['Ngày'], errors='coerce')

    df_raw = pd.concat([df_kiot, df_har], ignore_index=True)
    df_map = pd.concat([map_kiot, map_har], ignore_index=True)

    if df_raw.empty or 'Tên sản phẩm' not in df_raw.columns:
        return pd.DataFrame(), pd.DataFrame()

    df_raw['Tên gốc'] = df_raw['Tên sản phẩm'] 
    df_raw['Tên chuẩn'] = df_raw['Tên sản phẩm'].apply(clean_text)
    
    df_raw = df_raw[df_raw['Tên chuẩn'] != '']
    df_raw = df_raw[~df_raw['Tên chuẩn'].str.match(r'^\d{1,2}:\d{2}(:\d{2})?$')]
    if 'SKU' not in df_raw.columns: df_raw['SKU'] = ""
    else: df_raw = df_raw[~df_raw['SKU'].astype(str).str.strip().str.match(r'^\d{4}-\d{2}-\d{2}$')]

    if 'Số lượng' not in df_raw.columns: df_raw['Số lượng'] = 0
    if 'Thành tiền' not in df_raw.columns: df_raw['Thành tiền'] = 0
    df_raw['Số lượng'] = pd.to_numeric(df_raw['Số lượng'], errors='coerce').fillna(0)
    df_raw['Thành tiền'] = df_raw['Thành tiền'].apply(parse_money)

    # KẾT HỢP DỮ LIỆU TỪ UI MAPPING (JSON)
    local_map_dict = load_json(MAPPING_FILE, {})
    if local_map_dict:
        local_df = pd.DataFrame.from_dict(local_map_dict, orient='index').reset_index()
        local_df.rename(columns={'index': 'Tên chuẩn'}, inplace=True)
        df_map = pd.concat([df_map, local_df], ignore_index=True)

    cols_to_fill = ['Tên sản phẩm', 'Khu', 'Loại']
    for col in cols_to_fill:
        if col not in df_map.columns: df_map[col] = 'Chưa phân loại'

    df_map['Tên chuẩn'] = df_map['Tên sản phẩm'].apply(clean_text)
    df_map = df_map[df_map['Tên chuẩn'] != '']
    df_map = df_map.drop_duplicates(subset=['Tên chuẩn'], keep='last')

    df_merged = df_raw.merge(df_map[['Tên chuẩn', 'Khu', 'Loại']], on='Tên chuẩn', how='left')
    df_merged['Khu'] = df_merged['Khu'].fillna('Chưa phân loại').replace('', 'Chưa phân loại')
    df_merged['Loại'] = df_merged['Loại'].fillna('Chưa phân loại').replace('', 'Chưa phân loại')
    
    df_mapping_status = df_merged[['Nguồn', 'SKU', 'Tên gốc', 'Khu', 'Loại']].drop_duplicates(subset=['Tên gốc']).copy()

    if 'Ngày' in df_merged.columns:
        df_merged = df_merged.dropna(subset=['Ngày'])
        df_merged['Loại_lower'] = df_merged['Loại'].apply(clean_text)
        df_filtered = df_merged[~df_merged['Loại_lower'].isin(loai_bo_qua)].copy()
        
        df_filtered['Ngày báo cáo'] = df_filtered['Ngày'].dt.date
        df_filtered['Năm'] = df_filtered['Ngày'].dt.isocalendar().year
        df_filtered['Tháng'] = df_filtered['Ngày'].dt.month
        
        df_filtered['Tuần Label'] = df_filtered['Ngày'].dt.date.apply(get_week_label)
        df_filtered['Tuần ID'] = df_filtered['Ngày'].dt.isocalendar().week
        
        group_cols = ['Năm', 'Tháng', 'Tuần ID', 'Tuần Label', 'Ngày báo cáo', 'Khu', 'Loại']
        
        df_report = df_filtered.groupby(group_cols, as_index=False).agg({'Số lượng': 'sum', 'Thành tiền': 'sum'})
        df_report.rename(columns={'Số lượng': 'Số lượt bán', 'Thành tiền': 'Doanh thu'}, inplace=True)
        return df_report, df_mapping_status
    return pd.DataFrame(), df_mapping_status

def assign_group(row):
    nhom_dict = load_json(NHOM_FILE, NHOM_MAC_DINH)
    khu_name = clean_text(row['Khu'])
    
    # Ưu tiên 1: Tên Khu trùng chính xác với tên Nhóm (Ví dụ Khu = "Khu 12D")
    for group_name in nhom_dict.keys():
        if clean_text(group_name) == khu_name:
            return group_name
            
    # Ưu tiên 2: Tìm chuỗi con (Ví dụ Khu = "Vé 12D" chứa "12D")
    for group_name, khu_list in nhom_dict.items():
        for k in khu_list:
            if clean_text(k) in khu_name:
                return group_name
    return "Khu Khác"

# ---------------------------------------------------------
# GIAO DIỆN CHÍNH
# ---------------------------------------------------------
st.sidebar.header("⚙️ Nguồn dữ liệu")
gg_drive_url = st.sidebar.text_input("Link Google Drive chứa file Excel", "https://docs.google.com/spreadsheets/d/1v0j1mf6KbLf7ws1klEpBOi-6hv_xfv_9/edit?usp=sharing")

haravan_status = st.sidebar.text_input("Trạng thái thanh toán Haravan hợp lệ:", "paid")
loai_bo_qua_input = st.sidebar.text_input("Các loại vé/nhóm cần bỏ qua (cách nhau bởi dấu phẩy)", "vé đoàn, phụ kiện")
loai_bo_qua = [clean_text(x) for x in loai_bo_qua_input.split(',')]

if st.sidebar.button("🔄 LÀM MỚI DỮ LIỆU TỪ EXCEL"):
    st.cache_data.clear()
    st.rerun()

df_report = pd.DataFrame()
df_mapping_status = pd.DataFrame()

with st.spinner("Đang tải dữ liệu..."):
    df_kiot, df_har, map_kiot, map_har = load_data_from_google_drive(gg_drive_url)
    df_report, df_mapping_status = process_data(df_kiot, df_har, map_kiot, map_har, loai_bo_qua, haravan_status)

# ---------------------------------------------------------
# TABS GIAO DIỆN
# ---------------------------------------------------------
tab_dashboard, tab_mapping = st.tabs(["📊 DASHBOARD", "⚙️ PHÒNG ĐIỀU KHIỂN MAPPING"])

with tab_mapping:
    # --- PHẦN 1: QUẢN LÝ NHÓM ---
    st.markdown("## 1. Bản Đồ Nhóm Khu")
    st.info("Chỉnh sửa nhóm tuỳ ý. Bạn có thể bôi đen copy/paste như Google Sheets.")
    current_nhom = load_json(NHOM_FILE, NHOM_MAC_DINH)
    
    nhom_list = []
    for nhom, khus in current_nhom.items():
        for k in khus:
            nhom_list.append({"Khu": k, "Thuộc Nhóm": nhom})
    df_nhom = pd.DataFrame(nhom_list)
    
    edited_nhom = st.data_editor(df_nhom, num_rows="dynamic", key="nhom_editor", use_container_width=True)
    if st.button("💾 Lưu thay đổi Nhóm"):
        new_nhom_dict = {}
        for _, row in edited_nhom.iterrows():
            k = str(row['Khu']).strip()
            nh = str(row['Thuộc Nhóm']).strip()
            if k and nh and k != 'nan' and nh != 'nan':
                if nh not in new_nhom_dict: new_nhom_dict[nh] = []
                new_nhom_dict[nh].append(k)
        save_json(NHOM_FILE, new_nhom_dict)
        push_to_github_if_configured(NHOM_FILE)
        st.cache_data.clear()
        st.success("Đã lưu Nhóm vĩnh viễn!")
        st.rerun()
        
    if not is_github_connected() and os.path.exists(NHOM_FILE):
        with open(NHOM_FILE, 'r', encoding='utf-8') as f: st.download_button("📥 Tải file Backup Nhóm (Thủ công)", f, file_name=NHOM_FILE)
    
    st.divider()
    
    # --- PHẦN 2 & 3: PHÂN LOẠI SẢN PHẨM ---
    khu_options = ["Chưa phân loại"] + list(current_nhom.keys()) + [k for khus in current_nhom.values() for k in khus]
    khu_options = list(dict.fromkeys(khu_options)) # Xoá trùng lặp
    loai_options = ["Chưa phân loại", "Vé", "Combo", "Dịch vụ", "Khác"]
    
    col1, col2 = st.columns(2)
    
    def render_mapping_table(title, source, col_obj, unique_key):
        with col_obj:
            st.markdown(f"## {title}")
            only_missing = st.checkbox(f"Chỉ hiện SP Chưa phân loại", value=True, key=f"chk_{unique_key}")
            
            df_src = df_mapping_status[df_mapping_status['Nguồn'] == source].copy()
            if df_src.empty:
                st.info(f"Không có dữ liệu {source}")
                return
                
            if only_missing:
                df_src = df_src[df_src['Khu'] == 'Chưa phân loại']
            
            if df_src.empty:
                st.success("Tuyệt vời! Toàn bộ sản phẩm đã được phân loại.")
                return
                
            edited_df = st.data_editor(
                df_src[['SKU', 'Tên gốc', 'Khu', 'Loại']],
                column_config={
                    "SKU": st.column_config.TextColumn("SKU", disabled=True),
                    "Tên gốc": st.column_config.TextColumn("Tên sản phẩm", disabled=True),
                    "Khu": st.column_config.SelectboxColumn("Khu", options=khu_options),
                    "Loại": st.column_config.SelectboxColumn("Loại", options=loai_options)
                },
                use_container_width=True, key=f"editor_{unique_key}"
            )
            
            if st.button(f"💾 Lưu phân loại {source}", key=f"btn_{unique_key}"):
                changed = edited_df[(edited_df['Khu'] != df_src['Khu']) | (edited_df['Loại'] != df_src['Loại'])]
                if not changed.empty:
                    current_mapping = load_json(MAPPING_FILE, {})
                    for _, row in changed.iterrows():
                        ten_chuan = clean_text(row['Tên gốc'])
                        current_mapping[ten_chuan] = {
                            "Tên sản phẩm": row['Tên gốc'],
                            "Khu": row['Khu'],
                            "Loại": row['Loại']
                        }
                    save_json(MAPPING_FILE, current_mapping)
                    push_to_github_if_configured(MAPPING_FILE)
                    st.cache_data.clear()
                    st.success(f"Đã lưu {len(changed)} sản phẩm!")
                    st.rerun()
            
            if not is_github_connected() and os.path.exists(MAPPING_FILE):
                with open(MAPPING_FILE, 'r', encoding='utf-8') as f: 
                    st.download_button("📥 Tải file Backup Mapping (Thủ công)", f, file_name=MAPPING_FILE, key=f"dl_{unique_key}")

    render_mapping_table("2. Phân loại Kiot Viet", "Kiot Viet", col1, "kiot")
    render_mapping_table("3. Phân loại Haravan", "Haravan", col2, "har")

with tab_dashboard:
    if not df_report.empty:
        df_report['Nhóm Tab'] = df_report.apply(assign_group, axis=1)
        
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
        
        groups = sorted(list(df_final['Nhóm Tab'].unique()))
        if "Khu Khác" in groups: 
            groups.remove("Khu Khác")
            groups.append("Khu Khác")
            
        tabs = st.tabs(groups)
        
        def plot_dual_axis(df, x_col, title):
            fig = go.Figure()
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

        for i, group_name in enumerate(groups):
            with tabs[i]:
                df_group = df_final[df_final['Nhóm Tab'] == group_name]
                if df_group.empty:
                    st.info("Trống")
                    continue
                    
                khu_list = sorted(list(df_group['Khu'].unique()))
                selected_khu = st.multiselect(f"Chọn khu trong {group_name}:", khu_list, default=khu_list, key=f"ms_{i}")
                
                for khu in selected_khu:
                    df_khu = df_group[df_group['Khu'] == khu]
                    if df_khu.empty: continue
                    
                    st.markdown(f"### 🎯 {khu}")
                    
                    df_ngay = df_khu.groupby('Ngày báo cáo').agg({'Số lượt bán':'sum', 'Doanh thu':'sum'}).reset_index()
                    if not df_ngay.empty:
                        df_ngay['Ngày báo cáo'] = pd.to_datetime(df_ngay['Ngày báo cáo'])
                        try:
                            idx = pd.date_range(pd.to_datetime(date_range[0]), pd.to_datetime(date_range[1]))
                            df_ngay = df_ngay.set_index('Ngày báo cáo').reindex(idx, fill_value=0).reset_index()
                            df_ngay.rename(columns={'index': 'Ngày báo cáo'}, inplace=True)
                        except: pass
                        df_ngay['Ngày báo cáo'] = df_ngay['Ngày báo cáo'].dt.strftime('%d/%m')
                    
                    df_tuan = df_khu.groupby(['Năm', 'Tuần ID', 'Tuần Label']).agg({'Số lượt bán':'sum', 'Doanh thu':'sum'}).reset_index()
                    df_tuan = df_tuan.sort_values(['Năm', 'Tuần ID'])
                    
                    df_khu_thang = df_khu.copy()
                    df_khu_thang['Tháng-Năm'] = df_khu_thang['Tháng'].astype(str) + "/" + df_khu_thang['Năm'].astype(str)
                    df_thang = df_khu_thang.groupby(['Năm', 'Tháng', 'Tháng-Năm']).agg({'Số lượt bán':'sum', 'Doanh thu':'sum'}).reset_index()
                    df_thang = df_thang.sort_values(['Năm', 'Tháng'])
                    
                    df_nam = df_khu.groupby('Năm').agg({'Số lượt bán':'sum', 'Doanh thu':'sum'}).reset_index()
                    
                    c1, c2, c3, c4 = st.tabs(["📅 Ngày", "📆 Tuần", "🗓️ Tháng", "🎉 Năm"])
                    with c1: st.plotly_chart(plot_dual_axis(df_ngay, 'Ngày báo cáo', f"Doanh thu & Lượt bán Ngày - {khu}"), use_container_width=True)
                    with c2: st.plotly_chart(plot_dual_axis(df_tuan, 'Tuần Label', f"Doanh thu & Lượt bán Tuần - {khu}"), use_container_width=True)
                    with c3: st.plotly_chart(plot_dual_axis(df_thang, 'Tháng-Năm', f"Doanh thu & Lượt bán Tháng - {khu}"), use_container_width=True)
                    with c4: st.plotly_chart(plot_dual_axis(df_nam, 'Năm', f"Doanh thu & Lượt bán Năm - {khu}"), use_container_width=True)
