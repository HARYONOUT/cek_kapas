import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import os
import glob
import sqlite3
import datetime
from io import BytesIO
import hashlib

# -------------------------------------------------------------------
# 1. KONFIGURASI HALAMAN
# -------------------------------------------------------------------
st.set_page_config(
    page_title="Dashboard Alokasi Kuota Ujian Online",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# -------------------------------------------------------------------
# 2. CUSTOM CSS
# -------------------------------------------------------------------
st.markdown("""
    <style>
    .main { background-color: #f8f9fa; }
    .metric-card {
        background-color: white; padding: 20px; border-radius: 10px;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.05); border-left: 5px solid #007bff;
        margin-bottom: 20px;
    }
    .stAlert { border-radius: 10px; }
    div[data-testid="stMetricValue"] { font-size: 28px; font-weight: bold; }
    .custom-sub { font-size: 14px; color: #6c757d; margin-top: -10px; margin-bottom: 10px; }
    .update-badge {
        background-color: #e8f4fd; color: #1d8cf8; padding: 6px 12px; border-radius: 20px;
        font-size: 13px; font-weight: bold; display: inline-block; margin-bottom: 20px;
    }
    .visit-counter {
        background-color: #f1f3f5; color: #495057; padding: 4px 12px; border-radius: 20px;
        font-size: 14px; font-weight: 500; display: inline-block; margin-right: 10px;
    }
    </style>
""", unsafe_allow_html=True)

# -------------------------------------------------------------------
# 3. KONSTANTA GLOBAL
# -------------------------------------------------------------------
DB_NAME = "kuota_ujian.db"
SESI_REGULER = 5
SESI_TAP = 2

# -------------------------------------------------------------------
# 4. FUNGSI DATABASE (termasuk hit counter)
# -------------------------------------------------------------------
def init_db_if_needed():
    try:
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute("CREATE TABLE IF NOT EXISTS db_metadata (key TEXT PRIMARY KEY, value TEXT)")
        cursor.execute("CREATE TABLE IF NOT EXISTS master_ruang_tap (id_ruang TEXT, id_sekolah TEXT, Tanggal TEXT)")
        cursor.execute("CREATE TABLE IF NOT EXISTS app_visits (id INTEGER PRIMARY KEY, total_visits INTEGER)")
        cursor.execute("INSERT OR IGNORE INTO app_visits (id, total_visits) VALUES (1, 0)")
        conn.commit()
        conn.close()
    except Exception as e:
        st.error(f"Gagal inisialisasi database: {e}")

def increment_visit_count():
    try:
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute("UPDATE app_visits SET total_visits = total_visits + 1 WHERE id = 1")
        conn.commit()
        conn.close()
    except Exception as e:
        st.error(f"Gagal menambah hit counter: {e}")

def get_visit_count():
    try:
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute("SELECT total_visits FROM app_visits WHERE id = 1")
        row = cursor.fetchone()
        conn.close()
        return row[0] if row else 0
    except Exception:
        return 0

def get_metadata(key: str) -> str:
    try:
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute("SELECT value FROM db_metadata WHERE key = ?", (key,))
        row = cursor.fetchone()
        conn.close()
        return row[0] if row else "Belum pernah diperbarui"
    except Exception:
        return "Error"

def set_metadata(key: str, value: str) -> None:
    try:
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute("INSERT OR REPLACE INTO db_metadata (key, value) VALUES (?, ?)", (key, value))
        conn.commit()
        conn.close()
    except Exception as e:
        st.error(f"Gagal menyimpan metadata: {e}")

def is_db_ready() -> bool:
    if not os.path.exists(DB_NAME):
        return False
    try:
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table'")
        tables = {row[0] for row in cursor.fetchall()}
        required = {'master_peserta', 'master_wilayah', 'master_sekolah', 'master_ruang', 'master_ruang_tap'}
        conn.close()
        return required.issubset(tables)
    except Exception:
        return False

def reset_database():
    try:
        st.cache_data.clear()
        if os.path.exists(DB_NAME):
            os.remove(DB_NAME)
        init_db_if_needed()
        return True
    except Exception as e:
        st.error(f"Gagal reset database: {e}")
        return False

# -------------------------------------------------------------------
# 5. FUNGSI BACA BERKAS
# -------------------------------------------------------------------
def read_file(file_or_path):
    try:
        if isinstance(file_or_path, str):
            if file_or_path.endswith(('.xlsx', '.xls')):
                return pd.read_excel(file_or_path)
            else:
                return pd.read_csv(file_or_path)
        else:
            if file_or_path.name.endswith(('.xlsx', '.xls')):
                return pd.read_excel(file_or_path)
            else:
                return pd.read_csv(file_or_path)
    except Exception as e:
        st.error(f"Gagal membaca file: {e}")
        raise

# -------------------------------------------------------------------
# 6. FUNGSI PEMBERSIH DATA
# -------------------------------------------------------------------
def clean_ruang_tap(df: pd.DataFrame) -> pd.DataFrame:
    required = ['id_ruang', 'id_sekolah', 'Tanggal']
    if df.empty:
        return pd.DataFrame(columns=required)
    df.columns = df.columns.astype(str).str.strip()
    col_map = {}
    for req in required:
        candidates = [c for c in df.columns if req.lower() in c.lower()]
        if candidates:
            col_map[req] = candidates[0]
    if len(col_map) != 3:
        st.warning(f"File Ruang TAP harus memiliki kolom: {required}. Ditemukan: {df.columns.tolist()}")
        return pd.DataFrame(columns=required)
    df_clean = pd.DataFrame()
    df_clean['id_ruang'] = df[col_map['id_ruang']].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
    df_clean['id_sekolah'] = df[col_map['id_sekolah']].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
    df_clean['Tanggal'] = pd.to_datetime(df[col_map['Tanggal']], errors='coerce', dayfirst=True).dt.strftime('%Y-%m-%d')
    return df_clean.dropna(subset=['id_ruang', 'Tanggal'])

def standardize_ids(df: pd.DataFrame, id_columns: list) -> pd.DataFrame:
    df = df.copy()
    for col in id_columns:
        if col in df.columns:
            df[col] = df[col].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
    return df

def safe_numeric(df: pd.DataFrame, columns: list) -> pd.DataFrame:
    df = df.copy()
    for col in columns:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0).astype(float)
    return df

# -------------------------------------------------------------------
# 7. LOGIKA UTAMA: HITUNG KAPASITAS DINAMIS + MERGE DATA
# -------------------------------------------------------------------
@st.cache_data(ttl=300, show_spinner=False)
def load_and_process_data():
    conn = sqlite3.connect(DB_NAME)
    try:
        df_master = pd.read_sql_query("SELECT * FROM master_peserta", conn)
        df_wilayah = pd.read_sql_query("SELECT * FROM master_wilayah", conn)
        df_sekolah = pd.read_sql_query("SELECT * FROM master_sekolah", conn)
        df_ruang = pd.read_sql_query("SELECT * FROM master_ruang", conn)
        df_ruang_tap = pd.read_sql_query("SELECT * FROM master_ruang_tap", conn)
    except Exception as e:
        conn.close()
        st.error(f"Gagal membaca dari database: {e}")
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()
    conn.close()

    for df in [df_master, df_wilayah, df_sekolah, df_ruang]:
        df.columns = df.columns.str.strip()

    df_master = standardize_ids(df_master, ['kode_tuo'])
    df_sekolah = standardize_ids(df_sekolah, ['id_sekolah', 'id_wilayah'])
    df_wilayah = standardize_ids(df_wilayah, ['id_wilayah'])
    df_ruang = standardize_ids(df_ruang, ['id_sekolah', 'id_ruang'])
    df_ruang_tap = clean_ruang_tap(df_ruang_tap)

    numeric_cols = ['jml_s2', 'jml_tap', 'jml_s1_objektif', 'jml_s1_uraian']
    df_master = safe_numeric(df_master, numeric_cols)

    if '4212' not in df_sekolah['id_sekolah'].values:
        new_school = pd.DataFrame([{
            'id_sekolah': '4212', 'nama_sekolah': 'ITS NU PEKALONGAN',
            'id_wilayah': '33264', 'alamat': 'Pekalongan', 'ruang': 6
        }])
        df_sekolah = pd.concat([df_sekolah, new_school], ignore_index=True)
    else:
        df_sekolah.loc[df_sekolah['id_sekolah'] == '4212', 'id_wilayah'] = '33264'
    df_sekolah.loc[df_sekolah['id_sekolah'] == '4202', 'id_wilayah'] = '33755'

    existing_rooms = set(df_ruang[df_ruang['id_sekolah'] == '4212']['id_ruang'])
    needed_rooms = [f'4212{i}' for i in range(1, 6)] + ['42126']
    for rid in needed_rooms:
        if rid not in existing_rooms:
            kap = 30 if rid != '42126' else 15
            new_room = pd.DataFrame([{
                'id_ruang': rid, 'id_sekolah': '4212', 'nama_ruang': 'ITS NU PEKALONGAN',
                'kapasitas': kap, 'Ruang': rid[-1]
            }])
            df_ruang = pd.concat([df_ruang, new_room], ignore_index=True)

    if 'tgl_ujian' not in df_master.columns:
        st.error("Kolom 'tgl_ujian' tidak ditemukan di data peserta.")
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()
    df_master['tgl_ujian'] = pd.to_datetime(df_master['tgl_ujian'], errors='coerce').dt.strftime('%Y-%m-%d')
    unique_dates = df_master['tgl_ujian'].dropna().unique()
    if len(unique_dates) == 0:
        unique_dates = ['2026-06-20']

    df_master['total_tap_s2'] = df_master['jml_s2'] + df_master['jml_tap']
    tap_dict = df_master.groupby(['kode_tuo', 'tgl_ujian'])['total_tap_s2'].sum().to_dict()
    active_tap_rooms = set(zip(df_ruang_tap['id_ruang'], df_ruang_tap['Tanggal']))

    df_ruang['kapasitas'] = pd.to_numeric(df_ruang['kapasitas'], errors='coerce').fillna(0).astype(int)
    rooms_by_school = df_ruang.groupby('id_sekolah')[['id_ruang', 'kapasitas']].apply(lambda x: x.to_dict('records')).to_dict()

    capacity_records = []
    for sekolah in df_sekolah['id_sekolah'].unique():
        school_rooms = rooms_by_school.get(sekolah, [])
        if not school_rooms:
            continue
        for tgl in unique_dates:
            total_tap = tap_dict.get((sekolah, tgl), 0)
            total_cap = 0
            is_tap = False
            for rm in school_rooms:
                rid = rm['id_ruang']
                kap = rm['kapasitas']
                if total_tap == 0:
                    sesi = SESI_REGULER
                else:
                    if (rid, tgl) in active_tap_rooms:
                        sesi = SESI_TAP
                        is_tap = True
                    else:
                        sesi = SESI_REGULER
                total_cap += kap * sesi
            capacity_records.append({
                'id_sekolah': sekolah, 'tgl_ujian': tgl,
                'kapasitas_total': total_cap, 'is_tap': is_tap and total_tap > 0
            })
    df_capacity = pd.DataFrame(capacity_records)

    df_school_wil = pd.merge(df_sekolah, df_wilayah, on='id_wilayah', how='left')
    df_merged = pd.merge(df_master, df_school_wil, left_on='kode_tuo', right_on='id_sekolah', how='left')
    df_merged = pd.merge(df_merged, df_capacity, left_on=['kode_tuo', 'tgl_ujian'], right_on=['id_sekolah', 'tgl_ujian'], how='left')
    df_merged = df_merged.drop(columns=['id_sekolah_x', 'id_sekolah_y'], errors='ignore')
    if 'id_sekolah' not in df_merged.columns:
        df_merged['id_sekolah'] = df_merged['kode_tuo']

    df_merged['total_peserta'] = df_merged[numeric_cols].sum(axis=1)
    df_merged['kapasitas_total'] = pd.to_numeric(df_merged['kapasitas_total'], errors='coerce').fillna(0).astype(int)
    df_merged['sisa_kuota'] = df_merged['kapasitas_total'] - df_merged['total_peserta']
    df_merged['persentase_keterisian'] = np.where(
        df_merged['kapasitas_total'] > 0,
        (df_merged['total_peserta'] / df_merged['kapasitas_total'] * 100).round(1),
        0.0
    )
    df_merged['nama_sekolah'] = df_merged['nama_sekolah'].fillna(df_merged.get('nama_tuo', 'Tidak diketahui'))
    df_merged['sisa_kuota'] = df_merged['sisa_kuota'].astype(float)
    df_merged['persentase_keterisian'] = df_merged['persentase_keterisian'].astype(float)
    
    return df_merged, df_ruang, df_sekolah

# -------------------------------------------------------------------
# 8. FUNGSI STYLING TABEL
# -------------------------------------------------------------------
def get_color_style(pct, is_tap=False):
    try:
        if pd.isna(pct):
            return 'background-color: #f1f3f5; color: #adb5bd; text-align: center;'
        pct_val = float(pct)
    except (ValueError, TypeError):
        return 'background-color: #f1f3f5; color: #adb5bd; text-align: center;'
    
    if is_tap:
        return 'background-color: #cce5ff; color: #004085; font-weight: bold; text-align: center; border: 1.5px solid #b8daff;'
    if pct_val > 100.0:
        return 'background-color: #fff3cd; color: #856404; font-weight: bold; text-align: center;'
    if pct_val == 100.0:
        return 'background-color: #f8d7da; color: #721c24; font-weight: bold; text-align: center;'
    if pct_val >= 50.0:
        return 'background-color: #a9dfbf; color: #196f3d; font-weight: bold; text-align: center;'
    return 'background-color: #d4edda; color: #155724; font-weight: bold; text-align: center;'

def style_matrix(df_pct, df_tap):
    styled = pd.DataFrame('', index=df_pct.index, columns=df_pct.columns)
    for r in df_pct.index:
        for c in df_pct.columns:
            pct = df_pct.loc[r, c] if r in df_pct.index and c in df_pct.columns else np.nan
            is_tap_val = df_tap.loc[r, c] if r in df_tap.index and c in df_tap.columns else False
            styled.loc[r, c] = get_color_style(pct, is_tap_val)
    return styled

# -------------------------------------------------------------------
# 9. SIDEBAR & FILTER
# -------------------------------------------------------------------
def render_sidebar(df):
    st.sidebar.header("🔍 Filter Wilayah & Jadwal")
    kab_list = sorted(df['nama_kabupaten'].dropna().unique())
    selected_kab = st.sidebar.selectbox("Pilih Kabupaten / Kota:", ["SEMUA KABUPATEN"] + kab_list)
    tgl_list = sorted(df['tgl_ujian'].dropna().unique())
    selected_tgl = st.sidebar.multiselect("Pilih Tanggal Ujian:", tgl_list, default=tgl_list)
    status_filter = st.sidebar.radio(
        "Filter Status Ketersediaan:",
        ["Tampilkan Semua", "Hanya yang Penuh / Melebihi Kuota", "Hanya yang Aman (Tersisa)"]
    )
    return selected_kab, selected_tgl, status_filter

def apply_filters(df, kab, tgl_list, status):
    df_filtered = df.copy()
    if kab != "SEMUA KABUPATEN":
        df_filtered = df_filtered[df_filtered['nama_kabupaten'] == kab]
    if tgl_list:
        df_filtered = df_filtered[df_filtered['tgl_ujian'].isin(tgl_list)]
    if status == "Hanya yang Penuh / Melebihi Kuota":
        df_filtered = df_filtered[df_filtered['sisa_kuota'] <= 0]
    elif status == "Hanya yang Aman (Tersisa)":
        df_filtered = df_filtered[df_filtered['sisa_kuota'] > 0]
    return df_filtered

# -------------------------------------------------------------------
# 10. RENDER METRIK UTAMA
# -------------------------------------------------------------------
def render_metrics(df_filtered):
    total_peserta = df_filtered['total_peserta'].sum()
    unique_capacity = df_filtered[['kode_tuo', 'tgl_ujian', 'kapasitas_total']].drop_duplicates()
    total_daya_tampung = unique_capacity['kapasitas_total'].sum()
    sisa_total = total_daya_tampung - total_peserta
    sekolah_aktif = df_filtered['kode_tuo'].nunique()
    hari_dipakai = df_filtered['tgl_ujian'].nunique()
    
    col1, col2, col3, col4, col5 = st.columns(5)
    with col1:
        st.markdown(f"""
            <div class="metric-card" style="border-left-color: #28a745;">
                <p style="color: #6c757d; margin-bottom: 2px; font-size:14px;">Total Peserta</p>
                <h3 style="margin:0; font-size:24px;">{int(total_peserta)} Orang</h3>
            </div>
        """, unsafe_allow_html=True)
    with col2:
        st.markdown(f"""
            <div class="metric-card" style="border-left-color: #007bff;">
                <p style="color: #6c757d; margin-bottom: 2px; font-size:14px;">Daya Tampung Total</p>
                <h3 style="margin:0; font-size:24px;">{int(total_daya_tampung)} Kursi</h3>
            </div>
        """, unsafe_allow_html=True)
    with col3:
        warna = "#28a745" if sisa_total >= 0 else "#dc3545"
        st.markdown(f"""
            <div class="metric-card" style="border-left-color: {warna};">
                <p style="color: #6c757d; margin-bottom: 2px; font-size:14px;">Sisa Kuota Gabungan</p>
                <h3 style="margin:0; font-size:24px;">{int(sisa_total)} Kursi</h3>
            </div>
        """, unsafe_allow_html=True)
    with col4:
        st.markdown(f"""
            <div class="metric-card" style="border-left-color: #ffc107;">
                <p style="color: #6c757d; margin-bottom: 2px; font-size:14px;">Lokasi Ujian Aktif</p>
                <h3 style="margin:0; font-size:24px;">{sekolah_aktif} Lokasi</h3>
            </div>
        """, unsafe_allow_html=True)
    with col5:
        st.markdown(f"""
            <div class="metric-card" style="border-left-color: #17a2b8;">
                <p style="color: #6c757d; margin-bottom: 2px; font-size:14px;">Hari yang Digunakan</p>
                <h3 style="margin:0; font-size:24px;">{int(hari_dipakai)} Hari</h3>
            </div>
        """, unsafe_allow_html=True)

# -------------------------------------------------------------------
# 11. RENDER TABEL PIVOT
# -------------------------------------------------------------------
def render_pivot_table(df_filtered):
    st.subheader("🗓️ Visualisasi Sisa Kuota Lokasi per Tanggal (Kapasitas Dinamis)")
    st.markdown(
        "Kapasitas dihitung berdasarkan **Daya Tampung Ruang × 5 Sesi** (reguler) atau **× 2 Sesi** (jika ada ujian TAP/S2 pada tanggal tersebut)."
    )
    if df_filtered.empty:
        st.info("Tidak ada data dengan filter yang dipilih.")
        return
    
    view_option = st.radio(
        "Pilih Model Tampilan Data:",
        ["Tampilkan Sisa Kuota", "Tampilkan Rasio Keterisian (%)", "Tampilkan Format (Peserta / Kapasitas Maks)"],
        horizontal=True
    )
    
    df_filtered['persentase_keterisian'] = df_filtered['persentase_keterisian'].astype(float)
    df_pivot_pct = df_filtered.pivot_table(
        index='nama_sekolah', columns='tgl_ujian', values='persentase_keterisian', aggfunc='mean'
    )
    df_pivot_tap = df_filtered.pivot_table(
        index='nama_sekolah', columns='tgl_ujian', values='is_tap', aggfunc='max'
    ).fillna(False).astype(bool)
    
    if view_option == "Tampilkan Sisa Kuota":
        df_pivot = df_filtered.pivot_table(
            index='nama_sekolah', columns='tgl_ujian', values='sisa_kuota', aggfunc='sum'
        )
        styled = df_pivot.style.apply(lambda _: style_matrix(df_pivot_pct, df_pivot_tap), axis=None)
        st.dataframe(styled.format("{:.0f}", na_rep="-"), use_container_width=True, height=450)
    elif view_option == "Tampilkan Rasio Keterisian (%)":
        df_pivot = df_filtered.pivot_table(
            index='nama_sekolah', columns='tgl_ujian', values='persentase_keterisian', aggfunc='mean'
        )
        styled = df_pivot.style.apply(lambda _: style_matrix(df_pivot_pct, df_pivot_tap), axis=None)
        st.dataframe(styled.format("{:.1f}%", na_rep="-"), use_container_width=True, height=450)
    else:
        df_peserta = df_filtered.pivot_table(index='nama_sekolah', columns='tgl_ujian', values='total_peserta', aggfunc='sum')
        df_kapasitas = df_filtered.pivot_table(index='nama_sekolah', columns='tgl_ujian', values='kapasitas_total', aggfunc='sum')
        df_combined = df_peserta.copy()
        for col in df_peserta.columns:
            peserta_str = df_peserta[col].fillna(0).astype(int).astype(str)
            kap_str = df_kapasitas[col].fillna(0).astype(int).astype(str)
            df_combined[col] = peserta_str + " / " + kap_str
            df_combined.loc[df_peserta[col].isna(), col] = "-"
        styled = df_combined.style.apply(lambda _: style_matrix(df_pivot_pct, df_pivot_tap), axis=None)
        st.dataframe(styled, use_container_width=True, height=450)

# -------------------------------------------------------------------
# 12. TAB: GRAFIK, DETAIL RUANG, EKSPOR
# -------------------------------------------------------------------
def render_tabs(df_filtered, df_ruang):
    tab1, tab2, tab3 = st.tabs(["📊 Analisis Grafik", "📋 Struktur Kapasitas Ruang", "💾 Ekspor Laporan"])
    with tab1:
        st.subheader("Grafik Analisis Alokasi per Lokasi (Kapasitas Total)")
        if not df_filtered.empty:
            chart_data = df_filtered.groupby('nama_sekolah')[['total_peserta', 'kapasitas_total']].sum().reset_index()
            fig = go.Figure()
            fig.add_trace(go.Bar(x=chart_data['nama_sekolah'], y=chart_data['kapasitas_total'],
                                 name='Daya Tampung Maksimal', marker_color='#6c757d', opacity=0.6))
            fig.add_trace(go.Bar(x=chart_data['nama_sekolah'], y=chart_data['total_peserta'],
                                 name='Total Peserta Terdaftar', marker_color='#007bff'))
            fig.update_layout(barmode='group', xaxis_title='Nama Lokasi', yaxis_title='Jumlah Kursi',
                              legend_title='Keterangan', template='plotly_white', height=500)
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("Pilih filter terlebih dahulu.")
    with tab2:
        st.subheader("Rincian Kapasitas Ruangan & Skema Sesi")
        schools_active = df_filtered['kode_tuo'].unique()
        df_ruang_filtered = df_ruang[df_ruang['id_sekolah'].isin(schools_active)].copy()
        if not df_ruang_filtered.empty:
            show_cols = ['id_sekolah', 'nama_ruang', 'Ruang', 'kapasitas']
            available = [c for c in show_cols if c in df_ruang_filtered.columns]
            df_ruang_filtered['kapasitas_5_sesi'] = df_ruang_filtered['kapasitas'] * SESI_REGULER
            st.dataframe(df_ruang_filtered[available + ['kapasitas_5_sesi']],
                         use_container_width=True, hide_index=True)
        else:
            st.info("Tidak ada data ruang untuk lokasi yang difilter.")
    with tab3:
        st.subheader("Ekspor Hasil Analisis")
        export_df = df_filtered[['nama_kabupaten', 'nama_sekolah', 'tgl_ujian',
                                 'total_peserta', 'kapasitas_total', 'sisa_kuota', 'persentase_keterisian']].copy()
        export_df.columns = ['Kabupaten', 'Nama Lokasi', 'Tanggal Ujian',
                             'Total Peserta', 'Kapasitas Maksimal', 'Sisa Kuota', 'Persentase Keterisian (%)']
        csv = export_df.to_csv(index=False).encode('utf-8')
        st.download_button("📥 Unduh Laporan (.CSV)", data=csv,
                           file_name=f"laporan_alokasi_{datetime.datetime.now().strftime('%Y%m%d')}.csv",
                           mime="text/csv", use_container_width=True)
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            export_df.to_excel(writer, sheet_name='Alokasi Kuota', index=False)
        st.download_button("📊 Unduh Laporan (Excel)", data=output.getvalue(),
                           file_name=f"laporan_alokasi_{datetime.datetime.now().strftime('%Y%m%d')}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           use_container_width=True)

# -------------------------------------------------------------------
# 13. PANEL ADMIN (dengan password)
# -------------------------------------------------------------------
def check_password() -> bool:
    """Verifikasi password admin. Gunakan st.secrets jika ada, fallback ke hardcode."""
    if 'admin_authenticated' not in st.session_state:
        st.session_state.admin_authenticated = False
    
    if st.session_state.admin_authenticated:
        return True
    
    # Coba ambil dari secrets (untuk deployment)
    try:
        correct_password = st.secrets["admin"]["password"]
    except (FileNotFoundError, KeyError, AttributeError):
        # Fallback: hardcode password (ubah sesuai keinginan Anda)
        correct_password = "admin123"
    
    input_password = st.text_input("Masukkan Password Admin:", type="password", key="admin_pw")
    if st.button("Login sebagai Admin", key="admin_login"):
        if input_password == correct_password:
            st.session_state.admin_authenticated = True
            st.rerun()
        else:
            st.error("Password salah. Akses ditolak.")
    return False

def render_admin_panel():
    """Hanya tampilkan isi panel jika sudah terautentikasi."""
    with st.expander("⚙️ Konfigurasi Basis Data & Unggah File Master (Khusus Admin)", expanded=False):
        if not check_password():
            st.warning("🔒 Bagian ini dilindungi password. Silakan login di atas untuk mengakses konfigurasi database dan unggah file.")
            return
        
        # Setelah login, tampilkan seluruh konten admin
        st.success("✅ Mode Admin Aktif - Anda dapat melakukan perubahan data.")
        st.markdown("Bagian ini digunakan untuk melakukan update master data secara manual.")
        
        st.subheader("🚨 Fitur Pembersihan Database")
        if st.button("Hapus & Reset Semua Data Database (Mulai Baru)", type="primary"):
            if reset_database():
                st.success("✅ Database SQLite telah direset sepenuhnya! Silakan unggah file baru.")
                st.rerun()
        st.markdown("---")
        
        if is_db_ready():
            try:
                conn = sqlite3.connect(DB_NAME)
                n_peserta = conn.execute("SELECT COUNT(*) FROM master_peserta").fetchone()[0]
                n_wilayah = conn.execute("SELECT COUNT(*) FROM master_wilayah").fetchone()[0]
                n_sekolah = conn.execute("SELECT COUNT(*) FROM master_sekolah").fetchone()[0]
                n_ruang = conn.execute("SELECT COUNT(*) FROM master_ruang").fetchone()[0]
                n_ruang_tap = conn.execute("SELECT COUNT(*) FROM master_ruang_tap").fetchone()[0]
                conn.close()
                st.markdown("### 📊 Ringkasan Baris Data SQLite Aktif:")
                col1, col2, col3, col4, col5 = st.columns(5)
                col1.metric("Data Peserta", f"{n_peserta} baris")
                col2.metric("Master Wilayah", f"{n_wilayah} baris")
                col3.metric("Master Sekolah", f"{n_sekolah} baris")
                col4.metric("Master Ruang", f"{n_ruang} baris")
                col5.metric("Master Ruang TAP", f"{n_ruang_tap} baris")
                st.markdown("---")
            except Exception:
                pass
        
        col1, col2 = st.columns(2)
        with col1:
            st.subheader("Opsi 1: Muat Ulang File Lokal")
            files = glob.glob("*.csv") + glob.glob("*.xlsx") + glob.glob("*.xls")
            master_file = next((f for f in files if "master" in f.lower() and "ruang" not in f.lower() and "sekolah" not in f.lower() and "wilayah" not in f.lower()), None)
            wilayah_file = next((f for f in files if "wilayah" in f.lower()), None)
            sekolah_file = next((f for f in files if "sekolah" in f.lower()), None)
            ruang_file = next((f for f in files if "ruang" in f.lower() and "tap" not in f.lower()), None)
            ruang_tap_file = next((f for f in files if "ruang_tap" in f.lower() or "ruang tap" in f.lower()), None)
            if all([master_file, wilayah_file, sekolah_file, ruang_file]):
                st.success("📂 File CSV/Excel terdeteksi di direktori!")
                if st.button("🔄 Sinkronisasi Semua File Lokal ke SQLite"):
                    try:
                        df_m = read_file(master_file)
                        df_w = read_file(wilayah_file)
                        df_s = read_file(sekolah_file)
                        df_r = read_file(ruang_file)
                        df_rt = read_file(ruang_tap_file) if ruang_tap_file else pd.DataFrame()
                        conn = sqlite3.connect(DB_NAME)
                        df_m.to_sql('master_peserta', conn, if_exists='replace', index=False)
                        df_w.to_sql('master_wilayah', conn, if_exists='replace', index=False)
                        df_s.to_sql('master_sekolah', conn, if_exists='replace', index=False)
                        df_r.to_sql('master_ruang', conn, if_exists='replace', index=False)
                        df_rt_clean = clean_ruang_tap(df_rt) if not df_rt.empty else pd.DataFrame(columns=['id_ruang','id_sekolah','Tanggal'])
                        df_rt_clean.to_sql('master_ruang_tap', conn, if_exists='replace', index=False)
                        conn.close()
                        set_metadata("last_peserta_update", datetime.datetime.now().strftime("%d-%m-%Y %H:%M:%S WIB"))
                        st.success("✅ Sinkronisasi berhasil!")
                        st.cache_data.clear()
                        st.rerun()
                    except Exception as e:
                        st.error(f"Gagal sinkronisasi: {e}")
            else:
                st.info("Tidak ditemukan kelima file master di direktori.")
        with col2:
            st.subheader("Opsi 2: Unggah Manual (Satuan atau Sekaligus)")
            uploaded_master = st.file_uploader("1. Data Peserta", type=["csv","xlsx","xls"])
            uploaded_wilayah = st.file_uploader("2. Master Wilayah", type=["csv","xlsx","xls"])
            uploaded_sekolah = st.file_uploader("3. Master Sekolah", type=["csv","xlsx","xls"])
            uploaded_ruang = st.file_uploader("4. Master Ruang", type=["csv","xlsx","xls"])
            uploaded_ruang_tap = st.file_uploader("5. File Ruang TAP", type=["csv","xlsx","xls"])
            if st.button("💾 Simpan & Perbarui Berkas Terpilih ke SQLite"):
                try:
                    conn = sqlite3.connect(DB_NAME)
                    if uploaded_master:
                        df = read_file(uploaded_master)
                        df = safe_numeric(df, ['jml_s2','jml_tap','jml_s1_objektif','jml_s1_uraian'])
                        df.to_sql('master_peserta', conn, if_exists='replace', index=False)
                        set_metadata("last_peserta_update", datetime.datetime.now().strftime("%d-%m-%Y %H:%M:%S WIB"))
                        st.success("Data Peserta tersimpan.")
                    if uploaded_wilayah:
                        read_file(uploaded_wilayah).to_sql('master_wilayah', conn, if_exists='replace', index=False)
                        st.success("Master Wilayah tersimpan.")
                    if uploaded_sekolah:
                        read_file(uploaded_sekolah).to_sql('master_sekolah', conn, if_exists='replace', index=False)
                        st.success("Master Sekolah tersimpan.")
                    if uploaded_ruang:
                        df_r = read_file(uploaded_ruang)
                        if 'kapasitas' in df_r.columns:
                            df_r['kapasitas'] = pd.to_numeric(df_r['kapasitas'], errors='coerce').fillna(0).astype(int)
                        df_r.to_sql('master_ruang', conn, if_exists='replace', index=False)
                        st.success("Master Ruang tersimpan.")
                    if uploaded_ruang_tap:
                        df_rt = read_file(uploaded_ruang_tap)
                        df_rt_clean = clean_ruang_tap(df_rt)
                        df_rt_clean.to_sql('master_ruang_tap', conn, if_exists='replace', index=False)
                        st.success("Master Ruang TAP tersimpan.")
                    conn.close()
                    st.cache_data.clear()
                    st.rerun()
                except Exception as e:
                    st.error(f"Gagal menyimpan: {e}")

# -------------------------------------------------------------------
# 14. MAIN APP
# -------------------------------------------------------------------
def main():
    init_db_if_needed()
    
    if 'visit_counted' not in st.session_state:
        increment_visit_count()
        st.session_state['visit_counted'] = True
    
    total_visits = get_visit_count()
    
    st.title("📊 Dashboard Pemantauan Kuota Ujian Online (UO)")
    st.markdown("<p class='custom-sub'>Sistem monitoring ketersediaan daya tampung sekolah (Kapasitas Ruang x Sesi Ujian) vs jumlah peserta real-time berbasis SQLite</p>", unsafe_allow_html=True)
    
    col_left, col_right = st.columns([3, 1])
    with col_left:
        last_update = get_metadata("last_peserta_update")
        st.markdown(f"<div class='update-badge'>🕒 Data Peserta Terakhir Diperbarui: {last_update}</div>", unsafe_allow_html=True)
    with col_right:
        st.markdown(f"<div class='visit-counter'>👥 Total Dikunjungi: {total_visits} kali</div>", unsafe_allow_html=True)
    
    render_admin_panel()
    
    if not is_db_ready():
        st.info("👋 Selamat datang! Silakan unggah kelima file master melalui panel di atas untuk memulai.")
        st.image("https://images.unsplash.com/photo-1551288049-bebda4e38f71?auto=format&fit=crop&w=1200&q=80", caption="Sistem Visualisasi Manajemen Alokasi Ruang & Kuota")
        return
    
    with st.spinner("Memuat data dari database..."):
        df_final, df_ruang, _ = load_and_process_data()
    
    if df_final.empty:
        st.warning("Data tidak ditemukan. Pastikan file master sudah lengkap dan berisi data.")
        return
    
    selected_kab, selected_tgl, status_filter = render_sidebar(df_final)
    df_filtered = apply_filters(df_final, selected_kab, selected_tgl, status_filter)
    render_metrics(df_filtered)
    render_pivot_table(df_filtered)
    render_tabs(df_filtered, df_ruang)

if __name__ == "__main__":
    main()