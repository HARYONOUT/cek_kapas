import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import os
import glob
import sqlite3
import datetime

# Konfigurasi halaman Streamlit
st.set_page_config(
    page_title="Dashboard Alokasi Kuota Ujian Online",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS untuk tampilan premium & responsif
st.markdown("""
    <style>
    .main {
        background-color: #f8f9fa;
    }
    .metric-card {
        background-color: white;
        padding: 20px;
        border-radius: 10px;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.05);
        border-left: 5px solid #007bff;
        margin-bottom: 20px;
    }
    .stAlert {
        border-radius: 10px;
    }
    div[data-testid="stMetricValue"] {
        font-size: 28px;
        font-weight: bold;
    }
    .custom-sub {
        font-size: 14px;
        color: #6c757d;
        margin-top: -10px;
        margin-bottom: 10px;
    }
    .update-badge {
        background-color: #e8f4fd;
        color: #1d8cf8;
        padding: 6px 12px;
        border-radius: 20px;
        font-size: 13px;
        font-weight: bold;
        display: inline-block;
        margin-bottom: 20px;
    }
    </style>
""", unsafe_allow_html=True)

# Judul Utama Dashboard
st.title("📊 Dashboard Pemantauan Kuota Ujian Online (UO)")
st.markdown("<p class='custom-sub'>Sistem monitoring ketersediaan daya tampung sekolah (Kapasitas Ruang x Sesi Ujian) vs jumlah peserta real-time berbasis SQLite</p>", unsafe_allow_html=True)

DB_NAME = "kuota_ujian.db"

# Fungsi inisialisasi tabel SQLite tambahan jika diperlukan (backwards compatibility)
def inisialisasi_db_jika_perlu():
    try:
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS master_ruang_tap (
                id_ruang TEXT,
                id_sekolah TEXT,
                Tanggal TEXT
            )
        """)
        conn.commit()
        conn.close()
    except Exception:
        pass

# Jalankan inisialisasi database
inisialisasi_db_jika_perlu()

# Fungsi membaca berkas baik dalam format CSV maupun EXCEL (XLSX/XLS) secara cerdas
def baca_berkas(file_path_atau_buffer):
    if isinstance(file_path_atau_buffer, str):
        if file_path_atau_buffer.endswith(('.xlsx', '.xls')):
            return pd.read_excel(file_path_atau_buffer)
        return pd.read_csv(file_path_atau_buffer)
    else:
        if file_path_atau_buffer.name.endswith(('.xlsx', '.xls')):
            return pd.read_excel(file_path_atau_buffer)
        return pd.read_csv(file_path_atau_buffer)

# Fungsi pembersih & standardisasi berkas RUANG_TAP agar robust terhadap shift kolom / rename
def clean_ruang_tap(df):
    df_new = pd.DataFrame()
    if df.empty:
        return pd.DataFrame(columns=['id_ruang', 'id_sekolah', 'Tanggal'])
    
    # Bersihkan spasi pada header
    df.columns = df.columns.astype(str).str.strip()
    
    id_ruang_col = None
    id_sekolah_col = None
    tanggal_col = None
    
    # Deteksi kolom id_ruang dan id_sekolah secara dinamis
    for col in df.columns:
        c_lower = col.lower()
        if 'id_ruang' in c_lower:
            id_ruang_col = col
        elif 'id_sekolah' in c_lower:
            id_sekolah_col = col
            
    # Temukan kolom tanggal yang valid berdasarkan analisis isi datanya (menghindari kolom teks salah nama)
    # Tahap A: Cari kolom yang namanya mengandung indikasi tanggal DAN isi datanya memang format tanggal
    for col in df.columns:
        c_lower = col.lower()
        if 'tanggal' in c_lower or 'tgl' in c_lower or 'date' in c_lower:
            sample = df[col].astype(str).dropna()
            if sample.str.contains(r'(\d{4}[-/]\d{2}[-/]\d{2})|(\d{2}[-/]\d{2}[-/]\d{4})').any():
                tanggal_col = col
                break
                
    # Tahap B: Jika tidak ketemu, cari kolom mana saja yang isinya mengandung struktur tanggal (tanpa mempedulikan nama header)
    if not tanggal_col:
        for col in df.columns:
            sample = df[col].astype(str).dropna()
            if sample.str.contains(r'(\d{4}[-/]\d{2}[-/]\d{2})|(\d{2}[-/]\d{2}[-/]\d{4})').any():
                tanggal_col = col
                break
            
    # Tahap C: Jika tetap tidak ada yang lolos validasi isi, fallback ke pencarian teks biasa pada nama kolom
    if not tanggal_col:
        for col in df.columns:
            c_lower = col.lower()
            if 'tanggal' in c_lower or 'tgl' in c_lower or 'date' in c_lower:
                tanggal_col = col
                break
            
    # Fallback pencarian alternatif jika tidak pas 100%
    if not id_ruang_col and len(df.columns) > 0:
        id_ruang_col = df.columns[0]
    if not id_sekolah_col and len(df.columns) > 1:
        id_sekolah_col = df.columns[1]
    if not tanggal_col and len(df.columns) > 2:
        tanggal_col = df.columns[2]
            
    # Standardisasi data ke dataframe baru
    if id_ruang_col:
        df_new['id_ruang'] = df[id_ruang_col].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
    if id_sekolah_col:
        df_new['id_sekolah'] = df[id_sekolah_col].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
    if tanggal_col:
        df_new['Tanggal'] = pd.to_datetime(df[tanggal_col], errors='coerce', dayfirst=True).dt.strftime('%Y-%m-%d')
        
    return df_new.dropna(subset=['id_ruang', 'Tanggal'])

# Fungsi untuk menyimpan dan mengambil metadata dari SQLite (untuk melacak waktu pembaruan)
def dapatkan_metadata(key):
    try:
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute("CREATE TABLE IF NOT EXISTS db_metadata (key TEXT PRIMARY KEY, value TEXT)")
        cursor.execute("SELECT value FROM db_metadata WHERE key = ?", (key,))
        row = cursor.fetchone()
        conn.close()
        if row:
            return row[0]
    except Exception:
        pass
    return "Belum pernah diperbarui"

def simpan_metadata(key, value):
    try:
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute("CREATE TABLE IF NOT EXISTS db_metadata (key TEXT PRIMARY KEY, value TEXT)")
        cursor.execute("INSERT OR REPLACE INTO db_metadata (key, value) VALUES (?, ?)", (key, value))
        conn.commit()
        conn.close()
    except Exception:
        pass

# Menampilkan informasi tanggal dan jam pembaruan terakhir data peserta
waktu_update = dapatkan_metadata("last_peserta_update")
st.markdown(f"<div class='update-badge'>🕒 Data Peserta Terakhir Diperbarui: {waktu_update}</div>", unsafe_allow_html=True)

# Fungsi pembantu untuk menghasilkan warna gradasi baru (5 Kategori Warna termasuk Biru TAP)
def get_color_style(pct, is_tap=False):
    if pd.isna(pct):
        return 'background-color: #f1f3f5; color: #adb5bd; text-align: center;'
    
    # 0. Jika ada ujian TAP/S2/Essay pada tanggal tersebut -> Biru Premium
    if is_tap:
        return 'background-color: #cce5ff; color: #004085; font-weight: bold; text-align: center; border: 1.5px solid #b8daff;'
    
    # 1. Di atas 100% -> Kuning (Anomali data / Over-capacity)
    if pct > 100.0:
        return 'background-color: #fff3cd; color: #856404; font-weight: bold; text-align: center; border: 1.5px solid #ffeeba;'
        
    # 2. Tepat 100% -> Merah
    if pct == 100.0:
        return 'background-color: #f8d7da; color: #721c24; font-weight: bold; text-align: center; border: 1.5px solid #f5c6cb;'
        
    # 3. 50% s.d. 99% -> Hijau muda lebih tua sedikit
    if pct >= 50.0:
        return 'background-color: #a9dfbf; color: #196f3d; font-weight: bold; text-align: center;'
        
    # 4. Di bawah 50% -> Hijau Muda
    return 'background-color: #d4edda; color: #155724; font-weight: bold; text-align: center;'

# Fungsi untuk memeriksa apakah SQLite database sudah memiliki data yang lengkap
def cek_db_tersimpan():
    if not os.path.exists(DB_NAME):
        return False
    try:
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table'")
        tables = [r[0] for r in cursor.fetchall()]
        required_tables = {'master_peserta', 'master_wilayah', 'master_sekolah', 'master_ruang', 'master_ruang_tap'}
        conn.close()
        return required_tables.issubset(set(tables))
    except Exception:
        return False

# Fungsi menyimpan data mentah ke SQLite (Mengganti data lama secara penuh)
def simpan_ke_sqlite(df_m, df_w, df_s, df_r, df_rt=None, update_peserta_only=False):
    conn = sqlite3.connect(DB_NAME)
    if not update_peserta_only:
        df_m.to_sql('master_peserta', conn, if_exists='replace', index=False)
        df_w.to_sql('master_wilayah', conn, if_exists='replace', index=False)
        df_s.to_sql('master_sekolah', conn, if_exists='replace', index=False)
        df_r.to_sql('master_ruang', conn, if_exists='replace', index=False)
        if df_rt is not None:
            # Bersihkan dan standardisasi df_rt sebelum simpan
            df_rt_clean = clean_ruang_tap(df_rt)
            df_rt_clean.to_sql('master_ruang_tap', conn, if_exists='replace', index=False)
    else:
        # Jika hanya memperbarui peserta saja
        df_m.to_sql('master_peserta', conn, if_exists='replace', index=False)
        
    conn.close()
    
    # Rekam waktu jam update saat ini
    waktu_sekarang = datetime.datetime.now().strftime("%d-%m-%Y %H:%M:%S WIB")
    simpan_metadata("last_peserta_update", waktu_sekarang)
    
    # Hapus cache streamlit agar data baru langsung tampil tanpa tertahan cache lama
    st.cache_data.clear()

# Fungsi Deteksi File Lokal Otomatis & Cerdas
def cari_file_lokal():
    semua_files = glob.glob("*.csv") + glob.glob("*.xlsx") + glob.glob("*.xls")
    master_file = None
    wilayah_file = None
    sekolah_file = None
    ruang_file = None
    ruang_tap_file = None
    
    for f in semua_files:
        f_lower = f.lower()
        if "ruang_tap" in f_lower or "ruang tap" in f_lower:
            ruang_tap_file = f
        elif "master_ruang" in f_lower or "master ruang" in f_lower:
            ruang_file = f
        elif "master_sekolah" in f_lower or "master sekolah" in f_lower:
            sekolah_file = f
        elif "master_wilayah" in f_lower or "master wilayah" in f_lower:
            wilayah_file = f
        elif "master" in f_lower or "rekap_peserta" in f_lower or "worksheet" in f_lower:
            # Pastikan bukan file master lainnya
            if "ruang" not in f_lower and "sekolah" not in f_lower and "wilayah" not in f_lower:
                master_file = f
                
    return master_file, wilayah_file, sekolah_file, ruang_file, ruang_tap_file

# Fungsi Pembacaan, Penggabungan, dan Kalkulasi Data dari SQLite (Kapasitas Dinamis per Tanggal)
@st.cache_data
def load_and_process_data_from_db():
    conn = sqlite3.connect(DB_NAME)
    df_master = pd.read_sql_query("SELECT * FROM master_peserta", conn)
    df_wilayah = pd.read_sql_query("SELECT * FROM master_wilayah", conn)
    df_sekolah = pd.read_sql_query("SELECT * FROM master_sekolah", conn)
    df_ruang = pd.read_sql_query("SELECT * FROM master_ruang", conn)
    try:
        df_ruang_tap_raw = pd.read_sql_query("SELECT * FROM master_ruang_tap", conn)
    except Exception:
        df_ruang_tap_raw = pd.DataFrame(columns=['id_ruang', 'id_sekolah', 'Tanggal'])
    conn.close()
    
    # 1. Bersihkan nama kolom dari spasi tidak sengaja
    for df in [df_master, df_wilayah, df_sekolah, df_ruang]:
        df.columns = df.columns.str.strip()
        
    # Standardisasi df_ruang_tap secara eksplist dari database tanpa re-run column detection dinamis yang riskan
    df_ruang_tap = df_ruang_tap_raw.copy()
    df_ruang_tap.columns = df_ruang_tap.columns.astype(str).str.strip()
        
    # Helper untuk menstandardisasi ID / Kode agar tipe data COCOK saat digabung
    def standard_id(series):
        return series.astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
        
    # Standardisasi kolom relasi data
    if 'id_sekolah' in df_sekolah.columns:
        df_sekolah['id_sekolah'] = standard_id(df_sekolah['id_sekolah'])
    if 'id_sekolah' in df_ruang.columns:
        df_ruang['id_sekolah'] = standard_id(df_ruang['id_sekolah'])
    if 'id_ruang' in df_ruang.columns:
        df_ruang['id_ruang'] = standard_id(df_ruang['id_ruang']) # Standardisasi id_ruang agar bertipe string
    if 'kode_tuo' in df_master.columns:
        df_master['kode_tuo'] = standard_id(df_master['kode_tuo'])
        
    if 'id_wilayah' in df_wilayah.columns:
        df_wilayah['id_wilayah'] = standard_id(df_wilayah['id_wilayah'])
    if 'id_wilayah' in df_sekolah.columns:
        df_sekolah['id_wilayah'] = standard_id(df_sekolah['id_wilayah'])
        
    if not df_ruang_tap.empty:
        df_ruang_tap['id_ruang'] = standard_id(df_ruang_tap['id_ruang'])
        df_ruang_tap['id_sekolah'] = standard_id(df_ruang_tap['id_sekolah'])
        df_ruang_tap['Tanggal'] = df_ruang_tap['Tanggal'].astype(str).str.strip()
    else:
        df_ruang_tap = pd.DataFrame(columns=['id_ruang', 'id_sekolah', 'Tanggal'])
        
    # Standardisasi format tgl_ujian master peserta
    df_master['tgl_ujian'] = pd.to_datetime(df_master['tgl_ujian'], errors='coerce').dt.strftime('%Y-%m-%d')
    df_master['tgl_ujian'] = df_master['tgl_ujian'].astype(str).str.strip()
    unique_dates = df_master['tgl_ujian'].dropna().unique()
    if len(unique_dates) == 0:
        unique_dates = ['2026-06-20']
        
    # 2. Hitung Kapasitas Maksimal per Sekolah dengan Aturan Sesi Dinamis (2 Sesi jika ada TAP, else 5 Sesi)
    if 'kapasitas' in df_ruang.columns:
        df_ruang['kapasitas'] = pd.to_numeric(df_ruang['kapasitas'], errors='coerce').fillna(0).astype(int)
        
        # Buat kombinasi semua ruang dan semua tanggal ujian yang aktif
        df_dates_grid = pd.DataFrame({'tgl_ujian': unique_dates})
        df_ruang_copy = df_ruang.copy()
        df_ruang_copy['key'] = 1
        df_dates_grid['key'] = 1
        df_ruang_dates = pd.merge(df_ruang_copy, df_dates_grid, on='key').drop('key', axis=1)
        
        # Bersihkan whitespace
        df_ruang_dates['id_ruang'] = df_ruang_dates['id_ruang'].astype(str).str.strip()
        df_ruang_dates['tgl_ujian'] = df_ruang_dates['tgl_ujian'].astype(str).str.strip()
        
        # Merge dengan data ruang TAP untuk melacak ruang yang terpakai TAP pada tanggal ujian tertentu
        df_ruang_tap_clean = df_ruang_tap[['id_ruang', 'Tanggal']].copy()
        df_ruang_tap_clean.columns = ['id_ruang', 'tap_tgl_ujian']
        df_ruang_tap_clean['id_ruang'] = df_ruang_tap_clean['id_ruang'].astype(str).str.strip()
        df_ruang_tap_clean['tap_tgl_ujian'] = df_ruang_tap_clean['tap_tgl_ujian'].astype(str).str.strip()
        df_ruang_tap_clean['is_room_tap'] = True
        
        # PERBAIKAN STRUKTURAL: Konversi pencocokan tanggal gabungan menjadi objek Datetime murni agar 100% akurat
        df_ruang_dates['tgl_ujian_dt'] = pd.to_datetime(df_ruang_dates['tgl_ujian'], errors='coerce')
        df_ruang_tap_clean['tap_tgl_ujian_dt'] = pd.to_datetime(df_ruang_tap_clean['tap_tgl_ujian'], errors='coerce')
        
        df_ruang_dates = pd.merge(
            df_ruang_dates,
            df_ruang_tap_clean,
            left_on=['id_ruang', 'tgl_ujian_dt'],
            right_on=['id_ruang', 'tap_tgl_ujian_dt'],
            how='left'
        )
        df_ruang_dates['is_room_tap'] = df_ruang_dates['is_room_tap'].fillna(False)
        
        # Aturan Sesi: 2 Sesi jika terpakai TAP/S2/Essay, else 5 Sesi
        df_ruang_dates['kapasitas_sesi'] = np.where(
            df_ruang_dates['is_room_tap'],
            df_ruang_dates['kapasitas'] * 2,
            df_ruang_dates['kapasitas'] * 5
        ).astype(int)
        
        # Group by school & date untuk total kapasitas total real-time
        df_capacity = df_ruang_dates.groupby(['id_sekolah', 'tgl_ujian'])['kapasitas_sesi'].sum().reset_index()
        df_capacity.columns = ['id_sekolah', 'tgl_ujian', 'kapasitas_total']
        
        # Simpan metadata kapasitas 5 sesi standar untuk rincian detail (Tab 2)
        df_ruang['kapasitas_5_sesi'] = df_ruang['kapasitas'] * 5
    else:
        df_ruang['kapasitas'] = 0
        df_ruang['kapasitas_5_sesi'] = 0
        df_capacity = pd.DataFrame(columns=['id_sekolah', 'tgl_ujian', 'kapasitas_total'])
    
    # 3. Hubungkan data sekolah dengan data wilayah (Kabupaten)
    df_sch_wil = pd.merge(df_sekolah, df_wilayah, on='id_wilayah', how='left')
    
    # 4. Hubungkan data master peserta dengan sekolah & wilayah
    df_merged = pd.merge(df_master, df_sch_wil, left_on='kode_tuo', right_on='id_sekolah', how='left')
    
    # Bersihkan keys sebelum merge kapasitas
    if not df_capacity.empty:
        df_capacity['id_sekolah'] = df_capacity['id_sekolah'].astype(str).str.strip()
        df_capacity['tgl_ujian'] = df_capacity['tgl_ujian'].astype(str).str.strip()
        
    df_merged['kode_tuo'] = df_merged['kode_tuo'].astype(str).str.strip()
    df_merged['tgl_ujian'] = df_merged['tgl_ujian'].astype(str).str.strip()
    
    # 5. Gabungkan dengan kapasitas total sekolah per tanggal ujian yang dinamis
    df_merged = pd.merge(
        df_merged,
        df_capacity,
        left_on=['kode_tuo', 'tgl_ujian'],
        right_on=['id_sekolah', 'tgl_ujian'],
        how='left'
    )
    
    # 6. Tandai apakah sekolah tersebut memiliki Ujian TAP / S2 / Essay pada tanggal tertentu (is_tap)
    df_ruang_tap_tag = df_ruang_tap.copy()
    df_ruang_tap_tag['is_tap'] = True
    df_sch_tap_dates = df_ruang_tap_tag[['id_sekolah', 'Tanggal', 'is_tap']].drop_duplicates()
    df_sch_tap_dates.columns = ['tap_id_sekolah', 'tap_Tanggal', 'is_tap']
    df_sch_tap_dates['tap_id_sekolah'] = df_sch_tap_dates['tap_id_sekolah'].astype(str).str.strip()
    df_sch_tap_dates['tap_Tanggal'] = df_sch_tap_dates['tap_Tanggal'].astype(str).str.strip()
    
    # Gunakan robust datetime matching untuk highlight warna biru di tabel pivot
    df_merged['tgl_ujian_dt'] = pd.to_datetime(df_merged['tgl_ujian'], errors='coerce')
    df_sch_tap_dates['tap_Tanggal_dt'] = pd.to_datetime(df_sch_tap_dates['tap_Tanggal'], errors='coerce')
    
    df_merged = pd.merge(
        df_merged,
        df_sch_tap_dates,
        left_on=['kode_tuo', 'tgl_ujian_dt'],
        right_on=['tap_id_sekolah', 'tap_Tanggal_dt'],
        how='left'
    )
    df_merged['is_tap'] = df_merged['is_tap'].fillna(False).astype(bool)
    
    # Bersihkan sisa merge redundan dan kolom datetime sementara
    for col_to_drop in ['tap_id_sekolah', 'tap_Tanggal_dt', 'tgl_ujian_dt', 'id_sekolah_y']:
        if col_to_drop in df_merged.columns:
            df_merged = df_merged.drop(col_to_drop, axis=1)
            
    if 'id_sekolah_x' in df_merged.columns:
        df_merged = df_merged.rename(columns={'id_sekolah_x': 'id_sekolah'})
        
    df_merged['kapasitas_total'] = df_merged['kapasitas_total'].fillna(0).astype(int)
    df_merged['nama_sekolah'] = df_merged['nama_sekolah'].fillna(df_merged['nama_tuo'])
    
    # 7. Hitung Total Peserta Berdasarkan Kolom yang Tersedia
    participant_cols = ['jml_s2', 'jml_tap', 'jml_s1_objektif', 'jml_s1_uraian']
    for col in participant_cols:
        if col in df_merged.columns:
            df_merged[col] = pd.to_numeric(df_merged[col], errors='coerce').fillna(0).astype(int)
        else:
            df_merged[col] = 0
            
    # Sesuai request: jumlah peserta = jml_s2 + jml_tap + jml_s1_objektif + jml_s1_uraian
    df_merged['total_peserta'] = (
        df_merged['jml_s2'] + 
        df_merged['jml_tap'] + 
        df_merged['jml_s1_objektif'] + 
        df_merged['jml_s1_uraian']
    ).astype(int)
    df_merged['sisa_kuota'] = (df_merged['kapasitas_total'] - df_merged['total_peserta']).astype(int)
    
    # Persentase keterisian lokasi
    df_merged['persentase_keterisian'] = np.where(
        df_merged['kapasitas_total'] > 0,
        (df_merged['total_peserta'] / df_merged['kapasitas_total'] * 100).round(1),
        0.0
    )
    
    return df_merged, df_ruang, df_sekolah


# --- PINDAH KE ATAS & OTOMATIS TERSEMBUNYI (EXPANDER) ---
with st.expander("⚙️ Konfigurasi Basis Data & Unggah File Master (Khusus Admin)", expanded=False):
    st.markdown("Bagian ini digunakan untuk melakukan update master data. Anda dapat memperbarui berkas secara satuan atau bersamaan.")
    
    # Periksa status database SQLite
    db_siap = cek_db_tersimpan()
    loc_master, loc_wilayah, loc_sekolah, loc_ruang, loc_ruang_tap = cari_file_lokal()
    file_lokal_siap = loc_master and loc_wilayah and loc_sekolah and loc_ruang

    # Tampilkan Ringkasan Baris SQLite Aktif (Jika siap)
    if db_siap:
        try:
            conn = sqlite3.connect(DB_NAME)
            n_peserta = conn.execute("SELECT COUNT(*) FROM master_peserta").fetchone()[0]
            n_wilayah = conn.execute("SELECT COUNT(*) FROM master_wilayah").fetchone()[0]
            n_sekolah = conn.execute("SELECT COUNT(*) FROM master_sekolah").fetchone()[0]
            n_ruang = conn.execute("SELECT COUNT(*) FROM master_ruang").fetchone()[0]
            n_ruang_tap = conn.execute("SELECT COUNT(*) FROM master_ruang_tap").fetchone()[0]
            conn.close()
            
            st.markdown("### 📊 Ringkasan Baris Data SQLite Aktif:")
            col_stat1, col_stat2, col_stat3, col_stat4, col_stat5 = st.columns(5)
            col_stat1.metric("Data Peserta", f"{n_peserta} baris")
            col_stat2.metric("Master Wilayah", f"{n_wilayah} baris")
            col_stat3.metric("Master Sekolah", f"{n_sekolah} baris")
            col_stat4.metric("Master Ruang", f"{n_ruang} baris")
            col_stat5.metric("Master Ruang TAP", f"{n_ruang_tap} baris")
            st.markdown("---")
        except Exception:
            pass

    col_adm1, col_adm2 = st.columns(2)
    
    with col_adm1:
        st.subheader("Opsi 1: Muat Ulang File Lokal")
        if file_lokal_siap:
            st.success("📂 File CSV/Excel terdeteksi di direktori komputer!")
            if st.button("🔄 Sinkronisasi Semua File Lokal ke SQLite"):
                try:
                    df_m = baca_berkas(loc_master)
                    df_w = baca_berkas(loc_wilayah)
                    df_s = baca_berkas(loc_sekolah)
                    df_r = baca_berkas(loc_ruang)
                    df_rt = baca_berkas(loc_ruang_tap) if loc_ruang_tap else pd.DataFrame(columns=['id_ruang', 'id_sekolah', 'Tanggal'])
                    simpan_ke_sqlite(df_m, df_w, df_s, df_r, df_rt)
                    st.success("💾 Sukses sinkronisasi semua file lokal ke SQLite!")
                    st.rerun()
                except Exception as ex:
                    st.error(f"Gagal memindahkan data ke SQLite: {ex}")
        else:
            st.info("Tidak terdeteksi paket file lokal lengkap di direktori.")

    with col_adm2:
        st.subheader("Opsi 2: Unggah Manual (Satuan atau Sekaligus)")
        
        # File uploaders
        uploaded_master = st.file_uploader("1. Unggah Data Peserta Baru (Mendukung .xlsx, .xls, .csv)", type=["csv", "xlsx", "xls"])
        uploaded_wilayah = st.file_uploader("2. File Master Wilayah (CSV/Excel)", type=["csv", "xlsx", "xls"])
        uploaded_sekolah = st.file_uploader("3. File Master Sekolah (CSV/Excel)", type=["csv", "xlsx", "xls"])
        uploaded_ruang = st.file_uploader("4. File Master Ruang (CSV/Excel)", type=["csv", "xlsx", "xls"])
        uploaded_ruang_tap = st.file_uploader("5. File Ruang TAP (CSV/Excel)", type=["csv", "xlsx", "xls"])
        
        # Deteksi berkas apa saja yang siap diunggah
        files_to_update = {}
        if uploaded_master:
            files_to_update['master_peserta'] = uploaded_master
        if uploaded_wilayah:
            files_to_update['master_wilayah'] = uploaded_wilayah
        if uploaded_sekolah:
            files_to_update['master_sekolah'] = uploaded_sekolah
        if uploaded_ruang:
            files_to_update['master_ruang'] = uploaded_ruang
        if uploaded_ruang_tap:
            files_to_update['master_ruang_tap'] = uploaded_ruang_tap
            
        if files_to_update:
            st.markdown("### 📋 Berkas yang Siap Diperbarui:")
            for name, file in files_to_update.items():
                st.markdown(f"- **{name.replace('master_', '').replace('_', ' ').title()}**: `{file.name}` ✅")
                
            # Validasi inisialisasi awal
            if not db_siap and len(files_to_update) < 5:
                st.warning("⚠️ Basis data SQLite kosong. Untuk pembentukan pertama kali, mohon unggah kelima berkas master di atas secara lengkap sekaligus.")
            else:
                if st.button("💾 Simpan & Perbarui Berkas Terpilih ke SQLite"):
                    try:
                        conn = sqlite3.connect(DB_NAME)
                        
                        if uploaded_master:
                            df_m = baca_berkas(uploaded_master)
                            df_m.to_sql('master_peserta', conn, if_exists='replace', index=False)
                            # Rekam waktu jam update
                            waktu_sekarang = datetime.datetime.now().strftime("%d-%m-%Y %H:%M:%S WIB")
                            simpan_metadata("last_peserta_update", waktu_sekarang)
                            st.success("✅ Berhasil menyimpan Data Peserta!")
                            
                        if uploaded_wilayah:
                            df_w = baca_berkas(uploaded_wilayah)
                            df_w.to_sql('master_wilayah', conn, if_exists='replace', index=False)
                            st.success("✅ Berhasil menyimpan Master Wilayah!")
                            
                        if uploaded_sekolah:
                            df_s = baca_berkas(uploaded_sekolah)
                            df_s.to_sql('master_sekolah', conn, if_exists='replace', index=False)
                            st.success("✅ Berhasil menyimpan Master Sekolah!")
                            
                        if uploaded_ruang:
                            df_r = baca_berkas(uploaded_ruang)
                            df_r.to_sql('master_ruang', conn, if_exists='replace', index=False)
                            st.success("✅ Berhasil menyimpan Master Ruang!")
                            
                        if uploaded_ruang_tap:
                            df_rt = baca_berkas(uploaded_ruang_tap)
                            df_rt_clean = clean_ruang_tap(df_rt)
                            df_rt_clean.to_sql('master_ruang_tap', conn, if_exists='replace', index=False)
                            st.success("✅ Berhasil menyimpan Master Ruang TAP!")
                            
                        conn.commit()
                        conn.close()
                        
                        st.success("🔥 Basis data SQLite sukses diperbarui!")
                        st.cache_data.clear()
                        st.rerun()
                    except Exception as ex:
                        st.error(f"Gagal memproses berkas: {ex}")


# Cek kesiapan pemrosesan data utama
data_siap_proses = cek_db_tersimpan()

# --- FILTER DI SIDEBAR (HANYA UNTUK FILTER) ---
st.sidebar.header("🔍 Filter Wilayah & Jadwal")

if data_siap_proses:
    try:
        # Load data langsung dari SQLite terpadu
        df_final, df_ruang_raw, df_sekolah_raw = load_and_process_data_from_db()
        
        # Filter Kabupaten
        list_kabupaten = sorted(df_final['nama_kabupaten'].unique())
        pilih_kabupaten = st.sidebar.selectbox("Pilih Kabupaten / Kota:", ["SEMUA KABUPATEN"] + list_kabupaten)
        
        # Filter Tanggal Ujian
        list_tanggal = sorted(df_final['tgl_ujian'].unique())
        pilih_tanggal = st.sidebar.multiselect("Pilih Tanggal Ujian:", list_tanggal, default=list_tanggal)
        
        # Filter Status
        pilih_status = st.sidebar.radio(
            "Filter Status Ketersediaan:",
            ["Tampilkan Semua", "Hanya yang Penuh / Melebihi Kuota", "Hanya yang Aman (Tersisa)"]
        )
        
        # Terapkan Filter
        df_filtered = df_final.copy()
        if pilih_kabupaten != "SEMUA KABUPATEN":
            df_filtered = df_filtered[df_filtered['nama_kabupaten'] == pilih_kabupaten]
            
        if pilih_tanggal:
            df_filtered = df_filtered[df_filtered['tgl_ujian'].isin(pilih_tanggal)]
            
        if pilih_status == "Hanya yang Penuh / Melebihi Kuota":
            df_filtered = df_filtered[df_filtered['sisa_kuota'] <= 0]
        elif pilih_status == "Hanya yang Aman (Tersisa)":
            df_filtered = df_filtered[df_filtered['sisa_kuota'] > 0]
            
        # --- PANEL METRIK UTAMA (KPI) DENGAN 5 KOLOM ---
        col1, col2, col3, col4, col5 = st.columns(5)
        
        total_peserta = df_filtered['total_peserta'].sum()
        # Hitung kapasitas berdasarkan keunikan kombinasi sekolah & tanggal ujian yang ter-filter
        total_daya_tampung = df_filtered.drop_duplicates(subset=['kode_tuo', 'tgl_ujian'])['kapasitas_total'].sum()
        sisa_total_kuota = total_daya_tampung - total_peserta
        sekolah_aktif = df_filtered['kode_tuo'].nunique()
        hari_digunakan = df_filtered['tgl_ujian'].nunique()
        
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
                <p style="color: #6c757d; margin-bottom: 2px; font-size:14px;">Daya Tampung Total (Sesi Aktif)</p>
                <h3 style="margin:0; font-size:24px;">{int(total_daya_tampung)} Kursi</h3>
            </div>
            """, unsafe_allow_html=True)
            
        with col3:
            warna_sisa = "#28a745" if sisa_total_kuota >= 0 else "#dc3545"
            st.markdown(f"""
            <div class="metric-card" style="border-left-color: {warna_sisa};">
                <p style="color: #6c757d; margin-bottom: 2px; font-size:14px;">Sisa Kuota Gabungan</p>
                <h3 style="margin:0; font-size:24px;">{int(sisa_total_kuota)} Kursi</h3>
            </div>
            """, unsafe_allow_html=True)
            
        with col4:
            st.markdown(f"""
            <div class="metric-card" style="border-left-color: #ffc107;">
                <p style="color: #6c757d; margin-bottom: 2px; font-size:14px;">Lokasi Ujian Aktif</p>
                <h3 style="margin:0; font-size:24px;">{sekolah_aktif} Sekolah</h3>
            </div>
            """, unsafe_allow_html=True)

        with col5:
            st.markdown(f"""
            <div class="metric-card" style="border-left-color: #17a2b8;">
                <p style="color: #6c757d; margin-bottom: 2px; font-size:14px;">Hari yang Digunakan</p>
                <h3 style="margin:0; font-size:24px;">{int(hari_digunakan)} Hari</h3>
            </div>
            """, unsafe_allow_html=True)
            
        # --- TABEL UTAMA PIVOT DENGAN CONDITIONAL COLORING ---
        st.subheader("🗓️ Visualisasi Sisa Kuota Sekolah per Tanggal (Kapasitas Dinamis)")
        st.markdown(
            "Tabel pivot di bawah menampilkan **sisa kapasitas** untuk masing-masing sekolah. "
            "Kapasitas dihitung berdasarkan **Daya Tampung Ruang x 2 Sesi** jika terdapat ujian TAP/S2 pada tanggal tersebut, dan **x 5 Sesi** jika ujian reguler."
        )
        
        opsi_tampilan = st.radio(
            "Pilih Model Tampilan Data:",
            ["Tampilkan Sisa Kuota (Leb Jelas)", "Tampilkan Rasio Keterisian (%)", "Tampilkan Format (Peserta / Kapasitas Maks)"],
            horizontal=True
        )
        
        if df_filtered.empty:
            st.info("⚠️ Tidak ada data yang sesuai dengan kombinasi filter Anda saat ini.")
        else:
            # 1. Selalu buat Pivot Matrix Rasio Persentase Keterisian sebagai acuan pemetaan warna yang presisi
            df_pivot_pct = df_filtered.pivot_table(
                index=['nama_sekolah'],
                columns='tgl_ujian',
                values='persentase_keterisian',
                aggfunc='mean'
            )
            
            # 2. Pivot Matrix is_tap untuk identifikasi letak ujian TAP/S2/Essay
            df_pivot_tap = df_filtered.pivot_table(
                index=['nama_sekolah'],
                columns='tgl_ujian',
                values='is_tap',
                aggfunc='max'
            ).fillna(False).astype(bool)

            # Fungsi pembuat warna berbasis matriks persentase keterisian terpadu (menyertakan deteksi TAP)
            def style_matriks_by_percentage(val_df):
                out = pd.DataFrame('', index=val_df.index, columns=val_df.columns)
                for r in val_df.index:
                    for c in val_df.columns:
                        pct = df_pivot_pct.loc[r, c] if r in df_pivot_pct.index and c in df_pivot_pct.columns else np.nan
                        is_tap_val = df_pivot_tap.loc[r, c] if r in df_pivot_tap.index and c in df_pivot_tap.columns else False
                        out.loc[r, c] = get_color_style(pct, is_tap=is_tap_val)
                return out

            if opsi_tampilan == "Tampilkan Sisa Kuota (Leb Jelas)":
                df_pivot = df_filtered.pivot_table(
                    index=['nama_sekolah'],
                    columns='tgl_ujian',
                    values='sisa_kuota',
                    aggfunc='sum'
                )
                styled_df = df_pivot.style.apply(style_matriks_by_percentage, axis=None).format("{:.0f}", na_rep="-")
                st.dataframe(styled_df, use_container_width=True, height=450)
                
            elif opsi_tampilan == "Tampilkan Rasio Keterisian (%)":
                df_pivot = df_filtered.pivot_table(
                    index=['nama_sekolah'],
                    columns='tgl_ujian',
                    values='persentase_keterisian',
                    aggfunc='mean'
                )
                styled_df = df_pivot.style.apply(style_matriks_by_percentage, axis=None).format("{:.1f}%", na_rep="-")
                st.dataframe(styled_df, use_container_width=True, height=450)
                
            else: # Format Peserta / Kapasitas
                df_pivot_peserta = df_filtered.pivot_table(index='nama_sekolah', columns='tgl_ujian', values='total_peserta', aggfunc='sum')
                df_pivot_kapasitas = df_filtered.pivot_table(index='nama_sekolah', columns='tgl_ujian', values='kapasitas_total', aggfunc='sum')
                
                df_gabung = df_pivot_peserta.copy()
                for col in df_pivot_peserta.columns:
                    # Ambil nilai bulat (integer) bebas desimal `.0` & hilangkan koma ribuan
                    peserta_str = df_pivot_peserta[col].fillna(0).astype(int).astype(str)
                    kapasitas_str = df_pivot_kapasitas[col].fillna(0).astype(int).astype(str)
                    
                    df_gabung[col] = peserta_str + " / " + kapasitas_str
                    df_gabung.loc[df_pivot_peserta[col].isna(), col] = "-"
                    
                styled_df = df_gabung.style.apply(style_matriks_by_percentage, axis=None)
                st.dataframe(styled_df, use_container_width=True, height=450)

        # --- PANEL PANEL RINCIAN RUANGAN & EKSPOR DATA ---
        st.markdown("---")
        tab1, tab2, tab3 = st.tabs(["📊 Analisis Grafik", "📋 Struktur Kapasitas Ruang", "💾 Ekspor Laporan"])
        
        with tab1:
            st.subheader("Grafik Analisis Alokasi per Sekolah (Kapasitas Total)")
            if not df_filtered.empty:
                df_chart = df_filtered.groupby('nama_sekolah')[['total_peserta', 'kapasitas_total']].sum().reset_index()
                
                fig = go.Figure()
                fig.add_trace(go.Bar(
                    x=df_chart['nama_sekolah'],
                    y=df_chart['kapasitas_total'],
                    name='Daya Tampung Maksimal',
                    marker_color='#6c757d',
                    opacity=0.6
                ))
                fig.add_trace(go.Bar(
                    x=df_chart['nama_sekolah'],
                    y=df_chart['total_peserta'],
                    name='Total Peserta Terdaftar',
                    marker_color='#007bff'
                ))
                
                fig.update_layout(
                    barmode='group',
                    xaxis_title='Nama Sekolah',
                    yaxis_title='Jumlah Kursi',
                    legend_title='Keterangan',
                    template='plotly_white',
                    height=500
                )
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.info("Pilih kabupaten atau tanggal terlebih dahulu.")
                
        with tab2:
            st.subheader("Rincian Kapasitas Ruangan & Skema Sesi")
            st.markdown("Detail kapasitas dasar per ruangan serta total kapasitas kumulatif standar (5 Sesi).")
            
            sekolah_terpilih = df_filtered['kode_tuo'].unique()
            df_ruang_filtered = df_ruang_raw[df_ruang_raw['id_sekolah'].astype(str).str.replace(r'\.0$', '', regex=True).isin(sekolah_terpilih)].copy()
            
            if not df_ruang_filtered.empty:
                clean_cols = [c for c in df_ruang_filtered.columns if not c.startswith('Unnamed') and c.strip() != '']
                df_show = df_ruang_filtered[clean_cols].rename(columns={
                    'nama_ruang': 'Nama Ruang / Lokasi',
                    'kapasitas': 'Kapasitas Dasar (1 Sesi)',
                    'kapasitas_5_sesi': 'Kapasitas Total (5 Sesi)',
                    'Ruang': 'Nomor Ruangan'
                })
                # Urutkan kolom agar enak dibaca
                kolom_tampil = ['id_sekolah', 'Nama Ruang / Lokasi', 'Nomor Ruangan', 'Kapasitas Dasar (1 Sesi)', 'Kapasitas Total (5 Sesi)']
                kolom_tersedia = [col for col in kolom_tampil if col in df_show.columns]
                
                # Menghilangkan koma ribuan pada visualisasi rincian data table ruang
                st.dataframe(
                    df_show[kolom_tersedia], 
                    use_container_width=True, 
                    hide_index=True,
                    column_config={
                        "id_sekolah": st.column_config.TextColumn("ID Sekolah"),
                        "Nomor Ruangan": st.column_config.TextColumn("Nomor Ruangan"),
                        "Kapasitas Dasar (1 Sesi)": st.column_config.NumberColumn("Kapasitas Dasar (1 Sesi)", format="%d"),
                        "Kapasitas Total (5 Sesi)": st.column_config.NumberColumn("Kapasitas Total (5 Sesi)", format="%d")
                    }
                )
            else:
                st.info("Tidak ada rincian ruang untuk sekolah aktif saat ini.")
                
        with tab3:
            st.subheader("Ekspor Hasil Analisis")
            st.markdown("Unduh hasil filter visualisasi saat ini ke dalam berkas CSV untuk kemudahan laporan eksternal.")
            
            df_export = df_filtered[['nama_kabupaten', 'nama_sekolah', 'tgl_ujian', 'total_peserta', 'kapasitas_total', 'sisa_kuota', 'persentase_keterisian']].copy()
            df_export.columns = ['Kabupaten', 'Nama Sekolah', 'Tanggal Ujian', 'Total Peserta', 'Kapasitas Maksimal', 'Sisa Kuota', 'Persentase Keterisian (%)']
            
            csv_data = df_export.to_csv(index=False).encode('utf-8')
            
            st.download_button(
                label="📥 Unduh Laporan (.CSV)",
                data=csv_data,
                file_name=f"Laporan_Alokasi_UO_{pilih_kabupaten}.csv",
                mime="text/csv",
                use_container_width=True
            )
            
    except Exception as e:
        st.error(f"Terjadi kesalahan saat memproses data: {e}")
        st.info("Saran: Pastikan struktur kolom CSV Anda telah sesuai atau gunakan opsi Upload Manual di Sidebar.")
        st.code(str(e))
else:
    # Halaman Selamat Datang / Panduan Upload
    st.info("👋 Selamat datang di Dashboard Alokasi Kuota Ujian Online!")
    st.markdown("""
    ### Langkah Memulai Aplikasi:
    Aplikasi ini membutuhkan data master yang tersimpan di dalam **SQLite**. Silakan buka panel **"⚙️ Konfigurasi Basis Data & Unggah File Master"** di atas untuk memuat data pertama kali:
    
    1. **Upload Berkas Excel/CSV Baru**: Anda dapat langsung mengunggah berkas data peserta baru (bisa dalam format Excel **.xlsx**) serta data pelengkap lainnya pada panel di atas.
    2. **Sinkronisasi Berkas Lokal**: Jika Anda menjalankan aplikasi secara lokal dan berkas master berada dalam folder yang sama, Anda cukup klik tombol **"Sinkronisasi Semua File Lokal ke SQLite"** di atas.
    
    *Formula daya tampung ruangan saat ini menggunakan skema: **Kapasitas Ruang x Sesi Ujian** (Sesi dinamis berdasarkan skema TAP).*
    """)
    st.image("https://images.unsplash.com/photo-1551288049-bebda4e38f71?auto=format&fit=crop&w=1200&q=80", caption="Sistem Visualisasi Manajemen Alokasi Ruang & Kuota")