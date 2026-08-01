import streamlit as st
import pandas as pd
import plotly.express as px
import base64
import os
from PIL import Image, ImageOps
from datetime import datetime, timedelta

st.set_page_config(
    page_title="Portal Resmi RW 14 Griya Permata Raya",
    layout="wide",
    page_icon="🏛️",
    initial_sidebar_state="expanded"
)

st.markdown("""
<style>
.stApp { background: linear-gradient(135deg, #f0f4f8 0%, #e2e8f0 100%); background-attachment: fixed; }
div[data-testid="metric-container"] { background: rgba(255, 255, 255, 0.85); backdrop-filter: blur(10px); padding: 15px; border-radius: 12px; box-shadow: 0px 6px 15px rgba(0,0,0,0.04); border-left: 6px solid #0D47A1; }
div[data-testid="stMetricValue"], div[data-testid="stMetricValue"] > div { font-size: 38px !important; color: #0D47A1 !important; font-weight: 900 !important; }
div[data-testid="stMetricLabel"] p, div[data-testid="stMetricLabel"] > div, div[data-testid="stMetricLabel"] { font-size: 16px !important; font-weight: bold !important; color: #334155 !important; }
h3 { font-size: 24px !important; color: #0D47A1; font-weight: 800; }
button[data-baseweb="tab"] > div[data-testid="stMarkdownContainer"] > p { font-size: 13px !important; font-weight: bold; }
.stPlotlyChart { background: rgba(255, 255, 255, 0.9); border-radius: 12px; box-shadow: 0px 6px 15px rgba(0,0,0,0.06); padding: 15px; border: 1px solid #cbd5e1; }
</style>
""", unsafe_allow_html=True)

def ambil_logo_lokal(nama_file):
    if os.path.exists(nama_file):
        with open(nama_file, "rb") as img_file:
            encoded_string = base64.b64encode(img_file.read()).decode()
        ext = nama_file.split('.')[-1].lower()
        mime = 'jpeg' if ext in ['jpg', 'jpeg'] else 'png'
        return f"data:image/{mime};base64,{encoded_string}"
    return "https://cdn-icons-png.flaticon.com/512/3135/3135673.png"

sumber_logo = ambil_logo_lokal("logo rw.png")

def muat_dan_seragamkan_foto(path_file, ukuran=(300, 400)):
    try:
        img = Image.open(path_file)
        img = ImageOps.exif_transpose(img)
        img = ImageOps.fit(img, ukuran, method=Image.Resampling.LANCZOS, centering=(0.5, 0.5))
        return img
    except Exception:
        return path_file

def urutkan_data_warga(df):
    if df.empty:
        return df
    df.columns = df.columns.str.strip().str.upper()
    if "RT" in df.columns:
        df["RT_NUM"] = pd.to_numeric(df["RT"].astype(str).str.replace(r'[^0-9]', '', regex=True), errors="coerce").fillna(99)
    else:
        df["RT_NUM"] = 99
        
    kolom_nama = None
    for k in ["NAMA", "NAMA LENGKAP", "NAMA WARGA"]:
        if k in df.columns:
            kolom_nama = k
            break
            
    if kolom_nama:
        df["NAMA_STR"] = df[kolom_nama].astype(str).str.strip().str.upper()
        df = df.sort_values(by=["RT_NUM", "NAMA_STR"], ascending=[True, True]).reset_index(drop=True)
        df = df.drop(columns=["RT_NUM", "NAMA_STR"], errors="ignore")
    else:
        df = df.sort_values(by=["RT_NUM"], ascending=[True]).reset_index(drop=True)
        df = df.drop(columns=["RT_NUM"], errors="ignore")
    return df

@st.cache_data
def load_data():
    if os.path.exists("datawarga.xlsx"):
        df = pd.read_excel("datawarga.xlsx")
        df = urutkan_data_warga(df)
        if "UMUR" in df.columns and "USIA" not in df.columns: df.rename(columns={"UMUR": "USIA"}, inplace=True)
        if "STATUS" in df.columns and "STATUS PERKAWINAN" not in df.columns: df.rename(columns={"STATUS": "STATUS PERKAWINAN"}, inplace=True)
        if "STATUS NIKAH" in df.columns and "STATUS PERKAWINAN" not in df.columns: df.rename(columns={"STATUS NIKAH": "STATUS PERKAWINAN"}, inplace=True)
        if "NO KK" in df.columns and "NO. KK" not in df.columns: df.rename(columns={"NO KK": "NO. KK"}, inplace=True)
        if "PENDIDIKAN TERAKHIR" in df.columns and "PENDIDIKAN" not in df.columns: df.rename(columns={"PENDIDIKAN TERAKHIR": "PENDIDIKAN"}, inplace=True)
        
        if "STATUS PENDUDUK" not in df.columns:
            df["STATUS PENDUDUK"] = "Tetap"
        else:
            df["STATUS PENDUDUK"] = df["STATUS PENDUDUK"].fillna("Tetap").astype(str).str.strip().str.title()
            df["STATUS PENDUDUK"] = df["STATUS PENDUDUK"].apply(lambda x: x if x in ["Tetap", "Musiman"] else "Tetap")
        return df
    return pd.DataFrame()

@st.cache_data
def load_kas():
    if os.path.exists("datakas.xlsx"):
        df_kas = pd.read_excel("datakas.xlsx")
        df_kas.columns = df_kas.columns.str.strip().str.upper()
        if "TANGGAL" not in df_kas.columns: df_kas["TANGGAL"] = ""
        if "KETERANGAN" not in df_kas.columns: df_kas["KETERANGAN"] = ""
        
        kolom_masuk = [c for c in df_kas.columns if "PEMASUKAN" in c or "MASUK" in c]
        kolom_keluar = [c for c in df_kas.columns if "PENGELUARAN" in c or "KELUAR" in c]
        m_col = kolom_masuk[0] if kolom_masuk else df_kas.columns[2] if len(df_kas.columns) > 2 else "PEMASUKAN"
        k_col = kolom_keluar[0] if kolom_keluar else df_kas.columns[3] if len(df_kas.columns) > 3 else "PENGELUARAN"
        
        if m_col not in df_kas.columns: df_kas[m_col] = 0
        if k_col not in df_kas.columns: df_kas[k_col] = 0
        
        df_kas["TANGGAL"] = df_kas["TANGGAL"].fillna("").astype(str).str.replace(r'\.0$', '', regex=True)
        df_kas["KETERANGAN"] = df_kas["KETERANGAN"].fillna("").astype(str)
        df_kas[m_col] = pd.to_numeric(df_kas[m_col].astype(str).str.replace(r'[^0-9.-]', '', regex=True), errors="coerce").fillna(0)
        df_kas[k_col] = pd.to_numeric(df_kas[k_col].astype(str).str.replace(r'[^0-9.-]', '', regex=True), errors="coerce").fillna(0)
        df_kas["SALDO"] = (df_kas[m_col] - df_kas[k_col]).cumsum()
        
        df_kas = df_kas[["TANGGAL", "KETERANGAN", m_col, k_col, "SALDO"]]
        df_kas.columns = ["TANGGAL", "KETERANGAN", "PEMASUKAN", "PENGELUARAN", "SALDO"]
        return df_kas
    return pd.DataFrame(columns=["TANGGAL", "KETERANGAN", "PEMASUKAN", "PENGELUARAN", "SALDO"])

@st.cache_data
def load_kas_pemakaman():
    if os.path.exists("datakaspemakaman.xlsx"):
        df_kp = pd.read_excel("datakaspemakaman.xlsx")
        df_kp.columns = df_kp.columns.str.strip().str.upper()
        if "TANGGAL" not in df_kp.columns: df_kp["TANGGAL"] = ""
        if "KETERANGAN" not in df_kp.columns: df_kp["KETERANGAN"] = ""
        
        kolom_masuk_kp = [c for c in df_kp.columns if "PEMASUKAN" in c or "MASUK" in c]
        kolom_keluar_kp = [c for c in df_kp.columns if "PENGELUARAN" in c or "KELUAR" in c]
        m_col = kolom_masuk_kp[0] if kolom_masuk_kp else df_kp.columns[2] if len(df_kp.columns) > 2 else "PEMASUKAN"
        k_col = kolom_keluar_kp[0] if kolom_keluar_kp else df_kp.columns[3] if len(df_kp.columns) > 3 else "PENGELUARAN"
        
        if m_col not in df_kp.columns: df_kp[m_col] = 0
        if k_col not in df_kp.columns: df_kp[k_col] = 0
        
        df_kp["TANGGAL"] = df_kp["TANGGAL"].fillna("").astype(str).str.replace(r'\.0$', '', regex=True)
        df_kp["KETERANGAN"] = df_kp["KETERANGAN"].fillna("").astype(str)
        df_kp[m_col] = pd.to_numeric(df_kp[m_col].astype(str).str.replace(r'[^0-9.-]', '', regex=True), errors="coerce").fillna(0)
        df_kp[k_col] = pd.to_numeric(df_kp[k_col].astype(str).str.replace(r'[^0-9.-]', '', regex=True), errors="coerce").fillna(0)
        df_kp["SALDO"] = (df_kp[m_col] - df_kp[k_col]).cumsum()
        
        df_kp = df_kp[["TANGGAL", "KETERANGAN", m_col, k_col, "SALDO"]]
        df_kp.columns = ["TANGGAL", "KETERANGAN", "PEMASUKAN", "PENGELUARAN", "SALDO"]
        return df_kp
    return pd.DataFrame(columns=["TANGGAL", "KETERANGAN", "PEMASUKAN", "PENGELUARAN", "SALDO"])

@st.cache_data
def load_info():
    if os.path.exists("datainfo.xlsx"):
        df_info = pd.read_excel("datainfo.xlsx")
        df_info.columns = df_info.columns.str.strip().str.upper()
        return df_info
    return pd.DataFrame(columns=["TANGGAL", "JUDUL", "ISI / KATEGORI"])

@st.cache_data
def load_galeri_meta():
    if os.path.exists("datagaleri.xlsx"):
        df_g = pd.read_excel("datagaleri.xlsx")
        df_g.columns = df_g.columns.str.strip().str.upper()
        return df_g
    return pd.DataFrame(columns=["NAMA_FILE", "HARI", "TANGGAL", "BULAN", "TAHUN", "KETERANGAN"])

@st.cache_data
def load_saran():
    if os.path.exists("datasaran.xlsx"):
        df_s = pd.read_excel("datasaran.xlsx")
        df_s.columns = df_s.columns.str.strip().str.upper()
        return df_s
    return pd.DataFrame(columns=["WAKTU", "PENGIRIM", "JABATAN", "SARAN_PENDAPAT"])

@st.cache_data
def load_struktur():
    data_resmi = {
        "JABATAN / SEKSI": [
            "Ketua RW 014", "PKK & Posyandu", "Sekretaris", "Bendahara", 
            "Keamanan & Ketertiban", "Pembangunan & Lingkungan", "Olahraga", 
            "Sosial & Pemakaman", "Seni Budaya & Pemuda"
        ],
        "NAMA PENGURUS": [
            "Triyadi Sucipto", "Tim PKK / Posyandu RW 014", "Irvan Permana", "Aan Toni Fauyi", 
            "Dedi (RT 04), Uus (RT 04), Ali (RT 03), Tiktik (RT 07)", 
            "E. Rustandi (RT 06), Nahnu (RT 07), Dahlan (RT 03), Sugiyanto (RT 01), Mulyono (RT 05)", 
            "Mulyana (RT 05), Ateng (RT 03), Fajar (RT 01), Kris (RT 04), Apeng (RT 02), Mulyadi (RT 06)", 
            "Ust. Nanang (RT 03), E. Rustandi (RT 06), Ust. Juhendi (RT 07), Shulton (RT 04), Edi (RT 05), Baryanto (RT 01)", 
            "Uwa Tia (RT 06), Ridwan S (RT 01), Hary (RT 07)"
        ],
        "JOB DESCRIPTION (URAIAN TUGAS)": [
            "Memimpin, mengkoordinasikan, dan mengendalikan seluruh kegiatan penyelenggaraan rukun warga serta membina kerukunan.",
            "Menggerakkan partisipasi kaum ibu dalam bidang kesejahteraan keluarga, kesehatan anak, balita, dan lansia.",
            "Mengelola administrasi kesekretariatan, surat-menyurat, pendataan kependudukan, dan dokumentasi notulen.",
            "Bertanggung jawab penuh terhadap pengelolaan keuangan, pencatatan kas masuk/keluar, dan transparansi dana.",
            "Menjaga keamanan lingkungan perumahan, mengkoordinasikan siskamling, dan mengantisipasi gangguan ketertiban.",
            "Mengelola kebersihan fasum/fassos, merencanakan pemeliharaan infrastruktur, dan pelestarian lingkungan hijau.",
            "Mengembangkan minat dan bakat warga di bidang olahraga serta mempererat kebersamaan melalui turnamen.",
            "Mengurus pelayanan sosial kemasyarakatan, penanganan musibah warga, serta koordinasi proses pemakaman.",
            "Mengembangkan potensi seni budaya lokal serta merangkul karang taruna/pemuda dalam kegiatan positif."
        ],
        "BENTUK PROGRAM / KEGIATAN": [
            "• Koordinasi rutin RT\n• Musyawarah warga",
            "• Posyandu bulanan\n• Penyuluhan kesehatan & gizi\n• Dasawisma",
            "• Pendataan warga\n• Arsip surat & notulen rapat",
            "• Pengelolaan iuran & kas\n• Laporan keuangan bulanan",
            "• Pos ronda / siskamling\n• Pengawasan tamu\n• Penanganan darurat",
            "• Kerja bakti rutin\n• Pemeliharaan drainase",
            "• Senam sehat warga\n• Turnamen antar-RT",
            "• Pengelolaan dana sosial\n• Layanan takziah & pemakaman\n• Santunan duka",
            "• Peringatan Hari Besar (PHBN)\n• Kegiatan pemuda & seni"
        ]
    }
    return pd.DataFrame(data_resmi)

df = load_data()
df_kas = load_kas()
df_kas_pemakaman = load_kas_pemakaman()
df_info = load_info()
df_struktur = load_struktur()
df_galeri_meta = load_galeri_meta()
df_saran = load_saran()

st.sidebar.markdown("---")
waktu_sekarang = datetime.utcnow() + timedelta(hours=7)
hari_list = {"Monday": "Senin", "Tuesday": "Selasa", "Wednesday": "Rabu", "Thursday": "Kamis", "Friday": "Jumat", "Saturday": "Sabtu", "Sunday": "Minggu"}
bulan_list = {1: "Januari", 2: "Februari", 3: "Maret", 4: "April", 5: "Mei", 6: "Juni", 7: "Juli", 8: "Agustus", 9: "September", 10: "Oktober", 11: "November", 12: "Desember"}
tanggal_indo = f"{hari_list.get(waktu_sekarang.strftime('%A'), '')}, {waktu_sekarang.day} {bulan_list.get(waktu_sekarang.month, '')} {waktu_sekarang.year}"
jam_indo = waktu_sekarang.strftime("%H:%M:%S") + " WIB"

st.sidebar.markdown(f"""
<div style="background-color: #E3F2FD; padding: 10px; border-radius: 8px; text-align: center; border: 1px solid #90CAF9; margin-bottom: 10px;">
    <p style="margin: 0; font-size: 13px; color: #555; font-weight: bold;">📅 {tanggal_indo}</p>
    <p style="margin: 5px 0 0 0; font-size: 16px; color: #0D47A1; font-weight: 900;">⏰ {jam_indo}</p>
</div>
""", unsafe_allow_html=True)

st.sidebar.header("🛠️ Panel Filter Data RT & Status")
if not df.empty and "RT" in df.columns:
    df["RT_FORMAT"] = df["RT"].apply(lambda x: f"RT{int(x):02d}" if pd.notnull(x) and str(x).isdigit() else f"RT{str(x)}")
    semua_rt_format = sorted(df["RT_FORMAT"].dropna().unique(), key=lambda x: int(''.join(filter(str.isdigit, x)) or 0))
    pilihan_rt_format = st.sidebar.multiselect("Tampilkan Data RT:", options=semua_rt_format, default=semua_rt_format)
    semua_status_penduduk = ["Tetap", "Musiman"]
    pilihan_status_penduduk = st.sidebar.multiselect("Filter Status Penduduk:", options=semua_status_penduduk, default=semua_status_penduduk)
    
    df_filtered = df[df["RT_FORMAT"].isin(pilihan_rt_format)].copy()
    if "STATUS PENDUDUK" in df_filtered.columns and pilihan_status_penduduk:
        df_filtered = df_filtered[df_filtered["STATUS PENDUDUK"].astype(str).str.title().isin(pilihan_status_penduduk)]
else:
    df_filtered = pd.DataFrame()

password_input = st.sidebar.text_input("Masukkan Password Admin:", type="password")
admin_terverifikasi = (password_input == "V@nadminrw14")

st.markdown("<h2 style='color: #0D47A1;'>Portal Resmi & Dashboard Warga RW 14</h2>", unsafe_allow_html=True)

tab0, tab_struk, tab1, tab2, tab3, tab4, tab5, tab_rekap_rt, tab_update_kk, tab6, tab_kas_pemakaman_pub, tab7, tab8, tab9, tab10 = st.tabs([
    "🏠 Beranda", "👥 Struktur", "📋 Statistik", "👫 Demografi", "🎓 Pendidikan", 
    "🗂️ Data Warga", "🔍 Cari KK", "📊 Rekap RT", "📤 Update Data RT & KK", "💰 Kas RW", 
    "🪦 Kas Pemakaman", "📢 Info & Rapat", "🖼️ Galeri", "💬 Saran Pengurus", "⚙️ Edit & Upload (Admin)"
])

with tab0:
    st.subheader("👋 Selamat Datang di Portal Warga RW 14")
    st.write("Portal resmi layanan informasi dan kependudukan RW 14 Griya Permata Raya.")

with tab_struk:
    st.subheader("👥 Struktur Pengurus RW 14")
    if not df_struktur.empty:
        st.dataframe(df_struktur, use_container_width=True, hide_index=True)

with tab1:
    st.subheader("Angka Kunci Kependudukan Terkini")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("👥 Total Warga", f"{len(df_filtered)} Jiwa")
    kk_count = len(df_filtered[df_filtered["HUBUNGAN"].astype(str).str.upper() == "KEPALA KELUARGA"]) if "HUBUNGAN" in df_filtered.columns else 0
    c2.metric("👨‍💼 Kepala Keluarga", kk_count)
    laki_count = len(df_filtered[df_filtered["JENIS KELAMIN"].astype(str).str.upper() == "LAKI-LAKI"]) if "JENIS KELAMIN" in df_filtered.columns else 0
    c3.metric("👨 Laki-laki", laki_count)
    pr_count = len(df_filtered[df_filtered["JENIS KELAMIN"].astype(str).str.upper() == "PEREMPUAN"]) if "JENIS KELAMIN" in df_filtered.columns else 0
    c4.metric("👩 Perempuan", pr_count)

with tab6:
    st.subheader("💰 Transparansi Laporan Kas RW 14")
    if not df_kas.empty:
        st.dataframe(df_kas, use_container_width=True, hide_index=True)
    else:
        st.info("Data kas belum tersedia.")

with tab7:
    st.subheader("📢 Informasi Kegiatan & Hasil Rapat")
    if not df_info.empty:
        st.dataframe(df_info, use_container_width=True, hide_index=True)
    else:
        st.info("Belum ada informasi.")

with tab8:
    st.subheader("🖼️ Galeri Kegiatan Warga")
    if not df_galeri_meta.empty:
        st.dataframe(df_galeri_meta, use_container_width=True, hide_index=True)
    else:
        st.info("Belum ada foto galeri.")

with tab10:
    st.subheader("⚙️ Panel Admin")
    if admin_terverifikasi:
        st.success("✅ Admin Aktif")
    else:
        st.warning("Masukkan password admin di sidebar.")