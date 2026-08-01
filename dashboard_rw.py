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

# ================= SIDEBAR (WAKTU, JADWAL SHOLAT, MUSIK, FOTO RT) =================
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

st.sidebar.markdown("""
<div style="background: linear-gradient(135deg, #0D47A1, #1976D2); padding: 12px; border-radius: 10px; color: white; margin-bottom: 15px; box-shadow: 0px 4px 6px rgba(0,0,0,0.1);">
    <p style="margin: 0; font-size: 14px; font-weight: bold; text-align: center;">🕌 Jadwal Sholat & Pengingat</p>
    <p style="margin: 2px 0 10px 0; font-size: 11px; text-align: center; color: #E3F2FD;">Nanjung Mekar, Rancaekek, Kab. Bandung</p>
    <hr style="border-color: rgba(255,255,255,0.2); margin: 5px 0 8px 0;">
    <div style="font-size: 13px; display: flex; justify-content: space-between; margin-bottom: 4px;"><span>🌅 Subuh:</span> <b>04:44 WIB</b></div>
    <div style="font-size: 13px; display: flex; justify-content: space-between; margin-bottom: 4px;"><span>☀️ Dzuhur:</span> <b>12:00 WIB</b></div>
    <div style="font-size: 13px; display: flex; justify-content: space-between; margin-bottom: 4px;"><span>🌤️ Ashar:</span> <b>15:21 WIB</b></div>
    <div style="font-size: 13px; display: flex; justify-content: space-between; margin-bottom: 4px;"><span>🌇 Maghrib:</span> <b>17:58 WIB</b></div>
    <div style="font-size: 13px; display: flex; justify-content: space-between;"><span>🌙 Isya:</span> <b>19:06 WIB</b></div>
</div>
""", unsafe_allow_html=True)

st.sidebar.markdown("---")
st.sidebar.markdown("<p style='font-weight: bold; color: #0D47A1; margin-bottom: 5px; font-size: 14px;'>🎵 Pemutar Musik Latar:</p>", unsafe_allow_html=True)
file_musik = "backsound.mp3"
if os.path.exists(file_musik):
    with open(file_musik, "rb") as f:
        audio_bytes = f.read()
    audio_base64 = base64.b64encode(audio_bytes).decode()
    st.sidebar.markdown(f"""
    <audio controls loop style="width: 100%;">
        <source src="data:audio/mp3;base64,{audio_base64}" type="audio/mp3">
        Browser Anda tidak mendukung elemen audio.
    </audio>
    """, unsafe_allow_html=True)
else:
    st.sidebar.warning("⚠️ File 'backsound.mp3' belum ditemukan di GitHub.")

st.sidebar.header("🛠️ Panel Filter Data RT & Status")
if not df.empty and "RT" in df.columns:
    df["RT_FORMAT"] = df["RT"].apply(lambda x: f"RT{int(x):02d}" if pd.notnull(x) and str(x).isdigit() else f"RT{str(x)}")
    semua_rt_format = sorted(df["RT_FORMAT"].dropna().unique(), key=lambda x: int(''.join(filter(str.isdigit, x)) or 0))
    pilihan_rt_format = st.sidebar.multiselect("Tampilkan Data RT:", options=semua_rt_format, default=semua_rt_format)
    pilihan_rt_format = sorted(pilihan_rt_format, key=lambda x: int(''.join(filter(str.isdigit, x)) or 0))

    semua_status_penduduk = ["Tetap", "Musiman"]
    pilihan_status_penduduk = st.sidebar.multiselect("Filter Status Penduduk:", options=semua_status_penduduk, default=semua_status_penduduk)

    st.sidebar.markdown("---")
    st.sidebar.markdown("<p style='font-weight: bold; color: #0D47A1; margin-bottom: 10px; font-size: 15px;'>👨‍✈️ Profil Ketua RT Terpilih:</p>", unsafe_allow_html=True)

    daftar_ketua_rt_resmi = {
        "1": "M. Husni Mubarak", "2": "Casnanto", "3": "Ucok Yudho Hartono",
        "4": "Salya", "5": "Suwarno", "6": "Agus Hendra", "7": "Dodi Sunardi"
    }

    for rt_pilih in pilihan_rt_format:
        rt_num_clean = str(int(''.join(filter(str.isdigit, str(rt_pilih))) or 0))
        path_foto = None
        
        kemungkinan_nama = [
            f"RT {rt_num_clean.zfill(2)}.png", f"RT {rt_num_clean.zfill(2)}.PNG",
            f"RT {rt_num_clean}.png", f"RT {rt_num_clean}.PNG",
            f"RT 0{rt_num_clean}.png", f"RT {rt_num_clean}.jpg", f"RT {rt_num_clean}.JPG"
        ]
        
        for lokasi_file in kemungkinan_nama:
            if os.path.exists(lokasi_file):
                path_foto = lokasi_file
                break

        nama_ketua = daftar_ketua_rt_resmi.get(rt_num_clean, f"Ketua {rt_pilih}")

        if path_foto and os.path.exists(path_foto):
            st.sidebar.markdown(f"""
            <div style="background-color: #ffffff; padding: 10px; border-radius: 10px 10px 0 0; border: 2px solid #90CAF9; border-bottom: none; text-align: center;">
                <span style="background-color: #0D47A1; color: white; padding: 2px 8px; border-radius: 10px; font-weight: bold; font-size: 12px;">{rt_pilih}</span>
            </div>
            """, unsafe_allow_html=True)
            img_terluruskan = muat_dan_seragamkan_foto(path_foto, ukuran=(300, 400))
            st.sidebar.image(img_terluruskan, use_container_width=True)
            st.sidebar.markdown(f"""
            <div style="background-color: #ffffff; padding: 8px; border-radius: 0 0 10px 10px; border: 2px solid #90CAF9; border-top: none; text-align: center; margin-bottom: 15px; box-shadow: 0px 3px 8px rgba(0,0,0,0.08);">
                <p style="margin: 0; font-weight: bold; color: #0D47A1; font-size: 14px;">👨‍✈️ {nama_ketua}</p>
            </div>
            """, unsafe_allow_html=True)
        else:
            st.sidebar.markdown(f"""
            <div style="background-color: #ffffff; padding: 10px; border-radius: 10px; border: 1px dashed #1976D2; text-align: center; margin-bottom: 15px;">
                <p style="margin: 0; font-weight: bold; color: #0D47A1; font-size: 13px;">📌 {rt_pilih} - {nama_ketua}</p>
            </div>
            """, unsafe_allow_html=True)

    df_filtered = df[df["RT_FORMAT"].isin(pilihan_rt_format)].copy()
    if "STATUS PENDUDUK" in df_filtered.columns and pilihan_status_penduduk:
        df_filtered = df_filtered[df_filtered["STATUS PENDUDUK"].astype(str).str.title().isin(pilihan_status_penduduk)]
else:
    df_filtered = pd.DataFrame()

st.sidebar.markdown("---")
password_input = st.sidebar.text_input("Masukkan Password Admin:", type="password")
admin_terverifikasi = (password_input == "V@nadminrw14")

# ================= KEPALA HALAMAN & TABS UTAMA =================
col_logo, col_teks = st.columns([1, 6])
with col_logo:
    st.image(sumber_logo, width=70)
with col_teks:
    st.markdown("<h2 style='color: #0D47A1; font-weight: 900; margin: 0; padding-top: 5px; font-size: 20px;'>Portal Resmi & Dashboard Warga RW 14 Perum Griya Permata Raya Desa Nanjung Mekar Kec. Rancaekek Kab. Bandung</h2>", unsafe_allow_html=True)
    st.markdown("<p style='color: #333; font-weight: bold; margin: 3px 0 0 0; font-size: 14px;'>Pusat Layanan Informasi, Kependudukan, dan Transparansi Keuangan Lingkungan</p>", unsafe_allow_html=True)

st.write("---")

tab0, tab_struk, tab1, tab2, tab3, tab4, tab5, tab_rekap_rt, tab_update_kk, tab6, tab_kas_pemakaman_pub, tab7, tab8, tab9, tab10 = st.tabs([
    "🏠 Beranda", "👥 Struktur", "📋 Statistik", "👫 Demografi", "🎓 Pendidikan", 
    "🗂️ Data Warga", "🔍 Cari KK", "📊 Rekap RT", "📤 Update Data RT & KK", "💰 Kas RW", 
    "🪦 Kas Pemakaman", "📢 Info & Rapat", "🖼️ Galeri", "💬 Saran Pengurus", "⚙️ Edit & Upload (Admin)"
])

with tab0:
    st.subheader("👋 Selamat Datang di Portal Warga RW 14")
    col_p1, col_p2 = st.columns([2, 1])
    with col_p1:
        st.markdown("""
        ### 🌟 Sambutan Pengurus RW 14
        Assalamu’alaikum Warahmatullahi Wabarakatuh,  
        Selamat datang di website resmi **Portal & Dashboard Warga RW 14 Perum Griya Permata Raya Desa Nanjung Mekar Kec. Rancaekek Kab. Bandung**. Website ini dikembangkan khusus untuk memudahkan warga dan pengurus dalam mengakses informasi kependudukan secara transparan, akurat, dan cepat.
        """)
        st.markdown("---")
        st.markdown("### 🏛️ Jajaran Pengurus Inti RW 14")
        nama_rw = "Triyadi Sucipto"
        nama_sek = "Irvan Permana"
        nama_bend = "Aan Toni Fauyi"
        
        col_pengurus1, col_pengurus2, col_pengurus3 = st.columns(3)
        def cari_foto_pengurus_root(nama_file_dasar):
            for ext in ['.png', '.PNG', '.jpg', '.JPG', '.jpeg']:
                p = f"{nama_file_dasar}{ext}"
                if os.path.exists(p): return p
            return None

        with col_pengurus1:
            foto_rw = cari_foto_pengurus_root("KETUA RW")
            if foto_rw: st.image(muat_dan_seragamkan_foto(foto_rw, ukuran=(300, 400)), use_container_width=True)
            else: st.image("https://cdn-icons-png.flaticon.com/512/3135/3135673.png", use_container_width=True)
            st.markdown(f"<div style='text-align: center; font-weight: bold; color: #0D47A1; margin-top: 5px;'>{nama_rw}<br><span style='font-size: 12px; color: #555;'>Ketua RW 14</span></div>", unsafe_allow_html=True)

        with col_pengurus2:
            foto_sek = cari_foto_pengurus_root("SEKRETARIS")
            if foto_sek: st.image(muat_dan_seragamkan_foto(foto_sek, ukuran=(300, 400)), use_container_width=True)
            else: st.image("https://cdn-icons-png.flaticon.com/512/3135/3135673.png", use_container_width=True)
            st.markdown(f"<div style='text-align: center; font-weight: bold; color: #0D47A1; margin-top: 5px;'>{nama_sek}<br><span style='font-size: 12px; color: #555;'>Sekretaris</span></div>", unsafe_allow_html=True)

        with col_pengurus3:
            foto_bend = cari_foto_pengurus_root("BENDAHARA")
            if foto_bend: st.image(muat_dan_seragamkan_foto(foto_bend, ukuran=(300, 400)), use_container_width=True)
            else: st.image("https://cdn-icons-png.flaticon.com/512/3135/3135673.png", use_container_width=True)
            st.markdown(f"<div style='text-align: center; font-weight: bold; color: #0D47A1; margin-top: 5px;'>{nama_bend}<br><span style='font-size: 12px; color: #555;'>Bendahara</span></div>", unsafe_allow_html=True)

    with col_p2:
        st.markdown("""
        <div style="background-color: white; padding: 15px; border-radius: 10px; border: 1px solid #90CAF9;">
            <h4 style="color: #0D47A1; margin-top:0;">📞 Kontak Penting RW 14</h4>
            <p style="margin: 8px 0; font-size: 14px;">🚨 <b>Keamanan / Satpam:</b> 0812-XXXX-XXXX</p>
            <p style="margin: 8px 0; font-size: 14px;">🏥 <b>Kesehatan / Posyandu:</b> 0813-XXXX-XXXX</p>
            <hr style="margin: 10px 0;">
            <p style="margin: 0; font-size: 13px; color: #0D47A1; text-align: center; font-weight: 900;">GPR NGAHIJI</p>
        </div>
        """, unsafe_allow_html=True)

with tab_struk:
    st.subheader("👥 Struktur Pengurus RW 14")
    if not df_struktur.empty:
        st.dataframe(df_struktur, use_container_width=True, hide_index=True)

with tab1:
    st.subheader("Angka Kunci Kependudukan Terkini")
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("👥 Total Warga", f"{len(df_filtered)} Jiwa")
    kk_count = len(df_filtered[df_filtered["HUBUNGAN"].astype(str).str.upper() == "KEPALA KELUARGA"]) if "HUBUNGAN" in df_filtered.columns else 0
    col2.metric("👨‍💼 Kepala Keluarga", kk_count)
    laki_count = len(df_filtered[df_filtered["JENIS KELAMIN"].astype(str).str.upper() == "LAKI-LAKI"]) if "JENIS KELAMIN" in df_filtered.columns else 0
    col3.metric("👨 Laki-laki", laki_count)
    pr_count = len(df_filtered[df_filtered["JENIS KELAMIN"].astype(str).str.upper() == "PEREMPUAN"]) if "JENIS KELAMIN" in df_filtered.columns else 0
    col4.metric("👩 Perempuan", pr_count)

    st.write("---")
    st.subheader("Sebaran Penduduk per RT")
    if not df_filtered.empty and "RT_FORMAT" in df_filtered.columns:
        df_rt = df_filtered.groupby("RT_FORMAT").size().reset_index(name="Jumlah Warga")
        df_rt = df_rt.sort_values("RT_FORMAT")
        fig_rt = px.bar(df_rt, x="RT_FORMAT", y="Jumlah Warga", color="RT_FORMAT", text="Jumlah Warga", color_discrete_sequence=px.colors.qualitative.Vivid)
        fig_rt.update_layout(plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)", showlegend=False)
        st.plotly_chart(fig_rt, use_container_width=True)

with tab2:
    st.subheader("📊 Analisis Demografi Warga RW 14")
    if not df_filtered.empty and "JENIS KELAMIN" in df_filtered.columns:
        col_a, col_b = st.columns(2)
        with col_a:
            st.markdown("#### 🚻 Jenis Kelamin")
            fig_jk = px.pie(df_filtered, names="JENIS KELAMIN", hole=0.4, color_discrete_sequence=['#C71585', '#1B365D'])
            st.plotly_chart(fig_jk, use_container_width=True)
        with col_b:
            st.markdown("#### ☪️ Sebaran Agama")
            if "AGAMA" in df_filtered.columns:
                fig_agama = px.pie(df_filtered, names="AGAMA", hole=0.0)
                st.plotly_chart(fig_agama, use_container_width=True)

with tab3:
    st.subheader("🎓 Tingkat Pendidikan Warga RW 14")
    if not df_filtered.empty and "PENDIDIKAN" in df_filtered.columns:
        df_pendidikan = df_filtered["PENDIDIKAN"].astype(str).str.upper().value_counts().reset_index()
        df_pendidikan.columns = ["Tingkat Pendidikan", "Jumlah"]
        fig_pendidikan = px.bar(df_pendidikan, x="Tingkat Pendidikan", y="Jumlah", color="Tingkat Pendidikan", text_auto=True)
        st.plotly_chart(fig_pendidikan, use_container_width=True)

with tab4:
    st.subheader("🗂️ Data Seluruh Warga")
    if not df_filtered.empty:
        st.dataframe(df_filtered, use_container_width=True, hide_index=True)

with tab6:
    st.subheader("💰 Transparansi Laporan Kas RW 14")
    if not df_kas.empty:
        st.dataframe(df_kas, use_container_width=True, hide_index=True)
    else:
        st.info("Data kas belum tersedia.")

with tab_kas_pemakaman_pub:
    st.subheader("🪦 Transparansi Laporan Kas Pemakaman")
    if not df_kas_pemakaman.empty:
        st.dataframe(df_kas_pemakaman, use_container_width=True, hide_index=True)
    else:
        st.info("Data kas pemakaman belum tersedia.")

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
        st.success("✅ Login Admin Berhasil")
    else:
        st.warning("Masukkan password admin di sidebar sebelah kiri.")