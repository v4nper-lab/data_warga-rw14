import streamlit as st
import pandas as pd
import plotly.express as px
import base64
import os
from PIL import Image, ImageOps
from datetime import datetime, timedelta

# 1. PENGATURAN HALAMAN & KOSMETIK PORTAL
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
        m_col = kolom_masuk[0] if kolom_masuk else "PEMASUKAN"
        k_col = kolom_keluar[0] if kolom_keluar else "PENGELUARAN"
        
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
        m_col = kolom_masuk_kp[0] if kolom_masuk_kp else "PEMASUKAN"
        k_col = kolom_keluar_kp[0] if kolom_keluar_kp else "PENGELUARAN"
        
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

# ================= SIDEBAR =================
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

# ================= BREAKING NEWS & MOTIVASI DI ATAS JUDUL =================
teks_berita_online = "📢 SELAMAT DATANG DI PORTAL RESMI RW 14 GRIYA PERMATA RAYA &bull; "
if not df_info.empty:
    list_info_berita = []
    for _, r in df_info.iterrows():
        tgl_info = str(r.get("TANGGAL", ""))
        judul_info = str(r.get("JUDUL", ""))
        if judul_info and judul_info != "nan":
            list_info_berita.append(f"📌 [{tgl_info}] {judul_info}")
    if list_info_berita:
        teks_berita_online += " &nbsp;&bull;&nbsp; ".join(list_info_berita)
else:
    teks_berita_online += "Pengumuman dan agenda kegiatan lingkungan akan diperbarui secara berkala oleh Pengurus."

st.markdown(f"""
<div style="background: linear-gradient(135deg, #0D47A1, #1976D2); padding: 12px 18px; border-radius: 12px; margin-bottom: 15px; box-shadow: 0px 4px 10px rgba(0,0,0,0.15); border: 1px solid #90CAF9;">
    <div style="display: flex; align-items: center;">
        <span style="background-color: #ff9800; color: white; padding: 4px 10px; border-radius: 6px; font-size: 13px; font-weight: bold; margin-right: 12px; white-space: nowrap;">📰 BREAKING NEWS</span>
        <marquee behavior="scroll" direction="left" scrollamount="5" style="color: #ffffff; font-weight: 900; font-size: 16px; letter-spacing: 0.5px;">
            {teks_berita_online}
        </marquee>
    </div>
</div>
""", unsafe_allow_html=True)

st.markdown("""
<div style="background: linear-gradient(135deg, #E3F2FD, #BBDEFB); padding: 15px; border-radius: 12px; box-shadow: 0px 4px 10px rgba(0,0,0,0.06); margin-bottom: 15px; border: 2px solid #90CAF9;">
    <div style="background-color: #ffffff; padding: 6px 10px; border-radius: 8px; border: 1px solid #90CAF9; box-shadow: inset 0px 1px 3px rgba(0,0,0,0.05);">
        <marquee behavior="scroll" direction="left" scrollamount="5" style="color: #0D47A1; font-weight: bold; font-size: 14px;">
            🏡 Kepada seluruh Ketua RT RW 14 &nbsp;&bull;&nbsp; Mengurus data warga hari ini adalah investasi kemudahan untuk urusan sosial kemasyarakatan di masa depan &nbsp;&bull;&nbsp; Semangat terus melayani warga dengan sepenuh hati! ❤️
        </marquee>
    </div>
</div>
""", unsafe_allow_html=True)

# ================= KEPALA HALAMAN =================
col_logo, col_teks = st.columns([1, 6])
with col_logo:
    st.image(sumber_logo, width=70)
with col_teks:
    st.markdown("<h2 style='color: #0D47A1; font-weight: 900; margin: 0; padding-top: 5px; font-size: 20px;'>Portal Resmi & Dashboard Warga RW 14 Perum Griya Permata Raya Desa Nanjung Mekar Kec. Rancaekek Kab. Bandung</h2>", unsafe_allow_html=True)
    st.markdown("<p style='color: #333; font-weight: bold; margin: 3px 0 0 0; font-size: 14px;'>Pusat Layanan Informasi, Kependudukan, dan Transparansi Keuangan Lingkungan</p>", unsafe_allow_html=True)

st.write("---")

# ================= TABS UTAMA =================
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
    st.subheader("👥 Bagan & Uraian Tugas Struktur Pengurus RW 14")
    path_struktur_img = "struktur_rw.jpg" if os.path.exists("struktur_rw.jpg") else "struktur_rw.png"
    if os.path.exists(path_struktur_img):
        st.image(path_struktur_img, caption="Struktur Pengurus RW 014 Griya Permata Raya Periode 2024 - 2029", use_container_width=True)
    else:
        st.info("ℹ️ File gambar struktur belum diunggah.")
    st.write("---")
    if not df_struktur.empty: st.dataframe(df_struktur, use_container_width=True, hide_index=True)

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
        fig_rt.update_layout(plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)", showlegend=False, xaxis=dict(title="", tickfont=dict(size=16, color="black", weight="bold")), yaxis=dict(title="Jumlah Penduduk (Jiwa)")) 
        fig_rt.update_traces(textfont_size=18, textfont_color="black", textposition="outside")
        st.plotly_chart(fig_rt, use_container_width=True)

with tab2:
    st.subheader("📊 Analisis Demografi Warga RW 14")
    palet_agama_lain = ['#2980B9', '#A0522D', '#7F8C8D', '#3498DB', '#8B4513', '#95A5A6']
    if not df_filtered.empty:
        col_a, col_b, col_c = st.columns(3)
        with col_a:
            st.markdown("#### 🚻 Jenis Kelamin")
            if "JENIS KELAMIN" in df_filtered.columns:
                fig_jk = px.pie(df_filtered, names="JENIS KELAMIN", hole=0.4, color_discrete_sequence=['#C71585', '#1B365D'])
                fig_jk.update_layout(paper_bgcolor="rgba(0,0,0,0)", font=dict(size=13), legend=dict(orientation="h", yanchor="bottom", y=-0.2, xanchor="center", x=0.5))
                fig_jk.update_traces(textposition='inside', textfont_size=14, textfont_color="white", textinfo="label+percent")
                st.plotly_chart(fig_jk, use_container_width=True)
                
        with col_b:
            st.markdown("#### ☪️ Sebaran Agama")
            if "AGAMA" in df_filtered.columns:
                agama_list = df_filtered["AGAMA"].astype(str).str.title().unique()
                color_map = {ag: ("#2E8B57" if "ISLAM" in ag.upper() else palet_agama_lain[i % len(palet_agama_lain)]) for i, ag in enumerate(agama_list)}
                fig_agama = px.pie(df_filtered, names="AGAMA", hole=0.0, color="AGAMA", color_discrete_map=color_map)
                fig_agama.update_layout(paper_bgcolor="rgba(0,0,0,0)", font=dict(size=12), legend=dict(orientation="h", yanchor="bottom", y=-0.4, xanchor="center", x=0.5), margin=dict(t=20, b=80, l=10, r=10))
                teks_warna_list = ["white" if "ISLAM" in str(x).upper() else "black" for x in df_filtered["AGAMA"]]
                fig_agama.update_traces(textposition='inside', textfont_size=12, textfont_color=teks_warna_list, textinfo="label+percent")
                st.plotly_chart(fig_agama, use_container_width=True)

                df_agama_summary = df_filtered["AGAMA"].astype(str).str.title().value_counts().reset_index()
                df_agama_summary.columns = ["Agama", "Jumlah (Jiwa)"]
                df_agama_summary["Persentase (%)"] = ((df_agama_summary["Jumlah (Jiwa)"] / len(df_filtered)) * 100).round(2).astype(str) + "%"
                st.dataframe(df_agama_summary, use_container_width=True, hide_index=True)
                
        with col_c:
            st.markdown("#### 💍 Status Perkawinan")
            if "STATUS PERKAWINAN" in df_filtered.columns:
                df_status = df_filtered["STATUS PERKAWINAN"].astype(str).str.title().value_counts().reset_index()
                df_status.columns = ["Status", "Jumlah"]
                fig_status = px.bar(df_status, x="Status", y="Jumlah", color="Status", text_auto=True, color_discrete_sequence=['#1B365D', '#008080', '#D9822B', '#5C2D91', '#2E8B57'])
                fig_status.update_layout(paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", showlegend=False, font=dict(size=13))
                st.plotly_chart(fig_status, use_container_width=True)

with tab3:
    st.subheader("🎓 Tingkat Pendidikan Warga RW 14")
    if not df_filtered.empty and "PENDIDIKAN" in df_filtered.columns:
        df_pendidikan = df_filtered["PENDIDIKAN"].astype(str).str.upper().value_counts().reset_index()
        df_pendidikan.columns = ["Tingkat Pendidikan", "Jumlah"]
        fig_pendidikan = px.bar(df_pendidikan, x="Tingkat Pendidikan", y="Jumlah", color="Tingkat Pendidikan", text_auto=True, color_discrete_sequence=px.colors.qualitative.Prism)
        fig_pendidikan.update_layout(paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", showlegend=False, font=dict(size=14))
        st.plotly_chart(fig_pendidikan, use_container_width=True)

with tab4:
    st.subheader("🗂️ Data Seluruh Warga (Akses Khusus Pengurus)")
    if "warga_terbuka" not in st.session_state: st.session_state["warga_terbuka"] = False
    if not st.session_state["warga_terbuka"]:
        pass_warga = st.text_input("Kata Sandi Akses Data Warga:", type="password", key="pass_input_data_warga")
        if pass_warga == "ijindibuka": st.session_state["warga_terbuka"] = True; st.rerun()
        elif pass_warga != "": st.error("❌ Kata sandi salah!")
    else:
        if st.button("🔒 Kunci Kembali"): st.session_state["warga_terbuka"] = False; st.rerun()
        st.dataframe(df_filtered.drop(columns=["NO. KK", "NIK", "USIA_ANGKA", "Kelompok Usia", "RT_FORMAT"], errors="ignore"), use_container_width=True, hide_index=True)

with tab5:
    st.subheader("🔍 Pencarian KK")
    if "cari_terbuka" not in st.session_state: st.session_state["cari_terbuka"] = False
    if not st.session_state["cari_terbuka"]:
        pass_cari = st.text_input("Kata Sandi Akses Pencarian KK:", type="password", key="pass_input_cari_kk")
        if pass_cari == "ijindibuka": st.session_state["cari_terbuka"] = True; st.rerun()
        elif pass_cari != "": st.error("❌ Kata sandi salah!")
    else:
        if st.button("🔒 Kunci Kembali Pencarian"): st.session_state["cari_terbuka"] = False; st.rerun()
        kata_kunci = st.text_input("🔎 Masukkan Nama Warga:")
        if kata_kunci:
            hasil = df[df.astype(str).apply(lambda x: x.str.contains(kata_kunci, case=False)).any(axis=1)]
            st.dataframe(hasil, use_container_width=True, hide_index=True)

with tab6:
    st.subheader("💰 Transparansi Laporan Kas RW 14")
    if not df_kas.empty:
        val_m_sum = pd.to_numeric(df_kas["PEMASUKAN"], errors="coerce").fillna(0)
        val_k_sum = pd.to_numeric(df_kas["PENGELUARAN"], errors="coerce").fillna(0)
        total_masuk = val_m_sum.sum()
        total_keluar = val_k_sum.sum()
        saldo_akhir = (val_m_sum - val_k_sum).cumsum().iloc[-1] if not df_kas.empty else 0

        c1, c2, c3 = st.columns(3)
        c1.metric("💵 Total Pemasukan", f"Rp {total_masuk:,.0f}".replace(",", "."))
        c2.metric("💸 Total Pengeluaran", f"Rp {total_keluar:,.0f}".replace(",", "."))
        c3.metric("💰 Saldo Akhir", f"Rp {saldo_akhir:,.0f}".replace(",", "."))
        
        df_kas_display = df_kas.copy()
        df_kas_display["TANGGAL"] = df_kas_display["TANGGAL"].astype(str).str.replace(r'\.0$', '', regex=True)
        df_kas_display["PEMASUKAN"] = val_m_sum.apply(lambda x: f"Rp {x:,.0f}".replace(",", ".") if x > 0 else "Rp 0")
        df_kas_display["PENGELUARAN"] = val_k_sum.apply(lambda x: f"Rp {x:,.0f}".replace(",", ".") if x > 0 else "Rp 0")
        df_kas_display["SALDO"] = (val_m_sum - val_k_sum).cumsum().apply(lambda x: f"Rp {x:,.0f}".replace(",", "."))
        
        st.dataframe(df_kas_display, use_container_width=True, hide_index=True)
    else:
        st.info("Data kas belum tersedia.")

with tab_kas_pemakaman_pub:
    st.subheader("🪦 Transparansi Laporan Kas Pemakaman")
    if not df_kas_pemakaman.empty:
        val_m_kp = pd.to_numeric(df_kas_pemakaman["PEMASUKAN"], errors="coerce").fillna(0)
        val_k_kp = pd.to_numeric(df_kas_pemakaman["PENGELUARAN"], errors="coerce").fillna(0)
        tot_m_kp = val_m_kp.sum()
        tot_k_kp = val_k_kp.sum()
        saldo_kp = (val_m_kp - val_k_kp).cumsum().iloc[-1] if not df_kas_pemakaman.empty else 0

        c1, c2, c3 = st.columns(3)
        c1.metric("💵 Total Pemasukan", f"Rp {tot_m_kp:,.0f}".replace(",", "."))
        c2.metric("💸 Total Pengeluaran", f"Rp {tot_k_kp:,.0f}".replace(",", "."))
        c3.metric("💰 Saldo Akhir", f"Rp {saldo_kp:,.0f}".replace(",", "."))
        
        df_kp_display = df_kas_pemakaman.copy()
        df_kp_display["TANGGAL"] = df_kp_display["TANGGAL"].astype(str).str.replace(r'\.0$', '', regex=True)
        df_kp_display["PEMASUKAN"] = val_m_kp.apply(lambda x: f"Rp {x:,.0f}".replace(",", ".") if x > 0 else "Rp 0")
        df_kp_display["PENGELUARAN"] = val_k_kp.apply(lambda x: f"Rp {x:,.0f}".replace(",", ".") if x > 0 else "Rp 0")
        df_kp_display["SALDO"] = (val_m_kp - val_k_kp).cumsum().apply(lambda x: f"Rp {x:,.0f}".replace(",", "."))
        
        st.dataframe(df_kp_display, use_container_width=True, hide_index=True)
    else:
        st.info("Data kas pemakaman belum tersedia.")

with tab7:
    st.subheader("📢 Informasi Kegiatan & Hasil Rapat")
    if not df_info.empty:
        for _, r in df_info.iterrows():
            with st.expander(f"📌 [{r.get('TANGGAL', '')}] — {r.get('JUDUL', '')}"):
                st.write(r.get('ISI / KATEGORI', ''))
    else:
        st.info("Belum ada informasi.")

with tab8:
    st.subheader("🖼️ Galeri Kegiatan Warga")
    folder_galeri = "galeri"
    if os.path.exists(folder_galeri) and not df_galeri_meta.empty:
        mapping_bulan = {'Januari': 1, 'Februari': 2, 'Maret': 3, 'April': 4, 'Mei': 5, 'Juni': 6, 'Juli': 7, 'Agustus': 8, 'September': 9, 'Oktober': 10, 'November': 11, 'Desember': 12}
        df_g_sort = df_galeri_meta.copy()
        df_g_sort["BULAN_NUM"] = df_g_sort["BULAN"].map(mapping_bulan).fillna(0)
        df_g_sort["TAHUN_NUM"] = pd.to_numeric(df_g_sort["TAHUN"], errors="coerce").fillna(0)
        df_g_sort["TANGGAL_NUM"] = pd.to_numeric(df_g_sort["TANGGAL"], errors="coerce").fillna(0)
        df_g_sort = df_g_sort.sort_values(by=["TAHUN_NUM", "BULAN_NUM", "TANGGAL_NUM"], ascending=[False, False, False]).reset_index(drop=True)
        
        for (thn, bln), group_df in df_g_sort.groupby(["TAHUN", "BULAN"]):
            st.markdown(f"### 🗓️ Periode: {bln} {thn}")
            st.markdown("---")
            cols = st.columns(4)
            for idx, (_, row) in enumerate(group_df.iterrows()):
                nama_file = str(row.get("NAMA_FILE", ""))
                p_foto = os.path.join(folder_galeri, nama_file)
                if os.path.exists(p_foto):
                    with cols[idx % 4]:
                        st.image(muat_dan_seragamkan_foto(p_foto, ukuran=(400, 300)), use_container_width=True)
                        st.markdown(f"<p style='font-size: 12px; font-weight: bold;'>{row.get('KETERANGAN', '')}</p>", unsafe_allow_html=True)
    else:
        st.info("Belum ada foto galeri.")

with tab10:
    st.subheader("⚙️ Panel Admin")
    if admin_terverifikasi:
        menu_admin = st.selectbox(
            "Menu Admin:", 
            [
                "Data Warga", 
                "Struktur Organisasi", 
                "Laporan Kas RW", 
                "Laporan Kas Pemakaman/Sosial", 
                "Upload File PDF (Kas & Rapat)", 
                "Upload Foto Galeri", 
                "Edit Keterangan Galeri",
                "Hapus Foto Galeri"
            ]
        )
        if menu_admin == "Laporan Kas RW":
            st.markdown("💡 *Input atau edit data kas langsung di bawah ini, lalu klik tombol **Simpan Laporan Kas** di bawah.*")
            df_kas_edit = df_kas.drop(columns=["SALDO"], errors="ignore").copy()
            kas_terbaru = st.data_editor(df_kas_edit, num_rows="dynamic", use_container_width=True, key="editor_kas_rw_admin")
            
            if st.button("💾 Simpan Laporan Kas"):
                try:
                    if not kas_terbaru.empty:
                        m = [c for c in kas_terbaru.columns if "PEMASUKAN" in c or "MASUK" in c][0]
                        k = [c for c in kas_terbaru.columns if "PENGELUARAN" in c or "KELUAR" in c][0]
                        kas_terbaru["TANGGAL"] = kas_terbaru["TANGGAL"].astype(str).str.replace(r'\.0$', '', regex=True).replace('nan', '').str.strip()
                        kas_terbaru[m] = pd.to_numeric(kas_terbaru[m].astype(str).str.replace(r'[^0-9.-]', '', regex=True), errors="coerce").fillna(0)
                        kas_terbaru[k] = pd.to_numeric(kas_terbaru[k].astype(str).str.replace(r'[^0-9.-]', '', regex=True), errors="coerce").fillna(0)
                        kas_terbaru["SALDO"] = (kas_terbaru[m] - kas_terbaru[k]).cumsum()
                        kas_terbaru.to_excel("datakas.xlsx", index=False)
                        st.cache_data.clear()
                        st.success("✅ Laporan Kas RW berhasil disimpan dan diperbarui!")
                        st.rerun()
                except Exception as e:
                    st.error(f"❌ Gagal menyimpan: {e}")
        elif menu_admin == "Data Warga":
            ed = st.data_editor(df.drop(columns=["RT_FORMAT"], errors="ignore"), num_rows="dynamic", use_container_width=True)
            if st.button("Simpan Data Warga"):
                df_baru = pd.DataFrame(ed)
                df_terurut = urutkan_data_warga(df_baru)
                df_terurut.to_excel("datawarga.xlsx", index=False)
                st.cache_data.clear()
                st.success("✅ Data warga disimpan!")
                st.rerun()
        else:
            st.info("Pilih menu admin di atas sesuai kebutuhan.")
    else:
        st.warning("⚠️ Masukkan password admin di sidebar.")