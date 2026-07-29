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
.stApp { background-color: #F4F9F9; }
div[data-testid="metric-container"] { background-color: white; padding: 15px; border-radius: 10px; box-shadow: 0px 4px 6px rgba(0,0,0,0.05); border-left: 6px solid #1976D2; }
div[data-testid="stMetricValue"], div[data-testid="stMetricValue"] > div { font-size: 40px !important; color: #0D47A1 !important; font-weight: 900 !important; }
div[data-testid="stMetricLabel"] p, div[data-testid="stMetricLabel"] > div, div[data-testid="stMetricLabel"] { font-size: 18px !important; font-weight: bold !important; color: #2C3E50 !important; }
h3 { font-size: 24px !important; color: #0D47A1; }
button[data-baseweb="tab"] > div[data-testid="stMarkdownContainer"] > p { font-size: 13px !important; font-weight: bold; }
.stPlotlyChart { background-color: white; border-radius: 10px; box-shadow: 0px 4px 6px rgba(0,0,0,0.1); padding: 10px; }
</style>
""", unsafe_allow_html=True)

def ambil_logo_lokal(nama_file):
    if os.path.exists(nama_file):
        with open(nama_file, "rb") as img_file:
            encoded_string = base64.b64encode(img_file.read()).decode()
        ext = nama_file.split('.')[-1].lower()
        if ext in ['jpg', 'jpeg']: mime = 'jpeg'
        else: mime = 'png'
        return f"data:image/{mime};base64,{encoded_string}"
    else:
        return "https://cdn-icons-png.flaticon.com/512/3135/3135673.png"

sumber_logo = ambil_logo_lokal("logo rw.png")

# Fungsi untuk meluruskan orientasi foto secara presisi
def muat_dan_seragamkan_foto(path_file, ukuran=(300, 350)):
    try:
        img = Image.open(path_file)
        img = ImageOps.exif_transpose(img)
        img = ImageOps.fit(img, ukuran, method=Image.Resampling.LANCZOS, centering=(0.5, 0.5))
        return img
    except Exception:
        return path_file

@st.cache_data
def load_data():
    df = pd.read_excel("datawarga.xlsx")
    df.columns = df.columns.str.strip().str.upper()
    if "UMUR" in df.columns and "USIA" not in df.columns: df.rename(columns={"UMUR": "USIA"}, inplace=True)
    if "STATUS" in df.columns and "STATUS PERKAWINAN" not in df.columns: df.rename(columns={"STATUS": "STATUS PERKAWINAN"}, inplace=True)
    if "STATUS NIKAH" in df.columns and "STATUS PERKAWINAN" not in df.columns: df.rename(columns={"STATUS NIKAH": "STATUS PERKAWINAN"}, inplace=True)
    if "NO KK" in df.columns and "NO. KK" not in df.columns: df.rename(columns={"NO KK": "NO. KK"}, inplace=True)
    if "PENDIDIKAN TERAKHIR" in df.columns and "PENDIDIKAN" not in df.columns: df.rename(columns={"PENDIDIKAN TERAKHIR": "PENDIDIKAN"}, inplace=True)
    
    if "STATUS PENDUDUK" not in df.columns:
        df["STATUS PENDUDUK"] = "Tetap"
    else:
        df["STATUS PENDUDUK"] = df["STATUS PENDUDUK"].astype(str).str.strip().str.title()
        df["STATUS PENDUDUK"] = df["STATUS PENDUDUK"].apply(lambda x: x if x in ["Tetap", "Musiman"] else "Tetap")
    return df

@st.cache_data
def load_kas():
    if os.path.exists("datakas.xlsx"):
        df_kas = pd.read_excel("datakas.xlsx")
        df_kas.columns = df_kas.columns.str.strip().str.upper()
        return df_kas
    else:
        return pd.DataFrame(columns=["TANGGAL", "KETERANGAN", "PEMASUKAN", "PENGELUARAN", "SALDO"])

@st.cache_data
def load_kas_pemakaman():
    if os.path.exists("datakaspemakaman.xlsx"):
        df_kp = pd.read_excel("datakaspemakaman.xlsx")
        df_kp.columns = df_kp.columns.str.strip().str.upper()
        return df_kp
    else:
        return pd.DataFrame(columns=["TANGGAL", "KETERANGAN", "PEMASUKAN", "PENGELUARAN", "SALDO"])

@st.cache_data
def load_info():
    if os.path.exists("datainfo.xlsx"):
        df_info = pd.read_excel("datainfo.xlsx")
        df_info.columns = df_info.columns.str.strip().str.upper()
        return df_info
    else:
        return pd.DataFrame(columns=["TANGGAL", "JUDUL", "ISI / KATEGORI"])

@st.cache_data
def load_galeri_meta():
    if os.path.exists("datagaleri.xlsx"):
        df_g = pd.read_excel("datagaleri.xlsx")
        df_g.columns = df_g.columns.str.strip().str.upper()
        return df_g
    else:
        return pd.DataFrame(columns=["NAMA_FILE", "HARI", "TANGGAL", "BULAN", "TAHUN", "KETERANGAN"])

@st.cache_data
def load_saran():
    if os.path.exists("datasaran.xlsx"):
        df_s = pd.read_excel("datasaran.xlsx")
        df_s.columns = df_s.columns.str.strip().str.upper()
        return df_s
    else:
        return pd.DataFrame(columns=["WAKTU", "PENGIRIM", "JABATAN", "SARAN_PENDAPAT"])

# STRUKTUR PENGURUS LENGKAP DENGAN JOB DESCRIPTION & PROGRAM KERJA
@st.cache_data
def load_struktur():
    data_resmi = {
        "JABATAN / SEKSI": [
            "Ketua RW 014", 
            "PKK & Posyandu", 
            "Sekretaris", 
            "Bendahara", 
            "Keamanan & Ketertiban", 
            "Pembangunan & Lingkungan", 
            "Olahraga", 
            "Sosial & Pemakaman", 
            "Seni Budaya & Pemuda"
        ],
        "NAMA PENGURUS": [
            "Triyadi Sucipto", 
            "Tim PKK / Posyandu RW 014", 
            "Irvan Permana", 
            "Aan Toni Fauyi", 
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

# ================= WAKTU REAL-TIME & PENGINGAT SHOLAT DI SIDEBAR =================
st.sidebar.markdown("---")
waktu_sekarang = datetime.utcnow() + timedelta(hours=7)

hari_list = {"Monday": "Senin", "Tuesday": "Selasa", "Wednesday": "Rabu", "Thursday": "Kamis", "Friday": "Jumat", "Saturday": "Sabtu", "Sunday": "Minggu"}
bulan_list = {1: "Januari", 2: "Februari", 3: "Maret", 4: "April", 5: "Mei", 6: "Juni", 7: "Juli", 8: "Agustus", 9: "September", 10: "Oktober", 11: "November", 12: "Desember"}

nama_hari = hari_list.get(waktu_sekarang.strftime("%A"), waktu_sekarang.strftime("%A"))
nama_bulan = bulan_list.get(waktu_sekarang.month, "")
tanggal_indo = f"{nama_hari}, {waktu_sekarang.day} {nama_bulan} {waktu_sekarang.year}"
jam_indo = waktu_sekarang.strftime("%H:%M:%S") + " WIB"

st.sidebar.markdown(f"""
<div style="background-color: #E3F2FD; padding: 10px; border-radius: 8px; text-align: center; border: 1px solid #90CAF9; margin-bottom: 10px;">
    <p style="margin: 0; font-size: 13px; color: #555; font-weight: bold;">📅 {tanggal_indo}</p>
    <p style="margin: 5px 0 0 0; font-size: 16px; color: #0D47A1; font-weight: 900;">⏰ {jam_indo}</p>
</div>
""", unsafe_allow_html=True)

# Widget Pengingat Sholat Wilayah Kab. Bandung (Desa Nanjung Mekar, Kec. Rancaekek)
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

# ================= PEMUTAR MUSIK HTML5 LANGSUNG DI SIDEBAR =================
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
if "RT" in df.columns:
    df["RT_FORMAT"] = df["RT"].apply(lambda x: f"RT{int(x):02d}" if pd.notnull(x) and str(x).isdigit() else f"RT{str(x)}")
    semua_rt_format = sorted(df["RT_FORMAT"].dropna().unique(), key=lambda x: int(''.join(filter(str.isdigit, x)) or 0))
    pilihan_rt_format = st.sidebar.multiselect("Tampilkan Data RT:", options=semua_rt_format, default=semua_rt_format, key="filter_rt_sidebar_multiselect")
    pilihan_rt_format = sorted(pilihan_rt_format, key=lambda x: int(''.join(filter(str.isdigit, x)) or 0))

    semua_status_penduduk = ["Tetap", "Musiman"]
    pilihan_status_penduduk = st.sidebar.multiselect("Filter Status Penduduk:", options=semua_status_penduduk, default=semua_status_penduduk, key="filter_status_penduduk_sidebar")

    # ================= FOTO KETUA RT FORMAT VERTIKAL DI SIDEBAR =================
    st.sidebar.markdown("---")
    st.sidebar.markdown("<p style='font-weight: bold; color: #0D47A1; margin-bottom: 10px; font-size: 15px;'>👨‍✈️ Profil Ketua RT Terpilih:</p>", unsafe_allow_html=True)
    
    folder_foto_rt = "rt"
    if not os.path.exists(folder_foto_rt):
        os.makedirs(folder_foto_rt)

    daftar_ketua_rt_resmi = {
        "1": "M. Husni Mubarak",
        "2": "Casnanto",
        "3": "Ucok Yudho Hartono",
        "4": "Salya",
        "5": "Suwarno",
        "6": "Agus Hendra",
        "7": "Dodi Sunardi"
    }

    for rt_pilih in pilihan_rt_format:
        rt_num_clean = str(int(''.join(filter(str.isdigit, str(rt_pilih))) or 0))
        
        path_foto = None
        kemungkinan_nama = [
            os.path.join(folder_foto_rt, f"rt{rt_num_clean}.jpg"), os.path.join(folder_foto_rt, f"rt{rt_num_clean}.jpeg"), os.path.join(folder_foto_rt, f"rt{rt_num_clean}.png"),
            os.path.join(folder_foto_rt, f"rt0{rt_num_clean}.jpg"), os.path.join(folder_foto_rt, f"rt0{rt_num_clean}.png"),
            f"rt{rt_num_clean}.jpg", f"rt{rt_num_clean}.png", f"rt{rt_num_clean}.jpg", f"rt{rt_num_clean}.png"
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
            
            img_terluruskan = muat_dan_seragamkan_foto(path_foto, ukuran=(250, 300))
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
                <p style="margin: 4px 0 0 0; font-size: 11px; color: #777;"><i>(Foto belum diunggah via Admin)</i></p>
            </div>
            """, unsafe_allow_html=True)

    if not pilihan_rt_format:
        st.warning("⚠️ Silakan pilih minimal satu RT di menu sebelah kiri.")
        st.stop()
        
    df_filtered = df[df["RT_FORMAT"].isin(pilihan_rt_format)].copy()
    if "STATUS PENDUDUK" in df_filtered.columns and pilihan_status_penduduk:
        df_filtered = df_filtered[df_filtered["STATUS PENDUDUK"].astype(str).str.title().isin(pilihan_status_penduduk)]
else:
    st.error("Kolom 'RT' tidak ditemukan di Excel.")
    st.stop()

# ================= KONTROL KEAMANAN ADMIN (STABIL DENGAN SESSION STATE) =================
st.sidebar.markdown("---")
st.sidebar.markdown("<p style='font-weight: bold; color: red; font-size: 16px; margin-bottom: 5px;'>🔐 Menu Pengurus (Admin)</p>", unsafe_allow_html=True)

if "admin_logged_in" not in st.session_state:
    st.session_state["admin_logged_in"] = False

password_input = st.sidebar.text_input("Masukkan Password Admin:", type="password", key="input_password_admin_segel_unik_rw14")

if password_input == "V@nadminrw14":
    st.session_state["admin_logged_in"] = True
elif password_input != "":
    st.session_state["admin_logged_in"] = False
    st.sidebar.error("❌ Password salah!")

admin_terverifikasi = st.session_state["admin_logged_in"]
if admin_terverifikasi:
    st.sidebar.success("✅ Login Admin Berhasil!")

# =========================================================================
# ============ TEKS BERITA BERJALAN ONLINE (HURUF DIPERBESAR) ==============
# =========================================================================
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

# ================= KONTEN UTAMA PORTAL =================
st.markdown("""
<div style="background: linear-gradient(135deg, #E3F2FD, #BBDEFB); padding: 20px; border-radius: 15px; box-shadow: 0px 4px 10px rgba(0,0,0,0.06); margin-bottom: 20px; border: 2px solid #90CAF9;">
""", unsafe_allow_html=True)

# Teks Berjalan Motivasi RT (Di dalam frame utama)
st.markdown("""
<div style="background-color: #ffffff; padding: 6px 10px; border-radius: 8px; border: 1px solid #90CAF9; margin-bottom: 10px; box-shadow: inset 0px 1px 3px rgba(0,0,0,0.05);">
    <marquee behavior="scroll" direction="left" scrollamount="5" style="color: #0D47A1; font-weight: bold; font-size: 14px;">
        🏡 Kepada seluruh Ketua RT RW 14 &nbsp;&bull;&nbsp; Mengurus data warga hari ini adalah investasi kemudahan untuk urusan sosial kemasyarakatan di masa depan &nbsp;&bull;&nbsp; Semangat terus melayani warga dengan sepenuh hati! ❤️
    </marquee>
</div>
""", unsafe_allow_html=True)

col_logo, col_teks = st.columns([1, 6])
with col_logo:
    st.image(sumber_logo, width=70)
with col_teks:
    st.markdown("<h2 style='color: #0D47A1; font-weight: 900; margin: 0; padding-top: 5px; text-shadow: 1px 1px 2px rgba(255,255,255,0.8); font-size: 20px;'>Portal Resmi & Dashboard Warga RW 14 Perum Griya Permata Raya Desa Nanjung Mekar Kec. Rancaekek Kab. Bandung</h2>", unsafe_allow_html=True)
    st.markdown("<p style='color: #333; font-weight: bold; margin: 3px 0 0 0; font-size: 14px;'>Pusat Layanan Informasi, Kependudukan, dan Transparansi Keuangan Lingkungan</p>", unsafe_allow_html=True)

st.write("---")

# ================= MENU UTAMA WEBSITE PORTAL (JUMLAH TAB KONSISTEN) =================
tab0, tab_struk, tab1, tab2, tab3, tab4, tab5, tab_update_kk, tab6, tab7, tab8, tab9, tab10 = st.tabs([
    "🏠 Beranda", "👥 Struktur", "📋 Statistik", "👫 Demografi", "🎓 Pendidikan", 
    "🗂️ Data Warga", "🔍 Cari KK", "📤 Update Data RT & KK", "💰 Kas RW", 
    "📢 Info & Rapat", "🖼️ Galeri", "💬 Saran Pengurus", "⚙️ Edit & Upload (Admin)"
])

# ================= TAB 0: BERANDA / PROFIL =================
with tab0:
    st.subheader("👋 Selamat Datang di Portal Warga RW 14")
    
    col_p1, col_p2 = st.columns([2, 1])
    with col_p1:
        st.markdown("""
        ### 🌟 Sambutan Pengurus RW 14
        Assalamu’alaikum Warahmatullahi Wabarakatuh,  
        Selamat datang di website resmi **Portal & Dashboard Warga RW 14 Perum Griya Permata Raya Desa Nanjung Mekar Kec. Rancaekek Kab. Bandung**. Website ini dikembangkan khusus untuk memudahkan warga dan pengurus dalam mengakses informasi kependudukan secara transparan, akurat, dan cepat.
        
        Melalui portal digital ini, Anda dapat:
        * Melihat struktur kepengurusan RW dan profil Ketua RT secara vertikal di panel sebelah kiri.
        * Memeriksa statistik kependudukan dan tingkat pendidikan warga.
        * Mencari data Kartu Keluarga (KK) dengan mudah.
        * Memantau transparansi laporan keuangan kas RW dan laporan khusus seksi pemakaman.
        * Membaca hasil rapat, agenda kegiatan, dokumen PDF resmi, serta galeri foto lingkungan yang tersusun rapi berdasarkan arsip manual Admin.
        * Mengakses menu **Saran Pengurus** khusus untuk Ketua RT 01 s.d. 07 serta Pengurus Inti.
        
        Mari bersama-sama kita wujudkan kerukunan, keterbukaan, dan pelayanan warga yang semakin prima!
        """)
        
        # ================= TAMPILAN FOTO PENGURUS INTI DI BERANDA =================
        st.markdown("---")
        st.markdown("### 🏛️ Jajaran Pengurus Inti RW 14")
        
        nama_rw = "Triyadi Sucipto"
        nama_sek = "Irvan Permana"
        nama_bend = "Aan Toni Fauyi"
        if not df_struktur.empty:
            for _, row in df_struktur.iterrows():
                jab = str(row.get("JABATAN / SEKSI", "")).upper()
                nama_val = str(row.get("NAMA PENGURUS", ""))
                if "KETUA RW" in jab: nama_rw = nama_val
                elif "SEKRETARIS" in jab: nama_sek = nama_val
                elif "BENDAHARA" in jab: nama_bend = nama_val

        col_pengurus1, col_pengurus2, col_pengurus3 = st.columns(3)
        
        folder_pengurus = "pengurus"
        if not os.path.exists(folder_pengurus):
            os.makedirs(folder_pengurus)

        def cari_foto_pengurus(nama_file_dasar):
            for ext in ['.jpg', '.jpeg', '.png', '.JPG', '.PNG']:
                p = os.path.join(folder_pengurus, f"{nama_file_dasar}{ext}")
                if os.path.exists(p): return p
                p2 = f"{nama_file_dasar}{ext}"
                if os.path.exists(p2): return p2
            return None

        with col_pengurus1:
            foto_rw = cari_foto_pengurus("ketuarw")
            if foto_rw:
                img_rw = muat_dan_seragamkan_foto(foto_rw, ukuran=(280, 320))
                st.image(img_rw, use_container_width=True)
            else:
                st.image("https://cdn-icons-png.flaticon.com/512/3135/3135673.png", use_container_width=True)
            st.markdown(f"<div style='text-align: center; font-weight: bold; color: #0D47A1; margin-top: 5px;'>{nama_rw}<br><span style='font-size: 12px; color: #555;'>Ketua RW 14</span></div>", unsafe_allow_html=True)

        with col_pengurus2:
            foto_sek = cari_foto_pengurus("sekretaris")
            if foto_sek:
                img_sek = muat_dan_seragamkan_foto(foto_sek, ukuran=(280, 320))
                st.image(img_sek, use_container_width=True)
            else:
                st.image("https://cdn-icons-png.flaticon.com/512/3135/3135673.png", use_container_width=True)
            st.markdown(f"<div style='text-align: center; font-weight: bold; color: #0D47A1; margin-top: 5px;'>{nama_sek}<br><span style='font-size: 12px; color: #555;'>Sekretaris</span></div>", unsafe_allow_html=True)

        with col_pengurus3:
            foto_bend = cari_foto_pengurus("bendahara")
            if foto_bend:
                img_bend = muat_dan_seragamkan_foto(foto_bend, ukuran=(280, 320))
                st.image(img_bend, use_container_width=True)
            else:
                st.image("https://cdn-icons-png.flaticon.com/512/3135/3135673.png", use_container_width=True)
            st.markdown(f"<div style='text-align: center; font-weight: bold; color: #0D47A1; margin-top: 5px;'>{nama_bend}<br><span style='font-size: 12px; color: #555;'>Bendahara</span></div>", unsafe_allow_html=True)

    with col_p2:
        st.markdown("""
        <div style="background-color: white; padding: 15px; border-radius: 10px; border: 1px solid #90CAF9; box-shadow: 0px 2px 5px rgba(0,0,0,0.05);">
            <h4 style="color: #0D47A1; margin-top:0;">📞 Kontak Penting RW 14</h4>
            <p style="margin: 8px 0; font-size: 14px;">🚨 <b>Keamanan / Satpam:</b> 0812-XXXX-XXXX</p>
            <p style="margin: 8px 0; font-size: 14px;">🏥 <b>Kesehatan / Posyandu:</b> 0813-XXXX-XXXX</p>
            <p style="margin: 8px 0; font-size: 14px;">🧹 <b>Kebersihan / RT:</b> Hubungi RT Masing-masing</p>
            <hr style="margin: 10px 0;">
            <p style="margin: 0 0 4px 0; font-size: 12px; color: #666; text-align: center;"><b>RW 14 Bersih, Rukun, & Sejahtera</b></p>
            <p style="margin: 0; font-size: 13px; color: #0D47A1; text-align: center; font-weight: 900; letter-spacing: 1px;">GPR NGAHIJI</p>
        </div>
        """, unsafe_allow_html=True)

# ================= TAB STRUKTUR ORGANISASI =================
with tab_struk:
    st.subheader("👥 Bagan & Uraian Tugas (Job Description) Struktur Pengurus RW 14")
    st.markdown("Bagan organigram resmi, rincian tugas, serta bentuk program kerja dari masing-masing seksi Periode 2024 - 2029.")
    
    path_struktur_img = "struktur_rw.jpg"
    if not os.path.exists(path_struktur_img):
        path_struktur_img = "struktur_rw.png"
        
    if os.path.exists(path_struktur_img):
        st.image(path_struktur_img, caption="Struktur Pengurus RW 014 Griya Permata Raya Periode 2024 - 2029", use_container_width=True)
    else:
        st.info("ℹ️ File gambar struktur belum diunggah. Silakan upload file gambar dengan nama 'struktur_rw.jpg' ke folder utama project atau via menu Admin.")
    
    st.write("---")
    st.markdown("### 📋 Rincian Job Description & Program Kerja Masing-Masing Seksi")
    if not df_struktur.empty:
        st.dataframe(df_struktur, use_container_width=True, hide_index=True)
    else:
        st.info("ℹ️ Belum ada data struktur pengurus yang dimasukkan.")

    # ================= LAPORAN KHUSUS SEKSI PEMAKAMAN & SOSIAL =================
    st.write("---")
    st.subheader("🕊️ Laporan Khusus Keuangan Seksi Sosial & Pemakaman")
    st.markdown("Berikut adalah transparansi pencatatan khusus dana sosial, santunan duka, dan pemakaman warga RW 14 yang dikelola secara terpisah.")

    if not df_kas_pemakaman.empty:
        kolom_pemasukan_kp = [c for c in df_kas_pemakaman.columns if "PEMASUKAN" in c or "MASUK" in c]
        kolom_pengeluaran_kp = [c for c in df_kas_pemakaman.columns if "PENGELUARAN" in c or "KELUAR" in c]
        kolom_saldo_kp = [c for c in df_kas_pemakaman.columns if "SALDO" in c]

        total_masuk_kp = 0
        total_keluar_kp = 0
        
        if kolom_pemasukan_kp:
            df_kas_pemakaman["MASUK_ANGKA"] = pd.to_numeric(df_kas_pemakaman[kolom_pemasukan_kp[0]], errors="coerce").fillna(0)
            total_masuk_kp = df_kas_pemakaman["MASUK_ANGKA"].sum()
        if kolom_pengeluaran_kp:
            df_kas_pemakaman["KELUAR_ANGKA"] = pd.to_numeric(df_kas_pemakaman[kolom_pengeluaran_kp[0]], errors="coerce").fillna(0)
            total_keluar_kp = df_kas_pemakaman["KELUAR_ANGKA"].sum()

        saldo_akhir_kp = total_masuk_kp - total_keluar_kp
        if kolom_saldo_kp and not df_kas_pemakaman[kolom_saldo_kp[0]].dropna().empty:
            last_saldo = pd.to_numeric(df_kas_pemakaman[kolom_saldo_kp[0]], errors="coerce").iloc[-1]
            if pd.notnull(last_saldo): saldo_akhir_kp = last_saldo

        cp1, cp2, cp3 = st.columns(3)
        cp1.metric("💵 Total Pemasukan", f"Rp {total_masuk_kp:,.0f}".replace(",", "."))
        cp2.metric("💸 Total Pengeluaran", f"Rp {total_keluar_kp:,.0f}".replace(",", "."))
        cp3.metric("💰 Saldo Kas Pemakaman", f"Rp {saldo_akhir_kp:,.0f}".replace(",", "."))

        st.write("---")
        st.dataframe(df_kas_pemakaman.drop(columns=["MASUK_ANGKA", "KELUAR_ANGKA"], errors="ignore"), use_container_width=True, hide_index=True)
    else:
        st.info("ℹ️ Belum ada data transaksi khusus pemakaman yang dimasukkan.")

# ================= TAB 1: STATISTIK =================
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
    df_rt = df_filtered.groupby("RT_FORMAT").size().reset_index(name="Jumlah Warga")
    df_rt = df_rt.sort_values("RT_FORMAT")
    fig_rt = px.bar(df_rt, x="RT_FORMAT", y="Jumlah Warga", color="RT_FORMAT", text="Jumlah Warga", color_discrete_sequence=px.colors.qualitative.Vivid)
    fig_rt.update_layout(plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)", showlegend=False, xaxis=dict(title="", tickfont=dict(size=20, color="black", weight="bold")), yaxis=dict(title="Jumlah Penduduk (Jiwa)"), margin=dict(t=40)) 
    fig_rt.update_traces(textfont_size=24, textfont_color="black", textangle=0, textposition="outside", cliponaxis=False)
    st.plotly_chart(fig_rt, use_container_width=True)

# ================= TAB 2: DEMOGRAFI (WARNA ELEGAN & INTERAKTIF) =================
with tab2:
    st.subheader("📊 Analisis Demografi Warga RW 14")
    st.markdown("Statistik demografi kependudukan yang disajikan secara interaktif.")
    
    palet_elegan = ['#1B365D', '#008080', '#D9822B', '#5C2D91', '#2E8B57', '#C0392B', '#2980B9']

    col_a, col_b, col_c = st.columns(3)
    with col_a:
        st.markdown("#### 🚻 Jenis Kelamin")
        if "JENIS KELAMIN" in df_filtered.columns:
            fig_jk = px.pie(df_filtered, names="JENIS KELAMIN", hole=0.55, color_discrete_sequence=['#1B365D', '#D9822B'])
            fig_jk.update_layout(paper_bgcolor="rgba(0,0,0,0)", font=dict(size=14, family="sans-serif"), legend=dict(orientation="h", yanchor="bottom", y=-0.2, xanchor="center", x=0.5))
            fig_jk.update_traces(textfont_size=15, hoverinfo="label+percent+value", textinfo="label+percent")
            st.plotly_chart(fig_jk, use_container_width=True)
            
    with col_b:
        st.markdown("#### ☪️ Sebaran Agama")
        if "AGAMA" in df_filtered.columns:
            fig_agama = px.pie(df_filtered, names="AGAMA", hole=0.55, color_discrete_sequence=palet_elegan)
            fig_agama.update_layout(paper_bgcolor="rgba(0,0,0,0)", font=dict(size=14, family="sans-serif"), legend=dict(orientation="h", yanchor="bottom", y=-0.2, xanchor="center", x=0.5))
            fig_agama.update_traces(textfont_size=15, hoverinfo="label+percent+value", textinfo="label+percent")
            st.plotly_chart(fig_agama, use_container_width=True)
            
    with col_c:
        st.markdown("#### 💍 Status Perkawinan")
        if "STATUS PERKAWINAN" in df_filtered.columns:
            df_status = df_filtered["STATUS PERKAWINAN"].astype(str).str.title().value_counts().reset_index()
            df_status.columns = ["Status", "Jumlah"]
            fig_status = px.bar(df_status, x="Status", y="Jumlah", color="Status", text_auto=True, color_discrete_sequence=palet_elegan)
            fig_status.update_layout(paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", showlegend=False, font=dict(size=14, family="sans-serif"))
            fig_status.update_traces(textfont_size=16, textangle=0, marker_line_color='rgb(8,48,107)', marker_line_width=1.5)
            st.plotly_chart(fig_status, use_container_width=True)

# ================= TAB 3: PENDIDIKAN =================
with tab3:
    st.subheader("🎓 Tingkat Pendidikan Warga RW 14")
    st.markdown("Grafik sebaran tingkat pendidikan formal warga di lingkungan RW.")
    
    if "PENDIDIKAN" in df_filtered.columns:
        df_pendidikan = df_filtered["PENDIDIKAN"].astype(str).str.upper().value_counts().reset_index()
        df_pendidikan.columns = ["Tingkat Pendidikan", "Jumlah"]
        
        fig_pendidikan = px.bar(df_pendidikan, x="Tingkat Pendidikan", y="Jumlah", color="Tingkat Pendidikan", text_auto=True, color_discrete_sequence=px.colors.qualitative.Prism)
        fig_pendidikan.update_layout(paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", showlegend=False, font=dict(size=15))
        fig_pendidikan.update_traces(textfont_size=18, textangle=0)
        st.plotly_chart(fig_pendidikan, use_container_width=True)
    else:
        st.info("ℹ️ Kolom 'PENDIDIKAN' atau 'PENDIDIKAN TERAKHIR' belum tersedia di file Excel datawarga.")

# ================= TAB 4: SEMUA DATA (DIKUNCI SANDI: ijindibuka) =================
with tab4:
    st.subheader("🗂️ Data Seluruh Warga (Akses Khusus Pengurus)")
    
    if "warga_terbuka" not in st.session_state:
        st.session_state["warga_terbuka"] = False
        
    if not st.session_state["warga_terbuka"]:
        st.info("🔒 Menu ini dilindungi. Masukkan kata sandi pengurus untuk membuka data seluruh warga.")
        pass_warga = st.text_input("Kata Sandi Akses Data Warga:", type="password", key="pass_input_data_warga")
        if pass_warga == "ijindibuka":
            st.session_state["warga_terbuka"] = True
            st.success("✅ Akses diberikan!")
            st.rerun()
        elif pass_warga != "":
            st.error("❌ Kata sandi salah!")
    else:
        st.success("✅ Anda sedang dalam mode akses pengurus.")
        if st.button("🔒 Kunci Kembali Data Warga", key="btn_kunci_warga"):
            st.session_state["warga_terbuka"] = False
            st.rerun()
            
        st.markdown("---")
        kolom_dibuang = ["NO. KK", "NIK", "USIA_ANGKA", "Kelompok Usia", "RT_FORMAT"]
        df_tabel = df_filtered.drop(columns=kolom_dibuang, errors="ignore")
        st.dataframe(df_tabel, use_container_width=True, hide_index=True)

# ================= TAB 5: PENCARIAN KK (DIKUNCI SANDI: ijindibuka) =================
with tab5:
    st.subheader("🔍 Pencarian & Data per Kartu Keluarga (KK) (Akses Khusus Pengurus)")
    
    if "cari_terbuka" not in st.session_state:
        st.session_state["cari_terbuka"] = False
        
    if not st.session_state["cari_terbuka"]:
        st.info("🔒 Menu pencarian Kartu Keluarga dilindungi. Masukkan kata sandi pengurus untuk mengakses.")
        pass_cari = st.text_input("Kata Sandi Akses Pencarian KK:", type="password", key="pass_input_cari_kk")
        if pass_cari == "ijindibuka":
            st.session_state["cari_terbuka"] = True
            st.success("✅ Akses diberikan!")
            st.rerun()
        elif pass_cari != "":
            st.error("❌ Kata sandi salah!")
    else:
        st.success("✅ Anda sedang dalam mode akses pengurus.")
        if st.button("🔒 Kunci Kembali Pencarian KK", key="btn_kunci_cari"):
            st.session_state["cari_terbuka"] = False
            st.rerun()
            
        st.markdown("---")
        kata_kunci = st.text_input("🔎 Masukkan Nama Warga atau Kepala Keluarga:")
        
        if kata_kunci:
            kunci_bersih = kata_kunci.strip().lower()
            kolom_nama_opsi = [col for col in df.columns if "NAMA" in col]
            
            if kolom_nama_opsi:
                nama_kolom_aktif = kolom_nama_opsi[0]
                mask_nama = df[nama_kolom_aktif].astype(str).str.lower().str.contains(kunci_bersih, na=False)
                hasil_pencarian = df[mask_nama]
            else:
                hasil_pencarian = pd.DataFrame()
                nama_kolom_aktif = None
            
            if not hasil_pencarian.empty:
                kolom_kk_opsi = [col for col in df.columns if "KK" in col]
                if kolom_kk_opsi:
                    kk_kolom_aktif = kolom_kk_opsi[0]
                    list_kk = hasil_pencarian[kk_kolom_aktif].dropna().unique()
                    st.success(f"✅ Ditemukan {len(list_kk)} Kartu Keluarga terkait.")
                    
                    for kk in list_kk:
                        df_keluarga = df[df[kk_kolom_aktif] == kk].copy()
                        
                        nama_kepala = "Satu Keluarga"
                        kolom_hub = [col for col in df.columns if "HUBUNGAN" in col or "STATUS" in col]
                        if kolom_hub and kolom_nama_opsi:
                            kepala_df = df_keluarga[df_keluarga[kolom_hub[0]].astype(str).str.upper().str.contains("KEPALA KELUARGA", na=False)]
                            if not kepala_df.empty:
                                nama_kepala = "Keluarga Bpk/Ibu " + str(kepala_df.iloc[0][nama_kolom_aktif]).title()
                        
                        st.markdown(f"### 🏠 {nama_kepala}")
                        
                        def highlight_pencarian(row):
                            match = kunci_bersih in str(row.get(nama_kolom_aktif, "")).lower()
                            return ['background-color: #FFF9C4' if match else '' for _ in row]

                        kolom_dibuang = [kk_kolom_aktif, "NIK", "USIA_ANGKA", "Kelompok Usia", "RT_FORMAT"]
                        df_tampil = df_keluarga.drop(columns=[c for c in kolom_dibuang if c in df_keluarga.columns], errors="ignore")
                        
                        st.dataframe(df_tampil.style.apply(highlight_pencarian, axis=1), use_container_width=True, hide_index=True)
                else:
                    st.error("Kolom yang mengandung kata 'KK' tidak ditemukan di Excel.")
            else:
                st.warning(f"❌ Tidak ada warga dengan nama '{kata_kunci}' yang ditemukan.")

# ================= TAB 7: UPDATE DATA RT & UPLOAD DOKUMEN KK BARU =================
with tab_update_kk:
    st.subheader("📤 Menu Update Data Warga & Unggah Dokumen KK Baru (Khusus Ketua RT)")
    st.markdown("Menu ini digunakan oleh Ketua RT untuk melaporkan perubahan data warga (misalnya pembaruan KK lama ke KK baru) sekaligus mengirimkan dokumen bukti scan KK baru ke Pengurus RW.")

    if "rt_update_terverifikasi" not in st.session_state:
        st.session_state["rt_update_terverifikasi"] = False

    if not st.session_state["rt_update_terverifikasi"]:
        st.info("🔒 Masukkan kata sandi pengurus/Ketua RT untuk mengakses formulir pengajuan perubahan data warga.")
        pass_rt_up = st.text_input("Kata Sandi Akses Menu Update:", type="password", key="pass_input_rt_update")
        if pass_rt_up == "ijindibuka":
            st.session_state["rt_update_terverifikasi"] = True
            st.success("✅ Akses diberikan!")
            st.rerun()
        elif pass_rt_up != "":
            st.error("❌ Kata sandi salah!")
    else:
        st.success("✅ Mode Akses Ketua RT Aktif.")
        if st.button("🔒 Keluar / Kunci Menu Update", key="btn_kunci_rt_update"):
            st.session_state["rt_update_terverifikasi"] = False
            st.rerun()

        st.markdown("---")
        with st.form("form_update_kk_rt"):
            st.markdown("### 📝 Formulir Pengajuan Perubahan Data & Unggah KK Baru")
            
            col_u1, col_u2 = st.columns(2)
            with col_u1:
                pilih_rt_lapor = st.selectbox("Pilih RT:", ["RT 01", "RT 02", "RT 03", "RT 04", "RT 05", "RT 06", "RT 07"])
                nama_pelapor = st.text_input("Nama Ketua RT / Pengurus Pelapor:")
            with col_u2:
                no_kk_terkait = st.text_input("Nomor Kartu Keluarga (KK):")
                nama_kepala_keluarga = st.text_input("Nama Kepala Keluarga:")

            keterangan_perubahan = st.text_area("Keterangan Perubahan Data / Alasan (Contoh: Penambahan anggota keluarga baru, pemisahan KK, atau penggantian KK lama ke KK baru):")
            
            st.markdown("---")
            st.markdown("<b>📂 Unggah Dokumen Bukti (Scan KK Baru / Surat Pengantar / Dokumen Pendukung):</b>", unsafe_allow_html=True)
            dokumen_kk_up = st.file_uploader("Pilih File Dokumen (Format PDF, JPG, atau PNG):", type=["pdf", "jpg", "jpeg", "png"])

            submit_laporan_kk = st.form_submit_button("📤 Kirim Pengajuan ke Pengurus RW", type="primary")

            if submit_laporan_kk:
                if not nama_pelapor.strip() or not no_kk_terkait.strip() or not nama_kepala_keluarga.strip() or dokumen_kk_up is None:
                    st.error("❌ Mohon lengkapi semua kolom isian dan wajib mengunggah dokumen bukti KK baru!")
                else:
                    folder_pengajuan = "pengajuan_kk"
                    if not os.path.exists(folder_pengajuan):
                        os.makedirs(folder_pengajuan)

                    # Simpan file dokumen yang diunggah
                    nama_file_dokumen = f"{pilih_rt_lapor.replace(' ', '')}_{no_kk_terkait}_{dokumen_kk_up.name}"
                    path_simpan_dok = os.path.join(folder_pengajuan, nama_file_dokumen)
                    with open(path_simpan_dok, "wb") as f:
                        f.write(dokumen_kk_up.getbuffer())

                    # Catat ke dalam file excel rekap pengajuan
                    waktu_lapor = (datetime.utcnow() + timedelta(hours=7)).strftime("%d-%m-%Y %H:%M")
                    new_pengajuan = pd.DataFrame([{
                        "WAKTU": waktu_lapor,
                        "RT": pilih_rt_lapor,
                        "PELAPOR": nama_pelapor.strip(),
                        "NO_KK": no_kk_terkait.strip(),
                        "KEPALA_KELUARGA": nama_kepala_keluarga.strip(),
                        "KETERANGAN": keterangan_perubahan.strip(),
                        "DOKUMEN_FILE": nama_file_dokumen,
                        "STATUS": "Menunggu Verifikasi RW"
                    }])

                    file_excel_pengajuan = "datapengajuankk.xlsx"
                    if os.path.exists(file_excel_pengajuan):
                        df_p_exist = pd.read_excel(file_excel_pengajuan)
                        df_p_exist.columns = df_p_exist.columns.str.strip().str.upper()
                        df_p_updated = pd.concat([df_p_exist, new_pengajuan], ignore_index=True)
                    else:
                        df_p_updated = new_pengajuan

                    df_p_updated.to_excel(file_excel_pengajuan, index=False)
                    st.success("✅ Pengajuan perubahan data dan dokumen KK baru berhasil dikirim ke Pengurus RW!")
                    st.balloons()

        # Tampilkan riwayat pengajuan dokumen di RT tersebut
        st.markdown("---")
        st.markdown("### 📋 Riwayat Pengajuan Pembaruan KK yang Telah Dikirim")
        file_excel_pengajuan = "datapengajuankk.xlsx"
        if os.path.exists(file_excel_pengajuan):
            df_p_tampil = pd.read_excel(file_excel_pengajuan)
            df_p_tampil.columns = df_p_tampil.columns.str.strip().str.upper()
            if not df_p_tampil.empty:
                st.dataframe(df_p_tampil, use_container_width=True, hide_index=True)
                
                # Tombol Download untuk dokumen bukti yang telah dikirim
                st.markdown("#### 📥 Unduh Dokumen Bukti yang Diunggah")
                pilihan_dok_dl = st.selectbox("Pilih Dokumen untuk Diunduh:", options=df_p_tampil["DOKUMEN_FILE"].tolist(), key="sel_dl_dok_bukti")
                if pilihan_dok_dl:
                    path_target_dok = os.path.join("pengajuan_kk", pilihan_dok_dl)
                    if os.path.exists(path_target_dok):
                        with open(path_target_dok, "rb") as f_dok:
                            st.download_button(
                                label=f"📥 Download File: {pilihan_dok_dl}",
                                data=f_dok,
                                file_name=pilihan_dok_dl,
                                mime="application/octet-stream",
                                key="btn_download_dok_bukti"
                            )
            else:
                st.info("ℹ️ Belum ada riwayat pengajuan pembaruan KK.")
        else:
            st.info("ℹ️ Belum ada riwayat pengajuan pembaruan KK.")

# ================= TAB 6: KAS RW =================
with tab6:
    st.subheader("💰 Transparansi Laporan Kas RW 14")
    st.markdown("Berikut adalah ringkasan keuangan, rincian transaksi, serta dokumen PDF & Excel laporan keuangan resmi.")
    
    if not df_kas.empty:
        kolom_pemasukan = [c for c in df_kas.columns if "PEMASUKAN" in c or "MASUK" in c]
        kolom_pengeluaran = [c for c in df_kas.columns if "PENGELUARAN" in c or "KELUAR" in c]
        kolom_saldo = [c for c in df_kas.columns if "SALDO" in c]

        total_masuk = 0
        total_keluar = 0
        
        if kolom_pemasukan:
            df_kas["MASUK_ANGKA"] = pd.to_numeric(df_kas[kolom_pemasukan[0]], errors="coerce").fillna(0)
            total_masuk = df_kas["MASUK_ANGKA"].sum()
        if kolom_pengeluaran:
            df_kas["KELUAR_ANGKA"] = pd.to_numeric(df_kas[kolom_pengeluaran[0]], errors="coerce").fillna(0)
            total_keluar = df_kas["KELUAR_ANGKA"].sum()

        saldo_akhir = total_masuk - total_keluar
        if kolom_saldo and not df_kas[kolom_saldo[0]].dropna().empty:
            last_saldo = pd.to_numeric(df_kas[kolom_saldo[0]], errors="coerce").iloc[-1]
            if pd.notnull(last_saldo): saldo_akhir = last_saldo

        c1, c2, c3 = st.columns(3)
        c1.metric("💵 Total Pemasukan", f"Rp {total_masuk:,.0f}".replace(",", "."))
        c2.metric("💸 Total Pengeluaran", f"Rp {total_keluar:,.0f}".replace(",", "."))
        c3.metric("💰 Saldo Kas Bersih", f"Rp {saldo_akhir:,.0f}".replace(",", "."))
        
        st.write("---")
        st.dataframe(df_kas.drop(columns=["MASUK_ANGKA", "KELUAR_ANGKA"], errors="ignore"), use_container_width=True, hide_index=True)
    else:
        st.info("ℹ️ Belum ada data transaksi kas yang dimasukkan.")
        
    st.write("---")
    st.subheader("📄 Dokumen Laporan Kas Resmi (PDF & Excel)")
    
    folder_pdf_kas = "pdf_kas"
    if not os.path.exists(folder_pdf_kas):
        os.makedirs(folder_pdf_kas)
    daftar_pdf_kas = [f for f in os.listdir(folder_pdf_kas) if f.lower().endswith('.pdf')]
    
    if daftar_pdf_kas:
        for pdf_file in daftar_pdf_kas:
            path_pdf = os.path.join(folder_pdf_kas, pdf_file)
            st.markdown(f"**📂 PDF Kas: {pdf_file}**")
            with open(path_pdf, "rb") as f:
                st.download_button(
                    label=f"📥 Download Dokumen PDF: {pdf_file}",
                    data=f,
                    file_name=pdf_file,
                    mime="application/pdf",
                    key=f"dl_kas_{pdf_file}"
                )
            st.write("---")

    daftar_excel_kas = [f for f in os.listdir() if f.lower().endswith('.xlsx') and ('kas' in f.lower() or 'laporan' in f.lower())]
    if daftar_excel_kas:
        for excel_file in daftar_excel_kas:
            st.markdown(f"**📊 File Excel Kas: {excel_file}**")
            with open(excel_file, "rb") as f:
                st.download_button(
                    label=f"📥 Download File Excel Kas: {excel_file}",
                    data=f,
                    file_name=excel_file,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"dl_excel_kas_{excel_file}"
                )
            st.write("---")
    elif not daftar_pdf_kas:
        st.markdown("*Belum ada file dokumen PDF atau Excel laporan kas tambahan yang diunggah.*")

# ================= TAB 7: INFO & RAPAT =================
with tab7:
    st.subheader("📢 Informasi Kegiatan & Hasil Rapat RW 14")
    st.markdown("Pusat informasi resmi seputar hasil rapat pengurus, pengumuman warga, agenda kegiatan, serta dokumen PDF resmi.")
    
    if not df_info.empty:
        for index, row in df_info.iterrows():
            tgl = row.get("TANGGAL", "Agenda RW")
            judul = row.get("JUDUL", "Informasi Penting")
            isi = row.get("ISI / KATEGORI", "-")
            
            with st.expander(f"📌 [{tgl}] — {judul}"):
                st.write(isi)
    else:
        st.info("ℹ️ Belum ada pengumuman atau hasil rapat yang dipublikasikan.")

    st.write("---")
    st.subheader("📄 Dokumen & Notulen Hasil Rapat (PDF)")
    folder_pdf_info = "pdf_info"
    if not os.path.exists(folder_pdf_info):
        os.makedirs(folder_pdf_info)
    daftar_pdf_info = [f for f in os.listdir(folder_pdf_info) if f.lower().endswith('.pdf')]
    
    if daftar_pdf_info:
        for pdf_file in daftar_pdf_info:
            path_pdf = os.path.join(folder_pdf_info, pdf_file)
            st.markdown(f"**📂 {pdf_file}**")
            with open(path_pdf, "rb") as f:
                st.download_button(
                    label=f"📥 Download Dokumen PDF: {pdf_file}",
                    data=f,
                    file_name=pdf_file,
                    mime="application/pdf",
                    key=f"dl_info_{pdf_file}"
                )
            st.write("---")
    else:
        st.markdown("*Belum ada file PDF hasil rapat yang diunggah oleh pengurus.*")

# ================= TAB 8: GALERI KEGIATAN =================
with tab8:
    st.subheader("🖼️ Galeri Foto Kegiatan Warga RW 14")
    st.markdown("Dokumentasi foto kegiatan warga yang tersusun rapi berdasarkan Nama Hari, Tanggal, Bulan, dan Tahun kegiatan.")
    
    folder_galeri = "galeri"
    if not os.path.exists(folder_galeri):
        os.makedirs(folder_galeri)
        
    daftar_foto = [f for f in os.listdir(folder_galeri) if f.lower().endswith(('png', 'jpg', 'jpeg'))]
    
    if daftar_foto and not df_galeri_meta.empty:
        meta_dict = {}
        for _, row in df_galeri_meta.iterrows():
            f_name = str(row.get("NAMA_FILE", "")).strip()
            meta_dict[f_name] = {
                "hari": str(row.get("HARI", "")),
                "tanggal": str(row.get("TANGGAL", "")),
                "bulan": str(row.get("BULAN", "")),
                "tahun": str(row.get("TAHUN", "")),
                "ket": str(row.get("KETERANGAN", ""))
            }

        galeri_arsip = {}
        
        for nama_foto in daftar_foto:
            if nama_foto in meta_dict:
                h = meta_dict[nama_foto]["hari"]
                t = meta_dict[nama_foto]["tanggal"]
                b = meta_dict[nama_foto]["bulan"]
                th = meta_dict[nama_foto]["tahun"]
                
                if th and b and t:
                    if th not in galeri_arsip:
                        galeri_arsip[th] = {}
                    if b not in galeri_arsip[th]:
                        galeri_arsip[th][b] = {}
                    if t not in galeri_arsip[th][b]:
                        galeri_arsip[th][b][t] = []
                        
                    galeri_arsip[th][b][t].append((nama_foto, h))

        if galeri_arsip:
            for tahun in sorted(galeri_arsip.keys(), reverse=True):
                st.markdown(f"## 📅 Tahun {tahun}")
                for bulan in galeri_arsip[tahun].keys():
                    st.markdown(f"### 🗓️ Bulan: {bulan}")
                    for tanggal in sorted(galeri_arsip[tahun][bulan].keys(), reverse=True):
                        info_item = galeri_arsip[tahun][bulan][tanggal][0]
                        nama_hari_kegiatan = info_item[1]
                        
                        st.markdown(f"**📌 Tanggal Kegiatan: {nama_hari_kegiatan}, {tanggal} {bulan} {tahun}**")
                        
                        foto_list = galeri_arsip[tahun][bulan][tanggal]
                        cols = st.columns(3)
                        for idx, item in enumerate(foto_list):
                            nama_foto = item[0]
                            path_foto = os.path.join(folder_galeri, nama_foto)
                            
                            caption_teks = nama_foto.rsplit('.', 1)[0].replace('_', ' ').title()
                            if nama_foto in meta_dict and meta_dict[nama_foto]["ket"]:
                                caption_teks = meta_dict[nama_foto]["ket"]
                                
                            with cols[idx % 3]:
                                st.image(path_foto, caption=caption_teks, use_container_width=True)
                        st.write("---")
        else:
            st.info("ℹ️ Belum ada foto galeri yang diatur tanggal dan tahunnya oleh Admin.")
    else:
        st.info("ℹ️ Belum ada foto kegiatan di galeri. Pengurus dapat mengunggah dan mengatur jadwalnya melalui menu Admin.")

# ================= TAB 9: SARAN & PENDAPAT (KONSISTEN) =================
with tab9:
    st.subheader("💬 Kotak Aspirasi, Saran, & Pendapat Pengurus RW 14")
    st.markdown("Menu khusus bagi **Ketua RT 01 s.d. 07** serta **Pengurus Inti RW** untuk memberikan masukan, evaluasi, dan pendapat terkait pengembangan Portal RW 14.")

    if st.session_state.get("saran_terverifikasi", False):
        st.success(f"Anda sedang login sebagai: **{st.session_state.get('saran_nama')} ({st.session_state.get('saran_jabatan')})**")
        if st.button("🔄 Keluar / Ganti Akun Pengurus", key="btn_keluar_saran_pengurus"):
            st.session_state["saran_terverifikasi"] = False
            st.rerun()
            
        st.markdown("---")
        with st.form("form_kirim_saran"):
            st.markdown("### ✍️ Tuliskan Saran, Pendapat, atau Evaluasi Portal")
            isi_saran_input = st.text_area("Saran / Pendapat / Masukan Anda untuk Portal RW 14:")
            submit_saran = st.form_submit_button("📤 Kirim Saran & Pendapat", type="primary")
            
            if submit_saran:
                if not isi_saran_input.strip():
                    st.warning("⚠️ Saran atau pendapat tidak boleh kosong.")
                else:
                    waktu_kirim = (datetime.utcnow() + timedelta(hours=7)).strftime("%d-%m-%Y %H:%M")
                    new_saran_row = pd.DataFrame([{
                        "WAKTU": waktu_kirim,
                        "PENGIRIM": st.session_state.get('saran_nama'),
                        "JABATAN": st.session_state.get('saran_jabatan'),
                        "SARAN_PENDAPAT": isi_saran_input.strip()
                    }])
                    
                    if os.path.exists("datasaran.xlsx"):
                        df_s_existing = pd.read_excel("datasaran.xlsx")
                        df_s_existing.columns = df_s_existing.columns.str.strip().str.upper()
                        df_s_updated = pd.concat([df_s_existing, new_saran_row], ignore_index=True)
                    else:
                        df_s_updated = new_saran_row
                        
                    df_s_updated.to_excel("datasaran.xlsx", index=False)
                    st.cache_data.clear()
                    st.success("✅ Saran dan pendapat Anda berhasil dikirim dan tersimpan di sistem!")
                    st.rerun()
    else:
        with st.form("form_login_saran"):
            st.markdown("### 🔑 Verifikasi Identitas Pengurus")
            pilihan_jabatan_pengurus = st.selectbox(
                "Pilih Jabatan:", 
                ["-- Pilih Jabatan --", "Ketua RT 01", "Ketua RT 02", "Ketua RT 03", "Ketua RT 04", "Ketua RT 05", "Ketua RT 06", "Ketua RT 07", "Pengurus Inti (Ketua/Sekretaris/Bendahara)"],
                key="select_jabatan_saran_form"
            )
            nama_pengirim = st.text_input("Nama Lengkap Anda:", key="input_nama_saran_form")
            kode_akses_pengurus = st.text_input("Masukkan Kode Akses Pengurus:", type="password", key="input_kode_saran_form")
            
            submit_verif = st.form_submit_button("🔓 Verifikasi & Masuk")
            
        if submit_verif:
            kode_bersih = kode_akses_pengurus.strip()
            
            if pilihan_jabatan_pengurus == "-- Pilih Jabatan --" or not nama_pengirim.strip() or not kode_bersih:
                st.error("❌ Mohon lengkapi pilihan jabatan, nama, dan kode akses!")
            elif kode_bersih == "@pengurusrw14":
                st.session_state["saran_terverifikasi"] = True
                st.session_state["saran_nama"] = nama_pengirim.strip()
                st.session_state["saran_jabatan"] = pilihan_jabatan_pengurus
                st.success(f"✅ Verifikasi Berhasil! Selamat datang, {nama_pengirim.strip()} ({pilihan_jabatan_pengurus}).")
                st.rerun()
            else:
                st.error("❌ Kode akses pengurus salah!")

    st.markdown("---")
    st.markdown("### 📋 Daftar Saran & Pendapat Pengurus yang Masuk")
    if os.path.exists("datasaran.xlsx"):
        df_s_tampil = pd.read_excel("datasaran.xlsx")
        df_s_tampil.columns = df_s_tampil.columns.str.strip().str.upper()
        if not df_s_tampil.empty:
            for idx, row in df_s_tampil.iloc[::-1].iterrows():
                waktu_s = row.get("WAKTU", "-")
                pengirim_s = row.get("PENGIRIM", "-")
                jabatan_s = row.get("JABATAN", "-")
                pesan_s = row.get("SARAN_PENDAPAT", "-")
                
                st.markdown(f"""
                <div style="background-color: white; padding: 15px; border-radius: 8px; border-left: 5px solid #0D47A1; margin-bottom: 12px; box-shadow: 0px 2px 5px rgba(0,0,0,0.05);">
                    <p style="margin: 0; font-size: 13px; color: #555;"><b>👤 {pengirim_s}</b> ({jabatan_s}) &bull; <i>🕒 {waktu_s} WIB</i></p>
                    <p style="margin: 8px 0 0 0; font-size: 15px; color: #222; white-space: pre-wrap;">{pesan_s}</p>
                </div>
                """, unsafe_allow_html=True)
        else:
            st.info("ℹ️ Belum ada saran atau pendapat yang dikirimkan oleh pengurus.")
    else:
        st.info("ℹ️ Belum ada catatan saran pengurus.")

# ================= TAB 10: EDIT & UPLOAD ADMIN (KONSISTEN & AMAN) =================
with tab10:
    st.subheader("⚙️ Panel Pengaturan & Unggah Dokumen (Admin)")
    
    if admin_terverifikasi:
        st.success("✅ Status Admin Aktif. Silakan pilih menu pengelolaan di bawah ini:")
        menu_admin = st.selectbox(
            "Pilih Menu Pengelolaan:", 
            ["Data Warga", "Struktur Organisasi", "Laporan Kas RW", "Laporan Kas Pemakaman/Sosial", "Informasi & Hasil Rapat", "Upload Gambar Struktur RW", "Upload Foto Pengurus Inti", "Upload Foto Ketua RT", "Upload File PDF (Kas & Rapat)", "Upload Foto Galeri", "Hapus Foto Galeri"],
            key="selectbox_menu_admin_utama"
        )
        
        if menu_admin == "Data Warga":
            st.markdown("💡 *Pada kolom **Status Penduduk**, Anda dapat memilih **Tetap** atau **Musiman**.*")
            data_terbaru = st.data_editor(
                df.drop(columns=["RT_FORMAT"], errors="ignore"), 
                num_rows="dynamic", 
                use_container_width=True,
                column_config={
                    "STATUS PENDUDUK": st.column_config.SelectboxColumn(
                        "Status Penduduk",
                        help="Pilih status kependudukan warga",
                        options=["Tetap", "Musiman"],
                        required=True
                    )
                }
            )
            if st.button("💾 Simpan Perubahan Data Warga", type="primary", key="btn_simpan_warga_admin"):
                try:
                    data_terbaru.to_excel("datawarga.xlsx", index=False)
                    st.cache_data.clear()
                    st.success("✅ Data warga berhasil disimpan!")
                    st.rerun()
                except Exception as e:
                    st.error(f"❌ Gagal menyimpan data: {e}")
                    
        elif menu_admin == "Struktur Organisasi":
            struk_terbaru = st.data_editor(df_struktur, num_rows="dynamic", use_container_width=True)
            if st.button("💾 Simpan Perubahan Struktur Organisasi", type="primary", key="btn_simpan_struk_admin"):
                try:
                    struk_terbaru.to_excel("datastruktur.xlsx", index=False)
                    st.cache_data.clear()
                    st.success("✅ Struktur organisasi berhasil disimpan!")
                    st.rerun()
                except Exception as e:
                    st.error(f"❌ Gagal menyimpan struktur: {e}")
                    
        elif menu_admin == "Laporan Kas RW":
            kas_terbaru = st.data_editor(df_kas.drop(columns=["MASUK_ANGKA", "KELUAR_ANGKA"], errors="ignore"), num_rows="dynamic", use_container_width=True)
            if st.button("💾 Simpan Perubahan Kas RW", type="primary", key="btn_simpan_kas_admin"):
                try:
                    kas_terbaru.to_excel("datakas.xlsx", index=False)
                    st.cache_data.clear()
                    st.success("✅ Laporan Kas RW berhasil disimpan!")
                    st.rerun()
                except Exception as e:
                    st.error(f"❌ Gagal menyimpan kas: {e}")

        elif menu_admin == "Laporan Kas Pemakaman/Sosial":
            kp_terbaru = st.data_editor(df_kas_pemakaman.drop(columns=["MASUK_ANGKA", "KELUAR_ANGKA"], errors="ignore"), num_rows="dynamic", use_container_width=True)
            if st.button("💾 Simpan Perubahan Kas Pemakaman", type="primary", key="btn_simpan_kp_admin"):
                try:
                    kp_terbaru.to_excel("datakaspemakaman.xlsx", index=False)
                    st.cache_data.clear()
                    st.success("✅ Laporan Kas Pemakaman berhasil disimpan!")
                    st.rerun()
                except Exception as e:
                    st.error(f"❌ Gagal menyimpan kas pemakaman: {e}")
                    
        elif menu_admin == "Informasi & Hasil Rapat":
            info_terbaru = st.data_editor(df_info, num_rows="dynamic", use_container_width=True)
            if st.button("💾 Simpan Perubahan Informasi", type="primary", key="btn_simpan_info_admin"):
                try:
                    info_terbaru.to_excel("datainfo.xlsx", index=False)
                    st.cache_data.clear()
                    st.success("✅ Informasi kegiatan berhasil disimpan!")
                    st.rerun()
                except Exception as e:
                    st.error(f"❌ Gagal menyimpan informasi: {e}")

        elif menu_admin == "Upload Gambar Struktur RW":
            st.markdown("### 🖼️ Unggah Gambar Bagan Struktur Organisasi RW")
            struktur_img_up = st.file_uploader("Pilih File Gambar Struktur (JPG/PNG)", type=["jpg", "jpeg", "png"], key="up_str_img")
            if struktur_img_up is not None:
                path_simpan_str = "struktur_rw.jpg"
                with open(path_simpan_str, "wb") as f:
                    f.write(struktur_img_up.getbuffer())
                st.success("✅ Gambar struktur organisasi RW berhasil diunggah!")
                st.rerun()

        elif menu_admin == "Upload Foto Pengurus Inti":
            st.markdown("### 📸 Unggah Foto Pengurus Inti (Ketua RW, Sekretaris, Bendahara)")
            pilih_posisi = st.selectbox("Pilih Jabatan Pengurus:", ["Ketua RW", "Sekretaris", "Bendahara"], key="sel_pos_pengurus")
            
            mapping_nama = {"Ketua RW": "ketuarw", "Sekretaris": "sekretaris", "Bendahara": "bendahara"}
            file_key = mapping_nama[pilih_posisi]
            
            foto_pengurus_up = st.file_uploader(f"Pilih Foto untuk {pilih_posisi} (JPG/PNG)", type=["jpg", "jpeg", "png"], key="up_foto_pengurus")
            if foto_pengurus_up is not None:
                folder_pengurus_dir = "pengurus"
                if not os.path.exists(folder_pengurus_dir):
                    os.makedirs(folder_pengurus_dir)
                
                path_simpan_p = os.path.join(folder_pengurus_dir, f"{file_key}.jpg")
                with open(path_simpan_p, "wb") as f:
                    f.write(foto_pengurus_up.getbuffer())
                st.success(f"✅ Foto {pilih_posisi} berhasil diunggah dan disimpan!")
                st.rerun()
                    
        elif menu_admin == "Upload Foto Ketua RT":
            st.markdown("### 📸 Unggah Foto Profil Ketua RT")
            st.markdown("Pilih nomor RT dan unggah foto profil resminya:")
            
            pilih_rt_upload = st.selectbox("Pilih RT untuk Foto:", ["RT 01", "RT 02", "RT 03", "RT 04", "RT 05", "RT 06", "RT 07", "RT 08", "RT 09", "RT 10"], key="sel_rt_foto")
            rt_num_up = ''.join(filter(str.isdigit, pilih_rt_upload))
            
            foto_rt_upload = st.file_uploader(f"Pilih Foto untuk {pilih_rt_upload} (JPG/PNG)", type=["jpg", "jpeg", "png"], key="up_foto_rt_file")
            
            if foto_rt_upload is not None:
                folder_rt_dir = "rt"
                if not os.path.exists(folder_rt_dir):
                    os.makedirs(folder_rt_dir)
                
                nama_file_simpan = f"rt{rt_num_up}.jpg"
                path_simpan_rt = os.path.join(folder_rt_dir, nama_file_simpan)
                
                with open(path_simpan_rt, "wb") as f:
                    f.write(foto_rt_upload.getbuffer())
                
                st.success(f"✅ Foto untuk {pilih_rt_upload} berhasil diunggah dan disimpan!")
                st.rerun()
                    
        elif menu_admin == "Upload File PDF (Kas & Rapat)":
            st.markdown("Unggah dokumen resmi berformat PDF untuk Warga:")
            kategori_pdf = st.radio("Pilih Kategori Dokumen PDF:", ["Laporan Kas RW", "Hasil Rapat / Informasi RW"], key="radio_kat_pdf")
            pdf_upload = st.file_uploader("Pilih File PDF", type=["pdf"], key="up_pdf_file_rw")
            
            if pdf_upload is not None:
                if kategori_pdf == "Laporan Kas RW":
                    folder_tujuan = "pdf_kas"
                else:
                    folder_tujuan = "pdf_info"
                    
                if not os.path.exists(folder_tujuan):
                    os.makedirs(folder_tujuan)
                    
                path_simpan_pdf = os.path.join(folder_tujuan, pdf_upload.name)
                with open(path_simpan_pdf, "wb") as f:
                    f.write(pdf_upload.getbuffer())
                st.success(f"✅ Dokumen PDF '{pdf_upload.name}' berhasil diunggah!")
                st.rerun()
                    
        elif menu_admin == "Upload Foto Galeri":
            st.markdown("📦 **Upload Foto Kegiatan ke Galeri & Atur Jadwal Otomatis**")
            foto_upload = st.file_uploader("Pilih File Foto (JPG/PNG)", type=["jpg", "jpeg", "png"], key="up_foto_galeri_file")
            
            if foto_upload is not None:
                col_i1, col_i2 = st.columns(2)
                with col_i1:
                    input_tgl = st.selectbox("Pilih Tanggal:", ["-- Pilih --"] + [f"{i:02d}" for i in range(1, 32)], key="sel_tgl_galeri")
                    input_bln = st.selectbox("Pilih Bulan:", ["-- Pilih --", "Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"], key="sel_bln_galeri")
                with col_i2:
                    input_thn = st.selectbox("Pilih Tahun:", ["-- Pilih --", "2024", "2025", "2026", "2027", "2028"], key="sel_thn_galeri")
                    input_ket = st.text_input("Keterangan / Nama Kegiatan:", key="input_ket_galeri")
                
                bulan_to_angka = {
                    'Januari': 1, 'Februari': 2, 'Maret': 3, 'April': 4,
                    'Mei': 5, 'Juni': 6, 'Juli': 7, 'Agustus': 8,
                    'September': 9, 'Oktober': 10, 'November': 11, 'Desember': 12
                }
                
                input_hari = ""
                if input_tgl != "-- Pilih --" and input_bln != "-- Pilih --" and input_thn != "-- Pilih --":
                    try:
                        dt_pilih = datetime(int(input_thn), bulan_to_angka[input_bln], int(input_tgl))
                        hari_inggris = dt_pilih.strftime("%A")
                        hari_indo_map = {
                            'Monday': 'Senin', 'Tuesday': 'Selasa', 'Wednesday': 'Rabu',
                            'Thursday': 'Kamis', 'Friday': 'Jumat', 'Saturday': 'Sabtu', 'Sunday': 'Minggu'
                        }
                        input_hari = hari_indo_map.get(hari_inggris, "")
                        st.info(f"✨ Hari Terdeteksi Otomatis: **{input_hari}**")
                    except Exception:
                        st.warning("⚠️ Kombinasi tanggal yang dipilih tidak valid.")

                if st.button("💾 Simpan Foto & Atur Jadwal", type="primary", key="btn_simpan_galeri_admin"):
                    if input_tgl == "-- Pilih --" or input_bln == "-- Pilih --" or input_thn == "-- Pilih --" or not input_hari:
                        st.error("❌ Mohon pilih Tanggal, Bulan, dan Tahun dengan kombinasi yang valid!")
                    else:
                        folder_galeri = "galeri"
                        if not os.path.exists(folder_galeri):
                            os.makedirs(folder_galeri)
                        
                        path_simpan = os.path.join(folder_galeri, foto_upload.name)
                        with open(path_simpan, "wb") as f:
                            f.write(foto_upload.getbuffer())
                            
                        new_row = pd.DataFrame([{
                            "NAMA_FILE": foto_upload.name,
                            "HARI": input_hari,
                            "TANGGAL": input_tgl,
                            "BULAN": input_bln,
                            "TAHUN": input_thn,
                            "KETERANGAN": input_ket
                        }])
                        
                        if os.path.exists("datagaleri.xlsx"):
                            df_g_existing = pd.read_excel("datagaleri.xlsx")
                            df_g_existing.columns = df_g_existing.columns.str.strip().str.upper()
                            df_g_existing = df_g_existing[df_g_existing["NAMA_FILE"] != foto_upload.name]
                            df_g_updated = pd.concat([df_g_existing, new_row], ignore_index=True)
                        else:
                            df_g_updated = new_row
                            
                        df_g_updated.to_excel("datagaleri.xlsx", index=False)
                        st.cache_data.clear()
                        st.success(f"✅ Foto '{foto_upload.name}' berhasil diunggah dengan jadwal {input_hari}, {input_tgl}-{input_bln}-{input_thn}!")
                        st.rerun()

        elif menu_admin == "Hapus Foto Galeri":
            st.markdown("🗑️ **Kelola & Hapus Foto Galeri yang Tidak Diperlukan**")
            folder_galeri = "galeri"
            if os.path.exists(folder_galeri):
                daftar_foto_del = [f for f in os.listdir(folder_galeri) if f.lower().endswith(('png', 'jpg', 'jpeg'))]
                if daftar_foto_del:
                    foto_pilih_hapus = st.selectbox("Pilih Foto yang Ingin Dihapus:", options=daftar_foto_del, key="sel_hapus_galeri_foto")
                    
                    path_preview = os.path.join(folder_galeri, foto_pilih_hapus)
                    st.image(path_preview, width=300, caption=f"Preview: {foto_pilih_hapus}")
                    
                    if st.button("🗑️ Hapus Foto Ini Permanen", type="primary", key="btn_konfirmasi_hapus_galeri"):
                        try:
                            os.remove(path_preview)
                            if os.path.exists("datagaleri.xlsx"):
                                df_g_del = pd.read_excel("datagaleri.xlsx")
                                df_g_del.columns = df_g_del.columns.str.strip().str.upper()
                                df_g_del = df_g_del[df_g_del["NAMA_FILE"] != foto_pilih_hapus]
                                df_g_del.to_excel("datagaleri.xlsx", index=False)
                                
                            st.cache_data.clear()
                            st.success(f"✅ Foto '{foto_pilih_hapus}' berhasil dihapus secara permanen!")
                            st.rerun()
                        except Exception as e:
                            st.error(f"❌ Gagal menghapus foto: {e}")
                else:
                    st.info("ℹ️ Belum ada foto di dalam folder galeri.")
            else:
                st.info("ℹ️ Folder galeri belum tersedia.")
    else:
        st.warning("⚠️ **Akses Terbatas:** Silakan masukkan **Password Admin** yang benar di panel sidebar sebelah kiri untuk membuka menu pengaturan dan unggah dokumen ini.")

# Penutup blok div utama biru muda
st.markdown("</div>", unsafe_allow_html=True)

# ================= FOOTER PORTAL RESMI =================
st.markdown("""
<div style="text-align: center; padding: 20px; color: #555; font-size: 14px; border-top: 1px solid #ddd; margin-top: 30px;">
    <p style="margin: 0; font-weight: bold;">Portal Resmi RW 14 Perum Griya Permata Raya Desa Nanjung Mekar Kec. Rancaekek Kab. Bandung</p>
    <p style="margin: 5px 0 0 0; font-size: 12px; color: #777;">Dikelola oleh Pengurus RW 14 &bull; Didukung oleh Sistem Dashboard Digital Warga</p>
</div>
""", unsafe_allow_html=True)