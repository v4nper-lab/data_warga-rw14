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
    return df

@st.cache_data
def load_kas():
    if os.path.exists("datakas.xlsx"):
        df_kas = pd.read_excel("datakas.xlsx")
        df_kas.columns = df_kas.columns.str.strip().str.upper()
        return df_kas
    else:
        return pd.DataFrame(columns=["TANGGAL", "KETERANGAN", "JENIS", "JUMLAH"])

@st.cache_data
def load_info():
    if os.path.exists("datainfo.xlsx"):
        df_info = pd.read_excel("datainfo.xlsx")
        df_info.columns = df_info.columns.str.strip().str.upper()
        return df_info
    else:
        return pd.DataFrame(columns=["TANGGAL", "JUDUL", "ISI / KATEGORI"])

@st.cache_data
def load_struktur():
    if os.path.exists("datastruktur.xlsx"):
        df_struk = pd.read_excel("datastruktur.xlsx")
        df_struk.columns = df_struk.columns.str.strip().str.upper()
        return df_struk
    else:
        data_awal = {
            "JABATAN": ["Ketua RW 14", "Sekretaris", "Bendahara", "Keamanan & Ketertiban", "Pembangunan & Lingkungan", "Olahraga", "Sosial & Pemakaman", "Seni Budaya & Pemuda"],
            "NAMA PENGURUS": ["Triyadi Sucipto", "Irvan Permana", "Aan Toni Fauyi", "Dedi, Uus, Ali, Tiktik", "E. Rustandi, Nahnu, Dahlan, Sugiyanto, Mulyono", "Mulyana, Ateng, Fajar, Kris, Apeng, Mulyadi", "Ust. Nanang, E. Rustandi, Ust. Juhendi, Shulton, Edi, Baryanto", "Uwa Tia, Ridwan S, Hary"],
            "KONTAK / HP": ["0812xxxxxxxx", "0812xxxxxxxx", "0812xxxxxxxx", "-", "-", "-", "-", "-"]
        }
        return pd.DataFrame(data_awal)

df = load_data()
df_kas = load_kas()
df_info = load_info()
df_struktur = load_struktur()

# ================= WAKTU REAL-TIME DI SIDEBAR =================
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

st.sidebar.header("🛠️ Panel Filter Data RT")
if "RT" in df.columns:
    df["RT_FORMAT"] = df["RT"].apply(lambda x: f"RT{int(x):02d}" if pd.notnull(x) and str(x).isdigit() else f"RT{str(x)}")
    semua_rt_format = sorted(df["RT_FORMAT"].dropna().unique(), key=lambda x: int(''.join(filter(str.isdigit, x)) or 0))
    pilihan_rt_format = st.sidebar.multiselect("Tampilkan Data RT:", options=semua_rt_format, default=semua_rt_format)
    pilihan_rt_format = sorted(pilihan_rt_format, key=lambda x: int(''.join(filter(str.isdigit, x)) or 0))

    # ================= FOTO KETUA RT FORMAT VERTIKAL DI SIDEBAR =================
    st.sidebar.markdown("---")
    st.sidebar.markdown("<p style='font-weight: bold; color: #0D47A1; margin-bottom: 10px; font-size: 15px;'>👨‍✈️ Profil Ketua RT Terpilih:</p>", unsafe_allow_html=True)
    
    folder_foto_rt = "rt"
    if not os.path.exists(folder_foto_rt):
        os.makedirs(folder_foto_rt)

    for rt_pilih in pilihan_rt_format:
        rt_num = ''.join(filter(str.isdigit, rt_pilih))
        
        path_foto = None
        kemungkinan_nama = [
            os.path.join(folder_foto_rt, f"rt{rt_num}.jpg"), os.path.join(folder_foto_rt, f"rt{rt_num}.jpeg"), os.path.join(folder_foto_rt, f"rt{rt_num}.png"),
            os.path.join(folder_foto_rt, f"rt0{rt_num}.jpg"), os.path.join(folder_foto_rt, f"rt0{rt_num}.png"),
            f"rt{rt_num}.jpg", f"rt{rt_num}.png", f"rt0{rt_num}.jpg", f"rt0{rt_num}.png"
        ]
        
        for lokasi_file in kemungkinan_nama:
            if os.path.exists(lokasi_file):
                path_foto = lokasi_file
                break

        nama_ketua = f"Ketua {rt_pilih}"
        if not df_struktur.empty:
            for _, row in df_struktur.iterrows():
                jabatan_str = str(row.get("JABATAN", "")).upper()
                if f"RT {rt_num}" in jabatan_str or f"RT0{rt_num}" in jabatan_str or f"RT {rt_pilih}" in jabatan_str or f"RT{rt_num}" in jabatan_str:
                    nama_pengurus_val = row.get("NAMA PENGURUS", "")
                    if pd.notnull(nama_pengurus_val) and str(nama_pengurus_val).strip() != "":
                        nama_ketua = str(nama_pengurus_val).strip()
                        break

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
else:
    st.error("Kolom 'RT' tidak ditemukan di Excel.")
    st.stop()

# ================= KONTROL KEAMANAN ADMIN (LANGSUNG PASSWORD) =================
st.sidebar.markdown("---")
st.sidebar.subheader("🔐 Menu Pengurus (Admin)")
password_input = st.sidebar.text_input("Masukkan Password Admin:", type="password")

admin_terverifikasi = False
if password_input == "V@nadminrw14":
    admin_terverifikasi = True
    st.sidebar.success("✅ Login Admin Berhasil!")
elif password_input != "":
    st.sidebar.error("❌ Password salah!")

# ================= KONTEN UTAMA PORTAL =================
st.markdown("""
<div style="background: linear-gradient(135deg, #E3F2FD, #BBDEFB); padding: 25px; border-radius: 20px; box-shadow: 0px 6px 15px rgba(0,0,0,0.08); margin-bottom: 25px; border: 2px solid #90CAF9;">
""", unsafe_allow_html=True)

# Teks Berjalan Motivasi RT
st.markdown("""
<div style="background-color: #ffffff; padding: 8px 12px; border-radius: 8px; border: 1px solid #90CAF9; margin-bottom: 15px; box-shadow: inset 0px 1px 3px rgba(0,0,0,0.05);">
    <marquee behavior="scroll" direction="left" scrollamount="5" style="color: #0D47A1; font-weight: bold; font-size: 15px;">
        🏡 Kepada seluruh Ketua RT RW 14 &nbsp;&bull;&nbsp; Mengurus data warga hari ini adalah investasi kemudahan untuk urusan sosial kemasyarakatan di masa depan &nbsp;&bull;&nbsp; Semangat terus melayani warga dengan sepenuh hati! ❤️
    </marquee>
</div>
""", unsafe_allow_html=True)

col_logo, col_teks = st.columns([1, 6])
with col_logo:
    st.image(sumber_logo, width=80)
with col_teks:
    st.markdown("<h2 style='color: #0D47A1; font-weight: 900; margin: 0; padding-top: 5px; text-shadow: 1px 1px 2px rgba(255,255,255,0.8); font-size: 22px;'>Portal Resmi & Dashboard Warga RW 14 Perum Griya Permata Raya Desa Nanjung Mekar Kec. Rancaekek Kab. Bandung</h2>", unsafe_allow_html=True)
    st.markdown("<p style='color: #333; font-weight: bold; margin: 5px 0 0 0;'>Pusat Layanan Informasi, Kependudukan, dan Transparansi Keuangan Lingkungan</p>", unsafe_allow_html=True)

st.write("---")

# ================= MENU UTAMA WEBSITE PORTAL =================
if admin_terverifikasi:
    tab0, tab_struk, tab1, tab2, tab3, tab4, tab5, tab6, tab7, tab8, tab9 = st.tabs([
        "🏠 Beranda", "👥 Struktur", "📋 Statistik", "👫 Demografi", "🎓 Pendidikan", 
        "🗂️ Data Warga", "🔍 Cari KK", "💰 Kas RW", 
        "📢 Info & Rapat", "🖼️ Galeri", "⚙️ Edit & Upload (Admin)"
    ])
else:
    tab0, tab_struk, tab1, tab2, tab3, tab4, tab5, tab6, tab7, tab8 = st.tabs([
        "🏠 Beranda", "👥 Struktur", "📋 Statistik", "👫 Demografi", "🎓 Pendidikan", 
        "🗂️ Data Warga", "🔍 Cari KK", "💰 Kas RW", 
        "📢 Info & Rapat", "🖼️ Galeri"
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
        * Memantau transparansi laporan keuangan kas RW (lengkap dengan dokumen PDF resmi).
        * Membaca hasil rapat, agenda kegiatan, dokumen PDF resmi, serta galeri foto lingkungan.
        
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
                jab = str(row.get("JABATAN", "")).upper()
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
                img_rw = muat_dan_seragamkan_foto(foto_rw, ukuran=(300, 360))
                st.image(img_rw, use_container_width=True)
            else:
                st.image("https://cdn-icons-png.flaticon.com/512/3135/3135673.png", use_container_width=True)
            st.markdown(f"<div style='text-align: center; font-weight: bold; color: #0D47A1; margin-top: 5px;'>{nama_rw}<br><span style='font-size: 12px; color: #555;'>Ketua RW 14</span></div>", unsafe_allow_html=True)

        with col_pengurus2:
            foto_sek = cari_foto_pengurus("sekretaris")
            if foto_sek:
                img_sek = muat_dan_seragamkan_foto(foto_sek, ukuran=(300, 360))
                st.image(img_sek, use_container_width=True)
            else:
                st.image("https://cdn-icons-png.flaticon.com/512/3135/3135673.png", use_container_width=True)
            st.markdown(f"<div style='text-align: center; font-weight: bold; color: #0D47A1; margin-top: 5px;'>{nama_sek}<br><span style='font-size: 12px; color: #555;'>Sekretaris</span></div>", unsafe_allow_html=True)

        with col_pengurus3:
            foto_bend = cari_foto_pengurus("bendahara")
            if foto_bend:
                img_bend = muat_dan_seragamkan_foto(foto_bend, ukuran=(300, 360))
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

# ================= TAB STRUKTUR ORGANISASI (DENGAN GAMBAR ORGANIGRAM) =================
with tab_struk:
    st.subheader("👥 Bagan Struktur Organisasi Pengurus RW 14")
    st.markdown("Bagan organigram kepengurusan Rukun Warga (RW) 14 Perum Griya Permata Raya Periode 2024 - 2029.")
    
    path_struktur_img = "struktur_rw.jpg"
    if not os.path.exists(path_struktur_img):
        path_struktur_img = "struktur_rw.png"
        
    if os.path.exists(path_struktur_img):
        st.image(path_struktur_img, caption="Struktur Pengurus RW 014 Griya Permata Raya Periode 2024 - 2029", use_container_width=True)
    else:
        st.info("ℹ️ File gambar struktur belum diunggah. Silakan upload file gambar dengan nama 'struktur_rw.jpg' ke folder utama project atau via menu Admin.")
    
    st.write("---")
    st.markdown("### 📋 Rincian Tabel Pengurus")
    if not df_struktur.empty:
        st.dataframe(df_struktur, use_container_width=True, hide_index=True)
    else:
        st.info("ℹ️ Belum ada data struktur pengurus yang dimasukkan.")

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

# ================= TAB 2: DEMOGRAFI =================
with tab2:
    col_a, col_b, col_c = st.columns(3)
    with col_a:
        st.subheader("Jenis Kelamin")
        if "JENIS KELAMIN" in df_filtered.columns:
            fig_jk = px.pie(df_filtered, names="JENIS KELAMIN", hole=0.5, color_discrete_sequence=['#66b3ff','#ff9999'])
            fig_jk.update_layout(paper_bgcolor="rgba(0,0,0,0)", font=dict(size=16))
            fig_jk.update_traces(textfont_size=18)
            st.plotly_chart(fig_jk, use_container_width=True)
    with col_b:
        st.subheader("Sebaran Agama")
        if "AGAMA" in df_filtered.columns:
            fig_agama = px.pie(df_filtered, names="AGAMA", hole=0.5, color_discrete_sequence=px.colors.qualitative.Set3)
            fig_agama.update_layout(paper_bgcolor="rgba(0,0,0,0)", font=dict(size=16), legend=dict(font=dict(size=14)))
            fig_agama.update_traces(textfont_size=18)
            st.plotly_chart(fig_agama, use_container_width=True)
    with col_c:
        st.subheader("Status Perkawinan")
        if "STATUS PERKAWINAN" in df_filtered.columns:
            df_status = df_filtered["STATUS PERKAWINAN"].astype(str).str.title().value_counts().reset_index()
            df_status.columns = ["Status", "Jumlah"]
            fig_status = px.bar(df_status, x="Status", y="Jumlah", color="Status", text_auto=True, color_discrete_sequence=px.colors.qualitative.Set2)
            fig_status.update_layout(paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", showlegend=False, font=dict(size=16))
            fig_status.update_traces(textfont_size=18, textangle=0)
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

# ================= TAB 4: SEMUA DATA =================
with tab4:
    st.subheader("Tabel Seluruh Warga")
    st.markdown("💡 *Data sensitif (NIK & No. KK) disembunyikan untuk keamanan publik.*")
    kolom_dibuang = ["NO. KK", "NIK", "USIA_ANGKA", "Kelompok Usia", "RT_FORMAT"]
    df_tabel = df_filtered.drop(columns=kolom_dibuang, errors="ignore")
    st.dataframe(df_tabel, use_container_width=True, hide_index=True)

# ================= TAB 5: PENCARIAN KK =================
with tab5:
    st.subheader("🔍 Pencarian & Data per Kartu Keluarga (KK)")
    st.markdown("Ketik nama salah satu warga untuk melihat seluruh anggota keluarganya secara akurat.")
    
    kata_kunci = st.text_input("🔎 Masukkan Nama Warga:")
    
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

# ================= TAB 6: KAS RW =================
with tab6:
    st.subheader("💰 Transparansi Laporan Kas RW 14")
    st.markdown("Berikut adalah ringkasan keuangan, rincian transaksi, serta dokumen PDF laporan keuangan resmi.")
    
    if not df_kas.empty and "JUMLAH" in df_kas.columns and "JENIS" in df_kas.columns:
        df_kas["JUMLAH_ANGKA"] = pd.to_numeric(df_kas["JUMLAH"], errors="coerce").fillna(0)
        total_masuk = df_kas[df_kas["JENIS"].astype(str).str.upper().str.contains("MASUK", na=False)]["JUMLAH_ANGKA"].sum()
        total_keluar = df_kas[df_kas["JENIS"].astype(str).str.upper().str.contains("KELUAR", na=False)]["JUMLAH_ANGKA"].sum()
        saldo_akhir = total_masuk - total_keluar
        
        c1, c2, c3 = st.columns(3)
        c1.metric("💵 Total Pemasukan", f"Rp {total_masuk:,.0f}".replace(",", "."))
        c2.metric("💸 Total Pengeluaran", f"Rp {total_keluar:,.0f}".replace(",", "."))
        c3.metric("💰 Saldo Kas Bersih", f"Rp {saldo_akhir:,.0f}".replace(",", "."))
        
        st.write("---")
        st.dataframe(df_kas.drop(columns=["JUMLAH_ANGKA"], errors="ignore"), use_container_width=True, hide_index=True)
    else:
        st.info("ℹ️ Belum ada data transaksi kas yang dimasukkan.")
        
    st.write("---")
    st.subheader("📄 Dokumen Laporan Kas Resmi (PDF)")
    folder_pdf_kas = "pdf_kas"
    if not os.path.exists(folder_pdf_kas):
        os.makedirs(folder_pdf_kas)
    daftar_pdf_kas = [f for f in os.listdir(folder_pdf_kas) if f.lower().endswith('.pdf')]
    
    if daftar_pdf_kas:
        for pdf_file in daftar_pdf_kas:
            path_pdf = os.path.join(folder_pdf_kas, pdf_file)
            st.markdown(f"**📂 {pdf_file}**")
            with open(path_pdf, "rb") as f:
                st.download_button(
                    label=f"📥 Download & Lihat Dokumen: {pdf_file}",
                    data=f,
                    file_name=pdf_file,
                    mime="application/pdf",
                    key=f"dl_kas_{pdf_file}"
                )
            st.write("---")
    else:
        st.markdown("*Belum ada file PDF laporan kas yang diunggah oleh pengurus.*")

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
                    label=f"📥 Download & Lihat Dokumen: {pdf_file}",
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
    st.markdown("Dokumentasi foto kegiatan warga, kerja bakti, posyandu, dan acara kebersamaan di lingkungan Perum Griya Permata Raya.")
    
    folder_galeri = "galeri"
    if not os.path.exists(folder_galeri):
        os.makedirs(folder_galeri)
        
    daftar_foto = [f for f in os.listdir(folder_galeri) if f.lower().endswith(('png', 'jpg', 'jpeg'))]
    
    if daftar_foto:
        cols = st.columns(3)
        for idx, nama_foto in enumerate(daftar_foto):
            path_foto = os.path.join(folder_galeri, nama_foto)
            with cols[idx % 3]:
                st.image(path_foto, caption=nama_foto.rsplit('.', 1)[0].replace('_', ' ').title(), use_container_width=True)
    else:
        st.info("ℹ️ Belum ada foto kegiatan di galeri. Pengurus dapat mengunggah foto melalui menu Admin.")

# ================= TAB 9: EDIT & UPLOAD (HANYA ADMIN) =================
if admin_terverifikasi:
    with tab9:
        st.subheader("⚙️ Panel Pengaturan & Unggah Dokumen (Admin)")
        st.warning("⚠️ Anda berada dalam mode Admin. Anda dapat mengelola data warga, struktur, kas, informasi, galeri foto, foto pengurus, hingga mengunggah file PDF.")
        
        menu_admin = st.selectbox(
            "Pilih Menu Pengelolaan:", 
            ["Data Warga", "Struktur Organisasi", "Laporan Kas RW", "Informasi & Hasil Rapat", "Upload Gambar Struktur RW", "Upload Foto Pengurus Inti", "Upload Foto Ketua RT", "Upload File PDF (Kas & Rapat)", "Upload Foto Galeri"]
        )
        
        if menu_admin == "Data Warga":
            data_terbaru = st.data_editor(df.drop(columns=["RT_FORMAT"], errors="ignore"), num_rows="dynamic", use_container_width=True)
            if st.button("💾 Simpan Perubahan Data Warga", type="primary"):
                try:
                    data_terbaru.to_excel("datawarga.xlsx", index=False)
                    st.cache_data.clear()
                    st.success("✅ Data warga berhasil disimpan!")
                    st.rerun()
                except Exception as e:
                    st.error(f"❌ Gagal menyimpan data: {e}")
                    
        elif menu_admin == "Struktur Organisasi":
            struk_terbaru = st.data_editor(df_struktur, num_rows="dynamic", use_container_width=True)
            if st.button("💾 Simpan Perubahan Struktur Organisasi", type="primary"):
                try:
                    struk_terbaru.to_excel("datastruktur.xlsx", index=False)
                    st.cache_data.clear()
                    st.success("✅ Struktur organisasi berhasil disimpan!")
                    st.rerun()
                except Exception as e:
                    st.error(f"❌ Gagal menyimpan struktur: {e}")
                    
        elif menu_admin == "Laporan Kas RW":
            kas_terbaru = st.data_editor(df_kas.drop(columns=["JUMLAH_ANGKA"], errors="ignore"), num_rows="dynamic", use_container_width=True)
            if st.button("💾 Simpan Perubahan Kas RW", type="primary"):
                try:
                    kas_terbaru.to_excel("datakas.xlsx", index=False)
                    st.cache_data.clear()
                    st.success("✅ Laporan Kas RW berhasil disimpan!")
                    st.rerun()
                except Exception as e:
                    st.error(f"❌ Gagal menyimpan kas: {e}")
                    
        elif menu_admin == "Informasi & Hasil Rapat":
            info_terbaru = st.data_editor(df_info, num_rows="dynamic", use_container_width=True)
            if st.button("💾 Simpan Perubahan Informasi", type="primary"):
                try:
                    info_terbaru.to_excel("datainfo.xlsx", index=False)
                    st.cache_data.clear()
                    st.success("✅ Informasi kegiatan berhasil disimpan!")
                    st.rerun()
                except Exception as e:
                    st.error(f"❌ Gagal menyimpan informasi: {e}")

        elif menu_admin == "Upload Gambar Struktur RW":
            st.markdown("### 🖼️ Unggah Gambar Bagan Struktur Organisasi RW")
            struktur_img_up = st.file_uploader("Pilih File Gambar Struktur (JPG/PNG)", type=["jpg", "jpeg", "png"])
            if struktur_img_up is not None:
                path_simpan_str = "struktur_rw.jpg"
                with open(path_simpan_str, "wb") as f:
                    f.write(struktur_img_up.getbuffer())
                st.success("✅ Gambar struktur organisasi RW berhasil diunggah!")
                st.rerun()

        elif menu_admin == "Upload Foto Pengurus Inti":
            st.markdown("### 📸 Unggah Foto Pengurus Inti (Ketua RW, Sekretaris, Bendahara)")
            pilih_posisi = st.selectbox("Pilih Jabatan Pengurus:", ["Ketua RW", "Sekretaris", "Bendahara"])
            
            mapping_nama = {"Ketua RW": "ketuarw", "Sekretaris": "sekretaris", "Bendahara": "bendahara"}
            file_key = mapping_nama[pilih_posisi]
            
            foto_pengurus_up = st.file_uploader(f"Pilih Foto untuk {pilih_posisi} (JPG/PNG)", type=["jpg", "jpeg", "png"])
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
            
            pilih_rt_upload = st.selectbox("Pilih RT untuk Foto:", ["RT 01", "RT 02", "RT 03", "RT 04", "RT 05", "RT 06", "RT 07", "RT 08", "RT 09", "RT 10"])
            rt_num_up = ''.join(filter(str.isdigit, pilih_rt_upload))
            
            foto_rt_upload = st.file_uploader(f"Pilih Foto untuk {pilih_rt_upload} (JPG/PNG)", type=["jpg", "jpeg", "png"])
            
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
            kategori_pdf = st.radio("Pilih Kategori Dokumen PDF:", ["Laporan Kas RW", "Hasil Rapat / Informasi RW"])
            pdf_upload = st.file_uploader("Pilih File PDF", type=["pdf"])
            
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
            st.markdown("Unggah foto kegiatan baru ke galeri RW:")
            foto_upload = st.file_uploader("Pilih File Foto (JPG/PNG)", type=["jpg", "jpeg", "png"])
            if foto_upload is not None:
                folder_galeri = "galeri"
                if not os.path.exists(folder_galeri):
                    os.makedirs(folder_galeri)
                path_simpan = os.path.join(folder_galeri, foto_upload.name)
                with open(path_simpan, "wb") as f:
                    f.write(foto_upload.getbuffer())
                st.success(f"✅ Foto '{foto_upload.name}' berhasil diunggah ke galeri!")
                st.rerun()

# Penutup blok div utama biru muda
st.markdown("</div>", unsafe_allow_html=True)

# ================= FOOTER PORTAL RESMI =================
st.markdown("""
<div style="text-align: center; padding: 20px; color: #555; font-size: 14px; border-top: 1px solid #ddd; margin-top: 30px;">
    <p style="margin: 0; font-weight: bold;">Portal Resmi RW 14 Perum Griya Permata Raya Desa Nanjung Mekar Kec. Rancaekek Kab. Bandung</p>
    <p style="margin: 5px 0 0 0; font-size: 12px; color: #777;">Dikelola oleh Pengurus RW 14 &bull; Didukung oleh Sistem Dashboard Digital Warga</p>
</div>
""", unsafe_allow_html=True)