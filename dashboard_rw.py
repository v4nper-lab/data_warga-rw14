import streamlit as st
import pandas as pd
import plotly.express as px
import base64
import os
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
button[data-baseweb="tab"] > div[data-testid="stMarkdownContainer"] > p { font-size: 15px !important; font-weight: bold; }
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

@st.cache_data
def load_data():
    df = pd.read_excel("datawarga.xlsx")
    df.columns = df.columns.str.strip().str.upper()
    if "UMUR" in df.columns and "USIA" not in df.columns: df.rename(columns={"UMUR": "USIA"}, inplace=True)
    if "STATUS" in df.columns and "STATUS PERKAWINAN" not in df.columns: df.rename(columns={"STATUS": "STATUS PERKAWINAN"}, inplace=True)
    if "STATUS NIKAH" in df.columns and "STATUS PERKAWINAN" not in df.columns: df.rename(columns={"STATUS NIKAH": "STATUS PERKAWINAN"}, inplace=True)
    if "NO KK" in df.columns and "NO. KK" not in df.columns: df.rename(columns={"NO KK": "NO. KK"}, inplace=True)
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

df = load_data()
df_kas = load_kas()
df_info = load_info()

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

st.sidebar.header("🛠️ Panel Filter Data")
if "RT" in df.columns:
    df["RT_FORMAT"] = df["RT"].apply(lambda x: f"RT{int(x):02d}" if pd.notnull(x) and str(x).isdigit() else f"RT{str(x)}")
    semua_rt_format = sorted(df["RT_FORMAT"].dropna().unique())
    pilihan_rt_format = st.sidebar.multiselect("Tampilkan Data RT:", options=semua_rt_format, default=semua_rt_format)
    if not pilihan_rt_format:
        st.warning("⚠️ Silخاب silakan pilih minimal satu RT di menu sebelah kiri.")
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
    tab0, tab1, tab2, tab3, tab4, tab5, tab6, tab7, tab8, tab9 = st.tabs([
        "🏠 Beranda", "📋 Statistik", "👫 Demografi", "💼 Profesi", 
        "🗂️ Data Warga", "🔍 Cari KK", "💰 Kas RW", 
        "📢 Info & Rapat", "🖼️ Galeri Kegiatan", "⚙️ Edit Data (Admin)"
    ])
else:
    tab0, tab1, tab2, tab3, tab4, tab5, tab6, tab7, tab8 = st.tabs([
        "🏠 Beranda", "📋 Statistik", "👫 Demografi", "💼 Profesi", 
        "🗂️ Data Warga", "🔍 Cari KK", "💰 Kas RW", 
        "📢 Info & Rapat", "🖼️ Galeri Kegiatan"
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
        * Melihat statistik kependudukan dan sebaran RT.
        * Memeriksa data Kartu Keluarga (KK) dengan mudah menggunakan fitur pencarian nama.
        * Memantau transparansi laporan keuangan kas RW secara terbuka.
        * Membaca hasil rapat, pengumuman, dan agenda kegiatan lingkungan.
        * Menyimak dokumentasi foto kegiatan warga di menu Galeri.
        
        Mari bersama-sama kita wujudkan kerukunan, keterbukaan, dan pelayanan warga yang semakin prima!
        """)
    with col_p2:
        st.markdown("""
        <div style="background-color: white; padding: 15px; border-radius: 10px; border: 1px solid #90CAF9; box-shadow: 0px 2px 5px rgba(0,0,0,0.05);">
            <h4 style="color: #0D47A1; margin-top:0;">📞 Kontak Penting RW 14</h4>
            <p style="margin: 8px 0; font-size: 14px;">🚨 <b>Keamanan / Satpam:</b> 0812-XXXX-XXXX</p>
            <p style="margin: 8px 0; font-size: 14px;">🏥 <b>Kesehatan / Posyandu:</b> 0813-XXXX-XXXX</p>
            <p style="margin: 8px 0; font-size: 14px;">🧹 <b>Kebersihan / RT:</b> Hubungi RT Masing-masing</p>
            <hr style="margin: 10px 0;">
            <p style="margin: 0; font-size: 12px; color: #666; text-align: center;"><b>RW 14 Bersih, Rukun, & Sejahtera</b></p>
        </div>
        """, unsafe_allow_html=True)

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

# ================= TAB 3: PROFESI & USIA =================
with tab3:
    kolom_kiri2, kolom_kanan2 = st.columns(2)
    with kolom_kiri2:
        st.subheader("Pengelompokan Usia")
        if "USIA" in df_filtered.columns:
            df_filtered['USIA_ANGKA'] = pd.to_numeric(df_filtered['USIA'], errors='coerce')
            batas_usia = [0, 5, 12, 25, 45, 60, 150]
            label_usia = ['Balita (0-5)', 'Anak-anak (6-12)', 'Remaja (13-25)', 'Dewasa (26-45)', 'Pra-Lansia (46-60)', 'Lansia (>60)']
            df_filtered['Kelompok Usia'] = pd.cut(df_filtered['USIA_ANGKA'], bins=batas_usia, labels=label_usia, right=True)
            df_kelompok = df_filtered['Kelompok Usia'].value_counts().reset_index()
            df_kelompok.columns = ["Kelompok Usia", "Jumlah"]
            fig_kel_usia = px.bar(df_kelompok, x="Kelompok Usia", y="Jumlah", color="Kelompok Usia", text_auto=True, color_discrete_sequence=px.colors.qualitative.Safe)
            fig_kel_usia.update_layout(paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", showlegend=False, xaxis={'categoryorder':'array', 'categoryarray':label_usia}, font=dict(size=15))
            fig_kel_usia.update_traces(textfont_size=18, textangle=0)
            st.plotly_chart(fig_kel_usia, use_container_width=True)
    with kolom_kanan2:
        st.subheader("Top 10 Profesi Warga")
        if "PEKERJAAN" in df_filtered.columns:
            df_pekerjaan = df_filtered["PEKERJAAN"].astype(str).str.title().value_counts().reset_index().head(10)
            df_pekerjaan.columns = ["Profesi", "Jumlah"]
            fig_profesi = px.bar(df_pekerjaan, x="Jumlah", y="Profesi", orientation='h', color="Profesi", text_auto=True)
            fig_profesi.update_layout(yaxis={'categoryorder':'total ascending'}, paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", showlegend=False, font=dict(size=15))
            fig_profesi.update_traces(textfont_size=16)
            st.plotly_chart(fig_profesi, use_container_width=True)

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
    st.markdown("Berikut adalah ringkasan keuangan dan rincian transaksi Kas RW yang dapat diakses oleh seluruh warga.")
    
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

# ================= TAB 7: INFO & RAPAT =================
with tab7:
    st.subheader("📢 Informasi Kegiatan & Hasil Rapat RW 14")
    st.markdown("Pusat informasi resmi seputar hasil rapat pengurus, pengumuman warga, dan agenda kegiatan lingkungan.")
    
    if not df_info.empty:
        for index, row in df_info.iterrows():
            tgl = row.get("TANGGAL", "Agenda RW")
            judul = row.get("JUDUL", "Informasi Penting")
            isi = row.get("ISI / KATEGORI", "-")
            
            with st.expander(f"📌 [{tgl}] — {judul}"):
                st.write(isi)
    else:
        st.info("ℹ️ Belum ada pengumuman atau hasil rapat yang dipublikasikan.")

# ================= TAB 8: GALERI KEGIATAN =================
with tab8:
    st.subheader("🖼️ Galeri Foto Kegiatan Warga RW 14")
    st.markdown("Dokumentasi foto kegiatan warga, kerja bakti, posyandu, dan acara kebersamaan di lingkungan Perum Griya Permata Raya.")
    
    # Folder penyimpanan galeri
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

# ================= TAB 9: EDIT DATA (HANYA ADMIN) =================
if admin_terverifikasi:
    with tab9:
        st.subheader("⚙️ Panel Pengaturan & Edit Data (Admin)")
        st.warning("⚠️ Anda berada dalam mode Admin. Anda dapat memperbarui data warga, kas, informasi rapat, maupun mengunggah foto galeri.")
        
        menu_admin = st.selectbox("Pilih Data yang Ingin Dikelola:", ["Data Warga", "Laporan Kas RW", "Informasi & Hasil Rapat", "Upload Foto Galeri"])
        
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