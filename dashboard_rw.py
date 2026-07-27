import streamlit as st
import pandas as pd
import plotly.express as px
import base64
import os

# 1. PENGATURAN HALAMAN & KOSMETIK
st.set_page_config(page_title="Dashboard RW 14", layout="wide", page_icon="📊")

st.markdown("""
<style>
.stApp { background-color: #F4F9F9; }
div[data-testid="metric-container"] { background-color: white; padding: 15px; border-radius: 10px; box-shadow: 0px 4px 6px rgba(0,0,0,0.05); border-left: 6px solid #1976D2; }
div[data-testid="stMetricValue"], div[data-testid="stMetricValue"] > div { font-size: 45px !important; color: #0D47A1 !important; font-weight: 900 !important; }
div[data-testid="stMetricLabel"] p, div[data-testid="stMetricLabel"] > div, div[data-testid="stMetricLabel"] { font-size: 22px !important; font-weight: bold !important; color: #2C3E50 !important; }
h3 { font-size: 26px !important; color: #0D47A1; }
button[data-baseweb="tab"] > div[data-testid="stMarkdownContainer"] > p { font-size: 20px !important; font-weight: bold; }
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

df = load_data()

st.sidebar.header("🛠️ Panel Filter")
if "RT" in df.columns:
    df["RT_FORMAT"] = df["RT"].apply(lambda x: f"RT{int(x):02d}" if pd.notnull(x) and str(x).isdigit() else f"RT{str(x)}")
    semua_rt_format = sorted(df["RT_FORMAT"].dropna().unique())
    pilihan_rt_format = st.sidebar.multiselect("Tampilkan Data RT:", options=semua_rt_format, default=semua_rt_format)
    if not pilihan_rt_format:
        st.warning("⚠️ Silakan pilih minimal satu RT di menu sebelah kiri.")
        st.stop()
    df_filtered = df[df["RT_FORMAT"].isin(pilihan_rt_format)].copy()
else:
    st.error("Kolom 'RT' tidak ditemukan di Excel.")
    st.stop()

# ================= KONTROL KEAMANAN ADMIN DI SIDEBAR =================
st.sidebar.markdown("---")
st.sidebar.subheader("🔐 Menu Pengurus (Admin)")
mode_admin = st.sidebar.checkbox("Masuk Mode Admin (Edit Data)")

admin_terverifikasi = False
if mode_admin:
    password_input = st.sidebar.text_input("Masukkan Password Admin:", type="password")
    if password_input == "V@nadminrw14":
        admin_terverifikasi = True
        st.sidebar.success("✅ Login Admin Berhasil!")
    elif password_input != "":
        st.sidebar.error("❌ Password salah!")

# ================= BLOK UTAMA BACKGROUND BIRU MUDA =================
st.markdown("""
<div style="background: linear-gradient(135deg, #E3F2FD, #BBDEFB); padding: 25px; border-radius: 20px; box-shadow: 0px 6px 15px rgba(0,0,0,0.08); margin-bottom: 25px; border: 2px solid #90CAF9;">
""", unsafe_allow_html=True)

col_logo, col_teks = st.columns([1, 6])
with col_logo:
    st.image(sumber_logo, width=80)
with col_teks:
    st.markdown("<h2 style='color: #0D47A1; font-weight: 900; margin: 0; padding-top: 15px; text-shadow: 1px 1px 2px rgba(255,255,255,0.8);'>Dashboard Interaktif Data Warga RW 14</h2>", unsafe_allow_html=True)

st.write("---")

# ================= MENU TAB MENYESUAIKAN STATUS ADMIN =================
if admin_terverifikasi:
    tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs(["📋 Ringkasan", "👫 Demografi & Agama", "💼 Usia & Profesi", "🗂️ Semua Data", "🏠 Pencarian KK", "⚙️ Edit Data (Admin)"])
else:
    tab1, tab2, tab3, tab4, tab5 = st.tabs(["📋 Ringkasan", "👫 Demografi & Agama", "💼 Usia & Profesi", "🗂️ Semua Data", "🏠 Pencarian KK"])

# ================= TAB 1: RINGKASAN =================
with tab1:
    st.subheader("Angka Kunci Terkini")
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
    fig_rt.update_layout(plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)", showlegend=False, xaxis=dict(title="", tickfont=dict(size=24, color="black", weight="bold")), yaxis=dict(title="Jumlah Penduduk (Jiwa)"), margin=dict(t=40)) 
    fig_rt.update_traces(textfont_size=28, textfont_color="black", textangle=0, textposition="outside", cliponaxis=False)
    st.plotly_chart(fig_rt, use_container_width=True)

# ================= TAB 2: DEMOGRAFI & AGAMA =================
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

# ================= TAB 3: KELOMPOK USIA & PROFESI =================
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

# ================= TAB 4: TABEL DATA =================
with tab4:
    st.subheader("Tabel Seluruh Warga")
    st.markdown("💡 *Data sensitif (NIK & No. KK) disembunyikan untuk keamanan publik.*")
    kolom_dibuang = ["NO. KK", "NIK", "USIA_ANGKA", "Kelompok Usia", "RT_FORMAT"]
    df_tabel = df_filtered.drop(columns=kolom_dibuang, errors="ignore")
    st.dataframe(df_tabel, use_container_width=True, hide_index=True)

# ================= TAB 5: PENCARIAN KK =================
with tab5:
    st.subheader("🔍 Pencarian & Data per Kartu Keluarga (KK)")
    st.markdown("Ketik nama salah satu warga untuk melihat seluruh anggota keluarganya.")
    
    kata_kunci = st.text_input("🔎 Masukkan Nama Warga:")
    
    if kata_kunci:
        # Bersihkan spasi berlebih pada kata kunci ketikan
        kunci_bersih = kata_kunci.strip().lower()
        
        # Filter ketat khusus pada kolom NAMA (membersihkan spasi di Excel juga)
        if "NAMA" in df_filtered.columns:
            # Ubah kolom NAMA menjadi string kecil semua dan hilangkan spasi ganda
            nama_seragam = df_filtered["NAMA"].astype(str).str.lower().str.strip()
            mask_nama = nama_seragam.str.contains(kunci_bersih, na=False)
            
            # Jika user mengetik angka (misal No KK / NIK), izinkan pencarian umum
            if kunci_bersih.isdigit():
                mask_umum = df_filtered.astype(str).apply(lambda x: x.str.lower().str.contains(kunci_bersih, na=False)).any(axis=1)
                hasil_pencarian = df_filtered[mask_nama | mask_umum]
            else:
                hasil_pencarian = df_filtered[mask_nama]
        else:
            hasil_pencarian = pd.DataFrame()
        
        if not hasil_pencarian.empty:
            if "NO. KK" in df_filtered.columns:
                # Ambil daftar No KK yang benar-benar valid dari hasil temuan nama
                list_kk = hasil_pencarian["NO. KK"].dropna().unique()
                st.success(f"✅ Ditemukan {len(list_kk)} Kartu Keluarga terkait.")
                
                for kk in list_kk:
                    df_keluarga = df_filtered[df_filtered["NO. KK"] == kk].copy()
                    
                    # Pastikan sekali lagi: Apakah di dalam 1 KK ini benar-benar ada nama yang cocok?
                    # Jika karena suatu hal tidak ada, lewati KK tersebut
                    cek_lagi = df_keluarga["NAMA"].astype(str).str.lower().str.contains(kunci_bersih, na=False)
                    if not cek_lagi.any() and not kunci_bersih.isdigit():
                        continue
                        
                    nama_kepala = "Satu Keluarga"
                    if "HUBUNGAN" in df_keluarga.columns and "NAMA" in df_keluarga.columns:
                        kepala_df = df_keluarga[df_keluarga["HUBUNGAN"].astype(str).str.upper() == "KEPALA KELUARGA"]
                        if not kepala_df.empty:
                            nama_kepala = "Keluarga Bpk/Ibu " + str(kepala_df.iloc[0]["NAMA"]).title()
                    
                    st.markdown(f"### 🏠 {nama_kepala}")
                    
                    # Penanda warna kuning lembut untuk baris nama yang persis cocok
                    def highlight_pencarian(row):
                        match = kunci_bersih in str(row.get("NAMA", "")).lower()
                        return ['background-color: #FFF9C4' if match else '' for _ in row]

                    kolom_dibuang = ["NO. KK", "NIK", "USIA_ANGKA", "Kelompok Usia", "RT_FORMAT"]
                    df_tampil = df_keluarga.drop(columns=kolom_dibuang, errors="ignore")
                    
                    st.dataframe(df_tampil.style.apply(highlight_pencarian, axis=1), use_container_width=True, hide_index=True)
            else:
                st.error("Kolom 'NO. KK' tidak ditemukan.")
        else:
            st.warning(f"❌ Tidak ada warga dengan nama '{kata_kunci}' yang ditemukan.")

# ================= TAB 6: EDIT DATA (HANYA MUNCUL JIKA ADMIN LOGIN) =================
if admin_terverifikasi:
    with tab6:
        st.subheader("⚙️ Edit Data Warga (Admin)")
        st.warning("⚠️ Anda berada dalam mode Admin. Setiap perubahan akan memperbarui file **datawarga.xlsx**.")
        data_terbaru = st.data_editor(df.drop(columns=["RT_FORMAT"], errors="ignore"), num_rows="dynamic", use_container_width=True)
        if st.button("💾 Simpan Perubahan ke Excel", type="primary"):
            try:
                data_terbaru.to_excel("datawarga.xlsx", index=False)
                st.cache_data.clear()
                st.success("✅ Data berhasil disimpan!")
                st.rerun()
            except Exception as e:
                st.error(f"❌ Gagal menyimpan data: {e}")

# Penutup blok div utama biru muda
st.markdown("</div>", unsafe_allow_html=True)