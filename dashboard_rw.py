import streamlit as st
import pandas as pd
import os

st.set_page_config(page_title="Portal RW 14", layout="wide")

st.title("Portal Resmi RW 14 Griya Permata Raya")
st.write("Sistem kembali normal dan siap membaca data Anda.")

tab1, tab2, tab3 = st.tabs(["🏠 Beranda", "💰 Kas RW", "🖼️ Galeri"])

with tab1:
    st.subheader("Data Warga")
    if os.path.exists("datawarga.xlsx"):
        df = pd.read_excel("datawarga.xlsx")
        st.dataframe(df, use_container_width=True)
    else:
        st.info("File datawarga.xlsx siap dibaca.")

with tab2:
    st.subheader("Laporan Kas RW")
    if os.path.exists("datakas.xlsx"):
        df_kas = pd.read_excel("datakas.xlsx")
        st.dataframe(df_kas, use_container_width=True)
    else:
        st.info("File datakas.xlsx siap dibaca.")

with tab3:
    st.subheader("Galeri Kegiatan")
    if os.path.exists("datagaleri.xlsx"):
        df_gal = pd.read_excel("datagaleri.xlsx")
        st.dataframe(df_gal, use_container_width=True)
    else:
        st.info("File galeri siap dibaca.")