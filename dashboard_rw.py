with tab5:
    st.subheader("🔍 Pencarian Lembar Dokumen Kartu Keluarga (KK)")
    
    if "cari_terbuka" not in st.session_state: 
        st.session_state["cari_terbuka"] = False
        
    if not st.session_state["cari_terbuka"]:
        pass_cari = st.text_input("Kata Sandi Akses Pencarian KK:", type="password", key="pass_input_cari_kk")
        if pass_cari == "ijindibuka": 
            st.session_state["cari_terbuka"] = True
            st.rerun()
        elif pass_cari != "": 
            st.error("❌ Kata sandi salah!")
    else:
        col_btn1, col_btn2 = st.columns([1, 4])
        with col_btn1:
            if st.button("🔒 Kunci Kembali Pencarian"): 
                st.session_state["cari_terbuka"] = False
                st.rerun()
        with col_btn2:
            st.markdown("<button onclick='window.print()' style='background-color:#0D47A1; color:white; padding:8px 16px; border:none; border-radius:6px; font-weight:bold; cursor:pointer;'>🖨️ Cetak / Print Dokumen KK</button>", unsafe_allow_html=True)
        
        # Deteksi nama kolom secara otomatis
        kolom_nama = next((k for k in df.columns if "NAMA" in k), df.columns[1] if len(df.columns) > 1 else df.columns[0])
        kolom_hub = next((c for c in df.columns if "HUBUNGAN" in c or "STATUS KELUARGA" in c or "KEDUDUKAN" in c), None)
        kolom_alamat = next((k for k in df.columns if "ALAMAT" in k), "ALAMAT")

        kata_kunci = st.text_input("🔎 Ketik Nama Warga (Contoh: irma, asep, salya):", key="input_pencarian_blok_keluarga")
        
        if kata_kunci:
            kw = str(kata_kunci).strip().lower()
            
            df_temp = df.copy()
            df_temp["_COL_NAMA_"] = df_temp[kolom_nama].fillna("").astype(str).str.lower()
            
            # MEMBUAT BLOK KELUARGA OTOMATIS BERDASARKAN BARIS KEPALA KELUARGA
            if kolom_hub:
                kk_mask = df_temp[kolom_hub].fillna("").astype(str).str.upper().str.contains("KEPALA|KDH", regex=True)
                df_temp["_FAMILY_ID_"] = kk_mask.cumsum()
            else:
                df_temp["_FAMILY_ID_"] = df_temp[kolom_alamat].fillna("").astype(str).str.strip()
            
            # Cari baris warga yang sesuai dengan kata kunci nama
            matched_rows = df_temp[df_temp["_COL_NAMA_"].str.contains(kw, na=False)]
            
            if not matched_rows.empty:
                # Ambil daftar ID Keluarga unik dari orang-orang yang ditemukan
                family_ids = matched_rows["_FAMILY_ID_"].unique()
                
                for fam_id in family_ids:
                    if fam_id == 0 and kolom_hub: 
                        continue
                    
                    # TARIK HANYA ANGGOTA KELUARGA DALAM 1 BLOK KELUARGA TERSEBUT
                    keluarga_df = df[df_temp["_FAMILY_ID_"] == fam_id].copy()
                    
                    if not keluarga_df.empty:
                        # Baris pertama di blok keluarga pasti Kepala Keluarga
                        utama = keluarga_df.iloc[0]
                        
                        al_kk = utama.get(kolom_alamat, "-")
                        rt_kk = utama.get("RT", "-")
                        rw_kk = utama.get("RW", "14")
                        ds_kk = utama.get("DUSUN", "-")
                        
                        # Siapkan info siapa saja anggota yang cocok ditemukan di keluarga ini
                        anggota_cocok = matched_rows[matched_rows["_FAMILY_ID_"] == fam_id][kolom_nama].tolist()
                        info_pencarian = ", ".join([str(x) for x in anggota_cocok])
                        
                        # Tampilkan Header Info Kartu Keluarga
                        st.markdown(f"""
                        <div style="background-color: #ffffff; padding: 20px; border-radius: 12px; border: 2px solid #0D47A1; margin-bottom: 25px; box-shadow: 0px 4px 10px rgba(0,0,0,0.08);">
                            <div style="text-align: center; border-bottom: 2px solid #0D47A1; padding-bottom: 10px; margin-bottom: 15px;">
                                <h3 style="margin: 0; color: #0D47A1; font-size: 20px;">KARTU KELUARGA (KK)</h3>
                                <p style="margin: 3px 0 0 0; font-weight: bold; color: #555; font-size: 14px;">Alamat : {al_kk}</p>
                            </div>
                            <div style="display: flex; justify-content: space-between; font-size: 14px; margin-bottom: 15px; color: #333;">
                                <div>
                                    <p style="margin: 4px 0;"><b>Anggota Ditemukan :</b> {info_pencarian}</p>
                                    <p style="margin: 4px 0;"><b>RT / RW :</b> {rt_kk} / {rw_kk}</p>
                                </div>
                                <div>
                                    <p style="margin: 4px 0;"><b>Dusun / Desa :</b> {ds_kk} / Nanjung Mekar</p>
                                </div>
                            </div>
                            <p style="font-weight: bold; color: #0D47A1; margin-bottom: 8px; font-size: 14px;">📋 Daftar Anggota Keluarga (Kepala Keluarga & Keluarga Inti):</p>
                        </div>
                        """, unsafe_allow_html=True)
                        
                        # Tampilkan tabel anggota keluarga tersebut secara eksklusif
                        st.dataframe(keluarga_df, use_container_width=True, hide_index=True)
                        st.markdown("<hr style='margin: 30px 0;'>", unsafe_allow_html=True)
            else:
                st.warning(f"⚠️ Warga dengan nama '{kata_kunci}' tidak ditemukan di database.")