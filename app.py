import io
import time
from datetime import datetime
import pandas as pd
import plotly.express as px
import streamlit as st
from docx import Document
from google import genai
from google.genai import types

# Setup Halaman Streamlit
st.set_page_config(
    page_title="Sistem Intelkam - Analisis & Generasi Dokumen",
    layout="wide",
    page_icon="🛡",
)

st.title("🛡️ Sistem Informasi & Rekapitulasi Intelkam Polres Ciamis")
st.caption(
    "Dashboard Analisis, Pengkategorian Dokumen Intelijen, & Monitoring Kegiatan Kepolisian"
)

# Sidebar Input API Key & Konfigurasi Sistem
st.sidebar.header("⚙️ Pengaturan Sistem")

saved_api_key = ""
if "GEMINI_API_KEY" in st.secrets:
    saved_api_key = st.secrets["GEMINI_API_KEY"]

if saved_api_key:
    st.sidebar.success("🔑 API Key Permanen Terdeteksi & Aktif!")
    api_key = saved_api_key
else:
    api_key = st.sidebar.text_input(
        "Masukkan Gemini API Key:",
        type="password",
        help="Dapatkan di Google AI Studio (Gratis)",
    )

# Inisialisasi Database Lokal / Catatan Kegiatan (Mendukung data historis dari Januari)
if "db_kegiatan" not in st.session_state:
    # Contoh data awal untuk simulasi/pengujian bulan Januari - September
    initial_data = [
        {
            "Nomor": 1,
            "Nomor Dokumen": "LI/01/I/2026/IK",
            "Tanggal Dokumen": "2026-01-10",
            "Jenis Dokumen": "Laporan Informasi (LI)",
            "Nama/Judul Kegiatan": "Konser Musik Awal Tahun",
            "Kategori Kegiatan": "Konser/Hiburan",
            "Lokasi": "Alun-alun Ciamis",
            "Tanggal Kegiatan": "2026-01-10",
            "Bulan": "Januari",
            "Tahun": 2026,
            "Keterangan": "Aman dan terkendali",
        },
        {
            "Nomor": 2,
            "Nomor Dokumen": "SI/05/II/2026/IK",
            "Tanggal Dokumen": "2026-02-15",
            "Jenis Dokumen": "Surat Izin (SI)",
            "Nama/Judul Kegiatan": "Turnamen Sepak Bola Bupati Cup",
            "Kategori Kegiatan": "Kompetisi/Pertandingan",
            "Lokasi": "Stadion Galuh Ciamis",
            "Tanggal Kegiatan": "2026-02-15",
            "Bulan": "Februari",
            "Tahun": 2026,
            "Keterangan": "Berjalan lancar",
        },
    ]
    st.session_state.db_kegiatan = pd.DataFrame(initial_data)

# Tab Navigasi Utama
tab1, tab2, tab3 = st.tabs(
    [
        "📄 Upload & Generator Dokumen",
        "📊 Dasbor & Analisis Statistik",
        "🗂️ Database & Manajemen Data",
    ]
)

# ==========================================
# TAB 1: UPLOAD & GENERATOR DOKUMEN
# ==========================================
with tab1:
    st.subheader("1. Unggah Bahan / Berkas Kegiatan (Bisa Banyak Sekaligus)")
    uploaded_files = st.file_uploader(
        "Pilih file (PDF, TXT, DOCX, atau Foto/Scan Surat) - Dapat memilih lebih dari satu file",
        type=["pdf", "docx", "txt", "png", "jpg"],
        accept_multiple_files=True,
    )

    col1, col2 = st.columns(2)
    with col1:
        perintah = st.selectbox(
            "Pilih Format Output Dokumen (Kata Kunci):",
            [
                "skp (Surat Keterangan Kepolisian)",
                "si (Surat Izin)",
                "sttp (Surat Tanda Terima Pemberitahuan)",
                "li (Laporan Informasi)",
                "infosus (Informasi Khusus)",
                "kirkat (Perkiraan Keadaan Singkat)",
            ],
        )
    with col2:
        kategori_kegiatan = st.selectbox(
            "Kategori / Bidang Kegiatan:",
            [
                "Konser/Hiburan",
                "Kompetisi/Pertandingan",
                "Kegiatan Masyarakat",
                "Kegiatan Pemerintahan",
                "Lainnya",
            ],
        )

    tgl_kegiatan_input = st.date_input(
        "Tanggal Pelaksanaan Kegiatan:", value=datetime.today()
    )
    lokasi_input = st.text_input("Lokasi Kegiatan:", value="Wilayah Hukum Polres Ciamis")

    if st.button("🚀 Proses & Buat Dokumen"):
        if not api_key:
            st.error(
                "Silakan masukkan Gemini API Key terlebih dahulu di sidebar atau atur di Secrets!"
            )
        elif not uploaded_files:
            st.warning("Silakan unggah minimal satu berkas bahan kegiatan!")
        else:
            with st.spinner(
                "AI sedang menganalisis berkas dan menyusun draf baku sesuai format Sat Intelkam..."
            ):
                try:
                    client = genai.Client(api_key=api_key)
                    keyword_cmd = perintah.split()[0].lower()

                    for uploaded_file in uploaded_files:
                        bytes_data = uploaded_file.getvalue()
                        mime_type = uploaded_file.type

                        if mime_type == "text/plain":
                            file_part = bytes_data.decode("utf-8")
                        else:
                            file_part = types.Part.from_bytes(
                                data=bytes_data, mime_type=mime_type
                            )

                        prompt = f"""
                        Anda adalah Asisten Intelkam Polres Ciamis. Pelajari data yang diunggah dan buatkan draft teks produk intelijen sesuai kode perintah: {keyword_cmd}.
                        Aturan Ketat & Format Baku Sat Intelkam Polres Ciamis:
                        - Pejabat Penandatangan Resmi: KASAT INTELKAM POLRES CIAMIS, AKP RAHMAT KOMARA, S.H., M.H., AJUN KOMISARIS POLISI NRP 70030155.
                        - Gunakan bahasa baku dinas Kepolisian Republik Indonesia, lengkap, terstruktur, tidak disingkat sembarangan.
                        - Jika perintah 'skp': Buat Surat Keterangan Kepolisian sesuai format baku.
                        - Jika perintah 'si': Buat Surat Izin dengan struktur tabel 2 kolom dan 4 poin catatan/kewajiban.
                        - Jika perintah 'sttp': Buat Surat Tanda Terima Pemberitahuan sesuai format baku.
                        - Jika perintah 'li': Buat Laporan Informasi lengkap (Sumber, Hubungan, Cara, Waktu, Nilai A-1, Fakta 5W+1H, Analisa, Prediksi, Langkah, Rekomendasi).
                        - Jika perintah 'infosus': Buat Nota Dinas Pengantar dan Lembar Infosus berklasifikasi RAHASIA.
                        - Jika perintah 'kirkat': Buat Perkiraan Keadaan Intelijen Singkat.
                        """

                        models_to_try = [
                            "gemini-2.0-flash",
                            "gemini-1.5-flash",
                            "gemini-3.8-flash",
                        ]
                        response = None
                        last_error = None

                        for model_name in models_to_try:
                            try:
                                for attempt in range(3):
                                    try:
                                        response = client.models.generate_content(
                                            model=model_name,
                                            contents=[file_part, prompt],
                                        )
                                        if response:
                                            break
                                    except Exception as err:
                                        if (
                                            "503" in str(err)
                                            or "UNAVAILABLE" in str(err)
                                            or "404" in str(err)
                                        ):
                                            time.sleep(2)
                                        else:
                                            raise err
                                if response:
                                    break
                            except Exception as m_err:
                                last_error = m_err
                                continue

                        if not response:
                            raise Exception(
                                f"Gagal memproses file {uploaded_file.name}: {last_error}"
                            )

                        hasil_teks = response.text
                        st.success(f"Berhasil memproses: {uploaded_file.name}")
                        st.text_area(
                            f"Pratinjau: {uploaded_file.name}",
                            hasil_teks,
                            height=250,
                        )

                        # Export ke Word
                        doc = Document()
                        for p_para in hasil_teks.split("\n"):
                            doc.add_paragraph(p_para)
                        bio = io.BytesIO()
                        doc.save(bio)

                        st.download_button(
                            label=f"📥 Download Word ({uploaded_file.name})",
                            data=bio.getvalue(),
                            file_name=f"Produk_{keyword_cmd.upper()}_{uploaded_file.name.split('.')[0]}.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        )

                        # Otomatis catat ke database
                        bulan_nama = [
                            "Januari",
                            "Februari",
                            "Maret",
                            "April",
                            "Mei",
                            "Juni",
                            "Juli",
                            "Agustus",
                            "September",
                            "Oktober",
                            "November",
                            "Desember",
                        ][tgl_kegiatan_input.month - 1]

                         jenis_dok_map = {
                            "skp": "Surat Keterangan Kepolisian (SKP)",
                            "si": "Surat Izin (SI)",
                            "sttp": "Surat Tanda Terima Pemberitahuan",
                            "li": "Laporan Informasi (LI)",
                            "infosus": "Informasi Khusus",
                            "kirkat": "Perkiraan Keadaan Singkat",
                        }

                        new_row = {
                            "Nomor": len(st.session_state.db_kegiatan) + 1,
                            "Nomor Dokumen": f"{keyword_cmd.upper()}/{len(st.session_state.db_kegiatan)+1}/X/2026/IK",
                            "Tanggal Dokumen": datetime.today().strftime("%Y-%m-%d"),
                            "Jenis Dokumen": jenis_dok_map.get(keyword_cmd, keyword_cmd.upper()),
                            "Nama/Judul Kegiatan": uploaded_file.name.split(".")[0],
                            "Kategori Kegiatan": kategori_kegiatan,
                            "Lokasi": lokasi_input,
                            "Tanggal Kegiatan": tgl_kegiatan_input.strftime("%Y-%m-%d"),
                            "Bulan": bulan_nama,
                            "Tahun": tgl_kegiatan_input.year,
                            "Keterangan": "Selesai diproses sistem",
                        }
                        st.session_state.db_kegiatan = pd.concat(
                            [
                                st.session_state.db_kegiatan,
                                pd.DataFrame([new_row]),
                            ],
                            ignore_index=True,
                        )

                except Exception as e:
                    st.error(f"Terjadi kesalahan: {e}")

# ==========================================
# TAB 2: DASBOR & ANALISIS STATISTIK
# ==========================================
with tab2:
    st.subheader("📊 Dashboard Analisis Data Intelijen")

    if st.session_state.db_kegiatan.empty:
        st.info("Belum ada data kegiatan terdaftar.")
    else:
        df = st.session_state.db_kegiatan

        # Filter Global
        st.markdown("#### 🔍 Filter Data")
        col_f1, col_f2, col_f3 = st.columns(3)
        with col_f1:
            selected_tahun = st.multiselect(
                "Tahun:", options=df["Tahun"].unique(), default=df["Tahun"].unique()
            )
        with col_f2:
            selected_bulan = st.multiselect(
                "Bulan:", options=df["Bulan"].unique(), default=df["Bulan"].unique()
            )
        with col_f3:
            selected_kat = st.multiselect(
                "Kategori Kegiatan:",
                options=df["Kategori Kegiatan"].unique(),
                default=df["Kategori Kegiatan"].unique(),
            )

        df_filtered = df[
            df["Tahun"].isin(selected_tahun)
            & df["Bulan"].isin(selected_bulan)
            & df["Kategori Kegiatan"].isin(selected_kat)
        ]

        # Statistik Kartu Utama
        total_dok = len(df_filtered)
        tot_si = len(
            df_filtered[
                df_filtered["Jenis Dokumen"].str.contains("Surat Izin", case=False)
            ]
        )
        tot_li = len(
            df_filtered[
                df_filtered["Jenis Dokumen"].str.contains(
                    "Laporan Informasi", case=False
                )
            ]
        )
        tot_infosus = len(
            df_filtered[
                df_filtered["Jenis Dokumen"].str.contains("Informasi Khusus", case=False)
            ]
        )
        tot_kirkat = len(
            df_filtered[
                df_filtered["Jenis Dokumen"].str.contains(
                    "Perkiraan Keadaan", case=False
                )
            ]
        )
        tot_skp = len(
            df_filtered[
                df_filtered["Jenis Dokumen"].str.contains(
                    "Surat Keterangan Kepolisian", case=False
                )
            ]
        )

        c1, c2, c3, c4, c5, c6 = st.columns(6)
        c1.metric("Total Dokumen", total_dok)
        c2.metric("Surat Izin", tot_si)
        c3.metric("Laporan Info", tot_li)
        c4.metric("Infosus", tot_infosus)
        c5.metric("Kirkat", tot_kirkat)
        c6.metric("SKP", tot_skp)

        st.markdown("---")

        # Ringkasan Analisis Otomatis
        st.markdown("### 📝 Analisis Otomatis")
        if not df_filtered.empty:
            top_dok = df_filtered["Jenis Dokumen"].mode()[0] if not df_filtered["Jenis Dokumen"].empty else "-"
            top_kat = df_filtered["Kategori Kegiatan"].mode()[0] if not df_filtered["Kategori Kegiatan"].empty else "-"
            st.info(
                f"Berdasarkan filter aktif, tercatat total **{total_dok} dokumen intelijen**. "
                f"Jenis dokumen yang paling sering diterbitkan adalah **{top_dok}**, "
                f"sedangkan kategori kegiatan terbanyak ditangani adalah **{top_kat}**."
            )
        else:
            st.warning("Tidak ada data yang sesuai dengan filter yang dipilih.")

        st.markdown("---")

        # Grafik Grafik Batang (Jenis Dokumen & Kategori Kegiatan)
        col_g1, col_g2 = st.columns(2)
        with col_g1:
            st.markdown("#### Distribusi Jenis Dokumen")
            if not df_filtered.empty:
                dok_counts = df_filtered["Jenis Dokumen"].value_counts().reset_index()
                dok_counts.columns = ["Jenis Dokumen", "Jumlah"]
                fig_dok = px.bar(
                    dok_counts,
                    x="Jenis Dokumen",
                    y="Jumlah",
                    text="Jumlah",
                    color="Jenis Dokumen",
                )
                st.plotly_chart(fig_dok, use_container_width=True)
            else:
                st.write("Data kosong.")

        with col_g2:
            st.markdown("#### Sebaran Kategori Kegiatan")
            if not df_filtered.empty:
                kat_counts = (
                    df_filtered["Kategori Kegiatan"].value_counts().reset_index()
                )
                kat_counts.columns = ["Kategori Kegiatan", "Jumlah"]
                fig_kat = px.bar(
                    kat_counts,
                    x="Kategori Kegiatan",
                    y="Jumlah",
                    text="Jumlah",
                    color="Kategori Kegiatan",
                )
                st.plotly_chart(fig_kat, use_container_width=True)
            else:
                st.write("Data kosong.")

        # Tren Kegiatan Per Bulan (Jan - Des)
        st.markdown("### 📈 Tren Jumlah Kegiatan Per Bulan (Januari - Desember)")
        bulan_order = [
            "Januari",
            "Februari",
            "Maret",
            "April",
            "Mei",
            "Juni",
            "Juli",
            "Agustus",
            "September",
            "Oktober",
            "November",
            "Desember",
        ]
        if not df_filtered.empty:
            monthly_trend = (
                df_filtered.groupby("Bulan")
                .size()
                .reindex(bulan_order, fill_value=0)
                .reset_index()
            )
            monthly_trend.columns = ["Bulan", "Jumlah Kegiatan"]
            fig_trend = px.line(
                monthly_trend,
                x="Bulan",
                y="Jumlah Kegiatan",
                markers=True,
                text="Jumlah Kegiatan",
            )
            st.plotly_chart(fig_trend, use_container_width=True)

        # Analisis Silang
        st.markdown("### 🔀 Analisis Silang: Jenis Dokumen × Kategori Kegiatan")
        if not df_filtered.empty:
            cross_tab = pd.crosstab(
                df_filtered["Kategori Kegiatan"],
                df_filtered["Jenis Dokumen"],
                margins=True,
                margins_name="Total",
            )
            st.dataframe(cross_tab, use_container_width=True)

# ==========================================
# TAB 3: DATABASE & MANAJEMEN DATA
# ==========================================
with tab3:
    st.subheader("🗂️ Database Dokumen & Manajemen Hapus Data")

    if st.session_state.db_kegiatan.empty:
        st.info("Database kosong.")
    else:
        st.write("Draf dan riwayat dokumen yang telah diproses:")
        st.dataframe(st.session_state.db_kegiatan, use_container_width=True)

        st.markdown("---")
        st.markdown("#### 🗑️ Hapus Dokumen / Data")
        row_to_delete = st.selectbox(
            "Pilih Baris / Dokumen yang Ingin Dihapus (Berdasarkan Judul & Tanggal):",
            options=st.session_state.db_kegiatan.index,
            format_func=lambda x: f"[{x}] {st.session_state.db_kegiatan.loc[x, 'Nama/Judul Kegiatan']} ({st.session_state.db_kegiatan.loc[x, 'Tanggal Kegiatan']})",
        )

        if st.button("🗑️ Hapus Baris Terpilih"):
            st.session_state.db_kegiatan = (
                st.session_state.db_kegiatan.drop(row_to_delete)
                .reset_index(drop=True)
            )
            st.success("Data berhasil dihapus dari sistem!")
            st.rerun()

        # Tombol Export Excel/CSV
        st.markdown("---")
        csv_data = st.session_state.db_kegiatan.to_csv(index=False).encode("utf-8")
        st.download_button(
            label="📥 Export Data ke CSV",
            data=csv_data,
            file_name=f"Rekap_Intelkam_Ciamis_{datetime.today().strftime('%Y%m%d')}.csv",
            mime="text/csv",
        )

sempurnakan kode diatas dikolaborasikan dengan kode sebelumnya tanpa menghilangkan fungsi utama kode diatas untuk convert data menjadi format intelijen
