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
    "Aplikasi Pengolahan Produk Intelijen Baku (SKK, SI, STTP, LI, INFOSUS, KIRKAT) & Dasbor Rekapitulasi Kegiatan"
)

# Sidebar Input API Key & Konfigurasi Sistem
st.sidebar.header("⚙️ Pengaturan Sistem")

# Logika Membaca API Key Permanen dari secrets.toml atau input manual
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

# Inisialisasi Database Lokal / Catatan Kegiatan
if "db_kegiatan" not in st.session_state:
    st.session_state.db_kegiatan = pd.DataFrame(
        columns=[
            "Tanggal",
            "Nama Kegiatan",
            "Kategori",
            "Penanggung Jawab",
            "Jumlah Massa",
            "Produk Terbit",
        ]
    )

# Tab Navigasi
tab1, tab2 = st.tabs(
    ["📄 Upload & Penanganan Dokumen", "📊 Dasbor Rekapitulasi & Statistik"]
)

# TAB 1: UPLOAD & GENERATOR DOKUMEN
with tab1:
    st.subheader("1. Unggah Bahan / Berkas Kegiatan")
    uploaded_file = st.file_uploader(
        "Pilih file (PDF, TXT, DOCX, atau Foto/Scan Surat)",
        type=["pdf", "docx", "txt", "png", "jpg"],
    )

    col1, col2 = st.columns(2)
    with col1:
        perintah = st.selectbox(
            "Pilih Format Output yang Ingin Dibuat (Kata Kunci):",
            [
                "skk (Surat Keterangan Kepolisian / POA)",
                "si (Surat Izin Kepolisian)",
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
                "YANMAS / POA (Pengawasan Orang Asing / SKK)",
                "OLAHRAGA",
                "KEAGAMAAN",
                "SOSBUD / KERAMAIAN",
                "POLITIK / SOSPOL",
                "EKONOMI",
            ],
        )

    if st.button("🚀 Proses & Buat Dokumen"):
        if not api_key:
            st.error(
                "Silakan masukkan Gemini API Key terlebih dahulu di sidebar atau atur di Secrets!"
            )
        elif not uploaded_file:
            st.warning("Silakan unggah berkas bahan terlebih dahulu!")
        else:
            with st.spinner(
                "AI sedang menganalisis berkas dan menyusun draf baku sesuai format Sat Intelkam..."
            ):
                try:
                    client = genai.Client(api_key=api_key)

                    bytes_data = uploaded_file.getvalue()
                    mime_type = uploaded_file.type

                    # Penanganan tipe berkas
                    if mime_type == "text/plain":
                        file_part = bytes_data.decode("utf-8")
                    else:
                        file_part = types.Part.from_bytes(
                            data=bytes_data,
                            mime_type=mime_type,
                        )

                    # Ekstraksi kata kunci perintah
                    keyword_cmd = perintah.split()[0].lower()

                    prompt = f"""
                    Anda adalah Asisten Intelkam Polres Ciamis. Pelajari data yang diunggah dan buatkan draft teks produk intelijen sesuai kode perintah: {keyword_cmd}.
                    Aturan Ketat & Format Baku Sat Intelkam Polres Ciamis:
                    - Pejabat Penandatangan Resmi: KASAT INTELKAM POLRES CIAMIS, AKP RAHMAT KOMARA, S.H., M.H., AJUN KOMISARIS POLISI NRP 70030155.
                    - Gunakan bahasa baku dinas Kepolisian Republik Indonesia, lengkap, terstruktur, tidak disingkat sembarangan.
                    - Jika perintah 'skk': Buat Surat Keterangan Kepolisian (SKK) untuk Orang Asing/POA, sertakan detail identitas, Sponsor, Paspor, ITAS/ITAP, Penjamin, dan Masa Berlaku.
                    - Jika perintah 'si': Buat Surat Izin dengan struktur tabel 2 kolom (Pertimbangan, Dasar, Memperhatikan, Memberikan Izin) dan 4 poin catatan/kewajiban.
                    - Jika perintah 'sttp': Buat Surat Tanda Terima Pemberitahuan sesuai format baku.
                    - Jika perintah 'li': Buat Laporan Informasi dengan header baku (Sumber, Hubungan, Cara, Waktu, Nilai A-1), Fakta-Fakta 5W+1H, Analisa, Prediksi, Langkah-langkah, dan Rekomendasi secara kaya.
                    - Jika perintah 'infosus': Buat Nota Dinas Pengantar kepada Kapolres Ciamis dan Lembar Informasi Khusus berklasifikasi RAHASIA lengkap dengan distribusi baku.
                    - Jika perintah 'kirkat': Buat Perkiraan Keadaan Intelijen Singkat (Nota Dinas, Pendahuluan terperinci, Keadaan Sasaran, Analisa, Kesimpulan, dan Saran).
                    """

                    # Daftar model prioritas terbaru dengan mekanisme failover (mengatasi 404 / 503)
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
                            f"Server Google Gemini sedang sibuk atau model tidak tersedia. Detail: {last_error}"
                        )

                    st.success("Dokumen Berhasil Diproses!")
                    hasil_teks = response.text

                    st.subheader("📝 Pratinjau Teks Output:")
                    st.text_area("Hasil Teks Baku:", hasil_teks, height=350)

                    # Export ke .docx (Word)
                    doc = Document()
                    for paragraph in hasil_teks.split("\n"):
                        doc.add_paragraph(paragraph)

                    bio = io.BytesIO()
                    doc.save(bio)

                    st.download_button(
                        label="📥 Download Sebagai File Word (.docx)",
                        data=bio.getvalue(),
                        file_name=f"Produk_Intelkam_{keyword_cmd.upper()}_{datetime.now().strftime('%Y%m%d')}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    )

                    # Simpan data ke rekapitulasi lokal
                    new_entry = {
                        "Tanggal": datetime.now().strftime("%Y-%m-%d"),
                        "Nama Kegiatan": uploaded_file.name.split(".")[0],
                        "Kategori": kategori_kegiatan,
                        "Penanggung Jawab": "Tercatat di Dokumen",
                        "Jumlah Massa": "1 / Sesuai Data",
                        "Produk Terbit": keyword_cmd.upper(),
                    }
                    st.session_state.db_kegiatan = pd.concat(
                        [
                            st.session_state.db_kegiatan,
                            pd.DataFrame([new_entry]),
                        ],
                        ignore_index=True,
                    )

                except Exception as e:
                    st.error(f"Terjadi kesalahan saat memproses: {e}")

# TAB 2: DASBOR REKAPITULASI
with tab2:
    st.subheader("📊 Dasbor Rekapitulasi Kegiatan Intelkam")

    if st.session_state.db_kegiatan.empty:
        st.info(
            "Belum ada data kegiatan terdaftar. Silakan unggah dokumen di Tab 1."
        )
    else:
        col_f1, col_f2 = st.columns(2)
        with col_f1:
            kat_filter = st.multiselect(
                "Filter Kategori:",
                options=st.session_state.db_kegiatan["Kategori"].unique(),
                default=st.session_state.db_kegiatan["Kategori"].unique(),
            )

        df_filtered = st.session_state.db_kegiatan[
            st.session_state.db_kegiatan["Kategori"].isin(kat_filter)
        ]

        m1, m2, m3 = st.columns(3)
        m1.metric("Total Kegiatan", len(df_filtered))
        m2.metric(
            "Kategori Terbanyak",
            (
                df_filtered["Kategori"].mode()[0]
                if not df_filtered.empty
                else "-"
            ),
        )
        m3.metric(
            "Produk Terbit", len(df_filtered["Produk Terbit"].dropna())
        )

        st.markdown("---")

        c1, c2 = st.columns([6, 4])
        with c1:
            st.write("### Daftar Rekapitulasi Kegiatan")
            st.dataframe(df_filtered, use_container_width=True)

        with c2:
            st.write("### Grafik Sebaran Kategori")
            fig = px.pie(
                df_filtered,
                names="Kategori",
                title="Persentase Kegiatan per Bidang",
                hole=0.4,
            )
            st.plotly_chart(fig, use_container_width=True)
