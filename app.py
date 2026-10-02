import io
import time
from datetime import datetime

import pandas as pd
import plotly.express as px
import streamlit as st
from docx import Document
from google import genai
from google.genai import types

# ---------------------------------------------------------------------------
# Konfigurasi
# ---------------------------------------------------------------------------
st.set_page_config(
    page_title="Sistem Intelkam - Analisis & Generasi Dokumen",
    layout="wide",
    page_icon="🛡",
)

# Nama lengkap di antarmuka -> kode perintah internal untuk prompt
FORMAT_OUTPUT = {
    "SURAT KETERANGAN KEPOLISIAN": "skk",
    "SURAT IZIN": "si",
    "SURAT TANDA TERIMA PEMBERITAHUAN": "sttp",
    "LAPORAN INFORMASI": "li",
    "INFORMASI KHUSUS": "infosus",
    "PERKIRAAN KEADAAN SINGKAT": "kirkat",
}

KATEGORI = [
    "POLITIK",
    "EKONOMI",
    "SOSIAL BUDAYA",
    "KEAMANAN",
    "PENGAWASAN ORANG ASING (POA)",
]
KATEGORI_DEFAULT = "SOSIAL BUDAYA"

# Petunjuk isi tiap kategori agar deteksi otomatis lebih akurat
KATEGORI_KETERANGAN = {
    "POLITIK": "kegiatan partai politik, pemilu/pilkada, demonstrasi atau aksi massa bermuatan politik, dinamika sospol",
    "EKONOMI": "kegiatan perdagangan, perusahaan, harga kebutuhan pokok, ketenagakerjaan, investasi, usaha",
    "SOSIAL BUDAYA": "kegiatan keagamaan (pengajian, ibadah, hari besar), keramaian (konser, hajatan, festival, pertunjukan), olahraga, adat dan budaya",
    "KEAMANAN": "gangguan kamtibmas, konflik, kriminalitas, ancaman keamanan, pengamanan kegiatan",
    "PENGAWASAN ORANG ASING (POA)": "orang asing/WNA: paspor, ITAS/ITAP, sponsor, penjamin, SKK",
}

# Urutan prioritas model (ganti sesuai daftar model aktif di akun Anda)
MODELS_TO_TRY = ["gemini-3.8-flash", "gemini-2.5-flash"]

st.title("🛡️ Sistem Informasi & Rekapitulasi Intelkam Polres Ciamis")
st.caption(
    "Aplikasi Pengolahan Produk Intelijen Baku & Dasbor Rekapitulasi Kegiatan"
)

# ---------------------------------------------------------------------------
# Sidebar
# ---------------------------------------------------------------------------
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

# ---------------------------------------------------------------------------
# State
# ---------------------------------------------------------------------------
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
if "kategori_auto" not in st.session_state:
    st.session_state.kategori_auto = {}  # {signature berkas: kategori}
if "hasil_dokumen" not in st.session_state:
    st.session_state.hasil_dokumen = []  # hasil tetap tampil setelah klik download


# ---------------------------------------------------------------------------
# Fungsi bantu
# ---------------------------------------------------------------------------
def signature(f):
    return f"{f.name}_{f.size}"


def siapkan_berkas(f):
    """Ubah berkas unggahan menjadi konten yang bisa dikirim ke Gemini."""
    data = f.getvalue()
    name = f.name.lower()
    if name.endswith(".txt"):
        return data.decode("utf-8", errors="ignore")
    if name.endswith(".docx"):
        doc = Document(io.BytesIO(data))
        teks = "\n".join(p.text for p in doc.paragraphs)
        for tabel in doc.tables:
            for baris in tabel.rows:
                teks += "\n" + " | ".join(c.text for c in baris.cells)
        return teks
    return types.Part.from_bytes(data=data, mime_type=f.type)


def panggil_gemini(client, contents):
    """Panggil Gemini dengan failover model. 404 langsung pindah model."""
    last_error = None
    for model_name in MODELS_TO_TRY:
        for _ in range(3):
            try:
                resp = client.models.generate_content(
                    model=model_name, contents=contents
                )
                if resp and resp.text:
                    return resp
            except Exception as err:
                last_error = err
                msg = str(err)
                if "404" in msg or "NOT_FOUND" in msg:
                    break
                if "503" in msg or "UNAVAILABLE" in msg or "429" in msg:
                    time.sleep(2)
                    continue
                break
    raise Exception(
        f"Server Gemini sibuk atau model tidak tersedia. Detail: {last_error}"
    )


def deteksi_kategori(client, f):
    """Tentukan kategori kegiatan otomatis dari isi berkas."""
    prompt = (
        "Baca berkas ini lalu tentukan SATU kategori kegiatan yang paling sesuai "
        "dari daftar berikut beserta cakupannya:\n- "
        + "\n- ".join(f"{k}: {KATEGORI_KETERANGAN[k]}" for k in KATEGORI)
        + "\n\nJawab HANYA dengan salah satu teks kategori persis seperti di daftar, "
        "tanpa penjelasan tambahan."
    )
    try:
        resp = panggil_gemini(client, [siapkan_berkas(f), prompt])
        jawab = resp.text.strip().upper()
        for k in KATEGORI:
            if k.upper() == jawab or k.upper() in jawab:
                return k
        # Kecocokan longgar: kata pertama kategori
        for k in KATEGORI:
            if k.split()[0].upper() in jawab:
                return k
    except Exception:
        pass
    return KATEGORI_DEFAULT


def buat_prompt(keyword_cmd, kategori):
    return f"""
    Anda adalah Asisten Intelkam Polres Ciamis. Pelajari data yang diunggah dan buatkan draft teks produk intelijen sesuai kode perintah: {keyword_cmd}.
    Kategori / bidang kegiatan: {kategori}.
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


def buat_docx(teks):
    doc = Document()
    for paragraf in teks.replace("**", "").split("\n"):
        doc.add_paragraph(paragraf)
    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()


# ---------------------------------------------------------------------------
# Tab
# ---------------------------------------------------------------------------
tab1, tab2 = st.tabs(["📄 Produk Intelijen", "📊 Dasbor Rekapitulasi & Statistik"])

# TAB 1: PRODUK INTELIJEN
with tab1:
    st.subheader("1. Unggah Bahan / Berkas Kegiatan")
    uploaded_files = st.file_uploader(
        "Pilih satu atau beberapa file (PDF, TXT, DOCX, atau Foto/Scan Surat)",
        type=["pdf", "docx", "txt", "png", "jpg"],
        accept_multiple_files=True,
    )

    perintah = st.selectbox("Pilih Format Output yang Ingin Dibuat:", list(FORMAT_OUTPUT))

    # Deteksi kategori otomatis untuk setiap berkas baru
    kategori_pilihan = {}
    if uploaded_files:
        st.markdown("**Kategori / Bidang Kegiatan** (terpilih otomatis dari isi berkas, tetap bisa diubah):")
        client_deteksi = genai.Client(api_key=api_key) if api_key else None

        for f in uploaded_files:
            sig = signature(f)
            if sig not in st.session_state.kategori_auto:
                if client_deteksi:
                    with st.spinner(f"Mendeteksi kategori: {f.name}"):
                        st.session_state.kategori_auto[sig] = deteksi_kategori(
                            client_deteksi, f
                        )
                else:
                    st.session_state.kategori_auto[sig] = KATEGORI_DEFAULT

            terdeteksi = st.session_state.kategori_auto[sig]
            c_nama, c_kat = st.columns([4, 6])
            c_nama.write(f"📎 {f.name}")
            kategori_pilihan[sig] = c_kat.selectbox(
                f"Kategori {f.name}",
                KATEGORI,
                index=KATEGORI.index(terdeteksi),
                key=f"kat_{sig}",
                label_visibility="collapsed",
            )

        if not api_key:
            st.info("Masukkan API Key agar kategori dapat terdeteksi otomatis.")

    if st.button("🚀 Proses & Buat Dokumen"):
        if not api_key:
            st.error(
                "Silakan masukkan Gemini API Key terlebih dahulu di sidebar atau atur di Secrets!"
            )
        elif not uploaded_files:
            st.warning("Silakan unggah minimal satu berkas bahan terlebih dahulu!")
        else:
            client = genai.Client(api_key=api_key)
            keyword_cmd = FORMAT_OUTPUT[perintah]
            st.session_state.hasil_dokumen = []
            progres = st.progress(0.0)

            for i, f in enumerate(uploaded_files):
                sig = signature(f)
                kategori = kategori_pilihan.get(sig, KATEGORI_DEFAULT)
                try:
                    with st.spinner(f"Memproses {f.name} ({i + 1}/{len(uploaded_files)})..."):
                        resp = panggil_gemini(
                            client,
                            [siapkan_berkas(f), buat_prompt(keyword_cmd, kategori)],
                        )
                    teks = resp.text
                    st.session_state.hasil_dokumen.append(
                        {
                            "nama": f.name,
                            "perintah": perintah,
                            "kode": keyword_cmd,
                            "teks": teks,
                            "docx": buat_docx(teks),
                        }
                    )
                    st.session_state.db_kegiatan = pd.concat(
                        [
                            st.session_state.db_kegiatan,
                            pd.DataFrame(
                                [
                                    {
                                        "Tanggal": datetime.now().strftime("%Y-%m-%d"),
                                        "Nama Kegiatan": f.name.rsplit(".", 1)[0],
                                        "Kategori": kategori,
                                        "Penanggung Jawab": "Tercatat di Dokumen",
                                        "Jumlah Massa": "1 / Sesuai Data",
                                        "Produk Terbit": perintah,
                                    }
                                ]
                            ),
                        ],
                        ignore_index=True,
                    )
                except Exception as e:
                    st.error(f"Gagal memproses {f.name}: {e}")
                progres.progress((i + 1) / len(uploaded_files))

    # Tampilkan hasil (tetap ada setelah klik tombol download)
    if st.session_state.hasil_dokumen:
        st.success(f"{len(st.session_state.hasil_dokumen)} dokumen berhasil diproses!")
        st.subheader("📝 Hasil Dokumen:")
        for idx, h in enumerate(st.session_state.hasil_dokumen):
            with st.expander(f"{h['perintah']} — {h['nama']}", expanded=(idx == 0)):
                st.text_area(
                    "Hasil Teks Baku:", h["teks"], height=350, key=f"teks_{idx}"
                )
                st.download_button(
                    label="📥 Download Sebagai File Word (.docx)",
                    data=h["docx"],
                    file_name=f"{h['kode'].upper()}_{h['nama'].rsplit('.', 1)[0]}_{datetime.now().strftime('%Y%m%d')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    key=f"dl_{idx}",
                )

# TAB 2: DASBOR REKAPITULASI
with tab2:
    st.subheader("📊 Dasbor Rekapitulasi Kegiatan Intelkam")

    if st.session_state.db_kegiatan.empty:
        st.info("Belum ada data kegiatan terdaftar. Silakan unggah dokumen di tab Produk Intelijen.")
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
            df_filtered["Kategori"].mode()[0] if not df_filtered.empty else "-",
        )
        m3.metric("Produk Terbit", len(df_filtered["Produk Terbit"].dropna()))

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
