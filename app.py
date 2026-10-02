import base64
import io
import json
import os
import re
import time
from datetime import datetime

import pandas as pd
import plotly.express as px
import streamlit as st
from docx import Document

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

KATEGORI_KETERANGAN = {
    "POLITIK": "kegiatan partai politik, pemilu/pilkada, demonstrasi atau aksi massa bermuatan politik, dinamika sospol",
    "EKONOMI": "kegiatan perdagangan, perusahaan, harga kebutuhan pokok, ketenagakerjaan, investasi, usaha",
    "SOSIAL BUDAYA": "kegiatan keagamaan (pengajian, ibadah, hari besar), keramaian (konser, hajatan, festival, pertunjukan), olahraga, adat dan budaya",
    "KEAMANAN": "gangguan kamtibmas, konflik, kriminalitas, ancaman keamanan, pengamanan kegiatan",
    "PENGAWASAN ORANG ASING (POA)": "orang asing/WNA: paspor, ITAS/ITAP, sponsor, penjamin, SKK",
}

# Daftar AI yang didukung. Ubah daftar "models" sesuai model aktif di akun Anda.
# Model pertama dicoba lebih dulu, lalu berikutnya jika gagal.
PROVIDERS = {
    "Gemini (Google)": {
        "secret": "GEMINI_API_KEY",
        "models": ["gemini-3.8-flash"],
    },
    "Claude (Anthropic)": {
        "secret": "ANTHROPIC_API_KEY",
        "models": ["claude-sonnet-5-5"],
    },
    "ChatGPT (OpenAI)": {
        "secret": "OPENAI_API_KEY",
        "models": ["gpt-4.1"],
    },
}

st.title("🛡️ Sistem Informasi & Rekapitulasi Intelkam Polres Ciamis")
st.caption(
    "Aplikasi Pengolahan Produk Intelijen Baku & Dasbor Rekapitulasi Kegiatan"
)

# ---------------------------------------------------------------------------
# Sidebar: API key tiap AI
# ---------------------------------------------------------------------------
st.sidebar.header("⚙️ Pengaturan Sistem")

api_keys = {}
for nama, cfg in PROVIDERS.items():
    if cfg["secret"] in st.secrets and st.secrets[cfg["secret"]]:
        st.sidebar.success(f"🔑 {nama}: key permanen aktif")
        api_keys[nama] = st.secrets[cfg["secret"]]
    else:
        isi = st.sidebar.text_input(
            f"API Key {nama}:", type="password", key=f"key_{cfg['secret']}"
        )
        if isi:
            api_keys[nama] = isi

if api_keys:
    ai_utama = st.sidebar.selectbox("AI Utama:", list(api_keys))
    pakai_cadangan = st.sidebar.checkbox(
        "Pakai AI lain sebagai cadangan jika gagal",
        value=True,
        disabled=len(api_keys) < 2,
    )
    URUTAN_AI = [ai_utama] + (
        [n for n in api_keys if n != ai_utama] if pakai_cadangan else []
    )
else:
    URUTAN_AI = []
    st.sidebar.info("Isi minimal satu API Key untuk mulai.")

# ---------------------------------------------------------------------------
# State
# ---------------------------------------------------------------------------
KOLOM_DB = [
    "Tanggal",
    "Nama Kegiatan",
    "Kategori",
    "Penanggung Jawab",
    "Jumlah Massa",
    "Produk Terbit",
]
PRODUK_OPSI = list(FORMAT_OUTPUT) + ["LAINNYA"]
DB_FILE = "data_rekap.csv"


def simpan_db():
    """Simpan rekap ke file CSV lokal agar tidak hilang saat halaman di-refresh."""
    try:
        st.session_state.db_kegiatan.to_csv(DB_FILE, index=False, encoding="utf-8-sig")
    except Exception as e:
        st.sidebar.warning(f"Rekap gagal disimpan ke file: {e}")


if "db_kegiatan" not in st.session_state:
    if os.path.exists(DB_FILE):
        try:
            st.session_state.db_kegiatan = pd.read_csv(DB_FILE, dtype=str).fillna("")
        except Exception:
            st.session_state.db_kegiatan = pd.DataFrame(columns=KOLOM_DB)
    else:
        st.session_state.db_kegiatan = pd.DataFrame(columns=KOLOM_DB)
if "ver_rekap" not in st.session_state:
    st.session_state.ver_rekap = 0
if "import_preview" not in st.session_state:
    st.session_state.import_preview = None
if "kategori_auto" not in st.session_state:
    st.session_state.kategori_auto = {}
if "hasil_dokumen" not in st.session_state:
    st.session_state.hasil_dokumen = []


# ---------------------------------------------------------------------------
# Lapisan AI (satu antarmuka untuk semua penyedia)
# ---------------------------------------------------------------------------
def siapkan_berkas(f):
    """Bentuk netral: {'teks': ...} atau {'data': bytes, 'mime': ...}."""
    data = f.getvalue()
    name = f.name.lower()
    if name.endswith(".txt"):
        return {"teks": data.decode("utf-8", errors="ignore")}
    if name.endswith(".docx"):
        doc = Document(io.BytesIO(data))
        teks = "\n".join(p.text for p in doc.paragraphs)
        for tabel in doc.tables:
            for baris in tabel.rows:
                teks += "\n" + " | ".join(c.text for c in baris.cells)
        return {"teks": teks}
    return {"data": data, "mime": f.type, "nama": f.name}


def _gemini(key, model, berkas, prompt):
    from google import genai
    from google.genai import types

    client = genai.Client(api_key=key)
    if "teks" in berkas:
        isi = berkas["teks"]
    else:
        isi = types.Part.from_bytes(data=berkas["data"], mime_type=berkas["mime"])
    return client.models.generate_content(model=model, contents=[isi, prompt]).text


def _claude(key, model, berkas, prompt):
    import anthropic

    client = anthropic.Anthropic(api_key=key)
    if "teks" in berkas:
        blok = [{"type": "text", "text": berkas["teks"]}]
    else:
        b64 = base64.standard_b64encode(berkas["data"]).decode()
        tipe = "document" if berkas["mime"] == "application/pdf" else "image"
        blok = [
            {
                "type": tipe,
                "source": {
                    "type": "base64",
                    "media_type": berkas["mime"],
                    "data": b64,
                },
            }
        ]
    blok.append({"type": "text", "text": prompt})
    r = client.messages.create(
        model=model,
        max_tokens=8000,
        messages=[{"role": "user", "content": blok}],
    )
    return "".join(b.text for b in r.content if b.type == "text")


def _openai(key, model, berkas, prompt):
    from openai import OpenAI

    client = OpenAI(api_key=key)
    if "teks" in berkas:
        blok = [{"type": "input_text", "text": berkas["teks"]}]
    else:
        b64 = base64.standard_b64encode(berkas["data"]).decode()
        url = f"data:{berkas['mime']};base64,{b64}"
        if berkas["mime"] == "application/pdf":
            blok = [
                {
                    "type": "input_file",
                    "filename": berkas.get("nama", "berkas.pdf"),
                    "file_data": url,
                }
            ]
        else:
            blok = [{"type": "input_image", "image_url": url}]
    blok.append({"type": "input_text", "text": prompt})
    r = client.responses.create(
        model=model, input=[{"role": "user", "content": blok}]
    )
    return r.output_text


FUNGSI_AI = {
    "Gemini (Google)": _gemini,
    "Claude (Anthropic)": _claude,
    "ChatGPT (OpenAI)": _openai,
}


def panggil_ai(berkas, prompt):
    """Coba AI sesuai urutan; kembalikan (teks, label AI yang berhasil)."""
    if not URUTAN_AI:
        raise Exception("Belum ada API Key yang diisi.")
    errors = []
    for prov in URUTAN_AI:
        for model in PROVIDERS[prov]["models"]:
            for _ in range(3):
                try:
                    teks = FUNGSI_AI[prov](api_keys[prov], model, berkas, prompt)
                    if teks:
                        return teks, f"{prov} · {model}"
                    break
                except Exception as err:
                    msg = str(err)
                    errors.append(f"{prov}/{model}: {msg[:160]}")
                    sementara = (
                        "503" in msg
                        or "UNAVAILABLE" in msg
                        or "overloaded" in msg.lower()
                    )
                    if sementara:
                        time.sleep(2)
                        continue
                    break  # 404, auth, dll: langsung model/AI berikutnya
    if any("429" in e or "RESOURCE_EXHAUSTED" in e for e in errors):
        raise Exception(
            "Kuota API habis (429). Tunggu kuota reset, aktifkan billing, "
            "atau isi API Key AI lain di sidebar."
        )
    raise Exception("Semua AI gagal. Detail: " + " | ".join(errors[-3:]))


# ---------------------------------------------------------------------------
# Fungsi bantu aplikasi
# ---------------------------------------------------------------------------
def signature(f):
    return f"{f.name}_{f.size}"


def deteksi_kategori(f):
    prompt = (
        "Baca berkas ini lalu tentukan SATU kategori kegiatan yang paling sesuai "
        "dari daftar berikut beserta cakupannya:\n- "
        + "\n- ".join(f"{k}: {KATEGORI_KETERANGAN[k]}" for k in KATEGORI)
        + "\n\nJawab HANYA dengan salah satu teks kategori persis seperti di daftar, "
        "tanpa penjelasan tambahan."
    )
    try:
        teks, _ = panggil_ai(siapkan_berkas(f), prompt)
        jawab = teks.strip().upper()
        for k in KATEGORI:
            if k.upper() == jawab or k.upper() in jawab:
                return k
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


def rapikan_baris(d):
    """Normalisasi hasil ekstraksi AI ke kolom rekap."""
    kat = str(d.get("kategori", "")).strip().upper()
    kategori = next((k for k in KATEGORI if k.upper() == kat), None)
    if kategori is None:
        kategori = next((k for k in KATEGORI if k.split()[0].upper() in kat), KATEGORI_DEFAULT)
    prod = str(d.get("produk_terbit", "")).strip().upper()
    produk = next((p for p in FORMAT_OUTPUT if p == prod), None)
    if produk is None:
        produk = next((p for p in FORMAT_OUTPUT if prod and prod in p), "LAINNYA")
    return {
        "Tanggal": str(d.get("tanggal", "")).strip(),
        "Nama Kegiatan": str(d.get("nama_kegiatan", "")).strip(),
        "Kategori": kategori,
        "Penanggung Jawab": str(d.get("penanggung_jawab", "")).strip(),
        "Jumlah Massa": str(d.get("jumlah_massa", "")).strip(),
        "Produk Terbit": produk,
    }


BULAN_ID = {
    "januari": 1, "februari": 2, "maret": 3, "april": 4, "mei": 5, "juni": 6,
    "juli": 7, "agustus": 8, "september": 9, "oktober": 10, "november": 11,
    "desember": 12,
}


def cari_tanggal(teks):
    """Cari tanggal pertama di teks (YYYY-MM-DD, DD-MM-YYYY, atau '5 Mei 2026')."""
    pola = [
        (r"(\d{4})-(\d{1,2})-(\d{1,2})", lambda m: (m[0], m[1], m[2])),
        (r"(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})", lambda m: (m[2], m[1], m[0])),
        (
            r"(\d{1,2})\s+(" + "|".join(BULAN_ID) + r")\s+(\d{4})",
            lambda m: (m[2], BULAN_ID[m[1].lower()], m[0]),
        ),
    ]
    for rx, ambil in pola:
        for m in re.finditer(rx, teks, flags=re.I):
            try:
                y, mo, d = (int(x) for x in ambil(m.groups()))
                return datetime(y, mo, d).strftime("%Y-%m-%d")
            except Exception:
                continue
    return ""


def fallback_tanpa_ai(f):
    """Isi baris rekap tanpa AI: dari nama file (awalan LI/SI/dst.) dan isi teks."""
    nama = f.name.rsplit(".", 1)[0]
    produk = "LAINNYA"
    kode = {v.upper(): k for k, v in FORMAT_OUTPUT.items()}
    m = re.match(r"^\s*(SKK|SI|STTP|LI|INFOSUS|KIRKAT)\s*[=\-\u2013:_]+\s*(.*)$", nama, re.I)
    if m:
        produk = kode.get(m.group(1).upper(), "LAINNYA")
        nama = m.group(2).strip() or nama
    try:
        teks = siapkan_berkas(f).get("teks", "")
    except Exception:
        teks = ""
    return {
        "Tanggal": cari_tanggal(teks),
        "Nama Kegiatan": nama,
        "Kategori": KATEGORI_DEFAULT,
        "Penanggung Jawab": "",
        "Jumlah Massa": "",
        "Produk Terbit": produk,
    }


def ekstrak_data_lama(f):
    """Ekstrak data rekap dari satu dokumen lama memakai AI."""
    prompt = (
        "Baca dokumen intelijen ini lalu ekstrak data untuk rekapitulasi. "
        "Jawab HANYA satu objek JSON valid tanpa teks lain, dengan kunci:\n"
        '{"tanggal": "YYYY-MM-DD (tanggal surat/kegiatan, kosongkan jika tidak ada)", '
        '"nama_kegiatan": "...", '
        f'"kategori": "salah satu dari: {", ".join(KATEGORI)}", '
        '"penanggung_jawab": "nama penyelenggara/pemohon/penanggung jawab", '
        '"jumlah_massa": "perkiraan jumlah massa/peserta jika ada", '
        f'"produk_terbit": "salah satu dari: {", ".join(FORMAT_OUTPUT)}"}}\n'
        "Jangan mengarang: jika data tidak ada di dokumen, isi dengan string kosong."
    )
    teks, _ = panggil_ai(siapkan_berkas(f), prompt)
    awal, akhir = teks.find("{"), teks.rfind("}")
    if awal == -1 or akhir == -1:
        raise Exception("AI tidak mengembalikan JSON")
    data = json.loads(teks[awal : akhir + 1])
    baris = rapikan_baris(data)
    if not baris["Nama Kegiatan"]:
        baris["Nama Kegiatan"] = f.name.rsplit(".", 1)[0]
    return baris


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

    perintah = st.selectbox(
        "Pilih Format Output yang Ingin Dibuat:", list(FORMAT_OUTPUT)
    )

    kategori_pilihan = {}
    if uploaded_files:
        st.markdown(
            "**Kategori / Bidang Kegiatan** (terpilih otomatis dari isi berkas, tetap bisa diubah):"
        )
        for f in uploaded_files:
            sig = signature(f)
            if sig not in st.session_state.kategori_auto:
                if URUTAN_AI:
                    with st.spinner(f"Mendeteksi kategori: {f.name}"):
                        st.session_state.kategori_auto[sig] = deteksi_kategori(f)
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

        if not URUTAN_AI:
            st.info("Masukkan API Key agar kategori dapat terdeteksi otomatis.")

    if st.button("🚀 Proses & Buat Dokumen"):
        if not URUTAN_AI:
            st.error("Silakan isi minimal satu API Key di sidebar atau atur di Secrets!")
        elif not uploaded_files:
            st.warning("Silakan unggah minimal satu berkas bahan terlebih dahulu!")
        else:
            keyword_cmd = FORMAT_OUTPUT[perintah]
            st.session_state.hasil_dokumen = []
            progres = st.progress(0.0)

            for i, f in enumerate(uploaded_files):
                sig = signature(f)
                kategori = kategori_pilihan.get(sig, KATEGORI_DEFAULT)
                try:
                    with st.spinner(
                        f"Memproses {f.name} ({i + 1}/{len(uploaded_files)})..."
                    ):
                        teks, ai_dipakai = panggil_ai(
                            siapkan_berkas(f), buat_prompt(keyword_cmd, kategori)
                        )
                    st.session_state.hasil_dokumen.append(
                        {
                            "nama": f.name,
                            "perintah": perintah,
                            "kode": keyword_cmd,
                            "teks": teks,
                            "docx": buat_docx(teks),
                            "ai": ai_dipakai,
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
                    simpan_db()
                except Exception as e:
                    st.error(f"Gagal memproses {f.name}: {e}")
                progres.progress((i + 1) / len(uploaded_files))

    if st.session_state.hasil_dokumen:
        st.success(f"{len(st.session_state.hasil_dokumen)} dokumen berhasil diproses!")
        st.subheader("📝 Hasil Dokumen:")
        for idx, h in enumerate(st.session_state.hasil_dokumen):
            with st.expander(f"{h['perintah']} — {h['nama']}", expanded=(idx == 0)):
                st.caption(f"Dibuat oleh: {h['ai']}")
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

    # ---------------- Impor data lama ----------------
    with st.expander(
        "📥 Upload Berkas Lama (Januari s.d. sekarang)",
        expanded=st.session_state.db_kegiatan.empty,
    ):
        st.caption(
            "Unggah dokumen yang sudah pernah dibuat (PDF, DOCX, TXT, foto/scan) agar "
            "datanya otomatis diekstrak ke rekapitulasi, atau unggah file rekap "
            "CSV/Excel dengan kolom: " + ", ".join(KOLOM_DB) + "."
        )
        berkas_impor = st.file_uploader(
            "Pilih berkas lama",
            type=["pdf", "docx", "txt", "png", "jpg", "csv", "xlsx"],
            accept_multiple_files=True,
            key="uploader_impor",
        )

        if st.button("🔍 Analisis Berkas Lama"):
            if not berkas_impor:
                st.warning("Silakan unggah minimal satu berkas.")
            else:
                baris, gagal = [], []
                progres = st.progress(0.0)
                for i, f in enumerate(berkas_impor):
                    nama = f.name.lower()
                    try:
                        if nama.endswith((".csv", ".xlsx")):
                            df_in = (
                                pd.read_csv(f, dtype=str)
                                if nama.endswith(".csv")
                                else pd.read_excel(f, dtype=str)
                            ).fillna("")
                            for kol in KOLOM_DB:
                                if kol not in df_in.columns:
                                    df_in[kol] = ""
                            baris.extend(df_in[KOLOM_DB].to_dict("records"))
                        else:
                            if not URUTAN_AI:
                                raise Exception("Belum ada API Key untuk analisis AI.")
                            with st.spinner(f"Menganalisis {f.name} ({i + 1}/{len(berkas_impor)})..."):
                                baris.append(ekstrak_data_lama(f))
                            time.sleep(3)  # jeda agar tidak melewati batas permintaan per menit
                    except Exception as e:
                        gagal.append(f"{f.name}: {str(e)[:300]}")
                        baris.append(fallback_tanpa_ai(f))
                    progres.progress((i + 1) / len(berkas_impor))

                df_prev = pd.DataFrame(baris, columns=KOLOM_DB).fillna("")
                if not df_prev.empty:
                    tgl = pd.to_datetime(df_prev["Tanggal"], errors="coerce")
                    df_prev["Tanggal"] = tgl.dt.strftime("%Y-%m-%d").fillna("")
                st.session_state.import_preview = df_prev
                for g in gagal:
                    st.warning(f"AI gagal, baris diisi otomatis dari nama file/isi teks (periksa manual): {g}")

        prev = st.session_state.get("import_preview")
        if prev is not None and not prev.empty:
            st.markdown(
                "**Pratinjau hasil ekstraksi**: periksa dan koreksi dulu. "
                "Hasil AI bisa keliru pada tanggal atau nama. "
                "Tanggal wajib format `YYYY-MM-DD`. Baris yang tidak perlu bisa dihapus."
            )
            edited_prev = st.data_editor(
                prev,
                num_rows="dynamic",
                use_container_width=True,
                key="editor_preview",
                column_config={
                    "Kategori": st.column_config.SelectboxColumn(
                        "Kategori", options=KATEGORI
                    ),
                    "Produk Terbit": st.column_config.SelectboxColumn(
                        "Produk Terbit", options=PRODUK_OPSI
                    ),
                },
            )
            b1, b2 = st.columns(2)
            if b1.button("✅ Simpan ke Rekapitulasi"):
                df_baru = edited_prev.fillna("").copy()
                tgl = pd.to_datetime(df_baru["Tanggal"], errors="coerce")
                if df_baru.empty:
                    st.warning("Tidak ada baris untuk disimpan.")
                elif tgl.isna().any():
                    st.error(
                        f"{int(tgl.isna().sum())} baris punya tanggal kosong/tidak valid. "
                        "Perbaiki dulu di tabel."
                    )
                else:
                    df_baru["Tanggal"] = tgl.dt.strftime("%Y-%m-%d")
                    lama = st.session_state.db_kegiatan
                    gabung = pd.concat([lama, df_baru], ignore_index=True)
                    gabung = gabung.drop_duplicates(
                        subset=["Tanggal", "Nama Kegiatan", "Produk Terbit"],
                        keep="first",
                    ).reset_index(drop=True)
                    ditambah = len(gabung) - len(lama)
                    st.session_state.db_kegiatan = gabung
                    simpan_db()
                    st.session_state.import_preview = None
                    st.session_state.ver_rekap += 1
                    st.success(
                        f"{ditambah} data ditambahkan "
                        f"({len(df_baru) - ditambah} duplikat dilewati)."
                    )
                    st.rerun()
            if b2.button("❌ Batalkan"):
                st.session_state.import_preview = None
                st.rerun()

    # ---------------- Dasbor ----------------
    db = st.session_state.db_kegiatan

    if db.empty:
        st.info(
            "Belum ada data kegiatan terdaftar. Unggah dokumen di tab Produk Intelijen "
            "atau impor berkas lama di atas."
        )
    else:
        db_tgl = pd.to_datetime(db["Tanggal"], errors="coerce")
        tgl_min = db_tgl.min().date() if db_tgl.notna().any() else datetime.now().date()
        tgl_max = db_tgl.max().date() if db_tgl.notna().any() else datetime.now().date()

        col_f1, col_f2 = st.columns(2)
        with col_f1:
            kat_filter = st.multiselect(
                "Filter Kategori:",
                options=sorted(db["Kategori"].unique()),
                default=sorted(db["Kategori"].unique()),
            )
        with col_f2:
            periode = st.date_input(
                "Filter Periode:",
                value=(tgl_min, tgl_max),
                min_value=min(tgl_min, tgl_max),
                max_value=max(tgl_max, datetime.now().date()),
            )

        mask = db["Kategori"].isin(kat_filter)
        if isinstance(periode, (tuple, list)) and len(periode) == 2:
            mask &= (db_tgl.dt.date >= periode[0]) & (db_tgl.dt.date <= periode[1])
        df_filtered = db[mask]

        m1, m2, m3 = st.columns(3)
        m1.metric("Total Kegiatan", len(df_filtered))
        m2.metric(
            "Kategori Terbanyak",
            df_filtered["Kategori"].mode()[0] if not df_filtered.empty else "-",
        )
        m3.metric("Produk Terbit", len(df_filtered["Produk Terbit"].replace("", pd.NA).dropna()))

        st.markdown("---")

        c1, c2 = st.columns([6, 4])
        with c1:
            st.write("### Daftar Rekapitulasi Kegiatan")
            df_tampil = df_filtered.copy()
            df_tampil.insert(0, "Hapus", False)
            edited = st.data_editor(
                df_tampil,
                hide_index=True,
                use_container_width=True,
                disabled=KOLOM_DB,
                column_config={"Hapus": st.column_config.CheckboxColumn("Hapus")},
                key=f"editor_rekap_{st.session_state.ver_rekap}",
            )
            terpilih = edited.index[edited["Hapus"]]

            d1, d2 = st.columns(2)
            if d1.button(
                f"🗑️ Hapus {len(terpilih)} data terpilih",
                disabled=len(terpilih) == 0,
            ):
                st.session_state.db_kegiatan = db.drop(index=terpilih).reset_index(drop=True)
                simpan_db()
                st.session_state.ver_rekap += 1
                st.rerun()

            with d2:
                yakin = st.checkbox("Saya yakin menghapus SEMUA data")
                if st.button("⚠️ Hapus Semua Data", disabled=not yakin):
                    st.session_state.db_kegiatan = pd.DataFrame(columns=KOLOM_DB)
                    simpan_db()
                    st.session_state.ver_rekap += 1
                    st.rerun()

            st.download_button(
                "💾 Unduh Cadangan Rekap (CSV)",
                data=db.to_csv(index=False).encode("utf-8-sig"),
                file_name=f"rekap_intelkam_{datetime.now().strftime('%Y%m%d')}.csv",
                mime="text/csv",
            )

        with c2:
            st.write("### Grafik Sebaran Kategori")
            if df_filtered.empty:
                st.info("Tidak ada data pada filter ini.")
            else:
                fig = px.pie(
                    df_filtered,
                    names="Kategori",
                    title="Persentase Kegiatan per Bidang",
                    hole=0.4,
                )
                st.plotly_chart(fig, use_container_width=True)
