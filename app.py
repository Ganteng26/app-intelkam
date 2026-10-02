
import io
import re
import sqlite3
import hashlib
from datetime import datetime, date

import pandas as pd
import streamlit as st

# Optional document readers
try:
    from pypdf import PdfReader
except Exception:
    PdfReader = None

try:
    from docx import Document
except Exception:
    Document = None


# ============================================================
# CONFIG
# ============================================================
st.set_page_config(
    page_title="Dasbor Rekapitulasi Kegiatan Intelkam",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

DB_PATH = "intelkam.db"
UPLOAD_DIR = "uploads"

DOC_TYPES = [
    "Surat Izin (SI)",
    "Laporan Informasi (LI)",
    "Informasi Khusus (Infosus)",
    "Kirka/Kirkat",
    "Surat Keterangan Kepolisian (SKP)",
]

ACTIVITY_CATEGORIES = [
    "Konser/Hiburan",
    "Kompetisi/Pertandingan",
    "Kegiatan Masyarakat",
    "Kegiatan Pemerintahan",
    "Lainnya",
]

SUBCATEGORIES = {
    "Konser/Hiburan": [
        "Konser musik",
        "Pentas seni",
        "Festival",
        "Hiburan masyarakat",
        "Entertainment lainnya",
    ],
    "Kompetisi/Pertandingan": [
        "Kompetisi olahraga",
        "Turnamen",
        "Perlombaan",
        "Pertandingan",
        "Kejuaraan",
    ],
    "Kegiatan Masyarakat": [
        "Aksi/unjuk rasa",
        "Pengajian/keagamaan",
        "Seminar",
        "Rapat",
        "Sosialisasi",
        "Kegiatan organisasi",
        "Kegiatan masyarakat lainnya",
    ],
    "Kegiatan Pemerintahan": [
        "Kegiatan pemerintah daerah",
        "Kunjungan pejabat",
        "Upacara",
        "Peresmian",
        "Kegiatan kedinasan",
    ],
    "Lainnya": ["Lainnya"],
}

MONTHS = [
    "Januari", "Februari", "Maret", "April", "Mei", "Juni",
    "Juli", "Agustus", "September", "Oktober", "November", "Desember"
]


# ============================================================
# DATABASE
# ============================================================
def get_conn():
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    conn = get_conn()
    conn.execute("""
        CREATE TABLE IF NOT EXISTS documents (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            activity_id TEXT NOT NULL,
            file_name TEXT NOT NULL,
            file_hash TEXT UNIQUE,
            file_path TEXT,
            document_number TEXT,
            document_date TEXT,
            document_type TEXT NOT NULL,
            activity_name TEXT,
            category TEXT NOT NULL,
            subcategory TEXT,
            location TEXT,
            activity_date TEXT,
            month INTEGER,
            year INTEGER,
            description TEXT,
            status TEXT DEFAULT 'Lengkap',
            created_at TEXT NOT NULL
        )
    """)

    conn.execute("""
        CREATE TABLE IF NOT EXISTS trash (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            original_id INTEGER,
            activity_id TEXT,
            file_name TEXT,
            file_hash TEXT,
            file_path TEXT,
            document_number TEXT,
            document_date TEXT,
            document_type TEXT,
            activity_name TEXT,
            category TEXT,
            subcategory TEXT,
            location TEXT,
            activity_date TEXT,
            month INTEGER,
            year INTEGER,
            description TEXT,
            status TEXT,
            deleted_at TEXT NOT NULL
        )
    """)

    conn.commit()
    conn.close()


init_db()


# ============================================================
# HELPERS
# ============================================================
def sha256_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def save_uploaded_file(uploaded_file):
    import os
    os.makedirs(UPLOAD_DIR, exist_ok=True)
    file_hash = sha256_bytes(uploaded_file.getvalue())
    safe_name = re.sub(r"[^A-Za-z0-9._-]+", "_", uploaded_file.name)
    path = os.path.join(UPLOAD_DIR, f"{file_hash[:12]}_{safe_name}")
    with open(path, "wb") as f:
        f.write(uploaded_file.getvalue())
    return path, file_hash


def extract_text(uploaded_file):
    name = uploaded_file.name.lower()
    data = uploaded_file.getvalue()

    try:
        if name.endswith(".pdf"):
            if PdfReader is None:
                return ""
            reader = PdfReader(io.BytesIO(data))
            return "\n".join((page.extract_text() or "") for page in reader.pages)

        if name.endswith(".docx"):
            if Document is None:
                return ""
            doc = Document(io.BytesIO(data))
            parts = [p.text for p in doc.paragraphs]
            for table in doc.tables:
                for row in table.rows:
                    parts.append(" | ".join(cell.text for cell in row.cells))
            return "\n".join(parts)

        if name.endswith(".txt"):
            return data.decode("utf-8", errors="ignore")

    except Exception:
        return ""

    return ""


def normalize_text(text):
    return re.sub(r"\s+", " ", text or "").strip()


def find_first(patterns, text):
    for pattern in patterns:
        m = re.search(pattern, text, flags=re.I | re.M)
        if m:
            return normalize_text(m.group(1))
    return ""


def detect_doc_type(text, filename):
    s = f"{filename}\n{text}".lower()

    # Check distinctive labels first.
    if re.search(r"\binfosus\b|informasi\s+khusus", s):
        return "Informasi Khusus (Infosus)"
    if re.search(r"\blaporan\s+informasi\b|\bnomor\s*[:.]?\s*li[/\-]", s):
        return "Laporan Informasi (LI)"
    if re.search(r"\bsurat\s*[- ]?\s*izin\b|\bnomor\s*[:.]?\s*si[/\-]", s):
        return "Surat Izin (SI)"
    if re.search(r"\bkirkat\b|\bkirka\b", s):
        return "Kirka/Kirkat"
    if re.search(r"\bsurat\s+keterangan\s+kepolisian\b|\bskp\b", s):
        return "Surat Keterangan Kepolisian (SKP)"

    return "Laporan Informasi (LI)" if "informasi" in s else DOC_TYPES[0]


def detect_category(text):
    s = (text or "").lower()

    if any(k in s for k in [
        "konser", "festival", "pentas seni", "hiburan", "entertainment",
        "musik", "artis", "panggung"
    ]):
        return "Konser/Hiburan"

    if any(k in s for k in [
        "turnamen", "kompetisi", "pertandingan", "kejuaraan", "perlombaan",
        "liga", "futsal", "sepak bola", "basket", "voli", "badminton"
    ]):
        return "Kompetisi/Pertandingan"

    if any(k in s for k in [
        "unjuk rasa", "aksi", "pengajian", "seminar", "rapat", "sosialisasi",
        "organisasi", "masyarakat", "keagamaan", "ormas", "bazar"
    ]):
        return "Kegiatan Masyarakat"

    if any(k in s for k in [
        "pemerintah daerah", "pemda", "bupati", "wakil bupati", "upacara",
        "peresmian", "kunjungan pejabat", "kedinasan", "dinas"
    ]):
        return "Kegiatan Pemerintahan"

    return "Lainnya"


def detect_subcategory(category, text):
    s = (text or "").lower()
    mapping = {
        "Konser/Hiburan": [
            ("Konser musik", ["konser", "musik"]),
            ("Pentas seni", ["pentas seni"]),
            ("Festival", ["festival"]),
            ("Hiburan masyarakat", ["hiburan"]),
        ],
        "Kompetisi/Pertandingan": [
            ("Kompetisi olahraga", ["kompetisi", "olahraga"]),
            ("Turnamen", ["turnamen"]),
            ("Perlombaan", ["perlombaan", "lomba"]),
            ("Pertandingan", ["pertandingan"]),
            ("Kejuaraan", ["kejuaraan"]),
        ],
        "Kegiatan Masyarakat": [
            ("Aksi/unjuk rasa", ["unjuk rasa", "aksi"]),
            ("Pengajian/keagamaan", ["pengajian", "keagamaan"]),
            ("Seminar", ["seminar"]),
            ("Rapat", ["rapat"]),
            ("Sosialisasi", ["sosialisasi"]),
            ("Kegiatan organisasi", ["organisasi", "ormas"]),
        ],
        "Kegiatan Pemerintahan": [
            ("Kegiatan pemerintah daerah", ["pemerintah daerah", "pemda"]),
            ("Kunjungan pejabat", ["kunjungan pejabat"]),
            ("Upacara", ["upacara"]),
            ("Peresmian", ["peresmian"]),
            ("Kegiatan kedinasan", ["kedinasan", "dinas"]),
        ],
    }

    for label, keywords in mapping.get(category, []):
        if any(k in s for k in keywords):
            return label

    return SUBCATEGORIES[category][0] if SUBCATEGORIES.get(category) else "Lainnya"


def detect_date(text):
    # dd-mm-yyyy / dd/mm/yyyy / dd.mm.yyyy
    m = re.search(r"\b(0?[1-9]|[12]\d|3[01])[-/.](0?[1-9]|1[0-2])[-/.](20\d{2})\b", text or "")
    if m:
        d, mo, y = map(int, m.groups())
        try:
            return date(y, mo, d).isoformat()
        except ValueError:
            pass

    # Indonesian month names
    month_map = {
        "januari": 1, "februari": 2, "maret": 3, "april": 4,
        "mei": 5, "juni": 6, "juli": 7, "agustus": 8,
        "september": 9, "oktober": 10, "november": 11, "desember": 12
    }
    pattern = r"\b(\d{1,2})\s+(januari|februari|maret|april|mei|juni|juli|agustus|september|oktober|november|desember)\s+(20\d{2})\b"
    m = re.search(pattern, (text or "").lower())
    if m:
        d = int(m.group(1))
        mo = month_map[m.group(2)]
        y = int(m.group(3))
        try:
            return date(y, mo, d).isoformat()
        except ValueError:
            pass

    return ""


def detect_document_number(text):
    return find_first([
        r"(?:nomor|no\.?)\s*[:.]?\s*([A-Z]{0,8}\s*/?\s*[\w.-]+(?:/[\w.-]+){1,8})",
        r"\b((?:SI|LI|R/LI|R/ND|SI/|INFOSUS)[\w./-]{3,})\b",
    ], text)


def detect_activity_name(text):
    patterns = [
        r"(?:nama kegiatan|kegiatan|perihal|tentang)\s*[:\-]\s*([^\n]{5,150})",
        r"(?:acara|agenda)\s*[:\-]\s*([^\n]{5,150})",
    ]
    value = find_first(patterns, text)
    if value:
        return value[:150]
    return ""


def detect_location(text):
    return find_first([
        r"(?:tempat|lokasi)\s*[:\-]\s*([^\n]{3,150})",
        r"(?:bertempat di|berlokasi di)\s+([^\n.]{3,150})",
    ], text)[:150]


def get_status(row):
    required = [
        row.get("document_type"),
        row.get("category"),
        row.get("activity_name"),
        row.get("activity_date"),
    ]
    return "Lengkap" if all(str(x).strip() for x in required) else "Perlu Pemeriksaan"


def read_db_df():
    conn = get_conn()
    df = pd.read_sql_query("SELECT * FROM documents ORDER BY activity_date DESC, id DESC", conn)
    conn.close()
    return df


def filtered_df(df, year_filter, month_filter, type_filter, category_filter, search):
    out = df.copy()

    if out.empty:
        return out

    if year_filter != "Semua":
        out = out[out["year"] == int(year_filter)]

    if month_filter != "Semua":
        out = out[out["month"] == MONTHS.index(month_filter) + 1]

    if type_filter != "Semua":
        out = out[out["document_type"] == type_filter]

    if category_filter != "Semua":
        out = out[out["category"] == category_filter]

    if search:
        mask = pd.Series(False, index=out.index)
        for col in ["file_name", "document_number", "activity_name", "location", "description"]:
            if col in out.columns:
                mask = mask | out[col].fillna("").astype(str).str.contains(
                    search, case=False, na=False, regex=False
                )
        out = out[mask]

    return out


def move_to_trash(doc_id):
    conn = get_conn()
    row = conn.execute("SELECT * FROM documents WHERE id=?", (doc_id,)).fetchone()

    if not row:
        conn.close()
        return False

    conn.execute("""
        INSERT INTO trash (
            original_id, activity_id, file_name, file_hash, file_path,
            document_number, document_date, document_type, activity_name,
            category, subcategory, location, activity_date, month, year,
            description, status, deleted_at
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (
        row["id"], row["activity_id"], row["file_name"], row["file_hash"],
        row["file_path"], row["document_number"], row["document_date"],
        row["document_type"], row["activity_name"], row["category"],
        row["subcategory"], row["location"], row["activity_date"],
        row["month"], row["year"], row["description"], row["status"],
        datetime.now().isoformat(timespec="seconds")
    ))

    conn.execute("DELETE FROM documents WHERE id=?", (doc_id,))
    conn.commit()
    conn.close()
    return True


def permanently_delete_trash(trash_id):
    conn = get_conn()
    row = conn.execute("SELECT file_path FROM trash WHERE id=?", (trash_id,)).fetchone()
    if row and row["file_path"]:
        try:
            import os
            if os.path.exists(row["file_path"]):
                os.remove(row["file_path"])
        except Exception:
            pass
    conn.execute("DELETE FROM trash WHERE id=?", (trash_id,))
    conn.commit()
    conn.close()


def restore_trash(trash_id):
    conn = get_conn()
    row = conn.execute("SELECT * FROM trash WHERE id=?", (trash_id,)).fetchone()
    if not row:
        conn.close()
        return False

    exists = conn.execute(
        "SELECT id FROM documents WHERE file_hash=?",
        (row["file_hash"],)
    ).fetchone()

    if exists:
        conn.close()
        return False

    conn.execute("""
        INSERT INTO documents (
            activity_id, file_name, file_hash, file_path,
            document_number, document_date, document_type, activity_name,
            category, subcategory, location, activity_date, month, year,
            description, status, created_at
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (
        row["activity_id"], row["file_name"], row["file_hash"], row["file_path"],
        row["document_number"], row["document_date"], row["document_type"],
        row["activity_name"], row["category"], row["subcategory"],
        row["location"], row["activity_date"], row["month"], row["year"],
        row["description"], row["status"],
        datetime.now().isoformat(timespec="seconds")
    ))

    conn.execute("DELETE FROM trash WHERE id=?", (trash_id,))
    conn.commit()
    conn.close()
    return True


def insert_document(item):
    conn = get_conn()
    try:
        conn.execute("""
            INSERT INTO documents (
                activity_id, file_name, file_hash, file_path,
                document_number, document_date, document_type, activity_name,
                category, subcategory, location, activity_date, month, year,
                description, status, created_at
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            item["activity_id"], item["file_name"], item["file_hash"],
            item["file_path"], item["document_number"], item["document_date"],
            item["document_type"], item["activity_name"], item["category"],
            item["subcategory"], item["location"], item["activity_date"],
            item["month"], item["year"], item["description"], item["status"],
            datetime.now().isoformat(timespec="seconds")
        ))
        conn.commit()
        return True, ""
    except sqlite3.IntegrityError:
        return False, "Duplikat berdasarkan file/hash."
    finally:
        conn.close()


# ============================================================
# CSS
# ============================================================
st.markdown("""
<style>
.block-container { padding-top: 1.2rem; }
h1, h2, h3 { letter-spacing: -0.02em; }
.metric-card {
    padding: 16px;
    border-radius: 12px;
    background: rgba(128,128,128,.08);
    border: 1px solid rgba(128,128,128,.16);
}
.small-muted { color: #888; font-size: .85rem; }
</style>
""", unsafe_allow_html=True)


# ============================================================
# HEADER
# ============================================================
st.title("📊 Dasbor Rekapitulasi Kegiatan Intelkam")
st.caption("Monitoring dokumen, kegiatan, kategori, dan perkembangan bulanan.")

df_all = read_db_df()

# ============================================================
# SIDEBAR / FILTER
# ============================================================
with st.sidebar:
    st.header("🔎 Filter Data")

    years = sorted(
        [int(x) for x in df_all["year"].dropna().unique()],
        reverse=True
    ) if not df_all.empty else [datetime.now().year]

    year_filter = st.selectbox(
        "Tahun",
        ["Semua"] + [str(y) for y in years],
        index=1 if str(datetime.now().year) in [str(y) for y in years] else 0
    )

    month_filter = st.selectbox("Bulan", ["Semua"] + MONTHS)
    type_filter = st.selectbox("Jenis Dokumen", ["Semua"] + DOC_TYPES)
    category_filter = st.selectbox("Kategori Kegiatan", ["Semua"] + ACTIVITY_CATEGORIES)
    search = st.text_input("🔍 Pencarian", placeholder="Nomor, kegiatan, lokasi, file...")

    st.divider()
    st.caption("Data tersimpan pada database SQLite lokal: intelkam.db")

df = filtered_df(
    df_all, year_filter, month_filter,
    type_filter, category_filter, search
)

# ============================================================
# METRICS
# ============================================================
total_docs = len(df)
total_activities = df["activity_id"].nunique() if not df.empty else 0

type_counts = df["document_type"].value_counts() if not df.empty else pd.Series(dtype=int)
category_counts = df["category"].value_counts() if not df.empty else pd.Series(dtype=int)

most_category = category_counts.index[0] if len(category_counts) else "-"
most_category_count = int(category_counts.iloc[0]) if len(category_counts) else 0

# activity-level count is unique activity_id
st.markdown("### Ringkasan")

m1, m2, m3, m4, m5, m6 = st.columns(6)
m1.metric("Total Dokumen", total_docs)
m2.metric("Total Kegiatan", total_activities)
m3.metric("Surat Izin", int(type_counts.get(DOC_TYPES[0], 0)))
m4.metric("Laporan Informasi", int(type_counts.get(DOC_TYPES[1], 0)))
m5.metric("Infosus", int(type_counts.get(DOC_TYPES[2], 0)))
m6.metric("Kirkat / SKP", int(type_counts.get(DOC_TYPES[3], 0)) + int(type_counts.get(DOC_TYPES[4], 0)))

st.info(
    f"Kategori terbanyak: **{most_category}** ({most_category_count} dokumen/kegiatan terdata)."
    if most_category != "-" else
    "Belum ada data pada filter yang dipilih."
)

# ============================================================
# UPLOAD MASSAL
# ============================================================
st.divider()
st.subheader("📁 Import Dokumen Massal")

st.write(
    "Upload banyak PDF/DOCX sekaligus. Sistem akan membaca isi dokumen, "
    "mendeteksi jenis/kategori, memeriksa duplikasi, lalu menampilkan preview sebelum disimpan."
)

uploaded_files = st.file_uploader(
    "Tarik dan lepaskan file di sini",
    type=["pdf", "docx", "txt"],
    accept_multiple_files=True,
    help="Pilih banyak file sekaligus dari komputer."
)

if uploaded_files:
    if "import_items" not in st.session_state:
        st.session_state.import_items = {}

    for uf in uploaded_files:
        if uf.name not in st.session_state.import_items:
            text_content = extract_text(uf)
            text_content = normalize_text(text_content)

            doc_type = detect_doc_type(text_content, uf.name)
            category = detect_category(text_content)
            subcategory = detect_subcategory(category, text_content)
            activity_date = detect_date(text_content)

            if activity_date:
                dt = datetime.strptime(activity_date, "%Y-%m-%d")
                month = dt.month
                year = dt.year
            else:
                month = None
                year = None

            item = {
                "selected": True,
                "file_name": uf.name,
                "file_hash": sha256_bytes(uf.getvalue()),
                "file_path": "",
                "document_number": detect_document_number(text_content),
                "document_date": activity_date,
                "document_type": doc_type,
                "activity_name": detect_activity_name(text_content),
                "category": category,
                "subcategory": subcategory,
                "location": detect_location(text_content),
                "activity_date": activity_date,
                "month": month,
                "year": year,
                "description": "",
                "status": "Lengkap",
                "activity_id": "ACT-" + sha256_bytes(
                    (detect_activity_name(text_content) + "|" + (activity_date or "")).encode()
                )[:12],
            }

            # Check duplicate
            conn = get_conn()
            duplicate = conn.execute(
                "SELECT id, file_name FROM documents WHERE file_hash=?",
                (item["file_hash"],)
            ).fetchone()
            conn.close()

            item["duplicate"] = bool(duplicate)
            if duplicate:
                item["status"] = "Duplikat"

            if not item["activity_name"] or not item["activity_date"]:
                item["status"] = "Perlu Pemeriksaan"

            st.session_state.import_items[uf.name] = item

    items = list(st.session_state.import_items.values())

    st.write(f"**{len(items)} file** siap diperiksa.")

    # Summary
    ready = sum(1 for x in items if x["status"] == "Lengkap" and not x["duplicate"])
    review = sum(1 for x in items if x["status"] == "Perlu Pemeriksaan")
    duplicates = sum(1 for x in items if x["duplicate"])

    a, b, c = st.columns(3)
    a.metric("🟢 Siap Diimpor", ready)
    b.metric("🟡 Perlu Pemeriksaan", review)
    c.metric("🟠 Duplikat", duplicates)

    st.markdown("#### Preview Import")

    edited_items = []

    for idx, item in enumerate(items):
        with st.expander(
            f"{'🟠' if item['duplicate'] else '🟡' if item['status'] != 'Lengkap' else '🟢'} "
            f"{item['file_name']}",
            expanded=(item["status"] != "Lengkap" or item["duplicate"])
        ):
            c1, c2 = st.columns([1, 1])

            with c1:
                selected = st.checkbox(
                    "Import file ini",
                    value=item["selected"] and not item["duplicate"],
                    key=f"sel_{item['file_hash']}"
                )
                item["selected"] = selected

                item["document_type"] = st.selectbox(
                    "Jenis Dokumen",
                    DOC_TYPES,
                    index=DOC_TYPES.index(item["document_type"]) if item["document_type"] in DOC_TYPES else 0,
                    key=f"type_{item['file_hash']}"
                )

                item["document_number"] = st.text_input(
                    "Nomor Dokumen",
                    item["document_number"],
                    key=f"num_{item['file_hash']}"
                )

                item["activity_name"] = st.text_input(
                    "Nama/Judul Kegiatan",
                    item["activity_name"],
                    key=f"name_{item['file_hash']}"
                )

            with c2:
                item["category"] = st.selectbox(
                    "Kategori Kegiatan",
                    ACTIVITY_CATEGORIES,
                    index=ACTIVITY_CATEGORIES.index(item["category"]) if item["category"] in ACTIVITY_CATEGORIES else 4,
                    key=f"cat_{item['file_hash']}"
                )

                subs = SUBCATEGORIES[item["category"]]
                current_sub = item["subcategory"] if item["subcategory"] in subs else subs[0]
                item["subcategory"] = st.selectbox(
                    "Subkategori",
                    subs,
                    index=subs.index(current_sub),
                    key=f"sub_{item['file_hash']}"
                )

                parsed_date = None
                if item["activity_date"]:
                    try:
                        parsed_date = datetime.strptime(item["activity_date"], "%Y-%m-%d").date()
                    except Exception:
                        parsed_date = None

                chosen_date = st.date_input(
                    "Tanggal Kegiatan",
                    value=parsed_date,
                    key=f"date_{item['file_hash']}"
                )

                if chosen_date:
                    item["activity_date"] = chosen_date.isoformat()
                    item["month"] = chosen_date.month
                    item["year"] = chosen_date.year
                else:
                    item["activity_date"] = ""
                    item["month"] = None
                    item["year"] = None

                item["location"] = st.text_input(
                    "Lokasi",
                    item["location"],
                    key=f"loc_{item['file_hash']}"
                )

            item["description"] = st.text_area(
                "Keterangan",
                item["description"],
                key=f"desc_{item['file_hash']}"
            )

            # Recalculate status after edits
            item["status"] = get_status(item)
            st.caption(f"Status: **{item['status']}**")

            edited_items.append(item)

    st.session_state.import_items = {
        x["file_name"]: x for x in edited_items
    }

    col_a, col_b = st.columns([1, 1])

    with col_a:
        if st.button("🚀 Import Data Terpilih", type="primary", use_container_width=True):
            imported = 0
            skipped = 0
            failed = 0

            # Need uploaded file lookup for saving bytes
            uf_map = {u.name: u for u in uploaded_files}

            for item in edited_items:
                if not item["selected"]:
                    continue

                if item["duplicate"]:
                    skipped += 1
                    continue

                if item["status"] != "Lengkap":
                    skipped += 1
                    continue

                uf = uf_map.get(item["file_name"])
                if not uf:
                    failed += 1
                    continue

                path, file_hash = save_uploaded_file(uf)
                item["file_path"] = path
                item["file_hash"] = file_hash

                ok, _ = insert_document(item)
                if ok:
                    imported += 1
                else:
                    skipped += 1

            st.success(
                f"Import selesai: **{imported}** berhasil, "
                f"**{skipped}** dilewati, **{failed}** gagal."
            )
            st.session_state.import_items = {}
            st.rerun()

    with col_b:
        if st.button("🧹 Bersihkan Daftar Import", use_container_width=True):
            st.session_state.import_items = {}
            st.rerun()


# ============================================================
# CHARTS
# ============================================================
st.divider()
st.subheader("📈 Analisis")

left, right = st.columns(2)

with left:
    st.markdown("#### Jumlah Dokumen Berdasarkan Jenis")
    chart_types = pd.DataFrame({
        "Jenis Dokumen": DOC_TYPES,
        "Jumlah": [int(type_counts.get(x, 0)) for x in DOC_TYPES]
    }).set_index("Jenis Dokumen")
    st.bar_chart(chart_types, height=320)

with right:
    st.markdown("#### Sebaran Kategori Kegiatan")
    chart_cat = pd.DataFrame({
        "Kategori": ACTIVITY_CATEGORIES,
        "Jumlah": [int(category_counts.get(x, 0)) for x in ACTIVITY_CATEGORIES]
    }).set_index("Kategori")
    st.bar_chart(chart_cat, height=320)

st.markdown("#### 📅 Jumlah Kegiatan Per Bulan")

if not df.empty:
    monthly = (
        df.groupby("month")["activity_id"]
        .nunique()
        .reindex(range(1, 13), fill_value=0)
    )
else:
    monthly = pd.Series(0, index=range(1, 13))

monthly_df = pd.DataFrame({
    "Bulan": MONTHS,
    "Jumlah Kegiatan": monthly.values
}).set_index("Bulan")

st.line_chart(monthly_df, height=320)

# ============================================================
# CROSS ANALYSIS
# ============================================================
st.subheader("🔎 Analisis Silang: Jenis Dokumen × Kategori Kegiatan")

if not df.empty:
    cross = pd.crosstab(
        df["category"],
        df["document_type"]
    ).reindex(index=ACTIVITY_CATEGORIES, columns=DOC_TYPES, fill_value=0)

    cross["Total"] = cross.sum(axis=1)
    cross.loc["Total"] = cross.sum(axis=0)
else:
    cross = pd.DataFrame(
        0,
        index=ACTIVITY_CATEGORIES + ["Total"],
        columns=DOC_TYPES + ["Total"]
    )

st.dataframe(cross, use_container_width=True)

# ============================================================
# AUTO ANALYSIS
# ============================================================
st.subheader("📝 Analisis Data")

if total_docs:
    period = "data terpilih"
    if year_filter != "Semua" and month_filter != "Semua":
        period = f"{month_filter} {year_filter}"
    elif year_filter != "Semua":
        period = f"tahun {year_filter}"

    top_type = type_counts.index[0] if len(type_counts) else "-"
    top_type_count = int(type_counts.iloc[0]) if len(type_counts) else 0

    st.write(
        f"Pada {period} terdapat **{total_docs} dokumen** yang mencakup "
        f"**{total_activities} kegiatan unik**. "
        f"Jenis dokumen dengan jumlah terbanyak adalah **{top_type}** "
        f"sebanyak **{top_type_count} dokumen**. "
        f"Kategori kegiatan dengan jumlah terbanyak adalah **{most_category}** "
        f"sebanyak **{most_category_count} data**."
    )
else:
    st.info("Belum terdapat data yang sesuai dengan filter.")

# ============================================================
# DOCUMENT TABLE
# ============================================================
st.divider()
st.subheader("📋 Daftar Rekapitulasi Dokumen")

if df.empty:
    st.info("Belum ada dokumen pada filter yang dipilih.")
else:
    display_cols = [
        "id", "document_number", "activity_date", "document_type",
        "activity_name", "category", "subcategory", "location",
        "month", "year", "status", "file_name"
    ]

    display = df[display_cols].copy()
    display.columns = [
        "ID", "Nomor Dokumen", "Tanggal Kegiatan", "Jenis",
        "Nama Kegiatan", "Kategori", "Subkategori", "Lokasi",
        "Bulan", "Tahun", "Status", "File"
    ]

    st.dataframe(
        display,
        use_container_width=True,
        hide_index=True,
        column_config={
            "ID": st.column_config.NumberColumn(width="small"),
            "Tanggal Kegiatan": st.column_config.DateColumn(format="DD-MM-YYYY"),
            "File": st.column_config.TextColumn(width="medium"),
        }
    )

    st.markdown("#### Aksi Dokumen")

    ids = df["id"].tolist()
    selected_id = st.selectbox(
        "Pilih dokumen",
        ids,
        format_func=lambda x: (
            f"#{x} — "
            f"{df.loc[df['id'] == x, 'activity_name'].iloc[0] or df.loc[df['id'] == x, 'file_name'].iloc[0]}"
        )
    )

    ac1, ac2, ac3 = st.columns(3)

    with ac1:
        if st.button("👁️ Lihat Detail", use_container_width=True):
            row = df[df["id"] == selected_id].iloc[0]
            st.session_state["detail_id"] = int(selected_id)

    with ac2:
        if st.button("🗑️ Hapus ke Tempat Sampah", use_container_width=True):
            if move_to_trash(int(selected_id)):
                st.success("Dokumen dipindahkan ke Tempat Sampah.")
                st.rerun()

    with ac3:
        export_df = df.copy()
        csv_data = export_df.to_csv(index=False).encode("utf-8-sig")
        st.download_button(
            "📥 Export CSV",
            data=csv_data,
            file_name="rekap_intelkam.csv",
            mime="text/csv",
            use_container_width=True
        )

# ============================================================
# DETAIL
# ============================================================
if st.session_state.get("detail_id"):
    detail_id = st.session_state["detail_id"]
    detail_df = df_all[df_all["id"] == detail_id]

    if not detail_df.empty:
        row = detail_df.iloc[0]

        st.divider()
        st.subheader("📄 Detail Dokumen")

        d1, d2 = st.columns(2)

        with d1:
            st.write(f"**Nomor Dokumen:** {row['document_number'] or '-'}")
            st.write(f"**Jenis Dokumen:** {row['document_type']}")
            st.write(f"**Nama Kegiatan:** {row['activity_name'] or '-'}")
            st.write(f"**Kategori:** {row['category']}")
            st.write(f"**Subkategori:** {row['subcategory'] or '-'}")

        with d2:
            st.write(f"**Tanggal Kegiatan:** {row['activity_date'] or '-'}")
            st.write(f"**Lokasi:** {row['location'] or '-'}")
            st.write(f"**Status:** {row['status']}")
            st.write(f"**File:** {row['file_name']}")

        if row["file_path"]:
            try:
                with open(row["file_path"], "rb") as f:
                    file_bytes = f.read()
                st.download_button(
                    "📥 Buka / Download File Asli",
                    data=file_bytes,
                    file_name=row["file_name"],
                    mime="application/octet-stream"
                )
            except Exception:
                st.warning("File asli tidak ditemukan pada penyimpanan lokal.")

        if st.button("Tutup Detail"):
            del st.session_state["detail_id"]
            st.rerun()


# ============================================================
# TRASH
# ============================================================
st.divider()
st.subheader("🗑️ Tempat Sampah")

conn = get_conn()
trash_df = pd.read_sql_query(
    "SELECT * FROM trash ORDER BY deleted_at DESC",
    conn
)
conn.close()

if trash_df.empty:
    st.caption("Tempat Sampah kosong.")
else:
    st.write(f"**{len(trash_df)} dokumen** berada di Tempat Sampah.")

    st.dataframe(
        trash_df[[
            "id", "file_name", "document_number", "document_type",
            "activity_name", "deleted_at"
        ]].rename(columns={
            "id": "ID",
            "file_name": "File",
            "document_number": "Nomor Dokumen",
            "document_type": "Jenis",
            "activity_name": "Kegiatan",
            "deleted_at": "Dihapus"
        }),
        use_container_width=True,
        hide_index=True
    )

    trash_id = st.selectbox(
        "Pilih item di Tempat Sampah",
        trash_df["id"].tolist(),
        format_func=lambda x: trash_df.loc[
            trash_df["id"] == x, "file_name"
        ].iloc[0]
    )

    t1, t2 = st.columns(2)

    with t1:
        if st.button("♻️ Pulihkan", use_container_width=True):
            if restore_trash(int(trash_id)):
                st.success("Dokumen berhasil dipulihkan.")
                st.rerun()
            else:
                st.error("Gagal memulihkan. Kemungkinan file/hash sudah ada.")

    with t2:
        if st.button("❌ Hapus Permanen", use_container_width=True):
            permanently_delete_trash(int(trash_id))
            st.success("Dokumen dihapus permanen.")
            st.rerun()


# ============================================================
# FOOTER
# ============================================================
st.divider()
st.caption(
    "Dasbor Rekapitulasi Kegiatan Intelkam • Data tersimpan lokal pada SQLite. "
    "Pastikan melakukan backup file intelkam.db dan folder uploads secara berkala."
)
