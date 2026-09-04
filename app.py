import streamlit as st
import pandas as pd
import plotly.express as px
import re
import io
import os
import numpy as np


# ============================================================
# KONFIGURASI HALAMAN & CSS
# ============================================================

st.set_page_config(
    page_title="Dashboard Gama Statistika",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)


st.markdown("""
<style>

    /* ========================================================
       GLOBAL
       ======================================================== */

    .stApp {
        background-color: #f5f7fb;
    }

    section[data-testid="stSidebar"] {
        background-color: #111827;
    }

    section[data-testid="stSidebar"] * {
        color: white;
    }

    /* ========================================================
       MAIN TITLE
       ======================================================== */

    .main-title {
        font-size: 34px;
        font-weight: 800;
        color: #111827;
        margin-top: 5px;
        margin-bottom: 2px;
        letter-spacing: -0.5px;
    }

    .subtitle {
        font-size: 14px;
        color: #6b7280;
        margin-bottom: 18px;
    }

    /* ========================================================
       INFO BOX
       ======================================================== */

    .info-box {
        background: white;
        padding: 14px 18px;
        border-radius: 12px;
        border: 1px solid #e5e7eb;
        margin-bottom: 16px;
        color: #374151;
        font-size: 13px;
        line-height: 1.8;
        box-shadow: 0 2px 8px rgba(0,0,0,0.03);
    }

    /* ========================================================
       KPI CARD
       ======================================================== */

    .kpi-card {
        background: white;
        padding: 15px 16px;
        border-radius: 14px;
        border: 1px solid #e5e7eb;
        box-shadow: 0 3px 10px rgba(0,0,0,0.04);
        min-height: 92px;
        margin-bottom: 12px;
        transition: all 0.2s ease;
    }

    .kpi-card:hover {
        box-shadow: 0 5px 15px rgba(0,0,0,0.07);
        transform: translateY(-1px);
    }

    .kpi-title {
        color: #6b7280;
        font-size: 10px;
        font-weight: 700;
        margin-bottom: 8px;
        text-transform: uppercase;
        letter-spacing: 0.6px;
    }

    .kpi-value {
        color: #111827;
        font-size: 21px;
        font-weight: 800;
        line-height: 1.2;
    }

    .kpi-value-danger {
        color: #dc2626;
        font-size: 21px;
        font-weight: 800;
        line-height: 1.2;
    }

    .kpi-value-success {
        color: #16a34a;
        font-size: 21px;
        font-weight: 800;
        line-height: 1.2;
    }

    /* ========================================================
       SECTION TITLE
       ======================================================== */

    .section-title {
        font-size: 20px;
        font-weight: 800;
        color: #111827;
        margin-top: 22px;
        margin-bottom: 10px;
        padding-bottom: 5px;
    }

    /* ========================================================
       TABLE
       ======================================================== */

    div[data-testid="stDataFrame"] {
        border-radius: 10px;
        overflow: hidden;
        border: 1px solid #e5e7eb;
    }

    /* ========================================================
       SIDEBAR
       ======================================================== */

    section[data-testid="stSidebar"] {
        border-right: 1px solid #1f2937;
    }

    /* ========================================================
       BUTTON
       ======================================================== */

    .stDownloadButton button {
        border-radius: 9px;
        font-weight: 600;
    }

    /* ========================================================
       STREAMLIT ELEMENT SPACING
       ======================================================== */

    div.block-container {
        padding-top: 2rem;
        padding-bottom: 3rem;
    }

</style>
""", unsafe_allow_html=True)


# ============================================================
# LOAD FILE PATOKAN CLUSTER
# ============================================================

FILE_PATOKAN = "data_final_gama_streamlit.xlsx"


# ============================================================
# BATAS RANGE NILAI PROJECT
# ============================================================

BATAS_RENDAH = 1500000
BATAS_TINGGI = 2800000


# ============================================================
# NAMA BULAN
# ============================================================

nama_bulan_map = {
    1: "Januari",
    2: "Februari",
    3: "Maret",
    4: "April",
    5: "Mei",
    6: "Juni",
    7: "Juli",
    8: "Agustus",
    9: "September",
    10: "Oktober",
    11: "November",
    12: "Desember"
}


# ============================================================
# FUNGSI NORMALISASI TEKS
# ============================================================

def normalisasi_teks(teks):

    if pd.isna(teks):
        return ""

    return (
        str(teks)
        .strip()
        .lower()
        .replace("\n", " ")
        .replace("\r", " ")
        .replace("\t", " ")
    )


# ============================================================
# FUNGSI CARI KOLOM
# ============================================================

def cari_kolom(df, daftar_kata_kunci):

    cols_lower = {
        normalisasi_teks(c): c
        for c in df.columns
    }

    # Prioritas exact match
    for kata in daftar_kata_kunci:

        kata_lower = normalisasi_teks(kata)

        if kata_lower in cols_lower:
            return cols_lower[kata_lower]

    # Kemudian partial match
    for kata in daftar_kata_kunci:

        kata_lower = normalisasi_teks(kata)

        kata_bersih = (
            kata_lower
            .replace(" ", "")
            .replace("_", "")
            .replace("-", "")
            .replace("/", "")
            .replace(".", "")
        )

        if not kata_bersih:
            continue

        for col_key, col_orig in cols_lower.items():

            col_key_bersih = (
                col_key
                .replace(" ", "")
                .replace("_", "")
                .replace("-", "")
                .replace("/", "")
                .replace(".", "")
            )

            if kata_bersih in col_key_bersih:
                return col_orig

    return None


# ============================================================
# FUNGSI CARI KOLOM BERDASARKAN ISI
# ============================================================

def cari_kolom_kode_dari_isi(df):

    skor_kolom = {}

    pola_p = re.compile(
        r"\bp[\.\-\s]*(\d+)\b",
        re.IGNORECASE
    )

    for col in df.columns:

        try:

            sample = (
                df[col]
                .dropna()
                .astype(str)
                .head(5000)
            )

            if len(sample) == 0:
                continue

            jumlah_cocok = sample.apply(
                lambda x: bool(
                    pola_p.search(x.strip())
                )
            ).sum()

            if jumlah_cocok > 0:
                skor_kolom[col] = jumlah_cocok

        except Exception:
            continue

    if not skor_kolom:
        return None

    return max(
        skor_kolom,
        key=skor_kolom.get
    )


# ============================================================
# FUNGSI MENDETEKSI BARIS HEADER
# ============================================================

def cari_baris_header(df_raw):

    kata_kunci_header = [
        "nama",
        "klien",
        "pelanggan",
        "kode",
        "code",
        "id",
        "biaya",
        "omset",
        "bayar",
        "piutang",
        "tanggal",
        "tgl",
        "admin",
        "konsultan",
        "software",
        "metode",
        "kebutuhan",
        "fee",
        "project",
        "proyek",
        "pendapatan"
    ]

    skor_baris = []

    batas_scan = min(
        len(df_raw),
        30
    )

    for idx in range(batas_scan):

        nilai_baris = [
            normalisasi_teks(x)
            for x in df_raw.iloc[idx].tolist()
        ]

        skor = 0

        for nilai in nilai_baris:

            if not nilai:
                continue

            for kata in kata_kunci_header:

                if kata in nilai:
                    skor += 1
                    break

        jumlah_isi = sum(
            1 for x in nilai_baris
            if x != ""
        )

        if jumlah_isi >= 3:
            skor += 1

        skor_baris.append(skor)

    if not skor_baris:
        return 0

    skor_maksimum = max(
        skor_baris
    )

    if skor_maksimum == 0:
        return 0

    return skor_baris.index(
        skor_maksimum
    )


# ============================================================
# FUNGSI MEMBERSIHKAN NAMA KOLOM
# ============================================================

def bersihkan_nama_kolom(columns):

    hasil = []
    jumlah_nama = {}

    for i, col in enumerate(columns):

        nama = str(col).strip()

        if (
            not nama
            or nama.lower() == "nan"
            or nama.lower().startswith("unnamed")
        ):

            nama = f"Kolom_{i+1}"

        if nama in jumlah_nama:

            jumlah_nama[nama] += 1
            nama = f"{nama}_{jumlah_nama[nama]}"

        else:

            jumlah_nama[nama] = 1

        hasil.append(nama)

    return hasil


# ============================================================
# FUNGSI MEMBERSIHKAN BARIS DATA
# ============================================================

def bersihkan_baris_data(df):

    df = df.copy()

    df = df.dropna(
        how="all"
    )

    mask_kosong = (
        df.astype(str)
        .apply(
            lambda row:
            row.str.strip().eq("").all(),
            axis=1
        )
    )

    df = df[
        ~mask_kosong
    ]

    return df.reset_index(
        drop=True
    )


# ============================================================
# FUNGSI MEMBACA EXCEL
# ============================================================

def daftar_sheet_excel(uploaded_file):

    try:

        excel_file = pd.ExcelFile(
            uploaded_file
        )

        return excel_file.sheet_names

    except Exception:

        return []


# ============================================================
# FUNGSI BACA DATA DARI SHEET
# ============================================================

def baca_sheet_excel(
    uploaded_file,
    sheet_name
):

    try:

        df_raw = pd.read_excel(
            uploaded_file,
            sheet_name=sheet_name,
            header=None
        )

        if df_raw.empty:
            return None

        baris_header = cari_baris_header(
            df_raw
        )

        header = (
            df_raw
            .iloc[baris_header]
            .tolist()
        )

        header = bersihkan_nama_kolom(
            header
        )

        df = df_raw.iloc[
            baris_header + 1:
        ].copy()

        df.columns = header

        df = bersihkan_baris_data(
            df
        )

        return df

    except Exception as e:

        st.error(
            f"Gagal membaca sheet: {e}"
        )

        return None


# ============================================================
# FUNGSI BACA CSV
# ============================================================

def baca_csv_otomatis(
    uploaded_file
):

    try:

        try:

            df_raw = pd.read_csv(
                uploaded_file,
                header=None,
                encoding="utf-8"
            )

        except Exception:

            uploaded_file.seek(0)

            df_raw = pd.read_csv(
                uploaded_file,
                header=None,
                encoding="latin1"
            )

        if df_raw.empty:
            return None

        baris_header = cari_baris_header(
            df_raw
        )

        header = (
            df_raw
            .iloc[baris_header]
            .tolist()
        )

        header = bersihkan_nama_kolom(
            header
        )

        df = df_raw.iloc[
            baris_header + 1:
        ].copy()

        df.columns = header

        df = bersihkan_baris_data(
            df
        )

        return df

    except Exception as e:

        st.error(
            f"Gagal membaca CSV: {e}"
        )

        return None


# ============================================================
# FUNGSI CLEAN NUMERIC
# ============================================================

def clean_numeric(series):

    def parse_val(val):

        if pd.isna(val):
            return 0.0

        if isinstance(
            val,
            (int, float, np.integer, np.floating)
        ):

            if np.isnan(val) if isinstance(
                val,
                (float, np.floating)
            ) else False:

                return 0.0

            return float(val)

        val_str = (
            str(val)
            .strip()
            .lower()
            .replace("rp", "")
            .replace("idr", "")
            .strip()
        )

        if not val_str:
            return 0.0

        val_str = re.sub(
            r"[^\d,\.\-]",
            "",
            val_str
        )

        if not val_str:
            return 0.0

        if (
            "." in val_str
            and "," in val_str
        ):

            if (
                val_str.rfind(",")
                >
                val_str.rfind(".")
            ):

                val_str = (
                    val_str
                    .replace(".", "")
                    .replace(",", ".")
                )

            else:

                val_str = (
                    val_str
                    .replace(",", "")
                )

        elif "," in val_str:

            bagian = val_str.split(",")

            if len(
                bagian[-1]
            ) <= 2:

                val_str = (
                    val_str
                    .replace(",", ".")
                )

            else:

                val_str = (
                    val_str
                    .replace(",", "")
                )

        elif "." in val_str:

            bagian = val_str.split(".")

            if len(bagian) > 1 and all(
                len(x) == 3
                for x in bagian[1:]
            ):

                val_str = (
                    val_str
                    .replace(".", "")
                )

        try:

            return float(
                val_str
            )

        except Exception:

            return 0.0

    return series.apply(
        parse_val
    )


# ============================================================
# FUNGSI FORMAT RUPIAH
# ============================================================

def rupiah(value):

    if pd.isna(value):
        return "Rp0"

    try:

        val = float(value)

        return (
            f"Rp{val:,.0f}"
            .replace(",", ".")
        )

    except (
        ValueError,
        TypeError
    ):

        return "Rp0"


# ============================================================
# FUNGSI KATEGORI KLIEN
# ============================================================

def tentukan_kategori_klien(nama):

    if pd.isna(nama):
        return "Tidak Diketahui"

    val = (
        str(nama)
        .strip()
        .lower()
    )

    if (
        not val
        or val in [
            "nan",
            "none",
            "-",
            "null",
            ""
        ]
    ):

        return "Tidak Diketahui"

    match = re.search(
        r'p[\.\-\s]*(\d+)',
        val
    )

    if match:

        nomor = int(
            match.group(1)
        )

        if nomor == 1:
            return "Klien Baru"

        if nomor > 1:
            return "Repeat Order"

    if (
        "baru" in val
        or "new" in val
    ):

        return "Klien Baru"

    if (
        "repeat" in val
        or "ro" in val
        or "lama" in val
    ):

        return "Repeat Order"

    return "Tidak Diketahui"


# ============================================================
# FUNGSI NOMOR P
# ============================================================

def ambil_nomor_p(value):

    if pd.isna(value):
        return np.nan

    match = re.search(
        r'p[\.\-\s]*(\d+)',
        str(value).lower()
    )

    if match:

        return int(
            match.group(1)
        )

    return np.nan


# ============================================================
# FUNGSI KODE DASAR KLIEN
# ============================================================

def ambil_kode_dasar(value):

    if pd.isna(value):
        return ""

    text = str(value).strip()

    match = re.search(
        r'(p[\.\-\s]*\d+)',
        text.lower()
    )

    if match:

        return (
            match.group(1)
            .upper()
            .replace(" ", "")
            .replace("-", ".")
        )

    return text


# ============================================================
# LOAD CLUSTER REFERENCE
# ============================================================

@st.cache_data
def load_cluster_reference_series():

    if not os.path.exists(
        FILE_PATOKAN
    ):

        return None

    try:

        ref_df = pd.read_excel(
            FILE_PATOKAN
        )

        ref_df.columns = [
            str(c).strip()
            for c in ref_df.columns
        ]

        col_cluster_ref = cari_kolom(
            ref_df,
            [
                "label_cluster",
                "cluster_biaya",
                "cluster"
            ]
        )

        if col_cluster_ref:

            return (
                ref_df[
                    col_cluster_ref
                ]
                .reset_index(
                    drop=True
                )
            )

    except Exception as e:

        st.warning(
            f"File patokan cluster tidak dapat dibaca: {e}"
        )

    return None


# ============================================================
# FUNGSI URUTAN CLUSTER
# ============================================================

def urutan_cluster(label):

    label_lower = (
        str(label)
        .lower()
    )

    if (
        "rendah" in label_lower
        or "low" in label_lower
    ):

        return 1

    if (
        "menengah" in label_lower
        or "sedang" in label_lower
        or "medium" in label_lower
    ):

        return 2

    if (
        "tinggi" in label_lower
        or "high" in label_lower
    ):

        return 3

    return 99


# ============================================================
# FUNGSI KLASIFIKASI NILAI PROJECT
# ============================================================

def hitung_cluster_otomatis(nilai):

    if nilai >= BATAS_TINGGI:

        return (
            "Cluster Nilai Project Tinggi"
        )

    elif nilai >= BATAS_RENDAH:

        return (
            "Cluster Nilai Project Menengah"
        )

    else:

        return (
            "Cluster Nilai Project Rendah"
        )


# ============================================================
# SIDEBAR
# ============================================================

st.sidebar.markdown(
    "## 📊 Gama Statistika"
)

st.sidebar.markdown(
    "Dashboard Analisis Transaksi"
)

st.sidebar.divider()


uploaded_file = st.sidebar.file_uploader(
    "📁 Upload File Transaksi",
    type=[
        "xlsx",
        "xls",
        "csv"
    ]
)


# ============================================================
# TAMPILAN AWAL
# ============================================================

if uploaded_file is None:

    st.markdown(
        '<div class="main-title">'
        '📊 Dashboard Gama Statistika'
        '</div>',
        unsafe_allow_html=True
    )

    st.markdown(
        '<div class="subtitle">'
        'Dashboard analisis pelanggan, transaksi, '
        'omset, project, konsultan, software, '
        'dan segmentasi cluster.'
        '</div>',
        unsafe_allow_html=True
    )

    st.info(
        "👈 Silakan upload file Excel atau CSV "
        "di sidebar sebelah kiri."
    )

    st.stop()


# ============================================================
# BACA FILE
# ============================================================

nama_file = uploaded_file.name.lower()


# ============================================================
# EXCEL
# ============================================================

if (
    nama_file.endswith(".xlsx")
    or nama_file.endswith(".xls")
):

    daftar_sheet = daftar_sheet_excel(
        uploaded_file
    )

    if not daftar_sheet:

        st.error(
            "❌ Sheet Excel tidak ditemukan."
        )

        st.stop()

    st.sidebar.markdown(
        "### 📑 Pilih Sheet"
    )

    sheet_pilihan = st.sidebar.selectbox(
        "Sheet yang digunakan:",
        options=daftar_sheet
    )

    df = baca_sheet_excel(
        uploaded_file,
        sheet_pilihan
    )


# ============================================================
# CSV
# ============================================================

else:

    sheet_pilihan = None

    df = baca_csv_otomatis(
        uploaded_file
    )


# ============================================================
# CEK DATA
# ============================================================

if df is None or df.empty:

    st.warning(
        "⚠️ Tidak ditemukan data transaksi "
        "yang dapat dibaca."
    )

    st.stop()


all_cols = df.columns.tolist()


# ============================================================
# AUTO DETECT KOLOM
# ============================================================

auto_col_nama = cari_kolom(
    df,
    [
        "nama",
        "kode klien",
        "kode",
        "code",
        "klien",
        "pelanggan",
        "customer",
        "id klien"
    ]
)


if auto_col_nama is None:

    auto_col_nama = cari_kolom_kode_dari_isi(
        df
    )


auto_col_omset = cari_kolom(
    df,
    [
        "biaya_total_klien_bayar",
        "biaya_total",
        "total_biaya",
        "total biaya",
        "biaya",
        "omset",
        "omzet",
        "pendapatan",
        "bayar",
        "jumlah bayar"
    ]
)


auto_col_piutang = cari_kolom(
    df,
    [
        "piutang",
        "hutang",
        "sisa_bayar",
        "sisa bayar",
        "kekurangan"
    ]
)


COL_TANGGAL = cari_kolom(
    df,
    [
        "tanggal_deal/dp",
        "tanggal deal/dp",
        "tanggal_deal",
        "tanggal deal",
        "tanggal_bayar",
        "tanggal bayar",
        "tanggal",
        "tgl",
        "date"
    ]
)


COL_ADMIN = cari_kolom(
    df,
    [
        "admin",
        "nama_admin",
        "nama admin"
    ]
)


COL_KONSULTAN = cari_kolom(
    df,
    [
        "konsultan",
        "nama konsultan",
        "konsultan handle",
        "handle"
    ]
)


COL_SOFTWARE = cari_kolom(
    df,
    [
        "software",
        "aplikasi",
        "program",
        "metode",
        "tools"
    ]
)


COL_KEBUTUHAN = cari_kolom(
    df,
    [
        "kebutuhan",
        "jenis kebutuhan",
        "keperluan"
    ]
)


# ============================================================
# SIDEBAR PENGATURAN KOLOM
# ============================================================

st.sidebar.markdown(
    "### ⚙️ Pengaturan Kolom"
)


# ============================================================
# KOLOM NAMA / KODE
# ============================================================

default_idx_nama = (
    all_cols.index(
        auto_col_nama
    )
    if auto_col_nama in all_cols
    else 0
)


selected_col_nama = st.sidebar.selectbox(
    "Pilih Kolom Nama/Kode Klien:",
    options=all_cols,
    index=default_idx_nama
)


# ============================================================
# KOLOM OMSET
# ============================================================

default_idx_omset = (
    all_cols.index(
        auto_col_omset
    )
    if auto_col_omset in all_cols
    else 0
)


selected_col_omset = st.sidebar.selectbox(
    "Pilih Kolom Total Omset:",
    options=all_cols,
    index=default_idx_omset
)


# ============================================================
# KOLOM PIUTANG
# ============================================================

default_idx_piutang = (
    all_cols.index(
        auto_col_piutang
    )
    if auto_col_piutang in all_cols
    else 0
)


selected_col_piutang = st.sidebar.selectbox(
    "Pilih Kolom Piutang:",
    options=all_cols,
    index=default_idx_piutang
)


# ============================================================
# TRANSFORMASI DATA AWAL
# ============================================================

df["omset_clean"] = clean_numeric(
    df[
        selected_col_omset
    ]
)


df["piutang_clean"] = clean_numeric(
    df[
        selected_col_piutang
    ]
)


# ============================================================
# PARSING TANGGAL
# ============================================================

if COL_TANGGAL is not None:

    tanggal_asli = df[
        COL_TANGGAL
    ]

    df["tanggal_dt"] = pd.to_datetime(
        tanggal_asli,
        errors="coerce",
        dayfirst=True
    )

    mask_tanggal_gagal = (
        df["tanggal_dt"].isna()
        &
        tanggal_asli.notna()
    )

    if mask_tanggal_gagal.any():

        df.loc[
            mask_tanggal_gagal,
            "tanggal_dt"
        ] = pd.to_datetime(
            tanggal_asli[
                mask_tanggal_gagal
            ],
            errors="coerce"
        )

    df["tahun"] = (
        df[
            "tanggal_dt"
        ]
        .dt.year
        .fillna(0)
        .astype(int)
    )

    df["bulan"] = (
        df[
            "tanggal_dt"
        ]
        .dt.month
        .fillna(0)
        .astype(int)
    )

    df["nama_bulan"] = (
        df[
            "bulan"
        ]
        .map(
            nama_bulan_map
        )
        .fillna(
            "Tanpa Tanggal"
        )
    )

else:

    df["tanggal_dt"] = pd.NaT
    df["tahun"] = 0
    df["bulan"] = 0
    df["nama_bulan"] = "Tanpa Tanggal"


# ============================================================
# HAPUS BARIS YANG BUKAN DATA
# ============================================================

mask_data_valid = (
    df[
        selected_col_nama
    ].notna()
    |
    (
        df[
            "omset_clean"
        ] != 0
    )
    |
    df[
        "tanggal_dt"
    ].notna()
)


df = df[
    mask_data_valid
].copy()


df = df.reset_index(
    drop=True
)


# ============================================================
# DETEKSI KODE P
# ============================================================

df["nomor_p"] = (
    df[
        selected_col_nama
    ]
    .apply(
        ambil_nomor_p
    )
)


df["kode_klien"] = (
    df[
        selected_col_nama
    ]
    .apply(
        ambil_kode_dasar
    )
)


df["kategori_klien"] = (
    df[
        selected_col_nama
    ]
    .apply(
        tentukan_kategori_klien
    )
)


# ============================================================
# INJEKSI CLUSTER
# ============================================================

col_cluster_upload = cari_kolom(
    df,
    [
        "label_cluster",
        "cluster_biaya",
        "cluster"
    ]
)


if (
    col_cluster_upload
    and col_cluster_upload in df.columns
):

    df["label_cluster"] = (
        df[
            col_cluster_upload
        ]
        .astype(str)
        .replace(
            [
                "nan",
                "None",
                "none"
            ],
            "Cluster Tidak Ditemukan"
        )
    )

    sumber_cluster = (
        "Cluster dari file upload"
    )


else:

    cluster_series = (
        load_cluster_reference_series()
    )

    if (
        cluster_series is not None
        and len(cluster_series)
        == len(df)
    ):

        df["label_cluster"] = (
            cluster_series
            .reset_index(
                drop=True
            )
            .astype(str)
        )

        sumber_cluster = (
            "Cluster dari file patokan"
        )

    else:

        df["label_cluster"] = (
            df[
                "omset_clean"
            ]
            .apply(
                hitung_cluster_otomatis
            )
        )

        sumber_cluster = (
            "Klasifikasi otomatis "
            "berdasarkan batas nilai project"
        )


# ============================================================
# SIDEBAR FILTER
# ============================================================

st.sidebar.divider()

st.sidebar.markdown(
    "### 🔎 Filter Dashboard"
)


# ============================================================
# FILTER TAHUN
# ============================================================

tahun_list = sorted(
    [
        int(t)
        for t in df[
            "tahun"
        ].unique()
        if t != 0
    ]
)


tahun_pilihan = st.sidebar.multiselect(
    "Tahun",
    options=tahun_list,
    default=tahun_list
)


# ============================================================
# FILTER BULAN
# ============================================================

if tahun_pilihan:

    bulan_list = sorted(
        [
            int(b)
            for b in df[
                df[
                    "tahun"
                ].isin(
                    tahun_pilihan
                )
            ][
                "bulan"
            ].unique()
            if b != 0
        ]
    )

else:

    bulan_list = sorted(
        [
            int(b)
            for b in df[
                "bulan"
            ].unique()
            if b != 0
        ]
    )


bulan_pilihan = st.sidebar.multiselect(
    "Bulan",
    options=bulan_list,
    default=bulan_list,
    format_func=lambda x:
        nama_bulan_map.get(
            x,
            f"Bulan {x}"
        )
)


# ============================================================
# FILTER ADMIN
# ============================================================

if COL_ADMIN is not None:

    admin_list = sorted(
        df[
            COL_ADMIN
        ]
        .dropna()
        .astype(str)
        .unique()
        .tolist()
    )

    admin_pilihan = st.sidebar.multiselect(
        "Admin",
        options=admin_list,
        default=admin_list
    )

else:

    admin_list = []
    admin_pilihan = []


# ============================================================
# FILTER KONSULTAN
# ============================================================

if COL_KONSULTAN is not None:

    konsultan_list = sorted(
        df[
            COL_KONSULTAN
        ]
        .dropna()
        .astype(str)
        .unique()
        .tolist()
    )

    konsultan_pilihan = st.sidebar.multiselect(
        "Konsultan",
        options=konsultan_list,
        default=konsultan_list
    )

else:

    konsultan_list = []
    konsultan_pilihan = []


# ============================================================
# FILTER DATA
# ============================================================

df_filter = df.copy()


if (
    tahun_pilihan
    and len(tahun_pilihan)
    < len(tahun_list)
):

    df_filter = df_filter[
        df_filter[
            "tahun"
        ].isin(
            tahun_pilihan
        )
    ]


if (
    bulan_pilihan
    and len(bulan_pilihan)
    < len(bulan_list)
):

    df_filter = df_filter[
        df_filter[
            "bulan"
        ].isin(
            bulan_pilihan
        )
    ]


if (
    COL_ADMIN is not None
    and admin_pilihan
    and len(admin_pilihan)
    < len(admin_list)
):

    df_filter = df_filter[
        df_filter[
            COL_ADMIN
        ]
        .astype(str)
        .isin(
            admin_pilihan
        )
    ]


if (
    COL_KONSULTAN is not None
    and konsultan_pilihan
    and len(konsultan_pilihan)
    < len(konsultan_list)
):

    df_filter = df_filter[
        df_filter[
            COL_KONSULTAN
        ]
        .astype(str)
        .isin(
            konsultan_pilihan
        )
    ]


df_filter = df_filter.reset_index(
    drop=True
)


# ============================================================
# DASHBOARD HEADER
# ============================================================

st.markdown(
    '<div class="main-title">'
    '📊 Dashboard Gama Statistika'
    '</div>',
    unsafe_allow_html=True
)


st.markdown(
    '<div class="subtitle">'
    'Analisis pelanggan, transaksi, performa keuangan, '
    'project, konsultan, software, dan segmentasi cluster.'
    '</div>',
    unsafe_allow_html=True
)


st.markdown(
    f"""
    <div class="info-box">
    📌 Menampilkan <b>{len(df_filter):,}</b> project
    dari total <b>{len(df):,}</b> project.
    <br>
    📁 File: <b>{uploaded_file.name}</b>
    <br>
    📑 Sheet: <b>{sheet_pilihan if sheet_pilihan else "CSV"}</b>
    <br>
    🧩 Sumber cluster: <b>{sumber_cluster}</b>
    </div>
    """,
    unsafe_allow_html=True
)


# ============================================================
# METRIK UTAMA
# ============================================================

total_project = len(
    df_filter
)


total_omset = (
    df_filter[
        "omset_clean"
    ]
    .sum()
)


total_piutang = (
    df_filter[
        "piutang_clean"
    ]
    .sum()
)


jumlah_baru = (
    df_filter[
        "kategori_klien"
    ]
    == "Klien Baru"
).sum()


jumlah_repeat = (
    df_filter[
        "kategori_klien"
    ]
    == "Repeat Order"
).sum()


jumlah_tidak_diketahui = (
    df_filter[
        "kategori_klien"
    ]
    == "Tidak Diketahui"
).sum()


# ============================================================
# TOTAL KLIEN UNIK
# ============================================================

kode_valid = (
    df_filter[
        "kode_klien"
    ]
    .astype(str)
    .str.strip()
)


kode_valid = kode_valid[
    ~kode_valid.isin(
        [
            "",
            "nan",
            "none",
            "-"
        ]
    )
]


total_klien = (
    kode_valid
    .nunique()
)


if total_klien == 0:

    total_klien = (
        df_filter[
            selected_col_nama
        ]
        .dropna()
        .astype(str)
        .nunique()
    )


# ============================================================
# KPI TAMBAHAN
# ============================================================

aov = (
    total_omset / total_project
    if total_project > 0
    else 0
)


project_terbesar = (
    df_filter[
        "omset_clean"
    ].max()
    if total_project > 0
    else 0
)


project_terkecil = (
    df_filter[
        "omset_clean"
    ].min()
    if total_project > 0
    else 0
)


persen_baru = (
    jumlah_baru
    / total_project
    * 100
    if total_project > 0
    else 0
)


persen_repeat = (
    jumlah_repeat
    / total_project
    * 100
    if total_project > 0
    else 0
)


repeat_order_rate = (
    jumlah_repeat
    / total_klien
    * 100
    if total_klien > 0
    else 0
)


# ============================================================
# KPI UTAMA
# ============================================================

kpi1, kpi2, kpi3, kpi4 = st.columns(4)


def tampilkan_kpi(
    container,
    judul,
    nilai,
    kelas="kpi-value"
):

    with container:

        st.markdown(
            f"""
            <div class="kpi-card">
                <div class="kpi-title">
                    {judul}
                </div>
                <div class="{kelas}">
                    {nilai}
                </div>
            </div>
            """,
            unsafe_allow_html=True
        )


tampilkan_kpi(
    kpi1,
    "TOTAL PROJECT",
    f"{total_project:,}"
)


tampilkan_kpi(
    kpi2,
    "TOTAL OMSET",
    rupiah(total_omset)
)


tampilkan_kpi(
    kpi3,
    "KLIEN BARU",
    f"{jumlah_baru:,}"
)


tampilkan_kpi(
    kpi4,
    "REPEAT ORDER",
    f"{jumlah_repeat:,}"
)


# ============================================================
# KPI TAMBAHAN
# ============================================================

kpi5, kpi6, kpi7, kpi8 = st.columns(4)


tampilkan_kpi(
    kpi5,
    "RATA-RATA PROJECT",
    rupiah(aov)
)


tampilkan_kpi(
    kpi6,
    "PROJECT TERBESAR",
    rupiah(project_terbesar)
)


tampilkan_kpi(
    kpi7,
    "PROJECT TERKECIL",
    rupiah(project_terkecil)
)


tampilkan_kpi(
    kpi8,
    "TOTAL PIUTANG",
    rupiah(total_piutang),
    "kpi-value-danger"
)


# ============================================================
# OMSET BERDASARKAN KATEGORI KLIEN
# ============================================================

st.markdown(
    '<div class="section-title">'
    '💰 Omset Berdasarkan Kategori Klien'
    '</div>',
    unsafe_allow_html=True
)


omset_baru = (
    df_filter.loc[
        df_filter[
            "kategori_klien"
        ]
        == "Klien Baru",
        "omset_clean"
    ]
    .sum()
)


omset_repeat = (
    df_filter.loc[
        df_filter[
            "kategori_klien"
        ]
        == "Repeat Order",
        "omset_clean"
    ]
    .sum()
)


persen_omset_baru = (
    omset_baru
    / total_omset
    * 100
    if total_omset > 0
    else 0
)


persen_omset_repeat = (
    omset_repeat
    / total_omset
    * 100
    if total_omset > 0
    else 0
)


omset1, omset2, omset3, omset4 = st.columns(4)


tampilkan_kpi(
    omset1,
    "OMSET KLIEN BARU",
    rupiah(omset_baru)
)


tampilkan_kpi(
    omset2,
    "OMSET REPEAT ORDER",
    rupiah(omset_repeat)
)


tampilkan_kpi(
    omset3,
    "% OMSET KLIEN BARU",
    f"{persen_omset_baru:.1f}%"
)


tampilkan_kpi(
    omset4,
    "% OMSET REPEAT ORDER",
    f"{persen_omset_repeat:.1f}%"
)


# ============================================================
# OMSET KATEGORI CHART
# ============================================================

df_omset_kategori = pd.DataFrame({

    "Kategori": [
        "Klien Baru",
        "Repeat Order"
    ],

    "Omset": [
        omset_baru,
        omset_repeat
    ]

})


fig_omset_kategori = px.bar(
    df_omset_kategori,
    x="Kategori",
    y="Omset",
    text=df_omset_kategori[
        "Omset"
    ].apply(rupiah),
    labels={
        "Kategori": "Kategori Klien",
        "Omset": "Total Omset"
    },
    title=(
        "Perbandingan Omset "
        "Klien Baru vs Repeat Order"
    )
)


fig_omset_kategori.update_traces(
    textposition="outside"
)


fig_omset_kategori.update_layout(
    height=380,
    margin=dict(
        l=20,
        r=20,
        t=60,
        b=20
    )
)


st.plotly_chart(
    fig_omset_kategori,
    use_container_width=True
)


# ============================================================
# PERTUMBUHAN OMSET & PROJECT
# ============================================================

st.markdown(
    '<div class="section-title">'
    '📈 Pertumbuhan Performa'
    '</div>',
    unsafe_allow_html=True
)


df_bulanan_growth = df_filter[
    df_filter[
        "bulan"
    ] != 0
].copy()


if len(df_bulanan_growth) > 0:

    df_growth = (
        df_bulanan_growth
        .groupby(
            [
                "tahun",
                "bulan",
                "nama_bulan"
            ],
            as_index=False
        )
        .agg(
            omset=(
                "omset_clean",
                "sum"
            ),
            jumlah_project=(
                "omset_clean",
                "size"
            )
        )
    )


    df_growth[
        "periode_sort"
    ] = (
        df_growth[
            "tahun"
        ] * 100
        + df_growth[
            "bulan"
        ]
    )


    df_growth = (
        df_growth
        .sort_values(
            "periode_sort"
        )
        .reset_index(
            drop=True
        )
    )


    df_growth[
        "pertumbuhan_omset"
    ] = (
        df_growth[
            "omset"
        ]
        .pct_change()
        .replace(
            [np.inf, -np.inf],
            np.nan
        )
        * 100
    )


    df_growth[
        "pertumbuhan_project"
    ] = (
        df_growth[
            "jumlah_project"
        ]
        .pct_change()
        .replace(
            [np.inf, -np.inf],
            np.nan
        )
        * 100
    )


    growth_omset = (
        df_growth[
            "pertumbuhan_omset"
        ]
        .iloc[-1]
    )


    growth_project = (
        df_growth[
            "pertumbuhan_project"
        ]
        .iloc[-1]
    )


    if pd.isna(growth_omset):
        growth_omset = 0

    if pd.isna(growth_project):
        growth_project = 0


else:

    growth_omset = 0
    growth_project = 0


g1, g2 = st.columns(2)


tampilkan_kpi(
    g1,
    "PERSENTASE KLIEN BARU",
    f"{persen_baru:.1f}%"
)


tampilkan_kpi(
    g2,
    "PERSENTASE REPEAT",
    f"{persen_repeat:.1f}%"
)


# ============================================================
# OMSET BULANAN
# ============================================================

st.markdown(
    '<div class="section-title">'
    '📊 Performa Omset & Project Bulanan'
    '</div>',
    unsafe_allow_html=True
)


if len(df_bulanan_growth) > 0:

    df_omset = (
        df_bulanan_growth
        .groupby(
            [
                "tahun",
                "bulan",
                "nama_bulan"
            ],
            as_index=False
        )
        .agg(
            omset=(
                "omset_clean",
                "sum"
            ),
            jumlah_project=(
                "omset_clean",
                "size"
            ),
            klien_baru=(
                "kategori_klien",
                lambda x:
                (
                    x == "Klien Baru"
                ).sum()
            ),
            repeat_order=(
                "kategori_klien",
                lambda x:
                (
                    x == "Repeat Order"
                ).sum()
            )
        )
    )


    df_omset[
        "periode_sort"
    ] = (
        df_omset[
            "tahun"
        ] * 100
        + df_omset[
            "bulan"
        ]
    )


    df_omset = (
        df_omset
        .sort_values(
            "periode_sort"
        )
    )


    df_omset[
        "periode_label"
    ] = (
        df_omset[
            "nama_bulan"
        ]
        + " "
        + df_omset[
            "tahun"
        ].astype(str)
    )


    fig_omset = px.bar(
        df_omset,
        x="periode_label",
        y="omset",
        text=df_omset[
            "omset"
        ].apply(rupiah),
        labels={
            "periode_label": "Periode",
            "omset": "Total Omset"
        },
        title="Omset per Bulan"
    )


    fig_omset.update_traces(
        textposition="outside"
    )


    fig_omset.update_layout(
        height=430,
        margin=dict(
            l=20,
            r=20,
            t=60,
            b=20
        )
    )


    st.plotly_chart(
        fig_omset,
        use_container_width=True
    )


    fig_project_bulan = px.bar(
        df_omset,
        x="periode_label",
        y="jumlah_project",
        text="jumlah_project",
        labels={
            "periode_label": "Periode",
            "jumlah_project": "Jumlah Project"
        },
        title="Jumlah Project per Bulan"
    )


    fig_project_bulan.update_traces(
        textposition="outside"
    )


    fig_project_bulan.update_layout(
        height=400,
        margin=dict(
            l=20,
            r=20,
            t=60,
            b=20
        )
    )


    st.plotly_chart(
        fig_project_bulan,
        use_container_width=True
    )


else:

    st.info(
        "ℹ️ Tidak ada data tanggal untuk "
        "menampilkan analisis bulanan."
    )


# ============================================================
# KLIEN BARU VS REPEAT
# ============================================================

grafik1, grafik2 = st.columns(2)


with grafik1:

    st.markdown(
        '<div class="section-title">'
        '👥 Klien Baru vs Repeat Order'
        '</div>',
        unsafe_allow_html=True
    )


    if len(df_filter) > 0:

        df_kategori = (
            df_filter[
                "kategori_klien"
            ]
            .value_counts()
            .reset_index()
        )


        df_kategori.columns = [
            "kategori",
            "jumlah"
        ]


        color_map = {

            "Klien Baru": "#2563eb",

            "Repeat Order": "#16a34a",

            "Tidak Diketahui": "#9ca3af"

        }


        fig_kategori = px.pie(
            df_kategori,
            names="kategori",
            values="jumlah",
            hole=0.55,
            color="kategori",
            color_discrete_map=color_map
        )


        fig_kategori.update_layout(
            height=380,
            margin=dict(
                l=10,
                r=10,
                t=20,
                b=10
            )
        )


        st.plotly_chart(
            fig_kategori,
            use_container_width=True
        )


with grafik2:

    st.markdown(
        '<div class="section-title">'
        '📊 Distribusi Label Cluster'
        '</div>',
        unsafe_allow_html=True
    )


    if len(df_filter) > 0:

        df_dist_cluster = (
            df_filter[
                "label_cluster"
            ]
            .astype(str)
            .value_counts()
            .reset_index()
        )


        df_dist_cluster.columns = [
            "Label Cluster",
            "Jumlah Project"
        ]


        df_dist_cluster[
            "Persentase"
        ] = (
            df_dist_cluster[
                "Jumlah Project"
            ]
            / len(df_filter)
            * 100
        )


        fig_cluster_bar = px.bar(
            df_dist_cluster,
            x="Label Cluster",
            y="Jumlah Project",
            text=df_dist_cluster[
                "Persentase"
            ].apply(
                lambda x:
                f"{x:.1f}%"
            ),
            title="Distribusi Cluster"
        )


        fig_cluster_bar.update_traces(
            textposition="outside"
        )


        fig_cluster_bar.update_layout(
            height=380,
            margin=dict(
                l=10,
                r=10,
                t=40,
                b=20
            )
        )


        st.plotly_chart(
            fig_cluster_bar,
            use_container_width=True
        )


# ============================================================
# BATAS KLASIFIKASI
# ============================================================

st.markdown(
    '<div class="section-title">'
    '📏 Batas Nilai Klasifikasi'
    '</div>',
    unsafe_allow_html=True
)


# Variabel dibuat terlebih dahulu agar aman
# saat digunakan pada bagian download Excel.
batas_klasifikasi = pd.DataFrame()


if len(df_filter) > 0:

    batas_klasifikasi = (
        df_filter
        .groupby(
            "label_cluster"
        )
        .agg(
            jumlah_project=(
                "omset_clean",
                "size"
            ),
            total_omset=(
                "omset_clean",
                "sum"
            ),
            nilai_minimum=(
                "omset_clean",
                "min"
            ),
            nilai_maksimum=(
                "omset_clean",
                "max"
            ),
            rata_rata=(
                "omset_clean",
                "mean"
            )
        )
        .reset_index()
    )


    # ========================================================
    # PERSENTASE PROJECT PER KLASIFIKASI
    # ========================================================

    batas_klasifikasi[
        "persentase"
    ] = (
        batas_klasifikasi[
            "jumlah_project"
        ]
        / len(df_filter)
        * 100
    )


    # ========================================================
    # URUTAN CLUSTER
    # ========================================================

    batas_klasifikasi[
        "urutan"
    ] = (
        batas_klasifikasi[
            "label_cluster"
        ]
        .apply(
            urutan_cluster
        )
    )


    batas_klasifikasi = (
        batas_klasifikasi
        .sort_values(
            "urutan"
        )
        .drop(
            columns=[
                "urutan"
            ]
        )
        .reset_index(
            drop=True
        )
    )


    # ========================================================
    # VERSI UNTUK TAMPILAN
    # ========================================================

    batas_klasifikasi_view = (
        batas_klasifikasi.copy()
    )


    batas_klasifikasi_view[
        "total_omset"
    ] = (
        batas_klasifikasi_view[
            "total_omset"
        ]
        .apply(
            rupiah
        )
    )


    batas_klasifikasi_view[
        "nilai_minimum"
    ] = (
        batas_klasifikasi_view[
            "nilai_minimum"
        ]
        .apply(
            rupiah
        )
    )


    batas_klasifikasi_view[
        "nilai_maksimum"
    ] = (
        batas_klasifikasi_view[
            "nilai_maksimum"
        ]
        .apply(
            rupiah
        )
    )


    batas_klasifikasi_view[
        "rata_rata"
    ] = (
        batas_klasifikasi_view[
            "rata_rata"
        ]
        .apply(
            rupiah
        )
    )


    batas_klasifikasi_view[
        "persentase"
    ] = (
        batas_klasifikasi_view[
            "persentase"
        ]
        .apply(
            lambda x:
            f"{x:.1f}%"
        )
    )


    # ========================================================
    # RENAME KOLOM
    # ========================================================

    batas_klasifikasi_view.columns = [
        "Klasifikasi",
        "Jumlah Project",
        "Total Omset",
        "Nilai Minimum",
        "Nilai Maksimum",
        "Rata-rata Nilai Project",
        "Persentase"
    ]


    # ========================================================
    # TAMPILKAN TABEL
    # ========================================================

    st.dataframe(
        batas_klasifikasi_view,
        use_container_width=True,
        hide_index=True
    )


    # ========================================================
    # INFORMASI BATAS OTOMATIS
    # ========================================================

    st.info(
        f"""
**Batas klasifikasi otomatis:**

🔴 **Rendah:** < {rupiah(BATAS_RENDAH)}

🟡 **Menengah:** {rupiah(BATAS_RENDAH)} – < {rupiah(BATAS_TINGGI)}

🟢 **Tinggi:** ≥ {rupiah(BATAS_TINGGI)}

Sumber label saat ini: **{sumber_cluster}**
"""
    )


else:

    st.info(
        "Tidak ada data untuk menampilkan batas klasifikasi."
    )


# ============================================================
# DISTRIBUSI P.1, P.2, P.3, DST.
# ============================================================

st.markdown(
    '<div class="section-title">'
    '🔢 Distribusi Repeat Order P.1, P.2, P.3, dst.'
    '</div>',
    unsafe_allow_html=True
)


df_p = df_filter[
    df_filter[
        "nomor_p"
    ].notna()
].copy()


distribusi_p = pd.DataFrame()


if len(df_p) > 0:

    df_p[
        "nomor_p"
    ] = df_p[
        "nomor_p"
    ].astype(int)


    distribusi_p = (
        df_p[
            "nomor_p"
        ]
        .value_counts()
        .sort_index()
        .reset_index()
    )


    distribusi_p.columns = [
        "Nomor P",
        "Jumlah Project"
    ]


    distribusi_p[
        "Kode"
    ] = (
        "P."
        + distribusi_p[
            "Nomor P"
        ].astype(str)
    )


    distribusi_p[
        "Persentase"
    ] = (
        distribusi_p[
            "Jumlah Project"
        ]
        / len(df_p)
        * 100
    )


    distribusi_p_view = (
        distribusi_p[
            [
                "Kode",
                "Jumlah Project",
                "Persentase"
            ]
        ]
        .copy()
    )


    distribusi_p_view[
        "Persentase"
    ] = (
        distribusi_p_view[
            "Persentase"
        ]
        .apply(
            lambda x:
            f"{x:.1f}%"
        )
    )


    st.dataframe(
        distribusi_p_view,
        use_container_width=True,
        hide_index=True
    )


    fig_p = px.bar(
        distribusi_p,
        x="Kode",
        y="Jumlah Project",
        text="Jumlah Project",
        title="Jumlah Project Berdasarkan P"
    )


    fig_p.update_traces(
        textposition="outside"
    )


    fig_p.update_layout(
        height=400
    )


    st.plotly_chart(
        fig_p,
        use_container_width=True
    )


else:

    st.info(
        "Tidak ditemukan kode P.1, P.2, P.3, dst."
    )


# ============================================================
# ANALISIS SOFTWARE
# ============================================================

software_stats = pd.DataFrame()


if COL_SOFTWARE is not None:

    st.markdown(
        '<div class="section-title">'
        '💻 Analisis Software / Metode'
        '</div>',
        unsafe_allow_html=True
    )


    df_software = df_filter.copy()


    df_software[
        "software_clean"
    ] = (
        df_software[
            COL_SOFTWARE
        ]
        .fillna("Tidak Diketahui")
        .astype(str)
        .str.strip()
    )


    df_software[
        "software_clean"
    ] = (
        df_software[
            "software_clean"
        ]
        .replace(
            [
                "",
                "nan",
                "None"
            ],
            "Tidak Diketahui"
        )
    )


    software_stats = (
        df_software[
            "software_clean"
        ]
        .value_counts()
        .reset_index()
    )


    software_stats.columns = [
        "Software",
        "Jumlah Project"
    ]


    software_stats[
        "Persentase"
    ] = (
        software_stats[
            "Jumlah Project"
        ]
        / len(df_software)
        * 100
    )


    software_teratas = (
        software_stats.iloc[0]["Software"]
        if len(software_stats) > 0
        else "Tidak diketahui"
    )


    software_jumlah = (
        int(
            software_stats.iloc[0]["Jumlah Project"]
        )
        if len(software_stats) > 0
        else 0
    )


    s1, s2 = st.columns(2)


    tampilkan_kpi(
        s1,
        "SOFTWARE DOMINAN",
        str(software_teratas)
    )


    tampilkan_kpi(
        s2,
        "PROJECT SOFTWARE DOMINAN",
        f"{software_jumlah:,}"
    )


    software_view = software_stats.copy()


    software_view[
        "Persentase"
    ] = (
        software_view[
            "Persentase"
        ]
        .apply(
            lambda x:
            f"{x:.1f}%"
        )
    )


    st.dataframe(
        software_view,
        use_container_width=True,
        hide_index=True
    )


    fig_software = px.bar(
        software_stats.head(10),
        x="Software",
        y="Jumlah Project",
        text="Jumlah Project",
        title="Top 10 Software / Metode"
    )


    fig_software.update_traces(
        textposition="outside"
    )


    fig_software.update_layout(
        height=430
    )


    st.plotly_chart(
        fig_software,
        use_container_width=True
    )


# ============================================================
# ANALISIS KONSULTAN
# ============================================================

konsultan_stats = pd.DataFrame()


if COL_KONSULTAN is not None:

    st.markdown(
        '<div class="section-title">'
        '👨‍💼 Analisis Konsultan'
        '</div>',
        unsafe_allow_html=True
    )


    df_konsultan = df_filter.copy()


    df_konsultan[
        "konsultan_clean"
    ] = (
        df_konsultan[
            COL_KONSULTAN
        ]
        .fillna("Tidak Diketahui")
        .astype(str)
        .str.strip()
    )


    df_konsultan[
        "konsultan_clean"
    ] = (
        df_konsultan[
            "konsultan_clean"
        ]
        .replace(
            [
                "",
                "nan",
                "None"
            ],
            "Tidak Diketahui"
        )
    )


    konsultan_stats = (
        df_konsultan
        .groupby(
            "konsultan_clean",
            as_index=False
        )
        .agg(
            jumlah_project=(
                "omset_clean",
                "size"
            ),
            total_omset=(
                "omset_clean",
                "sum"
            ),
            rata_rata_project=(
                "omset_clean",
                "mean"
            )
        )
        .sort_values(
            "jumlah_project",
            ascending=False
        )
    )


    konsultan_stats[
        "persentase_project"
    ] = (
        konsultan_stats[
            "jumlah_project"
        ]
        / len(df_konsultan)
        * 100
    )


    konsultan_teratas = (
        konsultan_stats.iloc[0][
            "konsultan_clean"
        ]
        if len(konsultan_stats) > 0
        else "Tidak diketahui"
    )


    jumlah_project_konsultan = (
        int(
            konsultan_stats.iloc[0][
                "jumlah_project"
            ]
        )
        if len(konsultan_stats) > 0
        else 0
    )


    k1, k2 = st.columns(2)


    tampilkan_kpi(
        k1,
        "KONSULTAN DOMINAN",
        str(konsultan_teratas)
    )


    tampilkan_kpi(
        k2,
        "PROJECT DITANGANI",
        f"{jumlah_project_konsultan:,}"
    )


    konsultan_view = (
        konsultan_stats.copy()
    )


    konsultan_view[
        "total_omset"
    ] = (
        konsultan_view[
            "total_omset"
        ]
        .apply(
            rupiah
        )
    )


    konsultan_view[
        "rata_rata_project"
    ] = (
        konsultan_view[
            "rata_rata_project"
        ]
        .apply(
            rupiah
        )
    )


    konsultan_view[
        "persentase_project"
    ] = (
        konsultan_view[
            "persentase_project"
        ]
        .apply(
            lambda x:
            f"{x:.1f}%"
        )
    )


    konsultan_view.columns = [
        "Konsultan",
        "Jumlah Project",
        "Total Omset",
        "Rata-rata Project",
        "Persentase Project"
    ]


    st.dataframe(
        konsultan_view,
        use_container_width=True,
        hide_index=True
    )


    fig_konsultan = px.bar(
        konsultan_stats.head(10),
        x="konsultan_clean",
        y="jumlah_project",
        text="jumlah_project",
        title="Top 10 Konsultan Berdasarkan Jumlah Project"
    )


    fig_konsultan.update_traces(
        textposition="outside"
    )


    fig_konsultan.update_layout(
        height=430
    )


    st.plotly_chart(
        fig_konsultan,
        use_container_width=True
    )


# ============================================================
# TOP KLIEN
# ============================================================

st.markdown(
    '<div class="section-title">'
    '🏆 Top Klien'
    '</div>',
    unsafe_allow_html=True
)


df_top_klien = df_filter.copy()


df_top_klien[
    "klien_display"
] = (
    df_top_klien[
        selected_col_nama
    ]
    .fillna("Tidak Diketahui")
    .astype(str)
)


top_klien = (
    df_top_klien
    .groupby(
        "klien_display",
        as_index=False
    )
    .agg(
        jumlah_project=(
            "omset_clean",
            "size"
        ),
        total_omset=(
            "omset_clean",
            "sum"
        ),
        total_piutang=(
            "piutang_clean",
            "sum"
        )
    )
    .sort_values(
        [
            "jumlah_project",
            "total_omset"
        ],
        ascending=False
    )
    .head(10)
)


top_klien_view = top_klien.copy()


top_klien_view[
    "total_omset"
] = (
    top_klien_view[
        "total_omset"
    ]
    .apply(
        rupiah
    )
)


top_klien_view[
    "total_piutang"
] = (
    top_klien_view[
        "total_piutang"
    ]
    .apply(
        rupiah
    )
)


top_klien_view.columns = [
    "Klien",
    "Jumlah Project",
    "Total Omset",
    "Total Piutang"
]


st.dataframe(
    top_klien_view,
    use_container_width=True,
    hide_index=True
)


# ============================================================
# TOP PROJECT
# ============================================================

st.markdown(
    '<div class="section-title">'
    '💎 Top 10 Project Berdasarkan Nilai'
    '</div>',
    unsafe_allow_html=True
)


top_project = (
    df_filter
    .sort_values(
        "omset_clean",
        ascending=False
    )
    .head(10)
    .copy()
)


top_project_view = top_project[
    [
        selected_col_nama,
        "omset_clean",
        "kategori_klien",
        "label_cluster"
    ]
].copy()


top_project_view[
    "omset_clean"
] = (
    top_project_view[
        "omset_clean"
    ]
    .apply(
        rupiah
    )
)


top_project_view.columns = [
    "Klien/Kode",
    "Nilai Project",
    "Kategori Klien",
    "Cluster"
]


st.dataframe(
    top_project_view,
    use_container_width=True,
    hide_index=True
)


# ============================================================
# DETAIL SELURUH TRANSAKSI
# ============================================================

st.markdown(
    '<div class="section-title">'
    '📋 Detail Seluruh Transaksi'
    '</div>',
    unsafe_allow_html=True
)


df_tampil = df_filter.copy()


# ============================================================
# TAMBAHKAN INFORMASI OTOMATIS
# ============================================================

kolom_tambahan = [
    "nomor_p",
    "kode_klien",
    "kategori_klien",
    "omset_clean",
    "piutang_clean",
    "tanggal_dt",
    "tahun",
    "bulan",
    "nama_bulan",
    "label_cluster"
]


for kolom in kolom_tambahan:

    if kolom not in df_tampil.columns:
        df_tampil[kolom] = ""


# ============================================================
# PINDAHKAN KOLOM PENTING KE KANAN
# ============================================================

kolom_ke_kanan = [
    "nomor_p",
    "kode_klien",
    "kategori_klien",
    "omset_clean",
    "piutang_clean",
    "tanggal_dt",
    "tahun",
    "bulan",
    "nama_bulan",
    "label_cluster"
]


kolom_awal = [
    c
    for c in df_tampil.columns
    if c not in kolom_ke_kanan
]


df_tampil = df_tampil[
    kolom_awal
    + kolom_ke_kanan
]


# ============================================================
# BUAT DATA UNTUK TAMPILAN
# ============================================================

df_tampil_view = (
    df_tampil.copy()
)


# ============================================================
# FORMAT KOLOM OMSET
# ============================================================

if (
    "omset_clean"
    in df_tampil_view.columns
):

    df_tampil_view[
        "omset_clean"
    ] = (
        df_tampil_view[
            "omset_clean"
        ]
        .apply(
            rupiah
        )
    )


# ============================================================
# FORMAT PIUTANG
# ============================================================

if (
    "piutang_clean"
    in df_tampil_view.columns
):

    df_tampil_view[
        "piutang_clean"
    ] = (
        df_tampil_view[
            "piutang_clean"
        ]
        .apply(
            rupiah
        )
    )


# ============================================================
# FORMAT TANGGAL
# ============================================================

if (
    "tanggal_dt"
    in df_tampil_view.columns
):

    df_tampil_view[
        "tanggal_dt"
    ] = (
        pd.to_datetime(
            df_tampil_view[
                "tanggal_dt"
            ],
            errors="coerce"
        )
        .dt.strftime(
            "%d-%m-%Y"
        )
        .fillna("")
    )


# ============================================================
# KONVERSI STRING UNTUK TAMPILAN
# ============================================================

for col_item in (
    df_tampil_view.columns
):

    df_tampil_view[
        col_item
    ] = (
        df_tampil_view[
            col_item
        ]
        .fillna("")
        .astype(str)
    )


# ============================================================
# TAMPILKAN DATA
# ============================================================

st.dataframe(
    df_tampil_view,
    use_container_width=True,
    height=500
)

# ============================================================
# DOWNLOAD DATA & HASIL DASHBOARD
# ============================================================

st.markdown(
    '<div class="section-title">'
    '📥 Download Data & Hasil Dashboard'
    '</div>',
    unsafe_allow_html=True
)

col_dl1, col_dl2 = st.columns(2)


# ============================================================
# 1. DOWNLOAD DETAIL TRANSAKSI + CLUSTER (CSV)
# ============================================================

with col_dl1:

    csv_detail = (
        df_tampil
        .to_csv(index=False)
        .encode("utf-8-sig")
    )

    st.download_button(
        label="📥 Download Detail Transaksi + Cluster (CSV)",
        data=csv_detail,
        file_name="Detail_Transaksi_Label_Cluster.csv",
        mime="text/csv",
        use_container_width=True
    )


# ============================================================
# 2. DOWNLOAD HASIL DASHBOARD (EXCEL)
# ============================================================

with col_dl2:

    output = io.BytesIO()

    try:

        with pd.ExcelWriter(
            output,
            engine="openpyxl"
        ) as writer:

            # =================================================
            # SHEET 1 - RINGKASAN DASHBOARD
            # =================================================

            total_omset_export = 0

            if (
                isinstance(df, pd.DataFrame)
                and not df.empty
                and "biaya_total" in df.columns
            ):

                total_omset_export = (
                    pd.to_numeric(
                        df["biaya_total"],
                        errors="coerce"
                    )
                    .fillna(0)
                    .sum()
                )


            ringkasan_export = pd.DataFrame({

                "Keterangan": [
                    "Total Klien",
                    "Klien Baru",
                    "Repeat Order",
                    "Total Transaksi",
                    "Total Omset"
                ],

                "Nilai": [

                    int(total_klien)
                    if "total_klien" in locals()
                    else 0,

                    int(jumlah_klien_baru)
                    if "jumlah_klien_baru" in locals()
                    else 0,

                    int(jumlah_repeat_order)
                    if "jumlah_repeat_order" in locals()
                    else 0,

                    int(total_transaksi)
                    if "total_transaksi" in locals()
                    else 0,

                    total_omset_export
                ]
            })


            ringkasan_export.to_excel(
                writer,
                index=False,
                sheet_name="Ringkasan Dashboard"
            )


            # =================================================
            # SHEET 2 - DATA TRANSAKSI
            # =================================================

            if (
                isinstance(df_tampil, pd.DataFrame)
                and not df_tampil.empty
            ):

                df_tampil.to_excel(
                    writer,
                    index=False,
                    sheet_name="Data Transaksi"
                )


            # =================================================
            # SHEET 3 - BATAS KLASIFIKASI
            # =================================================

            if (
                isinstance(batas_klasifikasi, pd.DataFrame)
                and not batas_klasifikasi.empty
            ):

                batas_klasifikasi.to_excel(
                    writer,
                    index=False,
                    sheet_name="Batas Klasifikasi"
                )


        # =====================================================
        # PASTIKAN ADA SHEET YANG VISIBLE
        # =====================================================

        output.seek(0)


        # =====================================================
        # TOMBOL DOWNLOAD EXCEL
        # =====================================================

        st.download_button(

            label="📊 Download Hasil Dashboard (Excel)",

            data=output.getvalue(),

            file_name="Dashboard_Gama_Statistika.xlsx",

            mime=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),

            use_container_width=True
        )


    except Exception as e:

        st.error(
            f"❌ Gagal membuat file Excel: {e}"
        )