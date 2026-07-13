# -*- coding: utf-8 -*-
"""
SALES INTELLIGENCE DASHBOARD
Distributor Bahan Bangunan / Cat
Dibuat sebagai alat bantu manajemen untuk meningkatkan omset, AO, customer aktif,
dan mengidentifikasi peluang penjualan berbasis data.

VERSI REPO/DEPLOY:
    Data dibaca otomatis dari folder ./data di repository ini:
        - data/Data_Omset_Power_Query_tahun_2026.xlsx  (data transaksi)
        - data/DATA_TARGET_FUAD.xlsx                    (data target, opsional)
    Nama file bisa diganti lewat environment variable / edit konstanta di bawah,
    dan pengguna tetap bisa override dengan upload manual dari sidebar jika perlu.

Cara menjalankan lokal:
    pip install -r requirements.txt
    streamlit run app.py

Cara deploy ke Streamlit Community Cloud:
    1. Push folder ini (app.py, requirements.txt, data/) ke repo GitHub.
    2. Buka https://share.streamlit.io -> New app -> pilih repo & branch -> main file: app.py
    3. Deploy. Selesai.
"""

import os
import re
import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go

# =============================================================================
# KONFIGURASI FILE DATA BAWAAN (DI DALAM REPO)
# =============================================================================
DATA_DIR = "data"
DEFAULT_TRANSAKSI_FILE = os.path.join(DATA_DIR, "Data_Omset_Power_Query_tahun_2026.xlsx")
DEFAULT_TARGET_FILE = os.path.join(DATA_DIR, "DATA_TARGET_FUAD.xlsx")

# =============================================================================
# KONFIGURASI HALAMAN & STYLE
# =============================================================================
st.set_page_config(
    page_title="Sales Intelligence Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

COLOR_GOOD = "#1E8449"  # hijau - baik
COLOR_WARN = "#B7950B"  # kuning - perlu perhatian
COLOR_BAD = "#C0392B"  # merah - buruk
COLOR_PRIMARY = "#154360"  # biru navy - primary
COLOR_ACCENT = "#2E86C1"  # biru terang - accent
PALETTE = [
    "#154360",
    "#2E86C1",
    "#1E8449",
    "#B7950B",
    "#C0392B",
    "#7D3C98",
    "#117864",
    "#A04000",
    "#5D6D7E",
    "#873600",
]

CUSTOM_CSS = f"""
<style>
    .main {{ background-color: #F4F6F7; }}
    .block-container {{ padding-top: 1.2rem; }}
    div[data-testid="stMetric"] {{
        background-color: #FFFFFF;
        border: 1px solid #E5E8E8;
        border-left: 5px solid {COLOR_PRIMARY};
        border-radius: 10px;
        padding: 14px 16px 8px 16px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.08);
    }}
    div[data-testid="stMetricLabel"] {{ font-weight: 600; color: #34495E; }}
    .insight-box {{
        background-color: #EAF2F8;
        border-left: 5px solid {COLOR_ACCENT};
        padding: 12px 16px;
        border-radius: 8px;
        margin-bottom: 8px;
        font-size: 0.95rem;
    }}
    .insight-good {{ border-left-color: {COLOR_GOOD}; background-color: #E9F7EF; }}
    .insight-warn {{ border-left-color: {COLOR_WARN}; background-color: #FEF9E7; }}
    .insight-bad {{ border-left-color: {COLOR_BAD}; background-color: #FDEDEC; }}
    .page-title {{
        color: {COLOR_PRIMARY}; font-weight: 800; margin-bottom: 0px;
    }}
    .page-subtitle {{ color: #566573; margin-top: 0px; margin-bottom: 1rem; }}
    section[data-testid="stSidebar"] {{ background-color: #FFFFFF; }}
</style>
"""
st.markdown(CUSTOM_CSS, unsafe_allow_html=True)

REQUIRED_COLS = [
    "KOTA",
    "TGL FKTR",
    "SUPP",
    "NOMINAL",
    "DEPO",
    "SALES",
    "MONTH",
    "TANGGAL",
    "DIVISI",
    "AREA",
    "SBD/NON",
    "KD GRUP",
    "KD-SUPP",
    "TELE",
    "NOO/NPD",
]

# =============================================================================
# FUNGSI UTILITAS
# =============================================================================


def fmt_rp(x):
    """Format angka menjadi format Rupiah singkat."""
    if pd.isna(x):
        return "-"
    sign = "-" if x < 0 else ""
    x = abs(x)
    if x >= 1e9:
        return f"{sign}Rp {x/1e9:,.2f} M"
    if x >= 1e6:
        return f"{sign}Rp {x/1e6:,.1f} Jt"
    if x >= 1e3:
        return f"{sign}Rp {x/1e3:,.0f} Rb"
    return f"{sign}Rp {x:,.0f}"


def fmt_num(x):
    if pd.isna(x):
        return "-"
    return f"{x:,.0f}"


def fmt_pct(x):
    if pd.isna(x):
        return "-"
    return f"{x:,.1f}%"


def insight(text, level="info"):
    cls = {
        "good": "insight-good",
        "warn": "insight-warn",
        "bad": "insight-bad",
        "info": "insight-box",
    }.get(level, "insight-box")
    icon = {"good": "✅", "warn": "⚠️", "bad": "🔴", "info": "💡"}.get(level, "💡")
    # Teks disisipkan ke dalam <div> HTML mentah, sehingga sintaks markdown **bold**
    # tidak otomatis diproses oleh Streamlit. Konversi manual ke <strong> di sini.
    text = re.sub(r"\*\*(.+?)\*\*", r"<strong>\1</strong>", text)
    st.markdown(
        f'<div class="insight-box {cls}">{icon} {text}</div>', unsafe_allow_html=True
    )


def style_df(df_in, decimals=0, pct_cols=None):
    """Format kolom numerik dengan pemisah ribuan agar tabel lebih mudah dibaca.

    decimals   : jumlah desimal default untuk kolom numerik
    pct_cols   : daftar nama kolom yang berisi persentase (diformat 1 desimal + '%')
    """
    pct_cols = pct_cols or []
    num_cols = df_in.select_dtypes(include=[np.number]).columns
    fmt = {}
    for c in num_cols:
        if c in pct_cols:
            fmt[c] = "{:,.1f}%"
        else:
            fmt[c] = f"{{:,.{decimals}f}}"
    return df_in.style.format(fmt, na_rep="-")


@st.cache_data(show_spinner=False)
def load_data(file):
    """`file` bisa berupa path string (data bawaan repo) atau file upload Streamlit."""
    name = file.name if hasattr(file, "name") else str(file)
    if name.lower().endswith(".csv"):
        df = pd.read_csv(file)
    else:
        df = pd.read_excel(file)
    df.columns = [str(c).strip().upper() for c in df.columns]
    if "TGL FKTR" in df.columns:
        # coba format eksplisit "26-Jan-26" dulu (lebih cepat & tanpa warning),
        # baru fallback ke parser umum untuk format tanggal lain
        parsed = pd.to_datetime(df["TGL FKTR"], format="%d-%b-%y", errors="coerce")
        mask_fail = parsed.isna() & df["TGL FKTR"].notna()
        if mask_fail.any():
            parsed.loc[mask_fail] = pd.to_datetime(
                df.loc[mask_fail, "TGL FKTR"], errors="coerce", dayfirst=True
            )
        df["TGL FKTR"] = parsed
    if "NOMINAL" in df.columns:
        df["NOMINAL"] = (
            df["NOMINAL"].astype(str).str.replace(r"[^0-9\.\-]", "", regex=True)
        )
        df["NOMINAL"] = pd.to_numeric(df["NOMINAL"], errors="coerce").fillna(0)
    for c in [
        "KOTA",
        "SUPP",
        "DEPO",
        "SALES",
        "MONTH",
        "DIVISI",
        "AREA",
        "SBD/NON",
        "KD GRUP",
        "KD-SUPP",
        "TELE",
        "NOO/NPD",
    ]:
        if c in df.columns:
            df[c] = df[c].astype(str).str.strip()
            df[c] = df[c].replace({"nan": np.nan, "None": np.nan, "": np.nan})
    # "0" pada kolom TELE/AREA/NOO-NPD biasanya berarti "tidak ada / tidak berlaku",
    # bukan nama/kode sungguhan -> dianggap kosong agar tidak muncul sebagai kategori filter
    for c in ["TELE", "AREA", "NOO/NPD"]:
        if c in df.columns:
            df[c] = df[c].replace({"0": np.nan})
    return df


MONTH_COLS = [
    "JAN",
    "FEB",
    "MAR",
    "APR",
    "MAY",
    "JUN",
    "JUL",
    "AUG",
    "SEP",
    "OCT",
    "NOV",
    "DEC",
]


@st.cache_data(show_spinner=False)
def load_target_data(file):
    """Baca file target penjualan (format: kolom bulan JAN..DEC per baris salesman/supplier)."""
    df = pd.read_excel(file)
    df.columns = [str(c).strip().upper() for c in df.columns]
    for c in ["DEPO", "DIVISI", "SALES YG COVER", "SUPPLIER", "NAMA SALESMAN"]:
        if c in df.columns:
            df[c] = df[c].astype(str).str.strip()
    for c in MONTH_COLS:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)
    return df


def has(df, col):
    return col in df.columns


def calc_omset(d):
    return d["NOMINAL"].sum() if has(d, "NOMINAL") else np.nan  # type: ignore[reportAttributeAccessIssue]


def calc_ao(d):
    return d["KD GRUP"].nunique() if has(d, "KD GRUP") else np.nan


def calc_customer_aktif(d):
    return d["KD GRUP"].nunique() if has(d, "KD GRUP") else np.nan


def calc_faktur(d):
    return len(d)


def is_new_customer_mask(d):
    """True untuk baris dengan tanda NOO atau NPD (customer baru). Kosong = existing customer."""
    if not has(d, "NOO/NPD"):
        return pd.Series(False, index=d.index)
    return d["NOO/NPD"].notna() & (d["NOO/NPD"].astype(str).str.strip() != "")


def calc_noo(d):
    if has(d, "NOO/NPD") and has(d, "KD GRUP"):
        mask = is_new_customer_mask(d)
        return d.loc[mask, "KD GRUP"].nunique()
    return np.nan


def dimension_summary(d, dim_col):
    """Ringkasan Omset/AO/Customer Aktif/Faktur per dimensi (Sales, Supplier, Kota, dst)."""
    if not has(d, dim_col) or not has(d, "NOMINAL"):
        return pd.DataFrame()
    grp = (
        d.groupby(dim_col)
        .agg(OMSET=("NOMINAL", "sum"), JUMLAH_FAKTUR=("NOMINAL", "count"))
        .reset_index()
    )
    if has(d, "KD GRUP"):
        ao = d.groupby(dim_col)["KD GRUP"].nunique().reset_index(name="AO")
        grp = grp.merge(ao, on=dim_col, how="left")
    else:
        grp["AO"] = np.nan
    if has(d, "KD GRUP"):
        cust = (
            d.groupby(dim_col)["KD GRUP"].nunique().reset_index(name="CUSTOMER_AKTIF")
        )
        grp = grp.merge(cust, on=dim_col, how="left")
    else:
        grp["CUSTOMER_AKTIF"] = np.nan
    grp["AVG_OMSET_PER_AO"] = grp["OMSET"] / grp["AO"].replace(0, np.nan)
    grp["AVG_OMSET_PER_CUSTOMER"] = grp["OMSET"] / grp["CUSTOMER_AKTIF"].replace(
        0, np.nan
    )
    grp["AVG_OMSET_PER_FAKTUR"] = grp["OMSET"] / grp["JUMLAH_FAKTUR"].replace(0, np.nan)
    total = grp["OMSET"].sum()  # type: ignore[reportAttributeAccessIssue]
    grp["KONTRIBUSI_%"] = np.where(total != 0, grp["OMSET"] / total * 100, 0)
    return grp.sort_values("OMSET", ascending=False).reset_index(drop=True)


def noo_summary(d, dim_col):
    if not (has(d, "NOO/NPD") and has(d, dim_col) and has(d, "KD GRUP")):
        return pd.DataFrame()
    noo_df = d[is_new_customer_mask(d)]
    if noo_df.empty:
        return pd.DataFrame()
    grp = noo_df.groupby(dim_col)["KD GRUP"].nunique().reset_index(name="NOO_COUNT")
    return grp.sort_values("NOO_COUNT", ascending=False).reset_index(drop=True)


def monthly_trend(d):
    if has(d, "TGL FKTR") and d["TGL FKTR"].notna().any():
        tmp = d.dropna(subset=["TGL FKTR"]).copy()
        tmp["PERIOD"] = tmp["TGL FKTR"].dt.to_period("M").astype(str)
        grp = (
            tmp.groupby("PERIOD")
            .agg(OMSET=("NOMINAL", "sum"), JUMLAH_FAKTUR=("NOMINAL", "count"))
            .reset_index()
        )
        if has(tmp, "KD GRUP"):
            kd_grup_nunique = (
                tmp.groupby("PERIOD")["KD GRUP"]
                .nunique()
                .reset_index(name="_KD_GRUP_NUNIQUE")
            )
            grp = grp.merge(
                kd_grup_nunique.rename(columns={"_KD_GRUP_NUNIQUE": "AO"}),
                on="PERIOD",
                how="left",
            )
            grp = grp.merge(
                kd_grup_nunique.rename(columns={"_KD_GRUP_NUNIQUE": "CUSTOMER_AKTIF"}),
                on="PERIOD",
                how="left",
            )
        return grp.sort_values("PERIOD").reset_index(drop=True)
    if has(d, "MONTH"):
        grp = (
            d.groupby("MONTH")
            .agg(OMSET=("NOMINAL", "sum"), JUMLAH_FAKTUR=("NOMINAL", "count"))
            .reset_index()
        )
        grp = grp.rename(columns={"MONTH": "PERIOD"})
        return grp
    return pd.DataFrame()


def pareto_abc(d):
    if not (has(d, "KD GRUP") and has(d, "NOMINAL")):
        return pd.DataFrame()
    cust = d.groupby("KD GRUP")["NOMINAL"].sum().reset_index()  # type: ignore[reportAttributeAccessIssue]
    cust = cust.sort_values("NOMINAL", ascending=False).reset_index(drop=True)
    total = cust["NOMINAL"].sum()  # type: ignore[reportAttributeAccessIssue]
    cust["CUM_OMSET"] = cust["NOMINAL"].cumsum()
    cust["CUM_PCT"] = np.where(total != 0, cust["CUM_OMSET"] / total * 100, 0)

    def abc_class(pct):
        if pct <= 80:
            return "A"
        if pct <= 95:
            return "B"
        return "C"

    cust["ABC"] = cust["CUM_PCT"].apply(abc_class)
    return cust


def quadrant_label(x, y, x_med, y_med):
    if x >= x_med and y >= y_med:
        return "AO Tinggi - Omset Tinggi"
    if x >= x_med and y < y_med:
        return "AO Tinggi - Omset Rendah"
    if x < x_med and y >= y_med:
        return "AO Rendah - Omset Tinggi"
    return "AO Rendah - Omset Rendah"


def styled_bar(
    df_in,
    x_col,
    y_col,
    title,
    color=COLOR_PRIMARY,
    top_n=15,
    horizontal=True,
    text_fmt=None,
):
    d = df_in.head(top_n).copy()
    if horizontal:
        d = d.sort_values(y_col, ascending=True)
        fig = px.bar(
            d,
            x=y_col,
            y=x_col,
            orientation="h",
            title=title,
            text=y_col,
            color_discrete_sequence=[color],
        )
    else:
        fig = px.bar(
            d,
            x=x_col,
            y=y_col,
            title=title,
            text=y_col,
            color_discrete_sequence=[color],
        )
    # Label angka di atas/samping bar selalu pakai pemisah ribuan (co: 1,234,567)
    fig.update_traces(texttemplate="%{text:,.0f}", textposition="outside")
    # Sumbu angka (nilai) juga diberi pemisah ribuan agar tick label mudah dibaca
    if horizontal:
        fig.update_xaxes(tickformat=",.0f")
    else:
        fig.update_yaxes(tickformat=",.0f")
    fig.update_layout(
        height=420,
        margin=dict(l=10, r=10, t=50, b=10),
        plot_bgcolor="white",
        title_font_color=COLOR_PRIMARY,
    )
    return fig


def kpi_row(items):
    """items: list of (label, value, delta)"""
    cols = st.columns(len(items))
    for c, (label, value, delta) in zip(cols, items):
        c.metric(label, value, delta)


def page_header(title, subtitle):
    st.markdown(f'<h2 class="page-title">{title}</h2>', unsafe_allow_html=True)
    st.markdown(f'<p class="page-subtitle">{subtitle}</p>', unsafe_allow_html=True)


def growth_pct(series):
    """Hitung growth % periode terakhir vs sebelumnya dari sebuah series terurut waktu."""
    s = series.dropna()
    if len(s) < 2 or s.iloc[-2] == 0:
        return np.nan
    return (s.iloc[-1] - s.iloc[-2]) / s.iloc[-2] * 100


# =============================================================================
# SIDEBAR — SUMBER DATA & FILTER
# =============================================================================
st.sidebar.markdown("## 📊 Sales Intelligence")
st.sidebar.caption("Dashboard Distributor Bahan Bangunan / Cat")

col_refresh, col_info = st.sidebar.columns([1, 1])
with col_refresh:
    if st.button("🔄 Refresh Data", use_container_width=True,
                  help="Bersihkan cache dan baca ulang data (pakai ini setelah file data diganti/diupdate)"):
        st.cache_data.clear()
        st.rerun()
with col_info:
    if st.button("🧹 Clear Cache", use_container_width=True,
                  help="Bersihkan seluruh cache aplikasi tanpa langsung reload halaman"):
        st.cache_data.clear()
        st.cache_resource.clear()
        st.sidebar.success("Cache dibersihkan ✅")

st.sidebar.markdown("---")

st.sidebar.markdown("### 📁 Sumber Data")
default_transaksi_ada = os.path.exists(DEFAULT_TRANSAKSI_FILE)
default_target_ada = os.path.exists(DEFAULT_TARGET_FILE)

sumber = st.sidebar.radio(
    "Pilih sumber data",
    ["Data bawaan repo (default)", "Upload manual"],
    index=0 if default_transaksi_ada else 1,
    help=(
        "Data bawaan repo dibaca otomatis dari folder data/ di repository ini. "
        "Pilih 'Upload manual' untuk mencoba file lain tanpa mengubah repo."
    ),
)

uploaded = None
uploaded_target = None

if sumber == "Upload manual":
    uploaded = st.sidebar.file_uploader(
        "📁 Upload Data Transaksi", type=["csv", "xlsx", "xls"]
    )
    uploaded_target = st.sidebar.file_uploader(
        "🎯 Upload Data Target (opsional)", type=["xlsx", "xls"]
    )
    data_source = uploaded
    target_source = uploaded_target
else:
    if not default_transaksi_ada:
        st.sidebar.error(
            f"File data bawaan tidak ditemukan di `{DEFAULT_TRANSAKSI_FILE}`. "
            "Pastikan file sudah ada di folder data/ pada repo, atau pilih 'Upload manual'."
        )
    data_source = DEFAULT_TRANSAKSI_FILE if default_transaksi_ada else None
    target_source = DEFAULT_TARGET_FILE if default_target_ada else None
    if default_transaksi_ada:
        st.sidebar.success(f"✅ Memakai: `{DEFAULT_TRANSAKSI_FILE}`")
    if default_target_ada:
        st.sidebar.success(f"✅ Memakai: `{DEFAULT_TARGET_FILE}`")
    else:
        st.sidebar.caption("ℹ️ Data target opsional tidak ditemukan di repo.")

if data_source is None:
    st.markdown(
        '<h1 class="page-title">📊 Sales Intelligence Dashboard</h1>',
        unsafe_allow_html=True,
    )
    st.markdown(
        '<p class="page-subtitle">Alat bantu manajemen untuk meningkatkan omset, AO, customer aktif, '
        "dan mengidentifikasi peluang penjualan berbasis data.</p>",
        unsafe_allow_html=True,
    )
    st.info(
        "⬅️ Silakan upload file data transaksi (CSV / Excel) melalui sidebar, "
        "atau letakkan file di folder `data/` pada repository untuk pemuatan otomatis."
    )
    with st.expander("📋 Struktur Kolom yang Diharapkan", expanded=True):
        st.dataframe(
            pd.DataFrame(
                {
                    "Kolom": REQUIRED_COLS,
                    "Keterangan": [
                        "Kota Customer",
                        "Tanggal Faktur",
                        "Supplier",
                        "Nilai Penjualan",
                        "Cabang/Depo",
                        "Salesman",
                        "Bulan",
                        "Tanggal",
                        "Divisi",
                        "Area",
                        "Kategori SBD atau NON SBD",
                        "Kode Customer",
                        "Gabungan Customer + Supplier",
                        "Nama Telemarketing",
                        "Customer Baru (NOO/NPD)",
                    ],
                }
            ),
            use_container_width=True,
            hide_index=True,
        )
    st.stop()

df_raw = load_data(data_source)
missing = [c for c in REQUIRED_COLS if c not in df_raw.columns]
if missing:
    st.warning(
        f"⚠️ Kolom berikut tidak ditemukan di data Anda dan visual terkait akan dilewati: {', '.join(missing)}"
    )

st.sidebar.markdown("### 🔍 Filter Data")


def ms_filter(label, col):
    if col not in df_raw.columns:
        return []
    opts = sorted([o for o in df_raw[col].dropna().unique().tolist() if o != "nan"])
    return st.sidebar.multiselect(label, opts)


f_month = ms_filter("Bulan", "MONTH")
f_kota = ms_filter("Kota", "KOTA")
f_supp = ms_filter("Supplier", "SUPP")
f_sales = ms_filter("Sales", "SALES")
f_depo = ms_filter("Depo", "DEPO")
f_divisi = ms_filter("Divisi", "DIVISI")
f_area = ms_filter("Area", "AREA")
f_tele = ms_filter("Telemarketing", "TELE")

df = df_raw.copy()
for col, sel in [
    ("MONTH", f_month),
    ("KOTA", f_kota),
    ("SUPP", f_supp),
    ("SALES", f_sales),
    ("DEPO", f_depo),
    ("DIVISI", f_divisi),
    ("AREA", f_area),
    ("TELE", f_tele),
]:
    if sel and col in df.columns:
        df = df[df[col].isin(sel)]

if df.empty:
    st.warning(
        "Tidak ada data untuk kombinasi filter yang dipilih. Silakan ubah filter."
    )
    st.stop()

# Data target (opsional) — ikut disaring oleh filter Depo di sidebar agar konsisten
target_df = None
if target_source is not None:
    target_df = load_target_data(target_source)
    if f_depo and "DEPO" in target_df.columns:
        target_df = target_df[target_df["DEPO"].isin(f_depo)]

st.sidebar.markdown("---")
st.sidebar.markdown(f"**Baris data (setelah filter):** {len(df):,}")

st.sidebar.markdown("### 📑 Halaman")
page = st.sidebar.radio(
    "Pilih Halaman",
    [
        "1️⃣ Executive Dashboard",
        "2️⃣ Sales Performance",
        "3️⃣ Supplier Dashboard",
        "4️⃣ Customer Dashboard",
        "5️⃣ AO Dashboard",
        "6️⃣ NOO Dashboard",
        "7️⃣ Telemarketing Dashboard",
        "8️⃣ Area Dashboard",
        "9️⃣ Opportunity Dashboard",
        "🔟 KPI Efisiensi",
    ],
    label_visibility="collapsed",
)

# =============================================================================
# 1. EXECUTIVE DASHBOARD
# =============================================================================
if page.startswith("1️⃣"):
    page_header(
        "Executive Dashboard", "Ringkasan performa perusahaan secara keseluruhan"
    )

    omset = calc_omset(df)
    ao = calc_ao(df)
    cust_aktif = calc_customer_aktif(df)
    faktur = calc_faktur(df)
    noo = calc_noo(df)
    n_tele = df["TELE"].nunique() if has(df, "TELE") else np.nan

    kpi_row(
        [
            ("💰 Total Omset", fmt_rp(omset), None),
            ("🔁 Total AO", fmt_num(ao), None),
            ("👥 Customer Aktif", fmt_num(cust_aktif), None),
            ("🧾 Jumlah Faktur", fmt_num(faktur), None),
        ]
    )
    kpi_row(
        [
            (
                "📈 Avg Omset / Faktur",
                fmt_rp(omset / faktur if faktur else np.nan),
                None,
            ),
            ("📈 Avg Omset / AO", fmt_rp(omset / ao if ao else np.nan), None),
            ("🆕 Total NOO", fmt_num(noo), None),
            ("📞 Total Telemarketing", fmt_num(n_tele), None),
        ]
    )

    # =========================================================================
    # TARGET VS REALISASI (BULAN BERJALAN)
    # =========================================================================
    st.markdown("### 🎯 Target vs Realisasi (Bulan Berjalan)")
    if target_df is None:
        st.caption(
            "📌 Upload **Data Target** di sidebar (atau letakkan di data/) untuk menampilkan perbandingan target vs realisasi."
        )
    elif not has(df, "TGL FKTR") or not df["TGL FKTR"].notna().any():
        st.caption(
            "Kolom TGL FKTR tidak tersedia/tidak valid, target vs realisasi tidak bisa dihitung."
        )
    elif not has(target_df, "DEPO"):
        st.caption("Kolom DEPO tidak ditemukan di file Data Target.")
    else:
        max_date = df["TGL FKTR"].max()
        bulan_num = max_date.month
        bulan_col = MONTH_COLS[bulan_num - 1]
        bulan_nama_id = [
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
        ][bulan_num - 1]

        st.caption(
            f"📅 Update penjualan per **{max_date.strftime('%d %B %Y')}** — "
            f"target yang ditampilkan adalah target **bulan {bulan_nama_id} {max_date.year}** saja "
            f"(bukan akumulasi target 12 bulan)."
        )

        if bulan_col not in target_df.columns:
            st.caption(
                f"Kolom target untuk bulan {bulan_nama_id} ({bulan_col}) tidak ditemukan di file Data Target."
            )
        else:
            realisasi_bulan = df[
                (df["TGL FKTR"].dt.month == bulan_num)
                & (df["TGL FKTR"].dt.year == max_date.year)
            ]
            realisasi_by_depo = (
                realisasi_bulan.groupby("DEPO")["NOMINAL"]
                .sum()  # type: ignore[reportAttributeAccessIssue]
                .reset_index()
                .rename(columns={"NOMINAL": "REALISASI"})
            )
            target_by_depo = (
                target_df.groupby("DEPO")[bulan_col]
                .sum()  # type: ignore[reportAttributeAccessIssue]
                .reset_index()
                .rename(columns={bulan_col: "TARGET"})
            )
            tvr = target_by_depo.merge(
                realisasi_by_depo, on="DEPO", how="outer"
            ).fillna(0)
            tvr = tvr.sort_values("TARGET", ascending=False).reset_index(drop=True)
            tvr["ACHV_%"] = np.where(
                tvr["TARGET"] != 0, tvr["REALISASI"] / tvr["TARGET"] * 100, np.nan
            )

            total_target = tvr["TARGET"].sum()  # type: ignore[reportAttributeAccessIssue]
            total_realisasi = tvr["REALISASI"].sum()  # type: ignore[reportAttributeAccessIssue]
            total_achv = (
                total_realisasi / total_target * 100 if total_target else np.nan
            )

            kpi_row(
                [
                    (
                        f"🎯 Total Target ({bulan_nama_id})",
                        fmt_rp(total_target),
                        None,
                    ),
                    ("💰 Total Realisasi", fmt_rp(total_realisasi), None),
                    (
                        "📊 Pencapaian",
                        f"{total_achv:,.1f}%" if not pd.isna(total_achv) else "-",
                        None,
                    ),
                ]
            )

            fig_tvr = go.Figure()
            fig_tvr.add_trace(
                go.Bar(
                    x=tvr["DEPO"],
                    y=tvr["TARGET"],
                    name="Target",
                    marker_color=COLOR_PRIMARY,
                    text=tvr["TARGET"],
                    texttemplate="%{text:,.0f}",
                    textposition="outside",
                )
            )
            fig_tvr.add_trace(
                go.Bar(
                    x=tvr["DEPO"],
                    y=tvr["REALISASI"],
                    name="Realisasi",
                    marker_color=COLOR_ACCENT,
                    text=tvr["REALISASI"],
                    texttemplate="%{text:,.0f}",
                    textposition="outside",
                )
            )
            fig_tvr.update_layout(
                barmode="group",
                title=f"Target vs Realisasi per Depo — {bulan_nama_id} {max_date.year}",
                height=420,
                margin=dict(l=10, r=10, t=50, b=10),
                plot_bgcolor="white",
                title_font_color=COLOR_PRIMARY,
                legend=dict(
                    orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1
                ),
            )
            fig_tvr.update_yaxes(tickformat=",.0f")
            st.plotly_chart(fig_tvr, use_container_width=True)

            below = tvr[tvr["ACHV_%"] < 100].sort_values("ACHV_%")
            if not below.empty:
                worst = below.iloc[0]
                insight(
                    f"Depo **{worst['DEPO']}** pencapaian terendah bulan {bulan_nama_id}: "
                    f"{fmt_rp(worst['REALISASI'])} dari target {fmt_rp(worst['TARGET'])} "
                    f"({worst['ACHV_%']:.1f}%).",
                    "warn",
                )

            tvr_display = tvr.rename(
                columns={"DEPO": "Depo", "ACHV_%": "Pencapaian (%)"}
            )
            st.dataframe(
                style_df(tvr_display, decimals=0, pct_cols=["Pencapaian (%)"]),
                use_container_width=True,
                hide_index=True,
            )

    st.markdown("### 📈 Trend Omset Bulanan")
    tr = monthly_trend(df)
    if not tr.empty:
        fig = px.line(
            tr,
            x="PERIOD",
            y="OMSET",
            markers=True,
            color_discrete_sequence=[COLOR_PRIMARY],
        )
        fig.update_yaxes(tickformat=",.0f")
        fig.update_traces(hovertemplate="%{x}<br>Omset: %{y:,.0f}<extra></extra>")
        fig.update_layout(height=380, plot_bgcolor="white")
        st.plotly_chart(fig, use_container_width=True)
        g = growth_pct(tr["OMSET"])
        if not pd.isna(g):
            level = "good" if g > 0 else "bad"
            arrow = "naik" if g > 0 else "turun"
            insight(
                f"Omset periode terakhir {arrow} {abs(g):.1f}% dibanding periode sebelumnya.",
                level,
            )
    else:
        st.caption("Data tanggal/bulan tidak tersedia untuk trend.")

    c1, c2 = st.columns(2)
    with c1:
        s = dimension_summary(df, "KOTA")
        if not s.empty:
            st.plotly_chart(
                styled_bar(s, "KOTA", "OMSET", "Omset per Kota", COLOR_ACCENT),
                use_container_width=True,
            )
            top = s.iloc[0]
            insight(
                f"Kota **{top['KOTA']}** menyumbang {top['KONTRIBUSI_%']:.1f}% dari total omset perusahaan.",
                "info",
            )
    with c2:
        s = dimension_summary(df, "DEPO")
        if not s.empty:
            st.plotly_chart(
                styled_bar(s, "DEPO", "OMSET", "Omset per Depo", COLOR_PRIMARY),
                use_container_width=True,
            )

    s = dimension_summary(df, "DIVISI")
    if not s.empty:
        st.plotly_chart(
            styled_bar(s, "DIVISI", "OMSET", "Omset per Divisi", "#117864"),
            use_container_width=True,
        )

    s = dimension_summary(df, "SALES")
    if not s.empty:
        st.plotly_chart(
            styled_bar(s, "SALES", "OMSET", "Omset per Sales (Top 15)", COLOR_BAD),
            use_container_width=True,
        )
        bottom = s.tail(1).iloc[0] if len(s) > 1 else None
        if bottom is not None:
            insight(
                f"Sales dengan kontribusi omset terendah adalah **{bottom['SALES']}** ({fmt_rp(bottom['OMSET'])}). "
                f"Perlu evaluasi coverage area atau pendampingan.",
                "warn",
            )

# =============================================================================
# 2. SALES PERFORMANCE DASHBOARD
# =============================================================================
elif page.startswith("2️⃣"):
    page_header(
        "Sales Performance Dashboard", "Ranking dan analisis kuadran performa salesman"
    )

    s = dimension_summary(df, "SALES")
    if s.empty:
        st.warning("Kolom SALES tidak tersedia di data.")
    else:
        total_omset = s["OMSET"].sum()  # type: ignore[reportAttributeAccessIssue]
        kpi_row(
            [
                ("🏆 Jumlah Sales Aktif", fmt_num(len(s)), None),
                ("💰 Omset Tertinggi", fmt_rp(s["OMSET"].max()), s.iloc[0]["SALES"]),
                ("🔁 AO Tertinggi", fmt_num(s["AO"].max()), None),
                ("📈 Avg Omset/Sales", fmt_rp(total_omset / len(s)), None),
            ]
        )

        st.markdown("### 🎯 Target vs Realisasi per Salesman (Bulan Berjalan)")
        if target_df is None:
            st.caption(
                "📌 Upload **Data Target** di sidebar (atau letakkan di data/) untuk menampilkan perbandingan target vs realisasi per salesman."
            )
        elif not has(df, "TGL FKTR") or not df["TGL FKTR"].notna().any():
            st.caption(
                "Kolom TGL FKTR tidak tersedia/tidak valid, target vs realisasi tidak bisa dihitung."
            )
        elif not has(target_df, "NAMA SALESMAN"):
            st.caption("Kolom NAMA SALESMAN tidak ditemukan di file Data Target.")
        else:
            target_sales_df = target_df
            if f_sales:
                target_sales_df = target_sales_df[
                    target_sales_df["NAMA SALESMAN"].isin(f_sales)
                ]

            max_date_sm = df["TGL FKTR"].max()
            bulan_num_sm = max_date_sm.month
            bulan_col_sm = MONTH_COLS[bulan_num_sm - 1]
            bulan_nama_sm = [
                "Januari", "Februari", "Maret", "April", "Mei", "Juni",
                "Juli", "Agustus", "September", "Oktober", "November", "Desember",
            ][bulan_num_sm - 1]

            st.caption(
                f"📅 Update penjualan per **{max_date_sm.strftime('%d %B %Y')}** — "
                f"target yang ditampilkan adalah target **bulan {bulan_nama_sm} {max_date_sm.year}** saja "
                f"(bukan akumulasi target 12 bulan)."
            )

            if bulan_col_sm not in target_sales_df.columns:
                st.caption(
                    f"Kolom target untuk bulan {bulan_nama_sm} ({bulan_col_sm}) tidak ditemukan di file Data Target."
                )
            else:
                realisasi_bulan_sm = df[
                    (df["TGL FKTR"].dt.month == bulan_num_sm)
                    & (df["TGL FKTR"].dt.year == max_date_sm.year)
                ]
                realisasi_by_sales = (
                    realisasi_bulan_sm.groupby("SALES")["NOMINAL"]
                    .sum()  # type: ignore[reportAttributeAccessIssue]
                    .reset_index()
                    .rename(columns={"SALES": "SALESMAN", "NOMINAL": "REALISASI"})
                )
                target_by_sales = (
                    target_sales_df.groupby("NAMA SALESMAN")[bulan_col_sm]
                    .sum()  # type: ignore[reportAttributeAccessIssue]
                    .reset_index()
                    .rename(columns={"NAMA SALESMAN": "SALESMAN", bulan_col_sm: "TARGET"})
                )
                tvr_sm = target_by_sales.merge(
                    realisasi_by_sales, on="SALESMAN", how="outer"
                ).fillna(0)
                tvr_sm = tvr_sm.sort_values("TARGET", ascending=False).reset_index(
                    drop=True
                )
                tvr_sm["ACHV_%"] = np.where(
                    tvr_sm["TARGET"] != 0,
                    tvr_sm["REALISASI"] / tvr_sm["TARGET"] * 100,
                    np.nan,
                )

                total_target_sm = tvr_sm["TARGET"].sum()  # type: ignore[reportAttributeAccessIssue]
                total_realisasi_sm = tvr_sm["REALISASI"].sum()  # type: ignore[reportAttributeAccessIssue]
                total_achv_sm = (
                    total_realisasi_sm / total_target_sm * 100
                    if total_target_sm
                    else np.nan
                )

                kpi_row(
                    [
                        (
                            f"🎯 Total Target ({bulan_nama_sm})",
                            fmt_rp(total_target_sm),
                            None,
                        ),
                        ("💰 Total Realisasi", fmt_rp(total_realisasi_sm), None),
                        (
                            "📊 Pencapaian",
                            f"{total_achv_sm:,.1f}%"
                            if not pd.isna(total_achv_sm)
                            else "-",
                            None,
                        ),
                    ]
                )

                chart_height = max(420, 24 * len(tvr_sm))
                fig_tvr_sm = go.Figure()
                fig_tvr_sm.add_trace(
                    go.Bar(
                        y=tvr_sm["SALESMAN"],
                        x=tvr_sm["TARGET"],
                        name="Target",
                        orientation="h",
                        marker_color=COLOR_PRIMARY,
                    )
                )
                fig_tvr_sm.add_trace(
                    go.Bar(
                        y=tvr_sm["SALESMAN"],
                        x=tvr_sm["REALISASI"],
                        name="Realisasi",
                        orientation="h",
                        marker_color=COLOR_ACCENT,
                    )
                )
                fig_tvr_sm.update_layout(
                    barmode="group",
                    title=f"Target vs Realisasi per Salesman — {bulan_nama_sm} {max_date_sm.year}",
                    height=chart_height,
                    margin=dict(l=10, r=10, t=50, b=10),
                    plot_bgcolor="white",
                    title_font_color=COLOR_PRIMARY,
                    yaxis=dict(autorange="reversed"),
                    legend=dict(
                        orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1
                    ),
                )
                fig_tvr_sm.update_xaxes(tickformat=",.0f")
                st.plotly_chart(fig_tvr_sm, use_container_width=True)

                below_sm = tvr_sm[tvr_sm["ACHV_%"] < 100].sort_values("ACHV_%")
                if not below_sm.empty:
                    worst_sm = below_sm.iloc[0]
                    insight(
                        f"Salesman **{worst_sm['SALESMAN']}** pencapaian terendah bulan {bulan_nama_sm}: "
                        f"{fmt_rp(worst_sm['REALISASI'])} dari target {fmt_rp(worst_sm['TARGET'])} "
                        f"({worst_sm['ACHV_%']:.1f}%).",
                        "warn",
                    )

                tvr_sm_display = tvr_sm.rename(
                    columns={"SALESMAN": "Salesman", "ACHV_%": "Pencapaian (%)"}
                )
                st.dataframe(
                    style_df(
                        tvr_sm_display, decimals=0, pct_cols=["Pencapaian (%)"]
                    ),
                    use_container_width=True,
                    hide_index=True,
                )

        st.markdown("### 🏅 Ranking Sales")
        show = s.rename(
            columns={
                "SALES": "Sales",
                "OMSET": "Omset",
                "AO": "AO",
                "CUSTOMER_AKTIF": "Customer Aktif",
                "JUMLAH_FAKTUR": "Jml Faktur",
                "AVG_OMSET_PER_AO": "Avg Omset/AO",
                "AVG_OMSET_PER_CUSTOMER": "Avg Omset/Customer",
                "AVG_OMSET_PER_FAKTUR": "Avg Omset/Faktur",
                "KONTRIBUSI_%": "Kontribusi %",
            }
        )
        st.dataframe(
            style_df(show, decimals=0, pct_cols=["Kontribusi %"]),
            use_container_width=True,
            hide_index=True,
        )

        st.markdown("### 🎯 Quadrant Analysis — AO vs Omset")
        x_med, y_med = s["AO"].median(), s["OMSET"].median()
        s["QUADRANT"] = s.apply(
            lambda r: quadrant_label(r["AO"], r["OMSET"], x_med, y_med), axis=1
        )
        color_map = {
            "AO Tinggi - Omset Tinggi": COLOR_GOOD,
            "AO Tinggi - Omset Rendah": COLOR_BAD,
            "AO Rendah - Omset Tinggi": COLOR_ACCENT,
            "AO Rendah - Omset Rendah": COLOR_WARN,
        }
        fig = px.scatter(
            s,
            x="AO",
            y="OMSET",
            color="QUADRANT",
            text="SALES",
            color_discrete_map=color_map,
            size="CUSTOMER_AKTIF",
            hover_data=["JUMLAH_FAKTUR"],
        )
        fig.add_vline(x=x_med, line_dash="dash", line_color="gray")
        fig.add_hline(y=y_med, line_dash="dash", line_color="gray")
        fig.update_traces(textposition="top center")
        fig.update_xaxes(tickformat=",.0f")
        fig.update_yaxes(tickformat=",.0f")
        fig.update_layout(height=520, plot_bgcolor="white")
        st.plotly_chart(fig, use_container_width=True)

        cnt = s["QUADRANT"].value_counts()
        risky = s[s["QUADRANT"] == "AO Tinggi - Omset Rendah"].sort_values(
            "AO", ascending=False
        )
        if not risky.empty:
            names = ", ".join(risky["SALES"].head(3).tolist())
            insight(
                f"**{len(risky)} sales** memiliki AO tinggi tetapi omset rendah ({names}, dst) — "
                f"artinya sering transaksi namun nilai transaksi kecil. Perlu didorong untuk **upselling / cross-selling** "
                f"agar nilai per transaksi meningkat.",
                "warn",
            )
        strong = s[s["QUADRANT"] == "AO Tinggi - Omset Tinggi"].sort_values(
            "OMSET", ascending=False
        )
        if not strong.empty:
            insight(
                f"Sales terbaik (AO tinggi & omset tinggi): **{', '.join(strong['SALES'].head(3).tolist())}**. "
                f"Jadikan sebagai benchmark best-practice bagi sales lain.",
                "good",
            )
        few_cust_high_omset = s[
            (s["OMSET"] >= y_med) & (s["CUSTOMER_AKTIF"] < s["CUSTOMER_AKTIF"].median())
        ]
        if not few_cust_high_omset.empty:
            insight(
                f"Sales dengan omset tinggi tapi customer sedikit: "
                f"**{', '.join(few_cust_high_omset['SALES'].head(3).tolist())}** — omset terkonsentrasi pada sedikit customer, "
                f"berisiko tinggi jika salah satu customer besar hilang. Perlu ekspansi jumlah customer.",
                "warn",
            )

# =============================================================================
# 3. SUPPLIER DASHBOARD
# =============================================================================
elif page.startswith("3️⃣"):
    page_header("Supplier Dashboard", "Performa dan pertumbuhan tiap supplier")

    s = dimension_summary(df, "SUPP")
    if s.empty:
        st.warning("Kolom SUPP tidak tersedia di data.")
    else:
        kpi_row(
            [
                ("🏭 Jumlah Supplier", fmt_num(len(s)), None),
                ("💰 Omset Tertinggi", fmt_rp(s["OMSET"].max()), s.iloc[0]["SUPP"]),
                ("🔁 Total AO", fmt_num(s["AO"].sum()), None),  # type: ignore[reportAttributeAccessIssue]
                ("📈 Avg Omset/Supplier", fmt_rp(s["OMSET"].mean()), None),
            ]
        )

        c1, c2 = st.columns(2)
        with c1:
            st.plotly_chart(
                styled_bar(
                    s, "SUPP", "OMSET", "Ranking Omset Supplier (Top 15)", COLOR_PRIMARY
                ),
                use_container_width=True,
            )
        with c2:
            st.plotly_chart(
                styled_bar(s, "SUPP", "AO", "AO per Supplier (Top 15)", COLOR_ACCENT),
                use_container_width=True,
            )

        st.markdown("### 📊 Supplier Growth (Bulan Terakhir vs Sebelumnya)")
        if has(df, "TGL FKTR") and df["TGL FKTR"].notna().any():
            tmp = df.dropna(subset=["TGL FKTR"]).copy()
            tmp["PERIOD"] = tmp["TGL FKTR"].dt.to_period("M").astype(str)
            periods = sorted(tmp["PERIOD"].unique())
            if len(periods) >= 2:
                last, prev = periods[-1], periods[-2]
                piv = (
                    tmp[tmp["PERIOD"].isin([prev, last])]
                    .groupby(["SUPP", "PERIOD"])["NOMINAL"]
                    .sum()  # type: ignore[reportAttributeAccessIssue]
                    .unstack(fill_value=0)
                )
                piv["GROWTH_%"] = np.where(
                    piv[prev] != 0, (piv[last] - piv[prev]) / piv[prev] * 100, np.nan
                )
                piv = piv.sort_values("GROWTH_%")
                st.dataframe(
                    style_df(
                        piv.rename(
                            columns={prev: f"Omset {prev}", last: f"Omset {last}"}
                        ),
                        decimals=1,
                        pct_cols=["GROWTH_%"],
                    ),
                    use_container_width=True,
                )
                decline = (
                    piv.dropna(subset=["GROWTH_%"]).sort_values("GROWTH_%").head(3)
                )
                for supp_name, row in decline.iterrows():
                    if row["GROWTH_%"] < 0:
                        insight(
                            f"Supplier **{supp_name}** mengalami penurunan {abs(row['GROWTH_%']):.1f}% "
                            f"dibanding bulan sebelumnya.",
                            "bad",
                        )
            else:
                st.caption("Data belum cukup periode untuk menghitung growth.")
        else:
            st.caption("Kolom TGL FKTR tidak tersedia untuk menghitung growth bulanan.")

        st.markdown("### 💡 Insight Supplier")
        ao_med, omset_med = s["AO"].median(), s["OMSET"].median()
        push = s[(s["AO"] >= ao_med) & (s["OMSET"] < omset_med)].sort_values(
            "AO", ascending=False
        )
        if not push.empty:
            top3 = ", ".join(push["SUPP"].head(3).tolist())
            insight(
                f"Supplier dengan AO tinggi tetapi omset rendah (peluang upselling besar): **{top3}**.",
                "warn",
            )
        productive = s.sort_values("AVG_OMSET_PER_AO", ascending=False).head(3)
        insight(
            f"Supplier paling produktif (omset per AO tertinggi): **{', '.join(productive['SUPP'].tolist())}**.",
            "good",
        )

# =============================================================================
# 4. CUSTOMER DASHBOARD
# =============================================================================
elif page.startswith("4️⃣"):
    page_header("Customer Dashboard", "Distribusi, Pareto, dan segmentasi ABC customer")

    if not has(df, "KD GRUP"):
        st.warning("Kolom KD GRUP tidak tersedia di data.")
    else:
        n_cust = df["KD GRUP"].nunique()
        kpi_row(
            [
                ("👥 Distinct Customer (KD GRUP)", fmt_num(n_cust), None),
                (
                    "💰 Omset per Customer (Avg)",
                    fmt_rp(calc_omset(df) / n_cust if n_cust else np.nan),
                    None,
                ),
                (
                    "🧾 Faktur per Customer (Avg)",
                    f"{calc_faktur(df)/n_cust:.1f}" if n_cust else "-",
                    None,
                ),
                ("🆕 Total NOO", fmt_num(calc_noo(df)), None),
            ]
        )

        cust_omset = (
            df.groupby("KD GRUP")["NOMINAL"]
            .sum()  # type: ignore[reportAttributeAccessIssue]
            .reset_index()
            .sort_values("NOMINAL", ascending=False)
        )
        c1, c2 = st.columns(2)
        with c1:
            st.markdown("#### 🥇 Top 15 Customer")
            st.plotly_chart(
                styled_bar(
                    cust_omset.rename(columns={"NOMINAL": "OMSET"}),
                    "KD GRUP",
                    "OMSET",
                    "Top 15 Customer berdasarkan Omset",
                    COLOR_GOOD,
                ),
                use_container_width=True,
            )
        with c2:
            if has(df, "KOTA"):
                s = (
                    df.groupby("KOTA")["KD GRUP"]
                    .nunique()
                    .reset_index(name="JUMLAH_CUSTOMER")
                    .sort_values("JUMLAH_CUSTOMER", ascending=False)
                )
                st.plotly_chart(
                    styled_bar(
                        s, "KOTA", "JUMLAH_CUSTOMER", "Customer per Kota", COLOR_ACCENT
                    ),
                    use_container_width=True,
                )

        c3, c4 = st.columns(2)
        with c3:
            if has(df, "SALES"):
                s = (
                    df.groupby("SALES")["KD GRUP"]
                    .nunique()
                    .reset_index(name="JUMLAH_CUSTOMER")
                    .sort_values("JUMLAH_CUSTOMER", ascending=False)
                )
                st.plotly_chart(
                    styled_bar(
                        s,
                        "SALES",
                        "JUMLAH_CUSTOMER",
                        "Customer per Sales",
                        COLOR_PRIMARY,
                    ),
                    use_container_width=True,
                )
        with c4:
            if has(df, "SUPP"):
                s = (
                    df.groupby("SUPP")["KD GRUP"]
                    .nunique()
                    .reset_index(name="JUMLAH_CUSTOMER")
                    .sort_values("JUMLAH_CUSTOMER", ascending=False)
                )
                st.plotly_chart(
                    styled_bar(
                        s,
                        "SUPP",
                        "JUMLAH_CUSTOMER",
                        "Customer per Supplier",
                        COLOR_WARN,
                    ),
                    use_container_width=True,
                )

        st.markdown("### 📈 Pareto & Segmentasi ABC Customer")
        abc = pareto_abc(df)
        if not abc.empty:
            fig = go.Figure()
            top_n = min(50, len(abc))
            plot_d = abc.head(top_n)
            fig.add_trace(
                go.Bar(
                    x=plot_d["KD GRUP"],
                    y=plot_d["NOMINAL"],
                    name="Omset",
                    marker_color=COLOR_PRIMARY,
                    hovertemplate="%{x}<br>Omset: %{y:,.0f}<extra></extra>",
                )
            )
            fig.add_trace(
                go.Scatter(
                    x=plot_d["KD GRUP"],
                    y=plot_d["CUM_PCT"],
                    name="Kumulatif %",
                    yaxis="y2",
                    line=dict(color=COLOR_BAD),
                    hovertemplate="%{x}<br>Kumulatif: %{y:,.1f}%<extra></extra>",
                )
            )
            fig.update_layout(
                title=f"Diagram Pareto Customer (Top {top_n})",
                yaxis=dict(title="Omset", tickformat=",.0f"),
                yaxis2=dict(
                    title="Kumulatif %", overlaying="y", side="right", range=[0, 100]
                ),
                height=450,
                plot_bgcolor="white",
            )
            st.plotly_chart(fig, use_container_width=True)

            abc_count = abc["ABC"].value_counts().reindex(["A", "B", "C"]).fillna(0)
            abc_omset = (
                abc.groupby("ABC")["NOMINAL"].sum().reindex(["A", "B", "C"]).fillna(0)  # type: ignore[reportAttributeAccessIssue]
            )
            cc1, cc2, cc3 = st.columns(3)
            cc1.metric(
                "Kategori A (kontributor utama)",
                f"{int(abc_count['A'])} customer",
                fmt_rp(abc_omset["A"]),
            )
            cc2.metric(
                "Kategori B", f"{int(abc_count['B'])} customer", fmt_rp(abc_omset["B"])
            )
            cc3.metric(
                "Kategori C (long-tail)",
                f"{int(abc_count['C'])} customer",
                fmt_rp(abc_omset["C"]),
            )

            insight(
                f"**{int(abc_count['A'])} customer (kategori A)** menyumbang 80% dari total omset. "
                f"Fokuskan program retensi dan layanan prioritas pada kelompok ini.",
                "good",
            )
            insight(
                f"**{int(abc_count['C'])} customer (kategori C)** hanya menyumbang sisa omset kecil (long-tail) — "
                f"berpotensi untuk digarap lebih intensif atau dikonsolidasikan ke telemarketing.",
                "warn",
            )

            top1 = abc.iloc[0]
            insight(
                f"Customer terbesar adalah **{top1['KD GRUP']}** dengan omset {fmt_rp(top1['NOMINAL'])} "
                f"({top1['CUM_PCT']:.1f}% kumulatif).",
                "info",
            )

# =============================================================================
# 5. AO DASHBOARD
# =============================================================================
elif page.startswith("5️⃣"):
    page_header("AO Dashboard", "AO dihitung menggunakan Distinct Count KD GRUP")

    if not has(df, "KD GRUP"):
        st.warning("Kolom KD GRUP tidak tersedia di data.")
    else:
        total_ao = calc_ao(df)
        kpi_row(
            [
                ("🔁 Total AO", fmt_num(total_ao), None),
                (
                    "💰 Avg Omset per AO",
                    fmt_rp(calc_omset(df) / total_ao if total_ao else np.nan),
                    None,
                ),
                (
                    "👥 AO per Customer Aktif",
                    (
                        f"{total_ao/calc_customer_aktif(df):.2f}"
                        if calc_customer_aktif(df)
                        else "-"
                    ),
                    None,
                ),
                (
                    "🧾 AO per Faktur",
                    f"{total_ao/calc_faktur(df):.2f}" if calc_faktur(df) else "-",
                    None,
                ),
            ]
        )

        dims = [
            ("SALES", "AO per Sales"),
            ("SUPP", "AO per Supplier"),
            ("KOTA", "AO per Kota"),
            ("DEPO", "AO per Depo"),
            ("DIVISI", "AO per Divisi"),
            ("AREA", "AO per Area"),
        ]
        cols = st.columns(2)
        for i, (dim, title) in enumerate(dims):
            if has(df, dim):
                s = (
                    df.groupby(dim)["KD GRUP"]
                    .nunique()
                    .reset_index(name="AO")
                    .sort_values("AO", ascending=False)
                )
                with cols[i % 2]:
                    st.plotly_chart(
                        styled_bar(s, dim, "AO", title, PALETTE[i % len(PALETTE)]),
                        use_container_width=True,
                    )

        st.markdown("### 📈 Trend AO Bulanan")
        tr = monthly_trend(df)
        if not tr.empty and "AO" in tr.columns:
            fig = px.line(
                tr,
                x="PERIOD",
                y="AO",
                markers=True,
                color_discrete_sequence=[COLOR_ACCENT],
            )
            fig.update_yaxes(tickformat=",.0f")
            fig.update_traces(hovertemplate="%{x}<br>AO: %{y:,.0f}<extra></extra>")
            fig.update_layout(height=380, plot_bgcolor="white", title="Trend AO")
            st.plotly_chart(fig, use_container_width=True)
            g = growth_pct(tr["AO"])
            if not pd.isna(g):
                level = "good" if g > 0 else "bad"
                insight(
                    f"AO periode terakhir {'naik' if g>0 else 'turun'} {abs(g):.1f}% dibanding periode sebelumnya.",
                    level,
                )

# =============================================================================
# 6. NOO DASHBOARD
# =============================================================================
elif page.startswith("6️⃣"):
    page_header(
        "NOO Dashboard",
        "Customer baru (New Outlet/New Order) berdasarkan kolom NOO/NPD",
    )

    if not has(df, "NOO/NPD"):
        st.warning("Kolom NOO/NPD tidak tersedia di data.")
    else:
        new_mask = is_new_customer_mask(df)
        new_df = df[new_mask]
        total_noo = calc_noo(df)
        kpi_row(
            [
                ("🆕 Total Customer Baru (NOO/NPD)", fmt_num(total_noo), None),
                (
                    "👥 % dari Customer Aktif",
                    (
                        fmt_pct(total_noo / calc_customer_aktif(df) * 100)
                        if calc_customer_aktif(df)
                        else "-"
                    ),
                    None,
                ),
                ("💰 Omset dari Customer Baru", fmt_rp(new_df["NOMINAL"].sum()), None),  # type: ignore[reportAttributeAccessIssue]
                (
                    "🏭 Supplier Terlibat",
                    fmt_num(new_df["SUPP"].nunique()) if has(df, "SUPP") else "-",
                    None,
                ),
            ]
        )

        noo_only = df.loc[
            new_mask & df["NOO/NPD"].str.upper().str.contains("NOO", na=False),
            "KD GRUP",
        ].nunique()
        npd_only = df.loc[
            new_mask & df["NOO/NPD"].str.upper().str.contains("NPD", na=False),
            "KD GRUP",
        ].nunique()
        cbrk1, cbrk2 = st.columns(2)
        cbrk1.metric("Customer baru — tanda NOO", fmt_num(noo_only))
        cbrk2.metric("Customer baru — tanda NPD", fmt_num(npd_only))

        cols = st.columns(2)
        dims = [
            ("SALES", "NOO per Sales"),
            ("SUPP", "NOO per Supplier"),
            ("KOTA", "NOO per Kota"),
            ("DEPO", "NOO per Depo"),
        ]
        for i, (dim, title) in enumerate(dims):
            s = noo_summary(df, dim)
            if not s.empty:
                with cols[i % 2]:
                    st.plotly_chart(
                        styled_bar(
                            s, dim, "NOO_COUNT", title, PALETTE[i % len(PALETTE)]
                        ),
                        use_container_width=True,
                    )

        if has(df, "MONTH"):
            s = noo_summary(df, "MONTH")
            if not s.empty:
                st.markdown("### 📅 NOO per Bulan")
                noo_fig = px.bar(
                    s,
                    x="MONTH",
                    y="NOO_COUNT",
                    color_discrete_sequence=[COLOR_GOOD],
                    title="NOO per Bulan",
                    text="NOO_COUNT",
                )
                noo_fig.update_traces(
                    texttemplate="%{text:,.0f}",
                    textposition="outside",
                    hovertemplate="%{x}<br>NOO: %{y:,.0f}<extra></extra>",
                )
                noo_fig.update_yaxes(tickformat=",.0f")
                noo_fig.update_layout(height=380, plot_bgcolor="white")
                st.plotly_chart(noo_fig, use_container_width=True)

        st.markdown("### 🏅 Ranking Sales berdasarkan NOO")
        s = noo_summary(df, "SALES")
        if not s.empty:
            st.dataframe(
                style_df(
                    s.rename(columns={"SALES": "Sales", "NOO_COUNT": "Jumlah NOO"}),
                    decimals=0,
                ),
                use_container_width=True,
                hide_index=True,
            )
            insight(
                f"Sales dengan akuisisi customer baru terbanyak: **{s.iloc[0]['SALES']}** "
                f"({int(s.iloc[0]['NOO_COUNT'])} customer baru). Jadikan role model untuk strategi akuisisi.",
                "good",
            )

# =============================================================================
# 7. TELEMARKETING DASHBOARD
# =============================================================================
elif page.startswith("7️⃣"):
    page_header("Telemarketing Dashboard", "Performa tim telemarketing")

    if not has(df, "TELE"):
        st.warning("Kolom TELE tidak tersedia di data.")
    else:
        s = dimension_summary(df, "TELE")
        noo_s = noo_summary(df, "TELE")
        kpi_row(
            [
                ("📞 Jumlah Telemarketing", fmt_num(df["TELE"].nunique()), None),
                ("💰 Omset via Tele", fmt_rp(s["OMSET"].sum()), None),  # type: ignore[reportAttributeAccessIssue]
                ("🔁 Total AO via Tele", fmt_num(s["AO"].sum()), None),  # type: ignore[reportAttributeAccessIssue]
                (
                    "🆕 Total NOO via Tele",
                    fmt_num(noo_s["NOO_COUNT"].sum() if not noo_s.empty else 0),  # type: ignore[reportAttributeAccessIssue]
                    None,
                ),
            ]
        )

        c1, c2 = st.columns(2)
        with c1:
            st.plotly_chart(
                styled_bar(
                    s, "TELE", "OMSET", "Omset per Telemarketing", COLOR_PRIMARY
                ),
                use_container_width=True,
            )
        with c2:
            st.plotly_chart(
                styled_bar(s, "TELE", "AO", "AO per Telemarketing", COLOR_ACCENT),
                use_container_width=True,
            )

        st.markdown("### 🏅 Ranking & Kontribusi Telemarketing")
        show = s[
            ["TELE", "OMSET", "AO", "CUSTOMER_AKTIF", "JUMLAH_FAKTUR", "KONTRIBUSI_%"]
        ].rename(
            columns={
                "TELE": "Telemarketing",
                "OMSET": "Omset",
                "CUSTOMER_AKTIF": "Customer Aktif",
                "JUMLAH_FAKTUR": "Jml Faktur",
                "KONTRIBUSI_%": "Kontribusi %",
            }
        )
        st.dataframe(
            style_df(show, decimals=0, pct_cols=["Kontribusi %"]),
            use_container_width=True,
            hide_index=True,
        )
        if not s.empty:
            top = s.iloc[0]
            insight(
                f"Telemarketing **{top['TELE']}** memberikan kontribusi omset tertinggi sebesar {top['KONTRIBUSI_%']:.1f}%.",
                "good",
            )

# =============================================================================
# 8. AREA DASHBOARD
# =============================================================================
elif page.startswith("8️⃣"):
    page_header(
        "Area Dashboard",
        "Performa berdasarkan wilayah geografis dan struktur organisasi",
    )

    tabs = st.tabs(["Kota", "Area", "Divisi", "Depo"])
    for tab, dim in zip(tabs, ["KOTA", "AREA", "DIVISI", "DEPO"]):
        with tab:
            s = dimension_summary(df, dim)
            if s.empty:
                st.caption(f"Kolom {dim} tidak tersedia.")
                continue
            c1, c2 = st.columns(2)
            with c1:
                st.plotly_chart(
                    styled_bar(
                        s, dim, "OMSET", f"Omset per {dim.title()}", COLOR_PRIMARY
                    ),
                    use_container_width=True,
                )
            with c2:
                st.plotly_chart(
                    styled_bar(s, dim, "AO", f"AO per {dim.title()}", COLOR_ACCENT),
                    use_container_width=True,
                )
            st.dataframe(
                style_df(
                    s[
                        [
                            dim,
                            "OMSET",
                            "AO",
                            "CUSTOMER_AKTIF",
                            "JUMLAH_FAKTUR",
                            "KONTRIBUSI_%",
                        ]
                    ],
                    decimals=0,
                    pct_cols=["KONTRIBUSI_%"],
                ),
                use_container_width=True,
                hide_index=True,
            )

# =============================================================================
# 9. OPPORTUNITY DASHBOARD (paling penting)
# =============================================================================
elif page.startswith("9️⃣"):
    page_header(
        "Opportunity Dashboard",
        "🔎 Peluang otomatis untuk meningkatkan omset — halaman paling penting",
    )

    # --- Sales AO tinggi omset rendah ---
    s_sales = dimension_summary(df, "SALES")
    if not s_sales.empty:
        ao_med, om_med = s_sales["AO"].median(), s_sales["OMSET"].median()
        st.markdown("### 1️⃣ Sales: AO Tinggi tetapi Omset Rendah")
        r = s_sales[
            (s_sales["AO"] >= ao_med) & (s_sales["OMSET"] < om_med)
        ].sort_values("AO", ascending=False)
        st.dataframe(
            style_df(r[["SALES", "AO", "OMSET", "AVG_OMSET_PER_AO"]]),
            use_container_width=True,
            hide_index=True,
        )
        if not r.empty:
            insight(
                f"{len(r)} sales perlu didorong meningkatkan **nilai transaksi per kunjungan** (upselling), "
                f"bukan menambah frekuensi kunjungan.",
                "warn",
            )

        st.markdown("### 2️⃣ Sales: Omset Tinggi tetapi Customer Sedikit")
        cust_med = s_sales["CUSTOMER_AKTIF"].median()
        r2 = s_sales[
            (s_sales["OMSET"] >= om_med) & (s_sales["CUSTOMER_AKTIF"] < cust_med)
        ].sort_values("OMSET", ascending=False)
        st.dataframe(
            style_df(
                r2[["SALES", "OMSET", "CUSTOMER_AKTIF", "AVG_OMSET_PER_CUSTOMER"]]
            ),
            use_container_width=True,
            hide_index=True,
        )
        if not r2.empty:
            insight(
                f"{len(r2)} sales bergantung pada sedikit customer besar — risiko konsentrasi. "
                f"Perlu program akuisisi customer baru (NOO).",
                "bad",
            )

    # --- Supplier AO tinggi omset kecil ---
    s_supp = dimension_summary(df, "SUPP")
    if not s_supp.empty:
        st.markdown("### 3️⃣ Supplier: AO Tinggi tetapi Omset Kecil")
        ao_med_s, om_med_s = s_supp["AO"].median(), s_supp["OMSET"].median()
        r3 = s_supp[
            (s_supp["AO"] >= ao_med_s) & (s_supp["OMSET"] < om_med_s)
        ].sort_values("AO", ascending=False)
        st.dataframe(
            style_df(r3[["SUPP", "AO", "OMSET", "AVG_OMSET_PER_AO"]]),
            use_container_width=True,
            hide_index=True,
        )

        st.markdown("### 4️⃣ Supplier: Customer Banyak tetapi Kontribusi Kecil")
        cust_med_s = s_supp["CUSTOMER_AKTIF"].median()
        r4 = s_supp[
            (s_supp["CUSTOMER_AKTIF"] >= cust_med_s)
            & (s_supp["KONTRIBUSI_%"] < s_supp["KONTRIBUSI_%"].median())
        ].sort_values("CUSTOMER_AKTIF", ascending=False)
        st.dataframe(
            style_df(
                r4[["SUPP", "CUSTOMER_AKTIF", "OMSET", "KONTRIBUSI_%"]],
                pct_cols=["KONTRIBUSI_%"],
            ),
            use_container_width=True,
            hide_index=True,
        )
        if not r4.empty:
            insight(
                f"Supplier **{', '.join(r4['SUPP'].head(3).tolist())}** punya banyak customer namun kontribusi omset kecil "
                f"— peluang besar untuk program bundling/upselling.",
                "warn",
            )

    # --- Kota customer banyak omset kecil ---
    s_kota = dimension_summary(df, "KOTA")
    if not s_kota.empty:
        st.markdown("### 5️⃣ Kota: Customer Banyak tetapi Omset Kecil")
        cust_med_k, om_med_k = (
            s_kota["CUSTOMER_AKTIF"].median(),
            s_kota["OMSET"].median(),
        )
        r5 = s_kota[
            (s_kota["CUSTOMER_AKTIF"] >= cust_med_k) & (s_kota["OMSET"] < om_med_k)
        ].sort_values("CUSTOMER_AKTIF", ascending=False)
        st.dataframe(
            style_df(r5[["KOTA", "CUSTOMER_AKTIF", "OMSET", "AVG_OMSET_PER_CUSTOMER"]]),
            use_container_width=True,
            hide_index=True,
        )
        if not r5.empty:
            insight(
                f"Kota **{', '.join(r5['KOTA'].head(3).tolist())}** memiliki potensi besar — banyak customer tetapi "
                f"rata-rata omset per customer masih rendah.",
                "warn",
            )

    # --- Depo AO tinggi omset rendah ---
    s_depo = dimension_summary(df, "DEPO")
    if not s_depo.empty:
        st.markdown("### 6️⃣ Depo/Cabang: AO Tinggi tetapi Omset Rendah")
        ao_med_d, om_med_d = s_depo["AO"].median(), s_depo["OMSET"].median()
        r6 = s_depo[
            (s_depo["AO"] >= ao_med_d) & (s_depo["OMSET"] < om_med_d)
        ].sort_values("AO", ascending=False)
        st.dataframe(
            style_df(r6[["DEPO", "AO", "OMSET", "AVG_OMSET_PER_AO"]]),
            use_container_width=True,
            hide_index=True,
        )
        if not r6.empty:
            insight(
                f"Depo **{', '.join(r6['DEPO'].head(3).tolist())}** perlu perhatian khusus manajemen cabang.",
                "bad",
            )

    # --- Area growth terendah ---
    if has(df, "AREA") and has(df, "TGL FKTR") and df["TGL FKTR"].notna().any():
        st.markdown("### 7️⃣ Area dengan Pertumbuhan Paling Rendah")
        tmp = df.dropna(subset=["TGL FKTR"]).copy()
        tmp["PERIOD"] = tmp["TGL FKTR"].dt.to_period("M").astype(str)
        periods = sorted(tmp["PERIOD"].unique())
        if len(periods) >= 2:
            last, prev = periods[-1], periods[-2]
            piv = (
                tmp[tmp["PERIOD"].isin([prev, last])]
                .groupby(["AREA", "PERIOD"])["NOMINAL"]
                .sum()  # type: ignore[reportAttributeAccessIssue]
                .unstack(fill_value=0)
            )
            piv["GROWTH_%"] = np.where(
                piv[prev] != 0, (piv[last] - piv[prev]) / piv[prev] * 100, np.nan
            )
            piv = piv.sort_values("GROWTH_%")
            st.dataframe(
                style_df(piv, decimals=1, pct_cols=["GROWTH_%"]),
                use_container_width=True,
            )
            worst = piv.dropna(subset=["GROWTH_%"]).head(1)
            if not worst.empty:
                insight(
                    f"Area dengan pertumbuhan terendah: **{worst.index[0]}** "
                    f"({worst['GROWTH_%'].iloc[0]:.1f}%). Perlu investigasi penyebab penurunan.",
                    "bad",
                )
        else:
            st.caption("Data belum cukup periode untuk menghitung growth area.")

    # --- Top/Bottom 20 customer ---
    if has(df, "KD GRUP"):
        st.markdown("### 8️⃣–9️⃣ Top 20 & Bottom 20 Customer")
        cust_omset = (
            df.groupby("KD GRUP")["NOMINAL"]
            .sum()  # type: ignore[reportAttributeAccessIssue]
            .reset_index()
            .sort_values("NOMINAL", ascending=False)
        )
        c1, c2 = st.columns(2)
        with c1:
            st.markdown("**Top 20 Customer**")
            st.dataframe(
                style_df(cust_omset.head(20)), use_container_width=True, hide_index=True
            )
        with c2:
            st.markdown("**Bottom 20 Customer**")
            st.dataframe(
                style_df(cust_omset.tail(20)), use_container_width=True, hide_index=True
            )

    # --- Cross selling ---
    if has(df, "KD GRUP") and has(df, "SUPP"):
        st.markdown("### 🔟 Cross Selling Opportunity")
        supp_count = (
            df.groupby("KD GRUP")["SUPP"].nunique().reset_index(name="JUMLAH_SUPPLIER")
        )
        if supp_count.empty:
            st.caption("Tidak ada data customer/supplier untuk dianalisis pada filter saat ini.")
        else:
            single = supp_count[supp_count["JUMLAH_SUPPLIER"] == 1]
            multi = supp_count[supp_count["JUMLAH_SUPPLIER"] > 1]
            c1, c2 = st.columns(2)
            c1.metric("Customer hanya 1 Supplier (target cross-sell)", fmt_num(len(single)))
            c2.metric("Customer dengan multi-Supplier", fmt_num(len(multi)))
            max_supplier = supp_count["JUMLAH_SUPPLIER"].max()
            nbins = int(max_supplier) if pd.notna(max_supplier) and max_supplier > 0 else 1
            fig = px.histogram(
                supp_count,
                x="JUMLAH_SUPPLIER",
                nbins=nbins,
                title="Distribusi Jumlah Supplier per Customer",
                color_discrete_sequence=[COLOR_ACCENT],
            )
            fig.update_traces(
                hovertemplate="Jumlah Supplier: %{x}<br>Jumlah Customer: %{y:,.0f}<extra></extra>"
            )
            fig.update_yaxes(tickformat=",.0f")
            fig.update_layout(height=380, plot_bgcolor="white")
            st.plotly_chart(fig, use_container_width=True)
            pct_single = len(single) / len(supp_count) * 100 if len(supp_count) else 0
            insight(
                f"**{len(single)} customer ({pct_single:.1f}%)** hanya membeli dari 1 supplier — "
                f"berpotensi menjadi target utama program **cross-selling** ke supplier/merek lain.",
                "warn",
            )
            top_single = single.merge(
                df.groupby("KD GRUP")["NOMINAL"].sum().reset_index(), on="KD GRUP"  # type: ignore[reportAttributeAccessIssue]
            )
            top_single = top_single.sort_values("NOMINAL", ascending=False).head(15)
            st.markdown(
                "**Top Customer Single-Supplier (prioritas cross-sell, diurutkan omset)**"
            )
            st.dataframe(style_df(top_single), use_container_width=True, hide_index=True)

# =============================================================================
# 10. KPI EFISIENSI
# =============================================================================
elif page.startswith("🔟"):
    page_header("KPI Efisiensi", "Rasio efisiensi operasional penjualan")

    omset = calc_omset(df)
    ao = calc_ao(df)
    cust = calc_customer_aktif(df)
    faktur = calc_faktur(df)
    n_sales = df["SALES"].nunique() if has(df, "SALES") else np.nan
    n_supp = df["SUPP"].nunique() if has(df, "SUPP") else np.nan
    n_kota = df["KOTA"].nunique() if has(df, "KOTA") else np.nan
    n_depo = df["DEPO"].nunique() if has(df, "DEPO") else np.nan

    kpi_row(
        [
            ("Omset / AO", fmt_rp(omset / ao if ao else np.nan), None),
            ("Omset / Customer", fmt_rp(omset / cust if cust else np.nan), None),
            ("Omset / Faktur", fmt_rp(omset / faktur if faktur else np.nan), None),
            ("AO / Sales", f"{ao/n_sales:.1f}" if n_sales else "-", None),
        ]
    )
    kpi_row(
        [
            ("Customer / Sales", f"{cust/n_sales:.1f}" if n_sales else "-", None),
            ("Omset / Supplier", fmt_rp(omset / n_supp if n_supp else np.nan), None),
            ("Omset / Kota", fmt_rp(omset / n_kota if n_kota else np.nan), None),
            ("Omset / Depo", fmt_rp(omset / n_depo if n_depo else np.nan), None),
        ]
    )

    st.markdown("### 📊 Efisiensi per Sales (diurutkan Avg Omset/Customer)")
    s = dimension_summary(df, "SALES")
    if not s.empty:
        eff = s[
            [
                "SALES",
                "AVG_OMSET_PER_AO",
                "AVG_OMSET_PER_CUSTOMER",
                "AVG_OMSET_PER_FAKTUR",
            ]
        ].sort_values("AVG_OMSET_PER_CUSTOMER", ascending=False)
        st.dataframe(style_df(eff), use_container_width=True, hide_index=True)
        low_eff = eff.tail(3)
        insight(
            f"Sales dengan efisiensi omset per customer terendah: **{', '.join(low_eff['SALES'].tolist())}** — "
            f"perlu pendampingan teknik penjualan / bundling produk untuk meningkatkan rata-rata transaksi.",
            "warn",
        )

    st.markdown("### 🏭 Efisiensi per Supplier")
    s2 = dimension_summary(df, "SUPP")
    if not s2.empty:
        eff2 = s2[["SUPP", "AVG_OMSET_PER_AO", "AVG_OMSET_PER_CUSTOMER"]].sort_values(
            "AVG_OMSET_PER_AO", ascending=False
        )
        st.dataframe(style_df(eff2), use_container_width=True, hide_index=True)

st.sidebar.markdown("---")
st.sidebar.caption(
    "💡 Dashboard ini menghasilkan insight otomatis berbasis analisis kuadran, "
    "Pareto, dan growth month-over-month dari data yang diupload."
)
