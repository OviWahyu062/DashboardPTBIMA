import os
import base64
import sqlite3
from datetime import datetime
from io import BytesIO
from zoneinfo import ZoneInfo

import pandas as pd
import plotly.graph_objects as go
import streamlit as st


# ==============================
# KONFIGURASI DASAR
# ==============================

DB_PATH = "po_database.db"
LOGO_PATH = "logo_pelindo.png"

RAW_TABLE = "po_raw"
PROCESSED_TABLE = "po_processed"
UPLOAD_TABLE = "uploaded_files"

TIMEZONE = ZoneInfo("Asia/Jakarta")


# ==============================
# KONFIGURASI IDENTIFIKASI EXCEL
# ==============================

# Tidak menggunakan nama sheet tetap lagi.
# Sistem mencari berdasarkan nama kolom.

PO_REQUIRED_IDENTIFIER = [
    "Purchasing Document"
]


KK_REQUIRED_IDENTIFIER = [
    "Purchase Order",
    "Lama Proses PO"
]


REQUIRED_COLUMNS = [
    "Purchasing Document",
    "Item",
    "History",
    "Document Date",
    "Supplier",
    "Supplier Name PO",
    "Material PO",
    "Material Long Text",
    "Short Text (PO)",
    "Header Text PO",
    "SAP CS",
    "Acct Assignment Category PO",
    "Plant",
    "Net Order Value",
    "Net Order Price",
    "Order Quantity",
    "Still to be delivered (QTY)",
    "Still to be delivered (Value)",
    "Still to be invoiced (QTY)",
    "Still to be invoiced (Value)",
    "Purchase Requisition",
    "Requisitioner",
    "Material Group",
    "Valuation Price",
]


NUMERIC_COLUMNS = [
    "Net Order Value",
    "Net Order Price",
    "Order Quantity",
    "Still to be delivered (QTY)",
    "Still to be delivered (Value)",
    "Still to be invoiced (QTY)",
    "Still to be invoiced (Value)",
    "Valuation Price",
]

# ==============================
# SETTING HALAMAN
# ==============================

st.set_page_config(
    page_title="Dashboard Pengolahan PO",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed",
)


# ==============================
# CUSTOM CSS CORPORATE PELINDO
# ==============================

st.markdown(
    """
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');

        html, body, [class*="css"] {
            font-family: 'Inter', sans-serif;
        }

        .stApp {
            background:
                radial-gradient(circle at top right, rgba(0, 120, 184, 0.10), transparent 30%),
                linear-gradient(180deg, #F7FBFE 0%, #F3F7FB 100%);
        }

        .block-container {
            padding-top: 1.2rem;
            padding-left: 1.2rem;
            padding-right: 1.4rem;
            padding-bottom: 2rem;
            max-width: 1700px;
        }

        header[data-testid="stHeader"] {
            background: transparent;
            height: 0rem;
        }

        div[data-testid="stToolbar"],
        div[data-testid="stDecoration"],
        div[data-testid="collapsedControl"],
        button[data-testid="baseButton-header"],
        [data-testid="stSidebarCollapseButton"],
        section[data-testid="stSidebar"] {
            display: none !important;
            visibility: hidden !important;
            opacity: 0 !important;
            width: 0px !important;
            min-width: 0px !important;
            max-width: 0px !important;
        }

        #MainMenu {
            visibility: hidden;
        }

        footer {
            visibility: hidden;
        }

        .sidebar-logo-card {
            background: linear-gradient(180deg, #FFFFFF 0%, #F7FBFE 100%);
            border: 1px solid #DDEAF4;
            border-radius: 28px;
            padding: 26px 24px;
            margin-bottom: 18px;
            box-shadow: 0 18px 40px rgba(0, 72, 135, 0.12);
            text-align: center;
        }

        .sidebar-logo-card img {
            width: 100%;
            max-width: 220px;
            display: block;
            margin: auto;
        }

        .fallback-logo {
            text-align: center;
            padding: 8px 0;
        }

        .fallback-logo .pelindo {
            font-size: 31px;
            font-weight: 800;
            letter-spacing: 1px;
            color: #0078B8;
        }

        .fallback-logo .jasa {
            font-size: 13px;
            font-weight: 800;
            color: #16A7A0;
            margin-top: -4px;
        }

        .fallback-logo .bima {
            font-size: 12px;
            font-weight: 800;
            color: #E53935;
            margin-top: 2px;
        }

        .sidebar-nav-title-card {
            background: linear-gradient(180deg, #EAF6FC 0%, #D7EEF9 100%);
            border: 1px solid #B9D9EF;
            border-radius: 22px;
            padding: 16px 16px;
            margin-bottom: 14px;
            box-shadow: 0 12px 26px rgba(0, 120, 184, 0.12);
        }

        .sidebar-title {
            color: #092D55;
            font-size: 12px;
            font-weight: 800;
            letter-spacing: 0.8px;
            text-transform: uppercase;
        }

        div[role="radiogroup"] label {
            background: rgba(255, 255, 255, 0.96);
            border: 1px solid rgba(0, 120, 184, 0.18);
            padding: 14px 15px;
            border-radius: 18px;
            margin-bottom: 12px;
            color: #093057;
            font-weight: 700;
            font-size: 14px;
            transition: all 0.18s ease-in-out;
            box-shadow: 0 8px 18px rgba(0, 72, 135, 0.07);
        }

        div[role="radiogroup"] label:hover {
            background: #FFFFFF;
            border-color: #0078B8;
            transform: translateX(3px);
        }

        div[role="radiogroup"] label:has(input:checked) {
            background: linear-gradient(135deg, #0078B8, #005B9E);
            color: #FFFFFF !important;
            border-color: #0078B8;
            box-shadow: 0 14px 26px rgba(0, 120, 184, 0.26);
        }

        .sidebar-info {
            padding: 20px;
            border-radius: 24px;
            color: #FFFFFF;
            background:
                radial-gradient(circle at top left, rgba(255,255,255,0.32), rgba(255,255,255,0.06) 35%, transparent 52%),
                linear-gradient(135deg, #0089CE, #005B9E 68%, #003B70);
            box-shadow: 0 18px 40px rgba(0, 64, 116, 0.30);
            margin-top: 20px;
        }

        .sidebar-info-title {
            font-size: 16px;
            font-weight: 800;
            margin-bottom: 8px;
        }

        .sidebar-info-text {
            font-size: 12px;
            line-height: 1.6;
            opacity: 0.95;
        }

        .sidebar-user {
            margin-top: 18px;
            padding: 14px;
            border-radius: 18px;
            background: rgba(255,255,255,0.16);
            border: 1px solid rgba(255,255,255,0.25);
            color: #FFFFFF;
        }

        .sidebar-user-main {
            font-size: 13px;
            font-weight: 800;
        }

        .sidebar-user-sub {
            font-size: 11px;
            opacity: 0.9;
            margin-top: 2px;
        }

        .page-header {
            background:
                linear-gradient(135deg, rgba(255,255,255,0.98), rgba(255,255,255,0.92)),
                radial-gradient(circle at right, rgba(0, 120, 184, 0.11), transparent 45%);
            border: 1px solid #E1EBF5;
            border-radius: 26px;
            padding: 26px 30px;
            margin-bottom: 20px;
            box-shadow: 0 14px 34px rgba(0, 58, 112, 0.08);
            display: flex;
            align-items: center;
            justify-content: space-between;
            gap: 18px;
        }

        .page-title {
            color: #082E5F;
            font-size: 31px;
            font-weight: 800;
            line-height: 1.2;
            letter-spacing: -0.5px;
            margin: 0;
        }

        .page-subtitle {
            color: #637D96;
            font-size: 14px;
            font-weight: 500;
            margin-top: 8px;
            line-height: 1.5;
        }

        .header-actions {
            display: flex;
            justify-content: flex-end;
            gap: 10px;
            flex-wrap: wrap;
        }

        .header-badge {
            background: #F4F9FD;
            color: #092D55;
            border: 1px solid #D7E8F5;
            border-radius: 15px;
            padding: 10px 15px;
            font-size: 13px;
            font-weight: 700;
            white-space: nowrap;
        }

        .hero-card {
            background:
                linear-gradient(135deg, rgba(0, 91, 158, 0.98), rgba(0, 137, 206, 0.92)),
                radial-gradient(circle at right, rgba(255,255,255,0.22), transparent 45%);
            border-radius: 25px;
            padding: 24px 26px;
            color: #FFFFFF;
            margin-bottom: 20px;
            box-shadow: 0 18px 38px rgba(0, 83, 143, 0.25);
            border: 1px solid rgba(255,255,255,0.20);
        }

        .hero-title {
            font-size: 22px;
            font-weight: 800;
            margin-bottom: 6px;
        }

        .hero-text {
            font-size: 13px;
            opacity: 0.95;
            line-height: 1.6;
            max-width: 1100px;
        }

        .metric-card {
            background: #FFFFFF;
            border: 1px solid #E1EBF5;
            border-radius: 22px;
            padding: 18px 18px;
            min-height: 150px;
            box-shadow: 0 13px 30px rgba(0, 65, 120, 0.07);
            position: relative;
            overflow: hidden;
        }

        .metric-card::after {
            content: "";
            position: absolute;
            width: 120px;
            height: 120px;
            right: -50px;
            top: -50px;
            background: rgba(0, 120, 184, 0.08);
            border-radius: 999px;
            z-index: 0;
        }

        .metric-row {
            display: flex;
            align-items: flex-start;
            gap: 12px;
            position: relative;
            z-index: 2;
            width: 100%;
        }

        .metric-icon {
            width: 52px;
            height: 52px;
            border-radius: 17px;
            display: flex;
            align-items: center;
            justify-content: center;
            color: #FFFFFF;
            font-size: 24px;
            background: linear-gradient(135deg, #0078B8, #005B9E);
            box-shadow: 0 10px 24px rgba(0, 120, 184, 0.25);
            flex-shrink: 0;
        }

        .metric-icon.teal {
            background: linear-gradient(135deg, #16A7A0, #0B8D91);
        }

        .metric-icon.green {
            background: linear-gradient(135deg, #1EAD6F, #0F8F61);
        }

        .metric-icon.purple {
            background: linear-gradient(135deg, #4869E8, #2545B8);
        }

        .metric-content {
            min-width: 0;
            width: 100%;
            flex: 1;
        }

        .metric-label {
            color: #637D96;
            font-size: 12px;
            font-weight: 800;
            margin-bottom: 5px;
            line-height: 1.35;
        }

        .metric-value {
            color: #092D55;
            font-size: clamp(17px, 1.15vw, 23px);
            font-weight: 800;
            line-height: 1.22;
            max-width: 100%;
            white-space: normal;
            overflow-wrap: anywhere;
            word-break: break-word;
        }

        .metric-caption {
            margin-top: 16px;
            padding-top: 12px;
            border-top: 1px solid #EEF4F9;
            color: #6A829A;
            font-size: 12px;
            font-weight: 600;
            position: relative;
            z-index: 2;
            line-height: 1.45;
        }

        .section-title {
            color: #092D55;
            font-size: 19px;
            font-weight: 800;
            margin-bottom: 4px;
        }

        .section-subtitle {
            color: #6B7F95;
            font-size: 12px;
            font-weight: 500;
            margin-bottom: 14px;
        }

        div[data-testid="stVerticalBlockBorderWrapper"] {
            background: #FFFFFF;
            border: 1px solid #E1EBF5;
            border-radius: 22px;
            box-shadow: 0 12px 28px rgba(0, 65, 120, 0.06);
            padding: 18px;
        }

        .chart-corporate-title {
            background:
                linear-gradient(135deg, #FFFFFF, #F3FAFE),
                radial-gradient(circle at right, rgba(0, 120, 184, 0.12), transparent 40%);
            border: 1px solid #D8EAF5;
            border-left: 6px solid #0078B8;
            border-radius: 24px;
            padding: 18px 22px;
            color: #092D55;
            margin-bottom: 14px;
            box-shadow: 0 14px 30px rgba(0, 72, 135, 0.10);
        }

        .chart-corporate-title h3 {
            margin: 0;
            font-size: 22px;
            font-weight: 900;
            color: #092D55;
            letter-spacing: 0.3px;
        }

        .chart-corporate-title p {
            margin: 5px 0 0 0;
            font-size: 13px;
            color: #637D96;
        }

        .info-box {
            background: linear-gradient(135deg, #EAF6FC, #FFFFFF);
            border: 1px solid #CFE8F7;
            border-left: 5px solid #0078B8;
            border-radius: 17px;
            padding: 16px 18px;
            color: #092D55;
            font-size: 14px;
            margin-bottom: 18px;
            line-height: 1.6;
        }

        .success-box {
            background: linear-gradient(135deg, #E8F9F2, #FFFFFF);
            border: 1px solid #C8F0DC;
            border-left: 5px solid #16A86B;
            border-radius: 17px;
            padding: 16px 18px;
            color: #075239;
            font-size: 14px;
            margin-bottom: 18px;
            line-height: 1.6;
        }

        .warning-box {
            background: linear-gradient(135deg, #FFF8E7, #FFFFFF);
            border: 1px solid #F5DFAB;
            border-left: 5px solid #F3A712;
            border-radius: 17px;
            padding: 16px 18px;
            color: #704D00;
            font-size: 14px;
            margin-bottom: 18px;
            line-height: 1.6;
        }

        .danger-box {
            background: linear-gradient(135deg, #FFF0F0, #FFFFFF);
            border: 1px solid #FFD0D0;
            border-left: 5px solid #D93025;
            border-radius: 17px;
            padding: 16px 18px;
            color: #7A1515;
            font-size: 14px;
            margin-bottom: 18px;
            line-height: 1.6;
        }

        .stButton > button {
            background: linear-gradient(135deg, #0078B8, #005B9E);
            color: #FFFFFF;
            border: none;
            border-radius: 14px;
            padding: 0.75rem 1.15rem;
            font-weight: 800;
            box-shadow: 0 10px 22px rgba(0, 120, 184, 0.24);
        }

        .stButton > button:hover {
            background: linear-gradient(135deg, #006DA8, #004F8C);
            color: #FFFFFF;
            border: none;
        }

        .stDownloadButton > button {
            background: #FFFFFF;
            color: #005B9E;
            border: 1px solid #B9D9EF;
            border-radius: 14px;
            padding: 0.75rem 1.15rem;
            font-weight: 800;
        }

        .stDownloadButton > button:hover {
            background: #EAF6FC;
            color: #004F8C;
            border-color: #0078B8;
        }

        div[data-testid="stFileUploader"] {
            background: #FFFFFF;
            border: 2px dashed #B9D9EF;
            border-radius: 20px;
            padding: 18px;
        }

        div[data-testid="stDataFrame"] {
            border-radius: 17px;
            overflow: hidden;
        }

        .footer-main {
            background: linear-gradient(135deg, #003B70, #0078B8);
            color: #FFFFFF;
            padding: 15px 20px;
            border-radius: 20px;
            margin-top: 24px;
            display: flex;
            justify-content: space-between;
            align-items: center;
            font-size: 12px;
            box-shadow: 0 12px 26px rgba(0, 63, 115, 0.22);
        }

        .footer-main strong {
            font-weight: 800;
        }

        @media (max-width: 900px) {
            .page-header {
                flex-direction: column;
                align-items: flex-start;
            }

            .header-actions {
                justify-content: flex-start;
            }

            .footer-main {
                flex-direction: column;
                align-items: flex-start;
                gap: 6px;
            }
        }
    </style>
    """,
    unsafe_allow_html=True,
)


# ==============================
# FUNGSI WAKTU WIB
# ==============================

def get_now_wib():
    return datetime.now(TIMEZONE)


def get_now_wib_str():
    return get_now_wib().strftime("%Y-%m-%d %H:%M:%S WIB")


def get_now_wib_display():
    return get_now_wib().strftime("%d %b %Y %H:%M WIB")


# ==============================
# FUNGSI UMUM
# ==============================

def image_to_base64(path):
    if not os.path.exists(path):
        return None

    with open(path, "rb") as image_file:
        encoded = base64.b64encode(image_file.read()).decode()

    return encoded


def get_month_label(month_key):
    bulan = {
        "01": "Januari",
        "02": "Februari",
        "03": "Maret",
        "04": "April",
        "05": "Mei",
        "06": "Juni",
        "07": "Juli",
        "08": "Agustus",
        "09": "September",
        "10": "Oktober",
        "11": "November",
        "12": "Desember",
    }

    try:
        tahun, bulan_angka = month_key.split("-")
        return f"{bulan.get(bulan_angka, bulan_angka)} {tahun}"
    except Exception:
        return str(month_key)


def get_month_key_from_date(date_value):
    return date_value.strftime("%Y-%m")


def drop_empty_uploaded_rows(df):
    """
    Menghapus baris kosong dari Excel.
    Baris dianggap kosong jika:
    1. semua kolom penting kosong, atau
    2. Purchasing Document kosong.
    """
    if df is None or df.empty:
        return pd.DataFrame() if df is None else df

    df = df.copy()

    df = df.replace(r"^\s*$", pd.NA, regex=True)
    df = df.replace(["nan", "NaN", "None", "NONE", "NaT", "<NA>"], pd.NA)

    check_cols = [col for col in REQUIRED_COLUMNS if col in df.columns]

    if check_cols:
        df = df.dropna(how="all", subset=check_cols)
    else:
        df = df.dropna(how="all")

    if "Purchasing Document" in df.columns:
        key = df["Purchasing Document"].astype(str).str.strip()
        invalid_key = (
            key.eq("")
            | key.str.lower().isin(["nan", "none", "nat", "<na>"])
        )
        df = df[~invalid_key]

    df = df.reset_index(drop=True)

    return df


def drop_empty_kk_rows(df):
    if df is None or df.empty:
        return pd.DataFrame() if df is None else df

    df = df.copy()
    df = df.replace(r"^\s*$", pd.NA, regex=True)
    df = df.dropna(how="all")

    if "Purchase Order" in df.columns:
        key = df["Purchase Order"].astype(str).str.strip()
        invalid_key = (
            key.eq("")
            | key.str.lower().isin(["nan", "none", "nat", "<na>"])
        )
        df = df[~invalid_key]

    df = df.reset_index(drop=True)
    return df


# ==============================
# DATABASE
# ==============================

def get_connection():
    return sqlite3.connect(DB_PATH, check_same_thread=False)


def table_exists(conn, table_name):
    cursor = conn.cursor()
    cursor.execute(
        "SELECT name FROM sqlite_master WHERE type='table' AND name=?",
        (table_name,)
    )
    return cursor.fetchone() is not None


def get_table_columns(conn, table_name):
    if not table_exists(conn, table_name):
        return []

    cursor = conn.cursor()
    cursor.execute(f"PRAGMA table_info({table_name})")
    rows = cursor.fetchall()
    return [row[1] for row in rows]


def ensure_columns(conn, table_name, df):
    if not table_exists(conn, table_name):
        return

    existing_columns = get_table_columns(conn, table_name)
    cursor = conn.cursor()

    for col in df.columns:
        if col not in existing_columns:
            safe_col = col.replace('"', '""')
            cursor.execute(f'ALTER TABLE {table_name} ADD COLUMN "{safe_col}" TEXT')

    conn.commit()


def init_database():
    conn = get_connection()
    cursor = conn.cursor()

    cursor.execute(
        f"""
        CREATE TABLE IF NOT EXISTS {UPLOAD_TABLE} (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            file_name TEXT,
            upload_date TEXT,
            total_rows INTEGER,
            status TEXT,
            periode_data TEXT,
            periode_label TEXT
        )
        """
    )

    upload_columns = get_table_columns(conn, UPLOAD_TABLE)

    if "periode_data" not in upload_columns:
        cursor.execute(f"ALTER TABLE {UPLOAD_TABLE} ADD COLUMN periode_data TEXT")

    if "periode_label" not in upload_columns:
        cursor.execute(f"ALTER TABLE {UPLOAD_TABLE} ADD COLUMN periode_label TEXT")

    conn.commit()
    conn.close()


def save_upload_history(file_name, total_rows, periode_data, periode_label, status="Berhasil diproses"):
    conn = get_connection()
    cursor = conn.cursor()

    cursor.execute(
        f"""
        INSERT INTO {UPLOAD_TABLE}
        (file_name, upload_date, total_rows, status, periode_data, periode_label)
        VALUES (?, ?, ?, ?, ?, ?)
        """,
        (
            file_name,
            get_now_wib_str(),
            int(total_rows),
            status,
            periode_data,
            periode_label,
        ),
    )

    conn.commit()
    conn.close()


def save_dataframe_to_db(df_raw, df_processed, mode="append"):
    conn = get_connection()

    ensure_columns(conn, RAW_TABLE, df_raw)
    ensure_columns(conn, PROCESSED_TABLE, df_processed)

    df_raw.to_sql(RAW_TABLE, conn, if_exists=mode, index=False)
    df_processed.to_sql(PROCESSED_TABLE, conn, if_exists=mode, index=False)

    conn.close()


def read_table(table_name):
    conn = get_connection()

    try:
        df = pd.read_sql_query(f"SELECT * FROM {table_name}", conn)
    except Exception:
        df = pd.DataFrame()

    conn.close()

    if table_name in [RAW_TABLE, PROCESSED_TABLE]:
        df = drop_empty_uploaded_rows(df)

    return df


def get_available_months():
    df = read_table(PROCESSED_TABLE)

    if df.empty or "Periode Data" not in df.columns:
        return []

    months = (
        df["Periode Data"]
        .dropna()
        .astype(str)
        .replace("", pd.NA)
        .dropna()
        .unique()
        .tolist()
    )

    months = sorted(months, reverse=True)
    return months


def delete_data_by_month(month_key):
    conn = get_connection()
    cursor = conn.cursor()

    for table_name in [RAW_TABLE, PROCESSED_TABLE]:
        if table_exists(conn, table_name):
            columns = get_table_columns(conn, table_name)

            if "Periode Data" in columns:
                cursor.execute(
                    f'DELETE FROM {table_name} WHERE "Periode Data" = ?',
                    (month_key,)
                )

    if table_exists(conn, UPLOAD_TABLE):
        upload_columns = get_table_columns(conn, UPLOAD_TABLE)

        if "periode_data" in upload_columns:
            cursor.execute(
                f"DELETE FROM {UPLOAD_TABLE} WHERE periode_data = ?",
                (month_key,)
            )

    conn.commit()
    conn.close()


def filter_df_by_month(df, month_key):
    if df.empty:
        return df

    if month_key == "ALL":
        return df

    if "Periode Data" not in df.columns:
        return df.iloc[0:0]

    return df[df["Periode Data"].astype(str) == str(month_key)].copy()


def prepare_month_metadata(df, month_key, month_label, uploaded_file_name):
    df = df.copy()
    df["Periode Data"] = month_key
    df["Periode Label"] = month_label
    df["Nama File Upload"] = uploaded_file_name
    df["Tanggal Upload"] = get_now_wib_str()
    return df


# ==============================
# PENGOLAHAN DATA EXCEL
# ==============================


def find_po_sheet(xls):
    """
    Mencari sheet PO utama berdasarkan kolom
    Purchasing Document
    """

    for sheet in xls.sheet_names:

        try:
            df_temp = pd.read_excel(
                xls,
                sheet_name=sheet,
                nrows=5
            )

            columns = [
                str(col).strip()
                for col in df_temp.columns
            ]

            if "Purchasing Document" in columns:
                return sheet

        except Exception:
            continue

    return None



def find_kk_sheet(xls):
    """
    Mencari sheet KK Pemaketan berdasarkan:
    Purchase Order
    Lama Proses PO
    """

    for sheet in xls.sheet_names:

        try:
            df_temp = pd.read_excel(
                xls,
                sheet_name=sheet,
                nrows=5
            )

            columns = [
                str(col).strip()
                for col in df_temp.columns
            ]

            if (
                "Purchase Order" in columns
                and
                "Lama Proses PO" in columns
            ):
                return sheet

        except Exception:
            continue

    return None



def validate_columns(df):

    missing_columns = [
        col
        for col in REQUIRED_COLUMNS
        if col not in df.columns
    ]

    return missing_columns



def clean_numeric_columns(df):

    for col in NUMERIC_COLUMNS:

        if col in df.columns:

            df[col] = pd.to_numeric(
                df[col],
                errors="coerce"
            ).fillna(0)

    return df



def process_po_data(df_po, df_kk=None):

    df = df_po.copy()

    df = drop_empty_uploaded_rows(df)

    df = clean_numeric_columns(df)


    # ==============================
    # HITUNG TOTAL VALUE
    # ==============================

    df["Total Valuation Price"] = (
        df["Valuation Price"]
        *
        df["Order Quantity"]
    )


    header = (
        df["Header Text PO"]
        .fillna("")
        .astype(str)
    )


    df["PIR"] = contains_text(
        header,
        "PIR"
    )


    df["PENGADAAN LANGSUNG PENYELENGGARA"] = contains_text(
        header,
        "PENGADAAN LANGSUNG PENYELENGGARA"
    )


    df["PENGADAAN LANGSUNG NON PENYELENGGARA"] = contains_text(
        header,
        "PENGADAAN LANGSUNG NON PENYELENGGARA"
    )


    df["PENUNJUKAN LANGSUNG"] = contains_text(
        header,
        "PENUNJUKAN LANGSUNG"
    )


    df["tender"] = contains_text(
        header,
        "tender"
    )



    def get_status_final(row):

        requisitioner = (
            str(row.get("Requisitioner",""))
            .strip()
            .upper()
        )


        if requisitioner == "BIMA001":
            return "PENUNJUKAN LANGSUNG"


        elif row["PIR"]:
            return "PIR"


        elif row["PENGADAAN LANGSUNG PENYELENGGARA"]:
            return "PENGADAAN LANGSUNG PENYELENGGARA"


        elif row["PENGADAAN LANGSUNG NON PENYELENGGARA"]:
            return "PENGADAAN LANGSUNG NON PENYELENGGARA"


        elif row["PENUNJUKAN LANGSUNG"]:
            return "PENUNJUKAN LANGSUNG"


        elif row["tender"]:
            return "Tender Terbatas"


        else:
            return "Tidak Teridentifikasi"



    df["Status Final"] = df.apply(
        get_status_final,
        axis=1
    )



    def get_prj(row):

        header_text = (
            str(row.get("Header Text PO",""))
            .upper()
        )

        requisitioner = (
            str(row.get("Requisitioner",""))
            .strip()
            .upper()
        )


        if "PRJ" in header_text:
            return "Project"


        elif requisitioner == "BIMA001":
            return "Inbound"


        else:
            return "Pemeliharaan"



    df["PRJ"] = df.apply(
        get_prj,
        axis=1
    )



    # ==============================
    # EFISIENSI
    # ==============================

    df["Efisiensi"] = df.apply(
        lambda row:
        0
        if row["Valuation Price"] == 0
        else
        (
            row["Net Order Price"]
            -
            row["Valuation Price"]
        )
        *
        row["Order Quantity"],
        axis=1
    )



    df["Prosentase"] = df.apply(
        lambda row:
        0
        if (
            row["Valuation Price"] == 0
            or
            row["Net Order Price"] == 0
        )
        else
        (
            row["Net Order Price"]
            -
            row["Valuation Price"]
        )
        /
        row["Net Order Price"],
        axis=1
    )



    # ==============================
    # REVISI UTAMA
    # MENGAMBIL LAMA PROSES PO
    # ==============================

    df["Lama Proses PO"] = 0


    if (
        df_kk is not None
        and not df_kk.empty
    ):


        if (
            "Purchase Order" in df_kk.columns
            and
            "Lama Proses PO" in df_kk.columns
        ):


            kk_lookup = df_kk[
                [
                    "Purchase Order",
                    "Lama Proses PO"
                ]
            ].copy()



            # samakan format nomor PO

            kk_lookup["Purchase Order"] = (
                kk_lookup["Purchase Order"]
                .astype(str)
                .str.replace(
                    ".0",
                    "",
                    regex=False
                )
                .str.strip()
            )



            kk_lookup["Lama Proses PO"] = pd.to_numeric(
                kk_lookup["Lama Proses PO"],
                errors="coerce"
            )



            # mengikuti fungsi MATCH Excel
            # ambil data pertama

            kk_lookup = (
                kk_lookup
                .drop_duplicates(
                    subset=[
                        "Purchase Order"
                    ],
                    keep="first"
                )
            )



            df["_PO_KEY"] = (
                df["Purchasing Document"]
                .astype(str)
                .str.replace(
                    ".0",
                    "",
                    regex=False
                )
                .str.strip()
            )



            df = df.merge(
                kk_lookup,
                how="left",
                left_on="_PO_KEY",
                right_on="Purchase Order"
            )



            df["Lama Proses PO"] = pd.to_numeric(
                df["Lama Proses PO_y"],
                errors="coerce"
            ).fillna(0)



            df = df.drop(
                columns=[
                    "_PO_KEY",
                    "Purchase Order",
                    "Lama Proses PO_y"
                ],
                errors="ignore"
            )



            df = df.rename(
                columns={
                    "Lama Proses PO_x":
                    "Lama Proses PO"
                }
            )


    return drop_empty_uploaded_rows(df)



def read_excel_file(uploaded_file):

    xls = pd.ExcelFile(
        uploaded_file,
        engine="openpyxl"
    )


    po_sheet = find_po_sheet(xls)

    kk_sheet = find_kk_sheet(xls)



    if po_sheet is None:

        raise Exception(
            "Sheet PO tidak ditemukan. Pastikan ada kolom Purchasing Document."
        )


    df_po = pd.read_excel(
        xls,
        sheet_name=po_sheet
    )



    if kk_sheet is not None:

        df_kk = pd.read_excel(
            xls,
            sheet_name=kk_sheet
        )

    else:

        df_kk = pd.DataFrame()



    df_po = drop_empty_uploaded_rows(
        df_po
    )


    df_kk = drop_empty_kk_rows(
        df_kk
    )


    return (
        df_po,
        df_kk,
        xls.sheet_names
    )

# ==============================
# REKAP DASHBOARD
# ==============================

def make_rekap(df):

    df = drop_empty_uploaded_rows(df)


    if df.empty:

        return (
            pd.DataFrame(),
            pd.DataFrame(),
            pd.DataFrame()
        )


    # Pastikan Lama Proses PO selalu numeric

    if "Lama Proses PO" in df.columns:
        avg_lama_proses = (
            pd.to_numeric(
                df["Lama Proses PO"],
                errors="coerce"
            )
            .fillna(0)
            .mean()
        )
    else:
        avg_lama_proses = 0

    # ==============================
    # REKAP JUMLAH PAKET PO
    # ==============================

    paket_po = (
        df.groupby(
            "Status Final",
            dropna=False
        )
        [
            "Purchasing Document"
        ]
        .nunique()
        .reset_index(
            name="Jumlah Paket PO"
        )
        .sort_values(
            "Jumlah Paket PO",
            ascending=False
        )
    )



    # ==============================
    # REKAP EFISIENSI
    # ==============================

    efisiensi = (

        df.groupby(
            "Status Final",
            dropna=False
        )
        .agg(

            Total_Net_Order_Value=
            (
                "Net Order Value",
                "sum"
            ),

            Total_Valuation_Price=
            (
                "Total Valuation Price",
                "sum"
            ),

            Total_Efisiensi=
            (
                "Efisiensi",
                "sum"
            ),

            Rata_Rata_Prosentase=
            (
                "Prosentase",
                "mean"
            ),

        )

        .reset_index()

    )



    # ==============================
    # REKAP LAMA PROSES PO
    # ==============================

    lama_proses = (

        df.groupby(
            "Status Final",
            dropna=False
        )

        .agg(

            Rata_Rata_Lama_Proses_PO=
            (
                "Lama Proses PO",
                "mean"
            ),

            Maksimal_Lama_Proses_PO=
            (
                "Lama Proses PO",
                "max"
            ),

            Jumlah_Data=
            (
                "Purchasing Document",
                "count"
            ),

        )

        .reset_index()

    )



    return (
        paket_po,
        efisiensi,
        lama_proses
    )

# ==============================
# KOMPONEN UI
# ==============================

def render_custom_sidebar():
    logo_base64 = image_to_base64(LOGO_PATH)

    if logo_base64:
        st.markdown(
            f"""
            <div class="sidebar-logo-card">
                <img src="data:image/png;base64,{logo_base64}">
            </div>
            """,
            unsafe_allow_html=True,
        )
    else:
        st.markdown(
            """
            <div class="sidebar-logo-card">
                <div class="fallback-logo">
                    <div class="pelindo">PELINDO</div>
                    <div class="jasa">JASA MARITIM</div>
                    <div class="bima">BIMA</div>
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    st.markdown(
        """
        <div class="sidebar-nav-title-card">
            <div class="sidebar-title">Navigasi Sistem</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    menu_choice = st.radio(
        "Menu",
        [
            "Dashboard",
            "Entry Data Excel",
            "Hasil Pengolahan Data",
            "Database & Download",
        ],
        label_visibility="collapsed",
    )

    st.markdown(
        """
        <div class="sidebar-info">
            <div class="sidebar-info-title">Procurement Dashboard</div>
            <div class="sidebar-info-text">
                Sistem pengolahan data Purchase Order berbasis Excel,
                SQLite, dan visualisasi dashboard internal.
            </div>
            <div class="sidebar-user">
                <div class="sidebar-user-main">👤 Procurement Team</div>
                <div class="sidebar-user-sub">Internal User</div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    return menu_choice


def render_header(title, subtitle):
    today = get_now_wib_display()

    st.markdown(
        f"""
        <div class="page-header">
            <div>
                <h1 class="page-title">{title}</h1>
                <div class="page-subtitle">{subtitle}</div>
            </div>
            <div class="header-actions">
                <div class="header-badge">📅 {today}</div>
                <div class="header-badge">⚓ Pelindo Jasa Maritim - BIMA</div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_hero():
    st.markdown(
        """
        <div class="hero-card">
            <div class="hero-title">Dashboard Monitoring Pengadaan</div>
            <div class="hero-text">
                Sistem ini digunakan untuk mengunggah data Purchase Order dari Excel,
                mengolah data secara otomatis, menyimpan hasilnya ke database SQLite,
                serta menampilkan ringkasan Paket PO, Efisiensi, dan Lama Proses PO.
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_metric_card(title, value, caption, icon, color_class=""):
    st.markdown(
        f"""
        <div class="metric-card">
            <div class="metric-row">
                <div class="metric-icon {color_class}">{icon}</div>
                <div class="metric-content">
                    <div class="metric-label">{title}</div>
                    <div class="metric-value">{value}</div>
                </div>
            </div>
            <div class="metric-caption">{caption}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_section_title(title, subtitle):
    st.markdown(
        f"""
        <div class="section-title">{title}</div>
        <div class="section-subtitle">{subtitle}</div>
        """,
        unsafe_allow_html=True,
    )


def render_corporate_chart_title(title, subtitle):
    st.markdown(
        f"""
        <div class="chart-corporate-title">
            <h3>{title}</h3>
            <p>{subtitle}</p>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_info(text):
    st.markdown(
        f"""
        <div class="info-box">{text}</div>
        """,
        unsafe_allow_html=True,
    )


def render_success(text):
    st.markdown(
        f"""
        <div class="success-box">{text}</div>
        """,
        unsafe_allow_html=True,
    )


def render_warning(text):
    st.markdown(
        f"""
        <div class="warning-box">{text}</div>
        """,
        unsafe_allow_html=True,
    )


def render_danger(text):
    st.markdown(
        f"""
        <div class="danger-box">{text}</div>
        """,
        unsafe_allow_html=True,
    )


def render_footer():
    st.markdown(
        """
        <div class="footer-main">
            <div><strong>Sistem Pengelolaan Data Pengadaan</strong> • Pelindo Jasa Maritim - BIMA</div>
            <div>Streamlit • Pandas • SQLite</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


# ==============================
# CHART BAR CORPORATE
# ==============================

def make_long_horizontal_bar(
    df,
    x_col,
    y_col,
    title,
    x_title,
    bar_color="#0078B8",
    value_mode="number",
):
    if df.empty:
        fig = go.Figure()
        fig.update_layout(
            title="Tidak ada data",
            height=360,
            paper_bgcolor="#FFFFFF",
            plot_bgcolor="#FFFFFF",
            font=dict(color="#092D55", family="Inter"),
        )
        return fig

    data = df.copy()
    data[x_col] = pd.to_numeric(data[x_col], errors="coerce").fillna(0)
    data[y_col] = data[y_col].astype(str)
    data = data.sort_values(x_col, ascending=True)

    if value_mode == "currency":
        text_values = [format_rupiah_ringkas(v).replace("Rp ", "") for v in data[x_col]]
    elif value_mode == "day":
        text_values = [f"{float(v):.1f} hari" for v in data[x_col]]
    else:
        text_values = [format_number(v) for v in data[x_col]]

    height = max(420, 78 * len(data) + 140)

    fig = go.Figure()

    fig.add_trace(
        go.Bar(
            x=data[x_col],
            y=data[y_col],
            orientation="h",
            text=text_values,
            textposition="outside",
            marker=dict(
                color=bar_color,
                line=dict(color="#005B9E", width=1),
            ),
            hovertemplate="<b>%{y}</b><br>Nilai: %{text}<extra></extra>",
        )
    )

    fig.update_layout(
        title=dict(
            text=title,
            x=0.5,
            xanchor="center",
            font=dict(
                size=24,
                color="#092D55",
                family="Inter",
            ),
        ),
        height=height,
        margin=dict(l=240, r=120, t=80, b=65),
        paper_bgcolor="#FFFFFF",
        plot_bgcolor="#F7FBFE",
        font=dict(
            family="Inter",
            color="#092D55",
            size=12,
        ),
        xaxis=dict(
            title=dict(
                text=x_title,
                font=dict(color="#0078B8", size=12),
            ),
            gridcolor="#DDEAF4",
            zeroline=False,
            tickfont=dict(color="#092D55"),
        ),
        yaxis=dict(
            title="",
            gridcolor="#EEF5FA",
            tickfont=dict(color="#092D55", size=12),
        ),
        bargap=0.34,
    )

    fig.update_traces(
        textfont=dict(
            color="#092D55",
            size=12,
            family="Inter",
        ),
        cliponaxis=False,
    )

    return fig


def make_efficiency_comparison_bar(df):
    if df.empty:
        fig = go.Figure()
        fig.update_layout(
            title="Tidak ada data",
            height=420,
            paper_bgcolor="#FFFFFF",
            plot_bgcolor="#FFFFFF",
            font=dict(color="#092D55", family="Inter"),
        )
        return fig

    data = df.copy()

    for col in ["Total_Valuation_Price", "Total_Net_Order_Value", "Total_Efisiensi"]:
        data[col] = pd.to_numeric(data[col], errors="coerce").fillna(0)

    data["Status Final"] = data["Status Final"].astype(str)
    data = data.sort_values("Total_Net_Order_Value", ascending=True)

    height = max(480, 95 * len(data) + 150)

    fig = go.Figure()

    fig.add_trace(
        go.Bar(
            y=data["Status Final"],
            x=data["Total_Valuation_Price"],
            name="HPS/OE",
            orientation="h",
            marker=dict(color="#B9D9EF"),
            text=[format_rupiah_ringkas(v).replace("Rp ", "") for v in data["Total_Valuation_Price"]],
            textposition="outside",
            hovertemplate="<b>%{y}</b><br>HPS/OE: %{text}<extra></extra>",
        )
    )

    fig.add_trace(
        go.Bar(
            y=data["Status Final"],
            x=data["Total_Net_Order_Value"],
            name="PO",
            orientation="h",
            marker=dict(color="#0078B8"),
            text=[format_rupiah_ringkas(v).replace("Rp ", "") for v in data["Total_Net_Order_Value"]],
            textposition="outside",
            hovertemplate="<b>%{y}</b><br>PO: %{text}<extra></extra>",
        )
    )

    fig.add_trace(
        go.Bar(
            y=data["Status Final"],
            x=data["Total_Efisiensi"],
            name="Efisiensi",
            orientation="h",
            marker=dict(color="#16A7A0"),
            text=[format_rupiah_ringkas(v).replace("Rp ", "") for v in data["Total_Efisiensi"]],
            textposition="outside",
            hovertemplate="<b>%{y}</b><br>Efisiensi: %{text}<extra></extra>",
        )
    )

    fig.update_layout(
        title=dict(
            text="EFISIENSI",
            x=0.5,
            xanchor="center",
            font=dict(
                size=26,
                color="#092D55",
                family="Inter",
            ),
        ),
        barmode="group",
        height=height,
        margin=dict(l=250, r=135, t=85, b=70),
        paper_bgcolor="#FFFFFF",
        plot_bgcolor="#F7FBFE",
        font=dict(
            family="Inter",
            color="#092D55",
            size=12,
        ),
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="center",
            x=0.5,
            font=dict(color="#092D55"),
            bgcolor="rgba(255,255,255,0)",
        ),
        xaxis=dict(
            title=dict(
                text="Nilai Rupiah",
                font=dict(color="#0078B8", size=12),
            ),
            gridcolor="#DDEAF4",
            zeroline=True,
            zerolinecolor="#B9D9EF",
            tickfont=dict(color="#092D55"),
        ),
        yaxis=dict(
            title="",
            gridcolor="#EEF5FA",
            tickfont=dict(color="#092D55", size=12),
        ),
        bargap=0.22,
        bargroupgap=0.08,
    )

    fig.update_traces(
        textfont=dict(
            color="#092D55",
            size=11,
            family="Inter",
        ),
        cliponaxis=False,
    )

    return fig


# ==============================
# UI STREAMLIT
# ==============================

init_database()

sidebar_col, main_col = st.columns([1.15, 4.2], gap="large")

with sidebar_col:
    menu = render_custom_sidebar()

with main_col:
    if menu == "Dashboard":
        render_header(
            "Sistem Pengelolaan Data Pengadaan",
            "Dashboard ringkasan kinerja Purchase Order, efisiensi, dan lama proses PO."
        )

        render_hero()

        df = read_table(PROCESSED_TABLE)

        if df.empty:
            render_warning(
                "Belum ada data yang tersimpan. Silakan upload file Excel terlebih dahulu melalui menu <b>Entry Data Excel</b>."
            )

            col1, col2, col3, col4, col5 = st.columns(5)

            with col1:
                render_metric_card("Jumlah Data", "0", "Total baris data aktif", "🗄️")

            with col2:
                render_metric_card("Total Paket PO", "0", "Jumlah paket PO unik", "📦")

            with col3:
                render_metric_card("Net Order Value", "Rp 0", "Total nilai order", "💼", "teal")

            with col4:
                render_metric_card("Total Efisiensi", "Rp 0", "Total efisiensi", "📈", "green")

            with col5:
                render_metric_card("Rata-rata Lama PO", "0 Hari", "Rata-rata lama proses", "⏱️", "purple")

            st.write("")

            with st.container(border=True):
                render_section_title(
                    "Cara Menggunakan Sistem",
                    "Ikuti alur berikut agar data dashboard muncul."
                )
                st.markdown(
                    """
                    1. Buka menu **Entry Data Excel**.  
                    2. Pilih bulan data yang sesuai.  
                    3. Upload file Excel PO.  
                    4. Klik tombol **Proses dan Simpan ke Database**.  
                    5. Buka kembali menu **Dashboard** untuk melihat grafik dan rekap data.  
                    6. Gunakan menu **Database & Download** untuk mengunduh atau menghapus data berdasarkan bulan.
                    """
                )

        else:
            df = clean_dashboard_numeric(df)

            available_months = get_available_months()

            if available_months:
                month_options = ["ALL"] + available_months
                month_labels = ["Semua Data"] + [get_month_label(m) for m in available_months]

                selected_month_label = st.selectbox(
                    "Filter Bulan Data",
                    month_labels,
                    index=0
                )

                selected_month = month_options[month_labels.index(selected_month_label)]
                df = filter_df_by_month(df, selected_month)
                df = drop_empty_uploaded_rows(df)

            total_po = df["Purchasing Document"].nunique() if not df.empty else 0
            total_rows = len(df)
            total_order = df["Net Order Value"].sum() if not df.empty else 0
            total_efisiensi = df["Efisiensi"].sum() if not df.empty else 0
            if "Lama Proses PO" in df.columns:
                avg_lama_proses = (
                    pd.to_numeric(
                        df["Lama Proses PO"],
                        errors="coerce"
                    )
                    .fillna(0)
                    .mean()
                )
            else:
                avg_lama_proses = 0

            col1, col2, col3, col4, col5 = st.columns(5)

            with col1:
                render_metric_card("Jumlah Data", format_number(total_rows), "Total baris data aktif", "🗄️")

            with col2:
                render_metric_card("Total Paket PO", format_number(total_po), "Jumlah Purchasing Document unik", "📦")

            with col3:
                render_metric_card("Net Order Value", format_rupiah_ringkas(total_order), "Total Net Order Value", "💼", "teal")

            with col4:
                render_metric_card("Total Efisiensi", format_rupiah_ringkas(total_efisiensi), "Akumulasi nilai efisiensi", "📈", "green")

            with col5:
                render_metric_card("Rata-rata Lama PO", f"{avg_lama_proses:.1f} Hari", "Rata-rata lama proses PO", "⏱️", "purple")

            st.write("")

            if df.empty:
                render_warning("Tidak ada data pada bulan yang dipilih.")
            else:
                paket_po, efisiensi, lama_proses = make_rekap(df)

                st.write("")

                render_corporate_chart_title(
                    "PAKET PURCHASE ORDER (PO)",
                    "Visualisasi jumlah dokumen paket PO berdasarkan status final pengadaan."
                )

                fig_paket = make_long_horizontal_bar(
                    paket_po,
                    x_col="Jumlah Paket PO",
                    y_col="Status Final",
                    title="PAKET PURCHASE ORDER (PO)",
                    x_title="Jumlah Dokumen Paket PO",
                    bar_color="#0078B8",
                    value_mode="number"
                )
                st.plotly_chart(fig_paket, use_container_width=True)

                st.write("")

                render_corporate_chart_title(
                    "EFISIENSI",
                    "Perbandingan nilai HPS/OE, nilai PO, dan total efisiensi berdasarkan status final."
                )

                fig_efisiensi = make_efficiency_comparison_bar(efisiensi)
                st.plotly_chart(fig_efisiensi, use_container_width=True)

                st.write("")

                render_corporate_chart_title(
                    "LAMA PROSES PURCHASE ORDER (PO)",
                    "Visualisasi rata-rata lama proses PO per hari berdasarkan status final pengadaan."
                )

                fig_lama = make_long_horizontal_bar(
                    lama_proses,
                    x_col="Rata_Rata_Lama_Proses_PO",
                    y_col="Status Final",
                    title="LAMA PROSES PURCHASE ORDER (PO)",
                    x_title="Rata-Rata Lama Proses PO Per Hari",
                    bar_color="#16A7A0",
                    value_mode="day"
                )
                st.plotly_chart(fig_lama, use_container_width=True)

                st.write("")

                table1, table2, table3 = st.columns(3)

                with table1:
                    with st.container(border=True):
                        render_section_title("Ringkasan Paket PO", "Tabel jumlah paket PO.")
                        st.dataframe(paket_po, use_container_width=True, hide_index=True)

                with table2:
                    with st.container(border=True):
                        render_section_title("Ringkasan Efisiensi", "Tabel efisiensi pengadaan.")
                        st.dataframe(efisiensi, use_container_width=True, hide_index=True)

                with table3:
                    with st.container(border=True):
                        render_section_title("Ringkasan Lama Proses PO", "Tabel rata-rata proses PO.")
                        st.dataframe(lama_proses, use_container_width=True, hide_index=True)

        render_footer()

    elif menu == "Entry Data Excel":
        render_header(
            "Entry Data Excel",
            "Upload file Excel Purchase Order untuk diproses dan disimpan ke database."
        )

        render_info(
            """
            Pilih bulan data terlebih dahulu sebelum upload. Sistem akan menghapus baris kosong secara otomatis,
            sehingga baris kosong dari Excel tidak ikut masuk ke database dan tidak ikut dihitung di dashboard.
            """
        )

        with st.container(border=True):
            render_section_title("Periode Data", "Pilih bulan data yang sesuai dengan file Excel yang akan diunggah.")

            periode_tanggal = st.date_input(
                "Pilih Bulan Data",
                value=get_now_wib().date()
            )

            periode_data = get_month_key_from_date(periode_tanggal)
            periode_label = get_month_label(periode_data)

            st.info(f"Periode data yang dipilih: {periode_label}")
            st.caption(f"Waktu sistem saat ini: {get_now_wib_str()}")

        with st.container(border=True):
            render_section_title("Upload File Excel", "Gunakan file Excel dengan format .xlsx atau .xlsm.")

            uploaded_file = st.file_uploader(
                "Upload file Excel",
                type=["xlsx", "xlsm"]
            )

        if uploaded_file is not None:
            try:
                df_po, df_kk, sheet_names = read_excel_file(uploaded_file)

                total_setelah_bersih = len(df_po)

                render_success(
                    f"File <b>{uploaded_file.name}</b> berhasil dibaca. Sheet yang ditemukan: <b>{', '.join(sheet_names)}</b>."
                )

                st.info(f"Jumlah baris data valid yang terbaca: {format_number(total_setelah_bersih)} baris.")

                missing_columns = validate_columns(df_po)

                if missing_columns:
                    render_warning("Format file belum sesuai. Kolom berikut belum ditemukan pada file Excel.")

                    with st.container(border=True):
                        render_section_title("Kolom yang Belum Ditemukan", "Pastikan nama kolom pada Excel sama persis.")
                        st.write(missing_columns)

                elif df_po.empty:
                    render_warning("File berhasil dibaca, tetapi tidak ada baris data valid. Pastikan kolom Purchasing Document terisi.")

                else:
                    with st.container(border=True):
                        render_section_title("Preview Data Input", "Berikut 20 baris awal dari data valid yang akan diproses.")
                        st.dataframe(df_po.head(20), use_container_width=True)

                    replace_month_data = st.checkbox(
                        f"Hapus data lama untuk periode {periode_label} sebelum menyimpan data baru"
                    )

                    if st.button("Proses dan Simpan ke Database"):
                        if replace_month_data:
                            delete_data_by_month(periode_data)

                        df_processed = process_po_data(
    df_po,
    df_kk
)


                        # pastikan kolom Lama Proses PO tersimpan numeric
                        if "Lama Proses PO" in df_processed.columns:

                            df_processed["Lama Proses PO"] = pd.to_numeric(
                                df_processed["Lama Proses PO"],
                                errors="coerce"
                            ).fillna(0)


                        df_processed = drop_empty_uploaded_rows(
                            df_processed
                        )


                        df_po_save = prepare_month_metadata(
                            df_po,
                            periode_data,
                            periode_label,
                            uploaded_file.name
                        )


                        df_processed_save = prepare_month_metadata(
                            df_processed,
                            periode_data,
                            periode_label,
                            uploaded_file.name
                        )


                        df_po_save = drop_empty_uploaded_rows(
                            df_po_save
                        )


                        df_processed_save = drop_empty_uploaded_rows(
                            df_processed_save
                        )


# Validasi akhir Lama Proses PO

if "Lama Proses PO" in df_processed_save.columns:

    df_processed_save["Lama Proses PO"] = pd.to_numeric(
        df_processed_save["Lama Proses PO"],
        errors="coerce"
    ).fillna(0)


save_dataframe_to_db(
    df_po_save,
    df_processed_save,
    mode="append"
)

                        save_upload_history(
                            uploaded_file.name,
                            len(df_processed_save),
                            periode_data,
                            periode_label
                        )

                        render_success(
                            f"Data berhasil diproses dan disimpan ke database untuk periode <b>{periode_label}</b>. "
                            f"Jumlah baris data valid yang tersimpan: <b>{format_number(len(df_processed_save))}</b> baris."
                        )

                        with st.container(border=True):
                            render_section_title("Preview Hasil Pengolahan", "Kolom hasil olahan sudah ditambahkan ke data PO.")
                            st.dataframe(df_processed_save.head(20), use_container_width=True)

                        excel_bytes = dataframe_to_excel_bytes(
                            {
                                "Data Mentah": df_po_save,
                                "Hasil Pengolahan": df_processed_save,
                            }
                        )

                        st.download_button(
                            label=f"Download Hasil Pengolahan {periode_label}",
                            data=excel_bytes,
                            file_name=f"hasil_pengolahan_po_{periode_data}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        )

            except Exception as e:
                render_warning("File gagal diproses. Cek kembali format file Excel dan nama sheet.")
                st.exception(e)

        render_footer()

    elif menu == "Hasil Pengolahan Data":
        render_header(
            "Hasil Pengolahan Data",
            "Menampilkan data PO yang sudah diproses dan ditambahkan kolom hasil olahan."
        )

        df = read_table(PROCESSED_TABLE)

        if df.empty:
            render_warning("Belum ada data hasil pengolahan.")

        else:
            available_months = get_available_months()

            if available_months:
                month_options = ["ALL"] + available_months
                month_labels = ["Semua Data"] + [get_month_label(m) for m in available_months]

                selected_month_label = st.selectbox(
                    "Filter Bulan Data",
                    month_labels,
                    index=0
                )

                selected_month = month_options[month_labels.index(selected_month_label)]
                df = filter_df_by_month(df, selected_month)
                df = drop_empty_uploaded_rows(df)

            render_info(
                """
                Data berikut merupakan hasil pengolahan dari file Excel yang sudah di-upload.
                Baris kosong tidak ditampilkan dan tidak ikut dihitung.
                """
            )

            with st.container(border=True):
                render_section_title("Tabel Hasil Pengolahan", f"Jumlah baris data valid: {format_number(len(df))} baris.")
                st.dataframe(df, use_container_width=True)

            excel_bytes = dataframe_to_excel_bytes(
                {
                    "Hasil Pengolahan": df
                }
            )

            st.download_button(
                label="Download Hasil Pengolahan",
                data=excel_bytes,
                file_name="hasil_pengolahan_po.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

        render_footer()

    elif menu == "Database & Download":
        render_header(
            "Database & Download",
            "Menampilkan database tersimpan, download data bulanan, dan hapus data berdasarkan bulan."
        )

        available_months = get_available_months()

        with st.container(border=True):
            render_section_title("Download & Hapus Data Bulanan", "Pilih bulan data untuk mengunduh atau menghapus data yang tersimpan.")

            if not available_months:
                render_warning("Belum ada data bulanan yang tersimpan.")
            else:
                month_labels = [get_month_label(m) for m in available_months]

                selected_month_label = st.selectbox(
                    "Pilih Bulan Data",
                    month_labels
                )

                selected_month = available_months[month_labels.index(selected_month_label)]

                df_raw_all = read_table(RAW_TABLE)
                df_processed_all = read_table(PROCESSED_TABLE)

                df_raw_month = filter_df_by_month(df_raw_all, selected_month)
                df_processed_month = filter_df_by_month(df_processed_all, selected_month)

                df_raw_month = drop_empty_uploaded_rows(df_raw_month)
                df_processed_month = drop_empty_uploaded_rows(df_processed_month)

                paket_po, efisiensi, lama_proses = make_rekap(df_processed_month)

                col_a, col_b, col_c = st.columns(3)

                with col_a:
                    render_metric_card(
                        "Periode Data",
                        selected_month_label,
                        "Bulan data yang dipilih",
                        "📅"
                    )

                with col_b:
                    render_metric_card(
                        "Jumlah Data",
                        format_number(len(df_processed_month)),
                        "Total baris data aktif",
                        "🗄️"
                    )

                with col_c:
                    render_metric_card(
                        "Total Paket PO",
                        format_number(df_processed_month["Purchasing Document"].nunique() if not df_processed_month.empty else 0),
                        "Paket PO unik pada bulan ini",
                        "📦"
                    )

                monthly_excel = dataframe_to_excel_bytes(
                    {
                        "Data Mentah": df_raw_month,
                        "Hasil Pengolahan": df_processed_month,
                        "Rekap Paket PO": paket_po,
                        "Rekap Efisiensi": efisiensi,
                        "Rekap Lama Proses PO": lama_proses,
                    }
                )

                st.download_button(
                    label=f"Download Data Bulan {selected_month_label}",
                    data=monthly_excel,
                    file_name=f"database_po_{selected_month}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

                st.write("")

                render_danger(
                    f"""
                    <b>Hapus Data Bulanan</b><br>
                    Fitur ini akan menghapus data mentah, hasil pengolahan, dan riwayat upload untuk periode
                    <b>{selected_month_label}</b>. Data yang sudah dihapus tidak bisa dikembalikan dari aplikasi.
                    """
                )

                confirm_delete = st.checkbox(
                    f"Saya yakin ingin menghapus data periode {selected_month_label}"
                )

                if confirm_delete:
                    if st.button(f"Hapus Data {selected_month_label}"):
                        delete_data_by_month(selected_month)
                        st.success(f"Data periode {selected_month_label} berhasil dihapus.")
                        st.rerun()

        st.write("")

        with st.container(border=True):
            render_section_title("Riwayat Upload", "Daftar file Excel yang pernah diproses oleh sistem.")

            upload_history = read_table(UPLOAD_TABLE)

            if upload_history.empty:
                st.info("Belum ada riwayat upload.")
            else:
                st.dataframe(upload_history, use_container_width=True, hide_index=True)

        st.write("")

        tab1, tab2, tab3 = st.tabs(
            [
                "Data Mentah",
                "Data Hasil Pengolahan",
                "Download Semua Database",
            ]
        )

        with tab1:
            with st.container(border=True):
                render_section_title("Data Mentah Tersimpan", "Data asli hasil upload dari file Excel.")

                df_raw = read_table(RAW_TABLE)
                df_raw = drop_empty_uploaded_rows(df_raw)

                if df_raw.empty:
                    st.info("Belum ada data mentah tersimpan.")
                else:
                    st.dataframe(df_raw, use_container_width=True)

                    raw_excel = dataframe_to_excel_bytes(
                        {
                            "Data Mentah": df_raw
                        }
                    )

                    st.download_button(
                        label="Download Semua Data Mentah Excel",
                        data=raw_excel,
                        file_name="data_mentah_po_semua.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )

        with tab2:
            with st.container(border=True):
                render_section_title("Data Hasil Pengolahan Tersimpan", "Data PO yang sudah ditambahkan hasil perhitungan dan klasifikasi.")

                df_processed = read_table(PROCESSED_TABLE)
                df_processed = drop_empty_uploaded_rows(df_processed)

                if df_processed.empty:
                    st.info("Belum ada data hasil pengolahan tersimpan.")

                else:
                    st.dataframe(df_processed, use_container_width=True)

                    paket_po, efisiensi, lama_proses = make_rekap(df_processed)

                    processed_excel = dataframe_to_excel_bytes(
                        {
                            "Hasil Pengolahan": df_processed,
                            "Rekap Paket PO": paket_po,
                            "Rekap Efisiensi": efisiensi,
                            "Rekap Lama Proses PO": lama_proses,
                        }
                    )

                    st.download_button(
                        label="Download Semua Hasil Pengolahan & Rekap",
                        data=processed_excel,
                        file_name="hasil_pengolahan_dan_rekap_po_semua.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )

        with tab3:
            with st.container(border=True):
                render_section_title("Download Semua Database", "Export seluruh data mentah, hasil olahan, dan rekap dashboard.")

                df_raw = read_table(RAW_TABLE)
                df_processed = read_table(PROCESSED_TABLE)

                df_raw = drop_empty_uploaded_rows(df_raw)
                df_processed = drop_empty_uploaded_rows(df_processed)

                if df_processed.empty:
                    st.info("Belum ada database yang bisa diunduh.")

                else:
                    paket_po, efisiensi, lama_proses = make_rekap(df_processed)

                    excel_bytes = dataframe_to_excel_bytes(
                        {
                            "Data Mentah": df_raw,
                            "Hasil Pengolahan": df_processed,
                            "Rekap Paket PO": paket_po,
                            "Rekap Efisiensi": efisiensi,
                            "Rekap Lama Proses PO": lama_proses,
                        }
                    )

                    st.download_button(
                        label="Download Semua Data dalam Excel",
                        data=excel_bytes,
                        file_name="database_po_export_semua.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )

                if os.path.exists(DB_PATH):
                    with open(DB_PATH, "rb") as db_file:
                        st.download_button(
                            label="Download Database SQLite (.db)",
                            data=db_file,
                            file_name="po_database.db",
                            mime="application/octet-stream",
                        )

        render_footer()
