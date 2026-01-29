# app.py
# Streamlit app: Collection Efficiency Automation (DVC consolidate + Hub-wise split)
# + "Arrear_Advance" pivot sheet (cust_id wise) in the DVC output.
# + NEW MODULE: "LPC Report" (filtered detail + formatted Summary)
# IMPORTANT:
# - Arrear hub-wise saved as CSV (to avoid MemoryError).
# - LPC detail sheet is written by pandas (fast, low memory) + Summary formatted by openpyxl.

import os
import re
import zipfile
import tempfile
import time
import pandas as pd
import numpy as np
import streamlit as st

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.utils import get_column_letter
def try_import(module_name, func_name):
    try:
        mod = __import__(module_name, fromlist=[func_name])
        return getattr(mod, func_name), None
    except Exception as e:
        return None, str(e)

run_vault_cash_report, vault_err = try_import("Vault_Cash", "run_vault_cash_report")
run_pending_collection_entry, pending_err = try_import("Pending_Collection", "run_pending_collection_entry")
run_demand_verification, dv_err = try_import("Demand_Verification", "run_demand_verification")
run_map_ledger_difference, map_err = try_import("Branch_MAP", "run_map_ledger_difference")
run_from_streamlit, ex_err = try_import("Excess_Amount", "run_from_streamlit")
run_cpp_payable_vs_cpp_ledger_difference, cpp_err = try_import("CPP_Payable", "run_cpp_payable_vs_cpp_ledger_difference")
run_cms_recon_streamlit, cms_err = try_import("CMS_Recon", "run_cms_recon_streamlit")



# ============================================================
# PAGE UI
# ============================================================
st.set_page_config(page_title="Reporting_Team", page_icon="📊", layout="wide")
st.title("📊 Reporting_Team - Reports Automatisation")
st.markdown("""
<style>
/* --- TAB BUTTON STYLE (Compact + Bright Orange) --- */

/* Tab container */
.stTabs [data-baseweb="tab-list"]{
    gap: 4px;
    border-bottom: none !important;
    padding-bottom: 4px;
    flex-wrap: nowrap;
    overflow-x: auto;
}

/* Each tab */
.stTabs [data-baseweb="tab"]{
    background: #2b2b2b !important;          /* dark grey */
    border: 1px solid #ffffff !important;    /* white outline */
    border-radius: 4px !important;
    padding: 3px 6px !important;             /* smaller width */
    color: #ffffff !important;
    min-height: 26px !important;
}

/* Tab text */
.stTabs [data-baseweb="tab"] p{
    font-size: 11px !important;
    font-weight: 600 !important;
    margin: 0 !important;
    white-space: nowrap;
}

/* Active tab – Bright Orange */
.stTabs [aria-selected="true"]{
    background: #ff6a00 !important;          /* sharp bright orange */
    border: 1px solid #ff6a00 !important;
    color: #ffffff !important;
}

/* Remove default underline */
.stTabs [data-baseweb="tab-highlight"],
.stTabs [data-baseweb="tab-border"]{
    display: none !important;
}

/* ---------------- File Uploader Button Styling ---------------- */

div[data-testid="stFileUploader"] button{
    background-color: #ff6a00 !important;    /* same bright orange */
    color: white !important;
    border-radius: 6px !important;
    border: none !important;
    padding: 5px 12px !important;            /* compact button */
    font-weight: 600 !important;
}

div[data-testid="stFileUploader"] button:hover{
    background-color: #e65c00 !important;
}

/* ============================================================
   NEW: Make Upload Boxes Smaller (Height + Padding) everywhere
   ============================================================ */

/* Orange label line: "Upload ....xlsx" */
div[data-testid="stFileUploader"] label{
    color: #ff6a00 !important;               /* same as Browse button */
    font-weight: 700 !important;
    font-size: 13px !important;
}

/* Reduce uploader drop-zone height + padding */
div[data-testid="stFileUploader"] section{
    padding: 8px 10px !important;            /* smaller padding */
    min-height: 52px !important;             /* reduce height */
    border-radius: 10px !important;
}

/* Make inner content compact */
div[data-testid="stFileUploader"] section *{
    line-height: 1.15 !important;
}

/* "Drag and drop..." and helper text smaller */
div[data-testid="stFileUploader"] section small{
    font-size: 10px !important;
}

/* Uploaded file row compact */
div[data-testid="stFileUploader"] ul{
    margin-top: 4px !important;
}
div[data-testid="stFileUploader"] li{
    padding: 4px 8px !important;
    font-size: 11px !important;
    border-radius: 8px !important;
}

/* Optional: make browse button slightly smaller too */
div[data-testid="stFileUploader"] button{
    padding: 4px 10px !important;
    font-size: 13px !important;
}
/* ===== Force ultra-compact upload box ===== */

/* Outer wrapper */
div[data-testid="stFileUploader"]{
    margin-bottom: 6px !important;
}

/* Drop zone main area */
div[data-testid="stFileUploader"] section{
    min-height: 38px !important;     /* hard reduce height */
    padding: 6px 8px !important;     /* tight padding */
}

/* Icon + text row */
div[data-testid="stFileUploader"] section > div{
    gap: 6px !important;
}

/* Cloud icon size */
div[data-testid="stFileUploader"] svg{
    width: 20px !important;
    height: 20px !important;
}

/* Main text: "Drag and drop file here" */
div[data-testid="stFileUploader"] section strong,
div[data-testid="stFileUploader"] section span{
    font-size: 12px !important;
    line-height: 1.1 !important;
}

/* Helper text: "Limit 200MB..." */
div[data-testid="stFileUploader"] section small{
    font-size: 9px !important;
    margin-top: 0 !important;
    line-height: 1 !important;
}

/* Uploaded file row */
div[data-testid="stFileUploader"] li{
    padding: 3px 6px !important;
    font-size: 10px !important;
}
/* ===== HARD REDUCE uploader box HEIGHT (not only text) ===== */

/* Main drop area */
div[data-testid="stFileUploader"] section{
    padding: 4px 8px !important;
    min-height: 34px !important;
    height: 44px !important;          /* force actual height */
    display: flex !important;
    align-items: center !important;
}

/* Inner wrapper */
div[data-testid="stFileUploader"] section > div{
    padding: 0 !important;
    margin: 0 !important;
    width: 100% !important;
}

/* The clickable dropzone block */
div[data-testid="stFileUploader"] section div[role="button"]{
    padding: 0 !important;
    margin: 0 !important;
    min-height: 34px !important;
    height: 44px !important;          /* force height */
    display: flex !important;
    align-items: center !important;
}

/* Reduce left icon spacing */
div[data-testid="stFileUploader"] section svg{
    width: 18px !important;
    height: 18px !important;
    margin-right: 6px !important;
}
/* ===== Compact Browse Files button container ===== */

/* Reduce the height of the button wrapper */
div[data-testid="stFileUploader"] section div[role="button"] > div:last-child{
    padding: 0 !important;
    margin: 0 !important;
    display: flex !important;
    align-items: center !important;
}

/* Make the Browse button slimmer */
div[data-testid="stFileUploader"] button{
    padding: 3px 10px !important;   /* smaller height */
    font-size: 12px !important;
    line-height: 1.1 !important;
    border-radius: 5px !important;
}
/* ========= Report Dropdown Styling ========= */
div[data-testid="stSelectbox"] label{
    color: #ff6a00 !important;
    font-weight: 800 !important;
    font-size: 14px !important;
}

div[data-testid="stSelectbox"] div[role="combobox"]{
    background: #2b2b2b !important;
    border: 1px solid #ffffff !important;
    border-radius: 8px !important;
    padding: 6px 10px !important;
    min-height: 42px !important;
    color: #ffffff !important;
}

div[data-testid="stSelectbox"] div[role="combobox"]:focus-within{
    border: 1px solid #ff6a00 !important;
    box-shadow: 0 0 0 2px rgba(255,106,0,0.35) !important;
}
/* ===== Selectbox (Report Dropdown) – Match Browse Button Color ===== */

/* Main selectbox container */
div[data-baseweb="select"] > div {
    background-color: #ff6a00 !important;
    border: 1px solid #ff6a00 !important;
    color: white !important;
}

/* Text inside selectbox */
div[data-baseweb="select"] span {
    color: white !important;
    font-weight: 600 !important;
}

/* Dropdown arrow */
div[data-baseweb="select"] svg {
    fill: white !important;
}

/* Dropdown menu items */
ul[role="listbox"] {
    background-color: #1f1f1f !important;
}

/* Hovered option */
ul[role="listbox"] li:hover {
    background-color: #ff6a00 !important;
    color: white !important;
}

/* Selected option */
ul[role="listbox"] li[aria-selected="true"] {
    background-color: #ff6a00 !important;
    color: white !important;
}
/* Reduce width of report dropdown (search bar) */
div[data-testid="stSelectbox"]{
    max-width: 50% !important;   /* half width */
}

</style>
""", unsafe_allow_html=True)

report_options = [
    "1) Collection Efficiency Report",
    "2) Arrear Advance Report",
    "3) LPC Report",
    "4) Vault Cash Report",
    "5) Pending Collection Entry",
    "6) Demand Verification Report",
    "7) MAP Ledger Difference",
    "8) Excess Amount",
    "9) CPP Payable",
    "10) CMS and UPI Recon"
]

selected_report = st.selectbox(
    "🔎 Select Report (type to search)",
    report_options,
    index=0,
    key="report_selector"
)


# ============================================================
# HELPERS
# ============================================================
def save_uploaded_file(uploaded, path):
    with open(path, "wb") as f:
        f.write(uploaded.getbuffer())


def zip_folder(folder_path, zip_path):
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for root, _, files in os.walk(folder_path):
            for file in files:
                full_path = os.path.join(root, file)
                rel_path = os.path.relpath(full_path, folder_path)
                zf.write(full_path, rel_path)


def clean_name_for_folder(text):
    if text is None:
        return "BLANK"
    s = str(text).strip()
    if s == "" or s.lower() == "nan":
        return "BLANK"
    s = re.sub(r'[\\/:*?"<>|]', "", s)
    return s[:80]


def safe_sheet_title(name: str) -> str:
    name = re.sub(r'[:\\/?*\[\]]', "_", str(name))
    return name[:31]


def to_numeric_clean(series: pd.Series) -> pd.Series:
    return pd.to_numeric(
        series.astype(str)
              .str.replace(",", "", regex=False)
              .str.replace("%", "", regex=False),
        errors="coerce"
    )


def is_percent_column(col_name: str, series_numeric: pd.Series) -> bool:
    name = str(col_name).lower()
    if "%" in name or "efficiency" in name or "pct" in name or "percent" in name or "(y/x)" in name:
        return True
    s = series_numeric.dropna()
    if s.empty:
        return False
    mn, mx = s.min(), s.max()
    if mn >= -1 and mx <= 1.5 and (s % 1 != 0).any():
        return True
    return False


def recompute_grand_total(df: pd.DataFrame, hub_col: str) -> pd.DataFrame:
    if df.empty:
        return df

    first_col = df.columns[0]
    gt_mask = df[first_col].astype(str).str.strip().str.lower().eq("grand total")
    if not gt_mask.any():
        return df

    gt_row = df[gt_mask].iloc[:1].copy()
    data_rows = df[~gt_mask].copy()

    for col in df.columns:
        if col in (first_col, hub_col):
            continue
        num = to_numeric_clean(data_rows[col])
        if num.notna().sum() == 0:
            continue
        if is_percent_column(col, num):
            gt_row.iloc[0, df.columns.get_loc(col)] = float(num.mean())
        else:
            gt_row.iloc[0, df.columns.get_loc(col)] = float(num.sum())

    return pd.concat([gt_row, data_rows], ignore_index=True)


def build_format_map(df: pd.DataFrame, hub_col: str) -> dict:
    fmt_map = {}
    for idx, col in enumerate(df.columns, start=1):
        if col == hub_col:
            continue
        num = to_numeric_clean(df[col])
        if num.notna().sum() == 0:
            continue
        fmt_map[idx] = "0.00%" if is_percent_column(col, num) else "#,##0"
    return fmt_map


def write_df_fast(ws, df: pd.DataFrame, format_map: dict):
    for row in dataframe_to_rows(df, index=False, header=True):
        ws.append(row)

    header_font = Font(bold=True)
    for cell in ws[1]:
        cell.font = header_font

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    max_row = ws.max_row
    for col_idx, fmt in format_map.items():
        for r in range(2, max_row + 1):
            c = ws.cell(row=r, column=col_idx)
            if isinstance(c.value, (int, float)) and c.value is not None:
                c.number_format = fmt


def read_ce_sheets(path: str) -> dict:
    return pd.read_excel(path, sheet_name=None, engine="openpyxl")


def read_csv_with_fallback(path, compression=None):
    encodings = ["utf-8", "utf-8-sig", "cp1252", "latin1"]
    last_err = None
    for enc in encodings:
        try:
            return pd.read_csv(path, compression=compression, low_memory=False, encoding=enc)
        except Exception as e:
            last_err = e
    raise last_err


def read_arrear_any(path: str, original_name: str, use_csv: bool) -> pd.DataFrame:
    ext = os.path.splitext(original_name)[1].lower()

    if use_csv and ext == ".csv":
        return read_csv_with_fallback(path)

    if use_csv and ext == ".zip":
        return read_csv_with_fallback(path, compression="zip")

    return pd.read_excel(path, engine="openpyxl")


# ============================================================
# DVC CONSOLIDATE LOGIC (Reusable)
# ============================================================
def build_dvc_consolidate(jlg_path: str, il_path: str) -> pd.DataFrame:
    header_map_jlg = [
        'COMPANY_ID', 'ZoneName', 'region_name', 'unit_name', 'cluster_name',
        'state_id', 'branch_id', 'branch_name', 'CENTER_ID', 'center_name',
        'group_name', 'FirstNAme', 'LastName', 'facility_id', 'installment_no',
        'loan_id', 'cust_id', 'opening_principal_advance', 'opening_interest_advance',
        'overdue_principal_demand', 'overdue_interest_demand', 'current_principal_demand',
        'current_interest_demand', 'insurance_premium_due', 'total_demand',
        'overdue_principal_collected', 'overdue_interest_collected', 'current_principal_collected',
        'current_interest_collected', 'advance_principal_collected', 'advance_interest_collected',
        'insurance_premium_collected', 'total_collection', 'excess_amt_collected',
        'disbursement_date', 'loan_maturity_date', 'Last_Collection_Paid', 'collection_date',
        'repayment_frequency', 'Loan_tenure_Month', 'No_of_Inst_OD', 'Total_Inst_Amt_Paid',
        'No_of_Inst_Due', 'mobile_number', 'P2P', 'prod_category_id', 'product_id',
        'status', 'Installment_Date', 'ReportDate', 'PRECLOSURE_REASON',
        'od_days', 'Spouse_Name'
    ]

    jlg_df = pd.read_excel(jlg_path, engine="openpyxl", usecols=header_map_jlg)
    il_df = pd.read_excel(il_path, engine="openpyxl")

    il_df.rename(columns={
        'zone_name': 'ZoneName',
        'area_name': 'unit_name',
        'report_date': 'ReportDate'
    }, inplace=True)

    if 'branch_name' in il_df.columns:
        il_df['branch_id'] = il_df['branch_name'].astype(str).str[:5]
    else:
        il_df['branch_id'] = None

    il_aligned_df = pd.DataFrame({c: il_df[c] if c in il_df.columns else None for c in header_map_jlg})
    consolidated_df = pd.concat([jlg_df, il_aligned_df], ignore_index=True)

    consolidated_df.columns = consolidated_df.columns.str.strip().str.replace(' ', '_').str.lower()

    if 'branch_id' in consolidated_df.columns:
        cols = consolidated_df.columns.tolist()
        cols.insert(0, cols.pop(cols.index('branch_id')))
        consolidated_df = consolidated_df[cols]

    if 'zonename' in consolidated_df.columns:
        consolidated_df.rename(columns={'zonename': 'hub'}, inplace=True)

    if 'company_id' in consolidated_df.columns:
        consolidated_df.drop(columns=['company_id'], inplace=True)

    if 'status' in consolidated_df.columns:
        consolidated_df = consolidated_df[
            ~consolidated_df['status'].astype(str).str.lower().str.contains('death', na=False)
        ]

    num_cols = [
        'overdue_interest_collected', 'current_interest_collected',
        'overdue_principal_collected', 'current_principal_collected',
        'opening_interest_advance', 'opening_principal_advance',
        'current_interest_demand', 'current_principal_demand',
        'advance_interest_collected', 'advance_principal_collected',
        'total_collection'
    ]
    for c in num_cols:
        if c not in consolidated_df.columns:
            consolidated_df[c] = 0
        consolidated_df[c] = pd.to_numeric(
            consolidated_df[c].astype(str).str.replace(",", "", regex=False),
            errors="coerce"
        ).fillna(0)

    mask_i = consolidated_df['current_interest_collected'] < 0
    consolidated_df.loc[mask_i, 'overdue_interest_collected'] += consolidated_df.loc[mask_i, 'current_interest_collected']
    consolidated_df.loc[mask_i, 'current_interest_collected'] = 0

    mask_p = consolidated_df['current_principal_collected'] < 0
    consolidated_df.loc[mask_p, 'overdue_principal_collected'] += consolidated_df.loc[mask_p, 'current_principal_collected']
    consolidated_df.loc[mask_p, 'current_principal_collected'] = 0

    consolidated_df['advance_demand'] = consolidated_df['opening_interest_advance'] + consolidated_df['opening_principal_advance']
    consolidated_df['current_demand'] = consolidated_df['current_interest_demand'] + consolidated_df['current_principal_demand']
    consolidated_df['advance'] = consolidated_df['advance_interest_collected'] + consolidated_df['advance_principal_collected']
    consolidated_df['advance'] = consolidated_df[['advance', 'total_collection']].min(axis=1)

    consolidated_df['collection_exc_advance'] = consolidated_df['total_collection'] - consolidated_df['advance']
    consolidated_df['same_day_demand'] = (consolidated_df['current_demand'] - consolidated_df['advance_demand']).clip(lower=0)
    consolidated_df.loc[consolidated_df['current_demand'] == 0, 'same_day_demand'] = 0
    consolidated_df['same_day_repayment_received'] = consolidated_df[['same_day_demand', 'collection_exc_advance']].min(axis=1)
    consolidated_df['od_received'] = (consolidated_df['collection_exc_advance'] - consolidated_df['same_day_demand']).clip(lower=0)
    consolidated_df['same_day_arrear'] = (consolidated_df['same_day_demand'] - consolidated_df['same_day_repayment_received']).clip(lower=0)

    return consolidated_df


def build_arrear_advance_sheet(consolidated_df: pd.DataFrame) -> pd.DataFrame:
    df = consolidated_df.copy()

    for col in ["cust_id", "advance", "same_day_arrear", "same_day_demand"]:
        if col not in df.columns:
            df[col] = 0

    df["advance"] = pd.to_numeric(df["advance"], errors="coerce").fillna(0)
    df["same_day_arrear"] = pd.to_numeric(df["same_day_arrear"], errors="coerce").fillna(0)
    df["same_day_demand"] = pd.to_numeric(df["same_day_demand"], errors="coerce").fillna(0)

    df = df[df["same_day_demand"] != 0].copy()

    pivot = (
        df.groupby("cust_id", as_index=False)[["advance", "same_day_arrear"]]
          .sum()
    )

    pivot = pivot[(pivot["advance"] > 0) & (pivot["same_day_arrear"] > 0)].copy()
    pivot.sort_values(["same_day_arrear", "advance"], ascending=False, inplace=True)

    return pivot


# ============================================================
# LPC REPORT MODULE (NEW)
# ============================================================
def lpc_generate_report(in_path: str, out_path: str, sheet_name=0):
    # ---------- Read ----------
    df = pd.read_excel(in_path, sheet_name=sheet_name, engine="openpyxl")
    df.columns = df.columns.astype(str).str.strip()

    # ---------- Filters ----------
    if "Status" not in df.columns:
        raise ValueError("Column 'Status' not found in LPC file.")
    df["Status"] = df["Status"].astype(str).str.strip().str.upper()
    df = df[~df["Status"].isin(["DEATH", "DEATH APPROVED"])].copy()

    if "DISBURSEMENT_DATE" not in df.columns:
        raise ValueError("Column 'DISBURSEMENT_DATE' not found in LPC file.")
    cutoff = pd.Timestamp("2025-10-01")
    raw_date = df["DISBURSEMENT_DATE"]

    dt1 = pd.to_datetime(raw_date, errors="coerce", dayfirst=True)
    is_serial = pd.to_numeric(raw_date, errors="coerce").notna() & dt1.isna()
    serial_vals = pd.to_numeric(raw_date, errors="coerce")

    dt2 = dt1.copy()
    dt2.loc[is_serial] = pd.to_datetime(serial_vals.loc[is_serial], unit="D", origin="1899-12-30", errors="coerce")
    df["DISBURSEMENT_DATE"] = dt2
    df = df[df["DISBURSEMENT_DATE"] >= cutoff].copy()

    # ---------- Remarks ----------
    for col in ["Total_Arrear", "Total_OS", "Pending_LPC"]:
        if col not in df.columns:
            raise ValueError(f"Missing column for remarks: {col}")
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    df["Remarks"] = np.select(
        [
            (df["Pending_LPC"] <= 0),
            (df["Total_Arrear"] > 0) & (df["Total_OS"] > 0) & (df["Pending_LPC"] > 0),
            (df["Total_Arrear"] == 0) & (df["Total_OS"] == 0) & (df["Pending_LPC"] > 0),
            (df["Total_Arrear"] == 0) & (df["Total_OS"] > 0) & (df["Pending_LPC"] > 0),
        ],
        [
            "LPC Fully Collected",
            "Arrear & LPC Pending",
            "No Arrear & Outstanding but LPC Pending",
            "No Arrear but LPC Pending",
        ],
        default=""
    )

    # ---------- Pivot base ----------
    required = ["ZONE_NAME", "CLUSTER_NAME", "LOAN_ID", "LPC_Collected", "Pending_LPC"]
    for c in required:
        if c not in df.columns:
            raise ValueError(f"Missing column: {c}")

    df["LPC_Collected"] = pd.to_numeric(df["LPC_Collected"], errors="coerce").fillna(0)
    df["Pending_LPC"] = pd.to_numeric(df["Pending_LPC"], errors="coerce").fillna(0)

    if "Total_LPC_Due" not in df.columns:
        df["Total_LPC_Due"] = df["LPC_Collected"] + df["Pending_LPC"]
    else:
        df["Total_LPC_Due"] = pd.to_numeric(df["Total_LPC_Due"], errors="coerce").fillna(0)

    remark_order = [
        "Arrear & LPC Pending",
        "LPC Fully Collected",
        "No Arrear & Outstanding but LPC Pending",
        "No Arrear but LPC Pending"
    ]

    keys = df.groupby(["ZONE_NAME", "CLUSTER_NAME"], as_index=False).size()[["ZONE_NAME", "CLUSTER_NAME"]]
    summary_df = keys.copy()

    def block(rmk):
        d = df[df["Remarks"] == rmk]
        if d.empty:
            return pd.DataFrame(columns=["ZONE_NAME", "CLUSTER_NAME", "LPC Due", "LPC Collected", "LPC pending", "No. of loans"])
        return d.groupby(["ZONE_NAME", "CLUSTER_NAME"], as_index=False).agg(
            **{
                "LPC Due": ("Total_LPC_Due", "sum"),
                "LPC Collected": ("LPC_Collected", "sum"),
                "LPC pending": ("Pending_LPC", "sum"),
                "No. of loans": ("LOAN_ID", "count"),
            }
        )

    for rmk in remark_order:
        t = block(rmk).rename(columns={
            "LPC Due": f"{rmk}__LPC Due",
            "LPC Collected": f"{rmk}__LPC Collected",
            "LPC pending": f"{rmk}__LPC pending",
            "No. of loans": f"{rmk}__No. of loans",
        })
        summary_df = summary_df.merge(t, on=["ZONE_NAME", "CLUSTER_NAME"], how="left")

    summary_df = summary_df.fillna(0)

    summary_df["Total LPC Due"] = 0
    summary_df["Total LPC Collected"] = 0
    summary_df["Total LPC pending"] = 0
    summary_df["Total No.of loans"] = 0

    for rmk in remark_order:
        summary_df["Total LPC Due"] += summary_df[f"{rmk}__LPC Due"]
        summary_df["Total LPC Collected"] += summary_df[f"{rmk}__LPC Collected"]
        summary_df["Total LPC pending"] += summary_df[f"{rmk}__LPC pending"]
        summary_df["Total No.of loans"] += summary_df[f"{rmk}__No. of loans"]

    summary_df = summary_df.sort_values(["ZONE_NAME", "CLUSTER_NAME"]).reset_index(drop=True)

    # ---------- Write detail sheet by pandas (fast, low memory) ----------
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Member wise data", index=False)

    # ---------- Format summary with openpyxl ----------
    wb = load_workbook(out_path)
    if "Summary" in wb.sheetnames:
        wb.remove(wb["Summary"])
    ws = wb.create_sheet("Summary", 0)
    ws.sheet_view.showGridLines = False

    # Styles (local)
    thin = Side(style="thin", color="000000")
    thick = Side(style="medium", color="000000")
    border_thin = Border(left=thin, right=thin, top=thin, bottom=thin)

    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)
    align_right = Alignment(horizontal="right", vertical="center", wrap_text=True)

    font_title = Font(bold=True, size=12)
    font_head = Font(bold=True, size=10)
    font_bold = Font(bold=True)

    fill_title = PatternFill("solid", fgColor="D9EEF7")
    fill_group1 = PatternFill("solid", fgColor="FCE4D6")
    fill_group2 = PatternFill("solid", fgColor="C6EFCE")
    fill_group3 = PatternFill("solid", fgColor="EAD1F5")
    fill_group4 = PatternFill("solid", fgColor="CFEFFF")
    fill_grand = PatternFill("solid", fgColor="D9EEF7")

    def apply_border_range(ws_, r1, c1, r2, c2):
        for rr in range(r1, r2 + 1):
            for cc in range(c1, c2 + 1):
                ws_.cell(rr, cc).border = border_thin

    def fill_range(ws_, r1, c1, r2, c2, fill):
        for rr in range(r1, r2 + 1):
            for cc in range(c1, c2 + 1):
                ws_.cell(rr, cc).fill = fill

    def set_thick_block_outline(ws_, r1, c1, r2, c2):
        for cc in range(c1, c2 + 1):
            ws_.cell(r1, cc).border = Border(
                left=ws_.cell(r1, cc).border.left,
                right=ws_.cell(r1, cc).border.right,
                top=thick,
                bottom=ws_.cell(r1, cc).border.bottom,
            )
            ws_.cell(r2, cc).border = Border(
                left=ws_.cell(r2, cc).border.left,
                right=ws_.cell(r2, cc).border.right,
                top=ws_.cell(r2, cc).border.top,
                bottom=thick,
            )
        for rr in range(r1, r2 + 1):
            ws_.cell(rr, c1).border = Border(
                left=thick,
                right=ws_.cell(rr, c1).border.right,
                top=ws_.cell(rr, c1).border.top,
                bottom=ws_.cell(rr, c1).border.bottom,
            )
            ws_.cell(rr, c2).border = Border(
                left=ws_.cell(rr, c2).border.left,
                right=thick,
                top=ws_.cell(rr, c2).border.top,
                bottom=ws_.cell(rr, c2).border.bottom,
            )

    def paint_blocks(row_idx):
        fill_range(ws, row_idx, 2, row_idx, 5, fill_group1)
        fill_range(ws, row_idx, 6, row_idx, 9, fill_group2)
        fill_range(ws, row_idx, 10, row_idx, 13, fill_group3)
        fill_range(ws, row_idx, 14, row_idx, 17, fill_group4)

    def format_row_numbers(row_idx):
        for col in range(2, 22):
            cell = ws.cell(row_idx, col)
            cell.alignment = align_right
            cell.number_format = "0" if col in [5, 9, 13, 17, 21] else "#,##0"

    # Title row
    title = "Summary of Late Payment Charges Due, Collected & Pending as on January 21, 2026"
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=21)
    c = ws.cell(1, 1, title)
    c.font = font_title
    c.alignment = align_center
    fill_range(ws, 1, 1, 1, 21, fill_title)
    apply_border_range(ws, 1, 1, 1, 21)

    group_row, sub_row, data_start = 2, 3, 4

    ws.merge_cells(start_row=group_row, start_column=1, end_row=sub_row, end_column=1)
    ws.cell(group_row, 1, "HUB/STATE").font = font_head
    ws.cell(group_row, 1).alignment = align_center
    fill_range(ws, group_row, 1, sub_row, 1, fill_title)

    group_defs = [
        ("Arrear & LPC Pending", 2, 5, fill_group1),
        ("LPC Fully Collected", 6, 9, fill_group2),
        ("No Arrear & Outstanding but LPC Pending", 10, 13, fill_group3),
        ("No Arrear but LPC Pending", 14, 17, fill_group4),
    ]
    for text, c1, c2, fill in group_defs:
        ws.merge_cells(start_row=group_row, start_column=c1, end_row=group_row, end_column=c2)
        h = ws.cell(group_row, c1, text)
        h.font = font_head
        h.alignment = align_center
        fill_range(ws, group_row, c1, group_row, c2, fill)

    subs = ["LPC Due", "LPC Collected", "LPC pending", "No.of loans"]
    for start_col, fill in [(2, fill_group1), (6, fill_group2), (10, fill_group3), (14, fill_group4)]:
        for j, sh in enumerate(subs):
            cc = ws.cell(sub_row, start_col + j, sh)
            cc.font = font_head
            cc.alignment = align_center
            cc.fill = fill

    tot_heads = ["Total LPC Due", "Total LPC Collected", "Total LPC pending", "Total No.of loans"]
    for j, th in enumerate(tot_heads):
        cell = ws.cell(sub_row, 18 + j, th)
        cell.font = font_head
        cell.alignment = align_center
        cell.fill = fill_title
    for col in range(18, 22):
        ws.cell(group_row, col).fill = fill_title

    apply_border_range(ws, group_row, 1, sub_row, 21)

    # Data with zone headers, totals, grand total
    r = data_start
    current_zone = None

    def put_block(row, start_col, rmk):
        ws.cell(r, start_col, float(row[f"{rmk}__LPC Due"]))
        ws.cell(r, start_col + 1, float(row[f"{rmk}__LPC Collected"]))
        ws.cell(r, start_col + 2, float(row[f"{rmk}__LPC pending"]))
        ws.cell(r, start_col + 3, float(row[f"{rmk}__No. of loans"]))

    def write_zone_total(zone_name, zone_df, row_idx):
        ws.cell(row_idx, 1, f"{zone_name} Total").font = font_bold
        ws.cell(row_idx, 1).alignment = align_left

        mapping = {
            2: "Arrear & LPC Pending__LPC Due",
            3: "Arrear & LPC Pending__LPC Collected",
            4: "Arrear & LPC Pending__LPC pending",
            5: "Arrear & LPC Pending__No. of loans",
            6: "LPC Fully Collected__LPC Due",
            7: "LPC Fully Collected__LPC Collected",
            8: "LPC Fully Collected__LPC pending",
            9: "LPC Fully Collected__No. of loans",
            10: "No Arrear & Outstanding but LPC Pending__LPC Due",
            11: "No Arrear & Outstanding but LPC Pending__LPC Collected",
            12: "No Arrear & Outstanding but LPC Pending__LPC pending",
            13: "No Arrear & Outstanding but LPC Pending__No. of loans",
            14: "No Arrear but LPC Pending__LPC Due",
            15: "No Arrear but LPC Pending__LPC Collected",
            16: "No Arrear but LPC Pending__LPC pending",
            17: "No Arrear but LPC Pending__No. of loans",
            18: "Total LPC Due",
            19: "Total LPC Collected",
            20: "Total LPC pending",
            21: "Total No.of loans",
        }
        for col_idx, field in mapping.items():
            v = zone_df[field].sum()
            cell = ws.cell(row_idx, col_idx, float(v))
            cell.font = font_bold
            cell.alignment = align_right
            cell.number_format = "0" if col_idx in [5, 9, 13, 17, 21] else "#,##0"

        paint_blocks(row_idx)
        apply_border_range(ws, row_idx, 1, row_idx, 21)

    for _, row in summary_df.iterrows():
        zone, cluster = row["ZONE_NAME"], row["CLUSTER_NAME"]

        if current_zone is None or zone != current_zone:
            if current_zone is not None:
                prev_zone_df = summary_df[summary_df["ZONE_NAME"] == current_zone]
                write_zone_total(current_zone, prev_zone_df, r)
                r += 1

            # Zone header row (colored blocks)
            ws.cell(r, 1, zone).font = font_bold
            ws.cell(r, 1).alignment = align_left
            paint_blocks(r)
            apply_border_range(ws, r, 1, r, 21)
            r += 1
            current_zone = zone

        ws.cell(r, 1, cluster).alignment = align_left

        put_block(row, 2, "Arrear & LPC Pending")
        put_block(row, 6, "LPC Fully Collected")
        put_block(row, 10, "No Arrear & Outstanding but LPC Pending")
        put_block(row, 14, "No Arrear but LPC Pending")

        ws.cell(r, 18, float(row["Total LPC Due"]))
        ws.cell(r, 19, float(row["Total LPC Collected"]))
        ws.cell(r, 20, float(row["Total LPC pending"]))
        ws.cell(r, 21, float(row["Total No.of loans"]))

        paint_blocks(r)
        format_row_numbers(r)
        apply_border_range(ws, r, 1, r, 21)
        r += 1

    if current_zone is not None:
        last_zone_df = summary_df[summary_df["ZONE_NAME"] == current_zone]
        write_zone_total(current_zone, last_zone_df, r)
        r += 1

    # Grand Total
    ws.cell(r, 1, "Grand Total").font = font_bold
    ws.cell(r, 1).alignment = align_left

    col_map = {
        2: "Arrear & LPC Pending__LPC Due",
        3: "Arrear & LPC Pending__LPC Collected",
        4: "Arrear & LPC Pending__LPC pending",
        5: "Arrear & LPC Pending__No. of loans",
        6: "LPC Fully Collected__LPC Due",
        7: "LPC Fully Collected__LPC Collected",
        8: "LPC Fully Collected__LPC pending",
        9: "LPC Fully Collected__No. of loans",
        10: "No Arrear & Outstanding but LPC Pending__LPC Due",
        11: "No Arrear & Outstanding but LPC Pending__LPC Collected",
        12: "No Arrear & Outstanding but LPC Pending__LPC pending",
        13: "No Arrear & Outstanding but LPC Pending__No. of loans",
        14: "No Arrear but LPC Pending__LPC Due",
        15: "No Arrear but LPC Pending__LPC Collected",
        16: "No Arrear but LPC Pending__LPC pending",
        17: "No Arrear but LPC Pending__No. of loans",
        18: "Total LPC Due",
        19: "Total LPC Collected",
        20: "Total LPC pending",
        21: "Total No.of loans",
    }
    for col_idx, field in col_map.items():
        v = summary_df[field].sum()
        cell = ws.cell(r, col_idx, float(v))
        cell.font = font_bold
        cell.alignment = align_right
        cell.number_format = "0" if col_idx in [5, 9, 13, 17, 21] else "#,##0"

    fill_range(ws, r, 1, r, 21, fill_grand)
    apply_border_range(ws, r, 1, r, 21)
    r_end = r

    # Thick outlines
    set_thick_block_outline(ws, group_row, 1, r_end, 21)
    set_thick_block_outline(ws, group_row, 2, r_end, 5)
    set_thick_block_outline(ws, group_row, 6, r_end, 9)
    set_thick_block_outline(ws, group_row, 10, r_end, 13)
    set_thick_block_outline(ws, group_row, 14, r_end, 17)
    set_thick_block_outline(ws, group_row, 18, r_end, 21)

    # Freeze / widths
    ws.freeze_panes = "B4"
    ws.column_dimensions["A"].width = 26

    # Simple width set (limited scan for speed)
    for col in range(1, 22):
        letter = get_column_letter(col)
        ws.column_dimensions[letter].width = max(ws.column_dimensions[letter].width or 12, 12)

    wb.save(out_path)
    return df, summary_df


# ============================================================
# TAB 1: FULL AUTOMATION
# ============================================================
if selected_report == "1) Collection Efficiency Report":
    st.subheader("🏦 Collection Efficiency")
    st.caption("Upload 4 files → Run → Download DVC.xlsx (with Arrear_Advance sheet) + Hub-wise ZIP")

    c1, c2 = st.columns(2)
    with c1:
        ce_file = st.file_uploader(
            "1) Upload Collection Efficiency.xlsx",
            type=["xlsx"],
            help="Must contain a column named 'Hub' in at least one sheet.",
            key="ce"
        )
        jlg_file = st.file_uploader("2) Upload DVC JLG.xlsx", type=["xlsx"], key="jlg")
    with c2:
        il_file = st.file_uploader("3) Upload DVC IL.xlsx", type=["xlsx"], key="il")
        arrear_file = st.file_uploader(
            "4) Upload Arrear JLG & IL (ZIP/CSV/XLSX)",
            type=["zip", "csv", "xlsx"],
            help="Best: upload ZIP containing a single CSV with column 'zone_name'.",
            key="arrear"
        )

    use_csv_mode = st.checkbox("⚡ Use CSV mode for Arrear (faster)", value=True, key="csvmode")
    st.caption("Tip: If Arrear CSV is >200MB, ZIP it and upload the .zip (usually <200MB).")

    all_uploaded = all([ce_file, jlg_file, il_file, arrear_file])
    run_btn = st.button("🚀 Run Automation", disabled=not all_uploaded, use_container_width=True, key="run_full")

    if run_btn:
        try:
            progress = st.progress(0)
            status = st.empty()
            timer_box = st.empty()
            t0 = time.time()

            with tempfile.TemporaryDirectory() as workdir:
                ce_path = os.path.join(workdir, "Collection Efficiency.xlsx")
                jlg_path = os.path.join(workdir, "DVC JLG.xlsx")
                il_path = os.path.join(workdir, "DVC IL.xlsx")
                arrear_ext = os.path.splitext(arrear_file.name)[1].lower()
                arrear_path = os.path.join(workdir, f"Arrear{arrear_ext}")

                status.write("Saving uploaded files...")
                save_uploaded_file(ce_file, ce_path)
                save_uploaded_file(jlg_file, jlg_path)
                save_uploaded_file(il_file, il_path)
                save_uploaded_file(arrear_file, arrear_path)

                progress.progress(10)
                timer_box.write(f"Elapsed: {time.time()-t0:.1f}s")

                status.write("Creating DVC consolidated file...")
                consolidated_df = build_dvc_consolidate(jlg_path, il_path)

                status.write("Creating Arrear_Advance sheet (cust_id pivot)...")
                arrear_adv_df = build_arrear_advance_sheet(consolidated_df)

                dvc_output_path = os.path.join(workdir, "DVC.xlsx")
                with pd.ExcelWriter(dvc_output_path, engine="openpyxl") as writer:
                    consolidated_df.to_excel(writer, sheet_name="DVC_consolidate", index=False)
                    arrear_adv_df.to_excel(writer, sheet_name="Arrear_Advance", index=False)

                progress.progress(35)
                timer_box.write(f"Elapsed after DVC write: {time.time()-t0:.1f}s")

                status.write("Reading Collection Efficiency...")
                ce_sheets = read_ce_sheets(ce_path)

                hub_list = None
                for _, d in ce_sheets.items():
                    if "Hub" in d.columns:
                        hub_list = d["Hub"].dropna().unique()
                        break
                if hub_list is None or len(hub_list) == 0:
                    raise ValueError("No Hub values found in Collection Efficiency.xlsx (column 'Hub').")

                progress.progress(45)

                status.write("Preparing Arrear data...")
                arr_df = read_arrear_any(arrear_path, arrear_file.name, use_csv_mode)
                arr_df.columns = arr_df.columns.astype(str).str.strip()

                zone_col = None
                for c in arr_df.columns:
                    if c.strip().lower() == "zone_name":
                        zone_col = c
                        break
                if zone_col is None:
                    raise ValueError("Arrear file: Column 'zone_name' not found.")

                progress.progress(55)
                timer_box.write(f"Elapsed after Arrear ready: {time.time()-t0:.1f}s")

                status.write("Creating Hub-wise folders/files...")
                hubwise_root = os.path.join(workdir, "Hub wise")
                os.makedirs(hubwise_root, exist_ok=True)

                total_hubs = len(hub_list)
                dvc_hub_col = "hub" if "hub" in consolidated_df.columns else ("Hub" if "Hub" in consolidated_df.columns else None)

                for i, hub in enumerate(hub_list, start=1):
                    hub_folder = os.path.join(hubwise_root, clean_name_for_folder(hub))
                    os.makedirs(hub_folder, exist_ok=True)

                    wb_ce = Workbook()
                    wb_ce.remove(wb_ce.active)
                    ce_written = 0

                    for sh_name, df_sh in ce_sheets.items():
                        if "Hub" not in df_sh.columns:
                            continue

                        filtered = df_sh[df_sh["Hub"] == hub].copy()
                        if filtered.empty:
                            continue

                        filtered = recompute_grand_total(filtered, hub_col="Hub")

                        ws_ = wb_ce.create_sheet(title=safe_sheet_title(sh_name))
                        fmt_map = build_format_map(filtered, hub_col="Hub")
                        write_df_fast(ws_, filtered, fmt_map)
                        ce_written += 1

                    if ce_written == 0:
                        ws_ = wb_ce.create_sheet(title="Info")
                        ws_["A1"] = f"No Collection Efficiency data for HUB = {hub}"
                        ws_["A1"].font = Font(bold=True)

                    wb_ce.save(os.path.join(hub_folder, "Collection Efficiency.xlsx"))

                    wb_dvc = Workbook()
                    wb_dvc.remove(wb_dvc.active)

                    if dvc_hub_col is None:
                        ws_info = wb_dvc.create_sheet("Info")
                        ws_info["A1"] = "DVC output does not have hub column."
                        ws_info["A1"].font = Font(bold=True)
                    else:
                        dvc_hub_df = consolidated_df[
                            consolidated_df[dvc_hub_col].astype(str).str.strip() == str(hub).strip()
                        ].copy()

                        ws1 = wb_dvc.create_sheet("DVC_consolidate")
                        if dvc_hub_df.empty:
                            ws1["A1"] = f"No DVC data for HUB = {hub}"
                            ws1["A1"].font = Font(bold=True)
                        else:
                            fmt_map = build_format_map(dvc_hub_df, hub_col=dvc_hub_col)
                            write_df_fast(ws1, dvc_hub_df, fmt_map)

                        ws2 = wb_dvc.create_sheet("Arrear_Advance")
                        if dvc_hub_df.empty:
                            ws2["A1"] = "No data"
                            ws2["A1"].font = Font(bold=True)
                        else:
                            hub_adv = build_arrear_advance_sheet(dvc_hub_df)
                            for row in dataframe_to_rows(hub_adv, index=False, header=True):
                                ws2.append(row)
                            for cell in ws2[1]:
                                cell.font = Font(bold=True)
                            ws2.freeze_panes = "A2"
                            ws2.auto_filter.ref = ws2.dimensions

                    wb_dvc.save(os.path.join(hub_folder, "DVC.xlsx"))

                    ar_hub_df = arr_df[arr_df[zone_col].astype(str).str.strip() == str(hub).strip()].copy()
                    arrear_out_csv = os.path.join(hub_folder, "Arrear JLG & IL.csv")

                    if ar_hub_df.empty:
                        pd.DataFrame(columns=arr_df.columns).to_csv(arrear_out_csv, index=False, encoding="utf-8-sig")
                    else:
                        ar_hub_df.to_csv(arrear_out_csv, index=False, encoding="utf-8-sig")

                    progress.progress(55 + int(40 * (i / max(total_hubs, 1))))

                status.write("Packaging ZIP...")
                zip_path = os.path.join(workdir, "Hub_Wise_Output.zip")
                zip_folder(hubwise_root, zip_path)
                progress.progress(98)

                with open(dvc_output_path, "rb") as f:
                    dvc_bytes = f.read()
                with open(zip_path, "rb") as f:
                    zip_bytes = f.read()

            progress.progress(100)
            status.empty()
            timer_box.empty()
            st.success("✅ Done. Download outputs below.")

            d1, d2 = st.columns(2)
            with d1:
                st.download_button(
                    "⬇️ Download DVC.xlsx (with Arrear_Advance sheet)",
                    data=dvc_bytes,
                    file_name="DVC.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            with d2:
                st.download_button(
                    "⬇️ Download Hub-wise Output (ZIP)",
                    data=zip_bytes,
                    file_name="Hub_Wise_Output.zip",
                    mime="application/zip",
                    use_container_width=True
                )

        except Exception as e:
            st.error("Error occurred. Full details below:")
            st.exception(e)


# ============================================================
# TAB 2: ONLY DVC CONSOLIDATE + ARREAR ADVANCE REPORT
# ============================================================
if selected_report == "2) Arrear Advance Report":
    st.subheader("🏦 Arrear Advance Report")
    st.caption("Upload only JLG + IL → Download DVC.xlsx (DVC_consolidate + Arrear_Advance)")

    c1, c2 = st.columns(2)
    with c1:
        jlg2 = st.file_uploader("Upload DVC JLG.xlsx", type=["xlsx"], key="jlg2")
    with c2:
        il2 = st.file_uploader("Upload DVC IL.xlsx", type=["xlsx"], key="il2")

    run2 = st.button("🚀 Generate Report", disabled=not (jlg2 and il2), use_container_width=True, key="run2")

    if run2:
        try:
            progress = st.progress(0)
            status = st.empty()
            t0 = time.time()

            with tempfile.TemporaryDirectory() as workdir:
                jlg_path = os.path.join(workdir, "DVC JLG.xlsx")
                il_path = os.path.join(workdir, "DVC IL.xlsx")

                status.write("Saving uploaded files...")
                save_uploaded_file(jlg2, jlg_path)
                save_uploaded_file(il2, il_path)
                progress.progress(20)

                status.write("Building DVC consolidate...")
                consolidated_df = build_dvc_consolidate(jlg_path, il_path)
                progress.progress(60)

                status.write("Building Arrear_Advance (cust_id pivot)...")
                arrear_adv_df = build_arrear_advance_sheet(consolidated_df)
                progress.progress(85)

                out_path = os.path.join(workdir, "DVC.xlsx")
                with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
                    consolidated_df.to_excel(writer, sheet_name="DVC_consolidate", index=False)
                    arrear_adv_df.to_excel(writer, sheet_name="Arrear_Advance", index=False)

                with open(out_path, "rb") as f:
                    out_bytes = f.read()

            progress.progress(100)
            status.empty()
            st.success(f"✅ Report generated in {time.time()-t0:.1f}s")

            st.download_button(
                "⬇️ Download DVC.xlsx (with Arrear_Advance sheet)",
                data=out_bytes,
                file_name="DVC.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

        except Exception as e:
            st.error("Error occurred. Full details below:")
            st.exception(e)


# ============================================================
# TAB 3: LPC REPORT (NEW MODULE)
# ============================================================
if selected_report == "3) LPC Report":

    st.subheader("🏦LPC Report")
    st.caption("Upload LPC raw Excel → Download output with 'Summary' (formatted) + 'Member wise data' (filtered)")

    lpc_file = st.file_uploader("Upload LPC raw report (Excel)", type=["xlsx"], key="lpc_raw")
    run_lpc = st.button("🚀 Generate LPC Report", disabled=not bool(lpc_file), use_container_width=True, key="run_lpc")

    if run_lpc:
        try:
            progress = st.progress(0)
            status = st.empty()
            t0 = time.time()

            with tempfile.TemporaryDirectory() as workdir:
                in_path = os.path.join(workdir, "LPC_Raw.xlsx")
                save_uploaded_file(lpc_file, in_path)
                progress.progress(15)

                status.write("Running LPC filters + Summary formatting...")
                out_path = os.path.join(workdir, "LPC_filtered.xlsx")

                df_f, _ = lpc_generate_report(in_path, out_path, sheet_name=0)
                progress.progress(85)

                with open(out_path, "rb") as f:
                    out_bytes = f.read()

            progress.progress(100)
            status.empty()
            st.success(f"✅ LPC report generated. Filtered rows: {len(df_f)} | Time: {time.time()-t0:.1f}s")

            st.download_button(
                "⬇️ Download LPC_filtered.xlsx",
                data=out_bytes,
                file_name="LPC_filtered.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

        except Exception as e:
            st.error("Error occurred. Full details below:")
            st.exception(e)
if selected_report == "4) Vault Cash Report":

    st.subheader("🏦 Vault Cash Report")

    uploaded = st.file_uploader("Upload Vault Cash Report.xlsx", type=["xlsx"], key="vault_cash")

    if uploaded:
        import tempfile, os

        temp_dir = tempfile.mkdtemp()
        temp_path = os.path.join(temp_dir, uploaded.name)

        with open(temp_path, "wb") as f:
            f.write(uploaded.getbuffer())

        if st.button("Run Vault Cash Automation", key="run_vault"):
            try:
                run_vault_cash_report(temp_path)
                st.success("✅ Vault Cash Report generated successfully")

                with open(temp_path, "rb") as f:
                    st.download_button(
                        "⬇️ Download Updated Vault Cash Report",
                        data=f,
                        file_name="Vault Cash Report - Updated.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
            except Exception as e:
                st.error("Error occurred. Full details below:")
                st.exception(e)
if selected_report == "5) Pending Collection Entry":

    # Pending Collection Entry UI only
    st.subheader("🏦 Pending Collection Entry")

    uploaded = st.file_uploader(
        "Upload Pending Collection Entry.xlsx",
        type=["xlsx"],
        key="pending_collection"
    )
    if uploaded:
        import tempfile, os

        temp_dir = tempfile.mkdtemp()
        temp_path = os.path.join(temp_dir, uploaded.name)

        with open(temp_path, "wb") as f:
            f.write(uploaded.getbuffer())

        if st.button("Run Pending Collection Automation", key="run_pending_collection"):
            try:
                run_pending_collection_entry(temp_path)
                st.success("✅ Summary sheet created successfully")

                with open(temp_path, "rb") as f:
                    st.download_button(
                        "⬇️ Download Updated File",
                        data=f,
                        file_name="Pending Collection Entry - Updated.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
            except Exception as e:
                st.error("Error occurred. Full details below:")
                st.exception(e)
if selected_report == "6) Demand Verification Report":

    st.subheader("🏦 Demand Verification Report")

    uploaded = st.file_uploader(
        "Upload Demand Verification.xlsx",
        type=["xlsx"],
        key="demand_verification"
    )

    if uploaded:
        import tempfile, os

        temp_dir = tempfile.mkdtemp()
        temp_path = os.path.join(temp_dir, uploaded.name)

        with open(temp_path, "wb") as f:
            f.write(uploaded.getbuffer())

        if st.button("Run Demand Verification", key="run_demand_verification"):
            try:
                run_demand_verification(temp_path)
                st.success("✅ Demand Verification generated successfully")

                with open(temp_path, "rb") as f:
                    st.download_button(
                        "⬇️ Download Updated Demand Verification",
                        data=f,
                        file_name="Demand Verification - Updated.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
            except Exception as e:
                st.error("❌ Error occurred. Full details below:")
                st.exception(e)
if selected_report == "7) MAP Ledger Difference":

    st.subheader("📌 MAP vs Ledger – Difference Report")

    map_file = st.file_uploader("Upload MAP.xlsx (must have Consolidated sheet)", type=["xlsx"], key="map_xlsx")
    ledger_file = st.file_uploader("Upload MAP Ledger.xlsx", type=["xlsx"], key="map_ledger_xlsx")

    if map_file and ledger_file:
        import tempfile

        temp_dir = tempfile.mkdtemp()

        map_path = os.path.join(temp_dir, map_file.name)
        ledger_path = os.path.join(temp_dir, ledger_file.name)

        with open(map_path, "wb") as f:
            f.write(map_file.getbuffer())

        with open(ledger_path, "wb") as f:
            f.write(ledger_file.getbuffer())

        if st.button("Run MAP Ledger Difference", key="run_map_diff"):
            try:
                run_map_ledger_difference(map_path, ledger_path)
                st.success("✅ Difference sheet generated in MAP Ledger file!")

                with open(ledger_path, "rb") as f:
                    st.download_button(
                        "⬇️ Download Updated MAP Ledger.xlsx",
                        data=f,
                        file_name="MAP Ledger - Updated.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
            except Exception as e:
                st.error("❌ Error occurred. Details below:")
                st.exception(e)
if selected_report == "8) Excess Amount":
    st.subheader("💰 Excess Amount Received")

    il = st.file_uploader("Upload Repayment Summary IL.xlsx", type=["xlsx"], key="ex_il")
    jlg = st.file_uploader("Upload Repayment Summary JLG.xlsx", type=["xlsx"], key="ex_jlg")
    esc = st.file_uploader("Upload Escalation.xlsx", type=["xlsx"], key="ex_esc")

    if il and jlg and esc:
        import tempfile, os

        temp_dir = tempfile.mkdtemp()

        il_path = os.path.join(temp_dir, "Repayment Summary IL.xlsx")
        jlg_path = os.path.join(temp_dir, "Repayment Summary JLG.xlsx")
        esc_path = os.path.join(temp_dir, "Escalation.xlsx")

        with open(il_path, "wb") as f:
            f.write(il.getbuffer())
        with open(jlg_path, "wb") as f:
            f.write(jlg.getbuffer())
        with open(esc_path, "wb") as f:
            f.write(esc.getbuffer())

        if st.button("Run Excess Amount Report"):
            try:
                out_file = run_from_streamlit(temp_dir)

                st.success("Report generated successfully!")

                with open(out_file, "rb") as f:
                    st.download_button(
                        "⬇️ Download Consolidated Report",
                        data=f,
                        file_name=os.path.basename(out_file),
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
            except Exception as e:
                st.error("Error occurred")
                st.exception(e)
if selected_report == "9) CPP Payable":

    st.subheader("📑 CPP Payable vs CPP Ledger")

    cpp_payable = st.file_uploader(
        "Upload CPP Payable.xlsx",
        type=["xlsx"],
        key="cpp_payable"
    )

    cpp_ledger = st.file_uploader(
        "Upload CPP Ledger.xlsx",
        type=["xlsx"],
        key="cpp_ledger"
    )

    if cpp_payable and cpp_ledger:
        import tempfile, os

        temp_dir = tempfile.mkdtemp()
        cpp_payable_path = os.path.join(temp_dir, cpp_payable.name)
        cpp_ledger_path = os.path.join(temp_dir, cpp_ledger.name)

        with open(cpp_payable_path, "wb") as f:
            f.write(cpp_payable.getbuffer())

        with open(cpp_ledger_path, "wb") as f:
            f.write(cpp_ledger.getbuffer())

        if st.button("Run CPP Difference Report", key="run_cpp_diff"):
            try:
                run_cpp_payable_vs_cpp_ledger_difference(
                    cpp_payable_path,
                    cpp_ledger_path
                )

                st.success("✅ CPP Difference report generated successfully")

                with open(cpp_ledger_path, "rb") as f:
                    st.download_button(
                        "⬇️ Download Updated CPP Ledger.xlsx",
                        data=f,
                        file_name="CPP Ledger - Updated.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )

            except Exception as e:
                st.error("Error occurred while processing")
                st.exception(e)
if selected_report == "10) CMS and UPI Recon":

    st.subheader("🏦 CMS Recon")

    cms_format = st.file_uploader("Upload CMS Recon Sheet.xlsx", type=["xlsx"], key="cms_format")
    cms_ledger = st.file_uploader("Upload CMS Ledger.xlsx", type=["xlsx"], key="cms_ledger")
    cms_zip = st.file_uploader("Upload BS Folder (ZIP of all statements)", type=["zip"], key="cms_zip")

    if cms_format and cms_ledger and cms_zip:
        import tempfile, os, zipfile

        temp_dir = tempfile.mkdtemp()

        # Save uploaded files
        format_path = os.path.join(temp_dir, "CMS_Recon_Sheet.xlsx")
        ledger_path = os.path.join(temp_dir, "CMS_Ledger.xlsx")
        zip_path = os.path.join(temp_dir, "BS.zip")

        with open(format_path, "wb") as f:
            f.write(cms_format.getbuffer())

        with open(ledger_path, "wb") as f:
            f.write(cms_ledger.getbuffer())

        with open(zip_path, "wb") as f:
            f.write(cms_zip.getbuffer())

        # Extract ZIP -> statements folder
        statements_folder = os.path.join(temp_dir, "BS")
        os.makedirs(statements_folder, exist_ok=True)

        with zipfile.ZipFile(zip_path, "r") as z:
            z.extractall(statements_folder)

        # Handle case where ZIP contains an inner "BS" folder
        inner_bs = os.path.join(statements_folder, "BS")
        if os.path.isdir(inner_bs):
            statements_folder = inner_bs

        output_path = os.path.join(temp_dir, "Consolidate CMS Recon.xlsx")

        if st.button("Run CMS Recon", key="run_cms_recon"):
            try:
                run_cms_recon_streamlit(format_path, statements_folder, ledger_path, output_path)

                st.success("✅ CMS Recon created successfully")

                with open(output_path, "rb") as f:
                    st.download_button(
                        "⬇️ Download Consolidate CMS Recon.xlsx",
                        data=f,
                        file_name="Consolidate CMS Recon.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )

            except Exception as e:
                st.error("❌ Error occurred. Details below:")
                st.exception(e)
