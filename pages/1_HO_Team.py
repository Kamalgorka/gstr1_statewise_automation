import streamlit as st
import os
import zipfile
import pandas as pd
from openpyxl import load_workbook, Workbook
import warnings
import re
import shutil
import time
from datetime import datetime
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.worksheet.table import Table, TableStyleInfo
import tempfile

from ui import load_global_css

load_global_css()

warnings.filterwarnings("ignore", category=UserWarning)
st.set_page_config(
    page_title="HO Team Automations",
    page_icon="📊",
    layout="wide"
)

# ======================================================
# =================== ✅ GST CONFIG =====================
# ======================================================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

UPLOAD_FOLDER = os.path.join("/tmp", "uploads")
OUTPUT_ROOT = os.path.join("/tmp", "outputs")
TEMPLATE_FILE = os.path.join(BASE_DIR, "..", "GSTR1_Excel_Workbook_Template.xlsx")
LOG_FILE = os.path.join(OUTPUT_ROOT, "audit_log.xlsx")

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_ROOT, exist_ok=True)

if not os.path.exists(TEMPLATE_FILE):
    st.error(f"❌ Template file not found in app folder: {TEMPLATE_FILE}")
    st.stop()


def log_activity(action, month, filename):
    now = datetime.now()

    if not os.path.exists(LOG_FILE):
        wb = Workbook()
        ws = wb.active
        ws.title = "Audit Log"
        ws.append(["Date", "Time", "Action", "Month", "File Name", "Machine Name"])
        wb.save(LOG_FILE)

    wb = load_workbook(LOG_FILE)
    ws = wb.active

    ws.append([
        now.strftime("%d-%m-%Y"),
        now.strftime("%H:%M:%S"),
        action,
        month,
        filename,
        os.environ.get("COMPUTERNAME", "Unknown")
    ])

    wb.save(LOG_FILE)


STATE_LIST = [
    "PUNJAB", "RAJASTHAN", "HARYANA", "CHANDIGARH", "BIHAR",
    "UTTAR PRADESH", "HIMACHAL PRADESH", "JHARKHAND", "GUJARAT",
    "UTTARAKHAND", "MADHYA PRADESH", "WEST BENGAL", "ODISHA",
    "JAMMU & KASHMIR", "ANDHRA PRADESH", "CHHATTISGARH", "ASSAM", "MAHARASHTRA"
]

STATE_CODE_MAP = {
    "JAMMU & KASHMIR": "01-Jammu & Kashmir",
    "HIMACHAL PRADESH": "02-Himachal Pradesh",
    "PUNJAB": "03-Punjab",
    "CHANDIGARH": "04-Chandigarh",
    "UTTARAKHAND": "05-Uttarakhand",
    "HARYANA": "06-Haryana",
    "DELHI": "07-Delhi",
    "RAJASTHAN": "08-Rajasthan",
    "UTTAR PRADESH": "09-Uttar Pradesh",
    "BIHAR": "10-Bihar",
    "SIKKIM": "11-Sikkim",
    "ARUNACHAL PRADESH": "12-Arunachal Pradesh",
    "NAGALAND": "13-Nagaland",
    "MANIPUR": "14-Manipur",
    "MIZORAM": "15-Mizoram",
    "TRIPURA": "16-Tripura",
    "MEGHALAYA": "17-Meghalaya",
    "ASSAM": "18-Assam",
    "WEST BENGAL": "19-West Bengal",
    "JHARKHAND": "20-Jharkhand",
    "ODISHA": "21-Odisha",
    "CHHATTISGARH": "22-Chhattisgarh",
    "MADHYA PRADESH": "23-Madhya Pradesh",
    "GUJARAT": "24-Gujarat",
    "DAMAN & DIU": "25-Daman & Diu",
    "DADRA & NAGAR HAVELI & DAMAN & DIU": "26-Dadra & Nagar Haveli & Daman & Diu",
    "MAHARASHTRA": "27-Maharashtra",
    "KARNATAKA": "29-Karnataka",
    "GOA": "30-Goa",
    "LAKSHADWEEP": "31-Lakshadweep",
    "KERALA": "32-Kerala",
    "TAMIL NADU": "33-Tamil Nadu",
    "PUDUCHERRY": "34-Puducherry",
    "ANDAMAN & NICOBAR ISLANDS": "35-Andaman & Nicobar Islands",
    "TELANGANA": "36-Telangana",
    "ANDHRA PRADESH": "37-Andhra Pradesh",
    "LADAKH": "38-Ladakh",
    "FOREIGN COUNTRY": "96-Foreign Country",
    "OTHER TERRITORY": "97-Other Territory"
}


def format_place_of_supply(state_val):
    if pd.isna(state_val):
        return ""
    key = str(state_val).strip().upper()
    return STATE_CODE_MAP.get(key, state_val)


def clean_hsn(x):
    x = str(x).strip()
    if x.endswith(".0"):
        x = x[:-2]
    return x


def filter_state_month(df, state_upper, month_str):
    if "State" not in df.columns:
        return df.iloc[0:0]
    s = df["State"].astype(str).str.upper().str.strip()
    mask = (s == state_upper)
    if "Month" in df.columns:
        m = df["Month"].astype(str).str.lower().str.strip()
        mask = mask & m.str.contains(month_str.lower(), na=False)
    return df[mask]


def extract_last_num(x):
    nums = re.findall(r'\d+', str(x))
    return int(nums[-1]) if nums else 0


REQUIRED_SHEETS = [
    "GSTR1", "B2B", "CD Note Reg", "DN Note Reg", "Exempt Income",
    "B2C PF", "CD Note Unreg", "B2C Onboarding"
]

REQUIRED_COLUMNS = {
    "B2B": ["State", "Invoice Type", "Invoice Number", "Invoice Date", "Total Transaction Value", "State Place of Supply"],
    "CD Note Reg": ["State", "Month"],
    "DN Note Reg": ["State", "Month"],
    "B2C PF": ["State", "Month", "LPF"],
    "CD Note Unreg": ["State", "Month"],
    "Exempt Income": ["Row Labels", "Sum of Collection Intrest", "Month"],
}


def validate_excel(file_path: str):
    errors = []
    warnings_list = []

    try:
        xl = pd.ExcelFile(file_path)
        sheet_names = xl.sheet_names
    except Exception as e:
        return [f"❌ Unable to read Excel file. Error: {e}"], []

    missing_sheets = [s for s in REQUIRED_SHEETS if s not in sheet_names]
    if missing_sheets:
        errors.append("❌ Missing required sheet(s): " + ", ".join(missing_sheets))

    if errors:
        return errors, warnings_list

    for sh, cols in REQUIRED_COLUMNS.items():
        try:
            df_head = pd.read_excel(file_path, sheet_name=sh, nrows=1)
            missing_cols = [c for c in cols if c not in df_head.columns]
            if missing_cols:
                errors.append(f"❌ Sheet '{sh}' missing column(s): {', '.join(missing_cols)}")
        except Exception as e:
            errors.append(f"❌ Could not read sheet '{sh}'. Error: {e}")

    try:
        df_states = pd.read_excel(file_path, sheet_name="B2B", usecols=["State"])
        if df_states["State"].dropna().empty:
            warnings_list.append("⚠️ 'B2B' sheet has empty 'State' column. State-wise output may be blank.")
    except Exception:
        pass

    return errors, warnings_list


@st.cache_data(show_spinner=False)
def read_all_sheets_cached(file_bytes: bytes):
    from io import BytesIO
    bio = BytesIO(file_bytes)

    df_gstr1 = pd.read_excel(bio, sheet_name="GSTR1")
    bio.seek(0)
    df_b2b = pd.read_excel(bio, sheet_name="B2B")
    bio.seek(0)
    df_cd = pd.read_excel(bio, sheet_name="CD Note Reg")
    bio.seek(0)
    df_dn = pd.read_excel(bio, sheet_name="DN Note Reg")
    bio.seek(0)
    df_exempt = pd.read_excel(bio, sheet_name="Exempt Income")
    bio.seek(0)
    df_b2c_pf = pd.read_excel(bio, sheet_name="B2C PF")
    bio.seek(0)
    df_cdunreg = pd.read_excel(bio, sheet_name="CD Note Unreg")
    bio.seek(0)
    df_b2c_onboard = pd.read_excel(bio, sheet_name="B2C Onboarding")

    if "Month" in df_b2c_pf.columns:
        df_b2c_pf["Month_norm"] = df_b2c_pf["Month"].astype(str).str.lower().str.strip()
    else:
        df_b2c_pf["Month_norm"] = ""

    return df_gstr1, df_b2b, df_cd, df_dn, df_exempt, df_b2c_pf, df_cdunreg, df_b2c_onboard


def run_gstr_process(
    df_gstr1, df_b2b, df_cd, df_dn, df_exempt,
    df_b2c_pf, df_cdunreg, df_b2c_onboard,
    MONTH, MONTH_NORM,
    TEMPLATE_FILE, OUTPUT_FOLDER,
    progress_bar=None,
    status_box=None,
    eta_box=None
):
    total_states = len(STATE_LIST)
    start_all = time.time()

    for i, state in enumerate(STATE_LIST, start=1):
        wb = load_workbook(TEMPLATE_FILE)

        for sheet_name in wb.sheetnames:
            if sheet_name.lower() not in ["exemp", "docs"]:
                ws_tmp = wb[sheet_name]
                ws_tmp.delete_rows(4)

        df_state_b2b = filter_state_month(df_b2b, state, MONTH)
        df_cd_state = filter_state_month(df_cd, state, MONTH)
        df_dn_state = filter_state_month(df_dn, state, MONTH)

        ws = wb["b2b"]
        r = 4
        for _, row in df_state_b2b.iterrows():
            if str(row["Invoice Type"]).upper() == "B2B":
                ws[f"A{r}"] = row.get("Customer Billing GSTIN", "")
                ws[f"B{r}"] = row.get("Customer Billing Name", "")
                ws[f"C{r}"] = row.get("Invoice Number", "")
                ws[f"D{r}"] = row.get("Invoice Date", "")
                ws[f"E{r}"] = row.get("Total Transaction Value", 0)
                ws[f"F{r}"] = format_place_of_supply(row.get("State Place of Supply", ""))
                ws[f"G{r}"] = "N"
                ws[f"H{r}"] = ""
                ws[f"I{r}"] = "Regular B2B"
                ws[f"J{r}"] = ""
                ws[f"K{r}"] = 18
                ws[f"L{r}"] = row.get("Item Taxable Value", row.get("Total Transaction Value", 0))
                ws[f"M{r}"] = ""
                r += 1

        ws = wb["b2cl"]
        r = 4
        for _, row in df_state_b2b.iterrows():
            if str(row["Invoice Type"]).upper() == "B2CL":
                ws[f"A{r}"] = row.get("Invoice Number", "")
                ws[f"B{r}"] = row.get("Invoice Date", "")
                ws[f"C{r}"] = row.get("Total Transaction Value", 0)
                ws[f"D{r}"] = row.get("State Place of Supply", "")
                ws[f"E{r}"] = ""
                ws[f"F{r}"] = 18
                ws[f"G{r}"] = row.get("Item Taxable Value", row.get("Total Transaction Value", 0))
                ws[f"H{r}"] = ""
                ws[f"I{r}"] = ""
                r += 1

        ws = wb["b2cs"]
        r = 4

        df_b2cs_1 = df_state_b2b[df_state_b2b["Invoice Type"].astype(str).str.upper() == "B2CS"].copy()
        if not df_b2cs_1.empty:
            df_b2cs_1["POS"] = df_b2cs_1["State Place of Supply"].astype(str).str.upper()
            df_b2cs_1["TAX"] = df_b2cs_1.get("Item Taxable Value", df_b2cs_1["Total Transaction Value"])
            df_b2cs_1 = df_b2cs_1[["POS", "TAX"]]
        else:
            df_b2cs_1 = pd.DataFrame(columns=["POS", "TAX"])

        df_b2c_pf_state = df_b2c_pf[
            (df_b2c_pf["State"].astype(str).str.upper() == state) &
            (df_b2c_pf["Month_norm"].str.contains(MONTH_NORM, na=False))
        ].copy()

        df_b2cs_2 = pd.DataFrame(columns=["POS", "TAX"])
        if not df_b2c_pf_state.empty:
            df_b2cs_2["POS"] = df_b2c_pf_state["State"].astype(str).str.upper()
            df_b2cs_2["TAX"] = df_b2c_pf_state["LPF"].fillna(0)

        df_b2cs_all = pd.concat([df_b2cs_1, df_b2cs_2], ignore_index=True)
        df_b2cs_grp = df_b2cs_all.groupby("POS", as_index=False)["TAX"].sum() if not df_b2cs_all.empty else pd.DataFrame(columns=["POS", "TAX"])

        df_cdun_b2cs = df_cdunreg[
            (df_cdunreg["State"].astype(str).str.upper() == state) &
            (df_cdunreg["Month"].astype(str).str.lower().str.contains(MONTH_NORM, na=False))
        ].copy()

        df_cdun_b2cs_adj = pd.DataFrame(columns=["POS", "CD_TAX"])
        if not df_cdun_b2cs.empty:
            df_cdun_b2cs["POS"] = df_cdun_b2cs["State"].astype(str).str.upper()
            df_cdun_b2cs_adj = (
                df_cdun_b2cs
                .groupby("POS", as_index=False)["Item Taxable Value"]
                .sum()
                .rename(columns={"Item Taxable Value": "CD_TAX"})
            )

        df_b2cs_final = df_b2cs_grp.merge(df_cdun_b2cs_adj, on="POS", how="left").fillna(0) if not df_b2cs_grp.empty else pd.DataFrame(columns=["POS", "TAX", "CD_TAX"])
        if not df_b2cs_final.empty:
            df_b2cs_final["TAX"] = (df_b2cs_final["TAX"] - df_b2cs_final["CD_TAX"]).clip(lower=0)

            for _, row in df_b2cs_final.iterrows():
                ws[f"A{r}"] = "OE"
                ws[f"B{r}"] = format_place_of_supply(row["POS"])
                ws[f"C{r}"] = ""
                ws[f"D{r}"] = 18
                ws[f"E{r}"] = row["TAX"]
                ws[f"F{r}"] = ""
                ws[f"G{r}"] = ""
                ws[f"H{r}"] = ""
                r += 1

        ws = wb["cdnr"]
        r = 4
        notes_df = pd.concat([df_cd_state, df_dn_state], ignore_index=True)
        for _, row in notes_df.iterrows():
            note_type_full = str(row.get("Credit(C)/Debit(D) Note Type", "")).strip().upper()
            if note_type_full.startswith("C"):
                mapped_type = "C"
            elif note_type_full.startswith("D"):
                mapped_type = "D"
            else:
                mapped_type = ""

            gstin = row.get("Customer Billing GSTIN", row.get("My GSTIN", ""))
            receiver_name = row.get("Customer Billing Name", "")
            note_number = row.get("Credit/Debit Note Number", "")
            note_date = row.get("Credit/Debit Note Date", "")
            pos = row.get("State Place of Supply", row.get("State", state))
            total_val = row.get("Total Transaction Value", row.get("Invoice Amount", 0))
            taxable_val = row.get("Item Taxable Value", total_val)

            ws[f"A{r}"] = gstin
            ws[f"B{r}"] = receiver_name
            ws[f"C{r}"] = note_number
            ws[f"D{r}"] = note_date
            ws[f"E{r}"] = mapped_type
            ws[f"F{r}"] = format_place_of_supply(pos)
            ws[f"G{r}"] = "N"
            ws[f"H{r}"] = "Regular"
            ws[f"I{r}"] = total_val
            ws[f"J{r}"] = ""
            ws[f"K{r}"] = 18
            ws[f"L{r}"] = taxable_val
            ws[f"M{r}"] = ""
            r += 1

        ws = wb["exemp"]
        exempt_val = None
        if "Row Labels" in df_exempt.columns and "Sum of Collection Intrest" in df_exempt.columns:
            mask_state = df_exempt["Row Labels"].astype(str).str.upper().str.strip() == state
            if "Month" in df_exempt.columns:
                m = df_exempt["Month"].astype(str).str.lower().str.strip()
                mask_month = m.str.contains(MONTH_NORM, na=False)
                mask = mask_state & mask_month
            else:
                mask = mask_state

            if mask.any():
                exempt_val = df_exempt.loc[mask, "Sum of Collection Intrest"].sum()

        for col in ["B", "C", "D"]:
            for row_idx in range(4, 8):
                ws[f"{col}{row_idx}"] = ""
        ws["C7"] = float(exempt_val) if (exempt_val and not pd.isna(exempt_val) and exempt_val != 0) else ""

        out_path = os.path.join(OUTPUT_FOLDER, f"GSTR1_{MONTH}_{state}.xlsx")
        os.makedirs(OUTPUT_FOLDER, exist_ok=True)
        wb.save(out_path)

        if status_box is not None:
            status_box.info(f"✅ Completed: {state}  ({i}/{len(STATE_LIST)})")
        if progress_bar is not None:
            progress_bar.progress(i / len(STATE_LIST))
        if eta_box is not None:
            elapsed = time.time() - start_all
            avg_per_state = elapsed / max(i, 1)
            remaining = avg_per_state * (len(STATE_LIST) - i)
            eta_box.caption(f"⏳ Estimated remaining processing time: ~{int(remaining)} sec")

    return


# ======================================================
# =================== ✅ DAY BOOK CORE ==================
# ======================================================
TRANSACTION_WISE_STATEMENTS = {"ICICI 7021"}

BANK_NAME_MAP = {
    "ICICI 7021": "ICICI Bank - 008205007021 - C",
    "ICICI 6086": "ICICI Bank - 008205006086 - C",
}

DEFAULT_DEPARTMENT = "H0001"
DEFAULT_BRANCH_ID = "H0001"

BANK_LEDGER_PAIRS = {
    ("311701", "HO0001"), ("311701", "HO0002"), ("311701", "HO0004"), ("311701", "HO0005"),
    ("311701", "HO0006"), ("311701", "HO0008"), ("311701", "HO0009"), ("311701", "HO0010"),
    ("311701", "HO0011"), ("311701", "HO0012"), ("311701", "HO0014"), ("311701", "HO0015"),
    ("311701", "HO0016"), ("311701", "HO0018"), ("311701", "HO0019"), ("311701", "HO0020"),
    ("311702", ""), ("311702", "HO0001"),
    ("311703", ""), ("311703", "HO0002"), ("311703", "HO0003"),
    ("311704", ""), ("311704", "HO0001"),
    ("311705", ""), ("311705", "HO0001"), ("311705", "HO0002"),
    ("311706", ""), ("311706", "HO0001"),
    ("311707", ""), ("311707", "HO0001"), ("311707", "HO0002"),
    ("311708", ""), ("311709", ""), ("311709", "HO0001"),
    ("311710", ""), ("311711", ""), ("311711", "HO0001"), ("311711", "HO0002"),
    ("311712", ""), ("311712", "HO0001"),
    ("311713", "HO0001"),
    ("311714", ""),
    ("311715", ""), ("311715", "HO0001"),
    ("311716", ""), ("311716", "HO0001"), ("311716", "HO0002"),
    ("311717", ""), ("311717", "HO0002"),
    ("311718", ""), ("311718", "HO0001"),
    ("311719", ""), ("311719", "HO0001"),
    ("311720", "HO0001"),
}


def _clean_text(x):
    if pd.isna(x):
        return ""
    return str(x).strip()


def _find_header_row_xlsx(file_path, required_headers=None, max_scan_rows=80):
    preview = pd.read_excel(file_path, header=None, nrows=max_scan_rows)
    required_headers = [h.strip().lower() for h in (required_headers or [])]
    for i in range(len(preview)):
        row_vals = preview.iloc[i].astype(str).str.strip().str.lower().tolist()
        if all(any(req == cell for cell in row_vals) for req in required_headers):
            return i
    return None


def _safe_sheet_name(name: str, max_len: int = 31) -> str:
    bad = [":", "\\", "/", "?", "*", "[", "]"]
    for b in bad:
        name = name.replace(b, " ")
    name = " ".join(name.split())
    return name[:max_len] if len(name) > max_len else name


def _build_keyword_table(coa_df: pd.DataFrame) -> pd.DataFrame:
    coa_df.columns = [c.strip() for c in coa_df.columns]
    if "Account Name" not in coa_df.columns:
        raise ValueError("COA.xlsx must have column: 'Account Name'")
    refer_cols = [c for c in coa_df.columns if c.upper().startswith("REFER")]
    if not refer_cols:
        raise ValueError("No Refer columns found in COA (Refer1..ReferN)")

    rows = []
    for _, r in coa_df.iterrows():
        ledger = _clean_text(r.get("Account Name"))
        if not ledger:
            continue
        for col in refer_cols:
            kw = _clean_text(r.get(col))
            if kw:
                rows.append({"ledger": ledger, "keyword": kw})

    kw_df = pd.DataFrame(rows)
    if kw_df.empty:
        raise ValueError("No keywords found in Refer columns.")
    kw_df["kw_len"] = kw_df["keyword"].str.len()
    return kw_df.sort_values("kw_len", ascending=False).reset_index(drop=True)


def _map_ledger(description: str, kw_df: pd.DataFrame) -> str:
    d = str(description).upper()
    for _, r in kw_df.iterrows():
        if r["keyword"].upper() in d:
            return r["ledger"]
    return "UNMAPPED"


def _move_others_last_df(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df
    df = df.copy()
    df["_is_others"] = df["Ledger"].astype(str).str.strip().str.upper().eq("OTHERS")
    df = df.sort_values(["_is_others"], ascending=[True]).drop(columns=["_is_others"])
    return df.reset_index(drop=True)


def _make_daybook(receipts: pd.DataFrame, payments: pd.DataFrame) -> pd.DataFrame:
    receipts = receipts.reset_index(drop=True)
    payments = payments.reset_index(drop=True)
    max_len = max(len(receipts), len(payments))
    rows = []
    for i in range(max_len):
        r_led = receipts.loc[i, "Ledger"] if i < len(receipts) else ""
        r_amt = receipts.loc[i, "Amount"] if i < len(receipts) else ""
        p_led = payments.loc[i, "Ledger"] if i < len(payments) else ""
        p_amt = payments.loc[i, "Amount"] if i < len(payments) else ""
        rows.append([r_led, r_amt, p_led, p_amt])
    return pd.DataFrame(rows, columns=["RECEIPT", "AMOUNT", "PAYMENTS", "AMOUNT.1"])


def _build_coa_lookup(coa_df: pd.DataFrame):
    coa_df = coa_df.copy()
    coa_df.columns = [c.strip() for c in coa_df.columns]
    for needed in ["Account Name", "Account", "Sub Account"]:
        if needed not in coa_df.columns:
            raise ValueError(f"COA.xlsx must have column: '{needed}'")

    by_name = {}
    name_by_pair = {}
    for _, r in coa_df.iterrows():
        name = _clean_text(r["Account Name"])
        acc = _clean_text(r["Account"])
        sub = _clean_text(r["Sub Account"])
        if name:
            by_name[name] = (acc, sub)
        name_by_pair[(acc, sub)] = name
    return by_name, name_by_pair


def _pick_document_date(bank_df: pd.DataFrame):
    for col in ["Value Date", "Txn Posted Date", "Document Date"]:
        if col in bank_df.columns:
            ser = pd.to_datetime(bank_df[col], errors="coerce").dropna()
            if len(ser):
                return ser.iloc[0].date()
    return None


def _is_bank_pair(acc: str, sub: str) -> bool:
    return (_clean_text(acc), _clean_text(sub)) in BANK_LEDGER_PAIRS


def _bank_short(name: str) -> str:
    s = _clean_text(name)
    up = s.upper()
    digit_runs = re.findall(r"\d{4,}", s)
    last4 = digit_runs[-1][-4:] if digit_runs else ""
    brand = None
    for b in ["ICICI", "HDFC", "AXIS", "SBI", "IDBI", "PNB", "CANARA", "FEDERAL", "BANDHAN", "YES"]:
        if b in up:
            brand = b
            break
    if brand and last4:
        return f"{brand} {last4}"
    return last4 if last4 else s


def _cv_ref(to_bank: str, from_bank: str) -> str:
    return f"BEING ONLINE FUND TRANSFERRED TO {_bank_short(to_bank)} FROM {_bank_short(from_bank)}"


def _create_entry(daybook_df: pd.DataFrame, bank_display_name: str,
                  coa_by_name: dict, coa_name_by_pair: dict, doc_date):
    bank_acc, bank_sub = coa_by_name.get(bank_display_name, ("", ""))
    doc_date_val = doc_date.strftime("%d-%b-%y") if doc_date else ""

    rows = []

    def add_pair(journal_code, ref_text, debit_acc, debit_sub, credit_acc, credit_sub, amt):
        rows.append({
            "Journal Code": journal_code, "Sequence": 1, "Account": debit_acc, "Sub Account": debit_sub,
            "Department": DEFAULT_DEPARTMENT, "Document Date": doc_date_val,
            "Debit": amt, "Credit": "", "Supplier Id": "", "Customer Id": "", "SAC/HSN": "",
            "Reference": ref_text, "Branch Id": DEFAULT_BRANCH_ID, "Invoice Num": "", "Comments": ""
        })
        rows.append({
            "Journal Code": journal_code, "Sequence": 2, "Account": credit_acc, "Sub Account": credit_sub,
            "Department": DEFAULT_DEPARTMENT, "Document Date": doc_date_val,
            "Debit": "", "Credit": amt, "Supplier Id": "", "Customer Id": "", "SAC/HSN": "",
            "Reference": ref_text, "Branch Id": DEFAULT_BRANCH_ID, "Invoice Num": "", "Comments": ""
        })

    df = daybook_df.copy()

    # Receipt
    for _, r in df.iterrows():
        ledger = _clean_text(r.get("RECEIPT"))
        if not ledger or ledger.upper() == "TOTAL":
            continue
        try:
            amt = float(r.get("AMOUNT"))
        except Exception:
            continue
        if abs(amt) < 0.01:
            continue

        led_acc, led_sub = coa_by_name.get(ledger, ("", ""))

        # ✅ CV if ANY side is bank-pair
        is_cv = _is_bank_pair(bank_acc, bank_sub) or _is_bank_pair(led_acc, led_sub)

        if is_cv:
            other_bank = coa_name_by_pair.get((led_acc, led_sub), ledger)
            ref = _cv_ref(to_bank=bank_display_name, from_bank=other_bank)
            journal = "CV"
        else:
            ref = f"BEING {ledger} RECEIPT IN {bank_display_name}"
            journal = "BR"

        add_pair(journal, ref, bank_acc, bank_sub, led_acc, led_sub, amt)

    # Payment
    for _, r in df.iterrows():
        ledger = _clean_text(r.get("PAYMENTS"))
        if not ledger or ledger.upper() == "TOTAL":
            continue
        try:
            amt = float(r.get("AMOUNT.1"))
        except Exception:
            continue
        if abs(amt) < 0.01:
            continue

        led_acc, led_sub = coa_by_name.get(ledger, ("", ""))

        is_cv = _is_bank_pair(bank_acc, bank_sub) or _is_bank_pair(led_acc, led_sub)

        if is_cv:
            other_bank = coa_name_by_pair.get((led_acc, led_sub), ledger)
            ref = _cv_ref(to_bank=other_bank, from_bank=bank_display_name)
            journal = "CV"
        else:
            ref = f"BEING {ledger} PAYMENT FROM {bank_display_name}"
            journal = "BP"

        add_pair(journal, ref, led_acc, led_sub, bank_acc, bank_sub, amt)

    cols = [
        "Journal Code", "Sequence", "Account", "Sub Account", "Department", "Document Date",
        "Debit", "Credit", "Supplier Id", "Customer Id", "SAC/HSN", "Reference",
        "Branch Id", "Invoice Num", "Comments"
    ]
    return pd.DataFrame(rows, columns=cols)


def run_daybook_from_zip(zip_bytes: bytes, coa_bytes: bytes, output_path: str):
    coa = pd.read_excel(pd.io.common.BytesIO(coa_bytes))
    kw_df = _build_keyword_table(coa)
    coa_by_name, coa_name_by_pair = _build_coa_lookup(coa)

    wb = Workbook()
    wb.remove(wb.active)

    with tempfile.TemporaryDirectory() as tmpdir:
        zpath = os.path.join(tmpdir, "statements.zip")
        with open(zpath, "wb") as f:
            f.write(zip_bytes)

        with zipfile.ZipFile(zpath, "r") as z:
            z.extractall(tmpdir)

        files = []
        for root, _, fs in os.walk(tmpdir):
            for fn in fs:
                if fn.lower().endswith(".xlsx") and not fn.startswith("~$"):
                    files.append(os.path.join(root, fn))

        if not files:
            raise ValueError("No .xlsx statements found inside ZIP.")

        thin = Side(style="thin", color="000000")
        border = Border(left=thin, right=thin, top=thin, bottom=thin)
        header_fill = PatternFill("solid", fgColor="BFBFBF")
        amt_fill = PatternFill("solid", fgColor="FFF200")
        others_fill = PatternFill("solid", fgColor="F8CBAD")
        unmapped_fill = PatternFill("solid", fgColor="FFC7CE")
        bold = Font(bold=True)

        for fp in sorted(files):
            base = os.path.splitext(os.path.basename(fp))[0]
            bank_name = BANK_NAME_MAP.get(base, base)

            header_row = _find_header_row_xlsx(fp, required_headers=["Description", "Cr/Dr"])
            if header_row is None:
                header_row = 6

            bank = pd.read_excel(fp, header=header_row)
            bank.columns = [str(c).strip() for c in bank.columns]

            if "Description" not in bank.columns or "Cr/Dr" not in bank.columns:
                continue

            amt_candidates = [c for c in bank.columns if "Transaction Amount" in c]
            if not amt_candidates:
                continue
            amt_col = amt_candidates[0]

            bank["Cr/Dr"] = bank["Cr/Dr"].astype(str).str.strip().str.upper()
            bank[amt_col] = pd.to_numeric(bank[amt_col], errors="coerce").fillna(0)
            bank["Ledger"] = bank["Description"].astype(str).apply(lambda x: _map_ledger(x, kw_df))

            total_cr = bank.loc[bank["Cr/Dr"] == "CR", amt_col].sum()
            total_dr = bank.loc[bank["Cr/Dr"] == "DR", amt_col].sum()

            is_txn = base in TRANSACTION_WISE_STATEMENTS

            if not is_txn:
                cr = bank[bank["Cr/Dr"] == "CR"].groupby("Ledger", as_index=False)[amt_col].sum()
                cr = cr[cr["Ledger"] != "UNMAPPED"].rename(columns={amt_col: "Amount"})
                cr_diff = round(total_cr - cr["Amount"].sum(), 2) if len(cr) else round(total_cr, 2)
                if abs(cr_diff) > 0.01:
                    cr = pd.concat([cr, pd.DataFrame([{"Ledger": "Others", "Amount": cr_diff}])], ignore_index=True)

                dr = bank[bank["Cr/Dr"] == "DR"].groupby("Ledger", as_index=False)[amt_col].sum()
                dr = dr[dr["Ledger"] != "UNMAPPED"].rename(columns={amt_col: "Amount"})
                dr_diff = round(total_dr - dr["Amount"].sum(), 2) if len(dr) else round(total_dr, 2)
                if abs(dr_diff) > 0.01:
                    dr = pd.concat([dr, pd.DataFrame([{"Ledger": "Others", "Amount": dr_diff}])], ignore_index=True)

                cr = _move_others_last_df(cr.sort_values("Ledger").reset_index(drop=True))
                dr = _move_others_last_df(dr.sort_values("Ledger").reset_index(drop=True))

                daybook_df = _make_daybook(cr, dr)

            else:
                cr = bank[(bank["Cr/Dr"] == "CR") & (bank["Ledger"].str.upper() != "UNMAPPED")][["Ledger", amt_col]].copy()
                cr = cr.rename(columns={amt_col: "Amount"})
                dr = bank[(bank["Cr/Dr"] == "DR") & (bank["Ledger"].str.upper() != "UNMAPPED")][["Ledger", amt_col]].copy()
                dr = dr.rename(columns={amt_col: "Amount"})

                cr_diff = round(total_cr - cr["Amount"].sum(), 2) if len(cr) else round(total_cr, 2)
                dr_diff = round(total_dr - dr["Amount"].sum(), 2) if len(dr) else round(total_dr, 2)

                if abs(cr_diff) > 0.01:
                    cr = pd.concat([cr, pd.DataFrame([{"Ledger": "Others", "Amount": cr_diff}])], ignore_index=True)
                if abs(dr_diff) > 0.01:
                    dr = pd.concat([dr, pd.DataFrame([{"Ledger": "Others", "Amount": dr_diff}])], ignore_index=True)

                cr = _move_others_last_df(cr)
                dr = _move_others_last_df(dr)
                daybook_df = _make_daybook(cr, dr)

            mapping_cols = [c for c in ["Value Date", "Txn Posted Date"] if c in bank.columns]
            mapping_df = bank[mapping_cols + ["Description", "Cr/Dr", amt_col, "Ledger"]].copy()
            mapping_df = mapping_df.rename(columns={amt_col: "Amount"})

            doc_date = _pick_document_date(bank)
            entry_df = _create_entry(daybook_df, bank_name, coa_by_name, coa_name_by_pair, doc_date)

            db_name = _safe_sheet_name(f"DayBook_{bank_name}")
            mp_name = _safe_sheet_name(f"Mapping_{bank_name}")
            en_name = _safe_sheet_name(f"Entry_{bank_name}")

            def uniq(nm):
                if nm not in wb.sheetnames:
                    return nm
                i = 1
                while True:
                    nn = _safe_sheet_name(f"{nm}_{i}")
                    if nn not in wb.sheetnames:
                        return nn
                    i += 1

            db_name, mp_name, en_name = uniq(db_name), uniq(mp_name), uniq(en_name)

            # DayBook write
            ws = wb.create_sheet(db_name)
            ws["A1"] = bank_name
            ws["A1"].font = Font(bold=True, size=14)
            ws.merge_cells("A1:D1")
            ws["A1"].alignment = Alignment(horizontal="center")

            start_row = 3
            headers = ["RECEIPT", "AMOUNT", "PAYMENTS", "AMOUNT"]
            for c, h in enumerate(headers, start=1):
                cell = ws.cell(row=start_row, column=c, value=h)
                cell.font = bold
                cell.fill = header_fill
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.border = border

            for r in range(len(daybook_df)):
                rec_led = str(daybook_df.iloc[r, 0]).strip().upper()
                pay_led = str(daybook_df.iloc[r, 2]).strip().upper()
                row_vals = [daybook_df.iloc[r, 0], daybook_df.iloc[r, 1], daybook_df.iloc[r, 2], daybook_df.iloc[r, 3]]

                for c, val in enumerate(row_vals, start=1):
                    cell = ws.cell(row=start_row + 1 + r, column=c, value=val)
                    if c in [2, 4]:
                        cell.fill = amt_fill
                        cell.number_format = "#,##0.00"
                        cell.alignment = Alignment(horizontal="right", vertical="center")
                    else:
                        cell.alignment = Alignment(horizontal="left", vertical="center")

                    if c in [1, 2] and rec_led == "OTHERS":
                        cell.fill = others_fill
                        if c == 1:
                            cell.hyperlink = f"#'{mp_name}'!A1"
                            cell.font = Font(color="0563C1", underline="single")

                    if c in [3, 4] and pay_led == "OTHERS":
                        cell.fill = others_fill
                        if c == 3:
                            cell.hyperlink = f"#'{mp_name}'!A1"
                            cell.font = Font(color="0563C1", underline="single")

                    cell.border = border

            last_data_row = start_row + len(daybook_df)
            total_row = last_data_row + 1

            ws.cell(row=total_row, column=1, value="TOTAL").font = bold
            ws.cell(row=total_row, column=1).fill = header_fill
            ws.cell(row=total_row, column=1).border = border

            ws.cell(row=total_row, column=2, value=f"=SUM(B{start_row+1}:B{last_data_row})").font = bold
            ws.cell(row=total_row, column=2).fill = amt_fill
            ws.cell(row=total_row, column=2).number_format = "#,##0.00"
            ws.cell(row=total_row, column=2).alignment = Alignment(horizontal="right")
            ws.cell(row=total_row, column=2).border = border

            ws.cell(row=total_row, column=3, value="TOTAL").font = bold
            ws.cell(row=total_row, column=3).fill = header_fill
            ws.cell(row=total_row, column=3).border = border

            ws.cell(row=total_row, column=4, value=f"=SUM(D{start_row+1}:D{last_data_row})").font = bold
            ws.cell(row=total_row, column=4).fill = amt_fill
            ws.cell(row=total_row, column=4).number_format = "#,##0.00"
            ws.cell(row=total_row, column=4).alignment = Alignment(horizontal="right")
            ws.cell(row=total_row, column=4).border = border

            ws.column_dimensions["A"].width = 50
            ws.column_dimensions["B"].width = 18
            ws.column_dimensions["C"].width = 50
            ws.column_dimensions["D"].width = 18

            # Mapping write (only UNMAPPED visible)
            ws2 = wb.create_sheet(mp_name)
            for c, col_name in enumerate(mapping_df.columns, start=1):
                cell = ws2.cell(row=1, column=c, value=col_name)
                cell.font = bold
                cell.fill = header_fill
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.border = border

            for idx in range(len(mapping_df)):
                ledger_val = str(mapping_df.loc[idx, "Ledger"]).strip().upper()
                is_unmapped = ledger_val == "UNMAPPED"
                excel_row = idx + 2
                if not is_unmapped:
                    ws2.row_dimensions[excel_row].hidden = True

                for c, col_name in enumerate(mapping_df.columns, start=1):
                    val = mapping_df.loc[idx, col_name]
                    cell = ws2.cell(row=excel_row, column=c, value=val)
                    if col_name.upper() == "AMOUNT":
                        cell.number_format = "#,##0.00"
                        cell.alignment = Alignment(horizontal="right", vertical="center")
                    else:
                        cell.alignment = Alignment(horizontal="left", vertical="center")

                    if is_unmapped:
                        cell.fill = unmapped_fill
                    cell.border = border

            last_row = len(mapping_df) + 1
            last_col = len(mapping_df.columns)
            last_col_letter = chr(64 + last_col)
            table_ref = f"A1:{last_col_letter}{last_row}"
            tbl_name = f"Tbl_{abs(hash(mp_name)) % 10_000_000}"
            table = Table(displayName=tbl_name, ref=table_ref)
            table.tableStyleInfo = TableStyleInfo(
                name="TableStyleMedium2",
                showFirstColumn=False,
                showLastColumn=False,
                showRowStripes=False,
                showColumnStripes=False
            )
            ws2.add_table(table)
            ws2.auto_filter.ref = table_ref
            ws2.freeze_panes = "A2"

            # Entry write
            ws3 = wb.create_sheet(en_name)
            for c, col_name in enumerate(entry_df.columns, start=1):
                cell = ws3.cell(row=1, column=c, value=col_name)
                cell.font = bold
                cell.fill = header_fill
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.border = border

            for r in range(len(entry_df)):
                excel_row = r + 2
                for c, col_name in enumerate(entry_df.columns, start=1):
                    val = entry_df.iloc[r][col_name]
                    cell = ws3.cell(row=excel_row, column=c, value=val)
                    if col_name in ["Debit", "Credit"] and val not in ["", None]:
                        cell.number_format = "#,##0.00"
                        cell.alignment = Alignment(horizontal="right", vertical="center")
                    else:
                        cell.alignment = Alignment(horizontal="left", vertical="center")
                    cell.border = border

            ws3.freeze_panes = "A2"

        wb.save(output_path)
        return output_path


# ======================================================
# ====================== ✅ UI ==========================
# ======================================================
st.markdown("""
<style>
.big-title { font-size: 30px; font-weight: 700; color: #0A3D62; }
.sub-title { font-size: 16px; color: #576574; }
.card { background-color: #F8F9FA; padding: 20px; border-radius: 12px;
        box-shadow: 0px 2px 6px rgba(0,0,0,0.08); }
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="big-title">📌 HO Team Automations</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title">GSTR-1 + HO Day Book</div>', unsafe_allow_html=True)
st.markdown("<br>", unsafe_allow_html=True)

# ================= GST SECTION =================
st.markdown("## 📑 GSTR-1 State-wise Automation")
st.markdown('<div class="card">', unsafe_allow_html=True)
col1, col2 = st.columns([2, 1])

with col1:
    uploaded_file = st.file_uploader("📂 Upload GSTR1 format.xlsx", type=["xlsx"], key="gstr_upload")

with col2:
    month = st.selectbox("📅 Select Month", [
        "January", "February", "March", "April", "May", "June",
        "July", "August", "September", "October", "November", "December"
    ], key="gstr_month")

st.markdown('</div>', unsafe_allow_html=True)

progress_bar = st.progress(0)
status_box = st.empty()
eta_box = st.empty()

if st.button("🚀 Generate State-wise Excel Files", use_container_width=True, key="gstr_run"):
    if uploaded_file is None:
        st.error("❌ Please upload GSTR1 format.xlsx")
        st.stop()

    gstr1_path = os.path.join(UPLOAD_FOLDER, "GSTR1_format.xlsx")
    file_bytes = uploaded_file.getvalue()
    with open(gstr1_path, "wb") as f:
        f.write(file_bytes)

    errs, warns = validate_excel(gstr1_path)
    if errs:
        st.error("File validation failed. Please fix below issue(s):")
        for e in errs:
            st.write(e)
        st.stop()

    if warns:
        st.warning("Validation warning(s):")
        for w in warns:
            st.write(w)

    MONTH = month
    MONTH_NORM = month.lower()

    OUTPUT_FOLDER = os.path.join(OUTPUT_ROOT, MONTH)
    if os.path.exists(OUTPUT_FOLDER):
        shutil.rmtree(OUTPUT_FOLDER)
    os.makedirs(OUTPUT_FOLDER)

    progress_bar.progress(0)
    status_box.info("⏳ Starting processing...")
    eta_box.caption("")

    with st.spinner("📥 Reading Excel (cached for faster repeats)..."):
        df_gstr1, df_b2b, df_cd, df_dn, df_exempt, df_b2c_pf, df_cdunreg, df_b2c_onboard = read_all_sheets_cached(file_bytes)

    with st.spinner("⚙️ Processing state-wise files..."):
        run_gstr_process(
            df_gstr1, df_b2b, df_cd, df_dn, df_exempt,
            df_b2c_pf, df_cdunreg, df_b2c_onboard,
            MONTH, MONTH_NORM,
            TEMPLATE_FILE, OUTPUT_FOLDER,
            progress_bar=progress_bar,
            status_box=status_box,
            eta_box=eta_box
        )

    status_box.success("🎉 Processing completed!")

    zip_path = os.path.join(OUTPUT_ROOT, f"GSTR1_{MONTH}.zip")
    with st.spinner("📦 Preparing ZIP for download..."):
        with zipfile.ZipFile(zip_path, "w") as zipf:
            for file in os.listdir(OUTPUT_FOLDER):
                zipf.write(os.path.join(OUTPUT_FOLDER, file), arcname=file)

        log_activity(action="GSTR-1 Generated", month=MONTH, filename=f"GSTR1_{MONTH}.zip")

    st.success("✅ ZIP Ready for download.")
    st.markdown("## 📥 Download Summary")

    for file in os.listdir(OUTPUT_FOLDER):
        if not file.endswith(".xlsx"):
            continue

        state_name = file.replace(f"GSTR1_{MONTH}_", "").replace(".xlsx", "")
        file_path = os.path.join(OUTPUT_FOLDER, file)

        c1, c2 = st.columns([3, 1])
        with c1:
            st.markdown(f"✅ **{state_name}**")
        with c2:
            with open(file_path, "rb") as f:
                if st.download_button("⬇ Download", data=f, file_name=file, key=f"dl_{file}"):
                    log_activity(action="State File Downloaded", month=MONTH, filename=file)

    with open(zip_path, "rb") as z:
        st.download_button("⬇ Download GSTR1 ZIP", z, file_name=f"GSTR1_{MONTH}.zip", key="dl_zip", use_container_width=True)

# ================= DAY BOOK SECTION =================
st.markdown("---")
st.markdown("## 🏦 HO Day Book Automation")

colA, colB = st.columns([1, 1])
with colA:
    daybook_zip = st.file_uploader("📦 Upload Statement ZIP (statements.zip)", type=["zip"], key="db_zip")
with colB:
    coa_file = st.file_uploader("📄 Upload COA.xlsx", type=["xlsx"], key="db_coa")

if daybook_zip and coa_file:
    if st.button("🚀 Generate HO Day Book", use_container_width=True, key="db_run"):
        try:
            with st.spinner("⚙️ Processing Day Book..."):
                tmpdir = tempfile.mkdtemp()
                out_path = os.path.join(tmpdir, "HO_DayBook_AllStatements.xlsx")
                run_daybook_from_zip(daybook_zip.getvalue(), coa_file.getvalue(), out_path)

            st.success("✅ HO Day Book Generated Successfully!")
            with open(out_path, "rb") as f:
                st.download_button(
                    "⬇ Download HO Day Book Output",
                    data=f,
                    file_name="HO_DayBook_AllStatements.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    key="db_download"
                )
        except Exception as e:
            st.error("❌ Error while generating HO Day Book")
            st.exception(e)
