"""
UNIVERSAL BANK BOOKS REPORT V2
==============================

Drop this script into a folder with one or more bank statements and run it.

Supported inputs:
    .xlsx / .xls / .csv / .pdf

Output:
    UNIVERSAL_BANK_BOOKS_REPORT.xlsx

Target transaction format:
    Date
    Customer / Counterparty Name
    UTR / Reference Number
    Narration
    Debit
    Credit
    Balance
    Source File

Reports:
    COVER_REPORT
    TRANSACTION_REGISTER
    DATE_WISE_SUMMARY
    MONTH_WISE_SUMMARY
    CUSTOMER_WISE_SUMMARY
    CUSTOMER_DATE_WISE
    UTR_WISE_REGISTER
    CONTROL_TOTALS
    SOURCE_MAPPING
    REPORT_NOTES

DATA INTEGRITY
--------------
- No transaction is invented.
- No missing customer name or UTR is manufactured.
- If a bank has a separate customer / beneficiary / payer / payee field,
  that source value is used.
- If customer or UTR exists only in narration, the script attempts a
  best-effort extraction and records its source.
- Possible duplicates are flagged, never deleted automatically.
- Searchable/text PDFs work best.
- Scanned/image-only PDFs need OCR dependencies and should be manually verified.

Install:
    python -m pip install pandas openpyxl xlrd pdfplumber

Optional for scanned PDFs:
    python -m pip install pytesseract pdf2image pillow

Windows scanned-PDF OCR additionally needs Tesseract OCR and Poppler installed.
"""

import os
import glob
import re
from datetime import datetime

import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# ============================================================
# CONFIG
# ============================================================

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_FILE = os.path.join(BASE_DIR, "UNIVERSAL_BANK_BOOKS_REPORT.xlsx")

PATTERNS = (
    "*.xlsx", "*.XLSX",
    "*.xls", "*.XLS",
    "*.csv", "*.CSV",
    "*.pdf", "*.PDF",
)

files = []
for pattern in PATTERNS:
    files.extend(glob.glob(os.path.join(BASE_DIR, pattern)))

files = sorted({
    p for p in files
    if os.path.abspath(p).lower() != os.path.abspath(OUTPUT_FILE).lower()
    and not os.path.basename(p).startswith("~$")
})

if not files:
    raise FileNotFoundError(
        "No XLSX/XLS/CSV/PDF bank statement found in:\n" + BASE_DIR
    )


# ============================================================
# FORMAT
# ============================================================

HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
SUB_FILL = PatternFill("solid", fgColor="D9EAF7")
WARN_FILL = PatternFill("solid", fgColor="FCE4D6")

WHITE_BOLD = Font(color="FFFFFF", bold=True)
BOLD = Font(bold=True)

THIN = Side(style="thin", color="D9D9D9")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

MONEY = '₹#,##0.00;[Red](₹#,##0.00);-'
DATE_FMT = 'dd-mm-yyyy'
COUNT_FMT = '#,##0'


# ============================================================
# ALIASES
# ============================================================

ALIASES = {
    "date": [
        "date", "txn date", "transaction date", "tran date", "trn date",
        "posting date", "value date", "date of transaction"
    ],
    "narration": [
        "narration", "particulars", "trn particulars", "trn. particulars",
        "transaction particulars", "description", "details", "remarks",
        "transaction description", "transaction details"
    ],
    "debit": [
        "debit", "debit amount", "debit amt", "withdrawal",
        "withdrawal amount", "dr", "dr amount", "paid out"
    ],
    "credit": [
        "credit", "credit amount", "credit amt", "deposit",
        "deposit amount", "cr", "cr amount", "paid in"
    ],
    "balance": [
        "balance", "closing balance", "running balance",
        "available balance", "ledger balance"
    ],
    "customer": [
        "customer", "customer name", "counterparty", "party name",
        "payer", "payer name", "payee", "payee name",
        "beneficiary", "beneficiary name", "remitter name",
        "sender name", "receiver name"
    ],
    "utr": [
        "utr", "utr no", "utr number", "bank utr",
        "reference", "reference no", "reference number",
        "reference id", "ref id", "bank ref", "bank ref no",
        "bank reference", "rrn", "txn id", "transaction id",
        "bank transaction id"
    ],
    "amount": [
        "amount", "transaction amount", "txn amount"
    ],
    "type": [
        "type", "transaction type", "txn type",
        "dr cr", "cr dr", "debit credit", "credit debit"
    ],
}


# ============================================================
# HELPERS
# ============================================================

def text(v):
    if v is None:
        return ""
    if isinstance(v, float) and np.isnan(v):
        return ""
    return " ".join(
        str(v).replace("\ufeff", "").replace("\n", " ").replace("\r", " ").strip().split()
    )


def norm(v):
    return re.sub(r"[^a-z0-9]+", " ", text(v).lower()).strip()


def num(v):
    if v is None:
        return np.nan

    if isinstance(v, (int, float, np.integer, np.floating)):
        try:
            return float(v)
        except Exception:
            return np.nan

    s = text(v)
    if not s:
        return np.nan

    negative = s.startswith("(") and s.endswith(")")
    if negative:
        s = s[1:-1]

    s = (
        s.replace(",", "")
         .replace("₹", "")
         .replace("INR", "")
         .replace("Rs.", "")
         .replace("Rs", "")
         .strip()
    )

    # Remove common DR/CR suffixes.
    s = re.sub(r"\s*(DR|CR)\s*$", "", s, flags=re.I).strip()

    try:
        value = float(s)
        return -value if negative else value
    except Exception:
        return np.nan


def matches(v, field):
    n = norm(v)
    if not n:
        return False

    for alias in ALIASES[field]:
        a = norm(alias)

        if n == a:
            return True

        if len(a) >= 4 and (a in n or n in a):
            return True

    return False


def find_col(columns, field):
    # Exact first.
    lookup = {norm(c): c for c in columns}

    for alias in ALIASES[field]:
        a = norm(alias)
        if a in lookup:
            return lookup[a]

    # Flexible fallback.
    for col in columns:
        if matches(col, field):
            return col

    return None


def header_hits(values):
    hits = set()

    for field in ALIASES:
        if any(matches(v, field) for v in values):
            hits.add(field)

    return hits


def is_header(values):
    hits = header_hits(values)

    # Standard separate debit / credit statement.
    if "date" in hits and ("debit" in hits or "credit" in hits):
        return True

    # Amount + DR/CR type statement.
    if "date" in hits and "amount" in hits and "type" in hits:
        return True

    return False


def find_header(raw, max_rows=150):
    limit = min(len(raw), max_rows)

    best_row = None
    best_score = -1

    for i in range(limit):
        vals = raw.iloc[i].tolist()
        hits = header_hits(vals)

        if is_header(vals):
            score = (
                5 * ("date" in hits)
                + 3 * ("debit" in hits)
                + 3 * ("credit" in hits)
                + 2 * ("narration" in hits)
                + 2 * ("amount" in hits)
                + 2 * ("type" in hits)
                + 1 * ("balance" in hits)
                + 1 * ("customer" in hits)
                + 1 * ("utr" in hits)
            )

            if score > best_score:
                best_score = score
                best_row = i

    return best_row


def make_table(raw, header_row):
    headers = [text(x) for x in raw.iloc[header_row].tolist()]

    safe = []
    used = {}

    for i, h in enumerate(headers):
        base = h or f"Column_{i+1}"
        used[base] = used.get(base, 0) + 1
        safe.append(base if used[base] == 1 else f"{base}_{used[base]}")

    data = raw.iloc[header_row + 1:].copy()
    data.columns = safe

    blank = data.apply(
        lambda r: all(text(x) == "" for x in r.tolist()),
        axis=1
    )

    return data.loc[~blank].copy()


# ============================================================
# BEST-EFFORT UTR / NAME EXTRACTION
# ============================================================

def utr_from_narration(v):
    s = text(v)

    if not s:
        return ""

    patterns = [
        r"(?:UTR)[\s:/-]*([A-Z0-9]{8,35})",
        r"(?:RRN)[\s:/-]*([0-9]{8,25})",
        r"/XUTR/([^/\s]+)",
        r"(?:REF(?:ERENCE)?)[\s:/-]*([A-Z0-9]{8,35})",
    ]

    for p in patterns:
        m = re.search(p, s, flags=re.I)
        if m:
            return text(m.group(1))

    return ""


def customer_from_narration(v):
    s = text(v)

    if not s:
        return ""

    parts = [p.strip() for p in s.split("/") if p.strip()]

    # NEFT//XUTR/UTR/PARTY/BANK
    for i, p in enumerate(parts):
        if p.upper() == "XUTR" and i + 2 < len(parts):
            return text(parts[i + 2])

    # UPI/NAME/REF...
    if parts and parts[0].upper().startswith("UPI") and len(parts) >= 2:
        candidate = text(parts[1])
        if candidate and not candidate.isdigit():
            return candidate

    return ""


# ============================================================
# STANDARDIZE
# ============================================================

def standardize(table):
    date_col = find_col(table.columns, "date")
    narr_col = find_col(table.columns, "narration")
    debit_col = find_col(table.columns, "debit")
    credit_col = find_col(table.columns, "credit")
    bal_col = find_col(table.columns, "balance")
    cust_col = find_col(table.columns, "customer")
    utr_col = find_col(table.columns, "utr")
    amount_col = find_col(table.columns, "amount")
    type_col = find_col(table.columns, "type")

    if date_col is None:
        return pd.DataFrame()

    out = pd.DataFrame(index=table.index)

    out["Date"] = pd.to_datetime(
        table[date_col],
        errors="coerce",
        dayfirst=True
    ).dt.normalize()

    out["Narration"] = table[narr_col] if narr_col else ""
    out["Customer Name"] = table[cust_col] if cust_col else ""
    out["UTR / Reference"] = table[utr_col] if utr_col else ""

    out["Debit"] = table[debit_col].apply(num) if debit_col else np.nan
    out["Credit"] = table[credit_col].apply(num) if credit_col else np.nan
    out["Balance"] = table[bal_col].apply(num) if bal_col else np.nan

    # Fallback: Amount + Type
    if debit_col is None and credit_col is None and amount_col and type_col:
        amount = table[amount_col].apply(num)
        txn_type = table[type_col].astype(str).str.upper()

        out["Debit"] = np.where(
            txn_type.str.contains(r"\bDR\b|DEBIT|WITHDRAW", regex=True, na=False),
            amount,
            np.nan
        )

        out["Credit"] = np.where(
            txn_type.str.contains(r"\bCR\b|CREDIT|DEPOSIT", regex=True, na=False),
            amount,
            np.nan
        )

    customer_blank = (
        out["Customer Name"].fillna("").astype(str).str.strip().eq("")
    )

    out.loc[customer_blank, "Customer Name"] = out.loc[
        customer_blank, "Narration"
    ].apply(customer_from_narration)

    utr_blank = (
        out["UTR / Reference"].fillna("").astype(str).str.strip().eq("")
    )

    out.loc[utr_blank, "UTR / Reference"] = out.loc[
        utr_blank, "Narration"
    ].apply(utr_from_narration)

    out["Customer Source"] = np.where(
        cust_col is not None,
        "SOURCE COLUMN",
        np.where(
            out["Customer Name"].fillna("").astype(str).str.strip().ne(""),
            "EXTRACTED FROM NARRATION",
            "NOT AVAILABLE"
        )
    )

    out["UTR Source"] = np.where(
        utr_col is not None,
        "SOURCE COLUMN",
        np.where(
            out["UTR / Reference"].fillna("").astype(str).str.strip().ne(""),
            "EXTRACTED FROM NARRATION",
            "NOT AVAILABLE"
        )
    )

    # Actual transaction lines only.
    actual = out["Debit"].notna() | out["Credit"].notna()

    return out.loc[actual].copy()


# ============================================================
# INPUT READERS
# ============================================================

def read_tabular(path):
    ext = os.path.splitext(path)[1].lower()

    if ext == ".csv":
        last_error = None

        for enc in ("utf-8-sig", "utf-8", "cp1252", "latin1"):
            try:
                raw = pd.read_csv(
                    path,
                    header=None,
                    dtype=str,
                    keep_default_na=False,
                    encoding=enc,
                    engine="python"
                )
                return [("CSV", raw)]
            except Exception as e:
                last_error = e

        raise last_error

    engine = "xlrd" if ext == ".xls" else "openpyxl"
    book = pd.ExcelFile(path, engine=engine)

    result = []

    for sheet in book.sheet_names:
        try:
            raw = pd.read_excel(
                book,
                sheet_name=sheet,
                header=None,
                dtype=str,
                keep_default_na=False
            )
            result.append((sheet, raw))
        except Exception as e:
            print(f"    Sheet skipped: {sheet}: {e}")

    return result


def read_pdf_tables(path):
    import pdfplumber

    result = []

    with pdfplumber.open(path) as pdf:
        for page_no, page in enumerate(pdf.pages, start=1):
            tables = page.extract_tables() or []

            for table_no, table in enumerate(tables, start=1):
                if not table:
                    continue

                width = max(len(r or []) for r in table)
                rows = []

                for row in table:
                    row = list(row or [])
                    row += [""] * (width - len(row))
                    rows.append(row)

                result.append(
                    (
                        f"PDF Page {page_no} Table {table_no}",
                        pd.DataFrame(rows)
                    )
                )

    return result


def read_pdf_ocr_optional(path):
    """
    OCR fallback for image-only PDF.
    Conservative: attempts OCR text-to-table only when the OCR output itself
    has a recognizable tabular structure. It does not invent columns.
    """
    try:
        from pdf2image import convert_from_path
        import pytesseract
    except Exception:
        return []

    result = []

    try:
        images = convert_from_path(path, dpi=300)
    except Exception:
        return []

    for page_no, image in enumerate(images, start=1):
        try:
            data = pytesseract.image_to_data(
                image,
                output_type=pytesseract.Output.DATAFRAME,
                config="--psm 6"
            )
        except Exception:
            continue

        if data is None or data.empty:
            continue

        data = data.dropna(subset=["text"]).copy()
        data["text"] = data["text"].astype(str).str.strip()
        data = data[data["text"] != ""]

        # Rebuild physical text lines. This can help diagnostics but is not
        # allowed to fabricate debit/credit column boundaries.
        lines = (
            data.groupby(["block_num", "par_num", "line_num"])["text"]
            .apply(lambda s: " ".join(s))
            .tolist()
        )

        if lines:
            result.append(
                (
                    f"OCR Page {page_no}",
                    pd.DataFrame([[line] for line in lines])
                )
            )

    return result


# ============================================================
# LOAD ALL FILES
# ============================================================

all_frames = []
source_rows = []

print("=" * 80)
print("UNIVERSAL BANK BOOKS REPORT V2")
print("=" * 80)
print("Folder:", BASE_DIR)
print("Source files found:", len(files))
print("")

for path in files:
    base = os.path.basename(path)
    ext = os.path.splitext(path)[1].lower()

    print("Reading:", base)

    try:
        if ext == ".pdf":
            sources = read_pdf_tables(path)

            if not sources:
                sources = read_pdf_ocr_optional(path)
        else:
            sources = read_tabular(path)

    except Exception as e:
        print("  ERROR:", e)

        source_rows.append({
            "Source File": base,
            "Source Part": "",
            "Status": "READ ERROR",
            "Header Row": "",
            "Transaction Count": 0,
            "Note": str(e),
        })

        continue

    if not sources:
        source_rows.append({
            "Source File": base,
            "Source Part": "",
            "Status": "NO TABLE FOUND",
            "Header Row": "",
            "Transaction Count": 0,
            "Note": "PDF may be scanned/image-only or unsupported layout.",
        })
        print("  No usable table found.")
        continue

    for source_name, raw in sources:
        header = find_header(raw)

        if header is None:
            source_rows.append({
                "Source File": base,
                "Source Part": source_name,
                "Status": "UNRECOGNIZED",
                "Header Row": "",
                "Transaction Count": 0,
                "Note": "Could not identify Date + Debit/Credit or Date + Amount + DR/CR header.",
            })
            continue

        table = make_table(raw, header)
        txns = standardize(table)

        if txns.empty:
            source_rows.append({
                "Source File": base,
                "Source Part": source_name,
                "Status": "HEADER FOUND / NO TRANSACTIONS",
                "Header Row": header + 1,
                "Transaction Count": 0,
                "Note": "",
            })
            continue

        txns["Source File"] = base
        txns["Source Part"] = source_name
        txns["Source Row"] = (
            txns.index.to_series().astype(int) + header + 2
        ).values

        all_frames.append(txns)

        source_rows.append({
            "Source File": base,
            "Source Part": source_name,
            "Status": "RECOGNIZED",
            "Header Row": header + 1,
            "Transaction Count": len(txns),
            "Note": "",
        })

        print(f"  {source_name}: {len(txns):,} transaction(s)")


if not all_frames:
    raise ValueError(
        "\nNo bank transaction rows were recognized.\n"
        "Expected formats similar to:\n"
        "Date | Narration | Debit | Credit | Balance\n"
        "or Date | Amount | Type (DR/CR).\n\n"
        "For scanned PDFs, install OCR dependencies or convert the statement "
        "to searchable PDF / Excel for better audit accuracy."
    )

tx = pd.concat(all_frames, ignore_index=True, sort=False)

tx["Debit"] = pd.to_numeric(tx["Debit"], errors="coerce")
tx["Credit"] = pd.to_numeric(tx["Credit"], errors="coerce")
tx["Balance"] = pd.to_numeric(tx["Balance"], errors="coerce")

tx["Direction"] = np.where(
    tx["Debit"].notna(),
    "DEBIT",
    np.where(tx["Credit"].notna(), "CREDIT", "")
)

tx["Transaction Amount"] = np.where(
    tx["Direction"].eq("DEBIT"),
    tx["Debit"],
    tx["Credit"]
)

# Never automatically remove duplicates.
dup_key = (
    tx["Date"].astype(str)
    + "|" + tx["Debit"].fillna(0).astype(str)
    + "|" + tx["Credit"].fillna(0).astype(str)
    + "|" + tx["UTR / Reference"].fillna("").astype(str)
    + "|" + tx["Narration"].fillna("").astype(str)
)

tx["Possible Duplicate"] = np.where(
    dup_key.duplicated(keep=False),
    "YES",
    "NO"
)


# ============================================================
# SUMMARIES
# ============================================================

date_wise = (
    tx.groupby("Date", dropna=False)
    .agg(
        Transaction_Count=("Transaction Amount", "count"),
        Debit_Count=("Debit", lambda s: int(s.notna().sum())),
        Debit_Total=("Debit", "sum"),
        Credit_Count=("Credit", lambda s: int(s.notna().sum())),
        Credit_Total=("Credit", "sum"),
    )
    .reset_index()
)

date_wise["Net_Credit_Minus_Debit"] = (
    date_wise["Credit_Total"].fillna(0)
    - date_wise["Debit_Total"].fillna(0)
)


tx["Month"] = tx["Date"].dt.to_period("M").astype(str)

month_wise = (
    tx.groupby("Month", dropna=False)
    .agg(
        Transaction_Count=("Transaction Amount", "count"),
        Debit_Count=("Debit", lambda s: int(s.notna().sum())),
        Debit_Total=("Debit", "sum"),
        Credit_Count=("Credit", lambda s: int(s.notna().sum())),
        Credit_Total=("Credit", "sum"),
        Unique_Customers=("Customer Name", lambda s: s.replace("", np.nan).dropna().nunique()),
        Unique_UTRs=("UTR / Reference", lambda s: s.replace("", np.nan).dropna().nunique()),
    )
    .reset_index()
)

month_wise["Net_Credit_Minus_Debit"] = (
    month_wise["Credit_Total"].fillna(0)
    - month_wise["Debit_Total"].fillna(0)
)


customer_tx = tx[
    tx["Customer Name"].fillna("").astype(str).str.strip().ne("")
].copy()

if customer_tx.empty:
    customer_wise = pd.DataFrame(columns=[
        "Customer Name",
        "Transaction_Count",
        "Debit_Count",
        "Debit_Total",
        "Credit_Count",
        "Credit_Total",
        "Net_Credit_Minus_Debit",
        "First_Date",
        "Last_Date",
        "Unique_UTRs",
    ])
else:
    customer_wise = (
        customer_tx.groupby("Customer Name", dropna=False)
        .agg(
            Transaction_Count=("Transaction Amount", "count"),
            Debit_Count=("Debit", lambda s: int(s.notna().sum())),
            Debit_Total=("Debit", "sum"),
            Credit_Count=("Credit", lambda s: int(s.notna().sum())),
            Credit_Total=("Credit", "sum"),
            First_Date=("Date", "min"),
            Last_Date=("Date", "max"),
            Unique_UTRs=("UTR / Reference", lambda s: s.replace("", np.nan).dropna().nunique()),
        )
        .reset_index()
    )

    customer_wise["Net_Credit_Minus_Debit"] = (
        customer_wise["Credit_Total"].fillna(0)
        - customer_wise["Debit_Total"].fillna(0)
    )

    customer_wise = customer_wise[
        [
            "Customer Name",
            "Transaction_Count",
            "Debit_Count",
            "Debit_Total",
            "Credit_Count",
            "Credit_Total",
            "Net_Credit_Minus_Debit",
            "First_Date",
            "Last_Date",
            "Unique_UTRs",
        ]
    ]


if customer_tx.empty:
    customer_date_wise = pd.DataFrame(columns=[
        "Date",
        "Customer Name",
        "Transaction_Count",
        "Debit_Count",
        "Debit_Total",
        "Credit_Count",
        "Credit_Total",
        "Net_Credit_Minus_Debit",
    ])
else:
    customer_date_wise = (
        customer_tx.groupby(["Date", "Customer Name"], dropna=False)
        .agg(
            Transaction_Count=("Transaction Amount", "count"),
            Debit_Count=("Debit", lambda s: int(s.notna().sum())),
            Debit_Total=("Debit", "sum"),
            Credit_Count=("Credit", lambda s: int(s.notna().sum())),
            Credit_Total=("Credit", "sum"),
        )
        .reset_index()
    )

    customer_date_wise["Net_Credit_Minus_Debit"] = (
        customer_date_wise["Credit_Total"].fillna(0)
        - customer_date_wise["Debit_Total"].fillna(0)
    )


utr_register = tx[
    [
        "Date",
        "Customer Name",
        "UTR / Reference",
        "Narration",
        "Debit",
        "Credit",
        "Balance",
        "Source File",
    ]
].copy()


# ============================================================
# TOTALS
# ============================================================

total_count = len(tx)
debit_count = int(tx["Debit"].notna().sum())
credit_count = int(tx["Credit"].notna().sum())
debit_total = float(tx["Debit"].fillna(0).sum())
credit_total = float(tx["Credit"].fillna(0).sum())
net_total = credit_total - debit_total

valid_dates = tx["Date"].dropna()
first_date = valid_dates.min() if not valid_dates.empty else None
last_date = valid_dates.max() if not valid_dates.empty else None

customer_count = int(
    tx["Customer Name"].replace("", np.nan).dropna().nunique()
)

utr_count = int(
    tx["UTR / Reference"].replace("", np.nan).dropna().nunique()
)

missing_customer = int(
    tx["Customer Name"].fillna("").astype(str).str.strip().eq("").sum()
)

missing_utr = int(
    tx["UTR / Reference"].fillna("").astype(str).str.strip().eq("").sum()
)

duplicate_count = int(
    tx["Possible Duplicate"].eq("YES").sum()
)


# ============================================================
# EXCEL HELPERS
# ============================================================

wb = Workbook()
wb.remove(wb.active)


def excel_value(v):
    if isinstance(v, pd.Timestamp):
        return v.to_pydatetime()

    if isinstance(v, np.integer):
        return int(v)

    if isinstance(v, np.floating):
        return None if np.isnan(v) else float(v)

    if pd.isna(v):
        return None

    return v


def header_style(ws):
    for cell in ws[1]:
        cell.fill = HEADER_FILL
        cell.font = WHITE_BOLD
        cell.border = BORDER
        cell.alignment = Alignment(
            horizontal="center",
            vertical="center",
            wrap_text=True
        )


def auto_width(ws, cap=40):
    for c in range(1, ws.max_column + 1):
        max_len = 0

        for r in range(1, min(ws.max_row, 2500) + 1):
            v = ws.cell(r, c).value
            if v is not None:
                max_len = max(max_len, len(str(v)))

        ws.column_dimensions[get_column_letter(c)].width = min(
            max(max_len + 2, 11),
            cap
        )


def write_df(ws, frame, money_cols=(), date_cols=(), count_cols=()):
    if frame is None or frame.empty:
        ws["A1"] = "No data available"
        return

    for c, col in enumerate(frame.columns, start=1):
        ws.cell(1, c, col)

    for r, row in enumerate(frame.itertuples(index=False, name=None), start=2):
        for c, value in enumerate(row, start=1):
            ws.cell(r, c, excel_value(value))
            ws.cell(r, c).border = BORDER
            ws.cell(r, c).alignment = Alignment(vertical="center")

    header_style(ws)

    for c, col in enumerate(frame.columns, start=1):
        if col in money_cols:
            for r in range(2, ws.max_row + 1):
                ws.cell(r, c).number_format = MONEY

        if col in date_cols:
            for r in range(2, ws.max_row + 1):
                ws.cell(r, c).number_format = DATE_FMT

        if col in count_cols:
            for r in range(2, ws.max_row + 1):
                ws.cell(r, c).number_format = COUNT_FMT

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    auto_width(ws)


# ============================================================
# COVER_REPORT
# ============================================================

ws = wb.create_sheet("COVER_REPORT")
ws.sheet_view.showGridLines = False

ws.merge_cells("A1:D1")
ws["A1"] = "UNIVERSAL BANK BOOKS OF ACCOUNTS REPORT"
ws["A1"].font = Font(bold=True, size=16, color="1F4E78")
ws["A1"].alignment = Alignment(horizontal="center")

cover = [
    ("Total Transaction Count", total_count),
    ("Debit Transaction Count", debit_count),
    ("Total Debit", debit_total),
    ("Credit Transaction Count", credit_count),
    ("Total Credit", credit_total),
    ("Net Credit - Debit", net_total),
    ("First Transaction Date", first_date),
    ("Last Transaction Date", last_date),
    ("Unique Customers", customer_count),
    ("Unique UTR / References", utr_count),
    ("Transactions Missing Customer Name", missing_customer),
    ("Transactions Missing UTR / Reference", missing_utr),
    ("Possible Duplicate Rows", duplicate_count),
    ("Recognized Source Files", tx["Source File"].nunique()),
]

for r, (label, value) in enumerate(cover, start=3):
    ws.cell(r, 1, label)
    ws.cell(r, 2, excel_value(value))

    ws.cell(r, 1).border = BORDER
    ws.cell(r, 2).border = BORDER

    if label in ("Total Debit", "Total Credit", "Net Credit - Debit"):
        ws.cell(r, 2).number_format = MONEY

    if "Date" in label:
        ws.cell(r, 2).number_format = DATE_FMT

ws.column_dimensions["A"].width = 42
ws.column_dimensions["B"].width = 24


# ============================================================
# TRANSACTION_REGISTER
# ============================================================

register_cols = [
    "Date",
    "Customer Name",
    "UTR / Reference",
    "Narration",
    "Debit",
    "Credit",
    "Balance",
    "Direction",
    "Transaction Amount",
    "Customer Source",
    "UTR Source",
    "Possible Duplicate",
    "Source File",
    "Source Part",
    "Source Row",
]

ws = wb.create_sheet("TRANSACTION_REGISTER")

write_df(
    ws,
    tx[register_cols],
    money_cols=("Debit", "Credit", "Balance", "Transaction Amount"),
    date_cols=("Date",),
)


# ============================================================
# DATE_WISE_SUMMARY
# ============================================================

ws = wb.create_sheet("DATE_WISE_SUMMARY")

write_df(
    ws,
    date_wise,
    money_cols=("Debit_Total", "Credit_Total", "Net_Credit_Minus_Debit"),
    date_cols=("Date",),
    count_cols=("Transaction_Count", "Debit_Count", "Credit_Count"),
)


# ============================================================
# MONTH_WISE_SUMMARY
# ============================================================

ws = wb.create_sheet("MONTH_WISE_SUMMARY")

write_df(
    ws,
    month_wise,
    money_cols=("Debit_Total", "Credit_Total", "Net_Credit_Minus_Debit"),
    count_cols=(
        "Transaction_Count",
        "Debit_Count",
        "Credit_Count",
        "Unique_Customers",
        "Unique_UTRs",
    ),
)


# ============================================================
# CUSTOMER_WISE_SUMMARY
# ============================================================

ws = wb.create_sheet("CUSTOMER_WISE_SUMMARY")

write_df(
    ws,
    customer_wise,
    money_cols=("Debit_Total", "Credit_Total", "Net_Credit_Minus_Debit"),
    date_cols=("First_Date", "Last_Date"),
    count_cols=(
        "Transaction_Count",
        "Debit_Count",
        "Credit_Count",
        "Unique_UTRs",
    ),
)


# ============================================================
# CUSTOMER_DATE_WISE
# ============================================================

ws = wb.create_sheet("CUSTOMER_DATE_WISE")

write_df(
    ws,
    customer_date_wise,
    money_cols=("Debit_Total", "Credit_Total", "Net_Credit_Minus_Debit"),
    date_cols=("Date",),
    count_cols=("Transaction_Count", "Debit_Count", "Credit_Count"),
)


# ============================================================
# UTR_WISE_REGISTER
# ============================================================

ws = wb.create_sheet("UTR_WISE_REGISTER")

write_df(
    ws,
    utr_register,
    money_cols=("Debit", "Credit", "Balance"),
    date_cols=("Date",),
)


# ============================================================
# CONTROL_TOTALS
# ============================================================

ws = wb.create_sheet("CONTROL_TOTALS")

controls = [
    ["Metric", "Value"],
    ["Total Transaction Count", total_count],
    ["Debit Transaction Count", debit_count],
    ["Total Debit", debit_total],
    ["Credit Transaction Count", credit_count],
    ["Total Credit", credit_total],
    ["Net Credit - Debit", net_total],
    ["Unique Customers", customer_count],
    ["Unique UTR / References", utr_count],
    ["Missing Customer Name", missing_customer],
    ["Missing UTR / Reference", missing_utr],
    ["Possible Duplicate Rows", duplicate_count],
]

for r, row in enumerate(controls, start=1):
    for c, value in enumerate(row, start=1):
        ws.cell(r, c, excel_value(value))
        ws.cell(r, c).border = BORDER

header_style(ws)

for r in (4, 6, 7):
    ws.cell(r, 2).number_format = MONEY

ws.column_dimensions["A"].width = 38
ws.column_dimensions["B"].width = 24


# ============================================================
# SOURCE_MAPPING
# ============================================================

ws = wb.create_sheet("SOURCE_MAPPING")

write_df(
    ws,
    pd.DataFrame(source_rows),
    count_cols=("Header Row", "Transaction Count"),
)


# ============================================================
# REPORT_NOTES
# ============================================================

ws = wb.create_sheet("REPORT_NOTES")

notes = [
    ["Item", "Note"],
    ["Purpose", "Bank statement to Books-of-Accounts supporting transaction report."],
    ["Supported Inputs", "XLSX, XLS, CSV and PDF."],
    ["Core Fields", "Date, Customer/Counterparty Name, UTR/Reference, Narration, Debit, Credit and Balance."],
    ["Date-wise", "Daily transaction count, debit count/total, credit count/total and net movement."],
    ["Month-wise", "Monthly debit/credit totals plus unique customer and UTR counts."],
    ["Customer-wise", "Customer transaction count, debit/credit totals, first/last date and UTR count."],
    ["Customer + Date", "Daily totals for each available customer/counterparty."],
    ["Customer Name", "Uses source customer/payee/payer/beneficiary field where present; otherwise best-effort narration extraction."],
    ["UTR", "Uses source UTR/reference field where present; otherwise best-effort narration extraction."],
    ["Missing Fields", "Missing customer/UTR values remain unavailable. The script does not invent them."],
    ["Duplicates", "Possible duplicates are flagged only and never automatically deleted."],
    ["PDF", "Searchable bank PDFs are preferred. Scanned PDFs may require OCR and manual verification."],
    ["Generated On", datetime.now().strftime("%d-%m-%Y %H:%M:%S")],
]

for r, row in enumerate(notes, start=1):
    for c, value in enumerate(row, start=1):
        ws.cell(r, c, value)
        ws.cell(r, c).border = BORDER

header_style(ws)
ws.column_dimensions["A"].width = 28
ws.column_dimensions["B"].width = 110


# ============================================================
# SAVE
# ============================================================

print("")
print("=" * 80)
print("FINAL TOTALS")
print("=" * 80)
print(f"Transactions : {total_count:,}")
print(f"Debit Count  : {debit_count:,}")
print(f"Debit Total  : {debit_total:,.2f}")
print(f"Credit Count : {credit_count:,}")
print(f"Credit Total : {credit_total:,.2f}")
print(f"Customers    : {customer_count:,}")
print(f"UTR / Refs   : {utr_count:,}")
print(f"Missing Name : {missing_customer:,}")
print(f"Missing UTR  : {missing_utr:,}")
print("")
print("Saving:", OUTPUT_FILE)

wb.save(OUTPUT_FILE)

print("SUCCESS")
print("Output:", OUTPUT_FILE)
print("=" * 80)
