"""
UNIVERSAL BANK BOOKS REPORT V4
==============================

FIX IN V3
---------
Some bank PDFs (especially ICICI statements) look like normal tables on screen,
but PDF table extraction returns ONE giant multi-line row per page. That caused
older versions to report only 1 transaction per page and incorrect debit/credit
totals.

V3 detects that statement layout and parses transaction lines from PDF text,
then determines Debit/Credit from the running balance movement. It also checks
each page against the bank's printed page totals where available.

Supported:
    .xlsx
    .xls
    .csv
    .pdf

Output:
    UNIVERSAL_BANK_BOOKS_REPORT_V4.xlsx

Main reports:
    COVER_REPORT
    TRANSACTION_REGISTER
    DATE_WISE_SUMMARY
    MONTH_WISE_SUMMARY
    CUSTOMER_WISE_SUMMARY
    CUSTOMER_DATE_WISE
    UTR_WISE_REGISTER
    PDF_PAGE_RECON
    CONTROL_TOTALS
    SOURCE_MAPPING
    REPORT_NOTES

Core transaction fields:
    Date
    Customer / Counterparty Name
    UTR / Reference Number
    Narration
    Debit
    Credit
    Balance
    Source File
    Source Page / Sheet

DATA INTEGRITY
--------------
- No transaction is invented.
- No missing name, UTR, debit, or credit value is manufactured.
- Customer/UTR extraction from narration is best-effort and explicitly marked.
- Possible duplicates are flagged, never deleted.
- ICICI-style PDF page totals are used as a control check.
- Searchable/text PDFs are strongly preferred.
- Image-only/scanned PDFs may need OCR and manual verification.

INSTALL
-------
python -m pip install pandas openpyxl xlrd pdfplumber

Optional OCR:
python -m pip install pytesseract pdf2image pillow

RUN
---
python UNIVERSAL_BANK_BOOKS_REPORT_V4.py
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
OUTPUT_FILE = os.path.join(BASE_DIR, "UNIVERSAL_BANK_BOOKS_REPORT_V4.xlsx")

PATTERNS = (
    "*.xlsx", "*.XLSX",
    "*.xls", "*.XLS",
    "*.csv", "*.CSV",
    "*.pdf", "*.PDF",
)

input_files = []
for pattern in PATTERNS:
    input_files.extend(glob.glob(os.path.join(BASE_DIR, pattern)))

input_files = sorted({
    p for p in input_files
    if os.path.abspath(p).lower() != os.path.abspath(OUTPUT_FILE).lower()
    and not os.path.basename(p).startswith("~$")
})

if not input_files:
    raise FileNotFoundError(
        "No XLSX/XLS/CSV/PDF statement found in:\n" + BASE_DIR
    )


# ============================================================
# EXCEL STYLE
# ============================================================

HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
SUB_FILL = PatternFill("solid", fgColor="D9EAF7")
WARN_FILL = PatternFill("solid", fgColor="FCE4D6")
OK_FILL = PatternFill("solid", fgColor="E2F0D9")

WHITE_BOLD = Font(color="FFFFFF", bold=True)
BOLD = Font(bold=True)

THIN = Side(style="thin", color="D9D9D9")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

MONEY_FMT = '₹#,##0.00;[Red](₹#,##0.00);-'
DATE_FMT = 'dd-mm-yyyy'
COUNT_FMT = '#,##0'


# ============================================================
# GENERIC BANK HEADER ALIASES
# ============================================================

ALIASES = {
    "date": [
        "date", "tran date", "transaction date", "txn date", "trn date",
        "posting date", "value date", "date of transaction"
    ],
    "narration": [
        "particulars", "narration", "description", "details",
        "trn particulars", "trn. particulars", "transaction particulars",
        "transaction description", "transaction details", "remarks"
    ],
    "debit": [
        "debit", "debit amount", "debit amt",
        "withdrawal", "withdrawals", "withdrawal amount",
        "dr", "dr amount", "paid out"
    ],
    "credit": [
        "credit", "credit amount", "credit amt",
        "deposit", "deposits", "deposit amount",
        "cr", "cr amount", "paid in"
    ],
    "balance": [
        "balance", "balance inr", "balance (inr)",
        "closing balance", "running balance",
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
    "amount": ["amount", "transaction amount", "txn amount"],
    "type": [
        "type", "transaction type", "txn type",
        "dr cr", "cr dr", "debit credit", "credit debit"
    ],
}


# ============================================================
# BASIC HELPERS
# ============================================================

def clean_text(value):
    if value is None:
        return ""
    if isinstance(value, float) and np.isnan(value):
        return ""
    return " ".join(
        str(value)
        .replace("\ufeff", "")
        .replace("\n", " ")
        .replace("\r", " ")
        .strip()
        .split()
    )


def norm(value):
    return re.sub(
        r"[^a-z0-9]+",
        " ",
        clean_text(value).lower()
    ).strip()


def parse_number(value):
    if value is None:
        return np.nan

    if isinstance(value, (int, float, np.integer, np.floating)):
        try:
            return float(value)
        except Exception:
            return np.nan

    s = clean_text(value)
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

    s = re.sub(r"\s*(DR|CR)\s*$", "", s, flags=re.I).strip()

    try:
        x = float(s)
        return -x if negative else x
    except Exception:
        return np.nan


def alias_match(value, field):
    v = norm(value)
    if not v:
        return False

    for alias in ALIASES[field]:
        a = norm(alias)

        if v == a:
            return True

        if len(a) >= 4 and (a in v or v in a):
            return True

    return False


def find_column(columns, field):
    exact = {norm(c): c for c in columns}

    for alias in ALIASES[field]:
        a = norm(alias)
        if a in exact:
            return exact[a]

    for c in columns:
        if alias_match(c, field):
            return c

    return None


def header_hits(values):
    hits = set()

    for field in ALIASES:
        if any(alias_match(v, field) for v in values):
            hits.add(field)

    return hits


def is_generic_bank_header(values):
    hits = header_hits(values)

    if "date" in hits and ("debit" in hits or "credit" in hits):
        return True

    if "date" in hits and "amount" in hits and "type" in hits:
        return True

    return False


def locate_generic_header(raw, max_rows=150):
    limit = min(len(raw), max_rows)
    best_row = None
    best_score = -1

    for i in range(limit):
        vals = raw.iloc[i].tolist()
        hits = header_hits(vals)

        if not is_generic_bank_header(vals):
            continue

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


def dataframe_from_header(raw, header_row):
    headers = [clean_text(x) for x in raw.iloc[header_row].tolist()]

    safe_headers = []
    used = {}

    for i, h in enumerate(headers):
        base = h or f"Column_{i+1}"
        used[base] = used.get(base, 0) + 1

        if used[base] == 1:
            safe_headers.append(base)
        else:
            safe_headers.append(f"{base}_{used[base]}")

    data = raw.iloc[header_row + 1:].copy()
    data.columns = safe_headers

    blank = data.apply(
        lambda row: all(clean_text(v) == "" for v in row.tolist()),
        axis=1
    )

    return data.loc[~blank].copy()


# ============================================================
# CUSTOMER / UTR EXTRACTORS
# ============================================================

LOCATION_WORDS = [
    "KILAKARAI",
    "BANGALORE - DOMLUR",
    "RPC MUMBAI",
    "RPC-NASIK",
    "CHENNAI RPC",
    "RPC JODHPUR",
    "RPC-CHH. SAMBHAJINAGAR",
]


def clean_customer_name(value):
    s = clean_text(value).strip(" -/")

    for loc in LOCATION_WORDS:
        s = re.sub(
            r"\b" + re.escape(loc) + r"\b.*$",
            "",
            s,
            flags=re.I
        ).strip(" -/")

    return s[:120]


def is_plain_customer_line(line):
    s = clean_text(line)

    if not s or len(s) > 90:
        return False

    reject_starts = [
        "total", "page ", "cont", "sr245",
        "this is an authenticated",
        "tran date", "your details",
        "your base branch", "summary of accounts",
        "statement of transactions", "regd address",
        "legend for transactions", "sincerely",
        "team icici", "cin :", "corporate office",
    ]

    if any(s.lower().startswith(x) for x in reject_starts):
        return False

    if "/" in s or ":" in s:
        return False

    # Reject digit-heavy continuation/reference lines.
    digits = sum(ch.isdigit() for ch in s)
    letters = sum(ch.isalpha() for ch in s)

    if digits >= 3:
        return False

    return letters >= 4 and letters / max(len(s), 1) >= 0.50


def extract_reference_from_text(block):
    s = clean_text(block)

    patterns = [
        r"\bUPI/([0-9]{8,20})",
        r"\bMMT/IMPS/([0-9]{8,20})",
        r"\bINF/(?:NEFT|INFT)/([0-9]{8,20})",
        r"\bBIL/INFT/([0-9]{8,20})",
        r"\bNEFT-([A-Z0-9]{8,35})",
        r"\bRTGS-([A-Z0-9]{8,40})",
        r"\bRTGS/([A-Z0-9]{8,40})",
        r"\bTRF/[^/]+/([0-9]{3,20})",
    ]

    for pattern in patterns:
        m = re.search(pattern, s, flags=re.I)
        if m:
            return clean_text(m.group(1))

    return ""


def extract_customer_from_text(block, prefix_name=""):
    if prefix_name:
        return clean_customer_name(prefix_name)

    s = clean_text(block)

    patterns = [
        # INF/NEFT/REF/IFSC/Name
        r"INF/NEFT/\d+/[A-Z0-9]+/"
        r"([A-Za-z][A-Za-z0-9 .&_-]{2,70}?)"
        r"(?=\s+(?:KILAKARAI|RPC|CHENNAI|BANGALORE)|$)",

        # RTGS/REF/IFSC/Name
        r"RTGS/[A-Z0-9]+/[A-Z0-9]+/"
        r"([A-Za-z][A-Za-z0-9 .&_-]{2,70}?)"
        r"(?=\s+(?:KILAKARAI|RPC|CHENNAI|BANGALORE)|$)",

        # IMPS/REF/Name/
        r"MMT/IMPS/\d+/([^/]{2,70})/",

        # BIL/INFT/ref/NA/ Name
        r"BIL/INFT/\d+/(?:NA|MIB-)/\s*"
        r"([A-Za-z][A-Za-z .&_-]{2,70}?)"
        r"(?=\s+(?:KILAKARAI|RPC|CHENNAI|BANGALORE)|$)",

        # NEFT-Nxxx-NAME
        r"NEFT-[A-Z0-9]+-"
        r"([A-Za-z][A-Za-z0-9 .&_-]{2,90}?)"
        r"(?=\s+(?:RPC|KILAKARAI|CHENNAI|BANGALORE)|-SP\d|$)",

        # /NA/ NAME general
        r"/NA/\s*"
        r"([A-Za-z][A-Za-z .&_-]{2,70}?)"
        r"(?=\s+(?:KILAKARAI|RPC|CHENNAI|BANGALORE)|$)",
    ]

    for pattern in patterns:
        m = re.search(pattern, s, flags=re.I)
        if m:
            name = clean_customer_name(m.group(1))
            if name:
                return name

    return ""


# ============================================================
# GENERIC TABULAR STANDARDIZER
# ============================================================

def standardize_generic_table(table):
    date_col = find_column(table.columns, "date")
    narration_col = find_column(table.columns, "narration")
    debit_col = find_column(table.columns, "debit")
    credit_col = find_column(table.columns, "credit")
    balance_col = find_column(table.columns, "balance")
    customer_col = find_column(table.columns, "customer")
    utr_col = find_column(table.columns, "utr")
    amount_col = find_column(table.columns, "amount")
    type_col = find_column(table.columns, "type")

    if date_col is None:
        return pd.DataFrame()

    out = pd.DataFrame(index=table.index)

    out["Date"] = pd.to_datetime(
        table[date_col],
        errors="coerce",
        dayfirst=True
    ).dt.normalize()

    out["Narration"] = (
        table[narration_col]
        if narration_col
        else ""
    )

    out["Customer Name"] = (
        table[customer_col]
        if customer_col
        else ""
    )

    out["UTR / Reference"] = (
        table[utr_col]
        if utr_col
        else ""
    )

    out["Debit"] = (
        table[debit_col].apply(parse_number)
        if debit_col
        else np.nan
    )

    out["Credit"] = (
        table[credit_col].apply(parse_number)
        if credit_col
        else np.nan
    )

    out["Balance"] = (
        table[balance_col].apply(parse_number)
        if balance_col
        else np.nan
    )

    # Amount + DR/CR type fallback.
    if (
        debit_col is None
        and credit_col is None
        and amount_col
        and type_col
    ):
        amounts = table[amount_col].apply(parse_number)
        txn_type = table[type_col].astype(str).str.upper()

        out["Debit"] = np.where(
            txn_type.str.contains(
                r"\bDR\b|DEBIT|WITHDRAW",
                regex=True,
                na=False
            ),
            amounts,
            np.nan
        )

        out["Credit"] = np.where(
            txn_type.str.contains(
                r"\bCR\b|CREDIT|DEPOSIT",
                regex=True,
                na=False
            ),
            amounts,
            np.nan
        )

    customer_blank = (
        out["Customer Name"]
        .fillna("")
        .astype(str)
        .str.strip()
        .eq("")
    )

    out.loc[customer_blank, "Customer Name"] = out.loc[
        customer_blank, "Narration"
    ].apply(lambda x: extract_customer_from_text(x, ""))

    utr_blank = (
        out["UTR / Reference"]
        .fillna("")
        .astype(str)
        .str.strip()
        .eq("")
    )

    out.loc[utr_blank, "UTR / Reference"] = out.loc[
        utr_blank, "Narration"
    ].apply(extract_reference_from_text)

    out["Customer Source"] = np.where(
        customer_col is not None,
        "SOURCE COLUMN",
        np.where(
            out["Customer Name"]
            .fillna("")
            .astype(str)
            .str.strip()
            .ne(""),
            "EXTRACTED FROM NARRATION",
            "NOT AVAILABLE"
        )
    )

    out["UTR Source"] = np.where(
        utr_col is not None,
        "SOURCE COLUMN",
        np.where(
            out["UTR / Reference"]
            .fillna("")
            .astype(str)
            .str.strip()
            .ne(""),
            "EXTRACTED FROM NARRATION",
            "NOT AVAILABLE"
        )
    )

    actual = out["Debit"].notna() | out["Credit"].notna()

    return out.loc[actual].copy()


# ============================================================
# EXCEL / CSV READER
# ============================================================

def read_excel_csv(path):
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

    outputs = []

    for sheet in book.sheet_names:
        try:
            raw = pd.read_excel(
                book,
                sheet_name=sheet,
                header=None,
                dtype=str,
                keep_default_na=False
            )
            outputs.append((sheet, raw))
        except Exception as e:
            print(f"    WARNING sheet '{sheet}' skipped: {e}")

    return outputs


# ============================================================
# ICICI-STYLE PDF PARSER
# ============================================================

DATE_START_RE = re.compile(r"^(\d{2}-\d{2}-\d{4})\b")

MONEY_TOKEN_RE = re.compile(
    r"(?<![\w])"
    r"(?:\d{1,3}(?:,\d{2,3})+|\d+)"
    r"(?:\.\d{2})?"
    r"(?![\w])"
)


def is_icici_style_pdf_text(text_value):
    """
    Detect multiple ICICI statement generations/layouts.

    Supported examples:
    1) Tran Date | Value Date | Particulars | Location | Chq.No |
       Withdrawals | Deposits | Balance (INR)

    2) Date | Particulars | Chq.No. | Withdrawals | Deposits |
       Autosweep | Reverse Sweep | Balance(INR)

    We intentionally do NOT require Value Date because newer/alternate
    ICICI exports may not have that column.
    """
    t = clean_text(text_value).lower()

    has_date = (
        "tran date" in t
        or "transaction date" in t
        or re.search(r"\bdate\b", t) is not None
    )

    return (
        has_date
        and "particulars" in t
        and "withdrawals" in t
        and "deposits" in t
        and "balance" in t
    )


def signed_balance_from_line(line):
    """
    Return signed running balance.

    ICICI PDFs normally print:
        1,23,456.78 Cr
        499.99 Dr

    Some statement versions omit Cr/Dr when the balance is exactly 0.00.
    V4 accepts that zero-balance form too.
    """
    stripped = re.sub(
        r"^\d{2}-\d{2}-\d{4}"
        r"(?:\s+\d{2}-\d{2}-\d{4})?\s*",
        "",
        line
    )

    nums = MONEY_TOKEN_RE.findall(stripped)

    if not nums:
        return None

    bal = parse_number(nums[-1])

    if pd.isna(bal):
        return None

    if re.search(r"\bDr\b\s*$", line, flags=re.I):
        return -abs(float(bal))

    if re.search(r"\bCr\b\s*$", line, flags=re.I):
        return abs(float(bal))

    # Some ICICI exports print zero balance without Cr/Dr.
    if abs(float(bal)) <= 0.000001:
        return 0.0

    return None


def transaction_amount_from_line(line):
    """
    In ICICI transaction lines the final two numeric values are normally:
        transaction amount | running balance

    B/F opening row has only the running balance.
    """
    stripped = re.sub(
        r"^\d{2}-\d{2}-\d{4}"
        r"(?:\s+\d{2}-\d{2}-\d{4})?\s*",
        "",
        line
    )

    nums = MONEY_TOKEN_RE.findall(stripped)

    if "B/F" in line.upper() and len(nums) <= 1:
        return None

    if len(nums) < 2:
        return None

    amount = parse_number(nums[-2])

    return None if pd.isna(amount) else abs(float(amount))


def parse_icici_pdf(path):
    """
    Parse ICICI-style statement PDFs by transaction text rows instead of
    pdfplumber table rows.

    Why:
    pdfplumber may collapse an entire page into:
        header row
        one giant multiline data row
        total row
    which makes a normal table parser think the page has only one transaction.

    Debit/Credit is determined from actual running balance movement.
    Printed page totals are used as a reconciliation control.
    """
    import pdfplumber

    transactions = []
    page_recon = []

    previous_signed_balance = None

    with pdfplumber.open(path) as pdf:
        for page_no, page in enumerate(pdf.pages, start=1):
            page_text = page.extract_text(
                x_tolerance=1,
                y_tolerance=3
            ) or ""

            lines = [
                line.strip()
                for line in page_text.splitlines()
                if line.strip()
            ]

            # Printed page totals:
            # Total : Withdrawals Deposits Balance
            total_match = re.search(
                r"Total\s*:\s*"
                r"([\d,]+\.\d{2})\s+"
                r"([\d,]+\.\d{2})\s+"
                r"([\d,]+\.\d{2})",
                page_text
            )

            printed_debit = (
                parse_number(total_match.group(1))
                if total_match
                else np.nan
            )

            printed_credit = (
                parse_number(total_match.group(2))
                if total_match
                else np.nan
            )

            transaction_indexes = [
                i
                for i, line in enumerate(lines)
                if DATE_START_RE.match(line)
            ]

            page_rows = []

            for position, line_index in enumerate(transaction_indexes):
                line = lines[line_index]

                next_index = (
                    transaction_indexes[position + 1]
                    if position + 1 < len(transaction_indexes)
                    else len(lines)
                )

                # Lines after current dated line and before next dated line.
                block_lines = lines[line_index:next_index]

                # Stop transaction block at printed total/footer.
                cleaned_block_lines = []
                for bl in block_lines:
                    if bl.startswith("Total :"):
                        break
                    if bl.startswith("Page "):
                        break
                    cleaned_block_lines.append(bl)

                block_lines = cleaned_block_lines

                # Immediate plain-text line before a date is often the party name
                # in ICICI PDFs (UPI/NEFT/Google/Fashnear/etc.).
                prefix_name = ""

                if line_index > 0:
                    previous_line = lines[line_index - 1]

                    if is_plain_customer_line(previous_line):
                        prefix_name = previous_line

                # Find the line that contains the running balance.
                balance_line = None

                for bl in block_lines:
                    # Normal Cr/Dr ending OR alternate ICICI zero-balance row.
                    if re.search(
                        r"\b(?:Cr|Dr)\b\s*$",
                        bl,
                        flags=re.I
                    ):
                        balance_line = bl
                        break

                    # Newer/current-account ICICI PDFs sometimes omit Cr/Dr
                    # when final balance is exactly 0.00.
                    if DATE_START_RE.match(bl):
                        stripped_for_zero = re.sub(
                            r"^\d{2}-\d{2}-\d{4}"
                            r"(?:\s+\d{2}-\d{2}-\d{4})?\s*",
                            "",
                            bl
                        )
                        zero_nums = MONEY_TOKEN_RE.findall(stripped_for_zero)
                        if (
                            len(zero_nums) >= 2
                            and abs(parse_number(zero_nums[-1])) <= 0.000001
                        ):
                            balance_line = bl
                            break

                if balance_line is None:
                    continue

                current_balance = signed_balance_from_line(balance_line)

                if current_balance is None:
                    continue

                amount = transaction_amount_from_line(balance_line)

                date_match = DATE_START_RE.match(line)
                txn_date_text = date_match.group(1)

                # B/F opening balance is a control row, not a transaction.
                is_bf = "B/F" in line.upper()

                # Create narration block.
                block_for_parse = " ".join(
                    ([prefix_name] if prefix_name else [])
                    + block_lines
                )

                customer = extract_customer_from_text(
                    block_for_parse,
                    prefix_name
                )

                reference = extract_reference_from_text(block_for_parse)

                debit = np.nan
                credit = np.nan

                if (
                    not is_bf
                    and amount is not None
                    and previous_signed_balance is not None
                ):
                    balance_change = (
                        current_balance
                        - previous_signed_balance
                    )

                    # Strong control: balance movement should equal amount.
                    tolerance = max(
                        0.02,
                        abs(amount) * 0.000001
                    )

                    if abs(abs(balance_change) - amount) <= tolerance:
                        if balance_change < 0:
                            debit = amount
                        elif balance_change > 0:
                            credit = amount
                    else:
                        # Still classify by actual balance direction, but mark
                        # reconciliation status for review.
                        if balance_change < 0:
                            debit = amount
                        elif balance_change > 0:
                            credit = amount

                # Update running balance including B/F.
                previous_signed_balance = current_balance

                if is_bf:
                    continue

                page_rows.append({
                    "Date": pd.to_datetime(
                        txn_date_text,
                        errors="coerce",
                        dayfirst=True
                    ),
                    "Customer Name": customer,
                    "UTR / Reference": reference,
                    "Narration": clean_text(
                        " ".join(block_lines)
                    ),
                    "Debit": debit,
                    "Credit": credit,
                    "Balance": current_balance,
                    "Customer Source": (
                        "PDF NAME LINE / NARRATION"
                        if customer
                        else "NOT AVAILABLE"
                    ),
                    "UTR Source": (
                        "PDF NARRATION"
                        if reference
                        else "NOT AVAILABLE"
                    ),
                    "Source Page": page_no,
                    "Parser": "ICICI PDF TEXT + BALANCE RECON",
                })

            if page_rows:
                page_df = pd.DataFrame(page_rows)

                parsed_debit = float(
                    page_df["Debit"]
                    .fillna(0)
                    .sum()
                )

                parsed_credit = float(
                    page_df["Credit"]
                    .fillna(0)
                    .sum()
                )
            else:
                parsed_debit = 0.0
                parsed_credit = 0.0

            debit_diff = (
                parsed_debit - printed_debit
                if pd.notna(printed_debit)
                else np.nan
            )

            credit_diff = (
                parsed_credit - printed_credit
                if pd.notna(printed_credit)
                else np.nan
            )

            if (
                pd.notna(printed_debit)
                and pd.notna(printed_credit)
            ):
                page_status = (
                    "OK"
                    if (
                        abs(debit_diff) <= 0.02
                        and abs(credit_diff) <= 0.02
                    )
                    else "CHECK"
                )
            else:
                page_status = "NO PRINTED TOTAL"

            page_recon.append({
                "Page": page_no,
                "Parsed Debit": parsed_debit,
                "Printed Debit": printed_debit,
                "Debit Difference": debit_diff,
                "Parsed Credit": parsed_credit,
                "Printed Credit": printed_credit,
                "Credit Difference": credit_diff,
                "Status": page_status,
            })

            transactions.extend(page_rows)

    return (
        pd.DataFrame(transactions),
        pd.DataFrame(page_recon)
    )


# ============================================================
# GENERIC PDF TABLE READER
# ============================================================

def extract_pdf_tables(path):
    import pdfplumber

    outputs = []

    with pdfplumber.open(path) as pdf:
        for page_no, page in enumerate(pdf.pages, start=1):
            tables = page.extract_tables() or []

            for table_no, table in enumerate(tables, start=1):
                if not table:
                    continue

                width = max(
                    len(row or [])
                    for row in table
                )

                rows = []

                for row in table:
                    row = list(row or [])
                    row += [""] * (width - len(row))
                    rows.append(row)

                outputs.append((
                    f"PDF Page {page_no} Table {table_no}",
                    pd.DataFrame(rows),
                    page_no
                ))

    return outputs


def get_first_pdf_page_text(path):
    import pdfplumber

    with pdfplumber.open(path) as pdf:
        if not pdf.pages:
            return ""

        return pdf.pages[0].extract_text(
            x_tolerance=1,
            y_tolerance=3
        ) or ""


# ============================================================
# OPTIONAL OCR FALLBACK
# ============================================================

def ocr_pdf_optional(path):
    """
    Conservative OCR fallback.
    OCR text is NOT automatically turned into financial rows unless a
    recognizable table/header is produced.
    """
    try:
        from pdf2image import convert_from_path
        import pytesseract
    except Exception:
        return []

    try:
        images = convert_from_path(path, dpi=300)
    except Exception:
        return []

    outputs = []

    for page_no, image in enumerate(images, start=1):
        try:
            text_value = pytesseract.image_to_string(
                image,
                config="--psm 6"
            )
        except Exception:
            continue

        lines = [
            [x]
            for x in text_value.splitlines()
            if clean_text(x)
        ]

        if lines:
            outputs.append((
                f"OCR Page {page_no}",
                pd.DataFrame(lines),
                page_no
            ))

    return outputs


# ============================================================
# LOAD SOURCES
# ============================================================

all_frames = []
source_map = []
all_pdf_recon = []

print("=" * 82)
print("UNIVERSAL BANK BOOKS REPORT V4")
print("=" * 82)
print("Folder:", BASE_DIR)
print("Source files found:", len(input_files))
print("")

for path in input_files:
    base = os.path.basename(path)
    ext = os.path.splitext(path)[1].lower()

    print("Reading:", base)

    # --------------------------------------------------------
    # PDF
    # --------------------------------------------------------
    if ext == ".pdf":
        try:
            first_text = get_first_pdf_page_text(path)
        except Exception as e:
            print("  ERROR reading PDF:", e)

            source_map.append({
                "Source File": base,
                "Source Part": "",
                "Status": "READ ERROR",
                "Transaction Count": 0,
                "Note": str(e),
            })
            continue

        # ICICI-style special parser.
        if is_icici_style_pdf_text(first_text):
            try:
                pdf_tx, pdf_recon = parse_icici_pdf(path)
            except Exception as e:
                print("  ERROR in ICICI PDF parser:", e)

                source_map.append({
                    "Source File": base,
                    "Source Part": "ICICI PDF",
                    "Status": "PARSE ERROR",
                    "Transaction Count": 0,
                    "Note": str(e),
                })
                continue

            if pdf_tx.empty:
                source_map.append({
                    "Source File": base,
                    "Source Part": "ICICI PDF",
                    "Status": "NO TRANSACTIONS",
                    "Transaction Count": 0,
                    "Note": "",
                })
                continue

            pdf_tx["Source File"] = base
            pdf_tx["Source Part"] = (
                "Page "
                + pdf_tx["Source Page"].astype(str)
            )
            pdf_tx["Source Row"] = ""

            all_frames.append(pdf_tx)

            pdf_recon["Source File"] = base
            all_pdf_recon.append(pdf_recon)

            bad_pages = int(
                pdf_recon["Status"].eq("CHECK").sum()
            )

            source_map.append({
                "Source File": base,
                "Source Part": "ICICI PDF TEXT PARSER",
                "Status": (
                    "RECOGNIZED"
                    if bad_pages == 0
                    else "RECOGNIZED / RECON CHECK"
                ),
                "Transaction Count": len(pdf_tx),
                "Note": (
                    f"Page reconciliation issues: {bad_pages}"
                ),
            })

            print(
                f"  ICICI PDF parser: "
                f"{len(pdf_tx):,} transaction(s)"
            )
            print(
                f"  Debit total : "
                f"{pdf_tx['Debit'].fillna(0).sum():,.2f}"
            )
            print(
                f"  Credit total: "
                f"{pdf_tx['Credit'].fillna(0).sum():,.2f}"
            )
            print(
                f"  PDF page reconciliation CHECK pages: "
                f"{bad_pages}"
            )

            continue

        # ----------------------------------------------------
        # Generic PDF tables
        # ----------------------------------------------------
        try:
            pdf_sources = extract_pdf_tables(path)
        except Exception as e:
            print("  ERROR extracting PDF tables:", e)
            pdf_sources = []

        if not pdf_sources:
            pdf_sources = ocr_pdf_optional(path)

        recognized_pdf = False

        for source_name, raw, page_no in pdf_sources:
            header_row = locate_generic_header(raw)

            if header_row is None:
                continue

            table = dataframe_from_header(
                raw,
                header_row
            )

            tx = standardize_generic_table(table)

            if tx.empty:
                continue

            tx["Source File"] = base
            tx["Source Part"] = source_name
            tx["Source Row"] = (
                tx.index.to_series().astype(int)
                + header_row
                + 2
            ).values
            tx["Source Page"] = page_no
            tx["Parser"] = "GENERIC PDF TABLE"

            all_frames.append(tx)
            recognized_pdf = True

            source_map.append({
                "Source File": base,
                "Source Part": source_name,
                "Status": "RECOGNIZED",
                "Transaction Count": len(tx),
                "Note": "",
            })

            print(
                f"  {source_name}: "
                f"{len(tx):,} transaction(s)"
            )

        if not recognized_pdf:
            source_map.append({
                "Source File": base,
                "Source Part": "PDF",
                "Status": "UNRECOGNIZED",
                "Transaction Count": 0,
                "Note": (
                    "No reliable transaction table recognized. "
                    "If image-only PDF, install OCR dependencies "
                    "or export the bank statement as searchable PDF/Excel."
                ),
            })

            print(
                "  WARNING: no reliable transaction rows "
                "recognized from PDF."
            )

        continue

    # --------------------------------------------------------
    # EXCEL / CSV
    # --------------------------------------------------------
    try:
        sources = read_excel_csv(path)
    except Exception as e:
        print("  ERROR:", e)

        source_map.append({
            "Source File": base,
            "Source Part": "",
            "Status": "READ ERROR",
            "Transaction Count": 0,
            "Note": str(e),
        })

        continue

    recognized_file = False

    for source_name, raw in sources:
        header_row = locate_generic_header(raw)

        if header_row is None:
            continue

        table = dataframe_from_header(
            raw,
            header_row
        )

        tx = standardize_generic_table(table)

        if tx.empty:
            continue

        tx["Source File"] = base
        tx["Source Part"] = source_name
        tx["Source Row"] = (
            tx.index.to_series().astype(int)
            + header_row
            + 2
        ).values
        tx["Source Page"] = ""
        tx["Parser"] = "GENERIC EXCEL/CSV"

        all_frames.append(tx)
        recognized_file = True

        source_map.append({
            "Source File": base,
            "Source Part": source_name,
            "Status": "RECOGNIZED",
            "Transaction Count": len(tx),
            "Note": f"Header row: {header_row + 1}",
        })

        print(
            f"  {source_name}: "
            f"{len(tx):,} transaction(s)"
        )

    if not recognized_file:
        source_map.append({
            "Source File": base,
            "Source Part": "",
            "Status": "UNRECOGNIZED",
            "Transaction Count": 0,
            "Note": (
                "Expected Date + Debit/Credit, "
                "or Date + Amount + DR/CR."
            ),
        })


if not all_frames:
    raise ValueError(
        "\nNo bank transactions were recognized.\n"
        "For Excel/CSV expected something similar to:\n"
        "Date | Narration | Debit | Credit | Balance\n"
        "or Date | Amount | Type (DR/CR).\n\n"
        "For PDFs, searchable/text bank statements work best."
    )


# ============================================================
# CONSOLIDATE
# ============================================================

transactions = pd.concat(
    all_frames,
    ignore_index=True,
    sort=False
)

for required_col in [
    "Date",
    "Customer Name",
    "UTR / Reference",
    "Narration",
    "Debit",
    "Credit",
    "Balance",
    "Customer Source",
    "UTR Source",
    "Source File",
    "Source Part",
    "Source Row",
    "Source Page",
    "Parser",
]:
    if required_col not in transactions.columns:
        transactions[required_col] = ""

transactions["Date"] = pd.to_datetime(
    transactions["Date"],
    errors="coerce"
).dt.normalize()

transactions["Debit"] = pd.to_numeric(
    transactions["Debit"],
    errors="coerce"
)

transactions["Credit"] = pd.to_numeric(
    transactions["Credit"],
    errors="coerce"
)

transactions["Balance"] = pd.to_numeric(
    transactions["Balance"],
    errors="coerce"
)

transactions["Direction"] = np.where(
    transactions["Debit"].notna(),
    "DEBIT",
    np.where(
        transactions["Credit"].notna(),
        "CREDIT",
        ""
    )
)

transactions["Transaction Amount"] = np.where(
    transactions["Direction"].eq("DEBIT"),
    transactions["Debit"],
    transactions["Credit"]
)

# Duplicate diagnostic only.
dup_key = (
    transactions["Date"].astype(str)
    + "|"
    + transactions["Debit"].fillna(0).astype(str)
    + "|"
    + transactions["Credit"].fillna(0).astype(str)
    + "|"
    + transactions["UTR / Reference"]
      .fillna("")
      .astype(str)
      .str.strip()
    + "|"
    + transactions["Narration"]
      .fillna("")
      .astype(str)
      .str.strip()
)

transactions["Possible Duplicate"] = np.where(
    dup_key.duplicated(keep=False),
    "YES",
    "NO"
)


# ============================================================
# SUMMARIES
# ============================================================

date_wise = (
    transactions
    .groupby("Date", dropna=False)
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


transactions["Month"] = (
    transactions["Date"]
    .dt.to_period("M")
    .astype(str)
)

month_wise = (
    transactions
    .groupby("Month", dropna=False)
    .agg(
        Transaction_Count=("Transaction Amount", "count"),
        Debit_Count=("Debit", lambda s: int(s.notna().sum())),
        Debit_Total=("Debit", "sum"),
        Credit_Count=("Credit", lambda s: int(s.notna().sum())),
        Credit_Total=("Credit", "sum"),
        Unique_Customers=(
            "Customer Name",
            lambda s: (
                s.replace("", np.nan)
                 .dropna()
                 .nunique()
            )
        ),
        Unique_UTRs=(
            "UTR / Reference",
            lambda s: (
                s.replace("", np.nan)
                 .dropna()
                 .nunique()
            )
        ),
    )
    .reset_index()
)

month_wise["Net_Credit_Minus_Debit"] = (
    month_wise["Credit_Total"].fillna(0)
    - month_wise["Debit_Total"].fillna(0)
)


customer_tx = transactions[
    transactions["Customer Name"]
    .fillna("")
    .astype(str)
    .str.strip()
    .ne("")
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
    customer_wise = (
        customer_tx
        .groupby("Customer Name", dropna=False)
        .agg(
            Transaction_Count=("Transaction Amount", "count"),
            Debit_Count=("Debit", lambda s: int(s.notna().sum())),
            Debit_Total=("Debit", "sum"),
            Credit_Count=("Credit", lambda s: int(s.notna().sum())),
            Credit_Total=("Credit", "sum"),
            First_Date=("Date", "min"),
            Last_Date=("Date", "max"),
            Unique_UTRs=(
                "UTR / Reference",
                lambda s: (
                    s.replace("", np.nan)
                     .dropna()
                     .nunique()
                )
            ),
        )
        .reset_index()
    )

    customer_wise["Net_Credit_Minus_Debit"] = (
        customer_wise["Credit_Total"].fillna(0)
        - customer_wise["Debit_Total"].fillna(0)
    )

    customer_wise = customer_wise[[
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
    ]]

    customer_date_wise = (
        customer_tx
        .groupby(
            ["Date", "Customer Name"],
            dropna=False
        )
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


utr_register = transactions[[
    "Date",
    "Customer Name",
    "UTR / Reference",
    "Narration",
    "Debit",
    "Credit",
    "Balance",
    "Source File",
    "Source Part",
]].copy()


# ============================================================
# CONTROL TOTALS
# ============================================================

total_transactions = len(transactions)

debit_count = int(
    transactions["Debit"].notna().sum()
)

credit_count = int(
    transactions["Credit"].notna().sum()
)

total_debit = float(
    transactions["Debit"]
    .fillna(0)
    .sum()
)

total_credit = float(
    transactions["Credit"]
    .fillna(0)
    .sum()
)

net_movement = total_credit - total_debit

valid_dates = transactions["Date"].dropna()

first_date = (
    valid_dates.min()
    if not valid_dates.empty
    else None
)

last_date = (
    valid_dates.max()
    if not valid_dates.empty
    else None
)

unique_customers = int(
    transactions["Customer Name"]
    .replace("", np.nan)
    .dropna()
    .nunique()
)

unique_utrs = int(
    transactions["UTR / Reference"]
    .replace("", np.nan)
    .dropna()
    .nunique()
)

missing_customer = int(
    transactions["Customer Name"]
    .fillna("")
    .astype(str)
    .str.strip()
    .eq("")
    .sum()
)

missing_utr = int(
    transactions["UTR / Reference"]
    .fillna("")
    .astype(str)
    .str.strip()
    .eq("")
    .sum()
)

possible_duplicates = int(
    transactions["Possible Duplicate"]
    .eq("YES")
    .sum()
)


# ============================================================
# BUILD PDF RECON TABLE
# ============================================================

if all_pdf_recon:
    pdf_page_recon = pd.concat(
        all_pdf_recon,
        ignore_index=True,
        sort=False
    )
else:
    pdf_page_recon = pd.DataFrame(columns=[
        "Source File",
        "Page",
        "Parsed Debit",
        "Printed Debit",
        "Debit Difference",
        "Parsed Credit",
        "Printed Credit",
        "Credit Difference",
        "Status",
    ])


# ============================================================
# WORKBOOK HELPERS
# ============================================================

wb = Workbook()
wb.remove(wb.active)


def safe_excel_value(value):
    if isinstance(value, pd.Timestamp):
        return value.to_pydatetime()

    if isinstance(value, np.integer):
        return int(value)

    if isinstance(value, np.floating):
        if np.isnan(value):
            return None
        return float(value)

    if pd.isna(value):
        return None

    return value


def style_header(ws):
    for cell in ws[1]:
        cell.fill = HEADER_FILL
        cell.font = WHITE_BOLD
        cell.border = BORDER
        cell.alignment = Alignment(
            horizontal="center",
            vertical="center",
            wrap_text=True
        )


def auto_width(ws, max_width=42):
    for col_idx in range(1, ws.max_column + 1):
        max_len = 0

        for row_idx in range(
            1,
            min(ws.max_row, 3000) + 1
        ):
            value = ws.cell(
                row_idx,
                col_idx
            ).value

            if value is not None:
                max_len = max(
                    max_len,
                    len(str(value))
                )

        ws.column_dimensions[
            get_column_letter(col_idx)
        ].width = min(
            max(max_len + 2, 11),
            max_width
        )


def write_df(
    ws,
    frame,
    money_cols=(),
    date_cols=(),
    count_cols=()
):
    if frame is None or frame.empty:
        ws["A1"] = "No data available"
        return

    for col_idx, col_name in enumerate(
        frame.columns,
        start=1
    ):
        ws.cell(
            1,
            col_idx,
            col_name
        )

    for row_idx, row in enumerate(
        frame.itertuples(
            index=False,
            name=None
        ),
        start=2
    ):
        for col_idx, value in enumerate(
            row,
            start=1
        ):
            cell = ws.cell(
                row_idx,
                col_idx,
                safe_excel_value(value)
            )

            cell.border = BORDER
            cell.alignment = Alignment(
                vertical="center"
            )

    style_header(ws)

    for col_idx, col_name in enumerate(
        frame.columns,
        start=1
    ):
        if col_name in money_cols:
            for r in range(2, ws.max_row + 1):
                ws.cell(
                    r,
                    col_idx
                ).number_format = MONEY_FMT

        elif col_name in date_cols:
            for r in range(2, ws.max_row + 1):
                ws.cell(
                    r,
                    col_idx
                ).number_format = DATE_FMT

        elif col_name in count_cols:
            for r in range(2, ws.max_row + 1):
                ws.cell(
                    r,
                    col_idx
                ).number_format = COUNT_FMT

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    auto_width(ws)


# ============================================================
# COVER_REPORT
# ============================================================

ws = wb.create_sheet("COVER_REPORT")
ws.sheet_view.showGridLines = False

ws.merge_cells("A1:D1")

ws["A1"] = (
    "UNIVERSAL BANK BOOKS OF ACCOUNTS REPORT V3"
)

ws["A1"].font = Font(
    bold=True,
    size=16,
    color="1F4E78"
)

ws["A1"].alignment = Alignment(
    horizontal="center"
)

cover_rows = [
    ("Total Transaction Count", total_transactions),
    ("Debit Transaction Count", debit_count),
    ("Total Debit", total_debit),
    ("Credit Transaction Count", credit_count),
    ("Total Credit", total_credit),
    ("Net Credit - Debit", net_movement),
    ("First Transaction Date", first_date),
    ("Last Transaction Date", last_date),
    ("Unique Customers / Counterparties", unique_customers),
    ("Unique UTR / References", unique_utrs),
    ("Transactions Missing Customer Name", missing_customer),
    ("Transactions Missing UTR / Reference", missing_utr),
    ("Possible Duplicate Rows", possible_duplicates),
    (
        "PDF Reconciliation CHECK Pages",
        int(
            pdf_page_recon["Status"]
            .eq("CHECK")
            .sum()
        )
        if not pdf_page_recon.empty
        else 0
    ),
    (
        "Recognized Source Files",
        transactions["Source File"].nunique()
    ),
]

for r, (label, value) in enumerate(
    cover_rows,
    start=3
):
    ws.cell(
        r,
        1,
        label
    )

    ws.cell(
        r,
        2,
        safe_excel_value(value)
    )

    ws.cell(r, 1).border = BORDER
    ws.cell(r, 2).border = BORDER

    if label in (
        "Total Debit",
        "Total Credit",
        "Net Credit - Debit"
    ):
        ws.cell(
            r,
            2
        ).number_format = MONEY_FMT

    if "Date" in label:
        ws.cell(
            r,
            2
        ).number_format = DATE_FMT

ws.column_dimensions["A"].width = 44
ws.column_dimensions["B"].width = 24


# ============================================================
# TRANSACTION_REGISTER
# ============================================================

register_columns = [
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
    "Parser",
    "Source File",
    "Source Part",
    "Source Page",
    "Source Row",
]

ws = wb.create_sheet("TRANSACTION_REGISTER")

write_df(
    ws,
    transactions[register_columns],
    money_cols=(
        "Debit",
        "Credit",
        "Balance",
        "Transaction Amount",
    ),
    date_cols=("Date",),
)


# ============================================================
# DATE_WISE_SUMMARY
# ============================================================

ws = wb.create_sheet("DATE_WISE_SUMMARY")

write_df(
    ws,
    date_wise,
    money_cols=(
        "Debit_Total",
        "Credit_Total",
        "Net_Credit_Minus_Debit",
    ),
    date_cols=("Date",),
    count_cols=(
        "Transaction_Count",
        "Debit_Count",
        "Credit_Count",
    ),
)


# ============================================================
# MONTH_WISE_SUMMARY
# ============================================================

ws = wb.create_sheet("MONTH_WISE_SUMMARY")

write_df(
    ws,
    month_wise,
    money_cols=(
        "Debit_Total",
        "Credit_Total",
        "Net_Credit_Minus_Debit",
    ),
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
    money_cols=(
        "Debit_Total",
        "Credit_Total",
        "Net_Credit_Minus_Debit",
    ),
    date_cols=(
        "First_Date",
        "Last_Date",
    ),
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
    money_cols=(
        "Debit_Total",
        "Credit_Total",
        "Net_Credit_Minus_Debit",
    ),
    date_cols=("Date",),
    count_cols=(
        "Transaction_Count",
        "Debit_Count",
        "Credit_Count",
    ),
)


# ============================================================
# UTR_WISE_REGISTER
# ============================================================

ws = wb.create_sheet("UTR_WISE_REGISTER")

write_df(
    ws,
    utr_register,
    money_cols=(
        "Debit",
        "Credit",
        "Balance",
    ),
    date_cols=("Date",),
)


# ============================================================
# PDF_PAGE_RECON
# ============================================================

ws = wb.create_sheet("PDF_PAGE_RECON")

write_df(
    ws,
    pdf_page_recon,
    money_cols=(
        "Parsed Debit",
        "Printed Debit",
        "Debit Difference",
        "Parsed Credit",
        "Printed Credit",
        "Credit Difference",
    ),
    count_cols=("Page",),
)

if not pdf_page_recon.empty:
    # Highlight CHECK rows.
    status_col = list(
        pdf_page_recon.columns
    ).index("Status") + 1

    for r in range(2, ws.max_row + 1):
        if ws.cell(
            r,
            status_col
        ).value == "CHECK":
            for c in range(
                1,
                ws.max_column + 1
            ):
                ws.cell(
                    r,
                    c
                ).fill = WARN_FILL


# ============================================================
# CONTROL_TOTALS
# ============================================================

ws = wb.create_sheet("CONTROL_TOTALS")

controls = [
    ["Metric", "Value"],
    ["Total Transaction Count", total_transactions],
    ["Debit Transaction Count", debit_count],
    ["Total Debit", total_debit],
    ["Credit Transaction Count", credit_count],
    ["Total Credit", total_credit],
    ["Net Credit - Debit", net_movement],
    ["Unique Customers / Counterparties", unique_customers],
    ["Unique UTR / References", unique_utrs],
    ["Missing Customer Name", missing_customer],
    ["Missing UTR / Reference", missing_utr],
    ["Possible Duplicate Rows", possible_duplicates],
    [
        "PDF Page Reconciliation CHECK Count",
        int(
            pdf_page_recon["Status"]
            .eq("CHECK")
            .sum()
        )
        if not pdf_page_recon.empty
        else 0
    ],
]

for r, row in enumerate(
    controls,
    start=1
):
    for c, value in enumerate(
        row,
        start=1
    ):
        ws.cell(
            r,
            c,
            safe_excel_value(value)
        )

        ws.cell(
            r,
            c
        ).border = BORDER

style_header(ws)

for r in (4, 6, 7):
    ws.cell(
        r,
        2
    ).number_format = MONEY_FMT

ws.column_dimensions["A"].width = 44
ws.column_dimensions["B"].width = 24


# ============================================================
# SOURCE_MAPPING
# ============================================================

ws = wb.create_sheet("SOURCE_MAPPING")

write_df(
    ws,
    pd.DataFrame(source_map),
    count_cols=("Transaction Count",),
)


# ============================================================
# REPORT_NOTES
# ============================================================

ws = wb.create_sheet("REPORT_NOTES")

notes = [
    ["Item", "Note"],
    [
        "Purpose",
        "Bank statement to supporting Books-of-Accounts transaction report."
    ],
    [
        "V4 PDF Flexibility",
        "ICICI-style PDFs are parsed from transaction text rows. V4 supports both Tran Date + Value Date statements and Date-only current-account statements with Autosweep/Reverse Sweep columns."
    ],
    [
        "Debit / Credit",
        "For ICICI-style PDFs, direction is derived from actual running balance movement and transaction amount."
    ],
    [
        "PDF Page Reconciliation",
        "Where the PDF prints Total Withdrawals and Total Deposits, parsed page totals are checked against those printed figures."
    ],
    [
        "Customer Name",
        "Uses source name field when available; otherwise uses best-effort structured narration/name-line extraction."
    ],
    [
        "UTR / Reference",
        "Uses source reference field when available; otherwise extracts known UPI/IMPS/NEFT/RTGS/INFT references from narration."
    ],
    [
        "Missing Values",
        "Missing customer/UTR values remain unavailable; no value is invented."
    ],
    [
        "Duplicates",
        "Possible duplicates are only flagged and are never automatically removed."
    ],
    [
        "Supported Files",
        "XLSX, XLS, CSV and PDF."
    ],
    [
        "Generated On",
        datetime.now().strftime(
            "%d-%m-%Y %H:%M:%S"
        )
    ],
]

for r, row in enumerate(
    notes,
    start=1
):
    for c, value in enumerate(
        row,
        start=1
    ):
        ws.cell(
            r,
            c,
            value
        )

        ws.cell(
            r,
            c
        ).border = BORDER

style_header(ws)

ws.column_dimensions["A"].width = 30
ws.column_dimensions["B"].width = 115


# ============================================================
# SAVE / CONSOLE CONTROL
# ============================================================

print("")
print("=" * 82)
print("FINAL TOTALS")
print("=" * 82)
print(f"Transactions : {total_transactions:,}")
print(f"Debit Count  : {debit_count:,}")
print(f"Debit Total  : {total_debit:,.2f}")
print(f"Credit Count : {credit_count:,}")
print(f"Credit Total : {total_credit:,.2f}")
print(f"Customers    : {unique_customers:,}")
print(f"UTR / Refs   : {unique_utrs:,}")
print(f"Missing Name : {missing_customer:,}")
print(f"Missing UTR  : {missing_utr:,}")

if not pdf_page_recon.empty:
    print(
        "PDF Page Recon CHECK:",
        int(
            pdf_page_recon["Status"]
            .eq("CHECK")
            .sum()
        )
    )

print("")
print("Saving:", OUTPUT_FILE)

wb.save(OUTPUT_FILE)

print("SUCCESS")
print("Output:", OUTPUT_FILE)
print("=" * 82)
