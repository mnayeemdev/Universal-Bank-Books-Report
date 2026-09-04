"""
============================================================
UNIVERSAL BANK BOOKS REPORT V5
============================================================
Author   : Mohamed Nayeem
Copyright: © Mohamed Nayeem — All Rights Reserved
Instagram: @mohamednayeem7
============================================================

V5 — UNIVERSAL BANK SUPPORT
----------------------------
Supports ALL major Indian bank statement formats:
  SBI, HDFC, ICICI, Axis, Kotak Mahindra, PNB, BOB,
  Canara, Union Bank, IDBI, Yes Bank, IndusInd, Federal,
  RBL, Bandhan, AU Small Finance, Ujjivan, Equitas,
  Jana, Suryoday, Utkarsh, and any bank that produces
  a standard tabular statement.

File types:  .xlsx  .xls  .csv  .pdf

Output:  UNIVERSAL_BANK_BOOKS_REPORT_V5.xlsx

Sheets:
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

Core fields per transaction:
  Date, Customer Name, UTR / Reference, Narration,
  Debit, Credit, Balance, Direction, Transaction Amount,
  Customer Source, UTR Source, Possible Duplicate,
  Parser, Source File, Source Part, Source Page, Source Row

DATA INTEGRITY
  - No transaction is invented.
  - No missing name, UTR, debit, or credit is manufactured.
  - Customer/UTR extraction from narration is best-effort
    and explicitly marked.
  - Possible duplicates are flagged, never deleted.
  - PDF page totals are used as a control check.

INSTALL
  python -m pip install pandas openpyxl xlrd pdfplumber
Optional OCR:
  python -m pip install pytesseract pdf2image pillow

RUN
  python UNIVERSAL_BANK_BOOKS_REPORT_V5.py
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
# AUTHOR / BRANDING
# ============================================================
AUTHOR_NAME = "Mohamed Nayeem"
AUTHOR_COPYRIGHT = "© Mohamed Nayeem — All Rights Reserved"
AUTHOR_INSTAGRAM = "@mohamednayeem7"
VERSION = "V5"

# ============================================================
# CONFIG
# ============================================================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_FILE = os.path.join(
    BASE_DIR,
    f"UNIVERSAL_BANK_BOOKS_REPORT_{VERSION}.xlsx",
)
PATTERNS = (
    "*.xlsx", "*.XLSX",
    "*.xls",  "*.XLS",
    "*.csv",  "*.CSV",
    "*.pdf",  "*.PDF",
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
AUTHOR_FILL = PatternFill("solid", fgColor="F2F2F2")
WHITE_BOLD = Font(color="FFFFFF", bold=True)
BOLD = Font(bold=True)
THIN = Side(style="thin", color="D9D9D9")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
MONEY_FMT = '₹#,##0.00;[Red](₹#,##0.00);-'
DATE_FMT = "dd-mm-yyyy"
COUNT_FMT = "#,##0"

# ============================================================
# BANK DETECTION SIGNATURES  (PDF text fingerprints)
# ============================================================
BANK_SIGNATURES = {
    "ICICI Bank": [
        "icici bank", "icici", "tran date", "value date",
        "particulars", "withdrawals", "deposits",
    ],
    "State Bank of India (SBI)": [
        "state bank of india", "sbi", "txn date",
        "value date", "description", "ref no",
        "debit", "credit", "balance",
    ],
    "HDFC Bank": [
        "hdfc bank", "hdfc", "narration",
        "chq./ref.no", "value dt", "withdrawal amt",
        "deposit amt", "closing balance",
    ],
    "Axis Bank": [
        "axis bank", "axis", "particulars",
        "cheque no", "debit", "credit", "balance",
    ],
    "Kotak Mahindra Bank": [
        "kotak", "transaction date", "value date",
        "transaction particulars", "cheque number",
        "debit", "credit", "balance",
    ],
    "Punjab National Bank (PNB)": [
        "punjab national bank", "pnb", "txn date",
        "value date", "particulars", "debit",
        "credit", "balance",
    ],
    "Bank of Baroda (BOB)": [
        "bank of baroda", "bob", "transaction date",
        "particulars", "cheque no", "debit",
        "credit", "balance",
    ],
    "Canara Bank": [
        "canara bank", "txn date", "value date",
        "particulars", "debit", "credit", "balance",
    ],
    "Union Bank of India": [
        "union bank", "transaction date", "value date",
        "particulars", "cheque no", "debit",
        "credit", "balance",
    ],
    "IDBI Bank": [
        "idbi bank", "idbi", "transaction date",
        "value date", "particulars", "debit",
        "credit", "balance",
    ],
    "Yes Bank": [
        "yes bank", "transaction date", "value date",
        "particulars", "ref no", "debit",
        "credit", "balance",
    ],
    "IndusInd Bank": [
        "indusind", "transaction date", "value date",
        "particulars", "cheque no", "debit",
        "credit", "balance",
    ],
    "Federal Bank": [
        "federal bank", "transaction date", "value date",
        "particulars", "ref no", "debit",
        "credit", "balance",
    ],
    "RBL Bank": [
        "rbl bank", "rbl", "transaction date",
        "particulars", "ref no", "debit",
        "credit", "balance",
    ],
    "Bandhan Bank": [
        "bandhan bank", "bandhan", "transaction date",
        "particulars", "debit", "credit", "balance",
    ],
    "AU Small Finance Bank": [
        "au small finance", "au bank",
        "transaction date", "particulars",
        "debit", "credit", "balance",
    ],
}


def detect_bank_from_text(text):
    """Best-effort bank detection from first-page PDF text."""
    t = clean_text(text).lower()
    best_bank = "Unknown Bank"
    best_score = 0
    for bank, keywords in BANK_SIGNATURES.items():
        score = sum(1 for kw in keywords if kw in t)
        if score > best_score:
            best_score = score
            best_bank = bank
    return best_bank if best_score >= 2 else "Unknown Bank"


# ============================================================
# GENERIC BANK HEADER ALIASES  (expanded for all banks)
# ============================================================
ALIASES = {
    "date": [
        "date", "tran date", "transaction date", "txn date",
        "trn date", "posting date", "value date",
        "date of transaction", "txn. date", "trans date",
        "transaction dt", "tr date", "book date",
        "effective date", "entry date",
    ],
    "narration": [
        "particulars", "narration", "description", "details",
        "trn particulars", "trn. particulars",
        "transaction particulars", "transaction description",
        "transaction details", "remarks", "transaction narrative",
        "txn narration", "txn description", "details of transaction",
        "particular", "desc", "narrative",
    ],
    "debit": [
        "debit", "debit amount", "debit amt",
        "withdrawal", "withdrawals", "withdrawal amount",
        "withdrawal amt", "dr", "dr amount", "paid out",
        "debits", "dr.", "debit(rs)", "debit (rs)",
        "debit (inr)", "debit(inr)", "withdrawal(rs)",
        "withdrawals(rs)", "outflow", "money out",
    ],
    "credit": [
        "credit", "credit amount", "credit amt",
        "deposit", "deposits", "deposit amount",
        "deposit amt", "cr", "cr amount", "paid in",
        "credits", "cr.", "credit(rs)", "credit (rs)",
        "credit (inr)", "credit(inr)", "deposit(rs)",
        "deposits(rs)", "inflow", "money in",
    ],
    "balance": [
        "balance", "balance inr", "balance (inr)",
        "closing balance", "running balance",
        "available balance", "ledger balance",
        "bal", "bal (inr)", "bal(inr)",
        "closing bal", "balance(rs)", "balance (rs)",
        "end balance", "closing",
    ],
    "customer": [
        "customer", "customer name", "counterparty",
        "party name", "payer", "payer name",
        "payee", "payee name", "beneficiary",
        "beneficiary name", "remitter name",
        "sender name", "receiver name", "party",
        "name", "counterparty name",
    ],
    "utr": [
        "utr", "utr no", "utr number", "bank utr",
        "reference", "reference no", "reference number",
        "reference id", "ref id", "bank ref",
        "bank ref no", "bank reference", "rrn",
        "txn id", "transaction id", "bank transaction id",
        "ref no", "ref no.", "chq./ref.no",
        "chq/ref no", "cheque no", "cheque no.",
        "cheque number", "chq no", "chq. no",
        "ref number", "transaction ref", "txn ref",
    ],
    "amount": [
        "amount", "transaction amount", "txn amount",
        "amt", "transaction amt", "txn amt",
    ],
    "type": [
        "type", "transaction type", "txn type",
        "dr cr", "cr dr", "debit credit",
        "credit debit", "d/c", "c/d", "dr/cr",
    ],
}

# ============================================================
# UNIVERSAL DATE PATTERNS  (all Indian bank date formats)
# ============================================================
UNIVERSAL_DATE_PATTERNS = [
    # DD-MM-YYYY  (ICICI, SBI, PNB, Canara)
    (re.compile(r"\b(\d{2})-(\d{2})-(\d{4})\b"), "%d-%m-%Y"),
    # DD/MM/YYYY  (HDFC, Axis, Kotak, BOB, Union, IDBI)
    (re.compile(r"\b(\d{2})/(\d{2})/(\d{4})\b"), "%d/%m/%Y"),
    # DD.MM.YYYY  (some banks)
    (re.compile(r"\b(\d{2})\.(\d{2})\.(\d{4})\b"), "%d.%m.%Y"),
    # DD-MMM-YYYY  (SBI old, PNB old)
    (re.compile(
        r"\b(\d{2})-(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)-(\d{4})\b",
        re.I,
    ), "%d-%b-%Y"),
    # DD MMM YYYY
    (re.compile(
        r"\b(\d{2})\s+(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+(\d{4})\b",
        re.I,
    ), "%d %b %Y"),
    # DD/MMM/YYYY
    (re.compile(
        r"\b(\d{2})/(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)/(\d{4})\b",
        re.I,
    ), "%d/%b/%Y"),
    # YYYY-MM-DD  (some digital banks / ISO)
    (re.compile(r"\b(\d{4})-(\d{2})-(\d{2})\b"), "%Y-%m-%d"),
    # DD-MM-YY  (older formats)
    (re.compile(r"\b(\d{2})-(\d{2})-(\d{2})\b"), "%d-%m-%y"),
    # DD/MM/YY
    (re.compile(r"\b(\d{2})/(\d{2})/(\d{2})\b"), "%d/%m/%y"),
]

# Date-start regex for line detection (matches beginning of line)
DATE_LINE_PATTERNS = [
    re.compile(r"^(\d{2}-\d{2}-\d{4})\b"),
    re.compile(r"^(\d{2}/\d{2}/\d{4})\b"),
    re.compile(r"^(\d{2}\.\d{2}\.\d{4})\b"),
    re.compile(
        r"^(\d{2}-(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)-\d{4})\b",
        re.I,
    ),
    re.compile(
        r"^(\d{2}\s+(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+\d{4})\b",
        re.I,
    ),
    re.compile(r"^(\d{4}-\d{2}-\d{2})\b"),
    re.compile(r"^(\d{2}-\d{2}-\d{2})\b"),
    re.compile(r"^(\d{2}/\d{2}/\d{2})\b"),
]


def parse_date_flexible(text):
    """Try all known date formats and return a datetime or NaT."""
    s = clean_text(text).strip()
    if not s:
        return pd.NaT
    for pattern, fmt in UNIVERSAL_DATE_PATTERNS:
        m = pattern.search(s)
        if m:
            try:
                return pd.to_datetime(m.group(0), format=fmt)
            except Exception:
                continue
    # Last resort: let pandas guess
    try:
        return pd.to_datetime(s, dayfirst=True, errors="coerce")
    except Exception:
        return pd.NaT


def match_date_at_line_start(line):
    """Return the date string if the line starts with a date, else None."""
    for pat in DATE_LINE_PATTERNS:
        m = pat.match(line)
        if m:
            return m.group(1)
    return None


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
        .replace("\ufeff", " ")
        .replace("\n", " ")
        .replace("\r", " ")
        .strip()
        .split()
    )


def norm(value):
    return re.sub(
        r"[^a-z0-9]+",
        " ",
        clean_text(value).lower(),
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


MONEY_TOKEN_RE = re.compile(
    r"(?<![\w])"
    r"(?:\d{1,3}(?:,\d{2,3})+|\d+)"
    r"(?:\.\d{2})?"
    r"(?![\w])"
)


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
        axis=1,
    )
    return data.loc[~blank].copy()


# ============================================================
# CUSTOMER / UTR EXTRACTORS  (expanded for all banks)
# ============================================================
LOCATION_WORDS = [
    "KILAKARAI", "BANGALORE", "DOMLUR", "RPC MUMBAI",
    "RPC-NASIK", "CHENNAI RPC", "RPC JODHPUR",
    "RPC-CHH", "SAMBHAJINAGAR", "DELHI", "MUMBAI",
    "KOLKATA", "HYDERABAD", "PUNE", "AHMEDABAD",
    "JAIPUR", "LUCKNOW", "KANPUR", "NAGPUR",
    "INDORE", "THANE", "BHOPAL", "VISAKHAPATNAM",
    "PATNA", "VADODARA", "GHAZIABAD", "LUDHIANA",
    "AGRA", "NASHIK", "RANCHI", "MEERUT",
    "RAJKOT", "VARANASI", "SRINAGAR", "AURANGABAD",
    "DHANBAD", "AMRITSAR", "NAVI MUMBAI",
    "ALLAHABAD", "RANCHI", "COIMBATORE", "MADURAI",
    "TRICHY", "SALEM", "TIRUNELVELI", "ERODE",
    "VELLORE", "TUTICORIN", "THIRUVANANTHAPURAM",
    "KOCHI", "KOZHIKODE", "MANGALORE", "MYSORE",
    "HUBLI", "BELGAUM", "GULBARGA", "DAVANGERE",
    "BELLARY", "SHIMOGA", "TUMKUR", "RAICHUR",
    "BIJAPUR", "UDUPI", "HASSAN", "MANDYA",
]


def clean_customer_name(value):
    s = clean_text(value).strip(" -/")
    for loc in LOCATION_WORDS:
        s = re.sub(
            r"\b" + re.escape(loc) + r"\b.*$",
            "",
            s,
            flags=re.I,
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
        "opening balance", "closing balance",
        "brought forward", "carried forward",
        "date", "particulars", "narration",
        "description", "cheque", "reference",
        "debit", "credit", "balance",
        "withdrawal", "deposit", "amount",
        "account no", "account number",
        "branch", "ifsc", "micr",
        "customer id", "customer name",
        "statement period", "from", "to",
        "page no", "printed on",
    ]
    if any(s.lower().startswith(x) for x in reject_starts):
        return False
    if "/" in s or ":" in s:
        return False
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
        # HDFC patterns
        r"\bN\d{12,20}\b",
        r"\bIMPS\d{10,20}\b",
        # SBI patterns
        r"\bSBI[A-Z0-9]{10,25}\b",
        # Axis patterns
        r"\bAXIS[A-Z0-9]{8,25}\b",
        # Generic UTR (12-digit number)
        r"\bUTR\s*[:\-]?\s*(\d{12,16})\b",
        # Generic ref
        r"\bREF\s*[:\-]?\s*([A-Z0-9]{8,25})\b",
        # NEFT/RTGS generic
        r"\b(?:NEFT|RTGS)\s*[:\-/]\s*([A-Z0-9]{8,35})\b",
        # IMPS generic
        r"\bIMPS\s*[:\-/]\s*(\d{10,20})\b",
        # UPI ref
        r"\bUPI\s*[:\-/]\s*(\d{10,20})\b",
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
        r"(?=\s+(?:KILAKARAI|RPC|CHENNAI|BANGALORE|MUMBAI|DELHI)|$)",
        # RTGS/REF/IFSC/Name
        r"RTGS/[A-Z0-9]+/[A-Z0-9]+/"
        r"([A-Za-z][A-Za-z0-9 .&_-]{2,70}?)"
        r"(?=\s+(?:KILAKARAI|RPC|CHENNAI|BANGALORE|MUMBAI|DELHI)|$)",
        # IMPS/REF/Name/
        r"MMT/IMPS/\d+/([^/]{2,70})/",
        # BIL/INFT/ref/NA/ Name
        r"BIL/INFT/\d+/(?:NA|MIB-)/\s*"
        r"([A-Za-z][A-Za-z .&_-]{2,70}?)"
        r"(?=\s+(?:KILAKARAI|RPC|CHENNAI|BANGALORE|MUMBAI|DELHI)|$)",
        # NEFT-Nxxx-NAME
        r"NEFT-[A-Z0-9]+-"
        r"([A-Za-z][A-Za-z0-9 .&_-]{2,90}?)"
        r"(?=\s+(?:RPC|KILAKARAI|CHENNAI|BANGALORE|MUMBAI)|-SP\d|$)",
        # /NA/ NAME general
        r"/NA/\s*"
        r"([A-Za-z][A-Za-z .&_-]{2,70}?)"
        r"(?=\s+(?:KILAKARAI|RPC|CHENNAI|BANGALORE|MUMBAI|DELHI)|$)",
        # TO/ FROM name patterns (common in SBI, HDFC, Axis)
        r"\b(?:TO|FROM|A/C\s+OF|ACCT\s+OF|PAYMENT\s+TO|RECD\s+FROM)\s+"
        r"([A-Za-z][A-Za-z0-9 .&_-]{2,60}?)"
        r"(?=\s*(?:A/C|AC|ACC|IFSC|REF|UTR|NEFT|RTGS|IMPS|UPI|$))",
        # POS/ ECOM merchant name
        r"\b(?:POS|ECOM|ONLINE|SWIPE)\s+(?:AT\s+)?/?"
        r"([A-Za-z][A-Za-z0-9 .&_-]{2,50}?)"
        r"(?=\s*(?:ON|DT|DATE|REF|$))",
        # UPI-NAME@bank
        r"\bUPI[-/][A-Za-z0-9]+[-/]"
        r"([A-Za-z][A-Za-z0-9 .&_-]{2,40}?)"
        r"(?=\s*(?:@|REF|$))",
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
    out["Date"] = table[date_col].apply(parse_date_flexible)
    out["Narration"] = (
        table[narration_col] if narration_col else ""
    )
    out["Customer Name"] = (
        table[customer_col] if customer_col else ""
    )
    out["UTR / Reference"] = (
        table[utr_col] if utr_col else ""
    )
    out["Debit"] = (
        table[debit_col].apply(parse_number)
        if debit_col else np.nan
    )
    out["Credit"] = (
        table[credit_col].apply(parse_number)
        if credit_col else np.nan
    )
    out["Balance"] = (
        table[balance_col].apply(parse_number)
        if balance_col else np.nan
    )

    # Amount + DR/CR type fallback
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
                r"\bDR\b|DEBIT|WITHDRAW", regex=True, na=False
            ),
            amounts,
            np.nan,
        )
        out["Credit"] = np.where(
            txn_type.str.contains(
                r"\bCR\b|CREDIT|DEPOSIT", regex=True, na=False
            ),
            amounts,
            np.nan,
        )

    # Signed amount fallback (negative = debit)
    if (
        debit_col is None
        and credit_col is None
        and amount_col
        and type_col is None
    ):
        amounts = table[amount_col].apply(parse_number)
        out["Debit"] = np.where(amounts < 0, amounts.abs(), np.nan)
        out["Credit"] = np.where(amounts > 0, amounts, np.nan)

    # Best-effort customer extraction from narration
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
            out["Customer Name"].fillna("").astype(str).str.strip().ne(""),
            "EXTRACTED FROM NARRATION",
            "NOT AVAILABLE",
        ),
    )
    out["UTR Source"] = np.where(
        utr_col is not None,
        "SOURCE COLUMN",
        np.where(
            out["UTR / Reference"].fillna("").astype(str).str.strip().ne(""),
            "EXTRACTED FROM NARRATION",
            "NOT AVAILABLE",
        ),
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
                    engine="python",
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
                keep_default_na=False,
            )
            outputs.append((sheet, raw))
        except Exception as e:
            print(f"    WARNING sheet '{sheet}' skipped: {e}")
    return outputs


# ============================================================
# UNIVERSAL PDF TEXT PARSER  (works for ALL banks)
# ============================================================
def is_bank_statement_pdf(text):
    """Check if PDF text looks like any bank statement."""
    t = clean_text(text).lower()
    bank_keywords = [
        "bank", "statement", "account", "balance",
        "transaction", "debit", "credit", "deposit",
        "withdrawal", "particulars", "narration",
        "description", "cheque", "reference",
    ]
    score = sum(1 for kw in bank_keywords if kw in t)
    return score >= 3


def signed_balance_from_line(line):
    """
    Extract signed running balance from a transaction line.
    Handles:  1,23,456.78 Cr | 499.99 Dr | 0.00 | -1,234.56
    """
    # Strip leading dates
    stripped = line
    for pat in DATE_LINE_PATTERNS:
        stripped = pat.sub("", stripped, count=1)
    # Also strip a second date (value date)
    for pat in DATE_LINE_PATTERNS:
        stripped = pat.sub("", stripped, count=1)

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
    if abs(float(bal)) <= 0.000001:
        return 0.0
    # If no Cr/Dr suffix, return as-is (will be validated by movement)
    return float(bal)


def transaction_amount_from_line(line):
    """
    Extract transaction amount from a line.
    Usually the second-to-last money token (last is balance).
    """
    stripped = line
    for pat in DATE_LINE_PATTERNS:
        stripped = pat.sub("", stripped, count=1)
    for pat in DATE_LINE_PATTERNS:
        stripped = pat.sub("", stripped, count=1)

    nums = MONEY_TOKEN_RE.findall(stripped)
    if "B/F" in line.upper() and len(nums) <= 1:
        return None
    if len(nums) < 2:
        return None
    amount = parse_number(nums[-2])
    return None if pd.isna(amount) else abs(float(amount))


def parse_universal_pdf_text(path):
    """
    Universal PDF text parser for ALL bank statements.

    Strategy:
    1. Extract text from each page.
    2. Find lines starting with dates (all formats).
    3. Group lines between dates into transaction blocks.
    4. Extract money values from each block.
    5. Identify amount and balance using heuristics.
    6. Determine debit/credit from balance movement.
    7. Extract customer name and UTR from narration.
    8. Reconcile against printed page totals where available.
    """
    import pdfplumber

    transactions = []
    page_recon = []
    previous_signed_balance = None

    with pdfplumber.open(path) as pdf:
        for page_no, page in enumerate(pdf.pages, start=1):
            page_text = page.extract_text(
                x_tolerance=1,
                y_tolerance=3,
            ) or ""
            lines = [
                line.strip()
                for line in page_text.splitlines()
                if line.strip()
            ]

            # Printed page totals (various formats)
            total_match = re.search(
                r"(?:Total|Page\s+Total|Grand\s+Total)\s*[:\-]?\s*"
                r"([\d,]+\.\d{2})\s+"
                r"([\d,]+\.\d{2})",
                page_text,
                flags=re.I,
            )
            # Some banks print: Total Debit: xxx  Total Credit: yyy
            alt_total_match = re.search(
                r"(?:Total\s+)?(?:Debit|Withdrawal)s?\s*[:\-]?\s*"
                r"([\d,]+\.\d{2}).*?"
                r"(?:Total\s+)?(?:Credit|Deposit)s?\s*[:\-]?\s*"
                r"([\d,]+\.\d{2})",
                page_text,
                flags=re.I | re.S,
            )
            if total_match:
                printed_debit = parse_number(total_match.group(1))
                printed_credit = parse_number(total_match.group(2))
            elif alt_total_match:
                printed_debit = parse_number(alt_total_match.group(1))
                printed_credit = parse_number(alt_total_match.group(2))
            else:
                printed_debit = np.nan
                printed_credit = np.nan

            # Find all date-starting lines
            transaction_indexes = []
            for i, line in enumerate(lines):
                if match_date_at_line_start(line):
                    transaction_indexes.append(i)

            page_rows = []
            for position, line_index in enumerate(transaction_indexes):
                line = lines[line_index]
                next_index = (
                    transaction_indexes[position + 1]
                    if position + 1 < len(transaction_indexes)
                    else len(lines)
                )

                block_lines = lines[line_index:next_index]

                # Stop at footer/total lines
                cleaned_block = []
                for bl in block_lines:
                    bl_lower = bl.lower().strip()
                    if bl_lower.startswith("total"):
                        break
                    if bl_lower.startswith("page ") and len(bl) < 30:
                        break
                    if bl_lower.startswith("contd"):
                        break
                    if bl_lower.startswith("e. & o. e."):
                        break
                    cleaned_block.append(bl)
                block_lines = cleaned_block
                if not block_lines:
                    continue

                # Check for customer name line before date
                prefix_name = ""
                if line_index > 0:
                    prev = lines[line_index - 1]
                    if is_plain_customer_line(prev):
                        prefix_name = prev

                # Find balance line in block
                balance_line = None
                for bl in block_lines:
                    if re.search(r"\b(?:Cr|Dr)\b\s*$", bl, flags=re.I):
                        balance_line = bl
                        break
                    # Zero balance without Cr/Dr
                    if match_date_at_line_start(bl):
                        stripped_bl = bl
                        for pat in DATE_LINE_PATTERNS:
                            stripped_bl = pat.sub("", stripped_bl, count=1)
                        for pat in DATE_LINE_PATTERNS:
                            stripped_bl = pat.sub("", stripped_bl, count=1)
                        zero_nums = MONEY_TOKEN_RE.findall(stripped_bl)
                        if (
                            len(zero_nums) >= 2
                            and abs(parse_number(zero_nums[-1]) or 0)
                            <= 0.000001
                        ):
                            balance_line = bl
                            break
                    # Lines with enough money tokens (amount + balance)
                    nums_in_bl = MONEY_TOKEN_RE.findall(bl)
                    if len(nums_in_bl) >= 2:
                        balance_line = bl
                        break

                if balance_line is None:
                    continue

                current_balance = signed_balance_from_line(balance_line)
                if current_balance is None:
                    continue

                amount = transaction_amount_from_line(balance_line)
                date_str = match_date_at_line_start(line)
                txn_date = parse_date_flexible(date_str)
                if pd.isna(txn_date):
                    continue

                is_bf = "B/F" in line.upper() or "BROUGHT FORWARD" in line.upper()

                block_for_parse = " ".join(
                    ([prefix_name] if prefix_name else [])
                    + block_lines
                )
                customer = extract_customer_from_text(
                    block_for_parse, prefix_name
                )
                reference = extract_reference_from_text(block_for_parse)

                debit = np.nan
                credit = np.nan

                if (
                    not is_bf
                    and amount is not None
                    and previous_signed_balance is not None
                ):
                    balance_change = current_balance - previous_signed_balance
                    tolerance = max(0.02, abs(amount) * 0.000001)
                    if abs(abs(balance_change) - amount) <= tolerance:
                        if balance_change < 0:
                            debit = amount
                        elif balance_change > 0:
                            credit = amount
                    else:
                        # Fallback: classify by direction
                        if balance_change < 0:
                            debit = amount
                        elif balance_change > 0:
                            credit = amount
                elif not is_bf and amount is not None:
                    # First transaction (no previous balance)
                    # Try Cr/Dr suffix on the line
                    if re.search(r"\bDr\b", balance_line, flags=re.I):
                        debit = amount
                    elif re.search(r"\bCr\b", balance_line, flags=re.I):
                        credit = amount
                    else:
                        # Guess from balance vs amount
                        if current_balance < amount:
                            debit = amount
                        else:
                            credit = amount

                previous_signed_balance = current_balance

                if is_bf:
                    continue

                page_rows.append({
                    "Date": txn_date,
                    "Customer Name": customer,
                    "UTR / Reference": reference,
                    "Narration": clean_text(" ".join(block_lines)),
                    "Debit": debit,
                    "Credit": credit,
                    "Balance": current_balance,
                    "Customer Source": (
                        "PDF NAME LINE / NARRATION"
                        if customer else "NOT AVAILABLE"
                    ),
                    "UTR Source": (
                        "PDF NARRATION"
                        if reference else "NOT AVAILABLE"
                    ),
                    "Source Page": page_no,
                    "Parser": "UNIVERSAL PDF TEXT + BALANCE RECON",
                })

            # Page reconciliation
            if page_rows:
                page_df = pd.DataFrame(page_rows)
                parsed_debit = float(page_df["Debit"].fillna(0).sum())
                parsed_credit = float(page_df["Credit"].fillna(0).sum())
            else:
                parsed_debit = 0.0
                parsed_credit = 0.0

            debit_diff = (
                parsed_debit - printed_debit
                if pd.notna(printed_debit) else np.nan
            )
            credit_diff = (
                parsed_credit - printed_credit
                if pd.notna(printed_credit) else np.nan
            )
            if pd.notna(printed_debit) and pd.notna(printed_credit):
                page_status = (
                    "OK"
                    if abs(debit_diff) <= 0.02 and abs(credit_diff) <= 0.02
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

    return pd.DataFrame(transactions), pd.DataFrame(page_recon)


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
                width = max(len(row or []) for row in table)
                rows = []
                for row in table:
                    row = list(row or [])
                    row += [""] * (width - len(row))
                    rows.append(row)
                outputs.append((
                    f"PDF Page {page_no} Table {table_no}",
                    pd.DataFrame(rows),
                    page_no,
                ))
    return outputs


def get_first_pdf_page_text(path):
    import pdfplumber
    with pdfplumber.open(path) as pdf:
        if not pdf.pages:
            return ""
        return pdf.pages[0].extract_text(
            x_tolerance=1, y_tolerance=3
        ) or ""


# ============================================================
# OPTIONAL OCR FALLBACK
# ============================================================
def ocr_pdf_optional(path):
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
                image, config="--psm 6"
            )
        except Exception:
            continue
        lines = [
            [x] for x in text_value.splitlines() if clean_text(x)
        ]
        if lines:
            outputs.append((
                f"OCR Page {page_no}",
                pd.DataFrame(lines),
                page_no,
            ))
    return outputs


# ============================================================
# LOAD SOURCES
# ============================================================
all_frames = []
source_map = []
all_pdf_recon = []

print("=" * 82)
print(f"UNIVERSAL BANK BOOKS REPORT {VERSION}")
print(f"Author: {AUTHOR_NAME}")
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

        detected_bank = detect_bank_from_text(first_text)
        print(f"  Detected bank: {detected_bank}")

        if not is_bank_statement_pdf(first_text):
            source_map.append({
                "Source File": base,
                "Source Part": "PDF",
                "Status": "NOT A BANK STATEMENT",
                "Transaction Count": 0,
                "Note": "PDF does not appear to be a bank statement.",
            })
            print("  WARNING: PDF does not look like a bank statement.")
            continue

        # Try universal text parser first
        try:
            pdf_tx, pdf_recon = parse_universal_pdf_text(path)
        except Exception as e:
            print(f"  ERROR in universal PDF parser: {e}")
            pdf_tx = pd.DataFrame()
            pdf_recon = pd.DataFrame()

        if not pdf_tx.empty:
            pdf_tx["Source File"] = base
            pdf_tx["Source Part"] = (
                "Page " + pdf_tx["Source Page"].astype(str)
            )
            pdf_tx["Source Row"] = ""
            all_frames.append(pdf_tx)
            pdf_recon["Source File"] = base
            all_pdf_recon.append(pdf_recon)
            bad_pages = int(pdf_recon["Status"].eq("CHECK").sum())
            source_map.append({
                "Source File": base,
                "Source Part": f"UNIVERSAL PDF ({detected_bank})",
                "Status": (
                    "RECOGNIZED"
                    if bad_pages == 0
                    else "RECOGNIZED / RECON CHECK"
                ),
                "Transaction Count": len(pdf_tx),
                "Note": f"Page reconciliation issues: {bad_pages}",
            })
            print(
                f"  Universal PDF parser: "
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

        # Fallback: generic PDF table extraction
        print("  Text parser found no transactions. Trying table extraction...")
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
            table = dataframe_from_header(raw, header_row)
            tx = standardize_generic_table(table)
            if tx.empty:
                continue
            tx["Source File"] = base
            tx["Source Part"] = source_name
            tx["Source Row"] = (
                tx.index.to_series().astype(int)
                + header_row + 2
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
                "Note": f"Bank: {detected_bank}",
            })
            print(f"  {source_name}: {len(tx):,} transaction(s)")

        if not recognized_pdf:
            source_map.append({
                "Source File": base,
                "Source Part": "PDF",
                "Status": "UNRECOGNIZED",
                "Transaction Count": 0,
                "Note": (
                    f"Bank detected: {detected_bank}. "
                    "No reliable transaction table recognized. "
                    "If image-only PDF, install OCR dependencies "
                    "or export as searchable PDF/Excel."
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
        table = dataframe_from_header(raw, header_row)
        tx = standardize_generic_table(table)
        if tx.empty:
            continue
        tx["Source File"] = base
        tx["Source Part"] = source_name
        tx["Source Row"] = (
            tx.index.to_series().astype(int)
            + header_row + 2
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
        print(f"  {source_name}: {len(tx):,} transaction(s)")

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
        "For PDFs, searchable/text bank statements work best.\n"
        "Supported banks: SBI, HDFC, ICICI, Axis, Kotak, PNB, "
        "BOB, Canara, Union, IDBI, Yes, IndusInd, Federal, RBL, "
        "and more.\n"
    )

# ============================================================
# CONSOLIDATE
# ============================================================
transactions = pd.concat(all_frames, ignore_index=True, sort=False)

for required_col in [
    "Date", "Customer Name", "UTR / Reference", "Narration",
    "Debit", "Credit", "Balance", "Customer Source",
    "UTR Source", "Source File", "Source Part",
    "Source Row", "Source Page", "Parser",
]:
    if required_col not in transactions.columns:
        transactions[required_col] = ""

transactions["Date"] = pd.to_datetime(
    transactions["Date"], errors="coerce"
).dt.normalize()
transactions["Debit"] = pd.to_numeric(
    transactions["Debit"], errors="coerce"
)
transactions["Credit"] = pd.to_numeric(
    transactions["Credit"], errors="coerce"
)
transactions["Balance"] = pd.to_numeric(
    transactions["Balance"], errors="coerce"
)
transactions["Direction"] = np.where(
    transactions["Debit"].notna(),
    "DEBIT",
    np.where(transactions["Credit"].notna(), "CREDIT", ""),
)
transactions["Transaction Amount"] = np.where(
    transactions["Direction"].eq("DEBIT"),
    transactions["Debit"],
    transactions["Credit"],
)

# Duplicate diagnostic
dup_key = (
    transactions["Date"].astype(str)
    + "|"
    + transactions["Debit"].fillna(0).astype(str)
    + "|"
    + transactions["Credit"].fillna(0).astype(str)
    + "|"
    + transactions["UTR / Reference"]
    .fillna("").astype(str).str.strip()
    + "|"
    + transactions["Narration"]
    .fillna("").astype(str).str.strip()
)
transactions["Possible Duplicate"] = np.where(
    dup_key.duplicated(keep=False), "YES", "NO"
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
    transactions["Date"].dt.to_period("M").astype(str)
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
            lambda s: s.replace("", np.nan).dropna().nunique(),
        ),
        Unique_UTRs=(
            "UTR / Reference",
            lambda s: s.replace("", np.nan).dropna().nunique(),
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
    .fillna("").astype(str).str.strip().ne("")
].copy()

if customer_tx.empty:
    customer_wise = pd.DataFrame(columns=[
        "Customer Name", "Transaction_Count", "Debit_Count",
        "Debit_Total", "Credit_Count", "Credit_Total",
        "Net_Credit_Minus_Debit", "First_Date", "Last_Date",
        "Unique_UTRs",
    ])
    customer_date_wise = pd.DataFrame(columns=[
        "Date", "Customer Name", "Transaction_Count",
        "Debit_Count", "Debit_Total", "Credit_Count",
        "Credit_Total", "Net_Credit_Minus_Debit",
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
                lambda s: s.replace("", np.nan).dropna().nunique(),
            ),
        )
        .reset_index()
    )
    customer_wise["Net_Credit_Minus_Debit"] = (
        customer_wise["Credit_Total"].fillna(0)
        - customer_wise["Debit_Total"].fillna(0)
    )
    customer_wise = customer_wise[[
        "Customer Name", "Transaction_Count", "Debit_Count",
        "Debit_Total", "Credit_Count", "Credit_Total",
        "Net_Credit_Minus_Debit", "First_Date", "Last_Date",
        "Unique_UTRs",
    ]]
    customer_date_wise = (
        customer_tx
        .groupby(["Date", "Customer Name"], dropna=False)
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
    "Date", "Customer Name", "UTR / Reference", "Narration",
    "Debit", "Credit", "Balance", "Source File", "Source Part",
]].copy()

# ============================================================
# CONTROL TOTALS
# ============================================================
total_transactions = len(transactions)
debit_count = int(transactions["Debit"].notna().sum())
credit_count = int(transactions["Credit"].notna().sum())
total_debit = float(transactions["Debit"].fillna(0).sum())
total_credit = float(transactions["Credit"].fillna(0).sum())
net_movement = total_credit - total_debit
valid_dates = transactions["Date"].dropna()
first_date = valid_dates.min() if not valid_dates.empty else None
last_date = valid_dates.max() if not valid_dates.empty else None
unique_customers = int(
    transactions["Customer Name"]
    .replace("", np.nan).dropna().nunique()
)
unique_utrs = int(
    transactions["UTR / Reference"]
    .replace("", np.nan).dropna().nunique()
)
missing_customer = int(
    transactions["Customer Name"]
    .fillna("").astype(str).str.strip().eq("").sum()
)
missing_utr = int(
    transactions["UTR / Reference"]
    .fillna("").astype(str).str.strip().eq("").sum()
)
possible_duplicates = int(
    transactions["Possible Duplicate"].eq("YES").sum()
)

# ============================================================
# BUILD PDF RECON TABLE
# ============================================================
if all_pdf_recon:
    pdf_page_recon = pd.concat(
        all_pdf_recon, ignore_index=True, sort=False
    )
else:
    pdf_page_recon = pd.DataFrame(columns=[
        "Source File", "Page", "Parsed Debit", "Printed Debit",
        "Debit Difference", "Parsed Credit", "Printed Credit",
        "Credit Difference", "Status",
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
            wrap_text=True,
        )


def auto_width(ws, max_width=42):
    for col_idx in range(1, ws.max_column + 1):
        max_len = 0
        for row_idx in range(1, min(ws.max_row, 3000) + 1):
            value = ws.cell(row_idx, col_idx).value
            if value is not None:
                max_len = max(max_len, len(str(value)))
        ws.column_dimensions[
            get_column_letter(col_idx)
        ].width = min(max(max_len + 2, 11), max_width)


def write_df(ws, frame, money_cols=(), date_cols=(), count_cols=()):
    if frame is None or frame.empty:
        ws["A1"] = "No data available"
        return
    for col_idx, col_name in enumerate(frame.columns, start=1):
        ws.cell(1, col_idx, col_name)
    for row_idx, row in enumerate(
        frame.itertuples(index=False, name=None), start=2
    ):
        for col_idx, value in enumerate(row, start=1):
            cell = ws.cell(
                row_idx, col_idx, safe_excel_value(value)
            )
            cell.border = BORDER
            cell.alignment = Alignment(vertical="center")
    style_header(ws)
    for col_idx, col_name in enumerate(frame.columns, start=1):
        if col_name in money_cols:
            for r in range(2, ws.max_row + 1):
                ws.cell(r, col_idx).number_format = MONEY_FMT
        elif col_name in date_cols:
            for r in range(2, ws.max_row + 1):
                ws.cell(r, col_idx).number_format = DATE_FMT
        elif col_name in count_cols:
            for r in range(2, ws.max_row + 1):
                ws.cell(r, col_idx).number_format = COUNT_FMT
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    auto_width(ws)


# ============================================================
# COVER_REPORT
# ============================================================
ws = wb.create_sheet("COVER_REPORT")
ws.sheet_view.showGridLines = False

# Title
ws.merge_cells("A1:D1")
ws["A1"] = f"UNIVERSAL BANK BOOKS OF ACCOUNTS REPORT {VERSION}"
ws["A1"].font = Font(bold=True, size=16, color="1F4E78")
ws["A1"].alignment = Alignment(horizontal="center")

# Author block
ws.merge_cells("A2:D2")
ws["A2"] = (
    f"Author: {AUTHOR_NAME}  |  "
    f"{AUTHOR_COPYRIGHT}  |  "
    f"Instagram: {AUTHOR_INSTAGRAM}"
)
ws["A2"].font = Font(size=10, color="666666", italic=True)
ws["A2"].alignment = Alignment(horizontal="center")
ws["A2"].fill = AUTHOR_FILL

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
        int(pdf_page_recon["Status"].eq("CHECK").sum())
        if not pdf_page_recon.empty else 0,
    ),
    (
        "Recognized Source Files",
        transactions["Source File"].nunique(),
    ),
]

for r, (label, value) in enumerate(cover_rows, start=4):
    ws.cell(r, 1, label)
    ws.cell(r, 2, safe_excel_value(value))
    ws.cell(r, 1).border = BORDER
    ws.cell(r, 2).border = BORDER
    if label in ("Total Debit", "Total Credit", "Net Credit - Debit"):
        ws.cell(r, 2).number_format = MONEY_FMT
    if "Date" in label:
        ws.cell(r, 2).number_format = DATE_FMT

ws.column_dimensions["A"].width = 44
ws.column_dimensions["B"].width = 24

# ============================================================
# TRANSACTION_REGISTER
# ============================================================
register_columns = [
    "Date", "Customer Name", "UTR / Reference", "Narration",
    "Debit", "Credit", "Balance", "Direction",
    "Transaction Amount", "Customer Source", "UTR Source",
    "Possible Duplicate", "Parser", "Source File",
    "Source Part", "Source Page", "Source Row",
]
ws = wb.create_sheet("TRANSACTION_REGISTER")
write_df(
    ws,
    transactions[register_columns],
    money_cols=("Debit", "Credit", "Balance", "Transaction Amount"),
    date_cols=("Date",),
)

# ============================================================
# DATE_WISE_SUMMARY
# ============================================================
ws = wb.create_sheet("DATE_WISE_SUMMARY")
write_df(
    ws, date_wise,
    money_cols=("Debit_Total", "Credit_Total", "Net_Credit_Minus_Debit"),
    date_cols=("Date",),
    count_cols=("Transaction_Count", "Debit_Count", "Credit_Count"),
)

# ============================================================
# MONTH_WISE_SUMMARY
# ============================================================
ws = wb.create_sheet("MONTH_WISE_SUMMARY")
write_df(
    ws, month_wise,
    money_cols=("Debit_Total", "Credit_Total", "Net_Credit_Minus_Debit"),
    count_cols=(
        "Transaction_Count", "Debit_Count", "Credit_Count",
        "Unique_Customers", "Unique_UTRs",
    ),
)

# ============================================================
# CUSTOMER_WISE_SUMMARY
# ============================================================
ws = wb.create_sheet("CUSTOMER_WISE_SUMMARY")
write_df(
    ws, customer_wise,
    money_cols=("Debit_Total", "Credit_Total", "Net_Credit_Minus_Debit"),
    date_cols=("First_Date", "Last_Date"),
    count_cols=(
        "Transaction_Count", "Debit_Count", "Credit_Count",
        "Unique_UTRs",
    ),
)

# ============================================================
# CUSTOMER_DATE_WISE
# ============================================================
ws = wb.create_sheet("CUSTOMER_DATE_WISE")
write_df(
    ws, customer_date_wise,
    money_cols=("Debit_Total", "Credit_Total", "Net_Credit_Minus_Debit"),
    date_cols=("Date",),
    count_cols=("Transaction_Count", "Debit_Count", "Credit_Count"),
)

# ============================================================
# UTR_WISE_REGISTER
# ============================================================
ws = wb.create_sheet("UTR_WISE_REGISTER")
write_df(
    ws, utr_register,
    money_cols=("Debit", "Credit", "Balance"),
    date_cols=("Date",),
)

# ============================================================
# PDF_PAGE_RECON
# ============================================================
ws = wb.create_sheet("PDF_PAGE_RECON")
write_df(
    ws, pdf_page_recon,
    money_cols=(
        "Parsed Debit", "Printed Debit", "Debit Difference",
        "Parsed Credit", "Printed Credit", "Credit Difference",
    ),
    count_cols=("Page",),
)
if not pdf_page_recon.empty:
    status_col = list(pdf_page_recon.columns).index("Status") + 1
    for r in range(2, ws.max_row + 1):
        if ws.cell(r, status_col).value == "CHECK":
            for c in range(1, ws.max_column + 1):
                ws.cell(r, c).fill = WARN_FILL

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
        int(pdf_page_recon["Status"].eq("CHECK").sum())
        if not pdf_page_recon.empty else 0,
    ],
]
for r, row in enumerate(controls, start=1):
    for c, value in enumerate(row, start=1):
        ws.cell(r, c, safe_excel_value(value))
        ws.cell(r, c).border = BORDER
style_header(ws)
for r in (4, 6, 7):
    ws.cell(r, 2).number_format = MONEY_FMT
ws.column_dimensions["A"].width = 44
ws.column_dimensions["B"].width = 24

# ============================================================
# SOURCE_MAPPING
# ============================================================
ws = wb.create_sheet("SOURCE_MAPPING")
write_df(
    ws, pd.DataFrame(source_map),
    count_cols=("Transaction Count",),
)

# ============================================================
# REPORT_NOTES
# ============================================================
ws = wb.create_sheet("REPORT_NOTES")
notes = [
    ["Item", "Note"],
    [
        "Author",
        f"{AUTHOR_NAME}  |  {AUTHOR_COPYRIGHT}  |  "
        f"Instagram: {AUTHOR_INSTAGRAM}",
    ],
    [
        "Purpose",
        "Bank statement to supporting Books-of-Accounts "
        "transaction report.",
    ],
    [
        "V5 Universal Bank Support",
        "Supports ALL major Indian banks: SBI, HDFC, ICICI, "
        "Axis, Kotak, PNB, BOB, Canara, Union, IDBI, Yes, "
        "IndusInd, Federal, RBL, Bandhan, AU Small Finance, "
        "and any bank producing standard tabular statements.",
    ],
    [
        "V5 PDF Parser",
        "Universal PDF text parser handles all date formats "
        "(DD-MM-YYYY, DD/MM/YYYY, DD-MMM-YYYY, YYYY-MM-DD, "
        "etc.) and uses running balance movement to determine "
        "Debit/Credit for any bank layout.",
    ],
    [
        "Debit / Credit",
        "Direction is derived from actual running balance "
        "movement and transaction amount. Falls back to "
        "Cr/Dr suffix detection when balance history is "
        "unavailable.",
    ],
    [
        "PDF Page Reconciliation",
        "Where the PDF prints Total Withdrawals and Total "
        "Deposits, parsed page totals are checked against "
        "those printed figures.",
    ],
    [
        "Customer Name",
        "Uses source name field when available; otherwise "
        "uses best-effort structured narration/name-line "
        "extraction (NEFT, RTGS, IMPS, UPI, POS, TO/FROM "
        "patterns).",
    ],
    [
        "UTR / Reference",
        "Uses source reference field when available; "
        "otherwise extracts known UPI/IMPS/NEFT/RTGS/INFT "
        "references from narration.",
    ],
    [
        "Missing Values",
        "Missing customer/UTR values remain unavailable; "
        "no value is invented.",
    ],
    [
        "Duplicates",
        "Possible duplicates are only flagged and are never "
        "automatically removed.",
    ],
    [
        "Supported Files",
        "XLSX, XLS, CSV and PDF (text-based preferred).",
    ],
    [
        "Generated On",
        datetime.now().strftime("%d-%m-%Y %H:%M:%S"),
    ],
]
for r, row in enumerate(notes, start=1):
    for c, value in enumerate(row, start=1):
        ws.cell(r, c, value)
        ws.cell(r, c).border = BORDER
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
        int(pdf_page_recon["Status"].eq("CHECK").sum()),
    )
print("")
print(f"Author       : {AUTHOR_NAME}")
print(f"Copyright    : {AUTHOR_COPYRIGHT}")
print(f"Instagram    : {AUTHOR_INSTAGRAM}")
print("")
print("Saving:", OUTPUT_FILE)
wb.save(OUTPUT_FILE)
print("SUCCESS")
print("Output:", OUTPUT_FILE)
print("=" * 82)
