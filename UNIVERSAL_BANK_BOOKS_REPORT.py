"""
============================================================
UNIVERSAL BANK BOOKS REPORT V7.1
============================================================
Author   : Mohamed Nayeem
Copyright: © Mohamed Nayeem — All Rights Reserved
Instagram: @mohamednayeem7
============================================================

V7.1 — INDIA-WIDE BANK FORMAT COVERAGE + DIAGNOSTICS + AUDIT CONTROLS
----------------------------
Supports ALL major Indian bank statement formats:
  SBI, HDFC, ICICI, Axis, Kotak Mahindra, PNB, BOB,
  Canara, Union Bank, IDBI, Yes Bank, IndusInd, Federal,
  RBL, Bandhan, AU Small Finance, Ujjivan, Equitas,
  Jana, Suryoday, Utkarsh, and any bank that produces
  a standard tabular statement.

File types:  .xlsx  .xls  .csv  .pdf  .html  .htm  .txt

Output:  UNIVERSAL_BANK_BOOKS_REPORT_V7_1.xlsx

Sheets:
  COVER_REPORT
  ACCOUNT_PROFILE
  TRANSACTION_REGISTER
  REVIEW_REQUIRED
  DATE_WISE_SUMMARY
  MONTH_WISE_SUMMARY
  ACCOUNT_WISE_SUMMARY
  CUSTOMER_WISE_SUMMARY
  CUSTOMER_DATE_WISE
  TRANSACTION_TYPE_SUMMARY
  BANK_CHARGES_GST
  UTR_WISE_REGISTER
  UTR_DUPLICATE_CHECK
  BANK_RECONCILIATION
  PDF_PAGE_RECON
  DATA_QUALITY
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
  python -m pip install pandas openpyxl xlrd pdfplumber pyyaml lxml
Optional OCR:
  python -m pip install pytesseract pdf2image pillow

RUN
  python UNIVERSAL_BANK_BOOKS_REPORT_V7.py
"""

import os
import sys
import glob
import re
import yaml
import logging
import traceback
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
VERSION = "V7_1"
DISPLAY_VERSION = "V7.1"

# ============================================================
# CONFIG
# ============================================================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_FILE = os.path.join(
    BASE_DIR,
    "UNIVERSAL_BANK_BOOKS_REPORT_V7_1.xlsx",
)
PATTERNS = (
    "*.xlsx", "*.XLSX",
    "*.xls",  "*.XLS",
    "*.csv",  "*.CSV",
    "*.pdf",  "*.PDF",
    "*.html", "*.HTML", "*.htm", "*.HTM",
    "*.txt", "*.TXT",
)
input_files = []
for pattern in PATTERNS:
    input_files.extend(glob.glob(os.path.join(BASE_DIR, pattern)))
input_files = sorted({
    p for p in input_files
    if os.path.abspath(p).lower() != os.path.abspath(OUTPUT_FILE).lower()
    and not os.path.basename(p).startswith("~$")
    and not re.match(r"(?i)^UNIVERSAL_BANK_BOOKS_REPORT_V[0-9_.-]+\.xlsx$", os.path.basename(p))
    and not re.match(r"(?i)^audit_log_\d{8}_\d{6}\.txt$", os.path.basename(p))
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

# V7.1 — extended India-wide bank fingerprints. Generic parsing still works
# for an unknown/co-operative bank when standard statement fields are present.
BANK_SIGNATURES.update({
    "IDFC FIRST Bank": ["idfc first", "idfcfirst", "transaction date", "balance"],
    "Indian Bank": ["indian bank", "transaction date", "particulars", "balance"],
    "Indian Overseas Bank": ["indian overseas bank", "iob", "transaction date", "balance"],
    "Bank of India": ["bank of india", "transaction date", "particulars", "balance"],
    "Central Bank of India": ["central bank of india", "transaction date", "balance"],
    "UCO Bank": ["uco bank", "transaction date", "balance"],
    "Bank of Maharashtra": ["bank of maharashtra", "transaction date", "balance"],
    "Punjab & Sind Bank": ["punjab & sind bank", "punjab and sind bank", "transaction date"],
    "Karur Vysya Bank": ["karur vysya bank", "kvb", "transaction date", "balance"],
    "City Union Bank": ["city union bank", "cub", "transaction date", "balance"],
    "Tamilnad Mercantile Bank": ["tamilnad mercantile bank", "tmb", "transaction date", "balance"],
    "South Indian Bank": ["south indian bank", "sib", "transaction date", "balance"],
    "Karnataka Bank": ["karnataka bank", "transaction date", "balance"],
    "DCB Bank": ["dcb bank", "transaction date", "balance"],
    "CSB Bank": ["csb bank", "catholic syrian bank", "transaction date"],
    "Dhanlaxmi Bank": ["dhanlaxmi bank", "transaction date", "balance"],
    "Jammu & Kashmir Bank": ["jammu & kashmir bank", "j&k bank", "jk bank"],
    "Nainital Bank": ["nainital bank", "transaction date", "balance"],
    "Jana Small Finance Bank": ["jana small finance bank", "jana", "transaction date"],
    "Suryoday Small Finance Bank": ["suryoday small finance bank", "suryoday", "transaction date"],
    "Utkarsh Small Finance Bank": ["utkarsh small finance bank", "utkarsh", "transaction date"],
    "Ujjivan Small Finance Bank": ["ujjivan small finance bank", "ujjivan", "transaction date"],
    "Equitas Small Finance Bank": ["equitas small finance bank", "equitas", "transaction date"],
    "ESAF Small Finance Bank": ["esaf small finance bank", "esaf", "transaction date"],
    "Capital Small Finance Bank": ["capital small finance bank", "capital sfb", "transaction date"],
    "Unity Small Finance Bank": ["unity small finance bank", "unity sfb", "transaction date"],
    "Shivalik Small Finance Bank": ["shivalik small finance bank", "shivalik", "transaction date"],
    "Airtel Payments Bank": ["airtel payments bank", "airtel bank", "transaction date"],
    "India Post Payments Bank": ["india post payments bank", "ippb", "transaction date"],
    "Fino Payments Bank": ["fino payments bank", "fino bank", "transaction date"],
    "Jio Payments Bank": ["jio payments bank", "jio bank", "transaction date"],
    "NSDL Payments Bank": ["nsdl payments bank", "nsdl bank", "transaction date"],
})


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


# V7.1 — extra header variants seen across Indian retail/current-account exports.
_EXTRA_ALIASES = {
    "date": ["txn dt", "transaction dt.", "posting dt", "date & time", "transaction date & time"],
    "narration": ["transaction remarks", "transaction narration", "description/narration", "particulars / narration", "txn particulars / narration"],
    "debit": ["withdrawal amt.", "withdrawal amount(inr)", "debit amount(inr)", "debit amt(inr)", "dr amt", "dr amt."],
    "credit": ["deposit amt.", "deposit amount(inr)", "credit amount(inr)", "credit amt(inr)", "cr amt", "cr amt."],
    "balance": ["closing balance(inr)", "available bal", "ledger bal", "running bal", "balance amount", "balance amt"],
    "customer": ["remitter", "beneficiary / remitter", "counter party", "counter-party", "transaction party"],
    "utr": ["utr/ref no", "utr / ref no", "reference/cheque no", "chq/ref", "cheque/reference", "transaction reference no", "rrn/utr"],
    "amount": ["txn amt.", "transaction value", "value amount", "debit/credit amount", "cr/dr amount"],
    "type": ["dr/cr indicator", "debit/credit indicator", "txn nature", "nature", "transaction nature", "cr/dr"],
}
for _field, _values in _EXTRA_ALIASES.items():
    _seen = {re.sub(r"[^a-z0-9]+", " ", str(v).lower()).strip() for v in ALIASES.setdefault(_field, [])}
    for _value in _values:
        _key = re.sub(r"[^a-z0-9]+", " ", str(_value).lower()).strip()
        if _key not in _seen:
            ALIASES[_field].append(_value)
            _seen.add(_key)

# ============================================================
# V7 CONFIGURATION — VALIDATED AND MERGED WITH BUILT-IN ALIASES
# ============================================================
# Important: config aliases EXTEND the proven V6/V7 built-in aliases.
# They never replace them, so a short custom YAML cannot reduce bank support.
CONFIG_CANDIDATES = [
    os.path.join(BASE_DIR, "config.yaml"),
    os.path.join(BASE_DIR, "UNIVERSAL_BANK_BOOKS_CONFIG_V7.yaml"),
]
CONFIG_FILE = next((p for p in CONFIG_CANDIDATES if os.path.exists(p)), CONFIG_CANDIDATES[0])

DEFAULT_CONFIG = {
    "column_aliases": {},
    "rounding_precision": 2,
    "max_rows_per_sheet": 1000000,
    "preserve_raw_data": True,
    "review_confidence_ok": 85,
    "review_confidence_min": 65,
    # PG/GST validation is deliberately disabled for ordinary bank statements.
    # A bank transaction alone does not prove that a PG charge applies.
    "pg_validation_enabled": False,
    "pg_charge_rate": 0.008,
    "gst_rate": 0.18,
    "success_statuses": ["SUCCESS", "COMPLETED", "CAPTURED", "SETTLED"],
    "failed_statuses": ["FAILED", "DECLINED", "REJECTED", "ERROR"],
    "pending_statuses": ["PENDING", "INPROCESS", "AUTHORIZED"],
    "diagnostic_mode": True,
    "ocr_enabled": True,
    "allow_leading_serial_before_date": True,
    "unknown_bank_generic_parser": True,
}


def validate_config(cfg):
    if cfg is None:
        cfg = {}
    if not isinstance(cfg, dict):
        raise ValueError("config.yaml root must be a mapping/object.")

    merged = dict(DEFAULT_CONFIG)
    merged.update(cfg)

    aliases_cfg = merged.get("column_aliases", {})
    if aliases_cfg is None:
        aliases_cfg = {}
    if not isinstance(aliases_cfg, dict):
        raise ValueError("column_aliases must be a mapping of field -> list of aliases.")
    for field, values in aliases_cfg.items():
        if not isinstance(field, str):
            raise ValueError("Every column_aliases key must be text.")
        if not isinstance(values, list) or not all(isinstance(v, str) for v in values):
            raise ValueError(f"column_aliases.{field} must be a list of strings.")

    try:
        merged["rounding_precision"] = int(merged.get("rounding_precision", 2))
    except Exception:
        raise ValueError("rounding_precision must be an integer.")
    if not 0 <= merged["rounding_precision"] <= 8:
        raise ValueError("rounding_precision must be between 0 and 8.")

    try:
        merged["max_rows_per_sheet"] = int(merged.get("max_rows_per_sheet", 1000000))
    except Exception:
        raise ValueError("max_rows_per_sheet must be an integer.")
    # Excel supports 1,048,576 rows including header.
    if not 1000 <= merged["max_rows_per_sheet"] <= 1048575:
        raise ValueError("max_rows_per_sheet must be between 1,000 and 1,048,575 data rows.")

    for key in ("review_confidence_ok", "review_confidence_min"):
        try:
            merged[key] = int(merged.get(key, DEFAULT_CONFIG[key]))
        except Exception:
            raise ValueError(f"{key} must be an integer.")
        if not 0 <= merged[key] <= 100:
            raise ValueError(f"{key} must be between 0 and 100.")
    if merged["review_confidence_min"] > merged["review_confidence_ok"]:
        raise ValueError("review_confidence_min cannot be greater than review_confidence_ok.")

    for key in ("pg_charge_rate", "gst_rate"):
        try:
            merged[key] = float(merged.get(key, DEFAULT_CONFIG[key]))
        except Exception:
            raise ValueError(f"{key} must be numeric.")
        if merged[key] < 0 or merged[key] > 1:
            raise ValueError(f"{key} must be between 0 and 1 (for example 0.008 = 0.8%).")

    merged["preserve_raw_data"] = bool(merged.get("preserve_raw_data", True))
    merged["pg_validation_enabled"] = bool(merged.get("pg_validation_enabled", False))
    merged["diagnostic_mode"] = bool(merged.get("diagnostic_mode", True))
    merged["ocr_enabled"] = bool(merged.get("ocr_enabled", True))
    merged["allow_leading_serial_before_date"] = bool(merged.get("allow_leading_serial_before_date", True))
    merged["unknown_bank_generic_parser"] = bool(merged.get("unknown_bank_generic_parser", True))

    for key in ("success_statuses", "failed_statuses", "pending_statuses"):
        values = merged.get(key, DEFAULT_CONFIG[key])
        if not isinstance(values, list):
            raise ValueError(f"{key} must be a list.")
        merged[key] = [str(v).strip().upper() for v in values if str(v).strip()]

    merged["column_aliases"] = aliases_cfg
    return merged


if not os.path.exists(CONFIG_FILE):
    # Create a safe default and continue. No forced exit is needed.
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        yaml.safe_dump(DEFAULT_CONFIG, f, sort_keys=False, allow_unicode=True)

try:
    with open(CONFIG_FILE, "r", encoding="utf-8") as f:
        CONFIG = validate_config(yaml.safe_load(f))
except Exception as exc:
    raise ValueError(f"Invalid YAML configuration: {CONFIG_FILE}\\n{exc}") from exc

# Merge custom aliases into built-in aliases without deleting built-in support.
for field, custom_values in CONFIG.get("column_aliases", {}).items():
    existing = ALIASES.setdefault(field, [])
    seen = {str(v).strip().lower() for v in existing}
    for value in custom_values:
        key = str(value).strip().lower()
        if key and key not in seen:
            existing.append(value)
            seen.add(key)

ROUND_DECIMALS = CONFIG["rounding_precision"]
MAX_ROWS_PER_SHEET = CONFIG["max_rows_per_sheet"]
PRESERVE_RAW_DATA = CONFIG["preserve_raw_data"]
REVIEW_OK_THRESHOLD = CONFIG["review_confidence_ok"]
REVIEW_MIN_THRESHOLD = CONFIG["review_confidence_min"]
PG_VALIDATION_ENABLED = CONFIG["pg_validation_enabled"]
PG_CHARGE_RATE = CONFIG["pg_charge_rate"]
GST_RATE = CONFIG["gst_rate"]
DIAGNOSTIC_MODE = CONFIG["diagnostic_mode"]
OCR_ENABLED = CONFIG["ocr_enabled"]
ALLOW_LEADING_SERIAL_BEFORE_DATE = CONFIG["allow_leading_serial_before_date"]
UNKNOWN_BANK_GENERIC_PARSER = CONFIG["unknown_bank_generic_parser"]

# ============================================================
# V7 LOGGING / AUDIT TRAIL
# ============================================================
RUN_ID = datetime.now().strftime("%Y%m%d_%H%M%S")
LOG_FILE = os.path.join(BASE_DIR, f"audit_log_{RUN_ID}.txt")
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    handlers=[
        logging.FileHandler(LOG_FILE, encoding="utf-8"),
        logging.StreamHandler(sys.stdout),
    ],
)
logger = logging.getLogger("universal_bank_books_v7")
logger.info("Run ID: %s", RUN_ID)
logger.info("Config: %s", CONFIG_FILE)
logger.info("Output: %s", OUTPUT_FILE)

exceptions_list = []
raw_data_sheets = {}
raw_data_index_records = []
row_control_seed = []
parser_recon_records = []
format_diagnostic_records = []


def log_exception(category, description, source_file="", source_part="", source_row=""):
    record = {
        "Exception Type": str(category),
        "Description": str(description),
        "Source File": str(source_file),
        "Source Part": str(source_part),
        "Source Row": source_row,
        "Timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }
    exceptions_list.append(record)
    logger.warning("%s: %s | %s | %s", category, description, source_file, source_part)


def sanitize_sheet_name(name, used=None):
    value = re.sub(r"[\\/*?:\\[\\]]", "_", str(name or "Sheet"))
    value = value.strip(" '") or "Sheet"
    value = value[:31]
    if used is None:
        return value
    candidate = value
    n = 2
    while candidate.lower() in used:
        suffix = f"_{n}"
        candidate = value[:31-len(suffix)] + suffix
        n += 1
    used.add(candidate.lower())
    return candidate


def store_raw_data(df, source_file, source_part):
    if not PRESERVE_RAW_DATA or df is None or df.empty:
        return
    chunk_size = min(MAX_ROWS_PER_SHEET, 1048575)
    total = len(df)
    for start in range(0, total, chunk_size):
        chunk = df.iloc[start:start + chunk_size].copy()
        sheet_name = f"RAW_{len(raw_data_sheets) + 1:03d}"
        raw_data_sheets[sheet_name] = chunk
        raw_data_index_records.append({
            "Raw Sheet": sheet_name,
            "Source File": source_file,
            "Source Part": source_part,
            "Start Record": start + 1,
            "End Record": min(start + chunk_size, total),
            "Rows": len(chunk),
        })

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
    # DD.MM.YY
    (re.compile(r"\b(\d{2})\.(\d{2})\.(\d{2})\b"), "%d.%m.%y"),
    # DD-MMM-YY / DD MMM YY / DD/MMM/YY
    (re.compile(r"\b(\d{1,2})-(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)-(\d{2})\b", re.I), "%d-%b-%y"),
    (re.compile(r"\b(\d{1,2})\s+(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+(\d{2})\b", re.I), "%d %b %y"),
    (re.compile(r"\b(\d{1,2})/(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)/(\d{2})\b", re.I), "%d/%b/%y"),
    # YYYY/MM/DD
    (re.compile(r"\b(\d{4})/(\d{2})/(\d{2})\b"), "%Y/%m/%d"),
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
    """Return transaction date at line start; V7.1 tolerates a leading serial number."""
    text = str(line or "").strip()
    for pat in DATE_LINE_PATTERNS:
        m = pat.match(text)
        if m:
            return m.group(1)
    if ALLOW_LEADING_SERIAL_BEFORE_DATE:
        candidate = re.sub(r"^\s*\d{1,7}\s*(?:[|.:;-]\s*)?", "", text, count=1)
        for pat in DATE_LINE_PATTERNS:
            m = pat.match(candidate)
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
            return round(float(value), ROUND_DECIMALS)
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
        return round(-x if negative else x, ROUND_DECIMALS)
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
        # Avoid false matches such as generic "Amount" being treated as
        # both "Debit Amount" and "Credit Amount". Prefer alias->header
        # containment; allow reverse containment only for descriptive headers.
        if len(a) >= 4 and a in v:
            return True
        if len(v) >= 8 and len(v.split()) >= 2 and v in a:
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
    # V7.1: many co-operative/digital bank exports keep DR/CR inside the
    # Amount cells and therefore have no separate Type column.
    if "date" in hits and "amount" in hits and ("balance" in hits or "narration" in hits):
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

    # Amount-column DR/CR marker fallback (e.g. "1,250.00 DR").
    if (
        debit_col is None
        and credit_col is None
        and amount_col
        and type_col is None
    ):
        raw_amount = table[amount_col].astype(str)
        amounts = table[amount_col].apply(parse_number).abs()
        dr_marker = raw_amount.str.contains(r"\b(?:DR|DEBIT)\b", case=False, regex=True, na=False)
        cr_marker = raw_amount.str.contains(r"\b(?:CR|CREDIT)\b", case=False, regex=True, na=False)
        if bool((dr_marker | cr_marker).any()):
            out["Debit"] = np.where(dr_marker, amounts, np.nan)
            out["Credit"] = np.where(cr_marker, amounts, np.nan)
        else:
            signed_amounts = table[amount_col].apply(parse_number)
            out["Debit"] = np.where(signed_amounts < 0, signed_amounts.abs(), np.nan)
            out["Credit"] = np.where(signed_amounts > 0, signed_amounts, np.nan)

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
def _read_delimited_robust(path):
    """Read CSV/TXT exports with delimiter/preamble/encoding tolerance."""
    import csv
    last_error = None
    for enc in ("utf-8-sig", "utf-8", "cp1252", "latin1"):
        try:
            with open(path, "r", encoding=enc, errors="strict", newline="") as f:
                sample = f.read(65536)
                f.seek(0)
                # Prefer the delimiter that produces a stable multi-column structure.
                # This avoids commas inside Indian currency values (1,25,000.00)
                # being mistaken for CSV separators in pipe/tab exports.
                sample_lines = [ln for ln in sample.splitlines() if ln.strip()][:200]
                best = None
                for cand in ("|", "\t", ";", ","):
                    counts = [ln.count(cand) for ln in sample_lines]
                    positive = [c for c in counts if c > 0]
                    if not positive:
                        continue
                    from collections import Counter
                    mode_count, mode_freq = Counter(positive).most_common(1)[0]
                    coverage = mode_freq / max(len(sample_lines), 1)
                    score = (coverage, mode_count, mode_freq)
                    if best is None or score > best[0]:
                        best = (score, cand)
                delimiter = best[1] if best else ","
                rows = list(csv.reader(f, delimiter=delimiter))
            if not rows:
                return [("DELIMITED", pd.DataFrame())]
            width = max(len(r) for r in rows)
            padded = [r + [""] * (width - len(r)) for r in rows]
            return [("DELIMITED", pd.DataFrame(padded, dtype=str))]
        except Exception as exc:
            last_error = exc
    raise last_error


def read_excel_csv(path):
    ext = os.path.splitext(path)[1].lower()
    if ext in (".csv", ".txt"):
        return _read_delimited_robust(path)
    if ext in (".html", ".htm"):
        tables = pd.read_html(path, header=None, keep_default_na=False)
        return [(f"HTML_{i+1}", t.astype(str)) for i, t in enumerate(tables)]

    engine = "xlrd" if ext == ".xls" else "openpyxl"
    try:
        book = pd.ExcelFile(path, engine=engine)
        outputs = []
        for sheet in book.sheet_names:
            try:
                raw = pd.read_excel(book, sheet_name=sheet, header=None, dtype=str, keep_default_na=False)
                outputs.append((sheet, raw))
            except Exception as exc:
                logger.warning("Sheet '%s' skipped: %s", sheet, exc)
        if outputs:
            return outputs
    except Exception as excel_exc:
        # Many bank '.xls' downloads are actually HTML tables with an XLS extension.
        try:
            tables = pd.read_html(path, header=None, keep_default_na=False)
            if tables:
                logger.info("HTML-disguised Excel detected: %s", os.path.basename(path))
                return [(f"HTML_XLS_{i+1}", t.astype(str)) for i, t in enumerate(tables)]
        except Exception:
            pass
        raise excel_exc
    return []


# ============================================================
# UNIVERSAL PDF TEXT PARSER  (works for ALL banks)
# ============================================================
def is_bank_statement_pdf(text):
    """Broad bank-statement fingerprint; generic unknown-bank layouts are allowed in V7.1."""
    t = clean_text(text).lower()
    bank_keywords = [
        "bank", "statement", "account", "balance", "transaction", "debit", "credit",
        "deposit", "withdrawal", "particulars", "narration", "description", "cheque",
        "reference", "opening balance", "closing balance", "ifsc", "branch",
    ]
    score = sum(1 for kw in bank_keywords if kw in t)
    date_present = any(p.search(t) for p, _ in UNIVERSAL_DATE_PATTERNS)
    money_present = bool(re.search(r"\d[\d,]*\.\d{2}", t))
    if score >= 3:
        return True
    if UNKNOWN_BANK_GENERIC_PARSER and score >= 2 and date_present and money_present:
        return True
    return False


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
                balance_recon_status = "UNVERIFIED"
                balance_recon_difference = np.nan

                if (
                    not is_bf
                    and amount is not None
                    and previous_signed_balance is not None
                ):
                    balance_change = current_balance - previous_signed_balance
                    balance_recon_difference = abs(balance_change) - abs(amount)
                    tolerance = max(0.02, abs(amount) * 0.000001)
                    if abs(balance_recon_difference) <= tolerance:
                        balance_recon_status = "PASS"
                        if balance_change < 0:
                            debit = amount
                        elif balance_change > 0:
                            credit = amount
                    else:
                        # Audit-safe V7: keep direction unresolved instead of
                        # silently forcing debit/credit when the balance does not reconcile.
                        balance_recon_status = "CHECK"
                elif not is_bf and amount is not None:
                    # First transaction: use an explicit Dr/Cr marker only.
                    if re.search(r"\bDr\b", balance_line, flags=re.I):
                        debit = amount
                        balance_recon_status = "MARKER ONLY"
                    elif re.search(r"\bCr\b", balance_line, flags=re.I):
                        credit = amount
                        balance_recon_status = "MARKER ONLY"
                    else:
                        balance_recon_status = "UNRESOLVED"

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
                    "Balance Reconciliation Status": balance_recon_status,
                    "Balance Reconciliation Difference": balance_recon_difference,
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



def nonblank_nunique(series):
    """Count unique nonblank values without pandas downcasting warnings."""
    s = series.astype("string").str.strip()
    s = s[(s.notna()) & (s != "")]
    return int(s.nunique())


# ============================================================
# V6 — ACCOUNT PROFILE / CLASSIFICATION / AUDIT HELPERS
# ============================================================
def detect_bank_from_filename(filename):
    """Fallback bank detection when statement metadata is sparse."""
    n = norm(os.path.basename(filename))
    aliases = {
        "ICICI Bank": ["icici"],
        "State Bank of India (SBI)": ["sbi", "state bank"],
        "HDFC Bank": ["hdfc"],
        "Axis Bank": ["axis"],
        "Kotak Mahindra Bank": ["kotak"],
        "Punjab National Bank (PNB)": ["pnb", "punjab national"],
        "Bank of Baroda (BOB)": ["bob", "bank of baroda"],
        "Canara Bank": ["canara"],
        "Union Bank of India": ["union bank"],
        "IDBI Bank": ["idbi"],
        "Yes Bank": ["yes bank", "yesbank"],
        "IndusInd Bank": ["indusind"],
        "Federal Bank": ["federal"],
        "RBL Bank": ["rbl"],
        "Bandhan Bank": ["bandhan"],
        "AU Small Finance Bank": ["au small", "aubank"],
        "Jana Small Finance Bank": ["jana"],
        "Suryoday Small Finance Bank": ["suryoday"],
        "Utkarsh Small Finance Bank": ["utkarsh"],
        "Karur Vysya Bank": ["karur vysya", "kvb"],
        "Indian Overseas Bank": ["iob", "indian overseas"],
        "Bank of India": ["bank of india", "boi"],
    }
    aliases.update({
        "IDFC FIRST Bank": ["idfc first", "idfcfirst"],
        "Indian Bank": ["indian bank"],
        "Central Bank of India": ["central bank"],
        "UCO Bank": ["uco"],
        "Bank of Maharashtra": ["bank of maharashtra", "bom"],
        "Punjab & Sind Bank": ["punjab sind", "punjab and sind", "psb"],
        "City Union Bank": ["city union", "cub"],
        "Tamilnad Mercantile Bank": ["tamilnad mercantile", "tmb"],
        "South Indian Bank": ["south indian", "sib"],
        "Karnataka Bank": ["karnataka bank"],
        "DCB Bank": ["dcb"],
        "CSB Bank": ["csb", "catholic syrian"],
        "Dhanlaxmi Bank": ["dhanlaxmi"],
        "Jammu & Kashmir Bank": ["jammu kashmir", "j k bank", "jk bank"],
        "Ujjivan Small Finance Bank": ["ujjivan"],
        "Equitas Small Finance Bank": ["equitas"],
        "ESAF Small Finance Bank": ["esaf"],
        "Capital Small Finance Bank": ["capital sfb"],
        "Unity Small Finance Bank": ["unity sfb"],
        "Shivalik Small Finance Bank": ["shivalik"],
        "Airtel Payments Bank": ["airtel payments", "airtel bank"],
        "India Post Payments Bank": ["ippb", "india post payments"],
        "Fino Payments Bank": ["fino payments", "fino bank"],
        "Jio Payments Bank": ["jio payments", "jio bank"],
        "NSDL Payments Bank": ["nsdl payments", "nsdl bank"],
    })
    for bank, keys in aliases.items():
        if any(k in n for k in keys):
            return bank
    return "Unknown Bank"


def extract_account_profile_from_text(text_value, filename=""):
    """Best-effort account metadata extraction. Missing values remain blank."""
    raw = text_value or ""
    flat = clean_text(raw)
    bank = detect_bank_from_text(raw)
    if bank == "Unknown Bank":
        bank = detect_bank_from_filename(filename)

    def first_match(patterns):
        for pat in patterns:
            m = re.search(pat, raw, flags=re.I | re.M)
            if m:
                return clean_text(m.group(1))
        return ""

    account_number = first_match([
        r"(?:Account\s*(?:No\.?|Number)|A/C\s*(?:No\.?|Number))\s*[:\-]?\s*([A-Z0-9X*\-]{4,30})",
        r"(?:Account)\s*[:\-]?\s*([0-9X*]{6,30})",
    ])
    ifsc = first_match([
        r"(?:IFSC(?:\s*Code)?)\s*[:\-]?\s*([A-Z]{4}0[A-Z0-9]{6})",
    ])
    holder = first_match([
        r"(?:Customer\s*Name|Account\s*Holder(?:\s*Name)?|Name)\s*[:\-]\s*([^\n\r]{3,100})",
    ])
    period_from = first_match([
        r"(?:Statement\s*Period|Period)\s*(?:From)?\s*[:\-]?\s*(\d{1,2}[-/.][A-Za-z0-9]{2,9}[-/.][0-9]{2,4})",
        r"From\s*[:\-]?\s*(\d{1,2}[-/.][A-Za-z0-9]{2,9}[-/.][0-9]{2,4})",
    ])
    period_to = first_match([
        r"(?:Statement\s*Period|Period).*?(?:To)\s*[:\-]?\s*(\d{1,2}[-/.][A-Za-z0-9]{2,9}[-/.][0-9]{2,4})",
        r"To\s*[:\-]?\s*(\d{1,2}[-/.][A-Za-z0-9]{2,9}[-/.][0-9]{2,4})",
    ])

    masked = account_number[-4:] if len(account_number) >= 4 else account_number
    account_key = f"{bank} | {masked}" if masked else f"{bank} | {os.path.basename(filename)}"

    return {
        "Source File": os.path.basename(filename),
        "Bank Name": bank,
        "Account Holder": holder,
        "Account Number": account_number,
        "Account Last 4": masked,
        "IFSC": ifsc,
        "Statement From": period_from,
        "Statement To": period_to,
        "Account Key": account_key,
    }


def build_account_profiles(paths):
    """Read only statement headers/first page for account metadata."""
    profiles = []
    for path in paths:
        base = os.path.basename(path)
        ext = os.path.splitext(path)[1].lower()
        profile_text = ""
        try:
            if ext == ".pdf":
                profile_text = get_first_pdf_page_text(path)
            else:
                sources = read_excel_csv(path)
                chunks = []
                for _, raw in sources[:3]:
                    if raw is not None and not raw.empty:
                        subset = raw.head(50).fillna("").astype(str)
                        chunks.append("\n".join(" | ".join(row) for row in subset.values.tolist()))
                profile_text = "\n".join(chunks)
        except Exception:
            profile_text = ""
        profiles.append(extract_account_profile_from_text(profile_text, path))
    return pd.DataFrame(profiles).drop_duplicates(subset=["Source File"], keep="first")


def classify_transaction_type(narration):
    s = clean_text(narration).upper()
    rules = [
        ("UPI", ["UPI/", "UPI-", "UPI "]),
        ("IMPS", ["IMPS", "MMT/IMPS"]),
        ("RTGS", ["RTGS"]),
        ("NEFT", ["NEFT"]),
        ("CHEQUE", ["CHEQUE", "CHQ", "CTS"]),
        ("ATM WITHDRAWAL", ["ATM", "CASH WDL", "CASH WITHDRAWAL"]),
        ("CASH DEPOSIT", ["CASH DEP", "CASH DEPOSIT"]),
        ("POS / CARD", ["POS", "ECOM", "CARD", "SWIPE"]),
        ("BANK CHARGES", ["CHARGE", "CHGS", "FEE", "COMMISSION", "PENAL"]),
        ("GST / TAX", ["GST", "CGST", "SGST", "IGST", "TAX"]),
        ("INTEREST", ["INTEREST", "INT.PD", "INT PAID", "INT RECEIVED"]),
        ("REFUND / REVERSAL", ["REFUND", "REVERSAL", "REVERSED", "RETURN"]),
        ("INTERNAL TRANSFER", ["INFT", "INTERNAL TRANSFER", "OWN ACCOUNT", "SELF TRANSFER"]),
        ("NACH / ECS", ["NACH", "ECS", "ACH"]),
        ("BANK TRANSFER", ["TRF", "TRANSFER"]),
    ]
    for category, keywords in rules:
        if any(k in s for k in keywords):
            return category
    return "OTHER"


def explicit_gst_component(narration):
    """Extract GST amount only when narration explicitly states an amount."""
    s = clean_text(narration)
    patterns = [
        r"(?:GST|IGST|CGST|SGST)\s*(?:AMT|AMOUNT)?\s*[:\-]?\s*(?:RS\.?|INR|₹)?\s*([\d,]+(?:\.\d{1,2})?)",
        r"(?:RS\.?|INR|₹)?\s*([\d,]+(?:\.\d{1,2})?)\s*(?:GST|IGST|CGST|SGST)",
    ]
    for pat in patterns:
        m = re.search(pat, s, flags=re.I)
        if m:
            return parse_number(m.group(1))
    return np.nan


def calculate_confidence(row):
    score = 0
    reasons = []

    if pd.notna(row.get("Date")):
        score += 15
    else:
        reasons.append("Missing date")

    if pd.notna(row.get("Transaction Amount")):
        score += 20
    else:
        reasons.append("Missing amount")

    if clean_text(row.get("Direction")) in ("DEBIT", "CREDIT"):
        score += 15
    else:
        reasons.append("Direction unresolved")

    if clean_text(row.get("UTR / Reference")):
        score += 15
    else:
        reasons.append("Missing UTR/reference")

    if clean_text(row.get("Customer Name")):
        score += 10
    else:
        reasons.append("Missing customer")

    if pd.notna(row.get("Balance")):
        score += 10

    recon = clean_text(row.get("Balance Reconciliation Status")).upper()
    if recon == "PASS":
        score += 10
    elif recon == "CHECK":
        reasons.append("Balance movement mismatch")
    elif recon in ("UNRESOLVED", "UNVERIFIED"):
        reasons.append("Balance reconciliation unavailable")

    if clean_text(row.get("Possible Duplicate")) != "YES":
        score += 5
    else:
        reasons.append("Possible duplicate")

    return min(score, 100), "; ".join(dict.fromkeys(reasons))

def add_format_diagnostic(source_file, source_part, status, sample_text="", header_hits_value="", note=""):
    if not DIAGNOSTIC_MODE:
        return
    sample = clean_text(sample_text)[:1500]
    format_diagnostic_records.append({
        "Source File": source_file,
        "Source Part": source_part,
        "Status": status,
        "Detected Bank": detect_bank_from_text(sample_text) if sample_text else detect_bank_from_filename(source_file),
        "Header Hits": header_hits_value,
        "Sample / First Content": sample,
        "Note": note,
    })


# ============================================================
# LOAD SOURCES — V7.1 AUDIT / RAW / ROW-CONTROL
# ============================================================
all_frames = []
source_map = []
all_pdf_recon = []

logger.info("=" * 82)
logger.info("UNIVERSAL BANK BOOKS REPORT %s", VERSION)
logger.info("Author: %s", AUTHOR_NAME)
logger.info("Folder: %s", BASE_DIR)
logger.info("Source files found: %s", len(input_files))

for path in input_files:
    base = os.path.basename(path)
    ext = os.path.splitext(path)[1].lower()
    logger.info("Reading: %s", base)

    if ext == ".pdf":
        try:
            first_text = get_first_pdf_page_text(path)
        except Exception as exc:
            source_map.append({"Source File": base, "Source Part": "PDF", "Status": "READ ERROR", "Transaction Count": 0, "Note": str(exc)})
            log_exception("PDF Read Error", exc, base, "PDF")
            continue

        detected_bank = detect_bank_from_text(first_text)
        logger.info("  Detected bank: %s", detected_bank)

        if not is_bank_statement_pdf(first_text):
            source_map.append({"Source File": base, "Source Part": "PDF", "Status": "NOT A BANK STATEMENT", "Transaction Count": 0, "Note": "PDF does not appear to be a bank statement."})
            log_exception("Not Bank Statement", "PDF did not meet bank-statement fingerprint threshold", base, "PDF")
            add_format_diagnostic(base, "PDF", "FINGERPRINT REVIEW", first_text, note="Bank-statement fingerprint was weak; inspect first-page text.")
            row_control_seed.append({"Source File": base, "Source Part": "PDF", "Source Rows": 0, "Candidate Data Rows": 0, "Imported Rows": 0, "Processed Rows": 0, "Rejected Rows": 0, "Output Rows": 0, "Status": "REVIEW", "Note": "Not recognized as a bank statement"})
            continue

        try:
            pdf_tx, pdf_recon = parse_universal_pdf_text(path)
        except Exception as exc:
            log_exception("PDF Parse Error", traceback.format_exc(), base, "UNIVERSAL PDF TEXT")
            logger.error("  Universal PDF parser failed: %s", exc)
            pdf_tx = pd.DataFrame()
            pdf_recon = pd.DataFrame()

        if not pdf_tx.empty:
            pdf_tx["Source File"] = base
            pdf_tx["Source Part"] = "Page " + pdf_tx["Source Page"].astype(str)
            pdf_tx["Source Row"] = ""
            all_frames.append(pdf_tx)

            if not pdf_recon.empty:
                pdf_recon["Source File"] = base
                all_pdf_recon.append(pdf_recon)
                bad_pages = int(pdf_recon["Status"].eq("CHECK").sum())
            else:
                bad_pages = 0

            if PRESERVE_RAW_DATA:
                store_raw_data(pdf_tx.copy(), base, "PDF_PARSED_TRANSACTIONS")

            source_map.append({
                "Source File": base,
                "Source Part": f"UNIVERSAL PDF ({detected_bank})",
                "Status": "RECOGNIZED" if bad_pages == 0 else "RECOGNIZED / RECON CHECK",
                "Transaction Count": len(pdf_tx),
                "Note": f"Page reconciliation issues: {bad_pages}",
            })
            row_control_seed.append({
                "Source File": base,
                "Source Part": "UNIVERSAL PDF TEXT",
                "Source Rows": len(pdf_tx),
                "Candidate Data Rows": len(pdf_tx),
                "Imported Rows": len(pdf_tx),
                "Processed Rows": len(pdf_tx),
                "Rejected Rows": 0,
                "Output Rows": len(pdf_tx),
                "Status": "PASS" if bad_pages == 0 else "REVIEW",
                "Note": f"PDF page reconciliation CHECK count: {bad_pages}",
            })
            parser_recon_records.append({
                "Source File": base,
                "Parser": "UNIVERSAL PDF TEXT + BALANCE RECON",
                "Transactions": len(pdf_tx),
                "Debit Total": float(pdf_tx["Debit"].fillna(0).sum()),
                "Credit Total": float(pdf_tx["Credit"].fillna(0).sum()),
                "PDF CHECK Pages": bad_pages,
                "Status": "PASS" if bad_pages == 0 else "REVIEW",
            })
            logger.info("  Universal PDF parser: %s transaction(s)", f"{len(pdf_tx):,}")
            continue

        logger.info("  Text parser found no transactions. Trying PDF table extraction / OCR.")
        try:
            pdf_sources = extract_pdf_tables(path)
        except Exception as exc:
            pdf_sources = []
            log_exception("PDF Table Extract Error", traceback.format_exc(), base, "PDF TABLE")
            logger.error("  PDF table extraction failed: %s", exc)

        if not pdf_sources and OCR_ENABLED:
            pdf_sources = ocr_pdf_optional(path)

        recognized_pdf = False
        for source_name, raw, page_no in pdf_sources:
            if PRESERVE_RAW_DATA:
                store_raw_data(raw.copy(), base, source_name)

            header_row = locate_generic_header(raw)
            if header_row is None:
                sample = " | ".join(clean_text(x) for x in raw.head(8).fillna("").astype(str).values.flatten().tolist())
                add_format_diagnostic(base, source_name, "NO HEADER", sample, note="PDF table/OCR extracted but no supported header was detected.")
                row_control_seed.append({"Source File": base, "Source Part": source_name, "Source Rows": len(raw), "Candidate Data Rows": 0, "Imported Rows": 0, "Processed Rows": 0, "Rejected Rows": 0, "Output Rows": 0, "Status": "REVIEW", "Note": "No bank header detected"})
                continue

            table = dataframe_from_header(raw, header_row)
            tx = standardize_generic_table(table)
            if tx.empty:
                row_control_seed.append({"Source File": base, "Source Part": source_name, "Source Rows": len(raw), "Candidate Data Rows": len(table), "Imported Rows": 0, "Processed Rows": 0, "Rejected Rows": len(table), "Output Rows": 0, "Status": "REVIEW", "Note": "Header detected but no transactions"})
                continue

            tx["Source File"] = base
            tx["Source Part"] = source_name
            tx["Source Row"] = (tx.index.to_series().astype(int) + header_row + 2).values
            tx["Source Page"] = page_no
            tx["Parser"] = "GENERIC PDF TABLE / OCR"
            all_frames.append(tx)
            recognized_pdf = True
            source_map.append({"Source File": base, "Source Part": source_name, "Status": "RECOGNIZED", "Transaction Count": len(tx), "Note": f"Bank: {detected_bank}; header row: {header_row + 1}"})
            row_control_seed.append({"Source File": base, "Source Part": source_name, "Source Rows": len(raw), "Candidate Data Rows": len(table), "Imported Rows": len(tx), "Processed Rows": len(tx), "Rejected Rows": max(len(table) - len(tx), 0), "Output Rows": len(tx), "Status": "PASS", "Note": "Generic PDF table/OCR parser"})
            parser_recon_records.append({"Source File": base, "Parser": "GENERIC PDF TABLE / OCR", "Transactions": len(tx), "Debit Total": float(tx["Debit"].fillna(0).sum()), "Credit Total": float(tx["Credit"].fillna(0).sum()), "PDF CHECK Pages": "", "Status": "PASS"})

        if not recognized_pdf:
            source_map.append({"Source File": base, "Source Part": "PDF", "Status": "UNRECOGNIZED", "Transaction Count": 0, "Note": f"Bank detected: {detected_bank}. No reliable transaction table recognized."})
            log_exception("Unrecognized PDF Format", "No reliable transaction rows recognized", base, "PDF")
            add_format_diagnostic(base, "PDF", "UNRECOGNIZED", first_text, note="Text, table and OCR parsers did not produce reliable transactions.")
        continue

    # EXCEL / CSV
    try:
        sources = read_excel_csv(path)
    except Exception as exc:
        source_map.append({"Source File": base, "Source Part": "", "Status": "READ ERROR", "Transaction Count": 0, "Note": str(exc)})
        log_exception("File Read Error", traceback.format_exc(), base, "")
        continue

    recognized_file = False
    for source_name, raw in sources:
        if PRESERVE_RAW_DATA:
            store_raw_data(raw.copy(), base, source_name)

        header_row = locate_generic_header(raw)
        if header_row is None:
            sample = " | ".join(clean_text(x) for x in raw.head(8).fillna("").astype(str).values.flatten().tolist())
            add_format_diagnostic(base, source_name, "NO HEADER", sample, note="Excel/CSV/HTML/TXT source read successfully but no supported header was detected.")
            row_control_seed.append({"Source File": base, "Source Part": source_name, "Source Rows": len(raw), "Candidate Data Rows": 0, "Imported Rows": 0, "Processed Rows": 0, "Rejected Rows": 0, "Output Rows": 0, "Status": "REVIEW", "Note": "No bank header detected"})
            continue

        table = dataframe_from_header(raw, header_row)
        tx = standardize_generic_table(table)
        if tx.empty:
            row_control_seed.append({"Source File": base, "Source Part": source_name, "Source Rows": len(raw), "Candidate Data Rows": len(table), "Imported Rows": 0, "Processed Rows": 0, "Rejected Rows": len(table), "Output Rows": 0, "Status": "REVIEW", "Note": "Header detected but no transactions"})
            continue

        tx["Source File"] = base
        tx["Source Part"] = source_name
        tx["Source Row"] = (tx.index.to_series().astype(int) + header_row + 2).values
        tx["Source Page"] = ""
        tx["Parser"] = "GENERIC EXCEL/CSV"
        all_frames.append(tx)
        recognized_file = True
        source_map.append({"Source File": base, "Source Part": source_name, "Status": "RECOGNIZED", "Transaction Count": len(tx), "Note": f"Header row: {header_row + 1}"})
        row_control_seed.append({"Source File": base, "Source Part": source_name, "Source Rows": len(raw), "Candidate Data Rows": len(table), "Imported Rows": len(tx), "Processed Rows": len(tx), "Rejected Rows": max(len(table) - len(tx), 0), "Output Rows": len(tx), "Status": "PASS", "Note": f"Header row: {header_row + 1}"})
        parser_recon_records.append({"Source File": base, "Parser": "GENERIC EXCEL/CSV", "Transactions": len(tx), "Debit Total": float(tx["Debit"].fillna(0).sum()), "Credit Total": float(tx["Credit"].fillna(0).sum()), "PDF CHECK Pages": "", "Status": "PASS"})
        logger.info("  %s: %s transaction(s)", source_name, f"{len(tx):,}")

    if not recognized_file:
        source_map.append({"Source File": base, "Source Part": "", "Status": "UNRECOGNIZED", "Transaction Count": 0, "Note": "Expected Date + Debit/Credit, or Date + Amount + DR/CR."})
        log_exception("Unrecognized Format", "No supported bank header found", base, "")
        add_format_diagnostic(base, "", "UNRECOGNIZED", note="No supported header found in any source part.")

if not all_frames:
    raise ValueError(
        "\\nNo bank transactions were recognized.\\n"
        "Expected Date + Debit/Credit or Date + Amount + DR/CR.\\n"
        "Searchable/text PDFs are preferred; OCR is optional for scanned PDFs."
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
    "Balance Reconciliation Status", "Balance Reconciliation Difference",
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
# V6 — ENRICH TRANSACTIONS FOR BOOKS / AUDIT
# ============================================================
account_profile = build_account_profiles(input_files)
profile_lookup = account_profile.set_index("Source File").to_dict("index") if not account_profile.empty else {}

for col in ["Bank Name", "Account Holder", "Account Number", "Account Last 4", "IFSC", "Account Key"]:
    transactions[col] = transactions["Source File"].map(
        lambda f: profile_lookup.get(f, {}).get(col, "")
    )

transactions["Transaction Type"] = transactions["Narration"].apply(classify_transaction_type)
transactions["Explicit GST Component"] = transactions["Narration"].apply(explicit_gst_component)

confidence_values = transactions.apply(calculate_confidence, axis=1)
transactions["Confidence Score"] = confidence_values.apply(lambda x: x[0])
transactions["Review Reason"] = confidence_values.apply(lambda x: x[1])
transactions["Review Status"] = np.select(
    [
        transactions["Confidence Score"] >= REVIEW_OK_THRESHOLD,
        transactions["Confidence Score"] >= REVIEW_MIN_THRESHOLD,
    ],
    ["OK", "REVIEW"],
    default="CRITICAL REVIEW",
)

review_required = transactions[
    transactions["Review Status"].ne("OK")
    | transactions["Possible Duplicate"].eq("YES")
].copy()


# V7 row-count control: transparent source/candidate/import/output accounting.
row_control_df = pd.DataFrame(row_control_seed)
if not row_control_df.empty:
    duplicate_by_file = transactions.groupby("Source File")["Possible Duplicate"].apply(lambda s: int(s.eq("YES").sum())).to_dict()
    review_by_file = transactions.groupby("Source File")["Review Status"].apply(lambda s: int(s.ne("OK").sum())).to_dict()
    row_control_df["Duplicate Flagged"] = row_control_df["Source File"].map(duplicate_by_file).fillna(0).astype(int)
    row_control_df["Review Flagged"] = row_control_df["Source File"].map(review_by_file).fillna(0).astype(int)
    row_control_df["Difference"] = row_control_df["Imported Rows"] - row_control_df["Processed Rows"]
    row_control_df["Status"] = np.where(
        (row_control_df["Status"] == "PASS") & (row_control_df["Difference"] == 0),
        "PASS",
        row_control_df["Status"],
    )
else:
    row_control_df = pd.DataFrame(columns=[
        "Source File", "Source Part", "Source Rows", "Candidate Data Rows",
        "Imported Rows", "Processed Rows", "Rejected Rows", "Output Rows",
        "Duplicate Flagged", "Review Flagged", "Difference", "Status", "Note",
    ])

parser_reconciliation = pd.DataFrame(parser_recon_records)
raw_data_index = pd.DataFrame(raw_data_index_records)

# Safe PG/GST validation: disabled by default for bank statements.
if PG_VALIDATION_ENABLED:
    pg_gst_validation = transactions[[
        "Date", "Source File", "UTR / Reference", "Customer Name",
        "Transaction Amount", "Narration"
    ]].copy()
    pg_gst_validation["Expected PG Charge"] = (pg_gst_validation["Transaction Amount"] * PG_CHARGE_RATE).round(ROUND_DECIMALS)
    pg_gst_validation["Expected GST"] = (pg_gst_validation["Expected PG Charge"] * GST_RATE).round(ROUND_DECIMALS)
    pg_gst_validation["Validation Status"] = "MODEL ONLY - REQUIRES PG EVIDENCE"
else:
    pg_gst_validation = pd.DataFrame([{
        "Validation Status": "DISABLED",
        "Note": "Bank statements do not by themselves prove a payment-gateway charge. Set pg_validation_enabled: true only for a controlled PG-validation workflow.",
    }])

refund_chargeback_review = transactions[
    transactions["Narration"].fillna("").astype(str).str.contains(
        r"REFUND|REVERSAL|REVERSED|RETURN|CHARGEBACK|CHARGE BACK|DISPUTE",
        case=False, regex=True, na=False,
    )
].copy()

# Account-wise summary
account_wise = (
    transactions.groupby(
        ["Account Key", "Bank Name", "Account Holder", "Account Number", "IFSC"],
        dropna=False,
    )
    .agg(
        Transaction_Count=("Transaction Amount", "count"),
        Debit_Count=("Debit", lambda x: int(x.notna().sum())),
        Debit_Total=("Debit", "sum"),
        Credit_Count=("Credit", lambda x: int(x.notna().sum())),
        Credit_Total=("Credit", "sum"),
        First_Date=("Date", "min"),
        Last_Date=("Date", "max"),
        Unique_Customers=("Customer Name", nonblank_nunique),
        Unique_UTRs=("UTR / Reference", nonblank_nunique),
    )
    .reset_index()
)
account_wise["Net_Credit_Minus_Debit"] = account_wise["Credit_Total"].fillna(0) - account_wise["Debit_Total"].fillna(0)

# Transaction type summary
transaction_type_summary = (
    transactions.groupby("Transaction Type", dropna=False)
    .agg(
        Transaction_Count=("Transaction Amount", "count"),
        Debit_Count=("Debit", lambda x: int(x.notna().sum())),
        Debit_Total=("Debit", "sum"),
        Credit_Count=("Credit", lambda x: int(x.notna().sum())),
        Credit_Total=("Credit", "sum"),
    )
    .reset_index()
)
transaction_type_summary["Net_Credit_Minus_Debit"] = transaction_type_summary["Credit_Total"].fillna(0) - transaction_type_summary["Debit_Total"].fillna(0)

# Bank charges / GST-related rows. No GST split is invented.
charges_mask = transactions["Transaction Type"].isin(["BANK CHARGES", "GST / TAX"]) | transactions["Narration"].str.contains(
    r"GST|IGST|CGST|SGST|CHARGE|CHGS|FEE|COMMISSION", case=False, na=False, regex=True
)
bank_charges_gst = transactions.loc[charges_mask, [
    "Date", "Account Key", "Bank Name", "Narration", "Debit", "Credit",
    "Transaction Type", "Explicit GST Component", "UTR / Reference", "Source File",
]].copy()

# UTR duplicate / conflict control
utr_nonblank = transactions[
    transactions["UTR / Reference"].fillna("").astype(str).str.strip().ne("")
].copy()
if utr_nonblank.empty:
    utr_duplicate_check = pd.DataFrame(columns=[
        "UTR / Reference", "Occurrence_Count", "Unique_Amounts", "First_Date", "Last_Date",
        "Total_Debit", "Total_Credit", "Status",
    ])
else:
    utr_duplicate_check = (
        utr_nonblank.groupby("UTR / Reference", dropna=False)
        .agg(
            Occurrence_Count=("Transaction Amount", "count"),
            Unique_Amounts=("Transaction Amount", "nunique"),
            First_Date=("Date", "min"),
            Last_Date=("Date", "max"),
            Total_Debit=("Debit", "sum"),
            Total_Credit=("Credit", "sum"),
        )
        .reset_index()
    )
    utr_duplicate_check["Status"] = np.select(
        [
            (utr_duplicate_check["Occurrence_Count"] > 1) & (utr_duplicate_check["Unique_Amounts"] > 1),
            utr_duplicate_check["Occurrence_Count"] > 1,
        ],
        ["CRITICAL - SAME UTR DIFFERENT AMOUNT", "DUPLICATE / REPEATED UTR"],
        default="UNIQUE",
    )
    utr_duplicate_check = utr_duplicate_check[utr_duplicate_check["Occurrence_Count"] > 1].copy()

# Account reconciliation: derive opening balance from first transaction's post-balance.
recon_rows = []
for account_key, grp in transactions.sort_values(["Date"]).groupby("Account Key", dropna=False):
    grp = grp.sort_values(["Date"]).copy()
    with_bal = grp[grp["Balance"].notna()].copy()
    opening = np.nan
    closing = np.nan
    if not with_bal.empty:
        first = with_bal.iloc[0]
        first_bal = float(first["Balance"])
        first_debit = 0.0 if pd.isna(first["Debit"]) else float(first["Debit"])
        first_credit = 0.0 if pd.isna(first["Credit"]) else float(first["Credit"])
        opening = first_bal + first_debit - first_credit
        closing = float(with_bal.iloc[-1]["Balance"])
    debit_total_ac = float(grp["Debit"].fillna(0).sum())
    credit_total_ac = float(grp["Credit"].fillna(0).sum())
    expected_closing = opening + credit_total_ac - debit_total_ac if pd.notna(opening) else np.nan
    difference = closing - expected_closing if pd.notna(closing) and pd.notna(expected_closing) else np.nan
    status = "OK" if pd.notna(difference) and abs(difference) <= 0.02 else ("CHECK" if pd.notna(difference) else "BALANCE NOT AVAILABLE")
    sample = grp.iloc[0]
    recon_rows.append({
        "Account Key": account_key,
        "Bank Name": sample.get("Bank Name", ""),
        "Account Number": sample.get("Account Number", ""),
        "First Date": grp["Date"].min(),
        "Last Date": grp["Date"].max(),
        "Derived Opening Balance": opening,
        "Total Debit": debit_total_ac,
        "Total Credit": credit_total_ac,
        "Expected Closing Balance": expected_closing,
        "Actual Closing Balance": closing,
        "Difference": difference,
        "Status": status,
    })
bank_reconciliation = pd.DataFrame(recon_rows)

# Data quality summary
quality_metrics = [
    ("Total Transactions", len(transactions)),
    ("Missing Date", int(transactions["Date"].isna().sum())),
    ("Missing Amount", int(transactions["Transaction Amount"].isna().sum())),
    ("Missing Direction", int(transactions["Direction"].fillna("").eq("").sum())),
    ("Missing Customer Name", int(transactions["Customer Name"].fillna("").astype(str).str.strip().eq("").sum())),
    ("Missing UTR / Reference", int(transactions["UTR / Reference"].fillna("").astype(str).str.strip().eq("").sum())),
    ("Missing Balance", int(transactions["Balance"].isna().sum())),
    ("Possible Duplicate Rows", int(transactions["Possible Duplicate"].eq("YES").sum())),
    ("Review Rows", int(transactions["Review Status"].eq("REVIEW").sum())),
    ("Critical Review Rows", int(transactions["Review Status"].eq("CRITICAL REVIEW").sum())),
    ("Duplicate / Conflicting UTRs", int(len(utr_duplicate_check))),
]
data_quality = pd.DataFrame(quality_metrics, columns=["Metric", "Count"])
data_quality["Percent_of_Transactions"] = np.where(
    data_quality["Metric"].eq("Total Transactions"),
    100.0,
    np.where(len(transactions) > 0, data_quality["Count"] / len(transactions) * 100.0, 0.0),
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
            nonblank_nunique,
        ),
        Unique_UTRs=(
            "UTR / Reference",
            nonblank_nunique,
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
                nonblank_nunique,
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


def write_df_split(wb, frame, sheet_base, money_cols=(), date_cols=(), count_cols=()):
    """Write safely within Excel row/name limits. Returns created sheet names."""
    used = {ws.title.lower() for ws in wb.worksheets}
    base = sanitize_sheet_name(sheet_base)
    if frame is None or frame.empty:
        name = sanitize_sheet_name(base, used)
        ws = wb.create_sheet(name)
        ws["A1"] = "No data available"
        return [name]

    chunk_size = min(MAX_ROWS_PER_SHEET, 1048575)
    created = []
    if len(frame) <= chunk_size:
        name = sanitize_sheet_name(base, used)
        ws = wb.create_sheet(name)
        write_df(ws, frame, money_cols, date_cols, count_cols)
        return [name]

    index_records = []
    for n, start in enumerate(range(0, len(frame), chunk_size), start=1):
        chunk = frame.iloc[start:start + chunk_size]
        name = sanitize_sheet_name(f"{base}_{n:03d}", used)
        ws = wb.create_sheet(name)
        write_df(ws, chunk, money_cols, date_cols, count_cols)
        created.append(name)
        index_records.append([name, start + 1, min(start + chunk_size, len(frame)), len(chunk)])

    index_name = sanitize_sheet_name(f"{base}_INDEX", used)
    ws = wb.create_sheet(index_name)
    write_df(ws, pd.DataFrame(index_records, columns=["Sheet Name", "Start Record", "End Record", "Rows"]), count_cols=("Start Record", "End Record", "Rows"))
    return created


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
    ("Review Required Rows", int(len(review_required))),
    ("Critical Review Rows", int(transactions["Review Status"].eq("CRITICAL REVIEW").sum())),
    ("Duplicate / Conflicting UTR Count", int(len(utr_duplicate_check))),
    ("Bank Reconciliation CHECK Count", int(bank_reconciliation["Status"].eq("CHECK").sum()) if not bank_reconciliation.empty else 0),
    ("Bank Charges / GST Related Rows", int(len(bank_charges_gst))),
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
# ACCOUNT_PROFILE
# ============================================================
ws = wb.create_sheet("ACCOUNT_PROFILE")
write_df(
    ws,
    account_profile,
)

# ============================================================
# TRANSACTION_REGISTER
# ============================================================
register_columns = [
    "Date", "Bank Name", "Account Key", "Account Number",
    "Customer Name", "UTR / Reference", "Narration",
    "Transaction Type", "Debit", "Credit", "Balance", "Direction",
    "Transaction Amount", "Balance Reconciliation Status", "Balance Reconciliation Difference",
    "Confidence Score", "Review Status", "Review Reason",
    "Customer Source", "UTR Source", "Possible Duplicate", "Parser", "Source File",
    "Source Part", "Source Page", "Source Row",
]
write_df_split(
    wb,
    transactions[register_columns],
    "TRANSACTION_REGISTER",
    money_cols=("Debit", "Credit", "Balance", "Transaction Amount", "Balance Reconciliation Difference"),
    date_cols=("Date",),
)


# ============================================================
# REVIEW_REQUIRED
# ============================================================
review_columns = [
    "Date", "Bank Name", "Account Key", "Customer Name", "UTR / Reference",
    "Narration", "Debit", "Credit", "Balance", "Direction", "Transaction Amount",
    "Balance Reconciliation Status", "Balance Reconciliation Difference",
    "Confidence Score", "Review Status", "Review Reason", "Possible Duplicate",
    "Source File", "Source Part", "Source Page", "Source Row",
]
write_df_split(
    wb,
    review_required[review_columns] if not review_required.empty else review_required,
    "REVIEW_REQUIRED",
    money_cols=("Debit", "Credit", "Balance", "Transaction Amount", "Balance Reconciliation Difference"),
    date_cols=("Date",),
    count_cols=("Confidence Score",),
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
# ACCOUNT_WISE_SUMMARY
# ============================================================
ws = wb.create_sheet("ACCOUNT_WISE_SUMMARY")
write_df(
    ws,
    account_wise,
    money_cols=("Debit_Total", "Credit_Total", "Net_Credit_Minus_Debit"),
    date_cols=("First_Date", "Last_Date"),
    count_cols=("Transaction_Count", "Debit_Count", "Credit_Count", "Unique_Customers", "Unique_UTRs"),
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
# TRANSACTION_TYPE_SUMMARY
# ============================================================
ws = wb.create_sheet("TRANSACTION_TYPE_SUMMARY")
write_df(
    ws,
    transaction_type_summary,
    money_cols=("Debit_Total", "Credit_Total", "Net_Credit_Minus_Debit"),
    count_cols=("Transaction_Count", "Debit_Count", "Credit_Count"),
)

# ============================================================
# BANK_CHARGES_GST
# ============================================================
ws = wb.create_sheet("BANK_CHARGES_GST")
write_df(
    ws,
    bank_charges_gst,
    money_cols=("Debit", "Credit", "Explicit GST Component"),
    date_cols=("Date",),
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
# UTR_DUPLICATE_CHECK
# ============================================================
ws = wb.create_sheet("UTR_DUPLICATE_CHECK")
write_df(
    ws,
    utr_duplicate_check,
    money_cols=("Total_Debit", "Total_Credit"),
    date_cols=("First_Date", "Last_Date"),
    count_cols=("Occurrence_Count", "Unique_Amounts"),
)
if not utr_duplicate_check.empty:
    status_col = list(utr_duplicate_check.columns).index("Status") + 1
    for r in range(2, ws.max_row + 1):
        if str(ws.cell(r, status_col).value).startswith("CRITICAL"):
            for c in range(1, ws.max_column + 1):
                ws.cell(r, c).fill = WARN_FILL

# ============================================================
# BANK_RECONCILIATION
# ============================================================
ws = wb.create_sheet("BANK_RECONCILIATION")
write_df(
    ws,
    bank_reconciliation,
    money_cols=(
        "Derived Opening Balance", "Total Debit", "Total Credit",
        "Expected Closing Balance", "Actual Closing Balance", "Difference",
    ),
    date_cols=("First Date", "Last Date"),
)
if not bank_reconciliation.empty:
    status_col = list(bank_reconciliation.columns).index("Status") + 1
    for r in range(2, ws.max_row + 1):
        if ws.cell(r, status_col).value == "CHECK":
            for c in range(1, ws.max_column + 1):
                ws.cell(r, c).fill = WARN_FILL

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
# DATA_QUALITY
# ============================================================
ws = wb.create_sheet("DATA_QUALITY")
write_df(
    ws,
    data_quality,
    count_cols=("Count",),
)
if not data_quality.empty:
    pct_col = list(data_quality.columns).index("Percent_of_Transactions") + 1
    for r in range(2, ws.max_row + 1):
        ws.cell(r, pct_col).number_format = '0.00%'
        value = ws.cell(r, pct_col).value
        if isinstance(value, (int, float)):
            ws.cell(r, pct_col).value = value / 100.0

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
    ["Review Required Rows", int(len(review_required))],
    ["Critical Review Rows", int(transactions["Review Status"].eq("CRITICAL REVIEW").sum())],
    ["Duplicate / Conflicting UTR Count", int(len(utr_duplicate_check))],
    ["Bank Reconciliation CHECK Count", int(bank_reconciliation["Status"].eq("CHECK").sum()) if not bank_reconciliation.empty else 0],
    ["Bank Charges / GST Related Rows", int(len(bank_charges_gst))],
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
# V7 — ROW_COUNT_CONTROL
# ============================================================
write_df_split(
    wb, row_control_df, "ROW_COUNT_CONTROL",
    count_cols=("Source Rows", "Candidate Data Rows", "Imported Rows", "Processed Rows", "Rejected Rows", "Output Rows", "Duplicate Flagged", "Review Flagged", "Difference"),
)

# ============================================================
# V7 — SOURCE_RAW_INDEX + RAW SOURCE SHEETS
# ============================================================
write_df_split(wb, raw_data_index, "SOURCE_RAW_INDEX", count_cols=("Start Record", "End Record", "Rows"))
for raw_sheet_name, raw_df in raw_data_sheets.items():
    # RAW sheet names are generated as short safe names.
    write_df_split(wb, raw_df, raw_sheet_name)

# ============================================================
# V7 — PARSER_RECONCILIATION
# ============================================================
write_df_split(
    wb, parser_reconciliation, "PARSER_RECONCILIATION",
    money_cols=("Debit Total", "Credit Total"),
    count_cols=("Transactions", "PDF CHECK Pages"),
)

# ============================================================
# V7.1 — FORMAT_DIAGNOSTICS
# ============================================================
format_diagnostics_df = pd.DataFrame(format_diagnostic_records) if format_diagnostic_records else pd.DataFrame(columns=[
    "Source File", "Source Part", "Status", "Detected Bank", "Header Hits", "Sample / First Content", "Note"
])
write_df_split(wb, format_diagnostics_df, "FORMAT_DIAGNOSTICS")

# V7.1 — BANK_SUPPORT_MATRIX
bank_support_matrix = pd.DataFrame([
    {"Bank / Category": bank, "Detection": "BUILT-IN", "Parsing": "GENERIC + PDF TEXT/TABLE", "Notes": "Layout may still require review if scanned/protected/non-tabular."}
    for bank in sorted(BANK_SIGNATURES.keys())
] + [
    {"Bank / Category": "Other Indian bank / co-operative bank", "Detection": "GENERIC", "Parsing": "HEADER ALIAS + DATE/AMOUNT/BALANCE", "Notes": "Works when standard transaction fields are extractable; FORMAT_DIAGNOSTICS captures unsupported layouts."}
])
write_df_split(wb, bank_support_matrix, "BANK_SUPPORT_MATRIX")

# ============================================================
# V7 — EXCEPTIONS
# ============================================================
exceptions_df = pd.DataFrame(exceptions_list) if exceptions_list else pd.DataFrame(columns=[
    "Exception Type", "Description", "Source File", "Source Part", "Source Row", "Timestamp"
])
write_df_split(wb, exceptions_df, "EXCEPTIONS")

# ============================================================
# V7 — CONFIG_USED
# ============================================================
config_rows = []
for key, value in CONFIG.items():
    if key == "column_aliases":
        for field, aliases in value.items():
            config_rows.append({"Setting": f"column_aliases.{field}", "Value": " | ".join(map(str, aliases))})
    else:
        config_rows.append({"Setting": key, "Value": str(value)})
config_rows.append({"Setting": "config_file", "Value": CONFIG_FILE})
config_rows.append({"Setting": "run_id", "Value": RUN_ID})
write_df_split(wb, pd.DataFrame(config_rows), "CONFIG_USED")

# ============================================================
# V7 — AUDIT_LOG
# ============================================================
audit_log_df = pd.DataFrame([
    ["Run ID", RUN_ID],
    ["Generated On", datetime.now().strftime("%Y-%m-%d %H:%M:%S")],
    ["Author", AUTHOR_NAME],
    ["Version", DISPLAY_VERSION],
    ["Config File", CONFIG_FILE],
    ["Audit Log File", LOG_FILE],
    ["Source Files", len(input_files)],
    ["Transactions", len(transactions)],
    ["Review Required", len(review_required)],
    ["Raw Preservation", PRESERVE_RAW_DATA],
    ["PG Validation Enabled", PG_VALIDATION_ENABLED],
], columns=["Metric", "Value"])
write_df_split(wb, audit_log_df, "AUDIT_LOG")

# ============================================================
# V7 — PG_GST_VALIDATION (SAFE / CONFIG-CONTROLLED)
# ============================================================
write_df_split(
    wb, pg_gst_validation, "PG_GST_VALIDATION",
    money_cols=("Transaction Amount", "Expected PG Charge", "Expected GST"),
    date_cols=("Date",),
)

# ============================================================
# V7 — REFUND_CHARGEBACK_REVIEW (ACTUAL NARRATION ONLY)
# ============================================================
refund_cols = [c for c in [
    "Date", "Bank Name", "Account Key", "Customer Name", "UTR / Reference",
    "Narration", "Debit", "Credit", "Balance", "Transaction Amount",
    "Source File", "Source Part", "Source Page", "Source Row"
] if c in refund_chargeback_review.columns]
write_df_split(
    wb,
    refund_chargeback_review[refund_cols] if not refund_chargeback_review.empty else refund_chargeback_review,
    "REFUND_CHARGEBACK_REVIEW",
    money_cols=("Debit", "Credit", "Balance", "Transaction Amount"),
    date_cols=("Date",),
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
        "V7.1 India-wide Bank Support",
        "Supports ALL major Indian banks: SBI, HDFC, ICICI, "
        "Axis, Kotak, PNB, BOB, Canara, Union, IDBI, Yes, "
        "IndusInd, Federal, RBL, Bandhan, AU Small Finance, "
        "and any bank producing standard tabular statements.",
    ],
    [
        "V7.1 PDF Parser",
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
        "V7 Audit Controls",
        "Adds account profile, confidence scoring, review queue, transaction-type classification, bank charges/GST review, UTR duplicate controls, account-level reconciliation and data-quality metrics.",
    ],
    [
        "Confidence / Review",
        "Confidence score is a control indicator only. Low-confidence rows are surfaced in REVIEW_REQUIRED and are not silently deleted or altered.",
    ],
    [
        "Bank Reconciliation",
        "Derived opening balance is calculated from the first available post-transaction balance, then checked against total debit/credit and actual closing balance.",
    ],
    [
        "Bank Charges / GST",
        "Rows are classified from actual narration. GST component is populated only when an explicit GST amount is present in the narration; no GST split is invented.",
    ],
    [
        "Supported Files",
        "XLSX, XLS, CSV, PDF, HTML/HTM and delimited TXT. Text-based PDFs are preferred; OCR is optional for scanned PDFs.",
    ],
    [
        "V7.1 Config",
        "config.yaml is validated at startup. Custom aliases extend built-in aliases rather than replacing them.",
    ],
    [
        "V7 Row Control",
        "ROW_COUNT_CONTROL records source rows, candidate rows, imported/processed/output rows, rejected rows and review flags without treating metadata/header rows as missing transactions.",
    ],
    [
        "V7 Raw Traceability",
        "When preserve_raw_data is true, source tables and parsed PDF transactions are preserved in split RAW_### sheets with SOURCE_RAW_INDEX.",
    ],
    [
        "V7 PG/GST Safety",
        "PG validation is disabled by default. Ordinary bank transactions are never automatically charged 0.8% + GST merely because rates exist in config.yaml.",
    ],
    [
        "V7.1 Format Diagnostics",
        "Unknown or unusual layouts are recorded in FORMAT_DIAGNOSTICS rather than silently discarded. BANK_SUPPORT_MATRIX lists built-in bank fingerprints plus generic/co-operative-bank support.",
    ],
    [
        "Coverage Limitation",
        "No software can guarantee every bank statement layout. Password-protected, corrupt, image-only or highly non-tabular statements may require OCR/manual review. V7.1 is designed for maximum Indian-bank coverage with audit-safe fallbacks.",
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
print(f"Review Rows  : {len(review_required):,}")
print(f"UTR Issues   : {len(utr_duplicate_check):,}")
print(f"Recon CHECK  : {int(bank_reconciliation["Status"].eq("CHECK").sum()) if not bank_reconciliation.empty else 0:,}")
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
wb.properties.creator = AUTHOR_NAME
wb.properties.lastModifiedBy = AUTHOR_NAME
wb.properties.title = f"Universal Bank Books Report {DISPLAY_VERSION}"
wb.properties.subject = "Bank Statement Books of Accounts / Audit Control Report"
wb.properties.description = f"Prepared by {AUTHOR_NAME} | Instagram: {AUTHOR_INSTAGRAM} | Run ID: {RUN_ID}"
logger.info("Saving workbook: %s", OUTPUT_FILE)
wb.save(OUTPUT_FILE)
logger.info("SUCCESS: %s", OUTPUT_FILE)
print("SUCCESS")
print("Output:", OUTPUT_FILE)
print("=" * 82)
