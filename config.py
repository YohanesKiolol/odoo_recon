"""
config.py — loads all settings from .env
Script never hardcodes any value; everything is user-configurable.
"""

import os
import sys
from pathlib import Path
import warnings

# Suppress openpyxl stylesheet warnings globally
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

def _load_dotenv(dotenv_path):
    try:
        with open(dotenv_path, "r", encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if line and not line.startswith("#"):
                    if "=" in line:
                        k, v = line.split("=", 1)
                        v = v.strip()
                        if v.startswith('"') and v.endswith('"'):
                            v = v[1:-1]
                        elif v.startswith("'") and v.endswith("'"):
                            v = v[1:-1]
                        os.environ[k.strip()] = v
    except FileNotFoundError:
        pass

# In frozen (PyInstaller) mode, config.py lives in a temp _MEI* folder.
# The .env file is always next to the .exe (or next to config.py in dev).
if getattr(sys, "frozen", False):
    _base = Path(sys.executable).parent   # folder containing BankRekonsiliasi.exe
else:
    _base = Path(__file__).parent         # folder containing config.py (project root)

_load_dotenv(_base / ".env")



def _require(key: str) -> str:
    val = os.getenv(key)
    if not val:
        raise EnvironmentError(
            f"Missing required config: '{key}'\n"
            f"Copy .env.example to .env and fill in the value."
        )
    return val.strip()


def _optional(key: str, default: str = "") -> str:
    return (os.getenv(key) or default).strip()


# ── Odoo ──────────────────────────────────────────────────────────────────────
ODOO_URL                = _optional("ODOO_URL")
ODOO_DASHBOARD_URL      = _optional("ODOO_DASHBOARD_URL")
ODOO_PAYMENTS_URL       = _optional("ODOO_PAYMENTS_URL")
ODOO_JOURNAL_IMPORT_URL = _optional("ODOO_JOURNAL_IMPORT_URL")
ODOO_COMPANY_NAME       = _optional("ODOO_COMPANY_NAME", "Eyerizz Eyewear")
ODOO_JOURNAL_EDC        = _optional("ODOO_JOURNAL_EDC", "EDC Settlement Journal")
ODOO_JOURNAL_AR         = _optional("ODOO_JOURNAL_AR", "MDR Fees Journal")
JOURNAL_TOLERANCE       = float(_optional("JOURNAL_TOLERANCE", "200"))
ODOO_JOURNAL_ENTRIES_URL = _optional("ODOO_JOURNAL_ENTRIES_URL")
ODO_JOURNAL_EXCEL_PATH  = _base / _optional("ODO_JOURNAL_EXCEL_PATH", "input/Journal Entries (account.move).xlsx")

JOURNAL_TEMPLATE_DIR    = _optional("JOURNAL_TEMPLATE_DIR", "")
JOURNAL_EXCEL_PATTERN   = _optional("JOURNAL_EXCEL_PATTERN", "Journal Entry (account.move).xlsx")

ODO_EXCEL_PATH      = Path(_require("ODO_EXCEL_PATH"))
ODO_AMOUNT_COLUMN   = _require("ODO_AMOUNT_COLUMN")
ODO_NUMBER_COLUMN   = _optional("ODO_NUMBER_COLUMN", "Number")
ODO_GROUP_BCA       = _optional("ODO_GROUP_BCA", "")
ODO_GROUP_MANDIRI   = _optional("ODO_GROUP_MANDIRI", "")
ODO_GROUP_BRI       = _optional("ODO_GROUP_BRI", "")

# ── Dynamic Accounts ──────────────────────────────────────────────────────────
BANK_ACCOUNTS = {
    "bca": {},
    "mandiri": {},
    "bri": {}
}

import os
for k, v in os.environ.items():
    if k.startswith("ACCOUNT_"):
        parts = k.split("_")
        # Example formats: 
        # ACCOUNT_BCA_main_MID
        # ACCOUNT_BCA_main_EDC_DEBIT
        if len(parts) >= 4:
            bank = parts[1].lower()
            if bank in BANK_ACCOUNTS:
                alias = parts[2].lower()
                prop = "_".join(parts[3:]).lower()
                if alias not in BANK_ACCOUNTS[bank]:
                    BANK_ACCOUNTS[bank][alias] = {"mid": "", "group": "", "acc": "", "edc_debit": "", "edc_credit": ""}
                BANK_ACCOUNTS[bank][alias][prop] = v

# Fallbacks for legacy single-account setup
for bank_key, legacy_group in [("bca", ODO_GROUP_BCA), ("mandiri", ODO_GROUP_MANDIRI), ("bri", ODO_GROUP_BRI)]:
    if legacy_group:
        if "main" not in BANK_ACCOUNTS[bank_key]:
            BANK_ACCOUNTS[bank_key]["main"] = {"mid": "", "group": legacy_group, "acc": "", "edc_debit": "", "edc_credit": ""}
        elif not BANK_ACCOUNTS[bank_key]["main"].get("group"):
            BANK_ACCOUNTS[bank_key]["main"]["group"] = legacy_group


# ── Mandiri ───────────────────────────────────────────────────────────────────
MANDIRI_ZIP_DIR       = Path(_require("MANDIRI_ZIP_DIR"))
MANDIRI_ZIP_PATTERN   = _optional("MANDIRI_ZIP_PATTERN", "MSR_*.zip")
MANDIRI_MUTATION_PATTERN = _optional("MANDIRI_MUTATION_PATTERN", "account_statement")
MANDIRI_ZIP_PASSWORD  = _require("MANDIRI_ZIP_PASSWORD")
MANDIRI_AMOUNT_COLUMN = _require("MANDIRI_AMOUNT_COLUMN")
MANDIRI_NUMBER_COLUMN = _optional("MANDIRI_NUMBER_COLUMN", "AUTHCODE")

# ── BCA ───────────────────────────────────────────────────────────────────────
BCA_EXCEL_DIR      = Path(_require("BCA_EXCEL_DIR"))
BCA_EXCEL_PATTERN  = _require("BCA_EXCEL_PATTERN")
BCA_MUTATION_PATTERN = _optional("BCA_MUTATION_PATTERN", "CorpAcctTrxn")
BCA_EXCEL_PASSWORD = _require("BCA_EXCEL_PASSWORD")
BCA_AMOUNT_COLUMN  = _require("BCA_AMOUNT_COLUMN")
BCA_DATE_COLUMN    = _require("BCA_DATE_COLUMN")
BCA_NUMBER_COLUMN  = _optional("BCA_NUMBER_COLUMN", "Trace Number")

# ── BRI ───────────────────────────────────────────────────────────────────────
BRI_ZIP_DIR       = Path(_require("BRI_ZIP_DIR"))
BRI_ZIP_PATTERN   = _optional("BRI_ZIP_PATTERN", "")
BRI_MUTATION_PATTERN = _optional("BRI_MUTATION_PATTERN", "e-StatementBRImo")
BRI_PDF_PATTERN   = _optional("BRI_PDF_PATTERN", "detail")
BRI_AMOUNT_COLUMN = _require("BRI_AMOUNT_COLUMN")
BRI_NUMBER_COLUMN = _optional("BRI_NUMBER_COLUMN", "Keterangan")

# ── Output ────────────────────────────────────────────────────────────────────
BASE_DIR = Path(_optional("BASE_DIR", "."))
INPUT_DIR = Path(_optional("INPUT_DIR", "input"))
MUTATION_DIR = Path(_optional("MUTATION_DIR", "mutation"))
OUTPUT_DIR = Path(_optional("OUTPUT_DIR", "output"))
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
