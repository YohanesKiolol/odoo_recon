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
ODOO_DB                 = _optional("ODOO_DB", "production")
ODOO_API_KEY            = _optional("ODOO_API_KEY")
ODOO_DASHBOARD_URL      = _optional("ODOO_DASHBOARD_URL")
ODOO_PAYMENTS_URL       = _optional("ODOO_PAYMENTS_URL")
ODOO_JOURNAL_IMPORT_URL = _optional("ODOO_JOURNAL_IMPORT_URL")
ODOO_COMPANY_NAME       = _optional("ODOO_COMPANY_NAME", "Eyerizz Eyewear")
ODOO_JOURNAL_EDC        = _optional("ODOO_JOURNAL_EDC", "EDC Settlement Journal")
ODOO_JOURNAL_AR         = _optional("ODOO_JOURNAL_AR", "MDR Fees Journal")
ODOO_ACCOUNT_BANK_DIFF_INCOME = _optional("ODOO_ACCOUNT_BANK_DIFF_INCOME", "82005 Bank Difference Income")
ODOO_ACCOUNT_BANK_DIFF_LOSS   = _optional("ODOO_ACCOUNT_BANK_DIFF_LOSS", "8107 Bank Difference Loss")
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


# ── Predefined Accounts ───────────────────────────────────────────────────────
def get_predefined_accounts() -> dict[str, dict[str, str]]:
    """Parse predefined Odoo user accounts from .env."""
    import json
    accounts: dict[str, dict[str, str]] = {}

    # 1. JSON parsing: ODOO_PREDEFINED_ACCOUNTS='{"fransisca": "pass123", "admin": "pass456"}'
    raw_json = _optional("ODOO_PREDEFINED_ACCOUNTS")
    if raw_json:
        try:
            data = json.loads(raw_json)
            if isinstance(data, dict):
                for k, v in data.items():
                    if isinstance(v, dict):
                        accounts[k] = {
                            "username": str(v.get("username", v.get("user", k))).strip(),
                            "password": str(v.get("password", v.get("pass", ""))).strip(),
                        }
                    elif isinstance(v, str):
                        accounts[k] = {
                            "username": k,
                            "password": v.strip(),
                        }
            elif isinstance(data, list):
                for item in data:
                    if isinstance(item, dict):
                        label = str(item.get("label", item.get("username", item.get("user", "")))).strip()
                        u = str(item.get("username", item.get("user", ""))).strip()
                        p = str(item.get("password", item.get("pass", ""))).strip()
                        if label:
                            accounts[label] = {"username": u or label, "password": p}
        except Exception:
            pass

    # 2. Comma-separated: ODOO_ACCOUNTS="fransisca:pass123,admin@eyerizz.com:pass456"
    #                      or with custom label: ODOO_ACCOUNTS="Fransisca=fransisca:pass123,Admin=admin@eyerizz.com:pass456"
    raw_csv = _optional("ODOO_ACCOUNTS")
    if raw_csv:
        for pair in raw_csv.split(","):
            pair = pair.strip()
            if not pair:
                continue
            if "=" in pair and ":" in pair:
                label, up = pair.split("=", 1)
                u, p = up.split(":", 1)
                label, u, p = label.strip(), u.strip(), p.strip()
                if label and u:
                    accounts[label] = {"username": u, "password": p}
            elif ":" in pair:
                u, p = pair.split(":", 1)
                u = u.strip()
                p = p.strip()
                if u and u not in accounts:
                    accounts[u] = {"username": u, "password": p}

    # 3. Environment prefix variables: ODOO_ACCOUNT_*_PASS / ODOO_USER_*_PASS
    for k, v in os.environ.items():
        if (k.startswith("ODOO_ACCOUNT_") or k.startswith("ODOO_USER_")) and k.endswith("_PASS"):
            prefix = k[:-5]
            label = os.environ.get(f"{prefix}_LABEL", os.environ.get(f"{prefix}_NAME", prefix.split("_", 2)[-1].lower().capitalize()))
            user = os.environ.get(f"{prefix}_USER", os.environ.get(f"{prefix}_NAME", label.lower()))
            if label not in accounts:
                accounts[label] = {"username": user, "password": v.strip()}

    return accounts

PREDEFINED_ACCOUNTS = get_predefined_accounts()

# ── Supabase Cloud Sync ───────────────────────────────────────────────────────
SUPABASE_URL = _optional("SUPABASE_URL")
SUPABASE_KEY = _optional("SUPABASE_KEY")
