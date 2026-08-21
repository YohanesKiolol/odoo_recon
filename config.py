"""
config.py — loads all settings from company .env files with multi-company support.
"""

import os
import sys
import json
import warnings
from pathlib import Path

# Suppress openpyxl stylesheet warnings globally
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

# Base directory determination (supports PyInstaller frozen mode)
if getattr(sys, "frozen", False):
    if sys.platform == "darwin" and ".app/Contents/MacOS" in str(sys.executable):
        _base = Path(sys.executable).parents[2].parent
    else:
        _base = Path(sys.executable).parent
else:
    _base = Path(__file__).resolve().parent

BASE_DIR = _base
LAST_COMPANY_FILE = _base / ".last_company.json"


def _load_dotenv(dotenv_path: Path):
    """Load key-value pairs from a .env file into os.environ."""
    if not dotenv_path.exists():
        return
    try:
        content = None
        for enc in ("utf-8-sig", "utf-8", "cp1252", "latin-1"):
            try:
                with open(dotenv_path, "r", encoding=enc) as f:
                    content = f.readlines()
                break
            except (UnicodeDecodeError, LookupError, OSError):
                continue

        if content is None:
            with open(dotenv_path, "r", encoding="utf-8", errors="replace") as f:
                content = f.readlines()

        for line in content:
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
    except Exception:
        pass


def _require(key: str, default: str = "") -> str:
    val = os.getenv(key, default)
    if not val:
        raise EnvironmentError(
            f"Missing required config: '{key}'\n"
            f"Check your company .env file and fill in the value."
        )
    return val.strip()


def _optional(key: str, default: str = "") -> str:
    return (os.getenv(key) or default).strip()


def get_available_companies() -> dict[str, str]:
    """
    Scan root directory for company environment files (.env.<company_key> or .env).
    Returns dict of {company_key: display_name}.
    """
    companies = {}
    # Scan for .env.* files
    for env_file in sorted(_base.glob(".env.*")):
        if env_file.name in [".env.example", ".env.template", ".env.bak", ".env.tmp"]:
            continue
        key = env_file.name.split(".", 2)[-1].strip()
        if not key:
            continue
        # Extract COMPANY_NAME from file
        comp_name = key.replace("_", " ").title()
        try:
            with open(env_file, "r", encoding="utf-8", errors="ignore") as f:
                for line in f:
                    if line.strip().startswith("COMPANY_NAME="):
                        comp_name = line.strip().split("=", 1)[1].strip().strip('"\'')
                        break
        except Exception:
            pass
        companies[key] = comp_name

    # If no .env.* found, check standard .env
    if not companies and (_base / ".env").exists():
        comp_name = "Default Company"
        try:
            with open(_base / ".env", "r", encoding="utf-8", errors="ignore") as f:
                for line in f:
                    if line.strip().startswith("COMPANY_NAME="):
                        comp_name = line.strip().split("=", 1)[1].strip().strip('"\'')
                        break
        except Exception:
            pass
        companies["default"] = comp_name

    return companies


def get_last_selected_company() -> str:
    """Read last selected company key from local settings."""
    if LAST_COMPANY_FILE.exists():
        try:
            data = json.loads(LAST_COMPANY_FILE.read_text(encoding="utf-8"))
            key = data.get("company_key", "")
            if key:
                return key
        except Exception:
            pass
    avail = get_available_companies()
    if "eyerizz" in avail:
        return "eyerizz"
    return next(iter(avail), "default")


def save_selected_company(company_key: str):
    """Persist selected company key to .last_company.json."""
    try:
        LAST_COMPANY_FILE.write_text(json.dumps({"company_key": company_key}, indent=2), encoding="utf-8")
    except Exception:
        pass


# Global configuration state holders
COMPANY_NAME = "Eyerizz Eyewear"
COMPANY_KEY = "eyerizz"

ODOO_URL = ""
ODOO_DB = ""
ODOO_API_KEY = ""
ODOO_DASHBOARD_URL = ""
ODOO_PAYMENTS_URL = ""
ODOO_JOURNAL_IMPORT_URL = ""
ODOO_JOURNAL_ENTRIES_URL = ""
ODOO_COMPANY_NAME = ""
ODOO_JOURNAL_EDC = "EDC Settlement Journal"
ODOO_JOURNAL_AR = "MDR Fees Journal"
ODOO_ACCOUNT_BANK_DIFF_INCOME = "82005 Bank Difference Income"
ODOO_ACCOUNT_BANK_DIFF_LOSS = "8107 Bank Difference Loss"
JOURNAL_TOLERANCE = 200000.0

ODO_JOURNAL_EXCEL_PATH = _base / "input/Journal Entries (account.move).xlsx"
ODO_EXCEL_PATH = _base / "input/Payments (account.payment).xlsx"
ODO_AMOUNT_COLUMN = "Amount Signed"
ODO_NUMBER_COLUMN = "Number"
ODO_GROUP_BCA = "BCA EDC Sanur"
ODO_GROUP_MANDIRI = "Mandiri EDC"
ODO_GROUP_BRI = ""

JOURNAL_TEMPLATE_DIR = _base / "template"
JOURNAL_EXCEL_PATTERN = "Journal Entry (account.move).xlsx"

# Bank directories
MANDIRI_ZIP_DIR = _base / "input/mandiri"
MANDIRI_ZIP_PATTERN = ""
MANDIRI_MUTATION_PATTERN = ""
MANDIRI_ZIP_PASSWORD = ""
MANDIRI_AMOUNT_COLUMN = "AMOUNT"
MANDIRI_NUMBER_COLUMN = "AUTHCODE"

BCA_EXCEL_DIR = _base / "input/bca"
BCA_EXCEL_PATTERN = ""
BCA_MUTATION_PATTERN = ""
BCA_EXCEL_PASSWORD = ""
BCA_AMOUNT_COLUMN = "Original Amount"
BCA_DATE_COLUMN = "Transaction Date"
BCA_NUMBER_COLUMN = "Trace Number"

BRI_ZIP_DIR = _base / "input/bri"
BRI_ZIP_PATTERN = ""
BRI_MUTATION_PATTERN = ""
BRI_PDF_PATTERN = ""
BRI_AMOUNT_COLUMN = "AMT_TRX"
BRI_NUMBER_COLUMN = "REMARK_RK"

INPUT_DIR = _base / "input"
MUTATION_DIR = _base / "mutation"
OUTPUT_DIR = _base / "output"
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

SUPABASE_URL = ""
SUPABASE_KEY = ""
SERVICE_ROLE = ""
SUPABASE_ANON_KEY = ""
ENCRYPTION_KEY = ""

BANK_ACCOUNTS = {
    "bca": {},
    "mandiri": {},
    "bri": {}
}
PREDEFINED_ACCOUNTS = {}


def load_company_env(company_key: str = None) -> str:
    """Load configuration for the specified company key."""
    global COMPANY_NAME, COMPANY_KEY, ODOO_URL, ODOO_DB, ODOO_API_KEY
    global ODOO_DASHBOARD_URL, ODOO_PAYMENTS_URL, ODOO_JOURNAL_IMPORT_URL, ODOO_JOURNAL_ENTRIES_URL
    global ODOO_COMPANY_NAME, ODOO_JOURNAL_EDC, ODOO_JOURNAL_AR
    global ODOO_ACCOUNT_BANK_DIFF_INCOME, ODOO_ACCOUNT_BANK_DIFF_LOSS, JOURNAL_TOLERANCE
    global ODO_JOURNAL_EXCEL_PATH, ODO_EXCEL_PATH, ODO_AMOUNT_COLUMN, ODO_NUMBER_COLUMN
    global ODO_GROUP_BCA, ODO_GROUP_MANDIRI, ODO_GROUP_BRI
    global MANDIRI_ZIP_PASSWORD, BCA_EXCEL_PASSWORD
    global SUPABASE_URL, SUPABASE_KEY, SERVICE_ROLE, SUPABASE_ANON_KEY, ENCRYPTION_KEY, BANK_ACCOUNTS, PREDEFINED_ACCOUNTS

    if not company_key:
        company_key = get_last_selected_company()

    # Locate env file: prefer .env.<company_key>, fallback to .env
    target_env = _base / f".env.{company_key}"
    if not target_env.exists():
        if (_base / ".env").exists():
            target_env = _base / ".env"
        elif (_base / ".env.eyerizz").exists():
            target_env = _base / ".env.eyerizz"
            company_key = "eyerizz"

    _load_dotenv(target_env)
    save_selected_company(company_key)

    COMPANY_KEY = company_key
    COMPANY_NAME = _optional("COMPANY_NAME", company_key.replace("_", " ").title())

    # Odoo Config
    ODOO_URL = _optional("ODOO_URL")
    ODOO_DB = _optional("ODOO_DB")
    ODOO_API_KEY = _optional("ODOO_API_KEY")
    ODOO_DASHBOARD_URL = _optional("ODOO_DASHBOARD_URL")
    ODOO_PAYMENTS_URL = _optional("ODOO_PAYMENTS_URL")
    ODOO_JOURNAL_IMPORT_URL = _optional("ODOO_JOURNAL_IMPORT_URL")
    ODOO_JOURNAL_ENTRIES_URL = _optional("ODOO_JOURNAL_ENTRIES_URL")
    ODOO_COMPANY_NAME = _optional("ODOO_COMPANY_NAME", COMPANY_NAME)
    ODOO_JOURNAL_EDC = _optional("ODOO_JOURNAL_EDC", "EDC Settlement Journal")
    ODOO_JOURNAL_AR = _optional("ODOO_JOURNAL_AR", "MDR Fees Journal")
    ODOO_ACCOUNT_BANK_DIFF_INCOME = _optional("ODOO_ACCOUNT_BANK_DIFF_INCOME", "82005 Bank Difference Income")
    ODOO_ACCOUNT_BANK_DIFF_LOSS = _optional("ODOO_ACCOUNT_BANK_DIFF_LOSS", "8107 Bank Difference Loss")

    # Reconciliation Tolerance (User-Configurable in .env, defaults to 200,000)
    try:
        JOURNAL_TOLERANCE = float(_optional("JOURNAL_TOLERANCE", "200000"))
    except Exception:
        JOURNAL_TOLERANCE = 200000.0

    ODO_JOURNAL_EXCEL_PATH = _base / _optional("ODO_JOURNAL_EXCEL_PATH", "input/Journal Entries (account.move).xlsx")
    ODO_EXCEL_PATH = _base / _optional("ODO_EXCEL_PATH", "input/Payments (account.payment).xlsx")
    ODO_AMOUNT_COLUMN = _optional("ODO_AMOUNT_COLUMN", "Amount Signed")
    ODO_NUMBER_COLUMN = _optional("ODO_NUMBER_COLUMN", "Number")
    ODO_GROUP_BCA = _optional("ODO_GROUP_BCA", "BCA EDC Sanur")
    ODO_GROUP_MANDIRI = _optional("ODO_GROUP_MANDIRI", "Mandiri EDC")
    ODO_GROUP_BRI = _optional("ODO_GROUP_BRI", "")

    # Bank Passwords
    MANDIRI_ZIP_PASSWORD = _optional("MANDIRI_ZIP_PASSWORD")
    BCA_EXCEL_PASSWORD = _optional("BCA_EXCEL_PASSWORD")

    # Supabase
    SUPABASE_URL = _optional("SUPABASE_URL")
    SERVICE_ROLE = _optional("SERVICE_ROLE")
    SUPABASE_ANON_KEY = _optional("SUPABASE_ANON_KEY")
    ENCRYPTION_KEY = _optional("ENCRYPTION_KEY")
    # Admin app uses SERVICE_ROLE; falls back to anon key (sales portal)
    SUPABASE_KEY = SERVICE_ROLE or SUPABASE_ANON_KEY or _optional("SUPABASE_KEY")

    # Dynamic Bank Accounts parsing from environment
    BANK_ACCOUNTS = {}
    for k, v in os.environ.items():
        if k.startswith("ACCOUNT_"):
            parts = k.split("_")
            if len(parts) >= 4:
                bank = parts[1].lower()
                if bank not in BANK_ACCOUNTS:
                    BANK_ACCOUNTS[bank] = {}
                alias = parts[2].lower()
                prop = "_".join(parts[3:]).lower()
                if alias not in BANK_ACCOUNTS[bank]:
                    BANK_ACCOUNTS[bank][alias] = {"mid": "", "group": "", "acc": "", "edc_debit": "", "edc_credit": "", "store": ""}
                BANK_ACCOUNTS[bank][alias][prop] = v

    # Fallbacks for legacy groups (only if explicitly defined in environment)
    for bank_key_s, legacy_group in [("bca", ODO_GROUP_BCA), ("mandiri", ODO_GROUP_MANDIRI), ("bri", ODO_GROUP_BRI)]:
        if legacy_group and bank_key_s in BANK_ACCOUNTS:
            if "main" not in BANK_ACCOUNTS[bank_key_s]:
                BANK_ACCOUNTS[bank_key_s]["main"] = {"mid": "", "group": legacy_group, "acc": "", "edc_debit": "", "edc_credit": "", "store": ""}
            elif not BANK_ACCOUNTS[bank_key_s]["main"].get("group"):
                BANK_ACCOUNTS[bank_key_s]["main"]["group"] = legacy_group

    PREDEFINED_ACCOUNTS = get_predefined_accounts()

    # Pre-calculate active configured banks & subtitles in memory (0ms lookup)
    global CONFIGURED_BANKS, BANK_SUBTITLES
    configured = []
    subtitles = {"All": "All active"}
    for b_k, accs in BANK_ACCOUNTS.items():
        if accs:
            b_name = b_k.upper()
            configured.append(b_name)
            if list(accs.keys()) == ["main"]:
                info = accs["main"]
                types = []
                if any("creditcard" in k for k in info): types.append("Credit")
                if any("debitcard" in k for k in info):  types.append("Debit")
                if any("_qr" in k for k in info):        types.append("QR")
                subtitles[b_name] = " · ".join(types) if types else "EDC"
            else:
                parts = [a.upper() if len(a) <= 3 else a.capitalize() for a in accs.keys()]
                subtitles[b_name] = " · ".join(parts[:3])

    CONFIGURED_BANKS = configured
    BANK_SUBTITLES = subtitles
    return company_key



def get_journal_store(journal_name: str) -> str:
    """Determine the physical store location ('Sanur' or 'Seminyak') for an Odoo journal."""
    j_raw = str(journal_name or "").strip()
    j_upper = j_raw.upper()
    if not j_upper:
        return ""

    for bank_k, accounts in BANK_ACCOUNTS.items():
        for alias_k, acc_data in accounts.items():
            grp = str(acc_data.get("group") or "").strip().upper()
            st = str(acc_data.get("store") or "").strip()
            if grp and (grp == j_upper or grp in j_upper or j_upper in grp):
                if st:
                    return st

    if "SANUR" in j_upper or "LBF" in j_upper:
        return "Sanur"
    if "SEMINYAK" in j_upper or "NARA" in j_upper:
        return "Seminyak"
    if "MANDIRI" in j_upper:
        return "Sanur"

    return ""


def get_predefined_accounts() -> dict[str, dict[str, str]]:
    """Parse predefined Odoo user accounts from .env (supports API keys and passwords)."""
    accounts: dict[str, dict[str, str]] = {}

    raw_json = _optional("ODOO_PREDEFINED_ACCOUNTS")
    if raw_json:
        try:
            data = json.loads(raw_json)
            if isinstance(data, dict):
                for k, v in data.items():
                    if isinstance(v, dict):
                        key_val = str(v.get("api_key", v.get("key", v.get("password", v.get("pass", ""))))).strip()
                        accounts[k] = {
                            "username": str(v.get("username", v.get("user", k))).strip(),
                            "api_key": key_val,
                            "password": key_val,
                        }
                    elif isinstance(v, str):
                        accounts[k] = {
                            "username": k,
                            "api_key": v.strip(),
                            "password": v.strip(),
                        }
        except Exception:
            pass

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
                    accounts[label] = {"username": u, "api_key": p, "password": p}
            elif ":" in pair:
                u, p = pair.split(":", 1)
                u = u.strip()
                p = p.strip()
                if u and u not in accounts:
                    accounts[u] = {"username": u, "api_key": p, "password": p}

    for k, v in os.environ.items():
        if (k.startswith("ODOO_ACCOUNT_") or k.startswith("ODOO_USER_")) and (k.endswith("_KEY") or k.endswith("_API_KEY") or k.endswith("_PASS")):
            if k.endswith("_API_KEY"):
                prefix = k[:-8]
            elif k.endswith("_KEY"):
                prefix = k[:-4]
            else:
                prefix = k[:-5]
            label = os.environ.get(f"{prefix}_LABEL", os.environ.get(f"{prefix}_NAME", prefix.split("_", 2)[-1].lower().capitalize()))
            user = os.environ.get(f"{prefix}_USER", os.environ.get(f"{prefix}_NAME", label.lower()))
            if label not in accounts:
                accounts[label] = {"username": user, "api_key": v.strip(), "password": v.strip()}

    single_user = _optional("ODOO_USER")
    single_key = _optional("ODOO_API_KEY")
    if single_key and ("," in single_key or (":" in single_key and not single_user)):
        for item in single_key.split(","):
            item = item.strip()
            if not item:
                continue
            if "=" in item:
                u, k = item.split("=", 1)
            elif ":" in item:
                u, k = item.split(":", 1)
            else:
                continue
            u, k = u.strip(), k.strip()
            if u and k:
                accounts[u.capitalize()] = {"username": u, "api_key": k, "password": k}
    elif single_user and single_key and not accounts:
        accounts[single_user.capitalize()] = {"username": single_user, "api_key": single_key, "password": single_key}

    return accounts



# Initialize default company environment on module import
load_company_env()
