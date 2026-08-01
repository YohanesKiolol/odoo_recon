"""
config.py — loads all settings from .env
Script never hardcodes any value; everything is user-configurable.
"""

import os
import sys
from pathlib import Path

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
ODOO_JOURNAL_CREATE_URL = _optional("ODOO_JOURNAL_CREATE_URL")
ODOO_PAYMENTS_URL       = _optional("ODOO_PAYMENTS_URL")

ODO_EXCEL_PATH      = Path(_require("ODO_EXCEL_PATH"))
ODO_AMOUNT_COLUMN   = _require("ODO_AMOUNT_COLUMN")
ODO_NUMBER_COLUMN   = _optional("ODO_NUMBER_COLUMN", "Number")
ODO_GROUP_BCA       = _require("ODO_GROUP_BCA")
ODO_GROUP_MANDIRI   = _require("ODO_GROUP_MANDIRI")
ODO_GROUP_BRI       = _require("ODO_GROUP_BRI")

# ── Mandiri ───────────────────────────────────────────────────────────────────
MANDIRI_ZIP_DIR       = Path(_require("MANDIRI_ZIP_DIR"))
MANDIRI_ZIP_PATTERN   = _optional("MANDIRI_ZIP_PATTERN", "MSR_*.zip")
MANDIRI_ZIP_PASSWORD  = _require("MANDIRI_ZIP_PASSWORD")
MANDIRI_AMOUNT_COLUMN = _require("MANDIRI_AMOUNT_COLUMN")
MANDIRI_NUMBER_COLUMN = _optional("MANDIRI_NUMBER_COLUMN", "AUTHCODE")

# ── BCA ───────────────────────────────────────────────────────────────────────
BCA_EXCEL_DIR      = Path(_require("BCA_EXCEL_DIR"))
BCA_EXCEL_PATTERN  = _require("BCA_EXCEL_PATTERN")
BCA_EXCEL_PASSWORD = _require("BCA_EXCEL_PASSWORD")
BCA_AMOUNT_COLUMN  = _require("BCA_AMOUNT_COLUMN")
BCA_DATE_COLUMN    = _require("BCA_DATE_COLUMN")
BCA_NUMBER_COLUMN  = _optional("BCA_NUMBER_COLUMN", "Trace Number")

# ── BRI ───────────────────────────────────────────────────────────────────────
BRI_ZIP_DIR       = Path(_require("BRI_ZIP_DIR"))
BRI_ZIP_PATTERN   = _require("BRI_ZIP_PATTERN")
BRI_PDF_PATTERN   = _optional("BRI_PDF_PATTERN", "detail_")
BRI_AMOUNT_COLUMN = _require("BRI_AMOUNT_COLUMN")
BRI_NUMBER_COLUMN = _optional("BRI_NUMBER_COLUMN", "REMARK_RK")

# ── Output ────────────────────────────────────────────────────────────────────
OUTPUT_DIR = Path(_optional("OUTPUT_DIR", "output"))
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
