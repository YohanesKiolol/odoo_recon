"""
config.py — loads all settings from .env
Script never hardcodes any value; everything is user-configurable.
"""

import os
from pathlib import Path
from dotenv import load_dotenv

load_dotenv()


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
