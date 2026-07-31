"""
readers/bca_reader.py

Reads BCA bank transactions from a password-protected Excel file.

File discovery:
  - Searches BCA_EXCEL_DIR for an Excel file whose name contains BCA_EXCEL_PATTERN
  - Default pattern: 'ReportMerchantBCA_' (e.g. ReportMerchantBCA_20260720.xlsx)

Excel structure:
  - Password-protected (msoffcrypto-tool to decrypt)
  - Header at row 5 (rows 1-4 are preamble)
  - Column: 'Original Amount' — no thousand separator, ends with .00
  - Column: 'Transaction Date' — format DD/MM/YYYY (e.g. "17/06/2026")
  - Filter: only rows where Transaction Date matches the ODO date
"""

import io
from datetime import date
from pathlib import Path
from decimal import Decimal

import msoffcrypto
import openpyxl

from amount_utils import parse_amount, normalize_for_compare


def _parse_bca_date(raw) -> date | None:
    """Parse BCA date cell: '17/06/2026' → date(2026, 6, 17)"""
    if raw is None:
        return None
    # If Excel stored as actual date object
    if isinstance(raw, date):
        return raw
    from datetime import datetime
    s = str(raw).strip()
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%d %b %Y"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None

def _find_bca_excel(excel_dir: Path, excel_pattern: str) -> Path:
    """
    Find the BCA Excel file in excel_dir whose name contains excel_pattern (substring).
    Accepts .xlsx and .xls extensions.
    """
    candidates = [
        p for p in excel_dir.iterdir()
        if p.suffix.lower() in (".xlsx", ".xls")
        and excel_pattern.lower() in p.name.lower()
    ]
    if not candidates:
        all_xlsx = [p.name for p in excel_dir.glob("*.xlsx")] + [p.name for p in excel_dir.glob("*.xls")]
        raise FileNotFoundError(
            f"No BCA Excel containing '{excel_pattern}' found in: {excel_dir}\n"
            f"Available Excel files: {all_xlsx}\n"
            f"Check BCA_EXCEL_PATTERN in your .env file."
        )
    if len(candidates) > 1:
        print(f"  [WARN] Multiple BCA files matched '{excel_pattern}': {[c.name for c in candidates]}")
        print(f"  Using: {candidates[0].name}")
    return candidates[0]


def read_bca(
    excel_dir: Path,
    excel_pattern: str,
    password: str,
    amount_col: str,
    date_col: str,
    filter_date: date,
    number_col: str = "",
) -> list[dict]:
    """
    Read BCA transactions for a specific date.

    Args:
        excel_dir     : folder to search for the BCA Excel file
        excel_pattern : substring to match BCA filename (e.g. 'ReportMerchantBCA_')
        password      : file password
        amount_col    : column name for amount (e.g. 'Original Amount')
        date_col      : column name for date (e.g. 'Transaction Date')
        filter_date   : only return rows matching this date (from ODO A3)

    Returns list of transaction dicts.
    """
    if not excel_dir.exists():
        raise FileNotFoundError(
            f"BCA Excel directory not found: {excel_dir}\n"
            f"Check BCA_EXCEL_DIR in your .env file."
        )

    excel_path = _find_bca_excel(excel_dir, excel_pattern)
    print(f"  BCA file: {excel_path.name}")

    # ── Decrypt ───────────────────────────────────────────────────────────────
    print(f"  Decrypting {excel_path.name}...")
    decrypted = io.BytesIO()
    try:
        with open(excel_path, "rb") as f:
            office_file = msoffcrypto.OfficeFile(f)
            office_file.load_key(password=password)
            office_file.decrypt(decrypted)
    except Exception as e:
        raise RuntimeError(
            f"Cannot decrypt BCA Excel: {e}\n"
            f"Check BCA_EXCEL_PASSWORD in your .env file."
        )

    decrypted.seek(0)
    wb = openpyxl.load_workbook(decrypted, data_only=True)
    ws = wb.active
    assert ws is not None, "No active worksheet in BCA Excel"

    # ── Header at row 5 (index 4) ─────────────────────────────────────────────
    header_row = [cell.value for cell in ws[5]]
    headers = [str(h).strip() if h is not None else "" for h in header_row]

    def _find_col(col_name: str) -> int:
        col_name = col_name.strip()
        # Exact match first
        if col_name in headers:
            return headers.index(col_name)
        # Case-insensitive
        for i, h in enumerate(headers):
            if h.lower() == col_name.lower():
                return i
        raise ValueError(
            f"Column '{col_name}' not found in BCA Excel.\n"
            f"Available columns: {[h for h in headers if h]}"
        )

    amount_idx = _find_col(amount_col)
    date_idx   = _find_col(date_col)

    # Optional number column — soft fail if not present
    number_idx: int | None = None
    if number_col:
        try:
            number_idx = _find_col(number_col.strip())
        except ValueError:
            print(f"  [WARN] BCA: number column '{number_col}' not found — skipping")

    txns = []
    skipped_date = 0
    skipped_empty = 0

    # Data starts at row 6 (header was row 5)
    for row_num, row in enumerate(ws.iter_rows(min_row=6, values_only=True), start=6):
        if not any(c is not None for c in row):
            continue

        raw_date   = row[date_idx]   if date_idx   < len(row) else None
        raw_amount = row[amount_idx] if amount_idx < len(row) else None

        # Skip empty amount
        if raw_amount is None or str(raw_amount).strip() in ("", "0", "0.00"):
            skipped_empty += 1
            continue

        # Filter by date
        txn_date = _parse_bca_date(raw_date)
        if txn_date != filter_date:
            skipped_date += 1
            continue

        try:
            amount = normalize_for_compare(parse_amount(raw_amount))
        except ValueError as e:
            print(f"  [WARN] BCA row {row_num}: {e} — skipped")
            skipped_empty += 1
            continue

        # Description: try common BCA column names
        desc = ""
        for desc_col_name in ("Transaction Remark", "Keterangan", "Description", "Remark"):
            try:
                di = _find_col(desc_col_name)
                desc = str(row[di]).strip() if row[di] is not None else ""
                break
            except (ValueError, IndexError):
                continue

        number = ""
        if number_idx is not None and number_idx < len(row):
            number = str(row[number_idx]).strip() if row[number_idx] is not None else ""

        txns.append({
            "amount":      amount,
            "amount_raw":  raw_amount,
            "date":        str(txn_date),
            "description": desc,
            "number":      number,
            "source":      "Bank (BCA)",
        })

    print(f"  BCA: {len(txns)} rows matched date {filter_date}, "
          f"{skipped_date} other-date rows excluded, {skipped_empty} empty skipped")

    return txns
