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

def _find_bca_excels(excel_dir: Path, excel_pattern: str) -> list[Path]:
    """
    Find ALL BCA Excel files in excel_dir whose name contains excel_pattern.
    Returns sorted list (oldest first by filename).
    """
    candidates = sorted(
        p for p in excel_dir.iterdir()
        if p.suffix.lower() in (".xlsx", ".xls")
        and excel_pattern.lower() in p.name.lower()
        and not p.name.startswith("~$")   # skip Excel temp/lock files
    )
    if not candidates:
        all_xlsx = [p.name for p in excel_dir.glob("*.xlsx")] + [p.name for p in excel_dir.glob("*.xls")]
        raise FileNotFoundError(
            f"No BCA Excel containing '{excel_pattern}' found in: {excel_dir}\n"
            f"Available Excel files: {all_xlsx}\n"
            f"Check BCA_EXCEL_PATTERN in your .env file."
        )
    return candidates


def _read_one_bca(
    excel_path: Path,
    password: str,
    amount_col: str,
    date_col: str,
    number_col: str = "",
    filter_dates: set | None = None,   # if provided, only keep rows with matching date
) -> list[dict]:
    """Read transactions from a single BCA Excel file.
    If filter_dates is given, only rows whose date is in that set are returned.
    """
    print(f"  BCA file: {excel_path.name}")
    print(f"  Decrypting {excel_path.name}...")
    decrypted = io.BytesIO()
    try:
        with open(excel_path, "rb") as f:
            office_file = msoffcrypto.OfficeFile(f)
            office_file.load_key(password=password)
            office_file.decrypt(decrypted)
    except Exception as e:
        raise RuntimeError(
            f"Cannot decrypt BCA Excel '{excel_path.name}': {e}\n"
            f"Check BCA_EXCEL_PASSWORD in your .env file."
        )

    decrypted.seek(0)
    wb = openpyxl.load_workbook(decrypted, data_only=True)
    ws = wb.active
    assert ws is not None

    header_row = [cell.value for cell in ws[5]]
    headers = [str(h).strip() if h is not None else "" for h in header_row]

    def _find_col(col_name: str) -> int:
        col_name = col_name.strip()
        if col_name in headers:
            return headers.index(col_name)
        for i, h in enumerate(headers):
            if h.lower() == col_name.lower():
                return i
        raise ValueError(
            f"Column '{col_name}' not found in BCA Excel '{excel_path.name}'.\n"
            f"Available columns: {[h for h in headers if h]}"
        )

    amount_idx = _find_col(amount_col)
    date_idx   = _find_col(date_col)

    number_idx: int | None = None
    if number_col:
        try:
            number_idx = _find_col(number_col.strip())
        except ValueError:
            print(f"  [WARN] BCA: number column '{number_col}' not found — skipping")

    txns = []
    skipped_empty = 0

    for row_num, row in enumerate(ws.iter_rows(min_row=6, values_only=True), start=6):
        if not any(c is not None for c in row):
            continue

        raw_amount = row[amount_idx] if amount_idx < len(row) else None
        if raw_amount is None or str(raw_amount).strip() in ("", "0", "0.00"):
            skipped_empty += 1
            continue

        raw_date = row[date_idx] if date_idx < len(row) else None
        txn_date = _parse_bca_date(raw_date)

        if txn_date is None:
            skipped_empty += 1
            continue

        # Date filter: skip rows whose date is not in the allowed set
        if filter_dates is not None and txn_date not in filter_dates:
            continue

        try:
            amount = normalize_for_compare(parse_amount(raw_amount))
        except ValueError as e:
            print(f"  [WARN] BCA row {row_num}: {e} — skipped")
            skipped_empty += 1
            continue

        from datetime import date as dt_date, timedelta
        if isinstance(txn_date, dt_date):
            payment_date = txn_date + timedelta(days=1)
        else:
            payment_date = ""

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
            "date":        str(txn_date) if txn_date else str(raw_date or ""),
            "payment_date": str(payment_date) if payment_date else "",
            "description": desc,
            "number":      number,
            "filename":    excel_path.name,
            "source":      "Bank (BCA)",
        })

    print(f"    → {len(txns)} transactions ({skipped_empty} empty skipped)")
    return txns


def read_bca(
    excel_dir: Path,
    excel_pattern: str,
    password: str,
    amount_col: str,
    date_col: str,
    number_col: str = "",
    filter_date=None,          # single date (legacy, converted to set)
    filter_dates: set | None = None,  # set of dates from ODO — primary filter
) -> list[dict]:
    """
    Read BCA transactions from ALL matching Excel files in excel_dir.

    Date filtering:
      - filter_dates (set[date]): only keep rows matching any date in the set.
        Derived from ODO BCA transactions → preserves original single-date
        behavior when ODO has 1 date, and extends to multi-date automatically.
      - filter_date (date, legacy): converted to a one-element set.
      - If neither is provided: all dates are included (no filter).
    """
    # Backwards-compat: convert legacy single filter_date to set
    if filter_dates is None and filter_date is not None:
        filter_dates = {filter_date}

    if not excel_dir.exists():
        raise FileNotFoundError(
            f"BCA Excel directory not found: {excel_dir}\n"
            f"Check BCA_EXCEL_DIR in your .env file."
        )

    excel_files = _find_bca_excels(excel_dir, excel_pattern)
    print(f"  Found {len(excel_files)} BCA file(s): {[p.name for p in excel_files]}")

    all_txns: list[dict] = []
    for path in excel_files:
        all_txns.extend(
            _read_one_bca(path, password, amount_col, date_col, number_col,
                          filter_dates=filter_dates)
        )

    # ── Deduplicate overlapping files ─────────────────────────────────────────
    # Key: (date, trace_number) if trace_number exists, else (date, amount, desc)
    seen: set = set()
    deduped: list[dict] = []
    for t in all_txns:
        num = (t.get("number") or "").strip()
        if num:
            key = (t["date"], num)
        else:
            key = (t["date"], str(t["amount"]), t.get("description", ""))
        if key not in seen:
            seen.add(key)
            deduped.append(t)

    dupes = len(all_txns) - len(deduped)
    if dupes:
        print(f"  [INFO] BCA: {dupes} duplicate row(s) removed (overlapping files)")
    all_txns = deduped

    dates = sorted({t["date"] for t in all_txns if t["date"]})
    date_summary = (dates[0] if len(dates) == 1
                    else f"{dates[0]} – {dates[-1]}" if dates else "?")
    print(f"  ✅ BCA: {len(all_txns)} transactions loaded "
          f"({len(excel_files)} file(s), dates: {date_summary})")
    return all_txns
