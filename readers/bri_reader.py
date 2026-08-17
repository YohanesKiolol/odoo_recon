"""
readers/bri_reader.py

Reads BRI transactions from a ZIP file (no password).

ZIP discovery:
  - Searches BRI_ZIP_DIR for a ZIP file whose name contains BRI_ZIP_PATTERN
  - Default pattern: '_franswega7' (e.g. 20260720_franswega7_xxxx.zip)

ZIP structure:
  <bri_zip>/
  └── <some_folder>/
      ├── detail_****.pdf   ← matched by BRI_PDF_PATTERN (default: 'detail_')
      └── ******.pdf        ← ignored

PDF structure:
  - Has a table with column 'AMT_TRX'
  - Amount format: comma as thousand separator, NO .00 at the back
    (e.g. "1,500,000")
"""

import zipfile
import io
import re
from pathlib import Path
from decimal import Decimal

import pdfplumber

from amount_utils import parse_amount, normalize_for_compare


def _find_bri_zips(zip_dir: Path, zip_pattern: str) -> list[Path]:
    """
    Find ALL BRI ZIP files in zip_dir.

    Primary:  filename contains zip_pattern (fast, zero I/O).
    Fallback: if no filename matches, return all .zip files in zip_dir —
              since ZIPs are now stored in alias-specific subdirectories,
              the directory itself is the alias filter. This handles renamed ZIPs.
    """
    candidates = sorted(
        p for p in zip_dir.iterdir()
        if p.suffix.lower() == ".zip" and zip_pattern.lower() in p.name.lower()
    )
    if candidates:
        return candidates

    # Fallback: return all ZIPs in the alias-specific directory
    all_zips = sorted(p for p in zip_dir.iterdir() if p.suffix.lower() == ".zip")
    if all_zips:
        return all_zips

    raise FileNotFoundError(
        f"No BRI ZIP files found in: {zip_dir}\n"
        f"Check BRI_ZIP_DIR in your .env file."
    )


def _extract_detail_pdf(zip_path: Path, pdf_pattern: str) -> bytes:
    """
    Open ZIP (no password), find the PDF whose filename starts with pdf_pattern.
    Returns the raw PDF bytes.
    """
    with zipfile.ZipFile(zip_path, "r") as zf:
        all_names = zf.namelist()
        # Match filename (basename only) starting with pdf_pattern, case-insensitive
        detail_files = [
            n for n in all_names
            if Path(n).suffix.lower() == ".pdf"
            and Path(n).name.lower().startswith(pdf_pattern.lower())
        ]

        if not detail_files:
            raise FileNotFoundError(
                f"No PDF starting with '{pdf_pattern}' found inside {zip_path.name}.\n"
                f"Contents: {all_names}\n"
                f"Check BRI_PDF_PATTERN in your .env file."
            )

        if len(detail_files) > 1:
            print(f"  [WARN] Multiple matching PDFs found: {detail_files}")
            print(f"  Using: {detail_files[0]}")

        print(f"  PDF found: {detail_files[0]}")
        return zf.read(detail_files[0])

def _normalize_pdf_header(h: str) -> str:
    """
    Fix character-doubling in BRI PDF headers caused by pdfplumber reading bold
    glyphs twice: 'AAMMTT__TTRRXX' → 'AMT_TRX'.
    Strategy: if every consecutive pair of chars is identical, halve the string.
    Also strips newlines and extra whitespace.
    """
    if not h:
        return h
    # Normalize newlines and strip
    h = h.replace("\n", " ").replace("\r", "").strip()
    # Check if the string is fully doubled (every pair is same char)
    if len(h) >= 2 and len(h) % 2 == 0:
        halved = "".join(h[i] for i in range(0, len(h), 2))
        doubled_back = "".join(c * 2 for c in halved)
        if doubled_back == h:
            return halved
    return h


def _parse_any_date(s: str):
    """Parse a date string in various formats → datetime.date or None."""
    from datetime import datetime
    if not s or not s.strip():
        return None
    for fmt in ("%d-%m-%Y", "%Y-%m-%d", "%d/%m/%Y", "%Y/%m/%d", "%d %b %Y", "%d %B %Y"):
        try:
            return datetime.strptime(s.strip(), fmt).date()
        except ValueError:
            continue
    return None


def _parse_pdf_table(
    pdf_bytes: bytes,
    amount_col: str,
    number_col: str = "",
    filter_dates: set | None = None,
    source_file: str = "",   # filename tag for tracing
) -> list[dict]:
    """
    Extract table rows from PDF, find the AMT_TRX column.
    Returns list of transaction dicts.
    """
    amount_col = amount_col.strip()
    number_col = number_col.strip()
    txns = []
    skipped = 0

    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page_num, page in enumerate(pdf.pages, start=1):
            tables = page.extract_tables()
            if not tables:
                continue

            for table in tables:
                if not table or len(table) < 2:
                    continue

                # First row of table assumed to be header
                raw_header = table[0]
                headers = [
                    _normalize_pdf_header(h) if h is not None else ""
                    for h in raw_header
                ]

                # Find amount column (case-insensitive)
                amount_idx = None
                for i, h in enumerate(headers):
                    if h.upper() == amount_col.upper():
                        amount_idx = i
                        break

                if amount_idx is None:
                    # This table doesn't have our column — skip it
                    continue

                # Try to find a date/description/number column
                date_idx = None
                desc_idx = None
                number_idx = None
                payment_date_idx = None
                admin_fee_idx = None
                for i, h in enumerate(headers):
                    hu = h.upper()
                    if date_idx is None and "TGL" in hu and "TRX" in hu:
                        date_idx = i
                    if desc_idx is None and "REMARK" in hu and "RK" in hu:
                        desc_idx = i
                    if payment_date_idx is None and "TGL" in hu and "RK" in hu and "REMARK" not in hu:
                        payment_date_idx = i
                    if admin_fee_idx is None and "DISC" in hu and "AMT" in hu:
                        admin_fee_idx = i
                    if number_idx is None and ("AAPPRR" in hu or "RREEFF" in hu):
                        number_idx = i

                # Parse data rows
                for row_num, row in enumerate(table[1:], start=2):
                    if not row or len(row) <= amount_idx:
                        continue

                    raw_amount = row[amount_idx]
                    if not raw_amount or (isinstance(raw_amount, str) and raw_amount.strip() in ("", "-", "0")):
                        skipped += 1
                        continue

                    try:
                        amount = normalize_for_compare(parse_amount(raw_amount))
                    except ValueError as e:
                        print(f"  [WARN] BRI PDF page {page_num} row {row_num}: {e} — skipped")
                        skipped += 1
                        continue

                    date_val   = str(row[date_idx]).strip()   if date_idx   is not None and row[date_idx]   else ""
                    desc_val   = str(row[desc_idx]).strip()   if desc_idx   is not None and row[desc_idx]   else ""
                    number_val = str(row[number_idx]).strip() if number_idx is not None and row[number_idx] else ""
                    pay_val    = str(row[payment_date_idx]).strip() if payment_date_idx is not None and row[payment_date_idx] else ""

                    # Date filter
                    txn_date = _parse_any_date(date_val)
                    if filter_dates is not None:
                        if txn_date not in filter_dates:
                            skipped += 1
                            continue
                            
                    final_date = str(txn_date) if txn_date else date_val
                    
                    # Parse Payment Date
                    pay_date = _parse_any_date(pay_val)
                    
                    # Fallback to H+1 if payment date is missing
                    if not pay_date and txn_date:
                        from datetime import date as dt_date, timedelta
                        if isinstance(txn_date, dt_date):
                            pay_date = txn_date + timedelta(days=1)
                            
                    final_pay_date = str(pay_date) if pay_date else pay_val
                    
                    # Admin fee
                    admin_fee_val = str(row[admin_fee_idx]).strip() if admin_fee_idx is not None and len(row) > admin_fee_idx and row[admin_fee_idx] else ""
                    try:
                        admin_fee = parse_amount(admin_fee_val) if admin_fee_val else Decimal("0")
                    except:
                        admin_fee = Decimal("0")

                    txns.append({
                        "amount":      amount,
                        "amount_raw":  raw_amount,
                        "admin_fee":   admin_fee,
                        "date":        final_date,
                        "payment_date": final_pay_date,
                        "description": desc_val,
                        "number":      number_val,
                        "is_void":     "VOID" in desc_val.upper(),
                        "is_reversal": "REVERSAL" in desc_val.upper(),
                        "is_refund":   "REFUND" in desc_val.upper(),
                        "filename":    source_file,
                        "source":      "Bank (BRI)",
                    })

    if skipped:
        print(f"  [INFO] BRI PDF: {skipped} empty/zero rows skipped")

    return txns


def read_bri(
    zip_dir: Path,
    zip_pattern: str,
    pdf_pattern: str,
    amount_col: str,
    number_col: str = "",
    filter_dates: set | None = None,
) -> tuple[list[dict], list[dict]]:
    """
    Read BRI transactions from ALL matching ZIPs in zip_dir.
    If filter_dates is provided, only rows matching those dates are returned.
    """
    if not zip_dir.exists():
        raise FileNotFoundError(
            f"BRI ZIP directory not found: {zip_dir}\n"
            f"Check BRI_ZIP_DIR in your .env file."
        )

    zip_paths = _find_bri_zips(zip_dir, zip_pattern)
    print(f"  Found {len(zip_paths)} BRI ZIP(s): {[p.name for p in zip_paths]}")

    def _process_one_bri_zip(zip_path):
        print(f"  BRI ZIP: {zip_path.name}")
        pdf_bytes = _extract_detail_pdf(zip_path, pdf_pattern)
        print(f"  PDF extracted ({len(pdf_bytes):,} bytes), parsing table...")
        return _parse_pdf_table(pdf_bytes, amount_col, number_col,
                                filter_dates=filter_dates,
                                source_file=zip_path.name)

    all_txns: list[dict] = []
    if len(zip_paths) <= 1:
        for zip_path in zip_paths:
            all_txns.extend(_process_one_bri_zip(zip_path))
    else:
        from concurrent.futures import ThreadPoolExecutor
        with ThreadPoolExecutor(max_workers=min(len(zip_paths), 4), thread_name_prefix="bri_zip") as pool:
            futs = [pool.submit(_process_one_bri_zip, zp) for zp in zip_paths]
            for fut in futs:
                all_txns.extend(fut.result())

    # ── Exclude Voided/Reversed Transactions ──────────────────────────────────
    voided_keys = set()
    for t in all_txns:
        if t.get("is_void") or t.get("is_reversal"):
            num = (t.get("number") or "").strip()
            if num:
                voided_keys.add((t["date"], num))

    excluded_txns = []
    if voided_keys:
        filtered_txns = []
        for t in all_txns:
            num = (t.get("number") or "").strip()
            if num and (t["date"], num) in voided_keys:
                if t.get("is_void"):
                    t["exclusion_reason"] = "Void"
                elif t.get("is_reversal"):
                    t["exclusion_reason"] = "Reversal"
                elif t.get("is_refund"):
                    t["exclusion_reason"] = "Refund"
                else:
                    t["exclusion_reason"] = "Original transaction of Void/Reversal"
                excluded_txns.append(t)
                continue
            filtered_txns.append(t)
        
        print(f"  [INFO] BRI: Excluded {len(excluded_txns)} rows due to Void/Reversal found in description")
        all_txns = filtered_txns

    return all_txns, excluded_txns
