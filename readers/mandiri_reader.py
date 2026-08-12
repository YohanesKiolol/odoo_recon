"""
readers/mandiri_reader.py

Reads Mandiri bank transactions from 2 password-protected ZIP files.

ZIP file naming: MSR_****.zip
ZIP contents: one or more CSV files
CSV structure: header at row 6 (rows 1-5 are preamble), column 'AMOUNT'
Amount format: comma as thousand separator, ends with .00 (e.g. "1,500,000.00")
"""

import io
import csv
from pathlib import Path
from decimal import Decimal

import pyzipper

from amount_utils import parse_amount, normalize_for_compare


def _parse_any_date(s: str):
    """Parse a date string in various formats → datetime.date or None."""
    from datetime import datetime
    if not s or not s.strip():
        return None
    s = s.strip()
    for fmt in ("%d-%m-%Y", "%Y-%m-%d", "%d/%m/%Y", "%Y/%m/%d", "%d %b %Y", "%d %B %Y"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None


def _read_csv_from_bytes(
    data: bytes,
    amount_col: str,
    number_col: str = "",
    filter_dates: set | None = None,
    source_file: str = "",   # filename tag for tracing
) -> list[dict]:
    """
    Parse CSV bytes where header starts at row 6 (0-indexed row 5).
    Returns list of transaction dicts.
    """
    text = data.decode("utf-8", errors="replace")
    lines = text.splitlines()

    if len(lines) < 6:
        print(f"  [WARN] CSV has fewer than 6 lines, skipping")
        return []

    # Row 6 (index 5) is the header
    header_line = lines[5]
    data_lines  = lines[6:]

    file_category = "QR"
    header_upper = header_line.upper()
    if "CARD" in header_upper and "PRINCIPAL" in header_upper:
        file_category = "Debit Card"

    reader = csv.DictReader(
        [header_line] + data_lines,
        skipinitialspace=True,
    )

    # Normalize column names (strip whitespace)
    amount_col_stripped = amount_col.strip()

    txns = []
    skipped = 0

    for row_num, row in enumerate(reader, start=7):
        # Strip whitespace from all keys
        row = {k.strip(): str(v).strip() for k, v in row.items() if k is not None and v is not None}
        
        # Stop processing if we reach the TOTAL or SUMMARY GROUP section at the bottom
        first_val = str(list(row.values())[0]).upper() if row.values() else ""
        if first_val.startswith("TOTAL") or first_val.startswith("SUMMARY GROUP"):
            break

        raw = row.get(amount_col_stripped)
        if raw is None:
            # Try case-insensitive search
            for k, v in row.items():
                if k.upper() == amount_col_stripped.upper():
                    raw = v
                    break

        if not raw or raw.strip() in ("", "-", "0", "0.00"):
            skipped += 1
            continue

        try:
            amount = normalize_for_compare(parse_amount(raw))
        except ValueError as e:
            print(f"  [WARN] Mandiri CSV row {row_num}: {e} — skipped")
            skipped += 1
            continue

        # Date
        date_val = (
            row.get("TRXDATE") or
            row.get("TRANSACTION DATE") or
            row.get("TGL TRANSAKSI") or
            row.get("DATE") or ""
        )

        # Date filter: skip rows not in the allowed set
        txn_date = _parse_any_date(date_val)
        if filter_dates is not None:
            if txn_date not in filter_dates:
                continue
        
        final_date = str(txn_date) if txn_date else date_val

        desc_val = (
            row.get("DESCRIPTION") or
            row.get("KETERANGAN") or
            row.get("REMARK") or ""
        )

        # Reference number
        number_val = ""
        if number_col:
            nc = number_col.strip()
            number_val = row.get(nc) or ""
            if not number_val:
                for k, v in row.items():
                    if k.upper() == nc.upper():
                        number_val = v
                        break

        # Payment Date
        payment_date_val = (
            row.get("PAYMENT DATE") or
            row.get("TANGGAL BAYAR") or
            ""
        )
        payment_date = _parse_any_date(payment_date_val)
        
        # Fallback to H+1 if payment date is missing
        if not payment_date and txn_date:
            from datetime import date as dt_date, timedelta
            if isinstance(txn_date, dt_date):
                payment_date = txn_date + timedelta(days=1)
                
        final_payment_date = str(payment_date) if payment_date else payment_date_val

        # Admin fee
        admin_fee_val = row.get("MDR Amount") or row.get("MDR AMOUNT") or row.get("mdr amount") or ""
        try:
            admin_fee = parse_amount(admin_fee_val) if admin_fee_val else Decimal("0")
        except ValueError:
            admin_fee = Decimal("0")

        txns.append({
            "amount":      amount,
            "amount_raw":  raw,
            "admin_fee":   admin_fee,
            "date":        final_date,
            "payment_date": final_payment_date,
            "description": desc_val,
            "number":      number_val,
            "is_void":     "VOID" in desc_val.upper(),
            "is_reversal": "REVERSAL" in desc_val.upper(),
            "is_refund":   "REFUND" in desc_val.upper(),
            "filename":    source_file,
            "source":      "Bank (Mandiri)",
            "category":    file_category,
        })

    if skipped:
        print(f"  [INFO] Mandiri: {skipped} empty/zero rows skipped")

    return txns


def _find_mandiri_zips(zip_dir: Path, zip_pattern: str, password: str = "") -> list[Path]:
    """
    Find ALL Mandiri ZIP files in zip_dir.

    Primary:  filename matching zip_pattern (fast, zero I/O).
    Fallback: if no filename matches, return all .zip files in zip_dir —
              probing content via AES password if available. Handles renamed files.
    """
    candidates = sorted(p for p in zip_dir.glob(zip_pattern) if p.is_file() and not p.name.startswith("."))
    if candidates:
        return candidates

    # Fallback: check all .zip files in zip_dir
    all_zips = sorted(
        p for p in zip_dir.iterdir()
        if p.is_file() and p.suffix.lower() == ".zip" and not p.name.startswith(".")
    )
    if not all_zips:
        return []

    if password:
        try:
            from readers.file_detector import _probe_mandiri_zip
            probed = [z for z in all_zips if _probe_mandiri_zip(z, password)]
            if probed:
                return probed
        except Exception:
            pass

    return all_zips


def read_mandiri(
    zip_dir: Path,
    password: str,
    amount_col: str,
    number_col: str = "",
    zip_pattern: str = "MSR_*.zip",
    filter_dates: set | None = None,
) -> tuple[list[dict], list[dict]]:
    """
    Read all Mandiri transactions from ZIP files in zip_dir.
    If filter_dates is provided, only rows matching those dates are returned.
    """
    if not zip_dir.exists():
        raise FileNotFoundError(
            f"Mandiri ZIP directory not found: {zip_dir}\n"
            f"Check MANDIRI_ZIP_DIR in your .env file."
        )

    zip_files = _find_mandiri_zips(zip_dir, zip_pattern, password)
    if not zip_files:
        raise FileNotFoundError(
            f"No Mandiri ZIP files matching '{zip_pattern}' or fallback found in: {zip_dir}\n"
            f"Check MANDIRI_ZIP_PATTERN in your .env file."
        )

    print(f"  Found {len(zip_files)} Mandiri ZIP file(s): {[z.name for z in zip_files]}")

    all_txns: list[dict] = []
    pwd_bytes = password.encode("utf-8")

    for zip_path in zip_files:
        print(f"  Processing {zip_path.name}...")
        try:
            with pyzipper.AESZipFile(zip_path, "r") as zf:
                zf.setpassword(pwd_bytes)
                csv_names = [n for n in zf.namelist() if n.lower().endswith(".csv")]

                if not csv_names:
                    print(f"  [WARN] No CSV found in {zip_path.name}")
                    continue

                for csv_name in csv_names:
                    print(f"    Reading CSV: {csv_name}")
                    data = zf.read(csv_name)
                    txns = _read_csv_from_bytes(data, amount_col, number_col,
                                                filter_dates=filter_dates,
                                                source_file=zip_path.name)
                    print(f"    → {len(txns)} transactions")
                    all_txns.extend(txns)

        except RuntimeError as e:
            raise RuntimeError(
                f"Cannot open {zip_path.name}: wrong password or corrupt ZIP.\n"
                f"Check MANDIRI_ZIP_PASSWORD in your .env file.\nError: {e}"
            )

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
        
        print(f"  [INFO] Mandiri: Excluded {len(excluded_txns)} rows due to Void/Reversal found in description")
        all_txns = filtered_txns

    return all_txns, excluded_txns
