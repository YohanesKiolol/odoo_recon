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


def _read_csv_from_bytes(data: bytes, amount_col: str, number_col: str = "") -> list[dict]:
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
        row = {k.strip(): v.strip() for k, v in row.items() if k is not None}

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

        # Try to get date from common Mandiri CSV column names
        date_val = (
            row.get("TRXDATE") or
            row.get("TRANSACTION DATE") or
            row.get("TGL TRANSAKSI") or
            row.get("DATE") or ""
        )
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

        txns.append({
            "amount":      amount,
            "amount_raw":  raw,
            "date":        date_val,
            "description": desc_val,
            "number":      number_val,
            "source":      "Bank (Mandiri)",
        })

    if skipped:
        print(f"  [INFO] Mandiri: {skipped} empty/zero rows skipped")

    return txns


def read_mandiri(
    zip_dir: Path,
    password: str,
    amount_col: str,
    number_col: str = "",
    zip_pattern: str = "MSR_*.zip",
) -> list[dict]:
    """
    Read all Mandiri transactions from ZIP files in zip_dir.
    Only files matching zip_pattern (glob) are processed.

    Returns merged list of transaction dicts.
    """
    if not zip_dir.exists():
        raise FileNotFoundError(
            f"Mandiri ZIP directory not found: {zip_dir}\n"
            f"Check MANDIRI_ZIP_DIR in your .env file."
        )

    zip_files = sorted(zip_dir.glob(zip_pattern))
    if not zip_files:
        raise FileNotFoundError(
            f"No Mandiri ZIP files matching '{zip_pattern}' found in: {zip_dir}\n"
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
                    txns = _read_csv_from_bytes(data, amount_col, number_col)
                    print(f"    → {len(txns)} transactions")
                    all_txns.extend(txns)

        except RuntimeError as e:
            raise RuntimeError(
                f"Cannot open {zip_path.name}: wrong password or corrupt ZIP.\n"
                f"Check MANDIRI_ZIP_PASSWORD in your .env file.\nError: {e}"
            )

    return all_txns
