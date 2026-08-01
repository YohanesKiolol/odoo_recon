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
    Find ALL BRI ZIP files in zip_dir whose name contains zip_pattern.
    Returns sorted list.
    """
    candidates = sorted(
        p for p in zip_dir.iterdir()
        if p.suffix.lower() == ".zip" and zip_pattern.lower() in p.name.lower()
    )
    if not candidates:
        all_zips = [p.name for p in zip_dir.glob("*.zip")]
        raise FileNotFoundError(
            f"No BRI ZIP containing '{zip_pattern}' found in: {zip_dir}\n"
            f"Available ZIPs: {all_zips}\n"
            f"Check BRI_ZIP_PATTERN in your .env file."
        )
    return candidates


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
                for i, h in enumerate(headers):
                    hu = h.upper()
                    if date_idx is None and any(k in hu for k in ("DATE", "TGL", "TANGGAL")):
                        date_idx = i
                    if desc_idx is None and any(k in hu for k in ("DESC", "KET", "REMARK", "TRX_DESC", "KETERANGAN")):
                        desc_idx = i
                    if number_idx is None and number_col and hu == number_col.upper():
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

                    # Date filter
                    txn_date = _parse_any_date(date_val)
                    if filter_dates is not None:
                        if txn_date not in filter_dates:
                            skipped += 1
                            continue
                            
                    final_date = str(txn_date) if txn_date else date_val

                    txns.append({
                        "amount":      amount,
                        "amount_raw":  raw_amount,
                        "date":        final_date,
                        "description": desc_val,
                        "number":      number_val,
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
) -> list[dict]:
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

    all_txns: list[dict] = []
    for zip_path in zip_paths:
        print(f"  BRI ZIP: {zip_path.name}")
        pdf_bytes = _extract_detail_pdf(zip_path, pdf_pattern)
        print(f"  PDF extracted ({len(pdf_bytes):,} bytes), parsing table...")
        txns = _parse_pdf_table(pdf_bytes, amount_col, number_col,
                                filter_dates=filter_dates,
                                source_file=zip_path.name)
        all_txns.extend(txns)

    return all_txns
