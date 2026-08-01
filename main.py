"""
main.py — Odoo ↔ Bank Reconciliation (Multi-Bank)

Compares:
  BCA     : password-protected Excel (filtered by ODO date)
  Mandiri : 2 password-protected ZIPs containing CSVs
  BRI     : ZIP → folder → detail_****.pdf table

vs Odoo:
  Payments (account.payment).xlsx — grouped by bank transaction type

Usage:
    python main.py              # runs all 3 banks
    python main.py --bank bca   # run only BCA
    python main.py --bank mandiri bri  # run Mandiri and BRI only
"""

import sys
import argparse

# ── Dependency check ──────────────────────────────────────────────────────────
def _check_deps():
    missing = []
    for pkg in ["openpyxl", "dotenv", "pyzipper", "msoffcrypto", "pdfplumber"]:
        try:
            __import__(pkg)
        except ImportError:
            missing.append(pkg)
    if missing:
        print("=" * 60)
        print("ERROR: Dependencies not installed.")
        print("Run:  pip install -r requirements.txt")
        print(f"Missing: {missing}")
        print("=" * 60)
        sys.exit(1)

_check_deps()

from config import (
    ODO_EXCEL_PATH, ODO_AMOUNT_COLUMN, ODO_NUMBER_COLUMN,
    ODO_GROUP_BCA, ODO_GROUP_MANDIRI, ODO_GROUP_BRI,
    MANDIRI_ZIP_DIR, MANDIRI_ZIP_PATTERN, MANDIRI_ZIP_PASSWORD, MANDIRI_AMOUNT_COLUMN, MANDIRI_NUMBER_COLUMN,
    BCA_EXCEL_DIR, BCA_EXCEL_PATTERN, BCA_EXCEL_PASSWORD, BCA_AMOUNT_COLUMN, BCA_DATE_COLUMN, BCA_NUMBER_COLUMN,
    BRI_ZIP_DIR, BRI_ZIP_PATTERN, BRI_PDF_PATTERN, BRI_AMOUNT_COLUMN, BRI_NUMBER_COLUMN,
    OUTPUT_DIR,
)
from readers.odoo_reader    import read_odoo
from readers.mandiri_reader import read_mandiri
from readers.bca_reader     import read_bca
from readers.bri_reader     import read_bri
from reconciler             import reconcile, summary
from excel_writer           import write_report

ALL_BANKS = ["bca", "mandiri", "bri"]


def parse_args():
    parser = argparse.ArgumentParser(
        description="Reconcile Bank transactions vs Odoo by amount."
    )
    parser.add_argument(
        "--bank", nargs="+", choices=ALL_BANKS,
        help=f"Which banks to process. Default: all. Options: {ALL_BANKS}"
    )
    parser.add_argument(
        "--scan", action="store_true",
        help="Scan files and print a summary of dates and counts, then exit."
    )
    return parser.parse_args()


def _banner(text: str):
    print(f"\n{'─' * 60}")
    print(f"  {text}")
    print(f"{'─' * 60}")


def main():
    # Force UTF-8 output — Windows console defaults to cp1252 which
    # can't encode ↔, ✅, ❌ etc. and raises UnicodeEncodeError.
    import io as _io
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    elif hasattr(sys.stdout, "buffer"):
        sys.stdout = _io.TextIOWrapper(sys.stdout.buffer,
                                       encoding="utf-8", errors="replace")

    args    = parse_args()
    banks   = [b.lower() for b in args.bank] if args.bank else ALL_BANKS

    print()
    print("=" * 60)
    print("  Odoo ↔ Bank Reconciliation  (Multi-Bank)")
    print("=" * 60)
    print(f"  ODO file : {ODO_EXCEL_PATH}")
    print(f"  Banks    : {', '.join(b.upper() for b in banks)}")
    print(f"  Output   : {OUTPUT_DIR}")
    print("=" * 60)

    # ── Step 1: Read ODO ───────────────────────────────────────────────────────
    _banner("Reading Odoo transactions...")

    group_map = {}
    if "bca"     in banks: group_map[ODO_GROUP_BCA]     = "BCA"
    if "mandiri" in banks: group_map[ODO_GROUP_MANDIRI] = "Mandiri"
    if "bri"     in banks: group_map[ODO_GROUP_BRI]     = "BRI"

    try:
        odo_date, odo_bank_txns = read_odoo(
            excel_path  = ODO_EXCEL_PATH,
            amount_col  = ODO_AMOUNT_COLUMN,
            number_col  = ODO_NUMBER_COLUMN,
            group_map   = group_map,
        )
    except (FileNotFoundError, ValueError) as e:
        print(f"\n❌ ERROR (ODO): {e}\n")
        sys.exit(1)

    print(f"  ✅ ODO date: {odo_date}")
    for bank_key, txns in odo_bank_txns.items():
        print(f"     {bank_key:10} → {len(txns)} transactions")

    # ── Step 2: Read each bank ─────────────────────────────────────────────────
    bank_txns: dict[str, list[dict]] = {}

    if "bca" in banks:
        _banner("Reading BCA transactions...")
        try:
            # Derive allowed dates from ODO BCA transactions.
            # Single ODO date  → same filter as original behavior.
            # Multiple ODO dates → all are allowed automatically.
            from datetime import datetime as _dt
            bca_filter_dates = {
                _dt.strptime(t["date"], "%Y-%m-%d").date()
                for t in odo_bank_txns.get("BCA", [])
                if t.get("date")
            } or None   # None = no filter if ODO has no dates (shouldn't happen)

            bank_txns["BCA"] = read_bca(
                excel_dir     = BCA_EXCEL_DIR,
                excel_pattern = BCA_EXCEL_PATTERN,
                password      = BCA_EXCEL_PASSWORD,
                amount_col    = BCA_AMOUNT_COLUMN,
                date_col      = BCA_DATE_COLUMN,
                number_col    = BCA_NUMBER_COLUMN,
                filter_dates  = bca_filter_dates,
            )
        except FileNotFoundError as e:
            print(f"  ⚠️  No BCA files found, skipping BCA.")
        except Exception as e:
            print(f"\n❌ ERROR (BCA): {e}\n")
            sys.exit(1)

    if "mandiri" in banks:
        _banner("Reading Mandiri transactions...")
        try:
            from datetime import datetime as _dt
            mandiri_filter_dates = {
                _dt.strptime(t["date"], "%Y-%m-%d").date()
                for t in odo_bank_txns.get("Mandiri", [])
                if t.get("date")
            } or None
            bank_txns["Mandiri"] = read_mandiri(
                zip_dir     = MANDIRI_ZIP_DIR,
                password    = MANDIRI_ZIP_PASSWORD,
                amount_col  = MANDIRI_AMOUNT_COLUMN,
                number_col  = MANDIRI_NUMBER_COLUMN,
                zip_pattern = MANDIRI_ZIP_PATTERN,
                filter_dates = mandiri_filter_dates,
            )
            print(f"  ✅ Mandiri: {len(bank_txns['Mandiri'])} transactions loaded")
        except FileNotFoundError as e:
            print(f"  ⚠️  No Mandiri files found, skipping Mandiri.")
        except Exception as e:
            print(f"\n❌ ERROR (Mandiri): {e}\n")
            sys.exit(1)

    if "bri" in banks:
        _banner("Reading BRI transactions...")
        try:
            from datetime import datetime as _dt
            bri_filter_dates = {
                _dt.strptime(t["date"], "%Y-%m-%d").date()
                for t in odo_bank_txns.get("BRI", [])
                if t.get("date")
            } or None
            bank_txns["BRI"] = read_bri(
                zip_dir      = BRI_ZIP_DIR,
                zip_pattern  = BRI_ZIP_PATTERN,
                pdf_pattern  = BRI_PDF_PATTERN,
                amount_col   = BRI_AMOUNT_COLUMN,
                number_col   = BRI_NUMBER_COLUMN,
                filter_dates = bri_filter_dates,
            )
            print(f"  ✅ BRI: {len(bank_txns['BRI'])} transactions loaded")
        except FileNotFoundError as e:
            print(f"  ⚠️  No BRI files found, skipping BRI.")
        except Exception as e:
            print(f"\n❌ ERROR (BRI): {e}\n")
            sys.exit(1)

    # ── Handle Scan Mode ───────────────────────────────────────────────────────
    if args.scan:
        from collections import Counter
        _banner("SCAN SUMMARY")
        
        print("  [ ODO FILE ]")
        for bank_key, txns in odo_bank_txns.items():
            print(f"    {bank_key}:")
            dates = Counter(t.get("date", "Unknown") for t in txns)
            if not dates:
                print("      (No transactions)")
            for d, c in sorted(dates.items()):
                print(f"      {d} : {c} trxs")
        
        print("\n  [ BANK FILES ]")
        for bank_key in banks:
            b_key = bank_key.upper()
            if b_key == "MANDIRI": b_key = "Mandiri"
            
            txns = bank_txns.get(b_key, [])
            print(f"    {b_key}:")
            dates = Counter(t.get("date", "Unknown") for t in txns)
            if not dates:
                print("      (No transactions)")
            for d, c in sorted(dates.items()):
                print(f"      {d} : {c} trxs")
        
        print(f"\n{'─' * 60}\n")
        sys.exit(0)

    # ── Step 3: Reconcile per bank ─────────────────────────────────────────────
    _banner("Comparing transactions...")

    all_results: dict[str, list[dict]] = {}

    print(f"\n  {'Bank':<10} {'Done':>6} {'Bank Only':>10} {'ODO Only':>9} {'Total':>7}")
    print(f"  {'─'*10} {'─'*6} {'─'*10} {'─'*9} {'─'*7}")

    for bank_name in ["BCA", "Mandiri", "BRI"]:
        if bank_name not in bank_txns:
            continue

        b_txns = bank_txns[bank_name]
        o_txns = odo_bank_txns.get(bank_name, [])

        results = reconcile(b_txns, o_txns)
        stats   = summary(results)
        all_results[bank_name] = results

        print(f"  {bank_name:<10} {stats['done']:>6} {stats['bank_only']:>10} {stats['odo_only']:>9} {stats['total']:>7}")

    # ── Step 4: Write report ───────────────────────────────────────────────────
    _banner("Writing Excel report...")

    try:
        out_path = write_report(
            all_results,
            odo_date,
            OUTPUT_DIR,
            bank_txns=bank_txns,
            odo_bank_txns=odo_bank_txns,
        )
    except Exception as e:
        print(f"\n❌ ERROR writing report: {e}\n")
        sys.exit(1)

    print(f"\n✅ Report saved to:\n   {out_path.resolve()}")

    total_disc = sum(
        s["bank_only"] + s["odo_only"]
        for r in all_results.values()
        for s in [summary(r)]
    )
    if total_disc > 0:
        print(f"\n⚠️  {total_disc} discrepancies found — check 'Selisih (Semua)' sheet.")
    else:
        print("\n🎉 All transactions matched across all banks!")

    print()


if __name__ == "__main__":
    main()
