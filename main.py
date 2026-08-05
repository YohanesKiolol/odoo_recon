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
    for pkg in ["openpyxl", "pyzipper", "msoffcrypto", "pdfplumber"]:
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
from readers.mutation_reader import read_all_mutations
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
    parser.add_argument(
        "--all", action="store_true",
        help="Process all transactions including unknown groups (e.g., PayPal, Wise) into an 'Other' bucket."
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

    from config import BANK_ACCOUNTS
    group_map = {}
    for bank_key in banks:
        for alias, acc_info in BANK_ACCOUNTS.get(bank_key, {}).items():
            grp = acc_info.get("group")
            if grp:
                group_map[grp] = f"{bank_key}_{alias}"

    try:
        odo_date, odo_bank_txns = read_odoo(
            excel_path  = ODO_EXCEL_PATH,
            amount_col  = ODO_AMOUNT_COLUMN,
            number_col  = ODO_NUMBER_COLUMN,
            group_map   = group_map,
            include_others = args.all,
        )
    except (FileNotFoundError, ValueError) as e:
        if args.scan:
            print(f"  [i] ODO file not found, continuing scan for bank files...")
            odo_date = None
            odo_bank_txns = {}
        else:
            print(f"\n[!] ERROR (ODO): {e}\n")
            sys.exit(1)

    if odo_date:
        print(f"  [+] ODO date: {odo_date}")
    for bank_key, txns in odo_bank_txns.items():
        print(f"     {bank_key:10} → {len(txns)} transactions")

    # ── Step 2: Read each bank ─────────────────────────────────────────────────
    bank_txns: dict[str, list[dict]] = {}

    if "bca" in banks:
        _banner("Reading BCA transactions...")
        for alias, acc_info in BANK_ACCOUNTS.get("bca", {}).items():
            acc_key = f"bca_{alias}"
            target_dir = BCA_EXCEL_DIR / alias
            if alias == "main" and not target_dir.exists():
                target_dir = BCA_EXCEL_DIR
            try:
                from datetime import datetime as _dt
                filter_dates = None

                bank_txns[acc_key] = read_bca(
                    excel_dir     = target_dir,
                    excel_pattern = BCA_EXCEL_PATTERN,
                    password      = BCA_EXCEL_PASSWORD,
                    amount_col    = BCA_AMOUNT_COLUMN,
                    date_col      = BCA_DATE_COLUMN,
                    number_col    = BCA_NUMBER_COLUMN,
                    filter_dates  = filter_dates,
                )
            except FileNotFoundError as e:
                print(f"  [!] No BCA files found for {alias}, skipping.")
            except Exception as e:
                print(f"\n[!] ERROR (BCA {alias}): {e}\n")
                sys.exit(1)

    if "mandiri" in banks:
        _banner("Reading Mandiri transactions...")
        for alias, acc_info in BANK_ACCOUNTS.get("mandiri", {}).items():
            acc_key = f"mandiri_{alias}"
            target_dir = MANDIRI_ZIP_DIR / alias
            if alias == "main" and not target_dir.exists():
                target_dir = MANDIRI_ZIP_DIR
            try:
                from datetime import datetime as _dt
                filter_dates = None
                bank_txns[acc_key] = read_mandiri(
                    zip_dir     = target_dir,
                    password    = MANDIRI_ZIP_PASSWORD,
                    amount_col  = MANDIRI_AMOUNT_COLUMN,
                    number_col  = MANDIRI_NUMBER_COLUMN,
                    zip_pattern = MANDIRI_ZIP_PATTERN,
                    filter_dates = filter_dates,
                )
                print(f"  [+] Mandiri ({alias}): {len(bank_txns.get(acc_key, []))} transactions loaded")
            except FileNotFoundError as e:
                print(f"  [-] No Mandiri files found for {alias}, skipping.")
            except Exception as e:
                print(f"\n[!] ERROR (Mandiri {alias}): {e}\n")
                sys.exit(1)

    if "bri" in banks:
        _banner("Reading BRI transactions...")
        for alias, acc_info in BANK_ACCOUNTS.get("bri", {}).items():
            acc_key = f"bri_{alias}"
            target_dir = BRI_ZIP_DIR / alias
            if alias == "main" and not target_dir.exists():
                target_dir = BRI_ZIP_DIR
            try:
                from datetime import datetime as _dt
                filter_dates = None
                bank_txns[acc_key] = read_bri(
                    zip_dir      = target_dir,
                    zip_pattern  = acc_info.get("mid", "") or BRI_ZIP_PATTERN,
                    pdf_pattern  = BRI_PDF_PATTERN,
                    amount_col   = BRI_AMOUNT_COLUMN,
                    number_col   = BRI_NUMBER_COLUMN,
                    filter_dates = filter_dates,
                )
                print(f"  [+] BRI ({alias}): {len(bank_txns.get(acc_key, []))} transactions loaded")
            except FileNotFoundError as e:
                print(f"  [-] No BRI files found for {alias}, skipping.")
            except Exception as e:
                print(f"\n[!] ERROR (BRI {alias}): {e}\n")
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
        all_dates = []
        for acc_key, txns in bank_txns.items():
            print(f"    {acc_key}:")
            dates = Counter(t.get("date", "Unknown") for t in txns)
            if not dates:
                print("      (No transactions)")
            for d, c in sorted(dates.items()):
                print(f"      {d} : {c} trxs")
                d_str = str(d).strip()
                if d_str and d_str != "Unknown":
                    # basic check for YYYY-MM-DD
                    if len(d_str) >= 10 and d_str[0:4].isdigit():
                        all_dates.append(d_str[:10])
        
        print(f"\n{'─' * 60}\n")
        
        if all_dates:
            print(f"[DATE_RANGE]|{min(all_dates)}|{max(all_dates)}")
            
        sys.exit(0)

    # ── Step 3: Reconcile per bank ─────────────────────────────────────────────
    _banner("Comparing transactions...")

    all_results: dict[str, list[dict]] = {}

    print(f"\n  {'Bank':<10} {'Done':>6} {'Bank Only':>10} {'ODO Only':>9} {'Total':>7}")
    print(f"  {'─'*10} {'─'*6} {'─'*10} {'─'*9} {'─'*7}")

    # Make sure we process "other" if it exists in Odoo txns
    keys_to_process = list(bank_txns.keys())
    if "other" in odo_bank_txns and "other" not in keys_to_process:
        keys_to_process.append("other")

    for acc_key in keys_to_process:
        b_txns = bank_txns.get(acc_key, [])
        o_txns = odo_bank_txns.get(acc_key, [])
        
        if acc_key == "other":
            # Don't filter by valid bank dates since we don't have bank files for "other"
            o_txns_filtered = o_txns
        else:
            # Filter Odoo transactions: only keep dates that actually exist in the uploaded Bank files
            valid_bank_dates = {str(t.get("date")) for t in b_txns if t.get("date")}
            valid_odoo_dates = {str(t.get("date")) for t in o_txns if t.get("date")}
            
            missing_in_odoo = valid_bank_dates - valid_odoo_dates
            if missing_in_odoo:
                missing_str = ", ".join(sorted(missing_in_odoo))
                print(f"  [!] WARNING: {acc_key} has Bank data for {missing_str} but No Odoo data exists!")

            o_txns_filtered = [t for t in o_txns if str(t.get("date")) in valid_bank_dates]

        results = reconcile(b_txns, o_txns_filtered)
        stats   = summary(results)
        all_results[acc_key] = results

        print(f"  {acc_key:<10} {stats['done']:>6} {stats['bank_only']:>10} {stats['odo_only']:>9} {stats['total']:>7}")

    # ── Step 4: Parse Mutations ───────────────────────────────────────────────
    _banner("Reading Mutations...")
    mutations, unknown_mutations = read_all_mutations()
    print(f"  [+] Loaded {len(mutations)} recognized mutations")
    if unknown_mutations:
        print(f"  [!] Found {len(unknown_mutations)} unknown mutation patterns")

    # ── Step 5: Write report ───────────────────────────────────────────────────
    _banner("Writing Excel report...")

    try:
        out_path = write_report(
            all_results,
            odo_date,
            OUTPUT_DIR,
            bank_txns=bank_txns,
            odo_bank_txns=odo_bank_txns,
            mutations=mutations,
            unknown_mutations=unknown_mutations,
        )
    except Exception as e:
        print(f"\n[!] ERROR writing report: {e}\n")
        sys.exit(1)

    print(f"\n[+] Report saved to:\n   {out_path.resolve()}")

    total_disc = sum(
        s["bank_only"] + s["odo_only"]
        for r in all_results.values()
        for s in [summary(r)]
    )
    if total_disc > 0:
        print(f"\n[!] {total_disc} discrepancies found — check 'Selisih (Semua)' sheet.")
    else:
        print("\n[+] All transactions matched across all banks!")

    print()


if __name__ == "__main__":
    main()
