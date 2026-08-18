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
from pathlib import Path

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
    parser.add_argument(
        "--no-open", action="store_true",
        help="Do not auto-open the generated Excel file at the end."
    )
    return parser.parse_args()


def _banner(text: str):
    print(f"\n{'─' * 60}")
    print(f"  {text}")
    print(f"{'─' * 60}")


def scan_bank_date_range(banks=None, log_fn=None) -> tuple[str, str] | None:
    """Scan bank files in parallel and return the (min_date, max_date) ISO range."""
    import re, zipfile
    from collections import Counter
    from concurrent.futures import ThreadPoolExecutor
    from config import BANK_ACCOUNTS
    
    if not banks or any(b.lower() == "all" for b in banks):
        banks_to_scan = ALL_BANKS
    else:
        banks_to_scan = [b.lower() for b in banks]

    all_dates = []
    bank_details = {}

    def _log(msg):
        if log_fn:
            log_fn(msg)
        else:
            print(msg, end="")

    def _do_bca():
        bca_dates = []
        try:
            from readers.bca_reader import _find_bca_excels, _read_one_bca
            search_dirs = []
            for alias, acc_info in BANK_ACCOUNTS.get("bca", {}).items():
                target_dir = BCA_EXCEL_DIR / alias
                if alias == "main" and not target_dir.exists():
                    target_dir = BCA_EXCEL_DIR
                if target_dir.exists() and target_dir not in search_dirs:
                    search_dirs.append(target_dir)

            all_excels = []
            for sdir in search_dirs:
                all_excels.extend(_find_bca_excels(sdir, BCA_EXCEL_PATTERN))

            with ThreadPoolExecutor(max_workers=min(4, len(all_excels) or 1)) as bca_pool:
                bca_futs = [bca_pool.submit(_read_one_bca, f, BCA_EXCEL_PASSWORD, BCA_AMOUNT_COLUMN, BCA_DATE_COLUMN, BCA_NUMBER_COLUMN) for f in all_excels]
                for fut in bca_futs:
                    for r in fut.result():
                        d_str = str(r.get("date") or r.get("txn_date") or "").strip()
                        if len(d_str) >= 10 and d_str[0:4].isdigit():
                            all_dates.append(d_str[:10])
                            bca_dates.append(d_str[:10])
        except Exception:
            pass
        bank_details["bca"] = Counter(bca_dates)

    def _do_mandiri():
        man_dates = []
        try:
            from readers.mandiri_reader import extract_mandiri_dates_from_zip
            search_dirs = []
            for alias, acc_info in BANK_ACCOUNTS.get("mandiri", {}).items():
                target_dir = MANDIRI_ZIP_DIR / alias
                if alias == "main" and not target_dir.exists():
                    target_dir = MANDIRI_ZIP_DIR
                if target_dir.exists() and target_dir not in search_dirs:
                    search_dirs.append(target_dir)

            for sdir in search_dirs:
                for z_path in sdir.glob("*.zip"):
                    found = extract_mandiri_dates_from_zip(z_path, MANDIRI_ZIP_PASSWORD)
                    for d in found:
                        all_dates.append(d)
                        man_dates.append(d)
        except Exception:
            pass
        bank_details["mandiri"] = Counter(man_dates)

    def _do_bri():
        bri_dates = []
        try:
            search_dirs = []
            for alias, acc_info in BANK_ACCOUNTS.get("bri", {}).items():
                target_dir = BRI_ZIP_DIR / alias
                if alias == "main" and not target_dir.exists():
                    target_dir = BRI_ZIP_DIR
                if target_dir.exists() and target_dir not in search_dirs:
                    search_dirs.append(target_dir)

            for sdir in search_dirs:
                for z_path in sdir.glob("*.zip"):
                    try:
                        with zipfile.ZipFile(z_path, "r") as zf:
                            for name in zf.namelist():
                                m = re.search(r'(\d{4}-\d{2}-\d{2})', name)
                                if m:
                                    iso_d = m.group(1)
                                    all_dates.append(iso_d)
                                    bri_dates.append(iso_d)
                                else:
                                    m2 = re.search(r'((?:20\d{2})(?:0[1-9]|1[0-2])(?:0[1-9]|[12]\d|3[01]))', name)
                                    if m2:
                                        ds = m2.group(1)
                                        iso_d = f"{ds[:4]}-{ds[4:6]}-{ds[6:8]}"
                                        all_dates.append(iso_d)
                                        bri_dates.append(iso_d)
                    except Exception:
                        pass
        except Exception:
            pass
        bank_details["bri"] = Counter(bri_dates)

    futs = []
    with ThreadPoolExecutor(max_workers=3, thread_name_prefix="scan") as pool:
        if "bca" in banks_to_scan: futs.append(pool.submit(_do_bca))
        if "mandiri" in banks_to_scan: futs.append(pool.submit(_do_mandiri))
        if "bri" in banks_to_scan: futs.append(pool.submit(_do_bri))
        for f in futs:
            try:
                f.result()
            except Exception:
                pass

    _log("  [ BANK FILES ]\n")
    for b_key in sorted(bank_details.keys()):
        cnt = bank_details[b_key]
        _log(f"    {b_key.upper()}:\n")
        if not cnt:
            _log("      (No transactions found)\n")
        else:
            for d, c in sorted(cnt.items()):
                _log(f"      {d} : {c} trxs\n")
    _log(f"\n{'─' * 60}\n\n")

    if all_dates:
        min_d, max_d = min(all_dates), max(all_dates)
        _log(f"[DATE_RANGE]|{min_d}|{max_d}\n")
        return (min_d, max_d)
    return None


def run_reconciliation(banks=None, process_all=False, open_file=False) -> Path:
    """Execute the full multi-bank reconciliation pipeline in-process and return output Path."""
    from concurrent.futures import ThreadPoolExecutor
    from config import BANK_ACCOUNTS

    if not banks or any(b.lower() == "all" for b in banks):
        banks = ALL_BANKS
    else:
        banks = [b.lower() for b in banks]

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
            include_others = process_all,
        )
    except (FileNotFoundError, ValueError) as e:
        print(f"\n[!] ERROR (ODO): {e}\n")
        raise

    if odo_date:
        print(f"  [+] ODO date: {odo_date}")
    for bank_key, txns in odo_bank_txns.items():
        print(f"     {bank_key:10} → {len(txns)} transactions")

    # ── Step 2: Read each bank (parallel) ─────────────────────────────────────
    bank_txns: dict[str, list[dict]] = {}
    all_excluded_txns: list[dict] = []

    def _do_bca():
        frag, exc = {}, []
        _banner("Reading BCA transactions...")
        for alias, acc_info in BANK_ACCOUNTS.get("bca", {}).items():
            acc_key = f"bca_{alias}"
            target_dir = BCA_EXCEL_DIR / alias
            if alias == "main" and not target_dir.exists():
                target_dir = BCA_EXCEL_DIR
            try:
                b_txns, exc_txns = read_bca(
                    excel_dir=target_dir, excel_pattern=BCA_EXCEL_PATTERN,
                    password=BCA_EXCEL_PASSWORD, amount_col=BCA_AMOUNT_COLUMN,
                    date_col=BCA_DATE_COLUMN, number_col=BCA_NUMBER_COLUMN,
                    filter_dates=None,
                )
                frag[acc_key] = b_txns
                exc.extend(exc_txns)
            except FileNotFoundError:
                print(f"  [!] No BCA files for {alias}, skipping.")
            except Exception as e:
                print(f"\n[!] ERROR (BCA {alias}): {e}\n")
                raise
        return frag, exc

    def _do_mandiri():
        frag, exc = {}, []
        _banner("Reading Mandiri transactions...")
        for alias, acc_info in BANK_ACCOUNTS.get("mandiri", {}).items():
            acc_key = f"mandiri_{alias}"
            target_dir = MANDIRI_ZIP_DIR / alias
            if alias == "main" and not target_dir.exists():
                target_dir = MANDIRI_ZIP_DIR
            try:
                b_txns, exc_txns = read_mandiri(
                    zip_dir=target_dir, password=MANDIRI_ZIP_PASSWORD,
                    amount_col=MANDIRI_AMOUNT_COLUMN, number_col=MANDIRI_NUMBER_COLUMN,
                    zip_pattern=MANDIRI_ZIP_PATTERN, filter_dates=None,
                )
                frag[acc_key] = b_txns
                exc.extend(exc_txns)
                print(f"  [+] Mandiri ({alias}): {len(b_txns)} transactions loaded")
            except FileNotFoundError:
                print(f"  [-] No Mandiri files for {alias}, skipping.")
            except Exception as e:
                print(f"\n[!] ERROR (Mandiri {alias}): {e}\n")
                raise
        return frag, exc

    def _do_bri():
        frag, exc = {}, []
        _banner("Reading BRI transactions...")
        for alias, acc_info in BANK_ACCOUNTS.get("bri", {}).items():
            acc_key = f"bri_{alias}"
            target_dir = BRI_ZIP_DIR / alias
            if alias == "main" and not target_dir.exists():
                target_dir = BRI_ZIP_DIR
            try:
                b_txns, exc_txns = read_bri(
                    zip_dir=target_dir,
                    zip_pattern=acc_info.get("mid", "") or BRI_ZIP_PATTERN,
                    pdf_pattern=BRI_PDF_PATTERN,
                    amount_col=BRI_AMOUNT_COLUMN, number_col=BRI_NUMBER_COLUMN,
                    filter_dates=None,
                )
                frag[acc_key] = b_txns
                exc.extend(exc_txns)
                print(f"  [+] BRI ({alias}): {len(b_txns)} transactions loaded")
            except FileNotFoundError:
                print(f"  [-] No BRI files for {alias}, skipping.")
            except Exception as e:
                print(f"\n[!] ERROR (BRI {alias}): {e}\n")
                raise
        return frag, exc

    workers = {}
    with ThreadPoolExecutor(max_workers=4, thread_name_prefix="recon") as pool:
        if "bca" in banks:     workers["bca"]       = pool.submit(_do_bca)
        if "mandiri" in banks: workers["mandiri"]   = pool.submit(_do_mandiri)
        if "bri" in banks:     workers["bri"]       = pool.submit(_do_bri)
        workers["mutations"] = pool.submit(read_all_mutations)

        for key, fut in workers.items():
            if key == "mutations":
                continue
            try:
                frag, exc = fut.result()
                bank_txns.update(frag)
                all_excluded_txns.extend(exc)
            except Exception as e:
                print(f"\n[!] ERROR ({key}): {e}\n")
                raise

    # ── Step 3: Reconcile per bank ─────────────────────────────────────────────
    _banner("Comparing transactions...")

    all_results: dict[str, list[dict]] = {}

    print(f"\n  {'Bank':<10} {'Done':>6} {'Bank Only':>10} {'ODO Only':>9} {'Total':>7}")
    print(f"  {'─'*10} {'─'*6} {'─'*10} {'─'*9} {'─'*7}")

    keys_to_process = list(bank_txns.keys())
    if "other" in odo_bank_txns and "other" not in keys_to_process:
        keys_to_process.append("other")

    for acc_key in keys_to_process:
        b_txns = bank_txns.get(acc_key, [])
        o_txns = odo_bank_txns.get(acc_key, [])
        
        if acc_key != "other":
            valid_bank_dates = {str(t.get("date")) for t in b_txns if t.get("date")}
            valid_odoo_dates = {str(t.get("date")) for t in o_txns if t.get("date")}
            missing_in_odoo = valid_bank_dates - valid_odoo_dates
            if missing_in_odoo:
                missing_str = ", ".join(sorted(missing_in_odoo))
                print(f"  [!] WARNING: {acc_key} has Bank data for {missing_str} but No Odoo data exists!")

        results = reconcile(b_txns, o_txns)
        stats   = summary(results)
        all_results[acc_key] = results

        print(f"  {acc_key:<10} {stats['done']:>6} {stats['bank_only']:>10} {stats['odo_only']:>9} {stats['total']:>7}")

    # ── Step 4: Parse Mutations ───────────────────────────────────────────────
    _banner("Reading Mutations...")
    try:
        mutations, unknown_mutations = workers["mutations"].result()
    except Exception as e:
        print(f"  [!] Mutations read failed ({e}), falling back to empty.")
        mutations, unknown_mutations = [], []

    print(f"  [+] Loaded {len(mutations)} recognized mutations")
    if unknown_mutations:
        print(f"  [!] Found {len(unknown_mutations)} unknown mutation patterns")

    # ── Step 5: Write report ───────────────────────────────────────────────────
    _banner("Writing Excel report...")

    out_path = write_report(
        all_results,
        odo_date,
        OUTPUT_DIR,
        bank_txns=bank_txns,
        odo_bank_txns=odo_bank_txns,
        mutations=mutations,
        unknown_mutations=unknown_mutations,
        excluded_txns=all_excluded_txns,
    )

    print(f"\n[+] Report saved to:\n   {out_path.resolve()}")

    total_disc = sum(
        s["bank_only"] + s["odo_only"]
        for r in all_results.values()
        for s in [summary(r)]
    )
    if total_disc > 0:
        print(f"\n[!] {total_disc} discrepancies found — check 'Differences' sheet.")
    else:
        print("\n[+] All transactions matched across all banks!")

    if open_file:
        print()
        import os
        if os.name == 'nt':
            os.startfile(out_path)
        else:
            import subprocess
            subprocess.run(["open", str(out_path)])

    return out_path


def main():
    import io as _io
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    elif hasattr(sys.stdout, "buffer"):
        sys.stdout = _io.TextIOWrapper(sys.stdout.buffer,
                                       encoding="utf-8", errors="replace")

    args    = parse_args()
    banks   = [b.lower() for b in args.bank] if args.bank else ALL_BANKS

    if args.scan:
        _banner("SCAN SUMMARY")
        dates_range = scan_bank_date_range(banks)
        if dates_range:
            print(f"[DATE_RANGE]|{dates_range[0]}|{dates_range[1]}")
        sys.exit(0)

    try:
        run_reconciliation(banks=banks, process_all=args.all, open_file=not args.no_open)
    except Exception as e:
        print(f"\n[!] Reconciliation failed: {e}\n")
        sys.exit(1)


if __name__ == "__main__":
    main()
