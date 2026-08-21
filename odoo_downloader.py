"""
Direct XML-RPC Downloader for Odoo Customer Payments & Settlement Journal Entries.
Fast, headless, and browserless (<1 second execution).
"""
import argparse
import sys
import os
import openpyxl
from datetime import datetime
from pathlib import Path

from config import (
    ODO_EXCEL_PATH, BANK_ACCOUNTS, ODO_JOURNAL_EXCEL_PATH,
    ODOO_JOURNAL_EDC, ODOO_JOURNAL_AR,
    ODO_GROUP_BCA, ODO_GROUP_MANDIRI, ODO_GROUP_BRI
)
import odoo_inspector
from odoo_inspector import _normalize_date_to_iso


def _get_bank_journal_names(bank_key: str) -> list[str]:
    """Retrieve all configured journal group names for a bank from BANK_ACCOUNTS."""
    b_key = bank_key.lower()
    j_names = []
    for alias, info in BANK_ACCOUNTS.get(b_key, {}).items():
        grp = info.get("group")
        if grp and grp not in j_names:
            j_names.append(grp)
    legacy = {
        "bca": ODO_GROUP_BCA,
        "mandiri": ODO_GROUP_MANDIRI,
        "bri": ODO_GROUP_BRI,
    }.get(b_key, "")
    if legacy and legacy not in j_names:
        j_names.append(legacy)
    return j_names


def download_via_xmlrpc(
    date_from: str = "",
    date_to: str = "",
    banks: list[str] | str | None = None,
    status: str = "Posted",
    mode: str = "both",
    exact_dates_by_bank: dict[str, list[str]] | None = None,
    exact_dates: list[str] | None = None,
) -> bool:
    """
    Download Customer Payments and Journal Entries directly via Odoo XML-RPC.
    Supports exact date list filtering per bank to handle non-sequential dates.
    Generates exact Excel formats matching Odoo web exports.
    """
    try:
        iso_from = _normalize_date_to_iso(date_from)
        iso_to = _normalize_date_to_iso(date_to)

        # Resolve selected bank list
        if isinstance(banks, str):
            selected_banks = [b.strip() for b in banks.split(",") if b.strip()]
        elif isinstance(banks, list):
            selected_banks = [b.strip() for b in banks if b.strip()]
        else:
            selected_banks = []

        is_all_banks = not selected_banks or any(b.lower() == "all" for b in selected_banks) or len(selected_banks) >= 3

        # Auto-detect date range if missing
        if not iso_from or not iso_to:
            try:
                from main import scan_bank_dates_detailed
                banks_to_scan = selected_banks if not is_all_banks else None
                scan_details = scan_bank_dates_detailed(banks_to_scan)
                if scan_details.get("min_date") and scan_details.get("max_date"):
                    iso_from = _normalize_date_to_iso(scan_details["min_date"])
                    iso_to = _normalize_date_to_iso(scan_details["max_date"])
                    if not exact_dates_by_bank:
                        exact_dates_by_bank = scan_details.get("dates_by_bank")
                    if not exact_dates:
                        exact_dates = scan_details.get("all_dates")
            except Exception:
                pass

        if not iso_from:
            iso_from = datetime.now().strftime("%Y-%m-%d")
        if not iso_to:
            iso_to = iso_from
        
        print(f"\n[+] Direct XML-RPC Downloader active ({iso_from} to {iso_to})...")

        # Resolve matching journal names from config
        target_journal_names = []
        if not is_all_banks and selected_banks:
            for b in selected_banks:
                b_journals = _get_bank_journal_names(b)
                if b_journals:
                    target_journal_names.extend(b_journals)
                else:
                    target_journal_names.append(b)

        # ── 1. Customer Payments (account.payment) ──
        if mode in ("payments", "both", "auto_recon"):
            print(f"[+] Fetching Customer Payments via XML-RPC...")
            payments = []

            if exact_dates_by_bank:
                target_banks_list = ["bca", "mandiri", "bri"] if is_all_banks else [b.lower() for b in selected_banks]
                seen_payment_ids = set()
                
                for b_name in target_banks_list:
                    b_dates = exact_dates_by_bank.get(b_name, [])
                    b_journals = _get_bank_journal_names(b_name)
                    if not b_dates:
                        continue
                    
                    p_domain = [
                        ('date', 'in', b_dates),
                        ('payment_type', '=', 'inbound'),
                    ]
                    if b_journals:
                        p_domain.append(('journal_id.name', 'in', b_journals))
                    if status:
                        st_list = [s.strip().lower() for s in status.split(',') if s.strip()]
                        if len(st_list) == 1:
                            p_domain.append(('state', '=', st_list[0]))
                        elif len(st_list) > 1:
                            p_domain.append(('state', 'in', st_list))
                    
                    b_fetched = odoo_inspector._execute_kw(
                        'account.payment',
                        'search_read',
                        [p_domain],
                        {
                            'fields': [
                                'date', 'name', 'ref', 'pos_order_id', 'journal_id',
                                'payment_method_line_id', 'partner_id', 'amount', 'state', 'is_reconciled'
                            ],
                            'order': 'journal_id, date, name'
                        }
                    )
                    for p in b_fetched:
                        if p['id'] not in seen_payment_ids:
                            seen_payment_ids.add(p['id'])
                            payments.append(p)
                    print(f"    - {b_name.upper()} ({len(b_dates)} dates): {len(b_fetched)} payments")
            else:
                p_domain = [
                    ('payment_type', '=', 'inbound'),
                ]
                if iso_from:
                    p_domain.append(('date', '>=', iso_from))
                if iso_to:
                    p_domain.append(('date', '<=', iso_to))
                if status:
                    st_list = [s.strip().lower() for s in status.split(',') if s.strip()]
                    if len(st_list) == 1:
                        p_domain.append(('state', '=', st_list[0]))
                    elif len(st_list) > 1:
                        p_domain.append(('state', 'in', st_list))

                if target_journal_names:
                    p_domain.append(('journal_id.name', 'in', target_journal_names))

                payments = odoo_inspector._execute_kw(

                    'account.payment',
                    'search_read',
                    [p_domain],
                    {
                        'fields': [
                            'date', 'name', 'ref', 'pos_order_id', 'journal_id',
                            'payment_method_line_id', 'partner_id', 'amount', 'state', 'is_reconciled'
                        ],
                        'order': 'journal_id, date, name'
                    }
                )

            print(f"    -> Total retrieved: {len(payments)} payment transactions.")

            # Group payments by journal -> is_reconciled
            groups = {}
            total_amount = 0.0
            for p in payments:
                j_name = p['journal_id'][1] if p.get('journal_id') else "Other Journal"
                is_rec = bool(p.get('is_reconciled', False))
                groups.setdefault(j_name, {}).setdefault(is_rec, []).append(p)
                total_amount += float(p.get('amount', 0.0))

            # Build openpyxl workbook
            wb_p = openpyxl.Workbook()
            ws_p = wb_p.active
            ws_p.title = "Sheet1"

            headers = [
                'Date', 'Number', 'Reference', 'Linked POS Order', 'Shop',
                'Journal', 'Payment Method', 'Customer/Vendor', 'Amount Signed',
                'Amount Company Currency Signed', 'Status'
            ]
            ws_p.append(headers)

            dt_from_obj = datetime.strptime(iso_from, "%Y-%m-%d")
            month_str = dt_from_obj.strftime("%B %Y")
            day_str = dt_from_obj.strftime("%d %b %Y")

            # Row 2: Month Group Header
            ws_p.append([f"{month_str} ({len(payments)})", None, None, None, None, None, None, None, total_amount, total_amount, None])
            # Row 3: Day Group Header (Read by odoo_reader.py from Cell A3)
            ws_p.append([f"    {day_str} ({len(payments)})", None, None, None, None, None, None, None, total_amount, total_amount, None])

            # Group rows & Data rows
            for j_name, rec_dict in groups.items():
                total_j_count = sum(len(items) for items in rec_dict.values())
                j_subtotal = sum(sum(float(p.get('amount', 0.0)) for p in items) for items in rec_dict.values())
                # Journal group header (Row 4 etc.)
                ws_p.append([f"        {j_name} ({total_j_count})", None, None, None, None, None, None, None, j_subtotal, j_subtotal, None])

                # Subgroups for True and False (Row 5 etc. - read by readers/odoo_reader.py for is_reconciled)
                for is_rec in [True, False]:
                    if is_rec in rec_dict:
                        sub_list = rec_dict[is_rec]
                        sub_total = sum(float(p.get('amount', 0.0)) for p in sub_list)
                        rec_label = "True" if is_rec else "False"
                        ws_p.append([f"            {rec_label} ({len(sub_list)})", None, None, None, None, None, None, None, sub_total, sub_total, None])

                        for p in sub_list:
                            p_date = datetime.strptime(p['date'], "%Y-%m-%d") if isinstance(p['date'], str) else p['date']
                            p_num = p.get('name') or ""
                            p_ref = p.get('ref') or ""
                            pos_order = p['pos_order_id'][1] if p.get('pos_order_id') else ""
                            shop = ""
                            method = p['payment_method_line_id'][1] if p.get('payment_method_line_id') else ""
                            partner = p['partner_id'][1] if p.get('partner_id') else ""
                            amt = float(p.get('amount', 0.0))
                            state = p.get('state', 'posted')

                            ws_p.append([
                                p_date, p_num, p_ref, pos_order, shop,
                                j_name, method, partner, amt, amt, state
                            ])

            ODO_EXCEL_PATH.parent.mkdir(parents=True, exist_ok=True)
            try:
                wb_p.save(ODO_EXCEL_PATH)
            except Exception:
                tmp = ODO_EXCEL_PATH.with_suffix(".tmp.xlsx")
                wb_p.save(tmp)
                os.replace(tmp, ODO_EXCEL_PATH)
            finally:
                wb_p.close()
            print(f"    -> Successfully saved to '{ODO_EXCEL_PATH.name}'.")

        # ── 2. Journal Entries (account.move) ──
        if mode in ("journals", "both", "auto_recon"):
            print(f"[+] Fetching Journal Entries via XML-RPC...")
            settlement_journals = list(set([j for j in [ODOO_JOURNAL_EDC, ODOO_JOURNAL_AR, 'EDC Settlement Journal', 'MDR Fees Journal'] if j]))
            m_domain = [
                ('move_type', '=', 'entry'),
                ('state', 'in', ['posted', 'draft']),
            ]
            if exact_dates:
                m_domain.append(('date', 'in', exact_dates))
            elif exact_dates_by_bank:
                all_b_dates = sorted(list(set(d for dlist in exact_dates_by_bank.values() for d in dlist)))
                if all_b_dates:
                    m_domain.append(('date', 'in', all_b_dates))
                else:
                    m_domain.extend([('date', '>=', iso_from), ('date', '<=', iso_to)])
            else:
                m_domain.extend([('date', '>=', iso_from), ('date', '<=', iso_to)])

            if settlement_journals:
                m_domain.append(('journal_id.name', 'in', settlement_journals))

            moves = odoo_inspector._execute_kw(
                'account.move',
                'search_read',
                [m_domain],
                {
                    'fields': ['date', 'name', 'partner_id', 'ref', 'journal_id', 'company_id', 'amount_total_signed', 'amount_total', 'state'],
                    'order': 'date desc, name desc',
                    'limit': 5000
                }
            )

            print(f"    -> Retrieved {len(moves)} settlement journal entries.")

            wb_j = openpyxl.Workbook()
            ws_j = wb_j.active
            ws_j.title = "Sheet1"

            j_headers = ['Date', 'Number', 'Partner', 'Reference', 'Journal', 'Company', 'Total Signed', 'Status']
            ws_j.append(j_headers)

            for m in moves:
                m_date = datetime.strptime(m['date'], "%Y-%m-%d") if isinstance(m['date'], str) else m['date']
                m_name = m.get('name') or ""
                m_partner = m['partner_id'][1] if m.get('partner_id') else ""
                m_ref = m.get('ref') or ""
                m_journal = m['journal_id'][1] if m.get('journal_id') else ""
                m_company = m['company_id'][1] if m.get('company_id') else "Eyerizz Eyewear"
                m_total = float(m.get('amount_total_signed') or m.get('amount_total') or 0.0)
                m_state = str(m.get('state', '')).capitalize()

                ws_j.append([m_date, m_name, m_partner, m_ref, m_journal, m_company, m_total, m_state])

            ODO_JOURNAL_EXCEL_PATH.parent.mkdir(parents=True, exist_ok=True)
            try:
                wb_j.save(ODO_JOURNAL_EXCEL_PATH)
            except Exception:
                tmp = ODO_JOURNAL_EXCEL_PATH.with_suffix(".tmp.xlsx")
                wb_j.save(tmp)
                os.replace(tmp, ODO_JOURNAL_EXCEL_PATH)
            finally:
                wb_j.close()
            print(f"    -> Successfully saved to '{ODO_JOURNAL_EXCEL_PATH.name}'.")

        print("[+] Direct XML-RPC download completed successfully in < 1 second! ✅\n")
        return True

    except Exception as e:
        print(f"[!] XML-RPC Downloader Error: {e}")
        return False


def run_downloader():
    # Force UTF-8 encoding for console output
    import io as _io
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    elif hasattr(sys.stdout, "buffer"):
        sys.stdout = _io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

    parser = argparse.ArgumentParser(description="Odoo Direct XML-RPC Payment Downloader")
    parser.add_argument("--date-from", type=str, default="", help="Format: MM/DD/YYYY")
    parser.add_argument("--date-to", type=str, default="", help="Format: MM/DD/YYYY")
    parser.add_argument("--email", type=str, default="", help="Odoo Email (optional)")
    parser.add_argument("--password", type=str, default="", help="Odoo Password (optional)")
    parser.add_argument("--banks", type=str, default="BCA,Mandiri,BRI", help="Comma separated list of banks")
    parser.add_argument("--status", type=str, default="Posted", help="Comma separated list of payment statuses")
    parser.add_argument("--mode", type=str, choices=["payments", "journals", "both", "auto_recon"], default="both", help="Download mode")
    args = parser.parse_args()

    selected_banks = [b.strip() for b in args.banks.split(",")] if args.banks else []
    is_all_banks = not selected_banks or any(b.lower() == "all" for b in selected_banks) or len(selected_banks) >= 3
    
    exact_dates_by_bank = None
    all_exact_dates = None
    user_custom_dates = bool(args.date_from and args.date_to)

    if user_custom_dates:
        print(f"[+] Using manually specified date range: {args.date_from} to {args.date_to}")
    else:
        try:
            from main import scan_bank_dates_detailed
            banks_to_scan = selected_banks if not is_all_banks else None
            scan_details = scan_bank_dates_detailed(banks_to_scan)
            if scan_details.get("min_date") and scan_details.get("max_date"):
                args.date_from = scan_details["min_date"]
                args.date_to = scan_details["max_date"]
                exact_dates_by_bank = scan_details.get("dates_by_bank")
                all_exact_dates = scan_details.get("all_dates")
                print(f"[+] Bank dates detected: {scan_details['min_date']} to {scan_details['max_date']}")
                if exact_dates_by_bank:
                    for b_k, b_d in exact_dates_by_bank.items():
                        if b_d:
                            print(f"    - {b_k.upper()}: {len(b_d)} date(s) ({', '.join(b_d)})")
        except Exception as e:
            print(f"[!] Auto-scan note: {e}")

    # ── Execute Direct XML-RPC ──
    if args.mode == "auto_recon":
        # 1. Download Payments
        success_pay = download_via_xmlrpc(
            date_from=args.date_from,
            date_to=args.date_to,
            banks=args.banks,
            status=args.status,
            mode="payments",
            exact_dates_by_bank=exact_dates_by_bank,
            exact_dates=all_exact_dates,
        )
        if not success_pay:
            print("[!] Failed downloading payments via XML-RPC.")
            sys.exit(1)

        # 2. Run Reconciliation Engine in-process
        print("\n[+] Menjalankan rekonsiliasi in-memory...")
        from main import run_reconciliation
        banks_for_recon = selected_banks if not is_all_banks else None
        report_file = run_reconciliation(banks=banks_for_recon, process_all=False, open_file=False)

        # 3. Extract dates for Journal Entries in-process
        print("\n[+] Mengekstrak tanggal jurnal dari hasil rekonsiliasi...")
        from journal_checker import get_distinct_dates_from_recap
        extracted_dates = get_distinct_dates_from_recap(report_file) if report_file else []

        if extracted_dates:
            j_from = min(extracted_dates)
            j_to = max(extracted_dates)
            print(f"    -> Rentang tanggal jurnal otomatis: {j_from} s.d {j_to} (Exact: {len(extracted_dates)} dates)")
            download_via_xmlrpc(
                date_from=j_from,
                date_to=j_to,
                banks=args.banks,
                mode="journals",
                exact_dates=extracted_dates,
            )
        else:
            download_via_xmlrpc(
                date_from=args.date_from,
                date_to=args.date_to,
                banks=args.banks,
                mode="journals",
                exact_dates=all_exact_dates,
            )

        # 4. Check draft journals in-process
        print("\n[+] Mengecek status Journal Entries di Odoo...")
        from journal_checker import check_journal_entries
        check_journal_entries(report_file)
        return

    elif args.mode in ("payments", "journals", "both"):
        success = download_via_xmlrpc(
            date_from=args.date_from,
            date_to=args.date_to,
            banks=args.banks,
            status=args.status,
            mode=args.mode,
            exact_dates_by_bank=exact_dates_by_bank,
            exact_dates=all_exact_dates,
        )
        if success:
            print("\n✅ Download Completed Successfully!\n")
            return
        else:
            sys.exit(1)


if __name__ == "__main__":
    run_downloader()
