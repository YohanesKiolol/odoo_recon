import argparse
import sys
import os
import subprocess
from pathlib import Path
from playwright.sync_api import sync_playwright

# Import configs
from config import (
    ODOO_URL, ODOO_DASHBOARD_URL, ODOO_PAYMENTS_URL, ODO_EXCEL_PATH,
    BANK_ACCOUNTS, ODOO_JOURNAL_ENTRIES_URL, ODO_JOURNAL_EXCEL_PATH,
    ODOO_JOURNAL_EDC, ODOO_JOURNAL_AR
)


import openpyxl
from datetime import datetime, date

import odoo_inspector
from odoo_inspector import _normalize_date_to_iso


def download_via_xmlrpc(
    date_from: str,
    date_to: str,
    banks: list[str] | str | None = None,
    status: str = "Posted",
    mode: str = "both"
) -> bool:
    """
    Download Customer Payments and Journal Entries directly via Odoo XML-RPC.
    Generates exact Excel formats matching Odoo web exports.
    """
    try:
        iso_from = _normalize_date_to_iso(date_from)
        iso_to = _normalize_date_to_iso(date_to)
        
        print(f"\n[+] Direct XML-RPC Downloader active ({iso_from} to {iso_to})...")

        # Resolve selected bank list
        if isinstance(banks, str):
            selected_banks = [b.strip() for b in banks.split(",") if b.strip()]
        elif isinstance(banks, list):
            selected_banks = [b.strip() for b in banks if b.strip()]
        else:
            selected_banks = []

        is_all_banks = not selected_banks or any(b.lower() == "all" for b in selected_banks) or len(selected_banks) >= 3

        # Resolve matching journal names from config
        target_journal_names = []
        if not is_all_banks and selected_banks:
            for b in selected_banks:
                b_upper = b.upper()
                if b_upper in BANK_ACCOUNTS:
                    target_journal_names.extend(BANK_ACCOUNTS[b_upper].get("odoo_groups", []))
                else:
                    target_journal_names.append(b)

        # ── 1. Customer Payments (account.payment) ──
        if mode in ("payments", "both", "auto_recon"):
            print(f"[+] Fetching Customer Payments via XML-RPC...")
            p_domain = [
                ('date', '>=', iso_from),
                ('date', '<=', iso_to),
                ('payment_type', '=', 'inbound'),
            ]
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

            print(f"    -> Retreived {len(payments)} payment transactions.")

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
                ('date', '>=', iso_from),
                ('date', '<=', iso_to),
                ('move_type', '=', 'entry'),
                ('state', 'in', ['posted', 'draft']),
            ]
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

            print(f"    -> Retreived {len(moves)} settlement journal entries.")

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
    # Force UTF-8 — Windows CP1252 can't encode ──, ✅, ❌ etc.
    import io as _io
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    elif hasattr(sys.stdout, "buffer"):
        sys.stdout = _io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

    parser = argparse.ArgumentParser(description="Odoo Payment Downloader")
    parser.add_argument("--date-from", type=str, default="08/01/2026", help="Format: MM/DD/YYYY")
    parser.add_argument("--date-to", type=str, default="08/01/2026", help="Format: MM/DD/YYYY")
    parser.add_argument("--email", type=str, default="", help="Odoo Email (opsional untuk auto-login)")
    parser.add_argument("--password", type=str, default="", help="Odoo Password (opsional untuk auto-login)")
    parser.add_argument("--banks", type=str, default="BCA,Mandiri,BRI", help="Comma separated list of banks")
    parser.add_argument("--status", type=str, default="Posted", help="Comma separated list of payment statuses (e.g. Posted,In Payment)")
    parser.add_argument("--mode", type=str, choices=["payments", "journals", "both", "auto_recon"], default="both", help="Pilih mode download (payments/journals/both/auto_recon)")
    parser.add_argument("--playwright", action="store_true", help="Force browser automation using Playwright instead of direct XML-RPC")
    args = parser.parse_args()

    selected_banks = [b.strip() for b in args.banks.split(",")] if args.banks else []
    is_all_banks = not selected_banks or any(b.lower() == "all" for b in selected_banks) or len(selected_banks) >= 3
    
    if args.mode == "auto_recon":
        print("\n[+] Mode Auto-Recon: Mendeteksi tanggal bank secara otomatis...")
        try:
            from main import scan_bank_date_range
            banks_to_scan = selected_banks if not is_all_banks else None
            dates_range = scan_bank_date_range(banks_to_scan)
            if dates_range:
                args.date_from, args.date_to = dates_range
                print(f"[+] Tanggal otomatis terdeteksi: {args.date_from} s.d {args.date_to}")
        except Exception as e:
            print(f"[!] Failed auto-scan ({e}).")

    # ── Attempt Direct XML-RPC First (Fast & Headless) ──
    if not args.playwright:
        print("[+] Trying Fast Direct XML-RPC export...")

        if args.mode == "auto_recon":
            # 1. Download Payments
            success_pay = download_via_xmlrpc(
                date_from=args.date_from,
                date_to=args.date_to,
                banks=args.banks,
                status=args.status,
                mode="payments"
            )
            if not success_pay:
                print("[!] Failed downloading payments via XML-RPC. Falling back to Playwright...")
            else:
                # 2. Run Reconciliation Engine in-process
                print("\n[+] Menjalankan rekonsiliasi in-memory...")
                from main import run_reconciliation
                banks_for_recon = selected_banks if not is_all_banks else None
                report_file = run_reconciliation(banks=banks_for_recon, process_all=False, open_file=False)

                # 3. Extract dates for Journal Entries in-process
                print(f"[+] Mengekstrak tanggal dari {report_file.name} untuk filter Journal...")
                from journal_checker import extract_journal_date_range, check_journals
                j_dates = extract_journal_date_range(report_file)
                if j_dates:
                    j_date_from, j_date_to = j_dates
                    print(f"[+] Tanggal Journal terdeteksi: {j_date_from} s.d {j_date_to}")
                else:
                    j_date_from, j_date_to = args.date_from, args.date_to

                # 4. Download Journal Entries via XML-RPC
                download_via_xmlrpc(
                    date_from=j_date_from,
                    date_to=j_date_to,
                    banks=args.banks,
                    status=args.status,
                    mode="journals"
                )

                # 5. Run Journal Checker in-process
                print("\n[+] Menjalankan Journal Checker...")
                check_journals(report_file, skip_download=True)

                print(f"\n[+] Opening file {report_file.name}...")
                import os
                if os.name == 'nt':
                    os.startfile(str(report_file))
                else:
                    import subprocess
                    subprocess.run(["open", str(report_file)])
                return

        else:
            success = download_via_xmlrpc(
                date_from=args.date_from,
                date_to=args.date_to,
                banks=args.banks,
                status=args.status,
                mode=args.mode
            )
            if success:
                print("\n✅ Download Completed Successfully!\n")
                return

        print("[!] XML-RPC direct download could not complete. Falling back to Playwright browser automation...\n")
            
    is_headless = bool(args.email and args.password)
    
    selected_banks = [b.strip() for b in args.banks.split(",")] if args.banks else []

    # Force all URLs to use https:// to prevent Nginx 404 errors on login redirects
    odoo_url_https = ODOO_URL.replace("http://", "https://") if ODOO_URL else ""
    odoo_dash_https = ODOO_DASHBOARD_URL.replace("http://", "https://") if ODOO_DASHBOARD_URL else ""
    odoo_pay_https = ODOO_PAYMENTS_URL.replace("http://", "https://") if ODOO_PAYMENTS_URL else ""

    if not odoo_url_https or not odoo_dash_https or not odoo_pay_https:
        print("[!] Error: ODOO_URL, ODOO_DASHBOARD_URL, atau ODOO_PAYMENTS_URL belum di-set di .env")
        sys.exit(1)

    print("[+] Mengecek dan menginstall browser (butuh waktu beberapa menit pada run pertama)...")
    try:
        import subprocess
        import os
        
        # Override PyInstaller's default behavior (which tries to use the temporary _MEI folder)
        # by forcing Playwright to use a persistent browser path in the user's home directory.
        browsers_path = str(Path.home() / ".playwright_browsers")
        os.environ["PLAYWRIGHT_BROWSERS_PATH"] = browsers_path
        
        flags = getattr(subprocess, "CREATE_NO_WINDOW", 0) if os.name == 'nt' else 0
        if getattr(sys, "frozen", False):
            from playwright._impl._driver import compute_driver_executable, get_driver_env
            node_exe, cli_js = compute_driver_executable()
            driver_env = get_driver_env()
            driver_env["PLAYWRIGHT_BROWSERS_PATH"] = browsers_path
            subprocess.run([node_exe, cli_js, "install", "chromium"], check=True, env=driver_env, creationflags=flags)
        else:
            subprocess.run([sys.executable, "-m", "playwright", "install", "chromium"], check=True, creationflags=flags)
    except Exception as e:
        print(f"[!] Failed to check/install browser: {e}")

    print("[+] Menyiapkan browser (Playwright)...")
    
    with sync_playwright() as p:
        user_data_dir = str(Path(__file__).parent / "playwright_profile")
        
        import platform
        if platform.system() == "Windows":
            ua = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
        else:
            ua = "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
        
        context = None
        for channel in ["chrome", "msedge", None]:
            try:
                kwargs = {
                    "user_data_dir": user_data_dir,
                    "headless": is_headless,
                    "args": ["--disable-blink-features=AutomationControlled", "--test-type"],
                    "ignore_default_args": ["--no-sandbox", "--enable-automation"],
                    "user_agent": ua,
                    "viewport": {"width": 1920, "height": 1080}
                }
                if channel:
                    kwargs["channel"] = channel
                context = p.chromium.launch_persistent_context(**kwargs)
                break
            except Exception:
                continue

        if not context:
            possible_paths = []
            if platform.system() == "Windows":
                possible_paths = [
                    r"C:\Program Files\Google\Chrome\Application\chrome.exe",
                    r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
                    r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
                    r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
                ]
            elif platform.system() == "Darwin":
                possible_paths = [
                    "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
                    "/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge",
                ]
            for exe in possible_paths:
                if os.path.exists(exe):
                    try:
                        context = p.chromium.launch_persistent_context(
                            user_data_dir=user_data_dir,
                            headless=is_headless,
                            executable_path=exe,
                            args=["--disable-blink-features=AutomationControlled", "--test-type"],
                            ignore_default_args=["--no-sandbox", "--enable-automation"],
                            user_agent=ua,
                            viewport={"width": 1920, "height": 1080}
                        )
                        break
                    except Exception:
                        continue

        if not context:
            print("❌ Failed to launch browser (Chrome or Edge). Please ensure Google Chrome or Microsoft Edge is installed.")
            sys.exit(1)


        page = context.pages[0] if context.pages else context.new_page()
        page.add_init_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
        
        # -------------------------------------------------------------
        # Fix for Odoo proxy_mode bug: Odoo randomly redirects to http:// after login.
        # Playwright network interceptors cannot change protocols on top-level navigation,
        # so we inject a script that forces the browser to instantly redirect to https://
        # if it ever lands on an http:// page (like the 404 Nginx page).
        page.add_init_script("""
            if (window.location.protocol === 'http:') {
                window.location.href = window.location.href.replace('http:', 'https:');
            }
        """)
        # -------------------------------------------------------------

        print(f"[+] Membuka {odoo_url_https}")
        page.goto(odoo_url_https)

        if is_headless:
            print("[+] Mengisi form login Odoo di background...")
            try:
                page.wait_for_selector("#login", state="visible", timeout=10000)
                page.fill("#login", args.email)
                page.fill("#password", args.password)
                page.press("#password", "Enter")
                print("[+] Login submitted, waiting...")
            except Exception:
                print("[+] Already logged in or login form skipped.")
        else:
            print("\n[+] Waiting for manual login...")

        print("[+] Detecting Odoo dashboard...")
        try:
            page.wait_for_selector(".o_main_navbar, .o_web_client, .o_action_manager", state="visible", timeout=45000)
        except Exception:
            pass
        print("[+] Dashboard terdeteksi!")


        # -------------------------------------------------------------
        # ── Define XPaths ──
        # -------------------------------------------------------------
        xpath_filter_btn = "xpath=/html/body/div[1]/div/div[1]/div/div[2]/div/div[2]/button"
        xpath_favorites_filter = "xpath=/html/body/div[1]/div/div[1]/div/div[2]/div/div[2]/div/div[3]/span"
        xpath_custom_filter = "xpath=/html/body/div[1]/div/div[1]/div/div[2]/div/div[2]/div/div[1]/span[10]"
        xpath_field = "./div/div/div/div[2]/div/div[1]/div[1]/div/div"
        xpath_operator = "./div/div/div/div[2]/div/div[1]/div[2]/select"
        xpath_from = "./div/div/div/div[2]/div/div[1]/div[3]/div/div[1]/input"
        xpath_to   = "./div/div/div/div[2]/div/div[1]/div[3]/div/div[2]/input"
        xpath_add = "./footer/button[1]"
        xpath_gear = "xpath=/html/body/div[1]/div/div[1]/div/div[1]/div[2]/div[2]/div"
        xpath_export_all = "xpath=/html/body/div[1]/div/div[1]/div/div[1]/div[2]/div[2]/div/div/div/span[2]"

        # -------------------------------------------------------------
        # ── Download Payments (account.payment) ──
        # -------------------------------------------------------------
        if args.mode in ["payments", "both", "auto_recon"]:
            print(f"\n[+] Redirecting to page Payments: {odoo_pay_https}")
            page.goto(odoo_pay_https)

            # Tunggu loading SPA Odoo secara dinamis berdasarkan elemen tabel data
            print("[+] Waiting for Odoo data table to fully load...")
            page.wait_for_selector(".o_list_table, .o_kanban_view, .o_content", state="visible", timeout=30000)
            page.wait_for_timeout(2000) # Tambahan jeda 2 detik agar event listener Odoo siap

            print("\n[+] Starting filter automation (hingga langkah 5)...")
            
            # Langkah 1: Click dropdown Filter
            print("[+] Opening Filter menu...")
            print("[+] Opening Filter menu...")
            page.locator(xpath_filter_btn).click()
            page.wait_for_timeout(1000)
            
            # Langkah 1.5: Click Favorites Filter
            print("[+] Selecting Favorites Filter...")
            print("[+] Selecting Favorites Filter...")
            page.wait_for_selector(xpath_favorites_filter, state="visible", timeout=5000)
            page.locator(xpath_favorites_filter).click()
            page.wait_for_timeout(2000) # Tunggu filter favorit diaplikasikan dan tabel reload

            # Langkah 2: Click Add Custom Filter
            print("[+] Selecting 'Add Custom Filter'...")
            print("[+] Selecting 'Add Custom Filter'...")
            page.wait_for_selector(xpath_custom_filter, state="visible", timeout=5000)
            page.locator(xpath_custom_filter).click()
            page.wait_for_timeout(1000) # Beri jeda agar popover muncul

            # Langkah 3: Tunggu Pop Up Modal (Di Odoo 17 ini adalah popover dengan tag <main>)
            print("[+] Waiting for Pop-up Modal...")
            dialog = page.locator("main").last
            dialog.wait_for(state="visible", timeout=10000)

            # Langkah 4: Pilih Field Date
            print("[+] Filling filter field with 'Date'...")
            print("[+] Filling filter field with 'Date'...")
            field_loc = dialog.locator(f"xpath={xpath_field}")
            field_loc.click()
            page.wait_for_timeout(500)
            
            page.keyboard.type("Date")
            page.wait_for_timeout(500)
            page.keyboard.press("Enter")

            # Langkah 5: Ganti separator menjadi 'is between'
            print("[+] Changing separator to 'is between'...")
            print("[+] Changing separator to 'is between'...")
            operator_loc = dialog.locator(f"xpath={xpath_operator}")
            operator_loc.wait_for(state="visible", timeout=5000)
            
            try:
                operator_loc.select_option(label="is between")
            except Exception as e:
                print(f"[!] Failed to select 'is between': {e}")
                print("[+] Trying click/arrow fallback...")
                operator_loc.click()
                page.keyboard.type("between")
                page.keyboard.press("Enter")
                
            # Langkah 6: Input Tanggal (From & To)
            print(f"[+] Filling date range: {args.date_from} to {args.date_to}...")
            print(f"[+] Filling date range: {args.date_from} to {args.date_to}...")
            
            from_loc = dialog.locator(f"xpath={xpath_from}")
            from_loc.wait_for(state="visible", timeout=5000)
            from_loc.fill(args.date_from)
            page.wait_for_timeout(500)
            
            to_loc = dialog.locator(f"xpath={xpath_to}")
            to_loc.wait_for(state="visible", timeout=5000)
            to_loc.fill(args.date_to)
            to_loc.press("Enter")
            page.wait_for_timeout(500)

            # Track condition row index (Date is row 2)
            condition_row_idx = 2

            # Langkah 6.5: Tambahkan filter Journal jika tidak semua bank dipilih
            is_all_banks = any(b.lower() == "all" for b in selected_banks) or len(selected_banks) >= 3
            if not is_all_banks and selected_banks:
                print(f"[+] Adding Journal filter for bank: {', '.join(selected_banks)}...")
                
                # Click '+' icon to add a new rule row
                print("    -> Clicking '+' (Add Condition)...")
                try:
                    # Odoo 17 uses i.fa-plus for the add condition button
                    add_btn = dialog.locator("i.fa-plus").last
                    add_btn.wait_for(state="visible", timeout=3000)
                    add_btn.click()
                except Exception as e:
                    print(f"    [!] Fallback clicking '+' via XPath: {e}")
                    xpath_add_rule = "./div/div/div/div[2]/div/div[2]/button[1]"
                    dialog.locator(f"xpath={xpath_add_rule}").click()
                    
                page.wait_for_timeout(1500) # Tunggu render baris baru
                condition_row_idx += 1
                
                # Ensure 'Match ALL' is selected
                print("    -> Memastikan rule matching diset ke 'all'...")
                xpath_match_btn = "./div/div/div/div[1]/div/div/div/button"
                try:
                    match_btn = dialog.locator(f"xpath={xpath_match_btn}")
                    match_btn.wait_for(state="visible", timeout=3000)
                    match_btn.click(timeout=3000)
                    page.wait_for_timeout(500)
                    
                    xpath_match_all = "./div/div/div/div[1]/div/div/div/div/span[1]"
                    dialog.locator(f"xpath={xpath_match_all}").click(timeout=3000)
                    page.wait_for_timeout(500)
                except Exception as e:
                    print(f"    [!] Skip klik Match ALL: {e}")
                
                # Select 'Journal' field
                xpath_field_2 = f"./div/div/div/div[{condition_row_idx}]/div/div[1]/div[1]/div/div"
                dialog.locator(f"xpath={xpath_field_2}").click()
                page.wait_for_timeout(500)
                page.keyboard.type("Journal")
                page.wait_for_timeout(500)
                page.keyboard.press("Enter")
                
                # Select operator 'is in'
                xpath_operator_2 = f"./div/div/div/div[{condition_row_idx}]/div/div[1]/div[2]/select"
                op_loc = dialog.locator(f"xpath={xpath_operator_2}")
                op_loc.wait_for(state="visible", timeout=5000)
                try:
                    op_loc.select_option(label="is in")
                except Exception as e:
                    print(f"[!] Failed to select 'is in': {e}")
                    op_loc.click()
                    page.keyboard.type("in")
                    page.keyboard.press("Enter")
                
                page.wait_for_timeout(500)
                
                # Input journal values
                xpath_value_input = f"./div/div/div/div[{condition_row_idx}]/div/div[1]/div[3]/div/div/input"
                val_input = dialog.locator(f"xpath={xpath_value_input}")
                
                journal_names = []
                for bank in selected_banks:
                    bk = bank.lower()
                    for alias, acc_info in BANK_ACCOUNTS.get(bk, {}).items():
                        grp = acc_info.get("group")
                        if grp:
                            journal_names.append(grp)
                            
                for j_name in journal_names:
                        print(f"    -> Memasukkan '{j_name}'")
                        val_input.click()
                        val_input.fill(j_name)
                        # Tunggu dropdown Odoo muncul
                        page.wait_for_timeout(1000)
                        
                        # Opsi dropdown biasanya di dalam popover UI
                        # Cari elemen <a> (opsi) yang textnya mengandung j_name lalu click
                        try:
                            dropdown_opt = page.locator("a", has_text=j_name).first
                            dropdown_opt.wait_for(state="visible", timeout=3000)
                            dropdown_opt.click()
                        except Exception:
                            print(f"    [!] Dropdown option not found for '{j_name}', trying Enter fallback")
                            page.keyboard.press("Enter")
                            
                        page.wait_for_timeout(500)

            # Langkah 6.6: Tambahkan filter Status jika ditentukan (default: Posted)
            if args.status:
                status_list = [s.strip() for s in args.status.split(",") if s.strip()]
                print(f"[+] Adding Status filter: {', '.join(status_list)}...")
                try:
                    # Click '+' icon to add a new rule row
                    print("    -> Clicking '+' (Add Condition for Status)...")
                    try:
                        add_btn = dialog.locator("i.fa-plus").last
                        add_btn.wait_for(state="visible", timeout=3000)
                        add_btn.click()
                    except Exception as e:
                        print(f"    [!] Fallback clicking '+' via XPath: {e}")
                        xpath_add_rule = "./div/div/div/div[2]/div/div[2]/button[1]"
                        dialog.locator(f"xpath={xpath_add_rule}").click()
                        
                    page.wait_for_timeout(1500)
                    condition_row_idx += 1
                    
                    # Ensure 'Match ALL' is selected
                    xpath_match_btn = "./div/div/div/div[1]/div/div/div/button"
                    try:
                        match_btn = dialog.locator(f"xpath={xpath_match_btn}")
                        if match_btn.is_visible():
                            match_btn.click(timeout=2000)
                            page.wait_for_timeout(500)
                            xpath_match_all = "./div/div/div/div[1]/div/div/div/div/span[1]"
                            dialog.locator(f"xpath={xpath_match_all}").click(timeout=2000)
                            page.wait_for_timeout(500)
                    except Exception:
                        pass
                    
                    # Select 'Status' field
                    xpath_field_status = f"./div/div/div/div[{condition_row_idx}]/div/div[1]/div[1]/div/div"
                    dialog.locator(f"xpath={xpath_field_status}").click()
                    page.wait_for_timeout(500)
                    page.keyboard.type("Status")
                    page.wait_for_timeout(500)
                    page.keyboard.press("Enter")
                    page.wait_for_timeout(500)
                    
                    # Select operator 'is in' / 'is equal to'
                    xpath_operator_status = f"./div/div/div/div[{condition_row_idx}]/div/div[1]/div[2]/select"
                    op_status = dialog.locator(f"xpath={xpath_operator_status}")
                    op_status.wait_for(state="visible", timeout=5000)
                    try:
                        op_status.select_option(label="is in")
                    except Exception:
                        try:
                            op_status.select_option(label="is equal to")
                        except Exception:
                            op_status.click()
                            page.keyboard.type("in")
                            page.keyboard.press("Enter")
                    page.wait_for_timeout(500)
                    
                    # Value selector/input
                    xpath_status_val_box = f"./div/div/div/div[{condition_row_idx}]/div/div[1]/div[3]"
                    val_box = dialog.locator(f"xpath={xpath_status_val_box}")
                    val_input = val_box.locator("input")
                    val_select = val_box.locator("select")
                    
                    if val_input.count() > 0:
                        for st in status_list:
                            print(f"    -> Memasukkan Status '{st}'")
                            val_input.first.click()
                            val_input.first.fill(st)
                            page.wait_for_timeout(1000)
                            try:
                                dropdown_opt = page.locator("a, .dropdown-item", has_text=st).first
                                dropdown_opt.wait_for(state="visible", timeout=3000)
                                dropdown_opt.click()
                            except Exception:
                                page.keyboard.press("Enter")
                            page.wait_for_timeout(500)
                    elif val_select.count() > 0:
                        for st in status_list:
                            try:
                                val_select.first.select_option(label=st)
                            except Exception:
                                val_select.first.click()
                                page.keyboard.type(st)
                                page.keyboard.press("Enter")
                except Exception as e:
                    print(f"    [!] Failed to add Status filter: {e}")

            # Langkah 7: Klik tombol 'Add'
            print("[+] Mengklik tombol 'Add'...")
            # Path absolut dari user: /html/body/div[2]/div[2]/div/div/div/div/footer/button[1]
            # Karena kita menggunakan locator `main`, footernya sejajar dengan main. Jadi kita cari tombol Add di parent popover.
            popover = dialog.locator("..")
            add_loc = popover.locator(f"xpath={xpath_add}")
            add_loc.click()
            
            print("[+] Menunggu tabel data Odoo termuat dengan filter baru...")
            page.wait_for_timeout(2000) # Tunggu animasi modal tutup
            page.wait_for_selector(".o_list_table, .o_kanban_view, .o_content", state="visible", timeout=30000)

            # Langkah 7.5: Tambahkan Group By "Is Reconciled"
            print("[+] Menambahkan Group By 'Is Reconciled'...")
            try:
                # Re-open the search bar dropdown
                page.locator(xpath_filter_btn).click()
                page.wait_for_timeout(1000)
                
                # Select 'Is Reconciled' from the native select dropdown directly
                group_select = page.locator("select.o_add_custom_group_menu").last
                group_select.wait_for(state="visible", timeout=3000)
                
                # Click the select box first ("Add Custom Group") so Odoo registers the interaction
                group_select.click()
                page.wait_for_timeout(500)
                
                # Then select 'Is Reconciled'
                group_select.select_option(value="is_reconciled")
                
                # In Odoo, selecting from the custom group dropdown usually triggers the reload immediately.
                page.wait_for_timeout(2000)
                print("    -> Group By Is Reconciled successful!")
            except Exception as e:
                print(f"    [!] Failed to add Group By: {e}")

            def export_odoo_data(p_page, out_path, desc_name="data"):
                print(f"[+] Opening Action menu (⚙️) untuk {desc_name}...")
                p_page.wait_for_timeout(1000)
                gear_found = False
                gear_selectors = [
                    ".o_cp_action_menus i.fa-cog",
                    ".o_cp_action_menus button",
                    "button.o_dropdown_toggle:has(i.fa-cog)",
                    "button:has(i.fa-cog)",
                    ".o_action_manager .o_cp_action_menus",
                    xpath_gear,
                ]
                for sel in gear_selectors:
                    try:
                        loc = p_page.locator(sel).first
                        if loc.is_visible():
                            loc.click()
                            p_page.wait_for_timeout(1000)
                            gear_found = True
                            break
                    except Exception:
                        continue
                if not gear_found:
                    try:
                        p_page.locator(xpath_gear).click()
                        p_page.wait_for_timeout(1000)
                    except Exception as e:
                        print(f"    [!] Error clicking gear: {e}")

                print(f"[+] Mengklik 'Export All' untuk {desc_name}...")
                export_selectors = [
                    "span.dropdown-item:has-text('Export All')",
                    ".dropdown-menu span:has-text('Export All')",
                    ".dropdown-item:has-text('Export All')",
                    "span:text-is('Export All')",
                    ".o_dropdown_item:has-text('Export All')",
                    xpath_export_all,
                ]
                with p_page.expect_download(timeout=60000) as dl_info:
                    clicked = False
                    for sel in export_selectors:
                        try:
                            loc = p_page.locator(sel).first
                            if loc.is_visible():
                                loc.click()
                                clicked = True
                                break
                        except Exception:
                            continue
                    if not clicked:
                        p_page.locator(xpath_export_all).click()
                        
                dl = dl_info.value
                out_path.parent.mkdir(parents=True, exist_ok=True)
                if out_path.exists():
                    try: out_path.unlink()
                    except Exception: pass
                dl.save_as(str(out_path))
                print(f"[+] {desc_name} downloaded successfully to: {out_path}")

            # Langkah 8 & 9: Export Payments
            export_odoo_data(page, ODO_EXCEL_PATH, desc_name="Payments")
            # -------------------------------------------------------------
            # ── Auto Recon Middle Step (Run main.py & extract journal dates) ──
            # -------------------------------------------------------------
            if args.mode == "auto_recon":
                import glob, subprocess, os
                env = os.environ.copy()
                env["PYTHONIOENCODING"] = "utf-8"
                env["PYTHONUTF8"] = "1"
                
                print("\n[+] Menjalankan rekonsiliasi (main.py)...")
                recon_cmd = [sys.executable, "--worker", "--no-open"] if getattr(sys, "frozen", False) else [sys.executable, "main.py", "--no-open"]
                if selected_banks and not is_all_banks:
                    recon_cmd.append("--bank")
                    recon_cmd.extend([b.lower() for b in selected_banks])
                        
                flags = getattr(subprocess, "CREATE_NO_WINDOW", 0) if os.name == 'nt' else 0
                result = subprocess.run(recon_cmd, capture_output=True, text=True, encoding="utf-8", errors="replace", env=env, creationflags=flags)
                print(result.stdout, end="")
                if result.returncode != 0:
                    print(f"[!] main.py failed (exit {result.returncode}):")
                    if result.stderr:
                        print(result.stderr, end="")
                    raise SystemExit(result.returncode)
                
                list_of_files = glob.glob('output/Reconciliation_*.xlsx')
                if not list_of_files:
                    print("[!] Failed to find Reconciliation file!")
                    context.close()
                    sys.exit(1)
                    
                latest_file = max(list_of_files, key=os.path.getctime)
                print(f"[+] Mengekstrak tanggal dari {latest_file} untuk filter Journal...")
                
                try:
                    flags = getattr(subprocess, "CREATE_NO_WINDOW", 0) if os.name == 'nt' else 0
                    date_res = subprocess.run([sys.executable, "journal_checker.py", latest_file, "--get-dates"], capture_output=True, text=True, encoding="utf-8", errors="replace", env=env, creationflags=flags)
                    for line in date_res.stdout.splitlines():
                        if line.startswith("[DATE_RANGE]|"):
                            parts = line.split("|")
                            if len(parts) == 3:
                                args.date_from = parts[1]
                                args.date_to = parts[2]
                                print(f"[+] Tanggal Journal terdeteksi: {args.date_from} s.d {args.date_to}")
                except Exception as e:
                    print(f"[!] Failed to extract Journal date ({e}).")
                    
                print("\n[+] Melanjutkan download Journal Entries di browser yang sama...")

        # -------------------------------------------------------------
        # ── Download Journal Entries (account.move) ──
        # -------------------------------------------------------------
        if args.mode in ["journals", "both", "auto_recon"] and ODOO_JOURNAL_ENTRIES_URL:
            odoo_journal_https = ODOO_JOURNAL_ENTRIES_URL.replace("http://", "https://")
            print(f"\n[+] Navigasi ke halaman Journal Entries: {odoo_journal_https}")
            page.goto(odoo_journal_https)
            
            # Tunggu redirect dan load table
            page.wait_for_selector(".o_list_table, .o_kanban_view, .o_content", state="visible", timeout=45000)
            page.wait_for_timeout(2000)

            
            print("[+] Membersihkan filter default (misal: Posted)...")
            try:
                remove_btns = page.locator("button.o_facet_remove")
                count = remove_btns.count()
                for _ in range(count):
                    remove_btns.first.click(timeout=2000)
                    page.wait_for_timeout(500)
            except Exception as e:
                print(f"    [!] Failed to clear filter: {e}")
            
            print("[+] Opening Filter menu for Journal Entries...")
            page.locator(xpath_filter_btn).click()
            page.wait_for_timeout(1000)
            
            print("[+] Selecting 'Add Custom Filter'...")
            page.locator(xpath_custom_filter).click()
            page.wait_for_timeout(1000)
            
            print("[+] Mengatur filter Date...")
            dialog_j = page.locator("main").last
            dialog_j.wait_for(state="visible", timeout=10000)
            
            field_loc_j = dialog_j.locator(f"xpath={xpath_field}")
            field_loc_j.click()
            page.wait_for_timeout(500)
            page.keyboard.type("Date")
            page.wait_for_timeout(500)
            page.keyboard.press("Enter")
            
            operator_loc_j = dialog_j.locator(f"xpath={xpath_operator}")
            operator_loc_j.wait_for(state="visible", timeout=5000)
            try:
                operator_loc_j.select_option(label="is between")
            except:
                operator_loc_j.click()
                page.keyboard.type("between")
                page.keyboard.press("Enter")
                
            print(f"[+] Filling date range: {args.date_from} to {args.date_to}...")
            from_loc_j = dialog_j.locator(f"xpath={xpath_from}")
            from_loc_j.wait_for(state="visible", timeout=5000)
            from_loc_j.fill(args.date_from)
            page.wait_for_timeout(500)
            
            to_loc_j = dialog_j.locator(f"xpath={xpath_to}")
            to_loc_j.wait_for(state="visible", timeout=5000)
            to_loc_j.fill(args.date_to)
            to_loc_j.press("Enter")
            page.wait_for_timeout(500)
            
            print("[+] Menambahkan filter Journal (EDC & AR)...")
            try:
                add_btn_j = dialog_j.locator("i.fa-plus").last
                add_btn_j.wait_for(state="visible", timeout=3000)
                add_btn_j.click()
            except:
                xpath_add_rule_j = "./div/div/div/div[2]/div/div[2]/button[1]"
                dialog_j.locator(f"xpath={xpath_add_rule_j}").click()
                
            page.wait_for_timeout(1500)
            
            # Ensure Match ALL is selected
            xpath_match_btn_j = "./div/div/div/div[1]/div/div/div/button"
            try:
                match_btn_j = dialog_j.locator(f"xpath={xpath_match_btn_j}")
                match_btn_j.wait_for(state="visible", timeout=3000)
                match_btn_j.click(timeout=3000)
                page.wait_for_timeout(500)
                xpath_match_all_j = "./div/div/div/div[1]/div/div/div/div/span[1]"
                dialog_j.locator(f"xpath={xpath_match_all_j}").click(timeout=3000)
                page.wait_for_timeout(500)
            except:
                pass
                
            # Select Journal field
            xpath_field_2_j = "./div/div/div/div[3]/div/div[1]/div[1]/div/div"
            dialog_j.locator(f"xpath={xpath_field_2_j}").click()
            page.wait_for_timeout(500)
            page.keyboard.type("Journal")
            page.wait_for_timeout(500)
            page.keyboard.press("Enter")
            
            # Select operator 'is in'
            xpath_operator_2_j = "./div/div/div/div[3]/div/div[1]/div[2]/select"
            op_loc_j = dialog_j.locator(f"xpath={xpath_operator_2_j}")
            op_loc_j.wait_for(state="visible", timeout=5000)
            try:
                op_loc_j.select_option(label="is in")
            except:
                op_loc_j.click()
                page.keyboard.type("in")
                page.keyboard.press("Enter")
                
            page.wait_for_timeout(500)
            
            # Input journal values
            xpath_value_input_j = "./div/div/div/div[3]/div/div[1]/div[3]/div/div/input"
            val_input_j = dialog_j.locator(f"xpath={xpath_value_input_j}")
            
            for j_name in [ODOO_JOURNAL_EDC, ODOO_JOURNAL_AR]:
                if j_name:
                    print(f"    -> Memasukkan '{j_name}'")
                    val_input_j.click()
                    val_input_j.fill(j_name)
                    page.wait_for_timeout(1000)
                    try:
                        dropdown_opt = page.locator("a", has_text=j_name).first
                        dropdown_opt.wait_for(state="visible", timeout=3000)
                        dropdown_opt.click()
                    except:
                        page.keyboard.press("Enter")
                    page.wait_for_timeout(500)
            
            print("[+] Mengklik tombol 'Add'...")
            popover_j = dialog_j.locator("..")
            add_loc_j = popover_j.locator(f"xpath={xpath_add}")
            add_loc_j.click()
            
            page.wait_for_timeout(2000)
            page.wait_for_selector(".o_list_table, .o_kanban_view, .o_content", state="visible", timeout=30000)
            
            export_odoo_data(page, ODO_JOURNAL_EXCEL_PATH, desc_name="Journal Entries")

        context.close()
        
        # -------------------------------------------------------------
        # ── Auto Recon Final Step (Run journal checker) ──
        # -------------------------------------------------------------
        if args.mode == "auto_recon":
            print("\n[+] Menjalankan Journal Checker...")
            # Gunakan text=True, encoding='utf-8' atau jalankan dengan env
            env = os.environ.copy()
            env["PYTHONIOENCODING"] = "utf-8"
            env["PYTHONUTF8"] = "1"
            flags = getattr(subprocess, "CREATE_NO_WINDOW", 0) if os.name == 'nt' else 0
            subprocess.run([sys.executable, "journal_checker.py", latest_file, "--skip-download"], check=True, env=env, creationflags=flags)
            
            print(f"\n[+] Opening file {latest_file}...")
            if os.name == 'nt':
                os.startfile(latest_file)
            else:
                subprocess.run(["open", latest_file])

if __name__ == "__main__":
    run_downloader()
