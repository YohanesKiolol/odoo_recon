import argparse
import sys
from pathlib import Path
from playwright.sync_api import sync_playwright

# Import configs
from config import (
    ODOO_URL, ODOO_DASHBOARD_URL, ODOO_PAYMENTS_URL, ODO_EXCEL_PATH,
    BANK_ACCOUNTS
)


def run_downloader():
    parser = argparse.ArgumentParser(description="Odoo Payment Downloader")
    parser.add_argument("--date-from", type=str, default="08/01/2026", help="Format: MM/DD/YYYY")
    parser.add_argument("--date-to", type=str, default="08/01/2026", help="Format: MM/DD/YYYY")
    parser.add_argument("--email", type=str, default="", help="Odoo Email (opsional untuk auto-login)")
    parser.add_argument("--password", type=str, default="", help="Odoo Password (opsional untuk auto-login)")
    parser.add_argument("--banks", type=str, default="BCA,Mandiri,BRI", help="Comma separated list of banks")
    args = parser.parse_args()
    
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
        subprocess.run([sys.executable, "-m", "playwright", "install", "chromium"], check=True)
    except Exception as e:
        print(f"[!] Gagal mengecek/menginstall browser: {e}")

    print("[+] Menyiapkan browser (Playwright)...")
    
    with sync_playwright() as p:
        user_data_dir = str(Path(__file__).parent / "playwright_profile")
        
        try:
            context = p.chromium.launch_persistent_context(
                user_data_dir=user_data_dir,
                headless=is_headless,
                channel="chrome",
                args=["--disable-blink-features=AutomationControlled", "--test-type"],
                ignore_default_args=["--no-sandbox", "--enable-automation"],
                user_agent="Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
            )
        except Exception:
            context = p.chromium.launch_persistent_context(
                user_data_dir=user_data_dir,
                headless=is_headless,
                args=["--disable-blink-features=AutomationControlled", "--test-type"],
                ignore_default_args=["--no-sandbox", "--enable-automation"],
                user_agent="Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
            )

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
            page.wait_for_selector("xpath=/html/body/div[2]/main/div[1]/form/div[1]/input", state="visible", timeout=30000)
            page.fill("xpath=/html/body/div[2]/main/div[1]/form/div[1]/input", args.email)
            page.fill("xpath=/html/body/div[2]/main/div[1]/form/div[2]/input", args.password)
            page.click("xpath=/html/body/div[2]/main/div[1]/form/div[3]/button")
            print("[+] Login disubmit, menunggu proses...")
        else:
            print("\n[+] Menunggu Anda login secara manual...")
        
        # Strip protocol so it matches exactly
        dashboard_path = odoo_dash_https.split("://")[-1]
        
        page.wait_for_function(
            "path => window.location.href.includes(path)",
            arg=dashboard_path,
            timeout=0 if not is_headless else 30000
        )
        print("[+] Dashboard terdeteksi!")

        print(f"\n[+] Mengarahkan ke halaman Payments: {odoo_pay_https}")
        page.goto(odoo_pay_https)

        # Tunggu loading SPA Odoo secara dinamis berdasarkan elemen tabel data
        print("[+] Menunggu tabel data Odoo termuat sepenuhnya...")
        page.wait_for_selector(".o_list_table, .o_kanban_view, .o_content", state="visible", timeout=30000)
        page.wait_for_timeout(2000) # Tambahan jeda 2 detik agar event listener Odoo siap

        print("\n[+] Memulai otomatisasi filter (hingga langkah 5)...")
        
        # Langkah 1: Click dropdown Filter
        print("[+] Membuka menu Filter...")
        xpath_filter_btn = "xpath=/html/body/div[1]/div/div[1]/div/div[2]/div/div[2]/button"
        page.wait_for_selector(xpath_filter_btn, state="visible", timeout=15000)
        page.locator(xpath_filter_btn).click()
        page.wait_for_timeout(1000)
        
        # Langkah 1.5: Click Favorites Filter
        print("[+] Memilih Favorites Filter...")
        xpath_favorites_filter = "xpath=/html/body/div[1]/div/div[1]/div/div[2]/div/div[2]/div/div[3]/span"
        page.wait_for_selector(xpath_favorites_filter, state="visible", timeout=5000)
        page.locator(xpath_favorites_filter).click()
        page.wait_for_timeout(2000) # Tunggu filter favorit diaplikasikan dan tabel reload

        # Langkah 2: Click Add Custom Filter
        print("[+] Memilih 'Add Custom Filter'...")
        xpath_custom_filter = "xpath=/html/body/div[1]/div/div[1]/div/div[2]/div/div[2]/div/div[1]/span[10]"
        page.wait_for_selector(xpath_custom_filter, state="visible", timeout=5000)
        page.locator(xpath_custom_filter).click()
        page.wait_for_timeout(1000) # Beri jeda agar popover muncul

        # Langkah 3: Tunggu Pop Up Modal (Di Odoo 17 ini adalah popover dengan tag <main>)
        print("[+] Menunggu Pop-up Modal...")
        dialog = page.locator("main").last
        dialog.wait_for(state="visible", timeout=10000)

        # Langkah 4: Pilih Field Date
        print("[+] Mengisi field filter dengan 'Date'...")
        # Path relatif dari tag <main>
        xpath_field = "./div/div/div/div[2]/div/div[1]/div[1]/div/div"
        
        field_loc = dialog.locator(f"xpath={xpath_field}")
        field_loc.click()
        page.wait_for_timeout(500)
        
        page.keyboard.type("Date")
        page.wait_for_timeout(500)
        page.keyboard.press("Enter")

        # Langkah 5: Ganti separator menjadi 'is between'
        print("[+] Mengganti separator menjadi 'is between'...")
        xpath_operator = "./div/div/div/div[2]/div/div[1]/div[2]/select"
        operator_loc = dialog.locator(f"xpath={xpath_operator}")
        operator_loc.wait_for(state="visible", timeout=5000)
        
        try:
            operator_loc.select_option(label="is between")
        except Exception as e:
            print(f"[!] Gagal memilih 'is between' dengan label: {e}")
            print("[+] Mencoba fallback klik dan keyboard arrow...")
            operator_loc.click()
            page.keyboard.type("between")
            page.keyboard.press("Enter")
            
        # Langkah 6: Input Tanggal (From & To)
        print(f"[+] Mengisi rentang tanggal: {args.date_from} sampai {args.date_to}...")
        # Mengubah path menjadi relatif terhadap <main> agar aman dari perubahan DOM
        xpath_from = "./div/div/div/div[2]/div/div[1]/div[3]/div/div[1]/input"
        xpath_to   = "./div/div/div/div[2]/div/div[1]/div[3]/div/div[2]/input"
        
        from_loc = dialog.locator(f"xpath={xpath_from}")
        from_loc.wait_for(state="visible", timeout=5000)
        from_loc.fill(args.date_from)
        page.wait_for_timeout(500)
        
        to_loc = dialog.locator(f"xpath={xpath_to}")
        to_loc.wait_for(state="visible", timeout=5000)
        to_loc.fill(args.date_to)
        to_loc.press("Enter")
        page.wait_for_timeout(500)

        # Langkah 6.5: Tambahkan filter Journal jika tidak semua bank dipilih
        is_all_banks = any(b.lower() == "all" for b in selected_banks) or len(selected_banks) >= 3
        if not is_all_banks and selected_banks:
            print(f"[+] Menambahkan filter Journal untuk bank: {', '.join(selected_banks)}...")
            
            # Click '+' icon to add a new rule row
            print("    -> Meng-klik tombol '+' (Add Condition)...")
            try:
                # Odoo 17 uses i.fa-plus for the add condition button
                add_btn = dialog.locator("i.fa-plus").last
                add_btn.wait_for(state="visible", timeout=3000)
                add_btn.click()
            except Exception as e:
                print(f"    [!] Fallback klik '+' menggunakan XPath: {e}")
                xpath_add_rule = "./div/div/div/div[2]/div/div[2]/button[1]"
                dialog.locator(f"xpath={xpath_add_rule}").click()
                
            page.wait_for_timeout(1500) # Tunggu render baris baru
            
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
            xpath_field_2 = "./div/div/div/div[3]/div/div[1]/div[1]/div/div"
            dialog.locator(f"xpath={xpath_field_2}").click()
            page.wait_for_timeout(500)
            page.keyboard.type("Journal")
            page.wait_for_timeout(500)
            page.keyboard.press("Enter")
            
            # Select operator 'is in'
            xpath_operator_2 = "./div/div/div/div[3]/div/div[1]/div[2]/select"
            op_loc = dialog.locator(f"xpath={xpath_operator_2}")
            op_loc.wait_for(state="visible", timeout=5000)
            try:
                op_loc.select_option(label="is in")
            except Exception as e:
                print(f"[!] Gagal memilih 'is in' via select_option: {e}")
                op_loc.click()
                page.keyboard.type("in")
                page.keyboard.press("Enter")
            
            page.wait_for_timeout(500)
            
            # Input journal values
            xpath_value_input = "./div/div/div/div[3]/div/div[1]/div[3]/div/div/input"
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
                        print(f"    [!] Tidak menemukan opsi dropdown untuk '{j_name}', mencoba fallback Enter")
                        page.keyboard.press("Enter")
                        
                    page.wait_for_timeout(500)

        # Langkah 7: Klik tombol 'Add'
        print("[+] Mengklik tombol 'Add'...")
        # Path absolut dari user: /html/body/div[2]/div[2]/div/div/div/div/footer/button[1]
        # Karena kita menggunakan locator `main`, footernya sejajar dengan main. Jadi kita cari tombol Add di parent popover.
        popover = dialog.locator("..")
        xpath_add = "./footer/button[1]"
        add_loc = popover.locator(f"xpath={xpath_add}")
        add_loc.click()
        
        print("[+] Menunggu tabel data Odoo termuat dengan filter baru...")
        page.wait_for_timeout(2000) # Tunggu animasi modal tutup
        page.wait_for_selector(".o_list_table, .o_kanban_view, .o_content", state="visible", timeout=30000)

        # Langkah 8: Klik Icon Gear (Action Menu)
        print("[+] Membuka menu Action (⚙️)...")
        xpath_gear = "xpath=/html/body/div[1]/div/div[1]/div/div[1]/div[2]/div[2]/div"
        page.wait_for_selector(xpath_gear, state="visible", timeout=10000)
        page.locator(xpath_gear).click()
        page.wait_for_timeout(1000) # Tunggu menu dropdown terbuka

        # Langkah 9: Klik 'Export All' dan Tangkap Download
        print("[+] Mengklik 'Export All' dan mendownload file...")
        xpath_export_all = "xpath=/html/body/div[1]/div/div[1]/div/div[1]/div[2]/div[2]/div/div/div/span[2]"
        
        # Bersiap menangkap event download dari browser
        with page.expect_download(timeout=60000) as download_info:
            page.locator(xpath_export_all).click()
            
        download = download_info.value
        
        # Pastikan direktori tujuan ada
        ODO_EXCEL_PATH.parent.mkdir(parents=True, exist_ok=True)
        
        # Simpan file yang didownload ke lokasi yang ditentukan di .env
        download.save_as(str(ODO_EXCEL_PATH))
        print(f"[+] File berhasil didownload dan disimpan ke: {ODO_EXCEL_PATH}")

        context.close()


if __name__ == "__main__":
    run_downloader()
