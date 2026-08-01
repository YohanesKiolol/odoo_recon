import argparse
import sys
from pathlib import Path
from playwright.sync_api import sync_playwright

# Import configs
from config import ODOO_URL, ODOO_DASHBOARD_URL, ODOO_PAYMENTS_URL, ODO_EXCEL_PATH


def run_downloader():
    parser = argparse.ArgumentParser(description="Odoo Payment Downloader")
    parser.add_argument("--date-from", type=str, default="08/01/2026", help="Format: MM/DD/YYYY")
    parser.add_argument("--date-to", type=str, default="08/01/2026", help="Format: MM/DD/YYYY")
    args = parser.parse_args()

    if not ODOO_URL or not ODOO_DASHBOARD_URL or not ODOO_PAYMENTS_URL:
        print("❌ Error: ODOO_URL, ODOO_DASHBOARD_URL, atau ODOO_PAYMENTS_URL belum di-set di .env")
        sys.exit(1)

    print("🚀 Menyiapkan browser (Playwright)...")
    
    with sync_playwright() as p:
        user_data_dir = str(Path(__file__).parent / "playwright_profile")
        
        try:
            context = p.chromium.launch_persistent_context(
                user_data_dir=user_data_dir,
                headless=False,
                channel="chrome",
                args=["--disable-blink-features=AutomationControlled", "--test-type"],
                ignore_default_args=["--no-sandbox", "--enable-automation"],
                user_agent="Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
            )
        except Exception:
            context = p.chromium.launch_persistent_context(
                user_data_dir=user_data_dir,
                headless=False,
                args=["--disable-blink-features=AutomationControlled", "--test-type"],
                ignore_default_args=["--no-sandbox", "--enable-automation"],
                user_agent="Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
            )

        page = context.pages[0] if context.pages else context.new_page()
        page.add_init_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")

        # -------------------------------------------------------------
        # Fix 404 Nginx Loop by upgrading HTTP to HTTPS automatically
        # -------------------------------------------------------------
        def upgrade_to_https(route, request):
            if request.url.startswith("http://"):
                secure_url = request.url.replace("http://", "https://", 1)
                route.fulfill(status=301, headers={"Location": secure_url})
            else:
                route.continue_()

        page.route("**/*", upgrade_to_https)
        # -------------------------------------------------------------

        print(f"🌐 Membuka {ODOO_URL}")
        page.goto(ODOO_URL)

        print("\n⏳ Menunggu Anda login secara manual...")
        page.wait_for_function(
            "dashboardUrl => window.location.href.includes(dashboardUrl)",
            arg=ODOO_DASHBOARD_URL,
            timeout=0
        )
        print("✅ Dashboard terdeteksi!")

        print(f"\n➡️ Mengarahkan ke halaman Payments: {ODOO_PAYMENTS_URL}")
        page.goto(ODOO_PAYMENTS_URL)

        # Tunggu loading SPA Odoo secara dinamis berdasarkan elemen tabel data
        print("  ⏳ Menunggu tabel data Odoo termuat sepenuhnya...")
        page.wait_for_selector(".o_list_table, .o_kanban_view, .o_content", state="visible", timeout=30000)
        page.wait_for_timeout(2000) # Tambahan jeda 2 detik agar event listener Odoo siap

        print("\n🎉 Memulai otomatisasi filter (hingga langkah 5)...")
        
        # Langkah 1: Click dropdown Filter
        print("  1. Membuka menu Filter...")
        xpath_filter_btn = "xpath=/html/body/div[1]/div/div[1]/div/div[2]/div/div[2]/button"
        page.wait_for_selector(xpath_filter_btn, state="visible", timeout=15000)
        page.locator(xpath_filter_btn).click()
        page.wait_for_timeout(1000)
        
        # Langkah 1.5: Click Favorites Filter
        print("  1.5. Memilih Favorites Filter...")
        xpath_favorites_filter = "xpath=/html/body/div[1]/div/div[1]/div/div[2]/div/div[2]/div/div[3]/span"
        page.wait_for_selector(xpath_favorites_filter, state="visible", timeout=5000)
        page.locator(xpath_favorites_filter).click()
        page.wait_for_timeout(2000) # Tunggu filter favorit diaplikasikan dan tabel reload

        # Langkah 2: Click Add Custom Filter
        print("  2. Memilih 'Add Custom Filter'...")
        xpath_custom_filter = "xpath=/html/body/div[1]/div/div[1]/div/div[2]/div/div[2]/div/div[1]/span[10]"
        page.wait_for_selector(xpath_custom_filter, state="visible", timeout=5000)
        page.locator(xpath_custom_filter).click()
        page.wait_for_timeout(1000) # Beri jeda agar popover muncul

        # Langkah 3: Tunggu Pop Up Modal (Di Odoo 17 ini adalah popover dengan tag <main>)
        print("  3. Menunggu Pop-up Modal...")
        dialog = page.locator("main").last
        dialog.wait_for(state="visible", timeout=10000)

        # Langkah 4: Pilih Field Date
        print("  4. Mengisi field filter dengan 'Date'...")
        # Path relatif dari tag <main>
        xpath_field = "./div/div/div/div[2]/div/div[1]/div[1]/div/div"
        
        field_loc = dialog.locator(f"xpath={xpath_field}")
        field_loc.click()
        page.wait_for_timeout(500)
        
        page.keyboard.type("Date")
        page.wait_for_timeout(500)
        page.keyboard.press("Enter")

        # Langkah 5: Ganti separator menjadi 'is between'
        print("  5. Mengganti separator menjadi 'is between'...")
        xpath_operator = "./div/div/div/div[2]/div/div[1]/div[2]/select"
        operator_loc = dialog.locator(f"xpath={xpath_operator}")
        operator_loc.wait_for(state="visible", timeout=5000)
        
        try:
            operator_loc.select_option(label="is between")
        except Exception as e:
            print(f"    ⚠️ Gagal memilih 'is between' dengan label: {e}")
            print("    Mencoba fallback klik dan keyboard arrow...")
            operator_loc.click()
            page.keyboard.type("between")
            page.keyboard.press("Enter")
            
        # Langkah 6: Input Tanggal (From & To)
        print(f"  6. Mengisi rentang tanggal: {args.date_from} sampai {args.date_to}...")
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
        page.wait_for_timeout(500)

        # Langkah 7: Klik tombol 'Add'
        print("  7. Mengklik tombol 'Add'...")
        # Path absolut dari user: /html/body/div[2]/div[2]/div/div/div/div/footer/button[1]
        # Karena kita menggunakan locator `main`, footernya sejajar dengan main. Jadi kita cari tombol Add di parent popover.
        popover = dialog.locator("..")
        xpath_add = "./footer/button[1]"
        add_loc = popover.locator(f"xpath={xpath_add}")
        add_loc.click()
        
        print("  ⏳ Menunggu tabel data Odoo termuat dengan filter baru...")
        page.wait_for_timeout(2000) # Tunggu animasi modal tutup
        page.wait_for_selector(".o_list_table, .o_kanban_view, .o_content", state="visible", timeout=30000)

        # Langkah 8: Klik Icon Gear (Action Menu)
        print("  8. Membuka menu Action (⚙️)...")
        xpath_gear = "xpath=/html/body/div[1]/div/div[1]/div/div[1]/div[2]/div[2]/div"
        page.wait_for_selector(xpath_gear, state="visible", timeout=10000)
        page.locator(xpath_gear).click()
        page.wait_for_timeout(1000) # Tunggu menu dropdown terbuka

        # Langkah 9: Klik 'Export All' dan Tangkap Download
        print("  9. Mengklik 'Export All' dan mendownload file...")
        xpath_export_all = "xpath=/html/body/div[1]/div/div[1]/div/div[1]/div[2]/div[2]/div/div/div/span[2]"
        
        # Bersiap menangkap event download dari browser
        with page.expect_download(timeout=60000) as download_info:
            page.locator(xpath_export_all).click()
            
        download = download_info.value
        
        # Pastikan direktori tujuan ada
        ODO_EXCEL_PATH.parent.mkdir(parents=True, exist_ok=True)
        
        # Simpan file yang didownload ke lokasi yang ditentukan di .env
        download.save_as(str(ODO_EXCEL_PATH))
        print(f"✅ File berhasil didownload dan disimpan ke: {ODO_EXCEL_PATH}")

        context.close()


if __name__ == "__main__":
    run_downloader()
