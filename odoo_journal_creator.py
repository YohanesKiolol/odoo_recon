"""
odoo_journal_creator.py — Script to automate Odoo journal creation.
Phase 1: Testing and Navigation only.
"""

import os
import sys
import argparse
from pathlib import Path
from openpyxl import load_workbook
from playwright.sync_api import sync_playwright

# Import configs
from config import (
    ODOO_URL, ODOO_DASHBOARD_URL, ODOO_JOURNAL_CREATE_URL, OUTPUT_DIR
)

# Konfigurasi format referensi jurnal per bank
REF_FORMATS = {
    "BCA": "Settlement Journal EDC BCA for {tanggal}",
    "Mandiri": "Settlement Journal EDC MANDIRI for {tanggal}",
    "BRI": "Settlement Journal EDC BRI for {tanggal}"
}


def get_latest_excel_file() -> Path:
    """Find the most recently created .xlsx file in the output directory."""
    excel_files = list(OUTPUT_DIR.glob("reconciliation_*.xlsx"))
    if not excel_files:
        print(f"❌ Tidak ada file Excel yang ditemukan di '{OUTPUT_DIR}'.")
        sys.exit(1)
    # Sort by modification time, newest last
    latest_file = max(excel_files, key=os.path.getmtime)
    return latest_file


def read_sesuai_dates(filepath: Path) -> list[dict]:
    """Read 'Ringkasan Harian' and extract rows with 'Sesuai' status."""
    wb = load_workbook(filepath, data_only=True)
    if "Ringkasan Harian" not in wb.sheetnames:
        print("❌ Sheet 'Ringkasan Harian' tidak ditemukan di file Excel.")
        sys.exit(1)

    ws = wb["Ringkasan Harian"]
    results = []

    # Start reading from row 4 (where data begins)
    for row in range(4, ws.max_row + 1):
        bank = ws.cell(row=row, column=2).value
        tanggal = ws.cell(row=row, column=3).value
        status = ws.cell(row=row, column=7).value

        if not bank or not status:
            continue

        if "Sesuai" in str(status):
            results.append({
                "bank": bank,
                "tanggal": tanggal,
                "status": status
            })

    return results


def run_automation(selected_items: list[dict]):
    """Launch Playwright and navigate to Odoo."""
    if not ODOO_URL or not ODOO_DASHBOARD_URL or not ODOO_JOURNAL_CREATE_URL:
        print("❌ ODOO_URL, ODOO_DASHBOARD_URL, atau ODOO_JOURNAL_CREATE_URL tidak ada di .env")
        sys.exit(1)

    with sync_playwright() as p:
        user_data_dir = os.path.abspath("playwright_profile")
        
        try:
            # Gunakan persistent context agar cookie & HSTS cache tersimpan.
            # Ini mencegah Nginx 404 loop karena HTTP/HTTPS dan menghindari harus login tiap saat.
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

        # The persistent context automatically has one page
        page = context.pages[0] if context.pages else context.new_page()
        
        # Hide webdriver flag to bypass bot detection
        page.add_init_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")

        # -------------------------------------------------------------
        # Fix 404 Nginx Loop by upgrading HTTP to HTTPS automatically
        # -------------------------------------------------------------
        def upgrade_to_https(route, request):
            if request.url.startswith("http://"):
                secure_url = request.url.replace("http://", "https://", 1)
                # Playwright doesn't allow protocol change in route.continue_()
                # So we mock a 301 Redirect to force the browser to HTTPS
                route.fulfill(
                    status=301,
                    headers={"Location": secure_url}
                )
            else:
                route.continue_()

        # Apply the interceptor to ALL requests to guarantee HTTP -> HTTPS upgrade
        page.route("**/*", upgrade_to_https)
        # -------------------------------------------------------------

        print(f"🌐 Membuka {ODOO_URL}")
        page.goto(ODOO_URL)

        print("\n⏳ Menunggu Anda login secara manual...")
        print(f"Script akan lanjut setelah URL mengandung:\n{ODOO_DASHBOARD_URL}")
        
        # Wait until the URL contains the dashboard URL string
        page.wait_for_function(
            "dashboardUrl => window.location.href.includes(dashboardUrl)",
            arg=ODOO_DASHBOARD_URL,
            timeout=0
        )
        print("✅ Dashboard terdeteksi!")

        print(f"\n➡️ Mengarahkan ke form pembuatan jurnal: {ODOO_JOURNAL_CREATE_URL}")
        page.goto(ODOO_JOURNAL_CREATE_URL)
        
        print("\n🎉 Mulai Mengisi Form Jurnal (General Info)...")
        
        # XPaths provided by user
        xpath_ref = "/html/body/div[1]/div/div/div[2]/div/div[1]/div[2]/div[2]/div[1]/div/div[2]/div/input"
        xpath_date = "/html/body/div[1]/div/div/div[2]/div/div[1]/div[2]/div[2]/div[2]/div[1]/div[2]/div/div/input"
        xpath_journal = "/html/body/div[1]/div/div/div[2]/div/div[1]/div[2]/div[2]/div[2]/div[3]/div[2]/div/div[1]/div/div/input"

        for idx, item in enumerate(selected_items):
            bank = item["bank"]
            tanggal = item["tanggal"]  # expected string like 'YYYY-MM-DD' or similar
            
            # Use the global format dict, or a fallback if bank not in dict
            fmt = REF_FORMATS.get(bank, f"Settlement Journal EDC {bank.upper()} for {{tanggal}}")
            ref_text = fmt.format(tanggal=tanggal)
            
            print(f"\n📝 Memproses: {bank} - {tanggal}")
            
            # 1. Wait for the form fields to be visible
            page.wait_for_selector(f"xpath={xpath_ref}", state="visible", timeout=30000)
            
            # 2. Fill Reference
            page.locator(f"xpath={xpath_ref}").fill(ref_text)
            
            # 3. Fill Accounting Date
            # Odoo datepickers sometimes require clearing or pressing Enter
            page.locator(f"xpath={xpath_date}").fill(tanggal)
            page.keyboard.press("Escape") # close datepicker if open
            
            # 4. Fill Journal Selection
            # Clear existing, type, and press enter/click dropdown
            page.locator(f"xpath={xpath_journal}").fill("EDC Settlement Journal")
            # Wait for Odoo's dynamic dropdown to show matching result, then press enter
            page.wait_for_timeout(1000) 
            page.keyboard.press("Enter")
            
            print(f"✅ Selesai mengisi info untuk {bank}. (Hanya 1 item untuk testing)")
            break # STOP after 1 item for user verification

        input("\n[Cek browser. Tekan Enter di terminal ini untuk menutup browser dan keluar]")
        browser.close()


def main():
    parser = argparse.ArgumentParser(description="Automate Odoo Journal Creation")
    parser.add_argument("--file", type=str, help="Path ke file Excel rekonsiliasi (opsional)")
    args = parser.parse_args()

    # 1. Get File
    if args.file:
        file_path = Path(args.file)
        if not file_path.exists():
            print(f"❌ File tidak ditemukan: {file_path}")
            sys.exit(1)
    else:
        file_path = get_latest_excel_file()
    
    print(f"📁 Menggunakan file: {file_path.name}")

    # 2. Extract Data
    sesuai_items = read_sesuai_dates(file_path)
    if not sesuai_items:
        print("ℹ️ Tidak ada data dengan status 'Sesuai' di Ringkasan Harian.")
        sys.exit(0)

    # 3. Present CLI Selection
    print("\n📋 Data yang sesuai dan siap dibuat jurnal:")
    for idx, item in enumerate(sesuai_items, 1):
        print(f"  [{idx}] {item['bank']:<10} - {item['tanggal']}")

    print("\n💡 Secara default, semua data di atas akan diproses.")
    exclude_input = input("Masukkan nomor yang INGIN DIKECUALIKAN (pisahkan dengan koma), atau tekan Enter untuk lanjut semua: ")
    
    excludes = []
    if exclude_input.strip():
        try:
            excludes = [int(x.strip()) for x in exclude_input.split(",")]
        except ValueError:
            print("❌ Input tidak valid. Pastikan hanya memasukkan angka yang dipisah koma.")
            sys.exit(1)

    # 4. Filter selected items
    selected_items = []
    for idx, item in enumerate(sesuai_items, 1):
        if idx not in excludes:
            selected_items.append(item)

    if not selected_items:
        print("ℹ️ Tidak ada data yang dipilih.")
        sys.exit(0)

    # 5. Run Automation
    run_automation(selected_items)


if __name__ == "__main__":
    main()
