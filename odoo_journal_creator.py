"""
odoo_journal_creator.py — Script to automate Odoo journal creation via Excel Import.
Phase 2: Uploads the generated Excel file to the Odoo Import page.
"""

import os
import sys
import argparse
from pathlib import Path
from playwright.sync_api import sync_playwright

# Import configs
from config import (
    ODOO_URL, ODOO_DASHBOARD_URL, ODOO_JOURNAL_IMPORT_URL, OUTPUT_DIR
)
from journal_generator import generate_journal_import


def get_latest_excel_file() -> Path:
    """Find the most recently created .xlsx file in the output directory."""
    excel_files = list(OUTPUT_DIR.glob("reconciliation_*.xlsx"))
    if not excel_files:
        print(f"❌ No Excel file found in '{OUTPUT_DIR}'.")
        return None
    # Sort by modification time, newest last
    latest_file = max(excel_files, key=os.path.getmtime)
    return latest_file


def update_recon_file_status(recon_file: Path, config_path: Path):
    """Update Journal Status column in the Daily Summary sheet based on config."""
    from openpyxl import load_workbook
    import json
    
    try:
        if not config_path or not config_path.exists():
            print("⚠️ No config path provided, skipping excel status update.")
            return

        with open(config_path, "r") as f:
            selected_items = json.load(f)
            
        wb = load_workbook(recon_file)
        if "Daily Summary" in wb.sheetnames:
            ws = wb["Daily Summary"]
            updated = 0
            
            for item in selected_items:
                row_idx = item.get("row")
                if not row_idx:
                    continue
                
                created_edc = item.get("create_edc", False)
                created_ar = item.get("create_ar", False)
                
                new_status = ""
                if created_edc and created_ar:
                    new_status = "✅ Both"
                elif created_edc:
                    new_status = "✅ EDC"
                elif created_ar:
                    new_status = "✅ AR"
                else:
                    continue
                    
                current_val = str(ws.cell(row=row_idx, column=10).value)
                
                # Merge status
                if "EDC" in current_val and created_ar:
                    new_status = "✅ Both"
                elif "AR" in current_val and created_edc:
                    new_status = "✅ Both"
                elif "Both" in current_val:
                    new_status = "✅ Both"
                
                ws.cell(row=row_idx, column=10, value=new_status)
                
                # Keep original checkmark logic
                status9 = ws.cell(row=row_idx, column=9).value
                if status9 and "Match" in str(status9):
                    ws.cell(row=row_idx, column=9, value="✅ Journal Created")
                    
                updated += 1
                
            if updated > 0:
                wb.save(recon_file)
                print(f"✅ Successfully updated {updated} rows in '{recon_file.name}' with Journal Status.")
    except Exception as e:
        print(f"⚠️ Failed to update status in reconciliation file: {e}")

def run_import_automation(import_file: Path, recon_file: Path, config_path: Path = None):
    """Launch Playwright, navigate to Odoo, and upload the import file."""
    if not ODOO_URL or not ODOO_DASHBOARD_URL or not ODOO_JOURNAL_IMPORT_URL:
        print("❌ ODOO_URL, ODOO_DASHBOARD_URL, or ODOO_JOURNAL_IMPORT_URL missing from .env")
        sys.exit(1)

    with sync_playwright() as p:
        user_data_dir = os.path.abspath("playwright_profile")
        
        try:
            # Use persistent context to save cookies & HSTS cache
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

        print(f"🌐 Opening {ODOO_URL}")
        page.goto(ODOO_URL)

        print("\n⏳ Waiting for manual login...")
        print(f"Script will continue after URL includes:\n{ODOO_DASHBOARD_URL}")
        
        page.wait_for_function(
            "dashboardUrl => window.location.href.includes(dashboardUrl)",
            arg=ODOO_DASHBOARD_URL,
            timeout=0
        )
        print("✅ Dashboard terdeteksi!")

        print(f"\n➡️ Mengarahkan ke form Import Jurnal: {ODOO_JOURNAL_IMPORT_URL}")
        page.goto(ODOO_JOURNAL_IMPORT_URL)
        
        # Tunggu sampai halaman import termuat sepenuhnya
        page.wait_for_load_state("domcontentloaded")
        page.locator("xpath=/html/body/div[1]/div/div[1]/div/div[1]/div[1]/div[2]/span/span/button").wait_for(state="visible", timeout=30000)

        # Klik tombol Upload File (menggunakan FileChooser API Playwright)
        print(f"\n📤 Mengunggah file: {import_file.name} ...")
        xpath_upload_btn = "xpath=/html/body/div[1]/div/div[1]/div/div[1]/div[1]/div[2]/span/span/button"
        
        try:
            with page.expect_file_chooser(timeout=10000) as fc_info:
                page.locator(xpath_upload_btn).click()
            
            file_chooser = fc_info.value
            file_chooser.set_files(str(import_file.resolve()))
            print("✅ File berhasil diunggah!")
        except Exception as e:
            print(f"❌ Gagal mengunggah file. Pastikan tombol 'Upload File' terlihat.\nError: {e}")
            input("\n[Tekan Enter untuk menutup browser]")
            context.close()
            return

        # 4. Automate Test & Import
        print("🔍 Melakukan Test Validasi...")
        xpath_test_btn = "xpath=/html/body/div[1]/div/div[1]/div/div[1]/div[1]/div[2]/button[2]"
        xpath_valid_msg = "xpath=/html/body/div[1]/div/div[2]/div[2]/div/p"
        xpath_import_btn = "xpath=/html/body/div[1]/div/div[1]/div/div[1]/div[1]/div[2]/button[1]"
        
        try:
            # Tunggu tombol test muncul setelah upload file selesai diparse
            page.locator(xpath_test_btn).wait_for(state="visible", timeout=30000)
            
            # Click Test
            page.locator(xpath_test_btn).click()
            
            # Wait for validation message
            valid_locator = page.locator(xpath_valid_msg)
            valid_locator.wait_for(state="visible", timeout=15000)
            
            msg_text = valid_locator.inner_text()
            if "Everything seems valid" in msg_text:
                print("✅ Validasi sukses: 'Everything seems valid.'")
                print("🚀 Melakukan Import...")
                page.locator(xpath_import_btn).click()
                
                
                # Wait for page to redirect to Journal Entries list (action=account.action_move_journal_line)
                try:
                    page.wait_for_url("**/web#action=*", timeout=15000)
                    print("✅ Berhasil dialihkan ke halaman Journal Entries.")
                    # Update status in Excel since import is successful
                    update_recon_file_status(recon_file, config_path)
                except Exception as e:
                    print(f"⚠️ Import mungkin sukses, tapi timeout saat menunggu dialihkan ke Journal Entries: {e}")
                    
                print("🎉 Proses import selesai!")
                page.wait_for_timeout(2000)
            else:
                print(f"⚠️ Validasi Odoo menampilkan pesan lain: {msg_text}")
                print("Tolong periksa manual di browser.")
                input("\n[Tekan Enter di terminal ini untuk menutup browser]")
        except Exception as e:
            print(f"❌ Terjadi kesalahan saat melakukan Test/Import.\nError: {e}")
            input("\n[Tekan Enter di terminal ini untuk menutup browser]")
            
        context.close()


def main():
    parser = argparse.ArgumentParser(description="Automate Odoo Journal Creation via Import")
    parser.add_argument("--file", type=str, help="Path ke file Excel rekonsiliasi (opsional)")
    parser.add_argument("--config", type=str, help="Path ke file JSON konfigurasi jurnal (opsional)")
    parser.add_argument("--import-file", type=str, help="Path ke file import Excel yang sudah siap upload (opsional)")
    args = parser.parse_args()

    # 1. Get File
    if args.file:
        file_path = Path(args.file)
        if not file_path.exists():
            print(f"❌ File tidak ditemukan: {file_path}")
            sys.exit(1)
    else:
        file_path = get_latest_excel_file()
        if not file_path:
            sys.exit(1)
    
    print(f"📁 Memproses data dari: {file_path.name}")

    config_path = Path(args.config) if args.config else None

    # 2. Generate Import Excel or use existing
    if args.import_file:
        import_file = Path(args.import_file)
        if not import_file.exists():
            print(f"❌ File import tidak ditemukan: {import_file}")
            sys.exit(1)
    else:
        import_file = generate_journal_import(file_path, config_path)
        
        if not import_file:
            sys.exit(1)

    # 3. Run Browser Automation
    run_import_automation(import_file, file_path, config_path)


if __name__ == "__main__":
    main()
